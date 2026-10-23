/**
 * Morph rendering (slideshow spec 6.3, S3-6): one composited frame of a
 * Morph transition.
 *
 * The brief's `morphFrame(pairs, progress, compositor): void` interpolates
 * the matched objects' base geometry (position, size, rotation; shortest
 * path for rotation) linearly with progress, crossfades the two sides so
 * changed content blends rather than snaps, and fades unmatched objects at
 * their own base box. Fill and line colour interpolation rides the cross-
 * fade weights: the sides' colour transforms draw unmixed, and the mix
 * approximates the colour change.
 *
 * The stage passes the matcher's pairs in gallery order and the incoming
 * slide's watermark as `over` (drawn on top, unchanged); statics crossfade
 * as fade items and the outgoing watermark never draws (the two watermarks
 * are identical). Morph is always an approximation (spec 10). Canvas 2D
 * composes it with per-layer transforms and global alpha, which is what its
 * `frame` sees; the stage routes the effect through this module on WebGL2
 * and falls back to the catalogue's fade look-alike everywhere else
 * (diagnosed there), and reduced motion never reaches this module.
 */
import { mul } from '../player/hit-test'
import type { BaseGeometry } from '../engine/behaviours'
import type { Affine, Compositor, LayerDraw, LayerHandle, TargetState } from '../types'

export interface MorphLayerFrame {
  /** The layer's compositor handle. */
  handle: LayerHandle
  /** Base geometry in slide px (`p.base`), the box the bitmap paints. */
  geometry: BaseGeometry
  /** The layer's engine state at the frame's base (motion deltas composed). */
  state?: TargetState
  /** Split part the layer carries, if it paints only fill/line/text. */
  part?: 'fill' | 'line' | 'text'
}

/** One animation item: both sides animate into the interpolated box; a fade
 * carries one side at its own box; a static crossfades (from out, to in). */
export interface MorphFrameItem {
  from?: MorphLayerFrame
  to?: MorphLayerFrame
}

export interface MorphFrameSpec {
  /** Matched pairs, fades and statics, bottom-most first (gallery order). */
  pairs: readonly MorphFrameItem[]
  /** Draws kept unchanged on top (the incoming slide's watermark). */
  over?: readonly LayerDraw[]
}

const IDENTITY: Affine = [1, 0, 0, 1, 0, 0]

/** The interpolated box at `progress`, rotation on the shortest path. */
export function interpGeometry(from: BaseGeometry, to: BaseGeometry, progress: number): BaseGeometry {
  return {
    x: from.x + (to.x - from.x) * progress,
    y: from.y + (to.y - from.y) * progress,
    w: from.w + (to.w - from.w) * progress,
    h: from.h + (to.h - from.h) * progress,
    rot: from.rot + rotationDelta(from.rot, to.rot) * progress,
  }
}

/** Shortest-path signed rotation difference, in degrees. */
export function rotationDelta(fromDeg: number, toDeg: number): number {
  let d = (toDeg - fromDeg) % 360
  if (d > 180) d -= 360
  else if (d < -180) d += 360
  return d
}

/** An affine that maps the base box's content onto the target box: a
 * centre-preserving scale, then the remaining rotation (target minus the
 * base rotation the bitmap already paints) about the target centre. */
export function mapTransform(base: BaseGeometry, target: BaseGeometry): Affine {
  const sx = base.w > 0 ? target.w / base.w : 1
  const sy = base.h > 0 ? target.h / base.h : 1
  const bx = base.x + base.w / 2, by = base.y + base.h / 2
  const cx = target.x + target.w / 2, cy = target.y + target.h / 2
  const scale: Affine = [sx, 0, 0, sy, cx - sx * bx, cy - sy * by]
  const delta = rotationDelta(base.rot, target.rot) * Math.PI / 180
  if (Math.abs(delta) < 1e-9) return scale
  const c = Math.cos(delta), s = Math.sin(delta)
  return mul([c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy], scale)
}

/** One side's layer draw at weight `w` mapped onto the box `g`. */
function sideDraw(side: MorphLayerFrame, g: BaseGeometry, w: number): LayerDraw | undefined {
  const state = side.state ?? { visible: true, opacity: 1, transform: IDENTITY }
  if (!state.visible) return undefined
  const opacity = state.opacity * w
  if (!(opacity > 0)) return undefined
  return {
    handle: side.handle,
    state: { ...state, opacity, transform: mul(mapTransform(side.geometry, g), state.transform) },
    ...(side.part ? { part: side.part } : {}),
  }
}

export function morphFrame(pairs: MorphFrameSpec, progress: number, compositor: Compositor): void {
  const p = progress < 0 ? 0 : progress > 1 ? 1 : progress
  const out: LayerDraw[] = []
  for (const item of pairs.pairs) {
    if (item.from && item.to) {
      // The matched pair moves together into the interpolated box; the sides
      // crossfade, which also carries colour and content changes.
      const g = interpGeometry(item.from.geometry, item.to.geometry, p)
      const a = sideDraw(item.from, g, 1 - p)
      const b = sideDraw(item.to, g, p)
      if (a) out.push(a)
      if (b) out.push(b)
    } else {
      // An unmatched object fades at its own base box.
      const side = item.from ?? item.to
      if (!side) continue
      const d = sideDraw(side, side.geometry, item.from ? 1 - p : p)
      if (d) out.push(d)
    }
  }
  for (const draw of pairs.over ?? []) out.push(draw)
  compositor.frame(out)
}
