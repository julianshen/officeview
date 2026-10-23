/**
 * Morph rendering (slideshow spec 6.3, S3-6): one composited frame of a
 * Morph transition, on either backend (WebGL2 or Canvas 2D: per-layer
 * transforms plus crossfade, which is all a compositor `frame` needs).
 *
 * The brief's `morphFrame(pairs, progress, compositor): void` moves every
 * matched object through the linearly interpolated box (position, size,
 * rotation on the shortest path), and blends the two sides:
 *
 * - `under` (same outline): the two sides are summed alpha-correctly
 *   (`LayerDraw.sum`, PR #20 review 2): the outgoing side at `1 - p` and the
 *   incoming at `p`, added in premultiplied colour and alpha on the
 *   backend's scratch target, then composited source-over. The pair is
 *   exactly `(1 - p) * from + p * to` in colour and alpha, so a translucent
 *   pair keeps its opacity and an opaque one never dips. The stage uses it
 *   for the backgrounds and for matched shapes whose outline is the same.
 *   Without the scratch the outgoing side holds at full weight under the
 *   incoming one (`heldOpacity`): exact for opaque shapes.
 * - otherwise the sides crossfade (`1 - p` / `p`): changed content (text,
 *   a different outline or picture) blends rather than snaps.
 *
 * Unmatched objects fade at their own base box. Items paint in list order
 * (the stage emits them in z order), then `over` (the incoming watermark).
 * Morph is always an approximation (spec 10), diagnosed by the dispatch.
 * Reduced motion never reaches this module (`reducedTransition`).
 */
import { mul } from '../player/hit-test'
import type { BaseGeometry } from '../engine/behaviours'
import type { Affine, Compositor, LayerDraw, LayerHandle, TargetKey, TargetState } from '../types'

export interface MorphLayerFrame {
  /** The layer's compositor handle. */
  handle: LayerHandle
  /** Base geometry in slide px (`p.base`), the box the bitmap paints. */
  geometry: BaseGeometry
  /** The layer's engine state at the frame's base (motion deltas composed). */
  state?: TargetState
  /** Split part the layer carries, if it paints only fill/line/text. */
  part?: 'fill' | 'line' | 'text'
  /** The top-level object the layer belongs to (diagnostics and tests). */
  key?: TargetKey
}

/** One item: both sides move through the interpolated box; a lone side
 * fades at its own box. */
export interface MorphFrameItem {
  from?: MorphLayerFrame
  to?: MorphLayerFrame
  /** The outgoing side holds at full weight under the incoming one (same
   * outline: linear colour, no opacity dip); otherwise the sides crossfade. */
  under?: boolean
  /** The slides' backgrounds (bottom-most). */
  background?: boolean
  /** A matched object's two base boxes: its layers move through the
   * interpolated box (a lone side too). Defaults to the sides' own boxes
   * when both are present. */
  path?: { from: BaseGeometry; to: BaseGeometry }
}

export interface MorphFrameSpec {
  /** Items bottom-most first (z order). */
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

const rotation = (deg: number): Affine => {
  const r = deg * Math.PI / 180
  const c = Math.cos(r), s = Math.sin(r)
  return [c, s, -s, c, 0, 0]
}

/** An affine that maps the base box's content (its own rotation baked in)
 * onto the target box, decomposed so a rotated shape never shears:
 * `T(target centre) · R(target.rot) · S(sx, sy) · R(-base.rot) · T(-base centre)`. */
export function mapTransform(base: BaseGeometry, target: BaseGeometry): Affine {
  const sx = base.w > 0 ? target.w / base.w : 1
  const sy = base.h > 0 ? target.h / base.h : 1
  const bx = base.x + base.w / 2, by = base.y + base.h / 2
  const cx = target.x + target.w / 2, cy = target.y + target.h / 2
  // The target rotation as base + shortest delta, so the frame's rotation is continuous.
  const rot = base.rot + rotationDelta(base.rot, target.rot)
  let m: Affine = [1, 0, 0, 1, -bx, -by]
  m = mul(rotation(-base.rot), m)
  m = mul([sx, 0, 0, sy, 0, 0], m)
  m = mul(rotation(rot), m)
  m = mul([1, 0, 0, 1, cx, cy], m)
  return m
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
    const path = item.path ?? (item.from && item.to ? { from: item.from.geometry, to: item.to.geometry } : undefined)
    if (path) {
      // The matched pair moves together into the interpolated box.
      const g = interpGeometry(path.from, path.to, p)
      const a = item.from && sideDraw(item.from, g, 1 - p)
      const b = item.to && sideDraw(item.to, g, p)
      if (item.under && a && b) {
        // Summed alpha-correctly; the held weight is the backend's fallback.
        out.push({ ...a, sum: { with: b, heldOpacity: item.from!.state?.opacity ?? 1 } })
        continue
      }
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
