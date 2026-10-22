/**
 * Click targets (slideshow spec 4.3): a click is hit-tested against the
 * current render plan in authored z-order, with each target's current
 * transform, clip (its mask) and visibility. The topmost shape that has a
 * trigger or a hyperlink receives it; shapes without either never block one
 * below them (spec: "the topmost shape that has a trigger or a hyperlink").
 *
 * Geometry follows `paintNode`'s transform chain (the same walk as
 * `paintedBounds` in render/scene.ts): a node's local box `[0, w] × [0, h]`
 * maps to slide px through its ancestors' group transforms and its own
 * offset, rotation and flips. The point is mapped back through the inverse
 * of every matrix rather than the shapes forward.
 *
 * Masks clip in the target's own space over the unpadded box of the layer's
 * nodes (the compositor uses the padded raster bounds; the difference is the
 * stroke and anti-aliasing padding, a few px at the edge).
 */
import { emuToPx } from '../../core/geometry'
import type { PptxShape } from '../../pptx/types'
import { createRng } from '../rng'
import type { Affine, MaskState, RenderState, TargetState } from '../types'
import { isMaskRaster, maskPath, type MaskOp } from '../render/masks'
import type { PlanLayer, RenderPlan } from '../render/plan'

const IDENTITY: Affine = [1, 0, 0, 1, 0, 0]

/** `m1 ∘ m2`: apply `m2`, then `m1`. */
export function mul(m1: Affine, m2: Affine): Affine {
  const [a, b, c, d, e, f] = m1, [A, B, C, D, E, F] = m2
  return [a * A + c * B, b * A + d * B, a * C + c * D, b * C + d * D, a * E + c * F + e, b * E + d * F + f]
}

export function invert(m: Affine): Affine | undefined {
  const [a, b, c, d, e, f] = m
  const det = a * d - b * c
  if (!Number.isFinite(det) || Math.abs(det) < 1e-12) return undefined
  return [d / det, -b / det, -c / det, a / det, (c * f - d * e) / det, (b * e - a * f) / det]
}

export function apply(m: Affine, x: number, y: number): [number, number] {
  return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]]
}

/** A target's draw matrix in slide px: `translate(motion) ∘ transform` (the compositor contract). */
export function stateMatrix(s: TargetState | undefined): Affine {
  if (!s) return IDENTITY
  const t = s.transform.every(Number.isFinite) ? s.transform : IDENTITY
  const dx = s.motion && Number.isFinite(s.motion.dx) ? s.motion.dx : 0
  const dy = s.motion && Number.isFinite(s.motion.dy) ? s.motion.dy : 0
  return dx || dy ? mul([1, 0, 0, 1, dx, dy], t) : t
}

/** The local box `[0, w] × [0, h]` → parent space for one node; `undefined` when its transform is unusable. */
export function nodeMatrix(shape: PptxShape, parent: Affine): { m: Affine; w: number; h: number } | undefined {
  const x = emuToPx(shape.xEmu), y = emuToPx(shape.yEmu), w = emuToPx(shape.widthEmu), h = emuToPx(shape.heightEmu)
  if (shape.transformValid === false || ![x, y, w, h, shape.rotationDeg ?? 0].every(Number.isFinite) || w < 0 || h < 0) return undefined
  let m = mul(parent, [1, 0, 0, 1, x + w / 2, y + h / 2])
  if (shape.rotationDeg) {
    const r = shape.rotationDeg * Math.PI / 180, c = Math.cos(r), s = Math.sin(r)
    m = mul(m, [c, s, -s, c, 0, 0])
  }
  if (shape.flipH || shape.flipV) m = mul(m, [shape.flipH ? -1 : 1, 0, 0, shape.flipV ? -1 : 1, 0, 0])
  return { m: mul(m, [1, 0, 0, 1, -w / 2, -h / 2]), w, h }
}

/** A group's child space → parent space (after `nodeMatrix`). */
export function groupChildMatrix(shape: PptxShape, local: Affine): Affine | undefined {
  const g = shape.group
  if (!g || !(g.chExt.width > 0 && g.chExt.height > 0)) return undefined
  return mul(mul(local, [g.ext.width / g.chExt.width, 0, 0, g.ext.height / g.chExt.height, 0, 0]), [1, 0, 0, 1, -emuToPx(g.chOff.x), -emuToPx(g.chOff.y)])
}

/** Calls `visit` for every leaf with its local → slide matrix, in paint order. */
export function walkLeaves(nodes: readonly PptxShape[], parent: Affine, visit: (shape: PptxShape, m: Affine, w: number, h: number, path: readonly PptxShape[]) => void, path: readonly PptxShape[] = []): void {
  for (const shape of nodes) {
    const local = nodeMatrix(shape, parent)
    if (!local) continue
    if (shape.group) {
      const child = groupChildMatrix(shape, local.m)
      if (child) walkLeaves(shape.children ?? [], child, visit, [...path, shape])
      continue
    }
    visit(shape, local.m, local.w, local.h, path)
  }
}

/** The union of the nodes' unpadded boxes in the space `parent` maps into. */
function nodesBox(nodes: readonly PptxShape[]): { x: number; y: number; w: number; h: number } | undefined {
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity
  walkLeaves(nodes, IDENTITY, (_s, m, w, h) => {
    for (const [px, py] of [apply(m, 0, 0), apply(m, w, 0), apply(m, w, h), apply(m, 0, h)]) {
      minX = Math.min(minX, px); minY = Math.min(minY, py); maxX = Math.max(maxX, px); maxY = Math.max(maxY, py)
    }
  })
  return maxX > minX && maxY > minY ? { x: minX, y: minY, w: maxX - minX, h: maxY - minY } : undefined
}

function inOp(op: MaskOp, x: number, y: number): boolean {
  if (op.op === 'rect') return x >= op.x && x <= op.x + op.w && y >= op.y && y <= op.y + op.h
  if (op.op === 'sector') {
    const dx = x - op.cx, dy = y - op.cy
    if (dx * dx + dy * dy > op.r * op.r) return false
    const span = op.a1 - op.a0
    if (span >= 2 * Math.PI - 1e-9) return true
    let a = Math.atan2(dy, dx) - op.a0
    a = ((a % (2 * Math.PI)) + 2 * Math.PI) % (2 * Math.PI)
    return a <= span
  }
  const p = op.pts
  let inside = false
  for (let i = 0, j = p.length - 2; i < p.length; j = i, i += 2) {
    const xi = p[i], yi = p[i + 1], xj = p[j], yj = p[j + 1]
    if ((yi > y) !== (yj > y) && x < ((xj - xi) * (y - yi)) / (yj - yi) + xi) inside = !inside
  }
  return inside
}

/** Whether point (x, y) in the mask's box space (`box` in the same space) is shown by the mask. */
export function maskShows(mask: MaskState, box: { x: number; y: number; w: number; h: number }, x: number, y: number): boolean {
  const seed = typeof mask.params.seed === 'number' ? mask.params.seed : 0
  const m = maskPath(mask.kind, mask.progress, mask.params, box.w, box.h, createRng(seed))
  const lx = x - box.x, ly = y - box.y
  if (lx < 0 || ly < 0 || lx > box.w || ly > box.h) return false
  if (isMaskRaster(m)) {
    const cx = Math.min(m.width - 1, Math.floor((lx / box.w) * m.width))
    const cy = Math.min(m.height - 1, Math.floor((ly / box.h) * m.height))
    return (m.alpha[cy * m.width + cx] ?? 0) > 0
  }
  // Even-odd counts every covering sub-path; non-zero shows where any sub-path covers (the library's ops never cancel).
  const covering = m.ops.filter(op => inOp(op, lx, ly)).length
  return m.fillRule === 'evenodd' ? covering % 2 === 1 : covering > 0
}

const spidOf = (shape: PptxShape): string | undefined => shape.source?.id

/** The topmost interactive node under (x, y) in `nodes`, walking in reverse paint order. */
function hitNodes(nodes: readonly PptxShape[], parent: Affine, x: number, y: number, interactive: ReadonlySet<string>, owner?: string): string | undefined {
  for (let i = nodes.length - 1; i >= 0; i--) {
    const shape = nodes[i]
    const local = nodeMatrix(shape, parent)
    if (!local) continue
    const id = spidOf(shape)
    // A trigger or link on a group belongs to every leaf inside it (the innermost interactive node wins).
    const mine = id !== undefined && interactive.has(id) ? id : owner
    if (shape.group) {
      const child = groupChildMatrix(shape, local.m)
      const hit = child && hitNodes(shape.children ?? [], child, x, y, interactive, mine)
      if (hit) return hit
      continue
    }
    if (mine === undefined) continue
    const inv = invert(local.m)
    if (!inv) continue
    const [u, v] = apply(inv, x, y)
    if (u >= 0 && v >= 0 && u <= local.w && v <= local.h) return mine
  }
  return undefined
}

function hitLayers(layers: readonly PlanLayer[], state: RenderState, parent: Affine, x: number, y: number, interactive: ReadonlySet<string>): string | undefined {
  for (let i = layers.length - 1; i >= 0; i--) {
    const layer = layers[i]
    if (layer.kind === 'watermark') continue
    if (layer.kind === 'static') {
      const hit = hitNodes(layer.nodes, parent, x, y, interactive)
      if (hit) return hit
      continue
    }
    const s = state.get(layer.key)
    if (s && (!s.visible || !(s.opacity > 0))) continue
    const m = mul(parent, stateMatrix(s))
    if (s?.mask) {
      // The mask lives in the target's own space, over its layer box.
      const inv = invert(m)
      const box = nodesBox(layer.nodes)
      if (!inv || !box) continue
      const [lx, ly] = apply(inv, x, y)
      if (!maskShows(s.mask, box, lx, ly)) continue
    }
    const hit = layer.sub ? hitLayers(layer.sub, state, m, x, y, interactive) : hitNodes(layer.nodes, m, x, y, interactive)
    if (hit) return hit
  }
  return undefined
}

/**
 * The spid of the topmost trigger or link shape at slide-px point (x, y),
 * or `undefined` (the click advances). `interactive` holds the spids of the
 * slide's trigger shapes and hyperlinked shapes.
 */
export function hitTest(plan: RenderPlan, state: RenderState, x: number, y: number, interactive: ReadonlySet<string>): string | undefined {
  if (interactive.size === 0 || !Number.isFinite(x) || !Number.isFinite(y)) return undefined
  return hitLayers(plan.layers, state, IDENTITY, x, y, interactive)
}
