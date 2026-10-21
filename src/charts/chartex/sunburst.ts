/**
 * Sunburst (docs/charts/chartex-map.md §7; lazy chunk): one ring per depth, the root level innermost, each ring
 * R / levels thick; a node's angle is proportional to its subtree weight, children in source order from their parent's
 * start, clockwise from 12 o'clock. `parentLabelLayout` does not apply. Office does not document the algorithm: only
 * the angle and ring proportions are oracles; the hole-less centre and the label placement are project rules [U].
 */
import type { ChartModel } from '../model'
import type { ChartExHierarchy } from './hierarchy'
import type { Box, HierarchyPaintKit } from './treemap'
import { hierarchyDrawing, hierarchyLabelPx, nodeColor, nodeText, textColor } from './hierarchyStyle'

/** Four numbers per node, in pre-order: start angle, sweep (radians), inner and outer radius. */
export function sunburstLayout(tree: ChartExHierarchy, radius: number, out = new Float64Array(tree.count * 4)): Float64Array {
  const ring = tree.levels ? radius / tree.levels : 0
  // A parent's children start where it starts; `cursor` keeps each parent's next child angle.
  const cursor = new Float64Array(tree.count)
  let rootAngle = -Math.PI / 2
  for (let i = 0; i < tree.count; i++) {
    const p = tree.parent[i]
    const sweep = p < 0 ? (tree.total > 0 ? tree.weight[i] / tree.total * Math.PI * 2 : 0) : (tree.weight[p] > 0 ? out[p * 4 + 1] * tree.weight[i] / tree.weight[p] : 0)
    const start = p < 0 ? rootAngle : cursor[p]
    if (p < 0) rootAngle += sweep; else cursor[p] += sweep
    cursor[i] = start
    out[i * 4] = start; out[i * 4 + 1] = sweep; out[i * 4 + 2] = tree.depth[i] * ring; out[i * 4 + 3] = (tree.depth[i] + 1) * ring
  }
  return out
}

/** The geometry a sunburst paint uses for `area`: its centre, radius and segments. */
export function sunburstPaintGeometry(model: ChartModel, area: Box): { cx: number; cy: number; radius: number; segments: Float64Array } {
  const drawing = hierarchyDrawing(model)
  const cx = area.x + area.w / 2, cy = area.y + area.h / 2, radius = Math.max(1, Math.min(area.w, area.h) / 2 - 4)
  return { cx, cy, radius, segments: drawing ? sunburstLayout(drawing.tree, radius) : new Float64Array(0) }
}

/** Paints a prepared sunburst into `area` (paint-local geometry admitted by the caller). */
export function paintSunburst(ctx: CanvasRenderingContext2D, model: ChartModel, area: Box, kit: HierarchyPaintKit): void {
  const drawing = hierarchyDrawing(model)
  if (!drawing || !drawing.tree.count) return
  const { series, tree } = drawing
  const { cx, cy, radius, segments: s } = sunburstPaintGeometry(model, area)
  const box = { x: cx - radius, y: cy - radius, w: radius * 2, h: radius * 2 }
  for (let i = 0; i < tree.count; i++) {
    const start = s[i * 4], sweep = s[i * 4 + 1], r0 = s[i * 4 + 2], r1 = s[i * 4 + 3]
    if (!(sweep > 0)) continue
    ctx.beginPath()
    ctx.arc(cx, cy, r1, start, start + sweep)
    if (r0 > 0) ctx.arc(cx, cy, r0, start + sweep, start, true); else ctx.lineTo(cx, cy)
    ctx.closePath()
    const { fill, color } = nodeColor(drawing, i)
    if (kit.fillWith(ctx, fill, color, box)) ctx.fill()
    if (series.lineStyle) kit.strokeBorder(ctx, series.lineStyle, 'FFFFFF', box)
    else { ctx.strokeStyle = '#FFFFFF'; ctx.lineWidth = 1; ctx.stroke() }
  }
  const px = hierarchyLabelPx(model, series), font = kit.font(px)
  ctx.font = font; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'
  for (let i = 0; i < tree.count; i++) {
    const text = nodeText(series, tree, i)
    if (!text) continue
    const start = s[i * 4], sweep = s[i * 4 + 1], r0 = s[i * 4 + 2], r1 = s[i * 4 + 3], mid = (r0 + r1) / 2
    // Drawn horizontally at the segment's centre when it fits the ring's thickness and the arc at mid radius.
    const room = Math.min(r1 - r0, sweep >= Math.PI ? Infinity : 2 * mid * Math.sin(sweep / 2))
    if (kit.width(text, font) + 4 > room || px + 2 > (sweep >= Math.PI ? Infinity : 2 * mid * Math.sin(sweep / 2))) continue
    const angle = start + sweep / 2
    ctx.fillStyle = textColor(series, nodeColor(drawing, i).color)
    ctx.fillText(text, cx + Math.cos(angle) * mid, cy + Math.sin(angle) * mid)
  }
}
