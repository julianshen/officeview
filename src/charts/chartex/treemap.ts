/**
 * Treemap (docs/charts/chartex-map.md §7; lazy chunk): squarified rectangles (Bruls, Huizing and van Wijk), each node's
 * area proportional to its subtree weight within its parent's content area. `cx:parentLabelLayout`:
 * - `none`: parents carry no label; children fill the parent;
 * - `banner`: a strip one label line high across the top of each parent holds its label; children fill the rest;
 * - `overlapping` (the default when absent, unverified): children fill the parent and its label is drawn over them.
 * Office does not document its algorithm, so only the area proportions are oracles; the placement is a project rule.
 */
import type { ChartModel } from '../model'
import type { ChartFillStyle } from '../fills'
import type { ChartLineStyle } from '../style'
import type { ChartExHierarchy } from './hierarchy'
import { hierarchyDrawing, hierarchyLabelPx, isLeaf, nodeColor, nodeText, parentLabelLayout, textColor } from './hierarchyStyle'

export interface Box { x: number; y: number; w: number; h: number }

/**
 * Squarified layout of `tree` in `box`: four numbers per node (x, y, w, h), in pre-order. Siblings are placed in
 * descending weight (ties in source order), row by row along the shorter side of the space left, each row kept while
 * adding a sibling does not worsen its worst aspect ratio. A banner parent (`banner` > 0) lays its children out below
 * its strip of height min(banner, h).
 */
export function treemapLayout(tree: ChartExHierarchy, box: Box, layout: 'none' | 'banner' | 'overlapping', banner: number, out = new Float64Array(tree.count * 4)): Float64Array {
  const roots: number[] = []
  for (let i = 0; i < tree.count; i = tree.end[i]) roots.push(i)
  squarify(tree, roots, tree.total, box, out)
  for (let i = 0; i < tree.count; i++) {
    if (isLeaf(tree, i)) continue
    const x = out[i * 4], y = out[i * 4 + 1], w = out[i * 4 + 2], h = out[i * 4 + 3]
    const b = layout === 'banner' ? Math.min(Math.max(0, banner), h) : 0
    const children: number[] = []
    for (let c = i + 1; c < tree.end[i]; c = tree.end[c]) children.push(c)
    squarify(tree, children, tree.weight[i], { x, y: y + b, w, h: h - b }, out)
  }
  return out
}
function squarify(tree: ChartExHierarchy, items: number[], total: number, rect: Box, out: Float64Array): void {
  const order = items.slice().sort((a, b) => tree.weight[b] - tree.weight[a])
  let { x, y, w, h } = rect
  if (!(total > 0) || !(w > 0) || !(h > 0)) { for (const i of order) out.set([x, y, 0, 0], i * 4); return }
  const scale = (w * h) / total
  /** The worst aspect ratio of a row of areas summing to `sum` (largest `max`, smallest `min`) along `side`. */
  const worst = (sum: number, max: number, min: number, side: number) => { const s2 = side * side, a2 = sum * sum; return Math.max(s2 * max / a2, a2 / (s2 * min)) }
  let i = 0
  while (i < order.length) {
    const column = w >= h, side = column ? h : w
    let sum = tree.weight[order[i]] * scale, max = sum, min = sum, j = i + 1
    let ratio = worst(sum, max, min, side)
    while (j < order.length) {
      const a = tree.weight[order[j]] * scale
      const next = worst(sum + a, Math.max(max, a), Math.min(min, a), side)
      if (next > ratio) break
      sum += a; max = Math.max(max, a); min = Math.min(min, a); ratio = next; j++
    }
    const thickness = sum / side
    let offset = 0
    for (let k = i; k < j; k++) {
      const length = tree.weight[order[k]] * scale / thickness, at = order[k] * 4
      if (column) { out[at] = x; out[at + 1] = y + offset; out[at + 2] = thickness; out[at + 3] = length }
      else { out[at] = x + offset; out[at + 1] = y; out[at + 2] = length; out[at + 3] = thickness }
      offset += length
    }
    if (column) { x += thickness; w -= thickness } else { y += thickness; h -= thickness }
    i = j
  }
}

/** The geometry a treemap paint uses for `area` (the layout's chart area): the node rectangles and the banner height. */
export function treemapPaintGeometry(model: ChartModel, area: Box): { rects: Float64Array; banner: number; layout: 'none' | 'banner' | 'overlapping' } {
  const drawing = hierarchyDrawing(model)
  if (!drawing) return { rects: new Float64Array(0), banner: 0, layout: 'none' }
  const layout = parentLabelLayout(drawing), banner = layout === 'banner' ? Math.ceil(hierarchyLabelPx(model, drawing.series) + 6) : 0
  return { rects: treemapLayout(drawing.tree, area, layout, banner), banner, layout }
}

export interface HierarchyPaintKit {
  /** Sets the fill for `box` (an authored fill, else `color`); false for noFill. */
  fillWith(ctx: CanvasRenderingContext2D, fill: ChartFillStyle | undefined, color: string, box: Box): boolean
  /** Strokes the current path (or `rect`) with an authored border. */
  strokeBorder(ctx: CanvasRenderingContext2D, line: ChartLineStyle | undefined, color: string, box: Box, rect?: Box): void
  /** A measured width (layout measured every node label). */
  width(text: string, font: string): number
  font(px: number): string
}
/** Paints a prepared treemap into `area` (paint-local geometry admitted by the caller). */
export function paintTreemap(ctx: CanvasRenderingContext2D, model: ChartModel, area: Box, kit: HierarchyPaintKit): void {
  const drawing = hierarchyDrawing(model)
  if (!drawing || !drawing.tree.count) return
  const { series, tree } = drawing
  const { rects, banner, layout } = treemapPaintGeometry(model, area)
  const rect = (i: number): Box => ({ x: rects[i * 4], y: rects[i * 4 + 1], w: rects[i * 4 + 2], h: rects[i * 4 + 3] })
  // Parents first, so a banner strip keeps its branch colour and children paint over the rest.
  for (let i = 0; i < tree.count; i++) {
    const r = rect(i)
    if (!(r.w > 0 && r.h > 0)) continue
    const { fill, color } = nodeColor(drawing, i)
    if (kit.fillWith(ctx, fill, color, r)) ctx.fillRect(r.x, r.y, r.w, r.h)
    if (series.lineStyle) kit.strokeBorder(ctx, series.lineStyle, 'FFFFFF', r, r)
    else { ctx.strokeStyle = '#FFFFFF'; ctx.lineWidth = 1; ctx.strokeRect(r.x, r.y, r.w, r.h) }
  }
  const px = hierarchyLabelPx(model, series), font = kit.font(px)
  ctx.font = font
  const centred = (series.labels?.position ?? 'ctr') === 'ctr'
  const draw = (i: number, text: string, x: number, y: number, align: CanvasTextAlign, baseline: CanvasTextBaseline) => {
    ctx.fillStyle = textColor(series, nodeColor(drawing, i).color)
    ctx.textAlign = align; ctx.textBaseline = baseline
    ctx.fillText(text, x, y)
  }
  for (let i = 0; i < tree.count; i++) {
    const r = rect(i), leaf = isLeaf(tree, i)
    const text = leaf ? nodeText(series, tree, i) : layout === 'none' ? '' : tree.labels[i]
    if (!text) continue
    const width = kit.width(text, font)
    // A label that does not fit its rectangle (its banner, for a banner parent) is not drawn.
    if (width + 6 > r.w || px + 2 > (leaf || layout !== 'banner' ? r.h : banner)) continue
    if (leaf && centred) draw(i, text, r.x + r.w / 2, r.y + r.h / 2, 'center', 'middle')
    else if (!leaf && layout === 'banner') draw(i, text, r.x + 3, r.y + Math.min(banner, r.h) / 2, 'left', 'middle')
    else draw(i, text, r.x + 3, r.y + 3, 'left', 'top')
  }
}
