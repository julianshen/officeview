/**
 * Bubble charts (spec 4.1; lazy chart chunk). `bubble3D` is drawn flat.
 *
 * - x and y sit on two independent value axes (the `xy` binding, as scatter).
 * - Size: the largest bubble of the group has a diameter of a quarter of the
 *   plot's smaller side at `c:bubbleScale` 100% (0-300%, default 100).
 *   `c:sizeRepresents` area (default: radius ∝ √size) or w (radius ∝ size).
 * - A zero size draws no bubble. A negative size draws nothing unless
 *   `c:showNegBubbles` is 1; then a bubble of |size| with no fill, stroked in
 *   the series line colour, as Office draws negative bubbles.
 * - A size that is not a value (blank, error, unavailable) draws no bubble.
 *   A series without bubble-size data is unavailable (a required role, spec 6.4).
 *
 * Painting allocates nothing per point (two passes over the slots) and is
 * admitted against the painter's path work bound.
 */
import { attrs, getChildren, type XmlNode } from '../../core/xml'
import type { ChartFillStyle } from '../fills'
import type { ChartGroup, ChartPoints, ChartSeries } from '../model'
import { hiddenAt, valueAt } from '../points'
import { applyLineStyle, type ChartLineStyle } from '../style'
import { axisFraction, type ResolvedAxis } from '../axes'

const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]
const val = (node: XmlNode | undefined) => attrs(node).val

/** The bubble group options of a `c:bubbleChart` (Office defaults: 100%, area, negative bubbles hidden). */
export function readBubbleOptions(node: XmlNode): Pick<ChartGroup, 'bubbleScale' | 'sizeRepresents' | 'showNegBubbles'> {
  // ST_BubbleScale: an unsigned integer or a percentage ("150%"), 0-300.
  const raw = val(child(node, 'bubbleScale'))?.trim().replace(/%$/, '')
  const scale = raw !== undefined && raw !== '' && Number.isFinite(Number(raw)) ? Math.max(0, Math.min(300, Number(raw))) : 100
  const neg = child(node, 'showNegBubbles'), negVal = val(neg)
  return {
    bubbleScale: scale,
    sizeRepresents: val(child(node, 'sizeRepresents')) === 'w' ? 'w' : 'area',
    // The element's val defaults to true when present without one; an absent element is Office's 0.
    showNegBubbles: neg ? negVal === undefined || negVal === '1' || negVal === 'true' : false,
  }
}

interface Box { x: number; y: number; w: number; h: number }
/** What the bubble painter reads of a drawn series (the painter's display series). */
export interface BubbleSeries {
  source: ChartSeries
  color: string
  pointColors?: string[]
  hasX: boolean
  plot(slot: number, x?: boolean): number | 'gap' | 'span'
}
/** The painter's shared helpers, passed in (the family never imports the painter). */
export interface BubblePaintKit {
  plot: Box
  slots: number
  source(slot: number): number
  /** plotVisOnly: a size whose own cell is hidden is not drawn. */
  visibleOnly: boolean
  admitPath(cost: number): boolean
  fillWith(ctx: CanvasRenderingContext2D, fill: ChartFillStyle | undefined, color: string, box: Box): boolean
  strokeBorder(ctx: CanvasRenderingContext2D, line: ChartLineStyle | undefined, color: string, box: Box): void
}

/** The bubble size at a source index, or undefined (not a value, or hidden). */
function sizeAt(sizes: ChartPoints | undefined, i: number, visibleOnly: boolean): number | undefined {
  if (!sizes || i >= sizes.length || (visibleOnly && hiddenAt(sizes.hidden, i))) return undefined
  return valueAt(sizes, i)
}

/** The radius of a bubble of `size` when the largest drawn size is `largest` and its radius `big`. */
export function bubbleRadius(size: number, largest: number, big: number, represents: 'area' | 'w'): number {
  const ratio = Math.abs(size) / largest
  return represents === 'w' ? big * ratio : big * Math.sqrt(ratio)
}

/** Paints one bubble group on its x (`X`) and y (`V`) value axes. */
export function paintBubbles(ctx: CanvasRenderingContext2D, group: Pick<ChartGroup, 'bubbleScale' | 'sizeRepresents' | 'showNegBubbles'>, series: readonly BubbleSeries[], X: ResolvedAxis, V: ResolvedAxis, kit: BubblePaintKit): void {
  const { plot } = kit
  const showNeg = group.showNegBubbles === true
  const position = (s: BubbleSeries, slot: number): { x: number; y: number } | undefined => {
    const y = s.plot(slot), x = s.hasX ? s.plot(slot, true) : slot + 1
    if (typeof y !== 'number' || typeof x !== 'number' || !V.plottable(y) || !X.plottable(x)) return undefined
    const f = axisFraction(V, y, false), g = axisFraction(X, x, false)
    return { x: plot.x + g * plot.w, y: plot.y + plot.h - f * plot.h }
  }
  // The largest drawn size of the group sets the scale (negative sizes count only when they are drawn).
  let largest = 0
  for (const s of series) for (let slot = 0; slot < kit.slots; slot++) {
    const size = sizeAt(s.source.bubbleSizes, kit.source(slot), kit.visibleOnly)
    if (size === undefined || !(size > 0 || (showNeg && size < 0)) || !position(s, slot)) continue
    largest = Math.max(largest, Math.abs(size))
  }
  if (!(largest > 0) || !Number.isFinite(largest)) return
  const big = Math.min(plot.w, plot.h) / 8 * (group.bubbleScale ?? 100) / 100
  if (!(big > 0)) return
  for (const s of series) {
    if (!kit.admitPath(kit.slots * 4 + 1)) continue
    const src = s.source
    for (let slot = 0; slot < kit.slots; slot++) {
      const i = kit.source(slot), size = sizeAt(src.bubbleSizes, i, kit.visibleOnly)
      if (size === undefined || size === 0 || (size < 0 && !showNeg)) continue
      const at = position(s, slot)
      if (!at) continue
      const r = bubbleRadius(size, largest, big, group.sizeRepresents ?? 'area')
      if (!(r > 0) || !Number.isFinite(r)) continue
      const box = { x: at.x - r, y: at.y - r, w: r * 2, h: r * 2 }
      const point = src.dPt?.get(i)
      ctx.beginPath()
      ctx.arc(at.x, at.y, r, 0, Math.PI * 2)
      if (size < 0) {
        // Negative: no fill, the series line colour (an authored a:ln colour, else the series colour).
        ctx.save()
        try {
          ctx.strokeStyle = `#${s.color}`; ctx.lineWidth = 1
          if (applyLineStyle(ctx, point?.line ?? src.lineStyle, box)) ctx.stroke()
        } finally { ctx.restore() }
        continue
      }
      const color = point?.color ? s.pointColors?.[i] ?? s.color : s.color
      if (kit.fillWith(ctx, point?.fill ?? src.fill, color, box)) ctx.fill()
      kit.strokeBorder(ctx, point?.line ?? src.lineStyle, s.color, box)
    }
  }
}
