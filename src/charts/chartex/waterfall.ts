/**
 * The ChartEx waterfall family (spec 4.3 "Waterfall"; docs/charts/chartex-map.md §6). A display family: at most
 * CHART_POINT_LIMIT points are drawn (the model is capped and marked truncated at parse or fill).
 *
 * - Running total: acc₀ = 0. A value point i floats from acc to acc + vᵢ, then acc += vᵢ; it is an increase when
 *   vᵢ ≥ 0, else a decrease [D-wf].
 * - A subtotal point (its source index is in `cx:subtotals`) is drawn from 0 to the running total, anchored on the
 *   axis, and the accumulator continues from it. Its own cell value is not used for the height [U: the map's rule;
 *   no native part lists a subtotal].
 * - Blank, error, unavailable and hidden points draw no bar and contribute 0 [P].
 * - Colours: increase, decrease and total take the style palette entries 0, 1, 2 (accents 1-3; resolved at parse
 *   into `chartEx.palette`) [U]; a `dataPt@idx` fill or colour overrides one bar.
 * - Connector lines (`cx:visibility@connectorLines`, absent: on [U]) join the level after bar i to the next bar.
 *
 * Every function here streams over the display slots: nothing is allocated per point except by `waterfallLevels`,
 * whose buffers the paint workspace admits (`waterfallPaintBytes`) before they are built.
 */
import type { ChartExGroup, ChartExSeriesInfo, ChartPoint } from '../model'
import type { DisplaySeries, DisplayView } from '../display'
import type { ResolvedAxis } from '../axes'
import type { ChartFillStyle } from '../fills'
import type { ChartLineStyle } from '../style'
import { barGeometry } from '../barGeometry'

export type WaterfallKind = 'increase' | 'decrease' | 'total'
/** Diagnosed once when a waterfall error point is skipped (map §2.4 [P]). */
export const WATERFALL_ERROR_POINT = 'chart-waterfall-error-point'
export const WATERFALL_KINDS: readonly WaterfallKind[] = ['increase', 'decrease', 'total']
/** The legend entries Office shows for a waterfall (map §6) [U]. */
export const WATERFALL_LEGEND: readonly string[] = ['Increase', 'Decrease', 'Total']
/** The default increase, decrease and total colours (Office theme accents 1-3) when no palette was resolved. */
export const WATERFALL_DEFAULT_COLORS: readonly string[] = ['4472C4', 'ED7D31', 'A5A5A5']

/** Visit one drawn bar: its slot, source index, kind and the value-axis interval [from, to]. */
export type WaterfallVisit = (slot: number, source: number, kind: WaterfallKind, from: number, to: number) => void

/**
 * Streams the bars of one series over `slots` display slots. `source` is the display view's (monotonic) source
 * index of a slot, `point` the slot's point (undefined: hidden), `subtotals` the sorted, distinct subtotal indices.
 * `skipped` is told of every value-less, non-subtotal slot (its state; undefined when hidden).
 */
export function forEachWaterfallBar(slots: number, source: (slot: number) => number, point: (slot: number) => ChartPoint | undefined,
  subtotals: readonly number[], visit: WaterfallVisit, skipped?: (slot: number, state: ChartPoint['state'] | undefined) => void): void {
  let acc = 0, k = 0
  for (let slot = 0; slot < slots; slot++) {
    const i = source(slot)
    while (k < subtotals.length && subtotals[k] < i) k++
    if (k < subtotals.length && subtotals[k] === i) { visit(slot, i, 'total', 0, acc); continue }
    const p = point(slot)
    if (p?.state !== 'value') { skipped?.(slot, p?.state); continue }
    const from = acc
    acc += p.v
    visit(slot, i, p.v >= 0 ? 'increase' : 'decrease', from, acc)
  }
}

/** The subtotal indices of a series (sorted and distinct at parse), or none. */
export const waterfallSubtotals = (info: ChartExSeriesInfo | undefined): readonly number[] => info?.layoutPr.subtotals ?? []

/** The one series a waterfall group draws: its first visible (not `hidden`) series in the display view. */
export function waterfallSeries(group: ChartExGroup, view: DisplayView, g: number): { series: DisplaySeries; info: ChartExSeriesInfo } | undefined {
  for (const series of view.groupSeries(g)) {
    const info = group.chartEx.series[group.series.indexOf(series.series)]
    if (info && !info.hidden) return { series, info }
  }
  return undefined
}

/**
 * The value-axis extent the bars need (0 always included: totals are anchored there and acc₀ = 0), without any
 * per-point storage. `error` is told when an error point is skipped (map §2.4: diagnosed).
 */
export function waterfallExtent(group: ChartExGroup, view: DisplayView, g: number, see: (v: number) => void, error?: () => void): void {
  const drawn = waterfallSeries(group, view, g)
  see(0)
  if (!drawn) return
  forEachWaterfallBar(view.slots, view.source, slot => drawn.series.point(slot), waterfallSubtotals(drawn.info), (_s, _i, _k, from, to) => { see(from); see(to) },
    (_slot, state) => { if (state === 'error') error?.() })
}

/** The per-slot levels a paint or label pass reads (NaN: no bar). Admitted by `waterfallPaintBytes` before use. */
export interface WaterfallLevels { from: Float64Array; to: Float64Array; kind: Uint8Array }
/** Bytes of one `waterfallLevels` build over `slots` (two float buffers, one kind byte, the record). */
export const waterfallLevelBytes = (slots: number): number => 128 + slots * 17
/** The paint workspace a waterfall group needs: the bar pass and the label pass each build its levels once. */
export const waterfallPaintBytes = (slots: number): number => 2 * waterfallLevelBytes(slots)
export function waterfallLevels(group: ChartExGroup, view: DisplayView, g: number): WaterfallLevels {
  const from = new Float64Array(view.slots).fill(NaN), to = new Float64Array(view.slots).fill(NaN), kind = new Uint8Array(view.slots)
  const drawn = waterfallSeries(group, view, g)
  if (drawn) forEachWaterfallBar(view.slots, view.source, slot => drawn.series.point(slot), waterfallSubtotals(drawn.info), (slot, _i, k, a, b) => {
    from[slot] = a; to[slot] = b; kind[slot] = WATERFALL_KINDS.indexOf(k)
  })
  return { from, to, kind }
}
/** The colour of a kind: the parse-resolved palette entry, else the Office default. */
export const waterfallColor = (group: ChartExGroup, kind: WaterfallKind): string => {
  const k = WATERFALL_KINDS.indexOf(kind)
  return group.chartEx.palette?.[k] ?? WATERFALL_DEFAULT_COLORS[k]
}

interface Box { x: number; y: number; w: number; h: number }
/** What the cartesian painter hands the waterfall: its resolved axes and the shared paint helpers. */
export interface WaterfallPaint {
  group: ChartExGroup
  g: number
  view: DisplayView
  levels: WaterfallLevels
  category: ResolvedAxis
  between: boolean
  plot: Box
  /** A value's coordinate along the value axis. */
  at(v: number): number
  categoryIndex(slot: number): number
  admitPath(cost: number): boolean
  fillWith(ctx: CanvasRenderingContext2D, fill: ChartFillStyle | undefined, color: string, box: Box): boolean
  strokeBorder(ctx: CanvasRenderingContext2D, line: ChartLineStyle | undefined, color: string, box: Box, rect: Box): void
}
/** The connector colour and width (Office draws a thin grey line) [U]. */
export const WATERFALL_CONNECTOR = '#7F7F7F'
/** One bar's rectangle along the (standard-grouped, single series) category band rule shared with column bars. */
export function waterfallBar(p: Pick<WaterfallPaint, 'group' | 'category' | 'between' | 'plot' | 'at' | 'categoryIndex'>, slot: number, from: number, to: number): Box | undefined {
  const geometry = barGeometry({ grouping: 'standard', overlap: 0, gapWidth: p.group.gapWidth, series: { length: 1 } }, p.category, p.between, p.plot)
  const idx = p.categoryIndex(slot)
  if (!Number.isFinite(idx) || !Number.isFinite(from) || !Number.isFinite(to)) return undefined
  const start = geometry.start(idx, 0), a = p.at(from), b = p.at(to)
  return p.category.orientation === 'v' ? { x: Math.min(a, b), y: start, w: Math.abs(b - a), h: geometry.width } : { x: start, y: Math.min(a, b), w: geometry.width, h: Math.abs(b - a) }
}
/** Paints the bars (kind colours, `dataPt` overrides) and, when visible, the connector lines. */
export function paintWaterfall(ctx: CanvasRenderingContext2D, p: WaterfallPaint): void {
  const drawn = waterfallSeries(p.group, p.view, p.g)
  if (!drawn) return
  const { from, to, kind } = p.levels, count = from.length, src = drawn.series.series
  if (!p.admitPath(count * 4)) return
  for (let slot = 0; slot < count; slot++) {
    const bar = waterfallBar(p, slot, from[slot], to[slot])
    if (!bar) continue
    const i = p.view.source(slot), point = src.dPt?.get(i)
    const color = point?.color ?? waterfallColor(p.group, WATERFALL_KINDS[kind[slot]])
    if (p.fillWith(ctx, point?.fill, color, bar)) ctx.fillRect(bar.x, bar.y, bar.w, bar.h)
    p.strokeBorder(ctx, point?.line ?? src.lineStyle, color, bar, bar)
  }
  if (!drawn.info.layoutPr.visibility.connectorLines || count < 2 || !p.admitPath((count - 1) * 2)) return
  ctx.save()
  try {
    ctx.strokeStyle = WATERFALL_CONNECTOR; ctx.lineWidth = 0.75 * 96 / 72
    ctx.beginPath()
    for (let slot = 0; slot + 1 < count; slot++) {
      const a = waterfallBar(p, slot, from[slot], to[slot]), b = waterfallBar(p, slot + 1, from[slot + 1], to[slot + 1])
      if (!a || !b) continue
      // The level after bar `slot` (its end; a total's top) to the facing edge of the next bar.
      const level = p.at(to[slot])
      if (p.category.orientation === 'v') { ctx.moveTo(level, b.y > a.y ? a.y + a.h : a.y); ctx.lineTo(level, b.y > a.y ? b.y : b.y + b.h) }
      else { ctx.moveTo(b.x > a.x ? a.x + a.w : a.x, level); ctx.lineTo(b.x > a.x ? b.x : b.x + b.w, level) }
    }
    ctx.stroke()
  } finally { ctx.restore() }
}
