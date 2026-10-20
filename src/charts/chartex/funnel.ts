/**
 * The ChartEx funnel family (spec 4.3 "Funnel"; docs/charts/chartex-map.md §9). A display family: at most
 * CHART_POINT_LIMIT points are drawn (the model is capped and marked truncated at parse or fill).
 *
 * - One bar per point, in source order from top to bottom [U]; each bar is horizontally centred in the plot and its
 *   width is value / max(value) of the plot width [U]. The band gap is `catScaling@gapWidth` [U].
 * - Zero or negative values draw a zero-width bar; their label is still drawn [U]. Blank, error, unavailable and
 *   hidden points draw no bar [P].
 * - Only the one visible series is drawn; a `hidden` series is not; several visible series are unsupported (the
 *   fallback picture is kept) [U].
 * - The category labels sit left of the bars (the category axis's tick labels), measured in the layout.
 *
 * Bars are streamed: nothing is allocated per point.
 */
import type { ChartExGroup, ChartExSeriesInfo, ChartModel, ChartPoint } from '../model'
import type { DisplaySeries, DisplayView } from '../display'
import type { ChartFillStyle } from '../fills'
import type { ChartLineStyle } from '../style'

/** Several visible series in one display-family chart (map §9; for waterfall a [P] extension): unsupported. */
export const CHARTEX_SEVERAL_VISIBLE = 'Funnel and waterfall charts with more than one visible series are not supported'
/** The reason a display-family chart (waterfall, funnel) cannot be drawn from its series, if any. */
export function displayFamilyUnsupported(family: string, infos: readonly ChartExSeriesInfo[]): string | undefined {
  if (family !== 'funnel' && family !== 'waterfall') return undefined
  return infos.filter(info => !info.hidden).length > 1 ? CHARTEX_SEVERAL_VISIBLE : undefined
}

interface Box { x: number; y: number; w: number; h: number }
/** The key under which the layout keeps the funnel's measured category labels. */
export const FUNNEL_TEXTS = 'chartex-funnel'

/** The one series a funnel draws: its first visible (not `hidden`) series in the display view. */
export function funnelSeries(group: ChartExGroup, view: DisplayView, g = 0): DisplaySeries | undefined {
  for (const series of view.groupSeries(g)) {
    const info = group.chartEx.series[group.series.indexOf(series.series)]
    if (info && !info.hidden) return series
  }
  return undefined
}

/**
 * The bars of a funnel over `slots` display slots inside `plot`: band i is the i-th slice from the top, the bar is
 * `band / (1 + gapWidth/100)` thick, centred in its band, and `plot.w · max(v, 0) / max` wide, centred horizontally.
 * `bar(slot)` is undefined when the point has no value. The maximum is found once by a stream (no storage).
 */
export function funnelGeometry(slots: number, point: (slot: number) => ChartPoint | undefined, plot: Box, gapWidth: number): { max: number; band: number; bar(slot: number): Box | undefined } {
  let max = 0
  for (let slot = 0; slot < slots; slot++) { const p = point(slot); if (p?.state === 'value' && p.v > max) max = p.v }
  const band = plot.h / Math.max(1, slots), thick = band / (1 + Math.max(0, gapWidth) / 100), cx = plot.x + plot.w / 2
  return {
    max, band,
    bar: slot => {
      const p = slot >= 0 && slot < slots ? point(slot) : undefined
      if (p?.state !== 'value') return undefined
      const w = max > 0 ? plot.w * Math.max(0, p.v) / max : 0
      return { x: cx - w / 2, y: plot.y + band * slot + (band - thick) / 2, w, h: thick }
    },
  }
}

/** The funnel's category label strings (the drawn slots; ordinals when the series has none), or none. */
export function funnelTexts(model: ChartModel, view: DisplayView): string[] | undefined {
  const group = model.groups[0]
  if (group?.family !== 'funnel') return undefined
  const axis = model.axes.find(a => a.kind === 'category')
  if (!axis || axis.deleted || axis.tickLblPos === 'none') return []
  return Array.from({ length: view.slots }, (_, slot) => view.category(slot) ?? '')
}

/** What the painter hands the funnel: the plot, its measured labels and the shared paint helpers. */
export interface FunnelPaint {
  group: ChartExGroup
  view: DisplayView
  plot: Box
  /** The category labels (`FUNNEL_TEXTS`) and the label font, drawn right-aligned left of the plot. */
  labels: string[]
  font: string
  textColor: string
  admitPath(cost: number): boolean
  fillWith(ctx: CanvasRenderingContext2D, fill: ChartFillStyle | undefined, color: string, box: Box): boolean
  strokeBorder(ctx: CanvasRenderingContext2D, line: ChartLineStyle | undefined, color: string, box: Box, rect: Box): void
}
export function paintFunnel(ctx: CanvasRenderingContext2D, p: FunnelPaint): void {
  const drawn = funnelSeries(p.group, p.view)
  const slots = p.view.slots
  if (!p.admitPath(slots * 4)) return
  if (drawn) {
    const src = drawn.series, geometry = funnelGeometry(slots, slot => drawn.point(slot), p.plot, p.group.gapWidth)
    for (let slot = 0; slot < slots; slot++) {
      const bar = geometry.bar(slot)
      if (!bar || bar.w <= 0) continue
      const i = p.view.source(slot), point = src.dPt?.get(i)
      const color = point?.color ?? src.pointColors?.[i] ?? src.color
      if (p.fillWith(ctx, point?.fill ?? (point?.color ? undefined : src.fill), color, bar)) ctx.fillRect(bar.x, bar.y, bar.w, bar.h)
      p.strokeBorder(ctx, point?.line ?? src.lineStyle, color, bar, bar)
    }
  }
  if (!p.labels.length) return
  ctx.save()
  try {
    ctx.font = p.font; ctx.fillStyle = p.textColor; ctx.textAlign = 'right'; ctx.textBaseline = 'middle'
    const band = p.plot.h / Math.max(1, slots)
    for (let slot = 0; slot < Math.min(slots, p.labels.length); slot++) if (p.labels[slot]) ctx.fillText(p.labels[slot], p.plot.x - 6, p.plot.y + band * (slot + .5))
  } finally { ctx.restore() }
}

/**
 * A data label's anchor on a funnel bar (its text is centred there): `ctr` (the default) the bar centre; `inEnd`,
 * `inBase` inside the right and left ends; `outEnd` past the right end; `t`, `b`, `l`, `r` offset from the centre [U].
 */
export function funnelLabelAnchor(geometry: ReturnType<typeof funnelGeometry>, plot: Box, slot: number, position: string, textWidth: number): { x: number; y: number } | undefined {
  const bar = geometry.bar(slot)
  if (!bar) return undefined
  const cx = bar.x + bar.w / 2, cy = plot.y + geometry.band * (slot + .5), right = bar.x + bar.w, left = bar.x
  switch (position) {
    case 'inEnd': return { x: right - 6 - textWidth / 2, y: cy }
    case 'inBase': return { x: left + 6 + textWidth / 2, y: cy }
    case 'outEnd': return { x: right + 6 + textWidth / 2, y: cy }
    case 't': return { x: cx, y: cy - 8 }
    case 'b': return { x: cx, y: cy + 8 }
    case 'l': return { x: cx - 8, y: cy }
    case 'r': return { x: cx + 8, y: cy }
    default: return { x: cx, y: cy }
  }
}
