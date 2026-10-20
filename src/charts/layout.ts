/** Placement-only chart layout. The weak sidecar is never serialized with the model.
 * Placements share a pressure-evicted LRU per receiving workspace.
 * A conservative construction bound is reserved before measurement or graph construction.
 */
import { chartFont, labelFont } from './textStyle'
export { chartFont } from './textStyle'
import type { ChartModel, ChartManualLayout, ChartDiagnostic, ChartTextStyle } from './model'
import type { TextMeasure } from './format-api'
import type { ChartDocumentContext } from './context'
import { axisGraphBound, axisFraction, categoryFraction, type AxisGraph, type ResolvedAxis } from './axes'
import { displayView, type DisplayView } from './display'
import { chartModelOwner, ChartRefusal, onChartModelRelease } from './own'
import { createChartWorkspace, type ChartWorkspace } from './workspace'
import { hydrateAxes, seriesLabels } from './prepare'
import { CONTOUR_LEGEND_ENTRY_BYTES, CONTOUR_LEGEND_NAME_LIMIT, CONTOUR_LEGEND_TRUNCATED, contourEdges, contourLegendEntries, contourPaintBytes } from './families/contour'
import { stringBytes } from './points'
import { formatNumber } from '../xlsx/chartDescription'
import { CHART_TICK_LIMIT, dateTickLabel, dateUnitIndex } from './dates'
import { radarTexts } from './families/radar'
import { trendlineLegendEntries } from './trendlines'
import { WATERFALL_LEGEND, waterfallColor, waterfallPaintBytes, WATERFALL_KINDS } from './chartex/waterfall'
import { FUNNEL_TEXTS, funnelSeries, funnelTexts } from './chartex/funnel'
import type { ChartExGroup } from './model'

export interface Box { x: number; y: number; w: number; h: number }
export interface AxisLabel { index: number; text: string; position: number; extent: number; crossExtent: number; angle: number }
export interface ChartLayout {
  graph: AxisGraph; view: DisplayView; area: Box; plot: Box; title: Box; legend: Box
  readonly retainedBytes: number
  axisBands: Map<string, number>
  texts: Map<string, string[]>
  labelIndices: Map<string, Set<number>>
  fontPx: number; labels: Map<string, AxisLabel[]>; diagnostics: ChartDiagnostic[]
  /** The legend entries, built once (bounded and admitted with the layout) and reused by the paint. */
  legendEntries: Array<{ name: string; color?: string; line?: true }>
  width(text: string, font: string): number
}
export function textAngle(text: ChartTextStyle | undefined): number {
  return (text?.vertical === 'vert270' ? -90 : text?.vertical === 'vert' || text?.vertical === 'eaVert' ? 90 : text?.rotation ?? 0) * Math.PI / 180
}
export function resolveManualLayout(manual: ChartManualLayout | undefined, auto: Box, w: number, h: number, diagnostics: ChartDiagnostic[]): Box {
  if (!manual) return auto
  const x = manual.x === undefined ? auto.x : manual.x * w + (manual.xMode === 'edge' ? 0 : auto.x)
  const y = manual.y === undefined ? auto.y : manual.y * h + (manual.yMode === 'edge' ? 0 : auto.y)
  const width = manual.w === undefined ? auto.w : manual.w * w - (manual.wMode === 'edge' ? x : 0)
  const height = manual.h === undefined ? auto.h : manual.h * h - (manual.hMode === 'edge' ? y : 0)
  const left = Math.max(0, x), top = Math.max(0, y), right = Math.min(w, x + width), bottom = Math.min(h, y + height)
  if (![left, top, right, bottom].every(Number.isFinite) || right <= left || bottom <= top) {
    diagnostics.push({ code: 'chart-layout-invalid', reason: 'Manual layout is empty or inverted; automatic layout used' }); return auto
  }
  return { x: left, y: top, w: right - left, h: bottom - top }
}
export function labelPlace(a: ResolvedAxis, graph: AxisGraph): 'near' | 'far' | 'axis' {
  const reverse = a.crosses !== undefined && graph.axes.get(a.crosses)?.reversed === true
  const pos = a.axis.tickLblPos
  if (pos === undefined) return (a.axis.crosses === 'max') !== reverse ? 'far' : 'near'
  if (pos === 'low') return reverse ? 'far' : 'near'
  if (pos === 'high') return reverse ? 'near' : 'far'
  return 'axis'
}
export function axisSide(a: ResolvedAxis, graph: AxisGraph): 'l' | 'r' | 'b' | 't' {
  const place = labelPlace(a, graph)
  return place === 'axis' ? a.side : a.orientation === 'v' ? (place === 'near' ? 'l' : 'r') : (place === 'near' ? 'b' : 't')
}
export function axisTexts(a: ResolvedAxis, graph: AxisGraph, view: DisplayView): string[] {
  if (a.role === 'value') return a.ticks.map(t => formatNumber(a.labelValue(t), a.percent))
  if (a.role === 'date') return a.ticks.map(t => dateTickLabel(t, a.dateBase ?? 'days', graph.epoch))
  if (a.role === 'series') return seriesLabels(graph, view, a)
  return Array.from({ length: Math.min(a.slots, view.slots) }, (_, i) => view.category(i) ?? '')
}
/**
 * Legend entries; an entry without a colour has no swatch (a noFill contour band). A contour lists at most `limit`
 * bands (the rows its legend can show), each name cut to the label limit (`truncated` notes a cut). Drawn trendlines
 * follow the series (A6a; `line`: a line key).
 */
export function legendEntries(model: ChartModel, view: DisplayView, graph?: AxisGraph, limit = Infinity, truncated?: () => void): Array<{ name: string; color?: string; line?: true }> {
  const group = model.groups[0]
  if (group?.family === 'contour') {
    // Contour: the bands (spec 4.1), named by the shared value axis's tick labels.
    const value = graph?.groups[0]?.value
    if (!value) return []
    return contourLegendEntries(group, contourEdges(value), model.preparedAxes?.axes.find(p => p.id === value.id)?.labels ?? axisTexts(value, graph, view), edge => formatNumber(value.labelValue(edge), value.percent), limit, truncated)
  }
  // Waterfall (B2): the increase, decrease and total kinds in their colours (map §6 [U]).
  if (group?.kind === 'chartEx' && group.family === 'waterfall') return WATERFALL_LEGEND.map((name, k) => ({ name, color: waterfallColor(group as ChartExGroup, WATERFALL_KINDS[k]) }))
  // Funnel (B2): the drawn (visible) series only.
  if (group?.kind === 'chartEx' && group.family === 'funnel') { const s = funnelSeries(group as ChartExGroup, view)?.series; return s ? [{ name: s.name, color: s.color }] : [] }
  if (group?.family === 'pie' || group?.family === 'doughnut') {
    const s = view.groupSeries(0)[0]?.series
    return s ? Array.from({ length: Math.min(view.slots, 64) }, (_, i) => ({ name: view.category(i) ?? String(view.source(i) + 1), color: s.pointColors?.[view.source(i)] ?? s.color })) : []
  }
  // Drawn trendlines follow the series (A6a, spec 5.6/5.10).
  return [...view.series.map(({ series }) => ({ name: series.name, color: series.color })), ...trendlineLegendEntries(model)]
}
interface Entry {
  contexts: Map<ChartWorkspace, Map<string, Entry>>
  entries: Map<string, Entry>
  pins: number
  layout: ChartLayout
  workspace: ChartWorkspace
  tag: string
  key: string
  ref: WeakRef<Entry>
}
const defaults = new WeakMap<ChartModel, ChartWorkspace>()
const cache = new WeakMap<ChartModel, Map<ChartWorkspace, Map<string, Entry>>>()
// Weak links preserve collection of a model while maintaining insertion-order LRU.
const byWorkspace = new WeakMap<ChartWorkspace, Set<WeakRef<Entry>>>()
const fontEntries = new Set<WeakRef<Entry>>()
const collected = new FinalizationRegistry<{ workspace: ChartWorkspace; tag: string; ref: WeakRef<Entry> }>(e => {
  e.workspace.release(e.tag)
  byWorkspace.get(e.workspace)?.delete(e.ref)
  fontEntries.delete(e.ref)
})
const layoutEntries = new WeakMap<ChartLayout, Entry>()
let serial = 0

export function invalidateChartFonts(): void {
  for (const ref of fontEntries) { const entry = ref.deref(); if (entry) drop(entry) }
  fontEntries.clear()
}
function touch(entry: Entry): void {
  const lru = byWorkspace.get(entry.workspace)!
  lru.delete(entry.ref)
  lru.add(entry.ref)
}
function drop(entry: Entry): void {
  if (entry.entries.get(entry.key) === entry) entry.entries.delete(entry.key)
  if (!entry.entries.size) entry.contexts.delete(entry.workspace)
  entry.key = ''
  byWorkspace.get(entry.workspace)?.delete(entry.ref)
  fontEntries.delete(entry.ref)
  if (!entry.pins) {
    collected.unregister(entry)
    entry.workspace.release(entry.tag)
  }
}
/** Evict least-recently-used, unpinned cache entries before refusing construction or paint scratch. */
export function reserveChartLayoutWorkspace(workspace: ChartWorkspace, bytes: number, tag: string): boolean {
  if (workspace.reserve(bytes, tag)) return true
  for (const ref of byWorkspace.get(workspace) ?? []) {
    const entry = ref.deref()
    if (!entry) { byWorkspace.get(workspace)?.delete(ref); continue }
    if (entry.pins) continue
    drop(entry)
    if (workspace.reserve(bytes, tag)) return true
  }
  return false
}
export function holdChartLayout(layout: ChartLayout): () => void {
  const entry = layoutEntries.get(layout)!
  entry.pins++
  let released = false
  return () => {
    if (released) return
    released = true
    entry.pins--
    if (!entry.key) drop(entry)
  }
}
export function releaseChartLayouts(model: ChartModel): void {
  for (const placements of cache.get(model)?.values() ?? []) {
    for (const entry of placements.values()) drop(entry)
  }
  cache.delete(model)
  defaults.delete(model)
}
export function releaseWorkspaceLayouts(workspace: ChartWorkspace): void {
  for (const ref of byWorkspace.get(workspace) ?? []) { const entry = ref.deref(); if (entry) drop(entry) }
  byWorkspace.delete(workspace)
}
/** Scoped parse ledgers expose the receiving facet; paint and layout share this rule. */
export function ownedChartWorkspace(model: ChartModel): ChartWorkspace | undefined {
  const owner = chartModelOwner(model)?.workspace
  return owner && ((owner as { facet?: ChartWorkspace }).facet ?? owner)
}
/** The legend rows a w x h chart can show: vertical legends one per text line, horizontal ones one per minimal entry width. */
export function legendCapacity(model: ChartModel, w: number, h: number): number {
  const fontPx = model.fontSizePt * 96 / 72
  const vertical = model.legend === 'l' || model.legend === 'r' || model.legend === 'tr'
  return Math.max(1, Math.floor(vertical ? h / (fontPx + 6) : w / (fontPx * .7 + 12)) + 1)
}
/** Temporary construction bound, including labels/maps, legend widths and display closures. */
export function layoutBound(model: ChartModel, w = Infinity, h = Infinity): number {
  let slots = model.categories.length, series = 0, text = 0
  for (const group of model.groups) for (const s of group.series) { series++; slots = Math.max(slots, s.values.length); text += s.name.length * 2 }
  slots = Math.min(4096, Math.max(1, slots))
  for (const c of model.categories) text += c.length * 2
  // Contour legend entries: one per band the legend can show (at most the tick bound), each with its bounded name,
  // a transient endpoint cut and its measured width, plus a possible truncation diagnostic.
  if (model.groups[0]?.family === 'contour') text += Math.min(legendBands(model), legendCapacity(model, w, h)) * CONTOUR_LEGEND_ENTRY_BYTES + 512
  // Funnel category labels (B2): each drawn slot's label is retained in the texts map (record + UTF-16) and its
  // measured width (entry + UTF-16), priced from the strings the layout will select, before any is built.
  if (model.groups[0]?.family === 'funnel') { const view = displayView(model); for (let slot = 0; slot < view.slots; slot++) text += 8 + 112 + 4 * (view.category(slot)?.length ?? 0) }
  // Waterfall legend entries (B2): three fixed kind names and their widths.
  if (model.groups[0]?.family === 'waterfall') text += WATERFALL_LEGEND.reduce((n, name) => n + 64 + 2 * name.length + 128, 0)
  // Trendline legend entries (A6a): a generated or authored name (at most 255 + the series name's units) and its width.
  for (const group of model.groups) for (const s of group.series) text += (s.trendlines?.length ?? 0) * (2 * (264 + s.name.length) + 128)
  // The retained diagnostics (stage 1's and up to three layout ones) are priced as retainedBytes prices them.
  const diagnostics = (model.preparedAxes?.diagnostics ?? []).reduce((n, d) => n + 64 + d.code.length * 2 + d.reason.length * 2, 1024)
  return (model.preparedAxes?.dataLabels??[]).reduce((n,l)=>n+1024+l.text.length*4+l.runs.length*256,0) + axisGraphBound(model, slots) + 4096 + series * 512 + model.axes.length * (slots * 160 + 1000 * 128) + text + diagnostics
}
export function layoutChart(model: ChartModel, w: number, h: number, dpr: number, fontGeneration: number, measure: TextMeasure, context?: ChartDocumentContext, family = JSON.stringify(model.fontFamily)): ChartLayout {
  if (![w, h, dpr].every(v => Number.isFinite(v) && v > 0)) throw new RangeError('Invalid chart placement')
  let workspace = context?.workspace ?? ownedChartWorkspace(model) ?? defaults.get(model)
  if (!workspace) { workspace = createChartWorkspace('mobile'); defaults.set(model, workspace) }
  let contexts = cache.get(model)
  if (!contexts) { contexts = new Map(); cache.set(model, contexts) }
  let entries = contexts.get(workspace)
  if (!entries) { entries = new Map(); contexts.set(workspace, entries) }
  const key = JSON.stringify([w, h, dpr, fontGeneration, family])
  const hit = entries.get(key)
  if (hit) { touch(hit); return hit.layout }
  const tag = `chart-layout#${++serial}`, bound = layoutBound(model, w, h)
  if (!reserveChartLayoutWorkspace(workspace, bound, tag)) throw new ChartRefusal()
  try {
    const layout = buildLayout(model, w, h, measure, family)
    if (layout.retainedBytes > bound) throw new ChartRefusal()
    const retainedTag = `${tag}:retained`
    workspace.transfer(tag, retainedTag, layout.retainedBytes)
    const entry = { contexts, entries, pins: 0, layout, workspace, tag: retainedTag, key } as Entry
    entry.ref = new WeakRef(entry)
    layoutEntries.set(layout, entry)
    contexts.set(workspace, entries)
    entries.set(key, entry)
    let lru = byWorkspace.get(workspace)
    if (!lru) { lru = new Set(); byWorkspace.set(workspace, lru) }
    lru.add(entry.ref)
    fontEntries.add(entry.ref)
    collected.register(entry, { workspace, tag: retainedTag, ref: entry.ref }, entry)
    return layout
  } finally { workspace.release(tag) }
}

/** Pure lookup factory: its lexical scope contains no measuring callback or canvas. */
function cachedWidths(widths: Map<string, Map<string, number>>): ChartLayout['width'] {
  return (text, font) => {
    const width = widths.get(font)?.get(text)
    if (width === undefined) throw new Error('Chart text was not measured during layout')
    return width
  }
}
/** Tight retained bound: objects/maps/closures plus the actual ticks, label records and width entries.
 * Borrowed model storage is charged by the model owner; shared strings/ticks are conservatively priced again.
 */
function retainedBytes(layout: Omit<ChartLayout, 'width' | 'retainedBytes'>, widths: Map<string, Map<string, number>>): number {
  let bytes = 4096 + layout.view.series.length * 512 + layout.graph.groups.length * 256
  for (const axis of layout.graph.axes.values()) bytes += 1024 + (axis.ticks.length + axis.minorTicks.length) * 16
  for (const [id, strings] of layout.texts) bytes += 96 + id.length * 2 + strings.reduce((n, s) => n + 48 + s.length * 2, 0)
  for (const labels of layout.labels.values()) bytes += 128 + labels.length * 192
  // Legend entries: a contour's names and a trendline's generated name (A6a) are their own strings; other families'
  // reuse the model's (records only).
  // The waterfall's three kind names are not model strings either.
  const ownNames = layout.graph.groups[0]?.group.family === 'contour' || layout.graph.groups[0]?.group.family === 'waterfall'
  for (const entry of layout.legendEntries) bytes += 64 + (ownNames || entry.line ? stringBytes(entry.name) : 0)
  for (const [font, strings] of widths) {
    bytes += 96 + font.length * 2
    for (const text of strings.keys()) bytes += 64 + text.length * 2
  }
  for (const diagnostic of layout.diagnostics) bytes += 64 + diagnostic.code.length * 2 + diagnostic.reason.length * 2
  return bytes
}
/** Automatic thinning uses one interval; authored intervals retain a final overlap guard. */
function thinLabels(list: AxisLabel[], authored: number | undefined, gap: number): AxisLabel[] {
  const fits = (labels: AxisLabel[]) => labels.every((l, i) => !i || l.position - l.extent / 2 >= labels[i-1].position + labels[i-1].extent / 2 + gap)
  if (authored !== undefined) {
    let last = -Infinity
    return list.filter(label => {
      if (label.index % authored || label.position - label.extent / 2 < last + gap) return false
      last = label.position + label.extent / 2
      return true
    })
  }
  const byIndex = new Map(list.map(l => [l.index, l]))
  const max = Math.max(0, ...byIndex.keys())
  for (let stride = 1; stride <= max + 1; stride++) {
    const kept: AxisLabel[] = []
    for (let i = 0; i <= max; i += stride) { const label = byIndex.get(i); if (label) kept.push(label) }
    if (kept.length > 1 && kept[0].position > kept[kept.length-1].position) kept.reverse()
    if (fits(kept)) return kept
  }
  return []
}
function buildLayout(model: ChartModel, w: number, h: number, measure: TextMeasure, family: string): ChartLayout {
    const widths = new Map<string, Map<string, number>>()
    const width = (text: string, font: string) => {
      let byText = widths.get(font); if (!byText) widths.set(font, byText = new Map())
      let value = byText.get(text)
      if (value === undefined) { value = measure.width(text, font); byText.set(text, value) }
      return value
    }
    const view = displayView(model), graph = hydrateAxes(model, view)
    const diagnostics = [...graph.diagnostics], fontPx = model.fontSizePt * 96 / 72
    const pad = Math.min(10, w * .04, h * .04)
    let area: Box = { x: pad, y: pad, w: w - pad * 2, h: h - pad * 2 }
    const title = resolveManualLayout(model.titleLayout, { ...area, h: Math.min(96, fontPx * 1.4) + 6 }, w, h, diagnostics)
    if (model.title) area = { ...area, y: area.y + Math.min(96, fontPx * 1.4) + 6, h: area.h - Math.min(96, fontPx * 1.4) - 6 }
    let legendAuto = { ...area }
    // Built once, bounded to the rows the legend can show (admitted by layoutBound before this build), and retained.
    const entriesForLegend = model.legend ? legendEntries(model, view, graph, legendCapacity(model, w, h), () => {
      if (!diagnostics.some(d => d.code === CONTOUR_LEGEND_TRUNCATED)) diagnostics.push({ code: CONTOUR_LEGEND_TRUNCATED, reason: `Legend names longer than ${CONTOUR_LEGEND_NAME_LIMIT} characters are cut` })
    }) : [], line = fontPx + 6
    if (model.legend && entriesForLegend.length) {
      const sizes = entriesForLegend.map(e => width(e.name, chartFont(fontPx, family)) + fontPx * .7 + 12)
      if (model.legend === 'l' || model.legend === 'r' || model.legend === 'tr') {
        const size = Math.min(area.w * .4, Math.max(...sizes) + 6)
        const total = Math.min(entriesForLegend.length, Math.max(1, Math.floor(area.h / line))) * line
        legendAuto = { x: model.legend === 'l' ? area.x : area.x + area.w - size, y: model.legend === 'tr' ? area.y : area.y + (area.h - total) / 2, w: size, h: total }
        area = { ...area, x: area.x + (model.legend === 'l' ? size : 0), w: area.w - size }
      } else {
        const total = sizes.reduce((a, b) => a + b, 0)
        legendAuto = { x: area.x + Math.max(0, (area.w - total) / 2), y: model.legend === 't' ? area.y : area.y + area.h - line, w: Math.min(area.w, total), h: line }
        area = { ...area, y: area.y + (model.legend === 't' ? line : 0), h: area.h - line }
      }
    }
    const legend = resolveManualLayout(model.legendLayout, legendAuto, w, h, diagnostics)
    const placed = graph.groups.filter(g => g !== undefined), used = [...graph.axes.values()].filter(a => placed.some(g => g.horizontal === a || g.vertical === a))
    const valueAxes = used.filter(a => placed.some(g => g.value === a)), baseAxes = used.filter(a => !valueAxes.includes(a))
    const bands = { l: [] as number[], r: [] as number[], t: [] as number[], b: [] as number[] }
    const texts = new Map<string, string[]>(), axisBands = new Map<string, number>()
    for (const a of [...valueAxes, ...baseAxes].sort((p, q) => (p.orientation === 'v' ? 0 : 1) - (q.orientation === 'v' ? 0 : 1))) {
      const strings = model.preparedAxes?.axes.find(p => p.id === a.id)?.labels ?? axisTexts(a, graph, view); texts.set(a.id, strings)
      const px = (a.axis.text?.sizePt ?? model.fontSizePt) * 96 / 72, angle = textAngle(a.axis.text ?? model.text)
      let widest = 0
      for (const s of strings) widest = Math.max(widest, width(s, chartFont(px, family)))
      const cross = a.orientation === 'h' ? Math.abs(Math.sin(angle)) * widest + Math.abs(Math.cos(angle)) * px : Math.abs(Math.cos(angle)) * widest + Math.abs(Math.sin(angle)) * px
      const band = a.deleted || a.axis.tickLblPos === 'none' ? 0 : a.orientation === 'h' ? cross + 6 : a.role === 'value' ? cross + 6 + (a.axis.majorTickMark === 'out' || a.axis.majorTickMark === 'cross' ? 4 : 0) : Math.min(area.w * .35, cross + 8)
      axisBands.set(a.id, band)
      const side = axisSide(a, graph); bands[side].push(band)
      if (a.display?.label && !a.deleted) bands[side].push(fontPx + 6)
      if (a.axis.title) bands[side].push(fontPx + 6)
    }
    // Radar (spec 4.1): its radial tick labels and category labels are measured here; they sit on no planar axis.
    for (const [id, strings] of radarTexts(model, graph, view, a => model.preparedAxes?.axes.find(p => p.id === a.id)?.labels ?? axisTexts(a, graph, view))) {
      texts.set(id, strings)
      const a = graph.axes.get(id), px = (a?.axis.text?.sizePt ?? model.fontSizePt) * 96 / 72
      for (const s of strings) width(s, chartFont(px, family))
    }
    // Funnel (B2): its category labels sit left of the bars (it has no placed axis): measured here, in their band.
    const funnelLabels = funnelTexts(model, view)
    if (funnelLabels?.length) {
      texts.set(FUNNEL_TEXTS, funnelLabels)
      const axis = model.axes.find(a => a.kind === 'category'), px = (axis?.text?.sizePt ?? model.fontSizePt) * 96 / 72
      let widest = 0
      for (const s of funnelLabels) widest = Math.max(widest, width(s, chartFont(px, family)))
      if (widest > 0) bands.l.push(Math.min(area.w * .35, widest + 12))
    }
    const inset = (box: Box): Box => {
      let x = box.x, width = box.w, height = box.h, y = box.y + 4
      for (const b of bands.l) { x += b; width -= b }; width -= 8
      for (const b of bands.r) width -= b
      for (const b of bands.b) height -= b; height -= 4
      for (const b of bands.t) { y += b; height -= b }
      return { x, y, w: width, h: height }
    }
    const auto = inset(area)
    let plot = model.plotLayout?.target === 'inner' ? resolveManualLayout(model.plotLayout, auto, w, h, diagnostics) : inset(resolveManualLayout(model.plotLayout, area, w, h, diagnostics))
    if (plot.w <= 0 || plot.h <= 0) { plot = auto; if (model.plotLayout) diagnostics.push({ code: 'chart-layout-invalid', reason: 'Manual outer layout leaves no plot; automatic layout used' }) }
    const labels = new Map<string, AxisLabel[]>()
    for (const a of used) {
      const list: AxisLabel[] = [], strings = texts.get(a.id)!, angle = textAngle(a.axis.text ?? model.text)
      const px = (a.axis.text?.sizePt ?? model.fontSizePt) * 96 / 72
      const group = placed.find(g => g.category === a), between = !(group && (group.value.axis.crossBetween ?? a.axis.crossBetween) === 'midCat')
      for (let i = 0; i < strings.length; i++) {
        if (!strings[i]) continue
        const f = a.role === 'value' ? axisFraction(a, a.ticks[i], false) : a.role === 'date' ? categoryFraction(a, dateUnitIndex(a.ticks[i], a.dateBase!, graph.epoch) - a.dateStart!, between) : categoryFraction(a, i, between)
        const position = a.orientation === 'h' ? plot.x + f * plot.w : plot.y + plot.h - f * plot.h
        const measured = width(strings[i], chartFont(px, family))
        const extent = a.orientation === 'h' ? Math.abs(Math.cos(angle)) * measured + Math.abs(Math.sin(angle)) * px : Math.abs(Math.sin(angle)) * measured + Math.abs(Math.cos(angle)) * px
        const crossExtent = a.orientation === 'h' ? Math.abs(Math.sin(angle)) * measured + Math.abs(Math.cos(angle)) * px : Math.abs(Math.cos(angle)) * measured + Math.abs(Math.sin(angle)) * px
        list.push({ index: i, text: strings[i], position, extent, crossExtent, angle })
      }
      list.sort((a, b) => a.position - b.position)
      const kept = thinLabels(list, a.axis.tickLblSkip, a.role === 'value' ? 0 : 4)
      if (kept.length < list.length) diagnostics.push({ code: 'chart-label-skip', reason: `Axis ${a.id} labels thinned to avoid overlap` })
      labels.set(a.id, kept)
    }
    const labelIndices = new Map([...labels].map(([id, values]) => [id, new Set(values.map(v => v.index))]))
    for (const label of model.preparedAxes?.dataLabels??[]) for(const run of label.runs.length?label.runs:[{text:label.text}]) width(run.text,labelFont(run,fontPx,family))
    const data = { axisBands, texts, labelIndices, graph, view, area, plot, title, legend, fontPx, labels, diagnostics, legendEntries: entriesForLegend }
    return { ...data, retainedBytes: retainedBytes(data, widths), width: cachedWidths(widths) }
 }

onChartModelRelease(releaseChartLayouts)

/** The legend bands a contour can list: its prepared value axis's tick intervals, else the tick bound. */
function legendBands(model: ChartModel): number {
  const id = model.groups[0]?.binding.kind === 'surface' ? model.groups[0].binding.value : undefined
  const ticks = model.preparedAxes?.axes.find(a => a.id === id)?.ticks.length
  return ticks !== undefined ? Math.max(1, ticks) : CHART_TICK_LIMIT
}
/**
 * The paint-local buffers, admitted before the paint: per group its stack and series buffers (plus the two drop-line
 * extents, A4a); a contour group its band grid (`contourPaintBytes`, the dimensions the painter uses) and its band
 * edges. The legend entries are the layout's (retained there), never rebuilt by the paint.
 */
export const paintWorkspaceBytes = (model: ChartModel, view: DisplayView): number => 2048 + (model.preparedAxes?.trendlines?.length ?? 0) * 2048 + (model.preparedAxes?.dataLabels?.length ? view.slots*64 : 0) + model.groups.reduce((sum, group) => sum + 512 + view.slots * (group.dropLines ? 64 : 48) + (group.family === 'contour' ? contourPaintBytes(model, group) + 64 + (legendBands(model) + 1) * 8 : 0) + (group.family === 'waterfall' ? waterfallPaintBytes(view.slots) : 0), 0) + view.series.length * 256
