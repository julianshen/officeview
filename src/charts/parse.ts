/**
 * Chart part reader (lazy chunk): reads c:chartSpace into the chart model.
 * Moved from `src/xlsx/chart.ts` (A0a) behaviour-identically; the shared
 * default `parseChart` adapter runs it for XLSX, DOCX and PPTX.
 */
import { attrs, getChildren, orderedChildren, textOf, type XmlNode } from '../core/xml'
import { parseDrawingColor, resolveDrawingColor, type ThemeContext } from '../drawing/style'
import type { ChartParseContext } from './context'
import type { ChartAxis, ChartAxisBinding, ChartFamily, ChartGroup, ChartModel, ChartPoints, ChartRange, ChartSeries } from './model'
import { PointWriter, knownError, pointBytes } from './points'
import { ChartRefusal, adoptChartModel, modelOwner } from './own'

/** Bounds: series per chart, points per series drawn (display, spec 7). */
export const CHART_SERIES_LIMIT = 64
export const CHART_POINT_LIMIT = 4096
/** Points per series acquired for an analytic element (trendline, stdDev/stdErr error bars; spec 7). */
export const CHART_ANALYTIC_POINT_LIMIT = 65536
/** Characters kept per chart label (category, series name); longer labels are cut and the chart marked truncated. */
export const CHART_LABEL_LIMIT = 255
/** Hard cap on one fill's label storage (UTF-16 bytes). */
export const CHART_TEXT_BYTES_LIMIT = 4 * 1024 * 1024

const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]
const val = (node: XmlNode | undefined) => attrs(node).val
const finite = (raw: string | undefined): number | undefined => raw !== undefined && raw.trim() !== '' && Number.isFinite(Number(raw)) ? Number(raw) : undefined
const DEFAULT_ACCENTS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']

const GROUPS: Record<string, ChartFamily> = {
  barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area',
  pieChart: 'pie', pie3DChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter',
}
const UNSUPPORTED: Record<string, string> = {
  radarChart: 'Radar chart', bubbleChart: 'Bubble chart', stockChart: 'Stock chart', surfaceChart: 'Surface chart',
  surface3DChart: 'Surface chart', ofPieChart: 'Pie-of-pie chart',
}

function hex(color: { r: number; g: number; b: number } | undefined): string | undefined {
  if (!color) return undefined
  return [color.r, color.g, color.b].map(c => Math.max(0, Math.min(255, Math.round(c))).toString(16).padStart(2, '0')).join('').toUpperCase()
}
function accent(theme: ThemeContext | undefined, index: number): string {
  if (!Number.isSafeInteger(index) || index < 0) index = 0
  const slot = `accent${(index % 6) + 1}`
  return (theme?.palette[slot]?.replace(/^#/, '') ?? DEFAULT_ACCENTS[index % 6]).toUpperCase()
}
function solid(spPr: XmlNode | undefined, theme: ThemeContext | undefined, line = false): string | undefined | null {
  const holder = line ? child(spPr, 'ln') : spPr
  if (!holder) return undefined
  if (child(holder, 'noFill')) return null
  const fill = child(holder, 'solidFill')
  return fill ? hex(resolveDrawingColor(parseDrawingColor(fill), theme)) : undefined
}

function richText(node: XmlNode | undefined): string {
  const out: string[] = []
  for (const p of getChildren(child(node, 'rich'), 'p')) {
    let line = ''
    for (const [name, run] of orderedChildren(p)) {
      if (name === 'r' || name === 'fld') line += getChildren(run, 't').map(textOf).join('')
    }
    out.push(line)
  }
  return out.join('\n').trim()
}
function titleOf(title: XmlNode | undefined): string | undefined {
  if (!title) return undefined
  const tx = child(title, 'tx')
  return richText(tx) || cacheStrings(tx)[0] || undefined
}

/** Labels (categories, series names) keep CHART_LABEL_LIMIT characters; a longer one is cut and recorded. */
function capLabel(text: string, cut: { truncated: boolean }): string {
  if (text.length <= CHART_LABEL_LIMIT) return text
  cut.truncated = true
  return text.slice(0, CHART_LABEL_LIMIT)
}
/** Saved data of a role: its cache (numRef/strRef/multiLvlStrRef) or literal. */
const cacheOf = (node: XmlNode | undefined) => child(child(node, 'numRef'), 'numCache') ?? child(child(node, 'strRef'), 'strCache')
  ?? child(child(child(node, 'multiLvlStrRef'), 'multiLvlStrCache'), 'lvl') ?? child(node, 'numLit') ?? child(node, 'strLit')
/**
 * The saved points of a data source by `pt@idx`, sparse indices kept: a missing
 * index (or ptCount padding) is `missing` at its position, never compacted.
 * `cut.truncated` records points dropped by the per-series bound (declared count or indices past it).
 */
function cachePoints<T>(node: XmlNode | undefined, read: (pt: XmlNode) => T, missing: T, cut?: { truncated: boolean }, limit = CHART_POINT_LIMIT): T[] {
  const cache = cacheOf(node)
  const out: T[] = []
  const declared = finite(val(child(cache, 'ptCount'))) ?? 0
  if (cut && declared > limit) cut.truncated = true
  const count = Math.max(0, Math.min(limit, Number.isSafeInteger(declared) ? declared : 0))
  for (const pt of getChildren(cache, 'pt')) {
    const index = finite(attrs(pt).idx)
    if (index === undefined || index < 0 || !Number.isInteger(index)) continue
    if (index >= limit) { if (cut) cut.truncated = true; continue }
    out[index] = read(pt)
  }
  const length = Math.max(out.length, count)
  return Array.from({ length }, (_, i) => i in out ? out[i] : missing)
}
function cacheStrings(node: XmlNode | undefined, cut?: { truncated: boolean }): Array<string | undefined> {
  return cachePoints<string | undefined>(node, pt => textOf(child(pt, 'v')), undefined, cut)
}
/** The label storage a cached category role keeps (each label capped at CHART_LABEL_LIMIT), read before it is built. */
function labelBytes(node: XmlNode | undefined): number {
  let bytes = 0
  for (const pt of getChildren(cacheOf(node), 'pt')) bytes += Math.min(textOf(child(pt, 'v')).length, CHART_LABEL_LIMIT) * 2 + 32
  return bytes
}
/** The extent of a saved numeric role, read without building storage: its length (sparse indices kept, ptCount
 * padding included), whether a declared count or an index exceeds `limit`, and whether points carry formats. */
function cacheExtent(node: XmlNode | undefined, limit: number): { length: number; over: boolean; formats: boolean } {
  const cache = cacheOf(node)
  const declared = finite(val(child(cache, 'ptCount'))) ?? 0
  let length = Math.max(0, Math.min(limit, Number.isSafeInteger(declared) ? declared : 0)), over = declared > limit, formats = false
  for (const pt of getChildren(cache, 'pt')) {
    const index = finite(attrs(pt).idx)
    if (index === undefined || index < 0 || !Number.isInteger(index)) continue
    if (index >= limit) { over = true; continue }
    length = Math.max(length, index + 1)
    if (attrs(pt).formatCode) formats = true
  }
  return { length, over, formats }
}
/**
 * A saved numeric role (spec 5.8): a number is a value (with its `pt@formatCode`), a known error code (#N/A kept
 * distinct) an error, empty or other text blank; a missing index is blank at its position. The storage is
 * admitted through `charge` before it is built.
 */
function numericCache(node: XmlNode | undefined, extent: { length: number; formats: boolean }, charge: (bytes: number) => void): ChartPoints {
  charge(pointBytes(extent.length, extent.formats))
  const out = new PointWriter(extent.length, extent.formats, charge)
  for (const pt of getChildren(cacheOf(node), 'pt')) {
    const index = finite(attrs(pt).idx)
    if (index === undefined || index < 0 || !Number.isInteger(index) || index >= extent.length) continue
    const text = textOf(child(pt, 'v'))
    const v = finite(text)
    if (v !== undefined) out.value(index, v, attrs(pt).formatCode)
    else if (knownError(text) !== undefined) out.error(index, text.trim())
  }
  return out.done()
}

function hasCache(node: XmlNode | undefined): boolean {
  return Boolean(child(child(node, 'numRef'), 'numCache') ?? child(child(node, 'strRef'), 'strCache') ?? child(child(node, 'multiLvlStrRef'), 'multiLvlStrCache') ?? child(node, 'numLit') ?? child(node, 'strLit'))
}
function formulaOf(node: XmlNode | undefined): string | undefined {
  const ref = child(node, 'numRef') ?? child(node, 'strRef') ?? child(node, 'multiLvlStrRef')
  const f = textOf(child(ref, 'f')).trim()
  return f || undefined
}

const colIndex = (letters: string): number => { let n = 0; for (const ch of letters.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1 }
const CELL = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/
const NAME = /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/
/** A single-area reference (`Sheet1!$A$1:$A$5`, `'My sheet'!B2`) or a defined name. Multi-area references are unresolved. */
export function parseChartRange(formula: string): ChartRange | undefined {
  let text = formula.trim()
  while (text.startsWith('(') && text.endsWith(')')) text = text.slice(1, -1).trim()
  if (text.includes(',')) return undefined
  let sheet: string | undefined
  const bang = text.lastIndexOf('!')
  if (bang >= 0) {
    const raw = text.slice(0, bang)
    sheet = raw.startsWith("'") && raw.endsWith("'") ? raw.slice(1, -1).replace(/''/g, "'") : raw
    text = text.slice(bang + 1)
  }
  const [from, to = from] = text.split(':')
  const a = CELL.exec(from), b = CELL.exec(to)
  // A defined name (workbook or sheet scope), resolved against the workbook later.
  if ((!a || !b) && !text.includes(':') && NAME.test(text)) return { ...(sheet ? { sheet } : {}), definedName: text, r0: 0, c0: 0, r1: 0, c1: 0 }
  if (!a || !b) return undefined
  const c0 = colIndex(a[1]), r0 = Number(a[2]) - 1, c1 = colIndex(b[1]), r1 = Number(b[2]) - 1
  if (![c0, r0, c1, r1].every(n => Number.isSafeInteger(n) && n >= 0)) return undefined
  return { ...(sheet ? { sheet } : {}), r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) }
}

const AXIS_KINDS: Record<string, ChartAxis['kind']> = { catAx: 'category', dateAx: 'date', valAx: 'value', serAx: 'series' }
function axisOf(node: XmlNode, kind: ChartAxis['kind'], index: number): ChartAxis {
  const scaling = child(node, 'scaling')
  const formatCode = attrs(child(node, 'numFmt')).formatCode
  const title = titleOf(child(node, 'title'))
  const crossAx = val(child(node, 'crossAx'))
  return {
    // An axis without an axId keeps a positional id (A1a validates references).
    id: val(child(node, 'axId')) ?? `#${index}`,
    kind,
    ...(crossAx !== undefined ? { crossAx } : {}),
    ...(finite(val(child(scaling, 'min'))) !== undefined ? { min: finite(val(child(scaling, 'min'))) } : {}),
    ...(finite(val(child(scaling, 'max'))) !== undefined ? { max: finite(val(child(scaling, 'max'))) } : {}),
    ...(finite(val(child(node, 'majorUnit'))) ? { majorUnit: finite(val(child(node, 'majorUnit'))) } : {}),
    gridlines: Boolean(child(node, 'majorGridlines')),
    deleted: val(child(node, 'delete')) === '1' || val(child(node, 'delete')) === 'true',
    ...(title ? { title } : {}),
    ...(formatCode && formatCode !== 'General' ? { formatCode } : {}),
  }
}

/** A group's axis binding, recorded structurally from its `c:axId`s (A1a validates it). */
function bindingOf(family: ChartFamily, node: XmlNode, axes: readonly ChartAxis[]): ChartAxisBinding {
  if (family === 'pie' || family === 'doughnut') return { kind: 'none' }
  const ids = getChildren(node, 'axId').map(val).filter((id): id is string => id !== undefined)
  if (family === 'scatter') return { kind: 'xy', x: ids[0] ?? '', y: ids[1] ?? '' }
  const kindOf = (id: string) => axes.find(axis => axis.id === id)?.kind
  const category = ids.find(id => kindOf(id) === 'category' || kindOf(id) === 'date')
  const value = ids.find(id => kindOf(id) === 'value') ?? ''
  return { kind: 'cartesian', ...(category !== undefined ? { category } : {}), value }
}
/** A series feeds an analytic element: a trendline, or stdDev/stdErr error bars (spec 7). */
const isAnalytic = (ser: XmlNode) => Boolean(child(ser, 'trendline')) || getChildren(ser, 'errBars').some(bars => /^std(Dev|Err)$/.test(val(child(bars, 'errValType')) ?? ''))
const flag = (node: XmlNode | undefined, fallback: boolean) => { const v = val(node); return v === undefined ? fallback : v === '1' || v === 'true' }
/** `c16r3:dataDisplayOptions16/c16r3:dispNaAsBlank` in the chart's extension list. */
function dispNaAsBlank(chart: XmlNode): boolean | undefined {
  for (const ext of getChildren(child(chart, 'extLst'), 'ext')) {
    const node = child(child(ext, 'dataDisplayOptions16'), 'dispNaAsBlank')
    if (node) return flag(node, true)
  }
  return undefined
}

/** Read one chart part. `undefined` for a part that is not a chart document. */
/** `_context` is the chart's parse context (F0 threads it; A1a/D0 consume it). */
export function parseXlsxChart(root: XmlNode, theme: ThemeContext | undefined, context?: ChartParseContext): ChartModel | undefined {
  const chart = child(root, 'chart')
  const plot = child(chart, 'plotArea')
  if (!chart || !plot) return undefined
  // The model's point storage is charged to the document's chart workspace before it is built, and stays
  // charged for the model's lifetime (charts/own.ts). A refusal stops the parse and refunds what it charged.
  const owner = context?.workspace ? modelOwner(context.workspace) : undefined
  const charge = (bytes: number) => { if (owner && !owner.workspace.reserve(bytes, owner.tag)) throw new ChartRefusal() }
  try {
    const model = readChart(root, chart, plot, theme, charge)
    if (owner) adoptChartModel(model, owner)
    return model
  } catch (error) {
    owner?.workspace.release(owner.tag)
    if (!(error instanceof ChartRefusal)) throw error
    return { ...placeholderChart(titleOf(child(chart, 'title')), 'This chart needs more memory than is available for charts'), refused: true }
  }
}
function readChart(root: XmlNode, chart: XmlNode, plot: XmlNode, theme: ThemeContext | undefined, charge: (bytes: number) => void): ChartModel {
  const minorFont = theme?.fonts.minor.latin ?? 'Calibri'
  const sizeRaw = finite(attrs(getChildren(child(child(root, 'txPr'), 'p'), 'pPr').map(p => child(p, 'defRPr'))[0]).sz)
  const autoDeleted = val(child(chart, 'autoTitleDeleted')) === '1'
  let title = titleOf(child(chart, 'title'))
  const groups: ChartGroup[] = []
  let unsupported: string | undefined, truncated = false, seriesTotal = 0
  const cut = { truncated: false }
  let categories: string[] = [], categoryPending: ChartRange | undefined, categoryRef: ChartRange | undefined
  const axes: ChartAxis[] = []
  for (const [name, node] of orderedChildren(plot)) if (AXIS_KINDS[name]) axes.push(axisOf(node, AXIS_KINDS[name], axes.length))
  for (const [name, node] of orderedChildren(plot)) {
    if (UNSUPPORTED[name]) { unsupported ??= `${UNSUPPORTED[name]} is not supported`; continue }
    const type = GROUPS[name]
    if (!type) continue
    const grouping = (val(child(node, 'grouping')) ?? (type === 'bar' ? 'clustered' : 'standard')) as ChartGroup['grouping']
    const scatterStyle = val(child(node, 'scatterStyle')) ?? 'marker'
    const vary = type === 'pie' || type === 'doughnut' ? val(child(node, 'varyColors')) !== '0' : val(child(node, 'varyColors')) === '1'
    const series: ChartSeries[] = []
    for (const ser of getChildren(node, 'ser')) {
      if (seriesTotal >= CHART_SERIES_LIMIT) { truncated = true; break }
      // Series indices are unsigned integers; anything else falls back to the series position.
      const rawIndex = finite(val(child(ser, 'idx')))
      const index = rawIndex !== undefined && Number.isSafeInteger(rawIndex) && rawIndex >= 0 ? rawIndex : seriesTotal
      seriesTotal++
      const spPr = child(ser, 'spPr')
      const lineColor = solid(spPr, theme, true)
      const fillColor = solid(spPr, theme)
      const color = (type === 'line' || type === 'scatter' ? lineColor ?? fillColor : fillColor ?? lineColor) ?? accent(theme, index)
      const tx = child(ser, 'tx')
      const valueNode = child(ser, type === 'scatter' ? 'yVal' : 'val')
      const xNode = type === 'scatter' ? child(ser, 'xVal') : undefined
      const catNode = type === 'scatter' ? undefined : child(ser, 'cat')
      const pending: NonNullable<ChartSeries['pending']> = {}, refs: NonNullable<ChartSeries['refs']> = {}
      const rangeOf = (node: XmlNode | undefined) => { const f = formulaOf(node); return f ? parseChartRange(f) : undefined }
      // An uncached reference the viewer cannot read (multi-area, a formula) has no data: placeholder, never an empty chart.
      const unreadable = (node: XmlNode | undefined) => { const f = !hasCache(node) ? formulaOf(node) : undefined; if (f && !parseChartRange(f)) unsupported ??= `Chart data uses a reference the viewer cannot read (${f.slice(0, 80)})` }
      unreadable(valueNode); unreadable(xNode); unreadable(catNode)
      const txFormula = child(tx, 'strRef') && !child(child(tx, 'strRef'), 'strCache') ? formulaOf(tx) : undefined
      const nameRange = txFormula ? parseChartRange(txFormula) : undefined
      if (nameRange) pending.name = nameRange
      const literalName = textOf(child(tx, 'v')) || cacheStrings(tx)[0]
      // Analytic inputs are acquired in full (up to CHART_ANALYTIC_POINT_LIMIT) before any display cap;
      // more than that refuses the analytic element (spec 7), never a prefix: the series is then drawn capped.
      // Extents are read first, so nothing is built before the decision and its admission.
      let analytic: ChartSeries['analytic'] = isAnalytic(ser) ? true : undefined
      const limit = analytic ? CHART_ANALYTIC_POINT_LIMIT : CHART_POINT_LIMIT
      let valueExtent = cacheExtent(valueNode, limit), xExtent = xNode ? cacheExtent(xNode, limit) : undefined
      if (analytic && (valueExtent.over || xExtent?.over)) {
        analytic = 'refused'
        valueExtent = cacheExtent(valueNode, CHART_POINT_LIMIT); xExtent = xNode ? cacheExtent(xNode, CHART_POINT_LIMIT) : undefined
        cut.truncated = true
      } else if (valueExtent.over || xExtent?.over) cut.truncated = true
      const values = numericCache(valueNode, valueExtent, charge)
      const xValues = xNode && xExtent ? numericCache(xNode, xExtent, charge) : undefined
      // A role's reference is kept beside its cache: the cache gives the values, the reference the provenance.
      if (!hasCache(valueNode)) { const range = rangeOf(valueNode); if (range) pending.values = range } else { const range = rangeOf(valueNode); if (range) refs.values = range }
      if (xNode) { const range = rangeOf(xNode); if (range) { if (hasCache(xNode)) refs.xValues = range; else pending.xValues = range } }
      if (catNode && categories.length === 0 && !categoryPending) {
        if (hasCache(catNode)) { charge(labelBytes(catNode)); categories = cacheStrings(catNode, cut).map(v => capLabel(v ?? '', cut)); categoryRef = rangeOf(catNode) }
        else { const range = rangeOf(catNode); if (range) categoryPending = range }
      }
      // Display points past the bound are dropped at draw time (analytic series keep them for analysis).
      if (values.length > CHART_POINT_LIMIT || (xValues && xValues.length > CHART_POINT_LIMIT)) truncated = true
      const pointColors: string[] = []
      if (vary) for (let i = 0; i < Math.max(Math.min(values.length, CHART_POINT_LIMIT), 1) && i < 256; i++) pointColors[i] = accent(theme, i)
      for (const dPt of getChildren(ser, 'dPt')) {
        const i = finite(val(child(dPt, 'idx')))
        const c = solid(child(dPt, 'spPr'), theme)
        if (i !== undefined && i >= 0 && i < 256 && c) pointColors[i] = c
      }
      const markerNone = val(child(child(ser, 'marker'), 'symbol')) === 'none'
      const groupMarker = val(child(node, 'marker'))
      series.push({
        name: capLabel(literalName || `Series ${index + 1}`, cut),
        color,
        values,
        ...(xValues ? { xValues } : {}),
        ...(pointColors.length ? { pointColors } : {}),
        ...(Object.keys(pending).length ? { pending } : {}),
        ...(Object.keys(refs).length ? { refs } : {}),
        ...(analytic ? { analytic } : {}),
        ...(type === 'line' ? { line: lineColor !== null, markers: !markerNone && groupMarker !== '0' } : {}),
        ...(type === 'scatter' ? { line: lineColor !== null && /line|smooth/i.test(scatterStyle), markers: !markerNone && !/^(line|smooth)$/i.test(scatterStyle) } : {}),
      })
    }
    groups.push({
      family: type, grouping, series, binding: bindingOf(type, node, axes),
      ...(type === 'bar' ? { barDir: val(child(node, 'barDir')) === 'bar' ? 'bar' as const : 'col' as const } : {}),
      gapWidth: Math.max(0, Math.min(500, finite(val(child(node, 'gapWidth'))) ?? 150)),
      overlap: Math.max(-100, Math.min(100, finite(val(child(node, 'overlap'))) ?? (grouping === 'stacked' || grouping === 'percentStacked' ? 100 : 0))),
      ...(type === 'doughnut' ? { holeSize: Math.max(10, Math.min(90, finite(val(child(node, 'holeSize'))) ?? 50)) } : {}),
      ...(type === 'pie' || type === 'doughnut' ? { firstSliceAngle: finite(val(child(node, 'firstSliceAng'))) ?? 0 } : {}),
    })
  }
  const pies = groups.filter(g => g.family === 'pie' || g.family === 'doughnut')
  if (!groups.length) unsupported ??= 'This chart type is not supported'
  else if (!groups.some(g => g.series.length)) unsupported ??= 'This chart has no data series'
  else if (pies.length && groups.length > 1) unsupported ??= 'Combination pie charts are not supported'
  if (child(chart, 'title') && !title && !autoDeleted) title = groups.flatMap(g => g.series).length === 1 && !groups[0].series[0].pending?.name ? groups[0].series[0].name : val(child(root, 'lang'))?.startsWith('zh') ? '圖表標題' : 'Chart Title'
  const legend = child(chart, 'legend')
  const blanks = val(child(chart, 'dispBlanksAs'))
  const naAsBlank = dispNaAsBlank(chart)
  const model: ChartModel = {
    ...(title ? { title } : {}),
    groups,
    axes,
    categories,
    ...(categoryPending ? { categoryPending } : {}),
    ...(categoryRef ? { categoryRef } : {}),
    ...(legend ? { legend: (val(child(legend, 'legendPos')) ?? 'r') as ChartModel['legend'] } : {}),
    // Chart-level options as authored; absent stays undefined (Office defaults: gap, plot visible only).
    ...(child(chart, 'dispBlanksAs') ? { dispBlanksAs: blanks === 'gap' || blanks === 'span' ? blanks : 'zero' } : {}),
    ...(child(chart, 'plotVisOnly') ? { plotVisOnly: flag(child(chart, 'plotVisOnly'), true) } : {}),
    ...(naAsBlank !== undefined ? { dispNaAsBlank: naAsBlank } : {}),
    ...(child(root, 'date1904') ? { epoch: flag(child(root, 'date1904'), true) ? '1904' as const : '1900' as const } : {}),
    fontFamily: minorFont,
    fontSizePt: sizeRaw !== undefined ? Math.max(6, Math.min(36, sizeRaw / 100)) : 10,
    ...(unsupported ? { unsupported } : {}),
    ...(truncated || cut.truncated ? { truncated: true } : {}),
  }
  return model
}

/** A placeholder model for a chart that could not be read. */
export function placeholderChart(title: string | undefined, reason: string): ChartModel {
  return { ...(title ? { title } : {}), groups: [], axes: [], categories: [], fontFamily: 'Calibri', fontSizePt: 10, unsupported: reason }
}
