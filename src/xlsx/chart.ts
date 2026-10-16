/**
 * Worksheet charts (lazy chunk): reads c:chartSpace into an XlsxChartModel,
 * fills uncached series from worksheet cells and paints the common chart
 * types (bar/column clustered/stacked/100%, line, area, pie, doughnut,
 * scatter). Anything else becomes a placeholder card carrying the chart title
 * and an accessible description; nothing is ever left as silent blank space.
 *
 * Loaded through `chartLoader.ts` only when a sheet has a chart; importing it
 * registers the painter with the shared content painter.
 */
import { attrs, getChildren, orderedChildren, textOf, type XmlNode } from '../core/xml'
import { parseDrawingColor, resolveDrawingColor, type ThemeContext } from '../drawing/style'
import { paintChartPlaceholder, registerXlsxChartPainter } from '../drawing/content-paint'
import { chartDescription, formatNumber } from './chartDescription'
import type { XlsxChartAxis, XlsxChartGroup, XlsxChartModel, XlsxChartRange, XlsxChartSeries, XlsxChartType } from './chartModel'

/** Bounds: series per chart, points per series. */
export const CHART_SERIES_LIMIT = 64
export const CHART_POINT_LIMIT = 4096

const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]
const val = (node: XmlNode | undefined) => attrs(node).val
const finite = (raw: string | undefined): number | undefined => raw !== undefined && raw.trim() !== '' && Number.isFinite(Number(raw)) ? Number(raw) : undefined
const DEFAULT_ACCENTS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']

const GROUPS: Record<string, XlsxChartType> = {
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

/** Cached point values of a data source (numRef/strRef caches or literals). */
/** `cut.truncated` records points dropped by the per-series bound (declared count or indices past it). */
function cacheStrings(node: XmlNode | undefined, cut?: { truncated: boolean }): Array<string | undefined> {
  const cache = child(child(node, 'numRef'), 'numCache') ?? child(child(node, 'strRef'), 'strCache')
    ?? child(child(child(node, 'multiLvlStrRef'), 'multiLvlStrCache'), 'lvl') ?? child(node, 'numLit') ?? child(node, 'strLit')
  const out: Array<string | undefined> = []
  const declared = finite(val(child(cache, 'ptCount'))) ?? 0
  if (cut && declared > CHART_POINT_LIMIT) cut.truncated = true
  const count = Math.max(0, Math.min(CHART_POINT_LIMIT, Number.isSafeInteger(declared) ? declared : 0))
  for (const pt of getChildren(cache, 'pt')) {
    const index = finite(attrs(pt).idx)
    if (index === undefined || index < 0 || !Number.isInteger(index)) continue
    if (index >= CHART_POINT_LIMIT) { if (cut) cut.truncated = true; continue }
    out[index] = textOf(child(pt, 'v'))
  }
  const length = Math.max(out.length, count)
  return Array.from({ length }, (_, i) => out[i])
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
export function parseChartRange(formula: string): XlsxChartRange | undefined {
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

const numbers = (values: Array<string | undefined>) => values.map(v => finite(v))

function axisOf(node: XmlNode | undefined): XlsxChartAxis {
  const scaling = child(node, 'scaling')
  const formatCode = attrs(child(node, 'numFmt')).formatCode
  const title = titleOf(child(node, 'title'))
  return {
    ...(finite(val(child(scaling, 'min'))) !== undefined ? { min: finite(val(child(scaling, 'min'))) } : {}),
    ...(finite(val(child(scaling, 'max'))) !== undefined ? { max: finite(val(child(scaling, 'max'))) } : {}),
    ...(finite(val(child(node, 'majorUnit'))) ? { majorUnit: finite(val(child(node, 'majorUnit'))) } : {}),
    gridlines: Boolean(child(node, 'majorGridlines')),
    deleted: val(child(node, 'delete')) === '1' || val(child(node, 'delete')) === 'true',
    ...(title ? { title } : {}),
    ...(formatCode && formatCode !== 'General' ? { formatCode } : {}),
  }
}

/** Read one chart part. `undefined` for a part that is not a chart document. */
export function parseXlsxChart(root: XmlNode, theme: ThemeContext | undefined): XlsxChartModel | undefined {
  const chart = child(root, 'chart')
  const plot = child(chart, 'plotArea')
  if (!chart || !plot) return undefined
  const minorFont = theme?.fonts.minor.latin ?? 'Calibri'
  const sizeRaw = finite(attrs(getChildren(child(child(root, 'txPr'), 'p'), 'pPr').map(p => child(p, 'defRPr'))[0]).sz)
  const autoDeleted = val(child(chart, 'autoTitleDeleted')) === '1'
  let title = titleOf(child(chart, 'title'))
  const groups: XlsxChartGroup[] = []
  let unsupported: string | undefined, truncated = false, seriesTotal = 0
  const cut = { truncated: false }
  let categories: string[] = [], categoryPending: XlsxChartRange | undefined
  for (const [name, node] of orderedChildren(plot)) {
    if (UNSUPPORTED[name]) { unsupported ??= `${UNSUPPORTED[name]} is not supported`; continue }
    const type = GROUPS[name]
    if (!type) continue
    const grouping = (val(child(node, 'grouping')) ?? (type === 'bar' ? 'clustered' : 'standard')) as XlsxChartGroup['grouping']
    const scatterStyle = val(child(node, 'scatterStyle')) ?? 'marker'
    const vary = type === 'pie' || type === 'doughnut' ? val(child(node, 'varyColors')) !== '0' : val(child(node, 'varyColors')) === '1'
    const series: XlsxChartSeries[] = []
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
      const pending: NonNullable<XlsxChartSeries['pending']> = {}
      // An uncached reference the viewer cannot read (multi-area, a formula) has no data: placeholder, never an empty chart.
      const unreadable = (node: XmlNode | undefined) => { const f = !hasCache(node) ? formulaOf(node) : undefined; if (f && !parseChartRange(f)) unsupported ??= `Chart data uses a reference the viewer cannot read (${f.slice(0, 80)})` }
      unreadable(valueNode); unreadable(xNode); unreadable(catNode)
      const txFormula = child(tx, 'strRef') && !child(child(tx, 'strRef'), 'strCache') ? formulaOf(tx) : undefined
      const nameRange = txFormula ? parseChartRange(txFormula) : undefined
      if (nameRange) pending.name = nameRange
      const literalName = textOf(child(tx, 'v')) || cacheStrings(tx)[0]
      let values = numbers(cacheStrings(valueNode, cut))
      if (!hasCache(valueNode)) { const range = formulaOf(valueNode) && parseChartRange(formulaOf(valueNode)!); if (range) pending.values = range }
      let xValues = xNode ? numbers(cacheStrings(xNode, cut)) : undefined
      if (xNode && !hasCache(xNode)) { const range = formulaOf(xNode) && parseChartRange(formulaOf(xNode)!); if (range) pending.xValues = range }
      if (catNode && categories.length === 0 && !categoryPending) {
        if (hasCache(catNode)) categories = cacheStrings(catNode, cut).map(v => v ?? '')
        else { const range = formulaOf(catNode) && parseChartRange(formulaOf(catNode)!); if (range) categoryPending = range }
      }
      if (values.length > CHART_POINT_LIMIT) { values = values.slice(0, CHART_POINT_LIMIT); truncated = true }
      if (xValues && xValues.length > CHART_POINT_LIMIT) { xValues = xValues.slice(0, CHART_POINT_LIMIT); truncated = true }
      const pointColors: string[] = []
      if (vary) for (let i = 0; i < Math.max(values.length, 1) && i < 256; i++) pointColors[i] = accent(theme, i)
      for (const dPt of getChildren(ser, 'dPt')) {
        const i = finite(val(child(dPt, 'idx')))
        const c = solid(child(dPt, 'spPr'), theme)
        if (i !== undefined && i >= 0 && i < 256 && c) pointColors[i] = c
      }
      const markerNone = val(child(child(ser, 'marker'), 'symbol')) === 'none'
      const groupMarker = val(child(node, 'marker'))
      series.push({
        name: literalName || `Series ${index + 1}`,
        color,
        values,
        ...(xValues ? { xValues } : {}),
        ...(pointColors.length ? { pointColors } : {}),
        ...(Object.keys(pending).length ? { pending } : {}),
        ...(type === 'line' ? { line: lineColor !== null, markers: !markerNone && groupMarker !== '0' } : {}),
        ...(type === 'scatter' ? { line: lineColor !== null && /line|smooth/i.test(scatterStyle), markers: !markerNone && !/^(line|smooth)$/i.test(scatterStyle) } : {}),
      })
    }
    groups.push({
      type, grouping, series,
      ...(type === 'bar' ? { barDir: val(child(node, 'barDir')) === 'bar' ? 'bar' as const : 'col' as const } : {}),
      gapWidth: Math.max(0, Math.min(500, finite(val(child(node, 'gapWidth'))) ?? 150)),
      overlap: Math.max(-100, Math.min(100, finite(val(child(node, 'overlap'))) ?? (grouping === 'stacked' || grouping === 'percentStacked' ? 100 : 0))),
      ...(type === 'doughnut' ? { holeSize: Math.max(10, Math.min(90, finite(val(child(node, 'holeSize'))) ?? 50)) } : {}),
      ...(type === 'pie' || type === 'doughnut' ? { firstSliceAngle: finite(val(child(node, 'firstSliceAng'))) ?? 0 } : {}),
    })
  }
  const pies = groups.filter(g => g.type === 'pie' || g.type === 'doughnut')
  if (!groups.length) unsupported ??= 'This chart type is not supported'
  else if (!groups.some(g => g.series.length)) unsupported ??= 'This chart has no data series'
  else if (pies.length && groups.length > 1) unsupported ??= 'Combination pie charts are not supported'
  if (child(chart, 'title') && !title && !autoDeleted) title = groups.flatMap(g => g.series).length === 1 && !groups[0].series[0].pending?.name ? groups[0].series[0].name : 'Chart Title'
  const legend = child(chart, 'legend')
  const model: XlsxChartModel = {
    ...(title ? { title } : {}),
    groups,
    categories,
    ...(categoryPending ? { categoryPending } : {}),
    ...(legend ? { legend: (val(child(legend, 'legendPos')) ?? 'r') as XlsxChartModel['legend'] } : {}),
    categoryAxis: axisOf(child(plot, 'catAx') ?? child(plot, 'dateAx') ?? (groups[0]?.type === 'scatter' ? getChildren(plot, 'valAx')[0] : undefined)),
    valueAxis: axisOf(groups[0]?.type === 'scatter' ? getChildren(plot, 'valAx')[1] : child(plot, 'valAx')),
    fontFamily: minorFont,
    fontSizePt: sizeRaw !== undefined ? Math.max(6, Math.min(36, sizeRaw / 100)) : 10,
    ...(unsupported ? { unsupported } : {}),
    ...(truncated || cut.truncated ? { truncated: true } : {}),
  }
  return model
}

/** A placeholder model for a chart that could not be read. */
export function placeholderChart(title: string | undefined, reason: string): XlsxChartModel {
  return { ...(title ? { title } : {}), groups: [], categories: [], categoryAxis: { gridlines: false, deleted: false }, valueAxis: { gridlines: false, deleted: false }, fontFamily: 'Calibri', fontSizePt: 10, unsupported: reason }
}

export type ChartCellValue = string | number | boolean | null | undefined
/** Cell lookup for filling uncached references: `sheet` is the resolved sheet name. */
export interface ChartCell { value: ChartCellValue; /** The cell's formatted text (number/date format), as Excel shows it. */ text?: string }
/** `text`: the caller needs the formatted label (categories, names); values never format. */
export type ChartCellLookup = (sheet: string, row: number, col: number, text: boolean) => ChartCell | undefined

/** The one line of cells a chart consumes from a range (a 2-D range keeps its
 * last row/column, the level nearest the data), bounded by CHART_POINT_LIMIT.
 * Readers and the fill use this same sequence. */
export interface ChartLine { sheet: string; vertical: boolean; fixed: number; start: number; count: number; truncated: boolean; text: boolean }
export function chartLine(range: XlsxChartRange, ownerSheet: string, text: boolean): ChartLine {
  const rows = range.r1 - range.r0 + 1, cols = range.c1 - range.c0 + 1
  const vertical = cols === 1 || rows >= cols
  const length = vertical ? rows : cols
  return { sheet: range.sheet ?? ownerSheet, vertical, fixed: vertical ? range.c1 : range.r1, start: vertical ? range.r0 : range.c0, count: Math.min(length, CHART_POINT_LIMIT), truncated: length > CHART_POINT_LIMIT, text }
}
/** The cell coordinates of point `i` on a line. */
export const linePoint = (line: ChartLine, i: number): { row: number; col: number } => line.vertical ? { row: line.start + i, col: line.fixed } : { row: line.fixed, col: line.start + i }

export function chartHasPending(model: XlsxChartModel): boolean {
  return Boolean(model.categoryPending) || model.groups.some(g => g.series.some(s => s.pending))
}
/** The cells a range resolves to, row-major, bounded by CHART_POINT_LIMIT. */
function rangeCells(range: XlsxChartRange, sheet: string, lookup: ChartCellLookup, text: boolean): { values: Array<ChartCell | undefined>; truncated: boolean } {
  const line = chartLine(range, sheet, text)
  const values: Array<ChartCell | undefined> = []
  for (let i = 0; i < line.count; i++) { const at = linePoint(line, i); values.push(lookup(line.sheet, at.row, at.col, text)) }
  return { values, truncated: line.truncated }
}
const asNumber = (cell: ChartCell | undefined): number | undefined => numberOf(cell?.value)
const numberOf = (value: ChartCellValue): number | undefined => typeof value === 'number' && Number.isFinite(value) ? value : typeof value === 'boolean' ? undefined : typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value)) ? Number(value) : undefined
const rawText = (value: ChartCellValue): string => value === null || value === undefined ? '' : typeof value === 'number' ? String(Number(value.toPrecision(15))) : String(value)
const asText = (cell: ChartCell | undefined): string => cell?.text ?? rawText(cell?.value)

/** Resolves a defined name seen from a sheet to its formula (sheet scope first, then workbook). */
export type ChartNameResolver = (name: string, scopeSheet: string) => string | undefined
export interface ChartDefinedName { name: string; source: string; localSheetIndex?: number }
export function chartNameResolver(names: readonly ChartDefinedName[], sheets: ReadonlyArray<{ name: string; workbookIndex?: number }>): ChartNameResolver {
  return (name, scopeSheet) => {
    const upper = name.toUpperCase()
    const scope = sheets.find(sheet => sheet.name.toLowerCase() === scopeSheet.toLowerCase())
    const candidates = names.filter(entry => entry.name.toUpperCase() === upper)
    const local = scope?.workbookIndex !== undefined ? candidates.find(entry => entry.localSheetIndex === scope.workbookIndex) : undefined
    return (local ?? candidates.find(entry => entry.localSheetIndex === undefined))?.source
  }
}
/** Replace defined-name references by the ranges they name; a name that is
 * missing or not a single area turns the chart into a placeholder. */
export function resolveChartNames(model: XlsxChartModel, ownerSheet: string, resolve: ChartNameResolver | undefined): XlsxChartModel {
  let unsupported = model.unsupported
  const fix = (range: XlsxChartRange | undefined): XlsxChartRange | undefined => {
    if (!range?.definedName) return range
    const formula = resolve?.(range.definedName, range.sheet ?? ownerSheet)
    const target = formula ? parseChartRange(formula.replace(/^=/, '')) : undefined
    if (target && !target.definedName) return target
    unsupported ??= `Chart data name “${range.definedName}” ${formula ? 'is not a single worksheet range' : 'is not defined in this workbook'}`
    return undefined
  }
  if (!chartHasPending(model)) return model
  const categoryPending = fix(model.categoryPending)
  const groups = model.groups.map(group => ({ ...group, series: group.series.map(series => {
    if (!series.pending) return series
    const pending = Object.fromEntries(Object.entries(series.pending).map(([key, range]) => [key, fix(range)]).filter(([, range]) => range)) as NonNullable<XlsxChartSeries['pending']>
    const { pending: _old, ...rest } = series
    void _old
    return Object.keys(pending).length ? { ...rest, pending } : rest
  }) }))
  const { categoryPending: _cat, ...rest } = model
  void _cat
  return { ...rest, groups, ...(categoryPending ? { categoryPending } : {}), ...(unsupported ? { unsupported } : {}) }
}

/** Fill uncached references from worksheet cells (returns a new model). */
export function fillChartRanges(model: XlsxChartModel, ownerSheet: string, lookup: ChartCellLookup): XlsxChartModel {
  let truncated = model.truncated ?? false
  const sheetOf = (range: XlsxChartRange) => range.sheet ?? ownerSheet
  let categories = model.categories
  if (model.categoryPending) {
    const cells = rangeCells(model.categoryPending, sheetOf(model.categoryPending), lookup, true)
    truncated ||= cells.truncated
    categories = cells.values.map(asText)
  }
  const groups = model.groups.map(group => ({
    ...group,
    series: group.series.map(series => {
      if (!series.pending) return series
      const { pending, ...rest } = series
      const next: XlsxChartSeries = { ...rest }
      if (pending.values) { const cells = rangeCells(pending.values, sheetOf(pending.values), lookup, false); truncated ||= cells.truncated; next.values = cells.values.map(asNumber) }
      if (pending.xValues) { const cells = rangeCells(pending.xValues, sheetOf(pending.xValues), lookup, false); truncated ||= cells.truncated; next.xValues = cells.values.map(asNumber) }
      if (pending.name) { const text = asText(lookup(sheetOf(pending.name), pending.name.r0, pending.name.c0, true)); if (text) next.name = text }
      if (next.pointColors && next.values.length > next.pointColors.length) {
        const first = next.pointColors
        next.pointColors = Array.from({ length: Math.min(256, next.values.length) }, (_, i) => first[i] ?? first[i % first.length])
      }
      return next
    }),
  }))
  const { categoryPending, ...rest } = model
  void categoryPending
  return { ...rest, categories, groups, ...(truncated ? { truncated: true } : {}) }
}

// ---------------------------------------------------------------- painting
type FamilyCss = (family: string) => string
function niceScale(lo: number, hi: number, axis: XlsxChartAxis, percent: boolean): { min: number; max: number; step: number } {
  if (percent) return { min: axis.min ?? 0, max: axis.max ?? 1, step: axis.majorUnit ?? 0.2 }
  let min = axis.min ?? Math.min(0, lo), max = axis.max ?? Math.max(0, hi)
  if (!(max > min)) max = min + 1
  const raw = (max - min) / 5
  const base = 10 ** Math.floor(Math.log10(raw))
  const step = axis.majorUnit && axis.majorUnit > 0 ? axis.majorUnit : [1, 2, 5, 10].map(m => m * base).find(m => m >= raw) ?? base * 10
  if (axis.min === undefined) min = Math.floor(min / step) * step
  if (axis.max === undefined) max = Math.ceil(max / step) * step
  if (!(max > min)) max = min + step
  return { min, max, step }
}
const font = (sizePx: number, family: string, familyCss: FamilyCss, bold = false) => `${bold ? 'bold ' : ''}${sizePx}px ${familyCss(family)}`
const TEXT = '#595959', GRID = '#D9D9D9'

interface Box { x: number; y: number; w: number; h: number }
function legendEntries(model: XlsxChartModel): Array<{ name: string; color: string }> {
  const group = model.groups[0]
  if (group && (group.type === 'pie' || group.type === 'doughnut')) {
    const s = group.series[0]
    return (s?.values ?? []).map((_, i) => ({ name: model.categories[i] ?? String(i + 1), color: s.pointColors?.[i] ?? s.color })).slice(0, 64)
  }
  return model.groups.flatMap(g => g.series).map(s => ({ name: s.name, color: s.color }))
}

function paintLegend(ctx: CanvasRenderingContext2D, model: XlsxChartModel, area: Box, fontPx: number, family: string, familyCss: FamilyCss): Box {
  const entries = legendEntries(model)
  if (!model.legend || !entries.length) return area
  ctx.font = font(fontPx, family, familyCss)
  const sw = fontPx * 0.7, gap = 6
  const widths = entries.map(e => ctx.measureText(e.name).width + sw + gap * 2)
  const line = fontPx + 6
  if (model.legend === 'r' || model.legend === 'l' || model.legend === 'tr') {
    const width = Math.min(area.w * 0.4, Math.max(...widths) + gap)
    const total = Math.min(entries.length, Math.max(1, Math.floor(area.h / line))) * line
    const x = model.legend === 'l' ? area.x : area.x + area.w - width
    let y = model.legend === 'tr' ? area.y : area.y + (area.h - total) / 2
    ctx.textBaseline = 'middle'; ctx.textAlign = 'left'
    for (const entry of entries) {
      if (y + line > area.y + area.h + 0.5) break
      ctx.fillStyle = `#${entry.color}`
      ctx.fillRect(x + gap, y + (line - sw) / 2, sw, sw)
      ctx.fillStyle = TEXT
      ctx.fillText(entry.name, x + gap * 2 + sw, y + line / 2, width - sw - gap * 3)
      y += line
    }
    return model.legend === 'l' ? { x: area.x + width, y: area.y, w: area.w - width, h: area.h } : { ...area, w: area.w - width }
  }
  const total = widths.reduce((a, b) => a + b, 0)
  let x = area.x + Math.max(0, (area.w - total) / 2)
  const y = model.legend === 't' ? area.y : area.y + area.h - line
  ctx.textBaseline = 'middle'; ctx.textAlign = 'left'
  entries.forEach((entry, i) => {
    if (x + widths[i] > area.x + area.w + 0.5) return
    ctx.fillStyle = `#${entry.color}`
    ctx.fillRect(x + gap, y + (line - sw) / 2, sw, sw)
    ctx.fillStyle = TEXT
    ctx.fillText(entry.name, x + gap * 2 + sw, y + line / 2)
    x += widths[i]
  })
  return model.legend === 't' ? { x: area.x, y: area.y + line, w: area.w, h: area.h - line } : { ...area, h: area.h - line }
}

function paintPie(ctx: CanvasRenderingContext2D, model: XlsxChartModel, group: XlsxChartGroup, box: Box): void {
  const series = group.series[0]
  if (!series) return
  const values = series.values.map(v => v !== undefined && v > 0 ? v : 0)
  const total = values.reduce((a, b) => a + b, 0)
  if (!(total > 0)) return
  const r = Math.max(1, Math.min(box.w, box.h) / 2 - 4), cx = box.x + box.w / 2, cy = box.y + box.h / 2
  let angle = ((group.firstSliceAngle ?? 0) - 90) * Math.PI / 180
  values.forEach((value, i) => {
    if (!value) return
    const sweep = value / total * Math.PI * 2
    ctx.beginPath()
    ctx.moveTo(cx, cy)
    ctx.arc(cx, cy, r, angle, angle + sweep)
    ctx.closePath()
    ctx.fillStyle = `#${series.pointColors?.[i] ?? series.color}`
    ctx.fill()
    ctx.strokeStyle = '#FFFFFF'; ctx.lineWidth = 1; ctx.stroke()
    angle += sweep
  })
  if (group.type === 'doughnut') {
    ctx.beginPath()
    ctx.arc(cx, cy, r * (group.holeSize ?? 50) / 100, 0, Math.PI * 2)
    ctx.fillStyle = '#FFFFFF'
    ctx.fill()
  }
  void model
}

function paintCartesian(ctx: CanvasRenderingContext2D, model: XlsxChartModel, box: Box, fontPx: number, family: string, familyCss: FamilyCss): void {
  const scatter = model.groups.every(g => g.type === 'scatter')
  const horizontal = model.groups.some(g => g.type === 'bar' && g.barDir === 'bar')
  const percent = model.groups.some(g => g.grouping === 'percentStacked')
  const count = Math.max(model.categories.length, ...model.groups.flatMap(g => g.series.map(s => s.values.length)), 1)
  // Value extents, honouring stacking.
  let lo = Infinity, hi = -Infinity
  const stackedValues = new Map<XlsxChartGroup, { pos: number[]; neg: number[]; total: number[] }>()
  for (const group of model.groups) {
    if (group.grouping === 'stacked' || group.grouping === 'percentStacked') {
      const pos = new Array<number>(count).fill(0), neg = new Array<number>(count).fill(0), total = new Array<number>(count).fill(0)
      for (const s of group.series) s.values.forEach((v, i) => { if (v === undefined || i >= count) return; if (v >= 0) pos[i] += v; else neg[i] += v; total[i] += Math.abs(v) })
      stackedValues.set(group, { pos, neg, total })
      if (group.grouping === 'stacked') for (let i = 0; i < count; i++) { lo = Math.min(lo, neg[i]); hi = Math.max(hi, pos[i]) }
    } else for (const s of group.series) for (const v of s.values) if (v !== undefined) { lo = Math.min(lo, v); hi = Math.max(hi, v) }
  }
  if (!Number.isFinite(lo)) { lo = 0; hi = 1 }
  const scale = niceScale(lo, hi, model.valueAxis, percent)
  let xs: { min: number; max: number; step: number } | undefined
  if (scatter) {
    let xlo = Infinity, xhi = -Infinity
    for (const s of model.groups.flatMap(g => g.series)) (s.xValues?.length ? s.xValues : s.values.map((_, i) => i + 1)).forEach(v => { if (v !== undefined) { xlo = Math.min(xlo, v); xhi = Math.max(xhi, v) } })
    if (!Number.isFinite(xlo)) { xlo = 0; xhi = 1 }
    xs = niceScale(xlo, xhi, model.categoryAxis, false)
  }
  ctx.font = font(fontPx, family, familyCss)
  const ticks: number[] = []
  for (let v = scale.min, n = 0; v <= scale.max + scale.step / 1e6 && n < 200; v += scale.step, n++) ticks.push(Number(v.toPrecision(12)))
  const label = (v: number) => formatNumber(v, percent)
  const valueLabelWidth = model.valueAxis.deleted ? 0 : Math.max(...ticks.map(t => ctx.measureText(label(t)).width)) + 6
  const catLabelHeight = model.categoryAxis.deleted ? 0 : fontPx + 6
  const valueTitle = model.valueAxis.title, catTitle = model.categoryAxis.title
  const titleBand = fontPx + 6
  let plot: Box = { ...box }
  if (horizontal) {
    const catWidth = model.categoryAxis.deleted ? 0 : Math.min(box.w * 0.35, Math.max(0, ...model.categories.map(c => ctx.measureText(c).width)) + 8)
    plot = { x: box.x + catWidth + (catTitle ? titleBand : 0), y: box.y + 4, w: box.w - catWidth - (catTitle ? titleBand : 0) - 8, h: box.h - catLabelHeight - (valueTitle ? titleBand : 0) - 4 }
  } else {
    plot = { x: box.x + valueLabelWidth + (valueTitle ? titleBand : 0), y: box.y + 4, w: box.w - valueLabelWidth - (valueTitle ? titleBand : 0) - 8, h: box.h - catLabelHeight - (catTitle ? titleBand : 0) - 4 }
  }
  if (plot.w <= 4 || plot.h <= 4) return
  const vpos = (v: number) => (Math.min(scale.max, Math.max(scale.min, v)) - scale.min) / (scale.max - scale.min)
  // Value-axis gridlines and labels.
  ctx.lineWidth = 1
  for (const t of ticks) {
    const f = vpos(t)
    ctx.strokeStyle = GRID
    if (model.valueAxis.gridlines || t === scale.min) {
      ctx.beginPath()
      if (horizontal) { const x = Math.round(plot.x + f * plot.w) + 0.5; ctx.moveTo(x, plot.y); ctx.lineTo(x, plot.y + plot.h) }
      else { const y = Math.round(plot.y + plot.h - f * plot.h) + 0.5; ctx.moveTo(plot.x, y); ctx.lineTo(plot.x + plot.w, y) }
      ctx.stroke()
    }
    if (model.valueAxis.deleted) continue
    ctx.fillStyle = TEXT
    if (horizontal) { ctx.textAlign = 'center'; ctx.textBaseline = 'top'; ctx.fillText(label(t), plot.x + f * plot.w, plot.y + plot.h + 3) }
    else { ctx.textAlign = 'right'; ctx.textBaseline = 'middle'; ctx.fillText(label(t), plot.x - 4, plot.y + plot.h - f * plot.h) }
  }
  const zero = vpos(Math.max(scale.min, Math.min(scale.max, 0)))
  // Category axis line and labels.
  const band = (horizontal ? plot.h : plot.w) / count
  if (!scatter) {
    ctx.strokeStyle = '#BFBFBF'
    ctx.beginPath()
    if (horizontal) { const x = Math.round(plot.x + zero * plot.w) + 0.5; ctx.moveTo(x, plot.y); ctx.lineTo(x, plot.y + plot.h) }
    else { const y = Math.round(plot.y + plot.h - zero * plot.h) + 0.5; ctx.moveTo(plot.x, y); ctx.lineTo(plot.x + plot.w, y) }
    ctx.stroke()
    if (!model.categoryAxis.deleted && model.categories.length) {
      ctx.fillStyle = TEXT
      const widest = Math.max(...model.categories.map(c => ctx.measureText(c).width))
      const skip = horizontal ? Math.max(1, Math.ceil((fontPx + 2) / band)) : Math.max(1, Math.ceil((widest + 6) / band))
      model.categories.forEach((c, i) => {
        if (i % skip) return
        // Horizontal bars list categories top-down in reverse (Excel's default orientation).
        if (horizontal) { ctx.textAlign = 'right'; ctx.textBaseline = 'middle'; ctx.fillText(c, plot.x - 4, plot.y + plot.h - band * (i + 0.5)) }
        else { ctx.textAlign = 'center'; ctx.textBaseline = 'top'; ctx.fillText(c, plot.x + band * (i + 0.5), plot.y + plot.h + 3, skip * band) }
      })
    }
  } else if (xs && !model.categoryAxis.deleted) {
    ctx.fillStyle = TEXT; ctx.textAlign = 'center'; ctx.textBaseline = 'top'
    for (let v = xs.min, n = 0; v <= xs.max + xs.step / 1e6 && n < 200; v += xs.step, n++) {
      const x = plot.x + (v - xs.min) / (xs.max - xs.min) * plot.w
      ctx.fillText(formatNumber(Number(v.toPrecision(12))), x, plot.y + plot.h + 3)
      if (model.categoryAxis.gridlines) { ctx.strokeStyle = GRID; ctx.beginPath(); ctx.moveTo(Math.round(x) + 0.5, plot.y); ctx.lineTo(Math.round(x) + 0.5, plot.y + plot.h); ctx.stroke() }
    }
  }
  // Axis titles.
  ctx.fillStyle = TEXT
  if (catTitle) { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; if (horizontal) { ctx.save(); ctx.translate(box.x + fontPx, plot.y + plot.h / 2); ctx.rotate(-Math.PI / 2); ctx.fillText(catTitle, 0, 0); ctx.restore() } else ctx.fillText(catTitle, plot.x + plot.w / 2, box.y + box.h) }
  if (valueTitle) { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; if (horizontal) ctx.fillText(valueTitle, plot.x + plot.w / 2, box.y + box.h); else { ctx.save(); ctx.translate(box.x + fontPx, plot.y + plot.h / 2); ctx.rotate(-Math.PI / 2); ctx.fillText(valueTitle, 0, 0); ctx.restore() } }

  ctx.save()
  ctx.beginPath(); ctx.rect(plot.x, plot.y, plot.w, plot.h); ctx.clip()
  const at = (f: number) => horizontal ? plot.x + f * plot.w : plot.y + plot.h - f * plot.h
  const barGroups = model.groups.filter(g => g.type === 'bar')
  for (const group of model.groups) {
    const stacked = stackedValues.get(group)
    const share = (v: number, i: number) => group.grouping === 'percentStacked' && stacked ? (stacked.total[i] ? v / stacked.total[i] : 0) : v
    if (group.type === 'bar') {
      const n = group.grouping === 'clustered' ? Math.max(1, group.series.length) : 1
      const stride = 1 - group.overlap / 100
      const span = 1 + (n - 1) * stride
      const width = band / (span + group.gapWidth / 100)
      const posBase = new Array<number>(count).fill(0), negBase = new Array<number>(count).fill(0)
      group.series.forEach((s, si) => {
        s.values.forEach((raw, i) => {
          if (raw === undefined || i >= count) return
          const v = share(raw, i)
          let from = 0, to = v
          if (stacked) { const base = v >= 0 ? posBase : negBase; from = base[i]; to = base[i] + v; base[i] = to }
          const slot = group.grouping === 'clustered' ? si : 0
          const offset = (band - width * span) / 2 + width * slot * stride
          const a = at(vpos(from)), b = at(vpos(to))
          ctx.fillStyle = `#${s.pointColors?.[i] && barGroups.length === 1 && group.series.length === 1 ? s.pointColors[i] : s.color}`
          if (horizontal) ctx.fillRect(Math.min(a, b), plot.y + plot.h - band * (i + 1) + offset, Math.abs(b - a), width)
          else ctx.fillRect(plot.x + band * i + offset, Math.min(a, b), width, Math.abs(b - a))
        })
      })
    } else if (group.type === 'line' || group.type === 'area') {
      const base = new Array<number>(count).fill(0)
      const edge = group.type === 'area'
      const xAt = (i: number) => edge ? plot.x + (count > 1 ? i * plot.w / (count - 1) : plot.w / 2) : plot.x + band * (i + 0.5)
      const layers: Array<{ s: XlsxChartSeries; top: Array<number | undefined>; bottom: number[] }> = []
      for (const s of group.series) {
        const bottom = base.slice()
        const top = Array.from({ length: count }, (_, i) => {
          const raw = s.values[i]
          if (raw === undefined) return stacked && edge ? base[i] : undefined
          const v = share(raw, i)
          if (stacked) { base[i] += v; return base[i] }
          return v
        })
        layers.push({ s, top, bottom })
      }
      for (const { s, top, bottom } of layers) {
        ctx.strokeStyle = ctx.fillStyle = `#${s.color}`
        if (edge) {
          ctx.beginPath()
          let started = false
          top.forEach((v, i) => { if (v === undefined) return; const x = xAt(i), y = at(vpos(v)); if (!started) { ctx.moveTo(x, y); started = true } else ctx.lineTo(x, y) })
          if (!started) continue
          for (let i = count - 1; i >= 0; i--) if (top[i] !== undefined) ctx.lineTo(xAt(i), at(vpos(stacked ? bottom[i] : Math.max(scale.min, Math.min(scale.max, 0)))))
          ctx.closePath()
          ctx.globalAlpha = stacked ? 1 : 0.85
          ctx.fill()
          ctx.globalAlpha = 1
          continue
        }
        if (s.line !== false) {
          ctx.lineWidth = 2.25; ctx.lineJoin = 'round'
          ctx.beginPath()
          let pen = false
          top.forEach((v, i) => { if (v === undefined) { pen = false; return } const x = xAt(i), y = at(vpos(v)); if (!pen) { ctx.moveTo(x, y); pen = true } else ctx.lineTo(x, y) })
          ctx.stroke()
        }
        if (s.markers) top.forEach((v, i) => { if (v === undefined) return; ctx.beginPath(); ctx.arc(xAt(i), at(vpos(v)), 2.5, 0, Math.PI * 2); ctx.fill() })
      }
    } else if (group.type === 'scatter' && xs) {
      const xp = (v: number) => plot.x + (v - xs!.min) / (xs!.max - xs!.min) * plot.w
      for (const s of group.series) {
        const pts = s.values.map((v, i) => { const x = s.xValues?.length ? s.xValues[i] : i + 1; return v === undefined || x === undefined ? undefined : [xp(x), at(vpos(v))] as const })
        ctx.strokeStyle = ctx.fillStyle = `#${s.color}`
        if (s.line) {
          ctx.lineWidth = 2; ctx.beginPath()
          let pen = false
          for (const p of pts) { if (!p) { pen = false; continue } if (!pen) { ctx.moveTo(p[0], p[1]); pen = true } else ctx.lineTo(p[0], p[1]) }
          ctx.stroke()
        }
        if (s.markers !== false) for (const p of pts) if (p) { ctx.beginPath(); ctx.arc(p[0], p[1], 3, 0, Math.PI * 2); ctx.fill() }
      }
    }
  }
  ctx.restore()
}

/** Paint a chart into a w×h box at the origin of `ctx` (worksheet CSS px). */
export function paintXlsxChart(model: XlsxChartModel, ctx: CanvasRenderingContext2D, w: number, h: number, familyCss: FamilyCss = JSON.stringify): void {
  if (model.unsupported || !model.groups.length) { paintChartPlaceholder(model, ctx, w, h, familyCss); return }
  ctx.save()
  try {
    ctx.beginPath(); ctx.rect(0, 0, w, h); ctx.clip()
    ctx.fillStyle = '#FFFFFF'; ctx.fillRect(0, 0, w, h)
    ctx.strokeStyle = GRID; ctx.lineWidth = 1; ctx.strokeRect(0.5, 0.5, w - 1, h - 1)
    const fontPx = model.fontSizePt * 96 / 72
    const pad = Math.min(10, w * 0.04, h * 0.04)
    let area: Box = { x: pad, y: pad, w: w - pad * 2, h: h - pad * 2 }
    if (model.title) {
      const titlePx = fontPx * 1.4
      ctx.font = font(titlePx, model.fontFamily, familyCss)
      ctx.fillStyle = TEXT; ctx.textAlign = 'center'; ctx.textBaseline = 'top'
      ctx.fillText(model.title, w / 2, area.y, area.w)
      area = { ...area, y: area.y + titlePx + 6, h: area.h - titlePx - 6 }
    }
    area = paintLegend(ctx, model, area, fontPx, model.fontFamily, familyCss)
    if (area.w <= 8 || area.h <= 8) return
    const group = model.groups[0]
    if (group.type === 'pie' || group.type === 'doughnut') paintPie(ctx, model, group, area)
    else paintCartesian(ctx, model, area, fontPx, model.fontFamily, familyCss)
  } finally { ctx.restore() }
}

export { chartDescription }
registerXlsxChartPainter(paintXlsxChart)

/** Resident workbooks: fill every uncached chart reference from parsed cells,
 * after evaluation and semantics. Only the cells on the consumed lines are
 * visited (rows by binary search, cells by binary search), and labels are
 * formatted on demand: the workspace is bounded by the chart point limits. */
export function fillWorkbookCharts<Cell extends { row: number; col: number; value: ChartCellValue }, Sheet extends { name: string; workbookIndex?: number; rows: ReadonlyArray<{ index: number; cells: ReadonlyArray<Cell> }>; drawings?: XlsxDrawingLike[] }>(
  sheets: readonly Sheet[], options: { names?: readonly ChartDefinedName[]; text?: (sheet: Sheet, cell: Cell) => string; available?: (cell: Cell) => boolean } = {},
): void {
  const sheetByName = (name: string) => sheets.find(candidate => candidate.name.toLowerCase() === name.toLowerCase())
  const sortedRows = new WeakMap<object, boolean>()
  const ascending = <T>(items: ReadonlyArray<T>, key: (item: T) => number) => { for (let i = 1; i < items.length; i++) if (key(items[i - 1]) >= key(items[i])) return false; return true }
  const search = <T>(items: ReadonlyArray<T>, key: (item: T) => number, target: number, sorted: boolean): T | undefined => {
    if (!sorted) return items.find(item => key(item) === target)
    let lo = 0, hi = items.length - 1
    while (lo <= hi) { const mid = (lo + hi) >> 1, k = key(items[mid]); if (k === target) return items[mid]; if (k < target) lo = mid + 1; else hi = mid - 1 }
    return undefined
  }
  const isSorted = <T>(owner: object, items: ReadonlyArray<T>, key: (item: T) => number) => { let known = sortedRows.get(owner); if (known === undefined) { known = ascending(items, key); sortedRows.set(owner, known) } return known }
  const lookup: ChartCellLookup = (name, rowIndex, col, wantText) => {
    const sheet = sheetByName(name)
    if (!sheet) return undefined
    const row = search(sheet.rows, row => row.index, rowIndex, isSorted(sheet.rows, sheet.rows, row => row.index))
    const cell = row && search(row.cells, cell => cell.col, col, isSorted(row.cells, row.cells, cell => cell.col))
    if (!cell) return undefined
    if (options.available && !options.available(cell)) return { value: undefined }
    let text: string | undefined
    if (wantText && options.text) { try { text = options.text(sheet, cell) } catch { text = undefined } }
    return { value: cell.value, ...(text !== undefined ? { text } : {}) }
  }
  const resolve = chartNameResolver(options.names ?? [], sheets)
  for (const sheet of sheets) {
    try { fillDrawings(sheet.drawings, sheet.name, lookup, resolve) }
    catch { /* a chart that cannot be filled keeps its placeholder state; the workbook stays usable */ }
  }
}
/** Resolve defined names of a drawing tree in place (paged readers resolve before reading cells). */
export function resolveDrawingNames(drawings: XlsxDrawingLike[] | undefined, ownerSheet: string, resolve: ChartNameResolver | undefined): void {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: XlsxChartModel } | undefined
    if (content?.kind === 'xlsxChart' && content.chart) content.chart = resolveChartNames(content.chart, ownerSheet, resolve)
    if (node.children) resolveDrawingNames(node.children, ownerSheet, resolve)
  }
}
/** Whether a drawing tree names any defined name in its chart references. */
export function drawingsNeedNames(drawings: XlsxDrawingLike[] | undefined): boolean {
  return pendingRanges(drawings, '', [], true).some(range => range.definedName)
}
export interface XlsxDrawingLike { content?: { kind: string }; children?: XlsxDrawingLike[] }
/** Fill uncached chart references of a drawing tree in place. */
export function fillDrawings(drawings: XlsxDrawingLike[] | undefined, ownerSheet: string, lookup: ChartCellLookup, resolve?: ChartNameResolver): void {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: XlsxChartModel } | undefined
    if (content?.kind === 'xlsxChart' && content.chart && chartHasPending(content.chart)) content.chart = fillChartRanges(resolveChartNames(content.chart, ownerSheet, resolve), ownerSheet, lookup)
    if (node.children) fillDrawings(node.children, ownerSheet, lookup, resolve)
  }
}
/** The lines a drawing tree still needs (defined names resolved first), for paged readers. */
export function pendingLines(drawings: XlsxDrawingLike[] | undefined, ownerSheet: string, out: ChartLine[] = []): ChartLine[] {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: XlsxChartModel } | undefined
    const chart = content?.kind === 'xlsxChart' ? content.chart : undefined
    if (chart) {
      const add = (range: XlsxChartRange | undefined, text: boolean) => { if (range && !range.definedName) out.push(chartLine(range, ownerSheet, text)) }
      add(chart.categoryPending, true)
      for (const group of chart.groups) for (const s of group.series) {
        add(s.pending?.values, false); add(s.pending?.xValues, false); add(s.pending?.categories, true)
        if (s.pending?.name) add({ ...s.pending.name, r1: s.pending.name.r0, c1: s.pending.name.c0 }, true)
      }
    }
    if (node.children) pendingLines(node.children, ownerSheet, out)
  }
  return out
}
/** Every reference a drawing tree still needs, for paged readers. */
export function pendingRanges(drawings: XlsxDrawingLike[] | undefined, ownerSheet: string, out: Array<XlsxChartRange & { sheet: string }> = [], withNames = false): Array<XlsxChartRange & { sheet: string }> {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: XlsxChartModel } | undefined
    const chart = content?.kind === 'xlsxChart' ? content.chart : undefined
    if (chart) {
      const add = (range: XlsxChartRange | undefined) => { if (range && (withNames || !range.definedName)) out.push({ ...range, sheet: range.sheet ?? ownerSheet }) }
      add(chart.categoryPending)
      for (const group of chart.groups) for (const s of group.series) { add(s.pending?.values); add(s.pending?.xValues); add(s.pending?.name); add(s.pending?.categories) }
    }
    if (node.children) pendingRanges(node.children, ownerSheet, out, withNames)
  }
  return out
}

/** Charts that still need cells become placeholders with `reason` (a refused read). */
export function refuseCharts(drawings: XlsxDrawingLike[] | undefined, reason: string): void {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: XlsxChartModel } | undefined
    if (content?.kind === 'xlsxChart' && content.chart && chartHasPending(content.chart)) content.chart = { ...content.chart, unsupported: reason }
    if (node.children) refuseCharts(node.children, reason)
  }
}
/** Accounted bytes per chart cell read (map entry, key and value). */
export const CHART_CELL_BYTES = 160

/** Charts with an uncached reference to a sheet `reasonFor` names become placeholders with that reason. */
export function markUnreadableCharts(drawings: XlsxDrawingLike[] | undefined, ownerSheet: string, reasonFor: (sheet: string) => string | undefined): boolean {
  let marked = false
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: XlsxChartModel } | undefined
    if (content?.kind === 'xlsxChart' && content.chart && !content.chart.unsupported) {
      for (const line of pendingLines([node], ownerSheet)) {
        const reason = reasonFor(line.sheet)
        if (reason) { content.chart = { ...content.chart, unsupported: reason }; marked = true; break }
      }
    }
    if (node.children && markUnreadableCharts(node.children, ownerSheet, reasonFor)) marked = true
  }
  return marked
}
