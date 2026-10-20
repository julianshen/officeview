/**
 * Chart part reader (lazy chunk): reads c:chartSpace into the chart model.
 * Moved from `src/xlsx/chart.ts` (A0a) behaviour-identically; the shared
 * default `parseChart` adapter runs it for XLSX, DOCX and PPTX.
 */
import { attrs, getChildren, orderedChildren, textOf, type XmlNode } from '../core/xml'
import { parseDrawingColor, resolveDrawingColor, type ThemeContext } from '../drawing/style'
import type { ChartParseContext } from './context'
import type { ChartManualLayout, ChartTextStyle, ChartAxis, ChartAxisBinding, ChartFamily, ChartGroup, ChartModel, ChartPoints, ChartRange, ChartSeries } from './model'
import { PointWriter, knownError, pointBytes, stringBytes } from './points'
import { ChartRefusal, adoptChartModel, modelOwner } from './own'
import { admitChartModel } from './axisBinding'
import { readLabels } from './labels'
import { prepareChartAxes } from './prepare'
import { CHART_AXIS_UNPLACEABLE, placeAxes } from './axes'
import { hexColor, parseFillStyle } from './fills'
import { parseMarker } from './markers'
import { parseLineStyle, resolveSeriesColor, type ChartPointOverride } from './style'
import { chargeChartColor, chargeChartLine, chargeChartPoint } from './own'

/** Bounds: series per chart, points per series drawn (display, spec 7). Re-exported from the leaf limits module. */
import { CHART_POINT_LIMIT, CHART_LABEL_LIMIT, CHART_SERIES_LIMIT, CHART_ANALYTIC_POINT_LIMIT } from './limits'
export { CHART_SERIES_LIMIT, CHART_POINT_LIMIT, CHART_ANALYTIC_POINT_LIMIT, CHART_LABEL_LIMIT, CHART_TEXT_BYTES_LIMIT } from './limits'

const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]
const val = (node: XmlNode | undefined) => attrs(node).val
const finite = (raw: string | undefined): number | undefined => raw !== undefined && raw.trim() !== '' && Number.isFinite(Number(raw)) ? Number(raw) : undefined

const GROUPS: Record<string, ChartFamily> = {
  barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area',
  pieChart: 'pie', pie3DChart: 'pie', doughnutChart: 'doughnut', ofPieChart: 'ofPie', scatterChart: 'scatter',
}
const OF_PIE_TYPES = ['pie', 'bar'] as const
const OF_PIE_SPLIT_TYPES = ['auto', 'cust', 'percent', 'pos', 'val'] as const
const UNSUPPORTED: Record<string, string> = {
  radarChart: 'Radar chart', bubbleChart: 'Bubble chart', stockChart: 'Stock chart', surfaceChart: 'Surface chart',
  surface3DChart: 'Surface chart',
}

/** ofPie group options (spec 4.1): the secondary plot's shape, the split decision and its size. */
function ofPieOptions(node: XmlNode, charge: (b: number) => void): Partial<ChartGroup> {
  // The authored secondPiePt set (source indices): identity-bound one entry each and bounded by the
  // point limit, as Office applies them to split-like decisions; each surviving index is one charge,
  // the same per-entry rule the dPt loop below follows.
  const secondPiePt: number[] = []
  const seen = new Set<number>()
  for (const pt of getChildren(child(node, 'custSplit'), 'secondPiePt')) {
    const idx = finite(attrs(pt).idx)
    if (idx === undefined || !Number.isInteger(idx) || idx < 0 || idx >= CHART_POINT_LIMIT) continue
    if (seen.has(idx)) continue
    seen.add(idx)
    secondPiePt.push(idx)
    chargeChartPoint(charge)
  }
  const secondPieSize = finite(val(child(node, 'secondPieSize')))
  return {
    ofPieType: oneOf(val(child(node, 'ofPieType')), OF_PIE_TYPES),
    splitType: oneOf(val(child(node, 'splitType')), OF_PIE_SPLIT_TYPES),
    splitPos: finite(val(child(node, 'splitPos'))),
    ...(secondPieSize !== undefined ? { secondPieSize: Math.max(5, Math.min(200, secondPieSize)) } : {}),
    ...(secondPiePt.length ? { secondPiePt } : {}),
  }
}

const hex = hexColor
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

/** Equal saved category caches can share the already-owned strings. Compare the
 * original sparse sequence conservatively, without allocating a second string list. */
function sameCategoryCache(a: XmlNode, b: XmlNode | undefined): boolean {
  const ac = cacheOf(a), bc = cacheOf(b)
  if (!ac || !bc || val(child(ac, 'ptCount')) !== val(child(bc, 'ptCount'))) return false
  const ap = getChildren(ac, 'pt'), bp = getChildren(bc, 'pt')
  return ap.length === bp.length && ap.every((p, i) => attrs(p).idx === attrs(bp[i]).idx && textOf(child(p, 'v')) === textOf(child(bp[i], 'v')))
}
const sameRange = (a: ChartRange | undefined, b: ChartRange | undefined) => a && b
  && a.sheet === b.sheet && a.definedName === b.definedName && a.r0 === b.r0 && a.r1 === b.r1 && a.c0 === b.c0 && a.c1 === b.c1

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
  let length = Math.max(0, Math.min(limit, Number.isSafeInteger(declared) ? declared : 0)), over = declared > limit, formats = Boolean(textOf(child(cache, 'formatCode')))
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
    if (v !== undefined) out.value(index, v, attrs(pt).formatCode ?? textOf(child(cacheOf(node), 'formatCode')))
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
const oneOf = <T extends string>(raw: string | undefined, allowed: readonly T[]): T | undefined => allowed.includes(raw as T) ? raw as T : undefined
const TICK_MARKS = ['cross', 'in', 'out', 'none'] as const
const TIME_UNITS = ['days', 'months', 'years'] as const
/** `c:builtInUnit` divisors and default unit labels (spec 5.1). The labels for hundreds, thousands, millions, billions
 * and trillions are Office's words; the others are unverified against an Office reference. */
const BUILT_IN_UNITS: Record<string, readonly [number, string]> = {
  hundreds: [1e2, 'Hundreds'], thousands: [1e3, 'Thousands'], tenThousands: [1e4, 'Ten Thousands'], hundredThousands: [1e5, 'Hundred Thousands'],
  millions: [1e6, 'Millions'], tenMillions: [1e7, 'Ten Millions'], hundredMillions: [1e8, 'Hundred Millions'], billions: [1e9, 'Billions'], trillions: [1e12, 'Trillions'],
}
function dispUnitsOf(node: XmlNode | undefined): ChartAxis['dispUnits'] {
  if (!node) return undefined
  const builtIn = BUILT_IN_UNITS[val(child(node, 'builtInUnit')) ?? '']
  const custom = finite(val(child(node, 'custUnit')))
  const [divisor, unit] = builtIn ?? (custom !== undefined && custom > 0 ? [custom, `x${custom}`] as const : [undefined, ''])
  if (divisor === undefined) return undefined
  const lbl = child(node, 'dispUnitsLbl')
  const label = lbl ? richText(child(lbl, 'tx')) || unit : undefined
  return { divisor, unit, ...(label ? { label: label.slice(0, CHART_LABEL_LIMIT) } : {}) }
}
function textStyle(node: XmlNode | undefined): ChartTextStyle | undefined {
  if (!node) return undefined
  const body = attrs(child(node, 'bodyPr'))
  const size = finite(attrs(child(child(child(node, 'p'), 'pPr'), 'defRPr')).sz)
  const rot = finite(body.rot)
  return { ...(size !== undefined ? { sizePt: Math.max(6, Math.min(72, size / 100)) } : {}),
    ...(rot !== undefined ? { rotation: rot / 60000 } : {}), ...(oneOf(body.vert, ['horz', 'vert', 'vert270', 'eaVert', 'wordArtVert', 'wordArtVertRtl', 'mongolianVert'] as const) ? { vertical: body.vert } : {}) }
}
function manualLayout(node: XmlNode | undefined): ChartManualLayout | undefined {
  const manual = child(child(node, 'layout'), 'manualLayout')
  if (!manual) return undefined
  const layout: ChartManualLayout = {}
  for (const key of ['x', 'y', 'w', 'h'] as const) {
    const n = finite(val(child(manual, key))); if (n !== undefined) layout[key] = n
    const mode = oneOf(val(child(manual, `${key}Mode`)), ['edge', 'factor'] as const)
    if (mode) layout[`${key}Mode`] = mode
  }
  const target = oneOf(val(child(manual, 'layoutTarget')), ['inner', 'outer'] as const)
  if (target) layout.target = target
  return layout
}
function axisOf(node: XmlNode, kind: ChartAxis['kind'], index: number): ChartAxis {
  const scaling = child(node, 'scaling')
  const numFmt = attrs(child(node, 'numFmt'))
  const formatCode = numFmt.formatCode
  const title = titleOf(child(node, 'title'))
  const crossAx = val(child(node, 'crossAx'))
  const number = (name: string, from: XmlNode | undefined = node) => finite(val(child(from, name)))
  const positive = (name: string) => { const v = number(name); return v !== undefined && v > 0 ? v : undefined }
  const skip = (name: string) => { const v = number(name); return v !== undefined && Number.isInteger(v) && v >= 1 ? v : undefined }
  const fields = {
    position: oneOf(val(child(node, 'axPos')), ['b', 't', 'l', 'r'] as const),
    crosses: oneOf(val(child(node, 'crosses')), ['autoZero', 'min', 'max'] as const),
    crossesAt: number('crossesAt'),
    min: number('min', scaling), max: number('max', scaling),
    majorUnit: positive('majorUnit'), minorUnit: positive('minorUnit'),
    logBase: number('logBase', scaling),
    reversed: val(child(scaling, 'orientation')) === 'maxMin' ? true : undefined,
    crossBetween: oneOf(val(child(node, 'crossBetween')), ['between', 'midCat'] as const),
    majorTickMark: oneOf(val(child(node, 'majorTickMark')), TICK_MARKS),
    minorTickMark: oneOf(val(child(node, 'minorTickMark')), TICK_MARKS),
    tickLblPos: oneOf(val(child(node, 'tickLblPos')), ['nextTo', 'low', 'high', 'none'] as const),
    text: textStyle(child(node, 'txPr')),
    tickLblSkip: skip('tickLblSkip'), tickMarkSkip: skip('tickMarkSkip'),
    dispUnits: dispUnitsOf(child(node, 'dispUnits')),
    baseTimeUnit: oneOf(val(child(node, 'baseTimeUnit')), TIME_UNITS),
    majorTimeUnit: oneOf(val(child(node, 'majorTimeUnit')), TIME_UNITS),
    minorTimeUnit: oneOf(val(child(node, 'minorTimeUnit')), TIME_UNITS),
    auto: child(node, 'auto') ? flag(child(node, 'auto'), true) : undefined,
    minorGridlines: child(node, 'minorGridlines') ? true : undefined,
    sourceLinked: numFmt.sourceLinked !== undefined ? numFmt.sourceLinked === '1' || numFmt.sourceLinked === 'true' : undefined,
  }
  const present = Object.fromEntries(Object.entries(fields).filter(([, v]) => v !== undefined)) as Partial<ChartAxis>
  return {
    // An axis without an axId keeps a positional id.
    id: val(child(node, 'axId')) ?? `#${index}`,
    kind,
    ...(crossAx !== undefined ? { crossAx } : {}),
    ...present,
    gridlines: Boolean(child(node, 'majorGridlines')),
    deleted: val(child(node, 'delete')) === '1' || val(child(node, 'delete')) === 'true',
    ...(title ? { title } : {}),
    ...(formatCode !== undefined ? { formatCode: formatCode.length > 255 ? 'General' : formatCode } : {}),
  }
}

/** The retained size of one axis record: a fixed record (its scalar fields) plus its strings. */
const AXIS_RECORD_BYTES = 256
export const axisBytes = (axis: ChartAxis): number => AXIS_RECORD_BYTES + (axis.text ? 64 : 0) + [axis.id, axis.title, axis.formatCode, axis.dispUnits?.unit, axis.dispUnits?.label].reduce((sum, text) => sum + (text ? stringBytes(text) : 0), 0)
/** An implied axis (A1a): the plot area had no axis of that kind for a group without axIds; added once. */
function impliedAxis(axes: ChartAxis[], kind: ChartAxis['kind'], id: string): string {
  if (!axes.some(axis => axis.id === id)) axes.push({ id, kind, gridlines: false, deleted: false, implicit: true })
  return id
}
const horizontalPosition = (axis: ChartAxis | undefined) => axis?.position === 'b' || axis?.position === 't'
const verticalPosition = (axis: ChartAxis | undefined) => axis?.position === 'l' || axis?.position === 'r'
/**
 * A group's axis binding from its own `c:axId`s (spec 5.1): a category or date axis is the category axis, a value
 * axis the value axis; two value axes are x and y (or x and value), the horizontal one (`axPos` b/t) being x.
 * Unresolved ids are kept so validation names them. A group without axIds, or a plot area without any axis element
 * (never written by Office), binds the plot area's first axes of the right kinds, implied when there are none.
 */
function bindingOf(family: ChartFamily, node: XmlNode, axes: ChartAxis[]): ChartAxisBinding {
  if (family === 'pie' || family === 'doughnut' || family === 'ofPie') return { kind: 'none' }
  const ids = getChildren(node, 'axId').map(val).filter((id): id is string => id !== undefined)
  const axisOf = (id: string) => axes.find(axis => axis.id === id)
  // A plot area without any axis element binds the implied pair even when the group lists axIds.
  if (!ids.length || !axes.some(axis => !axis.implicit)) {
    const values = axes.filter(axis => axis.kind === 'value')
    if (family === 'scatter') {
      const x = values.find(horizontalPosition) ?? values[0], y = values.find(axis => axis !== x)
      return { kind: 'xy', x: x?.id ?? impliedAxis(axes, 'value', '#x'), y: y?.id ?? impliedAxis(axes, 'value', '#y') }
    }
    const category = axes.find(axis => axis.kind === 'category' || axis.kind === 'date')
    return { kind: 'cartesian', category: category?.id ?? impliedAxis(axes, 'category', '#cat'), value: values[0]?.id ?? impliedAxis(axes, 'value', '#val') }
  }
  const valueIds = ids.filter(id => axisOf(id)?.kind === 'value')
  const xFirst = (pair: string[]) => verticalPosition(axisOf(pair[0])) && horizontalPosition(axisOf(pair[1])) ? [pair[1], pair[0]] : pair
  if (family === 'scatter') {
    const [x, y] = valueIds.length >= 2 ? xFirst(valueIds) : [ids[0], ids[1]]
    return { kind: 'xy', x: x ?? '', y: y ?? '' }
  }
  const category = ids.find(id => { const kind = axisOf(id)?.kind; return kind === 'category' || kind === 'date' })
  if (category === undefined && valueIds.length >= 2) { const [x, value] = xFirst(valueIds); return { kind: 'cartesian', x, value } }
  const value = valueIds[0] ?? ids.find(id => id !== category) ?? ''
  return { kind: 'cartesian', category: category ?? ids.find(id => id !== value) ?? '', value }
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
    const model = prepareChartAxes(readChart(root, chart, plot, theme, charge, context), owner)
    if (owner) adoptChartModel(model, owner)
    return model
  } catch (error) {
    owner?.workspace.release(owner.tag)
    if (!(error instanceof ChartRefusal)) throw error
    return { ...placeholderChart(titleOf(child(chart, 'title')), 'This chart needs more memory than is available for charts'), refused: true }
  }
}
/**
 * `c:style` (1-48): a direct child of the chart space, or the usual Office wrapper
 * `mc:AlternateContent` (`c14:style`, which is 100 + the style, or the fallback `c:style`).
 */
function chartStyleId(root: XmlNode): number | undefined {
  const valid = (n: number | undefined) => n !== undefined && Number.isInteger(n) && n >= 1 && n <= 48 ? n : undefined
  const direct = valid(finite(val(child(root, 'style'))))
  if (direct !== undefined) return direct
  for (const alternate of getChildren(root, 'AlternateContent')) {
    const fallback = valid(finite(val(child(child(alternate, 'Fallback'), 'style'))))
    if (fallback !== undefined) return fallback
    for (const choice of getChildren(alternate, 'Choice')) {
      const c14 = finite(val(child(choice, 'style')))
      if (c14 !== undefined && valid(c14 - 100) !== undefined) return c14 - 100
    }
  }
  return undefined
}
/** The c14 `invertSolidFillFmt` colour of a series (the fill of its inverted negative points), from its extension list. */
function invertFill(ser: XmlNode, theme: ThemeContext | undefined): string | undefined {
  for (const ext of getChildren(child(ser, 'extLst'), 'ext')) {
    const fill = child(child(child(ext, 'invertSolidFillFmt'), 'spPr'), 'solidFill')
    if (fill) return hex(resolveDrawingColor(parseDrawingColor(fill), theme))
  }
  return undefined
}
const XLSX_PART = /^\/?xl\//i
function readChart(root: XmlNode, chart: XmlNode, plot: XmlNode, theme: ThemeContext | undefined, charge: (bytes: number) => void, context?: ChartParseContext): ChartModel {
  const minorFont = theme?.fonts.minor.latin ?? 'Calibri'
  const sizeRaw = finite(attrs(getChildren(child(child(root, 'txPr'), 'p'), 'pPr').map(p => child(p, 'defRPr'))[0]).sz)
  const autoDeleted = val(child(chart, 'autoTitleDeleted')) === '1'
  let title = titleOf(child(chart, 'title'))
  // Formatting (spec 5.4, 5.11): every retained fill, line, marker and point override is charged through
  // charts/own.ts before it is built.
  const styleId = chartStyleId(root)
  const parts = context?.styleParts
  const lineOf = (spPr: XmlNode | undefined) => parseLineStyle(child(spPr, 'ln'), theme, charge)
  const chartSpaceSpPr = child(root, 'spPr')
  const chartAreaFill = parseFillStyle(chartSpaceSpPr, theme, charge)
  const chartAreaLine = lineOf(chartSpaceSpPr)
  // A missing chart-area fill: white for an Excel chart; for DOCX and PPTX the c:style default (styles 41-48 are
  // the dark row: dk1), else white (spec 5.4).
  const darkStyle = !chartAreaFill && styleId !== undefined && styleId >= 41 && context?.partPath !== undefined && !XLSX_PART.test(context.partPath)
  const chartAreaDefault = darkStyle ? hex(resolveDrawingColor({ kind: 'scheme', value: 'dk1', transforms: [] }, theme)) ?? '000000' : undefined
  const titleSpPr = child(child(chart, 'title'), 'spPr')
  const titleFill = parseFillStyle(titleSpPr, theme, charge)
  const titleLine = lineOf(titleSpPr)
  const plotSpPr = child(plot, 'spPr')
  const plotAreaFill = parseFillStyle(plotSpPr, theme, charge)
  const plotAreaLine = lineOf(plotSpPr)
  const legend = child(chart, 'legend')
  const legendSpPr = child(legend, 'spPr')
  const legendFill = parseFillStyle(legendSpPr, theme, charge)
  const legendLine = lineOf(legendSpPr)
  // The series count of the whole chart: c:style's single-colour columns spread their shades over it.
  let seriesCount = 0
  for (const [name, node] of orderedChildren(plot)) if (GROUPS[name]) seriesCount += getChildren(node, 'ser').length
  seriesCount = Math.min(seriesCount, CHART_SERIES_LIMIT)
  const groups: ChartGroup[] = []
  let unsupported: string | undefined, truncated = false, seriesTotal = 0
  const cut = { truncated: false }
  let sharedCategoryNode: XmlNode | undefined
  let categories: string[] = [], categoryPending: ChartRange | undefined, categoryRef: ChartRange | undefined, categoryValues: ChartPoints | undefined
  const axes: ChartAxis[] = []
  for (const [name, node] of orderedChildren(plot)) if (AXIS_KINDS[name]) {
    // Each axis record the model retains (its scale, placement and label fields; spec 5.1) is charged before it is kept.
    const axis = axisOf(node, AXIS_KINDS[name], axes.length)
    charge(axisBytes(axis))
    axes.push(axis)
  }
  for (const [name, node] of orderedChildren(plot)) {
    if (UNSUPPORTED[name]) { unsupported ??= `${UNSUPPORTED[name]} is not supported`; continue }
    const type = GROUPS[name]
    if (!type) continue
    const grouping = (val(child(node, 'grouping')) ?? (type === 'bar' ? 'clustered' : 'standard')) as ChartGroup['grouping']
    const scatterStyle = val(child(node, 'scatterStyle')) ?? 'marker'
    const vary = type === 'pie' || type === 'doughnut' || type === 'ofPie' ? val(child(node, 'varyColors')) !== '0' : val(child(node, 'varyColors')) === '1'
    // Drop lines (line and area) and series lines (stacked bars): an element without its own a:ln strokes 1px grey.
    const connector = (name: string) => {
      const holder = child(node, name)
      if (!holder) return undefined
      return lineOf(child(holder, 'spPr')) ?? (chargeChartLine(charge), { noFill: false })
    }
    const dropLines = type === 'line' || type === 'area' ? connector('dropLines') : undefined
    const serLines = type === 'bar' || type === 'ofPie' ? connector('serLines') : undefined
    const series: ChartSeries[] = []
    const binding = bindingOf(type, node, axes)
    // A date axis reads its categories' saved numbers (serials in the chart epoch; spec 5.1).
    const dateAxis = binding.kind === 'cartesian' && axes.find(axis => axis.id === binding.category)?.kind === 'date'
    for (const ser of getChildren(node, 'ser')) {
      if (seriesTotal >= CHART_SERIES_LIMIT) { truncated = true; break }
      // Series indices are unsigned integers; anything else falls back to the series position.
      const rawIndex = finite(val(child(ser, 'idx')))
      const index = rawIndex !== undefined && Number.isSafeInteger(rawIndex) && rawIndex >= 0 ? rawIndex : seriesTotal
      seriesTotal++
      const spPr = child(ser, 'spPr')
      const lineColor = solid(spPr, theme, true)
      const fillColor = solid(spPr, theme)
      // Precedence (spec 5.11): an explicit spPr colour, else the style part, colour part, c:style, accent cycle.
      const color = (type === 'line' || type === 'scatter' ? lineColor ?? fillColor : fillColor ?? lineColor)
        ?? resolveSeriesColor({ index, count: seriesCount, line: type === 'line' || type === 'scatter', styleId, parts, theme })
      const fill = parseFillStyle(spPr, theme, charge)
      const lineStyle = lineOf(spPr)
      const markerStyle = parseMarker(child(ser, 'marker'), theme, charge)
      const invertIfNegative = child(ser, 'invertIfNegative') ? flag(child(ser, 'invertIfNegative'), false) : undefined
      const invertColor = invertIfNegative !== undefined ? invertFill(ser, theme) : undefined
      if (invertColor) chargeChartColor(charge)
      const rawExplosion = finite(val(child(ser, 'explosion')))
      const explosion = rawExplosion !== undefined ? Math.max(0, Math.min(400, rawExplosion)) : undefined
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
      let seriesCategories: string[] | undefined
      if(catNode && (categories.length>0 || categoryPending)) {
        if (hasCache(catNode)) {
          if (!sameCategoryCache(catNode, sharedCategoryNode)) { charge(labelBytes(catNode)); seriesCategories = cacheStrings(catNode, cut).map(v => capLabel(v ?? '', cut)) }
        }
        else { const range = rangeOf(catNode); if (range && !sameRange(range, categoryPending)) pending.categories = range }
      }
      if (catNode && categories.length === 0 && !categoryPending) {
        if (hasCache(catNode)) {
          sharedCategoryNode = catNode
          charge(labelBytes(catNode)); categories = cacheStrings(catNode, cut).map(v => capLabel(v ?? '', cut)); categoryRef = rangeOf(catNode)
          if (dateAxis && (child(child(catNode, 'numRef'), 'numCache') || child(catNode, 'numLit'))) categoryValues = numericCache(catNode, cacheExtent(catNode, CHART_POINT_LIMIT), charge)
        }
        else { const range = rangeOf(catNode); if (range) categoryPending = range }
      }
      // Display points past the bound are dropped at draw time (analytic series keep them for analysis).
      if (values.length > CHART_POINT_LIMIT || (xValues && xValues.length > CHART_POINT_LIMIT)) truncated = true
      const pointColors: string[] = []
      const pointFills: Array<import('./fills').ChartFillStyle | undefined> = []
      const dPtMap = new Map<number, ChartPointOverride>()
      // varyColors: each point takes the automatic colour of its index, by the same precedence as series.
      const pointCount = Math.max(Math.min(values.length, CHART_POINT_LIMIT), 1)
      if (vary) for (let i = 0; i < pointCount && i < 256; i++) pointColors[i] = resolveSeriesColor({ index: i, count: Math.min(pointCount, 256), styleId, parts, theme })
      // The ofPie series' synthetic Other slice is the point `n` (index n = the point count): it takes
      // the automatic colour of its index like any point.
      if (type === 'ofPie' && vary && pointCount < 256) pointColors[pointCount] = resolveSeriesColor({ index: pointCount, count: pointCount, styleId, parts, theme })
      for (const dPt of getChildren(ser, 'dPt')) {
        const i = finite(val(child(dPt, 'idx')))
        if (i === undefined || !Number.isInteger(i) || i < 0 || i >= 256) continue
        chargeChartPoint(charge)
        const ptSpPr = child(dPt, 'spPr')
        const ptFill = parseFillStyle(ptSpPr, theme, charge)
        const ptLine = lineOf(ptSpPr)
        const ptMarker = parseMarker(child(dPt, 'marker'), theme, charge)
        const ptInv = child(dPt, 'invertIfNegative') ? flag(child(dPt, 'invertIfNegative'), false) : undefined
        const ptExp = finite(val(child(dPt, 'explosion')))
        const c = solid(ptSpPr, theme)
        if (c) pointColors[i] = c
        if (ptFill) pointFills[i] = ptFill
        dPtMap.set(i, {
          ...(ptFill ? { fill: ptFill } : {}),
          ...(c ? { color: c } : {}),
          ...(ptLine ? { line: ptLine } : {}),
          ...(ptMarker ? { marker: ptMarker } : {}),
          ...(ptInv !== undefined ? { invertIfNegative: ptInv } : {}),
          ...(ptExp !== undefined ? { explosion: Math.max(0, Math.min(400, ptExp)) } : {}),
        })
      }
      // Required roles with neither saved data nor a reference (spec 6.4): values always; x values for scatter.
      const absentRoles: Array<'values' | 'xValues'> = []
      const sourceless = (node: XmlNode | undefined) => !node || (!hasCache(node) && !formulaOf(node))
      if (sourceless(valueNode)) absentRoles.push('values')
      if (type === 'scatter' && sourceless(xNode)) absentRoles.push('xValues')
      const markerNone = val(child(child(ser, 'marker'), 'symbol')) === 'none'
      const groupMarker = val(child(node, 'marker'))
      series.push({
        name: capLabel(literalName || `Series ${index + 1}`, cut),
        color,
        values,
        ...(seriesCategories ? {categories:seriesCategories} : {}),
        labels: readLabels(child(ser, 'dLbls'), charge),
        ...(xValues ? { xValues } : {}),
        ...(pointColors.length ? { pointColors } : {}),
        ...(fill ? { fill } : {}),
        ...(lineStyle ? { lineStyle } : {}),
        ...(markerStyle ? { markerStyle } : {}),
        ...(invertIfNegative !== undefined ? { invertIfNegative } : {}),
        ...(invertColor ? { invertColor } : {}),
        ...(explosion !== undefined ? { explosion } : {}),
        ...(dPtMap.size > 0 ? { dPt: dPtMap } : {}),
        ...(pointFills.some(Boolean) ? { pointFills } : {}),
        ...(Object.keys(pending).length ? { pending } : {}),
        ...(Object.keys(refs).length ? { refs } : {}),
        ...(analytic ? { analytic } : {}),
        ...(absentRoles.length ? { absent: absentRoles } : {}),
        ...(type === 'line' ? { line: lineColor !== null, markers: !markerNone && groupMarker !== '0' } : {}),
        ...(type === 'scatter' ? { line: lineColor !== null && /line|smooth/i.test(scatterStyle), markers: !markerNone && !/^(line|smooth)$/i.test(scatterStyle) } : {}),
      })
    }
    groups.push({
      family: type, grouping, series, binding,
      labels: readLabels(child(node, 'dLbls'), charge),
      ...(type === 'bar' ? { barDir: val(child(node, 'barDir')) === 'bar' ? 'bar' as const : 'col' as const } : {}),
      gapWidth: Math.max(0, Math.min(500, finite(val(child(node, 'gapWidth'))) ?? 150)),
      overlap: Math.max(-100, Math.min(100, finite(val(child(node, 'overlap'))) ?? (grouping === 'stacked' || grouping === 'percentStacked' ? 100 : 0))),
      ...(type === 'doughnut' ? { holeSize: Math.max(10, Math.min(90, finite(val(child(node, 'holeSize'))) ?? 50)) } : {}),
      ...(type === 'pie' || type === 'doughnut' || type === 'ofPie' ? { firstSliceAngle: finite(val(child(node, 'firstSliceAng'))) ?? 0 } : {}),
      ...(type === 'ofPie' ? ofPieOptions(node, charge) : {}),
      varyColors: vary,
      ...(dropLines ? { dropLines } : {}),
      ...(serLines ? { serLines } : {}),
    })
  }
  // Implied axes (a plot area without axis elements) are retained like authored ones.
  for (const axis of axes) if (axis.implicit) charge(axisBytes(axis))
  const pies = groups.filter(g => g.family === 'pie' || g.family === 'doughnut' || g.family === 'ofPie')
  if (!groups.length) unsupported ??= 'This chart type is not supported'
  else if (!groups.some(g => g.series.length)) unsupported ??= 'This chart has no data series'
  else if (pies.length && groups.length > 1) unsupported ??= 'Combination pie charts are not supported'
  if (child(chart, 'title') && !title && !autoDeleted) title = groups.flatMap(g => g.series).length === 1 && !groups[0].series[0].pending?.name ? groups[0].series[0].name : val(child(root, 'lang'))?.startsWith('zh') ? '圖表標題' : 'Chart Title'
  const blanks = val(child(chart, 'dispBlanksAs'))
  const naAsBlank = dispNaAsBlank(chart)
  const layouts = [plot, child(chart, 'title'), legend]
  for (const node of layouts) if (child(child(node, 'layout'), 'manualLayout')) charge(128)
  if (child(root, 'txPr')) charge(64)
  const plotLayout = manualLayout(plot), titleLayout = manualLayout(child(chart, 'title')), legendLayout = manualLayout(legend), text = textStyle(child(root, 'txPr'))
  const model: ChartModel = {
    ...(title ? { title } : {}),
    groups,
    axes,
    categories,
    ...(categoryValues ? { categoryValues } : {}),
    ...(categoryPending ? { categoryPending } : {}),
    ...(categoryRef ? { categoryRef } : {}),
    ...(legend ? { legend: (val(child(legend, 'legendPos')) ?? 'r') as ChartModel['legend'] } : {}),
    // Chart-level options as authored; absent stays undefined (Office defaults: gap, plot visible only).
    ...(child(chart, 'dispBlanksAs') ? { dispBlanksAs: blanks === 'gap' || blanks === 'span' ? blanks : 'zero' } : {}),
    ...(child(chart, 'plotVisOnly') ? { plotVisOnly: flag(child(chart, 'plotVisOnly'), true) } : {}),
    ...(naAsBlank !== undefined ? { dispNaAsBlank: naAsBlank } : {}),
    ...(child(root, 'date1904') ? { epoch: flag(child(root, 'date1904'), true) ? '1904' as const : '1900' as const } : {}),
    ...(plotLayout ? { plotLayout } : {}),
    ...(titleLayout ? { titleLayout } : {}),
    ...(legendLayout ? { legendLayout } : {}),
    ...(text ? { text } : {}),
    ...(chartAreaFill ? { chartAreaFill } : {}),
    ...(chartAreaLine ? { chartAreaLine } : {}),
    ...(plotAreaFill ? { plotAreaFill } : {}),
    ...(plotAreaLine ? { plotAreaLine } : {}),
    ...(legendFill ? { legendFill } : {}),
    ...(legendLine ? { legendLine } : {}),
    ...(titleFill ? { titleFill } : {}),
    ...(titleLine ? { titleLine } : {}),
    ...(styleId !== undefined ? { styleId } : {}),
    ...(chartAreaDefault ? { chartAreaDefault } : {}),
    fontFamily: minorFont,
    fontSizePt: sizeRaw !== undefined ? Math.max(6, Math.min(72, sizeRaw / 100)) : 10,
    ...(unsupported ? { unsupported } : {}),
    ...(truncated || cut.truncated ? { truncated: true } : {}),
  }
  return unsupported ? model : admitAxes(model, charge)
}

/**
 * Axis admission (spec 5.1): every group binding must name existing axes of the right kinds, and the groups must be
 * placeable together; otherwise the chart is a card with `chart-axis-unplaceable` (never drawn on a guessed axis).
 */
function admitAxes(model: ChartModel, charge: (bytes: number) => void): ChartModel {
  const admitted = admitChartModel(model)
  const placed = admitted.kind === 'native' ? placeAxes(model) : undefined
  const reason = admitted.kind !== 'native' ? admitted.reason : placed && !placed.ok ? placed.reason : undefined
  if (reason === undefined) return model
  charge(2 * stringBytes(reason))
  return { ...model, unsupported: `This chart's axes cannot be placed (${reason})`, diagnostics: [...model.diagnostics ?? [], { code: CHART_AXIS_UNPLACEABLE, reason }] }
}

/** A placeholder model for a chart that could not be read. */
export function placeholderChart(title: string | undefined, reason: string): ChartModel {
  return { ...(title ? { title } : {}), groups: [], axes: [], categories: [], fontFamily: 'Calibri', fontSizePt: 10, unsupported: reason }
}
