/**
 * Chart part reader (lazy chunk): reads c:chartSpace into the chart model.
 * Moved from `src/xlsx/chart.ts` (A0a) behaviour-identically; the shared
 * default `parseChart` adapter runs it for XLSX, DOCX and PPTX.
 */
import { attrs, getChildren, orderedChildren, textOf, type XmlNode } from '../core/xml'
import { parseDrawingColor, resolveDrawingColor, type ThemeContext } from '../drawing/style'
import type { ChartParseContext } from './context'
import type { ChartDiagnostic, ChartManualLayout, ChartTextStyle, ChartAxis, ChartAxisBinding, ChartFamily, ChartGroup, ChartModel, ChartPoints, ChartRange, ChartSeries } from './model'
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
import { readCategoryLevels } from './categories'
import { CHART_TITLE_LINE_LIMIT, CHART_TITLE_RUN_LIMIT, CHART_TITLE_TEXT_LIMIT, CHART_TITLE_TRUNCATED, runsText, titleRuns } from './titles'
import { markTrendlineLegendDeletions, readTrendlines } from './trendlines'
import { readErrorBars } from './errorBars'
import { contourGridDiagnostic, contourGridOf } from './families/contour'
import { readBubbleOptions } from './families/bubble'
import { readRadarStyle } from './families/radar'
import { readStockOptions } from './families/stock'

/** Bounds: series per chart, points per series drawn (display, spec 7). Re-exported from the leaf limits module. */
import { CHART_POINT_LIMIT, CHART_LABEL_LIMIT, CHART_SERIES_LIMIT, CHART_ANALYTIC_POINT_LIMIT, XLSX_CHART_PART } from './limits'
export { CHART_SERIES_LIMIT, CHART_POINT_LIMIT, CHART_ANALYTIC_POINT_LIMIT, CHART_LABEL_LIMIT, CHART_TEXT_BYTES_LIMIT } from './limits'

const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]

/**
 * Contour band formats (`c:bandFmts/c:bandFmt`, CT_BandFmt: a child `c:idx` and `c:spPr`), by band index: the
 * solid fill and line colours, null for noFill. Each retained colour is charged before it is kept.
 */
function parseBandFormats(node: XmlNode, theme: ThemeContext | undefined, charge: (bytes: number) => void): { bandColors?: Array<string | null>; bandLines?: Array<string | null> } {
  const fills: Array<string | null> = [], lines: Array<string | null> = []
  for (const bf of getChildren(child(node, 'bandFmts'), 'bandFmt')) {
    const idx = finite(val(child(bf, 'idx')))
    if (idx === undefined || !Number.isInteger(idx) || idx < 0 || idx >= 256) continue
    const spPr = child(bf, 'spPr')
    const fill = solid(spPr, theme), line = solid(spPr, theme, true)
    if (fill !== undefined) { chargeChartColor(charge); fills[idx] = fill }
    if (line !== undefined) { chargeChartColor(charge); lines[idx] = line }
  }
  return { ...(fills.length ? { bandColors: fills } : {}), ...(lines.length ? { bandLines: lines } : {}) }
}
const val = (node: XmlNode | undefined) => attrs(node).val
const finite = (raw: string | undefined): number | undefined => raw !== undefined && raw.trim() !== '' && Number.isFinite(Number(raw)) ? Number(raw) : undefined

const GROUPS: Record<string, ChartFamily> = {
  barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area',
  pieChart: 'pie', pie3DChart: 'pie', doughnutChart: 'doughnut', ofPieChart: 'ofPie', scatterChart: 'scatter',
  bubbleChart: 'bubble', radarChart: 'radar', surfaceChart: 'contour', stockChart: 'stock',
}
const OF_PIE_TYPES = ['pie', 'bar'] as const
const OF_PIE_SPLIT_TYPES = ['auto', 'cust', 'percent', 'pos', 'val'] as const
/** Chart types the engine does not draw, with the card's reason (spec 4.1; surface3D: the spec's card text). ofPie now draws natively. */
const UNSUPPORTED: Record<string, string> = {
  surface3DChart: '3-D surface charts are shown as a placeholder',
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
/** The explicit solid colour of an `spPr` fill (or its `a:ln` when `line`): `null` for noFill. */
export function solid(spPr: XmlNode | undefined, theme: ThemeContext | undefined, line = false): string | undefined | null {
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
/** The declared point count of a saved source; a multi-level cache declares it once, on the cache (spec 5.9). */
const declaredCount = (node: XmlNode | undefined, cache: XmlNode | undefined) =>
  finite(val(child(cache, 'ptCount') ?? child(child(child(node, 'multiLvlStrRef'), 'multiLvlStrCache'), 'ptCount')))
/**
 * The saved points of a data source by `pt@idx`, sparse indices kept: a missing
 * index (or ptCount padding) is `missing` at its position, never compacted.
 * `cut.truncated` records points dropped by the per-series bound (declared count or indices past it).
 */
function cachePoints<T>(node: XmlNode | undefined, read: (pt: XmlNode) => T, missing: T, cut?: { truncated: boolean }, limit = CHART_POINT_LIMIT): T[] {
  const cache = cacheOf(node)
  const out: T[] = []
  const declared = declaredCount(node, cache) ?? 0
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
  const declared = declaredCount(node, cache) ?? 0
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
/** Placement text properties of a DrawingML `txPr` (size, rotation, vertical). */
export function textStyle(node: XmlNode | undefined): ChartTextStyle | undefined {
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
    noMultiLvlLbl: kind === 'category' && child(node, 'noMultiLvlLbl') && flag(child(node, 'noMultiLvlLbl'), true) ? true as const : undefined,
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
  const xy = family === 'scatter' || family === 'bubble'
  // A plot area without any axis element binds the implied pair even when the group lists axIds.
  if (!ids.length || !axes.some(axis => !axis.implicit)) {
    const values = axes.filter(axis => axis.kind === 'value')
    if (xy) {
      const x = values.find(horizontalPosition) ?? values[0], y = values.find(axis => axis !== x)
      return { kind: 'xy', x: x?.id ?? impliedAxis(axes, 'value', '#x'), y: y?.id ?? impliedAxis(axes, 'value', '#y') }
    }
    if (family === 'contour') {
      const category = axes.find(axis => axis.kind === 'category' || axis.kind === 'date')
      const series = axes.find(axis => axis.kind === 'series')
      return { kind: 'surface', category: category?.id ?? impliedAxis(axes, 'category', '#cat'), series: series?.id ?? impliedAxis(axes, 'series', '#ser'), value: values[0]?.id ?? impliedAxis(axes, 'value', '#val') }
    }
    const category = axes.find(axis => axis.kind === 'category' || axis.kind === 'date')
    if (family === 'radar') return { kind: 'radial', value: values[0]?.id ?? impliedAxis(axes, 'value', '#val'), ...(category ? { category: category.id } : {}) }
    return { kind: 'cartesian', category: category?.id ?? impliedAxis(axes, 'category', '#cat'), value: values[0]?.id ?? impliedAxis(axes, 'value', '#val') }
  }
  const valueIds = ids.filter(id => axisOf(id)?.kind === 'value')
  const xFirst = (pair: string[]) => verticalPosition(axisOf(pair[0])) && horizontalPosition(axisOf(pair[1])) ? [pair[1], pair[0]] : pair
  if (xy) {
    const [x, y] = valueIds.length >= 2 ? xFirst(valueIds) : [ids[0], ids[1]]
    return { kind: 'xy', x: x ?? '', y: y ?? '' }
  }
  if (family === 'contour') {
    const category = ids.find(id => { const kind = axisOf(id)?.kind; return kind === 'category' || kind === 'date' })
    const series = ids.find(id => axisOf(id)?.kind === 'series')
    const value = valueIds[0] ?? ids.find(id => id !== category && id !== series) ?? ''
    return { kind: 'surface', category: category ?? '', series: series ?? '', value }
  }
  const category = ids.find(id => { const kind = axisOf(id)?.kind; return kind === 'category' || kind === 'date' })
  if (family === 'radar') return { kind: 'radial', value: valueIds[0] ?? ids.find(id => id !== category) ?? '', ...(category !== undefined ? { category } : {}) }
  if (category === undefined && valueIds.length >= 2) { const [x, value] = xFirst(valueIds); return { kind: 'cartesian', x, value } }
  const value = valueIds[0] ?? ids.find(id => id !== category) ?? ''
  return { kind: 'cartesian', category: category ?? ids.find(id => id !== value) ?? '', value }
}
/** A series feeds an analytic element: a trendline, or stdDev/stdErr error bars (spec 7). */
const statisticErrorBars = (ser: XmlNode) => getChildren(ser, 'errBars').some(bars => /^std(Dev|Err)$/.test(val(child(bars, 'errValType')) ?? ''))
const isAnalytic = (ser: XmlNode) => Boolean(child(ser, 'trendline')) || statisticErrorBars(ser)
const flag = (node: XmlNode | undefined, fallback: boolean) => { const v = val(node); return v === undefined ? fallback : v === '1' || v === 'true' }
const readSmooth = (el: XmlNode | undefined) => el ? flag(el, true) : undefined
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
const XLSX_PART = XLSX_CHART_PART
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
  const authoredSeries: number[] = []
  const cut = { truncated: false }
  let sharedCategoryNode: XmlNode | undefined
  let categoryLevels: ChartModel['categoryLevels'], categoryMultiLevel = false
  let categories: string[] = [], categoryPending: ChartRange | undefined, categoryRef: ChartRange | undefined, categoryValues: ChartPoints | undefined
  const axes: ChartAxis[] = []
  for (const [name, node] of orderedChildren(plot)) if (AXIS_KINDS[name]) {
    // Each axis record the model retains (its scale, placement and label fields; spec 5.1) is charged before it is kept.
    const axis = axisOf(node, AXIS_KINDS[name], axes.length)
    charge(axisBytes(axis))
    axes.push(axis)
  }
  for (const [name, node] of orderedChildren(plot)) {
    if (UNSUPPORTED[name]) { unsupported ??= UNSUPPORTED[name]; continue }
    const type = GROUPS[name]
    if (!type) continue
    const grouping = (val(child(node, 'grouping')) ?? (type === 'bar' ? 'clustered' : 'standard')) as ChartGroup['grouping']
    // A 100% stacked statistic normalises by every sibling's value at each index: all series of such a group are
    // analytic dependencies, acquired in full (spec 7) with the same refusal and unavailable rules.
    const stackedStatistic = grouping === 'percentStacked' && getChildren(node, 'ser').some(statisticErrorBars)
    const scatterStyle = val(child(node, 'scatterStyle')) ?? 'marker'
    const radarStyle = type === 'radar' ? readRadarStyle(node) : undefined
    const xy = type === 'scatter' || type === 'bubble'
    const vary = type === 'pie' || type === 'doughnut' || type === 'ofPie' ? val(child(node, 'varyColors')) !== '0' : val(child(node, 'varyColors')) === '1'
    const groupSmooth = readSmooth(child(node, 'smooth')) ?? (type === 'scatter' && /smooth/i.test(scatterStyle) ? true : undefined)
    const stockOpts = type === 'stock' ? readStockOptions(node, theme, charge) : undefined
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
      const lineLike = type === 'line' || type === 'scatter' || (type === 'radar' && radarStyle !== 'filled')
      const color = (lineLike ? lineColor ?? fillColor : fillColor ?? lineColor)
        ?? resolveSeriesColor({ index, count: seriesCount, line: lineLike, styleId, parts, theme })
      const fill = parseFillStyle(spPr, theme, charge)
      const lineStyle = lineOf(spPr)
      const markerStyle = parseMarker(child(ser, 'marker'), theme, charge)
      const trendlines = readTrendlines(ser, lineOf, charge)
      const invertIfNegative = child(ser, 'invertIfNegative') ? flag(child(ser, 'invertIfNegative'), false) : undefined
      const invertColor = invertIfNegative !== undefined ? invertFill(ser, theme) : undefined
      if (invertColor) chargeChartColor(charge)
      const rawExplosion = finite(val(child(ser, 'explosion')))
      const serSmooth = readSmooth(child(ser, 'smooth'))
      const smooth = serSmooth ?? groupSmooth
      const explosion = rawExplosion !== undefined ? Math.max(0, Math.min(400, rawExplosion)) : undefined
      const tx = child(ser, 'tx')
      const valueNode = child(ser, xy ? 'yVal' : 'val')
      const xNode = xy ? child(ser, 'xVal') : undefined
      const sizeNode = type === 'bubble' ? child(ser, 'bubbleSize') : undefined
      const catNode = xy ? undefined : child(ser, 'cat')
      const pending: NonNullable<ChartSeries['pending']> = {}, refs: NonNullable<ChartSeries['refs']> = {}
      const rangeOf = (node: XmlNode | undefined) => { const f = formulaOf(node); return f ? parseChartRange(f) : undefined }
      // Error bars (spec 5.7): a custom amount is a numeric role, its saved data read at the display bound (custom
      // amounts are drawn per displayed point, never analysed); without saved data its reference is pending.
      const errors = readErrorBars(ser, lineOf, charge, node => hasCache(node) ? { points: numericCache(node, { ...cacheExtent(node, CHART_POINT_LIMIT), formats: false }, charge) } : { range: rangeOf(node) })
      if (errors) Object.assign(pending, errors.pending)
      // An uncached reference the viewer cannot read (multi-area, a formula) has no data: placeholder, never an empty chart.
      const unreadable = (node: XmlNode | undefined) => { const f = !hasCache(node) ? formulaOf(node) : undefined; if (f && !parseChartRange(f)) unsupported ??= `Chart data uses a reference the viewer cannot read (${f.slice(0, 80)})` }
      unreadable(valueNode); unreadable(xNode); unreadable(sizeNode); unreadable(catNode)
      const txFormula = child(tx, 'strRef') && !child(child(tx, 'strRef'), 'strCache') ? formulaOf(tx) : undefined
      const nameRange = txFormula ? parseChartRange(txFormula) : undefined
      if (nameRange) pending.name = nameRange
      const literalName = textOf(child(tx, 'v')) || cacheStrings(tx)[0]
      // Analytic inputs are acquired in full (up to CHART_ANALYTIC_POINT_LIMIT) before any display cap;
      // more than that refuses the analytic element (spec 7), never a prefix: the series is then drawn capped.
      // Extents are read first, so nothing is built before the decision and its admission.
      let analytic: ChartSeries['analytic'] = stackedStatistic || isAnalytic(ser) ? true : undefined
      const limit = analytic ? CHART_ANALYTIC_POINT_LIMIT : CHART_POINT_LIMIT
      let valueExtent = cacheExtent(valueNode, limit), xExtent = xNode ? cacheExtent(xNode, limit) : undefined
      if (analytic && (valueExtent.over || xExtent?.over)) {
        analytic = 'refused'
        valueExtent = cacheExtent(valueNode, CHART_POINT_LIMIT); xExtent = xNode ? cacheExtent(xNode, CHART_POINT_LIMIT) : undefined
        cut.truncated = true
      } else if (valueExtent.over || xExtent?.over) cut.truncated = true
      const values = numericCache(valueNode, valueExtent, charge)
      const xValues = xNode && xExtent ? numericCache(xNode, xExtent, charge) : undefined
      // Bubble sizes are a display-only numeric role, admitted like the others before they are built.
      const sizeExtent = sizeNode ? cacheExtent(sizeNode, CHART_POINT_LIMIT) : undefined
      if (sizeExtent?.over) cut.truncated = true
      const bubbleSizes = sizeNode && sizeExtent ? numericCache(sizeNode, sizeExtent, charge) : undefined
      // A role's reference is kept beside its cache: the cache gives the values, the reference the provenance.
      if (!hasCache(valueNode)) { const range = rangeOf(valueNode); if (range) pending.values = range } else { const range = rangeOf(valueNode); if (range) refs.values = range }
      if (xNode) { const range = rangeOf(xNode); if (range) { if (hasCache(xNode)) refs.xValues = range; else pending.xValues = range } }
      if (sizeNode) { const range = rangeOf(sizeNode); if (range) { if (hasCache(sizeNode)) refs.bubbleSizes = range; else pending.bubbleSizes = range } }
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
          // Outer levels of a multi-level cache (spec 5.9), charged before they are kept.
          categoryLevels = readCategoryLevels(catNode, categories.length, charge, cut, CHART_LABEL_LIMIT)
          if (dateAxis && (child(child(catNode, 'numRef'), 'numCache') || child(catNode, 'numLit'))) categoryValues = numericCache(catNode, cacheExtent(catNode, CHART_POINT_LIMIT), charge)
        }
        else { const range = rangeOf(catNode); if (range) { categoryPending = range; if (child(catNode, 'multiLvlStrRef')) categoryMultiLevel = true } }
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
      // Required roles with neither saved data nor a reference (spec 6.4): values always; x values for scatter and
      // bubble; bubble sizes for bubble.
      const absentRoles: Array<'values' | 'xValues' | 'bubbleSizes'> = []
      const sourceless = (node: XmlNode | undefined) => !node || (!hasCache(node) && !formulaOf(node))
      if (sourceless(valueNode)) absentRoles.push('values')
      if (xy && sourceless(xNode)) absentRoles.push('xValues')
      if (type === 'bubble' && sourceless(sizeNode)) absentRoles.push('bubbleSizes')
      const markerNone = val(child(child(ser, 'marker'), 'symbol')) === 'none'
      const groupMarker = val(child(node, 'marker'))
      series.push({
        index,
        name: capLabel(literalName || `Series ${index + 1}`, cut),
        color,
        values,
        ...(seriesCategories ? {categories:seriesCategories} : {}),
        labels: type === 'contour' ? undefined : readLabels(child(ser, 'dLbls'), charge),
        ...(xValues ? { xValues } : {}),
        ...(bubbleSizes ? { bubbleSizes } : {}),
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
        ...(trendlines ? { trendlines } : {}),
        ...(errors?.bars.length ? { errorBars: errors.bars } : {}),
        ...(absentRoles.length ? { absent: absentRoles } : {}),
        ...(smooth !== undefined ? { smooth } : {}),
        ...(type === 'line' ? { line: lineColor !== null, markers: !markerNone && groupMarker !== '0' } : {}),
        ...(type === 'scatter' ? { line: lineColor !== null && /line|smooth/i.test(scatterStyle), markers: !markerNone && !/^(line|smooth)$/i.test(scatterStyle) } : {}),
        ...(type === 'radar' ? { line: lineColor !== null, markers: radarStyle === 'marker' && !markerNone } : {}),
        ...(type === 'stock' ? { markers: !markerNone && Boolean(child(ser, 'marker')) } : {}),
      })
    }
    // The contour bands' default colours: the accent cycle of the document's theme (and style/colour parts), charged.
    const bandAccents = () => Array.from({ length: 6 }, (_, index) => (chargeChartColor(charge), resolveSeriesColor({ index, count: 6, styleId, parts, theme })))
    authoredSeries.push(getChildren(node, 'ser').length)
    if (type === 'stock' && series.length !== 3 && series.length !== 4) {
      unsupported ??= `Stock charts require 3 (high, low, close) or 4 (open, high, low, close) series, got ${series.length}`
    }
    groups.push({
      family: type, grouping, series, binding,
      // CT_SurfaceChart and CT_SurfaceSer have no data labels.
      labels: type === 'contour' ? undefined : readLabels(child(node, 'dLbls'), charge),
      ...(type === 'bar' ? { barDir: val(child(node, 'barDir')) === 'bar' ? 'bar' as const : 'col' as const } : {}),
      gapWidth: Math.max(0, Math.min(500, stockOpts?.gapWidth ?? finite(val(child(node, 'gapWidth'))) ?? 150)),
      overlap: Math.max(-100, Math.min(100, finite(val(child(node, 'overlap'))) ?? (grouping === 'stacked' || grouping === 'percentStacked' ? 100 : 0))),
      ...(type === 'doughnut' ? { holeSize: Math.max(10, Math.min(90, finite(val(child(node, 'holeSize'))) ?? 50)) } : {}),
      ...(type === 'pie' || type === 'doughnut' || type === 'ofPie' ? { firstSliceAngle: finite(val(child(node, 'firstSliceAng'))) ?? 0 } : {}),
      ...(type === 'ofPie' ? ofPieOptions(node, charge) : {}),
      ...(type === 'contour' ? { wireframe: child(node, 'wireframe') ? flag(child(node, 'wireframe'), true) : false, ...parseBandFormats(node, theme, charge), bandAccents: bandAccents() } : {}),
      varyColors: vary,
      ...(groupSmooth !== undefined ? { smooth: groupSmooth } : {}),
      ...(type === 'bubble' ? readBubbleOptions(node) : {}),
      ...(radarStyle ? { radarStyle } : {}),
      ...(dropLines ? { dropLines } : {}),
      ...(serLines ? { serLines } : {}),
      ...(stockOpts?.hiLowLines ? { hiLowLines: stockOpts.hiLowLines } : {}),
      ...(stockOpts?.upDownBars ? { upDownBars: stockOpts.upDownBars } : {}),
    })
  }
  const allSeries = groups.flatMap(g => g.series), pointLegend = allSeries.length === 1 && groups.find(g => g.series.length === 1)?.varyColors
  // Deleted legend entries (c:legendEntry/c:delete), read once, bounded and charged before the set is built.
  const deleted = new Set<number>()
  for (const entry of getChildren(legend, 'legendEntry')) {
    const idx = finite(val(child(entry, 'idx')))
    // An entry index is a series c:idx (unsigned int) or a point index; the number of entries kept is bounded.
    if (deleted.size >= CHART_POINT_LIMIT) break
    if (child(entry, 'delete') && flag(child(entry, 'delete'), true) && idx !== undefined && Number.isInteger(idx) && idx >= 0 && idx <= 0xFFFFFFFF) { if (!deleted.has(idx)) { if (!deleted.size) charge(64); charge(4) } deleted.add(idx) }
  }
  // Trendline entry indices never coincide with another entry's key: after the highest series c:idx, or after every
  // point key of a point legend (source indices, uncapped: only the display keeps 64; Office's numbering is unverified).
  markTrendlineLegendDeletions(deleted, groups, pointLegend ? Math.max(categories.length, allSeries[0].values.length) : allSeries.reduce((n, s) => Math.max(n, s.index ?? 0), -1) + 1)
  // Implied axes (a plot area without axis elements) are retained like authored ones.
  for (const axis of axes) if (axis.implicit) charge(axisBytes(axis))
  const pies = groups.filter(g => g.family === 'pie' || g.family === 'doughnut' || g.family === 'ofPie')
  // The contour grid bound (spec 9): rows from the authored series count, columns from the categories.
  const contourClips = groups.flatMap((g, i) => g.family === 'contour' ? [contourGridOf({ categories }, g, authoredSeries[i])] : []).filter(grid => grid.clipped)
  if (contourClips.length) truncated = true
  if (!groups.length) unsupported ??= 'This chart type is not supported'
  else if (!groups.some(g => g.series.length)) unsupported ??= 'This chart has no data series'
  else if (pies.length && groups.length > 1) unsupported ??= 'Combination pie charts are not supported'
  else if (groups.some(g => g.family === 'radar') && groups.some(g => g.family !== 'radar')) unsupported ??= 'Combination radar charts are not supported'
  if (child(chart, 'title') && !title && !autoDeleted) title = groups.flatMap(g => g.series).length === 1 && !groups[0].series[0].pending?.name ? groups[0].series[0].name : val(child(root, 'lang'))?.startsWith('zh') ? '圖表標題' : 'Chart Title'
  // Title runs (formatting precedence: rPr, paragraph, txPr, style part, theme) and overlays; legend entries deleted.
  const titleNode = child(chart, 'title'), titleDiagnostics: ChartDiagnostic[] = []
  const resolved = title ? titleRuns(titleNode, title, theme, parts?.style?.title, charge) : undefined
  const runs = resolved?.runs
  // Runs cut by the title bounds: the title shown (painted and described) is what the runs keep, and it is reported.
  if (resolved?.truncated) {
    title = runsText(resolved.runs).replace(/\n/g, ' ').trim() || title
    truncated = true
    const reason = `The title keeps ${CHART_TITLE_RUN_LIMIT} runs, ${CHART_TITLE_LINE_LIMIT} lines and ${CHART_TITLE_TEXT_LIMIT} characters`
    charge(64 + stringBytes(CHART_TITLE_TRUNCATED) + stringBytes(reason))
    titleDiagnostics.push({ code: CHART_TITLE_TRUNCATED, reason })
  }
  const overlay = (node: XmlNode | undefined) => Boolean(child(node, 'overlay')) && flag(child(node, 'overlay'), true)
  const legendDeleted = deleted.size ? Uint32Array.from([...deleted].sort((a, b) => a - b)) : undefined
  const blanks = val(child(chart, 'dispBlanksAs'))
  const naAsBlank = dispNaAsBlank(chart)
  const layouts = [plot, child(chart, 'title'), legend]
  for (const node of layouts) if (child(child(node, 'layout'), 'manualLayout')) charge(128)
  if (child(root, 'txPr')) charge(64)
  const plotLayout = manualLayout(plot), titleLayout = manualLayout(child(chart, 'title')), legendLayout = manualLayout(legend), text = textStyle(child(root, 'txPr'))
  const model: ChartModel = {
    ...(title ? { title } : {}),
    ...(runs ? { titleRuns: runs } : {}),
    ...(title && overlay(titleNode) ? { titleOverlay: true as const } : {}),
    ...(legend && overlay(legend) ? { legendOverlay: true as const } : {}),
    ...(legendDeleted ? { legendDeleted } : {}),
    ...(categoryLevels ? { categoryLevels } : {}),
    groups,
    axes,
    categories,
    ...(categoryValues ? { categoryValues } : {}),
    ...(categoryPending ? { categoryPending } : {}),
    ...(categoryPending && categoryMultiLevel ? { categoryMultiLevel: true as const } : {}),
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
    ...(titleDiagnostics.length || contourClips.length ? { diagnostics: [...titleDiagnostics, ...contourClips.map(grid => { const d = contourGridDiagnostic(grid); charge(stringBytes(d.code) + stringBytes(d.reason)); return d })] } : {}),
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
