/**
 * The ChartEx (`cx:`, Office 2016+) part parser (spec 4.2, 4.3; lazy chunk), normative against
 * docs/charts/chartex-map.md. It produces the engine's one `ChartModel`, whose single group is a
 * discriminated `ChartExGroup` (charts/model.ts), with everything the family painters (B2-B5) need:
 *
 * - `cx:chartData/cx:data[@id]` sets are keyed by id; only the sets a series binds are acquired. Each
 *   dimension keeps its `@type`, its reference and its `cx:lvl` (sparse `pt@idx`; a missing index is blank).
 * - A series binds its data set through `cx:dataId`; its roles are chosen by dimension type (values: val, size,
 *   colorVal; categories: cat, entityId; the first of a type wins, a repeat is diagnosed). A `paretoLine` series
 *   has no data: it binds its owner through `ownerIdx` and is kept beside the group's series.
 * - Saved data (a single-level `cx:lvl` cache) is used as read. Without one (or with a multi-level cache, whose
 *   level order is unconfirmed), the role is a pending reference: XLSX resolves its hidden `_xlchart.*` defined
 *   name against the workbook (the XLSX fill); DOCX/PPTX read the embedded workbook (charts/dataPrecedence.ts).
 * - Analytic families (histogram, Pareto, box & whisker, treemap, sunburst) acquire every point up to 65,536 on
 *   every path (`analytic: 'chart'`); above that, or with a required point unavailable, the chart is refused.
 *   Waterfall and funnel are display families: the 4,096-point display cap applies (truncated).
 * - Styling is parsed and admitted now: chart-space, plot-area, plot-surface, title and legend `spPr`; chart,
 *   title, axis and label `txPr`; series `spPr`; `dataPt` overrides; data-label settings.
 * - Authored binning is validated here (map §5.1 step 7): invalid parameters refuse the chart.
 * - `layoutId` selects the family (map §3). Region map is recognised first and is always unsupported; its
 *   geography (`cx:binary`) is never read.
 *
 * Until a family's painter lands (B2-B5) the adapter answers `unsupported` for it, so an mc:Fallback is kept
 * (spec 6.3). Every retained allocation, text included, is bounded and charged to the document's chart workspace
 * before it is built (charts/own.ts); a refusal stops the parse and refunds the charge. Labels share one
 * chart-wide 4 MiB budget (spec 7); a label cut to 255 units or dropped past the budget marks the chart truncated
 * and is diagnosed (one bounded diagnostic per kind of cut, with its count).
 */
import { attrs, getChildren, orderedChildren, type XmlNode } from '../../core/xml'
import type { ThemeContext } from '../../drawing/style'
import type { ChartParseContext } from '../context'
import type {
  ChartAxis, ChartAxisBinding, ChartDiagnostic, ChartExBinning, ChartExData, ChartExDimension, ChartExFamily, ChartExGroup, ChartExLayout,
  ChartExLayoutProperties, ChartExLevel, ChartExSeriesInfo, ChartGroup, ChartModel, ChartPoints, ChartResult, ChartSeries, ChartTextStyle, ChartTickMark,
} from '../model'
import type { LabelControls, LabelOptions } from '../labels'
import type { ChartFillStyle } from '../fills'
import type { ChartLineStyle, ChartPointOverride } from '../style'
import { ChartRefusal, adoptChartModel, chargeChartPoint, chartModelOwner, modelOwner, releaseChartModel, succeedChartModel } from '../own'
import { CHART_ANALYTIC_POINT_LIMIT, CHART_LABEL_LIMIT, CHART_POINT_LIMIT, CHART_SERIES_LIMIT, CHART_TEXT_BYTES_LIMIT, axisBytes, parseChartRange, placeholderChart, solid, textStyle } from '../parse'
import { PointWriter, knownError, pointBytes, stringBytes } from '../points'
import { parseFillStyle } from '../fills'
import { parseLineStyle, resolveSeriesColor } from '../style'
import { readChartStyleParts } from '../styleParts'
import { applyDataPrecedence } from '../dataPrecedence'
import { CHART_DATA_NOT_AVAILABLE, CHART_TOO_MANY_POINTS, chartHasPending, chartHasRefs } from '../fill'
import { prepareChartAxes } from '../prepare'
import { displayFamilyUnsupported } from './funnel'
import { HISTOGRAM_FAMILIES, histogramStructure } from './histogram'

/** Families whose painter has landed. Every other family stays `unsupported` (its fallback is kept). */
export const CHARTEX_PAINTED: ReadonlySet<ChartExFamily> = new Set<ChartExFamily>(['waterfall', 'funnel', 'treemap', 'sunburst', 'boxWhisker', 'histogram', 'pareto'])
export const CHARTEX_PENDING = 'ChartEx charts are not supported yet'
export const CHARTEX_REGION_MAP = 'Region map charts are not supported'
export const CHARTEX_UNREADABLE = 'This chart could not be read'
export const CHARTEX_MULTI_LEVEL = 'multi-level cache order unconfirmed'
export const CHARTEX_TOO_MANY = CHART_TOO_MANY_POINTS
/** Map §5.1 step 7: every invalid authored binning refuses the chart with this reason. */
export const CHARTEX_INVALID_BINNING = 'invalid binning'
/** Map §5.1: `underflow`/`overflow="auto"` has unconfirmed semantics, so the chart is unsupported. */
export const CHARTEX_AUTO_BINNING = 'Automatic histogram overflow and underflow bins are not supported'
export const CHARTEX_BIN_LIMIT = 1000
const ANALYTIC: ReadonlySet<ChartExFamily> = new Set<ChartExFamily>(['histogram', 'pareto', 'boxWhisker', 'treemap', 'sunburst'])
/** Treemap and sunburst: a category path of every level (chartex-map §7), coloured by root-level branch. */
const HIERARCHY: ReadonlySet<ChartExFamily> = new Set<ChartExFamily>(['treemap', 'sunburst'])
/** Root-level branch colours resolved at parse (the palette cycles past them) [U: Office's branch colouring]. */
const HIERARCHY_PALETTE = 64
const LAYOUTS: readonly ChartExLayout[] = ['clusteredColumn', 'paretoLine', 'waterfall', 'funnel', 'boxWhisker', 'treemap', 'sunburst', 'regionMap']
const VALUE_TYPES = ['val', 'size', 'colorVal'], CATEGORY_TYPES = ['cat', 'entityId']
const LABEL_POSITIONS = ['bestFit', 'b', 'ctr', 'inBase', 'inEnd', 'l', 'outEnd', 'r', 't']
/** Retained record prices, charged before each record is built. */
const DATA_BYTES = 64, DIMENSION_BYTES = 128, LEVEL_BYTES = 96, SERIES_BYTES = 384, SLOT_BYTES = 8, TEXT_STYLE_BYTES = 96, LABEL_RECORD_BYTES = 256
/** The admission bound of one built label (CHART_LABEL_LIMIT + 1 units: one past the limit marks the cut). */
const LABEL_BOUND = stringBytes(' '.repeat(CHART_LABEL_LIMIT + 1))
const XLSX_PART = /^\/?xl\//i

export const isChartExGroup = (group: ChartGroup): group is ChartExGroup => group.kind === 'chartEx'

/** A structural failure: the chart is unsupported (or refused) with `message` (never a thrown parse error). */
class Unsupported extends Error { constructor(message: string, readonly refused = false) { super(message) } }
const child = (node: XmlNode | undefined, name: string): XmlNode | undefined => getChildren(node, name)[0]
/** A node's own text (already in the admitted tree): never a concatenation of descendants. */
const own = (node: XmlNode | undefined): string => { const text = node?.['#text']; return typeof text === 'string' ? text : '' }
const uint = (raw: string | undefined): number | undefined => raw !== undefined && /^\s*\d{1,15}\s*$/.test(raw) && Number.isSafeInteger(Number(raw)) ? Number(raw) : undefined
const finite = (raw: string | undefined): number | undefined => raw !== undefined && raw.trim() !== '' && Number.isFinite(Number(raw)) ? Number(raw) : undefined
const bool = (raw: string | undefined, fallback: boolean): boolean => raw === '1' || raw === 'true' ? true : raw === '0' || raw === 'false' ? false : fallback
/** Office writes some schema text elements as `@val` (map §12.7): `@val` first, then the text. */
const valOrText = (node: XmlNode | undefined): string | undefined => node ? attrs(node).val ?? own(node) : undefined

/** The family of the plot area's series (map §3), or the reason it has none. Cheap: no data, no text. */
export function chartExFamily(root: XmlNode): { family: ChartExFamily; byCategory?: true } | { unsupported: string } {
  if (child(root, 'dashboard') || !child(root, 'chart')) return { unsupported: CHARTEX_UNREADABLE }
  const nodes = getChildren(child(child(child(root, 'chart'), 'plotArea'), 'plotAreaRegion'), 'series')
  const layouts = nodes.map(node => attrs(node).layoutId as ChartExLayout)
  if (layouts.includes('regionMap')) return { unsupported: CHARTEX_REGION_MAP }
  if (!nodes.length || layouts.some(layout => !LAYOUTS.includes(layout))) return { unsupported: CHARTEX_UNREADABLE }
  const primary = layouts.filter(layout => layout !== 'paretoLine')
  if (!primary.length || primary.some(layout => layout !== primary[0])) return { unsupported: CHARTEX_UNREADABLE }
  const lines = nodes.filter(node => attrs(node).layoutId === 'paretoLine')
  // A paretoLine is valid only with ownerIdx naming a clusteredColumn series (document order, 0-based).
  if (lines.some(line => { const owner = uint(attrs(line).ownerIdx); return owner === undefined || attrs(nodes[owner]).layoutId !== 'clusteredColumn' })) return { unsupported: CHARTEX_UNREADABLE }
  const layout = primary[0]
  if (layout === 'clusteredColumn') {
    // Every column series carries the same discriminator: binning, or aggregation (By Category). Neither, or a
    // mix, is unsupported (map §3).
    const kinds = nodes.filter(node => attrs(node).layoutId === 'clusteredColumn').map(node => {
      const layoutPr = child(node, 'layoutPr')
      return child(layoutPr, 'binning') ? 'binning' : child(layoutPr, 'aggregation') ? 'aggregation' : undefined
    })
    if (kinds.some(kind => kind !== kinds[0]) || !kinds[0]) return { unsupported: CHARTEX_UNREADABLE }
    return { family: lines.length ? 'pareto' : 'histogram', ...(kinds[0] === 'aggregation' ? { byCategory: true as const } : {}) }
  }
  return { family: layout as ChartExFamily }
}

/** The parse's admission: workspace charge, the chart-wide label budget and the cut/drop counts. */
interface Admission {
  charge: (bytes: number) => void
  /** Admit one label of `length` units (from a string already in the tree): charges it, or drops it past the budget. */
  label: (length: number) => 'keep' | 'drop'
  cut: number
  dropped: number
  /** Display points dropped by the display cap (waterfall, funnel). */
  capped: boolean
}
function admission(charge: (bytes: number) => void): Admission {
  let used = 0
  const out: Admission = {
    charge, cut: 0, dropped: 0, capped: false,
    label: length => {
      const bytes = 32 + 2 * Math.min(length, CHART_LABEL_LIMIT)
      if (used + bytes > CHART_TEXT_BYTES_LIMIT) { out.dropped++; return 'drop' }
      charge(bytes)
      used += bytes
      if (length > CHART_LABEL_LIMIT) out.cut++
      return 'keep'
    },
  }
  return out
}
/**
 * Bounded text of runs (`a:p/a:r|a:fld/a:t`): the bound is charged (and stays charged) before anything is
 * built, then at most
 * CHART_LABEL_LIMIT + 1 units are gathered from the tree's own strings; the kept text joins the label budget.
 */
function boundedRuns(paragraphs: XmlNode[], admit: Admission): string | undefined {
  admit.charge(LABEL_BOUND)
  // `room` is what the bound still admits; every separator and run part is counted against it before it is
  // appended, the walk stops when it is spent, and a slice end is never negative.
  const CAP = CHART_LABEL_LIMIT + 1
  let text = '', length = 0
  outer: for (const [index, p] of paragraphs.entries()) {
    if (index && text) {
      length++
      if (text.length >= CAP) break
      text += '\n'
    }
    for (const [name, run] of orderedChildren(p)) {
      if (name !== 'r' && name !== 'fld') continue
      for (const t of getChildren(run, 't')) {
        const part = own(t)
        length += part.length
        const room = CAP - text.length
        if (room <= 0) break outer
        text += part.length > room ? part.slice(0, room) : part
        if (text.length >= CAP) break outer
      }
    }
  }
  return keepLabel(text.trim(), Math.max(length, text.length), admit)
}
/** A label from one tree string (`txData/v`, a `pt`, a name): its bound is decided before the slice is built. */
function boundedString(source: string, admit: Admission): string | undefined {
  admit.charge(LABEL_BOUND)
  return keepLabel(source.length > CHART_LABEL_LIMIT ? source.slice(0, CHART_LABEL_LIMIT + 1) : source, source.length, admit)
}
function keepLabel(text: string, length: number, admit: Admission): string | undefined {
  if (!text) return undefined
  if (admit.label(length) === 'drop') return undefined
  return text.length > CHART_LABEL_LIMIT ? text.slice(0, CHART_LABEL_LIMIT) : text
}
/** Title text (map §1): `tx/txData/v`, the `tx/rich` text, else the auto title's `txPr` run text. */
function titleText(title: XmlNode | undefined, admit: Admission): string | undefined {
  if (!title) return undefined
  const tx = child(title, 'tx')
  const v = own(child(child(tx, 'txData'), 'v')).trim()
  if (v) return boundedString(v, admit)
  const rich = getChildren(child(tx, 'rich'), 'p')
  return (rich.length ? boundedRuns(rich, admit) : undefined) ?? boundedRuns(getChildren(child(title, 'txPr'), 'p'), admit)
}
/** A `txPr`'s text style: placement (classic `textStyle`) plus the defRPr (or first run) colour, bold and italic. */
function textOf(txPr: XmlNode | undefined, theme: ThemeContext | undefined, admit: Admission): ChartTextStyle | undefined {
  if (!txPr) return undefined
  admit.charge(TEXT_STYLE_BYTES)
  const p = child(txPr, 'p')
  // Each property from the first run's rPr (the text as drawn), else the paragraph default (defRPr).
  const def = child(child(p, 'pPr'), 'defRPr'), run = child(child(p, 'r'), 'rPr')
  const a = { ...attrs(def), ...attrs(run) }, color = solid(run, theme) ?? solid(def, theme)
  const size = finite(a.sz)
  const style: ChartTextStyle = { ...textStyle(txPr), ...(size !== undefined ? { sizePt: Math.max(6, Math.min(72, size / 100)) } : {}), ...(color ? { color } : {}), ...(a.b ? { bold: bool(a.b, false) } : {}), ...(a.i ? { italic: bool(a.i, false) } : {}) }
  return Object.keys(style).length ? style : undefined
}

/** Read one ChartEx part into an owned model (a placeholder carrying `feature: 'ChartEx'` when it cannot be read). */
export function parseChartEx(root: XmlNode, theme: ThemeContext | undefined, context: ChartParseContext): ChartModel {
  const owner = modelOwner(context.workspace)
  const charge = (bytes: number) => { if (!owner.workspace.reserve(bytes, owner.tag)) throw new ChartRefusal() }
  const admit = admission(charge)
  let title: string | undefined
  /** A placeholder keeps (and owns) only its title. */
  const fail = (reason: string, refused = false): ChartModel => {
    owner.workspace.release(owner.tag)
    const kept = title && owner.workspace.reserve(stringBytes(title), owner.tag) ? title : undefined
    const model: ChartModel = { ...placeholderChart(kept, reason), feature: 'ChartEx', ...(refused ? { refused: true as const } : {}) }
    adoptChartModel(model, owner)
    return model
  }
  try {
    title = titleText(child(child(root, 'chart'), 'title'), admit)
    const selected = chartExFamily(root)
    if ('unsupported' in selected) return fail(selected.unsupported)
    const model = readChartEx(root, title, selected.family, selected.byCategory, theme, admit, context)
    adoptChartModel(model, owner)
    return model
  } catch (error) {
    if (error instanceof ChartRefusal) { title = undefined; return fail('This chart needs more memory than is available for charts', true) }
    if (error instanceof Unsupported) return fail(error.message, error.refused)
    owner.workspace.release(owner.tag)
    throw error
  }
}

function readChartEx(root: XmlNode, title: string | undefined, family: ChartExFamily, byCategory: true | undefined,
  theme: ThemeContext | undefined, admit: Admission, context: ChartParseContext): ChartModel {
  const { charge } = admit
  const chart = child(root, 'chart')!
  const plot = child(chart, 'plotArea')
  const region = child(plot, 'plotAreaRegion')
  const nodes = getChildren(region, 'series')
  if (nodes.length > CHART_SERIES_LIMIT) throw new Unsupported(CHARTEX_UNREADABLE)
  const analytic = ANALYTIC.has(family)
  const diagnostics: ChartDiagnostic[] = []
  const diagnose = (code: string, reason: string) => { if (diagnostics.some(d => d.code === code)) return; charge(stringBytes(code) + stringBytes(reason)); diagnostics.push({ code, reason }) }
  const fill = (spPr: XmlNode | undefined): ChartFillStyle | undefined => parseFillStyle(spPr, theme, charge)
  const line = (spPr: XmlNode | undefined): ChartLineStyle | undefined => parseLineStyle(child(spPr, 'ln'), theme, charge)

  // Axes (map §1): the kind is the scaling child, never the id or the order. Each record is priced first.
  const axes: ChartAxis[] = []
  for (const node of getChildren(plot, 'axis')) {
    const a = attrs(node), cat = child(node, 'catScaling'), val = child(node, 'valScaling')
    if (a.id === undefined || !cat === !val) throw new Unsupported(CHARTEX_UNREADABLE)
    const scaling = attrs(val), numFmt = attrs(child(node, 'numFmt')), tick = attrs(child(node, 'majorTickMarks')).type
    const unit = attrs(child(node, 'units')).unit
    const formatCode = numFmt.formatCode && numFmt.formatCode.length <= 255 ? numFmt.formatCode : undefined
    charge(axisBytes({ id: a.id, kind: 'value', deleted: false, gridlines: false, ...(formatCode ? { formatCode } : {}) }) + (unit ? stringBytes(unit.slice(0, 32)) : 0))
    const axisTitle = titleText(child(node, 'title'), admit)
    const text = textOf(child(node, 'txPr'), theme, admit)
    const titleStyle = textOf(child(child(node, 'title'), 'txPr'), theme, admit)
    axes.push({
      id: a.id, kind: cat ? 'category' : 'value', deleted: bool(a.hidden, false),
      gridlines: !!child(node, 'majorGridlines'), ...(child(node, 'minorGridlines') ? { minorGridlines: true } : {}),
      ...numberField('min', scaling.min), ...numberField('max', scaling.max), ...numberField('majorUnit', scaling.majorUnit), ...numberField('minorUnit', scaling.minorUnit),
      ...(axisTitle ? { title: axisTitle } : {}), ...(formatCode ? { formatCode, sourceLinked: bool(numFmt.sourceLinked, false) } : {}),
      ...(['cross', 'in', 'out', 'none'].includes(tick ?? '') ? { majorTickMark: tick as ChartTickMark } : {}),
      ...(child(node, 'tickLabels') ? {} : { tickLblPos: 'none' as const }), ...(unit ? { unit: unit.slice(0, 32) } : {}), ...(text ? { text } : {}), ...(titleStyle ? { titleText: titleStyle } : {}),
    })
  }
  const firstCategory = axes.find(axis => axis.kind === 'category'), firstValue = axes.find(axis => axis.kind === 'value')

  // Series, as authored. The paretoLine is kept apart: it has no data of its own.
  const dataSets = new Map<number, XmlNode>()
  for (const data of getChildren(child(root, 'chartData'), 'data')) {
    const id = uint(attrs(data).id)
    if (id === undefined || dataSets.has(id)) throw new Unsupported(CHARTEX_UNREADABLE)
    dataSets.set(id, data)
  }
  const acquired = new Map<number, ChartExData>()
  const parts = context.styleParts
  const series: ChartSeries[] = [], infos: ChartExSeriesInfo[] = []
  let paretoLine: ChartExSeriesInfo | undefined, gapWidth: number | undefined
  nodes.forEach((node, position) => {
    const a = attrs(node)
    const layout = a.layoutId as ChartExLayout
    const axisNodes = getChildren(node, 'axisId')
    charge(SERIES_BYTES + axisNodes.length * (SLOT_BYTES + stringBytes(' '.repeat(16))))
    const axisIds = axisNodes.map(valOrText).filter((id): id is string => !!id).map(id => id.slice(0, 16))
    const dataId = uint(valOrText(child(node, 'dataId')))
    const info: ChartExSeriesInfo = {
      layout, ...(dataId !== undefined ? { dataId } : {}),
      ...(uint(a.ownerIdx) !== undefined ? { ownerIdx: uint(a.ownerIdx) } : {}), hidden: bool(a.hidden, false),
      formatIdx: uint(a.formatIdx) ?? position, axisIds, layoutPr: layoutProperties(child(node, 'layoutPr'), family, charge), position,
    }
    if (layout === 'paretoLine') {
      // The line's colour and authored stroke (B3, map §5.2): it is drawn from its owner's data.
      const spPr = child(node, 'spPr'), lineStyle = line(spPr)
      paretoLine = { ...info, color: solid(spPr, theme) ?? resolveSeriesColor({ index: info.formatIdx, count: nodes.length, ...(parts ? { parts } : {}), theme }), ...(lineStyle ? { lineStyle } : {}) }
      return
    }
    if (info.dataId === undefined || !dataSets.has(info.dataId)) throw new Unsupported(CHARTEX_UNREADABLE)
    let data = acquired.get(info.dataId)
    if (!data) { data = readData(info.dataId, dataSets.get(info.dataId)!, analytic, admit, diagnose); acquired.set(info.dataId, data) }
    const tx = child(node, 'tx')
    const v = own(child(child(tx, 'txData'), 'v')).trim()
    const name = (v ? boundedString(v, admit) : boundedRuns(getChildren(child(tx, 'rich'), 'p'), admit)) ?? defaultName(series.length + 1, charge)
    const spPr = child(node, 'spPr')
    const explicit = solid(spPr, theme)
    const color = explicit ?? resolveSeriesColor({ index: info.formatIdx, count: nodes.length, ...(parts ? { parts } : {}), theme })
    series.push({
      ...seriesRoles(data, name, color, admit, HIERARCHY.has(family)),
      ...styleOf(node, spPr, fill, line, theme, admit, analytic ? CHART_ANALYTIC_POINT_LIMIT : CHART_POINT_LIMIT),
      ...(analytic ? { analytic: 'chart' as const } : {}),
    })
    infos.push(info)
    if (gapWidth === undefined) {
      const category = axes.find(axis => axis.kind === 'category' && axisIds.includes(axis.id)) ?? firstCategory
      gapWidth = finite(attrs(child(getChildren(plot, 'axis').find(node => attrs(node).id === category?.id), 'catScaling')).gapWidth)
    }
  })

  // Waterfall and funnel draw one series: several visible ones are unsupported (map §9; B2).
  const several = displayFamilyUnsupported(family, infos)
  if (several) throw new Unsupported(several)
  // Waterfall: the increase, decrease and total colours are the style palette's first three entries (map §6 [U]).
  let palette: string[] | undefined
  if (family === 'waterfall') {
    charge(3 * (SLOT_BYTES + stringBytes('000000')))
    palette = [0, 1, 2].map(index => resolveSeriesColor({ index, count: 3, ...(parts ? { parts } : {}), theme }))
  }
  const bound = infos[0]?.axisIds ?? []
  const category = axes.find(axis => axis.kind === 'category' && bound.includes(axis.id)) ?? firstCategory
  const value = axes.find(axis => axis.kind === 'value' && bound.includes(axis.id)) ?? firstValue
  const binding: ChartAxisBinding = category && value ? { kind: 'cartesian', category: category.id, value: value.id } : { kind: 'none' }
  const plotSpPr = child(plot, 'spPr'), plotAreaFill = fill(plotSpPr), plotAreaLine = line(plotSpPr)
  // Treemap/sunburst branches take the series colour cycle by root-level ordinal (style and colour parts, theme).
  if (HIERARCHY.has(family)) {
    charge(HIERARCHY_PALETTE * (SLOT_BYTES + stringBytes('000000')))
    palette = Array.from({ length: HIERARCHY_PALETTE }, (_, index) => resolveSeriesColor({ index, count: HIERARCHY_PALETTE, ...(parts ? { parts } : {}), theme }))
  }
  const group: ChartExGroup = {
    kind: 'chartEx', family, grouping: 'standard', series, binding, gapWidth: gapWidth !== undefined ? Math.round(gapWidth * 100) : 150, overlap: 0,
    chartEx: {
      data: [...acquired.values()], series: infos, ...(paretoLine ? { paretoLine } : {}), ...(byCategory ? { byCategory } : {}), ...(gapWidth !== undefined ? { gapWidth } : {}),
      ...(plotAreaFill ? { plotAreaFill } : {}), ...(plotAreaLine ? { plotAreaLine } : {}), ...(palette ? { palette } : {}),
    },
  }
  const legend = child(chart, 'legend'), legendPos = attrs(legend).pos ?? 'r'
  const surface = child(child(region, 'plotSurface'), 'spPr'), titleNode = child(chart, 'title'), space = child(root, 'spPr')
  const styling = {
    chartAreaFill: fill(space), chartAreaLine: line(space), plotAreaFill: fill(surface), plotAreaLine: line(surface),
    titleFill: fill(child(titleNode, 'spPr')), titleLine: line(child(titleNode, 'spPr')), legendFill: fill(child(legend, 'spPr')), legendLine: line(child(legend, 'spPr')),
    text: textOf(child(root, 'txPr'), theme, admit), titleText: textOf(child(titleNode, 'txPr'), theme, admit),
  }
  if (admit.cut) diagnose('chart-label-truncated', `${admit.cut} ${admit.cut === 1 ? 'label was' : 'labels were'} cut to ${CHART_LABEL_LIMIT} characters`)
  if (admit.dropped) diagnose('chart-label-budget', `${admit.dropped} ${admit.dropped === 1 ? 'label was' : 'labels were'} dropped past the chart's 4 MiB label budget`)
  const truncated = admit.cut > 0 || admit.dropped > 0 || admit.capped
  // A value role with neither saved data nor a reference: an analytic chart's data is not available (refused).
  const unavailable = series.some(s => s.absent?.includes('values'))
  if (unavailable && XLSX_PART.test(context.partPath)) throw new Unsupported(CHART_DATA_NOT_AVAILABLE, true)
  return {
    ...(title ? { title } : {}), groups: [group], axes, categories: series[0]?.categories ?? [],
    ...(legend && ['t', 'b', 'l', 'r'].includes(legendPos) ? { legend: legendPos as ChartModel['legend'] } : {}),
    fontFamily: theme?.fonts.minor.latin ?? 'Calibri', fontSizePt: styling.text?.sizePt ?? 10,
    ...Object.fromEntries(Object.entries(styling).filter(([, v]) => v !== undefined)),
    ...(truncated ? { truncated: true } : {}), ...(diagnostics.length ? { diagnostics } : {}),
  }
}
const defaultName = (n: number, charge: (bytes: number) => void): string => { const name = `Series ${n}`; charge(stringBytes(name)); return name }
const numberField = <K extends string>(key: K, raw: string | undefined): Partial<Record<K, number>> => {
  const value = finite(raw)
  return value === undefined ? {} : { [key]: value } as Record<K, number>
}

/** Series styling (map §4.2): `spPr` fill and line, `dataPt` overrides, data labels. */
/** `limit`: the source-index domain of overrides (the analytic input bound, or the display bound). Indices past the
 * display bound are kept in the sparse `dPt` map and label `points` record only; the dense arrays stay display-sized. */
function styleOf(node: XmlNode, spPr: XmlNode | undefined, fill: (spPr: XmlNode | undefined) => ChartFillStyle | undefined,
  line: (spPr: XmlNode | undefined) => ChartLineStyle | undefined, theme: ThemeContext | undefined, admit: Admission, limit: number): Partial<ChartSeries> {
  const out: Partial<ChartSeries> = {}
  const seriesFill = fill(spPr), seriesLine = line(spPr)
  if (seriesFill) out.fill = seriesFill
  if (seriesLine) out.lineStyle = seriesLine
  const dPt = new Map<number, ChartPointOverride>(), pointFills: Array<ChartFillStyle | undefined> = [], pointColors: string[] = []
  for (const point of getChildren(node, 'dataPt')) {
    const i = uint(attrs(point).idx)
    if (i === undefined || i >= limit || dPt.has(i)) continue
    chargeChartPoint(admit.charge)
    const ptSpPr = child(point, 'spPr'), ptFill = fill(ptSpPr), ptLine = line(ptSpPr), color = solid(ptSpPr, theme)
    if (ptFill && i < CHART_POINT_LIMIT) pointFills[i] = ptFill
    if (color && i < CHART_POINT_LIMIT) pointColors[i] = color
    // An override with nothing the model draws (effects only, as FN1) is not kept.
    if (ptFill || color || ptLine) dPt.set(i, { ...(ptFill ? { fill: ptFill } : {}), ...(color ? { color } : {}), ...(ptLine ? { line: ptLine } : {}) })
  }
  if (dPt.size) out.dPt = dPt
  if (pointFills.some(Boolean)) out.pointFills = pointFills
  if (pointColors.some(Boolean)) out.pointColors = pointColors
  const labels = dataLabels(child(node, 'dataLabels'), fill, line, theme, admit, limit)
  if (labels) out.labels = labels
  return out
}
/**
 * ChartEx data labels (map §4.2) as the shared label controls: `visibility` fields (an absent attribute is off at
 * the series level and inherited per point), `pos` (absent: the family default at paint), `separator`, `numFmt`,
 * `spPr`/`txPr`; `dataLabel@idx` overrides field by field; `dataLabelHidden@idx` deletes that point's label.
 */
function dataLabels(node: XmlNode | undefined, fill: (spPr: XmlNode | undefined) => ChartFillStyle | undefined,
  line: (spPr: XmlNode | undefined) => ChartLineStyle | undefined, theme: ThemeContext | undefined, admit: Admission, limit: number): LabelControls | undefined {
  if (!node) return undefined
  const options = (n: XmlNode, series: boolean): LabelOptions => {
    admit.charge(LABEL_RECORD_BYTES)
    const out: LabelOptions = {}
    const v = child(n, 'visibility'), va = attrs(v)
    if (v || series) {
      for (const [key, attr] of [['showSerName', 'seriesName'], ['showCatName', 'categoryName'], ['showVal', 'value']] as const) {
        if (va[attr] !== undefined) out[key] = bool(va[attr], false)
        else if (series) out[key] = false
      }
    }
    const pos = attrs(n).pos
    if (pos && LABEL_POSITIONS.includes(pos)) out.position = pos
    const separator = child(n, 'separator')
    if (separator) { const text = own(separator); admit.charge(stringBytes(text.slice(0, 16))); out.separator = text.slice(0, 16) }
    const fmt = attrs(child(n, 'numFmt'))
    if (fmt.formatCode !== undefined) { const code = fmt.formatCode.length > 255 ? 'General' : fmt.formatCode; admit.charge(stringBytes(code)); out.formatCode = code }
    if (fmt.sourceLinked !== undefined) out.sourceLinked = bool(fmt.sourceLinked, false)
    const spPr = child(n, 'spPr'), labelFill = fill(spPr), labelLine = line(spPr), text = textOf(child(n, 'txPr'), theme, admit)
    if (labelFill) out.fill = labelFill
    if (labelLine) out.line = labelLine
    if (text) out.text = text
    return out
  }
  const out: LabelControls = options(node, true)
  const points: Record<number, LabelOptions> = {}
  for (const p of getChildren(node, 'dataLabel')) {
    const i = uint(attrs(p).idx)
    if (i !== undefined && i < limit) points[i] = options(p, false)
  }
  for (const p of getChildren(node, 'dataLabelHidden')) {
    const i = uint(attrs(p).idx)
    if (i !== undefined && i < limit) { admit.charge(LABEL_RECORD_BYTES); points[i] = { delete: true } }
  }
  if (Object.keys(points).length) out.points = points
  return out
}

/** `cx:layoutPr` with the map's absence defaults (§4.1 visibility, §5.1 intervalClosed r, §8 exclusive); invalid
 * authored binning refuses the chart (§5.1 step 7). */
function layoutProperties(node: XmlNode | undefined, family: ChartExFamily, charge: (bytes: number) => void): ChartExLayoutProperties {
  const v = attrs(child(node, 'visibility'))
  const out: ChartExLayoutProperties = {
    visibility: { connectorLines: bool(v.connectorLines, true), meanLine: bool(v.meanLine, false), meanMarker: bool(v.meanMarker, true), nonoutliers: bool(v.nonoutliers, false), outliers: bool(v.outliers, true) },
  }
  const binning = child(node, 'binning')
  if (binning && (family === 'histogram' || family === 'pareto')) out.binning = binningOf(binning)
  else if (child(node, 'aggregation')) out.aggregation = true
  const subtotals = child(node, 'subtotals')
  if (subtotals) {
    const idx = getChildren(subtotals, 'idx')
    // The kept list plus the transient index list and set (charged up front; the transients stay charged).
    charge(idx.length * 6 * SLOT_BYTES)
    // Sorted and distinct, within the display bound (the waterfall walks them beside its slots).
    out.subtotals = [...new Set(idx.map(item => uint(valOrText(item))).filter((i): i is number => i !== undefined && i < CHART_POINT_LIMIT))].sort((a, b) => a - b)
  }
  const statistics = attrs(child(node, 'statistics')).quartileMethod
  if (child(node, 'statistics')) out.quartileMethod = statistics === 'inclusive' ? 'inclusive' : 'exclusive'
  const parent = attrs(child(node, 'parentLabelLayout')).val
  if (parent === 'none' || parent === 'banner' || parent === 'overlapping') out.parentLabelLayout = parent
  return out
}
/** Authored binning, validated (map §5.1). Data-dependent checks (Scott's h, a binSize-derived count) are the family's. */
function binningOf(node: XmlNode): ChartExBinning {
  const invalid = () => new Unsupported(CHARTEX_INVALID_BINNING, true)
  const b = attrs(node)
  if (b.intervalClosed !== undefined && b.intervalClosed !== 'l' && b.intervalClosed !== 'r') throw invalid()
  const intervalClosed = b.intervalClosed === 'l' ? 'l' : 'r'
  const outer = (raw: string | undefined): number | undefined => {
    if (raw === undefined) return undefined
    if (raw.trim() === 'auto') throw new Unsupported(CHARTEX_AUTO_BINNING)
    const value = finite(raw)
    if (value === undefined) throw invalid()
    return value
  }
  const underflow = outer(b.underflow), overflow = outer(b.overflow)
  if (underflow !== undefined && overflow !== undefined && (intervalClosed === 'r' ? underflow >= overflow : underflow > overflow)) throw invalid()
  const sizeNode = child(node, 'binSize'), countNode = child(node, 'binCount')
  if (sizeNode && countNode) throw invalid()
  const out: ChartExBinning = { intervalClosed, ...(underflow !== undefined ? { underflow } : {}), ...(overflow !== undefined ? { overflow } : {}) }
  if (sizeNode) {
    const size = finite(valOrText(sizeNode))
    if (size === undefined || size <= 0) throw invalid()
    out.binSize = size
  }
  if (countNode) {
    const count = uint(valOrText(countNode))
    const inner = count === undefined ? undefined : count - (underflow !== undefined ? 1 : 0) - (overflow !== undefined ? 1 : 0)
    // The count includes the outer bins [D-hist]; at least one inner bin; at most 1,000 bins in all (spec 6.5).
    if (count === undefined || inner === undefined || inner < 1 || count > CHARTEX_BIN_LIMIT) throw invalid()
    out.binCount = count
  }
  return out
}

/** One bound `cx:data` set. Analytic families acquire every level in full (never a prefix); display families cap. */
function readData(id: number, node: XmlNode, analytic: boolean, admit: Admission, diagnose: (code: string, reason: string) => void): ChartExData {
  admit.charge(DATA_BYTES)
  const dimensions: ChartExDimension[] = []
  const seen = new Set<string>()
  for (const [name, dim] of orderedChildren(node)) {
    if (name !== 'numDim' && name !== 'strDim') continue
    const a = attrs(dim), kind = name === 'numDim' ? 'num' : 'str'
    if (!a.type) throw new Unsupported(CHARTEX_UNREADABLE)
    const type = a.type.slice(0, 16)
    // Several dimensions of one type: the first wins (map §2.3), diagnosed once.
    if (seen.has(`${kind}:${type}`)) diagnose('chartex-duplicate-dimension', `A ChartEx data set repeats a "${type}" dimension; the first is used`)
    seen.add(`${kind}:${type}`)
    const formula = own(child(dim, 'f')).trim()
    const ref = formula ? parseChartRange(formula) : undefined
    admit.charge(DIMENSION_BYTES + stringBytes(type) + (ref?.sheet ? stringBytes(ref.sheet) : 0) + (ref?.definedName ? stringBytes(ref.definedName) : 0))
    // A multi-level cache is uninterpretable saved data (map §2.2: the level order is unconfirmed): it is ignored
    // and the role reads its referenced range; with no readable reference the chart is unsupported.
    const cached = getChildren(dim, 'lvl')
    if (cached.length > 1 && !ref) throw new Unsupported(CHARTEX_MULTI_LEVEL)
    const levels = cached.length > 1 ? [] : cached.map(level => readLevel(kind, level, analytic, admit))
    dimensions.push({ kind, type, ...(ref ? { ref } : {}), dir: attrs(child(dim, 'f')).dir === 'row' ? 'row' : 'col', levels })
  }
  return { id, dimensions }
}

/**
 * One `cx:lvl`: `ptCount` points; a missing `pt@idx` is blank; an index at or past `ptCount` is ignored. An
 * analytic level over 65,536 points refuses the chart; a display level keeps the first 4,096 (truncated).
 */
function readLevel(kind: 'num' | 'str', node: XmlNode, analytic: boolean, admit: Admission): ChartExLevel {
  const a = attrs(node)
  const declared = uint(a.ptCount)
  if (declared === undefined) throw new Unsupported(CHARTEX_UNREADABLE)
  if (analytic && declared > CHART_ANALYTIC_POINT_LIMIT) throw new Unsupported(CHARTEX_TOO_MANY, true)
  const length = analytic ? declared : Math.min(declared, CHART_POINT_LIMIT)
  if (length < declared) admit.capped = true
  const formatCode = a.formatCode && a.formatCode.length <= 255 ? a.formatCode : undefined
  admit.charge(LEVEL_BYTES + (formatCode ? stringBytes(formatCode) : 0))
  const name = a.name ? boundedString(a.name, admit) : undefined
  const head = { ptCount: declared, ...(name ? { name } : {}), ...(formatCode ? { formatCode } : {}) }
  const points = getChildren(node, 'pt')
  if (kind === 'num') {
    admit.charge(pointBytes(length))
    const out = new PointWriter(length, false, admit.charge)
    for (const pt of points) {
      const i = uint(attrs(pt).idx)
      if (i === undefined || i >= length) continue
      const text = own(pt), v = finite(text)
      if (v !== undefined) out.value(i, v)
      else if (knownError(text) !== undefined) out.error(i, text.trim())
    }
    return { ...head, kind: 'num', points: out.done() }
  }
  admit.charge(length * SLOT_BYTES)
  const values = new Array<string | undefined>(length).fill(undefined)
  for (const pt of points) {
    const i = uint(attrs(pt).idx)
    if (i === undefined || i >= length) continue
    const text = own(pt)
    // Each label is admitted against the chart's label budget before its (bounded) slice is built.
    if (admit.label(text.length) === 'keep') values[i] = text.length > CHART_LABEL_LIMIT ? text.slice(0, CHART_LABEL_LIMIT) : text
  }
  return { ...head, kind: 'str', values }
}

const EMPTY_POINTS = (): ChartPoints => ({ length: 0, states: new Uint8Array(0), values: new Float64Array(0) })
/**
 * A series' roles from its data set (map §2.3): values from the first val/size/colorVal dimension, categories from
 * the first cat/entityId dimension. A single-level cache is used as read; its concrete range is kept for
 * provenance (a multi-level cache was dropped by `readData`, so the reference is read instead). Without a cache, a
 * reference is pending (XLSX: the workbook's defined names; DOCX/PPTX: the embedded workbook). Neither: the role is
 * absent (values) or defaulted (categories).
 */
function seriesRoles(data: ChartExData, name: string, color: string, admit: Admission, hierarchy = false): ChartSeries {
  const value = data.dimensions.find(d => d.kind === 'num' && VALUE_TYPES.includes(d.type))
  const category = data.dimensions.find(d => d.kind === 'str' && CATEGORY_TYPES.includes(d.type))
  const out: ChartSeries = { name, color, values: EMPTY_POINTS() }
  const pending: NonNullable<ChartSeries['pending']> = {}, refs: NonNullable<ChartSeries['refs']> = {}
  const valueLevel = value?.levels[0]
  if (valueLevel?.kind === 'num') { out.values = valueLevel.points; if (value?.ref && !value.ref.definedName) refs.values = value.ref }
  else if (value?.ref) pending.values = value.ref
  else out.absent = ['values']
  const categoryLevel = category?.levels[0]
  if (categoryLevel?.kind === 'str') {
    // The role shares the level's admitted strings (no copy); only its slot array is new.
    admit.charge(categoryLevel.values.length * SLOT_BYTES)
    out.categories = categoryLevel.values.map(text => text ?? '')
    if (category?.ref && !category.ref.definedName) refs.categories = category.ref
  } else if (category?.ref && hierarchy) {
    // A treemap/sunburst path: every level of the reference is read (one per column, or per row), root first.
    pending.levels = category.ref
    out.levelDir = category.dir
  } else if (category?.ref) pending.categories = category.ref
  if (Object.keys(pending).length) out.pending = pending
  if (Object.keys(refs).length) out.refs = refs
  return out
}

/**
 * The ChartEx adapter (spec 6.2, 6.3): reads the part and shapes the representation. A family whose painter has
 * not landed, region map, and every unreadable part are `unsupported` (an mc:Fallback is kept; a card only when no
 * representation remains); a budget breach, invalid binning or unavailable analytic data is `refused`. A painted
 * DOCX/PPTX chart applies the data precedence (its own cache, else its embedded workbook); XLSX roles are filled
 * from the workbook by the XLSX fill.
 */
export async function parseChartExPart(root: XmlNode, theme: ThemeContext | undefined, context: ChartParseContext, painted: ReadonlySet<ChartExFamily> = CHARTEX_PAINTED): Promise<ChartResult> {
  const selected = chartExFamily(root)
  const paint = !('unsupported' in selected) && painted.has(selected.family)
  // Colours matter only to a painted chart: its style and colour parts are read first (spec 5.11).
  const style = paint ? await readChartStyleParts(context) : undefined
  if (style === 'refused') return { kind: 'refused', reason: 'This chart needs more memory than is available for charts', feature: 'ChartEx' }
  let model: ChartModel | undefined
  try {
    model = parseChartEx(root, theme, style ? { ...context, styleParts: style.parts } : context)
    style?.release()
    if (!model.unsupported && !paint) model = { ...placeholderAfter(model), unsupported: CHARTEX_PENDING }
    else if (!model.unsupported && !XLSX_PART.test(context.partPath)) model = await applyDataPrecedence(model, root, context)
    // Histogram and Pareto (B3): structural checks now, so the fallback is kept.
    const histogram = !('unsupported' in selected) && HISTOGRAM_FAMILIES.has(selected.family)
    if (paint && histogram && !model.unsupported) {
      const structural = histogramStructure(model)
      if (structural) { const next = { ...model, unsupported: structural }; succeedChartModel(model, next); model = next }
    }
    // Stage 1 of a painted chart whose data is complete (its own cache, or filled from the embedded workbook); an
    // XLSX chart with references is prepared by the workbook fill instead (a histogram's hidden-cell provenance).
    if (paint && !model.unsupported && !model.preparedAxes && !chartHasPending(model) && !(histogram && XLSX_PART.test(context.partPath) && chartHasRefs(model))) model = prepareOwned(model)
  } catch { if (model) releaseChartModel(model); model = undefined } finally { style?.release() }
  if (model && !model.unsupported) return { kind: 'native', model }
  const title = model?.title
  if (model) releaseChartModel(model)
  // A required point that is unavailable (no cache and no readable workbook value) refuses the chart (B1 rulings).
  return { kind: model?.refused || model?.unsupported === CHART_DATA_NOT_AVAILABLE ? 'refused' : 'unsupported', reason: model?.unsupported ?? CHARTEX_UNREADABLE, ...(title ? { title } : {}), feature: 'ChartEx' }
}
/** Stage 1 (charts/prepare.ts) of an owned model: the successor takes over its charge; a refusal refuses the chart. */
function prepareOwned(model: ChartModel): ChartModel {
  let next: ChartModel
  try { next = prepareChartAxes(model, chartModelOwner(model)) } catch (error) {
    if (!(error instanceof ChartRefusal)) throw error
    next = { ...model, preparedAxes: undefined, unsupported: 'This chart needs more memory than is available for charts', refused: true }
  }
  succeedChartModel(model, next)
  return next
}
/** A parsed model whose family is not painted: released here (its title is kept for the result). */
function placeholderAfter(model: ChartModel): ChartModel {
  const title = model.title
  releaseChartModel(model)
  return { ...placeholderChart(title, CHARTEX_PENDING), feature: 'ChartEx' }
}
