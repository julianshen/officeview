/**
 * The ChartEx (`cx:`, Office 2016+) part parser (spec 4.2, 4.3; lazy chunk), normative against
 * docs/charts/chartex-map.md. It produces the engine's one `ChartModel`, whose single group is a
 * discriminated `ChartExGroup` (charts/model.ts):
 *
 * - `cx:chartData/cx:data[@id]` sets are keyed by id; only the sets a series binds are acquired. Each
 *   dimension keeps its `@type`, its reference and every `cx:lvl` (sparse `pt@idx`; a missing index is blank).
 * - A series binds its data set through `cx:dataId`; its roles are chosen by dimension type (values: val, size,
 *   colorVal; categories: cat, entityId; the first of a type wins). A `paretoLine` series has no data: it binds
 *   its owner through `ownerIdx` and is kept beside the group's series.
 * - Saved data (a `cx:lvl` cache) is used as read. Without one, the role is a pending reference: XLSX resolves
 *   its hidden `_xlchart.*` defined name against the workbook (the XLSX fill); DOCX/PPTX read the embedded
 *   workbook (`cx:chartData/cx:externalData`, charts/dataPrecedence.ts).
 * - `layoutId` selects the family (map §3). Region map is recognised first and is always unsupported; its
 *   geography (`cx:binary`) is never read.
 *
 * Until a family's painter lands (B2-B5) the adapter answers `unsupported` for it, so an mc:Fallback is kept
 * (spec 6.3). Every retained allocation is charged to the document's chart workspace before it is made
 * (charts/own.ts); a refusal stops the parse and refunds the charge.
 */
import { attrs, getChildren, orderedChildren, textOf, type XmlNode } from '../../core/xml'
import type { ThemeContext } from '../../drawing/style'
import type { ChartParseContext } from '../context'
import type {
  ChartAxis, ChartAxisBinding, ChartExBinning, ChartExData, ChartExDimension, ChartExFamily, ChartExGroup, ChartExLayout,
  ChartExLayoutProperties, ChartExLevel, ChartExSeriesInfo, ChartGroup, ChartModel, ChartPoints, ChartResult, ChartSeries, ChartTickMark,
} from '../model'
import { ChartRefusal, adoptChartModel, modelOwner, releaseChartModel } from '../own'
import { CHART_ANALYTIC_POINT_LIMIT, CHART_LABEL_LIMIT, CHART_SERIES_LIMIT, axisBytes, parseChartRange, placeholderChart } from '../parse'
import { PointWriter, knownError, pointBytes, stringBytes } from '../points'
import { resolveSeriesColor } from '../style'
import { readChartStyleParts } from '../styleParts'
import { applyDataPrecedence } from '../dataPrecedence'

/** Families whose painter has landed. Every other family stays `unsupported` (its fallback is kept). */
export const CHARTEX_PAINTED: ReadonlySet<ChartExFamily> = new Set<ChartExFamily>()
export const CHARTEX_PENDING = 'ChartEx charts are not supported yet'
export const CHARTEX_REGION_MAP = 'Region map charts are not supported'
export const CHARTEX_UNREADABLE = 'This chart could not be read'
export const CHARTEX_MULTI_LEVEL = 'multi-level cache order unconfirmed'
export const CHARTEX_BIN_LIMIT = 1000
export const CHARTEX_TOO_MANY_BINS = 'This chart has more histogram bins than the viewer draws'
export const CHARTEX_TOO_MANY = 'This chart has too many data points to summarize'
const CHARTEX_DATA_UNAVAILABLE = "This chart's data is not available"
const LAYOUTS: readonly ChartExLayout[] = ['clusteredColumn', 'paretoLine', 'waterfall', 'funnel', 'boxWhisker', 'treemap', 'sunburst', 'regionMap']
const VALUE_TYPES = ['val', 'size', 'colorVal'], CATEGORY_TYPES = ['cat', 'entityId']
/** Retained record prices, charged before each record is built. */
const DATA_BYTES = 64, DIMENSION_BYTES = 128, LEVEL_BYTES = 96, SERIES_BYTES = 384, SLOT_BYTES = 8
const XLSX_PART = /^\/?xl\//i

export const isChartExGroup = (group: ChartGroup): group is ChartExGroup => group.kind === 'chartEx'

/** A structural failure: the chart is unsupported with `reason` (never a thrown parse error). */
class Unsupported extends Error {}
const child = (node: XmlNode | undefined, name: string): XmlNode | undefined => getChildren(node, name)[0]
const uint = (raw: string | undefined): number | undefined => raw !== undefined && /^\s*\d+\s*$/.test(raw) && Number.isSafeInteger(Number(raw)) ? Number(raw) : undefined
const finite = (raw: string | undefined): number | undefined => raw !== undefined && raw.trim() !== '' && Number.isFinite(Number(raw)) ? Number(raw) : undefined
const bool = (raw: string | undefined, fallback: boolean): boolean => raw === '1' || raw === 'true' ? true : raw === '0' || raw === 'false' ? false : fallback
/** Office writes some schema text elements as `@val` (map §12.7): `@val` first, then the text. */
const valOrText = (node: XmlNode | undefined): string | undefined => node ? attrs(node).val ?? textOf(node) : undefined
const auto = (raw: string | undefined): number | 'auto' | undefined => raw === 'auto' ? 'auto' : finite(raw)
const runText = (paragraphs: XmlNode[]): string => paragraphs.map(p => orderedChildren(p).filter(([name]) => name === 'r' || name === 'fld').map(([, run]) => getChildren(run, 't').map(textOf).join('')).join('')).join('\n').trim()

/** Title text (map §1): `tx/txData/v`, the `tx/rich` text, else the auto title's `txPr` run text. */
function titleText(title: XmlNode | undefined): string | undefined {
  if (!title) return undefined
  const tx = child(title, 'tx')
  const text = textOf(child(child(tx, 'txData'), 'v')).trim() || runText(getChildren(child(tx, 'rich'), 'p')) || runText(getChildren(child(title, 'txPr'), 'p'))
  return text ? text.slice(0, CHART_LABEL_LIMIT) : undefined
}
const seriesName = (series: XmlNode): string | undefined => {
  const tx = child(series, 'tx')
  const text = textOf(child(child(tx, 'txData'), 'v')).trim() || runText(getChildren(child(tx, 'rich'), 'p'))
  return text ? text.slice(0, CHART_LABEL_LIMIT) : undefined
}

/** The family of the plot area's series (map §3), or the reason it has none. Cheap: no data is read. */
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
    const layoutPr = child(nodes[layouts.indexOf('clusteredColumn')], 'layoutPr')
    const byCategory = child(layoutPr, 'binning') ? undefined : true as const
    return { family: lines.length ? 'pareto' : 'histogram', ...(byCategory ? { byCategory } : {}) }
  }
  return { family: layout as ChartExFamily }
}

/** Read one ChartEx part into a model (a placeholder carrying `feature: 'ChartEx'` when it cannot be read). */
export function parseChartEx(root: XmlNode, theme: ThemeContext | undefined, context?: ChartParseContext): ChartModel {
  const chart = child(root, 'chart')
  const title = titleText(child(chart, 'title'))
  const fail = (reason: string, refused = false): ChartModel => ({ ...placeholderChart(title, reason), feature: 'ChartEx', ...(refused ? { refused: true as const } : {}) })
  const selected = chartExFamily(root)
  if ('unsupported' in selected) return fail(selected.unsupported)
  const owner = context?.workspace ? modelOwner(context.workspace) : undefined
  const charge = (bytes: number) => { if (owner && !owner.workspace.reserve(bytes, owner.tag)) throw new ChartRefusal() }
  try {
    const model = readChartEx(root, chart!, title, selected.family, selected.byCategory, theme, charge, context)
    if (owner) adoptChartModel(model, owner)
    return model
  } catch (error) {
    owner?.workspace.release(owner.tag)
    if (error instanceof ChartRefusal) return fail('This chart needs more memory than is available for charts', true)
    // An analytic budget breach refuses the chart (never a prefix summarized; B1 rulings).
    if (error instanceof Unsupported) return fail(error.message, error.message === CHARTEX_TOO_MANY || error.message === CHARTEX_TOO_MANY_BINS)
    throw error
  }
}

function readChartEx(root: XmlNode, chart: XmlNode, title: string | undefined, family: ChartExFamily, byCategory: true | undefined,
  theme: ThemeContext | undefined, charge: (bytes: number) => void, context?: ChartParseContext): ChartModel {
  const plot = child(chart, 'plotArea')
  const nodes = getChildren(child(plot, 'plotAreaRegion'), 'series')
  if (nodes.length > CHART_SERIES_LIMIT) throw new Unsupported(CHARTEX_UNREADABLE)
  const xlsx = XLSX_PART.test(context?.partPath ?? '')
  if (title) charge(stringBytes(title))

  // Axes (map §1): the kind is the scaling child, never the id or the order.
  const axes: ChartAxis[] = []
  for (const node of getChildren(plot, 'axis')) {
    const a = attrs(node), cat = child(node, 'catScaling'), val = child(node, 'valScaling')
    if (a.id === undefined || !cat === !val) throw new Unsupported(CHARTEX_UNREADABLE)
    const scaling = attrs(val), numFmt = attrs(child(node, 'numFmt')), tick = attrs(child(node, 'majorTickMarks')).type
    const axisTitle = titleText(child(node, 'title')), unit = attrs(child(node, 'units')).unit
    const axis: ChartAxis = {
      id: a.id, kind: cat ? 'category' : 'value', deleted: bool(a.hidden, false),
      gridlines: !!child(node, 'majorGridlines'), ...(child(node, 'minorGridlines') ? { minorGridlines: true } : {}),
      ...numberField('min', scaling.min), ...numberField('max', scaling.max), ...numberField('majorUnit', scaling.majorUnit), ...numberField('minorUnit', scaling.minorUnit),
      ...(axisTitle ? { title: axisTitle } : {}), ...(numFmt.formatCode ? { formatCode: numFmt.formatCode, sourceLinked: bool(numFmt.sourceLinked, false) } : {}),
      ...(['cross', 'in', 'out', 'none'].includes(tick ?? '') ? { majorTickMark: tick as ChartTickMark } : {}),
      ...(child(node, 'tickLabels') ? {} : { tickLblPos: 'none' as const }), ...(unit ? { unit } : {}),
    }
    charge(axisBytes(axis) + (unit ? stringBytes(unit) : 0))
    axes.push(axis)
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
  const parts = context?.styleParts
  const series: ChartSeries[] = [], infos: ChartExSeriesInfo[] = []
  let paretoLine: ChartExSeriesInfo | undefined, gapWidth: number | undefined
  nodes.forEach((node, position) => {
    const a = attrs(node)
    const layout = a.layoutId as ChartExLayout
    const axisIds = getChildren(node, 'axisId').map(valOrText).filter((id): id is string => !!id)
    const dataId = uint(valOrText(child(node, 'dataId')))
    const info: ChartExSeriesInfo = {
      layout, ...(dataId !== undefined ? { dataId } : {}),
      ...(uint(a.ownerIdx) !== undefined ? { ownerIdx: uint(a.ownerIdx) } : {}), hidden: bool(a.hidden, false),
      formatIdx: uint(a.formatIdx) ?? position, axisIds, layoutPr: layoutProperties(child(node, 'layoutPr'), charge),
    }
    charge(SERIES_BYTES + axisIds.reduce((sum, id) => sum + stringBytes(id), 0))
    if (layout === 'paretoLine') { paretoLine = info; return }
    if (info.dataId === undefined || !dataSets.has(info.dataId)) throw new Unsupported(CHARTEX_UNREADABLE)
    let data = acquired.get(info.dataId)
    if (!data) { data = readData(info.dataId, dataSets.get(info.dataId)!, charge); acquired.set(info.dataId, data) }
    const name = seriesName(node) ?? `Series ${series.length + 1}`
    charge(stringBytes(name))
    const color = resolveSeriesColor({ index: info.formatIdx, count: nodes.length, ...(parts ? { parts } : {}), theme })
    series.push(seriesRoles(data, name, color, charge))
    infos.push(info)
    if (gapWidth === undefined) {
      const category = axes.find(axis => axis.kind === 'category' && axisIds.includes(axis.id)) ?? firstCategory
      gapWidth = finite(attrs(child(getChildren(plot, 'axis').find(node => attrs(node).id === category?.id), 'catScaling')).gapWidth)
    }
  })

  const bound = infos[0]?.axisIds ?? []
  const category = axes.find(axis => axis.kind === 'category' && bound.includes(axis.id)) ?? firstCategory
  const value = axes.find(axis => axis.kind === 'value' && bound.includes(axis.id)) ?? firstValue
  const binding: ChartAxisBinding = category && value ? { kind: 'cartesian', category: category.id, value: value.id } : { kind: 'none' }
  const group: ChartExGroup = {
    kind: 'chartEx', family, grouping: 'standard', series, binding, gapWidth: gapWidth !== undefined ? Math.round(gapWidth * 100) : 150, overlap: 0,
    chartEx: { data: [...acquired.values()], series: infos, ...(paretoLine ? { paretoLine } : {}), ...(byCategory ? { byCategory } : {}), ...(gapWidth !== undefined ? { gapWidth } : {}) },
  }
  const legend = child(chart, 'legend'), legendPos = attrs(legend).pos ?? 'r'
  // A value role with neither saved data nor a reference: the chart's data is not available (never blank).
  const unavailable = series.some(s => s.absent?.includes('values'))
  return {
    ...(title ? { title } : {}), groups: [group], axes, categories: series[0]?.categories ?? [],
    ...(legend && ['t', 'b', 'l', 'r'].includes(legendPos) ? { legend: legendPos as ChartModel['legend'] } : {}),
    fontFamily: theme?.fonts.minor.latin ?? 'Calibri', fontSizePt: 10,
    // XLSX: a required role with neither cache nor reference refuses the chart (B1 rulings). DOCX/PPTX are decided by
    // the data precedence (cache, then the embedded workbook), never rejected here for lacking a cache.
    ...(xlsx && unavailable ? { unsupported: CHARTEX_DATA_UNAVAILABLE, feature: 'ChartEx' as const, refused: true as const } : {}),
  }
}
const numberField = <K extends string>(key: K, raw: string | undefined): Partial<Record<K, number>> => {
  const value = finite(raw)
  return value === undefined ? {} : { [key]: value } as Record<K, number>
}

/** `cx:layoutPr` with the map's absence defaults (§4.1 visibility, §5.1 intervalClosed r, §8 exclusive). */
function layoutProperties(node: XmlNode | undefined, charge: (bytes: number) => void): ChartExLayoutProperties {
  const v = attrs(child(node, 'visibility'))
  const out: ChartExLayoutProperties = {
    visibility: { connectorLines: bool(v.connectorLines, true), meanLine: bool(v.meanLine, false), meanMarker: bool(v.meanMarker, true), nonoutliers: bool(v.nonoutliers, false), outliers: bool(v.outliers, true) },
  }
  const binning = child(node, 'binning')
  if (binning) {
    const b = attrs(binning)
    const size = finite(valOrText(child(binning, 'binSize'))), count = uint(valOrText(child(binning, 'binCount')))
    // An authored bin count above 1,000 is out of range and essential: the chart is refused (spec 6.5 table).
    // (A count derived from binSize depends on the data and is bounded by the histogram family task.)
    if (count !== undefined && count > CHARTEX_BIN_LIMIT) throw new Unsupported(CHARTEX_TOO_MANY_BINS)
    const under = auto(b.underflow), over = auto(b.overflow)
    const result: ChartExBinning = {
      intervalClosed: b.intervalClosed === 'l' ? 'l' : 'r', ...(size !== undefined && size > 0 ? { binSize: size } : {}),
      ...(count !== undefined && count > 0 ? { binCount: count } : {}), ...(under !== undefined ? { underflow: under } : {}), ...(over !== undefined ? { overflow: over } : {}),
    }
    out.binning = result
  } else if (child(node, 'aggregation')) out.aggregation = true
  const subtotals = child(node, 'subtotals')
  if (subtotals) {
    const idx = getChildren(subtotals, 'idx')
    charge(idx.length * SLOT_BYTES)
    out.subtotals = idx.map(item => uint(valOrText(item))).filter((i): i is number => i !== undefined)
  }
  const statistics = attrs(child(node, 'statistics')).quartileMethod
  if (child(node, 'statistics') || statistics) out.quartileMethod = statistics === 'inclusive' ? 'inclusive' : 'exclusive'
  const parent = attrs(child(node, 'parentLabelLayout')).val
  if (parent === 'none' || parent === 'banner' || parent === 'overlapping') out.parentLabelLayout = parent
  return out
}

/** One bound `cx:data` set: every dimension, each level acquired in full (analytic bound; never a prefix). */
function readData(id: number, node: XmlNode, charge: (bytes: number) => void): ChartExData {
  charge(DATA_BYTES)
  const dimensions: ChartExDimension[] = []
  for (const [name, dim] of orderedChildren(node)) {
    if (name !== 'numDim' && name !== 'strDim') continue
    const a = attrs(dim), kind = name === 'numDim' ? 'num' : 'str'
    if (!a.type) throw new Unsupported(CHARTEX_UNREADABLE)
    const formula = textOf(child(dim, 'f')).trim()
    const ref = formula ? parseChartRange(formula) : undefined
    charge(DIMENSION_BYTES + stringBytes(a.type) + (ref?.sheet ? stringBytes(ref.sheet) : 0) + (ref?.definedName ? stringBytes(ref.definedName) : 0))
    // A multi-level cache is uninterpretable saved data (map §2.2: the level order is unconfirmed): it is ignored
    // and the role reads its referenced range; with no readable reference the chart is unsupported.
    const cached = getChildren(dim, 'lvl')
    if (cached.length > 1 && !ref) throw new Unsupported(CHARTEX_MULTI_LEVEL)
    const levels = cached.length > 1 ? [] : cached.map(level => readLevel(kind, level, charge))
    dimensions.push({ kind, type: a.type, ...(ref ? { ref } : {}), dir: attrs(child(dim, 'f')).dir === 'row' ? 'row' : 'col', levels })
  }
  return { id, dimensions }
}

/** One `cx:lvl`: `ptCount` points; a missing `pt@idx` is blank; an index at or past `ptCount` is ignored. */
function readLevel(kind: 'num' | 'str', node: XmlNode, charge: (bytes: number) => void): ChartExLevel {
  const a = attrs(node)
  const declared = uint(a.ptCount)
  if (declared === undefined) throw new Unsupported(CHARTEX_UNREADABLE)
  // Analytic inputs are acquired in full up to the analytic bound; above it the chart is not summarized (spec 4.3).
  if (declared > CHART_ANALYTIC_POINT_LIMIT) throw new Unsupported(CHARTEX_TOO_MANY)
  charge(LEVEL_BYTES + (a.formatCode ? stringBytes(a.formatCode) : 0) + (a.name ? stringBytes(a.name) : 0))
  const head = { ptCount: declared, ...(a.name ? { name: a.name.slice(0, CHART_LABEL_LIMIT) } : {}), ...(a.formatCode ? { formatCode: a.formatCode.slice(0, 255) } : {}) }
  const points = getChildren(node, 'pt')
  if (kind === 'num') {
    charge(pointBytes(declared))
    const out = new PointWriter(declared, false, charge)
    for (const pt of points) {
      const i = uint(attrs(pt).idx)
      if (i === undefined || i >= declared) continue
      const text = textOf(pt), v = finite(text)
      if (v !== undefined) out.value(i, v)
      else if (knownError(text) !== undefined) out.error(i, text.trim())
    }
    return { ...head, kind: 'num', points: out.done() }
  }
  charge(declared * SLOT_BYTES)
  const values = new Array<string | undefined>(declared).fill(undefined)
  for (const pt of points) {
    const i = uint(attrs(pt).idx)
    if (i === undefined || i >= declared) continue
    const text = textOf(pt).slice(0, CHART_LABEL_LIMIT)
    charge(stringBytes(text))
    values[i] = text
  }
  return { ...head, kind: 'str', values }
}

const EMPTY_POINTS = (): ChartPoints => ({ length: 0, states: new Uint8Array(0), values: new Float64Array(0) })
/**
 * A series' roles from its data set (map §2.3): values from the first val/size/colorVal dimension, categories from
 * the first cat/entityId dimension. A single-level cache is used as read; its reference is kept for provenance (a
 * multi-level cache was dropped by `readData`, so the reference is read instead). Without a cache, a reference is pending (XLSX: the workbook's
 * defined names; DOCX/PPTX: the embedded workbook). Neither: the role is absent (values) or defaulted (categories).
 */
function seriesRoles(data: ChartExData, name: string, color: string, charge: (bytes: number) => void): ChartSeries {
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
    charge(categoryLevel.values.length * SLOT_BYTES)
    out.categories = categoryLevel.values.map(text => text ?? '')
    if (category?.ref && !category.ref.definedName) refs.categories = category.ref
  } else if (category?.ref) pending.categories = category.ref
  if (Object.keys(pending).length) out.pending = pending
  if (Object.keys(refs).length) out.refs = refs
  return out
}

/**
 * The ChartEx adapter (spec 6.2, 6.3): reads the part and shapes the representation. A family whose painter has
 * not landed, region map, and every unreadable part are `unsupported` (an mc:Fallback is kept; a card only when no
 * representation remains); a budget breach is `refused`. A painted DOCX/PPTX chart applies the data precedence
 * (its own cache, else its embedded workbook); XLSX roles are filled from the workbook by the XLSX fill.
 */
export async function parseChartExPart(root: XmlNode, theme: ThemeContext | undefined, context: ChartParseContext, painted: ReadonlySet<ChartExFamily> = CHARTEX_PAINTED): Promise<ChartResult> {
  const selected = chartExFamily(root)
  const title = titleText(child(child(root, 'chart'), 'title'))
  const unsupported = (reason: string): ChartResult => ({ kind: 'unsupported', reason, ...(title ? { title } : {}), feature: 'ChartEx' })
  if ('unsupported' in selected) return unsupported(selected.unsupported)
  const paint = painted.has(selected.family)
  // Colours matter only to a painted chart: its style and colour parts are read first (spec 5.11).
  const style = paint ? await readChartStyleParts(context) : undefined
  if (style === 'refused') return { kind: 'refused', reason: 'This chart needs more memory than is available for charts', ...(title ? { title } : {}), feature: 'ChartEx' }
  let model: ChartModel | undefined
  try {
    model = parseChartEx(root, theme, style ? { ...context, styleParts: style.parts } : context)
    style?.release()
    if (!model.unsupported && !paint) { releaseChartModel(model); return unsupported(CHARTEX_PENDING) }
    if (!model.unsupported && !XLSX_PART.test(context.partPath)) {
      model = await applyDataPrecedence(model, root, context)
    }
  } catch { if (model) releaseChartModel(model); model = undefined } finally { style?.release() }
  if (model && !model.unsupported) return { kind: 'native', model }
  if (model) releaseChartModel(model)
  // A required point that is unavailable (no cache and no readable workbook value) refuses the chart (B1 rulings).
  return { kind: model?.refused || model?.unsupported === CHARTEX_DATA_UNAVAILABLE ? 'refused' : 'unsupported', reason: model?.unsupported ?? CHARTEX_UNREADABLE, ...(title ? { title } : {}), feature: 'ChartEx' }
}
