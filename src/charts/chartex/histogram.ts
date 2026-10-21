/**
 * Histogram and Pareto (ChartEx `clusteredColumn` + `paretoLine`; chartex-map §5, spec 4.3; lazy chunk).
 *
 * Stage 1 (size-independent) turns the parsed ChartEx group into the engine's plain column (and line) groups,
 * so the shared axes, labels, layout and painter draw it:
 *
 * - **Binned** (`cx:binning`, map §5.1): the value points of every visible column series are pooled into one set
 *   of bins (one category axis) and each series is counted into them. Outer bins are taken first and removed
 *   before inner binning, closure `r`: underflow `v <= U`, overflow `v > O`, inner `(e_i, e_i+1]` (the first bin
 *   closed when U is unset); closure `l`: underflow `v < U`, overflow `v >= O`, inner `[e_i, e_i+1)` (the last
 *   closed when O is unset). The width is `binSize`, `(hi - lo) / k` for `binCount` (the count includes the
 *   outer bins), or Scott's rule `3.49 s n^(-1/3)` (sample s of the inner set). Membership is decided against the
 *   very edges the labels show. More than 1,000 bins (outer included), or an unusable Scott width, refuses the
 *   chart ("invalid binning"); an unavailable point refuses it ("This chart's data is not available").
 * - **By Category** (`cx:aggregation`): equal categories (exact text) are grouped in first-occurrence order and
 *   their values summed; blank and error values add nothing.
 * - **Pareto**: the owner (the column series the `paretoLine`'s `ownerIdx` names, in document order) is sorted
 *   descending with ties in their prior order; the line is the cumulative fraction on the `paretoLine`'s own axis
 *   (0..1, percent ticks, on the right edge).
 *
 * The derived model keeps no source provenance (hidden indices, references) and no per-point overrides (their
 * source-index meaning for bins is unconfirmed). Every retained allocation is charged before it is built; the
 * transient work (aggregation map, sort order, edges) is charged under a scratch tag released on return.
 */
import type { ChartAxis, ChartExBinning, ChartExGroup, ChartExSeriesInfo, ChartGroup, ChartModel, ChartPoints, ChartSeries } from '../model'
import type { ChartPointOverride } from '../style'
import type { LabelOptions } from '../labels'
import type { ChartWorkspace } from '../workspace'
import { ChartRefusal } from '../own'
import { PointWriter, POINT_UNAVAILABLE, POINT_VALUE, hiddenAt, pointBytes, stringBytes } from '../points'
import { createChartFormatter } from '../format'
import { chartEpoch } from '../dates'
import { CHART_BIN_LIMIT, CHART_DATA_NOT_AVAILABLE, CHART_INVALID_BINNING, CHART_POINT_LIMIT, CHART_UNREADABLE } from '../limits'

export const HISTOGRAM_FAMILIES: ReadonlySet<string> = new Set(['histogram', 'pareto'])
export const HISTOGRAM_INVALID_BINNING = CHART_INVALID_BINNING
export const HISTOGRAM_UNREADABLE = CHART_UNREADABLE
export const HISTOGRAM_BIN_LIMIT = CHART_BIN_LIMIT
const CATEGORY_TYPES = ['cat', 'entityId']
/** Retained record prices (a derived group, series or axis record) and a transient aggregation-map entry. */
const RECORD_BYTES = 512, MAP_ENTRY_BYTES = 96, SLOT_BYTES = 8
/** The bound of one built bin label: two formatted edges (each at most 256 units) and the brackets. */
const LABEL_BOUND = stringBytes(' '.repeat(2 * 256 + 4))

export const isHistogramGroup = (group: ChartGroup): group is ChartExGroup => group.kind === 'chartEx' && HISTOGRAM_FAMILIES.has(group.family)

/** The column series that take part in layout (map §2.3: a hidden series is hidden from layout), with their infos. */
const visibleColumns = (group: ChartExGroup) => group.chartEx.series.flatMap((info, i) => info.hidden || !group.series[i] ? [] : [{ info, series: group.series[i] }])

/** A structural reason the chart cannot be drawn (data-independent; the adapter keeps the fallback), or undefined. */
export function histogramStructure(model: ChartModel): string | undefined {
  const group = model.groups[0]
  if (!group || !isHistogramGroup(group)) return undefined
  const { paretoLine, byCategory, data } = group.chartEx
  const infos = visibleColumns(group).map(c => c.info)
  // A Pareto sorts one owner: the column series its ownerIdx names (document position, hidden series counted; map
  // §2.3, §5.2). Several visible columns have no single order; a hidden owner with another visible column has none.
  if (paretoLine && (infos.length > 1 || (infos.length === 1 && infos[0].position !== paretoLine.ownerIdx))) return HISTOGRAM_UNREADABLE
  // By Category needs the categories it groups (map §5.1 [P]).
  if (byCategory && infos.some(info => !data.find(d => d.id === info.dataId)?.dimensions.some(d => d.kind === 'str' && CATEGORY_TYPES.includes(d.type)))) return HISTOGRAM_UNREADABLE
  // One set of bins serves every visible series: their binning must agree.
  const key = (info: ChartExSeriesInfo) => JSON.stringify(info.layoutPr.binning ?? null)
  if (infos.some(info => key(info) !== key(infos[0]))) return HISTOGRAM_UNREADABLE
  return undefined
}

/**
 * Derived columns: labels, per-series heights and, for By Category, each column's source point per series (-1:
 * none; -2: several, that slot carries no point override). `undefined` sources (bins): no column is one source
 * point, so point overrides are not carried.
 */
interface Derived { labels: string[]; heights: Float64Array[]; sources: Array<Int32Array | undefined>; retained: number }
/** Values whose sums or edges pass the double range (spec 5.8: never NaN or Infinity drawn as a value). */
export const HISTOGRAM_OUT_OF_RANGE = "This chart's values are out of range"
export const HISTOGRAM_OVERRIDES_DROPPED = 'chart-histogram-point-overrides'

/**
 * Stage 1 for a histogram or Pareto model with its data in place: the derived model (plain column and line groups,
 * the bin or category labels), or the model marked unsupported/refused. `ownerTag`: the model owner's tag, which
 * takes over the retained bytes (absent: an unowned model; the charge ends with the call). Never re-derives a
 * derived model (it has no ChartEx group).
 */
export function deriveHistogram(model: ChartModel, workspace: ChartWorkspace, ownerTag?: string): ChartModel {
  const group = model.groups[0]
  if (!group || !isHistogramGroup(group)) return model
  const structural = histogramStructure(model)
  if (structural) return { ...model, preparedAxes: undefined, unsupported: structural }
  const scratch = `chart-histogram#${++jobs}`
  const charge = (bytes: number) => { if (!workspace.reserve(bytes, scratch)) throw new ChartRefusal() }
  try {
    const columns = visibleColumns(group), visible = columns.map(c => c.series)
    const visibleOnly = model.plotVisOnly !== false
    // Analytic inputs (map §2.4): an unavailable point refuses the chart; blank and error points are excluded.
    for (const s of visible) for (let i = 0; i < s.values.length; i++) if (s.values.states[i] === POINT_UNAVAILABLE) return refuse(model, CHART_DATA_NOT_AVAILABLE)
    const take = (s: ChartSeries, i: number) => s.values.states[i] === POINT_VALUE && !(visibleOnly && hiddenAt(s.values.hidden, i))
    const derived = group.chartEx.byCategory ? aggregate(model, visible, take, charge)
      : bin(model, group, columns[0]?.info.layoutPr.binning, visible, take, charge, workspace)
    if (typeof derived === 'string') return refuse(model, derived)
    // A Pareto's columns are sorted by its owner even when its line is hidden; the line is drawn only when visible.
    const pareto = group.chartEx.paretoLine
    const built = pareto ? sortPareto(derived, charge) : { ...derived, cumulative: undefined }
    if (typeof built === 'string') return refuse(model, built)
    // The derived records (groups, series, axes) and point storage, then the move of every retained byte to the owner.
    const records = RECORD_BYTES * (visible.length + 4)
    charge(records)
    const out = assemble(model, group, visible, built, pareto && !pareto.hidden ? pareto : undefined, charge)
    if (ownerTag) workspace.transfer(scratch, ownerTag, built.retained + records)
    return out
  } finally { workspace.release(scratch) }
}
let jobs = 0
const refuse = (model: ChartModel, reason: string): ChartModel => ({ ...model, preparedAxes: undefined, unsupported: reason, refused: true })

/**
 * The power-of-two divisor that keeps a sum of `count` terms of at most `maxAbs` inside the double range. Used only
 * for a sum whose plain form passed the range: dividing by a power of two is exact for every term it does not
 * push below the smallest normal, and such a term cannot change an overflowing sum.
 */
const sumScale = (maxAbs: number, count: number): number => maxAbs > 0 ? 2 ** Math.max(0, Math.ceil(Math.log2(maxAbs) + Math.log2(Math.max(1, count)) + 1 - 1023)) : 1

/**
 * By Category (map §5.1): equal category texts grouped in first-occurrence order (pooled across series), values
 * summed. Each slot is summed plainly; only a slot whose plain sum passes the double range is summed again with its
 * own power-of-two scale (small representable sums are never scaled). `sources[s][j]`: the source index of the one
 * point of series s in slot j, -1 none, -2 several (that slot carries no point override).
 */
function aggregate(model: ChartModel, visible: ChartSeries[], take: (s: ChartSeries, i: number) => boolean, charge: (bytes: number) => void): Derived | string {
  const index = new Map<string, number>(), labels: string[] = []
  for (const s of visible) {
    const categories = s.categories ?? model.categories
    const n = Math.max(s.values.length, categories.length)
    for (let i = 0; i < n; i++) {
      const text = categories[i] ?? ''
      if (index.has(text)) continue
      // A map entry (transient) and a label slot (retained; the text is the source's own string).
      charge(MAP_ENTRY_BYTES + SLOT_BYTES)
      index.set(text, labels.length)
      labels.push(text)
    }
  }
  charge(labels.length * 12 * visible.length)
  const sources: Array<Int32Array | undefined> = []
  const heights = visible.map(s => {
    const sums = new Float64Array(labels.length), source = new Int32Array(labels.length).fill(-1)
    const categories = s.categories ?? model.categories
    const slot = (i: number) => index.get(categories[i] ?? '')!
    for (let i = 0; i < Math.max(s.values.length, categories.length); i++) {
      const j = slot(i)
      // A column is one source point of this series while no second point of the series joins its slot.
      source[j] = source[j] === -1 ? i : -2
      if (i < s.values.length && take(s, i)) sums[j] += s.values.values[i]
    }
    // Slots whose plain sum passed the range: summed again, each with its own scale.
    const over = new Set<number>()
    for (let j = 0; j < sums.length; j++) if (!Number.isFinite(sums[j])) over.add(j)
    if (over.size) {
      const maxAbs = new Map<number, number>(), count = new Map<number, number>()
      for (let i = 0; i < s.values.length; i++) {
        const j = slot(i)
        if (!over.has(j) || !take(s, i)) continue
        maxAbs.set(j, Math.max(maxAbs.get(j) ?? 0, Math.abs(s.values.values[i])))
        count.set(j, (count.get(j) ?? 0) + 1)
      }
      const scaled = new Map<number, number>()
      for (const j of over) scaled.set(j, 0)
      for (let i = 0; i < s.values.length; i++) {
        const j = slot(i)
        if (!over.has(j) || !take(s, i)) continue
        scaled.set(j, scaled.get(j)! + s.values.values[i] / sumScale(maxAbs.get(j)!, count.get(j)!))
      }
      for (const j of over) sums[j] = scaled.get(j)! * sumScale(maxAbs.get(j) ?? 0, count.get(j) ?? 0)
    }
    sources.push(source)
    return sums
  })
  if (heights.some(h => h.some(v => !Number.isFinite(v)))) return HISTOGRAM_OUT_OF_RANGE
  return { labels, heights, sources, retained: labels.length * SLOT_BYTES }
}

/** Binned (map §5.1): pooled bins over every visible series; a string is the refusal reason. */
function bin(model: ChartModel, group: ChartExGroup, authored: ChartExBinning | undefined, visible: ChartSeries[], take: (s: ChartSeries, i: number) => boolean,
  charge: (bytes: number) => void, workspace: ChartWorkspace): Derived | string {
  const binning: ChartExBinning = authored ?? { intervalClosed: 'r' }
  const right = binning.intervalClosed !== 'l'
  const U = typeof binning.underflow === 'number' ? binning.underflow : undefined
  const O = typeof binning.overflow === 'number' ? binning.overflow : undefined
  if (binning.underflow === 'auto' || binning.overflow === 'auto') return HISTOGRAM_UNREADABLE
  const under = (v: number) => U !== undefined && (right ? v <= U : v < U)
  const over = (v: number) => O !== undefined && (right ? v > O : v >= O)
  // The inner set's extent, then its moments (Welford) on values scaled by one power of two, so a spread near the
  // double range cannot overflow the moments.
  let n = 0, min = Infinity, max = -Infinity
  for (const s of visible) for (let i = 0; i < s.values.length; i++) {
    if (!take(s, i)) continue
    const v = s.values.values[i]
    if (under(v) || over(v)) continue
    n++
    if (v < min) min = v
    if (v > max) max = v
  }
  const outer = (U !== undefined ? 1 : 0) + (O !== undefined ? 1 : 0)
  // Range, width and inner-bin count (map §5.1 steps 3-6). Differences that pass the double range are taken as the
  // difference of the quotients (`span`), never as Infinity.
  let k = 0, lo = 0, h = 0, single = false
  const span = (a: number, b: number, by: number) => { const q = (b - a) / by; return Number.isFinite(q) ? q : b / by - a / by }
  if (n > 0) {
    lo = U ?? min
    const hi = O ?? max
    if (hi === lo) { k = 1; single = true }
    else if (binning.binCount !== undefined) { k = binning.binCount - outer; h = span(lo, hi, k) }
    else {
      h = binning.binSize ?? scott(visible, take, under, over, n, Math.max(Math.abs(min), Math.abs(max)))
      if (!Number.isFinite(h) || !(h > 0)) return HISTOGRAM_INVALID_BINNING
      k = Math.max(1, Math.ceil(span(lo, hi, h)))
    }
    if (!(k >= 1) || !Number.isFinite(k)) return HISTOGRAM_INVALID_BINNING
  }
  const total = k + outer
  // Spec 7: more than 1,000 bins (outer included) is refused, before anything is allocated.
  if (total > HISTOGRAM_BIN_LIMIT) return HISTOGRAM_INVALID_BINNING
  charge((k + 1) * SLOT_BYTES)
  const edges = new Float64Array(k + 1)
  for (let i = 0; i <= k; i++) {
    const plain = lo + i * h
    // Halved and doubled only when the plain edge passes the double range (a subnormal width keeps its value).
    edges[i] = single ? lo : Number.isFinite(plain) ? plain : 2 * (lo / 2 + i * (h / 2))
  }
  if (binning.binCount !== undefined && !single && k) edges[k] = O ?? max
  // Validated geometry: finite, strictly increasing edges (an edge lost to precision or range is invalid binning).
  for (let i = 0; i <= k; i++) if (!Number.isFinite(edges[i]) || (!single && i && !(edges[i] > edges[i - 1]))) return HISTOGRAM_INVALID_BINNING
  const binOf = (v: number): number => {
    if (k === 1) return 0
    const q = span(lo, v, h)
    let i = right ? Math.ceil(q) - 1 : Math.floor(q)
    i = Number.isFinite(i) ? Math.max(0, Math.min(k - 1, i)) : 0
    if (right) { while (i > 0 && v <= edges[i]) i--; while (i < k - 1 && v > edges[i + 1]) i++ }
    else { while (i > 0 && v < edges[i]) i--; while (i < k - 1 && v >= edges[i + 1]) i++ }
    return i
  }
  const first = U !== undefined ? 1 : 0
  charge(total * 8 * visible.length)
  const heights = visible.map(s => {
    const counts = new Float64Array(total)
    for (let i = 0; i < s.values.length; i++) {
      if (!take(s, i)) continue
      const v = s.values.values[i]
      if (under(v)) counts[0]++
      else if (over(v)) counts[total - 1]++
      else if (k) counts[first + binOf(v)]++
    }
    return counts
  })
  // Labels (map §5.1 step 8), edges formatted with the category axis format, else the source format, else General.
  const formatter = createChartFormatter(workspace)
  const code = labelFormat(model, group, visible)
  const epoch = chartEpoch(model.epoch, model.workbookEpoch)
  const fmt = (v: number) => formatter.format(v, code, epoch)
  const labels: string[] = []
  let retained = 0
  const push = (build: () => string) => {
    charge(LABEL_BOUND)
    const text = build()
    labels.push(text)
    retained += stringBytes(text)
  }
  if (U !== undefined) push(() => `${right ? '≤' : '<'}${fmt(U)}`)
  for (let i = 0; i < k; i++) {
    const open = right ? !(i === 0 && U === undefined) : false
    const close = right ? true : i === k - 1 && O === undefined
    push(() => `${open ? '(' : '['}${fmt(edges[i])}, ${fmt(edges[i + 1])}${close ? ']' : ')'}`)
  }
  if (O !== undefined) push(() => `${right ? '>' : '≥'}${fmt(O)}`)
  // A bin is never one source point: overrides are not carried (map §5.1 [P]).
  return { labels, heights, sources: visible.map(() => undefined), retained: retained + labels.length * SLOT_BYTES }
}
/** Scott's width `3.49 s n^(-1/3)` (sample s of the inner set), its moments on values scaled to at most 1. */
function scott(visible: ChartSeries[], take: (s: ChartSeries, i: number) => boolean, under: (v: number) => boolean, over: (v: number) => boolean, n: number, maxAbs: number): number {
  if (n < 2) return NaN
  const scale = maxAbs > 0 ? 2 ** Math.min(1023, Math.ceil(Math.log2(maxAbs))) : 1
  let count = 0, mean = 0, m2 = 0
  for (const s of visible) for (let i = 0; i < s.values.length; i++) {
    if (!take(s, i)) continue
    const v = s.values.values[i]
    if (under(v) || over(v)) continue
    const x = v / scale
    count++
    const d = x - mean
    mean += d / count
    m2 += d * (x - mean)
  }
  return 3.49 * Math.sqrt(m2 / (n - 1)) * n ** (-1 / 3) * scale
}
/** The bin-label number format: the category axis's authored format, else the first value point's source format. */
function labelFormat(model: ChartModel, group: ChartExGroup, visible: ChartSeries[]): string | undefined {
  const id = group.binding.kind === 'cartesian' ? group.binding.category : undefined
  const axis = model.axes.find(a => a.id === id)
  if (axis?.formatCode && axis.sourceLinked !== true) return axis.formatCode
  for (const s of visible) for (let i = 0; i < s.values.length; i++) if (s.values.states[i] === POINT_VALUE) {
    const k = s.values.formatIndex?.[i]
    return k ? s.values.formats?.[k - 1] : undefined
  }
  return undefined
}

/**
 * Pareto (map §5.2): the owner's bars descending, ties in prior order (the source mapping moves with them); the
 * cumulative fraction of the total, on bars scaled by a power of two near their largest magnitude so the total cannot overflow.
 */
function sortPareto(derived: Derived, charge: (bytes: number) => void): (Derived & { cumulative: Float64Array }) | string {
  const owner = derived.heights[0] ?? new Float64Array(derived.labels.length)
  const n = derived.labels.length
  charge(n * 8 * (derived.heights.length + 2) + n * 4 * derived.sources.length)
  const order = Array.from({ length: n }, (_, i) => i).sort((a, b) => owner[b] - owner[a] || a - b)
  const labels = order.map(i => derived.labels[i])
  const heights = derived.heights.map(h => Float64Array.from(order, i => h[i]))
  const sources = derived.sources.map(src => src && Int32Array.from(order, i => src[i]))
  // Plain running sums when they stay in range; otherwise bars divided by a power of two near their largest magnitude.
  let plain = 0, inRange = true
  for (let i = 0; i < n; i++) { plain += heights[0]?.[i] ?? 0; if (!Number.isFinite(plain)) inRange = false }
  let largest = 0
  if (!inRange) for (let i = 0; i < n; i++) largest = Math.max(largest, Math.abs(owner[i]))
  const unit = !inRange && largest > 0 ? 2 ** Math.min(1023, Math.ceil(Math.log2(largest))) : 1
  let total = 0
  for (let i = 0; i < n; i++) total += owner[i] / unit
  const cumulative = new Float64Array(n)
  let running = 0
  for (let i = 0; i < n; i++) { running += (heights[0]?.[i] ?? 0) / unit; cumulative[i] = total > 0 ? running / total : 0 }
  if (cumulative.some(v => !Number.isFinite(v))) return HISTOGRAM_OUT_OF_RANGE
  return { labels, heights, sources, cumulative, retained: derived.retained }
}

const points = (values: Float64Array, built: Derived, charge: (bytes: number) => void): ChartPoints => {
  charge(pointBytes(values.length))
  built.retained += pointBytes(values.length)
  const out = new PointWriter(values.length)
  for (let i = 0; i < values.length; i++) out.value(i, values[i])
  return out.done()
}
const PARETO_CATEGORY = 'pareto-category', PARETO_PERCENT = 'pareto-percent', IMPLICIT_CATEGORY = 'histogram-category', IMPLICIT_VALUE = 'histogram-value'
/** Whether a series carries any point-level override (fill, colour, line, point label). */
const hasOverrides = (s: ChartSeries) => Boolean(s.dPt?.size || s.pointFills?.some(Boolean) || s.pointColors?.some(Boolean) || (s.labels?.points && Object.keys(s.labels.points).length))

/**
 * A derived series' point overrides (map §4.2): carried column by column where each column is one source point of
 * the series (`source[j]`: its source index); otherwise none.
 */
function remapOverrides(s: ChartSeries, source: Int32Array | undefined, built: Derived, charge: (bytes: number) => void): { series: Partial<ChartSeries>; dropped: number } {
  const { points: labelPoints, ...labels } = s.labels ?? {}
  const out: Partial<ChartSeries> = s.labels ? { labels } : {}
  if (!hasOverrides(s)) return { series: out, dropped: 0 }
  // Every source index an override addresses; those no single column carries are dropped (and counted).
  const addressed = new Set<number>([...s.dPt?.keys() ?? [], ...Object.keys(labelPoints ?? {}).map(Number)])
  s.pointFills?.forEach((fill, i) => { if (fill) addressed.add(i) })
  s.pointColors?.forEach((color, i) => { if (color) addressed.add(i) })
  if (!source) return { series: out, dropped: addressed.size }
  const carried = new Set<number>()
  for (let j = 0; j < source.length; j++) if (source[j] >= 0) carried.add(source[j])
  let dropped = 0
  for (const i of addressed) if (!carried.has(i)) dropped++
  const dPt = new Map<number, ChartPointOverride>(), fills: ChartSeries['pointFills'] = [], colors: string[] = [], pointLabels: Record<number, LabelOptions> = {}
  let kept = 0
  for (let j = 0; j < source.length; j++) {
    const i = source[j]
    if (i < 0) continue
    const point = s.dPt?.get(i), fill = s.pointFills?.[i], color = s.pointColors?.[i], label = labelPoints?.[i]
    if (!point && !fill && !color && !label) continue
    // The entry and slots are new; the override records are the parsed ones (shared, already owned).
    charge(4 * SLOT_BYTES)
    kept += 4 * SLOT_BYTES
    if (point) dPt.set(j, point)
    if (fill && j < CHART_POINT_LIMIT) fills[j] = fill
    if (color && j < CHART_POINT_LIMIT) colors[j] = color
    if (label) pointLabels[j] = label
  }
  built.retained += kept
  if (dPt.size) out.dPt = dPt
  if (fills.some(Boolean)) out.pointFills = fills
  if (colors.some(Boolean)) out.pointColors = colors
  if (Object.keys(pointLabels).length) out.labels = { ...labels, points: pointLabels }
  return { series: out, dropped }
}

/** The derived model: a plain clustered column group (and a Pareto line group), the labels as categories. */
function assemble(model: ChartModel, group: ChartExGroup, visible: ChartSeries[], built: Derived & { cumulative?: Float64Array }, pareto: ChartExSeriesInfo | undefined,
  charge: (bytes: number) => void): ChartModel {
  let axes = model.axes
  let binding = group.binding
  // A part without axes still draws on an implied primary pair (A1a `implicit`).
  if (binding.kind !== 'cartesian' || !binding.category) {
    axes = [...axes, { id: IMPLICIT_CATEGORY, kind: 'category', deleted: false, gridlines: false, implicit: true }, { id: IMPLICIT_VALUE, kind: 'value', deleted: false, gridlines: true, implicit: true }]
    binding = { kind: 'cartesian', category: IMPLICIT_CATEGORY, value: IMPLICIT_VALUE }
  }
  let dropped = 0
  const series: ChartSeries[] = visible.map((s, i) => {
    const remapped = remapOverrides(s, built.sources[i], built, charge)
    dropped += remapped.dropped
    return {
      name: s.name, color: s.color, values: points(built.heights[i], built, charge),
      ...(s.fill ? { fill: s.fill } : {}), ...(s.lineStyle ? { lineStyle: s.lineStyle } : {}), ...remapped.series,
    }
  })
  // Overrides that address a source point no single column draws are dropped, diagnosed once (map §4.2, §5.1 [P]).
  const diagnostics = [...model.diagnostics ?? []]
  if (dropped) {
    const reason = `${dropped} point ${dropped === 1 ? 'override is' : 'overrides are'} not drawn: a binned column, or a category grouped from several points, is not one source point`
    charge(stringBytes(HISTOGRAM_OVERRIDES_DROPPED) + stringBytes(reason))
    built.retained += stringBytes(HISTOGRAM_OVERRIDES_DROPPED) + stringBytes(reason)
    diagnostics.push({ code: HISTOGRAM_OVERRIDES_DROPPED, reason })
  }
  // Gap (map §5.1 step 9, §5.2): as authored; absent, 0 for binned columns (the histogram look) [P].
  const gapWidth = group.chartEx.gapWidth === undefined && !group.chartEx.byCategory ? 0 : group.gapWidth
  const groups: ChartGroup[] = [{ family: 'bar', barDir: 'col', grouping: 'clustered', series, binding, gapWidth, overlap: 0 }]
  if (pareto && built.cumulative) {
    const value = binding.value
    // The line's own axis: the value axis its axisId names, else an implied percent axis.
    const named = axes.find(a => a.kind === 'value' && a.id !== value && pareto.axisIds.includes(a.id))
    const percent: ChartAxis = { ...(named ?? { id: PARETO_PERCENT, kind: 'value', deleted: false, gridlines: false, unit: 'percentage' }), crossAx: PARETO_CATEGORY, crosses: 'max' }
    percent.min ??= 0
    percent.max ??= 1
    if (!percent.formatCode) percent.formatCode = '0%'
    axes = [...axes.filter(a => a !== named), percent, { id: PARETO_CATEGORY, kind: 'category', deleted: true, gridlines: false, crossAx: percent.id }]
    groups.push({
      family: 'line', grouping: 'standard', binding: { kind: 'cartesian', category: PARETO_CATEGORY, value: percent.id }, gapWidth: 150, overlap: 0,
      series: [{ name: visible[0]?.name ?? 'Pareto', color: pareto.color ?? '000000', values: points(built.cumulative, built, charge), line: true, markers: false, ...(pareto.lineStyle ? { lineStyle: pareto.lineStyle } : {}) }],
    })
  }
  const { categoryPending: _p, categoryRef: _r, categoryHidden: _h, categoryValues: _v, preparedAxes: _a, diagnostics: _d, ...rest } = model
  void _p, void _r, void _h, void _v, void _a, void _d
  const truncated = model.truncated || built.labels.length > CHART_POINT_LIMIT
  return { ...rest, groups, axes, categories: built.labels, derivedFamily: group.family, ...(diagnostics.length ? { diagnostics } : {}), ...(truncated ? { truncated: true } : {}) }
}
