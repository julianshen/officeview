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
import type { ChartAxis, ChartExGroup, ChartExSeriesInfo, ChartGroup, ChartModel, ChartPoints, ChartSeries } from '../model'
import type { ChartWorkspace } from '../workspace'
import { ChartRefusal, succeedChartModel } from '../own'
import { PointWriter, POINT_UNAVAILABLE, POINT_VALUE, hiddenAt, pointBytes, stringBytes } from '../points'
import { createChartFormatter } from '../format'
import { chartEpoch } from '../dates'
import { CHART_POINT_LIMIT } from '../parse'
import { CHART_DATA_NOT_AVAILABLE } from '../fill'

export const HISTOGRAM_FAMILIES: ReadonlySet<string> = new Set(['histogram', 'pareto'])
export const HISTOGRAM_INVALID_BINNING = 'invalid binning'
export const HISTOGRAM_UNREADABLE = 'This chart could not be read'
export const HISTOGRAM_BIN_LIMIT = 1000
const CATEGORY_TYPES = ['cat', 'entityId']
/** Retained record prices (a derived group, series or axis record) and a transient aggregation-map entry. */
const RECORD_BYTES = 512, MAP_ENTRY_BYTES = 96, SLOT_BYTES = 8
/** The bound of one built bin label: two formatted edges (each at most 256 units) and the brackets. */
const LABEL_BOUND = stringBytes(' '.repeat(2 * 256 + 4))

export const isHistogramGroup = (group: ChartGroup): group is ChartExGroup => group.kind === 'chartEx' && HISTOGRAM_FAMILIES.has(group.family)

/** A structural reason the chart cannot be drawn (data-independent; the adapter keeps the fallback), or undefined. */
export function histogramStructure(model: ChartModel): string | undefined {
  const group = model.groups[0]
  if (!group || !isHistogramGroup(group)) return undefined
  const { series: infos, paretoLine, byCategory, data } = group.chartEx
  // A Pareto sorts one owner: the column series its ownerIdx names (map §2.3, §5.2); several columns have no
  // single order.
  if (paretoLine && (infos.length !== 1 || infos[0].position !== paretoLine.ownerIdx)) return HISTOGRAM_UNREADABLE
  // By Category needs the categories it groups (map §5.1 [P]).
  if (byCategory && infos.some(info => !data.find(d => d.id === info.dataId)?.dimensions.some(d => d.kind === 'str' && CATEGORY_TYPES.includes(d.type)))) return HISTOGRAM_UNREADABLE
  // One set of bins serves every series: their binning must agree.
  const key = (info: ChartExSeriesInfo) => JSON.stringify(info.layoutPr.binning ?? null)
  if (infos.some(info => key(info) !== key(infos[0]))) return HISTOGRAM_UNREADABLE
  return undefined
}

interface Derived { labels: string[]; heights: Float64Array[]; retained: number }

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
    const visible = group.series.filter((_, i) => !group.chartEx.series[i]?.hidden)
    const visibleOnly = model.plotVisOnly !== false
    // Analytic inputs (map §2.4): an unavailable point refuses the chart; blank and error points are excluded.
    for (const s of visible) for (let i = 0; i < s.values.length; i++) if (s.values.states[i] === POINT_UNAVAILABLE) return refuse(model, CHART_DATA_NOT_AVAILABLE)
    const take = (s: ChartSeries, i: number) => s.values.states[i] === POINT_VALUE && !(visibleOnly && hiddenAt(s.values.hidden, i))
    const derived = group.chartEx.byCategory ? aggregate(model, visible, take, charge) : bin(model, group, visible, take, charge, workspace)
    if (typeof derived === 'string') return refuse(model, derived)
    const pareto = group.chartEx.paretoLine
    const built = pareto ? sortPareto(derived, charge) : { ...derived, cumulative: undefined }
    // The derived records (groups, series, axes) and point storage, then the move of every retained byte to the owner.
    const records = RECORD_BYTES * (visible.length + 4)
    charge(records)
    const out = assemble(model, group, visible, built, pareto, charge)
    if (ownerTag) workspace.transfer(scratch, ownerTag, built.retained + records)
    return out
  } finally { workspace.release(scratch) }
}
let jobs = 0
const refuse = (model: ChartModel, reason: string): ChartModel => ({ ...model, preparedAxes: undefined, unsupported: reason, refused: true })

/** By Category (map §5.1): equal category texts grouped in first-occurrence order (pooled across series), values summed. */
function aggregate(model: ChartModel, visible: ChartSeries[], take: (s: ChartSeries, i: number) => boolean, charge: (bytes: number) => void): Derived {
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
  charge(labels.length * 8 * visible.length)
  const heights = visible.map(s => {
    const sums = new Float64Array(labels.length)
    const categories = s.categories ?? model.categories
    for (let i = 0; i < s.values.length; i++) if (take(s, i)) sums[index.get(categories[i] ?? '')!] += s.values.values[i]
    return sums
  })
  return { labels, heights, retained: labels.length * SLOT_BYTES }
}

/** Binned (map §5.1): pooled bins over every visible series; a string is the refusal reason. */
function bin(model: ChartModel, group: ChartExGroup, visible: ChartSeries[], take: (s: ChartSeries, i: number) => boolean,
  charge: (bytes: number) => void, workspace: ChartWorkspace): Derived | string {
  const binning = group.chartEx.series[0]?.layoutPr.binning ?? { intervalClosed: 'r' as const }
  const right = binning.intervalClosed !== 'l'
  const U = typeof binning.underflow === 'number' ? binning.underflow : undefined
  const O = typeof binning.overflow === 'number' ? binning.overflow : undefined
  if (binning.underflow === 'auto' || binning.overflow === 'auto') return HISTOGRAM_UNREADABLE
  const under = (v: number) => U !== undefined && (right ? v <= U : v < U)
  const over = (v: number) => O !== undefined && (right ? v > O : v >= O)
  // The inner set's extent and moments (Welford), in one pass over every sample.
  let n = 0, mean = 0, m2 = 0, min = Infinity, max = -Infinity
  for (const s of visible) for (let i = 0; i < s.values.length; i++) {
    if (!take(s, i)) continue
    const v = s.values.values[i]
    if (under(v) || over(v)) continue
    n++
    const d = v - mean
    mean += d / n
    m2 += d * (v - mean)
    if (v < min) min = v
    if (v > max) max = v
  }
  const outer = (U !== undefined ? 1 : 0) + (O !== undefined ? 1 : 0)
  // Range, width and inner-bin count (map §5.1 steps 3-6).
  let k = 0, lo = 0, h = 0, single = false
  if (n > 0) {
    lo = U ?? min
    const hi = O ?? max
    if (hi === lo) { k = 1; single = true }
    else if (binning.binCount !== undefined) { k = binning.binCount - outer; h = (hi - lo) / k }
    else {
      h = binning.binSize ?? 3.49 * Math.sqrt(m2 / (n - 1)) * n ** (-1 / 3)
      if (!Number.isFinite(h) || !(h > 0)) return HISTOGRAM_INVALID_BINNING
      k = Math.max(1, Math.ceil((hi - lo) / h))
    }
    if (!(k >= 1) || !Number.isFinite(k)) return HISTOGRAM_INVALID_BINNING
  }
  const total = k + outer
  // Spec 7: more than 1,000 bins (outer included) is refused, before anything is allocated.
  if (total > HISTOGRAM_BIN_LIMIT) return HISTOGRAM_INVALID_BINNING
  charge((k + 1) * SLOT_BYTES)
  const edges = new Float64Array(k + 1)
  for (let i = 0; i <= k; i++) edges[i] = single ? lo : lo + i * h
  if (binning.binCount !== undefined && !single && k) edges[k] = O ?? max
  const binOf = (v: number): number => {
    if (k === 1) return 0
    let i = right ? Math.ceil((v - lo) / h) - 1 : Math.floor((v - lo) / h)
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
  return { labels, heights, retained: retained + labels.length * SLOT_BYTES }
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

/** Pareto (map §5.2): the owner's bars descending, ties in prior order; the cumulative fraction of the total. */
function sortPareto(derived: Derived, charge: (bytes: number) => void): Derived & { cumulative: Float64Array } {
  const owner = derived.heights[0] ?? new Float64Array(derived.labels.length)
  const n = derived.labels.length
  charge(n * 8 * (derived.heights.length + 2))
  const order = Array.from({ length: n }, (_, i) => i).sort((a, b) => owner[b] - owner[a] || a - b)
  const labels = order.map(i => derived.labels[i])
  const heights = derived.heights.map(h => Float64Array.from(order, i => h[i]))
  let total = 0
  for (let i = 0; i < n; i++) total += owner[i]
  const cumulative = new Float64Array(n)
  let running = 0
  for (let i = 0; i < n; i++) { running += heights[0]?.[i] ?? 0; cumulative[i] = total > 0 ? running / total : 0 }
  return { labels, heights, cumulative, retained: derived.retained }
}

const points = (values: Float64Array, built: Derived, charge: (bytes: number) => void): ChartPoints => {
  charge(pointBytes(values.length))
  built.retained += pointBytes(values.length)
  const out = new PointWriter(values.length)
  for (let i = 0; i < values.length; i++) out.value(i, values[i])
  return out.done()
}
const PARETO_CATEGORY = 'pareto-category', PARETO_PERCENT = 'pareto-percent', IMPLICIT_CATEGORY = 'histogram-category', IMPLICIT_VALUE = 'histogram-value'

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
  const series: ChartSeries[] = visible.map((s, i) => {
    const { points: _points, ...labels } = s.labels ?? {}
    void _points
    return {
      name: s.name, color: s.color, values: points(built.heights[i], built, charge),
      ...(s.fill ? { fill: s.fill } : {}), ...(s.lineStyle ? { lineStyle: s.lineStyle } : {}), ...(s.labels ? { labels } : {}),
    }
  })
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
  const { categoryPending: _p, categoryRef: _r, categoryHidden: _h, categoryValues: _v, preparedAxes: _a, ...rest } = model
  void _p, void _r, void _h, void _v, void _a
  const truncated = model.truncated || built.labels.length > CHART_POINT_LIMIT
  return { ...rest, groups, axes, categories: built.labels, ...(truncated ? { truncated: true } : {}) }
}

/**
 * The adapter's stage 1 for a histogram or Pareto whose data is complete (cached, or read from the embedded
 * workbook): structural checks, then `prepare` (which derives). The successor takes over the model's charge; a
 * memory refusal refuses the chart.
 */
export function prepareHistogramModel(model: ChartModel, prepare: (model: ChartModel) => ChartModel): ChartModel {
  let next: ChartModel
  try { next = prepare(model) } catch (error) {
    if (!(error instanceof ChartRefusal)) throw error
    next = { ...model, preparedAxes: undefined, unsupported: 'This chart needs more memory than is available for charts', refused: true }
  }
  succeedChartModel(model, next)
  return next
}
