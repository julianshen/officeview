/** Size-independent stage 1. Data-only axis snapshots survive structuredClone; closures live in stage 2. */
import { resolveAxes, type AxisGraph, type ResolvedAxis } from './axes'
import { displayView, displayFamily, displayLabelBytes, type DisplayForecast, type DisplayView } from './display'
const NO_FORECAST: DisplayForecast = { lead: 0, trail: 0 }
import { dateUnitIndex } from './dates'
import { prepareDataLabels, labelPreparationBound, labelBytes, type PreparedLabel } from './labels'
import { createChartFormatter, axisFormat } from './format'
import { valueAt, stringBytes, POINT_UNAVAILABLE } from './points'
import type { ChartModel, ChartDiagnostic } from './model'
import { ChartRefusal, onChartModelRelease } from './own'
import { createChartWorkspace, type ChartWorkspace } from './workspace'
import { prepareTrendlines, preparedTrendlineBytes, trendlineBound, trendlineExtents, trendlineForecast, type PreparedTrendline } from './trendlines'
import { chartEpoch } from './dates'
import { deriveHistogram, isHistogramGroup } from './chartex/histogram'
import { BOX_WHISKER_TOO_MANY_CATEGORIES, boxWhiskerCategoryLabels, prepareBoxWhisker, preparedBoxWhiskerBytes, boxWhiskerBound, type PreparedBoxWhiskerGroup } from './chartex/boxWhisker'
import { CHART_DATA_NOT_AVAILABLE, CHART_TOO_MANY_POINTS } from './limits'
import { deriveChartHierarchy, formatHierarchyValues, hierarchyBytes, type ChartExHierarchy } from './chartex/hierarchy'
import { errorBarBound, prepareErrorBars, preparedErrorBarBytes, type PreparedErrorBars } from './errorBars'

type PreparedAxis = Omit<ResolvedAxis, 'labelValue' | 'plottable'> & { labels: string[] }
export interface PreparedAxes {
  dataLabels: PreparedLabel[]
  /** Fitted trendlines (spec 5.6); absent when the chart has none. */
  trendlines?: PreparedTrendline[]
  /** Category slots the trendline forecasts add (the display view reads them); absent: none. */
  forecast?: DisplayForecast
  /** Treemap/sunburst: the derived tree (chartex/hierarchy.ts; chartex-map §7). */
  hierarchy?: ChartExHierarchy
  /** Validated error bar sets with their analytic amounts (spec 5.7); absent when the chart has none. */
  errorBars?: PreparedErrorBars[]
  axes: PreparedAxis[]
  groups: Array<{ category?: string; x?: string; value: string; horizontal: string; vertical: string; between: boolean } | undefined>
  diagnostics: ChartDiagnostic[]
  epoch: AxisGraph['epoch']
  boxWhisker?: PreparedBoxWhiskerGroup[]
}
interface PreparedOwner { workspace: ChartWorkspace; tag: string }
const owners = new WeakMap<PreparedAxes, PreparedOwner>()
const collected = new FinalizationRegistry<PreparedOwner>(owner => owner.workspace.release(owner.tag))
onChartModelRelease((model, next) => {
  const prepared = model.preparedAxes
  if (!prepared || prepared === next?.preparedAxes) return
  const owner = owners.get(prepared)
  if (owner) { owners.delete(prepared); collected.unregister(prepared); owner.workspace.release(owner.tag) }
})
let jobs = 0
export function prepareChartAxes(model: ChartModel, owner?: { workspace: ChartWorkspace; tag: string }): ChartModel {
  if (model.unsupported) return model
  if (model.categoryPending || model.groups.some(g => g.series.some(s => s.pending))) {
    const { preparedAxes: _previous, ...pending } = model
    return pending
  }
  if (model.groups.some(g => g.family === 'boxWhisker' && g.series.some(s => s.values?.states?.includes(POINT_UNAVAILABLE)))) {
    return { ...model, unsupported: CHART_DATA_NOT_AVAILABLE, refused: true }
  }
  const workspace = owner?.workspace ?? createChartWorkspace('mobile'), tag = `chart-prepare#${++jobs}`
  // ChartEx histogram and Pareto (B3): binned or aggregated into plain column and line groups first.
  if (model.groups.some(isHistogramGroup)) { model = deriveHistogram(model, workspace, owner?.tag); if (model.unsupported) return model }
  const boxDiagnostics: ChartDiagnostic[] = []
  let slots = model.categories.length
  for (const g of model.groups) for (const s of g.series) slots = Math.max(slots, s.values.length)
  slots = Math.min(4096, Math.max(1, slots))
  // The live resolved graph and the data-only snapshot coexist while building.
  // Automatic linear scales use a 1/2/5 step around five intervals (at most 16 ticks including endpoints).
  // Date/log/authored units keep the full 1,000-tick bound. Categories reuse input strings, without formatting.
  const axesBound = model.axes.reduce((n, a) => {
    const numeric = a.kind === 'value' || a.kind === 'date'
    const ticks = !numeric ? 0 : a.kind === 'date' || a.logBase !== undefined || a.majorUnit !== undefined ? 1000 : 16
    const minor = a.minorUnit === undefined ? 0 : 1000
    return n + 2048 + ticks * 640 + minor * 16 + Math.min(slots, model.categories.length) * 32 + (a.kind === 'date' ? slots * 16 : 0) + ((a.id.length + (a.crossAx?.length ?? 0)) * 8)
  }, 0)
  let bound = 4096 + labelPreparationBound(model) + trendlineBound(model) + boxWhiskerBound(model) + errorBarBound(model, model.groups.some(g => g.series.some(s => s.trendlines)) ? 4096 : slots) + axesBound + model.groups.reduce((n, g) => n + 1024 + g.series.length * 512, 0)
    + model.axes.length * model.categories.reduce((n, s) => n + s.length * 2, 0) + (displayFamily(model) ? model.axes.length * displayLabelBytes(slots) : 0)
  if (!workspace.reserve(bound, tag)) throw new ChartRefusal()
  // Box & whisker statistics (B4): charged through the workspace before each allocation; scratch per series is
  // released as soon as the series is summarised, retained lists join the prepared record's charge.
  const scratchTag = `${tag}:box-scratch`
  try {
    let boxWhisker: PreparedBoxWhiskerGroup[] | undefined
    if (model.groups.some(g => g.family === 'boxWhisker')) {
      const prepared = prepareBoxWhisker(model, {
        retain: bytes => { if (!workspace.reserve(bytes, tag)) throw new ChartRefusal(); bound += bytes },
        scratch: bytes => { if (!workspace.reserve(bytes, scratchTag)) throw new ChartRefusal() },
        releaseScratch: () => workspace.release(scratchTag),
      }, boxDiagnostics)
      // More distinct categories than the display limit refuse the chart: a prefix would misstate the statistics.
      if (prepared === BOX_WHISKER_TOO_MANY_CATEGORIES) { const { preparedAxes: _previous, ...rest } = model; return { ...rest, unsupported: CHART_TOO_MANY_POINTS, refused: true } }
      boxWhisker = prepared
    }
    // Treemap/sunburst (B5): the tree is derived first, its storage admitted under this stage's tag; an analytic
    // refusal (too many nodes, an unavailable weight) refuses the chart.
    const hierarchy = deriveChartHierarchy(model, bytes => { if (!workspace.reserve(bytes, tag)) throw new ChartRefusal() })
    if (hierarchy && 'refused' in hierarchy) return { ...model, preparedAxes: undefined, unsupported: hierarchy.refused, refused: true }
    // Trendlines first (A6a, spec 5.6): their forecast adds category slots and their extent joins the automatic scales.
    const formatter = createChartFormatter(workspace)
    if (hierarchy) formatHierarchyValues(model, hierarchy.tree, (v, code) => formatter.format(v, code, chartEpoch(model.epoch, model.workbookEpoch)), bytes => { if (!workspace.reserve(bytes, tag)) throw new ChartRefusal() })
    const trendDiagnostics: ChartDiagnostic[] = []
    const trendlines = model.groups.some(g => g.series.some(s => s.trendlines)) ? prepareTrendlines(model, formatter, chartEpoch(model.epoch, model.workbookEpoch), trendDiagnostics, workspace) : []
    let forecast = trendlines.length ? trendlineForecast(model, trendlines, displayView(model, NO_FORECAST).slots, trendDiagnostics) : undefined
    // The forecast's slots and their category label records are admitted before the view, graph and snapshot are
    // built; refused, the forecast is clipped (diagnosed), never the chart.
    if (forecast) {
      const extra = forecastBytes(model, forecast)
      if (workspace.reserve(extra, tag)) bound += extra
      else { forecast = undefined; trendDiagnostics.push({ code: 'chart-trendline-extension-clamped', reason: 'A trendline forecast needs more memory than is available for charts: it is clipped to the data' }) }
    }
    const view = displayView(model, forecast ?? NO_FORECAST)
    // Error bars (A6b, spec 5.7): their analytic amounts, and the drawn bar ends join the automatic scales.
    const errors = model.groups.some(g => g.series.some(s => s.errorBars)) ? prepareErrorBars(model, view, trendDiagnostics) : undefined
    const extents = [...trendlines.length ? trendlineExtents(trendlines) : [], ...errors?.extents ?? []]
    const graph = resolveAxes(model, view, undefined, undefined, extents.length ? extents : undefined)
    if(model.groups.some(g=>g.series.some(s=>[s.values,s.xValues,s.bubbleSizes].some(p=>p?.states.includes(POINT_UNAVAILABLE)))) && !graph.diagnostics.some(d=>d.code==='chart-point-unavailable'))graph.diagnostics.push({code:'chart-point-unavailable',reason:'A required saved chart point is unavailable'})
    const axes = [...graph.axes.values()].map(a => {
      const inputs = graph.groups.flatMap((g,i) => !g ? [] : g.category === a ? model.categoryValues ? [model.categoryValues] : [] : g.x === a ? model.groups[i].series.flatMap(s=>s.xValues?[s.xValues]:[]) : g.value === a ? model.groups[i].series.map(s=>s.values) : [])
      const code = axisFormat(a.axis,inputs) ?? (a.percent ? '0%' : a.role==='date' ? a.dateBase==='years'?'yyyy':a.dateBase==='months'?'mmm yyyy':'m/d/yyyy' : undefined)
      const { labelValue, plottable: _plottable, ...data } = a
      const bw = graph.groups.find(g => g?.category === a)?.group
      const labels = a.role === 'value' ? a.ticks.map(t => formatter.format(labelValue(t), code, graph.epoch)) : a.role === 'date' ? a.ticks.map(t => formatter.format(t, code, graph.epoch)) : a.role === 'series' ? seriesLabels(graph, view, a) : bw?.family === 'boxWhisker' ? boxWhiskerCategoryLabels(model.categories, bw) : Array.from({ length: Math.min(a.slots, view.slots, (displayFamily(model) ? view.slots : model.categories.length) + (forecast?.lead ?? 0)) }, (_, i) => view.category(i) ?? '')
      return { ...data, labels }
    })
    // A hierarchy family places and draws its own node labels (chartex/treemap.ts, sunburst.ts).
    const dataLabels = hierarchy ? [] : prepareDataLabels(model,formatter,graph.epoch,graph.diagnostics,view)
    graph.diagnostics.push(...trendDiagnostics, ...(hierarchy?.diagnostics ?? []), ...boxDiagnostics)
    const preparedAxes: PreparedAxes = { dataLabels, ...(trendlines.length ? { trendlines } : {}), ...(hierarchy ? { hierarchy: hierarchy.tree } : {}), ...(forecast ? { forecast } : {}), ...(errors?.bars.length ? { errorBars: errors.bars } : {}), ...(boxWhisker?.length ? { boxWhisker } : {}), axes, groups: graph.groups.map(g => g ? { ...(g.category ? { category: g.category.id } : {}), ...(g.x ? { x: g.x.id } : {}), value: g.value.id, horizontal: g.horizontal.id, vertical: g.vertical.id, between: g.between } : undefined), diagnostics: graph.diagnostics, epoch: graph.epoch }
    const bytes = preparedAxesBytes(preparedAxes)
    if (bytes > bound && !workspace.reserve(bytes - bound, tag)) throw new ChartRefusal()
    const retained = { workspace, tag: `${tag}:retained` }
    workspace.transfer(tag, retained.tag, bytes)
    owners.set(preparedAxes, retained)
    collected.register(preparedAxes, retained, preparedAxes)
    return { ...model, preparedAxes }
  } finally { workspace.release(scratchTag); workspace.release(tag) }
}
/**
 * What `forecast` slots add to preparation: per slot and axis a category label record (a string table entry) and
 * its construction share, plus the view's per-slot work.
 */
const forecastBytes = (model: ChartModel, forecast: DisplayForecast): number => (forecast.lead + forecast.trail) * (model.axes.length * 64 + 64)
/** A series axis's labels: the names of the series it places (a contour group's drawn rows), reusing the model's strings. */
export function seriesLabels(graph: AxisGraph, view: DisplayView, a: ResolvedAxis): string[] {
  const g = graph.groups.findIndex(p => p?.vertical === a || p?.horizontal === a)
  return g < 0 ? [] : view.groupSeries(g).slice(0, a.slots).map(s => s.series.name)
}
export function hydrateAxes(model: ChartModel, view: DisplayView): AxisGraph {
  const prepared = model.preparedAxes
  if (!prepared) return resolveAxes(model, view) // bounded compatibility for a hand-built/headless model
  const axes = new Map<string, ResolvedAxis>()
  for (const data of prepared.axes) axes.set(data.id, { ...data, labelValue: t => t / (data.display?.divisor ?? 1), plottable: v => !data.log || v > 0 })
  return { axes, diagnostics: prepared.diagnostics, epoch: prepared.epoch,
    groups: prepared.groups.map((g, i) => g ? { group: model.groups[i], ...(g.category ? { category: axes.get(g.category)! } : {}), ...(g.x ? { x: axes.get(g.x)! } : {}), value: axes.get(g.value)!, horizontal: axes.get(g.horizontal)!, vertical: axes.get(g.vertical)!, between: g.between } : undefined),
    categoryIndex: (axis, slot) => {
      if (axis.role !== 'date' || !axis.dateBase) return slot
      const v = valueAt(model.categoryValues, view.source(slot))
      return v === undefined ? NaN : dateUnitIndex(v, axis.dateBase, prepared.epoch) - (axis.dateStart ?? 0)
    },
  }
}

/** Footprint of the transferable stage-1 records (same pricing as their owner). */
export function preparedAxesBytes(preparedAxes: PreparedAxes | undefined): number {
  if (!preparedAxes) return 0
  const axes = preparedAxes.axes, graph = preparedAxes
  return 256 + hierarchyBytes(preparedAxes.hierarchy) + labelBytes(preparedAxes.dataLabels??[]) + (preparedAxes.trendlines ?? []).reduce((n, t) => n + preparedTrendlineBytes(t), 0) + (preparedAxes.errorBars ?? []).reduce((n, b) => n + preparedErrorBarBytes(b), 0) + preparedBoxWhiskerBytes(preparedAxes.boxWhisker) + preparedAxes.groups.length * 128 + axes.reduce((n, a) => n + 1024 + (a.ticks.length + a.minorTicks.length) * 8 + a.labels.reduce((n, s) => n + stringBytes(s), 0), 0) + graph.diagnostics.reduce((n, d) => n + 64 + stringBytes(d.reason), 0)
}
