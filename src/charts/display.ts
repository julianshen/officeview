/**
 * The display view (lazy chunk): what a chart draws from its filled model,
 * read through the compact storage without copying it (no per-point arrays).
 *
 * - The display bound (spec 7): at most CHART_POINT_LIMIT source indices are
 *   drawn. Analytic series keep every acquired point in the model; only the
 *   drawing is capped.
 * - `plotVisOnly` (spec 5.8; Office default on), per role and by each role's
 *   own source provenance, never a union across series:
 *   - a category whose own cell is hidden leaves the category axis (its slot
 *     and the points paired with it by index);
 *   - a point whose own value cell (or, for scatter, its x cell) is hidden is
 *     not drawn; a series whose every value point is hidden is not drawn;
 *   - `plotVisOnly=0` draws everything. Saved data without a worksheet
 *     reference has no provenance and is drawn as saved.
 */
import type { ChartModel, ChartPoints, ChartSeries, ChartPoint } from './model'
import { CHART_POINT_LIMIT } from './parse'
import { hiddenAt, valueAt, pointAt } from './points'

export interface DisplaySeries {
  series: ChartSeries
  /** The plotted number at a slot, or undefined (a gap: not a value, hidden, or past the series). */
  value(slot: number): number | undefined
  /** The x value at a slot (scatter), with the same rules. */
  x(slot: number): number | undefined
  /** The series' x role, when it has one. */
  hasX: boolean
  point(slot: number, x?: boolean): ChartPoint | undefined
}
export interface DisplayView {
  /** Category-axis slots drawn (at least 1). */
  slots: number
  /** The source index (category / point index in the model) a slot draws. */
  source(slot: number): number
  /** The category label at a slot (undefined past the categories). */
  category(slot: number): string | undefined
  series: DisplaySeries[]
  /** The series of one group, in order. */
  groupSeries(group: number): DisplaySeries[]
}

/** The k-th visible index given the sorted hidden indices (binary search on hidden[j] - j). */
function sourceOf(hidden: Uint32Array | undefined, slot: number): number {
  if (!hidden?.length) return slot
  let lo = 0, hi = hidden.length
  while (lo < hi) { const mid = (lo + hi) >>> 1; if (hidden[mid] - mid <= slot) lo = mid + 1; else hi = mid }
  return slot + lo
}
const allHidden = (points: ChartPoints) => points.length > 0 && (points.hidden?.length ?? 0) >= points.length

export function displayView(model: ChartModel): DisplayView {
  const visibleOnly = model.plotVisOnly !== false
  let extent = model.categories.length
  for (const group of model.groups) for (const s of group.series) extent = Math.max(extent, s.values.length)
  extent = Math.min(extent, CHART_POINT_LIMIT)
  // Hidden categories within the drawn extent leave the axis.
  const categoryHidden = visibleOnly ? model.categoryHidden : undefined
  let hiddenCategories = 0
  for (const i of categoryHidden ?? []) if (i < extent) hiddenCategories++
  const slots = Math.max(extent - hiddenCategories, 1)
  const source = (slot: number) => sourceOf(categoryHidden, slot)
  const make = (s: ChartSeries): DisplaySeries => {
    const read = (points: ChartPoints | undefined, slot: number) => {
      const i = source(slot)
      if (!points || i >= extent || i >= points.length) return undefined
      if (visibleOnly && (hiddenAt(points.hidden, i) || (s.xValues && hiddenAt(s.xValues.hidden, i)) || hiddenAt(s.values.hidden, i))) return undefined
      return valueAt(points, i)
    }
    const point = (slot:number, x=false): ChartPoint | undefined => {
      const i=source(slot), points=x?s.xValues:s.values
      if(i>=extent || (visibleOnly && (hiddenAt(s.values.hidden,i)||hiddenAt(s.xValues?.hidden,i))))return undefined
      return points ? pointAt(points,i) : {state:'blank'}
    }
    return { point, series: s, value: slot => read(s.values, slot), x: slot => read(s.xValues, slot), hasX: Boolean(s.xValues?.length) }
  }
  const byGroup = model.groups.map(group => group.series.filter(s => !(visibleOnly && allHidden(s.values))).map(make))
  return {
    slots,
    source,
    category: slot => { const i = source(slot); return i < extent ? model.categories[i] : undefined },
    series: byGroup.flat(),
    groupSeries: group => byGroup[group] ?? [],
  }
}
