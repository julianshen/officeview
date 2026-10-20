/**
 * The display model (lazy chunk): what a chart draws from its filled model.
 *
 * - `plotVisOnly` (spec 5.8; Office default on): points whose worksheet cell is
 *   in a hidden row or column (provenance recorded at fill: `hidden`,
 *   `categoryHidden`) are left out of the drawing, categories included; a
 *   series whose every point is hidden is left out. With `plotVisOnly=0` they
 *   are drawn. Saved (cached) data carries no provenance and is drawn as saved.
 * - The display bound (spec 7): at most CHART_POINT_LIMIT points per series
 *   are drawn. Analytic series keep every acquired point in the model; only
 *   the drawing is capped.
 *
 * The filled model itself is never changed: analysis reads all of it.
 */
import type { ChartModel, ChartSeries } from './model'
import { CHART_POINT_LIMIT } from './parse'

export function displayModel(model: ChartModel): ChartModel {
  const visibleOnly = model.plotVisOnly !== false
  const allHidden = (s: ChartSeries) => Boolean(s.hidden?.length && s.values.length && s.hidden.length >= s.values.length)
  const skip = new Set<number>(visibleOnly ? model.categoryHidden : [])
  if (visibleOnly) for (const group of model.groups) for (const s of group.series) if (!allHidden(s)) for (const i of s.hidden ?? []) skip.add(i)
  const over = model.groups.some(g => g.series.some(s => s.values.length > CHART_POINT_LIMIT || (s.xValues?.length ?? 0) > CHART_POINT_LIMIT))
  const dropsSeries = visibleOnly && model.groups.some(g => g.series.some(allHidden))
  if (!skip.size && !over && !dropsSeries) return model
  const keep = <T>(items: T[]) => (skip.size ? items.filter((_, i) => !skip.has(i)) : items).slice(0, CHART_POINT_LIMIT)
  return {
    ...model,
    categories: skip.size ? model.categories.filter((_, i) => !skip.has(i)) : model.categories,
    groups: model.groups.map(group => ({
      ...group,
      series: group.series.filter(s => !(visibleOnly && allHidden(s))).map(s => ({
        ...s,
        values: keep(s.values),
        ...(s.xValues ? { xValues: keep(s.xValues) } : {}),
        ...(s.pointColors && skip.size ? { pointColors: s.pointColors.filter((_, i) => !skip.has(i)) } : {}),
      })),
    })),
  }
}
