/**
 * Accessible chart descriptions (spec 6.6). Each one is bounded and names the chart type, title, series names and
 * value range; a placeholder card says `placeholder: <reason>`. It is shared by the engine, the sheet viewer and
 * the paged viewer's per-page chart list (through the paintables metadata). Never imported eagerly.
 */
import type { ChartModel } from './model'

export function formatNumber(value: number, percent = false): string {
  if (percent) return `${Number((value * 100).toPrecision(6))}%`
  const rounded = Number(value.toPrecision(10))
  const [int, frac] = String(Math.abs(rounded)).split('.')
  if (int.includes('e')) return String(rounded)
  const grouped = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  return `${rounded < 0 ? '-' : ''}${grouped}${frac ? '.' + frac : ''}`
}

/** The longest description (characters); longer text is cut with an ellipsis. */
export const CHART_DESCRIPTION_LIMIT = 1024
/** The longest title or series name kept in a description. */
const NAME_LIMIT = 120
const TYPE_NAMES: Record<string, string> = {
  bar: 'Bar chart', col: 'Column chart', line: 'Line chart', area: 'Area chart', pie: 'Pie chart', doughnut: 'Doughnut chart',
  ofPie: 'Pie of pie chart', scatter: 'Scatter chart', bubble: 'Bubble chart', radar: 'Radar chart', stock: 'Stock chart', contour: 'Surface chart',
  waterfall: 'Waterfall chart', histogram: 'Histogram chart', pareto: 'Pareto chart', boxWhisker: 'Box & whisker chart', funnel: 'Funnel chart',
  treemap: 'Treemap chart', sunburst: 'Sunburst chart',
}
const clip = (text: string, limit: number) => text.length > limit ? `${text.slice(0, limit - 1)}…` : text

/** The chart's type name ("Column chart"), or undefined for a card. */
export function chartTypeName(model: ChartModel): string | undefined {
  if (model.unsupported) return undefined
  const group = model.groups.find(g => g.family === 'stock') ?? model.groups[0]
  if (!group) return undefined
  // A histogram or Pareto is drawn as plain column (and line) groups after stage 1; it keeps its own name.
  if (model.derivedFamily) return TYPE_NAMES[model.derivedFamily] ?? 'Chart'
  return (group.family === 'bar' ? TYPE_NAMES[group.barDir === 'bar' ? 'bar' : 'col'] : TYPE_NAMES[group.family]) ?? 'Chart'
}

/** Accessible description of a chart (also used for the placeholder card). */
export function chartDescription(model: ChartModel): string {
  const name = model.title ? `Chart: ${clip(model.title, NAME_LIMIT)}.` : 'Chart.'
  const kind = chartTypeName(model)
  if (model.unsupported || !kind) return clip(`${name} placeholder: ${model.unsupported ?? 'This chart has no data to draw'}.`, CHART_DESCRIPTION_LIMIT)
  const group = model.groups.find(g => g.family === 'stock') ?? model.groups[0]
  const grouping = group.grouping === 'stacked' ? 'stacked ' : group.grouping === 'percentStacked' ? '100% stacked ' : ''
  // The first 8 series are read in place; the rest are only counted (no copy of the series list).
  let count = 0
  const series: typeof group.series = []
  for (const g of model.groups) for (const s of g.series) if (count++ < 8) series.push(s)
  const parts = [`${name} ${grouping ? grouping + kind.toLowerCase() : kind}, ${count} series`]
  if (model.categories.length) parts.push(`${model.categories.length} categories${model.categories.length <= 12 ? ` (${model.categories.map(c => clip(c, NAME_LIMIT)).join(', ')})` : ''}`)
  for (const s of series) {
    // Read from the compact storage (state 1: a value): the first 12 values, else the range.
    const first: number[] = []
    let n = 0, min = Infinity, max = -Infinity
    const values = s.values
    for (let i = 0; i < (values?.length ?? 0); i++) if (values.states[i] === 1) { const v = values.values[i]; if (n++ < 12) first.push(v); min = Math.min(min, v); max = Math.max(max, v) }
    const label = clip(s.name, NAME_LIMIT)
    parts.push(n ? `${label}: ${n <= 12 ? first.map(v => formatNumber(v)).join(', ') : `${formatNumber(min)} to ${formatNumber(max)}`}` : `${label}: no data`)
  }
  if (count > 8) parts.push(`${count - 8} more series`)
  return clip(parts.join('; ') + (model.truncated ? '; data truncated.' : '.'), CHART_DESCRIPTION_LIMIT)
}
