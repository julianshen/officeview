/** Accessible chart descriptions (small; shared by the lazy chart chunk and the viewer). */
import type { XlsxChartModel } from './chartModel'

export function formatNumber(value: number, percent = false): string {
  if (percent) return `${Number((value * 100).toPrecision(6))}%`
  const rounded = Number(value.toPrecision(10))
  const [int, frac] = String(Math.abs(rounded)).split('.')
  if (int.includes('e')) return String(rounded)
  const grouped = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  return `${rounded < 0 ? '-' : ''}${grouped}${frac ? '.' + frac : ''}`
}

const TYPE_NAMES: Record<string, string> = { bar: 'Bar chart', col: 'Column chart', line: 'Line chart', area: 'Area chart', pie: 'Pie chart', doughnut: 'Doughnut chart', scatter: 'Scatter chart', bubble: 'Bubble chart', radar: 'Radar chart', stock: 'Stock chart', treemap: 'Treemap chart', sunburst: 'Sunburst chart', boxWhisker: 'Box & whisker chart' }
/** Accessible description of a chart (also used for the placeholder card). */
export function chartDescription(model: XlsxChartModel): string {
  const name = model.title ? `Chart: ${model.title}.` : 'Chart.'
  if (model.unsupported) return `${name} ${model.unsupported}; it is not drawn.`
  const group = model.groups.find(g => g.family === 'stock') ?? model.groups[0]
  const kind = group.family === 'bar' ? TYPE_NAMES[group.barDir === 'bar' ? 'bar' : 'col'] : TYPE_NAMES[group.family]
  const grouping = group.grouping === 'stacked' ? 'stacked ' : group.grouping === 'percentStacked' ? '100% stacked ' : ''
  const series = model.groups.flatMap(g => g.series)
  const parts = [`${name} ${grouping ? grouping + kind.toLowerCase() : kind}, ${series.length} series`]
  if (model.categories.length) parts.push(`${model.categories.length} categories${model.categories.length <= 12 ? ` (${model.categories.join(', ')})` : ''}`)
  for (const s of series.slice(0, 8)) {
    // Read from the compact storage (state 1: a value): the first 12 values, else the range.
    const first: number[] = []
    let count = 0, min = Infinity, max = -Infinity
    for (let i = 0; i < s.values.length; i++) if (s.values.states[i] === 1) { const v = s.values.values[i]; if (count++ < 12) first.push(v); min = Math.min(min, v); max = Math.max(max, v) }
    parts.push(count ? `${s.name}: ${count <= 12 ? first.map(n => formatNumber(n)).join(', ') : `${formatNumber(min)} to ${formatNumber(max)}`}` : `${s.name}: no data`)
  }
  return parts.join('; ') + (model.truncated ? '; data truncated.' : '.')
}
