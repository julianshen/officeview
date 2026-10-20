/**
 * What the treemap and sunburst painters share with layout (lazy chunk): node colours, node label text, the label
 * font size, the strings layout measures (the paint reads only measured widths), the legend entries and the stage-2
 * byte bounds. Office documents neither family's drawing; everything here is a project rule [U] except the label
 * fields, which come from the authored `cx:dataLabels` (map §4.2).
 */
import type { ChartExGroup, ChartModel, ChartSeries } from '../model'
import type { ChartFillStyle } from '../fills'
import { formatNumber } from '../../xlsx/chartDescription'
import { isHierarchyGroup, type ChartExHierarchy } from './hierarchy'

export interface HierarchyDrawing { group: ChartExGroup; series: ChartSeries; tree: ChartExHierarchy }
/** The group, its drawn series (the first visible) and the derived tree, when the model is a prepared hierarchy. */
export function hierarchyDrawing(model: ChartModel): HierarchyDrawing | undefined {
  const group = model.groups[0], tree = model.preparedAxes?.hierarchy
  if (!isHierarchyGroup(group) || !tree) return undefined
  const visible = group.chartEx.series.findIndex(info => !info.hidden)
  const series = group.series[visible < 0 ? 0 : visible]
  return series ? { group, series, tree } : undefined
}
export const isLeaf = (tree: ChartExHierarchy, i: number): boolean => tree.end[i] === i + 1
/** The node label size in px: the label `txPr` size, else the chart's. */
export const hierarchyLabelPx = (model: ChartModel, series: ChartSeries): number => (series.labels?.text?.sizePt ?? model.fontSizePt) * 96 / 72
/** A node's data-label text from the authored fields (series name, category name, value; separator), '' when none. */
export function nodeText(series: ChartSeries, tree: ChartExHierarchy, i: number): string {
  const controls = series.labels
  if (!controls) return ''
  const parts: string[] = []
  if (controls.showSerName) parts.push(series.name)
  if (controls.showCatName && tree.labels[i]) parts.push(tree.labels[i])
  // The stage-1 text in the label number format (map §4.2), else General-like grouping for an unprepared tree.
  if (controls.showVal) parts.push(tree.values?.[i] ?? formatNumber(tree.weight[i]))
  return parts.join(controls.separator ?? ', ')
}
/** The treemap's parent-label layout: authored, else `overlapping` (the proposed default, unverified). */
export const parentLabelLayout = (drawing: HierarchyDrawing): 'none' | 'banner' | 'overlapping' => {
  const visible = drawing.group.chartEx.series.findIndex(info => !info.hidden)
  return drawing.group.chartEx.series[visible < 0 ? 0 : visible]?.layoutPr.parentLabelLayout ?? 'overlapping'
}
/** A node's colour: its root branch in the palette; an authored series fill or colour paints every node. */
export function nodeColor(drawing: HierarchyDrawing, i: number): { fill?: ChartFillStyle; color: string } {
  const { group, series, tree } = drawing
  if (series.fill) return { fill: series.fill, color: series.color }
  const palette = group.chartEx.palette
  return { color: palette?.length ? palette[tree.branch[i] % palette.length] : series.color }
}
/** Label text over a fill colour: the authored label colour, else dark on light fills and white on dark ones. */
export function textColor(series: ChartSeries, background: string): string {
  if (series.labels?.text?.color) return `#${series.labels.text.color}`
  const n = parseInt(background, 16)
  if (!Number.isFinite(n) || background.length !== 6) return '#FFFFFF'
  const channel = (v: number) => { const c = v / 255; return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4 }
  const luminance = 0.2126 * channel(n >> 16 & 255) + 0.7152 * channel(n >> 8 & 255) + 0.0722 * channel(n & 255)
  return luminance > 0.45 ? '#404040' : '#FFFFFF'
}
/** Every string a hierarchy paint may draw, measured during layout (the paint never measures). */
export function hierarchyTexts(model: ChartModel): { px: number; texts: string[] } | undefined {
  const drawing = hierarchyDrawing(model)
  if (!drawing) return undefined
  const { group, series, tree } = drawing, texts: string[] = []
  const parents = group.family === 'treemap' && parentLabelLayout(drawing) !== 'none'
  for (let i = 0; i < tree.count; i++) {
    const leaf = isLeaf(tree, i)
    const text = group.family === 'treemap' && !leaf ? (parents ? tree.labels[i] : '') : nodeText(series, tree, i)
    if (text) texts.push(text)
  }
  return { px: hierarchyLabelPx(model, series), texts }
}
/** Stage-2 construction and retention bound of a hierarchy's measured strings and legend (layoutBound adds it). */
export function hierarchyLayoutBytes(model: ChartModel): number {
  const measured = hierarchyTexts(model)
  if (!measured) return 0
  const tree = model.preparedAxes!.hierarchy!
  let bytes = 2048
  for (const text of measured.texts) bytes += 96 + text.length * 4
  for (let i = 0; i < tree.count; i++) if (tree.parent[i] < 0) bytes += 192 + tree.labels[i].length * 4
  return bytes
}
/** Paint-local buffers of a hierarchy paint (node geometry, the sort scratch), admitted before the paint. */
export const hierarchyPaintBytes = (model: ChartModel): number => {
  const tree = model.preparedAxes?.hierarchy
  return tree && isHierarchyGroup(model.groups[0]) ? 1024 + tree.count * 64 : 0
}
/** Legend entries: the root-level branches in their colours (bounded by the rows the legend shows). */
export function hierarchyLegend(model: ChartModel, limit: number): Array<{ name: string; color: string }> | undefined {
  const drawing = hierarchyDrawing(model)
  if (!drawing) return isHierarchyGroup(model.groups[0]) ? [] : undefined
  const out: Array<{ name: string; color: string }> = [], { tree } = drawing
  for (let i = 0; i < tree.count && out.length < limit; i = tree.end[i]) out.push({ name: tree.labels[i], color: nodeColor(drawing, i).color })
  return out
}
