/**
 * The treemap and sunburst hierarchy (docs/charts/chartex-map.md §7; lazy chunk).
 *
 * Acquisition: a hierarchy's category reference (`strDim type="cat"`, e.g. TM/SB `Sheet1!$B$4:$D$14`) is a path with
 * one level per column (`f@dir="col"`, the default) or per row (`dir="row"`), the root level first [P]. Every level
 * is read on every path (the resident and paged XLSX fills, the DOCX/PPTX embedded workbook): `levelLines` gives the
 * readers one line per level. A level over the analytic bound (65,536 points) or more than
 * CHART_HIERARCHY_LEVEL_LIMIT levels refuses the chart before anything is read.
 *
 * Derivation (stage 1, `buildHierarchy`), every rule a project rule [P] until an Office reference confirms it:
 * - a current path P is carried across rows; a non-blank cell that differs from P at its level replaces it and clears
 *   every deeper level, an equal one keeps the inherited deeper levels, a blank cell keeps P at its level;
 * - a row's depth is its deepest non-blank level; levels below it are ignored (a missing level makes a leaf); a row
 *   with no label at any level is excluded; a level cleared by an ancestor change is an unnamed node "";
 * - nodes are identified by their full path; duplicate paths are summed; a parent is the sum of its descendants; a
 *   path that is both a leaf and a prefix keeps its own weight as an unnamed child "";
 * - weights: a positive finite value is used; zero, negative, blank, error or non-numeric excludes the row (its cells
 *   still move P); an unavailable weight refuses the chart;
 * - more than CHART_HIERARCHY_NODE_LIMIT nodes (every node: internal and unnamed ones too) refuses the chart.
 *
 * The tree is an immutable, transferable record (typed arrays and strings) kept in the prepared stage-1 record
 * (`PreparedAxes.hierarchy`); its working set and retained arrays are charged before they are built.
 */
import type { ChartDiagnostic, ChartExGroup, ChartGroup, ChartModel, ChartPoints, ChartRange } from '../model'
import type { ChartLine } from '../fill'
import { CHART_DATA_NOT_AVAILABLE, CHART_TOO_MANY_POINTS } from '../fill'
import { CHART_ANALYTIC_POINT_LIMIT } from '../parse'
import { POINT_UNAVAILABLE, POINT_VALUE, stringBytes } from '../points'

/** Spec 6.5: treemap/sunburst nodes per chart. */
export const CHART_HIERARCHY_NODE_LIMIT = 4096
/** Levels of one hierarchy reference [P]: a wider category range refuses the chart before it is read. */
export const CHART_HIERARCHY_LEVEL_LIMIT = 64

export interface ChartExHierarchy {
  count: number
  /** Node labels in pre-order (parents before children, children in source order); '' is an unnamed node. */
  labels: string[]
  /** Pre-order index of each node's parent; -1 for a root-level node. */
  parent: Int32Array
  /** 0 for a root-level node. */
  depth: Uint16Array
  /** Subtree weight (a leaf: its summed weights). */
  weight: Float64Array
  /** Exclusive end of each node's subtree: its children are i + 1, end[i + 1], ... before end[i]. */
  end: Int32Array
  /** The ordinal of each node's root-level branch (its colour). */
  branch: Int32Array
  /** Levels drawn (the deepest depth + 1); 0 for an empty tree. */
  levels: number
  /** The sum of every root-level weight. */
  total: number
  /** Each node's value text in the label number format (stage 1; present when the labels show values). */
  values?: string[]
}
export type HierarchyResult = { tree: ChartExHierarchy; diagnostics: ChartDiagnostic[] } | { refused: string }

export const isHierarchyGroup = (group: ChartGroup | undefined): group is ChartExGroup =>
  group?.kind === 'chartEx' && (group.family === 'treemap' || group.family === 'sunburst')

/** Points per level of a hierarchy reference (rows for `col`, columns for `row`). */
export const levelPoints = (range: ChartRange, dir: 'col' | 'row'): number => dir === 'row' ? range.c1 - range.c0 + 1 : range.r1 - range.r0 + 1
/** Levels of a hierarchy reference (columns for `col`, rows for `row`). */
export const levelCount = (range: ChartRange, dir: 'col' | 'row'): number => dir === 'row' ? range.r1 - range.r0 + 1 : range.c1 - range.c0 + 1
/** A hierarchy reference past the acquisition bounds: refused before any cell is read. */
export const levelsOverLimit = (range: ChartRange, dir: 'col' | 'row'): boolean =>
  levelPoints(range, dir) > CHART_ANALYTIC_POINT_LIMIT || levelCount(range, dir) > CHART_HIERARCHY_LEVEL_LIMIT
/** The lines of a hierarchy reference, root level first: every point of every level (never a prefix). */
export function levelLines(range: ChartRange, ownerSheet: string, dir: 'col' | 'row'): ChartLine[] {
  const sheet = range.sheet ?? ownerSheet, count = levelPoints(range, dir)
  return Array.from({ length: levelCount(range, dir) }, (_, k): ChartLine => dir === 'row'
    ? { sheet, vertical: false, fixed: range.r0 + k, start: range.c0, count, truncated: false, text: true }
    : { sheet, vertical: true, fixed: range.c0 + k, start: range.r0, count, truncated: false, text: true })
}

/** Working-set price of one node while the tree is built (record, child map entry, path key). */
const NODE_WORK_BYTES = 320
/** Retained price of one node's arrays (parent, depth, weight, end, branch, label slot). */
const NODE_BYTES = 4 + 2 + 8 + 4 + 4 + 8
/** The retained footprint of a derived tree (its labels priced in full, though they share the level strings). */
export const hierarchyBytes = (tree: ChartExHierarchy | undefined): number =>
  tree ? 256 + tree.count * NODE_BYTES + tree.labels.reduce((n, label) => n + stringBytes(label), 0) + (tree.values ?? []).reduce((n, text) => n + 8 + stringBytes(text), 0) : 0

interface Node { label: string; parent: number; own: number; ownSeq: number; seq: number; children?: Map<string, number> }
const EMPTY = (): ChartExHierarchy => ({ count: 0, labels: [], parent: new Int32Array(0), depth: new Uint16Array(0), weight: new Float64Array(0), end: new Int32Array(0), branch: new Int32Array(0), levels: 0, total: 0 })

/**
 * Derives the tree of map §7 from category levels (root first; `undefined` or '' is a blank cell) and the weights
 * (the `size` points), row by row. `charge` admits each allocation before it is made (it throws to refuse).
 */
export function buildHierarchy(levels: ReadonlyArray<ReadonlyArray<string | undefined>>, weights: ChartPoints, charge: (bytes: number) => void = () => {}): HierarchyResult {
  const depthLimit = levels.length
  let rows = weights.length
  for (const level of levels) rows = Math.max(rows, level.length)
  // The working set is bounded by the node limit (plus the one refused node) and the current path.
  charge(Math.min(CHART_HIERARCHY_NODE_LIMIT + 1, rows * (depthLimit + 1)) * NODE_WORK_BYTES + depthLimit * 16 + 512)
  const path: Array<string | undefined> = new Array(depthLimit).fill(undefined)
  const roots = new Map<string, number>(), nodes: Node[] = []
  let seq = 0, excludedWeights = 0, excludedRows = 0, unnamed = 0, overflow = false
  const childOf = (parent: number, label: string): number => {
    const map = parent < 0 ? roots : (nodes[parent].children ??= new Map())
    let i = map.get(label)
    if (i === undefined) {
      if (nodes.length >= CHART_HIERARCHY_NODE_LIMIT) { overflow = true; return -1 }
      i = nodes.length
      nodes.push({ label, parent, own: 0, ownSeq: -1, seq: seq++ })
      map.set(label, i)
      if (!label) unnamed++
    }
    return i
  }
  for (let r = 0; r < rows; r++) {
    const state = r < weights.length ? weights.states[r] : 0
    // Map §7 order: a row with no label at any level is excluded (step 1) before its weight is looked at (step 6).
    let depth = -1
    for (let j = 0; j < depthLimit; j++) if (levels[j][r]) depth = j
    if (depth < 0) { excludedRows++; continue }
    if (state === POINT_UNAVAILABLE) return { refused: CHART_DATA_NOT_AVAILABLE }
    for (let j = 0; j <= depth; j++) {
      const cell = levels[j][r]
      if (cell && cell !== path[j]) { path[j] = cell; for (let k = j + 1; k < depthLimit; k++) path[k] = undefined }
    }
    const w = state === POINT_VALUE ? weights.values[r] : NaN
    if (!(w > 0 && Number.isFinite(w))) { excludedWeights++; continue }
    let at = -1
    for (let j = 0; j <= depth; j++) {
      at = childOf(at, path[j] ?? '')
      if (overflow) return { refused: CHART_TOO_MANY_POINTS }
    }
    const node = nodes[at]
    if (node.ownSeq < 0) node.ownSeq = seq++
    node.own += w
  }
  // A leaf that is also a prefix keeps its own weight as an unnamed child, placed where its own weight first appeared.
  for (let i = 0; i < nodes.length; i++) {
    const node = nodes[i]
    if (!(node.own > 0) || !node.children?.size) continue
    const fresh = !node.children.has('')
    const child = childOf(i, '')
    if (overflow) return { refused: CHART_TOO_MANY_POINTS }
    if (fresh) nodes[child].seq = node.ownSeq
    if (nodes[child].ownSeq < 0) nodes[child].ownSeq = node.ownSeq
    nodes[child].own += node.own
    node.own = 0
  }
  const count = nodes.length
  const diagnostics: ChartDiagnostic[] = []
  if (excludedWeights) diagnostics.push({ code: 'chart-hierarchy-weight-excluded', reason: `${excludedWeights} ${excludedWeights === 1 ? 'row has a weight' : 'rows have weights'} that ${excludedWeights === 1 ? 'is' : 'are'} not a positive number; excluded` })
  if (excludedRows) diagnostics.push({ code: 'chart-hierarchy-row-excluded', reason: `${excludedRows} ${excludedRows === 1 ? 'row has' : 'rows have'} no category at any level; excluded` })
  if (unnamed) diagnostics.push({ code: 'chart-hierarchy-unnamed', reason: `${unnamed} unnamed ${unnamed === 1 ? 'node' : 'nodes'} (a level with no category)` })
  for (const d of diagnostics) charge(64 + stringBytes(d.code) + stringBytes(d.reason))
  if (!count) return { tree: EMPTY(), diagnostics }
  // Subtree weights: a child is always created after its parent.
  charge(count * 8)
  const sums = new Float64Array(count)
  for (let i = count - 1; i >= 0; i--) { sums[i] += nodes[i].own; if (nodes[i].parent >= 0) sums[nodes[i].parent] += sums[i] }
  // Pre-order, children by first appearance.
  charge(count * NODE_BYTES + 256 + nodes.reduce((n, node) => n + stringBytes(node.label), 0))
  const tree: ChartExHierarchy = { count, labels: new Array<string>(count), parent: new Int32Array(count), depth: new Uint16Array(count), weight: new Float64Array(count), end: new Int32Array(count), branch: new Int32Array(count), levels: 0, total: 0 }
  const ordered = (map: Map<string, number> | undefined) => map ? [...map.values()].sort((a, b) => nodes[a].seq - nodes[b].seq) : []
  let next = 0, branch = -1
  const visit = (i: number, parent: number, depth: number) => {
    const at = next++
    if (depth === 0) branch++
    tree.labels[at] = nodes[i].label; tree.parent[at] = parent; tree.depth[at] = depth; tree.weight[at] = sums[i]; tree.branch[at] = branch
    tree.levels = Math.max(tree.levels, depth + 1)
    for (const c of ordered(nodes[i].children)) visit(c, at, depth + 1)
    tree.end[at] = next
  }
  for (const i of ordered(roots)) { visit(i, -1, 0); tree.total += sums[i] }
  return { tree, diagnostics }
}

/**
 * Stage 1 of a treemap/sunburst model: the tree of its first visible series (the size points; its filled levels, else
 * its one cached category level, else the ordinals 1..n [P]). Undefined for any other chart.
 */
export function deriveChartHierarchy(model: ChartModel, charge: (bytes: number) => void): HierarchyResult | undefined {
  const group = model.groups[0]
  if (!isHierarchyGroup(group)) return undefined
  const visible = group.chartEx.series.findIndex(info => !info.hidden)
  const series = group.series[visible < 0 ? 0 : visible]
  if (!series) return { tree: EMPTY(), diagnostics: [] }
  let levels: ReadonlyArray<ReadonlyArray<string | undefined>> | undefined = series.levels ?? (series.categories ? [series.categories] : undefined)
  if (!levels) {
    charge(series.values.length * (8 + stringBytes('65536')))
    levels = [Array.from({ length: series.values.length }, (_, i) => String(i + 1))]
  }
  const result = buildHierarchy(levels, series.values, charge)
  const others = group.chartEx.series.filter((info, i) => !info.hidden && i !== visible).length
  if (others && 'tree' in result) {
    const diagnostic = { code: 'chart-hierarchy-series-ignored', reason: `${others} further ${others === 1 ? 'series is' : 'series are'} not drawn: a ${group.family} draws one series` }
    charge(64 + stringBytes(diagnostic.code) + stringBytes(diagnostic.reason))
    result.diagnostics.push(diagnostic)
  }
  return result
}

/**
 * The value text of every node in the series' label number format (`dataLabels/numFmt`, unless source-linked; else
 * General), formatted at stage 1 so the paint never formats. Only when the drawn series' labels show values.
 */
export function formatHierarchyValues(model: ChartModel, tree: ChartExHierarchy, format: (value: number, code: string | undefined) => string, charge: (bytes: number) => void): void {
  const group = model.groups[0]
  if (!isHierarchyGroup(group) || !tree.count) return
  const visible = group.chartEx.series.findIndex(info => !info.hidden)
  const labels = group.series[visible < 0 ? 0 : visible]?.labels
  if (!labels?.showVal) return
  // Each text is bounded by the formatter's 256-unit output, admitted before it is built.
  charge(tree.count * (8 + stringBytes(' '.repeat(256))))
  const code = labels.sourceLinked !== true ? labels.formatCode : undefined
  tree.values = Array.from({ length: tree.count }, (_, i) => format(tree.weight[i], code))
}
