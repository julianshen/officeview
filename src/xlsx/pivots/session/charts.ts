/**
 * Pivot charts follow (plan Task 30, spec §6.6, §7.3, CODEX-3): a pivot chart's series, categories and caches rebuilt
 * from its pivot's current layout, and the session's chart gate built from that. Lazy chunk `pivots`.
 *
 * The mapping, read off Excel's saved charts (Task 24 chart cases, tests/fixtures/xlsx-pivots/interaction/chart-*.xlsx,
 * Excel 16.114; docs/validation/xlsx-pivots.md "Pivot charts follow", raw tables in
 * docs/validation/xlsx-pivots-stage3-measurements.md §(c)); pinned by tests/xlsx-pivots-charts.test.ts:
 * - **Series** are the column axis's leaf lines in order, grand totals excluded: one per Year × data field with Values on
 *   columns (6 in the seed), one per Year with Values on rows (chart-values-on-rows: 3). A series' name is its line's
 *   labels from the column header cells, outer first, joined by " - " ("2022 - Sum of Sales"; "2022").
 * - **Points (categories)** are the row axis's leaf lines in order: a line with no line of its own below it. A parent
 *   item's own line (compact, subtotal at the top), subtotal lines and the grand total are not points; a collapsed
 *   item is one point (chart-collapse-item: East, with no inner label); with Values on rows each data field line is a
 *   point and the subtotal lines per data field are not (24 points). Hidden items, value filters, sorts and page
 *   selections only change which lines there are (chart-hide-item 9, chart-value-filter 6, chart-sort's order,
 *   chart-page-select's values over chart-page-base's 4 single-level points).
 * - **Category levels**: the innermost level is each point's deepest label ('' for a collapsed point); each outer level
 *   has a label where its item starts (East at 0, North at 3, …). One row field gives single-level categories.
 * - **References** move with the location: the category reference is the label column from the first body row to the
 *   location's last row (grand row included: Pivot!$A$6:$A$22), each series' values its own column over the same rows.
 *   A value cell left blank is a blank point; the cache's format is General.
 * - **Formats (`pivotFmts`)**: Excel keeps every `c:pivotFmt` and the pivot's `chartFormats` when series disappear (6 in
 *   every case, values-on-rows' 3 series included), and no series carries its own `spPr`: `c:idx` and `c:order` are the
 *   series' position, its format the automatic one of that position. So a rebuilt series takes the saved chart's series
 *   at its position (colour included).
 *
 * Refused as `pivot-chart-mapping` (no case covers it; the pivot then stays display-only, CODEX-3): a field layout other
 * than rows [a] or [a, b] with Values last on columns after one column field, or rows [a, b, Values] with one column
 * field; more than one page field; grand totals off; a form with more than one label column (outline, tabular);
 * subtotal lines with Values on columns (subtotals at the bottom); a collapsed item with Values on rows or on the
 * column axis; an unfilled level; a label or value number format other than General; a text or error value cell; a chart that is not one clustered
 * bar/column group without varied colours, or that is uncached, truncated, date-categorised or has hidden-cell
 * provenance; trendlines or error bars; a series past the saved chart's series count (its format is unmeasured); a
 * series or category change on a chart whose series carry their own formatting, per-point formatting, deleted legend
 * entries or a title (an automatic title would follow the series), or under a single-colour style column (colours
 * shade by the series count); an empty chart.
 */
import type { ChartCategoryLevel, ChartModel, ChartRange, ChartSeries } from '../../../charts/model'
import { PointWriter, pointBytes, stringBytes } from '../../../charts/points'
import { refillChartModel } from '../../../charts/fill'
import { CHART_LABEL_LIMIT, CHART_POINT_LIMIT, CHART_SERIES_LIMIT } from '../../../charts/limits'
import { PIVOT_VALUES_FIELD, type PivotDiagnostic, type PivotTableDefinition } from '../model'
import type { EngineResult } from '../engine/axes'
import type { PivotCellValue, PivotLayout, PivotLayoutCell } from '../engine/place'
import type { ChartGate } from './session'

export const PIVOT_CHART_MAPPING = 'pivot-chart-mapping'
const refuse = (feature: string, message: string): { ok: false; diagnostic: PivotDiagnostic } => ({ ok: false, diagnostic: { kind: 'unsupported', feature, message } })
const unmapped = (def: PivotTableDefinition, why: string) => refuse(PIVOT_CHART_MAPPING, `The pivot chart of pivot table ${def.name} cannot follow its changes: ${why}`)

/** The chart's pivot name: `[book]Sheet!Name` → `Name` (the client links charts by it). */
const linkedName = (chart: ChartModel) => { const name = chart.pivot?.name ?? ''; return name.slice(name.lastIndexOf('!') + 1) }
const n15 = (v: number) => Number(v.toPrecision(15))

/** The measured field layouts: Values last on columns over one column field (rows [a] or [a, b]), or Values last on
 * rows [a, b, Values] under one column field. */
type Shape = 'valuesOnColumns' | 'valuesOnRows'
function shapeOf(def: PivotTableDefinition): Shape | string {
  const rows = def.rowFields, cols = def.colFields
  const plain = (fields: readonly number[]) => fields.every(field => field !== PIVOT_VALUES_FIELD)
  const shape: Shape | undefined = cols.length === 2 && cols[1] === PIVOT_VALUES_FIELD && plain(cols.slice(0, 1)) && (rows.length === 1 || rows.length === 2) && plain(rows) ? 'valuesOnColumns'
    : cols.length === 1 && plain(cols) && rows.length === 3 && rows[2] === PIVOT_VALUES_FIELD && plain(rows.slice(0, 2)) ? 'valuesOnRows' : undefined
  if (!shape) return `its fields (rows ${JSON.stringify(rows)}, columns ${JSON.stringify(cols)}) have no measured chart mapping`
  if (def.dataFields.length < 2) return 'a single data field has no measured chart mapping'
  if (def.pageFields.length > 1) return 'more than one page field has no measured chart mapping'
  if (!def.rowGrandTotals || !def.colGrandTotals) return 'grand totals turned off have no measured chart mapping'
  if (def.location.firstDataCol !== 1) return 'only the compact form (one label column) has a measured chart mapping'
  if (!rows.every(field => field === PIVOT_VALUES_FIELD || (def.fields[field]?.compact && def.fields[field]?.outline))) return 'only the compact form has a measured chart mapping'
  if (shape === 'valuesOnRows' && !(def.compactData && def.outlineData)) return 'only the compact form has a measured chart mapping'
  return shape
}

/** A label cell's text as the chart cache holds it: text as is, a General number as Excel writes it. */
function labelText(cell: PivotLayoutCell | undefined): string | undefined {
  const value: PivotCellValue = cell?.value ?? { kind: 'empty' }
  if (value.kind === 'empty') return ''
  if (value.kind === 'text' || value.kind === 'error') return value.value
  if (cell!.numFmtId !== undefined && cell!.numFmtId !== 0) return undefined
  return String(n15(value.value))
}

interface Point { line: number; items: readonly number[]; labels: string[] }
interface Built { categories: string[]; levels: ChartCategoryLevel[]; points: Point[]; series: Array<{ col: number; name: string }>; dataRow0: number }

/** The series and points of a layout under the measured mapping, or why it has none. */
function mapLayout(layout: PivotLayout, def: PivotTableDefinition, shape: Shape): Built | string {
  const { rect } = layout
  const dataRow0 = rect.minRow + def.location.firstDataRow, itemRow0 = rect.minRow + def.location.firstHeaderRow, dataCol0 = rect.minCol + def.location.firstDataCol
  const rows = layout.axes.rows, cols = layout.axes.cols
  if (rect.maxRow - dataRow0 + 1 !== rows.length || rect.maxCol - dataCol0 + 1 !== cols.length) return 'its layout does not place one row per line'
  const cells = new Map<number, PivotLayoutCell>()
  for (const cell of layout.cells) cells.set(cell.row * 16_384 + cell.col, cell)
  const cellAt = (row: number, col: number) => cells.get(row * 16_384 + col)
  const depth = def.rowFields.length, width = def.colFields.length

  const points: Point[] = []
  const labels: string[] = []
  for (let i = 0; i < rows.length; i++) {
    const line = rows[i]!
    if (line.type === 'grand') continue
    if (line.type === 'default' && shape === 'valuesOnRows') continue
    if (line.type !== 'data') return `a ${line.type} row line has no measured chart mapping`
    const own = line.items.length
    if (line.items.includes(-1) || own < 1 || own > depth) return 'a row line with unfilled levels has no measured chart mapping'
    // Compact form: each line writes its deepest label only, under its parents' own lines.
    if (line.first !== own - 1 || labels.length < own - 1) return 'a row line that writes more than one label has no measured chart mapping'
    const text = labelText(cellAt(dataRow0 + i, rect.minCol))
    if (text === undefined) return 'a row label with a number format has no measured chart mapping'
    labels.length = own - 1
    labels.push(text)
    const next = rows[i + 1]
    if (next && next.type === 'data' && next.first >= own) continue
    // A leaf with fewer levels than the axis is a collapsed item: measured for an outer item over one inner field.
    if (own < depth && !(shape === 'valuesOnColumns' && depth === 2 && own === 1)) return 'a collapsed item here has no measured chart mapping'
    points.push({ line: i, items: line.items, labels: labels.slice() })
  }
  const series: Built['series'] = []
  const heads: string[] = []
  for (let j = 0; j < cols.length; j++) {
    const line = cols[j]!
    if (line.type === 'grand') continue
    if (line.type !== 'data') return `a ${line.type} column line has no measured chart mapping`
    if (line.items.length !== width || line.items.includes(-1) || line.first > heads.length) return 'a collapsed or unfilled column line has no measured chart mapping'
    for (let level = line.first; level < width; level++) {
      const text = labelText(cellAt(itemRow0 + level, dataCol0 + j))
      if (text === undefined) return 'a column label with a number format has no measured chart mapping'
      heads[level] = text
    }
    series.push({ col: dataCol0 + j, name: heads.slice(0, width).join(' - ') })
  }
  if (!points.length || !series.length) return 'an empty chart has no measured mapping'
  if (points.length > CHART_POINT_LIMIT || series.length > CHART_SERIES_LIMIT) return 'the chart would pass the chart point or series bound'
  if ([...series.map(s => s.name), ...points.flatMap(p => p.labels)].some(text => text.length > CHART_LABEL_LIMIT)) return `a label longer than ${CHART_LABEL_LIMIT} characters has no measured chart mapping`
  const categories = points.map(p => p.items.length === depth ? p.labels[depth - 1]! : '')
  const levels: ChartCategoryLevel[] = []
  for (let level = depth - 2; level >= 0; level--) {
    const starts: number[] = [], texts: string[] = []
    points.forEach((p, k) => {
      const prev = points[k - 1]
      if (prev && p.items.slice(0, level + 1).every((item, m) => prev.items[m] === item)) return
      starts.push(k); texts.push(p.labels[level] ?? '')
    })
    levels.push({ starts: Uint32Array.from(starts), labels: texts })
  }
  return { categories, levels, points, series, dataRow0 }
}

/** Formatting a series carries of its own (anything but the automatic format of its position). */
const ownFormat = (s: ChartSeries) => Boolean(s.fill || s.lineStyle || s.markerStyle || s.dPt?.size || s.pointFills?.length || s.pointColors?.length
  || s.labels || s.invertColor || s.explosion !== undefined)
const perPoint = (s: ChartSeries) => Boolean(s.dPt?.size || s.pointFills?.length || s.pointColors?.length || s.labels?.points)

/**
 * The chart rebuilt over `layout` (the live definition `def`), through `refillChartModel`: series from the column
 * axis's leaf lines, categories from the row axis's leaf lines, per the mapping above. Pure: the chart is not changed.
 */
export function pivotChartSeries(chart: ChartModel, layout: PivotLayout, def: PivotTableDefinition): EngineResult<ChartModel> {
  if (!chart.pivot || linkedName(chart) !== def.name) return unmapped(def, 'the chart is not linked to this pivot table')
  if (chart.unsupported || chart.refused) return unmapped(def, 'the chart is shown as a placeholder')
  const group = chart.groups[0]
  if (chart.groups.length !== 1 || !group || group.kind || group.family !== 'bar' || group.grouping !== 'clustered' || group.varyColors) return unmapped(def, 'only a clustered column or bar chart has a measured mapping')
  if (chart.categoryPending || chart.categoryValues || chart.categoryHidden || chart.truncated) return unmapped(def, 'an uncached, date, truncated or hidden-cell chart has no measured mapping')
  const saved = group.series
  if (saved.some(s => s.pending || s.values.hidden || s.trendlines || s.errorBars || s.analytic)) return unmapped(def, 'uncached, hidden-cell or analytic series have no measured mapping')
  if (saved.some((s, k) => s.index !== undefined && s.index !== k)) return unmapped(def, 'series out of their positions have no measured mapping')
  const shape = shapeOf(def)
  if (shape !== 'valuesOnColumns' && shape !== 'valuesOnRows') return unmapped(def, shape)
  const built = mapLayout(layout, def, shape)
  if (typeof built === 'string') return unmapped(def, built)
  // pivotFmts: a rebuilt series takes the saved series at its position; past them there is no measured format.
  if (built.series.length > saved.length) return unmapped(def, `a series past the chart's ${saved.length} saved series has no measured format`)
  const sameSeries = built.series.length === saved.length && built.series.every((s, k) => s.name === saved[k]!.name)
  const sameCategories = built.categories.length === chart.categories.length && built.categories.every((c, k) => c === chart.categories[k])
  if (!sameSeries && (saved.some(ownFormat) || chart.legendDeleted?.length || chart.title !== undefined)) return unmapped(def, 'series that change under their own formatting, deleted legend entries or a title have no measured mapping')
  if (!sameCategories && saved.some(perPoint)) return unmapped(def, 'points that change under per-point formatting have no measured mapping')
  // Automatic colours by position hold for the accent-cycle style column (c:style 2 in every case); a single-colour
  // style column shades by the series count (charts/style.ts styleColor), unmeasured when the count changes.
  const id = chart.styleId, column = id !== undefined && Number.isInteger(id) && id >= 1 && id <= 48 ? (id - 1) % 8 : 1
  if (!sameSeries && column !== 1) return unmapped(def, 'series that change under a single-colour chart style have no measured colours')

  // Values: each series' column over the leaf lines (a blank cell a blank point; numbers General).
  const values: PivotLayoutCell[][] = []
  const cells = new Map<number, PivotLayoutCell>()
  for (const cell of layout.cells) cells.set(cell.row * 16_384 + cell.col, cell)
  for (const s of built.series) {
    const column: PivotLayoutCell[] = []
    for (const p of built.points) {
      const cell = cells.get((built.dataRow0 + p.line) * 16_384 + s.col)
      if (cell && cell.value.kind !== 'empty') {
        if (cell.value.kind === 'text' || cell.value.kind === 'error') return unmapped(def, 'a text or error value cell has no measured chart mapping')
        if (cell.value.kind === 'number' && cell.numFmtId !== undefined && cell.numFmtId !== 0) return unmapped(def, 'a value number format other than General has no measured chart mapping')
      }
      column.push(cell ?? { row: 0, col: 0, value: { kind: 'empty' } })
    }
    values.push(column)
  }
  const sheet = def.sheet, lastRow = layout.rect.maxRow
  const categoryRef: ChartRange = { sheet, r0: built.dataRow0, c0: layout.rect.minCol, r1: lastRow, c1: layout.rect.minCol + def.location.firstDataCol - 1 }
  const next = refillChartModel(chart, charge => {
    const n = built.points.length
    charge(built.categories.reduce((bytes, text) => bytes + stringBytes(text), 0))
    for (const level of built.levels) charge(64 + level.starts.length * 36 + level.labels.reduce((bytes, text) => bytes + text.length * 2, 0))
    const series = built.series.map((s, k): ChartSeries => {
      charge(256 + stringBytes(s.name) + pointBytes(n))
      const out = new PointWriter(n, 'lazy', charge)
      values[k]!.forEach((cell, i) => {
        if (cell.value.kind === 'number') out.value(i, cell.value.value, 'General')
      })
      const { pending: _pending, categories: _categories, absent: _absent, refs: _refs, ...rest } = saved[k]!
      void _pending, void _categories, void _absent, void _refs
      return { ...rest, index: k, name: s.name, values: out.done(), refs: { values: { sheet, r0: built.dataRow0, c0: s.col, r1: lastRow, c1: s.col } } }
    })
    const { preparedAxes: _prepared, categoryLevels: _levels, ...rest } = chart
    void _prepared, void _levels
    return { ...rest, categories: built.categories, ...(built.levels.length ? { categoryLevels: built.levels } : {}), categoryRef, groups: [{ ...group, series }] }
  })
  if (next.refused) return refuse('pivot-chart-memory', `The pivot chart of pivot table ${def.name} needs more memory than is available for charts`)
  return { ok: true, value: next }
}

/** Whether a rebuilt chart shows the saved chart's own data: its series' names and points, categories and levels. */
function showsSaved(saved: ChartModel, rebuilt: ChartModel): boolean {
  const a = saved.groups[0]!.series, b = rebuilt.groups[0]!.series
  const sameText = (x: readonly string[], y: readonly string[]) => x.length === y.length && x.every((t, k) => t === y[k])
  if (a.length !== b.length || !sameText(saved.categories, rebuilt.categories)) return false
  const levels = (m: ChartModel) => (m.categoryLevels ?? []).map(level => `${[...level.starts].join(',')}|${level.labels.join('\u0000')}`)
  if (!sameText(levels(saved), levels(rebuilt))) return false
  return a.every((s, k) => {
    const t = b[k]!
    if (s.name !== t.name || s.values.length !== t.values.length) return false
    for (let i = 0; i < s.values.length; i++) {
      if (s.values.states[i] !== t.values.states[i]) return false
      if (s.values.states[i] === 1 && n15(s.values.values[i]!) !== n15(t.values.values[i]!)) return false
    }
    return true
  })
}

/**
 * Task 25's chart gate: every linked chart rebuilt over the layout, or the first refusal (the pivot then refuses at
 * open, or the change at apply; CODEX-3). The first time the gate sees a chart is the session's open, on the saved
 * layout: the chart must then show exactly its own saved data, or it is not the default pivot mapping (series
 * reordered, a chart linked to a pivot it does not show) and refuses. Every later layout is rebuilt from the saved chart.
 */
const verified = new WeakSet<ChartModel>()
export const pivotChartGate: ChartGate = (charts, layout, def) => {
  const out: ChartModel[] = []
  for (const chart of charts) {
    const rebuilt = pivotChartSeries(chart, layout, def)
    if (!rebuilt.ok) return rebuilt
    if (!verified.has(chart)) {
      if (!showsSaved(chart, rebuilt.value)) return unmapped(def, 'its saved series and categories are not the pivot\'s default chart mapping')
      verified.add(chart)
    }
    out.push(rebuilt.value)
  }
  return { ok: true, value: out }
}
