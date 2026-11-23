/**
 * Stage 4a (S4 §6.4): the `sheet-layout` answer for one sheet, built from its servable snapshot only. It carries
 * paint-facing geometry and nothing else (parent §12's disclosure boundary): sizes, collapsed runs, frozen panes,
 * merges, the used extent, declared filter ranges with their header and totals counts, and counts. No cell value,
 * style, format, rule, drawing, defined name, `defaultRowPt` or `nonUniformFontRows` is read into it.
 */
import { cfExplicitBounds } from '../../../src/xlsx/cfLoader'
import { isValidCoordinate, MAX_COLS, MAX_ROWS, type SourceRect } from '../../../src/xlsx/view/coordinates'
import type { DeclaredSheetFilters } from '../../../src/xlsx/view/declaredFilters'
import type { AxisMetrics, AxisSizeInterval } from '../../../src/xlsx/view/contracts'
import type { XlsxSheet } from '../../../src/xlsx/types'
import { LAYOUT_LIMITS, LayoutTooLarge, type StreamSheetLayout } from '../../../src/stream/sheet-wire'
import type { SheetSnapshot } from './snapshot'

/** The snapshot's geometry is a diagnostic (its axes could not be built): there is no layout to serve. */
export class LayoutUnavailable extends Error { constructor() { super('sheet geometry unavailable'); this.name = 'LayoutUnavailable' } }

/**
 * An axis as the tiles paint it, split into its size intervals and its collapsed (zero-extent) runs. The axis's
 * effective overrides are already disjoint segments, so a hidden line's authored size is never among the intervals:
 * it is a zero segment, a collapsed run. Throws `LayoutTooLarge` before copying an axis that cannot fit.
 */
function splitAxis(axis: AxisMetrics): { sizes: AxisSizeInterval[]; collapsed: Array<{ start: number; end: number }> } {
  const ov = axis.overrides
  if (ov.length > 2 * LAYOUT_LIMITS.intervals) throw new LayoutTooLarge(`axis with ${ov.length} segments`)
  const sizes: AxisSizeInterval[] = [], collapsed: Array<{ start: number; end: number }> = []
  for (let i = 0; i < ov.length; i++) {
    const s = ov[i]!
    if (s.sizeCss === 0) collapsed.push({ start: s.start, end: s.end })
    else sizes.push({ start: s.start, end: s.end, sizeCss: s.sizeCss })
  }
  return { sizes, collapsed }
}

/** [lo, hi] lies inside one run of `runs` (ascending, disjoint). */
function within(runs: ReadonlyArray<{ start: number; end: number }>, lo: number, hi: number): boolean {
  let a = 0, b = runs.length
  while (a < b) { const mid = (a + b) >> 1; if (runs[mid]!.end < lo) a = mid + 1; else b = mid }
  const run = runs[a]
  return run !== undefined && run.start <= lo && hi <= run.end
}

/** The used extent, as the resident source store computes it (`ObjectXlsxSheetAdapter.reindex`): every stored cell,
 * every resolved sparkline host, a sparkline host's merge, and explicit CF bounds on a sheet with no cells. */
function usedOf(sheet: XlsxSheet): SourceRect | undefined {
  let r0 = MAX_ROWS, r1 = -1, c0 = MAX_COLS, c1 = -1
  const add = (ra: number, rb: number, ca: number, cb: number) => { r0 = Math.min(r0, ra); r1 = Math.max(r1, rb); c0 = Math.min(c0, ca); c1 = Math.max(c1, cb) }
  for (const row of sheet.rows) {
    if (!isValidCoordinate(row.index, 0)) continue
    for (const cell of row.cells) if (isValidCoordinate(cell.row, cell.col)) add(cell.row, cell.row, cell.col, cell.col)
  }
  for (const [row, hosts] of sheet.resolvedSparklines ?? []) for (const col of hosts.keys()) if (isValidCoordinate(row, col)) add(row, row, col, col)
  for (const m of sheet.mergeRanges ?? []) {
    if (!isValidCoordinate(m.minRow, m.minCol) || !isValidCoordinate(m.maxRow, m.maxCol) || m.minRow > m.maxRow || m.minCol > m.maxCol) continue
    if (sheet.resolvedSparklines?.get(m.minRow)?.has(m.minCol)) add(m.minRow, m.maxRow, m.minCol, m.maxCol)
  }
  const cf = cfExplicitBounds(sheet)
  if (cf && r1 < 0) add(cf.r0, cf.r1, cf.c0, cf.c1)
  return r1 < 0 ? undefined : { r0, r1, c0, c1 }
}

const validRect = (r: SourceRect) => isValidCoordinate(r.r0, r.c0) && isValidCoordinate(r.r1, r.c1) && r.r0 <= r.r1 && r.c0 <= r.c1

/** The declared filter ranges: the worksheet autoFilter (header 1, totals 0), then each table of the sheet. Ranges,
 * header and totals counts only; never a column, rule or table style. */
function filtersOf(sheet: XlsxSheet, declared: DeclaredSheetFilters | undefined): StreamSheetLayout['filters'] {
  const out: StreamSheetLayout['filters'] = []
  const af = declared?.worksheetAutoFilter?.range
  if (af) {
    const range = { r0: af.minRow, r1: af.maxRow, c0: af.minCol, c1: af.maxCol }
    if (validRect(range)) out.push({ kind: 'autoFilter', range, header: 1, totals: 0 })
  }
  for (const t of sheet.tables ?? []) {
    const e = t.extent
    const range = { r0: e.firstRow, r1: e.firstRow + e.rows - 1, c0: e.firstCol, c1: e.firstCol + e.cols - 1 }
    if (validRect(range)) out.push({ kind: 'table', range, header: t.headerRowCount === 1 ? 1 : 0, totals: t.totalsRowCount === 1 ? 1 : 0 })
  }
  return out
}

/**
 * The 4a layout of `s` (identity view): `viewGen` 0, `sort` null, `filterSummary` empty. `declared`: the sheet's
 * declared filter state from the same saved-values parse (its worksheet autoFilter range is not on the sheet).
 * `authorHidden` iff a collapsed row lies inside a declared filter range; `counts.totalRows` is `used.r1 + 1` (0 with
 * no used extent) and `counts.visibleRows` that minus the collapsed rows below it. Throws `LayoutUnavailable` for a
 * diagnostic geometry and `LayoutTooLarge` for an axis too fragmented to fit any answer.
 */
export function buildStreamLayout(s: SheetSnapshot, sheetId: string, declared?: DeclaredSheetFilters): StreamSheetLayout {
  const g = s.geometry
  if (g.kind !== 'ready') throw new LayoutUnavailable()
  const rows = splitAxis(g.metrics.rows), cols = splitAxis(g.metrics.cols)
  const sheet = s.sheet
  // Builder's call (S4 §6.4: gap positions, "no other hidden-row metadata"): a merge wholly inside one collapsed row
  // or column run paints nothing, so it does not travel.
  const merges: SourceRect[] = []
  for (const m of sheet.mergeRanges ?? []) {
    const r = { r0: m.minRow, r1: m.maxRow, c0: m.minCol, c1: m.maxCol }
    if (validRect(r) && !within(rows.collapsed, r.r0, r.r1) && !within(cols.collapsed, r.c0, r.c1)) merges.push(r)
  }
  const used = usedOf(sheet)
  const filters = filtersOf(sheet, declared)
  const authorHidden = rows.collapsed.some((c) => filters.some((f) => c.start <= f.range.r1 && f.range.r0 <= c.end))
  const totalRows = used ? used.r1 + 1 : 0
  let collapsedBelow = 0
  for (const c of rows.collapsed) if (c.start < totalRows) collapsedBelow += Math.min(c.end, totalRows - 1) - c.start + 1
  const f = sheet.frozen
  return {
    sheetId, viewGen: 0,
    defaults: { rowPx: g.metrics.rows.defaultSizeCss, colPx: g.metrics.cols.defaultSizeCss },
    rows: rows.sizes, cols: cols.sizes, collapsedRows: rows.collapsed, collapsedCols: cols.collapsed,
    ...(f ? { frozen: { rows: f.rows, cols: f.cols, ...(f.topRow !== undefined ? { topRow: f.topRow } : {}), ...(f.leftCol !== undefined ? { leftCol: f.leftCol } : {}) } } : {}),
    merges, ...(used ? { used } : {}), filters,
    sort: null, filterSummary: [], authorHidden, counts: { visibleRows: totalRows - collapsedBelow, totalRows },
  }
}
