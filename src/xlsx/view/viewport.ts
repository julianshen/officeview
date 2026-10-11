/**
 * Full-coordinate renderer seam (S2.3, spec §§5.3, 5.4, 11.1, 11.5).
 *
 * The interactive path never reads the legacy capped `rowY`/`colX` prefix
 * arrays: geometry is the S2.1 full-capacity axis/projection model, a viewport
 * plan locates its display window first and only then resolves source rows and
 * intersecting merges. Every source coordinate is painted at its own logical
 * offset (or its sorted display slot) or not at all; no missing geometry falls
 * back to the origin. Typed `AxisMetricsError` failures are mapped to diagnosed
 * fallbacks here so they never reach a caller as an unhandled exception.
 */
import type { XlsxCell, XlsxMergeRange, XlsxSheet } from '../types'
import type { AxisSizeInterval, LogicalViewport, RowOrder, RowVisibility, WorkerDiagnostic } from './contracts'
import { MAX_COLS, MAX_ROWS, isValidCoordinate, type SourceCell } from './coordinates'
import { AxisMetricsError, createAxisMetrics, type AxisMetricsErrorCode, type SparseAxisMetrics } from './axes'
import { createSheetMetrics, type ProjectedSheetMetrics } from './projection'
import { RectangleIndex } from './structure'

/** Excel default geometry in CSS pixels; identical to the static renderer so an
 * ordinary small sheet lays out the same in both paths. */
export const SHEET_DEFAULT_ROW_CSS = 20
export const SHEET_DEFAULT_COL_CSS = 64
const CHAR_CSS = 7
const COL_PADDING_CSS = 5

/** Stable diagnostic code per typed axis failure (spec §5.3, plan S2.3). */
export const AXIS_ERROR_DIAGNOSTICS: Readonly<Record<AxisMetricsErrorCode, string>> = Object.freeze({
  CONTRADICTORY_OVERRIDES: 'axis-contradictory-overrides',
  PRECISION_EXCEEDED: 'axis-precision-exceeded',
  TOTAL_NOT_FINITE: 'axis-total-not-finite',
  INVALID_METADATA: 'axis-invalid-metadata',
  INVALID_QUERY: 'axis-invalid-query',
})

/** Proposed §11.1 per-frame limits: individual cell-layout records and gridline
 * instances on desktop (D) and mobile (M). Beyond them a frame is coarse before
 * any cell is enumerated; the coarse layer itself belongs to S5.1. */
export const VIEWPORT_DETAIL_LIMITS = Object.freeze({
  D: Object.freeze({ records: 20_000, gridlines: 4096 }),
  M: Object.freeze({ records: 10_000, gridlines: 2048 }),
})
export type ViewportProfile = keyof typeof VIEWPORT_DETAIL_LIMITS

export interface SheetGeometryOptions { order?: RowOrder; visibility?: RowVisibility }
/** Built once per geometry generation (sheet + order + visibility), never per paint. */
export interface SheetGeometry {
  readonly kind: 'ready'
  readonly metrics: ProjectedSheetMetrics
  /** Fallbacks taken while building (dropped sizes, conflicting/invalid metadata). */
  readonly diagnostics: readonly WorkerDiagnostic[]
  /** Populated cells of one source row, ascending by column; undefined for blank rows. */
  cellsOf(sourceRow: number): readonly XlsxCell[] | undefined
  readonly merges: RectangleIndex<XlsxMergeRange>
}
export type SheetGeometryResult =
  | SheetGeometry
  | { readonly kind: 'diagnostic'; readonly diagnostics: readonly WorkerDiagnostic[] }

/** Logical CSS box of one painted record; merge anchors span their whole range. */
export interface CellPlacement { cell: XlsxCell; source: SourceCell; x: number; y: number; width: number; height: number; merge?: XlsxMergeRange }
export interface MergePlacement { range: XlsxMergeRange; x: number; y: number; width: number; height: number }
export interface ViewportDetail {
  kind: 'detail'
  /** Row-major (display slot, column) paint order. */
  cells: readonly CellPlacement[]
  merges: readonly MergePlacement[]
  /** Logical CSS positions of the visible grid boundaries. */
  rowLines: readonly number[]
  colLines: readonly number[]
  /** Full logical sheet extent (the worksheet continues past the used range). */
  extent: { width: number; height: number }
  diagnostics: readonly WorkerDiagnostic[]
}
export type ViewportPlan =
  | ViewportDetail
  | { kind: 'pending' | 'diagnostic' | 'coarse'; diagnostics: readonly WorkerDiagnostic[] }

export function axisDiagnostic(error: AxisMetricsError, axis: 'rows' | 'cols' | 'projection' | 'viewport'): WorkerDiagnostic {
  return { code: AXIS_ERROR_DIAGNOSTICS[error.code], message: `${axis}: ${error.message}` }
}

function validSize(value: number | undefined): value is number {
  return value !== undefined && Number.isFinite(value) && value >= 0
}

/** Try the authored axis, then without authored sizes, then plain defaults. Each
 * rejected attempt leaves its typed diagnostic; hidden intervals survive the
 * first fallback because they never contradict (all zero). */
function buildAxis(axis: 'rows' | 'cols', logicalCount: number, defaultSizeCss: number, overrides: AxisSizeInterval[], collapsed: Array<{ start: number; end: number }>, diagnostics: WorkerDiagnostic[]): SparseAxisMetrics | null {
  const attempts = [{ overrides, collapsed }, { overrides: [], collapsed }, { overrides: [], collapsed: [] }]
  for (const attempt of attempts) {
    try {
      return createAxisMetrics({ logicalCount, defaultSizeCss, ...attempt })
    } catch (error) {
      if (!(error instanceof AxisMetricsError)) throw error
      diagnostics.push(axisDiagnostic(error, axis))
    }
  }
  return null
}

/** Build the full-capacity geometry for a small object sheet. Authored-hidden rows
 * and columns collapse unconditionally (spec §5.3), with or without drawings. */
export function buildSheetGeometry(sheet: XlsxSheet, options: SheetGeometryOptions = {}): SheetGeometryResult {
  const diagnostics: WorkerDiagnostic[] = []
  const cellsByRow = new Map<number, XlsxCell[]>()
  const rowHidden = new Map<number, boolean>()
  const rowOverrides: AxisSizeInterval[] = []
  let skippedRows = 0, skippedCells = 0
  for (const row of sheet.rows) {
    if (!isValidCoordinate(row.index, 0)) { skippedRows++; continue }
    // Duplicate records of one row index are invalid OOXML; collapse only when
    // every record agrees, so a stray hidden duplicate never hides visible data.
    rowHidden.set(row.index, (rowHidden.get(row.index) ?? true) && row.hidden === true)
    if (row.customHeight && row.heightPt !== undefined) {
      if (validSize(row.heightPt)) rowOverrides.push({ start: row.index, end: row.index, sizeCss: Math.round(row.heightPt * (96 / 72)) })
      else diagnostics.push({ code: 'invalid-dimension', message: `row ${row.index + 1}: ignored height ${row.heightPt}` })
    }
    let cells = cellsByRow.get(row.index)
    for (const item of row.cells) {
      if (!isValidCoordinate(row.index, item.col)) { skippedCells++; continue }
      if (!cells) cellsByRow.set(row.index, cells = [])
      cells.push(item)
    }
  }
  // Stable: equal columns keep authored order, exactly as the static loop paints them.
  for (const cells of cellsByRow.values()) cells.sort((a, b) => a.col - b.col)
  const rowCollapsed: Array<{ start: number; end: number }> = []
  for (const [index, hidden] of rowHidden) {
    if (hidden) rowCollapsed.push({ start: index, end: index })
  }
  const conflicting = sheet.rows.filter(row => isValidCoordinate(row.index, 0) && row.hidden === true && rowHidden.get(row.index) === false)
  if (conflicting.length) diagnostics.push({ code: 'conflicting-row-metadata', message: `${new Set(conflicting.map(row => row.index)).size} row index(es) carry both hidden and visible records; kept visible` })
  if (skippedRows || skippedCells) diagnostics.push({ code: 'coordinate-outside-domain', message: `skipped ${skippedRows} row record(s) and ${skippedCells} cell(s) outside the worksheet domain` })

  const colOverrides: AxisSizeInterval[] = []
  const colCollapsed: Array<{ start: number; end: number }> = []
  for (const spec of sheet.cols) {
    if (!isValidCoordinate(0, spec.min) || !isValidCoordinate(0, spec.max) || spec.min > spec.max) {
      diagnostics.push({ code: 'coordinate-outside-domain', message: `column spec ${spec.min + 1}:${spec.max + 1} ignored` })
      continue
    }
    if (spec.hidden) colCollapsed.push({ start: spec.min, end: spec.max })
    else if (spec.widthChars !== undefined) {
      if (validSize(spec.widthChars)) colOverrides.push({ start: spec.min, end: spec.max, sizeCss: Math.round(spec.widthChars * CHAR_CSS + COL_PADDING_CSS) })
      else diagnostics.push({ code: 'invalid-dimension', message: `column spec ${spec.min + 1}:${spec.max + 1}: ignored width ${spec.widthChars}` })
    }
  }

  const rows = buildAxis('rows', MAX_ROWS, SHEET_DEFAULT_ROW_CSS, rowOverrides, rowCollapsed, diagnostics)
  const cols = buildAxis('cols', MAX_COLS, SHEET_DEFAULT_COL_CSS, colOverrides, colCollapsed, diagnostics)
  if (!rows || !cols) return { kind: 'diagnostic', diagnostics }
  let metrics: ProjectedSheetMetrics
  try {
    metrics = createSheetMetrics({ rows, cols, order: options.order, visibility: options.visibility })
  } catch (error) {
    // A projection that cannot be built must not silently show another order.
    if (!(error instanceof AxisMetricsError)) throw error
    diagnostics.push(axisDiagnostic(error, 'projection'))
    return { kind: 'diagnostic', diagnostics }
  }
  const mergeEntries: Array<{ rect: { r0: number; r1: number; c0: number; c1: number }; value: XlsxMergeRange }> = []
  for (const range of sheet.mergeRanges ?? []) {
    if (!isValidCoordinate(range.minRow, range.minCol) || !isValidCoordinate(range.maxRow, range.maxCol)
      || range.minRow > range.maxRow || range.minCol > range.maxCol) {
      diagnostics.push({ code: 'coordinate-outside-domain', message: 'merge range outside the worksheet domain ignored' })
      continue
    }
    mergeEntries.push({ rect: { r0: range.minRow, r1: range.maxRow, c0: range.minCol, c1: range.maxCol }, value: range })
  }
  return Object.freeze({
    kind: 'ready' as const,
    metrics,
    diagnostics: Object.freeze(diagnostics),
    cellsOf: (sourceRow: number) => cellsByRow.get(sourceRow),
    merges: RectangleIndex.build(mergeEntries),
  })
}

function firstCellAtOrAfter(cells: readonly XlsxCell[], col: number): number {
  let low = 0, high = cells.length
  while (low < high) {
    const mid = Math.floor((low + high) / 2)
    if (cells[mid].col < col) low = mid + 1
    else high = mid
  }
  return low
}

/** Plan one frame: locate the display window, enforce per-frame budgets from
 * O(log n) window counts, then resolve only visible source rows and the merges
 * intersecting them. Never scans source rows outside the window. */
export function planViewport(geometry: SheetGeometryResult | undefined, viewport: LogicalViewport, profile: ViewportProfile = 'D'): ViewportPlan {
  if (!geometry) return { kind: 'pending', diagnostics: [{ code: 'geometry-pending', message: 'sheet geometry has not been published' }] }
  if (geometry.kind !== 'ready') return { kind: 'diagnostic', diagnostics: geometry.diagnostics }
  try {
    return detailPlan(geometry, viewport, VIEWPORT_DETAIL_LIMITS[profile])
  } catch (error) {
    if (!(error instanceof AxisMetricsError)) throw error
    return { kind: 'diagnostic', diagnostics: [...geometry.diagnostics, axisDiagnostic(error, 'viewport')] }
  }
}

function detailPlan(geometry: SheetGeometry, viewport: LogicalViewport, limits: { records: number; gridlines: number }): ViewportPlan {
  const { metrics } = geometry
  const { x, y, width, height } = viewport.box
  if (!Number.isFinite(width) || !Number.isFinite(height) || width < 0 || height < 0) {
    throw new AxisMetricsError('INVALID_QUERY', 'Viewport size must be finite and nonnegative')
  }
  const rowWindow = metrics.rows.visibleBetween(y, y + height)
  const colWindow = metrics.cols.visibleBetween(x, x + width)
  const coarse = (message: string): ViewportPlan => ({ kind: 'coarse', diagnostics: [...geometry.diagnostics, { code: 'detail-limit', message }] })
  // Each visible row/column contributes one boundary instance; +2 closes the window.
  if (rowWindow.count + colWindow.count + 2 > limits.gridlines) {
    return coarse(`${rowWindow.count} rows x ${colWindow.count} columns exceed ${limits.gridlines} gridline instances`)
  }
  const extent = { width: metrics.cols.totalCss(), height: metrics.rows.totalCss() }
  const diagnostics = [...geometry.diagnostics]
  const rowLines: number[] = [], colLines: number[] = []
  const slots: number[] = []
  for (const run of rowWindow.runs) {
    for (let slot = run.start; slot <= run.end; slot++) { slots.push(slot); rowLines.push(metrics.rows.start(slot)) }
    rowLines.push(metrics.rows.start(run.end + 1))
  }
  let c0 = MAX_COLS, c1 = -1
  for (const run of colWindow.runs) {
    for (let col = run.start; col <= run.end; col++) colLines.push(metrics.cols.start(col))
    colLines.push(metrics.cols.start(run.end + 1))
    c0 = Math.min(c0, run.start)
    c1 = Math.max(c1, run.end)
  }
  const cells: Array<CellPlacement & { slot: number }> = []
  const merges: MergePlacement[] = []
  if (!slots.length || c1 < 0) return { kind: 'detail', cells, merges, rowLines, colLines, extent, diagnostics }

  // Merges never span a permutation band (the S3/S7 sort guard refuses that), so
  // outside bands display slots equal source rows and the slot envelope is the
  // source query. A merge that does touch a band is painted cell by cell.
  const page = geometry.merges.query({ r0: slots[0], r1: slots[slots.length - 1], c0, c1 }, { limit: limits.records })
  if (page.nextCursor !== undefined) return coarse(`more than ${limits.records} merged ranges intersect the frame`)
  const covering: XlsxMergeRange[] = []
  const bands = metrics.order.bands
  for (const range of page.items) {
    if (bands.some(band => band.r0 <= range.maxRow && band.r1 >= range.minRow)) {
      diagnostics.push({ code: 'merge-in-reordered-band', message: 'merged range inside a reordered band is painted per cell' })
      continue
    }
    const left = metrics.cols.start(range.minCol), top = metrics.rows.start(range.minRow)
    const placement = { range, x: left, y: top, width: metrics.cols.start(range.maxCol + 1) - left, height: metrics.rows.start(range.maxRow + 1) - top }
    covering.push(range)
    merges.push(placement)
    if (placement.width <= 0 || placement.height <= 0) continue
    const anchor = geometry.cellsOf(range.minRow)
    const at = anchor ? firstCellAtOrAfter(anchor, range.minCol) : -1
    // An anchor whose own row/column lies outside the window still owns the merged paint.
    if (anchor && at < anchor.length && anchor[at].col === range.minCol) {
      cells.push({ cell: anchor[at], source: { row: range.minRow, col: range.minCol }, slot: range.minRow, ...placement, merge: range })
    }
  }
  const coveredBy = (row: number, col: number) => covering.find(range => row >= range.minRow && row <= range.maxRow && col >= range.minCol && col <= range.maxCol)
  for (const slot of slots) {
    const source = metrics.order.displayToSource(slot)
    const rowCells = geometry.cellsOf(source)
    if (!rowCells) continue
    const top = metrics.rows.start(slot), rowHeight = metrics.rows.size(slot)
    for (let index = firstCellAtOrAfter(rowCells, c0); index < rowCells.length && rowCells[index].col <= c1; index++) {
      const item = rowCells[index]
      const colWidth = metrics.cols.size(item.col)
      // Collapsed rows/columns produce no fill, text or hit target (spec §5.3).
      if (colWidth <= 0 || coveredBy(source, item.col)) continue
      if (cells.length >= limits.records) return coarse(`more than ${limits.records} cell records in the frame`)
      cells.push({ cell: item, source: { row: source, col: item.col }, slot, x: metrics.cols.start(item.col), y: top, width: colWidth, height: rowHeight })
    }
  }
  cells.sort((a, b) => a.slot - b.slot || a.source.col - b.source.col)
  return { kind: 'detail', cells, merges, rowLines, colLines, extent, diagnostics }
}
