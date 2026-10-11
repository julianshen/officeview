/**
 * TODO(S1.2): interim object-backed workbook source for the resident viewer
 * session (spec §4.2 `ObjectXlsxSheetAdapter` semantics, `engine` value policy).
 *
 * It wraps the current parser output BY REFERENCE (no clone, no mutation),
 * indexes only populated rows, and keeps no per-column caches: every page is
 * produced from the row index and bounded by rows AND cells. Coverage it cannot
 * resolve (drawing anchors, dynamic-array evaluated extents, an unavailable
 * table inventory) is reported incomplete with an explicit diagnostic. S1.2
 * replaces this module with the real Object/Paged adapters behind the same
 * WorkbookSource/SourceStore interfaces.
 */
import type {
  AuthoritativeCell, AxisSizeInterval, CellRead, ColumnEntry, ColumnQuery, OverviewGroup, Page, PopulatedQuery,
  RangeCompletion, ReadOptions, SingleCellPayload, SourceFailure, SourceRowInterval, SourceStamp, SourceStore,
  SparseRowPage, StructuralResult, ValuePolicy, WorkerDiagnostic,
} from '../xlsx/view/contracts'
import type { XlsxCell, XlsxDocument, XlsxRow, XlsxSheet } from '../xlsx/types'
import type { TableMetadata } from '../xlsx/formula/types'
import { MAX_COLS, MAX_ROWS, cellId, isValidCoordinate, isValidSourceRect, parseA1, type SourceRect } from '../xlsx/view/coordinates'
import { createAxisMetrics } from '../xlsx/view/axes'
import { createIntervalSet } from '../xlsx/view/visibility'
import { INITIAL_GENERATION } from './viewerProtocol'
import {
  failedPage, failure, lowerBound, rowSetIntervals, sameStamp, type SheetGeometry, type WorkbookSource,
} from './viewerSession'

/** Row-page bounds: at most 200 rows and 262,144 cells (matches the fixture store). */
export const READ_ROWS_PAGE_ROWS = 200
export const READ_ROWS_MAX_CELLS = 262_144
export interface ObjectSourceOptions { readRowsMaxCells?: number }
/** Initial worker-side defaults; authoritative sheet metrics land with S2.3/S5. */
const DEFAULT_ROW_CSS = 20
const DEFAULT_COL_CSS = 64

const isPopulated = (cell: XlsxCell): boolean => cell.value !== null || cell.formula !== undefined
function toPayload(cell: XlsxCell): SingleCellPayload {
  const payload: SingleCellPayload = {
    id: cellId(cell.row, cell.col), row: cell.row, col: cell.col, value: cell.value, styleIndex: cell.styleIndex,
    valueProvenance: 'engine',
  }
  if (cell.hasCachedValue !== undefined) payload.hasCachedValue = cell.hasCachedValue
  if (cell.valueIsError !== undefined) payload.valueIsError = cell.valueIsError
  if (cell.formula !== undefined) payload.formula = cell.formula
  if (cell.sharedFormula !== undefined) payload.sharedFormula = { ...cell.sharedFormula }
  if (cell.arrayRef !== undefined) payload.arrayRef = cell.arrayRef
  if (cell.dynamicArray !== undefined) payload.dynamicArray = cell.dynamicArray
  if (cell.formulaType !== undefined) payload.formulaType = cell.formulaType
  if (cell.ca !== undefined) payload.ca = cell.ca
  return Object.freeze(payload)
}
function parseRange(ref: string): SourceRect | null {
  try {
    const [first, last = first] = ref.split(':')
    const a = parseA1(first), b = parseA1(last)
    return { r0: Math.min(a.row, b.row), r1: Math.max(a.row, b.row), c0: Math.min(a.col, b.col), c1: Math.max(a.col, b.col) }
  } catch {
    return null
  }
}
const intersects = (a: SourceRect, b: SourceRect): boolean =>
  a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1

class InterimObjectSheetStore implements SourceStore {
  readonly valuePolicy: ValuePolicy = 'engine'
  readonly extent: SourceStore['extent']
  private readonly rows: readonly XlsxRow[]
  private readonly rowIndexes: readonly number[]
  private readonly dynamicAnchors: readonly { row: number; col: number }[]
  private released = false
  /** `tables` undefined means the workbook table inventory is unavailable (the
   * parser's fail-closed signal), never "no tables". */
  constructor(
    readonly stamp: SourceStamp,
    private readonly sheet: XlsxSheet,
    private readonly tables: readonly TableMetadata[] | undefined,
    private readonly maxCells: number,
  ) {
    const rows = sheet.rows.filter(row => isValidCoordinate(row.index, 0) && row.cells.length)
      .sort((a, b) => a.index - b.index)
    this.rows = rows
    this.rowIndexes = rows.map(row => row.index)
    const dynamic: { row: number; col: number }[] = []
    let r0 = MAX_ROWS, r1 = -1, c0 = MAX_COLS, c1 = -1
    for (const row of rows) {
      for (const cell of row.cells) {
        if (!isValidCoordinate(cell.row, cell.col)) continue
        r0 = Math.min(r0, cell.row); r1 = Math.max(r1, cell.row)
        c0 = Math.min(c0, cell.col); c1 = Math.max(c1, cell.col)
        if (cell.dynamicArray) dynamic.push({ row: cell.row, col: cell.col })
      }
    }
    this.dynamicAnchors = dynamic
    this.extent = Object.freeze(r1 < 0 ? { rows: MAX_ROWS, cols: MAX_COLS } : { rows: MAX_ROWS, cols: MAX_COLS, used: { r0, r1, c0, c1 } })
  }
  /** Per-column caches retained by this store; it builds pages from the row index. */
  get cachedColumnCount(): number { return 0 }
  /** Coverage gaps this adapter can never close by waiting. */
  coverageDiagnostics(): WorkerDiagnostic[] {
    const out: WorkerDiagnostic[] = []
    if ((this.sheet.drawings?.length ?? 0) > 0) {
      out.push({ code: 'coverage-unresolvable', message: 'Drawing anchors are not available to the interim source adapter until S3.1 parser wiring' })
    }
    if (this.dynamicAnchors.length) {
      out.push({ code: 'coverage-unresolvable', message: 'Dynamic-array spill extents are not available to the interim source adapter until the S3.1 spill bridge', source: this.dynamicAnchors[0] })
    }
    if (this.tables === undefined) {
      out.push({ code: 'coverage-unresolvable', message: 'The workbook table inventory could not be read; table bands are unknown' })
    }
    return out
  }
  private guard(options: ReadOptions): SourceFailure | undefined {
    if (this.released) return failure('stale-source', 'Source released')
    if (options.signal?.aborted) return failure('cancelled', 'Request cancelled')
    if (!options.stamp || !sameStamp(options.stamp, this.stamp)) return failure('stale-source', 'Source owner changed', true)
    return undefined
  }
  private cellIn(row: XlsxRow, col: number): XlsxCell | undefined {
    return row.cells.find(cell => cell.col === col && cell.row === row.index)
  }
  async readCell(at: { row: number; col: number }, options: ReadOptions): Promise<CellRead> {
    const error = this.guard(options)
    if (error) return { kind: 'error', code: error.code, retryable: error.retryable }
    if (!isValidCoordinate(at.row, at.col)) return { kind: 'error', code: 'invalid-coordinate', retryable: false }
    const k = lowerBound(this.rowIndexes, at.row)
    const cell = this.rowIndexes[k] === at.row ? this.cellIn(this.rows[k], at.col) : undefined
    return cell ? { kind: 'cell', cell: toPayload(cell) } : { kind: 'blank' }
  }
  async readRows(range: SourceRect, cursor: string | undefined, options: ReadOptions): Promise<Page<SparseRowPage>> {
    const error = this.guard(options)
    if (error) return failedPage(this.stamp, error)
    if (!isValidSourceRect(range)) return failedPage(this.stamp, failure('invalid-source-rect', 'Invalid range'))
    const key = `objrows:${range.r0}:${range.r1}:${range.c0}:${range.c1}:`
    let at = lowerBound(this.rowIndexes, range.r0)
    if (cursor !== undefined) {
      const offset = cursor.startsWith(key) ? Number(cursor.slice(key.length)) : NaN
      if (!Number.isSafeInteger(offset) || offset < at || offset > this.rows.length) {
        return failedPage(this.stamp, failure('invalid-cursor', 'Invalid row cursor'))
      }
      at = offset
    }
    const items: SparseRowPage[] = []
    let visited = 0, cellCount = 0
    while (at < this.rows.length && this.rowIndexes[at] <= range.r1 && visited < READ_ROWS_PAGE_ROWS) {
      const row = this.rows[at]
      const cells: AuthoritativeCell[] = []
      for (const cell of row.cells) {
        if (cell.row === row.index && cell.col >= range.c0 && cell.col <= range.c1) cells.push(toPayload(cell))
      }
      // Cell bound: stop before a row would push the page past the cap (a row
      // has at most 16,384 cells, so a first row always fits a sane cap).
      if (items.length && cellCount + cells.length > this.maxCells) break
      at++
      visited++
      cellCount += cells.length
      if (cells.length) items.push({ row: row.index, cells: cells.sort((a, b) => a.col - b.col) })
    }
    const more = at < this.rows.length && this.rowIndexes[at] <= range.r1
    return { items, completeness: 'complete', stamp: this.stamp, ...(more ? { nextCursor: `${key}${at}` } : {}) }
  }
  /** Column page straight from the row index: populated entries plus explicit
   * blank runs covering every requested row; nothing is cached per column. */
  async queryColumn(request: ColumnQuery, options: ReadOptions): Promise<Page<ColumnEntry>> {
    const error = this.guard(options)
    if (error) return failedPage(this.stamp, error)
    if (!isValidCoordinate(0, request.col)) return failedPage(this.stamp, failure('invalid-coordinate', 'Invalid column'))
    if (!Number.isSafeInteger(request.limit) || request.limit <= 0) return failedPage(this.stamp, failure('invalid-source-rect', 'Invalid limit'))
    let intervals: SourceRowInterval[] | null
    try { intervals = rowSetIntervals(request.rows) } catch (cause) {
      return failedPage(this.stamp, failure('invalid-source-rect', (cause as Error).message))
    }
    if (!intervals) return failedPage(this.stamp, failure('unsupported', 'Resident row-set refs are not resolvable by the interim adapter'))
    const prefix = `objcol${request.col}`
    let at = 0, row = intervals[0]?.r0 ?? 0
    if (request.cursor !== undefined) {
      const parts = request.cursor.split(':')
      const i = Number(parts[1]), r = Number(parts[2])
      if (parts.length !== 3 || parts[0] !== prefix || !Number.isSafeInteger(i) || i < 0 || i >= intervals.length
        || !isValidCoordinate(r, 0) || r < intervals[i].r0 || r > intervals[i].r1) {
        return failedPage(this.stamp, failure('invalid-cursor', 'Invalid column cursor'))
      }
      at = i; row = r
    }
    const items: ColumnEntry[] = []
    while (at < intervals.length && items.length < request.limit) {
      const { r1 } = intervals[at]
      if (row < intervals[at].r0) row = intervals[at].r0
      // Next row at or after `row` that has a cell in this column, within r1.
      let k = lowerBound(this.rowIndexes, row), found: XlsxCell | undefined
      for (; k < this.rows.length && this.rowIndexes[k] <= r1; k++) {
        found = this.cellIn(this.rows[k], request.col)
        if (found) break
      }
      if (!found) {
        items.push({ kind: 'blank', r0: row, r1 })
        at++
        if (at < intervals.length) row = intervals[at].r0
        continue
      }
      if (found.row > row) {
        items.push({ kind: 'blank', r0: row, r1: found.row - 1 })
        if (items.length >= request.limit) { row = found.row; break }
      }
      items.push({ kind: 'cell', row: found.row, cell: toPayload(found) })
      row = found.row + 1
      if (row > r1) { at++; if (at < intervals.length) row = intervals[at].r0 }
    }
    return { items, completeness: 'complete', stamp: this.stamp, ...(at < intervals.length ? { nextCursor: `${prefix}:${at}:${row}` } : {}) }
  }
  async *scanPopulated(request: PopulatedQuery, options: ReadOptions): AsyncIterable<Page<AuthoritativeCell>> {
    const error = this.guard(options)
    if (error) { yield failedPage(this.stamp, error); return }
    if (!isValidSourceRect(request.range) || !Number.isSafeInteger(request.limit) || request.limit <= 0) {
      yield failedPage(this.stamp, failure('invalid-source-rect', 'Invalid populated query')); return
    }
    let items: AuthoritativeCell[] = []
    for (let at = lowerBound(this.rowIndexes, request.range.r0); at < this.rows.length && this.rowIndexes[at] <= request.range.r1; at++) {
      for (const cell of this.rows[at].cells) {
        if (cell.col < request.range.c0 || cell.col > request.range.c1 || !isPopulated(cell)) continue
        items.push(toPayload(cell))
        if (items.length === request.limit) {
          yield { items, completeness: 'complete', stamp: this.stamp }
          items = []
          const again = this.guard(options)
          if (again) { yield failedPage(this.stamp, again); return }
        }
      }
    }
    yield { items, completeness: 'complete', stamp: this.stamp }
  }
  async queryStructural(rect: SourceRect, options: ReadOptions): Promise<StructuralResult> {
    const base = { stamp: this.stamp, merges: [], tables: [], spills: [], drawings: [] }
    const error = this.guard(options) ?? (isValidSourceRect(rect) ? undefined : failure('invalid-source-rect', 'Invalid range'))
    if (error) return { ...base, completeness: 'incomplete', missing: [rect], error }
    const merges = this.sheet.mergeRanges
      .map(m => ({ id: `merge:${m.minRow}:${m.minCol}:${m.maxRow}:${m.maxCol}`, range: { r0: m.minRow, r1: m.maxRow, c0: m.minCol, c1: m.maxCol }, anchor: { row: m.minRow, col: m.minCol } }))
      .filter(m => isValidSourceRect(m.range) && intersects(m.range, rect))
    const tables = (this.tables ?? [])
      .map(t => ({ id: t.id, range: { r0: t.extent.firstRow, r1: t.extent.firstRow + t.extent.rows - 1, c0: t.extent.firstCol, c1: t.extent.firstCol + t.extent.cols - 1 }, headerRows: t.headerRowCount, totalsRows: t.totalsRowCount }))
      .filter(t => isValidSourceRect(t.range) && intersects(t.range, rect))
    const spills = []
    const missing: SourceRect[] = []
    const addMissing = (domain: SourceRect) => {
      const r0 = Math.max(domain.r0, rect.r0), r1 = Math.min(domain.r1, rect.r1)
      const c0 = Math.max(domain.c0, rect.c0), c1 = Math.min(domain.c1, rect.c1)
      if (r0 <= r1 && c0 <= c1) missing.push({ r0, r1, c0, c1 })
    }
    // A dynamic array's evaluated extent (its generated children) lives in the
    // formula engine and is not bridged yet (S3.1 spill wiring); the saved
    // footprint alone would let child rows escape the guards. Its potential
    // domain, anchor to sheet edge, stays unknown until that bridge lands.
    for (const anchor of this.dynamicAnchors) addMissing({ r0: anchor.row, r1: MAX_ROWS - 1, c0: anchor.col, c1: MAX_COLS - 1 })
    for (const row of this.rows) {
      for (const cell of row.cells) {
        if (cell.arrayRef === undefined) continue
        const saved = parseRange(cell.arrayRef)
        const anchorRect = { r0: cell.row, r1: cell.row, c0: cell.col, c1: cell.col }
        if (!intersects(saved ?? anchorRect, rect) && !intersects(anchorRect, rect)) continue
        spills.push({ anchor: { row: cell.row, col: cell.col }, kind: cell.dynamicArray ? 'dynamic' as const : 'fixed' as const, ...(saved ? { saved } : {}) })
      }
    }
    // Drawing anchors are not yet translated from parser output (S3.1 parser
    // wiring), and an unavailable table inventory cannot prove "no table".
    if ((this.sheet.drawings?.length ?? 0) > 0 || this.tables === undefined) addMissing(rect)
    return { ...base, merges, tables, spills, completeness: missing.length ? 'incomplete' : 'complete', missing }
  }
  async queryOverview(_request: unknown, options: ReadOptions): Promise<Page<OverviewGroup>> {
    return failedPage(this.stamp, this.guard(options) ?? failure('unsupported', 'Overview aggregates are owned by S1.2'))
  }
  async ensureComplete(rect: SourceRect, options: ReadOptions): Promise<RangeCompletion> {
    const error = this.guard(options)
    if (error) return { range: rect, completeness: 'incomplete', stamp: this.stamp, missing: [rect], error }
    return { range: rect, completeness: 'complete', stamp: this.stamp, missing: [] }
  }
  async release(): Promise<void> {
    this.released = true
  }
}

function sheetGeometry(sheet: XlsxSheet): SheetGeometry {
  const rowOverrides: AxisSizeInterval[] = []
  const hidden: SourceRowInterval[] = []
  for (const row of [...sheet.rows].sort((a, b) => a.index - b.index)) {
    if (!isValidCoordinate(row.index, 0)) continue
    if (row.heightPt !== undefined && Number.isFinite(row.heightPt) && row.heightPt >= 0) {
      rowOverrides.push({ start: row.index, end: row.index, sizeCss: row.heightPt * 96 / 72 })
    }
    if (row.hidden) hidden.push({ r0: row.index, r1: row.index })
  }
  const colOverrides: AxisSizeInterval[] = []
  for (const spec of sheet.cols) {
    if (spec.widthChars === undefined || !Number.isFinite(spec.widthChars) || spec.widthChars < 0) continue
    const start = Math.max(0, spec.min), end = Math.min(MAX_COLS - 1, spec.max)
    if (Number.isSafeInteger(start) && Number.isSafeInteger(end) && start <= end) {
      colOverrides.push({ start, end, sizeCss: Math.round(spec.widthChars * 7 + 5) })
    }
  }
  colOverrides.sort((a, b) => a.start - b.start)
  // Overlapping authored <col> ranges: the later start wins, keeping intervals disjoint.
  for (let i = 1; i < colOverrides.length; i++) {
    if (colOverrides[i].start <= colOverrides[i - 1].end) colOverrides[i - 1] = { ...colOverrides[i - 1], end: colOverrides[i].start - 1 }
  }
  const cols = colOverrides.filter(interval => interval.start <= interval.end)
  return Object.freeze({
    rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: DEFAULT_ROW_CSS, overrides: rowOverrides }),
    cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: DEFAULT_COL_CSS, overrides: cols }),
    rowOverrides: Object.freeze(rowOverrides),
    colOverrides: Object.freeze(cols),
    baseHidden: createIntervalSet(hidden),
  })
}
/** Wraps an already-parsed workbook by reference (no clone). Sheet ids use the
 * parser's stable workbook identity; documents are separated by docId. */
export function createObjectWorkbookSource(doc: XlsxDocument, docId: string, options: ObjectSourceOptions = {}): WorkbookSource {
  const maxCells = options.readRowsMaxCells ?? READ_ROWS_MAX_CELLS
  if (!Number.isSafeInteger(maxCells) || maxCells <= 0) throw new RangeError('Invalid readRows cell bound')
  const sheets = doc.sheets.map((sheet, index) => ({ sheet, sheetId: sheet.sheetId ?? String(index + 1) }))
  const ids = new Set<string>()
  for (const { sheetId } of sheets) {
    if (ids.has(sheetId)) throw new RangeError(`Duplicate sheet id ${sheetId}`)
    ids.add(sheetId)
  }
  const stores = new Map<string, InterimObjectSheetStore>()
  const geometries = new Map<string, SheetGeometry>()
  const find = (sheetId: string) => {
    const found = sheets.find(entry => entry.sheetId === sheetId)
    if (!found) throw new RangeError(`Unknown sheet ${sheetId}`)
    return found
  }
  const storeFor = (sheetId: string): InterimObjectSheetStore => {
    let store = stores.get(sheetId)
    if (!store) {
      const { sheet } = find(sheetId)
      const stamp = Object.freeze({ docId, sheetId, valueGen: INITIAL_GENERATION })
      store = new InterimObjectSheetStore(stamp, sheet, doc.tables?.filter(table => table.sheetId === sheetId), maxCells)
      stores.set(sheetId, store)
    }
    return store
  }
  return {
    valuePolicy: 'engine',
    sheets: Object.freeze(sheets.map(({ sheet, sheetId }) => Object.freeze({ sheetId, name: sheet.name }))),
    store: storeFor,
    geometry(sheetId) {
      let geometry = geometries.get(sheetId)
      if (!geometry) {
        geometry = sheetGeometry(find(sheetId).sheet)
        geometries.set(sheetId, geometry)
      }
      return geometry
    },
    coverageDiagnostics: sheetId => storeFor(sheetId).coverageDiagnostics(),
    async release() {
      await Promise.all([...stores.values()].map(store => store.release()))
      stores.clear()
      geometries.clear()
    },
  }
}
