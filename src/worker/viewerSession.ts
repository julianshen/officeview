/**
 * Resident document session (spec §§4.1, 6.3, 10.1–10.2; plan S4.1).
 *
 * One session owns one opened workbook: its SourceStore per sheet, the
 * per-sheet committed views, prepared columns and structural handles. Every
 * command after open reads this resident state; nothing reparses or clones the
 * workbook. Sort/filter build a private draft, pass the shared structural
 * guards and commit order/visibility/revisions in one synchronous step, so a
 * refusal or failure leaves every generation untouched.
 *
 * Scheduling, cancellation fan-out and publication races beyond owner checks
 * belong to S4.2; LRU/eviction of inactive sheets belongs to S4.3.
 */
import type {
  AuthoritativeCell, AxisMetrics, AxisSizeInterval, CacheIdentity, CellRead, ColumnEntry, ColumnQuery,
  CompactSourceRowSet, FilterSpec, ProfileId, GridSelection, IntervalSet, OperationRefusal, OverviewGroup, Page,
  PopulatedQuery, RangeCompletion, ReadOptions, RowOrder, SingleCellPayload, SortSpec, SourceFailure,
  SourceRowInterval, SourceStamp, SourceStore, SparseRowPage, StructuralResult, ValuePolicy, WorkbookOwner,
} from '../xlsx/view/contracts'
import type { XlsxCell, XlsxDocument, XlsxRow, XlsxSheet } from '../xlsx/types'
import type { TableMetadata } from '../xlsx/formula/types'
import {
  MAX_COLS, MAX_ROWS, cellId, isValidCoordinate, isValidSourceRect, parseA1, type SourceRect, type SourceRow,
} from '../xlsx/view/coordinates'
import { createAxisMetrics } from '../xlsx/view/axes'
import { createIdentityOrder } from '../xlsx/view/order'
import { createIntervalSet, normalizeFilter, projectVisibility, visibilityFromHidden, type CompactRowVisibility } from '../xlsx/view/visibility'
import { createSheetMetrics } from '../xlsx/view/projection'
import { hiddenSourceRowsFromDisplay, sortRowOrder, type SortOptions, type SortResult } from '../xlsx/view/sort'
import { StructuralIndex, type PopulatedCellSource } from '../xlsx/view/structure'
import { guardFilter, guardSort } from '../xlsx/view/operationGuard'
import { INITIAL_GENERATION as INITIAL_VALUE_GEN, isUint32Array, validateBitset } from './viewerProtocol'

const FULL_SHEET: SourceRect = Object.freeze({ r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 })
const READ_ROWS_PAGE = 200
const PREPARE_PAGE_LIMIT = 4096
/** Initial worker-side defaults; authoritative sheet metrics land with S2.3/S5. */
const DEFAULT_ROW_CSS = 20
const DEFAULT_COL_CSS = 64

const sameStamp = (a: SourceStamp, b: SourceStamp): boolean =>
  a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen
export const sameIdentity = (a: CacheIdentity, b: CacheIdentity): boolean =>
  a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen
  && a.fontGen === b.fontGen && a.viewGen === b.viewGen
/** Injective per-sheet cache key: JSON array encoding cannot alias ids that
 * contain a delimiter ("a:b"+"c" versus "a"+"b:c"), so documents never collide. */
export function cacheKey(identity: CacheIdentity): string {
  return JSON.stringify([identity.docId, identity.sheetId, identity.valueGen, identity.fontGen, identity.viewGen])
}
/** Aggregate (whole-workbook) identity: the global owner plus the complete ordered
 * sheet revision/value vector (§10.1). "A=1" and "B=1" never share a key. */
export function workbookIndexKey(owner: WorkbookOwner): string {
  return JSON.stringify([owner.docId, owner.globalViewGen, owner.workbookValueGen, owner.fontGen,
    owner.sheetVector.map(entry => [entry.sheetId, entry.revision, entry.valueGen])])
}
function sameWorkbookOwner(a: WorkbookOwner, b: WorkbookOwner): boolean {
  return workbookIndexKey(a) === workbookIndexKey(b)
}
const failure = (code: SourceFailure['code'], message: string, retryable = false): SourceFailure =>
  Object.freeze({ code, retryable, message })
function failedPage<T>(stamp: SourceStamp, error: SourceFailure): Page<T> {
  return { items: [], completeness: 'incomplete', stamp, error }
}
function lowerBound(sorted: readonly number[], value: number): number {
  let lo = 0, hi = sorted.length
  while (lo < hi) {
    const mid = Math.floor((lo + hi) / 2)
    if (sorted[mid] < value) lo = mid + 1
    else hi = mid
  }
  return lo
}
const isPopulated = (cell: XlsxCell): boolean => cell.value !== null || cell.formula !== undefined

/** Exact source membership as ascending inclusive intervals. Resident refs are
 * not resolvable by this interim adapter; the caller reports them unsupported. */
export function rowSetIntervals(rows: CompactSourceRowSet): SourceRowInterval[] | null {
  const out: SourceRowInterval[] = []
  const push = (r0: number, r1: number) => {
    if (!isValidCoordinate(r0, 0) || !isValidCoordinate(r1, 0) || r0 > r1) throw new RangeError('Invalid row set')
    const previous = out[out.length - 1]
    if (previous && r0 <= previous.r1) throw new RangeError('Row set is not ascending')
    if (previous && previous.r1 + 1 === r0) previous.r1 = r1
    else out.push({ r0, r1 })
  }
  switch (rows.kind) {
    case 'intervals':
      for (const interval of rows.intervals) push(interval.r0, interval.r1)
      return out
    case 'list':
      if (!isValidCoordinate(rows.r0, 0) || !isValidCoordinate(rows.r1, 0) || rows.r0 > rows.r1 || !isUint32Array(rows.rows)) {
        throw new RangeError('Invalid row list bounds')
      }
      for (const row of rows.rows) {
        if (row < rows.r0 || row > rows.r1) throw new RangeError('List row outside bounds')
        push(row, row)
      }
      return out
    case 'bitset': {
      // Full encoding check first: axis bounds, word count, zero padding, popcount.
      validateBitset(rows, 'source')
      for (let row = rows.r0; row <= rows.r1; row++) {
        const offset = row - rows.r0
        if ((rows.words[Math.floor(offset / 32)] >>> (offset % 32)) & 1) push(row, row)
      }
      return out
    }
    case 'ref':
      return null
    default:
      throw new RangeError('Unknown row set kind')
  }
}
interface ColumnSlice {
  rows: readonly number[]
  entries: readonly Extract<ColumnEntry, { kind: 'cell' | 'unavailable' }>[]
}
/** Bounded column page over a sparse sorted slice: populated entries plus explicit
 * blank runs covering every requested row, never a dense row list. */
/** Column cursors are positional over the requested membership and identical for
 * the store and the resident prepared view, so preparing a column between pages
 * keeps an outstanding cursor valid. */
const columnCursorPrefix = (col: number): string => `col${col}`
function columnPage(
  slice: ColumnSlice, intervals: readonly SourceRowInterval[], limit: number, cursor: string | undefined, prefix: string,
): { items: ColumnEntry[]; nextCursor?: string } {
  let at = 0, row = intervals[0]?.r0 ?? 0
  if (cursor !== undefined) {
    const parts = cursor.split(':')
    const i = Number(parts[1]), r = Number(parts[2])
    if (parts.length !== 3 || parts[0] !== prefix || !Number.isSafeInteger(i) || !Number.isSafeInteger(r)
      || i < 0 || i > intervals.length || !isValidCoordinate(r, 0)) throw new RangeError('Invalid column cursor')
    at = i; row = r
  }
  const items: ColumnEntry[] = []
  while (at < intervals.length && items.length < limit) {
    const { r0, r1 } = intervals[at]
    if (row < r0) row = r0
    if (row > r1) { at++; continue }
    const k = lowerBound(slice.rows, row)
    const next = k < slice.rows.length && slice.rows[k] <= r1 ? slice.rows[k] : -1
    if (next < 0) {
      items.push({ kind: 'blank', r0: row, r1 })
      at++
      continue
    }
    if (next > row) {
      items.push({ kind: 'blank', r0: row, r1: next - 1 })
      row = next
      continue
    }
    items.push(slice.entries[k])
    row++
  }
  // Normalize past exhausted intervals so a cursor never names row 1,048,576
  // (or any row outside its interval) and an exhausted query has no cursor.
  while (at < intervals.length && row > intervals[at].r1) {
    at++
    if (at < intervals.length) row = intervals[at].r0
  }
  return at < intervals.length ? { items, nextCursor: `${prefix}:${at}:${row}` } : { items }
}

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

/**
 * TODO(S1.2): interim object-backed SourceStore over the current parser output
 * (spec §4.2 `ObjectXlsxSheetAdapter` semantics, `engine` value policy). It indexes
 * only populated rows/cells and never mutates or clones the parsed sheet. S1.2
 * replaces it with the real Object/Paged adapters behind this same interface.
 */
class InterimObjectSheetStore implements SourceStore {
  readonly valuePolicy: ValuePolicy = 'engine'
  readonly extent: SourceStore['extent']
  private readonly rows: readonly XlsxRow[]
  private readonly rowIndexes: readonly number[]
  private readonly columns = new Map<number, ColumnSlice>()
  private released = false
  /** `tables` undefined means the workbook table inventory is unavailable (the
   * parser's fail-closed signal), never "no tables". */
  constructor(readonly stamp: SourceStamp, private readonly sheet: XlsxSheet, private readonly tables: readonly TableMetadata[] | undefined) {
    const rows = sheet.rows.filter(row => isValidCoordinate(row.index, 0) && row.cells.length)
      .sort((a, b) => a.index - b.index)
    this.rows = rows
    this.rowIndexes = rows.map(row => row.index)
    let r0 = MAX_ROWS, r1 = -1, c0 = MAX_COLS, c1 = -1
    for (const row of rows) {
      for (const cell of row.cells) {
        if (!isValidCoordinate(cell.row, cell.col)) continue
        r0 = Math.min(r0, cell.row); r1 = Math.max(r1, cell.row)
        c0 = Math.min(c0, cell.col); c1 = Math.max(c1, cell.col)
      }
    }
    this.extent = Object.freeze(r1 < 0 ? { rows: MAX_ROWS, cols: MAX_COLS } : { rows: MAX_ROWS, cols: MAX_COLS, used: { r0, r1, c0, c1 } })
  }
  private guard(options: ReadOptions): SourceFailure | undefined {
    if (this.released) return failure('stale-source', 'Source released')
    if (options.signal?.aborted) return failure('cancelled', 'Request cancelled')
    if (!options.stamp || !sameStamp(options.stamp, this.stamp)) return failure('stale-source', 'Source owner changed', true)
    return undefined
  }
  private cellAt(row: number, col: number): XlsxCell | undefined {
    const k = lowerBound(this.rowIndexes, row)
    if (this.rowIndexes[k] !== row) return undefined
    return this.rows[k].cells.find(cell => cell.col === col)
  }
  private column(col: number): ColumnSlice {
    let slice = this.columns.get(col)
    if (!slice) {
      const rows: number[] = [], entries: Extract<ColumnEntry, { kind: 'cell' }>[] = []
      for (const row of this.rows) {
        const cell = row.cells.find(item => item.col === col)
        if (cell && cell.row === row.index) { rows.push(row.index); entries.push({ kind: 'cell', row: row.index, cell: toPayload(cell) }) }
      }
      slice = { rows, entries }
      this.columns.set(col, slice)
    }
    return slice
  }
  async readCell(at: { row: number; col: number }, options: ReadOptions): Promise<CellRead> {
    const error = this.guard(options)
    if (error) return { kind: 'error', code: error.code, retryable: error.retryable }
    if (!isValidCoordinate(at.row, at.col)) return { kind: 'error', code: 'invalid-coordinate', retryable: false }
    const cell = this.cellAt(at.row, at.col)
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
    let visited = 0
    while (at < this.rows.length && this.rowIndexes[at] <= range.r1 && visited < READ_ROWS_PAGE) {
      const row = this.rows[at++]
      visited++
      const cells: AuthoritativeCell[] = []
      for (const cell of row.cells) {
        if (cell.row === row.index && cell.col >= range.c0 && cell.col <= range.c1) cells.push(toPayload(cell))
      }
      if (cells.length) items.push({ row: row.index, cells: cells.sort((a, b) => a.col - b.col) })
    }
    const more = at < this.rows.length && this.rowIndexes[at] <= range.r1
    return { items, completeness: 'complete', stamp: this.stamp, ...(more ? { nextCursor: `${key}${at}` } : {}) }
  }
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
    try {
      const page = columnPage(this.column(request.col), intervals, request.limit, request.cursor, columnCursorPrefix(request.col))
      return { ...page, completeness: 'complete', stamp: this.stamp }
    } catch (cause) {
      return failedPage(this.stamp, failure('invalid-cursor', (cause as Error).message))
    }
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
    for (const row of this.rows) {
      for (const cell of row.cells) {
        // A dynamic array's evaluated extent (its generated children) lives in the
        // formula engine and is not bridged yet (S3.1 spill wiring); the saved
        // footprint alone would let child rows escape the guards. Its potential
        // domain, anchor to sheet edge, stays unknown until that bridge lands.
        if (cell.dynamicArray) addMissing({ r0: cell.row, r1: MAX_ROWS - 1, c0: cell.col, c1: MAX_COLS - 1 })
        if (cell.arrayRef === undefined) continue
        const saved = parseRange(cell.arrayRef)
        const anchorRect = { r0: cell.row, r1: cell.row, c0: cell.col, c1: cell.col }
        if (!intersects(saved ?? anchorRect, rect) && !intersects(anchorRect, rect)) continue
        spills.push({ anchor: { row: cell.row, col: cell.col }, kind: cell.dynamicArray ? 'dynamic' as const : 'fixed' as const, ...(saved ? { saved } : {}) })
      }
    }
    // Drawing anchors are not yet translated from parser output (S3.1 parser
    // wiring), and an unavailable table inventory cannot prove "no table". Both
    // report the domain incomplete instead of a deceptively empty complete
    // result, so guards wait/refuse rather than authorize.
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
    this.columns.clear()
  }
}

/** Source-space dimensions and authored hidden rows; never per-row objects. */
export interface SheetGeometry {
  rows: AxisMetrics
  cols: AxisMetrics
  rowOverrides: readonly AxisSizeInterval[]
  colOverrides: readonly AxisSizeInterval[]
  baseHidden: IntervalSet
}
/** Resident workbook source boundary. S1 swaps the implementation (paged
 * adapters, blob/page handles) without changing the session. */
export interface WorkbookSource {
  readonly valuePolicy: ValuePolicy
  readonly sheets: readonly { sheetId: string; name: string }[]
  store(sheetId: string): SourceStore
  geometry(sheetId: string): SheetGeometry
  release(): Promise<void>
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
export function createObjectWorkbookSource(doc: XlsxDocument, docId: string): WorkbookSource {
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
  return {
    valuePolicy: 'engine',
    sheets: Object.freeze(sheets.map(({ sheet, sheetId }) => Object.freeze({ sheetId, name: sheet.name }))),
    store(sheetId) {
      let store = stores.get(sheetId)
      if (!store) {
        const { sheet } = find(sheetId)
        const stamp = Object.freeze({ docId, sheetId, valueGen: INITIAL_VALUE_GEN })
        store = new InterimObjectSheetStore(stamp, sheet, doc.tables?.filter(table => table.sheetId === sheetId))
        stores.set(sheetId, store)
      }
      return store
    },
    geometry(sheetId) {
      let geometry = geometries.get(sheetId)
      if (!geometry) {
        geometry = sheetGeometry(find(sheetId).sheet)
        geometries.set(sheetId, geometry)
      }
      return geometry
    },
    async release() {
      await Promise.all([...stores.values()].map(store => store.release()))
      stores.clear()
      geometries.clear()
    },
  }
}

/** S8 extension point: typed filter rule evaluation over the resident store.
 * The matcher is source-row based, so it survives later reorders. */
export type FilterEvaluation =
  | { kind: 'complete'; matches: (row: SourceRow, filter: FilterSpec) => boolean }
  | { kind: 'refused'; refusal: OperationRefusal }
  | { kind: 'incomplete'; missing: readonly SourceRect[] }
  | { kind: 'error'; error: SourceFailure }
  | { kind: 'cancelled' }
export type FilterEvaluator = (input: {
  store: SourceStore
  spec: FilterSpec
  nowMs: number
  signal: AbortSignal
}) => Promise<FilterEvaluation>
export interface PreparedBudget { maxColumns: number; maxBytes: number }
/** Prepared-column working set per admission profile (§12.4 worker pool). Mobile
 * keeps fewer resident key columns; unprepared keys still stream from the store,
 * so up-to-eight-key semantics are identical. */
export const PREPARED_BUDGETS: Readonly<Record<ProfileId, PreparedBudget>> = Object.freeze({
  D: Object.freeze({ maxColumns: 8, maxBytes: 64 * 1024 * 1024 }),
  M: Object.freeze({ maxColumns: 4, maxBytes: 32 * 1024 * 1024 }),
})
/** Spec §3 invariant 5: requested columns are prepared in batches of at most eight. */
export const MAX_PREPARE_BATCH = 8
const PREPARED_COLUMN_OVERHEAD = 256
/** Conservative resident estimate: entry/payload objects plus raw UTF-16 strings. */
function preparedEntryBytes(entry: Extract<ColumnEntry, { kind: 'cell' | 'unavailable' }>): number {
  const cell = entry.cell
  const text = (value: unknown) => typeof value === 'string' ? value.length * 2 : 0
  return 128 + text(cell.value) + text(cell.formula) + text(cell.arrayRef) + text(cell.formulaType)
}
export type PrepareOutcome =
  | { kind: 'complete' }
  | { kind: 'stale' }
  | { kind: 'cancelled' }
  | { kind: 'incomplete'; missing: readonly SourceRect[] }
  | { kind: 'error'; error: SourceFailure }
  | { kind: 'quota'; limitBytes: number; requiredBytes?: number }
  | { kind: 'resource-limit'; message: string }
export interface SessionOptions {
  /** Cooperative yield for long kernels; S4.2 replaces it with the scheduler. */
  yield?: SortOptions['yield']
  locale?: string
  filterEvaluator?: FilterEvaluator
  /** Admission profile from the open command; selects PREPARED_BUDGETS. */
  profile?: ProfileId
  /** Explicit override of the profile budget (tests, embedders). */
  preparedBudget?: PreparedBudget
}
/** Private draft; commitView publishes it atomically. */
export interface ViewDraft {
  order: RowOrder
  visibility: CompactRowVisibility
  sort: SortSpec | null
  filter: FilterSpec | null
  matches?: (row: SourceRow, filter: FilterSpec) => boolean
  selection?: GridSelection
}
export interface CommittedView extends Required<Omit<ViewDraft, 'matches'>> {
  identity: CacheIdentity
  workbookOwner: WorkbookOwner
  geometry: SheetGeometry
}
export type OperationOutcome =
  | { kind: 'committed'; view: CommittedView }
  | { kind: 'refused'; refusal: OperationRefusal }
  | { kind: 'stale' }
  | { kind: 'cancelled' }
  | { kind: 'incomplete'; missing: readonly SourceRect[] }
  | { kind: 'error'; error: SourceFailure }
  | { kind: 'quota'; requiredBytes: number; limitBytes: number }

interface SheetState {
  readonly sheetId: string
  revision: number
  valueGen: number
  order: RowOrder
  visibility: CompactRowVisibility
  sort: SortSpec | null
  filter: FilterSpec | null
  matches?: (row: SourceRow, filter: FilterSpec) => boolean
  selection: GridSelection
}
interface PreparedColumn extends ColumnSlice { stamp: SourceStamp; bytes: number }
const EMPTY_SELECTION: GridSelection = Object.freeze({ active: null, anchor: null, ranges: Object.freeze([]) })
const matchAll = () => true

export class DocumentSession {
  private globalGen = 0
  private readonly workbookValueGen = INITIAL_VALUE_GEN
  private readonly fontGen = 0
  private readonly states = new Map<string, SheetState>()
  /** Insertion order is recency order (oldest first). */
  private readonly prepared = new Map<string, PreparedColumn>()
  private preparedBytes = 0
  private readonly structures = new Map<string, StructuralIndex>()
  private published: WorkbookOwner | null = null
  private closed = false
  activeSheetId: string

  constructor(readonly docId: string, readonly source: WorkbookSource, private readonly options: SessionOptions = {}) {
    if (!source.sheets.length) throw new RangeError('Workbook has no sheets')
    this.activeSheetId = source.sheets[0].sheetId
  }
  get globalViewGen(): number { return this.globalGen }
  private get budget(): PreparedBudget {
    return this.options.preparedBudget ?? PREPARED_BUDGETS[this.options.profile ?? 'D']
  }
  get isClosed(): boolean { return this.closed }
  get publishedWorkbookIndex(): WorkbookOwner | null { return this.published }
  hasSheet(sheetId: string): boolean { return this.source.sheets.some(sheet => sheet.sheetId === sheetId) }
  /** Lazily materialized: an untouched sheet keeps no mask or order state. */
  private state(sheetId: string): SheetState {
    let state = this.states.get(sheetId)
    if (!state) {
      if (!this.hasSheet(sheetId)) throw new RangeError(`Unknown sheet ${sheetId}`)
      state = {
        sheetId, revision: 0, valueGen: INITIAL_VALUE_GEN, order: createIdentityOrder(),
        visibility: visibilityFromHidden(this.source.geometry(sheetId).baseHidden),
        sort: null, filter: null, selection: EMPTY_SELECTION,
      }
      this.states.set(sheetId, state)
    }
    return state
  }
  identity(sheetId: string): CacheIdentity {
    const state = this.states.get(sheetId)
    if (!state && !this.hasSheet(sheetId)) throw new RangeError(`Unknown sheet ${sheetId}`)
    return Object.freeze({
      docId: this.docId, sheetId, valueGen: state?.valueGen ?? INITIAL_VALUE_GEN, fontGen: this.fontGen, viewGen: state?.revision ?? 0,
    })
  }
  workbookOwner(): WorkbookOwner {
    return Object.freeze({
      docId: this.docId,
      globalViewGen: this.globalGen,
      workbookValueGen: this.workbookValueGen,
      fontGen: this.fontGen,
      sheetVector: Object.freeze(this.source.sheets.map(({ sheetId }) => {
        const state = this.states.get(sheetId)
        return Object.freeze({ sheetId, revision: state?.revision ?? 0, valueGen: state?.valueGen ?? INITIAL_VALUE_GEN })
      })),
    })
  }
  /** Read-only snapshot of a sheet's committed view. */
  view(sheetId: string): Readonly<SheetState> & { identity: CacheIdentity } {
    return { ...this.state(sheetId), identity: this.identity(sheetId) }
  }
  isCurrent(owner: CacheIdentity): boolean {
    return !this.closed && owner.docId === this.docId && this.hasSheet(owner.sheetId) && sameIdentity(owner, this.identity(owner.sheetId))
  }
  activate(sheetId: string): void {
    if (!this.hasSheet(sheetId)) throw new RangeError(`Unknown sheet ${sheetId}`)
    this.activeSheetId = sheetId
  }
  /** Atomic publication (§10.1): the sheet revision and the workbook-global
   * generation advance together with order/visibility/spec in one step. */
  async commitView(sheetId: string, draft: ViewDraft): Promise<CommittedView> {
    if (this.closed) throw new RangeError('Session closed')
    const state = this.state(sheetId)
    state.order = draft.order
    state.visibility = draft.visibility
    state.sort = draft.sort
    state.filter = draft.filter
    state.matches = draft.matches
    state.selection = draft.selection ?? state.selection
    state.revision++
    this.globalGen++
    return Object.freeze({
      identity: this.identity(sheetId), workbookOwner: this.workbookOwner(), geometry: this.source.geometry(sheetId),
      order: state.order, visibility: state.visibility, sort: state.sort, filter: state.filter, selection: state.selection,
    })
  }
  /** Captures the aggregate owner an async whole-workbook index is built under. */
  beginWorkbookIndex(): WorkbookOwner { return this.workbookOwner() }
  /** Publishes only when no sheet view, value or font changed since capture. */
  publishWorkbookIndex(captured: WorkbookOwner): boolean {
    if (this.closed || !sameWorkbookOwner(captured, this.workbookOwner())) return false
    this.published = captured
    return true
  }
  private preparedKey(stamp: SourceStamp, col: number): string {
    return JSON.stringify([stamp.docId, stamp.sheetId, stamp.valueGen, col])
  }
  isPrepared(sheetId: string, col: number): boolean {
    return this.prepared.has(this.preparedKey(this.source.store(sheetId).stamp, col))
  }
  /** Resident prepared columns, value-owned: they survive view commits and are
   * read from the store at most once per value generation while they stay in
   * the admitted working set. A batch is atomic: staged pages are charged
   * against the profile budget BEFORE they are retained, and nothing is kept
   * when admission fails. Room is made by dropping least-recently-used prepared
   * columns outside the batch (they are rebuildable caches; S4.3 owns the full
   * document-wide LRU across pages/views). */
  async prepareColumns(owner: CacheIdentity, cols: readonly number[], signal: AbortSignal): Promise<PrepareOutcome> {
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    const batch = [...new Set(cols)]
    if (batch.length > MAX_PREPARE_BATCH) {
      return { kind: 'resource-limit', message: `At most ${MAX_PREPARE_BATCH} columns are prepared per batch` }
    }
    for (const col of batch) {
      if (!isValidCoordinate(0, col)) return { kind: 'error', error: failure('invalid-coordinate', 'Invalid column') }
    }
    const budget = this.budget
    if (batch.length > budget.maxColumns) return { kind: 'quota', limitBytes: budget.maxBytes }
    const store = this.source.store(owner.sheetId)
    const stamp = store.stamp
    const keys = batch.map(col => this.preparedKey(stamp, col))
    let protectedBytes = 0
    for (const key of keys) {
      const existing = this.prepared.get(key)
      if (existing) protectedBytes += existing.bytes
    }
    const staged: { key: string; column: PreparedColumn }[] = []
    let stagedBytes = 0
    for (let i = 0; i < batch.length; i++) {
      const col = batch[i], key = keys[i]
      if (this.prepared.has(key)) continue
      const rows: number[] = [], entries: Extract<ColumnEntry, { kind: 'cell' | 'unavailable' }>[] = []
      let bytes = PREPARED_COLUMN_OVERHEAD
      let cursor: string | undefined
      const seen = new Set<string>()
      do {
        const page = await store.queryColumn({
          col, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: MAX_ROWS - 1 }] }, limit: PREPARE_PAGE_LIMIT, cursor,
        }, { signal, stamp })
        if (signal.aborted) return { kind: 'cancelled' }
        if (this.closed) return { kind: 'stale' }
        if (page.error) return page.error.code === 'cancelled' ? { kind: 'cancelled' } : { kind: 'error', error: page.error }
        if (!sameStamp(page.stamp, stamp)) return { kind: 'stale' }
        for (const item of page.items) {
          if (item.kind === 'unknown') return { kind: 'incomplete', missing: [{ r0: item.r0, r1: item.r1, c0: col, c1: col }] }
          if (item.kind === 'error') return { kind: 'error', error: failure(item.code, 'Column source failed', item.retryable) }
          if (item.kind === 'cell' || item.kind === 'unavailable') {
            bytes += preparedEntryBytes(item)
            // Admission before retention: the staged peak may never exceed the
            // budget left after the columns this batch must keep.
            if (protectedBytes + stagedBytes + bytes > budget.maxBytes) {
              return { kind: 'quota', limitBytes: budget.maxBytes, requiredBytes: protectedBytes + stagedBytes + bytes }
            }
            rows.push(item.row)
            entries.push(item)
          }
        }
        if (page.completeness !== 'complete') return { kind: 'incomplete', missing: [{ r0: 0, r1: MAX_ROWS - 1, c0: col, c1: col }] }
        cursor = page.nextCursor
        if (cursor !== undefined) {
          if (seen.has(cursor)) return { kind: 'error', error: failure('invalid-cursor', 'Column cursor made no progress') }
          seen.add(cursor)
        }
      } while (cursor !== undefined)
      staged.push({ key, column: { stamp, rows, entries, bytes } })
      stagedBytes += bytes
    }
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    const keep = new Set(keys)
    // Evict least-recently-used columns outside this batch until it fits.
    for (const [key, column] of this.prepared) {
      const fitsCount = this.prepared.size + staged.length <= budget.maxColumns
      const fitsBytes = this.preparedBytes + stagedBytes <= budget.maxBytes
      if (fitsCount && fitsBytes) break
      if (keep.has(key)) continue
      this.prepared.delete(key)
      this.preparedBytes -= column.bytes
    }
    for (const key of keys) {
      const existing = this.prepared.get(key)
      if (existing) { this.prepared.delete(key); this.prepared.set(key, existing) }
    }
    for (const { key, column } of staged) {
      this.prepared.set(key, column)
      this.preparedBytes += column.bytes
    }
    return { kind: 'complete' }
  }
  /** Retained prepared working set (document-wide, all sheets). */
  preparedStats(): { columns: number; bytes: number } {
    return { columns: this.prepared.size, bytes: this.preparedBytes }
  }
  /** The sheet store with prepared columns served from resident memory. */
  preparedStore(sheetId: string): SourceStore {
    const store = this.source.store(sheetId)
    const session = this
    return {
      get stamp() { return store.stamp },
      get valuePolicy() { return store.valuePolicy },
      get extent() { return store.extent },
      readCell: (at, options) => store.readCell(at, options),
      readRows: (range, cursor, options) => store.readRows(range, cursor, options),
      scanPopulated: (request, options) => store.scanPopulated(request, options),
      queryStructural: (rect, options) => store.queryStructural(rect, options),
      queryOverview: (request, options) => store.queryOverview(request, options),
      ensureComplete: (rect, options) => store.ensureComplete(rect, options),
      release: () => store.release(),
      async queryColumn(request, options) {
        const key = session.preparedKey(store.stamp, request.col)
        const prepared = session.prepared.get(key)
        if (!prepared) return store.queryColumn(request, options)
        session.prepared.delete(key)
        session.prepared.set(key, prepared)
        if (options.signal?.aborted) return failedPage(store.stamp, failure('cancelled', 'Request cancelled'))
        if (!options.stamp || !sameStamp(options.stamp, prepared.stamp)) return failedPage(store.stamp, failure('stale-source', 'Source owner changed', true))
        if (!Number.isSafeInteger(request.limit) || request.limit <= 0) return failedPage(store.stamp, failure('invalid-source-rect', 'Invalid limit'))
        let intervals: SourceRowInterval[] | null
        try { intervals = rowSetIntervals(request.rows) } catch (cause) {
          return failedPage(store.stamp, failure('invalid-source-rect', (cause as Error).message))
        }
        if (!intervals) return store.queryColumn(request, options)
        try {
          return { ...columnPage(prepared, intervals, request.limit, request.cursor, columnCursorPrefix(request.col)), completeness: 'complete', stamp: prepared.stamp }
        } catch (cause) {
          return failedPage(store.stamp, failure('invalid-cursor', (cause as Error).message))
        }
      },
    }
  }
  /** Structural handles are value-owned and shared by every guard (§6.3). Only a
   * complete index is cached; incomplete coverage is re-queried next time. */
  private async structure(sheetId: string, signal: AbortSignal): Promise<StructuralIndex> {
    const store = this.source.store(sheetId)
    const key = JSON.stringify([store.stamp.docId, store.stamp.sheetId, store.stamp.valueGen])
    const cached = this.structures.get(key)
    if (cached) return cached
    const index = new StructuralIndex(await store.queryStructural(FULL_SHEET, { signal, stamp: store.stamp }))
    if (index.completeness === 'complete') this.structures.set(key, index)
    return index
  }
  /** Populated-cell feed for outside-column guards over resident readRows pages. */
  private populated(sheetId: string, signal: AbortSignal): PopulatedCellSource {
    const store = this.source.store(sheetId)
    return {
      async iterate(range, cursor) {
        const page = await store.readRows(range, cursor, { signal, stamp: store.stamp })
        const items = page.items.flatMap(row => row.cells
          .filter(cell => cell.value !== null || cell.formula !== undefined)
          .map(cell => ({ row: cell.row, col: cell.col })))
        const complete = page.completeness === 'complete' && page.error === undefined
        return {
          items, completeness: complete ? 'complete' : 'incomplete', missing: complete ? [] : [range],
          ...(complete && page.nextCursor !== undefined ? { nextCursor: page.nextCursor } : {}),
        }
      },
    }
  }
  async sort(owner: CacheIdentity, spec: SortSpec | null, signal: AbortSignal): Promise<OperationOutcome> {
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    const state = this.state(owner.sheetId)
    if (spec !== null) {
      const guard = await guardSort(spec.range, await this.structure(owner.sheetId, signal), this.populated(owner.sheetId, signal))
      if (!guard.ok) return { kind: 'refused', refusal: guard.refusal }
    }
    const store = this.preparedStore(owner.sheetId)
    const result = await sortRowOrder(store, spec, {
      signal, stamp: store.stamp, locale: this.options.locale,
      hidden: hiddenSourceRowsFromDisplay(state.order, state.visibility),
      ...(this.options.yield ? { yield: this.options.yield } : {}),
    })
    if (result.kind !== 'complete') return sortFailure(result, spec)
    // Guards and kernels awaited: publish only if nothing committed meanwhile.
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    const geometry = this.source.geometry(owner.sheetId)
    const { visibility } = projectVisibility(geometry.baseHidden, state.filter, result.order, state.matches ?? matchAll)
    return { kind: 'committed', view: await this.commitView(owner.sheetId, {
      order: result.order, visibility, sort: spec, filter: state.filter, matches: state.matches,
    }) }
  }
  async filter(owner: CacheIdentity, spec: FilterSpec | null, nowMs: number, signal: AbortSignal): Promise<OperationOutcome> {
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    const state = this.state(owner.sheetId)
    let normalized: FilterSpec | null
    try { normalized = normalizeFilter(spec) } catch (cause) {
      return { kind: 'refused', refusal: { reason: 'invalid-rule', message: (cause as Error).message } }
    }
    let matches: ((row: SourceRow, filter: FilterSpec) => boolean) | undefined
    if (normalized) {
      const evaluator = this.options.filterEvaluator
      if (!evaluator) {
        return { kind: 'refused', refusal: { reason: 'unsupported', message: 'Typed filter rule evaluation is not available yet (S8)' } }
      }
      const evaluation = await evaluator({ store: this.preparedStore(owner.sheetId), spec: normalized, nowMs, signal })
      if (evaluation.kind !== 'complete') return evaluation
      matches = evaluation.matches
    }
    const geometry = this.source.geometry(owner.sheetId)
    const candidate = projectVisibility(geometry.baseHidden, normalized, state.order, matches ?? matchAll)
    const guardRange = normalized?.range ?? state.filter?.range
    if (guardRange) {
      const before = createSheetMetrics({ rows: geometry.rows, cols: geometry.cols, order: state.order, visibility: state.visibility })
      const after = createSheetMetrics({ rows: geometry.rows, cols: geometry.cols, order: state.order, visibility: candidate.visibility })
      const guard = await guardFilter({
        range: guardRange, candidateVisibility: candidate.visibility, currentVisibility: state.visibility,
        projectionBefore: before, projectionAfter: after, structure: await this.structure(owner.sheetId, signal),
      })
      if (!guard.ok) return { kind: 'refused', refusal: guard.refusal }
    }
    if (signal.aborted) return { kind: 'cancelled' }
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    return { kind: 'committed', view: await this.commitView(owner.sheetId, {
      order: state.order, visibility: candidate.visibility, sort: state.sort, filter: normalized, matches,
    }) }
  }
  /** Re-checks cancellation and the full owner after the await: a page computed
   * while a newer view committed is never published under the old owner. */
  async valueQuery(owner: CacheIdentity, query: ColumnQuery, signal: AbortSignal): Promise<Page<ColumnEntry> | { kind: 'stale' } | { kind: 'cancelled' }> {
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    const store = this.preparedStore(owner.sheetId)
    const page = await store.queryColumn(query, { signal, stamp: store.stamp })
    if (signal.aborted) return { kind: 'cancelled' }
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    return page
  }
  async release(): Promise<void> {
    if (this.closed) return
    this.closed = true
    this.states.clear()
    this.prepared.clear()
    this.preparedBytes = 0
    this.structures.clear()
    await this.source.release()
  }
}
function sortFailure(result: Exclude<SortResult, { kind: 'complete' }>, spec: SortSpec | null): OperationOutcome {
  switch (result.kind) {
    case 'cancelled': return { kind: 'cancelled' }
    case 'incomplete': return { kind: 'incomplete', missing: spec ? [{ r0: spec.range.r0, r1: spec.range.r1, c0: spec.range.c0, c1: spec.range.c1 }] : [] }
    case 'refused': return { kind: 'refused', refusal: result.refusal }
    case 'error': return result.error.code === 'cancelled' ? { kind: 'cancelled' } : { kind: 'error', error: result.error }
    case 'resource-limit': return { kind: 'quota', requiredBytes: result.requiredBytes, limitBytes: result.limitBytes }
  }
}
