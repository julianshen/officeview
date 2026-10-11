/**
 * Source adapters behind the async SourceStore (spec §§4.1-4.2, plan S1.2).
 *
 * `ObjectXlsxSheetAdapter` wraps the current parser output BY REFERENCE (no
 * clone, no mutation) for admitted small files: it indexes only populated
 * rows, keeps no per-column caches, and bounds every page by rows AND cells.
 * Coverage it cannot resolve (drawing anchors, dynamic-array evaluated extents,
 * an unavailable table inventory) is reported incomplete with a diagnostic.
 * The paged adapter for large files lives in ./pagedSource and answers the same
 * questions from stored pages; adapter parity is asserted within one value policy.
 */
import type {
  AuthoritativeCell, AxisSizeInterval, CellRead, ColumnEntry, ColumnQuery, OverviewGroup, OverviewQuery, Page, PopulatedQuery,
  RangeCompletion, ReadOptions, SingleCellPayload, SourceFailure, SourceRowInterval, SourceStamp, SourceStore,
  SparseRowPage, StructuralResult, ValuePolicy, WorkerDiagnostic,
} from './contracts'
import type { ColumnLanes } from './contracts'
import { countLoadEvaluatedFormulas, documentValuePolicy, isValueUnavailable, laneValue, SAVED_VALUES_DIAGNOSTIC, valueAuthority, valueStatus } from './values'
import type { XlsxCell, XlsxDocument, XlsxRow, XlsxSheet } from '../types'
import type { TableMetadata } from '../formula/types'
import { MAX_COLS, MAX_ROWS, cellId, isValidCoordinate, isValidSourceRect, parseA1, type SourceRect } from './coordinates'
import { createAxisMetrics } from './axes'
import { createIntervalSet } from './visibility'
import { INITIAL_GENERATION } from '../../worker/viewerProtocol'
import {
  failedPage, failure, lowerBound, rowSetIntervals, sameStamp, type SheetGeometry, type WorkbookSource,
} from '../../worker/viewerSession'
import { addRecord, answerOverview, emptyAggregate, type OverviewCellSource } from '../storage/overviewIndex'

/** Row-page bounds: at most 200 rows and 262,144 cells (matches the fixture store). */
export const READ_ROWS_PAGE_ROWS = 200
export const READ_ROWS_MAX_CELLS = 262_144
export interface ObjectSourceOptions {
  readRowsMaxCells?: number
  /** Fixed at admission (§4.4); 'cached' wraps a doc parsed with valuePolicy 'cached'. */
  valuePolicy?: ValuePolicy
}
/** Initial worker-side defaults; authoritative sheet metrics land with S2.3/S5. */
const DEFAULT_ROW_CSS = 20
const DEFAULT_COL_CSS = 64

const isPopulated = (cell: XlsxCell): boolean => cell.value !== null || cell.formula !== undefined
/** Under 'cached' a formula without a saved value is permanently unavailable (§4.4). */
const isUnavailable = (cell: XlsxCell, policy: ValuePolicy): boolean => isValueUnavailable(cell, policy)
function toPayload(cell: XlsxCell, policy: ValuePolicy = 'engine'): SingleCellPayload {
  const authority = valueAuthority(cell, policy)
  const payload: SingleCellPayload = {
    id: cellId(cell.row, cell.col), row: cell.row, col: cell.col, value: authority.value, styleIndex: cell.styleIndex,
    valueProvenance: authority.provenance,
  }
  if (cell.hasCachedValue !== undefined) payload.hasCachedValue = cell.hasCachedValue
  if (cell.valueIsError !== undefined) payload.valueIsError = authority.valueIsError
  if (cell.formula !== undefined) payload.formula = cell.formula
  if (cell.sharedFormula !== undefined) payload.sharedFormula = { ...cell.sharedFormula }
  if (cell.arrayRef !== undefined) payload.arrayRef = cell.arrayRef
  if (cell.dynamicArray !== undefined) payload.dynamicArray = cell.dynamicArray
  if (cell.formulaType !== undefined) payload.formulaType = cell.formulaType
  if (cell.ca !== undefined) payload.ca = cell.ca
  return Object.freeze(payload)
}
function toRowCell(cell: XlsxCell, policy: ValuePolicy): AuthoritativeCell {
  return isUnavailable(cell, policy) ? Object.freeze({ ...toPayload(cell, policy), unavailable: true as const }) : toPayload(cell, policy)
}

/** Accumulates one ColumnLanes run; rows must arrive strictly ascending. */
export class LaneBuilder {
  private rows: number[] = []
  private tags: number[] = []
  private numbers: number[] = []
  private styles: number[] = []
  private texts: string[] = []
  private textUnits = 0
  get count(): number { return this.rows.length }
  /** Same classification as the value index's cell-entry path (valueIndex.place). */
  add(row: number, value: string | number | boolean | null | undefined, valueIsError: boolean | undefined, styleIndex: number, unavailable: boolean): void {
    const { tag, num, text } = laneValue(value, valueIsError, unavailable)
    this.rows.push(row); this.tags.push(tag); this.numbers.push(num); this.styles.push(styleIndex); this.texts.push(text)
    this.textUnits += text.length
  }
  build(r0: number, r1: number): ColumnLanes {
    const n = this.rows.length
    const text = new Uint16Array(this.textUnits), textStart = new Uint32Array(n + 1)
    let at = 0
    for (let i = 0; i < n; i++) {
      textStart[i] = at
      const t = this.texts[i]
      // charCodeAt copies UTF-16 code units verbatim, lone surrogates included.
      for (let k = 0; k < t.length; k++) text[at++] = t.charCodeAt(k)
    }
    textStart[n] = at
    return {
      kind: 'lanes', r0, r1, rows: Uint32Array.from(this.rows), tags: Uint8Array.from(this.tags),
      numbers: Float64Array.from(this.numbers), styles: Uint32Array.from(this.styles), text, textStart,
    }
  }
}
/** Bytes a lanes run holds (row 4 + tag 1 + number 8 + style 4 per entry, 2 per code unit). */
export const laneBytes = (lanes: ColumnLanes): number => lanes.rows.length * 17 + lanes.text.length * 2 + lanes.textStart.length * 4

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

export class ObjectXlsxSheetAdapter implements SourceStore {
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
    readonly valuePolicy: ValuePolicy = 'engine',
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
  /** parse.ts turns an unreadable merge ref into an A1 merge; when the raw refs
   * are present, those are dropped (and diagnosed) exactly as the paged adapter does. */
  private validMerges(): XlsxSheet['mergeRanges'] {
    const raw = this.sheet.merges
    if (raw.length !== this.sheet.mergeRanges.length) return this.sheet.mergeRanges
    return this.sheet.mergeRanges.filter((_, i) => parseRange(raw[i]) !== null)
  }
  private get invalidMerges(): number {
    return this.sheet.mergeRanges.length - this.validMerges().length
  }
  /** Per-column caches retained by this store; it builds pages from the row index. */
  get cachedColumnCount(): number { return 0 }
  /** Coverage gaps this adapter can never close by waiting. */
  coverageDiagnostics(): WorkerDiagnostic[] {
    const out: WorkerDiagnostic[] = []
    if ((this.sheet.drawings?.length ?? 0) > 0) {
      out.push({ code: 'coverage-unresolvable', message: 'Drawing anchors are not available to the source adapter until S3.1 parser wiring' })
    }
    if (this.dynamicAnchors.length) {
      out.push({ code: 'coverage-unresolvable', message: 'Dynamic-array spill extents are not available to the source adapter until the S3.1 spill bridge', source: this.dynamicAnchors[0] })
    }
    if (this.invalidMerges) out.push(invalidMergeDiagnostic(this.invalidMerges))
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
    if (!cell) return { kind: 'blank' }
    const payload = toPayload(cell, this.valuePolicy)
    return isUnavailable(cell, this.valuePolicy) ? { kind: 'unavailable', cell: payload, reason: 'formula-value-unavailable' } : { kind: 'cell', cell: payload }
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
        if (cell.row === row.index && cell.col >= range.c0 && cell.col <= range.c1) cells.push(toRowCell(cell, this.valuePolicy))
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
    if (!intervals) return failedPage(this.stamp, failure('unsupported', 'Resident row-set refs are not resolvable by the object adapter'))
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
    if (request.encoding === 'lanes') return this.lanesPage(request, intervals, at, row, prefix)
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
      const payload = toPayload(found, this.valuePolicy)
      items.push(isUnavailable(found, this.valuePolicy)
        ? { kind: 'unavailable', row: found.row, cell: payload, reason: 'formula-value-unavailable' }
        : { kind: 'cell', row: found.row, cell: payload })
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
        items.push(toRowCell(cell, this.valuePolicy))
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
    const merges = this.validMerges()
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
  /** One lanes run per interval slice; a page covers at most `limit` rows. */
  private lanesPage(request: ColumnQuery, intervals: readonly SourceRowInterval[], at: number, row: number, prefix: string): Page<ColumnEntry> {
    const items: ColumnEntry[] = []
    let budget = request.limit
    while (at < intervals.length && budget > 0) {
      if (row < intervals[at].r0) row = intervals[at].r0
      const r1 = Math.min(intervals[at].r1, row + budget - 1)
      const lanes = new LaneBuilder()
      for (let k = lowerBound(this.rowIndexes, row); k < this.rows.length && this.rowIndexes[k] <= r1; k++) {
        const cell = this.cellIn(this.rows[k], request.col)
        if (cell) lanes.add(cell.row, cell.value, cell.valueIsError, cell.styleIndex, isUnavailable(cell, this.valuePolicy))
      }
      items.push(lanes.build(row, r1))
      budget -= r1 - row + 1
      row = r1 + 1
      if (row > intervals[at].r1) { at++; if (at < intervals.length) row = intervals[at].r0 }
    }
    return { items, completeness: 'complete', stamp: this.stamp, ...(at < intervals.length ? { nextCursor: `${prefix}:${at}:${row}` } : {}) }
  }
  async queryOverview(request: OverviewQuery, options: ReadOptions): Promise<Page<OverviewGroup>> {
    const error = this.guard(options)
    if (error) return failedPage(this.stamp, error)
    const source: OverviewCellSource = {
      aggregate: async (rect) => {
        const aggregate = emptyAggregate()
        for (let k = lowerBound(this.rowIndexes, rect.r0); k < this.rows.length && this.rowIndexes[k] <= rect.r1; k++) {
          for (const cell of this.rows[k].cells) {
            if (cell.row === this.rows[k].index && cell.col >= rect.c0 && cell.col <= rect.c1) addRecord(aggregate, cell.styleIndex, isPopulated(cell))
          }
        }
        return { kind: 'ok', aggregate }
      },
    }
    return answerOverview(request, source, this.stamp, options.signal)
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
  return buildGeometry([...sheet.rows].sort((a, b) => a.index - b.index), sheet.cols)
}
/** Source-space geometry from authored rows (ascending) and column specs; shared by both adapters. */
export function buildGeometry(
  rows: Iterable<{ index: number; heightPt?: number; hidden?: boolean }>,
  specs: readonly { min: number; max: number; widthChars?: number }[],
  hiddenIntervals: readonly SourceRowInterval[] = [],
): SheetGeometry {
  const rowOverrides: AxisSizeInterval[] = []
  const hidden: SourceRowInterval[] = []
  for (const row of rows) {
    if (!isValidCoordinate(row.index, 0)) continue
    if (row.heightPt !== undefined && Number.isFinite(row.heightPt) && row.heightPt >= 0) {
      rowOverrides.push({ start: row.index, end: row.index, sizeCss: row.heightPt * 96 / 72 })
    }
    if (row.hidden) {
      const last = hidden[hidden.length - 1]
      if (last && last.r1 + 1 === row.index) hidden[hidden.length - 1] = { r0: last.r0, r1: row.index }
      else if (!last || last.r1 < row.index) hidden.push({ r0: row.index, r1: row.index })
    }
  }
  const colOverrides: AxisSizeInterval[] = []
  for (const spec of specs) {
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
    baseHidden: createIntervalSet(hiddenIntervals.length ? hiddenIntervals : hidden),
  })
}
/** Wraps an already-parsed workbook by reference (no clone). Sheet ids use the
 * parser's stable workbook identity; documents are separated by docId. */
export { toPayload, isUnavailable, isPopulated, intersects, parseRange }
export const invalidMergeDiagnostic = (n: number): WorkerDiagnostic => ({ code: 'invalid-merge-ref', message: `Ignored ${n} merge${n === 1 ? '' : 's'} with an unreadable range` })
export function createObjectWorkbookSource(doc: XlsxDocument, docId: string, options: ObjectSourceOptions = {}): WorkbookSource {
  const maxCells = options.readRowsMaxCells ?? READ_ROWS_MAX_CELLS
  // The document's own policy decides (§4.4: one policy per document, never
  // mixed); an explicit option may only restate it.
  const policy = documentValuePolicy(doc)
  if (options.valuePolicy !== undefined && options.valuePolicy !== policy) {
    throw new RangeError(`A document parsed under the '${policy}' value policy cannot be served as '${options.valuePolicy}'`)
  }
  if (!Number.isSafeInteger(maxCells) || maxCells <= 0) throw new RangeError('Invalid readRows cell bound')
  const sheets = doc.sheets.map((sheet, index) => ({ sheet, sheetId: sheet.sheetId ?? String(index + 1) }))
  const ids = new Set<string>()
  for (const { sheetId } of sheets) {
    if (ids.has(sheetId)) throw new RangeError(`Duplicate sheet id ${sheetId}`)
    ids.add(sheetId)
  }
  const stores = new Map<string, ObjectXlsxSheetAdapter>()
  const geometries = new Map<string, SheetGeometry>()
  const find = (sheetId: string) => {
    const found = sheets.find(entry => entry.sheetId === sheetId)
    if (!found) throw new RangeError(`Unknown sheet ${sheetId}`)
    return found
  }
  const storeFor = (sheetId: string): ObjectXlsxSheetAdapter => {
    let store = stores.get(sheetId)
    if (!store) {
      const { sheet } = find(sheetId)
      const stamp = Object.freeze({ docId, sheetId, valueGen: INITIAL_GENERATION })
      store = new ObjectXlsxSheetAdapter(stamp, sheet, doc.tables?.filter(table => table.sheetId === sheetId), maxCells, policy)
      stores.set(sheetId, store)
    }
    return store
  }
  const status = valueStatus(policy, countLoadEvaluatedFormulas(doc))
  return {
    valuePolicy: policy,
    valueStatus: () => status,
    ...(status.message === undefined ? {} : { openDiagnostics: Object.freeze([{ code: SAVED_VALUES_DIAGNOSTIC, message: status.message }]) }),
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
    declaredFilters: () => doc.declaredFilters ?? [],
    async release() {
      await Promise.all([...stores.values()].map(store => store.release()))
      stores.clear()
      geometries.clear()
    },
  }
}
