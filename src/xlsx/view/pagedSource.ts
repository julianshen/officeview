/**
 * PagedXlsxSheetAdapter (spec §4.2, plan S1.2): the large-file SourceStore over
 * SourcePageStore pages written by the paged importer.
 *
 * Pages cover contiguous rows from 0, so the import frontier is one number per
 * sheet: rows at or before it are authoritative (absent records are proven
 * blank), rows after it are `unknown` while the import runs and a typed error
 * once it failed; nothing is ever reported blank before its page committed.
 * Values follow the `cached` policy (§4.4): saved values only, and a formula
 * without one is `unavailable`. Answers match ObjectXlsxSheetAdapter over the
 * same file parsed with valuePolicy 'cached' (adapter parity, §4.2).
 */
import type {
  AuthoritativeCell, CellRead, ColumnEntry, ColumnQuery, DrawingAnchor, OverviewGroup, OverviewQuery, Page, PopulatedQuery,
  RangeCompletion, ReadOptions, SingleCellPayload, SourceFailure, SourceRowInterval, SourceStamp, SourceStore,
  SparseRowPage, SpillExtent, StructuralResult, ValuePolicy, WorkerDiagnostic,
} from './contracts'
import { resolveDrawingAnchors } from './drawingBounds'
import { PackageError } from '../../core/package-error'
import { MAX_COLS, MAX_ROWS, cellId, isValidCoordinate, isValidSourceRect, type SourceCell, type SourceRect } from './coordinates'
import { INITIAL_GENERATION } from '../../worker/viewerProtocol'
import { failedPage, failure, rowSetIntervals, sameStamp, type SheetGeometry, type WorkbookSource } from '../../worker/viewerSession'
import type { MemoryPool, PoolEntry } from './tileCache'
import { LaneBuilder, READ_ROWS_MAX_CELLS, READ_ROWS_PAGE_ROWS, buildGeometry, intersects, invalidMergeDiagnostic } from './source'
import {
  F_BLOB, F_CACHED, F_ERROR, K_BOOL, K_EMPTY, K_ERROR, K_NUMBER, K_SST, K_TEXT, recordPopulated, textAt,
  type CellPagePayload, type PageSummary,
} from '../storage/sourcePages'
import { addRecord, answerOverview, emptyAggregate, mergeAggregate, type AggregateOutcome } from '../storage/overviewIndex'
import { isValueUnavailable, SAVED_VALUES_DIAGNOSTIC, SAVED_VALUES_STATUS, valueAuthority, valueStatus } from './values'
import type { PagedImportSession, PagedSheet, PagedWorkbook } from '../storage/pagedImport'
import type { DrawingMemory, ProfileId, SheetDrawings } from './contracts'
import { DRAWING_MEMORY_REFUSED, DrawingNodeLedger, drawingFootprint, unavailableDrawings } from './drawingPayload'
import type { XlsxDrawing, XlsxSheet } from '../types'
import { loadChartModule } from '../chartLoader'
import { cellDisplayText } from '../render'
import { semanticsFor } from './semanticsFor'

type RowStatus = { kind: 'ready' } | { kind: 'unknown'; pageId: string } | { kind: 'error'; failure: SourceFailure }

/** First index in [0,n) whose value is >= target. */
function lower(values: ArrayLike<number>, target: number, from = 0, to = values.length): number {
  let lo = from, hi = to
  while (lo < hi) {
    const mid = (lo + hi) >>> 1
    if (values[mid] < target) lo = mid + 1
    else hi = mid
  }
  return lo
}

/** A failure of an import that has ended: the import's own retryable flag
 * means "reopening may succeed"; re-reading this source never can. */
function ended(failure: SourceFailure | undefined): SourceFailure | undefined {
  return failure && failure.retryable ? Object.freeze({ ...failure, retryable: false }) : failure
}

export class PagedXlsxSheetAdapter implements SourceStore {
  readonly valuePolicy: ValuePolicy = 'cached'
  private released = false
  private drawingCache?: DrawingAnchor[]
  /** S13.2: the parsed drawing payload, retained (evictable) in the document pool so
   * repeated reads (stale retries, re-shown sheets) clone instead of re-parsing. */
  private drawingsCache?: { entry: PoolEntry; result: SheetDrawings }
  constructor(readonly stamp: SourceStamp, private readonly workbook: PagedWorkbook, private readonly sheet: PagedSheet, private readonly drawingState: PagedDrawingState = { nodes: new DrawingNodeLedger() }) {}

  get extent(): SourceStore['extent'] {
    const used = this.sheet.state === 'complete' ? this.sheet.used : undefined
    return used ? { rows: MAX_ROWS, cols: MAX_COLS, used: { ...used } } : { rows: MAX_ROWS, cols: MAX_COLS }
  }
  coverageDiagnostics(): WorkerDiagnostic[] {
    const out: WorkerDiagnostic[] = []
    if (this.sheet.drawingAnchors.unresolved > 0) out.push({ code: 'coverage-unresolvable', message: 'Some drawing anchors could not be read; their placement is unknown' })
    // No evaluator runs (`cached`): a dynamic array's saved ref is its published extent; without one it is unknown.
    const dynamic = this.sheet.arrayCells.find(a => a.dynamic && a.saved === undefined)
    if (dynamic) out.push({ code: 'coverage-unresolvable', message: 'A dynamic-array spill has no known extent', source: { row: dynamic.row, col: dynamic.col } })
    if (this.sheet.invalidMerges) out.push(invalidMergeDiagnostic(this.sheet.invalidMerges))
    if (this.sheet.state === 'complete' && this.sheet.tables === undefined) out.push({ code: 'coverage-unresolvable', message: 'The workbook table inventory could not be read; table bands are unknown' })
    return out
  }
  private guard(options: ReadOptions): SourceFailure | undefined {
    if (this.released) return failure('stale-source', 'Source released')
    // Abandoned or closed pages are gone: nothing may read as blank from here on.
    if (!this.workbook.store.live) return ended(this.workbook.failure) ?? failure('stale-source', 'Source pages were released')
    if (options.signal?.aborted) return failure('cancelled', 'Request cancelled')
    if (!options.stamp || !sameStamp(options.stamp, this.stamp)) return failure('stale-source', 'Source owner changed', true)
    return undefined
  }
  private status(row: number): RowStatus {
    if (row <= this.sheet.importedThroughRow) return { kind: 'ready' }
    if (this.sheet.state === 'incomplete' || this.workbook.completeness === 'incomplete') {
      return { kind: 'error', failure: ended(this.sheet.failure ?? this.workbook.failure) ?? failure('incomplete', 'The import did not finish') }
    }
    return { kind: 'unknown', pageId: `${this.sheet.sheetId}:pending:${this.sheet.importedThroughRow + 1}` }
  }
  /** Committed pages overlapping [r0,r1], in row order. */
  private pagesFor(r0: number, r1: number): readonly PageSummary[] {
    const pages = this.workbook.store.pages(this.sheet.sheetId)
    let lo = 0, hi = pages.length
    while (lo < hi) {
      const mid = (lo + hi) >>> 1
      if (pages[mid].r1 < r0) lo = mid + 1
      else hi = mid
    }
    const out: PageSummary[] = []
    for (let i = lo; i < pages.length && pages[i].r0 <= r1; i++) out.push(pages[i])
    return out
  }
  /** Cooperative cancellation at every page boundary: a cancelled query never
   * starts another page load and drops the one in flight. */
  private async load(page: PageSummary, signal: AbortSignal | undefined): Promise<CellPagePayload> {
    if (signal?.aborted) throw new PackageError('cancelled', 'Request cancelled')
    const payload = await this.workbook.store.page(this.sheet.sheetId, page.pageId)
    if (signal?.aborted) throw new PackageError('cancelled', 'Request cancelled')
    return payload
  }
  private isUnavailable(p: CellPagePayload, i: number): boolean {
    const meta = p.meta.get(i)
    return isValueUnavailable({ value: null, formula: meta?.formula, sharedFormula: meta?.sharedFormula, hasCachedValue: (p.flags[i] & F_CACHED) !== 0 }, 'cached')
  }
  private async valueOf(p: CellPagePayload, i: number): Promise<string | number | boolean | null> {
    switch (p.kinds[i]) {
      case K_EMPTY: return null
      case K_NUMBER: return p.nums[i]
      case K_SST: return (await this.workbook.store.sharedString(p.nums[i])) ?? ''
      // An oversized value is a blob record (§4.3.4), resolved lazily here so
      // every reader (cells, rows, scans, lanes, find, copy) sees the same value.
      case K_TEXT: case K_ERROR: return (p.flags[i] & F_BLOB) !== 0 ? this.workbook.store.blob(p.nums[i]) : textAt(p, i)
      case K_BOOL: return p.nums[i] === 1
      default: throw new RangeError('Unknown page value kind')
    }
  }
  private async payload(p: CellPagePayload, i: number): Promise<SingleCellPayload> {
    const row = p.rows[i], col = p.cols[i]
    const meta = p.meta.get(i)
    const hasCachedValue = (p.flags[i] & F_CACHED) !== 0
    const authority = valueAuthority({
      value: await this.valueOf(p, i), hasCachedValue, valueIsError: (p.flags[i] & F_ERROR) !== 0, formula: meta?.formula, sharedFormula: meta?.sharedFormula,
    }, this.valuePolicy)
    const out: SingleCellPayload = {
      id: cellId(row, col), row, col, value: authority.value, styleIndex: p.styles[i], valueProvenance: authority.provenance,
      hasCachedValue, valueIsError: authority.valueIsError,
    }
    if (meta) {
      if (meta.formula !== undefined) out.formula = meta.formula
      if (meta.sharedFormula !== undefined) out.sharedFormula = { ...meta.sharedFormula }
      if (meta.arrayRef !== undefined) out.arrayRef = meta.arrayRef
      if (meta.dynamicArray !== undefined) out.dynamicArray = meta.dynamicArray
      if (meta.formulaType !== undefined) out.formulaType = meta.formulaType
      if (meta.ca !== undefined) out.ca = meta.ca
    }
    return Object.freeze(out)
  }
  private async rowCell(p: CellPagePayload, i: number): Promise<AuthoritativeCell> {
    const cell = await this.payload(p, i)
    return this.isUnavailable(p, i) ? Object.freeze({ ...cell, unavailable: true as const }) : cell
  }
  /** Index of the record at (row,col) in a payload, or -1. */
  private find(p: CellPagePayload, row: number, col: number): number {
    const start = lower(p.rows, row)
    const end = lower(p.rows, row + 1, start)
    const at = lower(p.cols, col, start, end)
    return at < end && p.cols[at] === col ? at : -1
  }

  async readCell(at: { row: number; col: number }, options: ReadOptions): Promise<CellRead> {
    const error = this.guard(options)
    if (error) return { kind: 'error', code: error.code, retryable: error.retryable }
    if (!isValidCoordinate(at.row, at.col)) return { kind: 'error', code: 'invalid-coordinate', retryable: false }
    const status = this.status(at.row)
    if (status.kind === 'unknown') return { kind: 'unknown', pageId: status.pageId }
    if (status.kind === 'error') return { kind: 'error', code: status.failure.code, retryable: status.failure.retryable }
    const [page] = this.pagesFor(at.row, at.row)
    if (!page || page.aggregate.records === 0 || at.col < page.c0 || at.col > page.c1) return { kind: 'blank' }
    try {
      const p = await this.load(page, options.signal)
      const i = this.find(p, at.row, at.col)
      if (i < 0) return { kind: 'blank' }
      const cell = await this.payload(p, i)
      return this.isUnavailable(p, i) ? { kind: 'unavailable', cell, reason: 'formula-value-unavailable' } : { kind: 'cell', cell }
    } catch (cause) {
      const f = pageFailure(cause)
      return { kind: 'error', code: f.code, retryable: f.retryable }
    }
  }

  async readRows(range: SourceRect, cursor: string | undefined, options: ReadOptions): Promise<Page<SparseRowPage>> {
    const error = this.guard(options)
    if (error) return failedPage(this.stamp, error)
    if (!isValidSourceRect(range)) return failedPage(this.stamp, failure('invalid-source-rect', 'Invalid range'))
    const key = `pgrows:${range.r0}:${range.r1}:${range.c0}:${range.c1}:`
    let row = range.r0
    if (cursor !== undefined) {
      const next = cursor.startsWith(key) ? Number(cursor.slice(key.length)) : NaN
      if (!Number.isSafeInteger(next) || next < range.r0 || next > range.r1) return failedPage(this.stamp, failure('invalid-cursor', 'Invalid row cursor'))
      row = next
    }
    const items: SparseRowPage[] = []
    let visited = 0, cells = 0
    try {
      for (const page of this.pagesFor(row, range.r1)) {
        if (page.aggregate.records === 0 || page.c1 < range.c0 || page.c0 > range.c1) continue
        const p = await this.load(page, options.signal)
        let i = lower(p.rows, row)
        while (i < p.rows.length && p.rows[i] <= range.r1) {
          const r = p.rows[i]
          let end = i
          while (end < p.rows.length && p.rows[end] === r) end++
          const rowCells: AuthoritativeCell[] = []
          for (let k = lower(p.cols, range.c0, i, end); k < end && p.cols[k] <= range.c1; k++) rowCells.push(await this.rowCell(p, k))
          if (visited >= READ_ROWS_PAGE_ROWS || (items.length && cells + rowCells.length > READ_ROWS_MAX_CELLS)) {
            return { items, completeness: 'complete', stamp: this.stamp, nextCursor: `${key}${r}` }
          }
          visited++
          cells += rowCells.length
          if (rowCells.length) items.push({ row: r, cells: rowCells })
          i = end
        }
      }
    } catch (cause) {
      return { items, completeness: 'incomplete', stamp: this.stamp, error: pageFailure(cause) }
    }
    const status = this.status(range.r1)
    if (status.kind === 'error') return { items, completeness: 'incomplete', stamp: this.stamp, error: status.failure }
    if (status.kind === 'unknown') {
      // Committed rows are answered; the rest waits for the import frontier.
      const frontier = this.sheet.importedThroughRow + 1
      return { items, completeness: 'loading', stamp: this.stamp, nextCursor: `${key}${Math.max(row, frontier)}` }
    }
    return { items, completeness: 'complete', stamp: this.stamp }
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
    if (!intervals) return failedPage(this.stamp, failure('unsupported', 'Resident row-set refs are not resolvable by the paged adapter'))
    const prefix = `pgcol${request.col}`
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
    const lanes = request.encoding === 'lanes'
    const items: ColumnEntry[] = []
    let budget = request.limit
    let loading = false
    try {
      while (at < intervals.length && budget > 0) {
        if (row < intervals[at].r0) row = intervals[at].r0
        const status = this.status(row)
        if (status.kind !== 'ready') {
          // Everything from the frontier to the interval end is not yet known.
          const r1 = intervals[at].r1
          items.push(status.kind === 'unknown'
            ? { kind: 'unknown', r0: row, r1, pageId: status.pageId }
            : { kind: 'error', r0: row, r1, code: status.failure.code, retryable: status.failure.retryable })
          if (status.kind === 'unknown') loading = true
          budget--
          row = r1 + 1
          at++
          if (at < intervals.length) row = intervals[at].r0
          continue
        }
        const ready = Math.min(intervals[at].r1, this.sheet.importedThroughRow)
        if (lanes) {
          const r1 = Math.min(ready, row + budget - 1)
          const builder = new LaneBuilder()
          for (const page of this.pagesFor(row, r1)) {
            if (page.aggregate.records === 0 || request.col < page.c0 || request.col > page.c1) continue
            const p = await this.load(page, options.signal)
            for (let i = lower(p.rows, row); i < p.rows.length && p.rows[i] <= r1; i++) {
              if (p.cols[i] !== request.col) continue
              builder.add(p.rows[i], await this.valueOf(p, i), (p.flags[i] & F_ERROR) !== 0, p.styles[i], this.isUnavailable(p, i))
            }
          }
          items.push(builder.build(row, r1))
          budget -= r1 - row + 1
          row = r1 + 1
        } else {
          // Next record in this column at or after `row`, within the ready part.
          let found: { p: CellPagePayload; i: number } | undefined
          for (const page of this.pagesFor(row, ready)) {
            if (page.aggregate.records === 0 || request.col < page.c0 || request.col > page.c1) continue
            const p = await this.load(page, options.signal)
            for (let i = lower(p.rows, row); i < p.rows.length && p.rows[i] <= ready; i++) {
              if (p.cols[i] === request.col) { found = { p, i }; break }
            }
            if (found) break
          }
          if (!found) {
            items.push({ kind: 'blank', r0: row, r1: ready })
            budget--
            row = ready + 1
          } else {
            const r = found.p.rows[found.i]
            if (r > row) {
              items.push({ kind: 'blank', r0: row, r1: r - 1 })
              budget--
              row = r
              if (budget <= 0) break
            }
            const cell = await this.payload(found.p, found.i)
            items.push(this.isUnavailable(found.p, found.i)
              ? { kind: 'unavailable', row: r, cell, reason: 'formula-value-unavailable' }
              : { kind: 'cell', row: r, cell })
            budget--
            row = r + 1
          }
        }
        if (row > intervals[at].r1) { at++; if (at < intervals.length) row = intervals[at].r0 }
      }
    } catch (cause) {
      return { items, completeness: 'incomplete', stamp: this.stamp, error: pageFailure(cause) }
    }
    return {
      items, completeness: loading ? 'loading' : 'complete', stamp: this.stamp,
      ...(at < intervals.length ? { nextCursor: `${prefix}:${at}:${row}` } : {}),
    }
  }

  scanPopulated(request: PopulatedQuery, options: ReadOptions): AsyncIterable<Page<AuthoritativeCell>> {
    return this.scan(request, options, (p, i) => this.rowCell(p, i))
  }
  /** Same cells, order and pages as scanPopulated; coordinates only, no payload decode. */
  scanOccupied(request: PopulatedQuery, options: ReadOptions): AsyncIterable<Page<SourceCell>> {
    return this.scan(request, options, (p, i) => ({ row: p.rows[i], col: p.cols[i] }))
  }
  private async *scan<T>(request: PopulatedQuery, options: ReadOptions, item: (p: CellPagePayload, i: number) => T | Promise<T>): AsyncIterable<Page<T>> {
    const error = this.guard(options)
    if (error) { yield failedPage(this.stamp, error); return }
    const range = request.range
    if (!isValidSourceRect(range) || !Number.isSafeInteger(request.limit) || request.limit <= 0) {
      yield failedPage(this.stamp, failure('invalid-source-rect', 'Invalid populated query')); return
    }
    let items: T[] = []
    let row = range.r0
    for (;;) {
      try {
        for (const page of this.pagesFor(row, Math.min(range.r1, this.sheet.importedThroughRow))) {
          if (page.aggregate.populated === 0 || page.c1 < range.c0 || page.c0 > range.c1) { row = Math.max(row, page.r1 + 1); continue }
          const p = await this.load(page, options.signal)
          for (let i = lower(p.rows, row); i < p.rows.length && p.rows[i] <= range.r1; i++) {
            if (p.cols[i] < range.c0 || p.cols[i] > range.c1 || !recordPopulated(p, i)) continue
            const next = item(p, i)
            items.push(next instanceof Promise ? await next : next)
            if (items.length === request.limit) {
              yield { items, completeness: 'complete', stamp: this.stamp }
              items = []
              const again = this.guard(options)
              if (again) { yield failedPage(this.stamp, again); return }
            }
          }
          row = Math.max(row, page.r1 + 1)
        }
      } catch (cause) {
        yield { items, completeness: 'incomplete', stamp: this.stamp, error: pageFailure(cause) }
        return
      }
      if (row > range.r1) break
      const status = this.status(row)
      if (status.kind === 'error') { yield { items, completeness: 'incomplete', stamp: this.stamp, error: status.failure }; return }
      if (status.kind === 'ready') continue
      if (request.availability === 'loaded') { yield { items, completeness: 'loading', stamp: this.stamp }; return }
      // A scan waits for the import frontier instead of reporting rows it has not seen.
      const waited = await waitFor(this.workbook, options.signal)
      if (waited) { yield { items, completeness: 'incomplete', stamp: this.stamp, error: waited }; return }
    }
    yield { items, completeness: 'complete', stamp: this.stamp }
  }

  async queryStructural(rect: SourceRect, options: ReadOptions): Promise<StructuralResult> {
    const base = { stamp: this.stamp, merges: [], tables: [], spills: [], drawings: [] }
    const error = this.guard(options) ?? (isValidSourceRect(rect) ? undefined : failure('invalid-source-rect', 'Invalid range'))
    if (error) return { ...base, completeness: 'incomplete', missing: [rect], error }
    const sheet = this.sheet
    // Merges, tables and drawing references follow sheetData: unknown until the sheet finished.
    if (sheet.state !== 'complete') {
      const status = this.status(MAX_ROWS - 1)
      return {
        ...base, completeness: status.kind === 'error' ? 'incomplete' : 'loading', missing: [rect],
        ...(status.kind === 'error' ? { error: status.failure } : {}),
      }
    }
    const missing: SourceRect[] = []
    const addMissing = (domain: SourceRect) => {
      const r0 = Math.max(domain.r0, rect.r0), r1 = Math.min(domain.r1, rect.r1)
      const c0 = Math.max(domain.c0, rect.c0), c1 = Math.min(domain.c1, rect.c1)
      if (r0 <= r1 && c0 <= c1) missing.push({ r0, r1, c0, c1 })
    }
    const merges = sheet.mergeRects
      .map(m => ({ id: `merge:${m.r0}:${m.c0}:${m.r1}:${m.c1}`, range: m, anchor: { row: m.r0, col: m.c0 } }))
      .filter(m => isValidSourceRect(m.range) && intersects(m.range, rect))
    const tables = (sheet.tables ?? [])
      .map(t => ({ id: t.id, range: { r0: t.extent.firstRow, r1: t.extent.firstRow + t.extent.rows - 1, c0: t.extent.firstCol, c1: t.extent.firstCol + t.extent.cols - 1 }, headerRows: t.headerRowCount, totalsRows: t.totalsRowCount }))
      .filter(t => isValidSourceRect(t.range) && intersects(t.range, rect))
    const spills: SpillExtent[] = []
    for (const a of sheet.arrayCells) {
      // Saved values are authoritative (`cached`): the saved ref is a dynamic
      // array's published extent; without a readable one it may reach the sheet edge.
      if (a.dynamic && a.saved === undefined) addMissing({ r0: a.row, r1: MAX_ROWS - 1, c0: a.col, c1: MAX_COLS - 1 })
      if (!a.hasRef) continue
      const anchorRect = { r0: a.row, r1: a.row, c0: a.col, c1: a.col }
      if (!intersects(a.saved ?? anchorRect, rect) && !intersects(anchorRect, rect)) continue
      spills.push({ anchor: { row: a.row, col: a.col }, kind: a.dynamic ? 'dynamic' : 'fixed', ...(a.saved ? { saved: a.saved } : {}), ...(a.dynamic && a.saved ? { evaluated: a.saved } : {}) })
    }
    // Unreadable drawing anchors and an unavailable table inventory cannot prove "none here".
    if (sheet.drawingAnchors.unresolved > 0 || sheet.tables === undefined) addMissing(rect)
    const drawings = this.drawings().filter(d => d.kind === 'absolute' || intersects(d.bounds, rect))
    // Structure wrapped in mc:AlternateContent is unreadable, never known-empty (design note §3.4, RED 14a).
    if (sheet.metadata.structuralIncomplete.length) {
      return {
        ...base, merges, tables, spills, drawings, completeness: 'incomplete', missing: [rect],
        error: failure('unsupported', `${sheet.metadata.structuralIncomplete.join(', ')} metadata is wrapped in mc:AlternateContent and cannot be read`),
      }
    }
    return { ...base, merges, tables, spills, drawings, completeness: missing.length ? 'incomplete' : 'complete', missing }
  }
  /** Resolved once the sheet is complete (its row/column metadata is final), against the geometry the source publishes. */
  private drawings(): DrawingAnchor[] {
    const meta = this.sheet.metadata
    this.drawingCache ??= resolveDrawingAnchors(this.sheet.drawingAnchors.records,
      buildGeometry(meta.rowHeights.map(h => ({ index: h.row, heightPt: h.heightPt })), meta.cols, meta.hiddenRows, { hiddenCols: true }))
    return this.drawingCache
  }

  /** S13.2: the sheet's drawing content, read on demand once its metadata is final. */
  async readDrawings(options: ReadOptions, profile: ProfileId, memory?: DrawingMemory): Promise<SheetDrawings> {
    const error = this.guard(options)
    if (error) return unavailableDrawings(`Drawings could not be read: ${error.message}`, error)
    if (this.sheet.state !== 'complete') {
      const pending = this.sheet.failure ? ended(this.sheet.failure)! : failure('incomplete', 'The sheet is still loading', true)
      return unavailableDrawings(`Drawings are not available yet: ${pending.message}`, pending)
    }
    const meta = this.sheet.metadata
    if (!meta.drawingIds.length && !meta.legacyDrawingIds?.length) return { completeness: 'complete', drawings: [], images: [], diagnostics: [], bytes: 0 }
    const cached = this.drawingsCache
    if (cached?.entry.live) {
      cached.entry.touch()
      // The reply's buffers are transferred: it is a clone, charged while it lives.
      if (memory && !memory.set(cached.entry.bytes)) return memoryRefusal()
      return structuredClone(cached.result)
    }
    const controller = new AbortController()
    const abort = () => controller.abort()
    options.signal.addEventListener('abort', abort, { once: true })
    this.workbook.parts.signal.addEventListener('abort', abort, { once: true })
    if (options.signal.aborted || this.workbook.parts.signal.aborted) abort()
    try {
      const { loadPagedSheetDrawings, pagedDrawingLimits } = await import('../storage/pagedDrawings')
      const geometry = buildGeometry(meta.rowHeights.map(h => ({ index: h.row, heightPt: h.heightPt })), meta.cols, meta.hiddenRows, { hiddenCols: true })
      const nodes = this.drawingState.nodes, sheetId = this.sheet.sheetId
      const parsed = await loadPagedSheetDrawings(this.workbook.parts, { name: this.sheet.name, partPath: this.sheet.partPath, drawingIds: meta.drawingIds, legacyDrawingIds: meta.legacyDrawingIds },
        geometry, pagedDrawingLimits(profile, this.drawingState.payloadBytes), controller.signal, { memory, startNodes: nodes.start(sheetId), fillCharts: (drawings, owner) => this.fillCharts(drawings, owner, controller.signal) })
      nodes.record(sheetId, parsed.nodes)
      return this.retainDrawings(parsed.drawings, parsed.workingSet, memory)
    } catch (cause) {
      if (controller.signal.aborted) return unavailableDrawings('Drawing read cancelled', failure('cancelled', 'Request cancelled', true))
      const message = cause instanceof Error ? cause.message : String(cause)
      return unavailableDrawings(`Drawings could not be read: ${message}`, failure(cause instanceof PackageError ? cause.code : 'parse-error', message))
    } finally {
      options.signal.removeEventListener('abort', abort)
      this.workbook.parts.signal.removeEventListener('abort', abort)
    }
  }

  /** Fill uncached chart references from committed pages of the named sheets
   * (bounded per range by the chart point limit; an unread sheet stays empty). */
  private async fillCharts(drawings: XlsxDrawing[], owner: string, signal: AbortSignal): Promise<void> {
    const charts = await loadChartModule()
    const values = new Map<string, string | number | boolean | null>()
    const key = (sheet: string, row: number, col: number) => `${sheet.toLowerCase()}\u0000${row}\u0000${col}`
    for (const range of charts.pendingRanges(drawings, owner)) {
      const sheet = this.workbook.sheets.find(candidate => candidate.name.toLowerCase() === range.sheet.toLowerCase())
      if (!sheet) continue
      const r1 = Math.min(range.r1, range.r0 + charts.CHART_POINT_LIMIT - 1, sheet.importedThroughRow)
      const c1 = Math.min(range.c1, range.c0 + charts.CHART_POINT_LIMIT - 1)
      const pages = this.workbook.store.pages(sheet.sheetId)
      for (const page of pages) {
        if (page.r1 < range.r0 || page.r0 > r1 || page.aggregate.records === 0) continue
        if (signal.aborted) return
        const p = await this.workbook.store.page(sheet.sheetId, page.pageId)
        for (let i = lower(p.rows, range.r0); i < p.rows.length && p.rows[i] <= r1; i++) {
          if (p.cols[i] < range.c0 || p.cols[i] > c1) continue
          values.set(key(sheet.name, p.rows[i], p.cols[i]), await this.valueOf(p, i))
        }
      }
    }
    charts.fillDrawings(drawings, owner, (sheet, row, col) => values.get(key(sheet, row, col)))
  }

  /** Keeps a complete payload in the pool (when one is bound and it fits) and
   * answers with a clone; otherwise the parsed payload itself is the answer. */
  private retainDrawings(result: SheetDrawings, workingSet: number, memory?: DrawingMemory): SheetDrawings {
    const pool = this.drawingState.pool
    if (result.completeness !== 'complete' || !pool || pool.closed) return result
    const bytes = drawingFootprint(result)
    const holder: { entry?: PoolEntry } = {}
    holder.entry = pool.retain(`drawings:${this.stamp.docId}:${this.sheet.sheetId}`, {
      tier: 'rebuildable', bytes, owner: this.sheet.sheetId,
      onRelease: () => { if (this.drawingsCache?.entry === holder.entry) this.drawingsCache = undefined },
    }) ?? undefined
    if (!holder.entry) return result
    // The clone sits beside the retained copy until the reply is sent.
    if (memory && !memory.set(workingSet + bytes)) { holder.entry.release(); return result }
    this.drawingsCache = { entry: holder.entry, result }
    return structuredClone(result)
  }

  async queryOverview(request: OverviewQuery, options: ReadOptions): Promise<Page<OverviewGroup>> {
    const error = this.guard(options)
    if (error) return failedPage(this.stamp, error)
    return answerOverview(request, { aggregate: (rect, signal) => this.aggregate(rect, signal) }, this.stamp, options.signal)
  }
  /** Page summaries answer whole pages analytically; only a cut page is loaded. */
  private async aggregate(rect: SourceRect, signal: AbortSignal): Promise<AggregateOutcome> {
    const status = this.status(rect.r1)
    if (status.kind === 'unknown') return { kind: 'unknown', pageId: status.pageId }
    if (status.kind === 'error') return { kind: 'error', failure: status.failure }
    const sum = emptyAggregate()
    try {
      for (const page of this.pagesFor(rect.r0, rect.r1)) {
        if (page.aggregate.records === 0 || page.c1 < rect.c0 || page.c0 > rect.c1) continue
        if (rect.r0 <= page.r0 && page.r1 <= rect.r1 && rect.c0 <= page.c0 && page.c1 <= rect.c1) {
          mergeAggregate(sum, page.aggregate)
          continue
        }
        const p = await this.load(page, signal)
        for (let i = lower(p.rows, rect.r0); i < p.rows.length && p.rows[i] <= rect.r1; i++) {
          if (p.cols[i] >= rect.c0 && p.cols[i] <= rect.c1) addRecord(sum, p.styles[i], recordPopulated(p, i))
        }
      }
    } catch (cause) {
      return { kind: 'error', failure: pageFailure(cause) }
    }
    return { kind: 'ok', aggregate: sum }
  }

  async ensureComplete(rect: SourceRect, options: ReadOptions): Promise<RangeCompletion> {
    for (;;) {
      const error = this.guard(options) ?? (isValidSourceRect(rect) ? undefined : failure('invalid-source-rect', 'Invalid range'))
      if (error) return { range: rect, completeness: 'incomplete', stamp: this.stamp, missing: [rect], error }
      const status = this.status(rect.r1)
      if (status.kind === 'ready') return { range: rect, completeness: 'complete', stamp: this.stamp, missing: [] }
      if (status.kind === 'error') {
        const r0 = Math.max(rect.r0, this.sheet.importedThroughRow + 1)
        return { range: rect, completeness: 'incomplete', stamp: this.stamp, missing: [{ ...rect, r0 }], error: status.failure }
      }
      const waited = await waitFor(this.workbook, options.signal)
      if (waited) return { range: rect, completeness: 'incomplete', stamp: this.stamp, missing: [rect], error: waited }
    }
  }
  async release(): Promise<void> {
    this.released = true
    this.drawingsCache?.entry.release()
    this.drawingsCache = undefined
  }
}

/** Per-document drawing state shared by a paged source's sheet adapters (S13.2). */
export interface PagedDrawingState {
  /** Transfer budget override (tests, embedders). */
  payloadBytes?: number
  /** The document pool, once the session binds it. */
  pool?: MemoryPool
  nodes: DrawingNodeLedger
}
function memoryRefusal(): SheetDrawings {
  return unavailableDrawings('This sheet\'s drawings need more worker memory than the document has free and are not shown', failure('resource-limit', 'worker memory budget'), DRAWING_MEMORY_REFUSED)
}

function pageFailure(cause: unknown): SourceFailure {
  const code = (cause as { code?: unknown } | null)?.code
  if (code === 'storage-error' || code === 'quota' || code === 'parse-error' || code === 'cancelled') {
    return failure(code, (cause as Error).message, code !== 'parse-error')
  }
  return failure('page-error', cause instanceof Error ? cause.message : String(cause))
}
/** Waits for the next import change; a cancellation reply when the signal fires first. */
async function waitFor(workbook: PagedWorkbook, signal: AbortSignal | undefined): Promise<SourceFailure | undefined> {
  if (signal?.aborted) return failure('cancelled', 'Request cancelled')
  if (workbook.completeness !== 'loading') return undefined
  let onAbort: (() => void) | undefined
  const aborted = new Promise<'aborted'>((resolve) => {
    onAbort = () => resolve('aborted')
    signal?.addEventListener('abort', onAbort, { once: true })
  })
  try {
    const result = await Promise.race([workbook.changed().then(() => 'changed' as const), aborted])
    return result === 'aborted' ? failure('cancelled', 'Request cancelled') : undefined
  } finally {
    if (onAbort) signal?.removeEventListener('abort', onAbort)
  }
}

/** WorkbookSource over a paged import session; release() ends the session and removes its pages. */
export interface PagedSourceOptions {
  /** Drawing transfer budget override (tests, embedders); default drawingPayloadBudget(profile). */
  drawingPayloadBytes?: number
}
export function createPagedWorkbookSource(session: PagedImportSession, options: PagedSourceOptions = {}): WorkbookSource {
  const { workbook } = session
  const displaySheet: XlsxSheet = { name: '', rows: [], cols: [], merges: [], mergeRanges: [], semantics: semanticsFor(workbook.formatting.dateSystem) }
  const stores = new Map<string, PagedXlsxSheetAdapter>()
  const geometries = new Map<string, SheetGeometry>()
  const drawingState: PagedDrawingState = { nodes: new DrawingNodeLedger(), ...(options.drawingPayloadBytes !== undefined ? { payloadBytes: options.drawingPayloadBytes } : {}) }
  const find = (sheetId: string) => {
    const found = workbook.sheets.find(s => s.sheetId === sheetId)
    if (!found) throw new RangeError(`Unknown sheet ${sheetId}`)
    return found
  }
  const storeFor = (sheetId: string): PagedXlsxSheetAdapter => {
    let store = stores.get(sheetId)
    if (!store) {
      store = new PagedXlsxSheetAdapter(Object.freeze({ docId: workbook.docId, sheetId, valueGen: INITIAL_GENERATION }), workbook, find(sheetId), drawingState)
      stores.set(sheetId, store)
    }
    return store
  }
  return {
    valuePolicy: 'cached',
    // The formula count is only final once the import completed; until then
    // the status carries the spec's count-free wording.
    valueStatus: () => valueStatus('cached', workbook.completeness === 'complete' ? workbook.loadEvaluatedFormulas : undefined),
    // A getter, not a snapshot: read after the import completed it carries the
    // counted wording; read at open (import still running) the spec wording.
    get openDiagnostics() {
      const status = valueStatus('cached', workbook.completeness === 'complete' ? workbook.loadEvaluatedFormulas : undefined)
      return Object.freeze([{ code: SAVED_VALUES_DIAGNOSTIC, message: status.message ?? SAVED_VALUES_STATUS }])
    },
    sheets: Object.freeze(workbook.sheets.map(s => Object.freeze({ sheetId: s.sheetId, name: s.name }))),
    store: storeFor,
    geometry(sheetId) {
      const sheet = find(sheetId)
      let geometry = geometries.get(sheetId)
      if (!geometry) {
        // Row metadata arrives with the rows; cache only once the sheet is complete.
        geometry = buildGeometry(sheet.metadata.rowHeights.map(h => ({ index: h.row, heightPt: h.heightPt })), sheet.metadata.cols, sheet.metadata.hiddenRows, { hiddenCols: true })
        if (sheet.state === 'complete') geometries.set(sheetId, geometry)
      }
      return geometry
    },
    coverageDiagnostics: sheetId => storeFor(sheetId).coverageDiagnostics(),
    formatting: () => workbook.formatting,
    displayText: cell => cellDisplayText(displaySheet, { value: cell.value, style: workbook.formatting.styles[cell.styleIndex] }),
    declaredFilters: () => workbook.sheets.flatMap(s => s.declared ? [s.declared] : []),
    bindMemory(pool: MemoryPool) { workbook.store.bindMemory(pool); drawingState.pool = pool },
    async release() {
      await Promise.all([...stores.values()].map(store => store.release()))
      stores.clear()
      geometries.clear()
      await session.release()
    },
  }
}
