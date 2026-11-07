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
import { defaultColumnWidthPx } from '../columnWidth'
import type {
  AuthoritativeCell, CellRead, ColumnEntry, ColumnQuery, DrawingAnchor, OverviewGroup, OverviewQuery, Page, PopulatedQuery, RangeCompletion,
  ReadOptions, RemotePivotDisplay, SheetSparklines, SingleCellPayload, SourceFailure, SourceRowInterval, SourceStamp, SourceStore,
  SparseRowPage, SpillExtent, StructuralResult, ValueOverlay, ValuePolicy, WorkerDiagnostic,
} from './contracts'
import { resolveDrawingAnchors } from './drawingBounds'
import type { CfMemory, CfWorkBudget } from '../cf/limits'
import type { CfSheetRules, CfStamp, CfLiteral, CfValueSource, CfValuePage, CfLease, CfPrepared } from '../cf/model'
import { PackageError } from '../../core/package-error'
import { MAX_COLS, MAX_ROWS, cellId, formatA1, isValidCoordinate, isValidSourceRect, parseA1Range, type SourceCell, type SourceRect } from './coordinates'
import { uniqueRectList } from './rectIndex'
import { INITIAL_GENERATION } from '../../worker/viewerProtocol'
import { failedPage, failure, rowSetIntervals, sameStamp, type SheetGeometry, type WorkbookSource, type WorkerPivotAssets } from '../../worker/viewerSession'
import type { MemoryPool, PoolEntry } from './tileCache'
import { LaneBuilder, assertOverlayCellsProtected, overlayIncompleteColumnPage, READ_ROWS_MAX_CELLS, READ_ROWS_PAGE_ROWS, SourceStop, buildGeometry, requestBudget, withFrozen, intersects, invalidMergeDiagnostic, type StorePacing } from './source'
import {
  F_BLOB, F_CACHED, F_ERROR, K_BOOL, K_EMPTY, K_ERROR, K_NUMBER, K_SST, K_TEXT, recordPopulated, textAt, textPrefix,
  type CellPagePayload, type PageSummary,
} from '../storage/sourcePages'
import { addRecord, answerOverview, emptyAggregate, mergeAggregate, type AggregateOutcome } from '../storage/overviewIndex'
import { isValueUnavailable, SAVED_VALUES_DIAGNOSTIC, SAVED_VALUES_STATUS, valueAuthority, valueStatus } from './values'
import type { PagedImportSession, PagedSheet, PagedWorkbook } from '../storage/pagedImport'
import type { SheetMetadata } from '../storage/sheetRecords'
import { hostTask } from './scheduler'
import type { ChartWorkspace } from '../../charts/workspace'
import type { ChartModel } from '../../charts/model'
import type { DrawingMemory, ProfileId, SheetDrawings } from './contracts'
import { DRAWING_MEMORY_REFUSED, DrawingNodeLedger, drawingFootprint, unavailableDrawings } from './drawingPayload'
import type { XlsxCellStyle, XlsxDrawing, XlsxSheet } from '../types'
import type { XlsxCell } from '../types'
import type { SparklineAccess, SparklineResolveSheet } from '../sparklines/resolve'
import { resolveSparklineDataSheets } from '../sparklines/resolve'
import type { XlsxSparklineGroup } from '../sparklines/model'
import { loadChartModule } from '../chartLoader'
import { cellDisplayText } from '../render'
import { semanticsFor } from './semanticsFor'
import { MAX_PIVOT_DISPLAY_DXFS, MAX_REGION_STYLE_TABLES, REGION_STYLE_REPLY_OVERHEAD_BYTES, type RemoteRegionStyles, CF_REPLY_ENVELOPE_BYTES, estimateCfPreparationBytes, measureCfPropertiesBytes, versionCfRules, type CfReadRequest, type CfReadResult, type CfSharedDefinition, type CfReadStamp, type CfReadOptions } from './contracts'
import { estimatePayloadBytes } from './payloadBytes'
import { savedPivotRects } from '../pivots/stub'
import { SOURCE_ALLOCATION, adoptAllocation, type SourceAllocationOptions } from './allocationAccount'

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
function sameOverlayMerges(a: readonly SourceRect[] | undefined, b: readonly SourceRect[] | undefined): boolean {
  if (a === b) return true
  const left = a ?? [], right = b ?? []
  if (left.length !== right.length) return false
  const key = (rect: SourceRect) => `${rect.r0},${rect.c0},${rect.r1},${rect.c1}`
  const xOrder = [...left].map(key).sort(), yOrder = [...right].map(key).sort()
  for (let i = 0; i < xOrder.length; i++) {
    if (xOrder[i] !== yOrder[i]) return false
  }
  return true
}
interface OverlayState { overlay?: ValueOverlay; structureGen: number }
const overlayStates = new WeakMap<object, OverlayState>()
const workbookAdapters = new WeakMap<object, Map<string, PagedXlsxSheetAdapter>>()

/** Conservative structured-clone size for the sparkline model, recursive; depth is bounded by the sparkline model.
 * Repeated references are counted again. */
function sparklinePayloadTreeBytes(value: unknown): number {
  if (value === null || value === undefined) return 8
  if (typeof value === 'string') return 16 + value.length * 2
  if (typeof value !== 'object') return 8
  if (ArrayBuffer.isView(value)) return value.byteLength
  if (value instanceof ArrayBuffer) return value.byteLength
  if (Array.isArray(value)) {
    let bytes = 16 + value.length * 8
    for (const item of value) bytes += sparklinePayloadTreeBytes(item)
    return bytes
  }
  let bytes = 24
  for (const key in value) if (Object.prototype.hasOwnProperty.call(value, key)) {
    bytes += 16 + sparklinePayloadTreeBytes((value as Record<string, unknown>)[key])
  }
  return bytes
}

export class PagedXlsxSheetAdapter implements SourceStore {
  readonly valuePolicy: ValuePolicy = 'cached'
  private currentStamp: SourceStamp
  private released = false
  private drawingCache?: DrawingAnchor[]
  /** S13.2: the parsed drawing payload, retained (evictable) in the document pool so
   * repeated reads (stale retries, re-shown sheets) clone instead of re-parsing. */
  private drawingsCache?: { entry: PoolEntry; result: SheetDrawings; releaseCharts: () => void }
  /** Identity of the rule set whose shared-definition entry this adapter last
   * observed; replaced or released rules drop their share entry (live result
   * leases keep governing the bytes until their own release). */
  private sharedDefRules: CfSheetRules | undefined
  constructor(stamp: SourceStamp, private readonly workbook: PagedWorkbook, private readonly sheet: PagedSheet, private readonly drawingState: PagedDrawingState = { nodes: new DrawingNodeLedger() }, private readonly pacing: StorePacing = {}, registerForWorkbookReads = true) {
    this.currentStamp = stamp
    if (registerForWorkbookReads) {
      let adapters = workbookAdapters.get(workbook)
      if (!adapters) workbookAdapters.set(workbook, adapters = new Map())
      adapters.set(sheet.sheetId, this)
    }
  }

  get stamp(): SourceStamp { return this.currentStamp }
  private get valueOverlay(): ValueOverlay | undefined { return overlayStates.get(this)?.overlay }
  get structureGen(): number { return overlayStates.get(this)?.structureGen ?? 0 }
  setValueOverlay(overlay: ValueOverlay | undefined): void {
    const previous = overlayStates.get(this)
    // The overlay already installed is a no-op, checked before the (deferred T26 m1) assertion (Task 33 perf).
    if (previous?.overlay === overlay || (!previous && overlay === undefined)) return
    if (overlay) assertOverlayCellsProtected(overlay)
    const structureGen = (previous?.structureGen ?? 0) + (sameOverlayMerges(previous?.overlay?.merges, overlay?.merges) ? 0 : 1)
    overlayStates.set(this, { overlay, structureGen })
    this.currentStamp = Object.freeze({ ...this.currentStamp, valueGen: this.currentStamp.valueGen + 1 })
  }

  get extent(): SourceStore['extent'] {
    const used = this.sheet.state === 'complete' ? this.sheet.used : undefined
    const overlay = this.valueOverlay
    const bounds = overlay?.protectedRects ?? []
    const extra = bounds.length ? bounds.reduce((all, rect) => ({ r0: Math.min(all.r0, rect.r0), r1: Math.max(all.r1, rect.r1), c0: Math.min(all.c0, rect.c0), c1: Math.max(all.c1, rect.c1) }), { ...bounds[0] }) : undefined
    const merged = used && extra ? { r0: Math.min(used.r0, extra.r0), r1: Math.max(used.r1, extra.r1), c0: Math.min(used.c0, extra.c0), c1: Math.max(used.c1, extra.c1) } : used ?? extra
    return merged ? { rows: MAX_ROWS, cols: MAX_COLS, used: { ...merged } } : { rows: MAX_ROWS, cols: MAX_COLS }
  }
  private overlayCell(row: number, col: number): ReturnType<ValueOverlay['cell']> {
    const overlay = this.valueOverlay
    if (!overlay) return undefined
    for (const rect of overlay.rects) {
      if (row < rect.r0 || row > rect.r1 || col < rect.c0 || col > rect.c1) continue
      const cell = overlay.cell(row, col)
      if (cell !== undefined) return cell
    }
    return undefined
  }
  overlayStyle(at: { row: number; col: number }): XlsxCellStyle | null | undefined {
    const value = this.valueOverlay ? this.overlayCell(at.row, at.col) : undefined
    return value && value !== null ? (value.numFmtId !== undefined ? { numFmtId: value.numFmtId } : null) : undefined
  }
  private overlayPayload(row: number, col: number, value: NonNullable<ReturnType<ValueOverlay['cell']>>): SingleCellPayload {
    if (value === null) throw new RangeError('Blank overlay has no cell payload')
    return { id: cellId(row, col), row, col, value: value.value, styleIndex: 0, hasCachedValue: true, valueIsError: value.valueIsError ?? false, valueProvenance: 'engine' }
  }
  private mergeOverlayRows(items: SparseRowPage[], range: SourceRect, beforeRow = range.r1 + 1, fromRow = range.r0): { items: SparseRowPage[]; nextRow?: number } {
    const overlay = this.valueOverlay
    if (!overlay) return { items }
    const base = new Map(items.filter(item => item.row >= fromRow && item.row < beforeRow).map(item => [item.row, item]))
    const output: SparseRowPage[] = []
    let row = fromRow, visited = 0, cellCount = 0
    const nextCandidate = (at: number): number => {
      let next = Number.POSITIVE_INFINITY
      for (const value of base.keys()) if (value >= at) next = Math.min(next, value)
      for (const rect of overlay.rects) if (rect.r1 >= at && rect.c1 >= range.c0 && rect.c0 <= range.c1) next = Math.min(next, Math.max(at, rect.r0))
      return Math.min(next, beforeRow, range.r1 + 1)
    }
    while (row < beforeRow && row <= range.r1) {
      row = nextCandidate(row)
      if (row > range.r1 || row >= beforeRow) break
      if (visited >= READ_ROWS_PAGE_ROWS) return { items: output, nextRow: row }
      const cells = new Map<number, AuthoritativeCell>()
      for (const cell of base.get(row)?.cells ?? []) if (cell.col >= range.c0 && cell.col <= range.c1) cells.set(cell.col, cell)
      for (const rect of overlay.rects) {
        if (row < rect.r0 || row > rect.r1) continue
        for (let col = Math.max(range.c0, rect.c0); col <= Math.min(range.c1, rect.c1); col++) {
          const value = overlay.cell(row, col)
          if (value === undefined) continue
          cells.delete(col)
          if (value !== null) cells.set(col, this.overlayPayload(row, col, value))
        }
      }
      const rowCells = [...cells.values()].sort((a, b) => a.col - b.col)
      if (rowCells.length > READ_ROWS_MAX_CELLS) return { items: [], nextRow: row }
      if (visited > 0 && cellCount + rowCells.length > READ_ROWS_MAX_CELLS) return { items: output, nextRow: row }
      visited++; cellCount += rowCells.length
      if (rowCells.length) output.push({ row, cells: rowCells })
      row++
    }
    return { items: output }
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
    const overlay = this.overlayCell(at.row, at.col)
    if (overlay !== undefined) return overlay === null ? { kind: 'blank' } : { kind: 'cell', cell: this.overlayPayload(at.row, at.col, overlay) }
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

  /** Sparkline values distinguish numeric, boolean, empty, error and text, but
   * never inspect text content. Avoid materializing an SST/blob string solely
   * to map it to zero; this keeps a referenced huge text value bounded. */
  private async readSparklineCell(at: { row: number; col: number }, options: ReadOptions): Promise<CellRead> {
    const error = this.guard(options)
    if (error) return { kind: 'error', code: error.code, retryable: error.retryable }
    if (!isValidCoordinate(at.row, at.col)) return { kind: 'error', code: 'invalid-coordinate', retryable: false }
    const overlay = this.overlayCell(at.row, at.col)
    if (overlay !== undefined) return overlay === null ? { kind: 'blank' } : { kind: 'cell', cell: this.overlayPayload(at.row, at.col, overlay) }
    const status = this.status(at.row)
    if (status.kind === 'unknown') return { kind: 'unknown', pageId: status.pageId }
    if (status.kind === 'error') return { kind: 'error', code: status.failure.code, retryable: status.failure.retryable }
    const [page] = this.pagesFor(at.row, at.row)
    if (!page || page.aggregate.records === 0 || at.col < page.c0 || at.col > page.c1) return { kind: 'blank' }
    try {
      const payload = await this.load(page, options.signal), i = this.find(payload, at.row, at.col)
      if (i < 0) return { kind: 'blank' }
      const meta = payload.meta.get(i), hasCachedValue = (payload.flags[i] & F_CACHED) !== 0
      const value: XlsxCell['value'] = payload.kinds[i] === K_NUMBER ? payload.nums[i]
        : payload.kinds[i] === K_BOOL ? payload.nums[i] === 1
          : payload.kinds[i] === K_EMPTY ? null : ''
      const authority = valueAuthority({ value, hasCachedValue, valueIsError: (payload.flags[i] & F_ERROR) !== 0, formula: meta?.formula, sharedFormula: meta?.sharedFormula }, this.valuePolicy)
      const cell: SingleCellPayload = {
        id: cellId(at.row, at.col), row: at.row, col: at.col, value: authority.value,
        styleIndex: payload.styles[i], valueProvenance: authority.provenance, hasCachedValue, valueIsError: authority.valueIsError,
        ...(meta?.formula !== undefined ? { formula: meta.formula } : {}),
        ...(meta?.sharedFormula !== undefined ? { sharedFormula: { ...meta.sharedFormula } } : {}),
      }
      return authority.unavailable ? { kind: 'unavailable', cell, reason: 'formula-value-unavailable' } : { kind: 'cell', cell }
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
    const budget = requestBudget(() => this.guard(options), this.pacing)
    try {
      for (const page of this.pagesFor(row, range.r1)) {
        if (page.aggregate.records === 0 || page.c1 < range.c0 || page.c0 > range.c1) continue
        const p = await this.load(page, options.signal)
        let i = lower(p.rows, row)
        while (i < p.rows.length && p.rows[i] <= range.r1) {
          const r = p.rows[i]
          // Finding the row's end visits its records: work, like decoding.
          let end = i
          while (end < p.rows.length && p.rows[end] === r) {
            end++
            if (budget.spend(1)) await budget.pause()
          }
          const rowCells: AuthoritativeCell[] = []
          for (let k = lower(p.cols, range.c0, i, end); k < end && p.cols[k] <= range.c1; k++) {
            if (budget.spend(1)) await budget.pause()
            rowCells.push(await this.rowCell(p, k))
          }
          if (visited >= READ_ROWS_PAGE_ROWS || (items.length && cells + rowCells.length > READ_ROWS_MAX_CELLS)) {
            await budget.finish()
            const merged = this.mergeOverlayRows(items, range, r, row)
            return { items: merged.items, completeness: 'complete', stamp: this.stamp, nextCursor: `${key}${merged.nextRow ?? r}` }
          }
          visited++
          cells += rowCells.length
          if (rowCells.length) items.push({ row: r, cells: rowCells })
          i = end
        }
      }
      // A cancel or owner change landing after the last checkpoint never ends in success.
      await budget.finish()
    } catch (cause) {
      if (cause instanceof SourceStop) return failedPage(this.stamp, cause.failure)
      return { items, completeness: 'incomplete', stamp: this.stamp, error: pageFailure(cause) }
    } finally { budget.dispose() }
    const status = this.status(range.r1)
    if (status.kind === 'error') return { items, completeness: 'incomplete', stamp: this.stamp, error: status.failure }
    if (status.kind === 'unknown') {
      // Committed rows are answered; the rest waits for the import frontier.
      const frontier = this.sheet.importedThroughRow + 1
      const end = Math.max(row, frontier)
      const merged = this.mergeOverlayRows(items, range, end, row)
      return { items: merged.items, completeness: 'loading', stamp: this.stamp, nextCursor: `${key}${merged.nextRow ?? end}` }
    }
    const merged = this.mergeOverlayRows(items, range, range.r1 + 1, row)
    return { items: merged.items, completeness: 'complete', stamp: this.stamp, ...(merged.nextRow !== undefined ? { nextCursor: `${key}${merged.nextRow}` } : {}) }
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
    const startAt = at, startRow = row
    const overlay = this.valueOverlay
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
          if (!overlay) {
            for (const page of this.pagesFor(row, r1)) {
              if (page.aggregate.records === 0 || request.col < page.c0 || request.col > page.c1) continue
              const p = await this.load(page, options.signal)
              for (let i = lower(p.rows, row); i < p.rows.length && p.rows[i] <= r1; i++) {
                if (p.cols[i] !== request.col) continue
                builder.add(p.rows[i], await this.valueOf(p, i), (p.flags[i] & F_ERROR) !== 0, p.styles[i], this.isUnavailable(p, i))
              }
            }
          } else {
            const values = new Map<number, { value: string | number | boolean | null; error: boolean; style: number; unavailable: boolean }>()
            for (const page of this.pagesFor(row, r1)) {
              if (page.aggregate.records === 0 || request.col < page.c0 || request.col > page.c1) continue
              const p = await this.load(page, options.signal)
              for (let i = lower(p.rows, row); i < p.rows.length && p.rows[i] <= r1; i++) {
                if (p.cols[i] !== request.col) continue
                values.set(p.rows[i], { value: await this.valueOf(p, i), error: (p.flags[i] & F_ERROR) !== 0, style: p.styles[i], unavailable: this.isUnavailable(p, i) })
              }
            }
            for (const rect of overlay.rects) if (request.col >= rect.c0 && request.col <= rect.c1) {
              for (let current = Math.max(row, rect.r0); current <= Math.min(r1, rect.r1); current++) {
                const value = overlay.cell(current, request.col)
                if (value === undefined) continue
                if (value === null) values.delete(current)
                else values.set(current, { value: value.value, error: value.valueIsError ?? false, style: 0, unavailable: false })
              }
            }
            for (const [current, value] of [...values].sort((a, b) => a[0] - b[0])) builder.add(current, value.value, value.error, value.style, value.unavailable)
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
    const result: Page<ColumnEntry> = {
      items, completeness: loading ? 'loading' : 'complete', stamp: this.stamp,
      ...(at < intervals.length ? { nextCursor: `${prefix}:${at}:${row}` } : {}),
    }
    return overlay ? this.overlayColumnPage(result, request, intervals, startAt, startRow, prefix, overlay) : result
  }

  private overlayColumnPage(page: Page<ColumnEntry>, request: ColumnQuery, intervals: readonly SourceRowInterval[], startAt: number, startRow: number, prefix: string, overlay: ValueOverlay): Page<ColumnEntry> {
    if (!overlay || page.items.some(item => item.kind === 'lanes')) return page
    if (page.completeness !== 'complete') return overlayIncompleteColumnPage(page, request, intervals, startAt, prefix, overlay, row => this.overlayCell(row, request.col), (row, value) => this.overlayPayload(row, request.col, value))
    let endAt = intervals.length - 1, endRow = intervals[endAt]?.r1 ?? -1
    if (page.nextCursor) {
      const parts = page.nextCursor.split(':')
      const i = Number(parts[1]), r = Number(parts[2])
      if (parts.length === 3 && parts[0] === prefix && Number.isSafeInteger(i) && Number.isSafeInteger(r)) { endAt = i; endRow = r - 1 }
    }
    const cells = new Map<number, Extract<ColumnEntry, { kind: 'cell' | 'unavailable' }>>()
    for (const item of page.items) if (item.kind === 'cell' || item.kind === 'unavailable') cells.set(item.row, item)
    for (const rect of overlay.rects) {
      if (request.col < rect.c0 || request.col > rect.c1) continue
      for (let i = startAt; i <= endAt; i++) {
        const interval = intervals[i]
        if (!interval) continue
        for (let row = Math.max(interval.r0, rect.r0, i === startAt ? startRow : interval.r0); row <= Math.min(interval.r1, rect.r1, i === endAt ? endRow : interval.r1); row++) {
          const value = overlay.cell(row, request.col)
          if (value === undefined) continue
          if (value === null) cells.delete(row)
          else cells.set(row, { kind: 'cell', row, cell: this.overlayPayload(row, request.col, value) })
        }
      }
    }
    const output: ColumnEntry[] = []
    let continuation: { interval: number; row: number } | undefined
    for (let i = startAt; i <= endAt && !continuation; i++) {
      const interval = intervals[i]
      if (!interval) continue
      const r0 = Math.max(interval.r0, i === startAt ? startRow : interval.r0), r1 = Math.min(interval.r1, i === endAt ? endRow : interval.r1)
      let row = r0
      for (const entry of [...cells.values()].filter(cell => cell.row >= r0 && cell.row <= r1).sort((a, b) => a.row - b.row)) {
        if (entry.row > row) {
          if (output.length >= request.limit) { continuation = { interval: i, row }; break }
          output.push({ kind: 'blank', r0: row, r1: entry.row - 1 })
        }
        if (output.length >= request.limit) { continuation = { interval: i, row: entry.row }; break }
        output.push(entry); row = entry.row + 1
      }
      if (!continuation && row <= r1) {
        if (output.length >= request.limit) continuation = { interval: i, row }
        else output.push({ kind: 'blank', r0: row, r1 })
      }
    }
    const nextCursor = continuation ? `${prefix}:${continuation.interval}:${continuation.row}` : page.nextCursor
    const { nextCursor: _old, ...rest } = page
    return { ...rest, items: output, ...(nextCursor ? { nextCursor } : {}) }
  }

  scanPopulated(request: PopulatedQuery, options: ReadOptions & SourceAllocationOptions): AsyncIterable<Page<AuthoritativeCell>> {
    if (this.valueOverlay) return this.scanOverlay(request, options, cell => cell)
    return this.scan(request, options, (p, i) => this.rowCell(p, i))
  }
  /** Same cells, order and pages as scanPopulated; coordinates only, no payload decode. */
  scanOccupied(request: PopulatedQuery, options: ReadOptions & SourceAllocationOptions): AsyncIterable<Page<SourceCell>> {
    if (this.valueOverlay) return this.scanOverlay(request, options, cell => ({ row: cell.row, col: cell.col }))
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
    const budget = requestBudget(() => this.guard(options), this.pacing)
    try {
    for (;;) {
      try {
        for (const page of this.pagesFor(row, Math.min(range.r1, this.sheet.importedThroughRow))) {
          if (page.aggregate.populated === 0 || page.c1 < range.c0 || page.c0 > range.c1) { row = Math.max(row, page.r1 + 1); continue }
          const p = await this.load(page, options.signal)
          // Records are visited in batches up to the next checkpoint, each
          // batch charged once (spec §10.2 bound unchanged).
          for (let i = lower(p.rows, row); i < p.rows.length && p.rows[i] <= range.r1;) {
            let from = i
            const stop = Math.min(p.rows.length, i + budget.batch)
            for (; i < stop && p.rows[i] <= range.r1; i++) {
              if (p.cols[i] < range.c0 || p.cols[i] > range.c1 || !recordPopulated(p, i)) continue
              const next = item(p, i)
              items.push(next instanceof Promise ? await next : next)
              if (items.length === request.limit) {
                // No page escapes before its work is charged and checked: the
                // batch so far is spent (a due pause runs) and a check runs.
                if (budget.spend(i + 1 - from)) await budget.pause()
                from = i + 1
                budget.checkpoint()
                yield { items, completeness: 'complete', stamp: this.stamp }
                items = []
                const again = this.guard(options)
                if (again) { yield failedPage(this.stamp, again); return }
              }
            }
            if (budget.spend(i - from)) await budget.pause()
          }
          row = Math.max(row, page.r1 + 1)
        }
      } catch (cause) {
        if (cause instanceof SourceStop) { yield failedPage(this.stamp, cause.failure); return }
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
    try { await budget.finish() } catch (cause) {
      if (cause instanceof SourceStop) { yield failedPage(this.stamp, cause.failure); return }
      throw cause
    }
    yield { items, completeness: 'complete', stamp: this.stamp }
    } finally { budget.dispose() }
  }
  private async *scanOverlay<T>(request: PopulatedQuery, options: ReadOptions & SourceAllocationOptions, item: (cell: AuthoritativeCell) => T): AsyncIterable<Page<T>> {
    const error = this.guard(options)
    if (error) { yield failedPage(this.stamp, error); return }
    if (!isValidSourceRect(request.range) || !Number.isSafeInteger(request.limit) || request.limit <= 0) {
      yield failedPage(this.stamp, failure('invalid-source-rect', 'Invalid populated query')); return
    }
    let cursor: string | undefined, items: T[] = []
    let account = options[SOURCE_ALLOCATION]?.reserve()
    account?.charge(96)
    const publish = (page: Page<T>) => {
      if (account) adoptAllocation(account, page, account.bytes, [this.stamp])
      account?.release(); account = undefined
      return page
    }
    try { for (;;) {
      const page = await this.readRows(request.range, cursor, options)
      for (const row of page.items) for (const cell of row.cells) if (cell.value !== null || cell.formula !== undefined) {
        account?.charge(typeof cell.value === 'string' ? 384 + cell.value.length * 2 : 384)
        items.push(item(cell))
        if (items.length >= request.limit) { yield publish({ items, completeness: 'complete', stamp: this.stamp }); items = []; account = options[SOURCE_ALLOCATION]?.reserve(); account?.charge(96) }
      }
      if (page.error) { yield publish({ items, completeness: 'incomplete', stamp: this.stamp, error: page.error }); return }
      if (page.completeness !== 'complete') { yield publish({ items, completeness: page.completeness, stamp: this.stamp, ...(page.nextCursor ? { nextCursor: page.nextCursor } : {}) }); return }
      if (!page.nextCursor) break
      cursor = page.nextCursor
    }
    yield publish({ items, completeness: 'complete', stamp: this.stamp })
    } finally { account?.release() }
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
    const overlay = this.valueOverlay
    const merges = sheet.mergeRects
      .filter(m => !overlay?.rects.some(rect => m.r0 <= rect.r1 && rect.r0 <= m.r1 && m.c0 <= rect.c1 && rect.c0 <= m.c1))
      .map(m => ({ id: `merge:${m.r0}:${m.c0}:${m.r1}:${m.c1}`, range: m, anchor: { row: m.r0, col: m.c0 } }))
      .filter(m => isValidSourceRect(m.range) && intersects(m.range, rect))
    if (overlay) for (const merge of overlay.merges) if (intersects(merge, rect)) merges.push({ id: `pivot-merge:${merge.r0}:${merge.c0}:${merge.r1}:${merge.c1}`, range: merge, anchor: { row: merge.r0, col: merge.c0 } })
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
    // Every saved pivot is guarded with its page-field cells, overlay or not (S3-G). Vacated cells read
    // empty but stay guarded inside a saved area (S3-L): the guard is saved ∪ the overlay's protected
    // bounds, whatever the producer names, never the overlay alone.
    const groups = (sheet.metadata.pivotRegions ?? []).map(region => { const table = parseA1Range(region.ref); return table ? savedPivotRects(table, region.rowPageCount, region.colPageCount) : [] })
    const saved = uniqueRectList(groups.flat())
    const pivots = uniqueRectList([...saved, ...(overlay?.protectedRects ?? [])]).filter(pivot => intersects(pivot, rect))
    return { ...base, merges, tables, spills, drawings, ...(pivots?.length ? { pivots } : {}), completeness: missing.length ? 'incomplete' : 'complete', missing }
  }
  /** Resolved once the sheet is complete (its row/column metadata is final), against the geometry the source publishes. */
  private drawings(): DrawingAnchor[] {
    const meta = this.sheet.metadata
    this.drawingCache ??= resolveDrawingAnchors(this.sheet.drawingAnchors.records,
      buildGeometry(meta.rowHeights.map(h => ({ index: h.row, heightPt: h.heightPt })), meta.cols, meta.hiddenRows, { hiddenCols: true, columns: pagedColumns(this.workbook, meta) }))
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
      const geometry = buildGeometry(meta.rowHeights.map(h => ({ index: h.row, heightPt: h.heightPt })), meta.cols, meta.hiddenRows, { hiddenCols: true, columns: pagedColumns(this.workbook, meta) })
      const nodes = this.drawingState.nodes, sheetId = this.sheet.sheetId
      const parsed = await loadPagedSheetDrawings(this.workbook.parts, { name: this.sheet.name, partPath: this.sheet.partPath, drawingIds: meta.drawingIds, legacyDrawingIds: meta.legacyDrawingIds, slicerIds: meta.slicerRelationshipIds, timelineIds: meta.timelineRelationshipIds, stylesExtras: this.workbook.tableStyleExtras },
        geometry, pagedDrawingLimits(profile, this.drawingState.payloadBytes, this.workbook.pivotPartBytes), controller.signal, { memory, profile, chartDocument: this.drawingState.charts ??= {}, startNodes: nodes.start(sheetId), fillCharts: (drawings, owner, names, admit, workspace) => this.fillCharts(drawings, owner, controller.signal, names, admit, workspace as ChartWorkspace | undefined) })
      nodes.record(sheetId, parsed.nodes)
      // The payload now has its owner: the read's share stops holding its chart storage (D0). The retained entry
      // or the reply owns those bytes; the document facet keeps a retained chart's charge until the entry ends.
      parsed.chartScope?.handOff()
      const charts = parsed.chartScope ? await loadChartModule() : undefined
      const releaseCharts = () => { if (charts) forEachChart(parsed.drawings.drawings, chart => charts.releaseChartModel(chart)) }
      // A chart still waiting on another sheet's import is answered, never cached. An unretained reply keeps its
      // charts' facet charge while it lives: it is returned when the reply is released (collected after it is
      // posted or dropped; charts/own.ts), exactly once.
      if (parsed.drawings.pendingSheets?.length || parsed.uncacheable) return parsed.drawings
      return this.retainDrawings(parsed.drawings, parsed.workingSet - (parsed.payloadBytes ?? 0) - (parsed.modelBytes ?? 0), memory, releaseCharts)
    } catch (cause) {
      if (controller.signal.aborted) return unavailableDrawings('Drawing read cancelled', failure('cancelled', 'Request cancelled', true))
      const message = cause instanceof Error ? cause.message : String(cause)
      return unavailableDrawings(`Drawings could not be read: ${message}`, failure(cause instanceof PackageError ? cause.code : 'parse-error', message))
    } finally {
      options.signal.removeEventListener('abort', abort)
      this.workbook.parts.signal.removeEventListener('abort', abort)
    }
  }

  /** S8: transfer only the table descriptors and definitions this sheet uses.
   * The receiver grants the complete bounded payload before it is published. */
  async readRegionStyles(options: ReadOptions, grantBytes: number, memory?: DrawingMemory): Promise<RemoteRegionStyles> {
    const diagnostics: RemoteRegionStyles['diagnostics'][number][] = []
    const empty = (requiredBytes?: number): RemoteRegionStyles => ({ sheetId: this.sheet.sheetId, stamp: this.stamp, tables: [], diagnostics, ...(requiredBytes === undefined ? {} : { requiredBytes }) })
    const error = this.guard(options)
    if (error) {
      diagnostics.push({ kind: 'unsupported', feature: 'table-style-read-failed', message: error.message })
      return empty()
    }
    if (this.sheet.state !== 'complete') {
      diagnostics.push({ kind: 'unsupported', feature: 'table-style-read-failed', message: 'Table styles are not available until the sheet structure is complete.' })
      return empty()
    }
    const named = this.sheet.tables?.filter(table => Boolean(table.styleInfo?.name)) ?? []
    if (!named.length) return empty()
    if (!Number.isSafeInteger(grantBytes) || grantBytes <= 0 || named.length > MAX_REGION_STYLE_TABLES) {
      diagnostics.push({ kind: 'unsupported', feature: 'table-style-reply-limit', message: 'Table style metadata exceeds the receiver grant.' })
      return empty()
    }
    const tables: Array<RemoteRegionStyles['tables'][number]> = []
    // The built-in definitions load with the first styled table (spec §6.2), never with the paged worker.
    const { builtinTableStyle } = await import('../tableStyles/builtins')
    for (const table of named) {
      const name = table.styleInfo!.name!
      const def = this.workbook.tableStyleDefs?.get(name) ?? builtinTableStyle(name)
      if (!def?.table) {
        diagnostics.push({ kind: 'unsupported', feature: 'table-style-unknown', message: `Table "${table.displayName}" uses unknown table style "${name}"; stored cell formatting is used.` })
        continue
      }
      tables.push({ table: { id: table.id, displayName: table.displayName, sheetId: table.sheetId, extent: table.extent, headerRowCount: table.headerRowCount, totalsRowCount: table.totalsRowCount, styleInfo: table.styleInfo }, def })
    }
    const result: RemoteRegionStyles = { sheetId: this.sheet.sheetId, stamp: this.stamp, tables, ...(this.workbook.tableStyleTheme ? { theme: this.workbook.tableStyleTheme } : {}), diagnostics }
    const bytes = estimatePayloadBytes(result) + REGION_STYLE_REPLY_OVERHEAD_BYTES
    if (!Number.isSafeInteger(bytes) || bytes > grantBytes || (memory && !memory.set(bytes))) {
      diagnostics.splice(0, diagnostics.length, { kind: 'unsupported', feature: 'table-style-reply-limit', message: 'Table style metadata exceeds the receiver grant.' })
      return empty(bytes)
    }
    return result
  }

  /** S13: worker-owned pivot/control parts become bounded descriptors; the receiver builds display roles/buttons. */
  async readPivotDisplay(options: ReadOptions, grantBytes: number, memory?: DrawingMemory, profile: ProfileId = 'D'): Promise<RemotePivotDisplay> {
    const diagnostics: RemotePivotDisplay['diagnostics'][number][] = []
    const empty = (requiredBytes?: number): RemotePivotDisplay => ({ sheetId: this.sheet.sheetId, stamp: this.stamp, pivotTables: [], pivotStyles: [], diagnostics, ...(requiredBytes === undefined ? {} : { requiredBytes }) })
    const error = this.guard(options)
    if (error) { diagnostics.push({ kind: 'unsupported', feature: 'pivot-display-read-failed', message: error.message }); return empty() }
    if (this.sheet.state !== 'complete') { diagnostics.push({ kind: 'unsupported', feature: 'pivot-display-read-failed', message: 'Pivot display metadata is unavailable until the sheet structure is complete.' }); return empty() }
    if (!Number.isSafeInteger(grantBytes) || grantBytes <= 0) { diagnostics.push({ kind: 'unsupported', feature: 'pivot-display-reply-limit', message: 'Pivot display metadata exceeds the receiver grant.' }); return empty() }
    const pivotRegions = this.sheet.metadata.pivotRegions ?? []
    const parts = pivotRegions.length ? await (async () => {
      const { readPagedPivotDisplayParts } = await import('../storage/pagedDrawings')
      return readPagedPivotDisplayParts(this.workbook.parts, { sheetId: this.sheet.sheetId, sheetName: this.sheet.name, pivotRegions: new Map([[this.sheet.name, pivotRegions]]) }, profile, options.signal, memory, this.workbook.pivotPartBytes)
    })() : { pivotTables: [], diagnostics: [] }
    const displayDiagnostics = [...new Map([...(this.sheet.metadata.pivotDiagnostics ?? []), ...parts.diagnostics].map(item => [`${item.feature}:${'partPath' in item && item.partPath ? item.partPath : item.message}`, item])).values()]
    diagnostics.push(...displayDiagnostics)
    const pivotStyleNames = new Set(parts.pivotTables.map(table => table.styleInfo?.name).filter((name): name is string => Boolean(name)))
    const pivotStyles = [...pivotStyleNames].flatMap(name => { const style = this.workbook.tableStyleDefs?.get(name); return style ? [style] : [] })
    // The dxfs the pivots' `formats` apply (resident regions.ts reads them from the same styles.xml list).
    const dxfIds = new Set<number>()
    for (const table of parts.pivotTables) for (const format of table.formats) {
      if (format.action === 'formatting' && format.dxfId !== undefined && dxfIds.size < MAX_PIVOT_DISPLAY_DXFS) dxfIds.add(format.dxfId)
    }
    const dxfList = this.workbook.tableStyleExtras?.dxfs ?? []
    const pivotDxfs = [...dxfIds].flatMap(id => { const dxf = dxfList[id]; return dxf ? [{ id, dxf }] : [] })
    // Slicer and timeline appearance is resolved while reading drawings, which
    // retains the workbook's original style extras; these are not display-layer
    // descriptors and must not inflate the pivot-display reply.
    const result: RemotePivotDisplay = {
      sheetId: this.sheet.sheetId, stamp: this.stamp, pivotTables: parts.pivotTables, pivotStyles,
      ...(pivotDxfs.length ? { pivotDxfs } : {}),
      diagnostics,
      ...(this.workbook.tableStyleTheme ? { theme: this.workbook.tableStyleTheme } : {}),
    }
    const bytes = estimatePayloadBytes(result) + REGION_STYLE_REPLY_OVERHEAD_BYTES
    if (!Number.isSafeInteger(bytes) || bytes > grantBytes || (memory && !memory.set(bytes))) {
      diagnostics.push({ kind: 'unsupported', feature: 'pivot-display-reply-limit', message: 'Pivot display metadata exceeds the receiver grant.' })
      return empty(bytes)
    }
    return result
  }

  /**
   * Task 7: bounded CF preparation for the receiving CF layer. Validates the
   * receiver grant and the requested stamp, withholds while importing, admits
   * the bounded estimate before touching any page, then prepares through the
   * shared CF service and returns clone-safe winning states plus versioned
   * shared definitions.
   */
  async readConditionalFormatting(request: CfReadRequest, memory?: CfMemory, options?: CfReadOptions): Promise<CfReadResult> {
    // The CF service graph stays in a lazy chunk: it is only loaded when a
    // validated request actually builds. Account creation reserves nothing.
    // The workspace account is deliberately roomy: the receiver grant gates
    // the wire reply (so an exact-fit grant always admits an exact build),
    // while the transient workspace is bounded by the session caps.
    let account = memory
    if (!account) {
      const limits = await import('../cf/limits')
      account = limits.createCfMemory({})
    }
    const self = this
    // Single read of the sheet's rules: the same object feeds the live cfGen
    // below and the provider call, so they cannot disagree mid-call. The hash
    // is the worker's authority for definitions identity.
    const liveRules = this.sheet.conditionalFormatting
    // Release the previous rules' shared-definition identity when the sheet
    // replaced them (or dropped them); live result leases keep governing the
    // bytes until their own release.
    if (this.sharedDefRules !== undefined && this.sharedDefRules !== liveRules) {
      releaseSharedCfDefinitions(this.sharedDefRules)
    }
    this.sharedDefRules = liveRules
    // Admit the hash walk before normalizing rule content: even the adapter's
    // generation hash is construction, so a starved account refuses before
    // any rule value is read. Sizing is exact (lengths plus escape counts),
    // like the provider's: a flat per-key allowance cannot bound strings or
    // JSON escaping. The hold stays live across the provider call below, and
    // the admitted normalization travels with it so the version is computed
    // once per preparation instead of twice. One finally unwinds the hold on
    // every path out (idempotent on release).
    const hashBound = cfMetadataBound(liveRules, 0)
    if (hashBound === undefined) {
      return { ...cfReadError(this.sheet.sheetId, request.stamp, 'cf-resource-limit', 'cf-scope-too-large', 'Conditional formatting scope exceeds the bounded preparation area.'), requiredBytes: undefined }
    }
    const hashHold = account.reserve(hashBound)
    if (hashHold.kind !== 'reserved') {
      return { ...cfReadError(this.sheet.sheetId, request.stamp, 'cf-resource-limit', 'cf-budget-exceeded', 'Conditional formatting exceeds the admitted worker memory; retry when pressure drops.'), requiredBytes: undefined }
    }
    try {
      let liveCfGen: number
      let admitted: CfRemotePrepareOptions['admittedRules']
      if (liveRules && liveRules.rules.length > 0) {
        let exact: number | undefined
        try {
          exact = cfNormalizedBytes(liveRules.rules)
        } catch {
          return { ...cfReadError(this.sheet.sheetId, request.stamp, 'cf-preparation-failed', 'cf-preparation-withheld', 'Conditional formatting definitions could not be prepared.'), requiredBytes: undefined }
        }
        if (exact === undefined) {
          return { ...cfReadError(this.sheet.sheetId, request.stamp, 'cf-resource-limit', 'cf-scope-too-large', 'Conditional formatting definitions exceed the bounded preparation area.'), requiredBytes: undefined }
        }
        const grownTarget = hashBound + exact
        if (!Number.isSafeInteger(grownTarget) || !hashHold.grow(grownTarget)) {
          return { ...cfReadError(this.sheet.sheetId, request.stamp, 'cf-resource-limit', 'cf-budget-exceeded', 'Conditional formatting exceeds the admitted worker memory; retry when pressure drops.'), requiredBytes: undefined }
        }
        let normalized: string
        try {
          normalized = normalizeCfRules(liveRules.rules)
        } catch {
          return { ...cfReadError(this.sheet.sheetId, request.stamp, 'cf-preparation-failed', 'cf-preparation-withheld', 'Conditional formatting definitions could not be prepared.'), requiredBytes: undefined }
        }
        if (normalized.length * 2 > exact) {
          return { ...cfReadError(this.sheet.sheetId, request.stamp, 'cf-resource-limit', 'cf-budget-exceeded', 'Conditional formatting exceeds the admitted worker memory; retry when pressure drops.'), requiredBytes: undefined }
        }
        liveCfGen = versionCfRules(normalized)
        admitted = { rules: liveRules, normalized, version: liveCfGen }
      } else {
        liveCfGen = versionCfRules('[]')
      }
    const source: CfValueSource = {
      bounds: async (_id: string) => {
        // A complete sheet with no used extent is still readable: the service
        // clamps requested rects and synthesizes blanks for unstored cells, so
        // blank-targeting rules evaluate instead of pending forever. Only an
        // incomplete import withholds as pending.
        if (self.sheet.state !== 'complete') return undefined
        const used = self.sheet.used
        if (!used) return { firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 16383 }
        return { firstRow: used.r0, lastRow: used.r1, firstCol: used.c0, lastCol: used.c1 }
      },
      // Read the live generation on every call: the service re-reads this
      // stamp to withhold a preparation whose source committed mid-build.
      // cfGen is the live content hash (authoritative); view/clock/owner are
      // client-managed labels echoed for cache keying, not worker authority.
      stamp: (_id: string) => ({
        owner: request.stamp.owner,
        sheetId: self.sheet.sheetId,
        workbookValueGen: self.stamp.valueGen,
        cfGen: liveCfGen,
        viewGen: request.stamp.viewGen,
        structureGen: request.stamp.structureGen,
        clockGen: request.stamp.clockGen,
      }),
      pages: this.cfPages(request, account, options?.signal),
      }
      return await prepareCfRemote(request, this.sheet.sheetId, source, liveRules, account,
        admitted === undefined ? options : { ...options, admittedRules: { rules: liveRules, normalized: admitted.normalized, version: admitted.version } })
    } finally {
      try { hashHold.release() } catch { /* release-only path */ }
    }
  }

  /** CfValueSource over this sheet's paged rows: 64-cell pages admitted
   * before construction, released by the consumer after each page. */
  private cfPages(request: CfReadRequest, memory: CfMemory, signal?: AbortSignal): CfValueSource['pages'] {
    const self = this
    void request
    return pagedStoreCfPages({ readRows: (range, cursor, options) => self.readRows(range, cursor, options), get stamp() { return self.stamp } }, memory, signal)
  }

  /** S3: asynchronously snapshot only the source cells referenced by this
   * sheet's groups, then run the same synchronous resolver as resident XLSX.
   * The worker owns both the metadata and every page read (including cross-sheet
   * ranges); the receiver grant is checked before any point arrays are built. */
  async readSparklines(options: ReadOptions, grantBytes: number, memory?: DrawingMemory, replyEnvelopeBytes = 1024): Promise<SheetSparklines> {
    const groups = this.sheet.metadata.sparklineGroups ?? []
    const sourceDiagnostics = this.sheet.metadata.sparklineDiagnostics ?? []
    if (!groups.length) return { sparklines: [], groups: [], diagnostics: sourceDiagnostics.slice(0, 64).map(item => ({ ...item, message: item.message.slice(0, 256) })) }
    const sheets = this.workbook.sheets
    const byName = new Map(sheets.map(sheet => [sheet.name.toLowerCase(), sheet] as const))
    let sparkCount = 0, points = 0, axisPoints = 0, dataNameUnits = 0
    let planned = 0
    for (const group of groups) for (const spark of group.sparklines) {
      const area = spark.data
      if (!spark.host || !area) continue
      const target = area.sheet === undefined ? this.sheet : byName.get(area.sheet.toLowerCase())
      if (!target) continue
      const dataPoints = area.r0 === area.r1 ? area.c1 - area.c0 + 1 : area.r1 - area.r0 + 1
      const date = group.dateAxis ? group.dateRange : undefined
      const datePoints = date ? Math.min(dataPoints, date.r0 === date.r1 ? date.c1 - date.c0 + 1 : date.r1 - date.r0 + 1) : 0
      sparkCount++
      points += dataPoints
      axisPoints += datePoints
      dataNameUnits += target.name.length
      if (++planned % 128 === 0) {
        if (options.signal.aborted) return { sparklines: [], groups: [], diagnostics: [{ kind: 'unsupported', feature: 'sparkline-cancelled', message: 'Sparkline resolution was cancelled' }] }
        await new Promise<void>(resolve => setTimeout(resolve, 0))
      }
    }
    // Required wire bound: Σ(points × 8) + Σ(axis points × 8) + sparklines × 96
    // + groups × 256 + diagnostics UTF-16 × 2. The safe preflight extends that
    // formula for the other cloned fields: sourceRows/bridged/high/low arrays
    // (at most 64 B per point), per-record object/string overhead, the exact
    // bounded group tree (including its source definitions), and diagnostic
    // object slots. The axis positions are the additional Task 10 term.
    const groupTreeBytes = sparklinePayloadTreeBytes(groups)
    const baseBound = points * 8 + axisPoints * 8 + sparkCount * 96 + groups.length * 256 + 64 * 256 * 2
    const metadataBound = points * 64 + sparkCount * 512 + dataNameUnits * 2 + Math.max(0, groupTreeBytes - groups.length * 256) + 64 * 640
    const estimated = baseBound + metadataBound + replyEnvelopeBytes
    // Worker-only snapshot structures coexist with the wire result while it is
    // resolved. Reserve them before constructing coordinate sets or cell maps.
    const preparationBytes = estimated + (points + axisPoints) * 512
    if (!Number.isSafeInteger(preparationBytes) || estimated > grantBytes || memory && !memory.set(preparationBytes)) {
      return { sparklines: [], groups: [], diagnostics: [{ kind: 'unsupported', feature: 'sparkline-resource-limit', message: 'Sparkline data exceeds the receiver memory grant' }] }
    }
    // Keep room within the 64-entry host cap for resolver diagnostics generated
    // by the snapshot (missing data, incomplete groups, unsupported axes).
    const boundedSourceDiagnostics = sourceDiagnostics.slice(0, 48).map(item => ({ ...item, message: item.message.slice(0, 256) }))
    const stores = new Map<string, PagedXlsxSheetAdapter>()
    const sheetFor = (sheet: PagedSheet): PagedXlsxSheetAdapter => {
      let store = stores.get(sheet.sheetId)
      if (!store) {
        const existing = workbookAdapters.get(this.workbook)?.get(sheet.sheetId)
        store = existing && !existing.released ? existing : sheet.sheetId === this.sheet.sheetId ? this : new PagedXlsxSheetAdapter(
          Object.freeze({ ...this.stamp, sheetId: sheet.sheetId }), this.workbook, sheet, this.drawingState, this.pacing, false,
        )
        stores.set(sheet.sheetId, store)
      }
      return store
    }
    const cellsBySheet = new Map<string, Map<string, XlsxCell>>()
    const requested = new Map<string, Set<string>>()
    let collected = 0
    for (const group of groups) for (const spark of group.sparklines) {
      if (!spark.host || !spark.data) continue
      const target = spark.data.sheet === undefined ? this.sheet : byName.get(spark.data.sheet.toLowerCase())
      if (!target) continue
      const addRange = (sheet: PagedSheet, rect: { r0: number; r1: number; c0: number; c1: number }) => {
        let set = requested.get(sheet.sheetId)
        if (!set) requested.set(sheet.sheetId, set = new Set())
        if (rect.r0 === rect.r1) for (let col = rect.c0; col <= rect.c1; col++) set.add(`${rect.r0}:${col}`)
        else for (let row = rect.r0; row <= rect.r1; row++) set.add(`${row}:${rect.c0}`)
      }
      addRange(target, spark.data)
      const date = group.dateAxis ? group.dateRange : undefined
      if (date) {
        const dateTarget = date.sheet === undefined ? this.sheet : byName.get(date.sheet.toLowerCase())
        if (dateTarget) {
          const dataPoints = spark.data.r0 === spark.data.r1 ? spark.data.c1 - spark.data.c0 + 1 : spark.data.r1 - spark.data.r0 + 1
          const datePoints = date.r0 === date.r1 ? date.c1 - date.c0 + 1 : date.r1 - date.r0 + 1
          const pairedPoints = Math.min(dataPoints, datePoints)
          // Date pairing is positional; cells after the data range's length cannot contribute to slots or values.
          const pairedDate = date.r0 === date.r1
            ? { ...date, c1: date.c0 + pairedPoints - 1 }
            : { ...date, r1: date.r0 + pairedPoints - 1 }
          addRange(dateTarget, pairedDate)
        }
      }
      if (++collected % 128 === 0) {
        if (options.signal.aborted) return { sparklines: [], groups: [], diagnostics: [{ kind: 'unsupported', feature: 'sparkline-cancelled', message: 'Sparkline resolution was cancelled' }] }
        await new Promise<void>(resolve => setTimeout(resolve, 0))
      }
    }
    const pendingSheetIds = new Set<string>()
    for (const sheet of sheets) {
      const coords = requested.get(sheet.sheetId)
      if (!coords) continue
      const store = sheetFor(sheet), result = new Map<string, XlsxCell>()
      cellsBySheet.set(sheet.sheetId, result)
      let visited = 0
      for (const key of coords) {
        if (options.signal.aborted) return { sparklines: [], groups: [], diagnostics: [{ kind: 'unsupported', feature: 'sparkline-cancelled', message: 'Sparkline resolution was cancelled' }] }
        const [row, col] = key.split(':').map(Number)
        const read = await store.readSparklineCell({ row, col }, { signal: options.signal, stamp: store.stamp })
        if (read.kind === 'cell' || read.kind === 'unavailable') {
          const cell = read.cell
          result.set(key, { ...cell, row, col, ref: formatA1({ row, col }) })
        } else if (read.kind === 'unknown' || read.kind === 'error') {
          pendingSheetIds.add(sheet.sheetId)
          return { sparklines: [], groups: [], pendingSheetIds: [...pendingSheetIds], diagnostics: [{ kind: 'unsupported', feature: 'sparkline-data-unavailable', message: 'Sparkline data is not fully available in the worker' }] }
        }
        if (++visited % 128 === 0) await new Promise<void>(resolve => setTimeout(resolve, 0))
      }
    }
    const resolveSheet = (sheet: PagedSheet, isHost: boolean): SparklineResolveSheet => {
      const cells = cellsBySheet.get(sheet.sheetId) ?? new Map<string, XlsxCell>()
      const access: SparklineAccess = {
        rowHidden: row => sheet.metadata.hiddenRows.some(run => run.r0 <= row && row <= run.r1),
        colHidden: col => sheet.metadata.cols.some(spec => spec.hidden && spec.min <= col && col <= spec.max),
        cell: (row, col) => cells.get(`${row}:${col}`),
      }
      return {
        name: sheet.name, ...(isHost ? { sparklineGroups: groups } : {}),
        mergeRanges: sheet.mergeRects.map(range => ({ minRow: range.r0, maxRow: range.r1, minCol: range.c0, maxCol: range.c1 })),
        diagnostics: isHost ? [...boundedSourceDiagnostics] : [], maxDiagnostics: 64, access,
      }
    }
    const models = sheets.map(sheet => resolveSheet(sheet, sheet.sheetId === this.sheet.sheetId))
    resolveSparklineDataSheets(models, { cached: true })
    const hostModel = models.find(model => model.name === this.sheet.name)!
    const outputGroups: XlsxSparklineGroup[] = []
    const groupIndex = new Map<XlsxSparklineGroup, number>()
    const output: SheetSparklines['sparklines'] = []
    for (const row of hostModel.resolvedSparklines?.values() ?? []) for (const resolved of row.values()) {
      let index = groupIndex.get(resolved.group)
      if (index === undefined) { index = outputGroups.length; groupIndex.set(resolved.group, index); outputGroups.push(resolved.group) }
      output.push({
        host: resolved.host, dataSheet: resolved.dataSheet, sourceRows: resolved.sourceRows, groupIndex: index,
        values: Float64Array.from(resolved.values, value => value === null ? Number.NaN : value), bridged: resolved.bridged,
        firstIndex: resolved.firstIndex, lastIndex: resolved.lastIndex, highIndices: resolved.highIndices, lowIndices: resolved.lowIndices,
        min: resolved.min, max: resolved.max,
        ...(resolved.axis ? { axis: { positions: Float64Array.from(resolved.axis.positions), slot: resolved.axis.slot } } : {}),
      })
    }
    const diagnostics = (hostModel.diagnostics ?? []).slice(0, 64).map(item => ({ ...item, message: item.message.slice(0, 256) }))
    const actualBytes = sparklinePayloadTreeBytes({ sparklines: output, groups: outputGroups, diagnostics })
    // The final WorkerReply adds owner, diagnostics and replyBytes after this
    // source result; the kernel passes its exact payload-free envelope bound.
    if (actualBytes + replyEnvelopeBytes > grantBytes || memory && !memory.set(actualBytes + replyEnvelopeBytes)) {
      return { sparklines: [], groups: [], diagnostics: [{ kind: 'unsupported', feature: 'sparkline-resource-limit', message: 'Sparkline data exceeds the receiver memory grant' }] }
    }
    return { sparklines: output, groups: outputGroups, diagnostics }
  }

  /** Fill uncached chart references from committed pages, as a bounded stream:
   * 1. The point workspace (points x CHART_CELL_BYTES) is admitted before any
   *    page is touched; refusal makes the charts placeholders.
   * 2. The consumed cells are grouped by page from the ordered page summaries
   *    (binary search, no page load) with the label need merged per cell.
   * 3. Pages are processed in order, one at a time: each needed cell is
   *    extracted into compact chart storage (numbers in a Float64Array,
   *    labels capped at CHART_LABEL_LIMIT) and the page is released before the
   *    next is loaded, so no decoded page outlives its source charge.
   * 4. Every string is admitted before it is built, against a running
   *    reservation: a raw string by its capped prefix (a numeric-role string
   *    longer than the label limit is not a number and is never built), a
   *    formatted label by an upper bound of its format's output, then trued
   *    up to the actual label. The total is capped at CHART_TEXT_BYTES_LIMIT.
   *    Refusal rolls the reservation back and refuses the read.
   * Worst case: one decoded page plus the admitted chart storage.
   * Cells past a loading sheet's frontier stay unknown (pendingSheets); a
   * failed sheet turns its charts into placeholders and is never cached. */
  private async fillCharts(drawings: XlsxDrawing[], owner: string, signal: AbortSignal, names: readonly { name: string; source: string; localSheetIndex?: number }[] = [], admit: (bytes: number) => boolean = () => true, workspace?: ChartWorkspace): Promise<{ pendingSheets: string[]; uncacheable?: boolean; refused?: boolean; diagnostics?: WorkerDiagnostic[] }> {
    const charts = await loadChartModule()
    charts.stampWorkbookEpoch(drawings, this.workbook.formatting.dateSystem)
    charts.resolveDrawingNames(drawings, owner, charts.chartNameResolver(names, this.workbook.sheets))
    // An over-limit analytic chart is refused with its point-limit reason before any memory admission.
    charts.refuseAnalyticOverLimit(drawings)
    const sheetOf = (name: string) => this.workbook.sheets.find(candidate => candidate.name.toLowerCase() === name.toLowerCase())
    const diagnostics: WorkerDiagnostic[] = []
    const failed = charts.markUnreadableCharts(drawings, owner, name => {
      const sheet = sheetOf(name)
      if (!sheet || sheet.state !== 'incomplete') return undefined
      const reason = `Chart data on sheet “${sheet.name}” could not be read${sheet.failure?.message ? `: ${sheet.failure.message}` : ''}`
      if (!diagnostics.some(d => d.message === reason)) diagnostics.push({ code: 'chart-data-unavailable', message: reason })
      return reason
    })
    const done = (extra: { refused?: boolean } = {}) => ({ pendingSheets: [...pending], ...extra, ...(failed ? { uncacheable: true, diagnostics } : {}) })
    const pending = new Set<string>()
    const lines = charts.pendingLines(drawings, owner).filter(line => sheetOf(line.sheet)?.state !== 'incomplete')
    const points = lines.reduce((total, line) => total + line.count, 0)
    const base = points * (charts.CHART_CELL_BYTES + 8)
    if (!admit(base)) {
      charts.refuseCharts(drawings, 'This chart needs more worker memory than the document has free')
      return done()
    }
    // 2. Needed cells, grouped by page (summary search only).
    const key = (sheet: string, row: number, col: number) => `${sheet.toLowerCase()}\u0000${row}\u0000${col}`
    const slots = new Map<string, number>()
    const wantsText: boolean[] = []
    const groups = new Map<string, { sheetId: string; pageId: string; order: number; cells: Array<{ slot: number; row: number; col: number }> }>()
    for (const line of lines) {
      const sheet = sheetOf(line.sheet)
      if (!sheet) continue
      const lastRow = line.vertical ? line.start + line.count - 1 : line.fixed
      if (lastRow > sheet.importedThroughRow) pending.add(sheet.sheetId)
      const pages = this.workbook.store.pages(sheet.sheetId)
      let cursor = -1
      for (let i = 0; i < line.count; i++) {
        const at = charts.linePoint(line, i)
        if (at.row > sheet.importedThroughRow) break
        const id = key(sheet.name, at.row, at.col)
        const known = slots.get(id)
        if (known !== undefined) { wantsText[known] ||= line.text; continue }
        if (cursor < 0 || pages[cursor].r0 > at.row || pages[cursor].r1 < at.row) {
          let lo = 0, hi = pages.length
          while (lo < hi) { const mid = (lo + hi) >>> 1; if (pages[mid].r1 < at.row) lo = mid + 1; else hi = mid }
          cursor = lo < pages.length && pages[lo].r0 <= at.row ? lo : -1
        }
        if (cursor < 0) continue
        const summary = pages[cursor]
        if (summary.aggregate.records === 0 || at.col < summary.c0 || at.col > summary.c1) continue
        const slot = wantsText.length
        slots.set(id, slot)
        wantsText.push(line.text)
        const groupKey = `${sheet.sheetId}\u0000${summary.pageId}`
        let group = groups.get(groupKey)
        if (!group) groups.set(groupKey, group = { sheetId: sheet.sheetId, pageId: summary.pageId, order: cursor, cells: [] })
        group.cells.push({ slot, row: at.row, col: at.col })
      }
    }
    // 3./4. Stream the pages in order; strings admitted before they are built.
    const nums = new Float64Array(wantsText.length)
    // 0 unknown, 1 number, 2 text that is not a number, 3 present without a value,
    // 4 unavailable (a formula without a saved value), 5 an error cell (its code in `errors`; spec 5.8)
    const state = new Uint8Array(wantsText.length)
    const strings: Array<string | undefined> = new Array(wantsText.length)
    const errors: Array<string | undefined> = new Array(wantsText.length)
    const labels: Array<string | undefined> = new Array(wantsText.length)
    const formatCodes: Array<string | undefined> = new Array(wantsText.length)
    let reserved = base
    // Live charges of the records being read (a page, a string record) that the pool could not keep:
    // part of EVERY admission until the record is released (ownRecord / disown).
    // The unretained holders themselves, read at every admission: a lease's charge
    // can grow while it is held (its RAM owner released the shared buffer).
    const charges: { page?: { bytes: number }; string?: { bytes: number } } = {}
    const charged = () => (charges.page?.bytes ?? 0) + (charges.string?.bytes ?? 0)
    const reserve = (bytes: number): boolean => {
      if (reserved + bytes - base > charts.CHART_TEXT_BYTES_LIMIT || !admit(reserved + charged() + bytes)) return false
      reserved += bytes
      return true
    }
    const refuse = (release?: () => void) => { release?.(); admit(0); return done({ refused: true }) }
    const LIMIT = charts.CHART_LABEL_LIMIT
    // Formatting runs with an output cap (no builder grows past LIMIT + 1 units): the
    // capped output is admitted up front, and every further transient the formatter
    // reports through its allocation observer is admitted before it is allocated.
    const LABEL_OUTPUT_BYTES = (LIMIT + 1) * 2 * 2
    class LabelRefused extends Error {}
    const styles = this.workbook.formatting.styles
    const display: XlsxSheet = { name: '', rows: [], cols: [], merges: [], mergeRanges: [], semantics: semanticsFor(this.workbook.formatting.dateSystem) }
    /** One ownership path for pages and string records: pinned when the pool keeps
     * the record, otherwise charged to the read (refused when that charge cannot be
     * admitted). The charge stays in every admission until `disown`. */
    const ownRecord = <T extends { retained: boolean; bytes: number; release: () => void }>(held: T, slot: 'page' | 'string'): boolean => {
      if (held.retained) return true
      charges[slot] = held
      if (admit(reserved + charged())) return true
      held.release(); charges[slot] = undefined
      return false
    }
    const disown = (held: { release: () => void } | undefined, slot: 'page' | 'string') => { held?.release(); charges[slot] = undefined }
    // Number reading and label construction for one cell's text (record: typed storage).
    const takeText = (slot: number, kind: number, record: { text: Uint16Array; textStart: Uint32Array }, at: number, style: XlsxCellStyle | undefined): boolean => {
      const from = record.textStart[at], length = record.textStart[at + 1] - from
      if (kind === K_ERROR) {
        // Error identity (spec 5.8): the code is a short prefix, admitted before it is built.
        const units = Math.min(length, charts.CHART_ERROR_CODE_LIMIT)
        if (!reserve(units * 2)) return false
        errors[slot] = textPrefix(record, at, units).text
        state[slot] = 5
      } else {
        // The numeric reading streams the typed text (never built); labels are capped prefixes.
        const scanned = charts.scanChartNumber(length, k => record.text[from + k])
        if (scanned !== undefined) { nums[slot] = scanned; state[slot] = 1 } else state[slot] = 2
      }
      if (!wantsText[slot]) return true
      // Admitted, then built: at most LIMIT + 1 units (one past the limit marks truncation).
      if (!reserve(Math.min(length, LIMIT + 1) * 2)) return false
      const value = textPrefix(record, at, LIMIT + 1).text
      strings[slot] = value
      return label(slot, value, style)
    }
    const label = (slot: number, value: string | number | boolean | null, style: XlsxCellStyle | undefined): boolean => {
      if (!reserve(LABEL_OUTPUT_BYTES)) return false
      let live = 0, admitted = LABEL_OUTPUT_BYTES
      const allocation = {
        maxOutput: LIMIT + 1,
        allocate: (bytes: number) => {
          live += bytes
          if (live <= admitted) return
          // Transient formatter work is admitted against the read's memory (the hard cap bounds kept label storage).
          if (!admit(reserved - LABEL_OUTPUT_BYTES + live + charged())) throw new LabelRefused()
          admitted = live
        },
        release: (bytes: number) => { live -= bytes },
      }
      let text: string | undefined
      try { text = cellDisplayText(display, { value, style }, allocation) }
      catch (error) { if (error instanceof LabelRefused) return false; text = undefined }
      // Capped output keeps LIMIT + 1 units: one past the limit marks the label truncated (fillChartRanges cuts it).
      if (text !== undefined && text.length > LIMIT + 1) text = text.slice(0, LIMIT + 1)
      // True the reservation down to the label actually kept.
      reserved -= LABEL_OUTPUT_BYTES - (text?.length ?? 0) * 2
      labels[slot] = text
      return true
    }
    for (const group of [...groups.values()].sort((a, b) => a.sheetId === b.sheetId ? a.order - b.order : a.sheetId < b.sheetId ? -1 : 1)) {
      if (signal.aborted) return done()
      // Pass 1, page owned (its pool charge pinned; a page the pool could not keep
      // is charged to the read): numbers, inline text and their labels; shared
      // strings and oversized values are only noted, so no string record is
      // needed while the page is held.
      const deferred: Array<{ slot: number; kind: number; source: 'sst' | 'blob'; id: number; style: XlsxCellStyle | undefined }> = []
      const held = await this.workbook.store.pinnedPage(group.sheetId, group.pageId)
      if (!ownRecord(held, 'page')) return refuse()
      try {
        const p = held.page
        for (const cell of group.cells) {
          if (signal.aborted) return done()
          const index = this.find(p, cell.row, cell.col)
          if (index < 0) continue
          if (this.isUnavailable(p, index)) { state[cell.slot] = 4; continue }
          const kind = p.kinds[index], style = styles[p.styles[index]]
          formatCodes[cell.slot] = charts.sourceFormat(style)
          if (kind === K_SST) { deferred.push({ slot: cell.slot, kind, source: 'sst', id: p.nums[index], style }); continue }
          if (kind === K_TEXT || kind === K_ERROR) {
            if ((p.flags[index] & F_BLOB) !== 0) { deferred.push({ slot: cell.slot, kind, source: 'blob', id: p.nums[index], style }); continue }
            if (!takeText(cell.slot, kind, p, index, style)) return refuse()
            continue
          }
          const value = await this.valueOf(p, index)
          if (typeof value === 'number' && Number.isFinite(value)) { nums[cell.slot] = value; state[cell.slot] = 1 } else state[cell.slot] = 3
          if (wantsText[cell.slot] && !label(cell.slot, value, style)) return refuse()
        }
      } finally { disown(held, 'page') }
      // Pass 2, page released: string records one at a time. The needed ids are grouped
      // by record (SST chunk / oversized value) so each record is owned once, and the
      // previous record is released before the next is loaded: at most one string
      // record is owned at any moment.
      const byRecord = new Map<string, typeof deferred>()
      for (const item of deferred) {
        const recordKey = this.workbook.store.stringRecordKey(item.source, item.id)
        let items = byRecord.get(recordKey)
        if (!items) byRecord.set(recordKey, items = [])
        items.push(item)
      }
      for (const items of [...byRecord.entries()].sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0).map(([, list]) => list)) {
        if (signal.aborted) return done()
        const record = await this.workbook.store.pinnedString(items[0].source, items[0].id)
        if (!record) { for (const item of items) state[item.slot] = 3; continue }
        if (!ownRecord(record, 'string')) return refuse()
        try {
          for (const item of items) {
            if (signal.aborted) return done()
            // Entries of one record: an SST id's offset within the chunk follows from the first id's.
            const index = item.source === 'sst' ? record.index + (item.id - items[0].id) : record.index
            if (!takeText(item.slot, item.kind, record.record, index, item.style)) return refuse()
          }
        } finally { disown(record, 'string') }
      }
    }
    // The same point states as the resident fill (charts.cellPoint): unavailable and error keep their identity.
    charts.fillDrawings(drawings, owner, (sheet, row, col) => {
      const slot = slots.get(key(sheet, row, col))
      if (slot === undefined || state[slot] === 0) return undefined
      if (state[slot] === 4) return { value: undefined, unavailable: true }
      const number = state[slot] === 1 ? nums[slot] : null
      return { value: strings[slot] ?? (state[slot] === 1 ? nums[slot] : undefined), number, formatCode:formatCodes[slot], ...(state[slot] === 5 ? { error: errors[slot]! } : {}), ...(labels[slot] !== undefined ? { text: labels[slot] } : {}) }
    }, undefined, { hidden: (sheet, row, col) => hiddenCell(sheetOf(sheet)?.metadata, row, col), workspace, epoch: this.workbook.formatting.dateSystem })
    return done()
  }

  /** Keeps a complete payload in the pool (when one is bound and it fits) and
   * answers with a clone; otherwise the parsed payload itself is the answer.
   * `workingSet` is the read's charge WITHOUT the payload (D0): a retained payload
   * is owned by its pool entry (transferred, never charged twice), so the read's
   * share holds only the outgoing clone beside it: two live allocations. The
   * retained charts keep their document chart-facet charge until the entry ends
   * (eviction, close: `releaseCharts`); a payload that is not retained is the
   * reply itself and keeps its charge until that reply is released (collected;
   * charts/own.ts). Each model's charge is returned exactly once. */
  private retainDrawings(result: SheetDrawings, workingSet: number, memory?: DrawingMemory, releaseCharts: () => void = () => {}): SheetDrawings {
    const pool = this.drawingState.pool
    if (result.completeness !== 'complete' || !pool || pool.closed) return result
    const bytes = drawingFootprint(result)
    const holder: { entry?: PoolEntry } = {}
    holder.entry = pool.retain(`drawings:${this.stamp.docId}:${this.sheet.sheetId}`, {
      tier: 'rebuildable', bytes, owner: this.sheet.sheetId,
      onRelease: () => { releaseCharts(); if (this.drawingsCache?.entry === holder.entry) this.drawingsCache = undefined },
    }) ?? undefined
    if (!holder.entry) return result
    // The clone sits beside the retained copy until the reply is sent.
    // The clone cannot be admitted: the entry is dropped (an explicit release does not call onRelease) and the
    // parsed payload is the reply, keeping its charts' facet charge until it is released.
    if (memory && !memory.set(workingSet + bytes)) { holder.entry.release(); return result }
    this.drawingsCache = { entry: holder.entry, result, releaseCharts }
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
    const adapters = workbookAdapters.get(this.workbook)
    if (adapters?.get(this.sheet.sheetId) === this) adapters.delete(this.sheet.sheetId)
    this.drawingsCache?.entry.release()
    this.drawingsCache?.releaseCharts()
    this.drawingsCache = undefined
    if (this.sharedDefRules !== undefined) {
      releaseSharedCfDefinitions(this.sharedDefRules)
      this.sharedDefRules = undefined
    }
  }
}
/** Deterministic serialization of a rule set for shared-definition versioning:
 * identical rules share one charged copy; any change yields a new version. */
function normalizeCfRules(rules: readonly CfSheetRules['rules'][number][]): string {
  const normalized = [...rules].map(rule => ({
    id: rule.id,
    priority: rule.priority,
    documentOrder: rule.documentOrder,
    stopIfTrue: rule.stopIfTrue,
    ranges: [...rule.ranges].map(r => [r.firstRow, r.lastRow, r.firstCol, r.lastCol]),
    writes: [...rule.writes].sort(),
    dxf: Object.fromEntries(Object.keys(rule.dxf ?? {}).sort().map(k => [k, (rule.dxf as Record<string, unknown>)[k]])),
    definition: rule.definition,
  }))
  normalized.sort((a, b) => a.id - b.id || a.documentOrder - b.documentOrder)
  return JSON.stringify(normalized)
}

/** Content hash identifying a rule set for shared-definition versioning.
 * Test seam and provider helper: identical rule sets share one version. */
export function cfRulesVersion(rules: CfSheetRules | undefined): number {
  if (!rules || rules.rules.length === 0) return versionCfRules('[]')
  return versionCfRules(normalizeCfRules(rules.rules))
}

/** Exact UTF-16 units of one JSON string literal (quotes plus `"`/`\` escapes
 * and `\uXXXX` control escapes), matching JSON.stringify output. Reads only. */
function jsonQuotedUnits(value: string): number {
  let extra = 0
  for (let i = 0; i < value.length; i++) {
    const code = value.charCodeAt(i)
    if (code === 0x22 || code === 0x5c) extra += 1
    else if (code < 0x20) extra += 5
  }
  return 2 + value.length + extra
}
function jsonNumberUnits(value: number): number {
  if (!Number.isFinite(value)) return 4 // JSON.stringify emits null
  return String(value).length
}
/** Exact UTF-16 units of a JSON value, mirroring stringify semantics
 * (undefined/function/symbol properties dropped, array holes to null).
 * Returns undefined for unrepresentable shapes (depth/cycle/bigint), so the
 * caller refuses before serializing. Reads only; the `seen` set is the only
 * allocation and lives inside the caller's hold. */
function jsonValueUnits(value: unknown, depth: number, seen: Set<object>): number | undefined {
  if (value === null) return 4
  if (value === undefined) return undefined
  switch (typeof value) {
    case 'string': return jsonQuotedUnits(value)
    case 'number': return jsonNumberUnits(value)
    case 'boolean': return value ? 4 : 5
    case 'object': {
      if (depth > 64 || seen.has(value)) return undefined
      seen.add(value)
      try {
        if (Array.isArray(value)) {
          let units = 2
          for (let i = 0; i < value.length; i++) {
            if (i > 0) units += 1
            units += jsonValueUnits(value[i], depth + 1, seen) ?? 4
            if (!Number.isSafeInteger(units)) return undefined
          }
          return units
        }
        let units = 2, first = true
        for (const key of Object.keys(value as Record<string, unknown>)) {
          const item = jsonValueUnits((value as Record<string, unknown>)[key], depth + 1, seen)
          if (item === undefined) continue // dropped exactly like stringify
          if (!first) units += 1
          first = false
          units += jsonQuotedUnits(key) + 1 + item
          if (!Number.isSafeInteger(units)) return undefined
        }
        return units
      } finally {
        seen.delete(value)
      }
    }
    default: return undefined // function/symbol dropped; bigint refused above
  }
}
/** Skeleton punctuation of one normalized rule object, in normalizeCfRules
 * field order: the literal source of truth for the fixed per-rule units. */
const RULE_SKELETON = ['{"id":', ',"priority":', ',"documentOrder":', ',"stopIfTrue":', ',"ranges":', ',"writes":', ',"dxf":', ',"definition":', '}']
const RULE_SKELETON_UNITS = RULE_SKELETON.reduce((sum, part) => sum + part.length, 0)
/** Exact UTF-16 bytes of the normalized rule JSON (matching JSON.stringify
 * output byte-for-byte in size), computed by reading values but building
 * nothing: string lengths plus escape counts, numbers by literal length, the
 * definition by the bounded generic walk. Returns undefined for
 * unrepresentable shapes so the caller refuses before serializing. */
function cfNormalizedBytes(rules: readonly CfSheetRules['rules'][number][]): number | undefined {
  const seen = new Set<object>()
  let units = 2 // [...]
  for (let index = 0; index < rules.length; index++) {
    const rule = rules[index]!
    if (index > 0) units += 1
    units += RULE_SKELETON_UNITS
    units += jsonNumberUnits(rule.id) + jsonNumberUnits(rule.priority) + jsonNumberUnits(rule.documentOrder)
    units += rule.stopIfTrue ? 4 : 5
    units += 2 // ranges: [...]
    for (let ri = 0; ri < rule.ranges.length; ri++) {
      const range = rule.ranges[ri]!
      if (ri > 0) units += 1
      units += 5 + jsonNumberUnits(range.firstRow) + jsonNumberUnits(range.lastRow) + jsonNumberUnits(range.firstCol) + jsonNumberUnits(range.lastCol)
      if (!Number.isSafeInteger(units)) return undefined
    }
    units += 2 // writes: [...] (sorted order is the same multiset, same size)
    let wrote = false
    for (const name of rule.writes) {
      if (wrote) units += 1
      wrote = true
      units += jsonQuotedUnits(name)
      if (!Number.isSafeInteger(units)) return undefined
    }
    units += 2 // dxf: {...}
    const dxf = rule.dxf ?? {}
    let defined = false
    for (const key in dxf) {
      const item = jsonValueUnits((dxf as Record<string, unknown>)[key], 0, seen)
      if (item === undefined) {
        // stringify drops function/symbol/undefined values; bigint values
        // (nested or direct) refuse before serializing. A nested shape the
        // sizer misjudges still cannot slip through: normalization runs
        // inside try/catch with a byte-exact assertion below.
        const raw = (dxf as Record<string, unknown>)[key]
        if (raw !== undefined && typeof raw !== 'function' && typeof raw !== 'symbol' && typeof raw !== 'bigint') continue
        return undefined
      }
      if (defined) units += 1
      defined = true
      units += jsonQuotedUnits(key) + 1 + item
      if (!Number.isSafeInteger(units)) return undefined
    }
    const definition = jsonValueUnits(rule.definition, 0, seen)
    if (definition === undefined) return undefined
    units += definition
    units += 1 // closing }
    if (!Number.isSafeInteger(units)) return undefined
  }
  return units * 2
}

/** Read-only construction allowance for rule metadata: counts rules, ranges,
 * dxf keys and writes without reading values, copying or serializing, so no
 * rule content (including getter-backed test probes) is touched before the
 * returned bound is admitted. Exact per-value bytes join the working bound
 * later through the walked wire measures, inside the hold. */
function cfMetadataBound(rules: CfSheetRules | undefined, rectCount: number): number | undefined {
  let bound = 2048 + rectCount * 160
  if (!Number.isSafeInteger(bound)) return undefined
  if (rules) {
    bound += Math.min(rules.diagnostics.length, 64) * 2048
    for (const rule of rules.rules) {
      bound += 2048 + rule.ranges.length * 160 + rule.writes.size * 512
      const dxf = rule.dxf
      // Keys only: reading a value could evaluate arbitrary content, so each
      // entry carries a flat allowance instead of a measured size.
      if (dxf) for (const key in dxf) bound += 64 + key.length * 4 + 2048
      if (!Number.isSafeInteger(bound)) return undefined
    }
  }
  return bound
}

/** Cross-scope shared definitions: one frozen, charged array per live rule set
 * and version, so distinct viewport scopes with unchanged rules reuse the
 * same definitions instead of rebuilding (and recharging) them per scope.
 * The table entry owns its bytes through an explicit retained lease on the
 * shared account (admitted at commit); result leases charge the full reply
 * through post/drain, so the array is transiently double-counted while both
 * are live — conservative, never under-counted. Entries are committed only
 * on a fully admitted successful preparation — refusals and withholds never
 * populate the table — and the adapter releases an entry when its rules are
 * replaced or the store is released. WeakMap keys bound lifetime to the live
 * rules objects; explicit release ends the charge deterministically. */
const sharedCfDefinitions = new WeakMap<object, { version: number; definitions: readonly CfSharedDefinition[]; hits: number; lease?: CfLease<readonly CfSharedDefinition[]> }>()
function lookupSharedCfDefinitions(rules: CfSheetRules, version: number): readonly CfSharedDefinition[] | undefined {
  const cached = sharedCfDefinitions.get(rules)
  if (!cached || cached.version !== version) return undefined
  cached.hits++
  return cached.definitions
}
function commitSharedCfDefinitions(rules: CfSheetRules, version: number, definitions: readonly CfSharedDefinition[], lease: CfLease<readonly CfSharedDefinition[]> | undefined): void {
  sharedCfDefinitions.set(rules, { version, definitions, hits: 1, ...(lease === undefined ? {} : { lease }) })
}
/** Drops the shared-definition entry for replaced or released rules, ending
 * its charge; live result leases keep governing the result bytes until
 * their own release. */
export function releaseSharedCfDefinitions(rules: object): void {
  const cached = sharedCfDefinitions.get(rules)
  if (cached) {
    sharedCfDefinitions.delete(rules)
    if (cached.lease) {
      try { cached.lease.release() } catch { /* release-only path */ }
    }
  }
}
/** Test seam: identity-sharing lifecycle of one live rule set. */
export function sharedCfDefinitionStats(rules: object): { version: number; hits: number; charged: boolean } | undefined {
  const cached = sharedCfDefinitions.get(rules)
  return cached ? { version: cached.version, hits: cached.hits, charged: cached.lease !== undefined } : undefined
}
/** Deep-freezes a definition array: the outer array, each definition, its
 * properties and its writes. Shared arrays must be immutable at every
 * level, not just at the top. */
function freezeCfDefinitions(definitions: CfSharedDefinition[]): readonly CfSharedDefinition[] {
  for (const definition of definitions) {
    Object.freeze(definition.properties)
    Object.freeze(definition.writes)
    Object.freeze(definition)
  }
  return Object.freeze(definitions)
}

function cfReadError(sheetId: string, stamp: CfReadStamp, reason: string, feature: string, message: string, requiredBytes?: number, pendingSheetIds?: readonly string[]): CfReadResult {
  return {
    sheetId, stamp, kind: reason === 'cf-pending-data' ? 'pending' : 'unresolved', reason,
    states: [], definitions: [], definitionsVersion: versionCfRules('[]'),
    diagnostics: [{ kind: reason === 'cf-pending-data' ? 'pending' : 'resource-limit', feature, message }],
    ...(pendingSheetIds === undefined ? {} : { pendingSheetIds }),
    ...(requiredBytes === undefined ? {} : { requiredBytes }),
  }
}

/** Task 7: receiver-granted CF preparation over any CfValueSource. Validates the
 * grant and the requested stamp, withholds unimported scopes as pending, admits
 * the bounded estimate before touching any page, composes through the shared
 * CF service, and returns clone-safe winning states with versioned shared
 * definitions. Every early return releases what it admitted; the happy path
 * releases the service lease before returning plain data. */
export interface CfRemotePrepareOptions {
  readonly signal?: AbortSignal
  readonly budget?: () => CfWorkBudget
  readonly envelopeBytes?: number
  /** Receives the release of the live outgoing lease adopted on the shared
   * account: the caller holds it through post and releases after drain. */
  readonly retainOutgoing?: (release: () => void) => void
  /** Caller-admitted normalization: the adapter computed the version once
   * under its own hold (kept live across this call). The provider verifies
   * rules identity plus the content hash — both without reading rule values
   * — and skips its own sizing and serialization when they verify. */
  readonly admittedRules?: { rules: CfSheetRules | undefined; normalized: string; version: number }
}
export async function prepareCfRemote(
  request: CfReadRequest,
  sheetId: string,
  source: CfValueSource,
  rules: CfSheetRules | undefined,
  memory: CfMemory,
  options?: CfRemotePrepareOptions,
): Promise<CfReadResult> {
  const signal = options?.signal
  // The session forwards the kernel's measured envelope floor; the provider
  // uses the larger of that and its own framed wrapper bound.
  const envelopeBytes = Math.max(CF_REPLY_ENVELOPE_BYTES, options?.envelopeBytes ?? 0)
  const budget = options?.budget?.()
  const cancelled = () => { throw new PackageError('cancelled', 'CF remote preparation cancelled') }
  if (signal?.aborted) throw cancelled()
  const grantBytes = request.grantBytes
  if (!Number.isSafeInteger(grantBytes) || grantBytes <= 0) {
    return {
      ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-grant-required', 'Conditional formatting read requires a positive byte grant.'),
      requiredBytes: CF_REPLY_ENVELOPE_BYTES,
    }
  }
  for (const rect of request.rects) {
    if (!Number.isSafeInteger(rect.r0) || !Number.isSafeInteger(rect.r1) || !Number.isSafeInteger(rect.c0) || !Number.isSafeInteger(rect.c1)
      || rect.r0 < 0 || rect.r1 > 1048575 || rect.c0 < 0 || rect.c1 > 16383 || rect.r0 > rect.r1 || rect.c0 > rect.c1) {
      return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-invalid-rect', 'Conditional formatting read range is invalid.'), requiredBytes: undefined }
    }
  }
  if (request.stamp.sheetId !== sheetId) {
    return { ...cfReadError(sheetId, request.stamp, 'cf-stale-generation', 'cf-stale-source', 'Conditional formatting request targets a different sheet.'), requiredBytes: undefined }
  }
  // Bound validation-tier work upfront: the rule count caps every
  // normalization, definition build and per-cell bound below, so no
  // unbounded structure is ever built before admission. Both lazy CF modules
  // load before any reservation, so an import failure admits nothing.
  const limitsModule = await import('../cf/limits')
  const serviceModule = await import('../cf/service')
  if (rules && rules.rules.length > limitsModule.CF_LIMITS.rules) {
    return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-scope-too-large', 'Conditional formatting exceeds the bounded rule count.'), requiredBytes: undefined }
  }
  // The service refuses more than rangesPerRule source rects; refuse the same
  // bound here so an unbounded rect list can never reach the mapping below.
  if (request.rects.length > limitsModule.CF_LIMITS.rangesPerRule) {
    return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-scope-too-large', 'Conditional formatting scope exceeds the bounded preparation area.'), requiredBytes: undefined }
  }
  // Live stamp before any construction: a stale workbook generation withholds
  // without allocating.
  let live: CfStamp
  try {
    live = source.stamp(sheetId)
  } catch {
    return { ...cfReadError(sheetId, request.stamp, 'cf-stale-generation', 'cf-stale-source', 'Conditional formatting source stamp is unavailable.'), requiredBytes: undefined }
  }
  if (live.workbookValueGen !== request.stamp.workbookValueGen) {
    return {
      ...cfReadError(sheetId, { ...request.stamp, workbookValueGen: live.workbookValueGen }, 'cf-stale-generation', 'cf-stale-source', 'Conditional formatting source changed generation; re-request with the live stamp.'),
      requiredBytes: undefined,
    }
  }
  // Admit the metadata tier before touching rule content: the bound is a
  // pure read-only walk (counts only — no copies, no serialization, no value
  // reads), and the single reservation below covers every build until it
  // grows into the working hold. Concurrent preparations on a shared account
  // see each other's holds, so the cap is genuinely shared.
  const metaBound = cfMetadataBound(rules, request.rects.length)
  if (metaBound === undefined) {
    return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-scope-too-large', 'Conditional formatting scope exceeds the bounded preparation area.'), requiredBytes: undefined }
  }
  const working = memory.reserve(metaBound)
  if (working.kind !== 'reserved') {
    return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-budget-exceeded', 'Conditional formatting exceeds the admitted worker memory; retry when pressure drops.'), requiredBytes: undefined }
  }
  // A caller-admitted normalization arrives verified, not trusted: the rules
  // object must be identical and the content hash must match the admitted
  // string — both checks read no rule values. A mismatch falls back to the
  // provider's own admitted sizing and serialization below.
  const admittedInput = options?.admittedRules
  const admitted = admittedInput !== undefined && admittedInput.rules === rules
    && versionCfRules(admittedInput.normalized) === admittedInput.version
    ? admittedInput : undefined
  // Phase 2: exact serialization size by reading values (lengths and escape
  // counts only — no copies) inside the phase-1 hold, then grow the hold to
  // cover the build before serializing anything. A flat per-key allowance
  // cannot bound strings or JSON escaping; measuring first does. Skipped
  // when the caller admitted and verified the normalization above.
  let exactRules = 0
  if (!admitted && rules && rules.rules.length > 0) {
    try {
      const sized = cfNormalizedBytes(rules.rules)
      if (sized === undefined) {
        working.release()
        return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-scope-too-large', 'Conditional formatting definitions exceed the bounded preparation area.'), requiredBytes: undefined }
      }
      exactRules = sized
    } catch {
      working.release()
      return { ...cfReadError(sheetId, request.stamp, 'cf-preparation-failed', 'cf-preparation-withheld', 'Conditional formatting definitions could not be prepared.'), requiredBytes: undefined }
    }
    const grownTarget = metaBound + exactRules
    if (!Number.isSafeInteger(grownTarget) || !working.grow(grownTarget)) {
      working.release()
      return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-budget-exceeded', 'Conditional formatting exceeds the admitted worker memory; retry when pressure drops.'), requiredBytes: undefined }
    }
  }
  // Rules content hash is the worker's authority for definitions identity
  // (cfGen 0 means "unknown, adopt live" for cold starts). A known-but-wrong
  // version is withheld with the live stamp so the caller re-requests current
  // definitions instead of painting stale ones. Normalization runs at most
  // once per preparation — here, or once in the admitting caller — inside an
  // admitted hold, with a byte-exact backstop: a sizer/stringify mismatch
  // refuses here rather than building on a lie.
  let version: number
  if (admitted) {
    version = admitted.version
  } else if (rules && rules.rules.length > 0) {
    let normalized: string
    try {
      normalized = normalizeCfRules(rules.rules)
    } catch {
      working.release()
      return { ...cfReadError(sheetId, request.stamp, 'cf-preparation-failed', 'cf-preparation-withheld', 'Conditional formatting definitions could not be prepared.'), requiredBytes: undefined }
    }
    if (normalized.length * 2 > exactRules) {
      working.release()
      return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-budget-exceeded', 'Conditional formatting exceeds the admitted worker memory; retry when pressure drops.'), requiredBytes: undefined }
    }
    version = versionCfRules(normalized)
  } else {
    version = versionCfRules('[]')
  }
  if (rules && rules.rules.length > 0 && request.stamp.cfGen !== 0 && request.stamp.cfGen !== version) {
    working.release()
    return {
      ...cfReadError(sheetId, { ...request.stamp, workbookValueGen: live.workbookValueGen, cfGen: version }, 'cf-stale-generation', 'cf-stale-source', 'Conditional formatting definitions changed generation; re-request with the live stamp.'),
      requiredBytes: undefined,
    }
  }
  const liveStamp: CfReadStamp = { ...request.stamp, workbookValueGen: live.workbookValueGen, cfGen: version }
  // CF service ranges use firstRow/lastRow naming; the wire uses r0/r1/c0/c1.
  // Built inside the admitted hold: the rect list is request-bounded above.
  const cfRects = request.rects.map(r => ({ firstRow: r.r0, lastRow: r.r1, firstCol: r.c0, lastCol: r.c1 }))
  // Bounds before pages: unimported scopes stay pending without scanning anything.
  // Every early return and every throw below releases the admitted hold:
  // cancellation rethrows for the protocol, anything else fails closed.
  let bounds: { firstRow: number; lastRow: number; firstCol: number; lastCol: number } | undefined
  try {
    const b = await source.bounds(sheetId, budget, signal ?? new AbortController().signal)
    bounds = b === undefined ? undefined : { firstRow: b.firstRow, lastRow: b.lastRow, firstCol: b.firstCol, lastCol: b.lastCol }
  } catch (error) {
    try { working.release() } catch { /* release-only path */ }
    if (error instanceof PackageError && error.code === 'cancelled') throw error
    return cfReadError(sheetId, request.stamp, 'cf-pending-data', 'cf-source-unavailable', 'Conditional formatting bounds are not available.', undefined, [sheetId])
  }
  if (bounds === undefined) {
    working.release()
    return { ...cfReadError(sheetId, request.stamp, 'cf-pending-data', 'cf-source-pending', 'Conditional formatting source is still importing the requested range.'), pendingSheetIds: [sheetId] }
  }
  if (!rules || rules.rules.length === 0) {
    working.release()
    // Empty retained rules are not proof that CF is absent: capture/parse
    // overflow marks the sheet unknown via unresolvedSheet, and the common
    // service would withhold on it. Route the sentinel before the shortcut
    // so a fail-closed typed result — not a vacuous ready — comes back.
    const sentinel = rules?.unresolvedSheet
    if (sentinel) {
      const reason = sentinel.reason.startsWith('cf-') ? sentinel.reason : 'cf-preparation-failed'
      return {
        ...cfReadError(sheetId, liveStamp, reason, 'cf-sheet-scope-unresolved', `Conditional formatting sheet scope is unresolved (${sentinel.reason}); rules covering ${sentinel.writes} writes were withheld.`),
        requiredBytes: undefined,
      }
    }
    return {
      sheetId, stamp: liveStamp, kind: 'ready', states: [], definitions: [],
      definitionsVersion: version, diagnostics: [],
    }
  }
  // Upper bound before touching any page, measured with the same wire
  // arithmetic the reply is checked with: a fitting grant always admits an
  // exact build. The transient service workspace is bounded separately by
  // the session caps, so the grant only ever gates the wire reply.
  let cells = 0
  for (const r of cfRects) {
    cells += (r.lastRow - r.firstRow + 1) * (r.lastCol - r.firstCol + 1)
    if (!Number.isSafeInteger(cells) || cells > 1_000_000) {
      working.release()
      return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-scope-too-large', 'Conditional formatting scope exceeds the bounded preparation area.'), requiredBytes: undefined }
    }
  }
  // Shared definitions: one frozen array per live rule set and version, so
  // distinct viewport scopes with unchanged rules reuse the same definitions
  // instead of rebuilding (and recharging) them per scope. Previously
  // returned arrays stay intact for pinned readers after replacement. The
  // entry is committed only on a successful preparation below.
  let definitions: readonly CfSharedDefinition[] | undefined = lookupSharedCfDefinitions(rules, version)
  let builtFresh = false
  if (!definitions) {
    builtFresh = true
    const built: CfSharedDefinition[] = rules.rules.map((rule, index) => {
      const writes = [...rule.writes]
      const properties: Record<string, string | number | boolean | null> = {}
      for (const [key, value] of Object.entries(rule.dxf ?? {})) properties[key] = value as string | number | boolean | null
      return { id: `rule-${typeof rule.id === 'number' ? rule.id : index}`, version, properties, writes }
    })
    definitions = freezeCfDefinitions(built)
  }
  const diagnostics: CfReadResult['diagnostics'][number][] = rules.diagnostics.slice(0, 64).map(d => ({ kind: 'unsupported' as const, feature: d.feature, message: d.message.slice(0, 256) }))
  // Fixed framing is measured, not estimated: walk a template result with the
  // same walker the kernel publishes with, so preflight and post-build share
  // one framing model. Per-cell content stays bounded by the wire measure.
  // The template omits diagnostics on purpose: a grant that fits only the
  // base still completes ready with diagnostics omitted (never silently
  // partial data). Diagnostics ride the working hold through their own
  // walked size below instead of the grant gate. Every walk admits its
  // traversal workspace on the shared account instead of running unbounded.
  const templateBytes = limitsModule.measureCfWalkedBytes({
    sheetId, stamp: liveStamp, kind: 'ready', states: [], definitions,
    definitionsVersion: version, diagnostics: [],
  }, memory)
  const diagBytes = limitsModule.measureCfWalkedBytes(diagnostics, memory)
  // Per-cell bound matching the walked wire size exactly for full wins:
  // 24 object + 3x16 property slots + 2x8 numbers + 16 props slot + properties,
  // plus 8 for the states-array element slot the walker charges per item.
  // Partial wins undershoot; taint-heavy outliers are caught post-build.
  let perCell = 96
  for (const rule of rules.rules) perCell += measureCfPropertiesBytes((rule.dxf ?? {}) as Record<string, string | number | boolean | null>)
  if (templateBytes === undefined || diagBytes === undefined || !Number.isSafeInteger(perCell)) {
    working.release()
    return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-grant-required', 'Conditional formatting exceeds the receiver grant; retry with at least the estimated size.'), requiredBytes: undefined }
  }
  const estimate = estimateCfPreparationBytes(cells, perCell, 0, 0, 0)
  if (estimate === undefined) {
    working.release()
    return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-grant-required', 'Conditional formatting exceeds the receiver grant; retry with at least the estimated size.'), requiredBytes: undefined }
  }
  const boundedEstimate = templateBytes + estimate + envelopeBytes
  if (!Number.isSafeInteger(boundedEstimate) || boundedEstimate > grantBytes) {
    working.release()
    return {
      ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-grant-required', 'Conditional formatting exceeds the receiver grant; retry with at least the estimated size.'),
      requiredBytes: boundedEstimate,
    }
  }
  // Grow the metadata hold into the working hold: one reservation covers the
  // metadata build above plus the per-cell states, known diagnostics and
  // framing below, so the peak is admitted exactly once on the (possibly
  // shared) account. The grant gate stays diagnostics-free above.
  const workingTarget = boundedEstimate + diagBytes
  if (!Number.isSafeInteger(workingTarget) || !working.grow(workingTarget)) {
    working.release()
    return {
      ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-budget-exceeded', 'Conditional formatting exceeds the admitted worker memory; retry when pressure drops.'),
      requiredBytes: boundedEstimate,
    }
  }
  // Service and its lease live inside the working hold: everything built
  // below is covered by the pre-admitted bound, and a single finally unwinds
  // in reverse order (lease, service, working hold) exactly once.
  const service = serviceModule.createCfService(rules, memory, { budget: options?.budget })
  let prepared: CfLease<CfPrepared> | undefined
  try {
    if (signal?.aborted) throw cancelled()
    try {
      prepared = await service.prepareCf(
          // The service grant bounds the transient workspace (its own session
          // cap); the receiver grant above bounds the wire reply.
          { stamp: { ...liveStamp }, sourceRects: cfRects, signal: signal ?? new AbortController().signal, grantBytes: limitsModule.CF_LIMITS.scratchBytes },
          source,
          memory,
        )
    } catch (error) {
      if (error instanceof PackageError && error.code === 'cancelled') throw error
      if (signal?.aborted) throw cancelled()
        if (error instanceof CfPageFault) {
          if (error.fault === 'incomplete') {
            return { ...cfReadError(sheetId, liveStamp, 'cf-pending-data', 'cf-source-pending', 'Conditional formatting source is still importing the requested range.'), pendingSheetIds: [sheetId] }
          }
          if (error.fault === 'stale') {
            const live = source.stamp(sheetId)
            return {
              ...cfReadError(sheetId, { ...liveStamp, workbookValueGen: live.workbookValueGen }, 'cf-stale-generation', 'cf-stale-source', 'Conditional formatting source changed generation during preparation.'),
              requiredBytes: undefined,
            }
          }
          return {
            sheetId, stamp: liveStamp, kind: 'unresolved' as const, reason: 'cf-unavailable-values',
            states: [], definitions: [], definitionsVersion: version,
            diagnostics: [{ kind: 'unsupported' as const, feature: 'cf-unavailable-values', message: 'Conditional formatting withheld: the scope contains values unavailable to the cached reader.' }],
          }
        }
        const reason = error instanceof PackageError && error.code === 'resource-limit' ? 'cf-resource-limit' : 'cf-preparation-failed'
        return {
          sheetId, stamp: liveStamp, kind: 'unresolved' as const, reason,
          states: [], definitions: [], definitionsVersion: version,
          diagnostics: [{ kind: 'unsupported' as const, feature: 'cf-preparation-withheld', message: `Conditional formatting preparation withheld (${reason}).` }],
        }
    }
    // Revalidate liveness before publishing: a generation change during the
    // scan must withhold instead of publishing stale states.
    const liveAfter = source.stamp(sheetId)
    if (liveAfter.workbookValueGen !== request.stamp.workbookValueGen) {
      try { prepared.release() } catch { /* release-only path */ }
      prepared = undefined
      return {
        ...cfReadError(sheetId, { ...request.stamp, workbookValueGen: liveAfter.workbookValueGen }, 'cf-stale-generation', 'cf-stale-source', 'Conditional formatting source changed generation during preparation.'),
        requiredBytes: undefined,
      }
    }
    const states: CfReadResult['states'][number][] = []
    let unresolvedCells = 0
    let n = 0
    for (const rect of cfRects) {
      for (let row = rect.firstRow; row <= rect.lastRow; row++) {
        for (let col = rect.firstCol; col <= rect.lastCol; col++) {
          if (signal?.aborted) throw cancelled()
          if ((n++ & 63) === 0) await Promise.resolve()
          const at = prepared.value.at(row, col)
          if (at.kind === 'no-match') continue
          if (at.kind === 'unresolved') {
            unresolvedCells++
            states.push({ row, col, properties: {}, unresolvedWrites: [], unresolvedReason: at.reason })
            continue
          }
          const won = at.value.unresolvedWrites !== undefined ? [...at.value.unresolvedWrites] as string[] : undefined
          states.push({ row, col, properties: { ...at.value.properties }, ...(won === undefined ? {} : { unresolvedWrites: won }) })
        }
      }
    }
    if (unresolvedCells > 0) {
      // The post-build diagnostic was not in the preflight template: grow the
      // working hold to cover it before constructing, withholding on pressure.
      if (!working.grow(working.bytes + 1024)) {
        return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-budget-exceeded', 'Conditional formatting exceeds the admitted worker memory.'), requiredBytes: undefined }
      }
      diagnostics.push({ kind: 'unsupported', feature: 'cf-unresolved-cells', message: `${unresolvedCells} cells did not resolve and kept base formatting.` })
    }
    // Post-build measurement uses the same payload walker the kernel publishes
    // with, so preflight, publication and adoption share one framing model.
    // Every walk admits its traversal workspace on the shared account.
    const measureBuilt = (withDiagnostics: boolean): number | undefined => {
      const candidate: CfReadResult = {
        sheetId, stamp: liveStamp, kind: 'ready', states, definitions,
        definitionsVersion: version, diagnostics: withDiagnostics ? diagnostics : [],
      }
      return limitsModule.measureCfWalkedBytes(candidate, memory)
    }
    const fullBytes = measureBuilt(true)
    if (fullBytes === undefined) {
      return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-preparation-failed', 'Conditional formatting result could not be measured.'), requiredBytes: estimate }
    }
    const baseBytes = measureBuilt(false)
    if (baseBytes === undefined) {
      return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-preparation-failed', 'Conditional formatting result could not be measured.'), requiredBytes: estimate }
    }
    const fullWithEnvelope = fullBytes + envelopeBytes
    const baseWithEnvelope = baseBytes + envelopeBytes
    if (baseWithEnvelope > grantBytes) {
      return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-grant-required', 'Conditional formatting result exceeds the receiver grant.'), requiredBytes: fullWithEnvelope }
    }
    const omitted = fullWithEnvelope > grantBytes ? [] : diagnostics
    const result: CfReadResult = {
      sheetId, stamp: liveStamp, kind: 'ready', states, definitions,
      definitionsVersion: version, diagnostics: omitted,
    }
    // Transfer live outgoing ownership: adopt the wire result on the shared
    // account instead of refunding before return, so no window exists where
    // the live states sit uncharged. The caller holds the lease through post
    // and releases after drain; without an outbox the lease ends here.
    const wireBytes = limitsModule.measureCfWalkedBytes(result, memory)
    if (wireBytes === undefined || !Number.isSafeInteger(wireBytes + envelopeBytes)) {
      return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-budget-exceeded', 'Conditional formatting exceeds the admitted worker memory.'), requiredBytes: undefined }
    }
    let adopted: CfLease<CfReadResult>
    try {
      adopted = working.adopt(result, wireBytes + envelopeBytes)
    } catch {
      return { ...cfReadError(sheetId, request.stamp, 'cf-resource-limit', 'cf-budget-exceeded', 'Conditional formatting exceeds the admitted worker memory.'), requiredBytes: undefined }
    }
    // Publish only after all admission succeeds: the shared-definition entry
    // is committed here — with its own charged lease on the shared account —
    // never on a refusal or withhold path above. Lookups already counted
    // their hit; only fresh builds commit. If the definition charge cannot
    // be admitted the table stays empty (sharing skipped, result unaffected).
    if (builtFresh) {
      const defsBytes = limitsModule.measureCfWalkedBytes(definitions, memory)
      let defsLease: CfLease<readonly CfSharedDefinition[]> | undefined
      if (defsBytes !== undefined) {
        const defsHold = memory.reserve(defsBytes)
        if (defsHold.kind === 'reserved') {
          try {
            defsLease = defsHold.adopt(definitions, defsBytes)
          } catch {
            try { defsHold.release() } catch { /* release-only path */ }
          }
        }
      }
      commitSharedCfDefinitions(rules, version, definitions, defsLease)
    }
    // Adopting moves construction to retained; the finally below becomes a
    // no-op for the working hold while the outgoing lease stays live.
    if (options?.retainOutgoing) {
      const release = () => { try { adopted.release() } catch { /* release-only path */ } }
      try {
        options.retainOutgoing(release)
      } catch (error) {
        try { adopted.release() } catch { /* release-only path */ }
        throw error
      }
    } else {
      adopted.release()
    }
    return result
  } finally {
    if (prepared !== undefined) {
      try { prepared.release() } catch { /* release-only path */ }
    }
    service.dispose()
    try { working.release() } catch { /* release-only path */ }
  }
}

/** Minimal read surface the CF page bridge needs; the paged adapter supplies
 * it from its own read path, tests supply fakes. */
export interface CfPagedReadHost {
  readonly stamp: SourceStamp
  readRows(range: SourceRect, cursor: string | undefined, options: ReadOptions): Promise<Page<SparseRowPage>>
}

/** Typed bridge failures: incomplete pages withhold as pending (data may
 * arrive), unavailable cells withhold as unresolvable (unknown is not blank),
 * and stale pages withhold rather than publish a mixed-generation result. */
export class CfPageFault extends Error {
  constructor(
    public readonly fault: 'incomplete' | 'unavailable' | 'stale',
    public readonly sheetId: string,
    message: string,
  ) { super(message); this.name = 'CfPageFault' }
}

/** Admitted bytes of one bridge rect's chunk workspace: at most 64 chunk
 * entries in flight (object slots plus string references — the strings stay
 * owned by the store's page), admitted before the store is touched. */
const CF_PAGE_WORKSPACE_BYTES = 16384

/** CfValueSource pages over a readRows store: 64-cell pages admitted before
 * construction, released by the consumer after each page. Out-of-rect cells
 * never enter a page; incomplete scopes withhold, unavailable values withhold
 * (never silently blank), and stale pages withhold instead of publishing
 * mixed-generation results. */
export function pagedStoreCfPages(
  store: CfPagedReadHost,
  memory: CfMemory,
  signal?: AbortSignal,
): CfValueSource['pages'] {
  return async function* (id: string, rects: readonly import('../cf/model').SourceRect[], budget: CfWorkBudget, pageSignal: AbortSignal) {
    void id
    const combined = AbortSignal.any ? AbortSignal.any([signal, pageSignal].filter((s): s is AbortSignal => !!s)) : pageSignal
    // Generation pinned at open: every page must belong to it, or the whole
    // scan withholds rather than mixing generations.
    const openGen = store.stamp.valueGen
    const readRect = async function* (rect: (typeof rects)[number]): AsyncGenerator<CfLease<CfValuePage>> {
      let chunk: { row: number; col: number; value: CfLiteral }[] = []
      const flush = async function* (): AsyncGenerator<CfLease<CfValuePage>> {
        if (chunk.length === 0) return
        const spent = budget.spend(chunk.length)
        if (spent === 'refused') {
          if (combined?.aborted) throw new PackageError('cancelled', 'CF paged read cancelled')
          throw new PackageError('resource-limit', 'CF paged read exceeded the shared work budget', undefined, 'memory-budget')
        }
        if (spent === 'yield') {
          await hostTask()
          if (combined?.aborted) throw new PackageError('cancelled', 'CF paged read cancelled')
          budget.resume()
        }
        let bytes = 544 + chunk.length * 14
        for (const cell of chunk) {
          const v = cell.value
          if (v.kind === 'string') bytes += 16 + v.value.length * 2
          else if (v.kind === 'error') bytes += 16 + String(v.value).length * 2
          else if (v.kind === 'number') bytes += 92
          else if (v.kind === 'boolean') bytes += 94
          else bytes += 66
        }
        const hold = memory.reserve(bytes)
        if (hold.kind !== 'reserved') {
          throw new PackageError('resource-limit', 'CF paged read exceeded its admitted page', undefined, 'memory-budget')
        }
        try {
          const rows = new Uint32Array(chunk.length)
          const cols = new Uint16Array(chunk.length)
          const values = new Array<CfLiteral>(chunk.length)
          for (let i = 0; i < chunk.length; i++) {
            const cell = chunk[i]!
            rows[i] = cell.row
            cols[i] = cell.col
            values[i] = cell.value
          }
          yield hold.adopt({ rows, cols, values }, bytes)
        } catch (cause) {
          try { hold.release() } catch { /* release-only path */ }
          throw cause
        }
        chunk = []
      }
      let cursor: string | undefined
      for (;;) {
        if (combined?.aborted) throw new PackageError('cancelled', 'CF paged read cancelled')
        const page = await store.readRows(
          { r0: rect.firstRow, r1: rect.lastRow, c0: rect.firstCol, c1: rect.lastCol },
          cursor,
          { stamp: store.stamp, signal: combined ?? new AbortController().signal },
        )
        if (page.error) throw new PackageError('resource-limit', page.error.message || 'CF paged read failed on an incomplete page', undefined, 'memory-budget')
        if (page.completeness !== 'complete') throw new CfPageFault('incomplete', store.stamp.sheetId, 'Conditional formatting page is still importing; retry when complete')
        if (page.stamp.valueGen !== openGen) throw new CfPageFault('stale', store.stamp.sheetId, 'Conditional formatting source committed mid-scan; re-request at the live generation')
        for (const row of page.items) {
          for (const cell of row.cells) {
            if (cell.col < rect.firstCol || cell.col > rect.lastCol) continue
            // Unavailable is unknown, not blank: yielding it as blank would let
            // blank predicates match falsely and corrupt statistics. Withhold
            // the whole preparation instead.
            if (cell.unavailable) throw new CfPageFault('unavailable', store.stamp.sheetId, 'Conditional formatting scope contains values unavailable to the cached reader')
            const value: CfLiteral = cell.value === null ? { kind: 'blank' }
              : cell.valueIsError ? { kind: 'error', value: String(cell.value) }
              : typeof cell.value === 'number' ? { kind: 'number', value: cell.value }
              : typeof cell.value === 'boolean' ? { kind: 'boolean', value: cell.value }
              : { kind: 'string', value: cell.value }
            chunk.push({ row: row.row, col: cell.col, value })
            if (chunk.length >= 64) yield* flush()
          }
        }
        if (!page.nextCursor) break
        cursor = page.nextCursor
      }
      yield* flush()
    }
    for (const rect of rects) {
      // Admit the chunk workspace before the store is touched: chunk entries
      // and materialized values live inside this hold, released per rect.
      const workspace = memory.reserve(CF_PAGE_WORKSPACE_BYTES)
      if (workspace.kind !== 'reserved') {
        throw new PackageError('resource-limit', 'CF paged read exceeds its admitted page workspace', undefined, 'memory-budget')
      }
      try {
        yield* readRect(rect)
      } finally {
        try { workspace.release() } catch { /* release-only path */ }
      }
    }
  }
}

/** Per-document drawing state shared by a paged source's sheet adapters (S13.2). */
export interface PagedDrawingState {
  /** Transfer budget override (tests, embedders). */
  payloadBytes?: number
  /** The document pool, once the session binds it. */
  pool?: MemoryPool
  nodes: DrawingNodeLedger
  /** The document's chart facet (D0): one chart sublimit across every read and retained chart of the document. */
  charts?: { workspace?: unknown }
}
function memoryRefusal(): SheetDrawings {
  return unavailableDrawings('This sheet\'s drawings need more worker memory than the document has free and are not shown', failure('resource-limit', 'worker memory budget'), DRAWING_MEMORY_REFUSED)
}

function pageFailure(cause: unknown): SourceFailure {
  const code = (cause as { code?: unknown } | null)?.code
  // A record the source memory cannot admit right now (spec §12.4): typed and retryable.
  if (code === 'resource-limit') return failure('resource-limit', (cause as Error).message, true)
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
  /** Request-decode clock and host yield (tests inject both). */
  pacing?: StorePacing
}
/** Every chart model of a drawing tree. */
function forEachChart(drawings: readonly XlsxDrawing[], visit: (chart: ChartModel) => void): void {
  for (const node of drawings) {
    if (node.content?.kind === 'xlsxChart') visit(node.content.chart)
    if (node.children) forEachChart(node.children, visit)
  }
}
function collectPivotChartRefs(sheetId: string, drawings: readonly XlsxDrawing[], out: Array<{ sheetId: string; partPath: string; treePath: string; chart: ChartModel }>): void {
  for (const node of drawings) {
    // Only pivot charts (R2-4): the others are never linked, and would be held and charged with the assets.
    if (node.content?.kind === 'xlsxChart' && node.content.chart.pivot) out.push({ sheetId, partPath: node.source.partPath, treePath: node.source.treePath, chart: node.content.chart })
    if (node.children) collectPivotChartRefs(sheetId, node.children, out)
  }
}
/** Whether a cell sits in an authored hidden row or column (chart `plotVisOnly` provenance, as buildGeometry reads them). */
function hiddenCell(meta: SheetMetadata | undefined, row: number, col: number): boolean {
  if (!meta) return false
  const rows = meta.hiddenRows
  if (meta.unorderedRows) return rows.some(run => run.r0 <= row && row <= run.r1) || meta.cols.some(spec => spec.hidden && spec.min <= col && col <= spec.max)
  let lo = 0, hi = rows.length
  while (lo < hi) { const mid = (lo + hi) >>> 1; if (rows[mid].r1 < row) lo = mid + 1; else hi = mid }
  if (lo < rows.length && rows[lo].r0 <= row) return true
  return meta.cols.some(spec => spec.hidden && spec.min <= col && col <= spec.max)
}
export function createPagedWorkbookSource(session: PagedImportSession, options: PagedSourceOptions = {}): WorkbookSource {
  const { workbook } = session
  const displaySheet: XlsxSheet = { name: '', rows: [], cols: [], merges: [], mergeRanges: [], semantics: semanticsFor(workbook.formatting.dateSystem) }
  const stores = new Map<string, PagedXlsxSheetAdapter>()
  const geometries = new Map<string, SheetGeometry>()
  const pivotAssetCache = new Map<string, { assets: WorkerPivotAssets; entry?: PoolEntry; refs: number; retired: boolean }>()
  const currentPivotAssets = new Map<string, string>()
  // One build per sheet and generation (R2-4): a concurrent miss waits for it instead of retaining the same pool key.
  const pivotAssetBuilds = new Map<string, Promise<WorkerPivotAssets | undefined>>()
  const drawingState: PagedDrawingState = { nodes: new DrawingNodeLedger(), ...(options.drawingPayloadBytes !== undefined ? { payloadBytes: options.drawingPayloadBytes } : {}) }
  const find = (sheetId: string) => {
    const found = workbook.sheets.find(s => s.sheetId === sheetId)
    if (!found) throw new RangeError(`Unknown sheet ${sheetId}`)
    return found
  }
  const storeFor = (sheetId: string): PagedXlsxSheetAdapter => {
    let store = stores.get(sheetId)
    if (!store) {
      store = new PagedXlsxSheetAdapter(Object.freeze({ docId: workbook.docId, sheetId, valueGen: INITIAL_GENERATION }), workbook, find(sheetId), drawingState, options.pacing)
      stores.set(sheetId, store)
    }
    return store
  }
  /** Builds a sheet's pivot assets for its current generation (`pivotAssets` runs one build per key). */
  const buildPivotAssets = async (sheetId: string, signal: AbortSignal, profile: ProfileId, memory: Parameters<NonNullable<WorkbookSource['pivotAssets']>>[3]): Promise<WorkerPivotAssets | undefined> => {
      const sheet = find(sheetId)
      if (signal.aborted || sheet.state !== 'complete' || !sheet.metadata.pivotRegions?.length) return undefined
      const store = storeFor(sheetId)
      const generation = store.stamp.valueGen
      const cacheKey = JSON.stringify([sheetId, generation])
      const cached = pivotAssetCache.get(cacheKey)
      if (cached) {
        if (memory && !memory.set(0)) return undefined
        return cached.assets
      }
      const { readPagedPivotDisplayParts } = await import('../storage/pagedDrawings')
      const parsed = await readPagedPivotDisplayParts(workbook.parts, { sheetId, sheetName: sheet.name,
        pivotRegions: new Map([[sheet.name, sheet.metadata.pivotRegions]]) }, profile, signal, memory, workbook.pivotPartBytes)
      if (memory && !memory.set(0)) return undefined
      if (!parsed.pivotTables.length || !parsed.pivotCaches.size) return undefined
      const { createPivotRuntime } = await import('../pivots/session/runtime')
      const { openPart } = await import('../../core/zip-stream')
      const runtime = createPivotRuntime(parsed.pivotCaches, (path, readSignal) => {
        const entry = workbook.parts.inventory.find(path)
        return entry ? openPart(workbook.parts.src, entry, { signal: readSignal, budget: { charge: () => {} } }) : undefined
      }, () => workbook.formatting.dateSystem === '1904')
      const used = sheet.used
      if (!used || !store.scanOccupied) return undefined
      const occupied = new Set<number>()
      let cursor: string | undefined, count = 0
      const readOptions = { signal, stamp: store.stamp }
      do {
        const pages = store.scanOccupied({ range: used, limit: 8192, ...(cursor ? { cursor } : {}) }, readOptions)
        let next: string | undefined
        for await (const page of pages) {
          if (page.completeness !== 'complete' || page.error) return undefined
          for (const cell of page.items) {
            if (++count > 500_000) return undefined
            if (memory && !memory.set(count * 32 + 512)) return undefined
            occupied.add(cell.row * MAX_COLS + cell.col)
          }
          next = page.nextCursor
        }
        cursor = next
      } while (cursor !== undefined)
      const linkedCharts: Array<{ sheetId: string; partPath: string; treePath: string; chart: ChartModel }> = []
      for (const chartSheet of workbook.sheets) {
        if (chartSheet.state !== 'complete') return undefined
        if (!chartSheet.metadata.drawingIds.length && !chartSheet.metadata.legacyDrawingIds?.length) continue
        const chartStore = storeFor(chartSheet.sheetId)
        if (!chartStore.readDrawings) return undefined
        const drawings = await chartStore.readDrawings({ signal, stamp: chartStore.stamp }, profile, memory)
        if (drawings.completeness !== 'complete' || signal.aborted) return undefined
        collectPivotChartRefs(chartSheet.sheetId, drawings.drawings, linkedCharts)
        if (memory && !memory.set(0)) return undefined
      }
      if (store.stamp.valueGen !== generation || signal.aborted) return undefined
      const { pagedCellReader, pagedRangeReader } = await import('../pivots/engine/faithful')
      const { RECORD_LIMITS } = await import('../pivots/engine/records')
      // Merges, array and spill ranges and noted cells count too (S3-AC), charged with the metadata.
      const { pagedBlockers, structuralOccupancy } = await import('../pivots/session/occupancy')
      const structure = structuralOccupancy(pagedBlockers(sheet))
      const metadataBytes = 512 + count * 32 + structure.bytes + estimatePayloadBytes({ tables: parsed.pivotTables, caches: parsed.pivotCaches, linkedCharts })
      const previousKey = currentPivotAssets.get(sheetId)
      if (previousKey && previousKey !== cacheKey) {
        const previous = pivotAssetCache.get(previousKey)
        if (previous) {
          previous.retired = true
          if (previous.refs === 0) { previous.entry?.release(); pivotAssetCache.delete(previousKey) }
        }
      }
      currentPivotAssets.set(sheetId, cacheKey)
      const pool = drawingState.pool
      let chargedMetadataBytes = metadataBytes
      let entry: PoolEntry | undefined
      if (pool) {
        entry = pool.retain(`pivot-assets:${workbook.docId}:${sheetId}:${generation}`, {
          tier: 'source', owner: sheetId, bytes: metadataBytes, evictable: false,
          onRelease: () => {
            pivotAssetCache.delete(cacheKey)
            if (currentPivotAssets.get(sheetId) === cacheKey) currentPivotAssets.delete(sheetId)
          },
        }) ?? undefined
        if (!entry) throw Object.assign(new Error('Pivot metadata exceeds the worker memory budget.'), {
          code: 'resource-limit', requiredBytes: metadataBytes, limitBytes: pool.limitBytes,
        })
        chargedMetadataBytes = 0
      }
      const assets: WorkerPivotAssets = {
        tables: parsed.pivotTables, date1904: runtime.date1904, retainedBytes: 0, metadataBytes: chargedMetadataBytes,
        linkedCharts,
        pivots: workbook.sheets.flatMap(candidate => (candidate.metadata.pivotRegions ?? []).map(region => ({ sheet: candidate.name, ...(region.name !== undefined ? { name: region.name } : {}) }))),
        runtimeKey: runtime,
        retain: () => {
          const cachedAssets = pivotAssetCache.get(cacheKey)
          if (!cachedAssets) return () => {}
          cachedAssets.refs++
          let released = false
          return () => {
            if (released) return
            released = true
            cachedAssets.refs--
            // The last holder frees it (R2-2): an idle sheet keeps no pivot metadata charged; the next open rebuilds.
            if (cachedAssets.refs === 0) {
              cachedAssets.entry?.release()
              pivotAssetCache.delete(cacheKey)
            }
          }
        },
        cache: def => runtime.cache(def),
        records: (cache, readSignal, limits = RECORD_LIMITS[profile]) => runtime.records(cache, readSignal, limits),
        saved: (def, readSignal) => pagedCellReader(store, def, { signal: readSignal, stamp: store.stamp }, profile),
        // Task 32: a refreshOnLoad cache's source range, read page by page through that sheet's store (capped as the
        // saved cells are; an incomplete read marks the reader, and the refresh refuses).
        source: async (cache, readSignal) => {
          const { worksheetSourceRange } = await import('../pivots/session/refresh')
          const range = worksheetSourceRange(cache)
          const target = range ? workbook.sheets.find(candidate => candidate.name === range.sheet) : undefined
          if (!range || !target) return undefined
          const sourceStore = storeFor(target.sheetId)
          return pagedRangeReader(sourceStore, range.ref, { signal: readSignal, stamp: sourceStore.stamp }, profile)
        },
        occupied: (row, col) => occupied.has(row * MAX_COLS + col),
        blocked: (row, col) => structure.occupied(row, col),
        store: storeFor,
      }
      pivotAssetCache.set(cacheKey, { assets, ...(entry ? { entry } : {}), refs: 0, retired: false })
      return assets
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
    hasSparklines(sheetId) { return Boolean(find(sheetId).metadata.sparklineGroups?.length) },
    hasRegionStyles(sheetId) { return Boolean(find(sheetId).tables?.some(table => table.styleInfo?.name)) },
    regionStyleTableCount(sheetId) { return find(sheetId).tables?.filter(table => table.styleInfo?.name).length ?? 0 },
    hasPivotDisplay(sheetId) { const sheet = find(sheetId); return Boolean(sheet.metadata.pivotRegions?.length || sheet.metadata.pivotDiagnostics?.length) },
    pivotTableCount(sheetId) { return find(sheetId).metadata.pivotRegions?.length ?? 0 },
    async scanFormulaCellsOutsidePivots(signal?: AbortSignal) {
      for (const candidate of workbook.sheets) {
        if (signal?.aborted) return undefined
        if (candidate.state !== 'complete') return true // incomplete scan is conservative
        const pivots = (candidate.metadata.pivotRegions ?? []).flatMap(region => { const rect = parseA1Range(region.ref); return rect ? [rect] : [] })
        for (const summary of workbook.store.pages(candidate.sheetId)) {
          if (signal?.aborted) return undefined
          const lease = await workbook.store.pinnedPage(candidate.sheetId, summary.pageId)
          try {
            if (signal?.aborted) return undefined
            for (const [record, meta] of lease.page.meta) {
              if (signal?.aborted) return undefined
              if (!meta.formula && !meta.sharedFormula) continue
              const row = lease.page.rows[record]!, col = lease.page.cols[record]!
              if (!pivots.some(r => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1)) return true
            }
          } finally { lease.release() }
        }
      }
      return false
    },
    async pivotAssets(sheetId, signal, profile, memory): Promise<WorkerPivotAssets | undefined> {
      const sheet = find(sheetId)
      if (signal.aborted || sheet.state !== 'complete' || !sheet.metadata.pivotRegions?.length) return undefined
      const key = JSON.stringify([sheetId, storeFor(sheetId).stamp.valueGen])
      const building = pivotAssetBuilds.get(key)
      if (building) {
        const shared = await building
        // A build its own caller abandoned (aborted) is retried by this one.
        if (shared || signal.aborted) return shared
      }
      const build = buildPivotAssets(sheetId, signal, profile, memory)
      pivotAssetBuilds.set(key, build)
      try { return await build } finally { if (pivotAssetBuilds.get(key) === build) pivotAssetBuilds.delete(key) }
    },
    store: storeFor,
    geometry(sheetId) {
      const sheet = find(sheetId)
      let geometry = geometries.get(sheetId)
      if (!geometry) {
        // Row metadata arrives with the rows; cache only once the sheet is complete.
        geometry = withFrozen(buildGeometry(sheet.metadata.rowHeights.map(h => ({ index: h.row, heightPt: h.heightPt })), sheet.metadata.cols, sheet.metadata.hiddenRows, { hiddenCols: true, columns: pagedColumns(workbook, sheet.metadata) }), sheet.metadata.frozen)
        if (sheet.state === 'complete') geometries.set(sheetId, geometry)
      }
      return geometry
    },
    coverageDiagnostics: sheetId => storeFor(sheetId).coverageDiagnostics(),
    formatting: () => workbook.formatting,
    // A styled table/pivot region resolves over the cell's own style, as the viewer paints it (worker Find).
    displayText: (cell, _sheetId, allocation, region) => {
      const own = workbook.formatting.styles[cell.styleIndex]
      const overlayStyle = storeFor(_sheetId).overlayStyle({ row: cell.row, col: cell.col })
      const style = overlayStyle === null ? undefined : overlayStyle ?? own
      return cellDisplayText(displaySheet, { value: cell.value, style: region?.styleAt(cell.row, cell.col, style, 'screen') ?? style }, allocation)
    },
    declaredFilters: () => workbook.sheets.flatMap(s => s.declared ? [s.declared] : []),
    bindMemory(pool: MemoryPool) { workbook.store.bindMemory(pool); drawingState.pool = pool },
    async release() {
      for (const cached of pivotAssetCache.values()) cached.entry?.release()
      pivotAssetCache.clear()
      currentPivotAssets.clear()
      await Promise.all([...stores.values()].map(store => store.release()))
      stores.clear()
      geometries.clear()
      await session.release()
    },
  }
}

/** Excel column pixel model of a paged sheet: workbook MDW and its sheetFormatPr. */
function pagedColumns(workbook: PagedWorkbook, meta: { format?: { baseColWidth?: number; defaultColWidth?: number } }) {
  const mdw = workbook.formatting.maxDigitWidthPx ?? 7
  return { maxDigitWidthPx: mdw, defaultWidthPx: defaultColumnWidthPx(meta.format ?? {}, mdw) }
}
