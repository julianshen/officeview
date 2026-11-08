import { cfExplicitBounds } from '../cfLoader'
import { RectCover, uniqueRectList } from './rectIndex'
import { SOURCE_ALLOCATION, adoptAllocation, type AllocationReservation, type SourceAllocationOptions } from './allocationAccount'
import type { FormatAllocation } from '../formula/format'
/**
 * Source adapters behind the async SourceStore (spec §§4.1-4.2, plan S1.2).
 *
 * `ObjectXlsxSheetAdapter` wraps the current parser output BY REFERENCE (no
 * clone, no mutation) for admitted small files: it indexes only populated
 * rows, keeps no per-column caches, and bounds every page by rows AND cells.
 * Structure comes from parser metadata (S3.3): merges, tables, saved array
 * extents, engine spill footprints and drawing anchors. Coverage it cannot
 * resolve (unreadable drawing anchors, a dynamic array with no known extent,
 * an unavailable table inventory) is reported incomplete with a diagnostic.
 * The paged adapter for large files lives in ./pagedSource and answers the same
 * questions from stored pages; adapter parity is asserted within one value policy.
 */
import { DEFAULT_COLUMN_METRICS, columnWidthPx, type ColumnMetrics } from '../columnWidth'
import type {
  AuthoritativeCell, AxisSizeInterval, CellRead, ColumnEntry, ColumnQuery, OverviewGroup, OverviewQuery, Page, PopulatedQuery,
  DrawingAnchor, RangeCompletion, ReadOptions, SingleCellPayload, SourceFailure, SourceRowInterval, SourceStamp, SourceStore,
  SparseRowPage, SpillExtent, StructuralResult, ValuePolicy, ValueOverlay, ValueOverlayCell, WorkerDiagnostic,
} from './contracts'
import type { ColumnLanes, DrawingMemory, ProfileId, SheetDrawings } from './contracts'
import { DRAWING_MEMORY_REFUSED, drawingFootprint, drawingPayloadBudget, sheetDrawingPayload, unavailableDrawings } from './drawingPayload'
import { countLoadEvaluatedFormulas, documentValuePolicy, isValueUnavailable, laneValue, SAVED_VALUES_DIAGNOSTIC, valueAuthority, valueStatus } from './values'
import type { XlsxCell, XlsxDocument, XlsxDrawing, XlsxRow, XlsxSheet } from '../types'
import type { TableMetadata } from '../formula/types'
import { savedPivotRects } from '../pivots/stub'
import { engineSpillFootprints, type EngineSpillFootprint } from '../formula/workbook'
import { resolveDrawingAnchors } from './drawingBounds'
import { MAX_COLS, MAX_ROWS, cellId, formatA1, isValidCoordinate, isValidSourceRect, parseA1Range, type SourceCell, type SourceRect } from './coordinates'
import { createAxisMetrics } from './axes'
import { createIntervalSet } from './visibility'
import { INITIAL_GENERATION } from '../../worker/viewerProtocol'
import {
  failedPage, failure, lowerBound, rowSetIntervals, sameStamp, type SheetGeometry, type WorkbookSource, type WorkerPivotAssets,
} from '../../worker/viewerSession'
import { addRecord, answerOverview, emptyAggregate, type OverviewCellSource } from '../storage/overviewIndex'
import { cellDisplayText } from '../render'
import { CheckpointBudget, type CooperativeOptions } from './scheduler'
import { estimatePayloadBytes } from './payloadBytes'

/** Row-page bounds: at most 200 rows and 262,144 cells (matches the fixture store). */
export const READ_ROWS_PAGE_ROWS = 200
export const READ_ROWS_MAX_CELLS = 262_144
export interface ObjectSourceOptions {
  readRowsMaxCells?: number
  /** Drawing transfer budget override (S13.2); default drawingPayloadBudget(profile). */
  drawingPayloadBytes?: number
  /** Fixed at admission (§4.4); 'cached' wraps a doc parsed with valuePolicy 'cached'. */
  valuePolicy?: ValuePolicy
  /** Request-decode clock and host yield (tests inject both). */
  pacing?: StorePacing
}
export type StorePacing = Pick<CooperativeOptions, 'now' | 'yield'>
function collectObjectPivotCharts(sheetId: string, drawings: readonly XlsxDrawing[] | undefined, out: Array<{ sheetId: string; partPath: string; treePath: string; chart: import('../../charts/model').ChartModel }>): void {
  for (const node of drawings ?? []) {
    // Only pivot charts (R2-4): the others are never linked, and would be held and charged with the assets.
    if (node.content?.kind === 'xlsxChart' && node.content.chart.pivot) out.push({ sheetId, partPath: node.source.partPath, treePath: node.source.treePath, chart: node.content.chart })
    if (node.children) collectObjectPivotCharts(sheetId, node.children, out)
  }
}
/** A request's decode stopped at a checkpoint (cancelled or a stale owner). */
export class SourceStop extends Error {
  constructor(readonly failure: SourceFailure) { super(failure.message) }
}
/** Decode budget of one store request (spec §10.2, decision 10): cells are
 * decoded in steps of at most 4,096 with a real yield between steps, and the
 * request's guard (signal, owner) runs at every checkpoint, throwing SourceStop. */
export function requestBudget(guard: () => SourceFailure | undefined, pacing: StorePacing = {}): CheckpointBudget {
  return new CheckpointBudget(() => { const stop = guard(); if (stop) throw new SourceStop(stop) }, undefined, pacing, 'source')
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
  private capacity = 0
  private length = 0
  private stagedBytes = 0
  private finished = false
  constructor(private readonly account?: AllocationReservation) {}
  get count(): number { return this.length }
  private grow(): void {
    const capacity = Math.max(16, this.capacity * 2), bytes = capacity * 40 + 80
    this.account?.charge(bytes)
    try {
      const rows = new Array<number>(capacity), tags = new Array<number>(capacity), numbers = new Array<number>(capacity)
      const styles = new Array<number>(capacity), texts = new Array<string>(capacity)
      for (let i = 0; i < this.length; i++) {
        rows[i] = this.rows[i]; tags[i] = this.tags[i]; numbers[i] = this.numbers[i]
        styles[i] = this.styles[i]; texts[i] = this.texts[i]
      }
      this.rows = rows; this.tags = tags; this.numbers = numbers; this.styles = styles; this.texts = texts
      this.account?.release(this.stagedBytes)
      this.capacity = capacity
      this.stagedBytes = bytes
    } catch (error) { this.account?.release(bytes); throw error }
  }
  /** Same classification as the value index's cell-entry path (valueIndex.place). */
  add(row: number, value: string | number | boolean | null | undefined, valueIsError: boolean | undefined, styleIndex: number, unavailable: boolean): void {
    if (this.finished) throw new RangeError('Local column producer already finished')
    const { tag, num, text } = laneValue(value, valueIsError, unavailable)
    if (this.length === this.capacity) this.grow()
    const i = this.length++
    this.rows[i] = row; this.tags[i] = tag; this.numbers[i] = num; this.styles[i] = styleIndex; this.texts[i] = text
    this.textUnits += text.length
  }
  build(r0: number, r1: number): ColumnLanes {
    if (this.finished) throw new RangeError('Local column producer already finished')
    const n = this.length
    this.account?.charge(21 * n + 4 + 2 * this.textUnits + 512)
    const rows = new Uint32Array(n), tags = new Uint8Array(n), numbers = new Float64Array(n), styles = new Uint32Array(n)
    const text = new Uint16Array(this.textUnits), textStart = new Uint32Array(n + 1)
    let at = 0
    for (let i = 0; i < n; i++) {
      rows[i] = this.rows[i]; tags[i] = this.tags[i]; numbers[i] = this.numbers[i]; styles[i] = this.styles[i]
      textStart[i] = at
      const t = this.texts[i]
      for (let k = 0; k < t.length; k++) text[at++] = t.charCodeAt(k)
    }
    textStart[n] = at
    const result: ColumnLanes = { kind: 'lanes', r0, r1, rows, tags, numbers, styles, text, textStart }
    if (this.account) {
      this.rows = []; this.tags = []; this.numbers = []; this.styles = []; this.texts = []
      this.account.release(this.stagedBytes)
      this.stagedBytes = 0
      this.finished = true
    }
    return result
  }
}
/** Bytes a lanes run holds (row 4 + tag 1 + number 8 + style 4 per entry, 2 per code unit). */
export const laneBytes = (lanes: ColumnLanes): number => lanes.rows.length * 17 + lanes.text.length * 2 + lanes.textStart.length * 4

/** An A1 range (or cell) ref as a rectangle, null when unreadable (the shared parser, `parseA1Range`). */
const parseRange = (ref: string): SourceRect | null => parseA1Range(ref) ?? null
const intersects = (a: SourceRect, b: SourceRect): boolean =>
  a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1
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
// Keep sheets that never publish a pivot overlay free of per-store overlay
// fields; the common path performs only this WeakMap lookup.
const overlayStates = new WeakMap<object, OverlayState>()

export class ObjectXlsxSheetAdapter implements SourceStore {
  private baseExtent!: SourceStore['extent']
  private currentStamp: SourceStamp
  /** The sheet's row inventory, taken from its rows (`reindex`): again when a dependent spill adds or removes rows. */
  private rows!: readonly XlsxRow[]
  private rowIndexes!: readonly number[]
  private dynamicAnchors!: readonly { row: number; col: number }[]
  private readonly savedPivots: readonly SourceRect[]
  private drawingCache?: DrawingAnchor[]
  private released = false
  /** `tables` undefined means the workbook table inventory is unavailable (the
   * parser's fail-closed signal), never "no tables". `engineSpills` are the
   * engine's published footprints for this sheet; undefined under the `engine`
   * policy means no evaluation state is known, so dynamic extents stay unknown. */
  constructor(
    stamp: SourceStamp,
    private readonly sheet: XlsxSheet,
    private readonly tables: readonly TableMetadata[] | undefined,
    private readonly maxCells: number,
    readonly valuePolicy: ValuePolicy = 'engine',
    private readonly engineSpills?: readonly EngineSpillFootprint[],
    private readonly drawingPayloadBytes?: number,
    private readonly pacing: StorePacing = {},
  ) {
    this.currentStamp = stamp
    this.reindex()
    // The saved pivots' guarded rectangles, taken once from the file's definitions (T26 m2, Task 33 perf): a commit
    // replaces the live definitions with moved locations, and a structural query reparses nothing.
    const groups = [
      ...(sheet.pivotRegions ?? []).map(region => { const table = parseRange(region.ref); return table ? savedPivotRects(table, region.rowPageCount, region.colPageCount) : [] }),
      ...(sheet.pivotTables ?? []).map(pivot => {
        const table = parseRange(pivot.location.ref), pages = pivot.pageFields?.length ?? 0
        return table ? savedPivotRects(table, pages ? pivot.location.rowPageCount || pages : 0, pivot.location.colPageCount) : []
      }),
    ]
    this.savedPivots = Object.freeze(uniqueRectList(groups.flat()).map(rect => Object.freeze(rect)))
  }
  /** The row inventory, its index and the used extent, from the sheet's rows as they are now. */
  private reindex(): void {
    const sheet = this.sheet
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
    // A resolved sparkline is visible worksheet content even when its host has
    // no stored cell record; include it in range-aware navigation and queries.
    for (const [row, hosts] of sheet.resolvedSparklines ?? []) {
      for (const col of hosts.keys()) {
        if (!isValidCoordinate(row, col)) continue
        r0 = Math.min(r0, row); r1 = Math.max(r1, row)
        c0 = Math.min(c0, col); c1 = Math.max(c1, col)
      }
    }
    // A merged sparkline host occupies its full merged cell. Keep the used-range
    // extension narrow: unrelated merge ranges do not change navigation bounds.
    for (const merge of sheet.mergeRanges ?? []) {
      if (!isValidCoordinate(merge.minRow, merge.minCol) || !isValidCoordinate(merge.maxRow, merge.maxCol)
        || merge.minRow > merge.maxRow || merge.minCol > merge.maxCol
        || !sheet.resolvedSparklines?.get(merge.minRow)?.has(merge.minCol)) continue
      r0 = Math.min(r0, merge.minRow); r1 = Math.max(r1, merge.maxRow)
      c0 = Math.min(c0, merge.minCol); c1 = Math.max(c1, merge.maxCol)
    }
    const cf = cfExplicitBounds(sheet)
    if (cf) {
      const canMatchBlanks = sheet.conditionalFormatting?.rules.some(r =>
        (r.definition as any).type === 'containsBlanks' || (r.definition as any).type === 'expression'
      )
      if (canMatchBlanks || r1 < 0) {
        r0 = Math.min(r0, cf.r0); r1 = Math.max(r1, cf.r1); c0 = Math.min(c0, cf.c0); c1 = Math.max(c1, cf.c1)
      }
    }
    this.dynamicAnchors = dynamic
    this.baseExtent = Object.freeze(r1 < 0 ? { rows: MAX_ROWS, cols: MAX_COLS } : { rows: MAX_ROWS, cols: MAX_COLS, used: { r0, r1, c0, c1 } })
  }
  get completeness() { return 'complete' as const }
  get stamp(): SourceStamp { return this.currentStamp }
  private get valueOverlay(): ValueOverlay | undefined { return overlayStates.get(this)?.overlay }
  get structureGen(): number { return overlayStates.get(this)?.structureGen ?? 0 }
  get extent(): SourceStore['extent'] {
    const overlay = this.valueOverlay
    const base = this.baseExtent.used
    if (!overlay) return this.baseExtent
    const bounds = overlay.protectedRects
    const used = bounds.length ? bounds.reduce((all, rect) => ({ r0: Math.min(all.r0, rect.r0), r1: Math.max(all.r1, rect.r1), c0: Math.min(all.c0, rect.c0), c1: Math.max(all.c1, rect.c1) }), { ...bounds[0] }) : undefined
    if (!used && !base) return { rows: MAX_ROWS, cols: MAX_COLS }
    return { rows: MAX_ROWS, cols: MAX_COLS, used: base && used
      ? { r0: Math.min(base.r0, used.r0), r1: Math.max(base.r1, used.r1), c0: Math.min(base.c0, used.c0), c1: Math.max(base.c1, used.c1) }
      : used ?? base }
  }
  setValueOverlay(overlay: ValueOverlay | undefined): void {
    const previous = overlayStates.get(this)
    // The overlay already installed is a no-op, checked before the (deferred T26 m1) assertion (Task 33 perf).
    if (previous?.overlay === overlay || (!previous && overlay === undefined)) return
    if (overlay) assertOverlayCellsProtected(overlay)
    const structureGen = (previous?.structureGen ?? 0) + (sameOverlayMerges(previous?.overlay?.merges, overlay?.merges) ? 0 : 1)
    overlayStates.set(this, { overlay, structureGen })
    this.currentStamp = Object.freeze({ ...this.currentStamp, valueGen: this.currentStamp.valueGen + 1 })
  }
  noteValuesChanged(options?: { rows?: boolean }): void {
    // A spill that grew into rows the sheet did not have, or left rows empty: the inventory is taken again (PR #57).
    if (options?.rows) this.reindex()
    this.currentStamp = Object.freeze({ ...this.currentStamp, valueGen: this.currentStamp.valueGen + 1 })
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
  /** Drawing anchors whose placement is unknown: unreadable anchors/parts, or
   * parsed drawings on a sheet whose anchors were never extracted. */
  private get unresolvedDrawings(): number {
    const anchors = this.sheet.drawingAnchors
    return anchors === undefined ? this.sheet.drawings?.length ?? 0 : anchors.unresolved
  }
  /** Dynamic anchors whose spill extent is unknown (no saved ref, no engine footprint). */
  private unknownSpills(): { row: number; col: number }[] {
    const footprints = this.valuePolicy === 'cached' ? [] : this.engineSpills
    return this.dynamicAnchors.filter(at => {
      const cell = this.cellIn(this.rows[lowerBound(this.rowIndexes, at.row)], at.col)
      const saved = cell?.arrayRef === undefined ? null : parseRange(cell.arrayRef)
      if (footprints === undefined) return true
      return saved === null && !footprints.some(f => f.anchor.row === at.row && f.anchor.col === at.col)
    })
  }
  /** Coverage gaps this adapter can never close by waiting. */
  coverageDiagnostics(): WorkerDiagnostic[] {
    const out: WorkerDiagnostic[] = []
    if (this.unresolvedDrawings > 0) {
      out.push({ code: 'coverage-unresolvable', message: 'Some drawing anchors could not be read; their placement is unknown' })
    }
    const unknown = this.unknownSpills()
    if (unknown.length) {
      out.push({ code: 'coverage-unresolvable', message: 'A dynamic-array spill has no known extent', source: unknown[0] })
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
  private overlayPayload(row: number, col: number, value: NonNullable<ReturnType<ValueOverlay['cell']>>): SingleCellPayload {
    if (value === null) throw new RangeError('Blank overlay has no cell payload')
    return { id: cellId(row, col), row, col, value: value.value, styleIndex: 0, hasCachedValue: true, valueIsError: value.valueIsError ?? false, valueProvenance: 'engine' }
  }
  displayText(cell: AuthoritativeCell, allocation?: FormatAllocation): string {
    const k = lowerBound(this.rowIndexes, cell.row)
    const original = this.rowIndexes[k] === cell.row ? this.cellIn(this.rows[k], cell.col) : undefined
    const overlay = this.valueOverlay ? this.overlayCell(cell.row, cell.col) : undefined
    const style = overlay && overlay !== null ? (overlay.numFmtId !== undefined ? { numFmtId: overlay.numFmtId } : undefined) : original?.style
    return cellDisplayText(this.sheet, { value: cell.value, style }, allocation)
  }
  async readCell(at: { row: number; col: number }, options: ReadOptions): Promise<CellRead> {
    const error = this.guard(options)
    if (error) return { kind: 'error', code: error.code, retryable: error.retryable }
    if (!isValidCoordinate(at.row, at.col)) return { kind: 'error', code: 'invalid-coordinate', retryable: false }
    const overlay = this.overlayCell(at.row, at.col)
    if (overlay !== undefined) return overlay === null ? { kind: 'blank' } : { kind: 'cell', cell: this.overlayPayload(at.row, at.col, overlay) }
    const k = lowerBound(this.rowIndexes, at.row)
    const cell = this.rowIndexes[k] === at.row ? this.cellIn(this.rows[k], at.col) : undefined
    if (!cell) return { kind: 'blank' }
    const payload = toPayload(cell, this.valuePolicy)
    return isUnavailable(cell, this.valuePolicy) ? { kind: 'unavailable', cell: payload, reason: 'formula-value-unavailable' } : { kind: 'cell', cell: payload }
  }
  async readRows(range: SourceRect, cursor: string | undefined, options: ReadOptions & SourceAllocationOptions): Promise<Page<SparseRowPage>> {
    const error = this.guard(options)
    if (error) return failedPage(this.stamp, error)
    if (!isValidSourceRect(range)) return failedPage(this.stamp, failure('invalid-source-rect', 'Invalid range'))
    if (this.valueOverlay) return this.readRowsWithOverlay(range, cursor, options)
    const key = `objrows:${range.r0}:${range.r1}:${range.c0}:${range.c1}:`
    let at = lowerBound(this.rowIndexes, range.r0)
    if (cursor !== undefined) {
      const offset = cursor.startsWith(key) ? Number(cursor.slice(key.length)) : NaN
      if (!Number.isSafeInteger(offset) || offset < at || offset > this.rows.length) {
        return failedPage(this.stamp, failure('invalid-cursor', 'Invalid row cursor'))
      }
      at = offset
    }
    const account = options[SOURCE_ALLOCATION]?.reserve()
    account?.charge(96)
    const budget = requestBudget(() => this.guard(options), this.pacing)
    try {
    const items: SparseRowPage[] = []
    let visited = 0, cellCount = 0
    while (at < this.rows.length && this.rowIndexes[at] <= range.r1 && visited < READ_ROWS_PAGE_ROWS) {
      const row = this.rows[at]
      account?.charge(72)
      const cells: AuthoritativeCell[] = []
      for (const cell of row.cells) {
        if (budget.spend(1)) await budget.pause()
        if (cell.row === row.index && cell.col >= range.c0 && cell.col <= range.c1) {
          account?.charge(384 + (isUnavailable(cell, this.valuePolicy) ? 288 : 0) + (cell.sharedFormula ? 96 : 0))
          cells.push(toRowCell(cell, this.valuePolicy))
        }
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
    await budget.finish()
    const page = { items, completeness: 'complete' as const, stamp: this.stamp, ...(more ? { nextCursor: `${key}${at}` } : {}) }
    if (account) adoptAllocation(account, page, account.bytes, [this.stamp])
    return page
    } catch (stop) {
      if (stop instanceof SourceStop) return failedPage(this.stamp, stop.failure)
      throw stop
    } finally { account?.release(); budget.dispose() }
  }
  private async readRowsWithOverlay(range: SourceRect, cursor: string | undefined, options: ReadOptions & SourceAllocationOptions): Promise<Page<SparseRowPage>> {
    const key = `objrows-overlay:${range.r0}:${range.r1}:${range.c0}:${range.c1}:`
    let row = range.r0
    if (cursor !== undefined) {
      const next = cursor.startsWith(key) ? Number(cursor.slice(key.length)) : NaN
      if (!Number.isSafeInteger(next) || next < range.r0 || next > range.r1 + 1) return failedPage(this.stamp, failure('invalid-cursor', 'Invalid row cursor'))
      row = next
    }
    const account = options[SOURCE_ALLOCATION]?.reserve()
    account?.charge(96)
    const budget = requestBudget(() => this.guard(options), this.pacing)
    try {
      const items: SparseRowPage[] = []
      let visited = 0, cellCount = 0
      while (row <= range.r1 && visited < READ_ROWS_PAGE_ROWS) {
        const baseAt = lowerBound(this.rowIndexes, row)
        let candidate = baseAt < this.rowIndexes.length && this.rowIndexes[baseAt] <= range.r1 ? this.rowIndexes[baseAt] : range.r1 + 1
        for (const rect of this.valueOverlay!.rects) {
          if (rect.r1 < row || rect.c1 < range.c0 || rect.c0 > range.c1) continue
          candidate = Math.min(candidate, Math.max(row, rect.r0))
        }
        if (candidate > range.r1) break
        row = candidate
        if (budget.spend(1)) await budget.pause()
        const baseIndex = lowerBound(this.rowIndexes, row)
        const base = baseIndex < this.rowIndexes.length && this.rowIndexes[baseIndex] === row ? this.rows[baseIndex] : undefined
        const cells = new Map<number, AuthoritativeCell>()
        if (base) for (const cell of base.cells) if (cell.row === row && cell.col >= range.c0 && cell.col <= range.c1) {
          account?.charge(384 + (isUnavailable(cell, this.valuePolicy) ? 288 : 0) + (cell.sharedFormula ? 96 : 0))
          cells.set(cell.col, toRowCell(cell, this.valuePolicy))
        }
        const overlay = this.valueOverlay!
        for (const rect of overlay.rects) {
          if (row < rect.r0 || row > rect.r1) continue
          for (let col = Math.max(range.c0, rect.c0); col <= Math.min(range.c1, rect.c1); col++) {
            const value = overlay.cell(row, col)
            if (value === undefined) continue
            cells.delete(col)
            if (value !== null) {
              account?.charge(384)
              cells.set(col, this.overlayPayload(row, col, value))
            }
          }
        }
        const rowCells = [...cells.values()].sort((a, b) => a.col - b.col)
        if (rowCells.length > this.maxCells) return failedPage(this.stamp, failure('resource-limit', 'A source row exceeds the page cell limit'))
        if (items.length && cellCount + rowCells.length > this.maxCells) break
        visited++; cellCount += rowCells.length
        if (rowCells.length) items.push({ row, cells: rowCells })
        row++
      }
      const more = row <= range.r1 && this.valueOverlay!.rects.some(rect => rect.r1 >= row && rect.r0 <= range.r1 && rect.c1 >= range.c0 && rect.c0 <= range.c1)
        || row <= range.r1 && lowerBound(this.rowIndexes, row) < this.rowIndexes.length && this.rowIndexes[lowerBound(this.rowIndexes, row)] <= range.r1
      await budget.finish()
      const page = { items, completeness: 'complete' as const, stamp: this.stamp, ...(more ? { nextCursor: `${key}${row}` } : {}) }
      if (account) adoptAllocation(account, page, account.bytes, [this.stamp])
      return page
    } catch (stop) {
      if (stop instanceof SourceStop) return failedPage(this.stamp, stop.failure)
      throw stop
    } finally { account?.release(); budget.dispose() }
  }
  /** Column page straight from the row index: populated entries plus explicit
   * blank runs covering every requested row; nothing is cached per column. */
  async queryColumn(request: ColumnQuery, options: ReadOptions & SourceAllocationOptions): Promise<Page<ColumnEntry>> {
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
    const startAt = at, startRow = row
    const overlay = this.valueOverlay
    const account = options[SOURCE_ALLOCATION]?.reserve()
    account?.charge(96)
    try {
    const finish = (page: Page<ColumnEntry>) => { if (account) adoptAllocation(account, page, account.bytes, [this.stamp]); return page }
    if (request.encoding === 'lanes') return finish(overlay
      ? this.overlayLanesPage(request, intervals, at, row, prefix, account)
      : this.lanesPage(request, intervals, at, row, prefix, account))
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
        account?.charge(72)
        items.push({ kind: 'blank', r0: row, r1 })
        at++
        if (at < intervals.length) row = intervals[at].r0
        continue
      }
      if (found.row > row) {
        account?.charge(72)
        items.push({ kind: 'blank', r0: row, r1: found.row - 1 })
        if (items.length >= request.limit) { row = found.row; break }
      }
      account?.charge(384 + (isUnavailable(found, this.valuePolicy) ? 288 : 0) + (found.sharedFormula ? 96 : 0))
      const payload = toPayload(found, this.valuePolicy)
      items.push(isUnavailable(found, this.valuePolicy)
        ? { kind: 'unavailable', row: found.row, cell: payload, reason: 'formula-value-unavailable' }
        : { kind: 'cell', row: found.row, cell: payload })
      row = found.row + 1
      if (row > r1) { at++; if (at < intervals.length) row = intervals[at].r0 }
    }
    const page: Page<ColumnEntry> = { items, completeness: 'complete', stamp: this.stamp, ...(at < intervals.length ? { nextCursor: `${prefix}:${at}:${row}` } : {}) }
    return finish(overlay ? this.overlayColumnPage(page, request, intervals, startAt, startRow, prefix, overlay) : page)
    } finally { account?.release() }
  }
  async *scanPopulated(request: PopulatedQuery, options: ReadOptions & SourceAllocationOptions, checkOverlay = true): AsyncIterable<Page<AuthoritativeCell>> {
    if (checkOverlay && this.valueOverlay) {
      yield* this.scanOverlayRows(request, options, false) as AsyncIterable<Page<AuthoritativeCell>>
      return
    }
    const error = this.guard(options)
    if (error) { yield failedPage(this.stamp, error); return }
    if (!isValidSourceRect(request.range) || !Number.isSafeInteger(request.limit) || request.limit <= 0) {
      yield failedPage(this.stamp, failure('invalid-source-rect', 'Invalid populated query')); return
    }
    let account = options[SOURCE_ALLOCATION]?.reserve()
    account?.charge(96)
    const budget = requestBudget(() => this.guard(options), this.pacing)
    try {
    let items: AuthoritativeCell[] = []
    for (let at = lowerBound(this.rowIndexes, request.range.r0); at < this.rows.length && this.rowIndexes[at] <= request.range.r1; at++) {
      for (const cell of this.rows[at].cells) {
        try { if (budget.spend(1)) await budget.pause() } catch (stop) {
          if (!(stop instanceof SourceStop)) throw stop
          yield failedPage(this.stamp, stop.failure); return
        }
        if (cell.col < request.range.c0 || cell.col > request.range.c1 || !isPopulated(cell)) continue
        account?.charge(384 + (isUnavailable(cell, this.valuePolicy) ? 288 : 0) + (cell.sharedFormula ? 96 : 0))
        items.push(toRowCell(cell, this.valuePolicy))
        if (items.length === request.limit) {
          const page = { items, completeness: 'complete' as const, stamp: this.stamp }
          if (account) adoptAllocation(account, page, account.bytes, [this.stamp])
          account?.release()
          yield page
          account = options[SOURCE_ALLOCATION]?.reserve()
          account?.charge(96)
          items = []
          const again = this.guard(options)
          if (again) { yield failedPage(this.stamp, again); return }
        }
      }
    }
    try { await budget.finish() } catch (stop) {
      if (!(stop instanceof SourceStop)) throw stop
      yield failedPage(this.stamp, stop.failure); return
    }
    const page = { items, completeness: 'complete' as const, stamp: this.stamp }
    if (account) adoptAllocation(account, page, account.bytes, [this.stamp])
    account?.release()
    yield page
    } finally { account?.release(); budget.dispose() }
  }
  private async *scanOverlayRows(request: PopulatedQuery, options: ReadOptions & SourceAllocationOptions, occupiedOnly: boolean): AsyncIterable<Page<AuthoritativeCell | SourceCell>> {
    const error = this.guard(options)
    if (error) { yield failedPage(this.stamp, error); return }
    if (!isValidSourceRect(request.range) || !Number.isSafeInteger(request.limit) || request.limit <= 0) {
      yield failedPage(this.stamp, failure('invalid-source-rect', 'Invalid populated query')); return
    }
    let cursor: string | undefined
    let items: Array<AuthoritativeCell | SourceCell> = []
    let account = options[SOURCE_ALLOCATION]?.reserve()
    account?.charge(96)
    const publish = (page: Page<AuthoritativeCell | SourceCell>) => {
      if (account) adoptAllocation(account, page, account.bytes, [this.stamp])
      account?.release()
      account = undefined
      return page
    }
    // The inner row pages are not this scan's output: read them uncharged (the
    // published pages charge each cell once) so the caller never pays twice.
    const { [SOURCE_ALLOCATION]: _allocation, ...readOptions } = options
    try { for (;;) {
      const page = await this.readRows(request.range, cursor, readOptions)
      for (const row of page.items) for (const cell of row.cells) {
        if (!(cell.value !== null || cell.formula !== undefined)) continue
        account?.charge(occupiedOnly ? 16 : 384 + (cell.valueIsError ? 288 : 0))
        items.push(occupiedOnly ? { row: cell.row, col: cell.col } : cell)
        if (items.length === request.limit) {
          yield publish({ items, completeness: 'complete', stamp: this.stamp })
          items = []
          account = options[SOURCE_ALLOCATION]?.reserve()
          account?.charge(96)
        }
      }
      if (page.error) { yield publish({ items, completeness: 'incomplete', stamp: this.stamp, error: page.error }); return }
      if (page.completeness !== 'complete') { yield publish({ items, completeness: page.completeness, stamp: this.stamp, ...(page.nextCursor ? { nextCursor: page.nextCursor } : {}) }); return }
      if (!page.nextCursor) break
      cursor = page.nextCursor
    }
    yield publish({ items, completeness: 'complete', stamp: this.stamp })
    } finally { account?.release() }
  }
  async *scanOccupied(request: PopulatedQuery, options: ReadOptions & SourceAllocationOptions): AsyncIterable<Page<SourceCell>> {
    const overlay = this.valueOverlay
    if (overlay) {
      yield* this.scanOverlayRows(request, options, true) as unknown as AsyncIterable<Page<SourceCell>>
      return
    }
    for await (const page of this.scanPopulated(request, options, false)) yield { ...page, items: page.items.map(cell => ({ row: cell.row, col: cell.col })) }
  }
  async queryStructural(rect: SourceRect, options: ReadOptions & SourceAllocationOptions): Promise<StructuralResult> {
    const base = { stamp: this.stamp, merges: [], tables: [], spills: [], drawings: [] }
    const error = this.guard(options) ?? (isValidSourceRect(rect) ? undefined : failure('invalid-source-rect', 'Invalid range'))
    if (error) return { ...base, completeness: 'incomplete', missing: [rect], error }
    const account = options[SOURCE_ALLOCATION]?.reserve()
    // Sparse authored metadata bounds the new records; no full-axis materialization.
    let spillCount = this.engineSpills?.length ?? 0
    for (const row of this.rows) for (const cell of row.cells) if (cell.arrayRef !== undefined || cell.dynamicArray) spillCount++
    // Drawing anchors are the adapter's resolved cache (charged once, below):
    // a result only adds its own array slots.
    account?.charge(256 + 320 * (this.sheet.mergeRanges?.length ?? this.sheet.merges.length) + 256 * (this.tables?.length ?? 0) + 448 * spillCount + 16 * (this.sheet.drawingAnchors?.records.length ?? 0))
    try {
    const overlay = this.valueOverlay
    const merges = this.validMerges()
      .filter(m => !overlay?.rects.some(rect => m.minRow <= rect.r1 && rect.r0 <= m.maxRow && m.minCol <= rect.c1 && rect.c0 <= m.maxCol))
      .map(m => ({ id: `merge:${m.minRow}:${m.minCol}:${m.maxRow}:${m.maxCol}`, range: { r0: m.minRow, r1: m.maxRow, c0: m.minCol, c1: m.maxCol }, anchor: { row: m.minRow, col: m.minCol } }))
      .filter(m => isValidSourceRect(m.range) && intersects(m.range, rect))
    if (overlay) for (const merge of overlay.merges) if (intersects(merge, rect)) merges.push({ id: `pivot-merge:${merge.r0}:${merge.c0}:${merge.r1}:${merge.c1}`, range: merge, anchor: { row: merge.r0, col: merge.c0 } })
    const tables = (this.tables ?? [])
      .map(t => ({ id: t.id, range: { r0: t.extent.firstRow, r1: t.extent.firstRow + t.extent.rows - 1, c0: t.extent.firstCol, c1: t.extent.firstCol + t.extent.cols - 1 }, headerRows: t.headerRowCount, totalsRows: t.totalsRowCount }))
      .filter(t => isValidSourceRect(t.range) && intersects(t.range, rect))
    const spills: SpillExtent[] = []
    const missing: SourceRect[] = []
    const addMissing = (domain: SourceRect) => {
      const r0 = Math.max(domain.r0, rect.r0), r1 = Math.min(domain.r1, rect.r1)
      const c0 = Math.max(domain.c0, rect.c0), c1 = Math.min(domain.c1, rect.c1)
      if (r0 <= r1 && c0 <= c1) missing.push({ r0, r1, c0, c1 })
    }
    // A dynamic array with no known extent may reach anywhere from its anchor to
    // the sheet edge: that domain stays unknown, never assumed empty.
    for (const anchor of this.unknownSpills()) addMissing({ r0: anchor.row, r1: MAX_ROWS - 1, c0: anchor.col, c1: MAX_COLS - 1 })
    // Under `cached` the saved footprint is the published extent; under `engine`
    // the engine's footprint is, and the saved one stays in the union because a
    // retained (not recalculated) owner still shows its saved followers.
    const footprints = new Map<number, SourceRect>()
    if (this.valuePolicy === 'engine') for (const f of this.engineSpills ?? []) footprints.set(cellId(f.anchor.row, f.anchor.col), { r0: f.r0, r1: f.r1, c0: f.c0, c1: f.c1 })
    const add = (spill: SpillExtent) => {
      const anchorRect = { r0: spill.anchor.row, r1: spill.anchor.row, c0: spill.anchor.col, c1: spill.anchor.col }
      if ([spill.saved, spill.evaluated, anchorRect].some(extent => extent !== undefined && intersects(extent, rect))) spills.push(spill)
    }
    for (const row of this.rows) {
      for (const cell of row.cells) {
        if (cell.arrayRef === undefined && !(cell.dynamicArray && footprints.has(cellId(cell.row, cell.col)))) continue
        const saved = cell.arrayRef === undefined ? null : parseRange(cell.arrayRef)
        const anchor = { row: cell.row, col: cell.col }
        const evaluated = cell.dynamicArray ? (this.valuePolicy === 'cached' ? saved ?? undefined : footprints.get(cellId(cell.row, cell.col))) : undefined
        footprints.delete(cellId(cell.row, cell.col))
        add({ anchor, kind: cell.dynamicArray ? 'dynamic' : 'fixed', ...(saved ? { saved } : {}), ...(evaluated ? { evaluated } : {}) })
      }
    }
    // Owners the engine spilled without file array metadata (explicit dynamic mode).
    for (const [id, evaluated] of footprints) add({ anchor: { row: Math.floor(id / MAX_COLS), col: id % MAX_COLS }, kind: 'dynamic', evaluated })
    // Unreadable drawing anchors and an unavailable table inventory cannot prove "none here".
    if (this.unresolvedDrawings > 0 || this.tables === undefined) addMissing(rect)
    // The resolved anchors live as long as this adapter: charged once per
    // ledger, never again by each result that borrows them. A first resolution
    // is admitted before the cache is built and published; a refusal (or any
    // failure before its adoption) leaves no cache behind.
    const memory = options[SOURCE_ALLOCATION], records = this.sheet.drawingAnchors?.records.length ?? 0
    const fresh = this.drawingCache === undefined
    let cached: DrawingAnchor[]
    if (memory && records && (fresh || !memory.owns(this.drawingCache!))) {
      const cacheAccount = memory.reserve()
      try {
        cacheAccount.charge(320 * records)
        cached = this.drawings()
        if (cached.length) adoptAllocation(cacheAccount, cached)
      } catch (error) { if (fresh) this.drawingCache = undefined; throw error }
      finally { cacheAccount.release() }
    } else cached = this.drawings()
    const drawings = cached.filter(d => d.kind === 'absolute' || intersects(d.bounds, rect))
    // Every saved pivot is guarded with its page-field cells, overlay or not (S3-G). Vacated cells read
    // empty but stay guarded inside a saved area (S3-L): the guard is saved ∪ the overlay's protected
    // bounds, whatever the producer names, never the overlay alone.
    const pivots = uniqueRectList([...this.savedPivots, ...(overlay?.protectedRects ?? [])]).filter(pivot => intersects(pivot, rect))
    const result: StructuralResult = { ...base, merges, tables, spills, drawings, ...(pivots?.length ? { pivots } : {}), completeness: missing.length ? 'incomplete' : 'complete', missing }
    if (account) adoptAllocation(account, result, account.bytes, [this.stamp, ...drawings])
    return result
    } finally { account?.release() }
  }
  /** Resolved once: the parsed sheet behind this adapter never changes. Known gap
   * (pre-existing): this geometry omits `columnMetrics`, so anchors use the
   * pre-MDW column widths while the grid uses MDW widths (paged passes them). */
  private drawings(): DrawingAnchor[] {
    this.drawingCache ??= resolveDrawingAnchors(this.sheet.drawingAnchors?.records ?? [],
      withFrozen(buildGeometry([...this.sheet.rows].sort((a, b) => a.index - b.index), this.sheet.cols, [], { hiddenCols: true }), this.sheet.frozen))
    return this.drawingCache
  }
  /** One lanes run per interval slice; a page covers at most `limit` rows. */
  private lanesPage(request: ColumnQuery, intervals: readonly SourceRowInterval[], at: number, row: number, prefix: string, account?: AllocationReservation): Page<ColumnEntry> {
    const items: ColumnEntry[] = []
    let budget = request.limit
    while (at < intervals.length && budget > 0) {
      if (row < intervals[at].r0) row = intervals[at].r0
      const r1 = Math.min(intervals[at].r1, row + budget - 1)
      account?.charge(320)
      const lanes = new LaneBuilder(account)
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
  private overlayLanesPage(request: ColumnQuery, intervals: readonly SourceRowInterval[], at: number, row: number, prefix: string, account?: AllocationReservation): Page<ColumnEntry> {
    const items: ColumnEntry[] = []
    let budget = request.limit
    while (at < intervals.length && budget > 0) {
      if (row < intervals[at].r0) row = intervals[at].r0
      const r1 = Math.min(intervals[at].r1, row + budget - 1)
      account?.charge(320)
      const lanes = new LaneBuilder(account)
      for (let current = row; current <= r1; current++) {
        const overlay = this.overlayCell(current, request.col)
        const rowIndex = lowerBound(this.rowIndexes, current)
        const sourceRow = this.rowIndexes[rowIndex] === current ? this.rows[rowIndex] : undefined
        const cell = overlay === undefined && sourceRow ? this.cellIn(sourceRow, request.col) : undefined
        if (overlay === null) continue
        if (overlay !== undefined) lanes.add(current, overlay.value, overlay.valueIsError, 0, false)
        else if (cell) lanes.add(current, cell.value, cell.valueIsError, cell.styleIndex, isUnavailable(cell, this.valuePolicy))
      }
      items.push(lanes.build(row, r1))
      budget -= r1 - row + 1
      row = r1 + 1
      if (row > intervals[at].r1) { at++; if (at < intervals.length) row = intervals[at].r0 }
    }
    return { items, completeness: 'complete', stamp: this.stamp, ...(at < intervals.length ? { nextCursor: `${prefix}:${at}:${row}` } : {}) }
  }
  private overlayColumnPage(page: Page<ColumnEntry>, request: ColumnQuery, intervals: readonly SourceRowInterval[], startAt: number, startRow: number, prefix: string, overlay: ValueOverlay): Page<ColumnEntry> {
    if (!overlay || page.items.some(item => item.kind === 'lanes')) return page
    if (page.completeness !== 'complete') return overlayIncompleteColumnPage(page, request, intervals, startAt, prefix, overlay, row => this.overlayCell(row, request.col), (row, value) => this.overlayPayload(row, request.col, value))
    let endAt = intervals.length - 1, endRow = intervals[endAt]?.r1 ?? -1
    if (page.nextCursor) {
      const parts = page.nextCursor.split(':')
      const i = Number(parts[1]), r = Number(parts[2])
      if (parts.length === 3 && parts[0] === prefix && Number.isSafeInteger(i) && Number.isSafeInteger(r)) {
        endAt = i; endRow = r - 1
      }
    }
    const cells = new Map<number, Extract<ColumnEntry, { kind: 'cell' | 'unavailable' }>>()
    for (const item of page.items) if (item.kind === 'cell' || item.kind === 'unavailable') cells.set(item.row, item)
    for (const rect of overlay.rects) {
      if (request.col < rect.c0 || request.col > rect.c1) continue
      for (let i = startAt; i <= endAt; i++) {
        const interval = intervals[i]
        if (!interval) continue
        const r0 = Math.max(interval.r0, rect.r0, i === startAt ? startRow : interval.r0)
        const r1 = Math.min(interval.r1, rect.r1, i === endAt ? endRow : interval.r1)
        for (let row = r0; row <= r1; row++) {
          const value = overlay.cell(row, request.col)
          if (value === undefined) continue
          if (value === null) cells.delete(row)
          else cells.set(row, { kind: 'cell', row, cell: this.overlayPayload(row, request.col, value) })
        }
      }
    }
    const output: ColumnEntry[] = []
    let continuation: { interval: number; row: number } | undefined
    const limit = request.limit
    for (let i = startAt; i <= endAt && !continuation; i++) {
      const interval = intervals[i]
      if (!interval) continue
      const r0 = Math.max(interval.r0, i === startAt ? startRow : interval.r0)
      const r1 = Math.min(interval.r1, i === endAt ? endRow : interval.r1)
      if (r0 > r1) continue
      const events = [...cells.values()].filter(entry => entry.row >= r0 && entry.row <= r1).sort((a, b) => a.row - b.row)
      let row = r0
      for (const entry of events) {
        if (entry.row > row) {
          if (output.length >= limit) { continuation = { interval: i, row }; break }
          output.push({ kind: 'blank', r0: row, r1: entry.row - 1 })
        }
        if (output.length >= limit) { continuation = { interval: i, row: entry.row }; break }
        output.push(entry); row = entry.row + 1
      }
      if (!continuation && row <= r1) {
        if (output.length >= limit) continuation = { interval: i, row }
        else output.push({ kind: 'blank', r0: row, r1 })
      }
    }
    const nextCursor = continuation ? `${prefix}:${continuation.interval}:${continuation.row}` : page.nextCursor
    const { nextCursor: _oldCursor, ...rest } = page
    return { ...rest, items: output, ...(nextCursor ? { nextCursor } : {}) }
  }
  /** S13.2: the parsed drawings, cloned (the resident model keeps its own bytes
   * and workbook-wide picture indexes) and charged against the transfer budget. */
  async readDrawings(options: ReadOptions, profile: ProfileId, memory?: DrawingMemory): Promise<SheetDrawings> {
    const error = this.guard(options)
    if (error) return unavailableDrawings(`Drawings could not be read: ${error.message}`, error)
    // The clone is this read's working set in the document pool.
    if (memory && this.sheet.drawings?.length && !memory.set(drawingFootprint(this.sheet.drawings))) {
      return unavailableDrawings('This sheet\'s drawings need more worker memory than the document has free and are not shown', failure('resource-limit', 'worker memory budget'), DRAWING_MEMORY_REFUSED)
    }
    let drawings: XlsxDrawing[]
    try { drawings = this.sheet.drawings?.length ? structuredClone(this.sheet.drawings) : [] }
    catch (cause) { return unavailableDrawings(`Drawings could not be prepared: ${cause instanceof Error ? cause.message : String(cause)}`) }
    return sheetDrawingPayload(drawings, this.sheet.drawingTheme, [], this.drawingPayloadBytes ?? drawingPayloadBudget(profile))
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
  return withFrozen(buildGeometry([...sheet.rows].sort((a, b) => a.index - b.index), sheet.cols, [], { hiddenCols: true, ...(sheet.columnMetrics ? { columns: sheet.columnMetrics } : {}) }), sheet.frozen)
}
/** Geometry carrying the sheet's authored frozen panes (copied, frozen). */
export function withFrozen(geometry: SheetGeometry, frozen: SheetGeometry['frozen']): SheetGeometry {
  return frozen ? Object.freeze({ ...geometry, frozen: Object.freeze({ ...frozen }) }) : geometry
}
/** Source-space geometry from authored rows (ascending) and column specs; shared by both adapters. */
export function buildGeometry(
  rows: Iterable<{ index: number; heightPt?: number; hidden?: boolean }>,
  specs: readonly { min: number; max: number; widthChars?: number; hidden?: boolean }[],
  hiddenIntervals: readonly SourceRowInterval[] = [],
  /** S5.4 paged sources: authored-hidden columns collapse (zero width), as the object planner collapses them. */
  options: { hiddenCols?: boolean; columns?: ColumnMetrics } = {},
): SheetGeometry {
  const columns = options.columns ?? DEFAULT_COLUMN_METRICS
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
      colOverrides.push({ start, end, sizeCss: columnWidthPx(spec.widthChars, columns.maxDigitWidthPx) })
    }
  }
  colOverrides.sort((a, b) => a.start - b.start)
  // Overlapping authored <col> ranges: the later start wins, keeping intervals disjoint.
  for (let i = 1; i < colOverrides.length; i++) {
    if (colOverrides[i].start <= colOverrides[i - 1].end) colOverrides[i - 1] = { ...colOverrides[i - 1], end: colOverrides[i].start - 1 }
  }
  const authored = colOverrides.filter(interval => interval.start <= interval.end)
  // Worker layouts publish only overrides (the remote axis assumes 64px), so a
  // non-default sheet width is spelled out over every gap between them.
  const cols: AxisSizeInterval[] = []
  if (columns.defaultWidthPx === DEFAULT_COL_CSS) cols.push(...authored)
  else {
    let next = 0
    for (const interval of authored) {
      if (interval.start > next) cols.push({ start: next, end: interval.start - 1, sizeCss: columns.defaultWidthPx })
      cols.push(interval)
      next = interval.end + 1
    }
    if (next <= MAX_COLS - 1) cols.push({ start: next, end: MAX_COLS - 1, sizeCss: columns.defaultWidthPx })
  }
  const hiddenCols: Array<{ start: number; end: number }> = []
  if (options.hiddenCols) {
    for (const spec of [...specs].filter(spec => spec.hidden).sort((a, b) => a.min - b.min)) {
      const start = Math.max(0, spec.min), end = Math.min(MAX_COLS - 1, spec.max)
      if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start > end) continue
      const last = hiddenCols[hiddenCols.length - 1]
      if (last && start <= last.end + 1) last.end = Math.max(last.end, end)
      else hiddenCols.push({ start, end })
    }
  }
  return Object.freeze({
    rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: DEFAULT_ROW_CSS, overrides: rowOverrides }),
    cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: DEFAULT_COL_CSS, overrides: cols, ...(hiddenCols.length ? { collapsed: hiddenCols } : {}) }),
    rowOverrides: Object.freeze(rowOverrides),
    colOverrides: Object.freeze(cols),
    ...(hiddenCols.length ? { hiddenCols: Object.freeze(hiddenCols) } : {}),
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
  let footprints: ReturnType<typeof engineSpillFootprints> | null = null
  const spillFootprints = () => footprints === null ? (footprints = engineSpillFootprints(doc)) : footprints
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
      store = new ObjectXlsxSheetAdapter(stamp, sheet, doc.tables?.filter(table => table.sheetId === sheetId), maxCells, policy, spillFootprints()?.get(sheet) ?? (spillFootprints() ? [] : undefined), options.drawingPayloadBytes, options.pacing)
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
    scanFormulaCellsOutsidePivots(signal?: AbortSignal) {
      for (const { sheet } of sheets) {
        if (signal?.aborted) return undefined
        const refs = [...(sheet.pivotTables ?? []).map(def => String(def.location.ref)), ...(sheet.pivotRegions ?? []).map(region => region.ref)]
        const pivots = refs.flatMap(ref => { const rect = parseA1Range(ref); return rect ? [rect] : [] })
        for (const row of sheet.rows) for (const cell of row.cells) if (cell.formula || cell.sharedFormula) {
          if (signal?.aborted) return undefined
          if (!pivots.some(r => cell.row >= r.r0 && cell.row <= r.r1 && cell.col >= r.c0 && cell.col <= r.c1)) return true
        }
      }
      return false
    },
    async pivotAssets(sheetId, signal, profile, memory): Promise<WorkerPivotAssets | undefined> {
      if (signal.aborted) return undefined
      const sheet = find(sheetId).sheet
      if (!sheet.pivotTables?.length) return undefined
      const { pivotRuntimeFor } = await import('../pivotLoader')
      const runtime = pivotRuntimeFor(doc)
      if (!runtime) return undefined
      const { sheetCellReader } = await import('../pivots/engine/faithful')
      const cells = new Set<number>()
      const linkedCharts: Array<{ sheetId: string; partPath: string; treePath: string; chart: import('../../charts/model').ChartModel }> = []
      doc.sheets.forEach((candidate, index) => collectObjectPivotCharts(String(candidate.sheetId ?? index + 1), candidate.drawings, linkedCharts))
      let metadataBytes = 512 + estimatePayloadBytes(linkedCharts)
      if (memory && !memory.set(metadataBytes)) return undefined
      for (const row of sheet.rows) for (const cell of row.cells) {
        if (cell.value !== null || cell.formula !== undefined) {
          metadataBytes += 32
          if (memory && !memory.set(metadataBytes)) return undefined
          cells.add(cell.row * MAX_COLS + cell.col)
        }
      }
      const [{ RECORD_LIMITS }, { sheetBlockers, structuralOccupancy }] = await Promise.all([import('../pivots/engine/records'), import('../pivots/session/occupancy')])
      // Merges, array and spill ranges and noted cells count too (S3-AC), charged with the metadata.
      const structure = structuralOccupancy(sheetBlockers(sheet, spillFootprints()?.get(sheet)))
      metadataBytes += structure.bytes
      if (memory && !memory.set(metadataBytes)) return undefined
      return {
        tables: sheet.pivotTables, date1904: runtime.date1904, retainedBytes: runtime.retainedBytes, linkedCharts,
        pivots: doc.sheets.flatMap(candidate => (candidate.pivotTables ?? []).map(table => ({ sheet: candidate.name, name: String(table.name ?? '') }))),
        metadataBytes, runtimeKey: runtime,
        cache: def => runtime.cache(def),
        records: (cache, readSignal, limits = RECORD_LIMITS[profile]) => runtime.records(cache, readSignal, limits),
        saved: async () => sheetCellReader(sheet),
        source: async cache => {
          const { worksheetSourceRange } = await import('../pivots/session/refresh')
          const range = worksheetSourceRange(cache)
          const sourceSheet = range ? doc.sheets.find(candidate => candidate.name === range.sheet) : undefined
          return sourceSheet ? sheetCellReader(sourceSheet) : undefined
        },
        occupied: (row, col) => cells.has(row * MAX_COLS + col),
        blocked: (row, col) => structure.occupied(row, col),
        store: storeFor,
      }
    },
    store: storeFor,
    displayText: (cell, sheetId, allocation) => storeFor(sheetId).displayText(cell, allocation),
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
    // Compatibility-mode documents served by the worker style remote cells
    // through this table (spec §11.1), exactly as paged imports do.
    ...(doc.cellStyles ? { formatting: () => {
      const colors = doc.sheets.find(sheet => sheet.tableStyleSource)?.tableStyleSource?.colors
        ?? (doc.sheets[0]?.drawingTheme?.palette ? { theme: doc.sheets[0].drawingTheme.palette } : undefined)
      return { dateSystem: doc.semantics?.dateSystem ?? '1900', styles: doc.cellStyles!, ...(colors && (colors.indexed?.length || Object.keys(colors.theme ?? {}).length) ? { colors } : {}) }
    } } : {}),
    async release() {
      await Promise.all([...stores.values()].map(store => store.release()))
      stores.clear()
      geometries.clear()
    },
  }
}

/** The overlay over an incomplete (loading or failed) column page: answered entries take the
 * overlay, unknown and error runs pass through. A blank run that crosses an overlay column expands
 * into the owned cells, so the rewrite is cut at the request's limit with a continuation cursor
 * (queryColumn is bounded by its limit); a failed page is cut without one. */
export function overlayIncompleteColumnPage(page: Page<ColumnEntry>, request: ColumnQuery, intervals: readonly SourceRowInterval[], startAt: number, prefix: string, overlay: ValueOverlay,
  cellAt: (row: number) => ReturnType<ValueOverlay['cell']>, payload: (row: number, value: ValueOverlayCell) => SingleCellPayload): Page<ColumnEntry> {
  const output: ColumnEntry[] = []
  let resume: number | undefined
  const emit = (entry: ColumnEntry, from: number): boolean => {
    if (output.length >= request.limit) { resume = from; return false }
    output.push(entry); return true
  }
  items: for (const item of page.items) {
    if (item.kind === 'unknown' || item.kind === 'error') { if (!emit(item, item.r0)) break; continue }
    if (item.kind === 'cell' || item.kind === 'unavailable') {
      const value = cellAt(item.row)
      const entry: ColumnEntry = value === undefined ? item : value === null ? { kind: 'blank', r0: item.row, r1: item.row } : { kind: 'cell', row: item.row, cell: payload(item.row, value) }
      if (!emit(entry, item.row)) break
      continue
    }
    if (item.kind !== 'blank') { if (!emit(item, intervals[startAt]?.r0 ?? 0)) break; continue }
    const replacements: Array<{ row: number; value: ValueOverlayCell }> = []
    for (const rect of overlay.rects) if (request.col >= rect.c0 && request.col <= rect.c1) {
      for (let row = Math.max(item.r0, rect.r0); row <= Math.min(item.r1, rect.r1); row++) {
        const value = overlay.cell(row, request.col)
        if (value !== undefined && value !== null) replacements.push({ row, value })
      }
    }
    replacements.sort((a, b) => a.row - b.row)
    let at = item.r0
    for (const replacement of replacements) {
      if (replacement.row < at) continue
      if (replacement.row > at && !emit({ kind: 'blank', r0: at, r1: replacement.row - 1 }, at)) break items
      if (!emit({ kind: 'cell', row: replacement.row, cell: payload(replacement.row, replacement.value) }, replacement.row)) break items
      at = replacement.row + 1
    }
    if (at <= item.r1 && !emit({ kind: 'blank', r0: at, r1: item.r1 }, at)) break
  }
  if (resume === undefined) return { ...page, items: output }
  const { nextCursor: _cut, ...rest } = page
  let at = startAt
  while (at < intervals.length && intervals[at].r1 < resume) at++
  if (page.error || at >= intervals.length) return { ...rest, items: output }
  return { ...rest, items: output, nextCursor: `${prefix}:${at}:${Math.max(resume, intervals[at].r0)}` }
}

/** The overlay contract (ValueOverlay.protectedRects): every non-null owned cell lies inside the protected
 * bounds, which also bound the extent (and so Find and copy). A producer that breaks it is refused before
 * anything is published. Owned rectangles inside one protected rectangle are accepted without a scan; only
 * the parts outside (vacated cells, which must be null) are read. */
export function assertOverlayCellsProtected(overlay: ValueOverlay): void {
  // Indexed (Task 33 perf): a rectangle's containment and a cell's cover are point queries, not scans of every bound.
  const bounds = new RectCover(overlay.protectedRects)
  for (const rect of overlay.rects) {
    if (bounds.contains(rect)) continue
    for (let row = rect.r0; row <= rect.r1; row++) for (let col = rect.c0; col <= rect.c1; col++) {
      if (bounds.has(row, col)) continue
      const value = overlay.cell(row, col)
      if (value !== undefined && value !== null) throw new TypeError(`ValueOverlay cell R${row}C${col} (${formatA1({ row, col })}) has a value outside protectedRects`)
    }
  }
}
