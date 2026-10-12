/**
 * Resident document session (spec §§4.1, 6.3, 10.1–10.2; plan S4.1).
 *
 * One session owns one opened workbook: its SourceStore per sheet, the
 * per-sheet committed views, prepared columns and structural handles. Every
 * command after open reads this resident state; nothing reparses or clones the
 * workbook. Sort/filter build a private draft, pass the shared structural
 * guards and commit order/visibility/counts/revisions in one synchronous step,
 * so a refusal or failure leaves every generation untouched.
 *
 * S4.2 scheduling: every long job (preparation, populated guard scans, sort
 * kernels, visibility projection) yields on the shared CooperativeSlice budget
 * and rechecks cancellation and its owner after every yield. Drafts are built
 * privately, charged to a transient ledger while they live, and published only
 * through commitDraft, which validates the captured base owner. Progress is
 * reported from completed authoritative batches; the kernel decides whether it
 * may still be published.
 *
 * S4.3 lifecycle: one document-wide MemoryPool (§12.4 worker budget) charges
 * every retained view, prepared column and structural handle of every sheet
 * plus every transient draft/guard/metrics buffer, so the budget is aggregate.
 * Inactive sheet views are evictable: the compact logical state (specs,
 * selection, revision, counts and the sort recipe) stays, and the order/mask
 * buffers are rebuilt under the SAME identity on reactivation. A sort pins
 * exactly the rows hidden when it ran, and visibility is a source-row predicate,
 * so {sort spec, filter+matcher at sort time} reproduces the order; a digest by
 * source identity verifies it. Owner-free drafts (commitView with an arbitrary
 * order) have no recipe and are never evicted. Every retained view is admitted
 * under the hard limit (evicting first, else ViewQuotaError/quota with nothing
 * changed); untouched sheets hold no mask until a view or operation needs one.
 * A value publication drops an applied sort's buffers on resident and evicted
 * views alike; the next ensureView re-derives it from the new values and
 * reports `view-reapplied`. A released session never revives: no lazy state,
 * activation, admission or publication after release.
 */
import type { DeclaredSheetFilters } from '../xlsx/view/declaredFilters'
import type {
  AxisMetrics, AxisSizeInterval, CacheIdentity, ColumnEntry, ColumnQuery, CompactSourceRowSet, FilterSpec,
  AuthoritativeCell, GridSelection, GridSelectionPart, IntervalSet, OperationRefusal, Page, ProfileId, RowOrder, SingleCellPayload, SortSpec,
  SheetLayout, SourceFailure, WorkbookFormatting, SourceReadQuery, SourceReadResult, SourceRowInterval, SourceStamp, SourceStore, ValuePolicy, WorkbookOwner, WorkerDiagnostic,
} from '../xlsx/view/contracts'
import { MAX_COLS, MAX_ROWS, cellId, isValidCoordinate, type SourceRect, type SourceRow } from '../xlsx/view/coordinates'
import { createIdentityOrder } from '../xlsx/view/order'
import {
  normalizeFilter, projectVisibility, projectVisibilitySteps, visibilityFromHidden, type CompactRowVisibility,
} from '../xlsx/view/visibility'
import { createSheetMetrics } from '../xlsx/view/projection'
import { hiddenSourceRowsFromDisplay, sortRowOrder, type SortResult } from '../xlsx/view/sort'
import { CooperativeSlice, type CooperativeOptions, type YieldFunction } from '../xlsx/view/scheduler'
import { StructuralIndex, type PopulatedCellSource } from '../xlsx/view/structure'
import { guardFilter, guardSort, type GuardOutcome } from '../xlsx/view/operationGuard'
import { MemoryPool, WORKER_BUDGETS, type PoolEntry } from '../xlsx/view/tileCache'
import { INITIAL_GENERATION as INITIAL_VALUE_GEN, isUint32Array, validateBitset } from './viewerProtocol'
import { filterValueDomain, scanValueDomain, type ValueStatus } from '../xlsx/view/values'

const FULL_SHEET: SourceRect = Object.freeze({ r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 })
const PREPARE_PAGE_LIMIT = 4096
/** S5.4 read bounds: populated cells per find page, open scan continuations per document. */
export const MAX_SCAN_PAGE = 4096
export const MAX_OPEN_SCANS = 4
interface OpenScan {
  key: string
  sheetId: string
  valueGen: number
  iterator: AsyncIterator<Page<AuthoritativeCell>>
  controller: AbortController
  /** One page read ahead, so the last page carries no continuation. */
  ahead?: IteratorResult<Page<AuthoritativeCell>>
}

export const sameStamp = (a: SourceStamp, b: SourceStamp): boolean =>
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
export const failure = (code: SourceFailure['code'], message: string, retryable = false): SourceFailure =>
  Object.freeze({ code, retryable, message })
/** §4.1 column page ordering violation; the session never stages such lanes. */
const UNORDERED_PAGE = failure('page-error', 'Column pages must cover requested rows in ascending, contiguous order')
export function failedPage<T>(stamp: SourceStamp, error: SourceFailure): Page<T> {
  return { items: [], completeness: 'incomplete', stamp, error }
}
export function lowerBound(sorted: ArrayLike<number>, value: number, length = sorted.length): number {
  let lo = 0, hi = length
  while (lo < hi) {
    const mid = Math.floor((lo + hi) / 2)
    if (sorted[mid] < value) lo = mid + 1
    else hi = mid
  }
  return lo
}

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
/** Source-space dimensions and authored hidden rows; never per-row objects. */
export interface SheetGeometry {
  rows: AxisMetrics
  cols: AxisMetrics
  rowOverrides: readonly AxisSizeInterval[]
  colOverrides: readonly AxisSizeInterval[]
  baseHidden: IntervalSet
  /** Authored-hidden (collapsed) column intervals, when the source reports them. */
  hiddenCols?: readonly { start: number; end: number }[]
}
/** Resident workbook source boundary. S1 swaps the implementation (paged
 * adapters, blob/page handles) without changing the session. */
export interface WorkbookSource {
  readonly valuePolicy: ValuePolicy
  readonly sheets: readonly { sheetId: string; name: string }[]
  store(sheetId: string): SourceStore
  geometry(sheetId: string): SheetGeometry
  /** Coverage this backend can never resolve (e.g. the interim adapter before
   * S3.1 wiring): attached to incomplete replies so the UI can say why. */
  coverageDiagnostics?(sheetId: string): readonly WorkerDiagnostic[]
  /** S4.3: decoded source pages register in the session pool's `source` tier. */
  bindMemory?(pool: MemoryPool): void
  /** Notices attached to the open reply (e.g. compatibility mode, saved values). */
  readonly openDiagnostics?: readonly WorkerDiagnostic[]
  /** Value policy status (§4.4): the saved-values message exists only for
   * cached documents; a paged import learns its formula count as it runs. */
  valueStatus?(): ValueStatus
  /** Ownership-aware sort accounting (§12.4): true only when the sheet's raw
   * text cells stay retained AND charged to this session's pool (pinned) for
   * the session's lifetime, so a sort need not charge a second copy. */
  retainsKeyStrings?(sheetId: string): boolean
  /** S8.3 declared file filter/sort diagnostics; both adapters build identical records. */
  declaredFilters?(): readonly DeclaredSheetFilters[]
  /** S5.4: style table and date system of a source whose cells carry style ids only. */
  formatting?(): WorkbookFormatting
  release(): Promise<void>
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
  /** The operation's shared yield budget (S4.2); S8 scans step it per row. */
  cooperate?: CooperativeSlice<string>
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
/** Fixed per-column charge (descriptor, dictionary map, extras map). */
export const PREPARED_COLUMN_OVERHEAD = 256
/** Exact lane bytes per populated entry: row Uint32 + tag Uint8 + number Float64
 * + string id Uint32 + style Uint32 + flags Uint8. */
export const PREPARED_ENTRY_LANE_BYTES = 22
/** Charge per unique dictionary string beyond its UTF-16 payload. */
export const PREPARED_STRING_OVERHEAD = 48
/** Charge per sparse extras record (formula metadata) beyond its strings. */
export const PREPARED_EXTRA_OVERHEAD = 96

const TAG_NULL = 0, TAG_NUMBER = 1, TAG_STRING = 2, TAG_BOOLEAN = 3
const F_CACHED_SET = 1, F_CACHED = 2, F_ERROR_SET = 4, F_ERROR = 8, F_PROV_ENGINE = 16, F_PROV_CACHED = 32, F_UNAVAILABLE = 64
type Extras = Pick<SingleCellPayload, 'formula' | 'sharedFormula' | 'arrayRef' | 'dynamicArray' | 'formulaType' | 'ca'>
type PopulatedEntry = Extract<ColumnEntry, { kind: 'cell' | 'unavailable' }>
/** Compact resident column: typed lanes plus a string dictionary and sparse
 * formula extras. No per-row objects are retained; pages rebuild payloads. */
interface PreparedColumn {
  stamp: SourceStamp
  col: number
  count: number
  rows: Uint32Array
  tags: Uint8Array
  numbers: Float64Array
  stringIds: Uint32Array
  styles: Uint32Array
  flags: Uint8Array
  strings: readonly string[]
  extras: ReadonlyMap<number, Extras>
  laneBytes: number
  stringBytes: number
  extraBytes: number
  bytes: number
}
export interface PreparedFootprint {
  entries: number
  laneBytes: number
  stringBytes: number
  uniqueStrings: number
  extraObjects: number
  extraBytes: number
  bytes: number
}
const text = (value: unknown) => typeof value === 'string' ? value.length * 2 : 0
/** Growable lanes for one column under construction; every growth is charged. */
class ColumnBuilder {
  capacity = 0
  count = 0
  rows = new Uint32Array(0)
  tags = new Uint8Array(0)
  numbers = new Float64Array(0)
  stringIds = new Uint32Array(0)
  styles = new Uint32Array(0)
  flags = new Uint8Array(0)
  readonly strings: string[] = []
  private readonly dictionary = new Map<string, number>()
  readonly extras = new Map<number, Extras>()
  stringBytes = 0
  extraBytes = 0
  constructor(readonly col: number) {}
  /** Bytes currently held, including spare capacity. */
  get heldBytes(): number {
    return PREPARED_COLUMN_OVERHEAD + this.capacity * PREPARED_ENTRY_LANE_BYTES + this.stringBytes + this.extraBytes
  }
  /** Peak bytes above heldBytes if the entry were appended now. A growth step
   * allocates the new lanes (max(16, 2C) slots) while the old C slots are still
   * held for the copy, so the whole new allocation is the transient peak. */
  costOf(entry: PopulatedEntry): number {
    let cost = this.count === this.capacity ? Math.max(16, this.capacity * 2) * PREPARED_ENTRY_LANE_BYTES : 0
    const value = entry.cell.value
    if (typeof value === 'string' && !this.dictionary.has(value)) cost += value.length * 2 + PREPARED_STRING_OVERHEAD
    const cell = entry.cell
    if (cell.formula !== undefined || cell.sharedFormula !== undefined || cell.arrayRef !== undefined
      || cell.dynamicArray !== undefined || cell.formulaType !== undefined || cell.ca !== undefined) {
      cost += PREPARED_EXTRA_OVERHEAD + text(cell.formula) + text(cell.arrayRef) + text(cell.formulaType) + text(cell.sharedFormula?.ref)
    }
    return cost
  }
  private grow(): void {
    const capacity = Math.max(16, this.capacity * 2)
    const resize = <T extends Uint8Array | Uint32Array | Float64Array>(old: T, make: (n: number) => T): T => {
      const next = make(capacity)
      next.set(old)
      return next
    }
    this.rows = resize(this.rows, n => new Uint32Array(n))
    this.tags = resize(this.tags, n => new Uint8Array(n))
    this.numbers = resize(this.numbers, n => new Float64Array(n))
    this.stringIds = resize(this.stringIds, n => new Uint32Array(n))
    this.styles = resize(this.styles, n => new Uint32Array(n))
    this.flags = resize(this.flags, n => new Uint8Array(n))
    this.capacity = capacity
  }
  append(entry: PopulatedEntry): void {
    if (this.count === this.capacity) this.grow()
    const k = this.count++, cell = entry.cell, value = cell.value
    this.rows[k] = entry.row
    this.styles[k] = cell.styleIndex
    if (value === null) this.tags[k] = TAG_NULL
    else if (typeof value === 'number') { this.tags[k] = TAG_NUMBER; this.numbers[k] = value }
    else if (typeof value === 'boolean') { this.tags[k] = TAG_BOOLEAN; this.numbers[k] = value ? 1 : 0 }
    else {
      this.tags[k] = TAG_STRING
      let id = this.dictionary.get(value)
      if (id === undefined) {
        id = this.strings.length
        this.strings.push(value)
        this.dictionary.set(value, id)
        this.stringBytes += value.length * 2 + PREPARED_STRING_OVERHEAD
      }
      this.stringIds[k] = id
    }
    let flags = entry.kind === 'unavailable' ? F_UNAVAILABLE : 0
    if (cell.hasCachedValue !== undefined) flags |= F_CACHED_SET | (cell.hasCachedValue ? F_CACHED : 0)
    if (cell.valueIsError !== undefined) flags |= F_ERROR_SET | (cell.valueIsError ? F_ERROR : 0)
    if (cell.valueProvenance === 'engine') flags |= F_PROV_ENGINE
    else if (cell.valueProvenance === 'cached') flags |= F_PROV_CACHED
    this.flags[k] = flags
    if (cell.formula !== undefined || cell.sharedFormula !== undefined || cell.arrayRef !== undefined
      || cell.dynamicArray !== undefined || cell.formulaType !== undefined || cell.ca !== undefined) {
      const extra: Extras = {}
      if (cell.formula !== undefined) extra.formula = cell.formula
      if (cell.sharedFormula !== undefined) extra.sharedFormula = { ...cell.sharedFormula }
      if (cell.arrayRef !== undefined) extra.arrayRef = cell.arrayRef
      if (cell.dynamicArray !== undefined) extra.dynamicArray = cell.dynamicArray
      if (cell.formulaType !== undefined) extra.formulaType = cell.formulaType
      if (cell.ca !== undefined) extra.ca = cell.ca
      this.extras.set(k, Object.freeze(extra))
      this.extraBytes += PREPARED_EXTRA_OVERHEAD + text(cell.formula) + text(cell.arrayRef) + text(cell.formulaType) + text(cell.sharedFormula?.ref)
    }
  }
  /** Exact-length lanes (one transient copy when capacity exceeds the count). */
  finish(stamp: SourceStamp): PreparedColumn {
    const n = this.count
    const exact = <T extends Uint8Array | Uint32Array | Float64Array>(lane: T): T => (lane.length === n ? lane : lane.slice(0, n)) as T
    const laneBytes = n * PREPARED_ENTRY_LANE_BYTES
    return Object.freeze({
      stamp, col: this.col, count: n,
      rows: exact(this.rows), tags: exact(this.tags), numbers: exact(this.numbers),
      stringIds: exact(this.stringIds), styles: exact(this.styles), flags: exact(this.flags),
      strings: Object.freeze(this.strings), extras: this.extras,
      laneBytes, stringBytes: this.stringBytes, extraBytes: this.extraBytes,
      bytes: PREPARED_COLUMN_OVERHEAD + laneBytes + this.stringBytes + this.extraBytes,
    })
  }
}
/** Rebuilds one bounded page entry from the lanes. */
function preparedEntry(column: PreparedColumn, k: number): PopulatedEntry {
  const row = column.rows[k], tag = column.tags[k], flags = column.flags[k]
  const value = tag === TAG_NULL ? null : tag === TAG_NUMBER ? column.numbers[k]
    : tag === TAG_BOOLEAN ? column.numbers[k] !== 0 : column.strings[column.stringIds[k]]
  const cell: SingleCellPayload = { id: cellId(row, column.col), row, col: column.col, value, styleIndex: column.styles[k] }
  if (flags & F_CACHED_SET) cell.hasCachedValue = (flags & F_CACHED) !== 0
  if (flags & F_ERROR_SET) cell.valueIsError = (flags & F_ERROR) !== 0
  if (flags & F_PROV_ENGINE) cell.valueProvenance = 'engine'
  else if (flags & F_PROV_CACHED) cell.valueProvenance = 'cached'
  const extra = column.extras.get(k)
  if (extra) Object.assign(cell, extra, extra.sharedFormula ? { sharedFormula: { ...extra.sharedFormula } } : {})
  Object.freeze(cell)
  return flags & F_UNAVAILABLE
    ? { kind: 'unavailable', row, cell, reason: 'formula-value-unavailable' }
    : { kind: 'cell', row, cell }
}
/** Membership remaining after a position (interval index + next row). */
function residualIntervals(intervals: readonly SourceRowInterval[], at: number, row: number): SourceRowInterval[] {
  if (at >= intervals.length) return []
  return [{ r0: Math.max(row, intervals[at].r0), r1: intervals[at].r1 }, ...intervals.slice(at + 1)]
}
/** FNV-1a over the normalized query; ties a session cursor to its exact query and
 * source owner. (Bitwise arithmetic here hashes text, never a cell ID.) */
function queryFingerprint(col: number, stamp: SourceStamp, intervals: readonly SourceRowInterval[]): string {
  const textKey = JSON.stringify([col, stamp.docId, stamp.sheetId, stamp.valueGen, intervals.map(i => [i.r0, i.r1])])
  let hash = 0x811c9dc5
  for (let i = 0; i < textKey.length; i++) {
    hash ^= textKey.charCodeAt(i)
    hash = Math.imul(hash, 0x01000193) >>> 0
  }
  return `${hash.toString(36)}.${intervals.length}`
}

export type PrepareOutcome =
  | { kind: 'complete' }
  | { kind: 'stale' }
  | { kind: 'cancelled' }
  | { kind: 'incomplete'; missing: readonly SourceRect[] }
  | { kind: 'error'; error: SourceFailure }
  | { kind: 'quota'; limitBytes: number; requiredBytes: number; diagnostics?: readonly WorkerDiagnostic[] }
  | { kind: 'resource-limit'; message: string }
/** Yield phases of the session's own kernels; sort phases come from sort.ts. */
export type SessionPhase = 'prepare' | 'guard' | 'project'
export interface SessionOptions {
  /** Injected host yield shared by every long kernel (S4.2 scheduler). */
  yield?: YieldFunction<string>
  /** Injected clock for the 10 ms yield budget (deterministic tests). */
  now?: () => number
  /** Kernel-assigned document epoch: one per open, never reused. */
  epoch?: number
  locale?: string
  filterEvaluator?: FilterEvaluator
  /** Admission profile from the open command; selects PREPARED_BUDGETS. */
  profile?: ProfileId
  /** Explicit override of the profile budget (tests, embedders). */
  preparedBudget?: PreparedBudget
  /** Explicit document-wide worker budget (tests, embedders). */
  workerBudgetBytes?: number
}
/** Scanned value-domain entries per shared-budget charge (filter precheck). */
const DOMAIN_SCAN_CHARGE = 256
/** Visible display slots and the logical row axis, committed with the view. */
export interface ViewCounts { readonly visible: number; readonly total: number }
/** Private draft; commitDraft publishes it atomically. */
export interface ViewDraft {
  order: RowOrder
  visibility: CompactRowVisibility
  sort: SortSpec | null
  filter: FilterSpec | null
  matches?: (row: SourceRow, filter: FilterSpec) => boolean
  selection?: GridSelection
  /** Derived from visibility when omitted. */
  counts?: ViewCounts
}
export interface CommittedView extends Required<Omit<ViewDraft, 'matches'>> {
  identity: CacheIdentity
  workbookOwner: WorkbookOwner
  geometry: SheetGeometry
}
/** Completed authoritative work, reported by long operations. The kernel decides
 * whether it may still be published (owner, epoch and cancellation checks). */
export interface OperationHooks {
  progress?(completed: number, total: number): void
}
/** Sort/filter report three completed stages each (guard or rule evaluation,
 * kernel or projection, projection or geometry guard) before committing. */
export const VIEW_OPERATION_STAGES = 3
export interface SessionMemoryStats {
  /** Bytes charged by in-flight drafts and staged preparation right now. */
  transientBytes: number
  peakTransientBytes: number
  preparedBytes: number
  viewBytes: number
  structureBytes: number
  sourceBytes: number
  retainedBytes: number
  totalBytes: number
  peakBytes: number
  limitBytes: number
  evictions: number
}
/** Outcome of materializing a sheet view (S4.3). `diagnostics` reports a
 * re-applied sort when the view was evicted across a value publication. */
export type EnsureViewOutcome =
  | { kind: 'resident'; view: SheetViewSnapshot; diagnostics?: readonly WorkerDiagnostic[] }
  | { kind: 'stale' }
  | { kind: 'cancelled' }
  | { kind: 'incomplete'; missing: readonly SourceRect[] }
  | { kind: 'error'; error: SourceFailure }
  | { kind: 'quota'; requiredBytes: number; limitBytes: number }
/** Thrown by the synchronous view() when an evicted order needs an async
 * rebuild: await ensureView() first. */
export class ViewNotResidentError extends RangeError {
  constructor(readonly sheetId: string) {
    super(`Sheet view ${sheetId} is not resident; await ensureView`)
    this.name = 'ViewNotResidentError'
  }
}
/** A view's buffers could not be admitted under the document budget (S4.3);
 * thrown by owner-free commits and the synchronous view(). Nothing changed. */
export class ViewQuotaError extends RangeError {
  constructor(readonly requiredBytes: number, readonly limitBytes: number) {
    super(`Sheet view needs ${requiredBytes} bytes beyond the ${limitBytes}-byte worker budget`)
    this.name = 'ViewQuotaError'
  }
}
/** Smallest admissible worker budget: one full-axis sheet mask. */
export const MIN_WORKER_BUDGET_BYTES = Math.ceil(MAX_ROWS / 32) * 4 + (Math.ceil(MAX_ROWS / 32 / 32) + 1) * 4
/** A configured worker budget that could never hold one sheet view. */
export class WorkerBudgetError extends RangeError {
  constructor(readonly budgetBytes: number) {
    super(`Worker budget ${budgetBytes} B is below one sheet mask (${MIN_WORKER_BUDGET_BYTES} B)`)
    this.name = 'WorkerBudgetError'
  }
}
/** Per body row of a sort: index + scratch + order + inverse Uint32 lanes. */
export const SORT_ROW_ESTIMATE_BYTES = 16
/** Per body row and key: tag Uint8 + number Float64 + string id Uint32. */
export const SORT_KEY_ESTIMATE_BYTES = 13
/** Structural handle charge: fixed part plus one record estimate per entry. */
export const STRUCTURE_BASE_BYTES = 1024
export const STRUCTURE_RECORD_BYTES = 256
export type OperationOutcome =
  | { kind: 'committed'; view: CommittedView }
  | { kind: 'refused'; refusal: OperationRefusal }
  | { kind: 'stale' }
  | { kind: 'cancelled' }
  | { kind: 'incomplete'; missing: readonly SourceRect[]; diagnostics?: readonly WorkerDiagnostic[] }
  | { kind: 'error'; error: SourceFailure }
  | { kind: 'quota'; requiredBytes: number; limitBytes: number }
/** Read-only committed view snapshot. */
export interface SheetViewSnapshot {
  readonly sheetId: string
  readonly revision: number
  readonly valueGen: number
  readonly order: RowOrder
  readonly visibility: CompactRowVisibility
  readonly sort: SortSpec | null
  readonly filter: FilterSpec | null
  readonly selection: GridSelection
  readonly identity: CacheIdentity
  readonly counts: ViewCounts
}

/** How a committed order is reproduced after eviction: the filter (and its
 * captured matcher, never re-evaluated) that was active when the sort ran. */
interface ViewRecipe {
  readonly sortFilter: FilterSpec | null
  readonly sortMatches?: (row: SourceRow, filter: FilterSpec) => boolean
}
/** Immutable committed sheet record: a commit swaps the whole record, so no
 * observer can ever read a mix of old and new order/mask/counts/revision.
 * Eviction swaps in a compact record (order/visibility null) that keeps the
 * logical state; rebuilding swaps the buffers back without a new revision. */
interface SheetState {
  readonly sheetId: string
  readonly revision: number
  readonly valueGen: number
  /** null: evicted (identity orders are free and never evicted). */
  readonly order: RowOrder | null
  /** null: evicted, rebuilt from order + filter + matcher. */
  readonly visibility: CompactRowVisibility | null
  readonly counts: ViewCounts
  readonly sort: SortSpec | null
  readonly filter: FilterSpec | null
  readonly matches?: (row: SourceRow, filter: FilterSpec) => boolean
  readonly selection: GridSelection
  /** null: not reproducible (owner-free drafts); such views are never evicted. */
  readonly recipe: ViewRecipe | null
  /** Value generation the order was computed under. */
  readonly orderValueGen: number
  /** Digest by source identity, taken when an order is evicted. */
  readonly digest?: string
}
interface ResidentState extends SheetState {
  readonly order: RowOrder
  readonly visibility: CompactRowVisibility
}
const isResidentState = (state: SheetState): state is ResidentState => state.order !== null && state.visibility !== null
type Halt = Exclude<EnsureViewOutcome, { kind: 'resident' }>
const viewKey = (sheetId: string) => JSON.stringify(['view', sheetId])
const structureKey = (sheetId: string) => JSON.stringify(['structure', sheetId])
function orderBytes(order: RowOrder): number {
  let bytes = 0
  for (const entry of order.bands) bytes += entry.sourceAtOffset.byteLength + entry.displayOffsetOfSource.byteLength
  return bytes
}
const maskBytes = (visibility: CompactRowVisibility): number => visibility.words.byteLength + visibility.summaries.byteLength
/** FNV-1a over band bounds and source rows: the order by source identity. */
function orderDigest(order: RowOrder): string {
  let hash = 0x811c9dc5
  const mix = (value: number) => { hash = Math.imul(hash ^ value, 0x01000193) >>> 0 }
  for (const entry of order.bands) {
    mix(entry.r0); mix(entry.r1)
    for (let offset = 0; offset < entry.sourceAtOffset.length; offset++) mix(entry.sourceAtOffset[offset])
  }
  return `${order.bands.length}:${hash.toString(36)}`
}
/** Upper bound of createSheetMetrics' dense row sections for an order: per
 * band Float64 sizes + Float64 prefix + Uint32 positive counts. */
function metricsUpperBound(order: RowOrder): number {
  let bytes = 0
  for (const entry of order.bands) bytes += 20 * (entry.r1 - entry.r0 + 1) + 12
  return bytes
}
const countsOf = (visibility: CompactRowVisibility): ViewCounts =>
  Object.freeze({ visible: visibility.rank(visibility.logicalCount), total: MAX_ROWS })
/** Abandons a cooperative guard at a yield when its request was cancelled or
 * its owner superseded; caught by the operation and returned as its outcome. */
class GuardHalt extends Error {
  constructor(readonly outcome: { kind: 'cancelled' } | { kind: 'stale' }) { super(outcome.kind) }
}
/** One operation's share of the document pool's transient ledger; `set` is
 * absolute and false when the bytes cannot be admitted. */
interface Charge { set(bytes: number): boolean; release(): void }
/** Abandons a guard whose scratch buffer could not be admitted. */
class QuotaHalt extends Error {
  constructor(readonly requiredBytes: number) { super('quota') }
}
const EMPTY_SELECTION: GridSelection = Object.freeze({ active: null, anchor: null, ranges: Object.freeze([]) })
const matchAll = () => true
/** Deep-freezes plain selection data; typed arrays cannot be frozen and are
 * owned by the session once committed (deltas carry accounted copies). */
function freezeSelection(selection: GridSelection): GridSelection {
  const freeze = (value: unknown): unknown => {
    if (typeof value !== 'object' || value === null || ArrayBuffer.isView(value) || Object.isFrozen(value)) return value
    for (const entry of Object.values(value)) freeze(entry)
    return Object.freeze(value)
  }
  const parts = (ranges: readonly GridSelectionPart[]) => ranges.map(range => ({ rows: { ...range.rows }, cols: range.cols.map(col => ({ ...col })) }))
  return freeze({
    active: selection.active ? { ...selection.active } : null,
    anchor: selection.anchor ? { ...selection.anchor } : null,
    ranges: parts(selection.ranges),
    ...(selection.gesture ? { gesture: { ...selection.gesture, anchor: { ...selection.gesture.anchor }, prefix: parts(selection.gesture.prefix), ...(selection.gesture.part ? { part: parts([selection.gesture.part])[0] } : {}) } } : {}),
  }) as GridSelection
}
/** Guard refusals that are really waits or source failures are not refusals. */
function guardOutcome(guard: Extract<GuardOutcome, { ok: false }>, diagnostics: readonly WorkerDiagnostic[]): OperationOutcome {
  const refusal = guard.refusal
  if (refusal.reason === 'incomplete') return { kind: 'incomplete', missing: refusal.missing, diagnostics }
  if (refusal.reason === 'source-error') {
    return { kind: 'error', error: failure(refusal.code, refusal.message, refusal.retryable) }
  }
  return { kind: 'refused', refusal }
}

export class DocumentSession {
  private globalGen = 0
  private workbookValueGen = INITIAL_VALUE_GEN
  private fontGen = 0
  private peakTransient = 0
  private readonly states = new Map<string, SheetState>()
  /** Insertion order is recency order (oldest first). */
  private readonly prepared = new Map<string, PreparedColumn>()
  private readonly preparedEntries = new Map<string, PoolEntry>()
  private preparedBytes = 0
  private structureBytes = 0
  private readonly structureEntries = new Map<string, PoolEntry>()
  /** The single document-wide worker budget (§12.4): every sheet, operation and
   * retained cache of this document charges it. */
  readonly memory: MemoryPool
  /** Preparation transactions run one at a time (admission/eviction are atomic). */
  private prepareQueue: Promise<unknown> = Promise.resolve()
  /** One structural handle per sheet, valid only for its value generation; a
   * newer generation replaces (evicts) the stale entry, so this stays bounded
   * by the sheet count. */
  private readonly structures = new Map<string, StructuralIndex>()
  private published: WorkbookOwner | null = null
  private closed = false
  /** Insertion order is recency order; bounded by MAX_OPEN_SCANS. */
  private readonly scans = new Map<string, OpenScan>()
  private scanSerial = 0
  activeSheetId: string

  constructor(readonly docId: string, readonly source: WorkbookSource, private readonly options: SessionOptions = {}) {
    if (!source.sheets.length) throw new RangeError('Workbook has no sheets')
    this.activeSheetId = source.sheets[0].sheetId
    const limitBytes = options.workerBudgetBytes ?? WORKER_BUDGETS[options.profile ?? 'D']
    // Below one full-axis mask no sheet could ever be viewed: refuse up front.
    if (!(limitBytes >= MIN_WORKER_BUDGET_BYTES)) throw new WorkerBudgetError(limitBytes)
    this.memory = new MemoryPool({ limitBytes })
    this.memory.protect(this.activeSheetId)
    source.bindMemory?.(this.memory)
  }
  get globalViewGen(): number { return this.globalGen }
  /** Document epoch assigned by the kernel at open (1 when standalone). */
  get epoch(): number { return this.options.epoch ?? 1 }
  private get budget(): PreparedBudget {
    return this.options.preparedBudget ?? PREPARED_BUDGETS[this.options.profile ?? 'D']
  }
  get isClosed(): boolean { return this.closed }
  get publishedWorkbookIndex(): WorkbookOwner | null { return this.published }
  hasSheet(sheetId: string): boolean { return this.source.sheets.some(sheet => sheet.sheetId === sheetId) }
  /** Lazily materialized: an untouched sheet keeps no mask or order state. A
   * released session never re-creates state. */
  private state(sheetId: string): SheetState {
    if (this.closed) throw new RangeError('Session closed')
    let state = this.states.get(sheetId)
    if (!state) {
      if (!this.hasSheet(sheetId)) throw new RangeError(`Unknown sheet ${sheetId}`)
      // Compact until a view or operation needs the mask: the identity order is
      // free and the counts follow from the authored hidden intervals.
      let hidden = 0
      for (const interval of this.source.geometry(sheetId).baseHidden.intervals) hidden += interval.r1 - interval.r0 + 1
      state = Object.freeze({
        sheetId, revision: 0, valueGen: INITIAL_VALUE_GEN, order: createIdentityOrder(), visibility: null,
        counts: Object.freeze({ visible: MAX_ROWS - hidden, total: MAX_ROWS }), sort: null, filter: null, selection: EMPTY_SELECTION,
        recipe: Object.freeze({ sortFilter: null }), orderValueGen: INITIAL_VALUE_GEN,
      })
      this.states.set(sheetId, state)
    }
    return state
  }
  /** Admits a resident view's buffers under the hard limit (evicting other
   * entries first); throws ViewQuotaError and changes nothing when they cannot
   * fit. Callers retain BEFORE installing the record, so a refusal is atomic.
   * Operation commits uncharge their transient draft first, so the same bytes
   * are re-admitted, never counted twice and never forced. */
  private retainView(state: ResidentState): void {
    const bytes = orderBytes(state.order) + maskBytes(state.visibility)
    const entry = this.memory.retain(viewKey(state.sheetId), {
      tier: 'view', bytes, owner: state.sheetId, evictable: state.recipe !== null,
      onRelease: reason => { if (reason === 'evicted') this.evictView(state.sheetId) },
    })
    if (!entry) throw new ViewQuotaError(bytes, this.memory.limitBytes)
  }
  /** Drops the buffers of the sheet's CURRENT record (a value publication may
   * have swapped the record since it was retained); the logical state stays. */
  private evictView(sheetId: string): void {
    const state = this.states.get(sheetId)
    if (!state || !isResidentState(state)) return
    const identity = state.order.kind === 'identity'
    this.states.set(sheetId, Object.freeze({
      ...state, visibility: null, order: identity ? state.order : null,
      // The committed order's digest is taken once and kept across rebuilds, so
      // every rebuild is checked against the order the user committed.
      digest: identity ? state.digest : state.digest ?? orderDigest(state.order),
    }))
  }
  private pinView(sheetId: string): () => void {
    return this.memory.get(viewKey(sheetId))?.pin() ?? (() => {})
  }
  /** Whether the sheet's order and mask buffers are materialized. */
  isResident(sheetId: string): boolean {
    const state = this.states.get(sheetId)
    return state !== undefined && isResidentState(state)
  }
  /** Synchronous materialization: a mask-only eviction is rebuilt in place; an
   * evicted order needs the async ensureView. */
  private residentSync(sheetId: string): ResidentState {
    const state = this.state(sheetId)
    if (isResidentState(state)) return state
    if (state.order === null) throw new ViewNotResidentError(sheetId)
    const visibility = projectVisibility(this.source.geometry(sheetId).baseHidden, state.filter, state.order, state.matches ?? matchAll).visibility
    if (visibility.rank(visibility.logicalCount) !== state.counts.visible) {
      throw new RangeError(`Sheet view ${sheetId} could not be rebuilt identically`)
    }
    const next: ResidentState = Object.freeze({ ...state, order: state.order, visibility })
    this.retainView(next)
    this.states.set(sheetId, next)
    return next
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
  /** Frozen snapshot of a sheet's committed view; never the live mutable state.
   *
   * IMPORTANT: throws ViewNotResidentError whenever the sheet's order buffers
   * are absent — after LRU eviction of an inactive sheet AND after a value
   * publication on any sheet with an applied sort (its order is re-derived from
   * the new values). Asynchronous callers must `await ensureView(sheetId)`
   * first; the kernel never calls view() (activateSheet/sort/filter rebuild
   * through ensureView/materialize). It may also throw ViewQuotaError when a
   * lazily built mask cannot be admitted. */
  view(sheetId: string): SheetViewSnapshot {
    return this.snapshot(this.residentSync(sheetId))
  }
  private snapshot(state: ResidentState): SheetViewSnapshot {
    const sheetId = state.sheetId
    return Object.freeze({
      sheetId: state.sheetId, revision: state.revision, valueGen: state.valueGen, order: state.order,
      visibility: state.visibility, sort: state.sort, filter: state.filter, selection: state.selection,
      identity: this.identity(sheetId), counts: state.counts,
    })
  }
  isCurrent(owner: CacheIdentity): boolean {
    return !this.closed && owner.docId === this.docId && this.hasSheet(owner.sheetId) && sameIdentity(owner, this.identity(owner.sheetId))
  }
  /** Selects the active sheet; its view is protected from eviction. Callers
   * that need its buffers await ensureView (the kernel does on activateSheet). */
  activate(sheetId: string): void {
    if (this.closed) throw new RangeError('Session closed')
    if (!this.hasSheet(sheetId)) throw new RangeError(`Unknown sheet ${sheetId}`)
    this.activeSheetId = sheetId
    this.memory.protect(sheetId)
  }
  /** Materializes a sheet view, rebuilding evicted buffers by recipe under the
   * same identity: no revision, generation or owner changes. */
  async ensureView(sheetId: string, signal: AbortSignal): Promise<EnsureViewOutcome> {
    if (this.closed) return { kind: 'stale' }
    if (!this.hasSheet(sheetId)) throw new RangeError(`Unknown sheet ${sheetId}`)
    const result = await this.materialize(sheetId, signal)
    if ('kind' in result) return result
    return { kind: 'resident', view: this.snapshot(result.state), ...(result.diagnostics ? { diagnostics: result.diagnostics } : {}) }
  }
  /** The rebuild: the sort recipe reproduces the order (verified by digest when
   * the values are unchanged), then the mask is projected from the committed
   * filter and matcher (verified by count). Nothing is installed on mismatch. */
  private async materialize(sheetId: string, signal: AbortSignal): Promise<{ state: ResidentState; diagnostics?: readonly WorkerDiagnostic[] } | Halt> {
    if (signal.aborted) return { kind: 'cancelled' }
    if (this.closed) return { kind: 'stale' }
    const state = this.state(sheetId)
    if (isResidentState(state)) return { state }
    if (state.order !== null) {
      try { return { state: this.residentSync(sheetId) } } catch (cause) {
        if (cause instanceof ViewQuotaError) return this.refusal(cause.requiredBytes)
        return { kind: 'error', error: failure('page-error', (cause as Error).message) }
      }
    }
    const recipe = state.recipe!
    const geometry = this.source.geometry(sheetId)
    const charge = this.charge()
    try {
      const identityOrder = createIdentityOrder()
      // The rows the sort pinned: hidden under the filter active when it ran.
      const sortMask = recipe.sortFilter
        ? projectVisibility(geometry.baseHidden, recipe.sortFilter, identityOrder, recipe.sortMatches ?? matchAll).visibility
        : visibilityFromHidden(geometry.baseHidden)
      if (!charge.set(maskBytes(sortMask))) return this.refusal(maskBytes(sortMask))
      this.memory.trim(Math.max(0, this.memory.limitBytes - this.sortEstimate(state.sort)))
      const store = this.preparedStore(sheetId)
      const result = await sortRowOrder(store, state.sort, {
        signal, stamp: store.stamp, locale: this.options.locale,
        hidden: hiddenSourceRowsFromDisplay(identityOrder, sortMask),
        residentWorkerBytes: this.memory.totalBytes, maxWorkerBytes: this.memory.limitBytes,
        ...(this.source.retainsKeyStrings?.(sheetId) === true ? { sourceRetainsStrings: true } : {}),
        ...(this.options.yield ? { yield: this.options.yield } : {}),
        ...(this.options.now ? { now: this.options.now } : {}),
      })
      if (result.kind !== 'complete') {
        const outcome = sortFailure(result, state.sort)
        return outcome.kind === 'refused' ? { kind: 'error', error: failure('page-error', 'Sort refused while rebuilding an applied view') }
          : outcome as Halt
      }
      if (signal.aborted) return { kind: 'cancelled' }
      const current = this.closed ? undefined : this.states.get(sheetId)
      if (current !== state) {
        // A concurrent rebuild may already have restored the same logical view.
        return current && isResidentState(current) && current.revision === state.revision && current.valueGen === state.valueGen
          ? { state: current } : { kind: 'stale' }
      }
      const sameValues = state.orderValueGen === state.valueGen
      if (sameValues && state.digest !== undefined && orderDigest(result.order) !== state.digest) {
        return { kind: 'error', error: failure('page-error', `Sheet view ${sheetId} rebuilt a different order`) }
      }
      if (!charge.set(maskBytes(sortMask) + result.ledger.orderBytes + result.ledger.inverseBytes)) {
        return this.refusal(result.ledger.orderBytes + result.ledger.inverseBytes)
      }
      const visibility = projectVisibility(geometry.baseHidden, state.filter, result.order, state.matches ?? matchAll).visibility
      if (visibility.rank(visibility.logicalCount) !== state.counts.visible) {
        return { kind: 'error', error: failure('page-error', `Sheet view ${sheetId} rebuilt different visible rows`) }
      }
      const next: ResidentState = Object.freeze({
        ...state, order: result.order, visibility, orderValueGen: state.valueGen,
        // Kept when verified; a re-applied sort starts a new baseline.
        digest: sameValues && state.digest !== undefined ? state.digest : orderDigest(result.order),
      })
      // The rebuilt buffers move from the transient ledger to the retained view.
      charge.set(0)
      try { this.retainView(next) } catch (cause) {
        if (cause instanceof ViewQuotaError) return this.refusal(cause.requiredBytes)
        throw cause
      }
      this.states.set(sheetId, next)
      return sameValues ? { state: next } : {
        state: next,
        diagnostics: [{ code: 'view-reapplied', message: 'The sort was re-applied to newer values after its evicted order could not be reproduced' }],
      }
    } finally {
      charge.release()
    }
  }
  /** Rough sort working set, used only to make room before the kernel runs;
   * the sort kernel's own aggregate budget stays the authority. */
  private sortEstimate(spec: SortSpec | null): number {
    if (!spec) return 0
    const rows = Math.max(0, spec.range.r1 - spec.range.totalsRows - (spec.range.r0 + spec.range.headerRows) + 1)
    return rows * (SORT_ROW_ESTIMATE_BYTES + SORT_KEY_ESTIMATE_BYTES * spec.keys.length)
  }
  /** A charge that could not be admitted: stale once closed, else quota. */
  private refusal(requiredBytes: number): { kind: 'stale' } | { kind: 'quota'; requiredBytes: number; limitBytes: number } {
    return this.closed ? { kind: 'stale' } : { kind: 'quota', requiredBytes, limitBytes: this.memory.limitBytes }
  }
  /** Unconditional publication of a draft onto the current state (owner-free
   * callers such as selection drafts); worker operations use commitDraft. Only
   * an unchanged or identity order stays reproducible (and evictable). Rejects
   * with ViewQuotaError, changing nothing, when the view cannot be admitted. */
  async commitView(sheetId: string, draft: ViewDraft): Promise<CommittedView> {
    if (this.closed) throw new RangeError('Session closed')
    const previous = this.state(sheetId)
    const recipe = draft.order === previous.order ? previous.recipe
      : draft.order.kind === 'identity' && draft.sort === null ? Object.freeze({ sortFilter: null }) : null
    return this.install(sheetId, draft, recipe, draft.order === previous.order ? previous.orderValueGen : previous.valueGen)
  }
  /** One validated view transaction (§§6, 10.1): publishes only when the session
   * is open and `base` is still the sheet's current owner, so a draft built
   * under a superseded value/font/view generation can never land. Returns null
   * and changes nothing otherwise. The recipe defaults to the previous one (a
   * draft that keeps the committed order). Throws ViewQuotaError, changing
   * nothing, when the view cannot be admitted under the document budget. */
  commitDraft(base: CacheIdentity, draft: ViewDraft, recipe?: ViewRecipe | null): CommittedView | null {
    if (!this.isCurrent(base)) return null
    const previous = this.state(base.sheetId)
    if (recipe === undefined) {
      return draft.order === previous.order ? this.install(base.sheetId, draft, previous.recipe, previous.orderValueGen)
        : this.install(base.sheetId, draft, null, previous.valueGen)
    }
    return this.install(base.sheetId, draft, recipe, previous.valueGen)
  }
  /** The single synchronous commit step: one immutable record swap plus the
   * workbook-global generation, with nothing awaited in between. */
  private install(sheetId: string, draft: ViewDraft, recipe: ViewRecipe | null, orderValueGen: number): CommittedView {
    const previous = this.state(sheetId)
    const next: ResidentState = Object.freeze({
      sheetId, revision: previous.revision + 1, valueGen: previous.valueGen,
      order: draft.order, visibility: draft.visibility, counts: draft.counts ?? countsOf(draft.visibility),
      sort: draft.sort, filter: draft.filter, matches: draft.matches,
      selection: draft.selection ? freezeSelection(draft.selection) : previous.selection,
      recipe, orderValueGen,
    })
    // Admission first (throws ViewQuotaError): a refused commit changes nothing.
    this.retainView(next)
    this.states.set(sheetId, next)
    this.globalGen++
    return Object.freeze({
      identity: this.identity(sheetId), workbookOwner: this.workbookOwner(), geometry: this.source.geometry(sheetId),
      order: next.order, visibility: next.visibility, counts: next.counts, sort: next.sort, filter: next.filter,
      selection: next.selection,
    })
  }
  /** Source/evaluation value publication for one sheet (§10.1). The store has
   * already published the new stamp; the session follows it so identity and
   * store stay in lockstep. Prepared columns and structural handles of the dead
   * stamp become unreachable and are dropped and uncharged here. */
  publishValues(sheetId: string): void {
    if (this.closed) throw new RangeError('Session closed')
    const state = this.state(sheetId)
    const stamp = this.source.store(sheetId).stamp
    if (stamp.valueGen <= state.valueGen) throw new RangeError('A value publication must advance the source stamp')
    // One rule for resident and evicted views: an applied sort is re-derived
    // from the new values on next use (ensureView reports `view-reapplied`), so
    // its buffers are dropped here. Orders without a recipe are kept as is.
    const rederive = state.order !== null && state.order.kind !== 'identity' && state.recipe !== null
    if (rederive) this.memory.get(viewKey(sheetId))?.release()
    this.states.set(sheetId, Object.freeze({
      ...state, valueGen: stamp.valueGen,
      ...(rederive || (state.order === null) ? { order: null, visibility: null, digest: undefined } : {}),
    }))
    this.workbookValueGen++
    for (const [key, column] of [...this.prepared]) {
      if (column.stamp.sheetId !== sheetId || sameStamp(column.stamp, stamp)) continue
      this.dropPrepared(key)
    }
    this.dropStructure(sheetId)
  }
  /** Font readiness (§10.1): advances fontGen only; no view revision, order or
   * mask changes, but every owner captured before it is superseded. */
  publishFonts(): void {
    if (this.closed) throw new RangeError('Session closed')
    this.fontGen++
  }
  /** Document-wide ledger (one pool across sheets and operations). */
  memoryStats(): SessionMemoryStats {
    const pool = this.memory.stats()
    return {
      transientBytes: pool.transientBytes, peakTransientBytes: this.peakTransient, preparedBytes: this.preparedBytes,
      viewBytes: pool.byTier.view, structureBytes: this.structureBytes, sourceBytes: pool.byTier.source,
      retainedBytes: pool.retainedBytes, totalBytes: pool.totalBytes, peakBytes: pool.peakBytes, limitBytes: pool.limitBytes,
      evictions: pool.evictions,
    }
  }
  /** A per-operation share of the pool's transient ledger, released in the
   * operation's finally on every outcome (commit, refusal, cancel, stale).
   * Growing a charge evicts retained caches first; false when it cannot fit. */
  private charge(): Charge {
    const share = this.memory.charge()
    return {
      set: bytes => {
        if (!share.set(bytes)) return false
        this.peakTransient = Math.max(this.peakTransient, this.memory.transientBytes)
        return true
      },
      release: () => share.release(),
    }
  }
  private slice(): CooperativeSlice<string> {
    const options: CooperativeOptions<string> = {}
    if (this.options.now) options.now = this.options.now
    if (this.options.yield) options.yield = this.options.yield
    return new CooperativeSlice<string>(options)
  }
  /** Shared post-yield recheck: cancellation first, then ownership. */
  private halted(owner: CacheIdentity, signal: AbortSignal): { kind: 'cancelled' } | { kind: 'stale' } | null {
    if (signal.aborted) return { kind: 'cancelled' }
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    return null
  }
  /** Cooperative run of the single S2.1 projection (projectVisibilitySteps):
   * yields on the shared budget per reported chunk of projected rows and
   * rechecks cancellation and ownership after every yield. */
  private async projectDraft(
    owner: CacheIdentity, filter: FilterSpec | null, order: RowOrder,
    matches: (row: SourceRow, filter: FilterSpec) => boolean,
    slice: CooperativeSlice<string>, signal: AbortSignal, charge: Charge, heldBytes: number,
  ): Promise<{ kind: 'projected'; visibility: CompactRowVisibility } | { kind: 'cancelled' } | { kind: 'stale' } | { kind: 'quota'; requiredBytes: number; limitBytes: number }> {
    const baseHidden = this.source.geometry(owner.sheetId).baseHidden
    // The draft mask words (one bit per display slot) live while projecting.
    const maskWords = Math.ceil(MAX_ROWS / 32) * 4
    if (!charge.set(heldBytes + maskWords)) return this.refusal(heldBytes + maskWords)
    const steps = projectVisibilitySteps(baseHidden, filter, order, matches, MAX_ROWS)
    let next = steps.next()
    for (; !next.done; next = steps.next()) {
      if (slice.due(next.value)) {
        await slice.pause('project')
        const halt = this.halted(owner, signal)
        if (halt) return halt
      }
    }
    const visibility = next.value.visibility
    if (!charge.set(heldBytes + maskBytes(visibility))) return this.refusal(heldBytes + maskBytes(visibility))
    return { kind: 'projected', visibility }
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
  /** Resident prepared columns of a sheet's current value generation, ascending. */
  preparedColumns(sheetId: string): number[] {
    const stamp = this.source.store(sheetId).stamp
    return [...this.prepared.values()].filter(column => sameStamp(column.stamp, stamp)).map(column => column.col).sort((a, b) => a - b)
  }
  /** Measured retention of one prepared column (lane byteLengths + charges). */
  preparedFootprint(sheetId: string, col: number): PreparedFootprint | undefined {
    const column = this.prepared.get(this.preparedKey(this.source.store(sheetId).stamp, col))
    if (!column) return undefined
    const laneBytes = column.rows.byteLength + column.tags.byteLength + column.numbers.byteLength
      + column.stringIds.byteLength + column.styles.byteLength + column.flags.byteLength
    return {
      entries: column.count, laneBytes, stringBytes: column.stringBytes, uniqueStrings: column.strings.length,
      extraObjects: column.extras.size, extraBytes: column.extraBytes, bytes: column.bytes,
    }
  }
  /** Cached structural handles (diagnostic). */
  structureCacheSize(): number { return this.structures.size }
  /** Retained prepared working set (document-wide, all sheets). */
  preparedStats(): { columns: number; bytes: number } {
    return { columns: this.prepared.size, bytes: this.preparedBytes }
  }
  /** Resident prepared columns, value-owned: they survive view commits and are
   * read from the store at most once per value generation while they stay in
   * the admitted working set. Transactions are serialized, so two overlapping
   * batches never double-charge a column or evict each other mid-flight. A batch
   * is atomic: staged lanes are charged against the profile budget BEFORE they
   * are retained (including spare capacity and the per-column overhead), and
   * nothing is kept when admission fails. Room is made by dropping least-recently
   * used prepared columns outside the batch; admitted columns are `rebuildable`
   * entries of the document pool, which may evict them under aggregate pressure
   * (S4.3) and charges their staging to the same budget. */
  prepareColumns(owner: CacheIdentity, cols: readonly number[], signal: AbortSignal, hooks: OperationHooks = {}): Promise<PrepareOutcome> {
    const run = this.prepareQueue.then(() => this.prepareTransaction(owner, cols, signal, hooks))
    this.prepareQueue = run.catch(() => undefined)
    return run
  }
  /** Staged lanes are charged to the transient ledger while they are built and
   * released on every outcome; admitted columns move to the prepared ledger.
   * Each populated entry or blank run is one unit of the shared yield budget.
   * Progress counts rows covered by completed pages (a resident column counts
   * whole), over `batch × MAX_ROWS`. */
  private async prepareTransaction(owner: CacheIdentity, cols: readonly number[], signal: AbortSignal, hooks: OperationHooks): Promise<PrepareOutcome> {
    if (!Array.isArray(cols) || !cols.every(col => isValidCoordinate(0, col))) {
      return { kind: 'error', error: failure('invalid-coordinate', 'Columns must be valid zero-based column indexes') }
    }
    if (signal.aborted) return { kind: 'cancelled' }
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    const batch = [...new Set(cols)]
    if (batch.length > MAX_PREPARE_BATCH) {
      return { kind: 'resource-limit', message: `At most ${MAX_PREPARE_BATCH} columns are prepared per batch` }
    }
    const budget = this.budget
    if (batch.length > budget.maxColumns) {
      return {
        kind: 'quota', limitBytes: budget.maxBytes, requiredBytes: batch.length * PREPARED_COLUMN_OVERHEAD,
        diagnostics: [{ code: 'prepared-column-limit', message: `Profile working set holds at most ${budget.maxColumns} prepared columns` }],
      }
    }
    // Admitted: acknowledge with zero completed batches over the batch total.
    hooks.progress?.(0, batch.length * MAX_ROWS)
    const slice = this.slice()
    const charge = this.charge()
    try {
      return await this.stageColumns(owner, batch, signal, hooks, slice, charge, budget)
    } finally {
      charge.release()
      slice.dispose()
    }
  }
  private async stageColumns(
    owner: CacheIdentity, batch: readonly number[], signal: AbortSignal, hooks: OperationHooks,
    slice: CooperativeSlice<string>, charge: Charge, budget: PreparedBudget,
  ): Promise<PrepareOutcome> {
    const store = this.source.store(owner.sheetId)
    const stamp = store.stamp
    const keys = batch.map(col => this.preparedKey(stamp, col))
    let protectedBytes = 0
    for (const key of keys) protectedBytes += this.prepared.get(key)?.bytes ?? 0
    const staged: { key: string; column: PreparedColumn }[] = []
    let stagedBytes = 0, covered = 0
    const total = batch.length * MAX_ROWS
    const quota = (required: number): PrepareOutcome => ({ kind: 'quota', limitBytes: budget.maxBytes, requiredBytes: required })
    for (let i = 0; i < batch.length; i++) {
      const col = batch[i], key = keys[i]
      if (this.prepared.has(key)) { covered += MAX_ROWS; continue }
      const builder = new ColumnBuilder(col)
      // Reserve the column overhead before reading anything.
      if (protectedBytes + stagedBytes + builder.heldBytes > budget.maxBytes) return quota(protectedBytes + stagedBytes + builder.heldBytes)
      if (!charge.set(stagedBytes + builder.heldBytes)) return this.prepareRefusal(stagedBytes + builder.heldBytes)
      let cursor: string | undefined
      // Next row the full-axis query must cover (§4.1 column page ordering).
      let expected = 0
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
          // The lanes are binary-searched later, so the order is enforced, not
          // assumed: entries must cover rows ascending and contiguously.
          const first = 'row' in item ? item.row : item.r0, last = 'row' in item ? item.row : item.r1
          if (first !== expected || !(last >= first) || last >= MAX_ROWS) return { kind: 'error', error: UNORDERED_PAGE }
          expected = last + 1
          if (item.kind === 'unknown') return { kind: 'incomplete', missing: [{ r0: item.r0, r1: item.r1, c0: col, c1: col }] }
          if (item.kind === 'error') return { kind: 'error', error: failure(item.code, 'Column source failed', item.retryable) }
          // Lanes are never requested here; treating one as a blank run would hide its values.
          if (item.kind === 'lanes') return { kind: 'error', error: failure('page-error', 'Unrequested column lanes') }
          if (item.kind === 'cell' || item.kind === 'unavailable') {
            const next = protectedBytes + stagedBytes + builder.heldBytes + builder.costOf(item)
            if (next > budget.maxBytes) return quota(next)
            builder.append(item)
            if (!charge.set(stagedBytes + builder.heldBytes)) return this.prepareRefusal(stagedBytes + builder.heldBytes)
          }
          if (slice.due()) {
            await slice.pause('prepare')
            const halt = this.halted(owner, signal)
            if (halt) return halt
          }
        }
        if (page.completeness !== 'complete') return { kind: 'incomplete', missing: [{ r0: 0, r1: MAX_ROWS - 1, c0: col, c1: col }] }
        // Validated contiguous from row 0: `expected` is the covered prefix.
        if (expected > 0) hooks.progress?.(covered + expected, total)
        cursor = page.nextCursor
        if (cursor !== undefined) {
          if (seen.has(cursor)) return { kind: 'error', error: failure('invalid-cursor', 'Column cursor made no progress') }
          seen.add(cursor)
        }
      } while (cursor !== undefined)
      if (expected !== MAX_ROWS) return { kind: 'error', error: UNORDERED_PAGE }
      covered += MAX_ROWS
      // Trimming spare capacity briefly holds both lane sets; that peak is charged too.
      const exactLanes = builder.count * PREPARED_ENTRY_LANE_BYTES
      const trim = builder.capacity > builder.count ? exactLanes : 0
      const peak = protectedBytes + stagedBytes + builder.heldBytes + trim
      if (peak > budget.maxBytes) return quota(peak)
      if (!charge.set(stagedBytes + builder.heldBytes + trim)) return this.prepareRefusal(stagedBytes + builder.heldBytes + trim)
      const column = builder.finish(stamp)
      staged.push({ key, column })
      stagedBytes += column.bytes
      if (!charge.set(stagedBytes)) return this.prepareRefusal(stagedBytes)
    }
    if (signal.aborted) return { kind: 'cancelled' }
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    const keep = new Set(keys)
    // Evict least-recently-used columns outside this batch until it fits.
    for (const [key] of [...this.prepared]) {
      const fitsCount = this.prepared.size + staged.length <= budget.maxColumns
      const fitsBytes = this.preparedBytes + stagedBytes <= budget.maxBytes
      if (fitsCount && fitsBytes) break
      if (keep.has(key)) continue
      this.dropPrepared(key)
    }
    for (const key of keys) {
      const existing = this.prepared.get(key)
      if (existing) { this.prepared.delete(key); this.prepared.set(key, existing); this.preparedEntries.get(key)?.touch() }
    }
    // The staged bytes move from the transient ledger to retained entries.
    charge.set(0)
    for (const { key, column } of staged) {
      const entry = this.memory.retain(JSON.stringify(['prepared', key]), {
        tier: 'rebuildable', bytes: column.bytes, owner: owner.sheetId, onRelease: () => this.dropPrepared(key),
      })
      if (!entry) return this.prepareRefusal(column.bytes)
      this.prepared.set(key, column)
      this.preparedEntries.set(key, entry)
      this.preparedBytes += column.bytes
    }
    return { kind: 'complete' }
  }
  /** The single uncharge point of a prepared column (owner drop or pool
   * eviction/close): removes it from the working set exactly once. */
  private dropPrepared(key: string): void {
    const column = this.prepared.get(key)
    if (!column) return
    this.prepared.delete(key)
    this.preparedBytes -= column.bytes
    const entry = this.preparedEntries.get(key)
    this.preparedEntries.delete(key)
    entry?.release()
  }
  private prepareRefusal(requiredBytes: number): PrepareOutcome {
    return this.closed ? { kind: 'stale' } : { kind: 'quota', limitBytes: this.memory.limitBytes, requiredBytes }
  }
  /** Column page from resident lanes over a residual membership. */
  private preparedPage(column: PreparedColumn, residual: readonly SourceRowInterval[], limit: number): ColumnEntry[] {
    const items: ColumnEntry[] = []
    let at = 0, row = residual[0]?.r0 ?? 0
    while (at < residual.length && items.length < limit) {
      const { r0, r1 } = residual[at]
      if (row < r0) row = r0
      if (row > r1) { at++; continue }
      const k = lowerBound(column.rows, row, column.count)
      const next = k < column.count && column.rows[k] <= r1 ? column.rows[k] : -1
      if (next < 0) { items.push({ kind: 'blank', r0: row, r1 }); at++; continue }
      if (next > row) { items.push({ kind: 'blank', r0: row, r1: next - 1 }); row = next; continue }
      items.push(preparedEntry(column, k))
      row++
    }
    return items
  }
  /**
   * The sheet store with prepared columns served from resident lanes and with
   * SESSION-OWNED pagination: cursors are positional over the requested
   * membership (`vq1:<query fingerprint>:<interval>:<row>`), and every page asks
   * the backing store for the residual membership without a backend cursor. A
   * cursor therefore stays valid when its column is prepared or evicted between
   * pages, whatever cursor encoding the store uses (§4.1 cursors are opaque).
   */
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
      queryColumn: (request, options) => session.pagedColumn(store, request, options),
    }
  }
  private async pagedColumn(store: SourceStore, request: ColumnQuery, options: { signal: AbortSignal; stamp: SourceStamp }): Promise<Page<ColumnEntry>> {
    const stamp = store.stamp
    if (options.signal?.aborted) return failedPage(stamp, failure('cancelled', 'Request cancelled'))
    if (!options.stamp || !sameStamp(options.stamp, stamp)) return failedPage(stamp, failure('stale-source', 'Source owner changed', true))
    if (!isValidCoordinate(0, request.col)) return failedPage(stamp, failure('invalid-coordinate', 'Invalid column'))
    if (!Number.isSafeInteger(request.limit) || request.limit <= 0) return failedPage(stamp, failure('invalid-source-rect', 'Invalid limit'))
    let intervals: SourceRowInterval[] | null
    try { intervals = rowSetIntervals(request.rows) } catch (cause) {
      return failedPage(stamp, failure('invalid-source-rect', (cause as Error).message))
    }
    // Resident row-set refs are resolved only by the backend (S1.2).
    if (!intervals) return store.queryColumn(request, options)
    const fingerprint = queryFingerprint(request.col, stamp, intervals)
    let at = 0, row = intervals[0]?.r0 ?? 0
    if (request.cursor !== undefined) {
      const parts = request.cursor.split(':')
      const i = Number(parts[2]), r = Number(parts[3])
      if (parts.length !== 4 || parts[0] !== 'vq1' || parts[1] !== fingerprint || !Number.isSafeInteger(i)
        || i < 0 || i >= intervals.length || !isValidCoordinate(r, 0) || r < intervals[i].r0 || r > intervals[i].r1) {
        return failedPage(stamp, failure('invalid-cursor', 'Invalid or foreign column cursor'))
      }
      at = i; row = r
    }
    const residual = residualIntervals(intervals, at, row)
    if (!residual.length) return { items: [], completeness: 'complete', stamp }
    const key = this.preparedKey(stamp, request.col)
    const prepared = this.prepared.get(key)
    let items: readonly ColumnEntry[]
    let completeness: Page<ColumnEntry>['completeness'] = 'complete'
    if (prepared) {
      this.prepared.delete(key)
      this.prepared.set(key, prepared)
      this.preparedEntries.get(key)?.touch()
      items = this.preparedPage(prepared, residual, request.limit)
    } else {
      const page = await store.queryColumn({ col: request.col, rows: { kind: 'intervals', intervals: residual }, limit: request.limit }, options)
      if (page.error) return { ...page, nextCursor: undefined }
      if (!sameStamp(page.stamp, stamp)) return failedPage(stamp, failure('stale-source', 'Column page owner changed', true))
      items = page.items
      completeness = page.completeness
    }
    // Advance the position by the rows these entries cover; entries must cover the
    // residual membership contiguously, in order, or the page is malformed.
    let j = 0, next = residual[0].r0
    for (const item of items) {
      const first = 'row' in item ? item.row : item.r0, last = 'row' in item ? item.row : item.r1
      if (j >= residual.length || first !== next || last < first || last > residual[j].r1) {
        return failedPage(stamp, failure('page-error', 'Column page does not cover its rows contiguously'))
      }
      next = last + 1
      if (next > residual[j].r1) { j++; if (j < residual.length) next = residual[j].r0 }
    }
    if (j < residual.length && !items.length) return failedPage(stamp, failure('page-error', 'Column page made no progress'))
    return {
      items, completeness, stamp,
      ...(j < residual.length ? { nextCursor: `vq1:${fingerprint}:${at + j}:${next}` } : {}),
    }
  }
  /** Structural handles are value-owned and shared by every guard (§6.3). Only a
   * complete index is cached; incomplete coverage is re-queried next time. */
  private async structure(sheetId: string, signal: AbortSignal): Promise<StructuralIndex> {
    const store = this.source.store(sheetId)
    const stamp = store.stamp
    const cached = this.structures.get(sheetId)
    if (cached && sameStamp(cached.stamp, stamp)) return cached
    const index = new StructuralIndex(await store.queryStructural(FULL_SHEET, { signal, stamp }))
    // Cache only a complete index for the generation it was queried under.
    if (index.completeness === 'complete' && sameStamp(index.stamp, stamp) && !this.closed) {
      this.dropStructure(sheetId)
      const bytes = STRUCTURE_BASE_BYTES + index.recordCount * STRUCTURE_RECORD_BYTES
      const entry = this.memory.retain(structureKey(sheetId), {
        tier: 'rebuildable', bytes, owner: sheetId, onRelease: () => this.dropStructure(sheetId),
      })
      // Not admitted: the index still serves this operation, uncached.
      if (entry) {
        this.structures.set(sheetId, index)
        this.structureEntries.set(sheetId, entry)
        this.structureBytes += bytes
      }
    } else if (cached && !sameStamp(cached.stamp, stamp)) this.dropStructure(sheetId)
    return index
  }
  /** The single uncharge point of a cached structural handle. */
  private dropStructure(sheetId: string): void {
    this.structures.delete(sheetId)
    const entry = this.structureEntries.get(sheetId)
    if (!entry) return
    this.structureEntries.delete(sheetId)
    this.structureBytes -= entry.bytes
    entry.release()
  }
  /** Populated-cell feed for outside-column guards over resident readRows pages.
   * Honors the caller's per-page limit: a store page larger than the limit is
   * served in slices, resuming at an offset within that same store page. With a
   * slice, each served page (at most 200 source rows) steps the shared budget. */
  populatedSource(sheetId: string, signal: AbortSignal, slice?: CooperativeSlice<string>): PopulatedCellSource {
    const store = this.source.store(sheetId)
    return {
      async iterate(range, cursor, limit) {
        if (limit !== undefined && (!Number.isSafeInteger(limit) || limit <= 0)) throw new RangeError('Invalid populated page limit')
        let storeCursor: string | undefined, skip = 0
        if (cursor !== undefined) {
          let decoded: unknown
          try { decoded = JSON.parse(cursor) } catch { decoded = undefined }
          if (!Array.isArray(decoded) || decoded.length !== 2 || !(decoded[0] === null || typeof decoded[0] === 'string')
            || !Number.isSafeInteger(decoded[1]) || decoded[1] < 0) throw new RangeError('Invalid populated cursor')
          storeCursor = decoded[0] ?? undefined
          skip = decoded[1]
        }
        const page = await store.readRows(range, storeCursor, { signal, stamp: store.stamp })
        const all = page.items.flatMap(row => row.cells
          .filter(cell => cell.value !== null || cell.formula !== undefined)
          .map(cell => ({ row: cell.row, col: cell.col })))
        const complete = page.completeness === 'complete' && page.error === undefined
        if (!complete) return { items: [], completeness: 'incomplete', missing: [range] }
        const end = limit === undefined ? all.length : Math.min(all.length, skip + limit)
        const items = all.slice(skip, end)
        const nextCursor = end < all.length ? JSON.stringify([storeCursor ?? null, end])
          : page.nextCursor !== undefined ? JSON.stringify([page.nextCursor, 0]) : undefined
        await slice?.step('guard', Math.max(1, items.length))
        return { items, completeness: 'complete', missing: [], ...(nextCursor !== undefined ? { nextCursor } : {}) }
      },
    }
  }
  private coverageDiagnostics(sheetId: string): readonly WorkerDiagnostic[] {
    return this.source.coverageDiagnostics?.(sheetId) ?? []
  }
  /** Sort draft: structural guard, cooperative key sort, cooperative mask
   * projection, then one validated commit. The draft order/inverse and mask are
   * charged to the transient ledger until the operation ends. */
  async sort(owner: CacheIdentity, spec: SortSpec | null, signal: AbortSignal, hooks: OperationHooks = {}): Promise<OperationOutcome> {
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    const resident = this.residentNow(owner.sheetId) ?? await this.residentFor(owner, signal)
    if ('kind' in resident) return resident
    const state = resident.state
    // The captured view must not be evicted while this draft reads it.
    const unpin = this.pinView(owner.sheetId)
    const slice = this.slice()
    const charge = this.charge()
    try {
      if (spec !== null) {
        const guard = await guardSort(spec.range, await this.structure(owner.sheetId, signal), this.populatedSource(owner.sheetId, signal, slice))
        // Hand over to the sort kernel's own budget at a yield.
        await slice.flush('guard')
        const halt = this.halted(owner, signal)
        if (halt) return halt
        if (!guard.ok) return guardOutcome(guard, this.coverageDiagnostics(owner.sheetId))
      }
      hooks.progress?.(1, VIEW_OPERATION_STAGES)
      const store = this.preparedStore(owner.sheetId)
      // Make room for the kernel's working set, then let its own budget admit
      // against the AGGREGATE document pool (every sheet's retained bytes).
      this.memory.trim(Math.max(0, this.memory.limitBytes - this.sortEstimate(spec)))
      const result = await sortRowOrder(store, spec, {
        signal, stamp: store.stamp, locale: this.options.locale,
        hidden: hiddenSourceRowsFromDisplay(state.order, state.visibility),
        residentWorkerBytes: this.memory.totalBytes, maxWorkerBytes: this.memory.limitBytes,
        ...(this.source.retainsKeyStrings?.(owner.sheetId) === true ? { sourceRetainsStrings: true } : {}),
        ...(this.options.yield ? { yield: this.options.yield } : {}),
        ...(this.options.now ? { now: this.options.now } : {}),
      })
      if (result.kind !== 'complete') return sortFailure(result, spec)
      // Guards and kernels awaited: continue only if nothing superseded the base.
      const halt = this.halted(owner, signal)
      if (halt) return halt
      hooks.progress?.(2, VIEW_OPERATION_STAGES)
      const orderBytes = result.ledger.orderBytes + result.ledger.inverseBytes
      if (!charge.set(orderBytes)) return this.refusal(orderBytes)
      const projected = await this.projectDraft(owner, state.filter, result.order, state.matches ?? matchAll, slice, signal, charge, orderBytes)
      if (projected.kind !== 'projected') return projected
      hooks.progress?.(3, VIEW_OPERATION_STAGES)
      // The draft's bytes become the retained view: uncharge them first so the
      // commit never counts them twice. Recipe: the sort pinned exactly the rows
      // hidden by the filter it ran under.
      charge.set(0)
      return this.commitOutcome(() => this.commitDraft(owner, {
        order: result.order, visibility: projected.visibility, sort: spec, filter: state.filter, matches: state.matches,
      }, Object.freeze({ sortFilter: state.filter, sortMatches: state.matches })))
    } finally {
      charge.release()
      slice.dispose()
      unpin()
    }
  }
  /** Maps a validated commit to an outcome; an unadmittable view is quota. */
  private commitOutcome(commit: () => CommittedView | null): OperationOutcome {
    try {
      const view = commit()
      return view ? { kind: 'committed', view } : { kind: 'stale' }
    } catch (cause) {
      if (cause instanceof ViewQuotaError) return this.refusal(cause.requiredBytes)
      throw cause
    }
  }
  /** Resident view without awaiting (no extra scheduling turn for the common
   * case); null when an evicted order needs the async rebuild. */
  private residentNow(sheetId: string): { state: ResidentState } | null {
    const state = this.state(sheetId)
    if (isResidentState(state)) return { state }
    if (state.order === null) return null
    try { return { state: this.residentSync(sheetId) } } catch { return null }
  }
  /** Materializes the owner's sheet view for an operation and rechecks the owner. */
  private async residentFor(owner: CacheIdentity, signal: AbortSignal): Promise<{ state: ResidentState } | OperationOutcome> {
    const resident = await this.materialize(owner.sheetId, signal)
    if ('kind' in resident) return resident
    if (signal.aborted) return { kind: 'cancelled' }
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    return { state: resident.state }
  }
  /** Filter draft: rule evaluation, cooperative mask projection, the drawing
   * geometry guard over before/after metrics, then one validated commit. */
  async filter(owner: CacheIdentity, spec: FilterSpec | null, nowMs: number, signal: AbortSignal, hooks: OperationHooks = {}): Promise<OperationOutcome> {
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    let normalized: FilterSpec | null
    try { normalized = normalizeFilter(spec) } catch (cause) {
      return { kind: 'refused', refusal: { reason: 'invalid-rule', message: (cause as Error).message } }
    }
    const evaluator = this.options.filterEvaluator
    if (normalized && !evaluator) {
      return { kind: 'refused', refusal: { reason: 'unsupported', message: 'Typed filter rule evaluation is not available yet (S8)' } }
    }
    const resident = this.residentNow(owner.sheetId) ?? await this.residentFor(owner, signal)
    if ('kind' in resident) return resident
    const state = resident.state
    const unpin = this.pinView(owner.sheetId)
    const slice = this.slice()
    const charge = this.charge()
    const metricsCharge = this.charge()
    try {
      let matches: ((row: SourceRow, filter: FilterSpec) => boolean) | undefined
      if (normalized && evaluator) {
        // §4.4: an unavailable (cache-less formula) value inside the rule
        // columns' body refuses the filter with its count and first address;
        // unavailable cells elsewhere never block. Checked before any rule runs.
        // Engine documents have no unavailable values (§4.4): nothing to scan.
        const domain = this.source.store(owner.sheetId).valuePolicy === 'cached' ? filterValueDomain(normalized) : null
        if (domain) {
          const store = this.preparedStore(owner.sheetId)
          // Its own yield budget, charged in batches of scanned entries (like
          // the guard's 32-row words) and created on first charge: a small
          // domain samples no clock, and the rule/guard phases keep the
          // operation slice's accounting unchanged.
          let scanSlice: CooperativeSlice<string> | undefined
          let pending = 0
          let scan: Awaited<ReturnType<typeof scanValueDomain>>
          try {
            scan = await scanValueDomain(store, domain, { signal, stamp: store.stamp }, async units => {
              pending += units
              if (pending < DOMAIN_SCAN_CHARGE) return
              scanSlice ??= this.slice()
              const due = scanSlice.due(pending)
              pending = 0
              if (due) await scanSlice.pause('filter-domain')
            })
          } finally { scanSlice?.dispose() }
          const halt = this.halted(owner, signal)
          if (halt) return halt
          switch (scan.kind) {
            case 'clear': break
            case 'refused': return { kind: 'refused', refusal: scan.refusal }
            case 'incomplete': return { kind: 'incomplete', missing: scan.missing }
            case 'cancelled': return { kind: 'cancelled' }
            case 'error': return scan.error.code === 'stale-source' ? { kind: 'stale' } : { kind: 'error', error: scan.error }
          }
        }
        const evaluation = await evaluator({ store: this.preparedStore(owner.sheetId), spec: normalized, nowMs, signal, cooperate: slice })
        if (evaluation.kind !== 'complete') return evaluation
        const halt = this.halted(owner, signal)
        if (halt) return halt
        matches = evaluation.matches
      }
      hooks.progress?.(1, VIEW_OPERATION_STAGES)
      const projected = await this.projectDraft(owner, normalized, state.order, matches ?? matchAll, slice, signal, charge, 0)
      if (projected.kind !== 'projected') return projected
      hooks.progress?.(2, VIEW_OPERATION_STAGES)
      const geometry = this.source.geometry(owner.sheetId)
      const guardRange = normalized?.range ?? state.filter?.range
      if (guardRange) {
        // Admit both transient projections (dense sections per order band) first.
        const metricsBytes = 2 * metricsUpperBound(state.order)
        if (!metricsCharge.set(metricsBytes)) return this.refusal(metricsBytes)
        const before = createSheetMetrics({ rows: geometry.rows, cols: geometry.cols, order: state.order, visibility: state.visibility })
        const after = createSheetMetrics({ rows: geometry.rows, cols: geometry.cols, order: state.order, visibility: projected.visibility })
        let guard: GuardOutcome
        try {
          guard = await guardFilter({
            range: guardRange, candidateVisibility: projected.visibility, currentVisibility: state.visibility,
            projectionBefore: before, projectionAfter: after, structure: await this.structure(owner.sheetId, signal),
            scratch: {
              // A separate share: the guard bitmap is live alongside the draft mask.
              charge: bytes => {
                const scratch = this.charge()
                if (!scratch.set(bytes)) throw new QuotaHalt(bytes)
                return () => scratch.release()
              },
            },
            cooperate: {
              due: units => slice.due(units),
              pause: async () => {
                await slice.pause('guard')
                const halt = this.halted(owner, signal)
                if (halt) throw new GuardHalt(halt)
              },
            },
          })
        } catch (cause) {
          if (cause instanceof GuardHalt) return cause.outcome
          if (cause instanceof QuotaHalt) return this.refusal(cause.requiredBytes)
          throw cause
        }
        if (!guard.ok) return guardOutcome(guard, this.coverageDiagnostics(owner.sheetId))
      }
      const halt = this.halted(owner, signal)
      if (halt) return halt
      hooks.progress?.(3, VIEW_OPERATION_STAGES)
      // The draft mask becomes the retained view (no double count at commit).
      charge.set(0)
      metricsCharge.set(0)
      return this.commitOutcome(() => this.commitDraft(owner, {
        order: state.order, visibility: projected.visibility, sort: state.sort, filter: normalized, matches,
      }))
    } finally {
      metricsCharge.release()
      charge.release()
      slice.dispose()
      unpin()
    }
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
  /** Compact source-space layout of a sheet (S5.4): what the main thread needs
   * to build the same axis metrics and base visibility as this session. */
  layout(sheetId: string): SheetLayout {
    const geometry = this.source.geometry(sheetId)
    const used = this.source.store(sheetId).extent.used
    return {
      rowOverrides: geometry.rowOverrides.map(interval => ({ ...interval })),
      colOverrides: geometry.colOverrides.map(interval => ({ ...interval })),
      hidden: geometry.baseHidden.intervals.map(interval => ({ ...interval })),
      ...(geometry.hiddenCols?.length ? { hiddenCols: geometry.hiddenCols.map(interval => ({ ...interval })) } : {}),
      ...(used ? { used: { ...used } } : {}),
    }
  }
  /** One bounded source read for the paged viewer (S5.4), owned by the sheet's
   * committed view: a newer commit, a close or a cancellation withholds it.
   * Store pages carry their own bounds (rows: READ_ROWS_PAGE_ROWS/MAX_CELLS);
   * a scan page is capped at MAX_SCAN_PAGE and continues a resident iterator. */
  async read(owner: CacheIdentity, query: SourceReadQuery, signal: AbortSignal): Promise<SourceReadResult | { kind: 'stale' } | { kind: 'cancelled' }> {
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    const store = this.source.store(owner.sheetId)
    const options = { signal, stamp: { ...store.stamp } }
    let result: SourceReadResult
    switch (query?.op) {
      case 'cell': result = { op: 'cell', read: await store.readCell(query.at, options) }; break
      case 'rows': result = { op: 'rows', page: await store.readRows(query.range, query.cursor, options) }; break
      case 'structural': result = { op: 'structural', result: await store.queryStructural(query.rect, options) }; break
      case 'complete': result = { op: 'complete', result: await store.ensureComplete(query.rect, options) }; break
      case 'scan': result = { op: 'scan', page: await this.scanPage(owner, store, query, signal) }; break
      case 'layout': result = { op: 'layout', layout: this.layout(owner.sheetId) }; break
      default: result = { op: 'rows', page: failedPage(store.stamp, failure('unsupported', 'Unknown read operation')) }
    }
    if (signal.aborted) return { kind: 'cancelled' }
    if (!this.isCurrent(owner)) return { kind: 'stale' }
    return result
  }
  /** Open scan continuations (bounded by MAX_OPEN_SCANS). */
  get openScans(): number { return this.scans.size }
  /** Ends an abandoned continuation; owned by the document, not a view generation. */
  releaseScan(cursor: string): boolean {
    const scan = this.scans.get(cursor)
    if (!scan) return false
    scan.controller.abort()
    this.scans.delete(cursor)
    return true
  }
  private async scanPage(owner: CacheIdentity, store: SourceStore, query: Extract<SourceReadQuery, { op: 'scan' }>, signal: AbortSignal): Promise<Page<AuthoritativeCell>> {
    const request = { range: query.request.range, limit: Math.min(MAX_SCAN_PAGE, query.request.limit) }
    const key = JSON.stringify([owner.sheetId, store.stamp.valueGen, request.range, request.limit])
    let scan: OpenScan | undefined
    if (query.cursor !== undefined) {
      scan = this.scans.get(query.cursor)
      this.scans.delete(query.cursor)
      // A cursor belongs to its query and source generation (§4.1).
      if (!scan || scan.key !== key) {
        scan?.controller.abort()
        return failedPage(store.stamp, failure('invalid-cursor', 'Scan cursor expired; restart the scan'))
      }
    } else {
      const controller = new AbortController()
      scan = { key, sheetId: owner.sheetId, valueGen: store.stamp.valueGen, controller,
        iterator: store.scanPopulated(request, { signal: controller.signal, stamp: { ...store.stamp } })[Symbol.asyncIterator]() }
    }
    const open = scan
    // Cancelling the command abandons its scan; the client restarts from page 0.
    const abort = () => open.controller.abort()
    signal.addEventListener('abort', abort, { once: true })
    let step: IteratorResult<Page<AuthoritativeCell>>, ahead: IteratorResult<Page<AuthoritativeCell>> | undefined
    try {
      step = open.ahead ?? await open.iterator.next()
      open.ahead = undefined
      if (!step.done && step.value.completeness !== 'incomplete' && !signal.aborted) ahead = await open.iterator.next()
    } finally { signal.removeEventListener('abort', abort) }
    if (step.done) return { items: [], completeness: 'complete', stamp: store.stamp }
    const page = step.value
    if (!ahead || ahead.done || signal.aborted) { open.controller.abort(); return page }
    open.ahead = ahead
    const cursor = `scan:${this.epoch}:${++this.scanSerial}`
    this.scans.set(cursor, open)
    while (this.scans.size > MAX_OPEN_SCANS) {
      const oldest = this.scans.keys().next().value as string
      this.scans.get(oldest)!.controller.abort()
      this.scans.delete(oldest)
    }
    return { ...page, nextCursor: cursor }
  }
  /** Close: every retained entry is released through the pool ('closed'), so
   * views, prepared columns and structural handles are uncharged exactly once;
   * in-flight operations release their transient shares in their own finally. */
  async release(): Promise<void> {
    if (this.closed) return
    this.closed = true
    this.memory.close()
    this.states.clear()
    this.prepared.clear()
    this.preparedEntries.clear()
    this.preparedBytes = 0
    this.structures.clear()
    this.structureEntries.clear()
    this.structureBytes = 0
    for (const scan of this.scans.values()) scan.controller.abort()
    this.scans.clear()
    await this.source.release()
  }
}
function sortFailure(result: Exclude<SortResult, { kind: 'complete' }>, spec: SortSpec | null): OperationOutcome {
  switch (result.kind) {
    case 'cancelled': return { kind: 'cancelled' }
    case 'incomplete': return { kind: 'incomplete', missing: spec ? [{ r0: spec.range.r0, r1: spec.range.r1, c0: spec.range.c0, c1: spec.range.c1 }] : [] }
    case 'refused': return { kind: 'refused', refusal: result.refusal }
    case 'error':
      if (result.error.code === 'cancelled') return { kind: 'cancelled' }
      // The kernel's stamp check: the source owner moved mid-sort.
      if (result.error.code === 'stale-source') return { kind: 'stale' }
      return { kind: 'error', error: result.error }
    case 'resource-limit': return { kind: 'quota', requiredBytes: result.requiredBytes, limitBytes: result.limitBytes }
  }
}
