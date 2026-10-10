import type { XlsxCell, XlsxDocument } from '../types'
import type { CellId, SourceCell, SourceCol, SourceRect, SourceRow } from './coordinates'

/** Source-space contracts, spec §4.1; limits bound output, never semantic membership. */
export const MAX_OVERVIEW_GROUPS_DESKTOP = 2048
export const MAX_OVERVIEW_GROUPS_MOBILE = 1024
export type Completeness = 'complete' | 'loading' | 'incomplete'
export type SourceErrorCode =
  | 'invalid-coordinate' | 'invalid-source-rect' | 'invalid-cursor' | 'stale-source'
  | 'cancelled' | 'incomplete' | 'quota' | 'resource-limit' | 'unsupported'
  | 'parse-error' | 'storage-error' | 'page-error'

/** A policy is fixed at document admission, spec §4.4. */
export type ValuePolicy = 'engine' | 'cached'
export interface AuthoritativeCell extends Pick<XlsxCell,
  'value' | 'hasCachedValue' | 'valueIsError' | 'formula' | 'styleIndex'
  | 'sharedFormula' | 'arrayRef' | 'dynamicArray' | 'formulaType' | 'ca' | 'style'> {
  id: CellId
  row: SourceRow
  col: SourceCol
  valueProvenance?: 'engine' | 'cached'
}
/** Unavailable is permanent; unknown is a pending page, never a proven blank (§4.4). */
export type CellRead =
  | { kind: 'cell'; cell: AuthoritativeCell }
  | { kind: 'blank' }
  | { kind: 'unknown'; pageId: string }
  | { kind: 'error'; code: SourceErrorCode; retryable: boolean }
  | { kind: 'unavailable'; cell: AuthoritativeCell; reason: 'formula-value-unavailable' }

export interface SourceStamp { docId: string; sheetId: string; valueGen: number }
export interface ReadOptions { signal: AbortSignal; stamp: SourceStamp }
/** Cursors belong to the original query and stamp; a changed owner invalidates them (§4.1). */
export interface Page<T> {
  items: readonly T[]
  nextCursor?: string
  completeness: Completeness
  stamp: SourceStamp
}
export interface SourceRowInterval { r0: SourceRow; r1: SourceRow }
export interface SourceColInterval { c0: SourceCol; c1: SourceCol }

/** Exact source membership, spec §§4.1, 8.1. Buffers are bounded by an inclusive
 * band; refs stay resident in the worker and must not be copied to the UI.
 * Intervals are disjoint; lists contain unique members; bitsets exclude padding. */
export type CompactSourceRowSet =
  | { kind: 'intervals'; intervals: readonly SourceRowInterval[] }
  | { kind: 'list'; r0: SourceRow; r1: SourceRow; rows: Uint32Array }
  | { kind: 'bitset'; r0: SourceRow; r1: SourceRow; words: Uint32Array; count: number }
  | { kind: 'ref'; refId: string; count: number; stamp: SourceStamp }

export interface OverviewDomain {
  groupId: string
  rows: CompactSourceRowSet
  cols: readonly SourceColInterval[]
}
export interface OverviewQuery {
  groups: readonly OverviewDomain[]
  /** Positive bounded output: ≤2048 desktop / ≤1024 mobile (§4.1). */
  maxGroups: number
  cursor?: string
}
/** Raw aggregates have source ownership and no viewport geometry or sampled text (§4.1). */
export interface OverviewGroup {
  groupId: string
  populatedCount: number
  countKind: 'exact' | 'estimated'
  style: { kind: 'uniform'; styleIndex: number } | { kind: 'mixed' }
  content: 'blank' | 'populated' | 'mixed'
}
/** Sparse, bounded pages and range predicates, spec §4.1. */
export interface SparseRowPage {
  row: SourceRow
  cells: readonly AuthoritativeCell[]
  range: SourceRect
  limit: number
  cursor?: string
}
export interface ColumnQuery {
  col: SourceCol
  rows: CompactSourceRowSet
  limit: number
  cursor?: string
}
export type ColumnEntry =
  | { kind: 'cell'; row: SourceRow; cell: AuthoritativeCell }
  | { kind: 'blank'; r0: SourceRow; r1: SourceRow }
  | { kind: 'unknown'; r0: SourceRow; r1: SourceRow; pageId: string }
  | { kind: 'unavailable'; row: SourceRow; cell: AuthoritativeCell; reason: 'formula-value-unavailable' }
  | { kind: 'error'; r0: SourceRow; r1: SourceRow; code: SourceErrorCode; retryable: boolean }
export interface PopulatedQuery { range: SourceRect; limit: number; cursor?: string }
export interface SourceFailure { code: SourceErrorCode; retryable: boolean; message: string }
export interface RangeCompletion {
  range: SourceRect
  completeness: Completeness
  stamp: SourceStamp
  missing: readonly SourceRect[]
  error?: SourceFailure
}
export interface SourceStore {
  readonly stamp: SourceStamp
  readonly extent: { rows: 1048576; cols: 16384; used?: SourceRect }
  readCell(at: SourceCell, options: ReadOptions): Promise<CellRead>
  readRows(range: SourceRect, cursor: string | undefined, options: ReadOptions): Promise<Page<SparseRowPage>>
  queryColumn(request: ColumnQuery, options: ReadOptions): Promise<Page<ColumnEntry>>
  scanPopulated(request: PopulatedQuery, options: ReadOptions): AsyncIterable<Page<AuthoritativeCell>>
  queryStructural(rect: SourceRect, options: ReadOptions): Promise<StructuralResult>
  /** Complete sources terminate with a complete coarse result or explicit resource/error (§4.1). */
  queryOverview(request: OverviewQuery, options: ReadOptions): Promise<Page<OverviewGroup>>
  ensureComplete(rect: SourceRect, options: ReadOptions): Promise<RangeCompletion>
  release(): Promise<void>
}

/** Interval/range shapes only; no per-covered-cell expansion, spec §4.5. */
export interface MergeExtent { id: string; range: SourceRect; anchor: SourceCell }
export interface TableExtent {
  id: string
  range: SourceRect
  headerRows: 0 | 1
  totalsRows: 0 | 1
}
export interface SpillExtent {
  anchor: SourceCell
  kind: 'fixed' | 'dynamic'
  saved?: SourceRect
  evaluated?: SourceRect
}
export interface DrawingMarker extends SourceCell { offsetXEmu: number; offsetYEmu: number }
export interface EmuBox { x: number; y: number; width: number; height: number }
export type DrawingAnchor =
  | { id: string; kind: 'oneCell'; from: DrawingMarker; extent: EmuBox; bounds: SourceRect }
  | { id: string; kind: 'twoCell'; from: DrawingMarker; to: DrawingMarker; extent: EmuBox; bounds: SourceRect }
  | { id: string; kind: 'absolute'; extent: EmuBox }
export interface StructuralResult {
  stamp: SourceStamp
  completeness: Completeness
  merges: readonly MergeExtent[]
  tables: readonly TableExtent[]
  spills: readonly SpillExtent[]
  drawings: readonly DrawingAnchor[]
  missing: readonly SourceRect[]
  error?: SourceFailure
}
/** Refusals never publish a changed view owner, spec §§4.4, 6.1, 6.3. */
export type OperationRefusal =
  | { reason: 'formula-value-unavailable'; count: number; firstAddress: SourceCell }
  | { reason: 'merge' | 'fixed-array' | 'dynamic-spill' | 'outside-column-data'; range: SourceRect }
  | { reason: 'drawing-impact'; drawingId: string }
  | { reason: 'incomplete'; missing: readonly SourceRect[] }
  | { reason: 'ambiguous-selection' | 'invalid-range' | 'invalid-rule' | 'resource-limit' | 'unsupported'; message: string }
export type RefusalReason = OperationRefusal['reason']

/** Compact row state and source-identity selection, spec §§5.1, 8.1. */
export interface RowOrder {
  kind: 'identity' | 'bands'
  bands: readonly RowPermutationBand[]
  displayToSource(displayRow: number): SourceRow
  sourceToDisplay(sourceRow: SourceRow): number
}
export interface RowPermutationBand {
  r0: SourceRow
  r1: SourceRow
  /** Exactly the band length; inverse offsets have the same length (§5.1). */
  sourceAtOffset: Uint32Array
  displayOffsetOfSource: Uint32Array
}
export interface RowVisibility {
  visible(displayRow: number): boolean
  rank(endExclusive: number): number
  select(visibleOrdinal: number): number | null
  next(displayRow: number, direction: -1 | 1): number | null
}
export interface IntervalSet {
  intervals: readonly SourceRowInterval[]
  contains(row: SourceRow): boolean
}
export interface GridSelection {
  active: SourceCell | null
  anchor: SourceCell | null
  ranges: readonly GridSelectionPart[]
}
export interface GridSelectionPart { rows: CompactSourceRowSet; cols: readonly SourceColInterval[] }
export interface SheetView {
  sheetId: string
  revision: number
  order: RowOrder
  baseHidden: IntervalSet
  visibility: RowVisibility
  sort: SortSpec | null
  filter: FilterSpec | null
  selection: GridSelection
}

/** Required dependencies of SheetView/worker commands, spec §§6.1, 7.1, 7.4. */
export interface FilterRange extends SourceRect { headerRows: 0 | 1; totalsRows: 0 | 1 }
export interface SortKey { col: SourceCol; order: 'asc' | 'desc' }
export interface SortSpec { range: FilterRange; keys: readonly SortKey[]; caseSensitive: boolean }
export type NumberOp = 'eq' | 'neq' | 'gt' | 'gte' | 'lt' | 'lte' | 'between'
export type TextOp = 'eq' | 'neq' | 'begins' | 'ends' | 'contains' | 'not-contains'
export type DateOp = 'eq' | 'neq' | 'before' | 'after' | 'between'
export interface CalendarDay { year: number; month: number; day: number }
export type DatePreset = 'today' | 'yesterday' | 'tomorrow' | 'thisWeek' | 'thisMonth' | 'thisYear'
export type TypedCondition =
  | { kind: 'number'; op: NumberOp; first: number; second?: number }
  | { kind: 'text'; op: TextOp; first: string; wildcards: boolean }
  | { kind: 'date'; op: DateOp; first: CalendarDay; second?: CalendarDay }
  | { kind: 'datePreset'; preset: DatePreset }
export interface CompoundFilter { kind: 'compound'; connective: 'and' | 'or'; first: TypedCondition; second: TypedCondition }
export interface ValuesFilter { kind: 'values'; selection: ValueSelection; showBlanks: boolean }
export type ColumnRule = ValuesFilter | TypedCondition | CompoundFilter
export interface FilterSpec { range: FilterRange; columns: ReadonlyMap<SourceCol, ColumnRule> }
export interface PagedKeySetRef { refId: string; count: number; owner: SourceStamp }
export interface ValueSearchPredicate { text: string; locale: string }
export type ValueSelection =
  | { mode: 'include'; keys: PagedKeySetRef }
  | { mode: 'allExcept'; excluded: PagedKeySetRef }
  | { mode: 'queryPatches'; base: 'all' | 'none'; patches: readonly ValuePatch[] }
export interface ValuePatch {
  domain: { kind: 'keys'; keys: PagedKeySetRef } | { kind: 'search'; query: ValueSearchPredicate; owner: SourceStamp }
  selected: boolean
}

/** Float64 logical geometry and bounded visible runs, spec §§5.3–5.4. */
export interface AxisSizeInterval { start: number; end: number; sizeCss: number }
export interface AxisRun { start: number; end: number; startCss: number; sizeCss: number }
export interface AxisWindow {
  runs: Iterable<AxisRun>
  count: number
  nextCursor?: string
}
export interface LogicalCellBox { x: number; y: number; width: number; height: number }
export interface AxisMetrics {
  logicalCount: number
  defaultSizeCss: number
  overrides: readonly AxisSizeInterval[]
  totalCss(): number
  start(index: number): number
  size(index: number): number
  locate(logicalCss: number): number | null
  visibleBetween(startCss: number, endCss: number): AxisWindow
}
export interface SheetMetrics {
  rows: AxisMetrics
  cols: AxisMetrics
  order: RowOrder
  sourceCellToLogical(at: SourceCell): LogicalCellBox | null
  logicalPointToSource(x: number, y: number): SourceCell | null
}
export const MAX_PHYSICAL_SCROLL_CSS_PX = 1_000_000
/** Shape only: physical extents stay bounded after zoom/device-scale changes (§5.4). */
export interface ScrollWindow {
  logicalOrigin: { x: number; y: number }
  physicalPosition: { x: number; y: number }
  physicalExtent: { width: number; height: number }
  zoom: number
  deviceScale: number
}
export interface LogicalViewport extends LogicalCellBox {}

/** Every asynchronous publication captures the relevant complete owner, spec §10.1. */
export interface CacheIdentity { docId: string; sheetId: string; valueGen: number; fontGen: number; viewGen: number }
export interface WorkbookOwner {
  docId: string
  globalViewGen: number
  workbookValueGen: number
  fontGen: number
  sheetVector: readonly { sheetId: string; revision: number; valueGen: number }[]
}
/** Open-time source descriptors only; subsequent commands use resident data (§§4.2, 10.2). */
export type SourceHandle =
  | { kind: 'blob'; blob: Blob }
  | { kind: 'paged'; handleId: string }
  | { kind: 'object'; document: XlsxDocument }
export type ProfileId = 'D' | 'M'
export interface SearchQuery { text: string; caseSensitive: boolean; sheetIds?: readonly string[] }
export type WorkerCommand =
  | { kind: 'open'; requestId: string; docId: string; source: SourceHandle; profile: ProfileId }
  | { kind: 'activateSheet'; requestId: string; owner: WorkbookOwner; sheetId: string }
  | { kind: 'prepareColumns'; requestId: string; owner: CacheIdentity; cols: Uint16Array }
  | { kind: 'sort'; requestId: string; owner: CacheIdentity; spec: SortSpec | null }
  | { kind: 'filter'; requestId: string; owner: CacheIdentity; spec: FilterSpec | null; nowMs: number }
  | { kind: 'viewport'; requestId: string; owner: CacheIdentity; window: LogicalViewport }
  | { kind: 'valueQuery'; requestId: string; owner: CacheIdentity; query: ColumnQuery }
  | { kind: 'search'; requestId: string; owner: WorkbookOwner; query: SearchQuery; cursor?: string }
  | { kind: 'cancel'; requestId: string; targetRequestId: string }
  | { kind: 'close'; requestId: string; docId: string }

/** Transferable compact deltas, never method-bearing metrics/masks (§10.2). */
export interface ViewCommitDelta {
  bands: readonly RowPermutationBand[]
  visibility: readonly { r0: SourceRow; r1: SourceRow; words: Uint32Array }[]
  metrics: { rows: readonly AxisSizeInterval[]; cols: readonly AxisSizeInterval[] }
  counts: { visible: number; total: number }
  selection: GridSelection
  owner: CacheIdentity
  workbookOwner: WorkbookOwner
}
export interface SearchHit { sheetId: string; cellId: CellId; source: SourceCell; text: string }
export interface WorkerDiagnostic { code: string; message: string; source?: SourceCell; partId?: string; pageId?: string }
export type WorkerOwner = CacheIdentity | WorkbookOwner | { docId: string }
/** DTO v1: every reply carries its captured request owner (§10.2).
 * Progress/failures use the command's owner; completion narrows by result scope. */
export interface WorkerReplyBase<Owner extends WorkerOwner = WorkerOwner> {
  requestId: string
  owner: Owner
  diagnostics: readonly WorkerDiagnostic[]
}
export type WorkerResult =
  | { kind: 'open'; valuePolicy: ValuePolicy; sheets: readonly { sheetId: string; name: string }[] }
  | { kind: 'activateSheet'; sheetId: string }
  | { kind: 'prepareColumns'; cols: Uint16Array }
  | { kind: 'view'; delta: ViewCommitDelta }
  | { kind: 'viewport'; cells: readonly { source: SourceCell; bounds: LogicalCellBox }[] }
  | { kind: 'valueQuery'; page: Page<ColumnEntry> }
  | { kind: 'search'; items: readonly SearchHit[]; nextCursor?: string; unavailableCount: number }
  | { kind: 'close' }
/** A document ID alone cannot own a sheet/workbook publication (§§10.1–10.2). */
export type WorkerCompletion =
  | (WorkerReplyBase<{ docId: string }> & {
    kind: 'complete'
    result: Extract<WorkerResult, { kind: 'open' | 'close' }>
  })
  | (WorkerReplyBase<CacheIdentity> & {
    kind: 'complete'
    result: Extract<WorkerResult, { kind: 'prepareColumns' | 'view' | 'viewport' | 'valueQuery' }>
  })
  | (WorkerReplyBase<WorkbookOwner> & {
    kind: 'complete'
    result: Extract<WorkerResult, { kind: 'activateSheet' | 'search' }>
  })
export type WorkerReply = WorkerCompletion | (WorkerReplyBase & (
  | { kind: 'progress'; completed: number; total?: number }
  | { kind: 'cancelled' }
  | { kind: 'incomplete'; missing: readonly SourceRect[]; refusal?: OperationRefusal }
  | { kind: 'quota'; resource: 'memory' | 'storage'; limitBytes?: number; requiredBytes?: number }
  | { kind: 'error'; code: SourceErrorCode; retryable: boolean; refusal?: OperationRefusal }
))
