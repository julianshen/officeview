import type { XlsxCell, XlsxCellStyle, XlsxDocument, XlsxDrawing, XlsxImage } from '../types'
import type { ThemeContext } from '../../drawing/style'
import type { RangeDetection } from './rangeDetection'
import type { FindPage } from './find'
import type { CellId, DisplaySlot, SourceCell, SourceCol, SourceRect, SourceRow } from './coordinates'
import type { SparklineDiagnostic, XlsxSparklineGroup } from '../sparklines/model'
import type { TableMetadata } from '../formula/types'
import type { TableStyleDef, TableStyleDiagnostic } from '../tableStyles/model'

/** Source-space contracts, spec §4.1; limits bound output, never semantic membership. */
export const MAX_OVERVIEW_GROUPS_DESKTOP = 2048
export const MAX_OVERVIEW_GROUPS_MOBILE = 1024
export type Completeness = 'complete' | 'loading' | 'incomplete'
export const OPERATION_REFUSED_CODE = 'operation-refused'
/** Shared source/worker error-code registry, orthogonal to reply kinds (§§4.1, 10.2).
 * Quota/cancelled/incomplete codes preserve source failures in an error reply;
 * operation-refused denotes the guarded worker refusal with its required payload. */
export const SOURCE_ERROR_CODES = [
  'invalid-coordinate', 'invalid-source-rect', 'invalid-cursor', 'stale-source',
  'cancelled', 'incomplete', 'quota', 'resource-limit', 'unsupported',
  'parse-error', 'storage-error', 'page-error', OPERATION_REFUSED_CODE,
] as const
export type SourceErrorCode = typeof SOURCE_ERROR_CODES[number]

/** A policy is fixed at document admission, spec §4.4. */
export type ValuePolicy = 'engine' | 'cached'
export interface AuthoritativeCell extends Pick<XlsxCell,
  'value' | 'hasCachedValue' | 'valueIsError' | 'formula' | 'styleIndex'
  | 'sharedFormula' | 'arrayRef' | 'dynamicArray' | 'formulaType' | 'ca'> {
  id: CellId
  row: SourceRow
  col: SourceCol
  valueProvenance?: 'engine' | 'cached'
  /** Explicit authority state in rows/scans; never infer from cache flags in engine mode (§4.4). */
  unavailable?: true
}
/** Single-cell payloads use the enclosing discriminant, never the bulk marker (§4.4).
 * Omit alone accepts existing objects with extra properties; optional never also
 * excludes unavailable:true on an already-typed AuthoritativeCell. */
export type SingleCellPayload = Omit<AuthoritativeCell, 'unavailable'> & { unavailable?: never }
/** readCell returns kind='unavailable' for permanent missing values; unknown is
 * a pending page, never blank. Raw readRows/scanPopulated and future column lanes
 * use AuthoritativeCell.unavailable; current ColumnEntry uses its discriminant. */
export type CellRead =
  | { kind: 'cell'; cell: SingleCellPayload }
  | { kind: 'blank' }
  | { kind: 'unknown'; pageId: string }
  | { kind: 'error'; code: SourceErrorCode; retryable: boolean }
  | { kind: 'unavailable'; cell: SingleCellPayload; reason: 'formula-value-unavailable' }

export interface SourceStamp { docId: string; sheetId: string; valueGen: number }
export interface ReadOptions { signal: AbortSignal; stamp: SourceStamp }
/** Cursors belong to the original query/stamp; owner changes invalidate them (§4.1).
 * readRows/queryColumn/scanPopulated/queryOverview publish expected operational
 * failures here (including partial items), with completeness='incomplete'. They
 * must not disguise a failed query as a complete empty page or an untyped throw. */
export interface Page<T> {
  items: readonly T[]
  nextCursor?: string
  completeness: Completeness
  stamp: SourceStamp
  error?: SourceFailure
}
export interface SourceRowInterval { r0: SourceRow; r1: SourceRow }
export interface SourceColInterval { c0: SourceCol; c1: SourceCol }

/** Shared bounded mask encoding (§§4.1, 5.1, 8.1): inclusive [r0,r1],
 * words.length=ceil((r1-r0+1)/32). Bit i (LSB-first within each Uint32 word)
 * represents row r0+i. Padding bits are zero; count equals the popcount. */
export interface BoundedRowBitset<Space extends 'source' | 'display' = 'source'> {
  kind: 'bitset'
  /** Explicit space tag preserves type/wire identity even though both aliases are numbers. */
  space: Space
  r0: Space extends 'source' ? SourceRow : DisplaySlot
  r1: Space extends 'source' ? SourceRow : DisplaySlot
  words: Uint32Array
  count: number
}
/** Exact source membership (§§4.1, 8.1), never display order or enclosing bounds.
 * Intervals are ascending, non-overlapping, with inclusive r1. Lists hold ABSOLUTE
 * source rows, strictly ascending/unique, all inside inclusive [r0,r1]. Bitsets
 * use BoundedRowBitset's encoding in source space. Refs stay resident in the worker;
 * stamp is the source/value owner, the same role as PagedKeySetRef.owner (§7.4). */
export type CompactSourceRowSet =
  | { kind: 'intervals'; intervals: readonly SourceRowInterval[] }
  | { kind: 'list'; r0: SourceRow; r1: SourceRow; rows: Uint32Array }
  | BoundedRowBitset<'source'>
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
/** Raw source aggregates have no viewport geometry or sampled text (§4.1). */
export interface OverviewAggregate {
  groupId: string
  populatedCount: number
  countKind: 'exact' | 'estimated'
  style: { kind: 'uniform'; styleIndex: number } | { kind: 'mixed' }
}
/** Each unloaded group is explicit, without guessed counts/styles (§4.1, plan S5.1). */
export type OverviewGroup =
  | (OverviewAggregate & ({ content: 'blank' } | { content: 'populated' } | { content: 'mixed' }))
  | { groupId: string; content: 'unknown'; pageId: string }
/** Sparse response item only (§4.1): range belongs to readRows, continuation to
 * Page.nextCursor. Adapter/profile page bounds apply even without a caller limit. */
export interface SparseRowPage {
  row: SourceRow
  cells: readonly AuthoritativeCell[]
}
export interface ColumnQuery {
  col: SourceCol
  rows: CompactSourceRowSet
  /** Entries per page; for 'lanes' it bounds the rows one lanes entry covers. */
  limit: number
  cursor?: string
  /** 'lanes' asks for typed-lane entries (§4.1); default 'entries' never yields them. */
  encoding?: 'entries' | 'lanes'
}
/** Type lane tags of a ColumnLanes entry. The numbers are shared with the
 * S8.2 value index (SOURCE_* in valueIndex.ts), so lanes copy straight in. */
export const LANE_NUMBER = 1
export const LANE_STRING = 2
export const LANE_BOOLEAN = 3
export const LANE_ERROR = 4
/** A present but value-less record (style-only cell); still blank for values. */
export const LANE_BLANK = 5
export const LANE_UNAVAILABLE = 7
/** Typed lanes over inclusive [r0,r1] (§4.1): rows listed in `rows` (strictly
 * ascending, inside [r0,r1]) carry one tag/number/style each; every other row
 * of the run is blank (the default-blank run). Strings and error texts are raw
 * UTF-16 code units, `text[textStart[i]..textStart[i+1])`, so lone surrogates
 * survive; numbers hold finite numbers and booleans as 1/0. Requested only via
 * ColumnQuery.encoding='lanes'; formulas stay on readCell/readRows. */
export interface ColumnLanes {
  kind: 'lanes'
  r0: SourceRow
  r1: SourceRow
  rows: Uint32Array
  tags: Uint8Array
  numbers: Float64Array
  styles: Uint32Array
  text: Uint16Array
  textStart: Uint32Array
}
export type ColumnEntry =
  | { kind: 'cell'; row: SourceRow; cell: SingleCellPayload }
  | { kind: 'blank'; r0: SourceRow; r1: SourceRow }
  | { kind: 'unknown'; r0: SourceRow; r1: SourceRow; pageId: string }
  | { kind: 'unavailable'; row: SourceRow; cell: SingleCellPayload; reason: 'formula-value-unavailable' }
  | { kind: 'error'; r0: SourceRow; r1: SourceRow; code: SourceErrorCode; retryable: boolean }
  | ColumnLanes
export interface PopulatedQuery {
  range: SourceRect; limit: number; cursor?: string
  /** Snapshot the loaded frontier; terminate with loading instead of waiting for import. */
  availability?: 'loaded'
}
export interface SourceFailure { code: SourceErrorCode; retryable: boolean; message: string }
export interface RangeCompletion {
  range: SourceRect
  completeness: Completeness
  stamp: SourceStamp
  missing: readonly SourceRect[]
  error?: SourceFailure
}
/** Pacing contract (spec §10.2 provider request decode, decision 10): a
 * provider decodes a scanPopulated/scanOccupied or readRows request in steps of
 * at most YIELD_WORK_UNITS cells with a real macrotask yield between steps, and
 * checks the request's signal and source owner at checkpoints at most
 * CHECK_WORK_UNITS cells apart and after every yield, ending the request with a
 * cancelled or stale-source page (the shipped stores use requestBudget). The
 * request's limit and the number of requests are unchanged. queryColumn is
 * bounded by its limit. External custom providers must meet the same contract. */
export interface SourceStore {
  readonly stamp: SourceStamp
  readonly valuePolicy: ValuePolicy
  readonly extent: { rows: 1048576; cols: 16384; used?: SourceRect }
  readCell(at: SourceCell, options: ReadOptions): Promise<CellRead>
  readRows(range: SourceRect, cursor: string | undefined, options: ReadOptions): Promise<Page<SparseRowPage>>
  /** Ordering contract (§4.1; enforced by the worker session since S4.2): the
   * items of a page, and of successive pages of one query, cover the requested
   * membership in ASCENDING source-row order, contiguously and without overlap
   * or duplicates (each cell/unavailable entry is one row; blank/unknown/error
   * runs are inclusive [r0,r1]). Consumers binary-search lanes built from them;
   * a page that violates this is rejected with a nonretryable `page-error`. */
  queryColumn(request: ColumnQuery, options: ReadOptions): Promise<Page<ColumnEntry>>
  scanPopulated(request: PopulatedQuery, options: ReadOptions): AsyncIterable<Page<AuthoritativeCell>>
  /** Optional (S10.2): exactly scanPopulated's cells, order and page boundaries,
   * coordinates only, so occupancy probes (range detection) skip payload decode. */
  scanOccupied?(request: PopulatedQuery, options: ReadOptions): AsyncIterable<Page<SourceCell>>
  queryStructural(rect: SourceRect, options: ReadOptions): Promise<StructuralResult>
  /** Complete sources terminate with a complete coarse result or explicit resource/error (§4.1). */
  queryOverview(request: OverviewQuery, options: ReadOptions): Promise<Page<OverviewGroup>>
  ensureComplete(rect: SourceRect, options: ReadOptions): Promise<RangeCompletion>
  /** Optional (S10.2 main ledger): the consumer has decoded `payload` (a page
   * this store returned) into its own charged structure, or dropped it. A store
   * whose payloads are charged while reachable ends that charge now. Call it
   * only after the owner's own charge is in place. */
  settle?(payload: object): void
  release(): Promise<void>
  /** S13.2: the sheet's drawing content for painting (shapes, text, theme and
   * pictures), bounded by the drawing transfer budget of `profile`. Optional:
   * a store without it reports drawings as unsupported (a visible diagnostic).
   * `memory` is the caller's transient share of the document pool: the store
   * sizes it to its working set and refuses (a diagnostic) what does not fit. */
  readDrawings?(options: ReadOptions, profile: ProfileId, memory?: DrawingMemory): Promise<SheetDrawings>
  /** S3: resolves this sheet's sparkline groups from worker-owned cell stores. */
  readSparklines?(options: ReadOptions, grantBytes: number, memory?: DrawingMemory, replyEnvelopeBytes?: number): Promise<SheetSparklines>
  /** S8: bounded table-style descriptors for the receiving region-style layer. */
  readRegionStyles?(options: ReadOptions, grantBytes: number, memory?: DrawingMemory): Promise<RemoteRegionStyles>
}
/** Absolute transient bytes of one drawing read; false when they cannot be admitted. */
export interface DrawingMemory { set(bytes: number): boolean }

/** Clone-safe, render-ready sparkline transfer. NaN encodes a null point. */
export interface TransferredSparkline {
  host: SourceCell
  dataSheet: string
  sourceRows: number[]
  groupIndex: number
  values: Float64Array
  bridged: number[]
  firstIndex: number
  lastIndex: number
  highIndices: number[]
  lowIndices: number[]
  min: number
  max: number
  axis?: { positions: Float64Array; slot: number }
}
export interface SheetSparklines {
  sparklines: TransferredSparkline[]
  groups: XlsxSparklineGroup[]
  diagnostics: SparklineDiagnostic[]
  /** Referenced sheet pages that were still loading when this snapshot ran. */
  pendingSheetIds?: string[]
}

/** Worker-resolved table style definitions, bounded by the receiver's grant. */
export interface RemoteRegionStyles {
  sheetId: string
  stamp: SourceStamp
  tables: readonly { table: Pick<TableMetadata, 'id' | 'displayName' | 'sheetId' | 'extent' | 'headerRowCount' | 'totalsRowCount' | 'styleInfo'>; def: TableStyleDef }[]
  theme?: ThemeContext
  /** Reserved for future per-cell pivot region styles; empty for tables. */
  cells?: Int32Array
  diagnostics: readonly TableStyleDiagnostic[]
}

/** S13.2: one sheet's paintable drawings as the resident parser produces them,
 * with sheet-local picture indexes. Whatever the budget or a part cap kept out
 * is named in `diagnostics` (never silently omitted); `bytes` are the binary
 * asset bytes carried. An incomplete result carries the failure. */
export interface SheetDrawings {
  completeness: 'complete' | 'incomplete'
  drawings: readonly XlsxDrawing[]
  /** Indexed by each drawing's `imageIndex`. */
  images: readonly XlsxImage[]
  theme?: ThemeContext
  diagnostics: readonly WorkerDiagnostic[]
  bytes: number
  error?: SourceFailure
  /** Sheets whose cells a chart reads but whose import has not reached them yet:
   * the payload is not cached and is read again once they complete. */
  pendingSheets?: readonly string[]
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
  /** The structural source itself failed; distinct from waiting on missing ranges. */
  | { reason: 'source-error'; code: SourceErrorCode; retryable: boolean; message: string }
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
  /** sourceAtOffset[d] is the ABSOLUTE source row at band display offset d (§5.1). */
  sourceAtOffset: Uint32Array
  /** displayOffsetOfSource[s-r0] is source row s's display offset; both arrays
   * have exactly r1-r0+1 entries and are inverses (§5.1). */
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
  /** Serializable current gesture; preserves the Ctrl prefix across view copies. */
  gesture?: GridSelectionGesture
}
export interface GridSelectionGesture {
  kind: 'cells' | 'row' | 'column'
  anchor: SourceCell
  /** With `part`: only the prefix cells hidden under `part`; the full prefix is
   * `ranges \ part` plus these. Without `part` (legacy): the full prefix. */
  prefix: readonly GridSelectionPart[]
  /** The part the current gesture applied over the prefix. */
  part?: GridSelectionPart
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
/** owner names the same source/value ownership role as a row-set ref stamp (§7.4). */
export interface PagedKeySetRef { refId: string; count: number; owner: SourceStamp }
/** Dropdown value search: contains, optionally with S8.1 `*`/`?`/`~` wildcards. */
export interface ValueSearchPredicate { text: string; locale: string; wildcards?: boolean }
export type ValueSelection =
  | { mode: 'include'; keys: PagedKeySetRef }
  | { mode: 'allExcept'; excluded: PagedKeySetRef }
  | { mode: 'queryPatches'; base: 'all' | 'none'; patches: readonly ValuePatch[] }
export interface ValuePatch {
  domain: { kind: 'keys'; keys: PagedKeySetRef } | { kind: 'search'; query: ValueSearchPredicate; owner: SourceStamp }
  selected: boolean
}

/** Float64 logical geometry; start/end are inclusive axis indexes (§§5.3–5.4). */
export interface AxisSizeInterval { start: number; end: number; sizeCss: number }
/** A uniform run over inclusive axis indexes [start,end], with logical CSS origin. */
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
  zoom: SheetZoom
  deviceScale: number
  /** Authored frozen panes: requested counts and the logical extents pinned at
   * this zoom/viewport (0 = that axis is not pinned). See frozenPanes.ts. */
  frozen?: { rows: number; cols: number; left: number; top: number }
}
export const SHEET_MIN_ZOOM = 0.25
export const SHEET_MAX_ZOOM = 4
/** Absolute sheet-mode CSS scale: 1=100%, finite in [0.25,4] at admission (§5.5).
 * Runtime zoom validation belongs to S2.4; page fit-width zoom is a different policy. */
export type SheetZoom = number
/** A logical CSS box plus the exact projection used for detail admission (§§5.5, 11.1). */
export interface LogicalViewport {
  box: LogicalCellBox
  zoom: SheetZoom
  deviceScale: number
}
export type ViewportCell = CellRead & { source: SourceCell; bounds: LogicalCellBox }
export type ViewportOverviewGroup = OverviewGroup & { bounds: LogicalCellBox }
export type ViewportContent =
  | { kind: 'detail'; cells: readonly ViewportCell[] }
  | { kind: 'overview'; groups: readonly ViewportOverviewGroup[] }
  | { kind: 'unknown'; pageId: string }
/** Completed viewports contain only resolved detail/coarse records (§§4.1, 10.4).
 * Permanent unavailable cells are resolved; loading/error records are not. */
export type CompleteViewportContent =
  | { kind: 'detail'; cells: readonly Exclude<ViewportCell, { kind: 'unknown' | 'error' }>[] }
  | { kind: 'overview'; groups: readonly Exclude<ViewportOverviewGroup, { content: 'unknown' }>[] }
export type ViewportResult = (
  | { kind: 'viewport'; completeness: 'complete'; content: CompleteViewportContent; error?: never }
  | { kind: 'viewport'; completeness: 'loading' | 'incomplete'; content: ViewportContent; error?: SourceFailure }
) & {
  /** Request-local overflow evidence; never a row-cache completeness proof. */
  overflowProof?: import('./overflowProof').OverflowProof
  /** queryPayloadBytes of this result with its diagnostics, measured by the
   * worker's paced walker (spec §10.2), so main checks the bound without a walk. */
  queryBytes?: number
  /** Mixed detail/coarse panes and exact active inspection share the same owner. */
  groups?: readonly ViewportOverviewGroup[]
  active?: ViewportCell
  merges?: import('./viewport').StoreViewport['merges']
  coarse?: import('./viewport').StoreViewport['coarse']
  rowLines?: readonly number[]
  colLines?: readonly number[]
  work?: import('./viewport').ViewportWork
}

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
/** S5.4 bounded source reads of worker-owned (paged) sheets: one store call per
 * command, owned by the sheet's current committed CacheIdentity. `rows` serves
 * viewport records and copy reads, `scan` find paging (`cursor` continues the
 * worker-resident scan the previous page opened), `cell` Go To/inspection. */
export type SourceReadQuery =
  | { op: 'cell'; at: SourceCell }
  | { op: 'rows'; range: SourceRect; cursor?: string }
  | { op: 'scan'; request: PopulatedQuery; cursor?: string }
  | { op: 'structural'; rect: SourceRect }
  | { op: 'complete'; rect: SourceRect }
  | { op: 'overview'; request: OverviewQuery }
  /** Document-owned cleanup of an abandoned worker search continuation. */
  | { op: 'releaseSearch'; cursor: string }
  /** The sheet's current compact layout (it grows while a paged import runs). */
  | { op: 'layout' }
  /** Ends a scan continuation the caller abandoned (iterator return/cancel). */
  | { op: 'releaseScan'; cursor: string }
  /** S13.2: the sheet's drawing content (SourceStore.readDrawings). */
  /** `replyByteLimit` (internal, FC-T2): the receiving grant main reserved for
   * the reply, in estimatePayloadBytes units over the WHOLE final reply. The
   * worker replaces a reply over it with a payload-free quota reply. Omitted:
   * no grant (legacy callers); every RemoteWorkbook dispatch sets it. */
  | { op: 'drawings'; replyByteLimit?: number }
  /** S3: receiver-granted sparkline resolution; sheet id is on the read owner. */
  | { op: 'sparklines'; grantBytes: number }
  /** S8: receiver-granted table region styles. */
  | { op: 'regionStyles'; grantBytes: number }
  /** S10.2: operation range detection (sort admission includes the structural
   * guard) run next to the store, so a paged sheet is not scanned strip by
   * strip across the worker boundary. Same detectors, same outcome. */
  | { op: 'detectRange'; kind: 'sort' | 'filter'; selection: GridSelection; override?: { headerRows: 0 | 1; totalsRows?: 0 } | null }
export type SourceReadResult =
  | { op: 'cell'; read: CellRead }
  | { op: 'rows'; page: Page<SparseRowPage> }
  | { op: 'scan'; page: Page<AuthoritativeCell> }
  | { op: 'structural'; result: StructuralResult }
  | { op: 'complete'; result: RangeCompletion }
  | { op: 'overview'; page: Page<OverviewGroup> }
  | { op: 'releaseSearch'; released: boolean }
  | { op: 'layout'; layout: SheetLayout }
  | { op: 'releaseScan'; released: boolean }
  | { op: 'drawings'; result: SheetDrawings }
  | { op: 'sparklines'; result: SheetSparklines }
  | { op: 'regionStyles'; result: RemoteRegionStyles }
  | { op: 'detectRange'; detection: RangeDetection }
/** S5.4: resolved cell styles (indexed by cell styleIndex) and the date system
 * of a paged workbook, published with open so worker-backed cells format like
 * resident ones. Bounded by the styles part's small-part cap. */
export interface WorkbookFormatting {
  dateSystem: '1900' | '1904'
  styles: readonly XlsxCellStyle[]
  /** The default font's maximum digit width (column-width model); absent: 7px. */
  maxDigitWidthPx?: number
}
/** S5.4: committed value-list keys of a `values` rule, shipped with a worker
 * filter command because the dropdown's key store lives on the main thread.
 * `scope` is the exact value domain (valueScope); bounded by that store. */
export interface FilterKeySet {
  refId: string
  scope: string
  run: { count: number; kinds: Uint8Array; numbers: Float64Array; texts: readonly (string | undefined)[] }
}
/** Compact source-space layout published with an activated sheet (S5.4): the
 * main thread builds the same axis metrics and base visibility as the worker. */
export interface SheetLayout {
  rowOverrides: readonly AxisSizeInterval[]
  colOverrides: readonly AxisSizeInterval[]
  hidden: readonly SourceRowInterval[]
  /** Authored-hidden columns: zero width for paint, navigation, find and copy. */
  hiddenCols?: readonly { start: number; end: number }[]
  used?: SourceRect
  /** S3: metadata-only fast path; false lets sheets without sparklines avoid
   * any per-frame preparation or worker request. */
  hasSparklines?: boolean
  /** S8: table metadata fast path; absent/false has no per-frame style work. */
  hasRegionStyles?: boolean
  /** Authored frozen panes. */
  frozen?: { rows: number; cols: number; topRow?: number; leftCol?: number }
}
export type WorkerCommand =
  | { kind: 'open'; requestId: string; docId: string; source: SourceHandle; profile: ProfileId }
  | { kind: 'activateSheet'; requestId: string; owner: WorkbookOwner; sheetId: string }
  | { kind: 'prepareColumns'; requestId: string; owner: CacheIdentity; cols: Uint16Array }
  | { kind: 'sort'; requestId: string; owner: CacheIdentity; spec: SortSpec | null }
  | { kind: 'filter'; requestId: string; owner: CacheIdentity; spec: FilterSpec | null; nowMs: number; keySets?: readonly FilterKeySet[] }
  | { kind: 'viewport'; requestId: string; owner: CacheIdentity; window: LogicalViewport; scope?: 'frame' | 'tile' }
  | { kind: 'valueQuery'; requestId: string; owner: CacheIdentity; query: ColumnQuery }
  | { kind: 'search'; requestId: string; owner: WorkbookOwner; query: SearchQuery; cursor?: string }
  | { kind: 'read'; requestId: string; owner: CacheIdentity; query: SourceReadQuery }
  | { kind: 'cancel'; requestId: string; targetRequestId: string }
  | { kind: 'close'; requestId: string; docId: string }

/** Transferable compact deltas, never method-bearing metrics/masks (§10.2). */
export interface ViewCommitDelta {
  bands: readonly RowPermutationBand[]
  /** Same bounded mask encoding, but indexes are DISPLAY slots (§5.1). */
  visibility: readonly BoundedRowBitset<'display'>[]
  metrics: { rows: readonly AxisSizeInterval[]; cols: readonly AxisSizeInterval[] }
  counts: { visible: number; total: number }
  selection: GridSelection
  owner: CacheIdentity
  workbookOwner: WorkbookOwner
}
export interface SearchHit { sheetId: string; cellId: CellId; source: SourceCell; text: string }
export interface WorkerDiagnostic { code: string; message: string; source?: SourceCell; partId?: string; pageId?: string }
export type WorkerOwner = CacheIdentity | WorkbookOwner | { docId: string }
export type WorkerCommandKind = WorkerCommand['kind']
/** Open/close capture document context. Cancel captures the target request's
 * owner from resident state; every owner-bearing command preserves its scope (§10.2). */
export type WorkerCommandOwner<K extends WorkerCommandKind> =
  Extract<WorkerCommand, { kind: K }> extends { owner: infer Owner extends WorkerOwner }
    ? Owner : K extends 'cancel' ? WorkerOwner : { docId: string }
/** Freeze addition to §10.2: commandKind echoes the originating command kind
 * for owner correlation on ALL replies. S4.1's wire format MUST populate it. */
export type WorkerReplyBase<K extends WorkerCommandKind> =
  K extends WorkerCommandKind ? {
    commandKind: K
    requestId: string
    owner: WorkerCommandOwner<K>
    diagnostics: readonly WorkerDiagnostic[]
    /** FC-T2 (internal): on a reply to a granted drawings read, its exact
     * measured size (this field included), never over the grant. */
    replyBytes?: number
  } : never
export type WorkerResult =
  | { kind: 'open'; valuePolicy: ValuePolicy; sheets: readonly { sheetId: string; name: string }[]; formatting?: WorkbookFormatting }
  | { kind: 'activateSheet'; sheetId: string; layout?: SheetLayout }
  | { kind: 'prepareColumns'; cols: Uint16Array }
  | { kind: 'view'; delta: ViewCommitDelta }
  | ViewportResult
  | { kind: 'valueQuery'; page: Page<ColumnEntry> }
  | { kind: 'search'; items: readonly SearchHit[]; nextCursor?: string; unavailableCount: number; page?: FindPage }
  | { kind: 'read'; result: SourceReadResult }
  | { kind: 'close' }
/** Cancel is non-completing: it terminates with cancelled/error, not complete (§10.2).
 * Sort/filter publish one view delta; other completions match their command. */
export type WorkerResultFor<K extends WorkerCommandKind> =
  K extends 'cancel' ? never : Extract<WorkerResult, { kind: K extends 'sort' | 'filter' ? 'view' : K }>
/** Refusals use the existing error reply with the meaningful operation-refused
 * transport code (§6.3). No missing range or source failure is invented. The
 * captured owner is unchanged; emit one status/diagnostic, not a view completion. */
export type WorkerReply<K extends WorkerCommandKind = WorkerCommandKind> =
  K extends WorkerCommandKind ? WorkerReplyBase<K> & (
    | { kind: 'progress'; completed: number; total?: number }
    | (K extends 'cancel' ? never : { kind: 'complete'; result: WorkerResultFor<K> })
    | { kind: 'cancelled' }
    | { kind: 'incomplete'; missing: readonly SourceRect[] }
    | { kind: 'quota'; resource: 'memory' | 'storage'; limitBytes?: number; requiredBytes?: number }
    | { kind: 'error'; code: Exclude<SourceErrorCode, typeof OPERATION_REFUSED_CODE>; retryable: boolean; refusal?: never }
    | (K extends 'sort' | 'filter' ? {
      kind: 'error'; code: typeof OPERATION_REFUSED_CODE; retryable: false; refusal: OperationRefusal
    } : never)
  ) : never
export type WorkerCompletion = Extract<WorkerReply, { kind: 'complete' }>
