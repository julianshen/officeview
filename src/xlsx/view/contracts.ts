import type { XlsxCell, XlsxCellStyle, XlsxDocument, XlsxDrawing, XlsxImage } from '../types'
import type { ThemeContext } from '../../drawing/style'
import type { RangeDetection } from './rangeDetection'
import type { FindPage } from './find'
import type { CellId, DisplaySlot, SourceCell, SourceCol, SourceRect, SourceRow } from './coordinates'
import type { SparklineDiagnostic, XlsxSparklineGroup } from '../sparklines/model'
import type { TableMetadata } from '../formula/types'
import type { DxfStyle, TableStyleDef, TableStyleDiagnostic } from '../tableStyles/model'
import type { PivotTableDefinition, PivotDiagnostic } from '../pivots/model'
import type { PivotDelta } from '../pivots/session/delta'
import type { ChartModel } from '../../charts/model'
import type { CfMemory, CfWorkBudget } from '../cf/limits'

/** Shared bounds for paged table-style descriptors and decoded role caches. */
/** The main-ledger grant a pivot open or apply reserves for its reply's fit cells (Task 29): the worker stops the
 * cells within it (none when the grant is 0). */
export const PIVOT_FIT_REPLY_BYTES = 4 * 1024 * 1024
/** Shared bounded fan-out limits for slicer/timeline protocol, worker and adapters. */
export const MAX_LINKED_CONTROL_TARGETS = 8
export const MAX_LINKED_CONTROL_ITEMS = 10_000
/** Linked apply's reply gets a bounded main-ledger reservation; the worker grant is partitioned 1/3 for that reply. */
export const LINKED_CONTROL_REPLY_BYTES = 4 * 1024 * 1024
export const LINKED_CONTROL_WORKER_GRANT_BYTES = LINKED_CONTROL_REPLY_BYTES * 3
export const MAX_REGION_STYLE_TABLES = 512
/** Largest built-in table-style descriptor measured across all 144 table/pivot
 * names: 25,154 B excluding the shared workbook theme; 32 KiB leaves 1.3x headroom. */
export const REGION_STYLE_DESCRIPTOR_MAX_BYTES_PER_TABLE = 32_768
export const REGION_STYLE_REPLY_OVERHEAD_BYTES = 1_024
/** Retained custom styles parsed from styles.xml are capped independently of source size. */
export const MAX_PAGED_TABLE_STYLE_DEFS_BYTES = 2 * 1024 * 1024
/** Distinct `formats` dxfs one pivot-display reply carries (styles.xml holds at most MAX_DXFS = 65,536). */
export const MAX_PIVOT_DISPLAY_DXFS = 4_096
export const MAX_REGION_STYLE_MEMO_ENTRIES = 4_096
/** A pivot with `formats` keeps its overrides for at most this many cells (pivots/regions.ts), at this estimate each;
 * the receiver charges the allowance once per formatted pivot. */
export const PIVOT_FORMAT_MEMO_ENTRIES = 4_096
export const PIVOT_FORMAT_MEMO_ENTRY_ESTIMATE_BYTES = 128
export const PIVOT_FORMAT_MEMO_BYTES = PIVOT_FORMAT_MEMO_ENTRIES * PIVOT_FORMAT_MEMO_ENTRY_ESTIMATE_BYTES
export const REGION_STYLE_MEMO_ENTRY_ESTIMATE_BYTES = 1_024
export const REGION_STYLE_LAYER_BASE_BYTES = 4_096
export const REGION_STYLE_REGION_ESTIMATE_BYTES = 512
export const REGION_STYLE_LAYER_MAX_BYTES = REGION_STYLE_LAYER_BASE_BYTES
  + MAX_REGION_STYLE_TABLES * REGION_STYLE_REGION_ESTIMATE_BYTES
  + MAX_REGION_STYLE_MEMO_ENTRIES * REGION_STYLE_MEMO_ENTRY_ESTIMATE_BYTES
/** The fixed part of a decoded region layer's charge (receiver and worker Find alike): the shared style
 * memo, each table/pivot region and pivot style, and each formatted pivot's own overrides memo
 * (pivots/regions.ts). A pivot's saved layout is added by tableStyles/remoteLayer.ts. */
export function regionLayerFixedBytes(payload: RemoteRegionStyles | RemotePivotDisplay): number {
  const count = 'pivotTables' in payload ? payload.pivotTables.length + payload.pivotStyles.length : payload.tables.length
  if (!count) return 0
  const formatted = 'pivotTables' in payload ? payload.pivotTables.filter(def => def.formats.some(format => format.action === 'formatting' && format.dxfId !== undefined)).length : 0
  return REGION_STYLE_LAYER_BASE_BYTES + count * REGION_STYLE_REGION_ESTIMATE_BYTES + MAX_REGION_STYLE_MEMO_ENTRIES * REGION_STYLE_MEMO_ENTRY_ESTIMATE_BYTES + formatted * PIVOT_FORMAT_MEMO_BYTES
}

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
/** Pool estimate for each cell owned by a committed worker pivot overlay. */
export const PAGED_CELL_BYTES_ESTIMATE = 128
/** Committed pivot values. `undefined` means the coordinate is not owned;
 * `null` explicitly clears a cell inside an owned rectangle. */
export interface ValueOverlayCell { value: XlsxCell['value']; valueIsError?: boolean; numFmtId?: number }
export interface ValueOverlay {
  /** Cells the overlay replaces, including explicit null blanks (vacated cells). */
  readonly rects: readonly SourceRect[]
  /** The pivots' protected bounds (at least the current output): guarded from sheet sort/filter and
   * joined to the extent. Every non-null `cell()` lies inside them (the extent bounds Find and copy);
   * `setValueOverlay` refuses an overlay that breaks this with a TypeError naming the cell. Vacated cells
   * are in `rects` and read empty; the stores keep every saved pivot area (with its page-field cells)
   * guarded on their own, so vacated cells inside one stay protected whatever is named here (S3-G, S3-L).
   * Required: a fallback to `rects` would grow the extent over vacated cells. The stores derive the used
   * extent from these bounds; no separate row/column count is part of the overlay. */
  readonly protectedRects: readonly SourceRect[]
  cell(row: number, col: number): ValueOverlayCell | null | undefined
  readonly merges: readonly SourceRect[]
}
/** The used extent (exclusive row/column counts) an overlay adds: its protected bounds, as the stores' `extent`. */
export function overlayUsedCounts(rects: readonly SourceRect[]): { usedRows: number; usedCols: number } {
  let usedRows = 0, usedCols = 0
  for (const rect of rects) { usedRows = Math.max(usedRows, rect.r1 + 1); usedCols = Math.max(usedCols, rect.c1 + 1) }
  return { usedRows, usedCols }
}
/** Bounding worksheet rectangle occupied by saved/current pivot areas. `null` means no overlay extent. */
export function overlayUsedExtent(rects: readonly SourceRect[]): SourceRect | null {
  if (!rects.length) return null
  return rects.reduce((bounds, rect) => ({
    r0: Math.min(bounds.r0, rect.r0), r1: Math.max(bounds.r1, rect.r1),
    c0: Math.min(bounds.c0, rect.c0), c1: Math.max(bounds.c1, rect.c1),
  }), { ...rects[0]! })
}
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
  /** Whole-sheet completeness; undefined cannot certify statistics bounds. */
  readonly completeness?: Completeness
  readonly stamp: SourceStamp
  /** Structural generation for geometry-affecting metadata, including overlay merges. */
  readonly structureGen?: number
  readonly valuePolicy: ValuePolicy
  readonly extent: { rows: 1048576; cols: 16384; used?: SourceRect }
  /** Publish a committed pivot overlay; stores advance valueGen atomically. */
  setValueOverlay?(overlay: ValueOverlay | undefined): void
  /** Resident: the sheet's cell values changed in place (a dependent recompute after a pivot change wrote formula
   * results, or put the saved ones back); advances valueGen so value-keyed consumers (tiles, find, filter value
   * lists) re-read. `rows`: its row set changed too (a dependent spill grew into new rows or left rows empty), so the
   * store takes its row inventory and used extent again. */
  noteValuesChanged?(options?: { rows?: boolean }): void
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
  /** S13: clone-safe table/pivot/control descriptors for the main-thread display layer. */
  readPivotDisplay?(options: ReadOptions, grantBytes: number, memory?: DrawingMemory, profile?: ProfileId): Promise<RemotePivotDisplay>
  /** Conditional-formatting preparation for the receiving CF layer: bounded,
   * clone-safe winning cell states plus the shared definitions they were
   * composed from. Absent when the store has no CF path (callers must
   * feature-detect and never touch the store for such sheets). */
  readConditionalFormatting?(request: CfReadRequest, memory?: CfMemory, options?: CfReadOptions): Promise<CfReadResult>
}
/** Read options for a CF preparation: the shared budget lives in the called
 * memory account; callers cancel through the abort signal. `retainOutgoing`,
 * when supplied, receives the release of the live outgoing lease the
 * provider adopts on the shared account: the caller holds it through post
 * and releases after drain instead of the provider refunding at return. */
export interface CfReadOptions {
  readonly signal?: AbortSignal
  readonly budget?: () => CfWorkBudget
  readonly envelopeBytes?: number
  readonly retainOutgoing?: (release: () => void) => void
  readonly dateContext?: import('../cf/predicates').CfDateContext
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
  diagnostics: readonly TableStyleDiagnostic[]
  /** Exact worker estimate when the current receiver grant was insufficient. */
  requiredBytes?: number
}

/** S13 receiver-granted descriptors. Roles, buttons and drawing controls are composed on the main thread. */
export interface RemotePivotDisplay {
  sheetId: string
  stamp: SourceStamp
  pivotTables: readonly PivotTableDefinition[]
  pivotStyles: readonly TableStyleDef[]
  /** The styles.xml dxfs the pivots' `formats` name, by dxfId (bounded; counted in the reply estimate). */
  pivotDxfs?: readonly { id: number; dxf: DxfStyle }[]
  theme?: ThemeContext
  diagnostics: readonly (TableStyleDiagnostic | PivotDiagnostic)[]
  requiredBytes?: number
}

/** Receiver-granted CF preparation (Task 7): clone-safe winning cell states
 * plus the shared definitions they were composed from. Plain data only —
 * no Sets, Maps, or class instances — so structuredClone transfers it
 * without a transfer list. */
export interface CfReadStamp {
  readonly owner: string
  readonly sheetId: string
  readonly workbookValueGen: number
  readonly cfGen: number
  readonly viewGen: number
  readonly structureGen: number
  readonly clockGen: number
}
export interface CfReadRequest {
  readonly rects: readonly SourceRect[]
  readonly stamp: CfReadStamp
  readonly grantBytes: number
}
export interface CfCellStateWire {
  readonly row: number
  readonly col: number
  readonly properties: Record<string, string | number | boolean | null>
  readonly unresolvedWrites?: readonly string[]
  readonly unresolvedReason?: string
  readonly scale?: Readonly<{ ruleId: number; argb: number; fraction?: number }>
  /** Ruling S3-S: the per-cell resolved bar descriptor, numbers and booleans only; `rtl` encodes the direction. */
  readonly bar?: Readonly<{ ruleId: number; axisFraction: number; startFraction: number; endFraction: number; fillColor: number
    borderColor?: number; axisColor?: number; isGradient: boolean; isNegative: boolean; showValue: boolean; rtl: boolean }>
}
export interface CfSharedDefinition {
  readonly id: string
  readonly version: number
  readonly properties: Record<string, string | number | boolean | null>
  readonly writes: readonly string[]
}
export interface CfReadDiagnostic {
  readonly kind: 'unsupported' | 'resource-limit' | 'cancelled' | 'pending' | 'stale'
  readonly feature: string
  readonly message: string
}
export interface CfReadResult {
  /** The stamp this preparation is valid for (authoritative; mirrors a live source read). */
  readonly stamp: CfReadStamp
  readonly sheetId: string
  readonly kind: 'ready' | 'pending' | 'unresolved'
  readonly reason?: string
  readonly states: readonly CfCellStateWire[]
  readonly definitions: readonly CfSharedDefinition[]
  /** Hash of the normalized rule set the states were composed from; consumers
   * pin entries by version and keep the old version charged until release. */
  readonly definitionsVersion: number
  readonly diagnostics: readonly CfReadDiagnostic[]
  /** Sheets whose import had not reached the requested rects yet. */
  readonly pendingSheetIds?: readonly string[]
  /** Exact estimate when the receiver grant was too small; retry with at least this. */
  readonly requiredBytes?: number
}
/** Reply envelope accounted on top of every CF read reply: the kernel-added
 * command/owner/request/stamp framing around the payload (a 1,042 B empty
 * wrapper with short identities, 594 B of which is wrapper overhead),
 * measured with the same walker the kernel publishes with, plus margin.
 * The preflight and post-build checks share this bound, so a fitting grant
 * always admits an exact build; outliers fall back to the kernel walk. */
export const CF_REPLY_ENVELOPE_BYTES = 1152
function cfWireStringBytes(value: string): number {
  return 16 + value.length * 2
}
function cfWireValueBytes(value: string | number | boolean | null | undefined): number {
  return typeof value === 'string' ? cfWireStringBytes(value) : 8
}
/** Measured bytes of one properties object, mirroring the repository payload
 * walker exactly (24 per object, 16 per property slot, keys counted once):
 * shared helpers keep preflight and publication on one model. Exported so
 * providers bound per-cell wire cost with the same arithmetic the reply is
 * measured with. */
export function measureCfPropertiesBytes(properties: Record<string, string | number | boolean | null>): number {
  let bytes = 24
  for (const key of Object.keys(properties)) {
    void key
    bytes += 16 + cfWireValueBytes(properties[key])
  }
  return bytes
}
/** Mirrors the Task 5/6 service's retained per-rule property accounting so the
 * read-side estimate can never admit less than the producer checks. */
export function cfRulePropertyBytes(rules: readonly { dxf?: Record<string, unknown> | undefined; writes: Iterable<string> }[]): number {
  let bytes = 0
  for (const rule of rules) {
    const dxf: Record<string, unknown> = rule.dxf ?? {}
    for (const key of Object.keys(dxf)) {
      const value = dxf[key]
      bytes += 192 + key.length * 4 + (typeof value === 'string' ? value.length * 2 : 16)
    }
    for (const w of rule.writes) bytes += 128 + String(w).length * 4
  }
  return bytes
}
/** Measured bytes of one wire cell state, mirroring the repository payload
 * walker exactly (24 per object; 16 per property slot plus the walked value;
 * arrays 16 + 8 per item plus walked elements; strings 16 + 2 per char;
 * numbers, booleans and null 8; no cross-cell dedup). */
export function measureCfCellBytes(cell: Pick<CfCellStateWire, 'properties'> & Partial<Pick<CfCellStateWire, 'unresolvedWrites' | 'unresolvedReason' | 'scale' | 'bar'>>): number {
  // {row, col, properties} plus optional unresolvedWrites/unresolvedReason.
  let bytes = 24 + (16 + 8) + (16 + 8) + (16 + measureCfPropertiesBytes(cell.properties))
  if (cell.unresolvedWrites !== undefined) {
    bytes += 16 + 16 + 8 * cell.unresolvedWrites.length
    for (const w of cell.unresolvedWrites) bytes += cfWireStringBytes(w)
  }
  if (cell.unresolvedReason !== undefined) bytes += 16 + cfWireStringBytes(cell.unresolvedReason)
  if (cell.scale !== undefined) {
    // 16 prop slot + 24 scale object + 24 ruleId + 24 argb = 88 bytes
    bytes += 16 + 24 + (16 + 8) + (16 + 8)
    if (cell.scale.fraction !== undefined) bytes += 16 + 8
  }
  if (cell.bar !== undefined) {
    // 16 prop slot + 24 bar object + nine required number/boolean fields at (16 + 8) each, plus each optional colour.
    bytes += 16 + 24 + 9 * (16 + 8)
    if (cell.bar.borderColor !== undefined) bytes += 16 + 8
    if (cell.bar.axisColor !== undefined) bytes += 16 + 8
  }
  return bytes
}
/** Measured bytes of one shared definition (counted once however many cells reference it). */
export function measureCfDefinitionBytes(definition: Pick<CfSharedDefinition, 'id' | 'version' | 'properties' | 'writes'>): number {
  let bytes = 24 + (16 + cfWireStringBytes(definition.id)) + (16 + 8) + (16 + measureCfPropertiesBytes(definition.properties))
  bytes += 16 + 16 + 8 * definition.writes.length
  for (const w of definition.writes) bytes += cfWireStringBytes(w)
  return bytes
}
export function measureCfDiagnosticBytes(diagnostic: Pick<CfReadDiagnostic, 'kind' | 'feature' | 'message'>): number {
  return 24 + (16 + cfWireStringBytes(diagnostic.kind)) + (16 + cfWireStringBytes(diagnostic.feature)) + (16 + cfWireStringBytes(diagnostic.message))
}
/** Upper bound the provider checks before touching the source or building:
 * envelope plus diagnostics plus each shared definition once plus a per-cell
 * wire bound (framing included) the provider derives from the same wire
 * measures, so a fitting grant always admits an exact build. Undefined when
 * the bound itself is unrepresentable (fails closed instead of overflowing). */
export function estimateCfPreparationBytes(cellCount: number, cellBytes: number, definitionBytes: number, diagnosticBytes: number, envelopeBytes: number): number | undefined {
  for (const n of [cellCount, cellBytes, definitionBytes, diagnosticBytes, envelopeBytes]) {
    if (!Number.isSafeInteger(n) || n < 0) return undefined
  }
  const cells = cellCount * cellBytes
  if (!Number.isSafeInteger(cells)) return undefined
  const total = envelopeBytes + diagnosticBytes + definitionBytes + cells
  return Number.isSafeInteger(total) ? total : undefined
}
/** FNV-1a hash of the normalized rule set: identical definitions share one
 * cache entry and one charged copy; any change yields a new version. */
export function versionCfRules(normalized: string): number {
  let hash = 0x811c9dc5
  for (let i = 0; i < normalized.length; i++) {
    hash ^= normalized.charCodeAt(i)
    hash = Math.imul(hash, 0x01000193)
  }
  return hash >>> 0
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
  /** Pivot output rectangles are protected from sheet-level sort/filter. */
  pivots?: readonly SourceRect[]
  missing: readonly SourceRect[]
  error?: SourceFailure
}
/** Refusals never publish a changed view owner, spec §§4.4, 6.1, 6.3. */
export type OperationRefusal =
  | { reason: 'pivot'; range: SourceRect; message: string }
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
  /** S13 receiver-granted pivot/control descriptors. */
  | { op: 'pivotDisplay'; grantBytes: number }
  /** Task 7: receiver-granted conditional-formatting preparation — bounded
   * per-cell winning states plus shared rule definitions. */
  | { op: 'conditionalFormatting'; rects: readonly SourceRect[]; stamp: CfReadStamp; grantBytes: number }
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
  | { op: 'pivotDisplay'; result: RemotePivotDisplay }
  | { op: 'conditionalFormatting'; result: CfReadResult }
  | { op: 'detectRange'; detection: RangeDetection }
/** S5.4: resolved cell styles (indexed by cell styleIndex) and the date system
 * of a paged workbook, published with open so worker-backed cells format like
 * resident ones. Bounded by the styles part's small-part cap. */
export interface WorkbookFormatting {
  /** Shared workbook palette, included in the formatting payload's ownership. */
  colors?: import('../sheetColors').SheetColorContext
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
  /** Task 9 CF metadata fast path; absent/false means no per-batch CF reads. */
  hasConditionalFormatting?: boolean
  /** S8: table metadata fast path; absent/false has no per-frame style work. */
  hasRegionStyles?: boolean
  /** Number of table descriptors a bounded regionStyles read may return. */
  regionStyleTableCount?: number
  /** S13 metadata-only fast path; absent/false has no pivot display work. */
  hasPivotDisplay?: boolean
  pivotTableCount?: number
  /** Authored frozen panes. */
  frozen?: { rows: number; cols: number; topRow?: number; leftCol?: number }
}
/** Control selection is a separate workbook-scoped revision; it is not a page-field or item-filter delta. */
export type LinkedControlSelection =
  | { kind: 'slicerSelection'; name: string; selectedItemIds: readonly number[] }
  | { kind: 'slicerClear'; name: string }
  | { kind: 'timelineSelection'; name: string; start: number; end: number }
  | { kind: 'timelineClear'; name: string }
  | { kind: 'tableSlicerSelection'; name: string; selectedItemIds: readonly number[] }
export interface LinkedControlPivotTarget {
  sheetId: string
  pivot: number
  name: string
  cacheId: number
  baseRevision: number
  /** Monotonic per-pivot request sequence; distinct from the workbook transaction revision. */
  requestRevision: number
  delta: PivotDelta | 'reset'
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
  | { kind: 'pivotOpen'; requestId: string; docId: string; sheetId: string; pivot: number; grantBytes: number; /** Main's grant for the fit cells (Task 29 r2; absent: PIVOT_FIT_REPLY_BYTES, 0: none). */ fitGrantBytes?: number }
  | { kind: 'pivotApply'; requestId: string; docId: string; sheetId: string; pivot: number; /** Latest revision the receiver has committed. */ baseRevision: number; /** Unique receiver request sequence, used only to cancel its stage. */ revision: number; delta: PivotDelta | 'reset'; grantBytes: number; /** As pivotOpen's. */ fitGrantBytes?: number }
  | { kind: 'pivotCommit'; requestId: string; docId: string; sheetId: string; pivot: number; revision: number }
  | { kind: 'pivotFormulaStatus'; requestId: string; docId: string; sheetId: string }
  | { kind: 'pivotRelease'; requestId: string; docId: string; sheetId: string; pivot: number }
  | { kind: 'controlApply'; requestId: string; docId: string; revision: number; grantBytes: number; selection: LinkedControlSelection; targets: readonly LinkedControlPivotTarget[] }
  | { kind: 'controlCommit'; requestId: string; docId: string; revision: number }
  | { kind: 'controlCancel'; requestId: string; docId: string; revision: number }
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
  | { kind: 'open'; valuePolicy: ValuePolicy; sheets: readonly { sheetId: string; name: string }[]; hasFormulaCellsOutsidePivots?: boolean; formatting?: WorkbookFormatting }
  | { kind: 'activateSheet'; sheetId: string; layout?: SheetLayout }
  | { kind: 'prepareColumns'; cols: Uint16Array }
  | { kind: 'view'; delta: ViewCommitDelta }
  | ViewportResult
  | { kind: 'valueQuery'; page: Page<ColumnEntry> }
  | { kind: 'search'; items: readonly SearchHit[]; nextCursor?: string; unavailableCount: number; page?: FindPage }
  | { kind: 'read'; result: SourceReadResult }
  | { kind: 'pivotOpen'; /** Last worker-committed wire revision. */ revision: number; def: PivotTableDefinition; /** The menus' model (Task 28), built where the cache lives. */ menu: import('../pivots/session/menu').PivotMenuModel; /** The session's value cells a fit reads (Task 29: the main thread's autofit fits them); absent past the payload cap. */ cells?: readonly { row: number; col: number; value: ValueOverlayCell | null }[] }
  | { kind: 'pivotApply'; /** Worker-assigned candidate revision; commit only this value. */ revision: number; /** A refreshed pivot's one-line notice (Task 32). */ notice?: string; /** The candidate is the pivot's saved view (Ruling S3-AK: a sheet switch may release it). */ atSaved?: true; def: PivotTableDefinition; menu: import('../pivots/session/menu').PivotMenuModel; rects: readonly SourceRect[]; protectedRects: readonly SourceRect[]; overlayExtent: SourceRect | null; merges: readonly SourceRect[]; usedRows: number; usedCols: number; charts: readonly { readonly sheetId: string; readonly partPath: string; readonly treePath: string; readonly chart: ChartModel }[]; /** As pivotOpen's. */ cells?: readonly { row: number; col: number; value: ValueOverlayCell | null }[] }
  | { kind: 'pivotCommit'; revision: number; valueGen: number; hasFormulaCellsOutsidePivots?: boolean } | { kind: 'pivotCommit'; refused: 'stale' }
  | { kind: 'pivotFormulaStatus'; hasFormulaCellsOutsidePivots?: boolean }
  | { kind: 'pivotRelease'; released: boolean; valueGen?: number }
  | { kind: 'controlApply'; revision: number; selection: LinkedControlSelection; targets: readonly { id: string; sheetId: string; pivot: number; revision: number; result: Extract<WorkerResult, { kind: 'pivotApply' }> }[] }
  | { kind: 'controlCommit'; revision: number; selection: LinkedControlSelection; valueGens: readonly { sheetId: string; valueGen: number }[] }
  | { kind: 'controlCancel'; cancelled: boolean }
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
