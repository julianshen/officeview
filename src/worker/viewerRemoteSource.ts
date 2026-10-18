/**
 * Main-thread WorkbookSource over the resident viewer worker (plan S5.4; spec
 * §§4.1, 10.1–10.4). A paged workbook's pages, views and filter evaluation stay
 * in the worker; this adapter implements the `InteractiveSheetDoc` source
 * contract with one bounded command per store call:
 *
 *  - readRows/readCell/scanPopulated/queryStructural/ensureComplete → `read`;
 *  - queryColumn → `valueQuery`; sort/filter commits → `sort`/`filter`.
 *
 * Every command carries the sheet's current committed CacheIdentity, learned
 * from the worker's own view deltas, so the worker withholds replies for an
 * older view, a closed document or a cancelled request; the client ignores
 * replies it no longer waits for. Cancelling a caller's signal sends `cancel`
 * and settles the call at once as cancelled. Nothing transfers a whole sheet:
 * row pages, scan pages and column pages are bounded by the worker.
 *
 * Cell records for painting live in a bounded row cache keyed by 256-column
 * chunks. A chunk is known only after a complete row page covered it; unknown
 * chunks render as pending tiles and are requested in bounded runs. Main-safe:
 * imports the protocol/client and pure view modules, never the session.
 */
import type {
  AuthoritativeCell, CacheIdentity, CellRead, ColumnEntry, ColumnQuery, FilterKeySet, FilterSpec, LogicalCellBox, OperationRefusal, Page,
  PopulatedQuery, RangeCompletion, ReadOptions, RowOrder, SheetLayout, SortSpec, SourceFailure, SourceReadQuery, SourceReadResult,
  SourceStamp, SourceStore, SparseRowPage, StructuralResult, ValuePolicy, ViewCommitDelta, WorkbookFormatting, WorkbookOwner, WorkerDiagnostic, WorkerReply,
  ProfileId, GridSelection, OverviewGroup, OverviewQuery, LogicalViewport, SearchQuery,
} from '../xlsx/view/contracts'
import type { RangeDetection } from '../xlsx/view/rangeDetection'
import { createDocumentOverflowCache } from '../xlsx/view/documentOverflow'
import type { OverflowCache } from '../xlsx/view/overflowCache'
import { MainMemoryLedger, axisBytes, leaseTargets, type LedgerHold } from '../xlsx/view/mainLedger'
import type { XlsxCell, XlsxCellStyle, XlsxMergeRange } from '../xlsx/types'
import { MAX_COLS, MAX_ROWS, formatA1, isValidCoordinate, type SourceCell, type SourceRect } from '../xlsx/view/coordinates'
import { createAxisMetrics } from '../xlsx/view/axes'
import { adoptRowOrder, createIdentityOrder, createRowOrder, isUint32, localUint32, type CompactRowOrder } from '../xlsx/view/order'
import { createIntervalSet, type CompactRowVisibility } from '../xlsx/view/visibility'
import type { ProjectedSheetMetrics } from '../xlsx/view/projection'
import { savedValuesStatus, type ValueStatus } from '../xlsx/view/values'
import { MIN_DETAIL_DEVICE_PX, SHEET_DEFAULT_COL_CSS, SHEET_DEFAULT_ROW_CSS, TEXT_OVERFLOW_REACH_CSS, VIEWPORT_DETAIL_LIMITS } from '../xlsx/view/viewport'
import { MAIN_THREAD_BUDGETS } from '../xlsx/view/tileCache'
import { ViewerNotSentError, type ViewerClient } from './viewerClient'
import { replyLimitEnvelope } from './viewerReplyLimit'
import { initialIdentity, initialWorkbookOwner, reply, visibilityFromDisplayBitsets, type CommandOf } from './viewerProtocol'
import type { SheetGeometry, WorkbookSource } from './viewerSession'
import { RemoteDrawingCache, type PagedSheetDrawings } from './viewerRemoteDrawings'
import { DRAWING_MEMORY_REFUSED, drawingPayloadBudget, unavailableDrawings } from '../xlsx/view/drawingPayload'
import { admitDrawingOwners } from '../xlsx/view/drawingOwners'
import { FindError, workbookFindKey, type FindOptions, type FindPage, type FindSession } from '../xlsx/view/find'
import { queryPayloadBytes, WORKER_QUERY_BYTES } from './viewerQueryBudget'

export const ROW_CHUNK_COLS = 256
/** One document-wide byte ledger for every sheet's painting row cache: a
 * quarter of the main-thread profile budget (§12.4; D 16 MiB, M 8 MiB). */
export function rowCacheBudget(profile: ProfileId): number { return MAIN_THREAD_BUDGETS[profile] / 4 }
const DRAWING_MEMORY_MESSAGE = "This sheet's drawings need more main-thread memory than the document has free and are not shown"
/** FC-T2: no receiving grant fits now; nothing (more) is sent, and the cache may ask again. */
const DRAWING_WORKER_MEMORY_MESSAGE = "This sheet's drawings could not be prepared within the viewer worker's free memory and are not shown yet"
function drawingGrantRefusal(command: CommandOf<'read'>, message = DRAWING_MEMORY_MESSAGE): WorkerReply<'read'> {
  return reply(command, command.owner, { kind: 'complete', result: { kind: 'read', result: { op: 'drawings', result: unavailableDrawings(message, { code: 'resource-limit', retryable: true, message }, DRAWING_MEMORY_REFUSED) } } })
}
/** The bytes a granted reply adopts from its grant: its stamped size, or the whole grant. */
function grantedBytes(reply: WorkerReply, grant: number): number {
  const stamped = (reply as { replyBytes?: unknown }).replyBytes
  return typeof stamped === 'number' && Number.isSafeInteger(stamped) && stamped >= 0 && stamped <= grant ? stamped : grant
}
/** The bytes a non-granted reply is charged without walking it (spec §10.2):
 * the size the worker measured and stamped on it, or, when that is missing or
 * invalid, `maximum` (the whole main-thread budget: no reply main can hold is
 * larger). */
function stampedBytes(reply: WorkerReply, maximum: number): number {
  const stamped = (reply as { replyBytes?: unknown }).replyBytes
  return typeof stamped === 'number' && Number.isSafeInteger(stamped) && stamped >= 0 ? stamped : maximum
}
/** Status text when a sort's view cannot fit the main-thread budget (stream F). */
export const SORT_MEMORY_REFUSAL = 'Not enough memory to apply the sort on this device'
let viewReplySerial = 0
/** Main bytes a sorted-view reply over `rows` rows can need: its band and
 * inverse (4 B each per row, adopted in place), the wire and the local
 * full-axis visibility words with summaries, and object overhead. */
export function viewReplyBytes(rows: number): number {
  const visibility = Math.ceil(MAX_ROWS / 32) * 4
  return 8 * Math.max(0, rows) + 2 * (visibility + Math.ceil(visibility / 32) + 64) + 64 * 1024
}
/** Accounted bytes of one cached cell: record overhead plus UTF-16 text. */
export function cachedCellBytes(cell: { value: unknown; formula?: string }): number {
  return 96 + (typeof cell.value === 'string' ? cell.value.length * 2 : 0) + (cell.formula ? cell.formula.length * 2 : 0)
}
/** Detail rows one tile may request; beyond it the frame is coarse and reads no cells. */
export const MAX_PENDING_ROWS = 2048
/** Source rows two missing runs may be apart and still share one bounded read. */
const COALESCE_GAP_ROWS = 32
/** Row pages one load run may pull before yielding to the next paint. */
const PAGES_PER_RUN = 4
/** Retry delay while the worker's import frontier has not reached a row. */
const LOADING_RETRY_MS = 200

export type PagedCell = AuthoritativeCell & XlsxCell
export interface RemoteSheetView {
  revision: number
  order: RowOrder
  /** DISPLAY-space mask; absent while the view is the authored base. */
  visibility?: CompactRowVisibility
  filter: FilterSpec | null
  sort: SortSpec | null
}
export type RemoteCommit =
  | { kind: 'committed'; view: RemoteSheetView }
  | { kind: 'refused'; refusal: OperationRefusal }
  | { kind: 'stale' | 'cancelled' | 'incomplete' }
  | { kind: 'error'; message: string }
export interface RemoteOpen {
  docId: string
  valuePolicy: ValuePolicy
  sheets: readonly { sheetId: string; name: string }[]
  diagnostics: readonly WorkerDiagnostic[]
  layouts: ReadonlyMap<string, SheetLayout>
  formatting?: WorkbookFormatting
  profile?: ProfileId
  /** @internal Row-cache ledger override (tests); default rowCacheBudget(profile). */
  rowCacheBytes?: number
}
export interface RemoteStats { commands: Record<string, number> }

const failure = (code: SourceFailure['code'], message: string, retryable = false): SourceFailure => ({ code, retryable, message })
const CLOSED: unique symbol = Symbol('closed')
const CANCELLED: unique symbol = Symbol('cancelled')
type Settled<K extends RemoteKind> = WorkerReply<K> | typeof CLOSED | typeof CANCELLED
type RemoteKind = 'read' | 'valueQuery' | 'sort' | 'filter' | 'prepareColumns' | 'viewport' | 'search'

/** Opens `blob` in the worker behind `client`: open, then one activation per
 * sheet (each publishes its compact layout), then the first sheet again. */
export async function openRemoteWorkbook(client: ViewerClient, docId: string, blob: Blob, profile: 'D' | 'M'): Promise<WorkerReply<'open'> | RemoteOpen> {
  const open = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId, source: { kind: 'blob', blob }, profile })
  if (open.kind !== 'complete' || open.result.kind !== 'open') return open
  const { sheets, valuePolicy, formatting } = open.result
  const layouts = new Map<string, SheetLayout>()
  const owner = initialWorkbookOwner(docId, sheets)
  for (const sheet of [...sheets.slice(1), ...sheets.slice(0, 1)]) {
    const active = await client.send({ kind: 'activateSheet', requestId: client.nextRequestId(), owner, sheetId: sheet.sheetId })
    if (active.kind !== 'complete' || active.result.kind !== 'activateSheet' || !active.result.layout) {
      layouts.set(sheet.sheetId, { rowOverrides: [], colOverrides: [], hidden: [] })
      continue
    }
    layouts.set(sheet.sheetId, active.result.layout)
  }
  return { docId, valuePolicy, sheets, diagnostics: open.diagnostics, layouts, profile, ...(formatting ? { formatting } : {}) }
}

export function isRemoteWorkbook(source: WorkbookSource | undefined): source is RemoteWorkbook {
  return source instanceof RemoteWorkbook
}

export class RemoteWorkbook implements WorkbookSource {
  queryNotice?: string
  private noticeScope?: string
  private readonly latestQueries = new Map<string, string>()
  readonly valuePolicy: ValuePolicy
  readonly sheets: readonly { sheetId: string; name: string }[]
  readonly openDiagnostics: readonly WorkerDiagnostic[]
  readonly docId: string
  readonly profile: ProfileId
  private readonly format: WorkbookFormatting
  readonly stats: RemoteStats = { commands: {} }
  /** Document-wide row-cache ledger; insertion order is recency order. */
  readonly ledger: RowCacheLedger
  /** S10.2 §12.4: the document's main-thread application ledger (row cache,
   * adopted view buffers, in-flight replies; the viewer charges its own pools). */
  readonly mainLedger = new MainMemoryLedger()
  private overflow: { profile: ProfileId; cache: OverflowCache } | undefined
  private closedOverflow: OverflowCache | undefined
  private readonly identities = new Map<string, CacheIdentity>()
  private readonly views = new Map<string, RemoteSheetView>()
  /** The workbook's own hold on each sheet's current adopted view (§12.4):
   * released when a newer view replaces it or the workbook closes. */
  private readonly viewHolds = new Map<string, LedgerHold>()
  private readonly geometries = new Map<string, SheetGeometry>()
  // Layout keys disappear on replacement/release; obsolete column axes are
  // retained only by actual consumers, whose reachability charge remains live.
  private columnMetrics = new WeakMap<SheetLayout, ReturnType<typeof createAxisMetrics>>()
  private readonly stores = new Map<string, RemoteSheetStore>()
  private readonly listeners = new Set<() => void>()
  private readonly structures = new Map<string, { gen: number; merges: readonly XlsxMergeRange[]; state: 'loading' | 'done'; timer?: ReturnType<typeof setTimeout> }>()
  private readonly layouts: Map<string, SheetLayout>
  private statusInFlight = false
  private owner: WorkbookOwner
  private statusMessage: string | undefined
  private statusRefresh = 0
  private closed = false
  /** Per sheet: the worker's import failure reported on structural reads (quota, storage, damaged part). */
  private readonly failures = new Map<string, { code: string; message: string }>()
  /** Sheets whose structure arrived complete (the only path to "complete"). */
  private readonly completed = new Set<string>()
  /** S13.2: each sheet's drawing content, read once its structure is complete. */
  private readonly drawingCache: RemoteDrawingCache

  constructor(private readonly client: ViewerClient, open: RemoteOpen, private readonly onRelease: () => Promise<void> | void = () => {}) {
    this.docId = open.docId
    this.valuePolicy = open.valuePolicy
    this.sheets = open.sheets
    this.openDiagnostics = open.diagnostics.filter(d => d.code !== 'cached-value-policy')
    this.statusMessage = open.diagnostics.find(d => d.code === 'cached-value-policy')?.message
    this.owner = initialWorkbookOwner(open.docId, open.sheets)
    this.layouts = new Map(open.layouts)
    for (const { sheetId } of open.sheets) if (!this.layouts.has(sheetId)) this.layouts.set(sheetId, { rowOverrides: [], colOverrides: [], hidden: [] })
    this.profile = open.profile ?? 'D'
    this.format = open.formatting ?? { dateSystem: '1900', styles: [] }
    this.ledger = new RowCacheLedger(open.rowCacheBytes ?? rowCacheBudget(this.profile), this.mainLedger)
    for (const { sheetId } of open.sheets) this.identities.set(sheetId, initialIdentity(open.docId, sheetId))
    this.drawingCache = new RemoteDrawingCache({
      identity: sheetId => this.identity(sheetId), loadState: sheetId => this.loadState(sheetId), nextRequestId: () => this.nextRequestId(),
      request: (command, signal) => this.request(command, signal),
      // Adopted/evicted content re-keys the sheet's geometry and tiles.
      changed: sheetId => { (this.merges(sheetId) as { gen: number }).gen++; this.notify() },
    }, 2 * drawingPayloadBudget(this.profile))
  }
  /** S13.2: the sheet's drawing content (viewerRemoteDrawings.ts). */
  drawings(sheetId: string): PagedSheetDrawings { return this.drawingCache.get(sheetId) }

  get isClosed(): boolean { return this.closed }
  /** @internal S10.2 tests: document-owned buffers this workbook still holds. */
  retainedResources(): { stores: number; views: number; geometries: number; layouts: number; structures: number } {
    return { stores: this.stores.size, views: this.views.size, geometries: this.geometries.size, layouts: this.layouts.size, structures: this.structures.size }
  }
  /** Resolved style of a cell style index (the paged reader keeps ids only). */
  styleOf(index: number | undefined): XlsxCellStyle | undefined { return this.format.styles[index ?? 0] }
  formatting(): WorkbookFormatting { return this.format }
  /** Row pages one tile / one frame may have in flight (§11.1 page caps). */
  get pageLimits(): { tile: number; frame: number } {
    const limits = VIEWPORT_DETAIL_LIMITS[this.profile]
    return { tile: limits.tile.pages, frame: limits.pages }
  }
  /** Page requests painting work holds (all sheets) right now: row-cache runs
   * and worker tile plans. Credits return when a request settles. */
  pagesInFlight = 0
  private pageFrame: PageFrame = { used: 0, tiles: new Map() }
  /** Starts a frame's cumulative page allowance (§11.1: per frame and per
   * tile, D 32/8, M 16/4). Called once per presented frame. */
  beginPageFrame(): void { this.pageFrame = { used: 0, tiles: new Map() } }
  /** The frame whose allowance new painting requests spend. */
  get currentPageFrame(): PageFrame { return this.pageFrame }
  /**
   * Admits `pages` page requests for `tile` against ONE budget shared by row
   * reads and worker plans: the frame's and the tile's cumulative allowance
   * (continuations included; nothing is recycled within a frame) and the
   * outstanding requests across frames. Ring (prefetch) work stays within
   * half the frame on both counts, so visible work always finds room.
   */
  admitPages(frame: PageFrame, tile: string, pages: number, ring: boolean, commit = true): boolean {
    const limits = this.pageLimits
    const ceiling = ring ? limits.frame / 2 : limits.frame
    const tileUsed = frame.tiles.get(tile) ?? 0
    if (this.pagesInFlight + pages > ceiling || frame.used + pages > ceiling || tileUsed + pages > limits.tile) return false
    if (!commit) return true
    this.pagesInFlight += pages
    frame.used += pages
    frame.tiles.set(tile, tileUsed + pages)
    return true
  }
  /** Outstanding credits only: explicit reads (inspection, copy) outside any frame. */
  reservePages(pages: number): boolean {
    if (this.pagesInFlight + pages > this.pageLimits.frame) return false
    this.pagesInFlight += pages
    return true
  }
  releasePages(pages: number): void { this.pagesInFlight -= pages }
  /** One occupancy cache across all sheets and geometry generations. A closed
   * workbook hands out an already-closed owner (no charge): its reads refuse with
   * OverflowResourceError('closed'), which planning turns into a coarse frame and
   * a notice, so render, memo and layout-effect callers never throw. */
  overflowFor(profile: ProfileId = this.overflow?.profile ?? this.profile): OverflowCache {
    if (this.closed) {
      if (!this.closedOverflow) { this.closedOverflow = createDocumentOverflowCache(this.mainLedger, profile); this.closedOverflow.close() }
      return this.closedOverflow
    }
    if (this.overflow?.profile === profile && !this.overflow.cache.closed) return this.overflow.cache
    this.overflow?.cache.close()
    const cache = createDocumentOverflowCache(this.mainLedger, profile)
    this.overflow = { profile, cache }
    return cache
  }

  /** The active sheet keeps its cache; inactive sheets' caches are released. */
  activate(sheetId: string): void {
    for (const [id, store] of this.stores) if (id !== sheetId) store.clear()
  }
  identity(sheetId: string): CacheIdentity { return this.identities.get(sheetId) ?? initialIdentity(this.docId, sheetId) }
  workbookOwner(): WorkbookOwner { return this.owner }
  /** Known sheets have stable layout identities. Unknown IDs intentionally
   * use an uncached empty fallback rather than growing the layout map. */
  layout(sheetId: string): SheetLayout { return this.layouts.get(sheetId) ?? { rowOverrides: [], colOverrides: [], hidden: [] } }
  /** The worker's committed view of a sheet (revision = its viewGen). */
  view(sheetId: string): RemoteSheetView {
    return this.views.get(sheetId) ?? { revision: this.identity(sheetId).viewGen, order: createIdentityOrder(), filter: null, sort: null }
  }
  store(sheetId: string): RemoteSheetStore {
    let store = this.stores.get(sheetId)
    if (!store) {
      if (!this.sheets.some(sheet => sheet.sheetId === sheetId)) throw new RangeError(`Unknown sheet ${sheetId}`)
      this.stores.set(sheetId, store = new RemoteSheetStore(this, sheetId))
    }
    return store
  }
  /** Column-only geometry for proof validation. The caller admits construction
   * first; sharing this axis with geometry() avoids allocating row summaries. */
  columns(sheetId: string): ReturnType<typeof createAxisMetrics> {
    const layout = this.layout(sheetId)
    let columns = this.columnMetrics.get(layout)
    if (!columns) {
      columns = createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: SHEET_DEFAULT_COL_CSS, overrides: layout.colOverrides, ...(layout.hiddenCols?.length ? { collapsed: layout.hiddenCols } : {}) })
      this.mainLedger.chargeWhileReachable('geometry', [columns], axisBytes(columns))
      this.columnMetrics.set(layout, columns)
    }
    return columns
  }
  geometry(sheetId: string): SheetGeometry {
    let geometry = this.geometries.get(sheetId)
    if (!geometry) {
      const layout = this.layout(sheetId)
      geometry = Object.freeze({
        rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: SHEET_DEFAULT_ROW_CSS, overrides: layout.rowOverrides }),
        cols: this.columns(sheetId),
        rowOverrides: layout.rowOverrides, colOverrides: layout.colOverrides, baseHidden: createIntervalSet(layout.hidden),
        ...(layout.hiddenCols?.length ? { hiddenCols: layout.hiddenCols } : {}),
      })
      this.geometries.set(sheetId, geometry)
      // Columns have their own shared lifetime charge; this geometry owns
      // only its row summaries and authored-hidden intervals.
      this.mainLedger.chargeWhileReachable('geometry', [geometry], axisBytes(geometry.rows) + layout.hidden.length * 32 + (layout.hiddenCols?.length ?? 0) * 32)
    }
    return geometry
  }
  valueStatus(): ValueStatus {
    if (this.valuePolicy !== 'cached') return { policy: this.valuePolicy, formulas: 0, message: undefined }
    // Stated limit: the count is parsed from the worker's saved-values status
    // text (savedValuesStatus wording); WorkerDiagnostic carries no number field.
    const counted = this.statusMessage?.match(/— (\d+) formulas? not recalculated/)
    // The progressive import learns its formula count as it runs; completions
    // carry the latest status, so an idle viewer asks with one tiny read.
    // An ended import (failure or cancellation) never learns more: no probe.
    if (!counted && !this.closed && !this.failures.size && !this.statusInFlight && Date.now() - this.statusRefresh > 500) {
      this.statusRefresh = Date.now()
      this.statusInFlight = true
      const store = this.store(this.sheets[0].sheetId)
      void store.ensureComplete({ r0: 0, r1: 0, c0: 0, c1: 0 }, { signal: new AbortController().signal, stamp: store.stamp }).finally(() => { this.statusInFlight = false })
    }
    return { policy: 'cached', formulas: counted ? Number(counted[1]) : undefined, message: this.statusMessage ?? savedValuesStatus('cached'), ...(this.failures.size ? { ended: true } : {}) }
  }
  /** Merged ranges of a sheet from one whole-sheet structural read (bounded by
   * the sheet's merge records). `gen` grows when they arrive; until then, and
   * while the import has not reached the sheet's end, the list is empty. */
  merges(sheetId: string): { gen: number; merges: readonly XlsxMergeRange[] } {
    let entry = this.structures.get(sheetId)
    if (!entry) {
      this.structures.set(sheetId, entry = { gen: 0, merges: [], state: 'loading' })
      void this.loadStructure(sheetId)
    }
    return entry
  }
  private async loadStructure(sheetId: string): Promise<void> {
    const entry = this.structures.get(sheetId)
    if (!entry || this.closed) return
    const store = this.store(sheetId)
    const result = await store.queryStructural({ r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }, { signal: new AbortController().signal, stamp: store.stamp })
    if (this.closed) return
    if (result.completeness === 'complete') {
      // Structure completes with the sheet: its final row metadata and used
      // range replace the layout captured at open (a progressive import grows it).
      const layout = await this.fetchLayout(sheetId)
      if (this.closed) return
      const relaid = layout !== undefined && this.setLayout(sheetId, layout)
      if (result.merges.length) entry.merges = result.merges.map(merge => ({ minRow: merge.range.r0, maxRow: merge.range.r1, minCol: merge.range.c0, maxCol: merge.range.c1 }))
      if (relaid || result.merges.length) entry.gen++
      entry.state = 'done'
      this.completed.add(sheetId)
      this.failures.delete(sheetId)
      this.drawingCache.structureComplete(sheetId)
      // The load state changed (loading -> complete) even when geometry did not.
      this.notify()
      return
    }
    // Structure follows sheetData: wait for the import, unless it failed for good.
    // A failure the worker reports (even a retryable quota/storage one) is shown
    // at once: the sheet is incomplete, never an empty complete body (LIFE-03).
    const error = result.error, code = error?.code
    // A request-level cancel is retryable; a cancelled import is not and is
    // recorded like any other terminal failure (never "complete").
    const recorded = error && code !== 'incomplete' && code !== 'stale-source' && !(code === 'cancelled' && error.retryable)
    if (recorded && this.failures.get(sheetId)?.message !== error.message) {
      this.failures.set(sheetId, { code: error.code, message: error.message })
      this.notify()
    }
    // A page write that failed ended the import and abandoned its pages: no
    // reread can recover (only reopening can), so polling stops here too.
    if (error && code !== 'incomplete' && (!error.retryable || code === 'quota' || code === 'storage-error')) { entry.state = 'done'; return }
    entry.timer = setTimeout(() => { entry.timer = undefined; void this.loadStructure(sheetId) }, 500)
  }
  /** Load state of a sheet's source: complete once its structure has arrived;
   * incomplete with the worker's message once the import reported a failure. */
  loadState(sheetId: string): { completeness: 'complete' | 'loading' | 'incomplete'; failure?: string; code?: string } {
    const failure = this.failures.get(sheetId)
    if (failure !== undefined) return { completeness: 'incomplete', failure: failure.message, code: failure.code }
    this.merges(sheetId)
    if (this.completed.has(sheetId)) return { completeness: 'complete' }
    // Done without completing (a terminal error left unrecorded): never complete.
    return { completeness: this.structures.get(sheetId)?.state === 'done' ? 'incomplete' : 'loading' }
  }
  /** Replaces a sheet's layout; true when it changed. Callers bump the
   * structure generation so the viewer rebuilds geometry and drops tiles. */
  private setLayout(sheetId: string, layout: SheetLayout): boolean {
    if (JSON.stringify(layout) === JSON.stringify(this.layout(sheetId))) return false
    this.layouts.set(sheetId, layout)
    this.geometries.delete(sheetId)
    this.stores.get(sheetId)?.relayout()
    return true
  }
  /** Publishes a newer layout (the import's final metadata) and its generation. */
  publishLayout(sheetId: string, layout: SheetLayout): void {
    if (this.closed || !this.setLayout(sheetId, layout)) return
    const entry = this.merges(sheetId) as { gen: number }
    entry.gen++
    this.drawingCache.invalidate(sheetId)
    this.notify()
  }
  private async fetchLayout(sheetId: string): Promise<SheetLayout | undefined> {
    const out = await this.request({ kind: 'read', requestId: this.nextRequestId(), owner: this.identity(sheetId), query: { op: 'layout' } })
    return typeof out === 'object' && out.kind === 'complete' && out.result.kind === 'read' && out.result.result.op === 'layout' ? out.result.result.layout : undefined
  }
  subscribe(listener: () => void): () => void {
    this.listeners.add(listener)
    return () => { this.listeners.delete(listener) }
  }
  notify(): void { for (const listener of [...this.listeners]) listener() }

  /** One command with caller cancellation. Resolves CLOSED after release and
   * CANCELLED as soon as `signal` aborts (the worker gets a `cancel`). */
  /** `reservation`: a 'replies' owner admitted for this reply before the command
   * was sent; it ends exactly when the reply's own charge starts (never both). */
  async request<K extends RemoteKind>(command: CommandOf<K>, signal?: AbortSignal, progress?: (completed: number, total: number) => void, reservation?: string): Promise<Settled<K>> {
    if (this.closed) return CLOSED
    if (signal?.aborted) return CANCELLED
    const head = command as unknown as { kind: K; requestId: string; owner: CacheIdentity | WorkbookOwner; query?: { op: string } }
    const key = head.kind === 'read' ? `read:${head.query?.op}` : head.kind
    const scope = JSON.stringify([head.owner.docId, 'sheetId' in head.owner ? head.owner.sheetId : null, key])
    this.latestQueries.set(scope, head.requestId)
    if ((head.kind === 'search' || head.kind === 'sort' || head.kind === 'filter') && this.queryNotice) { this.queryNotice = undefined; this.notify() }
    this.stats.commands[key] = (this.stats.commands[key] ?? 0) + 1
    // FC-T2: a drawings reply is admitted BEFORE it exists. The grant is
    // reserved synchronously (concurrent reads see each other's credit) and the
    // worker never posts a reply over it; see drawingGrant.
    const drawings = key === 'read:drawings' ? command as CommandOf<'read'> : undefined
    let grant: string | undefined
    if (drawings) {
      // A duplicate id must not touch the credit of the read that owns it.
      if (this.client.isPending(head.requestId) || this.mainLedger.ownerBytes('replies', `drawing-grant:${head.requestId}`) > 0) return CLOSED
      const admitted = this.drawingGrant(drawings)
      if (typeof admitted !== 'string') return admitted as Settled<K>
      grant = admitted
      command = { ...drawings, query: { op: 'drawings', replyByteLimit: this.mainLedger.ownerBytes('replies', grant) } } as CommandOf<K>
    }
    const unsubscribe = progress ? this.client.onProgress(head.requestId, p => progress(p.completed, p.total ?? p.completed)) : undefined
    const sent: Promise<Settled<K>> = this.client.send<K>(command).then(reply => {
      // Charged until no part of the payload is reachable: a consumer that keeps
      // a page (a scan generator across tasks) keeps the charge. An admitted
      // reservation becomes that charge (no gap, no double count); any excess
      // over it is charged on top.
      // No reply is walked here: the worker stamped its exact size on it (a
      // granted one within its grant). A missing or invalid size is charged as
      // the whole grant, or as the whole main budget for any other reply.
      const pool = drawings ? 'drawings' : 'replies', targets = leaseTargets(reply)
      const bytes = grant ? grantedBytes(reply, this.mainLedger.ownerBytes('replies', grant)) : stampedBytes(reply, MAIN_THREAD_BUDGETS[this.profile])
      const held = grant ?? reservation
      const moved = held ? Math.min(bytes, this.mainLedger.ownerBytes('replies', held)) : 0
      if (moved > 0) this.mainLedger.adoptWhileReachable('replies', held!, pool, targets, moved)
      // The unused grant goes back before owner metadata is admitted on the same cap.
      if (grant) this.mainLedger.release('replies', grant)
      if (moved === 0 || bytes > moved) this.mainLedger.chargeWhileReachable(pool, targets, bytes - moved)
      if (drawings && reply.kind === 'quota') return this.drawingRetry(drawings, reply as WorkerReply<'read'>, signal) as Promise<Settled<K>>
      if (drawings && reply.kind === 'complete' && reply.result.kind === 'read' && reply.result.result.op === 'drawings' && reply.result.result.result.completeness === 'complete'
        && !admitDrawingOwners(reply.result.result.result, this.mainLedger, MAIN_THREAD_BUDGETS[this.profile])) {
        // The original raw reply remains charged until collected. No ownership
        // links were made; the existing cache displays this typed refusal.
        return { ...reply, result: { kind: 'read', result: { op: 'drawings', result: unavailableDrawings(DRAWING_MEMORY_MESSAGE, { code: 'resource-limit', retryable: false, message: DRAWING_MEMORY_MESSAGE }, DRAWING_MEMORY_REFUSED) } } } as WorkerReply<K>
      }
      this.capture(reply, scope, !drawings)
      return reply
    }, error => {
      // A grant is refunded only when the command provably never left: a
      // disposed client may still have its reply in flight (a closed
      // workbook's ledger keeps that harmless overstatement).
      if (grant ? error instanceof ViewerNotSentError : reservation) this.mainLedger.release('replies', (grant ?? reservation)!)
      return CLOSED
    })
    if (!signal) return sent.finally(() => unsubscribe?.())
    return new Promise<Settled<K>>(resolve => {
      const abort = () => {
        if (this.latestQueries.get(scope) === head.requestId) this.latestQueries.delete(scope)
        resolve(CANCELLED)
        if (!this.closed) void this.client.send({ kind: 'cancel', requestId: this.client.nextRequestId(), targetRequestId: head.requestId }).catch(() => {})
      }
      signal.addEventListener('abort', abort, { once: true })
      void sent.then(reply => { signal.removeEventListener('abort', abort); unsubscribe?.(); resolve(this.closed ? CLOSED : reply) })
    })
  }
  /**
   * FC-T2 receiving grant of one drawings read: min(headroom, a quarter of the
   * main budget, the size of the worker's binary drawing budget), or exactly
   * the size a retry asks for. Bounded so a read in
   * flight never starves the other main allocators (tiles, row cache, sorts);
   * a denser reply gets one retry at its measured size (drawingRetry). Below
   * the payload-free envelope nothing is sent and the refusal is retryable:
   * main pressure is transient (another sheet's grant settles in moments).
   */
  private drawingGrant(command: CommandOf<'read'>): string | WorkerReply<'read'> {
    const headroom = MAIN_THREAD_BUDGETS[this.profile] - this.mainLedger.bytes
    const asked = command.query.op === 'drawings' ? command.query.replyByteLimit : undefined
    const limit = asked ?? Math.min(headroom, MAIN_THREAD_BUDGETS[this.profile] / 4)
    if (limit > headroom || limit < replyLimitEnvelope(command)) return drawingGrantRefusal(command)
    const grant = `drawing-grant:${command.requestId}`
    this.mainLedger.set('replies', grant, limit)
    return grant
  }
  /** The worker measured a reply over its grant: ask once more (the worker's
   * source cache makes the re-read cheap) with exactly that size, adjusted for
   * the new requestId's length, when it fits now; otherwise refuse retryably. */
  private drawingRetry(command: CommandOf<'read'>, quota: WorkerReply<'read'>, signal?: AbortSignal): Promise<Settled<'read'>> | WorkerReply<'read'> {
    const required = quota.kind === 'quota' ? quota.requiredBytes : undefined
    // Unmeasurable under the worker's current pressure: the cache asks later.
    if (required === undefined) return drawingGrantRefusal(command, DRAWING_WORKER_MEMORY_MESSAGE)
    if (this.closed || signal?.aborted || command.query.op !== 'drawings' || command.query.replyByteLimit !== undefined) return drawingGrantRefusal(command)
    const requestId = this.nextRequestId()
    const replyByteLimit = required + 2 * Math.max(0, requestId.length - command.requestId.length)
    return this.request({ ...command, requestId, query: { op: 'drawings', replyByteLimit } }, signal)
  }
  nextRequestId(): string { return this.client.nextRequestId() }
  private capture(reply: WorkerReply, scope: string, rangeNotice: boolean): void {
    const status = reply.diagnostics?.find(d => d.code === 'cached-value-policy')
    if (status && status.message !== this.statusMessage) { this.statusMessage = status.message; this.notify() }
    // Drawing failures have their own bounded retry and visible diagnostics.
    if (!rangeNotice) return
    const timeout = reply.diagnostics.find(d => d.code === 'query-timeout')
    if (this.closed || this.latestQueries.get(scope) !== reply.requestId || !reply.owner) return
    const current = 'sheetId' in reply.owner ? this.identity(reply.owner.sheetId) : this.owner
    const matches = 'sheetId' in reply.owner && 'sheetId' in current
      ? reply.owner.docId === current.docId && reply.owner.sheetId === current.sheetId && reply.owner.viewGen === current.viewGen && reply.owner.valueGen === current.valueGen && reply.owner.fontGen === current.fontGen
      : 'sheetVector' in reply.owner && workbookFindKey(reply.owner) === workbookFindKey(this.owner)
    if (matches && timeout) { this.noticeScope = scope; this.queryNotice = timeout.message; this.notify() }
    else if (matches && reply.kind === 'complete' && this.noticeScope === scope && this.queryNotice) { this.queryNotice = undefined; this.noticeScope = undefined; this.notify() }
  }

  async viewport(owner: CacheIdentity, window: LogicalViewport, signal: AbortSignal, scope?: 'frame' | 'tile'): Promise<Settled<'viewport'>> {
    return this.request({ kind: 'viewport', requestId: this.nextRequestId(), owner, window, ...(scope ? { scope } : {}) }, signal)
  }

  /** The worker owns raw search; the UI controller separately guards its local
   * font/layout owner before accepting the returned source identities. */
  findSession(hostOwner: WorkbookOwner, query: SearchQuery, options: FindOptions): FindSession {
    const owner = this.workbookOwner(), controller = new AbortController()
    if (workbookFindKey({ ...hostOwner, fontGen: owner.fontGen }) !== workbookFindKey(owner)) throw new FindError('stale', 'Host and worker search owners differ')
    let cursor: string | undefined, closed = false
    const abort = () => controller.abort()
    options.signal?.addEventListener('abort', abort, { once: true })
    if (options.signal?.aborted) abort()
    const guard = () => {
      if (closed || this.closed) throw new FindError('closed', 'Search closed')
      if (controller.signal.aborted) throw new FindError('cancelled', 'Search cancelled')
      if (workbookFindKey(owner) !== workbookFindKey(this.owner) || options.currentOwner && workbookFindKey(options.currentOwner()) !== workbookFindKey(hostOwner)) throw new FindError('stale', 'Search owner changed')
    }
    return {
      nextPage: async (): Promise<FindPage> => {
        guard()
        const out = await this.request({ kind: 'search', requestId: this.nextRequestId(), owner, query, ...(cursor ? { cursor } : {}) }, controller.signal)
        if (typeof out === 'object' && out.kind === 'complete') cursor = out.result.nextCursor
        guard()
        if (out === CANCELLED) throw new FindError('cancelled', 'Search cancelled')
        if (out === CLOSED) throw new FindError('closed', 'Search closed')
        if (out.kind === 'cancelled') {
          const timeout = out.diagnostics.find(d => d.code === 'query-timeout')
          if (timeout) throw new FindError('resource-limit', timeout.message)
          throw new FindError('cancelled', 'Search cancelled')
        }
        if (out.kind === 'error' && out.code === 'stale-source') throw new FindError('stale', 'Search owner changed')
        if (out.kind === 'complete' && out.result.page) {
          cursor = out.result.nextCursor
          return { ...out.result.page, owner: hostOwner }
        }
        throw new Error(out.diagnostics[0]?.message ?? 'Worker search could not complete. Retry or narrow the search.')
      },
      close: () => {
        if (closed) return
        closed = true
        controller.abort()
        options.signal?.removeEventListener('abort', abort)
        if (cursor && !this.closed) void this.request({ kind: 'read', requestId: this.nextRequestId(), owner: this.identity(this.sheets[0].sheetId), query: { op: 'releaseSearch', cursor } })
        cursor = undefined
      },
    }
  }

  /** Worker sort: prepare the key columns (best effort; unprepared keys stream), then commit. */
  async sort(owner: CacheIdentity, spec: SortSpec | null, signal: AbortSignal, progress?: (completed: number, total: number) => void): Promise<RemoteCommit> {
    // Main-budget admission (stream F): the sorted view arrives as one reply
    // (order bands and their inverse, visibility) that the view then adopts in
    // place. Reserve it before anything is sent; if it cannot fit, refuse now,
    // while main and worker still agree on the current view.
    const reservation = spec?.keys.length ? this.reserveViewReply(spec.range.r1 - spec.range.r0 + 1) : undefined
    if (reservation === null) return { kind: 'refused', refusal: { reason: 'resource-limit', message: SORT_MEMORY_REFUSAL } }
    try { return await this.sortReserved(owner, spec, signal, progress, reservation) }
    finally { if (reservation) this.mainLedger.release('replies', reservation) }
  }
  /** Reserves the main bytes a view reply over `rows` sorted rows can need
   * (bands + inverse + wire and local visibility); null when they do not fit. */
  private reserveViewReply(rows: number): string | null {
    const bytes = viewReplyBytes(rows), budget = MAIN_THREAD_BUDGETS[this.profile]
    if (this.mainLedger.bytes + bytes > budget) return null
    const owner = `view-reply:${++viewReplySerial}`
    this.mainLedger.set('replies', owner, bytes)
    return owner
  }
  private async sortReserved(owner: CacheIdentity, spec: SortSpec | null, signal: AbortSignal, progress: ((completed: number, total: number) => void) | undefined, reservation: string | undefined): Promise<RemoteCommit> {
    if (spec?.keys.length) {
      const cols = new Uint16Array([...new Set(spec.keys.map(key => key.col))])
      const prepared = await this.request({ kind: 'prepareColumns', requestId: this.nextRequestId(), owner, cols }, signal)
      if (prepared === CANCELLED) return { kind: 'cancelled' }
      if (prepared === CLOSED) return { kind: 'stale' }
      if (prepared.kind === 'cancelled') return prepared.diagnostics.some(d => d.code === 'query-timeout')
        ? { kind: 'error', message: prepared.diagnostics[0].message } : { kind: 'cancelled' }
      if (prepared.kind === 'error' && prepared.code === 'stale-source') return { kind: 'stale' }
    }
    return this.commit(await this.request({ kind: 'sort', requestId: this.nextRequestId(), owner, spec }, signal, progress, reservation), owner.sheetId, { sort: spec }, reservation)
  }
  async filter(owner: CacheIdentity, spec: FilterSpec | null, keySets: readonly FilterKeySet[], signal: AbortSignal, progress?: (completed: number, total: number) => void): Promise<RemoteCommit> {
    const command: CommandOf<'filter'> = { kind: 'filter', requestId: this.nextRequestId(), owner, spec, nowMs: Date.now(), ...(keySets.length ? { keySets } : {}) }
    return this.commit(await this.request(command, signal, progress), owner.sheetId, { filter: spec })
  }
  private commit(out: WorkerReply<'sort' | 'filter'> | typeof CLOSED | typeof CANCELLED, sheetId: string, spec: { sort?: SortSpec | null; filter?: FilterSpec | null }, reservation?: string): RemoteCommit {
    if (out === CANCELLED) return { kind: 'cancelled' }
    if (out === CLOSED) return { kind: 'stale' }
    switch (out.kind) {
      case 'complete': {
        if (out.result.kind !== 'view') return { kind: 'error', message: 'Unexpected view reply' }
        const view = this.adopt(sheetId, out.result.delta, spec, out, reservation)
        // Adoption moved its bytes to the view; a stale reply has no consumer.
        this.mainLedger.settle(out)
        return view ? { kind: 'committed', view } : { kind: 'stale' }
      }
      case 'cancelled': return out.diagnostics.some(d => d.code === 'query-timeout') ? { kind: 'error', message: out.diagnostics[0].message } : { kind: 'cancelled' }
      case 'incomplete': return { kind: 'incomplete' }
      case 'quota': return { kind: 'refused', refusal: { reason: 'resource-limit', message: 'View exceeds the worker memory budget.' } }
      case 'error':
        if (out.code === 'operation-refused' && out.refusal) return { kind: 'refused', refusal: out.refusal }
        if (out.code === 'stale-source') return { kind: 'stale' }
        return { kind: 'error', message: out.diagnostics[0]?.message ?? `View update failed (${out.code}).` }
      default: return { kind: 'error', message: 'Unexpected view reply' }
    }
  }
  /** Adopts a committed worker view. The worker is the authority: a delta is
   * adopted whenever it is newer than what this side holds, even when the UI
   * that asked for it has moved on, so both sides never diverge. */
  private adopt(sheetId: string, delta: ViewCommitDelta, spec: { sort?: SortSpec | null; filter?: FilterSpec | null }, reply?: object, reservation?: string): RemoteSheetView | null {
    if (this.closed || delta.owner.docId !== this.docId || delta.owner.sheetId !== sheetId) return null
    const previous = this.view(sheetId)
    if (delta.owner.viewGen <= this.identity(sheetId).viewGen) return null
    // Zero-copy (stream F): the wire bands already carry their inverse; the view
    // takes those buffers and the reply's charge for them moves to the view, so
    // adoption adds only the local visibility words. Invalid or foreign bands
    // fall back to the copying path below.
    let adopted: CompactRowOrder | null = null
    try { adopted = adoptRowOrder(delta.bands) } catch { adopted = null }
    if (adopted) return this.publish(sheetId, delta, spec, previous, adopted, reply, reservation)
    // Staging: the order's local bands and inverse (makeBand's slice is their
    // only copy) coexist with the wire bands (reply) and the previous view.
    // Wire bands may come from another realm: re-view them, never pre-copy.
    // Non-typed input is still converted, and that copy is staged as well.
    let bandBytes = 0, copiedBytes = 0
    for (const band of delta.bands) {
      const bytes = band.sourceAtOffset.length * 4
      bandBytes += bytes
      if (!isUint32(band.sourceAtOffset)) copiedBytes += bytes
    }
    this.mainLedger.set('viewBuffers', `staging:${sheetId}`, bandBytes * 2 + copiedBytes)
    const inputs = delta.bands.map(band => ({ r0: band.r0, r1: band.r1, sourceAtOffset: isUint32(band.sourceAtOffset) ? localUint32(band.sourceAtOffset) : new Uint32Array(band.sourceAtOffset) }))
    const view: RemoteSheetView = {
      revision: delta.owner.viewGen,
      order: inputs.length ? createRowOrder(inputs) : createIdentityOrder(),
      visibility: visibilityFromDisplayBitsets(delta.visibility),
      filter: 'filter' in spec ? spec.filter ?? null : previous.filter,
      sort: 'sort' in spec ? spec.sort ?? null : previous.sort,
    }
    this.identities.set(sheetId, delta.owner)
    if (delta.workbookOwner.globalViewGen >= this.owner.globalViewGen) this.owner = delta.workbookOwner
    this.views.set(sheetId, view)
    // The adopted view (order bands with their display inverse, visibility words
    // and summaries) is charged while held (spec §§10.1, 12.4): by this workbook
    // while it is current, and by every display, copy or operation that keeps it
    // (holdView). A replaced view nothing else holds ends at once; one still
    // held stays charged, so a replacement's overlap counts. The identity order
    // is a shared singleton and never names the charge.
    const visibilityBytes = view.visibility ? view.visibility.words.byteLength + view.visibility.summaries.byteLength : 0
    const targets = [view, ...(view.order.kind === 'identity' ? [] : [view.order]), ...(view.visibility ? [view.visibility] : [])]
    // The staged local bands ARE the new order's bands: end the staging charge as
    // the view takes them over, so the swap never counts them twice (stream F: a
    // 4 MiB double count at the F2 1M-row sort pushed the M ledger past 32 MiB).
    this.mainLedger.release('viewBuffers', `staging:${sheetId}`)
    // As in publish(), what is left of the reservation covered the local
    // visibility words built above; it ends as they are charged.
    if (reservation) this.mainLedger.release('replies', reservation)
    this.mainLedger.chargeWhileHeld('viewBuffers', targets, bandBytes * 2 + visibilityBytes)
    const replaced = this.viewHolds.get(sheetId)
    this.viewHolds.set(sheetId, this.mainLedger.hold(view)!)
    replaced?.release()
    this.notify()
    return view
  }
  /** Publishes a view whose order took the reply's own band buffers. */
  private publish(sheetId: string, delta: ViewCommitDelta, spec: { sort?: SortSpec | null; filter?: FilterSpec | null }, previous: RemoteSheetView, order: CompactRowOrder, reply: object | undefined, reservation: string | undefined): RemoteSheetView {
    const view: RemoteSheetView = {
      revision: delta.owner.viewGen,
      order,
      visibility: visibilityFromDisplayBitsets(delta.visibility),
      filter: 'filter' in spec ? spec.filter ?? null : previous.filter,
      sort: 'sort' in spec ? spec.sort ?? null : previous.sort,
    }
    this.identities.set(sheetId, delta.owner)
    if (delta.workbookOwner.globalViewGen >= this.owner.globalViewGen) this.owner = delta.workbookOwner
    this.views.set(sheetId, view)
    const bandBytes = order.diagnostics.typedArrayBytes
    const visibilityBytes = view.visibility ? view.visibility.words.byteLength + view.visibility.summaries.byteLength : 0
    const targets = [view, ...(order.kind === 'identity' ? [] : [order]), ...(view.visibility ? [view.visibility] : [])]
    // The band bytes were admitted with the reply: transfer them, never re-charge.
    const token = reply ? this.mainLedger.leaseOf(reply) : undefined
    const moved = token ? Math.min(bandBytes, this.mainLedger.ownerBytes('replies', token)) : 0
    if (moved > 0) this.mainLedger.adoptWhileReachable('replies', token!, 'viewBuffers', targets, moved, undefined, true)
    // What is left of the reservation was kept for the local visibility words:
    // it ends as they are charged.
    if (reservation) this.mainLedger.release('replies', reservation)
    if (moved === 0 || bandBytes + visibilityBytes > moved) this.mainLedger.chargeWhileHeld('viewBuffers', targets, bandBytes + visibilityBytes - moved)
    const replaced = this.viewHolds.get(sheetId)
    this.viewHolds.set(sheetId, this.mainLedger.hold(view)!)
    replaced?.release()
    this.notify()
    return view
  }
  /** Another owner of an adopted view or any of its buffers (a viewer's
   * committed display, a copy session): it stays charged until released.
   * Null for the base view, which owns no buffers. */
  holdView(view: RemoteSheetView): LedgerHold | null { return this.mainLedger.hold(view) }

  /** The host (OfficeFile) owns the worker document; a viewer that borrows this
   * source releases nothing, so a remount (or StrictMode's dev re-mount) keeps
   * reading the same open document. `close()` ends it. */
  async release(): Promise<void> {}
  async close(): Promise<void> {
    if (this.closed) return
    this.closed = true
    this.listeners.clear()
    for (const entry of this.structures.values()) if (entry.timer) clearTimeout(entry.timer)
    this.drawingCache.close()
    this.overflow?.cache.close()
    this.overflow = undefined
    // Owned caches go with the document: nothing answers from them after close.
    // Document-owned buffers are dropped here; their lifetime charges end when
    // nothing (not even a host still holding this closed workbook) reaches them.
    for (const store of this.stores.values()) store.clear()
    this.stores.clear()
    this.views.clear()
    for (const hold of this.viewHolds.values()) hold.release()
    this.viewHolds.clear()
    this.geometries.clear()
    this.columnMetrics = new WeakMap()
    this.layouts.clear()
    this.structures.clear()
    this.latestQueries.clear()
    this.failures.clear()
    try {
      await Promise.race([this.client.send({ kind: 'close', requestId: this.client.nextRequestId(), docId: this.docId }), new Promise(resolve => setTimeout(resolve, 1000))])
    } catch { /* the worker is gone */ }
    this.client.dispose()
    await this.onRelease()
  }
}

class StaleViewportPublicationError extends Error {
  readonly name = 'StaleViewportPublicationError'
  readonly code = 'stale-source'
  readonly retryable = true
  constructor() { super('Viewport row cache changed; re-stage the publication') }
}

interface RowEntry { chunks: Set<number>; cells: PagedCell[]; adoptionMeasurement?: number }
/** Per-frame page budgets are shared across sheet stores, so tile identity
 * includes the sheet as well as the complete logical box. */
export const pageTileKey = (sheetId: string, box: LogicalCellBox): string => JSON.stringify([sheetId, box.x, box.y, box.width, box.height])
/** One frame's cumulative page spending, in total and per tile key. */
export interface PageFrame { used: number; tiles: Map<string, number> }
/** Where a painting run's pages are charged. */
interface PageAdmission { frame: PageFrame; tile: string; ring: boolean }

/** One byte ledger for the painting row caches of every sheet of a document
 * (§12.4). Insertion order is recency; charging past the limit evicts the
 * least recently used rows of any sheet. A row larger than the whole budget is
 * never cached (its reads go to the worker each time). */
interface RowCharge { store: RemoteSheetStore; row: number; bytes: number; recency: number; touched: number }
export class RowCacheLedger {
  bytes = 0
  peakBytes = 0
  private reserved = 0
  private recency = 0
  private stateGeneration = 0
  /** Cache data and charge changes invalidate held publication snapshots;
   * reads, LRU touches and reservation bookkeeping do not. */
  get generation(): number { return this.stateGeneration }
  /** @internal Shared by row-data mutations and charge ownership changes. */
  advanceGeneration(): void {
    if (!Number.isSafeInteger(this.stateGeneration + 1)) throw new Error('Row cache generation exhausted')
    this.stateGeneration++
  }
  private nextRecency(): number {
    if (!Number.isSafeInteger(this.recency + 1)) throw new Error('Row cache recency exhausted')
    return ++this.recency
  }
  private readonly reservations = new Map<object, { bytes: number; pinned: ReadonlySet<string> }>()
  reserve(bytes: number, store?: RemoteSheetStore, rows: Iterable<number> = []): { close(): void } | undefined {
    if (!Number.isSafeInteger(bytes) || bytes < 0 || this.reserved + bytes > this.limitBytes) return undefined
    const pinned = new Set<string>()
    if (store) for (const row of rows) pinned.add(this.key(store, row))
    // Promise retained growth without evicting until publication charges it.
    // Abandoned stages and failed feasibility checks leave warm rows intact.
    if (this.bytes + this.reserved + bytes > this.limitBytes) {
      let protectedBytes = 0
      for (const key of pinned) protectedBytes += this.entries.get(key)?.bytes ?? 0
      // Only active pins can prevent eviction. Deduplicate against new pins
      // and earlier reservations without scanning residents or copying a union.
      for (const [token, item] of this.reservations) for (const key of item.pinned) {
        if (pinned.has(key)) continue
        let counted = false
        for (const [previousToken, previous] of this.reservations) {
          if (previousToken === token) break
          if (previous.pinned.has(key)) { counted = true; break }
        }
        if (!counted) protectedBytes += this.entries.get(key)?.bytes ?? 0
      }
      if (protectedBytes + this.reserved + bytes > this.limitBytes) return undefined
    }
    this.reserved += bytes
    const token = { close: () => {
      if (!this.reservations.delete(token)) return
      this.reserved -= bytes
    } }
    this.reservations.set(token, { bytes, pinned })
    return token
  }
  viewportRollbackBytes(): number { return 128 }
  /** Journal only the rows this publication can replace. Independent evictions
   * stay evicted; restoring their former accounting would create phantom owners.
   * @internal The no-argument whole-ledger form is retained for tests only. */
  beginViewportPublication(store?: RemoteSheetStore, rows: Iterable<number> = [], reverted: (row: number) => boolean = () => true): { rollback(): void; close(): void } {
    // Compatibility for explicit whole-ledger transactions. Tile publication
    // always supplies its touched rows and never snapshots the cache Map.
    if (!store) {
      const snapshot = new Map(this.entries)
      let live = true
      const close = () => { live = false; snapshot.clear() }
      return { close, rollback: () => {
        if (!live) return
        this.advanceGeneration()
        this.entries.clear()
        this.bytes = 0
        for (const [key, entry] of snapshot) { this.entries.set(key, entry); this.bytes += entry.bytes }
        this.report()
        close()
      } }
    }
    const entries = new Map<number, RowCharge | undefined>()
    for (const row of rows) {
      const old = this.entries.get(this.key(store, row))
      entries.set(row, old ? { ...old } : undefined)
    }
    let live = true
    const close = () => { live = false; entries.clear() }
    return { close, rollback: () => {
      if (!live) return
      this.advanceGeneration()
      const restored: Array<[string, RowCharge]> = []
      for (const [row, old] of entries) {
        if (!reverted(row)) continue
        const key = this.key(store, row), current = this.entries.get(key)
        if (current) { this.entries.delete(key); this.bytes -= current.bytes }
        if (old) {
          // Explicit touches after publication began own their newer recency.
          if (current && current.touched > old.touched) { old.recency = current.recency; old.touched = current.touched }
          restored.push([key, old])
          this.bytes += old.bytes
        }
      }
      if (!restored.length) { this.report(); close(); return }
      restored.sort((a, b) => a[1].recency - b[1].recency)
      // Rotate the existing Map once, merging only the journaled rows. The
      // initial count prevents the iterator from revisiting appended entries.
      const count = this.entries.size
      let index = 0
      const iterator = this.entries.entries()
      for (let i = 0; i < count; i++) {
        const [key, entry] = iterator.next().value as [string, RowCharge]
        this.entries.delete(key)
        while (index < restored.length && restored[index][1].recency < entry.recency) {
          const [restoredKey, restoredEntry] = restored[index++]
          this.entries.set(restoredKey, restoredEntry)
        }
        this.entries.set(key, entry)
      }
      while (index < restored.length) {
        const [key, entry] = restored[index++]
        this.entries.set(key, entry)
      }
      this.report()
      close()
    } }
  }
  private isPinned(key: string, pinned: ReadonlySet<string>): boolean {
    if (pinned.has(key)) return true
    for (const item of this.reservations.values()) if (item.pinned.has(key)) return true
    return false
  }
  private evict(limit: number, pinned: ReadonlySet<string> = new Set()): void {
    for (const [key, entry] of this.entries) {
      if (this.bytes <= limit) break
      const protectedRow = this.isPinned(key, pinned)
      if (protectedRow) continue
      this.entries.delete(key)
      this.bytes -= entry.bytes
      entry.store.drop(entry.row)
    }
  }
  private readonly entries = new Map<string, RowCharge>()
  constructor(readonly limitBytes: number, private readonly main?: MainMemoryLedger) {}
  private report(): void { this.main?.set('rowCache', 'rows', this.bytes) }
  private key(store: RemoteSheetStore, row: number): string { return `${store.sheetId}\u0000${row}` }
  /** @internal Current retained ownership, including ordinary loads while a stage waits. */
  rowBytes(store: RemoteSheetStore, row: number): number | undefined { return this.entries.get(this.key(store, row))?.bytes }
  charge(store: RemoteSheetStore, row: number, bytes: number, reservation?: { close(): void }): void {
    const key = this.key(store, row)
    const old = this.entries.get(key), recency = this.nextRecency()
    this.advanceGeneration()
    if (old) { this.bytes -= old.bytes; this.entries.delete(key) }
    this.entries.set(key, { store, row, bytes, recency, touched: old?.touched ?? 0 })
    this.bytes += bytes
    const own = reservation ? this.reservations.get(reservation)?.bytes ?? 0 : 0
    this.evict(this.limitBytes - this.reserved + own)
    // A concurrent ordinary read cannot consume capacity promised to stages.
    if (this.bytes + this.reserved - own > this.limitBytes && this.entries.get(key)) {
      this.entries.delete(key)
      this.bytes -= bytes
      store.drop(row)
    }
    this.peakBytes = Math.max(this.peakBytes, this.bytes)
    this.report()
  }
  touch(store: RemoteSheetStore, row: number): void {
    const key = this.key(store, row)
    const entry = this.entries.get(key)
    if (entry) { entry.recency = this.nextRecency(); entry.touched = entry.recency; this.entries.delete(key); this.entries.set(key, entry) }
  }
  release(store: RemoteSheetStore, row: number): void {
    const key = this.key(store, row)
    const entry = this.entries.get(key)
    if (entry) { this.advanceGeneration(); this.entries.delete(key); this.bytes -= entry.bytes; this.report() }
  }
}

/** SourceStore proxy of one remote sheet plus its bounded painting row cache. */
export class RemoteSheetStore implements SourceStore {
  private readonly rows = new Map<number, RowEntry>()
  private readonly viewportStages = new Set<() => void>()
  private readonly loading = new Set<string>()
  // Mutable state lives in one shared holder: find wraps stores with
  // Object.create, and a primitive written through a wrapper would diverge.
  private readonly state: { retry?: ReturnType<typeof setTimeout>; extent: SourceStore['extent']; epoch: number; adoptionMeasurement: number } = { extent: { rows: MAX_ROWS, cols: MAX_COLS }, epoch: 0, adoptionMeasurement: 0 }

  constructor(private readonly workbook: RemoteWorkbook, readonly sheetId: string) {
    this.relayout()
  }
  get extent(): SourceStore['extent'] { return this.state.extent }
  /** Re-reads the used range from the workbook's current layout. */
  relayout(): void {
    const used = this.workbook.layout(this.sheetId).used
    this.state.extent = { rows: MAX_ROWS, cols: MAX_COLS, ...(used ? { used } : {}) }
  }
  get stamp(): SourceStamp {
    const identity = this.workbook.identity(this.sheetId)
    return { docId: identity.docId, sheetId: identity.sheetId, valueGen: identity.valueGen }
  }
  get valuePolicy(): ValuePolicy { return this.workbook.valuePolicy }
  /** Source guards every call applies before any cache or command (§§4.1, 10.1). */
  private guard(options: ReadOptions): SourceFailure | undefined {
    if (this.workbook.isClosed) return failure('stale-source', 'The document was closed')
    if (options.signal.aborted) return failure('cancelled', 'Read cancelled', true)
    const owner = this.stamp
    if (options.stamp.docId !== owner.docId || options.stamp.sheetId !== owner.sheetId || options.stamp.valueGen !== owner.valueGen) {
      return failure('stale-source', 'Read stamp does not match the source owner', true)
    }
    return undefined
  }

  /** Sends one `read` for the current view owner. */
  private async read(query: SourceReadQuery, options: ReadOptions): Promise<SourceReadResult | SourceFailure> {
    const refused = this.guard(options)
    if (refused) return refused
    const owner = this.workbook.identity(this.sheetId)
    const out = await this.workbook.request({ kind: 'read', requestId: this.workbook.nextRequestId(), owner, query }, options.signal)
    if (out === CANCELLED) return failure('cancelled', 'Read cancelled', true)
    if (out === CLOSED) return failure('stale-source', 'The document was closed')
    switch (out.kind) {
      case 'complete': return out.result.kind === 'read' ? out.result.result : failure('page-error', 'Unexpected read reply')
      case 'cancelled': return failure('cancelled', out.diagnostics[0]?.message ?? 'Read cancelled', true)
      case 'incomplete': return failure('incomplete', 'Source range is incomplete', true)
      case 'quota': return failure('quota', 'Worker memory budget exceeded')
      case 'error': return failure(out.code, out.diagnostics[0]?.message ?? `Read failed (${out.code})`, out.retryable)
      default: return failure('page-error', 'Unexpected read reply')
    }
  }
  private failedPage<T>(error: SourceFailure): Page<T> {
    return { items: [], completeness: 'incomplete', stamp: this.stamp, error }
  }

  async readCell(at: SourceCell, options: ReadOptions): Promise<CellRead> {
    if (!isValidCoordinate(at.row, at.col)) return { kind: 'error', code: 'invalid-coordinate', retryable: false }
    const refused = this.guard(options)
    if (refused) return { kind: 'error', code: refused.code, retryable: refused.retryable }
    // Copy and inspection read cell by cell in row-major order: one bounded
    // chunk page answers the following cells of the row locally. Stated limit:
    // copy therefore costs one worker round trip per new row chunk.
    const chunk = Math.floor(at.col / ROW_CHUNK_COLS)
    if (!this.known(at.row, chunk, chunk)) await this.load([at.row], at.col, at.col, options.signal, options.stamp)
    const after = this.guard(options)
    if (after) return { kind: 'error', code: after.code, retryable: after.retryable }
    if (this.known(at.row, chunk, chunk)) {
      const cell = this.rows.get(at.row)!.cells.find(item => item.col === at.col)
      if (!cell) return { kind: 'blank' }
      const { unavailable, ref: _ref, style: _style, ...payload } = cell
      return unavailable ? { kind: 'unavailable', cell: payload, reason: 'formula-value-unavailable' } : { kind: 'cell', cell: payload }
    }
    const out = await this.read({ op: 'cell', at }, options)
    if ('code' in out) return { kind: 'error', code: out.code, retryable: out.retryable }
    return out.op === 'cell' ? out.read : { kind: 'error', code: 'page-error', retryable: false }
  }
  async readRows(range: SourceRect, cursor: string | undefined, options: ReadOptions): Promise<Page<SparseRowPage>> {
    const out = await this.read({ op: 'rows', range, ...(cursor !== undefined ? { cursor } : {}) }, options)
    if ('code' in out) return this.failedPage(out)
    return out.op === 'rows' ? out.page : this.failedPage(failure('page-error', 'Unexpected read reply'))
  }
  async queryColumn(request: ColumnQuery, options: ReadOptions): Promise<Page<ColumnEntry>> {
    const refused = this.guard(options)
    if (refused) return this.failedPage(refused)
    const owner = this.workbook.identity(this.sheetId)
    // The caller's row membership is borrowed: the command carries its own copy
    // (the client transfers whole-buffer typed arrays).
    const query = structuredClone(request)
    const out = await this.workbook.request({ kind: 'valueQuery', requestId: this.workbook.nextRequestId(), owner, query }, options.signal)
    if (out === CANCELLED) return this.failedPage(failure('cancelled', 'Query cancelled', true))
    if (out === CLOSED) return this.failedPage(failure('stale-source', 'The document was closed'))
    if (out.kind === 'complete' && out.result.kind === 'valueQuery') return out.result.page
    if (out.kind === 'error') return this.failedPage(failure(out.code, out.diagnostics[0]?.message ?? 'Value query failed', out.retryable))
    return this.failedPage(failure(out.kind === 'cancelled' ? 'cancelled' : 'incomplete', 'Value query did not complete', true))
  }
  async *scanPopulated(request: PopulatedQuery, options: ReadOptions): AsyncIterable<Page<AuthoritativeCell>> {
    let cursor: string | undefined, finished = false
    try {
      for (;;) {
        const out = await this.read({ op: 'scan', request: structuredClone(request), ...(cursor !== undefined ? { cursor } : {}) }, options)
        // A failed or cancelled page already ended its continuation in the worker.
        if ('code' in out) { finished = true; yield this.failedPage(out); return }
        if (out.op !== 'scan') { finished = true; yield this.failedPage(failure('page-error', 'Unexpected read reply')); return }
        const { nextCursor, ...page } = out.page
        cursor = nextCursor
        if (nextCursor === undefined || page.completeness === 'incomplete') finished = true
        yield page
        if (finished) return
      }
    } finally {
      // An abandoned iteration (return(), a thrown consumer) releases its
      // worker continuation instead of leaving it to the LRU cap (§11.4).
      if (!finished && cursor !== undefined) void this.releaseScan(cursor)
    }
  }
  /** SourceStore.settle: a consumer decoded a page this store returned. */
  settle(payload: object): void { this.workbook.mainLedger.settle(payload) }
  private async releaseScan(cursor: string): Promise<void> {
    if (this.workbook.isClosed) return
    await this.workbook.request({ kind: 'read', requestId: this.workbook.nextRequestId(), owner: this.workbook.identity(this.sheetId), query: { op: 'releaseScan', cursor } })
  }
  /** S10.2: range detection runs next to the worker's store; one read, no
   * strip pages across the worker boundary. Failures become the detector's own
   * typed outcomes (cancelled, or a source-error refusal). */
  async detectRange(kind: 'sort' | 'filter', selection: GridSelection, override: { headerRows: 0 | 1; totalsRows?: 0 } | null, options: ReadOptions): Promise<RangeDetection> {
    const out = await this.read({ op: 'detectRange', kind, selection, override }, options)
    if ('code' in out) return out.code === 'cancelled' ? { kind: 'cancelled' } : { kind: 'refused', refusal: { reason: 'source-error', code: out.code, retryable: out.retryable, message: out.message } }
    if (out.op !== 'detectRange') return { kind: 'refused', refusal: { reason: 'source-error', code: 'page-error', retryable: false, message: 'Unexpected read reply' } }
    return out.detection
  }
  async queryStructural(rect: SourceRect, options: ReadOptions): Promise<StructuralResult> {
    const out = await this.read({ op: 'structural', rect }, options)
    if (!('code' in out) && out.op === 'structural') return out.result
    const error = 'code' in out ? out : failure('page-error', 'Unexpected read reply')
    return { stamp: this.stamp, completeness: 'incomplete', merges: [], tables: [], spills: [], drawings: [], missing: [rect], error }
  }
  async queryOverview(request: OverviewQuery, options: ReadOptions): Promise<Page<OverviewGroup>> {
    const budget = WORKER_QUERY_BYTES[this.workbook.profile]
    const cap = VIEWPORT_DETAIL_LIMITS[this.workbook.profile].overviewGroups
    const refused = this.guard(options)
    if (refused) return this.failedPage(refused)
    if (!Array.isArray(request.groups) || request.groups.length > cap || queryPayloadBytes(request, budget) > budget) return this.failedPage(failure('resource-limit', 'Overview request exceeds its bounded wire budget'))
    const charge = `overview-clone:${this.workbook.nextRequestId()}`
    this.workbook.mainLedger.set('replies', charge, queryPayloadBytes(request, budget))
    try {
      const out = await this.read({ op: 'overview', request: structuredClone(request) }, options)
      if ('code' in out) return this.failedPage(out)
      return out.op === 'overview' ? out.page : this.failedPage(failure('page-error', 'Unexpected overview reply'))
    } finally { this.workbook.mainLedger.release('replies', charge) }
  }
  async ensureComplete(rect: SourceRect, options: ReadOptions): Promise<RangeCompletion> {
    const out = await this.read({ op: 'complete', rect }, options)
    if (!('code' in out) && out.op === 'complete') return out.result
    const error = 'code' in out ? out : failure('page-error', 'Unexpected read reply')
    return { range: rect, completeness: 'incomplete', stamp: this.stamp, missing: [rect], error }
  }
  async release(): Promise<void> {}

  // ---- painting row cache ----

  /** Adopt resolved viewport cells for accessibility/inspection. This proves
   * only these coordinates, never the blank remainder of a row-cache chunk. */
  adoptViewport(cells: readonly import('../xlsx/view/contracts').ViewportCell[]): void {
    const token = this.stageViewport(cells, new AbortController().signal)
    try { token?.commit(() => {}) } finally { token?.close() }
  }
  /** Captured before the viewport await. Clear closes outstanding requests as
   * well as stages, so a delivered old reply cannot acquire a fresh epoch. */
  captureViewportRequest(signal: AbortSignal): { live(): boolean; close(): void } {
    const epoch = this.state.epoch
    let open = true
    const close = () => { if (!open) return; open = false; this.viewportStages.delete(close); signal.removeEventListener('abort', close) }
    this.viewportStages.add(close)
    signal.addEventListener('abort', close, { once: true })
    if (signal.aborted) close()
    return { close, live: () => open && !signal.aborted && !this.workbook.isClosed && epoch === this.state.epoch }
  }
  /** Scalar estimate before grouping, copying or sorting. Existing rows are
   * still owned by the row ledger and rollback keeps them alive during staging. */
  viewportRollbackBytes(): number { return this.workbook.ledger.viewportRollbackBytes() }
  viewportAdoptionBytes(cells: readonly import('../xlsx/view/contracts').ViewportCell[]): number {
    let bytes = 128
    // The numeric marker fits the accounted 64 B row record. It avoids a
    // payload-sized Set before joint admission, including for reordered rows.
    const measurement = this.state.adoptionMeasurement + 1
    if (!Number.isSafeInteger(measurement)) throw new Error('Viewport adoption measurement revision exhausted')
    this.state.adoptionMeasurement = measurement
    for (const item of cells) {
      if (item.kind !== 'cell' && item.kind !== 'unavailable') continue
      bytes += 256 + 3 * cachedCellBytes(item.cell)
      const old = this.rows.get(item.source.row)
      // New rows have no old payload to scan. Their per-cell allowance also
      // covers a journal slot if several incoming records share that new row.
      bytes += 192
      if (old && old.adoptionMeasurement !== measurement) {
        old.adoptionMeasurement = measurement
        for (const cell of old.cells) bytes += 128 + cachedCellBytes(cell)
      }
    }
    return bytes
  }
  /** The caller has admitted the joint main allocation before invoking this
   * lower-cache admission. A declined row adoption does not decline paint.
   * Invalidation is deliberately coarse and document-wide: even another sheet's
   * row change refuses a held token. Continuous loading can starve deferred
   * adoption; callers should re-stage and commit synchronously when possible. */
  stageViewport(cells: readonly import('../xlsx/view/contracts').ViewportCell[], signal: AbortSignal, request?: { live(): boolean }): { readonly error?: SourceFailure; commit(publish: () => void): boolean; close(): void } | undefined {
    // Staging/journal scratch belongs to the caller's joint main reservation.
    // The lower row ledger promises only the merged cache's retained growth.
    let reservation: { close(): void } | undefined
    const epoch = this.state.epoch, owner = this.workbook.identity(this.sheetId), generation = this.workbook.ledger.generation
    let error: SourceFailure | undefined
    const stale = () => {
      const cause = new StaleViewportPublicationError()
      // Every reader of the token shares this one failure: freeze it.
      error ??= Object.freeze(failure(cause.code, cause.message, cause.retryable))
      return cause
    }
    let live = true, committing = false
    const staged = new Map<number, RowEntry>(), previous = new Map<number, RowEntry | undefined>()
    const close = () => { if (!committing) { staged.clear(); previous.clear() }; if (!live) return; live = false; reservation?.close(); signal.removeEventListener('abort', close); this.viewportStages.delete(close) }
    this.viewportStages.add(close)
    signal.addEventListener('abort', close, { once: true })
    try {
      if (signal.aborted || this.workbook.isClosed || request && !request.live()) { close(); return undefined }
      const incoming = new Map<number, Map<number, Extract<import('../xlsx/view/contracts').ViewportCell, { kind: 'cell' | 'unavailable' }>>>()
      for (const item of cells) {
        if (item.kind !== 'cell' && item.kind !== 'unavailable') continue
        let row = incoming.get(item.source.row)
        if (!row) { row = new Map(); incoming.set(item.source.row, row) }
        row.set(item.source.col, item)
      }
      let growth = 0
      for (const [row, items] of incoming) {
        let bytes = 64
        for (const old of this.rows.get(row)?.cells ?? []) if (!items.has(old.col)) bytes += cachedCellBytes(old)
        for (const item of items.values()) bytes += cachedCellBytes(item.cell)
        growth += bytes - (this.workbook.ledger.rowBytes(this, row) ?? 0)
      }
      reservation = this.workbook.ledger.reserve(Math.max(0, growth), this, incoming.keys())
      if (!reservation) { close(); return undefined }
      for (const [row, items] of incoming) {
        const old = this.rows.get(row), existing = old?.cells ?? []
        previous.set(row, old)
        const added: PagedCell[] = []
        let sorted = true
        for (const item of items.values()) {
          const style = this.workbook.styleOf(item.cell.styleIndex)
          const cell = { ...item.cell, ...(item.cell.sharedFormula ? { sharedFormula: { ...item.cell.sharedFormula } } : {}), ref: formatA1(item.source), value: item.cell.value ?? null, ...(item.kind === 'unavailable' ? { unavailable: true } : {}), ...(style ? { style } : {}) } as PagedCell
          if (added.length && added[added.length - 1].col > cell.col) sorted = false
          added.push(cell)
        }
        // Ordinary viewport rows are ordered already. Only reordered input
        // needs sorting; the old row stays sorted and is merged once.
        if (!sorted) added.sort((a, b) => a.col - b.col)
        const merged: PagedCell[] = []
        let i = 0, j = 0
        while (i < existing.length && j < added.length) {
          const oldCol = existing[i].col, newCol = added[j].col
          if (oldCol < newCol) merged.push(existing[i++])
          else {
            if (oldCol === newCol) i++
            merged.push(added[j++])
          }
        }
        while (i < existing.length) merged.push(existing[i++])
        while (j < added.length) merged.push(added[j++])
        staged.set(row, { ...(old ?? { chunks: new Set<number>() }), cells: merged })
      }
      incoming.clear()
      return { close, get error() { return error }, commit: publish => {
        if (!live || signal.aborted || this.workbook.isClosed || request && !request.live() || epoch !== this.state.epoch || this.workbook.identity(this.sheetId) !== owner) { close(); return false }
        if (this.workbook.ledger.generation !== generation) { stale(); close(); return false }
        committing = true
        const reverted = new Set<number>()
        const accounting = this.workbook.ledger.beginViewportPublication(this, staged.keys(), row => reverted.has(row))
        try {
          const charges = [...staged].map(([row, next]) => {
            const oldBytes = this.workbook.ledger.rowBytes(this, row)
            const bytes = 64 + next.cells.reduce((sum, cell) => sum + cachedCellBytes(cell), 0)
            return { row, bytes, growth: bytes - (oldBytes ?? 0) }
          }).sort((a, b) => Number(a.growth > 0) - Number(b.growth > 0))
          if (staged.size) this.workbook.ledger.advanceGeneration()
          for (const [row, next] of staged) this.rows.set(row, next)
          const publicationGeneration = this.workbook.ledger.generation
          publish()
          // Publication callbacks may replace cache data or charge ownership
          // before this batch starts charging; do not overwrite that newer owner.
          if (this.workbook.ledger.generation !== publicationGeneration) throw stale()
          // Charge shrinking rows first: their freed bytes fund growing rows
          // without evicting warm data for an intermediate publication shape.
          for (const { row, bytes } of charges) {
            // A prior row's charge callback can replace a later row in the batch.
            if (this.rows.get(row) !== staged.get(row)) throw stale()
            this.workbook.ledger.charge(this, row, bytes, reservation)
            // Reporting that charge may cancel, clear or replace the source owner.
            if (!live || signal.aborted || this.workbook.isClosed || request && !request.live() || epoch !== this.state.epoch || this.workbook.identity(this.sheetId) !== owner) throw new Error('Viewport owner changed during publication')
          }
          // An empty batch still invokes publish, so it needs the same owner guard.
          if (!live || signal.aborted || this.workbook.isClosed || request && !request.live() || epoch !== this.state.epoch || this.workbook.identity(this.sheetId) !== owner) throw new Error('Viewport owner changed during publication')
          // The last charge callback can replace an already-charged row too.
          for (const [row, next] of staged) if (this.rows.get(row) !== next) throw stale()
          accounting.close()
          close()
          return true
        } catch (cause) {
          // A cache clear owns the newer epoch and its released rows. Restoring
          // the publication snapshot would revive accounting for those rows.
          if (epoch !== this.state.epoch) accounting.close()
          else {
            for (const [row, old] of previous) {
              if (this.rows.get(row) !== staged.get(row)) {
                if (!this.rows.has(row)) this.workbook.ledger.release(this, row)
                continue
              }
              reverted.add(row)
              if (old) this.rows.set(row, old)
              else this.rows.delete(row)
            }
            accounting.rollback()
          }
          close()
          if (cause instanceof StaleViewportPublicationError) { error ??= Object.freeze(failure(cause.code, cause.message, cause.retryable)); return false }
          throw cause
        } finally { committing = false; close() }
      } }
    } catch (cause) { close(); throw cause }
  }
  /** Cached populated cells of a source row (ascending columns); only the
   * known chunks contribute. Unknown chunks are reported by `pendingIn`. */
  cellsOf(row: number): readonly PagedCell[] | undefined {
    if (this.workbook.isClosed) return undefined
    const entry = this.rows.get(row)
    if (entry) this.workbook.ledger.touch(this, row)
    return entry && entry.cells.length ? entry.cells : undefined
  }
  /** Up to `limit` source rows the row cache holds now, in cache order; no
   * read is issued (bounded auto-fit, spec §5.6). */
  knownRows(limit: number): number[] {
    const out: number[] = []
    for (const row of this.rows.keys()) { if (out.length >= limit) break; out.push(row) }
    return out
  }
  /** Whether the cell's row chunk has been read (blank or populated). */
  isKnown(at: SourceCell): boolean {
    const chunk = Math.floor(at.col / ROW_CHUNK_COLS)
    return this.known(at.row, chunk, chunk)
  }
  cellAt(at: SourceCell): PagedCell | undefined {
    return this.workbook.isClosed ? undefined : this.rows.get(at.row)?.cells.find(cell => cell.col === at.col)
  }
  /** Cached rows (tests and diagnostics). */
  get cachedRows(): number { return this.rows.size }
  private known(row: number, k0: number, k1: number): boolean {
    const entry = this.rows.get(row)
    if (!entry) return false
    for (let k = k0; k <= k1; k++) if (!entry.chunks.has(k)) return false
    return true
  }
  /** Drops one row (ledger eviction). */
  drop(row: number): void { if (this.rows.delete(row)) this.workbook.ledger.advanceGeneration() }
  /** Releases the whole cache (inactive sheet, close); in-flight loads land nowhere. */
  clear(): void {
    for (const close of this.viewportStages) close()
    for (const row of this.rows.keys()) this.workbook.ledger.release(this, row)
    this.rows.clear()
    this.state.epoch++
  }
  /**
   * Logical boxes inside `box` whose detail rows are not yet known, in display
   * order; requests them as a side effect. Detail admission comes first (spec
   * §11.1): rows/columns projecting below four device pixels at `scale` are
   * summarized by the planner and never read. Requests are coalesced into runs
   * and every page (continuations included) is admitted against the frame's
   * shared per-tile/per-frame allowance (RemoteWorkbook.admitPages); a ring
   * tile (`ring`) stays within the ring share. The rest stay pending and are
   * asked for by a later frame.
   */
  pendingIn(metrics: ProjectedSheetMetrics, box: LogicalCellBox, scale: number, ring = false): LogicalCellBox[] {
    const found = this.uncoveredIn(metrics, box, scale)
    if (!found || 'oversized' in found) return []
    const { k0, k1, missing, out } = found
    if (missing.length) {
      const admission: PageAdmission = { frame: this.workbook.currentPageFrame, tile: pageTileKey(this.sheetId, box), ring }
      for (const run of coalesce(missing)) {
        if (!this.workbook.admitPages(admission.frame, admission.tile, 1, ring, false)) break
        void this.loadRun(run.r0, run.r1, k0, k1, new AbortController().signal, this.stamp, admission)
      }
    }
    return out
  }
  /** What the row cache can say about a box, without requesting anything:
   * 'proven' when it already knows every detail cell (or the box has none),
   * 'unknown' when `pendingIn` would read rows for it, and 'oversized' when it
   * spans more rows than the row cache may cover, so only a worker plan can. */
  coverage(metrics: ProjectedSheetMetrics, box: LogicalCellBox, scale: number): 'proven' | 'unknown' | 'oversized' {
    const found = this.uncoveredIn(metrics, box, scale)
    if (!found) return 'proven'
    if ('oversized' in found) return 'oversized'
    return found.missing.length ? 'unknown' : 'proven'
  }
  /** The pendingIn query without the loads: boxes and missing rows, null when
   * the box has no detail rows or columns to read, or oversized when it spans
   * more rows than the row cache may cover. */
  private uncoveredIn(metrics: ProjectedSheetMetrics, box: LogicalCellBox, scale: number):
    { k0: number; k1: number; missing: number[]; out: LogicalCellBox[] } | { oversized: true } | null {
    const detail = (sizeCss: number) => sizeCss * scale >= MIN_DETAIL_DEVICE_PX - 1e-9
    const rows = metrics.rows.visibleBetween(box.y, box.y + box.height)
    if (rows.count === 0) return null
    const cols = metrics.cols.visibleBetween(Math.max(0, box.x - TEXT_OVERFLOW_REACH_CSS), box.x + box.width + TEXT_OVERFLOW_REACH_CSS)
    let c0 = -1, c1 = -1
    for (const run of cols.runs) { if (!detail(run.sizeCss)) continue; if (c0 < 0) c0 = run.start; c1 = run.end }
    if (c0 < 0) return null
    // Too many rows to cover only matters when some of them are detail rows;
    // an all-subpixel band is summarized by the planner and never read.
    if (rows.count > MAX_PENDING_ROWS) {
      for (const run of rows.runs) if (detail(run.sizeCss)) return { oversized: true }
      return null
    }
    const k0 = Math.floor(c0 / ROW_CHUNK_COLS), k1 = Math.floor(c1 / ROW_CHUNK_COLS)
    const out: LogicalCellBox[] = []
    const missing: number[] = []
    for (const run of rows.runs) {
      if (!detail(run.sizeCss)) continue
      for (let slot = run.start; slot <= run.end; slot++) {
        const source = metrics.order.displayToSource(slot)
        if (this.known(source, k0, k1)) continue
        missing.push(source)
        const y = metrics.rows.start(slot), height = metrics.rows.size(slot)
        const last = out[out.length - 1]
        if (last && Math.abs(last.y + last.height - y) < 1e-6) last.height += height
        else out.push({ x: box.x, y, width: box.width, height })
      }
    }
    return { k0, k1, missing, out }
  }
  /** Loads the chunks covering [c0,c1] of `rows` (source rows) in runs. */
  async load(rows: readonly number[], c0: number, c1: number, signal?: AbortSignal, stamp: SourceStamp = this.stamp): Promise<void> {
    const k0 = Math.floor(c0 / ROW_CHUNK_COLS), k1 = Math.floor(c1 / ROW_CHUNK_COLS)
    const runs = coalesce(rows.filter(row => !this.known(row, k0, k1)), 0)
    await Promise.all(runs.map(run => this.loadRun(run.r0, run.r1, k0, k1, signal ?? new AbortController().signal, stamp)))
  }
  private async loadRun(r0: number, r1: number, k0: number, k1: number, signal: AbortSignal, stamp: SourceStamp, admission?: PageAdmission): Promise<void> {
    const key = `${stamp.valueGen}:${r0}:${r1}:${k0}:${k1}`
    if (this.loading.has(key)) return
    // A painting run spends its frame's allowance page by page; an explicit
    // read (inspection, copy) holds an outstanding credit only.
    const admit = () => admission ? this.workbook.admitPages(admission.frame, admission.tile, 1, admission.ring) : this.workbook.reservePages(1)
    if (admission) { if (!admit()) return } else this.workbook.pagesInFlight++
    this.loading.add(key)
    let held = 1
    const epoch = this.state.epoch
    const range = { r0, r1, c0: k0 * ROW_CHUNK_COLS, c1: Math.min(MAX_COLS - 1, (k1 + 1) * ROW_CHUNK_COLS - 1) }
    let changed = false
    try {
      let cursor: string | undefined
      // Pages cover the run in ascending rows (§4.1 ordering contract): every
      // row up to a page's last item, or the run end on the final page, is known.
      let covered = r0 - 1
      for (let pages = 0; pages < PAGES_PER_RUN; pages++) {
        // A continuation is another page request: without a credit the run
        // stops here, its later rows stay unknown and a later frame asks again.
        if (pages > 0) { if (!admit()) break; held++ }
        const page = await this.readRows(range, cursor, { signal, stamp })
        // A page for an older owner, a cleared cache or a closed document never lands.
        if (this.workbook.isClosed || epoch !== this.state.epoch || page.stamp.valueGen !== this.stamp.valueGen) return
        if (page.completeness !== 'complete') {
          if (page.completeness === 'loading') this.scheduleRetry()
          return
        }
        const last = page.nextCursor === undefined ? r1 : page.items.length ? Math.min(r1, page.items[page.items.length - 1].row) : covered
        const byRow = new Map(page.items.map(item => [item.row, item.cells]))
        for (let row = covered + 1; row <= last; row++) changed = this.put(row, k0, k1, byRow.get(row) ?? []) || changed
        // The row cache now holds (and charges) its own copies of these cells.
        this.workbook.mainLedger.settle(page)
        covered = Math.max(covered, last)
        cursor = page.nextCursor
        if (cursor === undefined) break
      }
    } finally {
      this.loading.delete(key)
      this.workbook.releasePages(held)
      if (changed) this.workbook.notify()
    }
  }
  /** Caches one row's chunk; false when the row alone exceeds the ledger. */
  private put(row: number, k0: number, k1: number, cells: readonly AuthoritativeCell[]): boolean {
    const entry = this.rows.get(row) ?? { chunks: new Set<number>(), cells: [] as PagedCell[] }
    const lo = k0 * ROW_CHUNK_COLS, hi = (k1 + 1) * ROW_CHUNK_COLS - 1
    const kept = entry.cells.filter(cell => cell.col < lo || cell.col > hi)
    const added = cells.map(cell => {
      const style = this.workbook.styleOf(cell.styleIndex)
      return { ...cell, ref: formatA1(cell), value: cell.value ?? null, styleIndex: cell.styleIndex ?? 0, ...(style ? { style } : {}) } as PagedCell
    })
    const next = [...kept, ...added].sort((a, b) => a.col - b.col)
    const bytes = 64 + next.reduce((sum, cell) => sum + cachedCellBytes(cell), 0)
    if (bytes > this.workbook.ledger.limitBytes) {
      this.workbook.ledger.release(this, row)
      this.rows.delete(row)
      return false
    }
    entry.cells = next
    for (let k = k0; k <= k1; k++) entry.chunks.add(k)
    this.rows.set(row, entry)
    this.workbook.ledger.charge(this, row, bytes)
    return this.rows.has(row)
  }
  private scheduleRetry(): void {
    if (this.state.retry !== undefined) return
    this.state.retry = setTimeout(() => { this.state.retry = undefined; if (!this.workbook.isClosed) this.workbook.notify() }, LOADING_RETRY_MS)
  }
}

/** Ascending unique rows → runs, merging runs at most `gap` rows apart (one
 * bounded read covers the gap; its rows become known too). */
function coalesce(rows: readonly number[], gap = COALESCE_GAP_ROWS): Array<{ r0: number; r1: number }> {
  const sorted = [...new Set(rows)].sort((a, b) => a - b)
  const runs: Array<{ r0: number; r1: number }> = []
  for (const row of sorted) {
    const last = runs[runs.length - 1]
    if (last && row - last.r1 <= gap + 1) last.r1 = row
    else runs.push({ r0: row, r1: row })
  }
  return runs
}
