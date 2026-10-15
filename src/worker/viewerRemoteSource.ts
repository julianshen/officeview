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
  ProfileId, GridSelection,
} from '../xlsx/view/contracts'
import type { RangeDetection } from '../xlsx/view/rangeDetection'
import { createDocumentOverflowCache } from '../xlsx/view/documentOverflow'
import type { OverflowCache } from '../xlsx/view/overflowCache'
import { MainMemoryLedger, axisBytes, estimatePayloadBytes, leaseTargets } from '../xlsx/view/mainLedger'
import type { XlsxCell, XlsxCellStyle, XlsxMergeRange } from '../xlsx/types'
import { MAX_COLS, MAX_ROWS, formatA1, isValidCoordinate, type SourceCell, type SourceRect } from '../xlsx/view/coordinates'
import { createAxisMetrics } from '../xlsx/view/axes'
import { createIdentityOrder, createRowOrder } from '../xlsx/view/order'
import { createIntervalSet, type CompactRowVisibility } from '../xlsx/view/visibility'
import type { ProjectedSheetMetrics } from '../xlsx/view/projection'
import { savedValuesStatus, type ValueStatus } from '../xlsx/view/values'
import { MIN_DETAIL_DEVICE_PX, SHEET_DEFAULT_COL_CSS, SHEET_DEFAULT_ROW_CSS, TEXT_OVERFLOW_REACH_CSS, VIEWPORT_DETAIL_LIMITS } from '../xlsx/view/viewport'
import { MAIN_THREAD_BUDGETS } from '../xlsx/view/tileCache'
import type { ViewerClient } from './viewerClient'
import { initialIdentity, initialWorkbookOwner, visibilityFromDisplayBitsets, type CommandOf } from './viewerProtocol'
import type { SheetGeometry, WorkbookSource } from './viewerSession'
import { RemoteDrawingCache, type PagedSheetDrawings } from './viewerRemoteDrawings'
import { drawingPayloadBudget } from '../xlsx/view/drawingPayload'

export const ROW_CHUNK_COLS = 256
/** One document-wide byte ledger for every sheet's painting row cache: a
 * quarter of the main-thread profile budget (§12.4; D 16 MiB, M 8 MiB). */
export function rowCacheBudget(profile: ProfileId): number { return MAIN_THREAD_BUDGETS[profile] / 4 }
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
type RemoteKind = 'read' | 'valueQuery' | 'sort' | 'filter' | 'prepareColumns'

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
  private readonly geometries = new Map<string, SheetGeometry>()
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
  /** Requests currently in flight for painting loads (all sheets). */
  loadsInFlight = 0
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
  geometry(sheetId: string): SheetGeometry {
    let geometry = this.geometries.get(sheetId)
    if (!geometry) {
      const layout = this.layout(sheetId)
      geometry = Object.freeze({
        rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: SHEET_DEFAULT_ROW_CSS, overrides: layout.rowOverrides }),
        cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: SHEET_DEFAULT_COL_CSS, overrides: layout.colOverrides, ...(layout.hiddenCols?.length ? { collapsed: layout.hiddenCols } : {}) }),
        rowOverrides: layout.rowOverrides, colOverrides: layout.colOverrides, baseHidden: createIntervalSet(layout.hidden),
        ...(layout.hiddenCols?.length ? { hiddenCols: layout.hiddenCols } : {}),
      })
      this.geometries.set(sheetId, geometry)
      // Axis summaries and authored-hidden intervals live as long as anything holds this geometry.
      this.mainLedger.chargeWhileReachable('geometry', [geometry], axisBytes(geometry.rows) + axisBytes(geometry.cols) + layout.hidden.length * 32 + (layout.hiddenCols?.length ?? 0) * 32)
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
  async request<K extends RemoteKind>(command: CommandOf<K>, signal?: AbortSignal, progress?: (completed: number, total: number) => void): Promise<Settled<K>> {
    if (this.closed) return CLOSED
    if (signal?.aborted) return CANCELLED
    const head = command as unknown as { kind: K; requestId: string; query?: { op: string } }
    const key = head.query ? `read:${head.query.op}` : head.kind
    this.stats.commands[key] = (this.stats.commands[key] ?? 0) + 1
    const unsubscribe = progress ? this.client.onProgress(head.requestId, p => progress(p.completed, p.total ?? p.completed)) : undefined
    const sent: Promise<Settled<K>> = this.client.send<K>(command).then(reply => {
      // Charged until no part of the payload is reachable: a consumer that keeps
      // a page (a scan generator across tasks) keeps the charge.
      this.mainLedger.chargeWhileReachable('replies', leaseTargets(reply), estimatePayloadBytes(reply))
      this.capture(reply)
      return reply
    }, () => CLOSED)
    if (!signal) return sent.finally(() => unsubscribe?.())
    return new Promise<Settled<K>>(resolve => {
      const abort = () => {
        resolve(CANCELLED)
        if (!this.closed) void this.client.send({ kind: 'cancel', requestId: this.client.nextRequestId(), targetRequestId: head.requestId }).catch(() => {})
      }
      signal.addEventListener('abort', abort, { once: true })
      void sent.then(reply => { signal.removeEventListener('abort', abort); unsubscribe?.(); resolve(this.closed ? CLOSED : reply) })
    })
  }
  nextRequestId(): string { return this.client.nextRequestId() }
  private capture(reply: WorkerReply): void {
    const status = reply.diagnostics?.find(d => d.code === 'cached-value-policy')
    if (status && status.message !== this.statusMessage) { this.statusMessage = status.message; this.notify() }
  }

  /** Worker sort: prepare the key columns (best effort; unprepared keys stream), then commit. */
  async sort(owner: CacheIdentity, spec: SortSpec | null, signal: AbortSignal, progress?: (completed: number, total: number) => void): Promise<RemoteCommit> {
    if (spec?.keys.length) {
      const cols = new Uint16Array([...new Set(spec.keys.map(key => key.col))])
      const prepared = await this.request({ kind: 'prepareColumns', requestId: this.nextRequestId(), owner, cols }, signal)
      if (prepared === CANCELLED) return { kind: 'cancelled' }
      if (prepared === CLOSED) return { kind: 'stale' }
      if (prepared.kind === 'error' && prepared.code === 'stale-source') return { kind: 'stale' }
    }
    return this.commit(await this.request({ kind: 'sort', requestId: this.nextRequestId(), owner, spec }, signal, progress), owner.sheetId, { sort: spec })
  }
  async filter(owner: CacheIdentity, spec: FilterSpec | null, keySets: readonly FilterKeySet[], signal: AbortSignal, progress?: (completed: number, total: number) => void): Promise<RemoteCommit> {
    const command: CommandOf<'filter'> = { kind: 'filter', requestId: this.nextRequestId(), owner, spec, nowMs: Date.now(), ...(keySets.length ? { keySets } : {}) }
    return this.commit(await this.request(command, signal, progress), owner.sheetId, { filter: spec })
  }
  private commit(out: WorkerReply<'sort' | 'filter'> | typeof CLOSED | typeof CANCELLED, sheetId: string, spec: { sort?: SortSpec | null; filter?: FilterSpec | null }): RemoteCommit {
    if (out === CANCELLED) return { kind: 'cancelled' }
    if (out === CLOSED) return { kind: 'stale' }
    switch (out.kind) {
      case 'complete': {
        if (out.result.kind !== 'view') return { kind: 'error', message: 'Unexpected view reply' }
        const view = this.adopt(sheetId, out.result.delta, spec)
        // The adopted view now carries the bytes (charged in adopt): end the reply's lease.
        if (view) this.mainLedger.settle(out)
        return view ? { kind: 'committed', view } : { kind: 'stale' }
      }
      case 'cancelled': return { kind: 'cancelled' }
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
  private adopt(sheetId: string, delta: ViewCommitDelta, spec: { sort?: SortSpec | null; filter?: FilterSpec | null }): RemoteSheetView | null {
    if (this.closed || delta.owner.docId !== this.docId || delta.owner.sheetId !== sheetId) return null
    const previous = this.view(sheetId)
    if (delta.owner.viewGen <= this.identity(sheetId).viewGen) return null
    // Staging: the local band copies coexist with the wire bands (reply), the
    // previous view and the order being built from them.
    const bandBytes = delta.bands.reduce((sum, band) => sum + band.sourceAtOffset.length * 4, 0)
    this.mainLedger.set('viewBuffers', `staging:${sheetId}`, bandBytes)
    const view: RemoteSheetView = {
      revision: delta.owner.viewGen,
      // Wire bands may come from another realm: rebuild them as local arrays.
      order: delta.bands.length ? createRowOrder(delta.bands.map(band => ({ r0: band.r0, r1: band.r1, sourceAtOffset: new Uint32Array(band.sourceAtOffset) }))) : createIdentityOrder(),
      visibility: visibilityFromDisplayBitsets(delta.visibility),
      filter: 'filter' in spec ? spec.filter ?? null : previous.filter,
      sort: 'sort' in spec ? spec.sort ?? null : previous.sort,
    }
    this.identities.set(sheetId, delta.owner)
    if (delta.workbookOwner.globalViewGen >= this.owner.globalViewGen) this.owner = delta.workbookOwner
    this.views.set(sheetId, view)
    // The adopted view (order bands with their display inverse, visibility words
    // and summaries) is charged for its own lifetime: the previous view stays
    // charged until nothing holds it, so a replacement's overlap counts.
    const visibilityBytes = view.visibility ? view.visibility.words.byteLength + view.visibility.summaries.byteLength : 0
    this.mainLedger.chargeWhileReachable('viewBuffers', [view.order, ...(view.visibility ? [view.visibility] : [])], bandBytes * 2 + visibilityBytes)
    this.mainLedger.release('viewBuffers', `staging:${sheetId}`)
    this.notify()
    return view
  }

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
    this.geometries.clear()
    this.layouts.clear()
    this.structures.clear()
    this.failures.clear()
    try {
      await Promise.race([this.client.send({ kind: 'close', requestId: this.client.nextRequestId(), docId: this.docId }), new Promise(resolve => setTimeout(resolve, 1000))])
    } catch { /* the worker is gone */ }
    this.client.dispose()
    await this.onRelease()
  }
}

interface RowEntry { chunks: Set<number>; cells: PagedCell[] }

/** One byte ledger for the painting row caches of every sheet of a document
 * (§12.4). Insertion order is recency; charging past the limit evicts the
 * least recently used rows of any sheet. A row larger than the whole budget is
 * never cached (its reads go to the worker each time). */
export class RowCacheLedger {
  bytes = 0
  peakBytes = 0
  private readonly entries = new Map<string, { store: RemoteSheetStore; row: number; bytes: number }>()
  constructor(readonly limitBytes: number, private readonly main?: MainMemoryLedger) {}
  private report(): void { this.main?.set('rowCache', 'rows', this.bytes) }
  private key(store: RemoteSheetStore, row: number): string { return `${store.sheetId}\u0000${row}` }
  charge(store: RemoteSheetStore, row: number, bytes: number): void {
    const key = this.key(store, row)
    const old = this.entries.get(key)
    if (old) { this.bytes -= old.bytes; this.entries.delete(key) }
    this.entries.set(key, { store, row, bytes })
    this.bytes += bytes
    while (this.bytes > this.limitBytes) {
      const [oldest, entry] = this.entries.entries().next().value as [string, { store: RemoteSheetStore; row: number; bytes: number }]
      this.entries.delete(oldest)
      this.bytes -= entry.bytes
      entry.store.drop(entry.row)
    }
    this.peakBytes = Math.max(this.peakBytes, this.bytes)
    this.report()
  }
  touch(store: RemoteSheetStore, row: number): void {
    const key = this.key(store, row)
    const entry = this.entries.get(key)
    if (entry) { this.entries.delete(key); this.entries.set(key, entry) }
  }
  release(store: RemoteSheetStore, row: number): void {
    const key = this.key(store, row)
    const entry = this.entries.get(key)
    if (entry) { this.entries.delete(key); this.bytes -= entry.bytes; this.report() }
  }
}

/** SourceStore proxy of one remote sheet plus its bounded painting row cache. */
export class RemoteSheetStore implements SourceStore {
  private readonly rows = new Map<number, RowEntry>()
  private readonly loading = new Set<string>()
  // Mutable state lives in one shared holder: find wraps stores with
  // Object.create, and a primitive written through a wrapper would diverge.
  private readonly state: { retry?: ReturnType<typeof setTimeout>; extent: SourceStore['extent']; epoch: number } = { extent: { rows: MAX_ROWS, cols: MAX_COLS }, epoch: 0 }

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
      case 'cancelled': return failure('cancelled', 'Read cancelled', true)
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
  async queryOverview(): Promise<Page<never>> {
    // Coarse frames paint declared neutral summaries on the main thread; the
    // worker overview query is not part of this transport.
    return this.failedPage(failure('unsupported', 'Overview summaries are not served by the paged viewer transport'))
  }
  async ensureComplete(rect: SourceRect, options: ReadOptions): Promise<RangeCompletion> {
    const out = await this.read({ op: 'complete', rect }, options)
    if (!('code' in out) && out.op === 'complete') return out.result
    const error = 'code' in out ? out : failure('page-error', 'Unexpected read reply')
    return { range: rect, completeness: 'incomplete', stamp: this.stamp, missing: [rect], error }
  }
  async release(): Promise<void> {}

  // ---- painting row cache ----

  /** Cached populated cells of a source row (ascending columns); only the
   * known chunks contribute. Unknown chunks are reported by `pendingIn`. */
  cellsOf(row: number): readonly PagedCell[] | undefined {
    if (this.workbook.isClosed) return undefined
    const entry = this.rows.get(row)
    if (entry) this.workbook.ledger.touch(this, row)
    return entry && entry.cells.length ? entry.cells : undefined
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
  drop(row: number): void { this.rows.delete(row) }
  /** Releases the whole cache (inactive sheet, close); in-flight loads land nowhere. */
  clear(): void {
    for (const row of this.rows.keys()) this.workbook.ledger.release(this, row)
    this.rows.clear()
    this.state.epoch++
  }
  /**
   * Logical boxes inside `box` whose detail rows are not yet known, in display
   * order; requests them as a side effect. Detail admission comes first (spec
   * §11.1): rows/columns projecting below four device pixels at `scale` are
   * summarized by the planner and never read. Requests are coalesced into runs
   * and bounded per tile and per frame (in flight); the rest stay pending and
   * are asked for once earlier loads land.
   */
  pendingIn(metrics: ProjectedSheetMetrics, box: LogicalCellBox, scale: number): LogicalCellBox[] {
    const detail = (sizeCss: number) => sizeCss * scale >= MIN_DETAIL_DEVICE_PX - 1e-9
    const rows = metrics.rows.visibleBetween(box.y, box.y + box.height)
    if (rows.count === 0 || rows.count > MAX_PENDING_ROWS) return []
    const cols = metrics.cols.visibleBetween(Math.max(0, box.x - TEXT_OVERFLOW_REACH_CSS), box.x + box.width + TEXT_OVERFLOW_REACH_CSS)
    let c0 = -1, c1 = -1
    for (const run of cols.runs) { if (!detail(run.sizeCss)) continue; if (c0 < 0) c0 = run.start; c1 = run.end }
    if (c0 < 0) return []
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
    if (missing.length) {
      const limits = this.workbook.pageLimits
      const budget = Math.min(limits.tile, limits.frame - this.workbook.loadsInFlight)
      const runs = coalesce(missing).slice(0, Math.max(0, budget))
      for (const run of runs) void this.loadRun(run.r0, run.r1, k0, k1, new AbortController().signal, this.stamp)
    }
    return out
  }
  /** Loads the chunks covering [c0,c1] of `rows` (source rows) in runs. */
  async load(rows: readonly number[], c0: number, c1: number, signal?: AbortSignal, stamp: SourceStamp = this.stamp): Promise<void> {
    const k0 = Math.floor(c0 / ROW_CHUNK_COLS), k1 = Math.floor(c1 / ROW_CHUNK_COLS)
    const runs = coalesce(rows.filter(row => !this.known(row, k0, k1)), 0)
    await Promise.all(runs.map(run => this.loadRun(run.r0, run.r1, k0, k1, signal ?? new AbortController().signal, stamp)))
  }
  private async loadRun(r0: number, r1: number, k0: number, k1: number, signal: AbortSignal, stamp: SourceStamp): Promise<void> {
    const key = `${stamp.valueGen}:${r0}:${r1}:${k0}:${k1}`
    if (this.loading.has(key)) return
    this.loading.add(key)
    this.workbook.loadsInFlight++
    const epoch = this.state.epoch
    const range = { r0, r1, c0: k0 * ROW_CHUNK_COLS, c1: Math.min(MAX_COLS - 1, (k1 + 1) * ROW_CHUNK_COLS - 1) }
    let changed = false
    try {
      let cursor: string | undefined
      // Pages cover the run in ascending rows (§4.1 ordering contract): every
      // row up to a page's last item, or the run end on the final page, is known.
      let covered = r0 - 1
      for (let pages = 0; pages < PAGES_PER_RUN; pages++) {
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
      this.workbook.loadsInFlight--
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
