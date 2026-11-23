/**
 * Stage 4a (S4 §6.1): the viewer source of a tiles-mode stream. It holds the manifest's sheet list, each sheet's
 * streamed `sheet-layout` (§6.4) and the server's `viewGen`, and nothing else: no `ViewerClient`, no row, overflow,
 * CF, drawing or sparkline cache. Cells, drawings and every styled pixel are inside the tiles the server paints.
 *
 * `store(sheetId)` is a `StreamSheetStore`: it carries the `stamp`, `extent`, `completeness` and `valuePolicy` the
 * viewer reads, and every content method rejects the typed `stream-no-cell-data`. That rejection is a guard that makes
 * an unrouted feature visible in tests, not the security boundary (the client has no data to leak).
 *
 * Stage 4b (P11–P13): the sheet's session sizes and `viewGen` live in the held wire layout (`sessionRows`/`sessionCols`),
 * so `streamSheetGeometry` and `geometry()` read one value. `applyView` moves it forward (never back); a rejected
 * resize, or a `stale-view` copy, a11y or fit, re-fetches `sheet-layout` for the authoritative view. The text methods
 * (find, copy, a11y, fit, resize) go to the session's `RemoteSheets.view`; nothing here indexes or serializes text.
 *
 * Widened, never subclassed (R16): this class implements `ViewerSheetSource` only, so a paged-only member of
 * `RemoteWorkbook` cannot be reached on a stream by inheritance.
 */
import type {
  AuthoritativeCell, CacheIdentity, CellRead, ColumnEntry, ColumnQuery, FilterKeySet, FilterSpec, OverviewGroup, OverviewQuery, Page,
  PopulatedQuery, RangeCompletion, ReadOptions, SheetLayout, SortSpec, SourceStamp, SourceStore, SparseRowPage, StructuralResult, ValuePolicy,
} from '../../xlsx/view/contracts'
import type { XlsxDocument, XlsxMergeRange } from '../../xlsx/types'
import type { SheetGeometry } from '../../worker/viewerSession'
import type { RemoteCommit, RemoteSheetView } from '../../worker/viewerRemoteSource'
import type { RemoteOfficeDoc } from '../../stream/remote/remote-doc'
import { StreamSessionError } from '../../stream/remote/errors'
import type { SheetFitRequest, SheetResizeArgs, SheetViewChange, StreamSheetLayout } from '../../stream/sheet-wire'
import type { FindOptions } from '../../xlsx/view/find'
import type { GridSelection, SearchQuery, WorkbookOwner } from '../../xlsx/view/contracts'
import { createAxisMetrics } from '../../xlsx/view/axes'
import { mergeAxisOverrides, NO_SESSION_SIZES, type SessionSizes } from '../../xlsx/view/axisSizes'
import { MAX_COLS, MAX_ROWS, type SourceCell, type SourceRect } from '../../xlsx/view/coordinates'
import { createIdentityOrder } from '../../xlsx/view/order'
import { createIntervalSet } from '../../xlsx/view/visibility'
import { SHEET_DEFAULT_COL_CSS, SHEET_DEFAULT_ROW_CSS } from '../../xlsx/view/viewport'
import type { ViewerSheetSource } from './viewerSource'
import { streamFindSession, wireSelection, type StreamFindSession } from './streamText'

type RemoteSheets = NonNullable<RemoteOfficeDoc['sheets']>

/** A content read on a stream source: the client holds no cell data (S4 §6.1). */
export class StreamNoCellData extends Error {
  readonly code = 'stream-no-cell-data'
  readonly retryable = false
  constructor(readonly method: string) {
    super(`stream-no-cell-data: ${method} has no cell data on a streamed sheet`)
    this.name = 'StreamNoCellData'
  }
}

const refuse = (method: string): Promise<never> => Promise.reject(new StreamNoCellData(method))

/** A stream sheet's store: identity and extent only; every content method rejects `stream-no-cell-data`. */
export class StreamSheetStore implements SourceStore {
  readonly completeness = 'complete' as const
  constructor(private readonly workbook: StreamWorkbook, readonly sheetId: string) {}
  get stamp(): SourceStamp { return { docId: this.workbook.docId, sheetId: this.sheetId, valueGen: 0 } }
  get structureGen(): number { return this.workbook.merges(this.sheetId).gen }
  get valuePolicy(): ValuePolicy { return this.workbook.valuePolicy }
  get extent(): SourceStore['extent'] {
    const used = this.workbook.layout(this.sheetId).used
    return { rows: MAX_ROWS, cols: MAX_COLS, ...(used ? { used } : {}) }
  }
  readCell(_at: SourceCell, _options: ReadOptions): Promise<CellRead> { return refuse('readCell') }
  readRows(_range: SourceRect, _cursor: string | undefined, _options: ReadOptions): Promise<Page<SparseRowPage>> { return refuse('readRows') }
  queryColumn(_request: ColumnQuery, _options: ReadOptions): Promise<Page<ColumnEntry>> { return refuse('queryColumn') }
  // eslint-disable-next-line require-yield
  async *scanPopulated(_request: PopulatedQuery, _options: ReadOptions): AsyncIterable<Page<AuthoritativeCell>> { throw new StreamNoCellData('scanPopulated') }
  // eslint-disable-next-line require-yield
  async *scanOccupied(_request: PopulatedQuery, _options: ReadOptions): AsyncIterable<Page<SourceCell>> { throw new StreamNoCellData('scanOccupied') }
  queryStructural(_rect: SourceRect, _options: ReadOptions): Promise<StructuralResult> { return refuse('queryStructural') }
  queryOverview(_request: OverviewQuery, _options: ReadOptions): Promise<Page<OverviewGroup>> { return refuse('queryOverview') }
  ensureComplete(_rect: SourceRect, _options: ReadOptions): Promise<RangeCompletion> { return refuse('ensureComplete') }
  async release(): Promise<void> {}
}

/** The layout a sheet shows before its `sheet-layout` arrives: default sizes, nothing hidden, nothing merged. */
function emptyWireLayout(sheetId: string): StreamSheetLayout {
  return Object.freeze({
    sheetId, viewGen: 0, defaults: { rowPx: SHEET_DEFAULT_ROW_CSS, colPx: SHEET_DEFAULT_COL_CSS }, rows: [], cols: [], collapsedRows: [], collapsedCols: [],
    merges: [], filters: [], sort: null, filterSummary: [], authorHidden: false, counts: { visibleRows: 0, totalRows: 0 },
  })
}
/** The viewer's `SheetLayout` of a wire layout: collapsed runs become the authored-hidden rows and columns, and every
 * paged-content flag is false (sparklines, CF, tables and pivots are painted inside the server's tiles). */
function viewerLayout(l: StreamSheetLayout): SheetLayout {
  return Object.freeze({
    rowOverrides: l.rows, colOverrides: l.cols,
    hidden: l.collapsedRows.map(run => ({ r0: run.start, r1: run.end })),
    ...(l.collapsedCols.length ? { hiddenCols: l.collapsedCols.map(run => ({ start: run.start, end: run.end })) } : {}),
    ...(l.used ? { used: l.used } : {}),
    ...(l.frozen ? { frozen: l.frozen } : {}),
    hasSparklines: false, hasConditionalFormatting: false, hasRegionStyles: false, regionStyleTableCount: 0, hasPivotDisplay: false, pivotTableCount: 0,
  })
}

interface SheetState {
  wire: StreamSheetLayout
  layout: SheetLayout
  /** Mutated in place, as `RemoteWorkbook.merges` is: `gen` grows when a layout arrives. */
  structure: { gen: number; merges: readonly XlsxMergeRange[]; sourceMerges: readonly XlsxMergeRange[] }
  view: RemoteSheetView
  geometry?: SheetGeometry
  /** `wire`'s session sizes as maps (rebuilt when `wire` changes). */
  sizes?: SessionSizes
  /** The latest `refresh` of this sheet (an older one's answer is dropped). */
  refreshing?: object
  loaded: boolean
  failure?: { code: string; message: string }
  inFlight?: Promise<void>
}

export class StreamWorkbook implements ViewerSheetSource {
  readonly kind = 'stream' as const
  /** Tiles paint the saved values; nothing recomputes on a stream (S4 §6.2). */
  readonly valuePolicy: ValuePolicy = 'cached'
  readonly sheets: readonly { sheetId: string; name: string }[]
  private readonly states = new Map<string, SheetState>()
  private readonly stores = new Map<string, StreamSheetStore>()
  private readonly listeners = new Set<() => void>()
  private readonly aborts = new Set<AbortController>()
  private closed = false

  constructor(private readonly remoteSheets: RemoteSheets, readonly docId: string) {
    this.sheets = remoteSheets.list.map(sheet => ({ sheetId: sheet.sheetId, name: sheet.name }))
  }

  private state(sheetId: string): SheetState {
    let state = this.states.get(sheetId)
    if (!state) {
      if (!this.sheets.some(sheet => sheet.sheetId === sheetId)) throw new RangeError(`Unknown sheet ${sheetId}`)
      const wire = emptyWireLayout(sheetId)
      state = { wire, layout: viewerLayout(wire), structure: { gen: 0, merges: [], sourceMerges: [] }, view: { revision: 0, order: createIdentityOrder(), filter: null, sort: null }, loaded: false }
      this.states.set(sheetId, state)
    }
    return state
  }

  /** Fetches the sheet's `sheet-layout` (one request per activation; a second activation while one is in flight joins
   * it) and holds it. Never rejects: a failure is reported through `loadState`, and listeners are notified. */
  activate(sheetId: string): Promise<void> {
    const state = this.state(sheetId)
    if (this.closed) return Promise.resolve()
    if (state.inFlight) return state.inFlight
    const abort = new AbortController()
    this.aborts.add(abort)
    const run = (async () => {
      try {
        const wire = await this.remoteSheets.layout(sheetId, abort.signal)
        if (this.closed || abort.signal.aborted) return
        this.adopt(state, wire)
      } catch (error) {
        if (this.closed || abort.signal.aborted) return
        const code = typeof (error as { code?: unknown })?.code === 'string' ? (error as { code: string }).code : 'layout-failed'
        state.failure = { code, message: error instanceof Error ? error.message : String(error) }
      } finally {
        this.aborts.delete(abort)
        state.inFlight = undefined
      }
      this.notify()
    })()
    state.inFlight = run
    return run
  }
  private adopt(state: SheetState, wire: StreamSheetLayout): void {
    // A layout asked for before a resize landed may arrive after it: never roll the view back. A first layout that
    // old is replaced by a fresh one.
    if (wire.viewGen < state.wire.viewGen) { if (!state.loaded) void this.refresh(state.wire.sheetId); return }
    state.wire = wire
    state.layout = viewerLayout(wire)
    const merges = wire.merges.map((r): XlsxMergeRange => ({ minRow: r.r0, maxRow: r.r1, minCol: r.c0, maxCol: r.c1 }))
    state.structure.merges = merges
    state.structure.sourceMerges = merges
    state.structure.gen++
    this.viewMoved(state)
    state.loaded = true
    state.failure = undefined
  }
  /** The wire's sizes or viewGen changed: derived values are rebuilt on demand. */
  private viewMoved(state: SheetState): void {
    state.geometry = undefined
    state.sizes = undefined
    if (state.view.revision !== state.wire.viewGen) state.view = { revision: state.wire.viewGen, order: createIdentityOrder(), filter: null, sort: null }
  }
  /** A fresh `sheet-layout` (never joined to an activation in flight), adopted unless a newer refresh or view
   * superseded it. Never rejects; a failure leaves the sheet as it was. */
  private async refresh(sheetId: string): Promise<void> {
    const state = this.state(sheetId)
    if (this.closed) return
    const token = {}
    state.refreshing = token
    const abort = new AbortController()
    this.aborts.add(abort)
    try {
      const wire = await this.remoteSheets.layout(sheetId, abort.signal)
      if (this.closed || abort.signal.aborted || state.refreshing !== token || wire.viewGen < state.wire.viewGen) return
      this.adopt(state, wire)
      this.notify()
    } catch { /* the sheet keeps the view it has */ } finally {
      this.aborts.delete(abort)
      if (state.refreshing === token) state.refreshing = undefined
    }
  }

  // ── Stage 4b: session sizes, viewGen and the text ops ────────────────────

  /** The session's sizes of the sheet (stable between changes). */
  sizes(sheetId: string): SessionSizes {
    const state = this.state(sheetId)
    const { sessionRows, sessionCols } = state.wire
    if (!sessionRows?.length && !sessionCols?.length) return NO_SESSION_SIZES
    return state.sizes ??= Object.freeze({ rows: new Map(sessionRows ?? []), cols: new Map(sessionCols ?? []) })
  }
  /** Applies the server's `sheet-view` answer: the sheet's viewGen and session sizes move, and listeners are notified.
   * A change older than the held view is ignored (idempotent for the one already applied). */
  applyView(v: SheetViewChange): void {
    const state = this.state(v.sheetId)
    if (v.viewGen < state.wire.viewGen) return
    const c = v.change
    let { sessionRows, sessionCols } = state.wire
    if ('reset' in c) { sessionRows = undefined; sessionCols = undefined }
    else {
      const list = (c.axis === 'row' ? sessionRows : sessionCols) ?? []
      const next = list.filter(([i]) => i !== c.index)
      next.push([c.index, c.size])
      next.sort((a, b) => a[0] - b[0])
      if (c.axis === 'row') sessionRows = next
      else sessionCols = next
    }
    const { sessionRows: _r, sessionCols: _c, ...rest } = state.wire
    state.wire = Object.freeze({ ...rest, viewGen: v.viewGen, ...(sessionRows?.length ? { sessionRows } : {}), ...(sessionCols?.length ? { sessionCols } : {}) })
    this.viewMoved(state)
    this.notify()
  }
  /** The session serves find, copy, a11y, fit and resize (`manifest.sheetView === 1`). */
  get textOps(): boolean { return !!this.remoteSheets.view }
  private ops(): NonNullable<RemoteSheets['view']> {
    if (this.closed) throw new StreamSessionError('closed')
    const view = this.remoteSheets.view
    if (!view) throw new StreamSessionError('unsupported-op', 'this stream does not serve sheet text')
    return view
  }
  /** F8: a `stale-view` refusal re-fetches the layout (its viewGen and sizes) before the call rejects; never retried. */
  private async staleRefresh<T>(sheetId: string, run: () => Promise<T>): Promise<T> {
    try { return await run() } catch (error) {
      if ((error as { code?: unknown })?.code === 'stale-view') await this.refresh(sheetId)
      throw error
    }
  }
  /** A workbook find over the server's answer (no local index); pages of `FIND_PAGE_HITS` hits without text. */
  findSession(owner: WorkbookOwner, query: SearchQuery, options: FindOptions): StreamFindSession {
    return streamFindSession(this.ops(), owner, this.sheets.map(sheet => sheet.sheetId), query, options.signal)
  }
  /** The selection's display text as TSV. A selection over the wire caps rejects `copy-too-large` with no `limit`,
   * without sending ('Selection too large to copy.'). */
  async copy(sheetId: string, selection: GridSelection): Promise<string> {
    const ops = this.ops()
    const sel = wireSelection(selection)
    if (sel === 'too-large') throw new StreamSessionError('copy-too-large', 'Selection too large to copy.')
    return this.staleRefresh(sheetId, () => ops.copy(sheetId, this.state(sheetId).wire.viewGen, sel))
  }
  /** A11y text of a window, asked under `viewGen` (the caller files the answer under it; F8). */
  async a11y(sheetId: string, viewGen: number, rows: Array<[number, number]>, cols: [number, number]): Promise<Array<{ row: number; col: number; text: string }>> {
    const ops = this.ops()
    return this.staleRefresh(sheetId, () => ops.a11y(sheetId, viewGen, rows, cols))
  }
  async fit(req: Omit<SheetFitRequest, 'viewGen'>): Promise<number> {
    const ops = this.ops()
    return this.staleRefresh(req.sheetId, () => ops.fit({ ...req, viewGen: this.state(req.sheetId).wire.viewGen }))
  }
  /** Resizes at the current viewGen and applies the answer. Any rejection but `closed` or an abort re-fetches the
   * layout first (P13: the server may have applied a resize whose answer was lost); it is never retried. */
  async resize(args: SheetResizeArgs): Promise<SheetViewChange> {
    const ops = this.ops()
    let change: SheetViewChange
    try {
      change = await ops.resize({ ...args, viewGen: this.state(args.sheetId).wire.viewGen })
    } catch (error) {
      const code = (error as { code?: unknown; name?: unknown })?.code
      if (code !== 'closed' && (error as { name?: unknown })?.name !== 'AbortError') await this.refresh(args.sheetId)
      throw error
    }
    this.applyView(change)
    return change
  }

  /** The sheet's layout as received (or the empty default before it arrives), for `streamSheetGeometry`. */
  wireLayout(sheetId: string): StreamSheetLayout { return this.state(sheetId).wire }
  layout(sheetId: string): SheetLayout { return this.state(sheetId).layout }
  merges(sheetId: string): { gen: number; merges: readonly XlsxMergeRange[]; sourceMerges: readonly XlsxMergeRange[] } { return this.state(sheetId).structure }
  identity(sheetId: string): CacheIdentity { return { docId: this.docId, sheetId, valueGen: 0, fontGen: 0, viewGen: this.state(sheetId).wire.viewGen } }
  /** The identity view (sort and filter are 4c). */
  view(sheetId: string): RemoteSheetView { return this.state(sheetId).view }
  store(sheetId: string): StreamSheetStore {
    let store = this.stores.get(sheetId)
    if (!store) {
      this.state(sheetId)
      this.stores.set(sheetId, store = new StreamSheetStore(this, sheetId))
    }
    return store
  }
  /** The axes `streamSheetGeometry` builds (the server's), and the authored-hidden rows. */
  geometry(sheetId: string): SheetGeometry {
    const state = this.state(sheetId)
    if (!state.geometry) {
      const l = state.wire, layout = state.layout
      // The session sizes merge over the authored ones exactly as `streamSheetGeometry` merges them.
      const sized = (base: StreamSheetLayout['rows'], session: Array<[number, number]> | undefined) => session?.length ? mergeAxisOverrides(base, new Map(session)) : base
      state.geometry = Object.freeze({
        rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: l.defaults.rowPx, overrides: sized(l.rows, l.sessionRows), ...(l.collapsedRows.length ? { collapsed: l.collapsedRows } : {}) }),
        cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: l.defaults.colPx, overrides: sized(l.cols, l.sessionCols), ...(l.collapsedCols.length ? { collapsed: l.collapsedCols } : {}) }),
        rowOverrides: layout.rowOverrides, colOverrides: layout.colOverrides, baseHidden: createIntervalSet(layout.hidden),
        ...(layout.hiddenCols?.length ? { hiddenCols: layout.hiddenCols } : {}),
      })
    }
    return state.geometry
  }
  loadState(sheetId: string): { completeness: 'complete' | 'loading' | 'incomplete'; failure?: string; code?: string } {
    const state = this.state(sheetId)
    if (state.failure) return { completeness: 'incomplete', failure: state.failure.message, code: state.failure.code }
    return { completeness: state.loaded ? 'complete' : 'loading' }
  }
  hasSparklines(_sheetId: string): boolean { return false }
  hasConditionalFormatting(_sheetId: string): boolean { return false }
  hasRegionStyles(_sheetId: string): boolean { return false }
  regionStyleTableCount(_sheetId: string): number { return 0 }
  hasPivotDisplay(_sheetId: string): boolean { return false }
  pivotTableCount(_sheetId: string): number { return 0 }
  /** 4c brings `sheet-sort` / `sheet-filter`; until then a stream view cannot be sorted or filtered. */
  sort(_owner: CacheIdentity, _spec: SortSpec | null, _signal: AbortSignal, _progress?: (completed: number, total: number) => void): Promise<RemoteCommit> { return refuse('sort') }
  filter(_owner: CacheIdentity, _spec: FilterSpec | null, _keySets: readonly FilterKeySet[], _signal: AbortSignal, _progress?: (completed: number, total: number) => void): Promise<RemoteCommit> { return refuse('filter') }
  /** The session's tile protocol (Task 13 batches it); rejects `closed` once this source is closed. */
  tiles(): RemoteSheets['tiles'] {
    return (req, onTile, signal) => this.closed ? Promise.reject(Object.assign(new Error('closed'), { code: 'closed' })) : this.remoteSheets.tiles(req, onTile, signal)
  }

  get isClosed(): boolean { return this.closed }
  subscribe(listener: () => void): () => void {
    this.listeners.add(listener)
    return () => { this.listeners.delete(listener) }
  }
  private notify(): void { for (const listener of [...this.listeners]) listener() }
  /** The viewer releases its source on unmount; the stream's owner closes it (as `RemoteWorkbook.release`). */
  async release(): Promise<void> {}
  /** Ends every layout fetch; later activations do nothing. */
  close(): void {
    if (this.closed) return
    this.closed = true
    for (const abort of this.aborts) abort.abort()
    this.aborts.clear()
    this.notify()
  }
}

/** The shell `XlsxDocument` of a stream: sheet names (and wire ids) only; no rows, cells, styles or drawings. */
export function streamShellDoc(list: ReadonlyArray<{ sheetId: string; name: string }>): XlsxDocument {
  return { sheets: list.map((sheet, index) => ({ name: sheet.name, sheetId: sheet.sheetId, workbookIndex: index, rows: [], cols: [], merges: [], mergeRanges: [] })) }
}
