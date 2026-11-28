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
 * Stage 4c (P7, P8): with `RemoteSheets.sortOps` (`manifest.sheetSort === 1`) the sheet's view is the server's: `sort` and
 * `filter` send the op and adopt the `sheet-order` its commit answers (`applyOrder`: the order, and the mask over the
 * layout's authored-hidden rows); a layout whose `sort`/`filterSummary` is not the identity view is fetched with its
 * order and both are adopted together. Any refusal of a sort or filter (a lost answer, a cancel, `stale-view`,
 * `view-refused`, a policy refusal after a commit) re-fetches the layout and the order, so the view equals the server's.
 * A resize keeps the order and mask. Nothing here sorts, filters or lists values.
 *
 * Widened, never subclassed (R16): this class implements `ViewerSheetSource` only, so a paged-only member of
 * `RemoteWorkbook` cannot be reached on a stream by inheritance.
 */
import type {
  AuthoritativeCell, CacheIdentity, CellRead, ColumnEntry, ColumnQuery, ColumnRule, FilterKeySet, FilterRange, FilterSpec, OperationRefusal, OverviewGroup, OverviewQuery, Page,
  PopulatedQuery, RangeCompletion, ReadOptions, SheetLayout, SortSpec, SourceStamp, SourceStore, SparseRowPage, StructuralResult, ValuePolicy,
} from '../../xlsx/view/contracts'
import type { XlsxDocument, XlsxMergeRange } from '../../xlsx/types'
import type { SheetGeometry } from '../../worker/viewerSession'
import type { RemoteCommit, RemoteSheetView } from '../../worker/viewerRemoteSource'
import type { RemoteOfficeDoc } from '../../stream/remote/remote-doc'
import { StreamSessionError } from '../../stream/remote/errors'
import type {
  SheetFitRequest, SheetOrder, SheetResizeArgs, SheetViewChange, StreamSheetLayout, ValuesPage, WireDetection, WireFilterSpec, WireRange, WireRule, WireSelection, WireSortSpec,
} from '../../stream/sheet-wire'
import type { FindOptions } from '../../xlsx/view/find'
import type { GridSelection, SearchQuery, WorkbookOwner } from '../../xlsx/view/contracts'
import { createAxisMetrics } from '../../xlsx/view/axes'
import { mergeAxisOverrides, NO_SESSION_SIZES, type SessionSizes } from '../../xlsx/view/axisSizes'
import { MAX_COLS, MAX_ROWS, type SourceCell, type SourceRect } from '../../xlsx/view/coordinates'
import { createIdentityOrder, createRowOrder } from '../../xlsx/view/order'
import { CompactRowVisibility, createIntervalSet, visibilityFromHidden } from '../../xlsx/view/visibility'
import { SHEET_DEFAULT_COL_CSS, SHEET_DEFAULT_ROW_CSS } from '../../xlsx/view/viewport'
import type { ViewerSheetSource } from './viewerSource'
import { streamFindSession, wireSelection, type StreamFindSession } from './streamText'

type RemoteSheets = NonNullable<RemoteOfficeDoc['sheets']>
type RemoteSortOps = NonNullable<RemoteSheets['sortOps']>

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

/** A layout whose session view is the identity view (no sort, no filter): it has no order to fetch. */
const identityWire = (l: StreamSheetLayout) => l.sort === null && l.filterSummary.length === 0
const identityView = (revision: number): RemoteSheetView => ({ revision, order: createIdentityOrder(), filter: null, sort: null })
const toWireRange = (r: FilterRange): WireRange => ({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1, header: r.headerRows, totals: r.totalsRows })
const toFilterRange = (r: WireRange): FilterRange => ({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1, headerRows: r.header, totalsRows: r.totals })
const isWireFilter = (spec: FilterSpec | WireFilterSpec): spec is WireFilterSpec => Array.isArray(spec.columns)
const isWireSort = (spec: SortSpec | WireSortSpec): spec is WireSortSpec => 'header' in spec.range
const sameRange = (a: FilterRange, b: WireRange) => a.r0 === b.r0 && a.r1 === b.r1 && a.c0 === b.c0 && a.c1 === b.c1 && a.headerRows === b.header && a.totalsRows === b.totals
/** A local refusal a stream answers without sending. */
const localRefusal = (reason: 'invalid-rule' | 'unsupported', message: string): RemoteCommit => ({ kind: 'refused', refusal: { reason, message } })
/** The local rule of a wire rule, for the view's `filter`: typed rules as sent; a values rule (its selection lives on the
 * server, P15) as an all-values placeholder with its blanks flag; `committed` as the column's rule in `previous`. */
function localRule(rule: WireRule, previous: ColumnRule | undefined): ColumnRule | undefined {
  switch (rule.kind) {
    case 'committed': return previous
    case 'values': return { kind: 'values', selection: { mode: 'queryPatches', base: 'all', patches: [] }, showBlanks: rule.blanks }
    default: return rule
  }
}
/** The wire rule of a local typed or compound rule; undefined for a values rule (it has no wire form without a list). */
const wireRule = (rule: ColumnRule): WireRule | undefined => rule.kind === 'values' ? undefined : rule

interface SheetState {
  wire: StreamSheetLayout
  layout: SheetLayout
  /** Mutated in place, as `RemoteWorkbook.merges` is: `gen` grows when a layout arrives. */
  structure: { gen: number; merges: readonly XlsxMergeRange[]; sourceMerges: readonly XlsxMergeRange[] }
  view: RemoteSheetView
  geometry?: SheetGeometry
  /** `wire`'s session sizes as maps (rebuilt when `wire` changes). */
  sizes?: SessionSizes
  /** Task 9 fix round 1: the committed filter's range from the last order head (null: no filter, or not yet known). */
  filterRange?: WireRange | null
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
        const { wire, order } = await this.fetchView(sheetId, abort.signal)
        if (this.closed || abort.signal.aborted) return
        this.adopt(state, wire, order)
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
  /** The sheet's layout and, when it is not the identity view on a `sortOps` session, the order of that same viewGen
   * (P7). A view that moves between the two answers is fetched again (twice at most), then rejects `stale-view`. */
  private async fetchView(sheetId: string, signal: AbortSignal): Promise<{ wire: StreamSheetLayout; order?: SheetOrder }> {
    for (let attempt = 0; ; attempt++) {
      const wire = await this.remoteSheets.layout(sheetId, signal)
      const sortOps = this.remoteSheets.sortOps
      if (!sortOps || identityWire(wire)) return { wire }
      const order = await sortOps.order(sheetId, signal)
      if (order.head.viewGen === wire.viewGen) return { wire, order }
      if (attempt >= 2 || order.head.viewGen < wire.viewGen) throw new StreamSessionError('stale-view', 'the sheet view moved while it was fetched')
    }
  }
  /** Adopts a layout with its view: the order of the same viewGen, else (an identity layout) the identity view. */
  private adopt(state: SheetState, wire: StreamSheetLayout, order?: SheetOrder): void {
    // A layout asked for before a resize landed may arrive after it: never roll the view back. A first layout that
    // old is replaced by a fresh one.
    if (wire.viewGen < state.wire.viewGen) { if (!state.loaded) void this.refresh(state.wire.sheetId); return }
    state.wire = wire
    state.layout = viewerLayout(wire)
    const merges = wire.merges.map((r): XlsxMergeRange => ({ minRow: r.r0, maxRow: r.r1, minCol: r.c0, maxCol: r.c1 }))
    state.structure.merges = merges
    state.structure.sourceMerges = merges
    state.structure.gen++
    if (!order) state.filterRange = null
    state.view = order ? this.viewOf(state, order) : identityView(wire.viewGen)
    this.viewMoved(state)
    state.loaded = true
    state.failure = undefined
  }
  /** The wire's sizes or viewGen changed: derived values are rebuilt on demand, and the view keeps its order and mask
   * under the new revision (a resize moves no row). */
  private viewMoved(state: SheetState): void {
    state.geometry = undefined
    state.sizes = undefined
    if (state.view.revision !== state.wire.viewGen) state.view = { ...state.view, revision: state.wire.viewGen }
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
      const { wire, order } = await this.fetchView(sheetId, abort.signal)
      if (this.closed || abort.signal.aborted || state.refreshing !== token || wire.viewGen < state.wire.viewGen) return
      this.adopt(state, wire, order)
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
   * A change at or below the held viewGen is ignored: that view already holds it (the server answers a no-op resize
   * with the current viewGen), so applying a change again is idempotent. */
  applyView(v: SheetViewChange): void {
    const state = this.state(v.sheetId)
    if (v.viewGen <= state.wire.viewGen) return
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
  /** The sheet's view: the identity view, or (4c) the server's order and mask under its viewGen. */
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

  // ── Stage 4c: sort, filter, value lists, detection and the order ─────────

  /** The session serves sort, filter, value lists and detection (`manifest.sheetSort === 1`). */
  get sortOps(): boolean { return !!this.remoteSheets.sortOps }
  /** Number and date conditions are served (`manifest.typedFilters`, `'full'` only). */
  get typedFilters(): boolean { return this.remoteSheets.sortOps?.typedFilters === true }
  /** The local view of an order: its bands, and a full-sheet mask whose slots outside the order's mask show as the identity
   * view does (the layout's authored-hidden rows hidden, mapped through the bands), as the server's `visibility` is. */
  private viewOf(state: SheetState, o: SheetOrder, own?: { sort?: SortSpec | null; filter?: FilterSpec | null }): RemoteSheetView {
    const order = o.bands.length ? createRowOrder(o.bands.map(b => ({ r0: b.r0, r1: b.r1, sourceAtOffset: b.sourceAtOffset }))) : createIdentityOrder()
    const base = visibilityFromHidden(createIntervalSet(state.wire.collapsedRows.map(run => ({ r0: run.start, r1: run.end }))))
    const words = base.words.slice()
    const set = (slot: number, on: boolean) => { const w = slot >>> 5, bit = 2 ** (slot & 31); words[w] = on ? (words[w]! | bit) >>> 0 : (words[w]! & ~bit) >>> 0 }
    for (const b of o.bands) for (let i = 0; i < b.sourceAtOffset.length; i++) set(b.r0 + i, base.visible(b.sourceAtOffset[i]!))
    if (o.mask) for (let i = 0; i <= o.mask.r1 - o.mask.r0; i++) set(o.mask.r0 + i, ((o.mask.words[i >>> 5]! >>> (i & 31)) & 1) === 1)
    const previous = state.view, head = o.head
    state.filterRange = head.filterRange ?? (head.filterSummary.length ? state.filterRange ?? null : null)
    const sort = own && 'sort' in own ? own.sort ?? null
      : !head.sort ? null
        : previous.sort && sameRange(previous.sort.range, head.sort.range) && previous.sort.caseSensitive === (head.caseSensitive ?? previous.sort.caseSensitive) && previous.sort.keys.length === head.sort.keys.length && previous.sort.keys.every((k, i) => k.col === head.sort!.keys[i]!.col && k.order === head.sort!.keys[i]!.order) ? previous.sort
          : { range: toFilterRange(head.sort.range), keys: head.sort.keys.map(k => ({ col: k.col, order: k.order })), caseSensitive: head.caseSensitive ?? false }
    // A filter this client did not send is kept only while the summary still names its columns and kinds.
    const kept = previous.filter && previous.filter.columns.size === head.filterSummary.length && head.filterSummary.every(c => previous.filter!.columns.get(c.col)?.kind === c.kind) ? previous.filter : null
    const filter = own && 'filter' in own ? own.filter ?? null : head.filterSummary.length ? kept : null
    return { revision: head.viewGen, order, visibility: new CompactRowVisibility(MAX_ROWS, words), sort, filter }
  }
  /** Adopts an order the server sent (a sort or filter's commit, or `sheet-order`): the view, the layout's viewGen,
   * sort, filter summary and counts move, and the session sizes stay. A `viewGen` at or below the held one is ignored. */
  applyOrder(o: SheetOrder): void { this.takeOrder(o) }
  private takeOrder(o: SheetOrder, own?: { sort?: SortSpec | null; filter?: FilterSpec | null }): void {
    const state = this.state(o.head.sheetId)
    if (o.head.viewGen <= state.wire.viewGen) return
    state.view = this.viewOf(state, o, own)
    state.wire = Object.freeze({ ...state.wire, viewGen: o.head.viewGen, sort: o.head.sort, filterSummary: o.head.filterSummary, counts: { visibleRows: o.head.counts.visible, totalRows: o.head.counts.total } })
    this.viewMoved(state)
    this.notify()
  }
  /** The sheet's committed filter as the server names it (P15's `committed` placeholders after a refetch): its range and
   * summary, or null for no filter (or a range not yet known). */
  filterHead(sheetId: string): { range: WireRange; columns: ReadonlyArray<{ col: number; kind: string }> } | null {
    const state = this.state(sheetId)
    return state.filterRange && state.wire.filterSummary.length ? { range: state.filterRange, columns: state.wire.filterSummary } : null
  }
  /** Re-fetches the layout and, when it is not the identity view, its order (P7); never rejects. */
  refreshOrder(sheetId: string): Promise<void> { return this.refresh(sheetId) }
  private sortOpsOrThrow(): RemoteSortOps {
    if (this.closed) throw new StreamSessionError('closed')
    return this.remoteSheets.sortOps!
  }
  /**
   * P8: progress `0 / 1` as the op is sent and `1 / 1` on its answer, whose order is adopted. Any rejection but `closed`
   * re-fetches the layout and the order first (the commit may stand although its answer was lost, Review Focus 1); then
   * `view-refused` is `refused` with its local refusal, `stale-view` is `stale`, a cancel or an abort is `cancelled`,
   * and any other code rejects with its `StreamSessionError` (a caller reads `stale-filter-values` by code).
   */
  private async commitOp(sheetId: string, progress: ((completed: number, total: number) => void) | undefined, run: () => Promise<SheetOrder>, own: { sort?: SortSpec | null; filter?: FilterSpec | null }): Promise<RemoteCommit> {
    progress?.(0, 1)
    let order: SheetOrder
    try { order = await run() } catch (error) {
      const e = error as { code?: unknown; name?: unknown; refusal?: OperationRefusal }
      if (e?.code === 'closed' || this.closed) throw error
      await this.refresh(sheetId)
      if (e?.code === 'view-refused' && e.refusal) return { kind: 'refused', refusal: e.refusal }
      if (e?.code === 'stale-view') return { kind: 'stale' }
      if (e?.code === 'cancelled' || e?.name === 'AbortError') return { kind: 'cancelled' }
      throw error
    }
    progress?.(1, 1)
    this.takeOrder(order, own)
    return { kind: 'committed', view: this.state(sheetId).view }
  }
  /** A view-only sort at the owner's viewGen (a stale owner is `stale` without sending). `null` clears the sort (the filter
   * stays) by `keys: []` over the held sort's range; with no sort held it changes nothing. */
  async sort(owner: CacheIdentity, spec: SortSpec | WireSortSpec | null, signal: AbortSignal, progress?: (completed: number, total: number) => void): Promise<RemoteCommit> {
    if (!this.remoteSheets.sortOps) return refuse('sort')
    const ops = this.sortOpsOrThrow(), sheetId = owner.sheetId
    if (owner.viewGen !== this.state(sheetId).wire.viewGen) return { kind: 'stale' }
    if (!spec) {
      const held = this.state(sheetId).wire.sort
      if (!held) return { kind: 'committed', view: this.state(sheetId).view }
      const clear: WireSortSpec = { range: held.range, keys: [], caseSensitive: false }
      return this.commitOp(sheetId, progress, () => ops.sort(sheetId, owner.viewGen, clear, signal), { sort: null })
    }
    const wire: WireSortSpec = isWireSort(spec) ? spec : { range: toWireRange(spec.range), keys: spec.keys.map(k => ({ col: k.col, order: k.order })), caseSensitive: spec.caseSensitive }
    const local: SortSpec = isWireSort(spec) ? { range: toFilterRange(spec.range), keys: spec.keys, caseSensitive: spec.caseSensitive } : spec
    return this.commitOp(sheetId, progress, () => ops.sort(sheetId, owner.viewGen, wire, signal), { sort: local })
  }
  /** A view-only filter at the owner's viewGen. `keySets` is ignored: a values rule travels in its wire form (P15),
   * so a local `FilterSpec` may carry typed rules only (a local values rule is refused without sending). `null` clears
   * the filter over the held filter's range; with no filter held it changes nothing. */
  async filter(owner: CacheIdentity, spec: FilterSpec | WireFilterSpec | null, _keySets: readonly FilterKeySet[], signal: AbortSignal, progress?: (completed: number, total: number) => void): Promise<RemoteCommit> {
    if (!this.remoteSheets.sortOps) return refuse('filter')
    const ops = this.sortOpsOrThrow(), sheetId = owner.sheetId, state = this.state(sheetId)
    if (owner.viewGen !== state.wire.viewGen) return { kind: 'stale' }
    let wire: WireFilterSpec, local: FilterSpec | null
    if (!spec) {
      const range = state.view.filter ? toWireRange(state.view.filter.range) : state.filterRange
      if (!range) return state.wire.filterSummary.length ? localRefusal('unsupported', "This filter's range is not known; reopen the sheet and try again.") : { kind: 'committed', view: state.view }
      wire = { range, columns: [] }
      local = null
    } else if (isWireFilter(spec)) {
      wire = spec
      const columns = new Map<number, ColumnRule>()
      for (const [col, rule] of spec.columns) {
        const r = localRule(rule, state.view.filter?.columns.get(col))
        if (r) columns.set(col, r)
      }
      local = spec.columns.length ? { range: toFilterRange(spec.range), columns } : null
    } else {
      const columns: WireFilterSpec['columns'] = []
      for (const [col, rule] of [...spec.columns].sort((a, b) => a[0] - b[0])) {
        const r = wireRule(rule)
        if (!r) return localRefusal('invalid-rule', 'A values filter needs its value list on a streamed sheet.')
        columns.push([col, r])
      }
      wire = { range: toWireRange(spec.range), columns }
      local = columns.length ? spec : null
    }
    return this.commitOp(sheetId, progress, () => ops.filter(sheetId, owner.viewGen, wire, signal), { filter: local })
  }
  /** A value-list page of the sheet at its held viewGen (P6): a first page, or a later one by its cursor. A `stale-view`
   * re-fetches the layout and the order before it rejects; `stale-filter-values` rejects as it is (the caller re-lists). */
  async values(sheetId: string, req: { range: WireRange; col: number; search?: { text: string; wildcards: boolean } } | { cursor: string }, signal?: AbortSignal): Promise<ValuesPage> {
    if (!this.remoteSheets.sortOps) return refuse('values')
    const ops = this.sortOpsOrThrow(), viewGen = this.state(sheetId).wire.viewGen
    const ask = 'cursor' in req ? { sheetId, viewGen, cursor: req.cursor } : { sheetId, viewGen, range: req.range, col: req.col, ...(req.search ? { search: req.search } : {}) }
    return this.staleRefresh(sheetId, () => ops.values(ask, signal))
  }
  /** The server's range detection for a sort or filter from `selection` (P4). */
  async detect(sheetId: string, kind: 'sort' | 'filter', selection: WireSelection, override?: { header: 0 | 1 }, signal?: AbortSignal): Promise<WireDetection> {
    if (!this.remoteSheets.sortOps) return refuse('detect')
    const viewGen = this.state(sheetId).wire.viewGen
    return this.sortOpsOrThrow().detect(sheetId, viewGen, kind, selection, override, signal)
  }
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
