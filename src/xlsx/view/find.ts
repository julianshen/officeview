/**
 * Global workbook find without a full bitmap (S5.3, spec §§10.1, 11.4).
 *
 * Search streams every populated cell of every requested worksheet through
 * its SourceStore (`scanPopulated`), independent of which sheets or tiles were
 * ever painted: no paint replay, no layout record and no bitmap is built to
 * search. A cell matches when it is visible under the sheet's single
 * display-space visibility mask (§5.2: authored hidden ∪ filter) and, when the
 * host supplies one, its column predicate.
 *
 * Match semantics follow the existing logical search (spec §11.4,
 * `core/search.ts`): `String.prototype.toLowerCase` substring matching unless
 * `caseSensitive`, accents significant, an empty query matches nothing. The
 * S8.5 canonical text matcher is deliberately NOT used: §11.4 binds find to the
 * DOCX/PPTX search contract, and the canonical form has no mapping back to raw
 * UTF-16 offsets for highlights. Hits are cell-level (Excel Find All); `start`
 * and `end` name the first occurrence in the cell's searched text, which is
 * the host's formatted display text when `textOf` is supplied and the raw
 * logical text otherwise.
 *
 * Ordering is workbook sheet order, then current display slot, then column;
 * source IDs are the stable identity. The scan visits source rows in order;
 * a permutation band maps source rows [r0,r1] onto display slots [r0,r1], so
 * hits are held only until their band (or row) is complete, then sorted by
 * slot and released. Scanning is lazy: a page scans only until it can fill
 * `pageSize` hits, so retained hits stay bounded by one page plus one band.
 *
 * Ownership: a session captures the aggregate WorkbookOwner (global view
 * generation, workbook value and font generations and the complete ordered
 * sheet revision/value vector) and validates every sheet identity against it.
 * After every await (source page, cooperative yield) it re-checks closure,
 * cancellation and the host's current owner; a mismatch rejects with
 * FindError('stale') so no page from an older owner is ever published.
 *
 * Seams (documented, not wired here): the resident worker has no `search`
 * command handler yet (contracts §10.2 defines it; S4 owns the kernel), and
 * OfficeDoc's large-mode canvas path is S5.2. OfficeDoc keeps its existing
 * paint-replay search for every format until S5.2 mounts the sheet viewport;
 * at that point it should drive FindController with the session's owner and
 * call navigateToHit/hitGeometry for the focused hit only.
 */
import type { XlsxCell } from '../types'
import type {
  AuthoritativeCell, CacheIdentity, Page, RowOrder, RowPermutationBand, RowVisibility, ScrollWindow, SearchHit, SearchQuery,
  SheetMetrics, SourceStamp, SourceStore, WorkbookOwner,
} from './contracts'
import { MAX_COLS, MAX_ROWS, cellId, type SourceCell, type SourceRect } from './coordinates'
import { CooperativeSlice, type CooperativeOptions } from './scheduler'
import { buildCellRecords, rawTextOf, type CellLayout, type CellLayoutHost, type LayoutCache, type Rect } from './textLayout'
import { navigateSelection, type SheetNavigationSession } from './selection'
import { createScrollWindow, navigateTo, type HiddenTargetReason, type ViewportSize } from './scrollWindow'

/** Initial bounded hit page (spec §11.4). */
export const FIND_PAGE_HITS = 200
/** Source cells requested per scanPopulated page; one S4.2 yield budget. */
export const FIND_SCAN_LIMIT = 4096
const FULL_SHEET: SourceRect = Object.freeze({ r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 })

export type FindErrorCode = 'stale' | 'cancelled' | 'closed'
export class FindError extends Error {
  constructor(readonly code: FindErrorCode, message: string) {
    super(message)
    this.name = 'FindError'
  }
}

export interface FindSheetInput {
  sheetId: string
  store: SourceStore
  /** Must agree with the owner's vector entry: viewGen=revision, same valueGen/fontGen. */
  identity: CacheIdentity
  order: RowOrder
  /** DISPLAY-space mask: the single §5.2 predicate (authored hidden ∪ filter). */
  visibility: RowVisibility
  /** Hidden columns, when the host tracks them (zero-width/authored hidden). */
  isColumnVisible?: (col: number) => boolean
}
/** `sheets` in workbook order; the owner is the aggregate state they belong to. */
export interface FindInput { owner: WorkbookOwner; sheets: readonly FindSheetInput[] }
export type FindPhase = 'find-scan'
export interface FindOptions extends CooperativeOptions<FindPhase> {
  pageSize?: number
  scanLimit?: number
  signal?: AbortSignal
  /** Host display formatter; defaults to the raw logical text. */
  textOf?: (cell: AuthoritativeCell, sheetId: string) => string
  /** The host's current aggregate owner, checked after every await. */
  currentOwner?: () => WorkbookOwner
}
export interface FindHit extends SearchHit { start: number; end: number; displaySlot: number }
export interface FindPage {
  owner: WorkbookOwner
  hits: readonly FindHit[]
  /** The whole requested domain has been scanned and every hit emitted. */
  complete: boolean
  /** Only when complete AND every source page was authoritative. */
  totalKnown?: number
  cursor?: string
  unavailableCount: number
  /** Sheets with a loading/failed source page: results are not final (§10.4). */
  incomplete: readonly string[]
}
export interface FindStats {
  scanPages: number
  cellsVisited: number
  textsFormatted: number
  hitsRecorded: number
  maxBufferedHits: number
  bufferedHits: number
  yields: number
}

/** Aggregate (whole-workbook) identity, the same encoding as the worker's
 * workbookIndexKey: "A=1" and "B=1" never share a key (§10.1). JSON array
 * encoding keeps ids containing delimiters injective. */
export function workbookFindKey(owner: WorkbookOwner): string {
  return JSON.stringify([owner.docId, owner.globalViewGen, owner.workbookValueGen, owner.fontGen,
    owner.sheetVector.map(entry => [entry.sheetId, entry.revision, entry.valueGen])])
}
export function findCacheKey(owner: WorkbookOwner, query: SearchQuery): string {
  return JSON.stringify([workbookFindKey(owner), query.text, query.caseSensitive, query.sheetIds ?? null])
}

/** Lowercased text plus, only when lowering changed the length, the raw
 * [start,end) of every lowered code unit. A code point that lowercases to the
 * same number of units maps unit by unit, so surrogate halves keep their own
 * positions; an expansion (İ → i̇) maps each produced unit to the whole
 * originating code point, so a match inside it still covers that character. */
function lowered(text: string): { text: string; from?: Uint32Array; to?: Uint32Array } {
  const lower = text.toLowerCase()
  if (lower.length === text.length) return { text: lower }
  const from = new Uint32Array(lower.length), to = new Uint32Array(lower.length)
  let out = 0
  for (let at = 0; at < text.length;) {
    const width = text.codePointAt(at)! > 0xffff ? 2 : 1
    const piece = text.slice(at, at + width).toLowerCase()
    for (let k = 0; k < piece.length; k++, out++) {
      if (piece.length === width) { from[out] = at + k; to[out] = at + k + 1 }
      else { from[out] = at; to[out] = at + width }
    }
    at += width
  }
  return { text: lower, from, to }
}
function firstMatch(text: string, needle: string, caseSensitive: boolean): { start: number; end: number } | null {
  if (caseSensitive) {
    const at = text.indexOf(needle)
    return at < 0 ? null : { start: at, end: at + needle.length }
  }
  const hay = lowered(text)
  const at = hay.text.indexOf(needle)
  if (at < 0) return null
  return hay.from ? { start: hay.from[at], end: hay.to![at + needle.length - 1] } : { start: at, end: at + needle.length }
}
function bandOf(bands: readonly RowPermutationBand[], row: number): RowPermutationBand | undefined {
  let lo = 0, hi = bands.length
  while (lo < hi) {
    const mid = Math.floor((lo + hi) / 2)
    if (bands[mid].r1 < row) lo = mid + 1
    else hi = mid
  }
  return lo < bands.length && bands[lo].r0 <= row ? bands[lo] : undefined
}
const byDisplay = (a: FindHit, b: FindHit): number => a.displaySlot - b.displaySlot || a.source.col - b.source.col

interface ScanState {
  sheet: FindSheetInput
  stamp: SourceStamp
  iterator?: AsyncIterator<Page<AuthoritativeCell>>
  cursor?: string
  last?: Page<AuthoritativeCell>
  /** Group of the held hits: a permutation band, or a source row outside bands. */
  group?: RowPermutationBand | number
}

class FindSlice extends CooperativeSlice<FindPhase> {
  constructor(options: FindOptions, private readonly guard: () => void) { super(options) }
  override check(): void { this.guard() }
}

/**
 * One find request over one aggregate owner. Pages are pulled with nextPage;
 * a session never retains layouts or text beyond its bounded hit buffers.
 */
export class WorkbookFind {
  readonly stats: FindStats = { scanPages: 0, cellsVisited: 0, textsFormatted: 0, hitsRecorded: 0, maxBufferedHits: 0, bufferedHits: 0, yields: 0 }
  private readonly key: string
  private readonly needle: string
  private readonly pageSize: number
  private readonly scanLimit: number
  private readonly sheets: readonly FindSheetInput[]
  private readonly slice: FindSlice
  private ready: FindHit[] = []
  private held: FindHit[] = []
  private scan: ScanState | undefined
  private nextSheet = 0
  private emitted = 0
  private unavailable = 0
  private readonly incompleteSheets = new Set<string>()
  private validated = false
  private done = false
  private closed = false

  constructor(private readonly input: FindInput, private readonly query: SearchQuery, private readonly options: FindOptions = {}) {
    this.key = workbookFindKey(input.owner)
    this.needle = query.caseSensitive ? query.text : query.text.toLowerCase()
    this.pageSize = options.pageSize ?? FIND_PAGE_HITS
    this.scanLimit = options.scanLimit ?? FIND_SCAN_LIMIT
    if (!Number.isSafeInteger(this.pageSize) || this.pageSize <= 0) throw new RangeError('Invalid find page size')
    if (!Number.isSafeInteger(this.scanLimit) || this.scanLimit <= 0) throw new RangeError('Invalid find scan limit')
    const wanted = query.sheetIds ? new Set(query.sheetIds) : undefined
    this.sheets = input.sheets.filter(sheet => !wanted || wanted.has(sheet.sheetId))
    this.slice = new FindSlice(options, () => this.guard())
  }

  /** Closure, cancellation and owner currency: run after every await. */
  private guard(): void {
    if (this.closed) throw new FindError('closed', 'Find session closed')
    if (this.options.signal?.aborted) throw new FindError('cancelled', 'Find cancelled')
    const current = this.options.currentOwner?.()
    if (current && workbookFindKey(current) !== this.key) throw new FindError('stale', 'Workbook view, value or font owner changed')
  }

  private validate(): void {
    const owner = this.input.owner
    const ids = new Set<string>()
    for (const sheet of this.input.sheets) {
      const entry = owner.sheetVector.find(item => item.sheetId === sheet.sheetId)
      const { identity, store } = sheet
      if (ids.has(sheet.sheetId) || !entry || identity.docId !== owner.docId || identity.sheetId !== sheet.sheetId
        || identity.viewGen !== entry.revision || identity.valueGen !== entry.valueGen || identity.fontGen !== owner.fontGen
        || store.stamp.docId !== owner.docId || store.stamp.sheetId !== sheet.sheetId || store.stamp.valueGen !== entry.valueGen) {
        throw new FindError('stale', `Sheet ${sheet.sheetId} does not belong to the captured workbook owner`)
      }
      ids.add(sheet.sheetId)
    }
  }

  async nextPage(): Promise<FindPage> {
    this.guard()
    if (!this.validated) {
      this.validate()
      this.validated = true
    }
    try {
      if (!this.needle) this.done = true
      while (!this.done && this.ready.length < this.pageSize) await this.scanStep()
    } catch (error) {
      throw this.translate(error)
    }
    this.guard()
    const hits = this.ready.length <= this.pageSize ? this.ready : this.ready.slice(0, this.pageSize)
    this.ready = this.ready.length <= this.pageSize ? [] : this.ready.slice(this.pageSize)
    this.emitted += hits.length
    this.stats.bufferedHits = this.ready.length + this.held.length
    this.stats.yields = this.slice.yields
    const complete = this.done && this.ready.length === 0
    const incomplete = [...this.incompleteSheets]
    return {
      owner: this.input.owner, hits, complete,
      ...(complete && !incomplete.length ? { totalKnown: this.emitted } : {}),
      ...(complete ? {} : { cursor: `${this.nextSheet}:${this.emitted}` }),
      unavailableCount: this.unavailable,
      incomplete,
    }
  }

  close(): void {
    if (this.closed) return
    this.closed = true
    void this.scan?.iterator?.return?.()
    this.scan = undefined
    this.ready = []
    this.held = []
    this.stats.bufferedHits = 0
    this.slice.dispose()
  }

  private translate(error: unknown): unknown {
    if (error instanceof FindError) return error
    if (this.closed) return new FindError('closed', 'Find session closed')
    if (this.options.signal?.aborted) return new FindError('cancelled', 'Find cancelled')
    return error
  }

  /** Pull one source page (or advance to the next sheet) and match it. */
  private async scanStep(): Promise<void> {
    if (!this.scan) {
      if (this.nextSheet >= this.sheets.length) {
        this.done = true
        return
      }
      const sheet = this.sheets[this.nextSheet++]
      this.scan = { sheet, stamp: { docId: sheet.identity.docId, sheetId: sheet.sheetId, valueGen: sheet.identity.valueGen } }
    }
    const scan = this.scan
    if (!scan.iterator) {
      const range = scan.sheet.store.extent.used ?? FULL_SHEET
      const signal = this.options.signal ?? new AbortController().signal
      scan.iterator = scan.sheet.store.scanPopulated({ range, limit: this.scanLimit, ...(scan.cursor !== undefined ? { cursor: scan.cursor } : {}) }, { signal, stamp: scan.stamp })[Symbol.asyncIterator]()
    }
    const next = await scan.iterator.next()
    this.guard()
    if (next.done) {
      // A store may end one iteration with a continuation cursor (§4.1).
      const cursor = scan.last?.nextCursor
      scan.iterator = undefined
      scan.last = undefined
      if (cursor !== undefined && cursor !== scan.cursor) {
        scan.cursor = cursor
        return
      }
      this.release()
      this.scan = undefined
      return
    }
    const page = next.value
    scan.last = page
    this.stats.scanPages++
    if (page.error?.code === 'stale-source' || page.stamp.docId !== scan.stamp.docId || page.stamp.sheetId !== scan.stamp.sheetId || page.stamp.valueGen !== scan.stamp.valueGen) {
      throw new FindError('stale', `Source owner of sheet ${scan.sheet.sheetId} changed`)
    }
    if (page.error?.code === 'cancelled') throw new FindError('cancelled', 'Find cancelled')
    if (page.completeness !== 'complete' || page.error) this.incompleteSheets.add(scan.sheet.sheetId)
    for (const cell of page.items) {
      this.visit(scan, cell)
      if (this.slice.due(1)) await this.slice.pause('find-scan')
    }
  }

  private visit(scan: ScanState, cell: AuthoritativeCell): void {
    this.stats.cellsVisited++
    const { sheet } = scan
    const slot = sheet.order.sourceToDisplay(cell.row)
    if (!sheet.visibility.visible(slot)) return
    if (sheet.isColumnVisible && !sheet.isColumnVisible(cell.col)) return
    if (cell.unavailable) {
      this.unavailable++
      return
    }
    let text: string
    if (this.options.textOf) {
      this.stats.textsFormatted++
      text = this.options.textOf(cell, sheet.sheetId)
    } else text = rawTextOf(cell.value)
    const match = firstMatch(text, this.needle, this.query.caseSensitive)
    if (!match) return
    this.stats.hitsRecorded++
    const group = bandOf(sheet.order.bands, cell.row) ?? cell.row
    if (group !== scan.group) {
      this.release()
      scan.group = group
    }
    this.held.push({ sheetId: sheet.sheetId, cellId: cellId(cell.row, cell.col), source: { row: cell.row, col: cell.col }, text, start: match.start, end: match.end, displaySlot: slot })
    this.stats.maxBufferedHits = Math.max(this.stats.maxBufferedHits, this.ready.length + this.held.length)
  }

  /** A band/row group is complete: its hits move to the ready queue in display order. */
  private release(): void {
    if (!this.held.length) return
    this.held.sort(byDisplay)
    for (const hit of this.held) this.ready.push(hit)
    this.held = []
    if (this.scan) this.scan.group = undefined
  }
}

/** S6.1 navigator shape for a hit: its sheet and SOURCE cell. */
export function matchTarget(hit: FindHit): { sheetId: string; cell: SourceCell } {
  return { sheetId: hit.sheetId, cell: { row: hit.source.row, col: hit.source.col } }
}

export interface HitGeometry { record: CellLayout; rects: Rect[] }
/**
 * Lazy canonical geometry for ONE focused/visible hit: the same S5.1 layout
 * record the painter replays, plus the highlight band of [start,end). Hidden
 * hits have no geometry. Plain runs map offsets proportionally along the
 * measured run (the core/search.ts approximation); laid-out (wrapped/rotated)
 * text, or a hit whose searched text differs from the record's formatted text,
 * highlights the whole text clip.
 */
export function hitGeometry(hit: FindHit, cell: XlsxCell, metrics: SheetMetrics, owner: CacheIdentity, host: CellLayoutHost, cache: LayoutCache): HitGeometry | null {
  const box = metrics.sourceCellToLogical(hit.source)
  if (!box) return null
  const [record] = buildCellRecords([{ cell, source: hit.source, x: box.x, y: box.y, width: box.width, height: box.height }], owner, host, cache)
  const text = record.text[0]
  const body = record.body.text
  if (!text || !body) return { record, rects: [] }
  if (body.kind !== 'plain' || text.formattedText !== hit.text || !body.text.length) return { record, rects: [{ ...text.clip }] }
  const x0 = box.x + (body.tx - record.body.x), unit = body.width / body.text.length
  return { record, rects: [{ x: x0 + unit * hit.start, y: box.y, width: unit * (hit.end - hit.start), height: box.height }] }
}

export interface FindNavigation {
  session: SheetNavigationSession
  /** Per-sheet scroll windows; zoom lives in each window and the retained state. */
  windows: Map<string, ScrollWindow>
  viewport: ViewportSize
}
export type NavigateHitResult =
  | { kind: 'visible'; sheetId: string; window: ScrollWindow; announcement: string }
  | { kind: 'hidden'; sheetId: string; window: ScrollWindow; hiddenReason: HiddenTargetReason; announcement: string }
  /** The sheet's projection was evicted (S4.3): S5.2 rebuilds it, then retries. */
  | { kind: 'needs-projection'; sheetId: string }

/**
 * Navigate to a hit through the active-sheet navigator and its ScrollWindow:
 * activates the hit's sheet, selects the SOURCE cell and scrolls it into view.
 * Zoom is never touched, and other sheets' retained selection/zoom/windows are
 * left as they were. A hit hidden by a later view change opens read-only
 * inspection and moves nothing.
 */
export function navigateToHit(navigation: FindNavigation, hit: FindHit): NavigateHitResult {
  const { session, windows, viewport } = navigation
  const state = session.activate(hit.sheetId)
  const projection = state.projection
  if (!projection) return { kind: 'needs-projection', sheetId: hit.sheetId }
  const request = { addressSpace: 'source' as const, at: { row: hit.source.row, col: hit.source.col } }
  const window = windows.get(hit.sheetId) ?? createScrollWindow(projection, viewport, { zoom: state.zoom })
  const selected = navigateSelection(state.selection, projection, request, state.view)
  if (selected.kind === 'hidden') {
    if (selected.selection !== state.selection) session.setSelection(hit.sheetId, selected.selection)
    return { kind: 'hidden', sheetId: hit.sheetId, window, hiddenReason: selected.hiddenReason!, announcement: selected.announcement }
  }
  const moved = navigateTo(window, { metrics: projection, viewport, view: { ...state.view, selection: selected.selection } }, request, state.selection.active)
  session.setSelection(hit.sheetId, selected.selection)
  windows.set(hit.sheetId, moved.window)
  return { kind: 'visible', sheetId: hit.sheetId, window: moved.window, announcement: selected.announcement }
}

export interface FindPublication {
  key: string
  owner: WorkbookOwner
  query: SearchQuery
  hits: readonly FindHit[]
  complete: boolean
  totalKnown?: number
  unavailableCount: number
  incomplete: readonly string[]
}
export interface FindControllerOptions {
  currentOwner: () => WorkbookOwner
  onPublish?: (state: FindPublication) => void
  /** Completed results kept by aggregate key (LRU). */
  cacheEntries?: number
}

/**
 * Replace/cancel/publish policy for the UI. One run is current: starting
 * another aborts it. Pages publish only while their run is current and the
 * host's owner still equals the run's captured owner; completed results are
 * cached by findCacheKey, so an unchanged owner+query is served without a
 * rescan while any view/value/font change misses.
 */
export class FindController {
  readonly stats = { cacheHits: 0, staleRejected: 0, superseded: 0 }
  published: FindPublication | null = null
  private generation = 0
  private active: AbortController | undefined
  private readonly cache = new Map<string, FindPublication>()
  constructor(private readonly options: FindControllerOptions) {}

  cancel(): void {
    this.generation++
    this.active?.abort()
    this.active = undefined
  }

  async run(input: FindInput, query: SearchQuery, options: FindOptions = {}): Promise<FindPublication | null> {
    this.cancel()
    const token = this.generation
    const key = findCacheKey(input.owner, query)
    if (workbookFindKey(this.options.currentOwner()) !== workbookFindKey(input.owner)) {
      this.stats.staleRejected++
      return null
    }
    const cached = this.cache.get(key)
    if (cached) {
      this.cache.delete(key)
      this.cache.set(key, cached)
      this.stats.cacheHits++
      this.publish(cached)
      return cached
    }
    const abort = new AbortController()
    this.active = abort
    const outer = options.signal
    const onOuter = () => abort.abort()
    outer?.addEventListener('abort', onOuter, { once: true })
    if (outer?.aborted) abort.abort()
    const session = new WorkbookFind(input, query, { ...options, signal: abort.signal, currentOwner: this.options.currentOwner })
    const hits: FindHit[] = []
    try {
      for (;;) {
        const page = await session.nextPage()
        if (token !== this.generation) {
          this.stats.superseded++
          return null
        }
        hits.push(...page.hits)
        const state: FindPublication = {
          key, owner: input.owner, query, hits: hits.slice(), complete: page.complete,
          ...(page.totalKnown !== undefined ? { totalKnown: page.totalKnown } : {}),
          unavailableCount: page.unavailableCount, incomplete: page.incomplete,
        }
        this.publish(state)
        if (page.complete) {
          if (page.totalKnown !== undefined) this.remember(key, state)
          return state
        }
      }
    } catch (error) {
      if (!(error instanceof FindError)) throw error
      if (error.code === 'stale') this.stats.staleRejected++
      else if (token !== this.generation) this.stats.superseded++
      return null
    } finally {
      outer?.removeEventListener('abort', onOuter)
      session.close()
      if (this.active === abort) this.active = undefined
    }
  }

  private publish(state: FindPublication): void {
    this.published = state
    this.options.onPublish?.(state)
  }

  private remember(key: string, state: FindPublication): void {
    this.cache.set(key, state)
    const limit = this.options.cacheEntries ?? 4
    while (this.cache.size > limit) this.cache.delete(this.cache.keys().next().value!)
  }
}
