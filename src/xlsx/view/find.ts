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
 * hits are held only until their band (or row) is complete, then ordered
 * cooperatively (sorted runs plus budgeted merge passes) and released. Waiting
 * hits are compact display keys (8-12 bytes each) under a hard budget that
 * refuses with FindError('resource-limit') before it is exceeded; a hit object
 * (with its text, re-read through readCell) is built only when its page is
 * emitted. Scanning is lazy: a page scans only until it can fill `pageSize`.
 *
 * Ownership: a session captures the aggregate WorkbookOwner (global view
 * generation, workbook value and font generations and the complete ordered
 * sheet revision/value vector) and validates every sheet identity against it.
 * After every await (source page, cooperative yield) it re-checks closure,
 * cancellation and the host's current owner; a mismatch rejects with
 * FindError('stale') so no page from an older owner is ever published.
 *
 * The controller can use the same bounded page interface locally or through
 * the resident worker. Materialize hit geometry only for the focused hit.
 */
import type { XlsxCell } from '../types'
import type {
  AuthoritativeCell, CacheIdentity, Page, RowOrder, RowPermutationBand, RowVisibility, ScrollWindow, SearchHit, SearchQuery,
  SheetMetrics, SourceErrorCode, SourceStamp, SourceStore, WorkbookOwner,
} from './contracts'
import { MAX_COLS, MAX_ROWS, cellId, type SourceCell, type SourceRect } from './coordinates'
import { CooperativeSlice, createHostYield, type CooperativeOptions, type HostYield } from './scheduler'
import { buildCellRecords, rawTextOf, type CellLayout, type CellLayoutHost, type LayoutCache, type Rect } from './textLayout'
import { navigateSelection, type SheetNavigationSession } from './selection'
import { createScrollWindow, navigateTo, type HiddenTargetReason, type ViewportSize } from './scrollWindow'

/** Initial bounded hit page (spec §11.4). */
export const FIND_PAGE_HITS = 200
/** Source cells requested per scanPopulated page; one S4.2 yield budget. */
export const FIND_SCAN_LIMIT = 4096
/** Hits waiting for emission (held band + ready queue), 12 bytes each: 3 MiB. */
export const FIND_HIT_BUFFER_BUDGET = 262_144
const FULL_SHEET: SourceRect = Object.freeze({ r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 })

/** resource-limit: the bounded hit buffer would be exceeded (§11.4 bounded storage). */
export type FindErrorCode = 'stale' | 'cancelled' | 'closed' | 'resource-limit'
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
export type FindPhase = 'find-scan' | 'find-order' | 'find-page'
export interface FindOptions extends CooperativeOptions<FindPhase> {
  pageSize?: number
  scanLimit?: number
  /** Waiting-hit budget; exceeding it rejects with FindError('resource-limit'). */
  hitBufferBudget?: number
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
  /** Typed source failures behind `incomplete`, for error/retry presentation. */
  failures: readonly FindSourceFailure[]
}
/** A source failure kept with its query: hits found before/around it stay inspectable. */
export interface FindSourceFailure { sheetId: string; code: SourceErrorCode; retryable: boolean; message: string; range?: SourceRect }
export interface FindStats {
  scanPages: number
  cellsVisited: number
  textsFormatted: number
  hitsRecorded: number
  maxBufferedHits: number
  bufferedHits: number
  /** Retained bytes of waiting hits (compact keys only, never hit objects). */
  bufferedBytes: number
  /** Hit objects built: only for emitted pages. */
  hitsMaterialized: number
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
const ORDER_RUN = 4096

/** Ready hits as parallel compact arrays: sheet index + display key
 * (displaySlot * MAX_COLS + col, exact below 2^34). */
class HitQueue {
  private sheets = new Uint32Array(64)
  private keys = new Float64Array(64)
  private head = 0
  private tail = 0
  get size(): number { return this.tail - this.head }
  push(sheet: number, keys: Float64Array, count: number): void {
    if (this.tail + count > this.keys.length) {
      const live = this.size, capacity = Math.max(64, 2 ** Math.ceil(Math.log2(live + count)))
      const sheets = new Uint32Array(capacity), next = new Float64Array(capacity)
      sheets.set(this.sheets.subarray(this.head, this.tail))
      next.set(this.keys.subarray(this.head, this.tail))
      this.sheets = sheets
      this.keys = next
      this.head = 0
      this.tail = live
    }
    this.sheets.fill(sheet, this.tail, this.tail + count)
    this.keys.set(keys.subarray(0, count), this.tail)
    this.tail += count
  }
  shift(): { sheet: number; key: number } {
    const entry = { sheet: this.sheets[this.head], key: this.keys[this.head] }
    this.head++
    if (this.head === this.tail) this.head = this.tail = 0
    return entry
  }
  clear(): void {
    this.sheets = new Uint32Array(64)
    this.keys = new Float64Array(64)
    this.head = this.tail = 0
  }
}

interface ScanState {
  sheet: FindSheetInput
  stamp: SourceStamp
  range?: SourceRect
  iterator?: AsyncIterator<Page<AuthoritativeCell>>
  cursor?: string
  last?: Page<AuthoritativeCell>
  /** Group of the held hits: a permutation band, or a source row outside bands. */
  group?: RowPermutationBand | number
}

class FindSlice extends CooperativeSlice<FindPhase> {
  constructor(options: CooperativeOptions<FindPhase>, private readonly guard: () => void) { super(options) }
  override check(): void { this.guard() }
}

/**
 * One find request over one aggregate owner. Pages are pulled with nextPage;
 * a session never retains layouts or text beyond its bounded hit buffers.
 */
export class WorkbookFind {
  readonly stats: FindStats = { scanPages: 0, cellsVisited: 0, textsFormatted: 0, hitsRecorded: 0, maxBufferedHits: 0, bufferedHits: 0, bufferedBytes: 0, hitsMaterialized: 0, yields: 0 }
  private readonly key: string
  private readonly needle: string
  private readonly pageSize: number
  private readonly scanLimit: number
  /** Requested, supplied sheets in owner (workbook) order; set by validate(). */
  private sheets: readonly FindSheetInput[] = []
  private readonly slice: FindSlice
  private readonly budget: number
  private readonly ready = new HitQueue()
  /** Display keys of the current band/row group, unsorted until it completes. */
  private held = new Float64Array(64)
  private heldCount = 0
  private scan: ScanState | undefined
  private nextSheet = 0
  private emitted = 0
  private unavailable = 0
  private readonly incompleteSheets = new Set<string>()
  private readonly failures: FindSourceFailure[] = []
  private readonly failureKeys = new Set<string>()
  private host: HostYield | undefined
  /** Settles the pending yield early: a disposed port (MessageChannel) never
   * delivers, so close/abort must wake the waiter for its guard to reject. */
  private wake: (() => void) | undefined
  /** Aborted on close() or caller abort: every source read uses this signal. */
  private readonly lifetime = new AbortController()
  private readonly onAbort = () => { this.wake?.(); this.lifetime.abort() }
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
    this.budget = options.hitBufferBudget ?? FIND_HIT_BUFFER_BUDGET
    if (!Number.isSafeInteger(this.budget) || this.budget <= 0) throw new RangeError('Invalid find hit buffer budget')
    this.host = options.yield ? undefined : createHostYield()
    const inner = options.yield ?? this.host!.yield
    const wrapped = (phase: FindPhase) => new Promise<void>((resolve, reject) => {
      this.wake = resolve
      inner(phase).then(resolve, reject)
    }).finally(() => { this.wake = undefined })
    this.slice = new FindSlice({ ...(options.awaitYield ? { awaitYield: options.awaitYield } : {}), ...(options.now ? { now: options.now } : {}), yield: wrapped }, () => this.guard())
    options.signal?.addEventListener('abort', this.onAbort, { once: true })
    if (options.signal?.aborted) this.lifetime.abort()
  }

  /** Settles when the session is closed/aborted even if the source ignores
   * its signal; the following guard() then rejects with closed/cancelled. */
  private interruptible<T>(work: Promise<T>): Promise<T | undefined> {
    const signal = this.lifetime.signal
    if (signal.aborted) return Promise.resolve(undefined)
    return new Promise<T | undefined>((resolve, reject) => {
      const stop = () => resolve(undefined)
      signal.addEventListener('abort', stop, { once: true })
      work.then(value => { signal.removeEventListener('abort', stop); resolve(value) },
        error => { signal.removeEventListener('abort', stop); reject(error) })
    })
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
    // The domain is every requested sheet of the owner, in workbook order. A
    // requested sheet the caller did not supply is explicitly incomplete, so a
    // partial input can never pass for (or be cached as) the whole workbook.
    const vector = owner.sheetVector.map(entry => entry.sheetId)
    for (const id of this.query.sheetIds ?? []) {
      if (!vector.includes(id)) throw new FindError('stale', `Sheet ${id} is not part of the captured workbook owner`)
    }
    const wanted = this.query.sheetIds ? new Set(this.query.sheetIds) : undefined
    const sheets: FindSheetInput[] = []
    for (const id of vector) {
      if (wanted && !wanted.has(id)) continue
      const sheet = this.input.sheets.find(item => item.sheetId === id)
      if (sheet) sheets.push(sheet)
      else this.fail({ sheetId: id, code: 'incomplete', retryable: true, message: `Sheet ${id} was not supplied to find` })
    }
    this.sheets = sheets
  }

  async nextPage(): Promise<FindPage> {
    this.guard()
    if (!this.validated) {
      this.validate()
      this.validated = true
    }
    const hits: FindHit[] = []
    try {
      if (!this.needle) this.done = true
      while (hits.length < this.pageSize) {
        while (!this.done && this.ready.size < this.pageSize - hits.length) await this.scanStep()
        if (!this.ready.size) break
        const hit = await this.materialize()
        if (hit) hits.push(hit)
      }
    } catch (error) {
      throw this.translate(error)
    }
    this.guard()
    this.emitted += hits.length
    this.account()
    this.stats.yields = this.slice.yields
    const complete = this.done && this.ready.size === 0
    // A finished scan never yields again; its host ports go now, not at close.
    if (this.done) this.releaseHost()
    const incomplete = [...this.incompleteSheets]
    return {
      owner: this.input.owner, hits, complete,
      ...(complete && !incomplete.length ? { totalKnown: this.emitted } : {}),
      ...(complete ? {} : { cursor: `${this.nextSheet}:${this.emitted}` }),
      unavailableCount: this.unavailable,
      incomplete,
      failures: this.failures.slice(),
    }
  }

  close(): void {
    if (this.closed) return
    this.closed = true
    void this.scan?.iterator?.return?.()
    this.scan = undefined
    this.ready.clear()
    this.held = new Float64Array(0)
    this.heldCount = 0
    this.account()
    this.options.signal?.removeEventListener('abort', this.onAbort)
    this.wake?.()
    this.lifetime.abort()
    this.releaseHost()
  }

  private releaseHost(): void {
    this.host?.dispose()
    this.host = undefined
  }

  private fail(failure: FindSourceFailure): void {
    this.incompleteSheets.add(failure.sheetId)
    const key = JSON.stringify([failure.sheetId, failure.code, failure.message])
    if (this.failureKeys.has(key)) return
    this.failureKeys.add(key)
    this.failures.push(failure)
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
      const range = scan.range ??= scan.sheet.store.extent.used ?? FULL_SHEET
      const signal = this.lifetime.signal
      scan.iterator = scan.sheet.store.scanPopulated({ range, limit: this.scanLimit, ...(scan.cursor !== undefined ? { cursor: scan.cursor } : {}) }, { signal, stamp: scan.stamp })[Symbol.asyncIterator]()
    }
    const next = await this.interruptible(scan.iterator.next())
    this.guard()
    if (!next) throw new FindError('closed', 'Find session closed')
    if (next.done) {
      // A store may end one iteration with a continuation cursor (§4.1).
      const cursor = scan.last?.nextCursor
      scan.iterator = undefined
      scan.last = undefined
      if (cursor !== undefined && cursor !== scan.cursor) {
        scan.cursor = cursor
        return
      }
      await this.release()
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
    if (page.error) this.fail({ sheetId: scan.sheet.sheetId, code: page.error.code, retryable: page.error.retryable, message: page.error.message, range: scan.range })
    for (const cell of page.items) {
      const key = this.visit(scan, cell)
      if (key !== undefined) {
        const group = bandOf(scan.sheet.order.bands, cell.row) ?? cell.row
        if (group !== scan.group) {
          await this.release()
          scan.group = group
        }
        this.hold(key)
      }
      if (this.slice.due(1)) await this.slice.pause('find-scan')
    }
  }

  /** Matches one scanned cell; returns its display key when it is a hit. */
  private visit(scan: ScanState, cell: AuthoritativeCell): number | undefined {
    this.stats.cellsVisited++
    const { sheet } = scan
    const slot = sheet.order.sourceToDisplay(cell.row)
    if (!sheet.visibility.visible(slot)) return undefined
    if (sheet.isColumnVisible && !sheet.isColumnVisible(cell.col)) return undefined
    if (cell.unavailable) {
      this.unavailable++
      return undefined
    }
    let text: string
    if (this.options.textOf) {
      this.stats.textsFormatted++
      text = this.options.textOf(cell, sheet.sheetId)
    } else text = rawTextOf(cell.value)
    if (!firstMatch(text, this.needle, this.query.caseSensitive)) return undefined
    this.stats.hitsRecorded++
    return slot * MAX_COLS + cell.col
  }

  private hold(key: number): void {
    if (this.heldCount + this.ready.size >= this.budget) {
      throw new FindError('resource-limit', `Find would buffer more than ${this.budget} hits; narrow the search or the sheets`)
    }
    if (this.heldCount === this.held.length) {
      const grown = new Float64Array(Math.min(this.budget, Math.max(64, this.held.length * 2)))
      grown.set(this.held)
      this.held = grown
    }
    this.held[this.heldCount++] = key
    this.stats.maxBufferedHits = Math.max(this.stats.maxBufferedHits, this.heldCount + this.ready.size)
    this.account()
  }

  private account(): void {
    this.stats.bufferedHits = this.heldCount + this.ready.size
    this.stats.bufferedBytes = 8 * this.heldCount + 12 * this.ready.size
  }

  /** A band/row group is complete: order its keys cooperatively (sorted 4,096-key
   * runs, then budgeted merge passes) and append them to the ready queue. */
  private async release(): Promise<void> {
    const count = this.heldCount
    if (!count || !this.scan) return
    let source = this.held.subarray(0, count)
    for (let lo = 0; lo < count; lo += ORDER_RUN) {
      source.subarray(lo, Math.min(count, lo + ORDER_RUN)).sort()
      await this.slice.step('find-order', Math.min(ORDER_RUN, count - lo))
    }
    if (count > ORDER_RUN) {
      let target = new Float64Array(count)
      for (let width = ORDER_RUN; width < count; width *= 2) {
        for (let lo = 0; lo < count; lo += 2 * width) {
          const mid = Math.min(count, lo + width), hi = Math.min(count, lo + 2 * width)
          let i = lo, j = mid, k = lo
          while (k < hi) {
            target[k++] = j >= hi || (i < mid && source[i] <= source[j]) ? source[i++] : source[j++]
            if (this.slice.due(1)) await this.slice.pause('find-order')
          }
        }
        const swap = source
        source = target
        target = swap
      }
    }
    this.ready.push(this.nextSheet - 1, source, count)
    this.heldCount = 0
    this.scan.group = undefined
    this.account()
  }

  /** Builds the next ready hit from one source read: only emitted hits become
   * objects. A changed or failed cell is stale or a typed failure, never a guess. */
  private async materialize(): Promise<FindHit | undefined> {
    const { sheet: index, key } = this.ready.shift()
    this.account()
    const sheet = this.sheets[index]
    const slot = Math.floor(key / MAX_COLS), col = key - slot * MAX_COLS
    const row = sheet.order.displayToSource(slot)
    const stamp = { docId: sheet.identity.docId, sheetId: sheet.sheetId, valueGen: sheet.identity.valueGen }
    const read = await this.interruptible(sheet.store.readCell({ row, col }, { signal: this.lifetime.signal, stamp }))
    this.guard()
    if (!read) throw new FindError('closed', 'Find session closed')
    if (read.kind === 'error' && read.code === 'cancelled') throw new FindError('cancelled', 'Find cancelled')
    if (read.kind === 'error' && read.code !== 'stale-source') {
      this.fail({ sheetId: sheet.sheetId, code: read.code, retryable: read.retryable, message: `Cell ${row},${col} could not be read`, range: { r0: row, r1: row, c0: col, c1: col } })
      return undefined
    }
    if (read.kind === 'unknown') {
      this.fail({ sheetId: sheet.sheetId, code: 'incomplete', retryable: true, message: `Cell ${row},${col} is loading`, range: { r0: row, r1: row, c0: col, c1: col } })
      return undefined
    }
    if (read.kind !== 'cell') throw new FindError('stale', `Source cell ${row},${col} of sheet ${sheet.sheetId} changed`)
    const text = this.options.textOf ? this.options.textOf(read.cell, sheet.sheetId) : rawTextOf(read.cell.value)
    const match = firstMatch(text, this.needle, this.query.caseSensitive)
    if (!match) throw new FindError('stale', `Source cell ${row},${col} of sheet ${sheet.sheetId} changed`)
    this.stats.hitsMaterialized++
    return { sheetId: sheet.sheetId, cellId: cellId(row, col), source: { row, col }, text, start: match.start, end: match.end, displaySlot: slot }
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
export function hitGeometry(hit: FindHit, cell: XlsxCell, metrics: SheetMetrics, owner: CacheIdentity, host: CellLayoutHost, cache: LayoutCache, overflow?: { left: number; right: number }): HitGeometry | null {
  const box = metrics.sourceCellToLogical(hit.source)
  if (!box) return null
  // `overflow` is the allowance the painter used (viewport.cellTextOverflow), so
  // the record — and its clip — is the painted one.
  const [record] = buildCellRecords([{ cell, source: hit.source, x: box.x, y: box.y, width: box.width, height: box.height, ...(overflow ? { overflow } : {}) }], owner, host, cache)
  const text = record.text[0]
  const body = record.body.text
  if (!text || !body) return { record, rects: [] }
  if (text.formattedText !== hit.text || !body.text.length) return { record, rects: [{ ...text.clip }] }
  if (body.kind === 'layout') {
    // One unrotated, unwrapped line (aligned right/center, possibly overflowing):
    // the band of [start,end) along its placed segments, like plain runs.
    const line = body.layout.lines.length === 1 && !body.rotation && !body.stacked ? body.layout.lines[0] : undefined
    const span = line && layoutSpan(line.segments, hit.start, hit.end)
    if (!span) return { record, rects: [{ ...text.clip }] }
    const dx = box.x - record.body.x
    const left = Math.max(dx + span.x0, text.clip.x), right = Math.min(dx + span.x1, text.clip.x + text.clip.width)
    if (!(right > left)) return { record, rects: [{ x: box.x, y: box.y, width: box.width, height: box.height }] }
    return { record, rects: [{ x: left, y: box.y, width: right - left, height: box.height }] }
  }
  const x0 = box.x + (body.tx - record.body.x), unit = body.width / body.text.length
  // Highlight only painted ink: intersect with the text clip. A match whose ink
  // is fully clipped is still a hit; it highlights its own cell box.
  const left = Math.max(x0 + unit * hit.start, text.clip.x), right = Math.min(x0 + unit * hit.end, text.clip.x + text.clip.width)
  if (!(right > left)) return { record, rects: [{ x: box.x, y: box.y, width: box.width, height: box.height }] }
  return { record, rects: [{ x: left, y: box.y, width: right - left, height: box.height }] }
}

/** Body-local horizontal extent of source offsets [start,end) over placed
 * segments (proportional within a segment); undefined unless every touched
 * segment is a plain horizontal translation. */
function layoutSpan(segments: readonly { width: number; sourceStart: number; sourceEnd: number; x: number; transform?: { a: number; b: number; c: number; d: number; e: number } }[], start: number, end: number): { x0: number; x1: number } | undefined {
  let x0 = Infinity, x1 = -Infinity
  for (const segment of segments) {
    const from = Math.max(start, segment.sourceStart), to = Math.min(end, segment.sourceEnd), length = segment.sourceEnd - segment.sourceStart
    if (!(to > from) || !(length > 0)) continue
    const t = segment.transform
    if (t && (t.a !== 1 || t.b !== 0 || t.c !== 0 || t.d !== 1)) return undefined
    const origin = t ? t.e : segment.x, unit = segment.width / length
    x0 = Math.min(x0, origin + unit * (from - segment.sourceStart))
    x1 = Math.max(x1, origin + unit * (to - segment.sourceStart))
  }
  return x1 > x0 ? { x0, x1 } : undefined
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

/** One published page: `hits` holds this page only (at most the page size). */
export interface FindPublication {
  key: string
  owner: WorkbookOwner
  query: SearchQuery
  hits: readonly FindHit[]
  pageIndex: number
  complete: boolean
  totalKnown?: number
  cursor?: string
  unavailableCount: number
  incomplete: readonly string[]
  failures: readonly FindSourceFailure[]
}
export interface FindControllerOptions {
  currentOwner: () => WorkbookOwner
  onPublish?: (state: FindPublication) => void
  /** Complete single-page results kept by aggregate key (LRU), at most a page each. */
  cacheEntries?: number
  createSession?: (input: FindInput, query: SearchQuery, options: FindOptions) => FindSession
}
export interface FindSession { nextPage(): Promise<FindPage>; close(): void }
interface FindRun {
  token: number
  key: string
  input: FindInput
  query: SearchQuery
  session: FindSession
  abort: AbortController
  detach: () => void
  pageIndex: number
}

/**
 * Replace/cancel/publish policy for the UI. One run is current: starting
 * another, or cancel(), closes it. Pages are pulled one at a time (run, then
 * next) and each publication carries that page only, so the controller never
 * holds more than the engine's bounded page. Every publication, cached ones
 * included, first re-checks the run token, abort and the host's current owner.
 * Only complete single-page results are cached (by findCacheKey), so the cache
 * is bounded by entries × one page; any view/value/font change misses it.
 */
export class FindController {
  readonly stats = { cacheHits: 0, staleRejected: 0, superseded: 0 }
  published: FindPublication | null = null
  private generation = 0
  private current: FindRun | undefined
  private readonly cache = new Map<string, FindPublication>()
  constructor(private readonly options: FindControllerOptions) {}

  /** A run is open and has further pages to pull. */
  get running(): boolean { return this.current !== undefined }

  /** Closes the current run and releases its session and published page. */
  cancel(): void {
    this.generation++
    this.close()
    this.published = null
  }

  dispose(): void {
    this.cancel()
    this.cache.clear()
  }

  /** Replaces the current run. The last published page stays visible until the
   * replacement publishes (spec §10.1); cancel() is the explicit release. */
  async run(input: FindInput, query: SearchQuery, options: FindOptions = {}): Promise<FindPublication | null> {
    this.generation++
    this.close()
    const token = this.generation
    const key = findCacheKey(input.owner, query)
    if (!this.publishable(token, input.owner, options.signal)) return null
    const cached = this.cache.get(key)
    if (cached) {
      this.cache.delete(key)
      this.cache.set(key, cached)
      this.stats.cacheHits++
      this.publish(cached)
      return cached
    }
    const abort = new AbortController()
    const outer = options.signal
    const onOuter = () => abort.abort()
    outer?.addEventListener('abort', onOuter, { once: true })
    const sessionOptions = { ...options, signal: abort.signal, currentOwner: this.options.currentOwner }
    const session = this.options.createSession?.(input, query, sessionOptions) ?? new WorkbookFind(input, query, sessionOptions)
    this.current = { token, key, input, query, session, abort, detach: () => outer?.removeEventListener('abort', onOuter), pageIndex: 0 }
    return this.pull(this.current)
  }

  /** The next page of the current run, or null when there is none. */
  async next(): Promise<FindPublication | null> {
    return this.current ? this.pull(this.current) : null
  }

  private async pull(run: FindRun): Promise<FindPublication | null> {
    let page: FindPage
    try {
      page = await run.session.nextPage()
    } catch (error) {
      this.end(run)
      // A resource refusal is an outcome the UI must show, not an empty result.
      if (!(error instanceof FindError) || error.code === 'resource-limit') throw error
      if (error.code === 'stale') this.stats.staleRejected++
      else this.stats.superseded++
      return null
    }
    // The session's own last guard ran before this continuation: any owner
    // change, abort or replacement in between must still block publication.
    if (!this.publishable(run.token, run.input.owner, run.abort.signal)) {
      this.end(run)
      return null
    }
    const state: FindPublication = {
      key: run.key, owner: run.input.owner, query: run.query, hits: page.hits, pageIndex: run.pageIndex++, complete: page.complete,
      ...(page.totalKnown !== undefined ? { totalKnown: page.totalKnown } : {}),
      ...(page.cursor !== undefined ? { cursor: page.cursor } : {}),
      unavailableCount: page.unavailableCount, incomplete: page.incomplete, failures: page.failures,
    }
    this.publish(state)
    if (page.complete) {
      if (state.pageIndex === 0 && page.totalKnown !== undefined) this.remember(run.key, state)
      this.end(run)
    }
    return state
  }

  private end(run: FindRun): void {
    if (this.current === run) this.close()
  }

  private close(): void {
    const run = this.current
    this.current = undefined
    if (!run) return
    run.abort.abort()
    run.detach()
    run.session.close()
  }

  /** Checked immediately before every publication, cached ones included (§10.1). */
  private publishable(token: number, owner: WorkbookOwner, signal: AbortSignal | undefined): boolean {
    if (token !== this.generation || signal?.aborted) {
      this.stats.superseded++
      return false
    }
    if (workbookFindKey(this.options.currentOwner()) !== workbookFindKey(owner)) {
      this.stats.staleRejected++
      return false
    }
    return true
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
