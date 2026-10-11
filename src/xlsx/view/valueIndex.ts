import type { ColumnEntry, FilterRange, PagedKeySetRef, SourceFailure, SourceStamp, SourceStore, ValueSearchPredicate, ValuesFilter } from './contracts'
import type { ColumnClassification, PreparedValue } from './filter'
import { FilterDiagnostic, TextFilterMatcher, WorkBudget, type FilterLimits } from './textFilter'
import {
  KEY_BOOLEAN, KEY_DATE, KEY_ERROR, KEY_NUMBER, KEY_STRING, NumberTable, ValueIndexError, loadKeySet,
  sameStamp, valueKeyString, type LaneAccount, type ValueKeySet, type ValueRunStore,
} from './valueRuns'

export { ValueIndexError } from './valueRuns'

/** Complete semantic value universe for one filter column (spec §§7.3–7.4).
 *
 * The worker owns a ValueIndexCache per document session. S4.x maps dropdown
 * open/search/page/OK onto ValueIndexCache.open → build → page/search and
 * ValueSelectionDraft; only bounded ValuePage records (ordinals and bounded
 * labels) cross to the main thread, and full keys stay worker-side (key()).
 * The display limit applies to pages, never to construction, search or rules. */

/** Source lane tags (same numbering as sort.ts); 0 is an uncovered cell. */
export const SOURCE_NUMBER = 1, SOURCE_STRING = 2, SOURCE_BOOLEAN = 3, SOURCE_ERROR = 4
export const SOURCE_BLANK = 5, SOURCE_UNKNOWN = 6, SOURCE_UNAVAILABLE = 7
export const VALUE_DISPLAY_LIMIT = 10_000
export const VALUE_PAGE_LIMIT = 500
export const VALUE_LABEL_LIMIT = 256
/** One bounded body page in typed lanes; texts hold string/error raw text at
 * the cell offset and styles name the display-format owner only. */
export interface ValueLanePage {
  kind: 'values'
  pageId: string
  count: number
  tags: Uint8Array
  numbers: Float64Array
  texts: readonly (string | undefined)[]
  styles?: Uint32Array
}
export type ValuePageRead = ValueLanePage | { kind: 'unknown'; pageId: string } | { kind: 'error'; failure: SourceFailure }
/** Body pages of one column in a fixed order; page numbers are stable under
 * stamp, and stamp is read live so an owner change is observable. */
export interface ValueColumnSource {
  readonly stamp: SourceStamp
  readonly pageCount: number
  readPage(page: number, signal: AbortSignal): Promise<ValuePageRead>
}
/** Every input of key identity. Display labels are deliberately absent. */
export interface ValueIndexIdentity {
  docId: string
  sheetId: string
  valueGen: number
  range: FilterRange
  col: number
  classification: ColumnClassification
  classificationGen: number
}
export type ValueCompleteness = 'complete' | 'loading' | 'incomplete'
export interface ValueIndexLedger {
  sourcePageReads: number
  mergedPages: number
  unknownPages: number
  /** Typed-array allocation steps (parallel lanes grown together count once). */
  allocations: number
  allocatedBytes: number
  /** Currently retained index charge in the cache's worker-memory pool. */
  chargedBytes: number
  peakBytes: number
}
export type ValueLabeler = (kind: 'number' | 'date' | 'boolean', value: number, style: number) => string
export type ValueEntryKind = 'number' | 'date' | 'string' | 'boolean' | 'error'
/** Bounded display record; the exact key is inspected through ValueIndex.key. */
export interface ValueEntryView { ordinal: number; kind: ValueEntryKind; label: string; labelTruncated: boolean; count: number }
/** Matches over the first searchedEntries ordinals; complete only once that is
 * the whole complete universe. Bit i (LSB-first per word) is ordinal i. Only
 * the producing index instance accepts it, and it holds a pool charge until
 * released (ValueIndex.release, draft close, or index disposal). */
export interface ValueSearchResult {
  readonly query: ValueSearchPredicate
  readonly owner: SourceStamp
  readonly matches: Uint32Array
  readonly count: number
  readonly searchedEntries: number
  readonly complete: boolean
}
export interface ValuePageRequest { offset: number; limit: number; search?: ValueSearchResult; label?: ValueLabeler }
export interface ValuePage { entries: readonly ValueEntryView[]; total: number; displayed: number; truncated: boolean; completeness: ValueCompleteness }
export interface ValueProgress { phase: 'pages'; completed: number; total: number }
export interface ValueSearchProgress { searched: number; total: number; matches: number }
interface Pacing {
  /** Maximum uninterrupted work before an event-loop yield (default 8 ms). */
  sliceMs?: number
  now?: () => number
  yield?: () => Promise<void>
}
export interface ValueBuildOptions extends Pacing {
  signal?: AbortSignal
  onProgress?: (progress: ValueProgress) => void
}
export interface ValueSearchOptions extends Pacing {
  signal?: AbortSignal
  label?: ValueLabeler
  limits?: Omit<FilterLimits, 'signal'>
  onProgress?: (progress: ValueSearchProgress) => void
  /** Continue a result over ordinals merged after it was produced; its own
   * label/limits policy is reused unless overridden. */
  resume?: ValueSearchResult
}

const MiB = 1024 * 1024
const LANE_BYTES = 17 // kind 1 + number 8 + style 4 + count 4
// String header + Map entry (as sort.ts charges) + the texts[] reference slot.
const STRING_ENTRY_BYTES = 40 + 128 + 8
const KEY_PATCH_BYTES = 48 // one Map<ordinal, boolean> entry
const RUN_ELEMENT_BYTES = 17 // kind 1 + number 8 + texts reference 8
const PACE_STRIDE = 16
const MERGED = 1, UNKNOWN = 2 // page state 0 = unread
const words = (bits: number): number => Math.ceil(bits / 32)
const hasBit = (set: Uint32Array, at: number): boolean => (set[at >>> 5] & (1 << (at & 31))) !== 0
const setBit = (set: Uint32Array, at: number): void => { set[at >>> 5] |= 1 << (at & 31) }
const clearBit = (set: Uint32Array, at: number): void => { set[at >>> 5] &= ~(1 << (at & 31)) }
function popcount(value: number): number {
  value -= (value >>> 1) & 0x55555555
  value = (value & 0x33333333) + ((value >>> 2) & 0x33333333)
  return (Math.imul((value + (value >>> 4)) & 0x0f0f0f0f, 0x01010101) >>> 24)
}
function boundedLabel(text: string): { label: string; labelTruncated: boolean } {
  if (text.length <= VALUE_LABEL_LIMIT) return { label: text, labelTruncated: false }
  let end = VALUE_LABEL_LIMIT
  const code = text.charCodeAt(end - 1)
  if (code >= 0xd800 && code <= 0xdbff) end--
  return { label: text.slice(0, end), labelTruncated: true }
}
/** Booleans have canonical labels; number/date text needs the caller's
 * formatted display lane, exactly as S8.1 text conditions refuse to guess. */
const defaultLabel: ValueLabeler = (kind, value) => {
  if (kind === 'boolean') return value ? 'TRUE' : 'FALSE'
  throw new FilterDiagnostic('missing-display-text', 'Missing formatted display text lane')
}
function defaultYield(): Promise<void> {
  const host = globalThis as { setImmediate?: (callback: () => void) => unknown }
  return new Promise(resolve => host.setImmediate ? host.setImmediate(resolve) : setTimeout(resolve, 0))
}
/** Time-sliced cooperative yielding: an abort scheduled as a task runs within
 * one slice even when every awaited page resolves immediately. */
class Pacer {
  private last: number
  private readonly now: () => number
  private readonly sliceMs: number
  private readonly pause: () => Promise<void>
  constructor(options: Pacing) {
    this.now = options.now ?? (() => performance.now())
    this.sliceMs = options.sliceMs ?? 8
    if (!Number.isFinite(this.sliceMs) || this.sliceMs < 0) throw new ValueIndexError('invalid-rule', 'Invalid time slice')
    this.pause = options.yield ?? defaultYield
    this.last = this.now()
  }
  get due(): boolean { return this.now() - this.last >= this.sliceMs }
  async yield(): Promise<void> {
    await this.pause()
    this.last = this.now()
  }
}
/** Exact value domain of an index: rules and key runs are valid only there. */
export function valueScope(identity: ValueIndexIdentity): string {
  const { range } = identity
  return JSON.stringify([identity.docId, identity.sheetId, identity.valueGen, range.r0, range.r1, range.c0, range.c1,
    range.headerRows, range.totalsRows, identity.col, identity.classification.kind, identity.classification.hasNumeric, identity.classificationGen])
}
function checkSignal(signal: AbortSignal | undefined, what: string): void {
  if (signal?.aborted) throw new ValueIndexError('cancelled', `${what} cancelled`)
}
function diagnostic(error: unknown): unknown {
  if (error instanceof FilterDiagnostic && error.code !== 'missing-display-text') {
    return new ValueIndexError(error.code === 'cancelled' ? 'cancelled' : error.code === 'resource-limit' ? 'resource-limit' : 'incomplete', error.message)
  }
  if (error instanceof RangeError) return new ValueIndexError('invalid-rule', error.message)
  return error
}

/** A retained charge in the pool, owned by a draft, search result or evaluator. */
export class ValueReservation implements LaneAccount {
  bytes = 0
  private released = false
  constructor(private readonly cache: ValueIndexCache, private readonly owner: ValueIndex | null) {}
  get closed(): boolean { return this.released }
  charge(bytes: number): void {
    if (this.released) throw new ValueIndexError('closed', 'Reservation already released')
    this.cache.charge(this.owner, bytes)
    this.bytes += bytes
  }
  release(bytes = this.bytes): void {
    if (this.released) return
    const amount = Math.min(bytes, this.bytes)
    this.cache.release(amount)
    this.bytes -= amount
  }
  close(): void {
    this.release()
    this.released = true
  }
}

/** Aggregate worker-memory pool for every cached index plus retained search
 * results, draft patches/key sets, evaluators and transient commit buffers;
 * idle LRU indexes are evicted before admission fails. */
export class ValueIndexCache {
  readonly limitBytes: number
  private used = 0
  private readonly indexes = new Map<string, ValueIndex>()
  constructor(options: { limitBytes?: number } = {}) {
    this.limitBytes = options.limitBytes ?? 512 * MiB
    if (!Number.isFinite(this.limitBytes) || this.limitBytes < 0) throw new RangeError('Invalid value index memory limit')
  }
  get usedBytes(): number { return this.used }
  get size(): number { return this.indexes.size }
  /** Returns the cached (complete or resumable) index for exactly this identity. */
  open(identity: ValueIndexIdentity, source: ValueColumnSource): ValueIndex {
    const stamp = { docId: identity.docId, sheetId: identity.sheetId, valueGen: identity.valueGen }
    if (!sameStamp(source.stamp, stamp)) throw new ValueIndexError('stale-source', 'Value source owner does not match the index identity')
    const key = valueScope(identity)
    let index = this.indexes.get(key)
    if (index && !index.disposed) {
      this.indexes.delete(key)
      index.attach(source)
    } else index = new ValueIndex(identity, source, this, key)
    this.indexes.set(key, index)
    return index
  }
  invalidate(predicate: (identity: ValueIndexIdentity) => boolean = () => true): void {
    for (const index of [...this.indexes.values()]) if (predicate(index.identity)) index.dispose()
  }
  reserve(owner: ValueIndex | null = null): ValueReservation { return new ValueReservation(this, owner) }
  /** @internal */
  charge(owner: ValueIndex | null, bytes: number): void {
    if (this.used + bytes > this.limitBytes) {
      for (const index of [...this.indexes.values()]) {
        if (this.used + bytes <= this.limitBytes) break
        if (index !== owner && index.idle) index.dispose()
      }
    }
    if (this.used + bytes > this.limitBytes) {
      throw new ValueIndexError('quota', 'Value index worker memory quota exceeded', {
        resource: 'memory', requiredBytes: this.used + bytes, limitBytes: this.limitBytes, retryable: false,
      })
    }
    this.used += bytes
  }
  /** @internal */
  release(bytes: number): void { this.used -= bytes }
  /** @internal */
  forget(key: string, index: ValueIndex): void {
    if (this.indexes.get(key) === index) this.indexes.delete(key)
  }
}

interface SearchRecord { reservation: ValueReservation; label?: ValueLabeler; limits?: Omit<FilterLimits, 'signal'> }

export class ValueIndex implements LaneAccount {
  readonly ledger: ValueIndexLedger = { sourcePageReads: 0, mergedPages: 0, unknownPages: 0, allocations: 0, allocatedBytes: 0, chargedBytes: 0, peakBytes: 0 }
  readonly stamp: SourceStamp
  readonly scope: string
  blankCount = 0
  unavailableCount = 0
  disposed = false
  private source: ValueColumnSource
  private pins = 0
  private pending: Promise<ValueCompleteness> | undefined
  private pages: Uint8Array | undefined
  private unmerged = 0
  private count = 0
  private kinds = new Uint8Array(0)
  private numbers = new Float64Array(0)
  private styles = new Uint32Array(0)
  private counts = new Uint32Array(0)
  private texts: string[] = []
  private numberIds: NumberTable | undefined
  private readonly booleanIds = new Int32Array([-1, -1])
  private strings = new Map<string, number>()
  private errors = new Map<string, number>()
  private readonly numberKind: number
  private readonly results = new WeakMap<ValueSearchResult, SearchRecord>()
  private readonly reservations = new Set<ValueReservation>()
  constructor(readonly identity: ValueIndexIdentity, source: ValueColumnSource, private readonly cache: ValueIndexCache, key: string) {
    this.stamp = Object.freeze({ docId: identity.docId, sheetId: identity.sheetId, valueGen: identity.valueGen })
    this.scope = key
    this.source = source
    this.numberKind = identity.classification.kind === 'date' ? KEY_DATE : KEY_NUMBER
  }
  get size(): number { return this.count }
  get idle(): boolean { return !this.pins && !this.pending }
  get completeness(): ValueCompleteness {
    if (this.identity.classification.kind === 'loading' || !this.pages || this.unmerged) return 'loading'
    return this.unavailableCount ? 'incomplete' : 'complete'
  }
  /** @internal Charges a typed allocation step before it is made. */
  charge(bytes: number): void {
    this.retain(bytes)
    this.ledger.allocations++
    this.ledger.allocatedBytes += bytes
  }
  private retain(bytes: number): void {
    this.cache.charge(this, bytes)
    this.ledger.chargedBytes += bytes
    this.ledger.peakBytes = Math.max(this.ledger.peakBytes, this.ledger.chargedBytes)
  }
  /** @internal */
  release(bytes: number): void
  /** Releases a search result's retained bitset charge. */
  release(result: ValueSearchResult): void
  release(target: number | ValueSearchResult): void {
    if (typeof target === 'number') {
      this.cache.release(target)
      this.ledger.chargedBytes -= target
      return
    }
    const record = this.results.get(target)
    if (!record) return
    record.reservation.close()
    this.reservations.delete(record.reservation)
  }
  /** @internal */
  attach(source: ValueColumnSource): void { this.source = source }
  /** @internal Drafts pin the index against memory eviction, not invalidation. */
  pin(delta: 1 | -1): void { this.pins += delta }
  /** @internal */
  reserve(): ValueReservation { return this.cache.reserve(this) }
  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    this.cache.release(this.ledger.chargedBytes)
    this.ledger.chargedBytes = 0
    for (const reservation of this.reservations) reservation.close()
    this.reservations.clear()
    this.cache.forget(this.scope, this)
    this.kinds = new Uint8Array(0); this.numbers = new Float64Array(0); this.styles = new Uint32Array(0); this.counts = new Uint32Array(0)
    this.texts = []; this.strings = new Map(); this.errors = new Map(); this.numberIds = undefined
  }
  /** Throws stale-source (and disposes) once invalidated or the live source
   * owner differs; every cached path and publication checks this. */
  assertLive(): void {
    if (!this.disposed && !sameStamp(this.source.stamp, this.stamp)) this.dispose()
    if (this.disposed) throw new ValueIndexError('stale-source', 'Value index was invalidated; retry')
  }
  /** Merges unread/unknown pages in cancellable, time-sliced page batches. A
   * page merges atomically, so cancellation or a source error leaves resumable state. */
  build(options: ValueBuildOptions = {}): Promise<ValueCompleteness> {
    try { this.assertLive() } catch (error) { return Promise.reject(error) }
    if (this.identity.classification.kind === 'loading') return Promise.resolve('loading')
    if (this.pages && !this.unmerged) return Promise.resolve(this.completeness)
    if (this.pending) return this.pending
    this.pending = this.run(options).catch(error => {
      // Partial merges after a quota/owner failure cannot be resumed honestly.
      if (error instanceof ValueIndexError && (error.code === 'quota' || error.code === 'stale-source')) this.dispose()
      throw error
    }).finally(() => { this.pending = undefined })
    return this.pending
  }
  private check(signal: AbortSignal | undefined): void {
    checkSignal(signal, 'Value index build')
    this.assertLive()
  }
  private async run(options: ValueBuildOptions): Promise<ValueCompleteness> {
    const { signal, onProgress } = options, pacer = new Pacer(options)
    const source = this.source, pageCount = source.pageCount
    if (!this.pages) {
      if (!Number.isSafeInteger(pageCount) || pageCount < 0) throw new ValueIndexError('page-error', 'Invalid value source page count')
      this.charge(pageCount + LANE_BYTES * 1024)
      this.pages = new Uint8Array(pageCount)
      this.unmerged = pageCount
      this.grow(1024)
      this.numberIds = new NumberTable(this, 2048)
    } else if (pageCount !== this.pages.length) throw new ValueIndexError('page-error', 'Value source page count changed')
    const readSignal = signal ?? new AbortController().signal
    const progress = () => onProgress?.({ phase: 'pages', completed: pageCount - this.unmerged, total: pageCount })
    for (let page = 0; page < pageCount; page++) {
      if (this.pages[page] === MERGED) continue
      this.check(signal)
      const read = await source.readPage(page, readSignal)
      this.ledger.sourcePageReads++
      this.check(signal)
      if (read.kind === 'error') {
        throw new ValueIndexError(read.failure.code, read.failure.message, { retryable: read.failure.retryable })
      }
      if (read.kind === 'unknown' || !this.admissible(read)) this.pages[page] = UNKNOWN
      else {
        this.merge(read)
        this.pages[page] = MERGED
        this.unmerged--
        this.ledger.mergedPages++
      }
      if (pacer.due) {
        progress()
        await pacer.yield()
        this.check(signal)
      }
    }
    let unknown = 0
    for (let page = 0; page < pageCount; page++) if (this.pages[page] === UNKNOWN) unknown++
    this.ledger.unknownPages = unknown
    progress()
    return this.completeness
  }
  /** Validates a lane page before any merge; unknown cells defer the whole page. */
  private admissible(page: ValueLanePage): boolean {
    const { count, tags, numbers, texts, styles } = page
    if (!Number.isSafeInteger(count) || count < 0 || tags.length !== count || numbers.length < count || (styles && styles.length < count)) {
      throw new ValueIndexError('page-error', 'Malformed value lane page')
    }
    let unknown = false
    for (let i = 0; i < count; i++) {
      const tag = tags[i]
      if (tag === SOURCE_UNKNOWN) unknown = true
      else if (tag < SOURCE_NUMBER || tag > SOURCE_UNAVAILABLE) throw new ValueIndexError('page-error', 'Uncovered value lane cell')
      else if ((tag === SOURCE_STRING || tag === SOURCE_ERROR) && typeof texts[i] !== 'string') throw new ValueIndexError('page-error', 'Missing text for value lane cell')
    }
    return !unknown
  }
  private grow(capacity: number): void {
    const old = this.kinds.length
    if (capacity <= old) return
    if (old) this.charge(capacity * LANE_BYTES)
    const kinds = new Uint8Array(capacity), numbers = new Float64Array(capacity), styles = new Uint32Array(capacity), counts = new Uint32Array(capacity)
    kinds.set(this.kinds); numbers.set(this.numbers); styles.set(this.styles); counts.set(this.counts)
    this.kinds = kinds; this.numbers = numbers; this.styles = styles; this.counts = counts
    if (old) this.release(old * LANE_BYTES)
  }
  private append(kind: number, value: number, style: number): number {
    if (this.count === this.kinds.length) this.grow(this.kinds.length * 2)
    const ordinal = this.count++
    this.kinds[ordinal] = kind
    this.numbers[ordinal] = value
    this.styles[ordinal] = style
    return ordinal
  }
  private intern(map: Map<string, number>, kind: number, text: string, style: number): number {
    const found = map.get(text)
    if (found !== undefined) return found
    this.retain(text.length * 2 + STRING_ENTRY_BYTES)
    const ordinal = this.append(kind, this.texts.length, style)
    this.texts.push(text)
    map.set(text, ordinal)
    return ordinal
  }
  private merge(page: ValueLanePage): void {
    const { count, tags, numbers, texts, styles } = page
    for (let i = 0; i < count; i++) {
      const style = styles ? styles[i] : 0
      let ordinal: number
      switch (tags[i]) {
        case SOURCE_BLANK: this.blankCount++; continue
        case SOURCE_UNAVAILABLE: this.unavailableCount++; continue
        case SOURCE_NUMBER: {
          const value = numbers[i]
          if (!Number.isFinite(value)) {
            // No finite engine representation: the same raw key as a #NUM! cell.
            ordinal = this.intern(this.errors, KEY_ERROR, '#NUM!', style)
            break
          }
          const key = value === 0 ? 0 : value
          ordinal = this.numberIds!.get(key)
          if (ordinal < 0) {
            ordinal = this.append(this.numberKind, key, style)
            this.numberIds!.set(key, ordinal)
          }
          break
        }
        case SOURCE_BOOLEAN: {
          const bit = numbers[i] ? 1 : 0
          ordinal = this.booleanIds[bit]
          if (ordinal < 0) ordinal = this.booleanIds[bit] = this.append(KEY_BOOLEAN, bit, style)
          break
        }
        case SOURCE_STRING: ordinal = this.intern(this.strings, KEY_STRING, texts[i]!, style); break
        default: ordinal = this.intern(this.errors, KEY_ERROR, texts[i]!, style)
      }
      this.counts[ordinal]++
    }
  }
  /** @internal Exact key lanes for membership and run writing. */
  entryKind(ordinal: number): number { return this.kinds[ordinal] }
  /** @internal */
  entryNumber(ordinal: number): number { return this.numbers[ordinal] }
  /** @internal */
  entryText(ordinal: number): string | undefined {
    const kind = this.kinds[ordinal]
    return kind === KEY_STRING || kind === KEY_ERROR ? this.texts[this.numbers[ordinal]] : undefined
  }
  /** Full exact key, for worker-side inspection only (never in display pages). */
  key(ordinal: number): string {
    this.assertLive()
    if (!Number.isSafeInteger(ordinal) || ordinal < 0 || ordinal >= this.count) throw new ValueIndexError('invalid-cursor', 'Value ordinal outside the index')
    const kind = this.kinds[ordinal]
    return valueKeyString(kind, kind === KEY_STRING || kind === KEY_ERROR ? 0 : this.numbers[ordinal], this.entryText(ordinal))
  }
  /** Ordinal of an exact raw value, or -1 (blanks are never entries). */
  find(value: PreparedValue): number {
    this.assertLive()
    switch (value.kind) {
      case 'number':
        if (!Number.isFinite(value.value)) return this.errors.get('#NUM!') ?? -1
        return this.numberIds ? this.numberIds.get(value.value === 0 ? 0 : value.value) : -1
      case 'boolean': return this.booleanIds[value.value ? 1 : 0]
      case 'string': return this.strings.get(value.value) ?? -1
      case 'error': return this.errors.get(value.value) ?? -1
      default: return -1
    }
  }
  private searchText(ordinal: number, label: ValueLabeler): string {
    const kind = this.kinds[ordinal]
    if (kind === KEY_STRING || kind === KEY_ERROR) return this.texts[this.numbers[ordinal]]
    return label(kind === KEY_DATE ? 'date' : kind === KEY_BOOLEAN ? 'boolean' : 'number', this.numbers[ordinal], this.styles[ordinal])
  }
  private view(ordinal: number, label: ValueLabeler): ValueEntryView {
    const kind = this.kinds[ordinal]
    const name: ValueEntryKind = kind === KEY_NUMBER ? 'number' : kind === KEY_DATE ? 'date' : kind === KEY_BOOLEAN ? 'boolean' : kind === KEY_STRING ? 'string' : 'error'
    return { ordinal, kind: name, ...boundedLabel(this.searchText(ordinal, label)), count: this.counts[ordinal] }
  }
  /** @internal Only results produced by this instance (and still held) are accepted. */
  ownsSearch(result: ValueSearchResult): void {
    this.assertLive()
    const record = this.results.get(result)
    if (!record || record.reservation.closed || result.searchedEntries > this.count) {
      throw new ValueIndexError('stale-source', 'Search result belongs to another or a released value index result')
    }
  }
  /** A bounded display page. Only the first VALUE_DISPLAY_LIMIT entries of the
   * universe or search result are displayable; totals stay exact. */
  page(request: ValuePageRequest): ValuePage {
    this.assertLive()
    const { offset, limit, search } = request
    if (!Number.isSafeInteger(offset) || !Number.isSafeInteger(limit) || offset < 0 || limit < 0 || limit > VALUE_PAGE_LIMIT) {
      throw new ValueIndexError('invalid-cursor', 'Invalid value page request')
    }
    if (search) this.ownsSearch(search)
    const label = request.label ?? defaultLabel
    const total = search ? search.count : this.count, displayed = Math.min(total, VALUE_DISPLAY_LIMIT)
    const end = Math.min(displayed, offset + limit), entries: ValueEntryView[] = []
    if (!search) for (let ordinal = offset; ordinal < end; ordinal++) entries.push(this.view(ordinal, label))
    else if (offset < end) {
      let rank = 0
      const set = search.matches
      for (let word = 0; word < words(search.searchedEntries) && rank < end; word++) {
        const bits = set[word]
        if (!bits) continue
        const ones = popcount(bits)
        if (rank + ones <= offset) { rank += ones; continue }
        for (let bit = 0; bit < 32 && rank < end; bit++) {
          if (!(bits & (1 << bit))) continue
          if (rank >= offset) entries.push(this.view(word * 32 + bit, label))
          rank++
        }
      }
    }
    return { entries, total, displayed, truncated: total > VALUE_DISPLAY_LIMIT, completeness: this.completeness }
  }
  /** Searches every merged value with the S8.1 host-collated `contains`
   * semantics over raw text (strings, errors) or display labels (numbers,
   * dates, booleans). Labels never become keys. Time-sliced and cancellable. */
  async search(query: ValueSearchPredicate, options: ValueSearchOptions = {}): Promise<ValueSearchResult> {
    this.assertLive()
    if (!query || typeof query.text !== 'string' || typeof query.locale !== 'string') throw new ValueIndexError('invalid-rule', 'Invalid value search')
    const { signal, resume, onProgress } = options
    let inherited: SearchRecord | undefined
    if (resume) {
      this.ownsSearch(resume)
      inherited = this.results.get(resume)
      if (resume.query.text !== query.text || resume.query.locale !== query.locale) throw new ValueIndexError('invalid-rule', 'Search resume changed its query')
    }
    const labelOption = options.label ?? inherited?.label, limits = options.limits ?? inherited?.limits
    const from = resume?.searchedEntries ?? 0, end = this.count, pacer = new Pacer(options)
    const reservation = this.reserve()
    reservation.charge(words(end) * 4)
    this.reservations.add(reservation)
    try {
      const matches = new Uint32Array(words(end))
      if (resume) matches.set(resume.matches.subarray(0, words(from)))
      let count = resume?.count ?? 0
      const label = labelOption ?? defaultLabel
      try {
        const budget = new WorkBudget({ ...limits, signal })
        const contains = query.text === '' ? undefined : new TextFilterMatcher(query.locale).prepare('contains', query.text, false, budget)
        for (let ordinal = from; ordinal < end; ordinal++) {
          budget.resetAllowance()
          // An empty literal is a subsequence of every candidate.
          if (!contains || contains(this.searchText(ordinal, label), budget)) { setBit(matches, ordinal); count++ }
          if ((ordinal - from) % PACE_STRIDE === PACE_STRIDE - 1 && pacer.due) {
            onProgress?.({ searched: ordinal + 1, total: end, matches: count })
            await pacer.yield()
            checkSignal(signal, 'Value search')
            this.assertLive()
          }
        }
      } catch (error) { throw diagnostic(error) }
      checkSignal(signal, 'Value search')
      this.assertLive()
      onProgress?.({ searched: end, total: end, matches: count })
      const result: ValueSearchResult = Object.freeze({
        query: Object.freeze({ text: query.text, locale: query.locale }), owner: this.stamp, matches, count,
        searchedEntries: end, complete: end === this.count && this.completeness === 'complete',
      })
      this.results.set(result, { reservation, label: labelOption, limits })
      return result
    } catch (error) {
      reservation.close()
      this.reservations.delete(reservation)
      throw error
    }
  }
}

type Layer = { kind: 'search'; result: ValueSearchResult; selected: boolean } | { kind: 'keys'; entries: Map<number, boolean> }
export interface ValueDraftPatchSummary { kind: 'keys' | 'search'; selected: boolean; count: number }
export type ValueCommit = { kind: 'unchanged' | 'committed'; rule: ValuesFilter | null }
/** A dropdown draft pinned to one index owner: base plus ordered patches where
 * the last matching patch wins (spec §7.4). Nothing is stored before OK. Its
 * key set, patches, searches and commit buffers are charged to the pool and
 * released on close. */
export class ValueSelectionDraft {
  closed = false
  private base: 'all' | 'none' | 'committed'
  private layers: Layer[] = []
  private showBlanks: boolean
  private dirty = false
  private keyBytes = 0
  private readonly controllers = new Set<AbortController>()
  private readonly owned = new Set<ValueSearchResult>()
  private current: ValueSearchResult | undefined
  private constructor(
    readonly index: ValueIndex,
    readonly committed: ValuesFilter | null,
    private readonly store: ValueRunStore,
    private committedKeys: ValueKeySet | undefined,
    private readonly reservation: ValueReservation,
  ) {
    this.base = committed ? 'committed' : 'all'
    this.showBlanks = committed ? committed.showBlanks : true
    index.pin(1)
  }
  /** Opens on the applied rule; a rule owned by another generation or column is stale. */
  static async open(index: ValueIndex, committed: ValuesFilter | null, store: ValueRunStore, signal?: AbortSignal): Promise<ValueSelectionDraft> {
    index.assertLive()
    const reservation = index.reserve()
    try {
      let keys: ValueKeySet | undefined
      if (committed) {
        const ref = committedRef(committed)
        if (!sameStamp(ref.owner, index.stamp)) throw new ValueIndexError('stale-source', 'Applied values rule belongs to another source generation')
        keys = await loadKeySet(store, ref, index.scope, signal, reservation)
        checkSignal(signal, 'Value draft open')
        index.assertLive()
      }
      return new ValueSelectionDraft(index, committed, store, keys, reservation)
    } catch (error) {
      reservation.close()
      throw error
    }
  }
  private open(): void {
    if (this.closed) throw new ValueIndexError('closed', 'Value draft is closed')
  }
  private ordinal(ordinal: number): void {
    if (!Number.isSafeInteger(ordinal) || ordinal < 0 || ordinal >= this.index.size) throw new ValueIndexError('invalid-cursor', 'Value ordinal outside the index')
  }
  get blanks(): boolean { return this.showBlanks }
  private baseSelected(ordinal: number): boolean {
    if (this.base !== 'committed') return this.base === 'all'
    const index = this.index, member = this.committedKeys!.has(index.entryKind(ordinal), index.entryNumber(ordinal), index.entryText(ordinal))
    return this.committed!.selection.mode === 'include' ? member : !member
  }
  isSelected(ordinal: number): boolean {
    this.open()
    this.index.assertLive()
    this.ordinal(ordinal)
    for (let at = this.layers.length - 1; at >= 0; at--) {
      const layer = this.layers[at]
      if (layer.kind === 'keys') {
        const selected = layer.entries.get(ordinal)
        if (selected !== undefined) return selected
      } else if (ordinal < layer.result.searchedEntries && hasBit(layer.result.matches, ordinal)) return layer.selected
    }
    return this.baseSelected(ordinal)
  }
  /** An individual checkbox is an exact raw-key patch. */
  toggle(ordinal: number, selected: boolean): void {
    this.open()
    this.index.assertLive()
    this.ordinal(ordinal)
    let top = this.layers[this.layers.length - 1]
    if (!top || top.kind !== 'keys') this.layers.push(top = { kind: 'keys', entries: new Map() })
    if (!top.entries.has(ordinal)) {
      this.reservation.charge(KEY_PATCH_BYTES)
      this.keyBytes += KEY_PATCH_BYTES
    }
    top.entries.set(ordinal, selected)
    this.dirty = true
  }
  setBlanks(selected: boolean): void {
    this.open()
    this.showBlanks = selected
    this.dirty = true
  }
  /** Without search: the complete domain and blanks. With search: exactly the
   * matching subset (also over values merged later), preserving other choices. */
  selectAll(selected: boolean, search?: ValueSearchResult, options: { includeBlanks?: boolean } = {}): void {
    this.open()
    this.index.assertLive()
    if (!search) {
      this.base = selected ? 'all' : 'none'
      this.layers = []
      this.showBlanks = selected
      this.dropCommittedKeys()
      this.reservation.release(this.keyBytes)
      this.keyBytes = 0
      this.releaseUnused()
    } else {
      this.index.ownsSearch(search)
      this.layers.push({ kind: 'search', result: search, selected })
      if (options.includeBlanks) this.showBlanks = selected
    }
    this.dirty = true
  }
  private dropCommittedKeys(): void {
    if (!this.committedKeys) return
    this.reservation.release(this.committedKeys.bytes)
    this.committedKeys = undefined
  }
  /** Releases searches this draft produced that are neither displayed nor
   * referenced by a layer. */
  private releaseUnused(): void {
    for (const result of this.owned) {
      if (result === this.current || this.layers.some(layer => layer.kind === 'search' && layer.result === result)) continue
      this.index.release(result)
      this.owned.delete(result)
    }
  }
  /** Search changes only the displayed list; Cancel aborts it. The previous
   * displayed search is released unless a Select All patch retains it. */
  async search(query: ValueSearchPredicate, options: ValueSearchOptions = {}): Promise<ValueSearchResult> {
    this.open()
    const controller = new AbortController(), outer = options.signal
    const forward = () => controller.abort()
    if (outer?.aborted) controller.abort()
    outer?.addEventListener('abort', forward)
    this.controllers.add(controller)
    try {
      const result = await this.index.search(query, { ...options, signal: controller.signal })
      if (this.closed) {
        this.index.release(result)
        throw new ValueIndexError('cancelled', 'Value draft closed during search')
      }
      this.owned.add(result)
      this.current = result
      this.releaseUnused()
      return result
    } finally {
      outer?.removeEventListener('abort', forward)
      this.controllers.delete(controller)
    }
  }
  /** Final membership bitset over the universe; the caller releases its charge. */
  private membership(): Uint32Array {
    const index = this.index, size = index.size
    this.reservation.charge(words(size) * 4)
    const bits = new Uint32Array(words(size))
    if (this.base === 'all') { bits.fill(0xffffffff); if (size % 32) bits[bits.length - 1] = 2 ** (size % 32) - 1 }
    else if (this.base === 'committed') for (let ordinal = 0; ordinal < size; ordinal++) if (this.baseSelected(ordinal)) setBit(bits, ordinal)
    for (const layer of this.layers) {
      if (layer.kind === 'keys') {
        for (const [ordinal, selected] of layer.entries) (selected ? setBit : clearBit)(bits, ordinal)
        continue
      }
      const matches = layer.result.matches, count = words(layer.result.searchedEntries)
      for (let word = 0; word < count; word++) bits[word] = layer.selected ? bits[word] | matches[word] : bits[word] & ~matches[word]
    }
    return bits
  }
  selectedCount(): { count: number; complete: boolean } {
    this.open()
    this.index.assertLive()
    const bits = this.membership()
    let count = 0
    for (const word of bits) count += popcount(word)
    this.reservation.release(bits.length * 4)
    return { count, complete: this.index.completeness === 'complete' }
  }
  describe(): { base: 'all' | 'none'; patches: readonly ValueDraftPatchSummary[]; blanks: boolean } {
    this.open()
    const patches: ValueDraftPatchSummary[] = []
    let base: 'all' | 'none' = this.base === 'none' ? 'none' : 'all'
    if (this.base === 'committed') {
      const selection = this.committed!.selection, include = selection.mode === 'include'
      base = include ? 'none' : 'all'
      patches.push({ kind: 'keys', selected: include, count: this.committedKeys!.size })
    }
    for (const layer of this.layers) {
      if (layer.kind === 'search') { patches.push({ kind: 'search', selected: layer.selected, count: layer.result.count }); continue }
      let on = 0
      for (const selected of layer.entries.values()) if (selected) on++
      if (on) patches.push({ kind: 'keys', selected: true, count: on })
      if (layer.entries.size - on) patches.push({ kind: 'keys', selected: false, count: layer.entries.size - on })
    }
    return { base, patches, blanks: this.showBlanks }
  }
  private ready(signal: AbortSignal | undefined): void {
    this.open()
    checkSignal(signal, 'Value commit')
    this.index.assertLive()
  }
  /** OK: compacts the draft over the complete universe into include or
   * all-minus-exclusions refs; refuses while the universe is incomplete.
   * Cancellation and owner changes are checked before every publication. */
  async commit(signal?: AbortSignal): Promise<ValueCommit> {
    this.ready(signal)
    const index = this.index
    if (index.completeness !== 'complete') throw new ValueIndexError('incomplete', 'Value universe is incomplete')
    if (!this.dirty) { this.close(); return { kind: 'unchanged', rule: this.committed } }
    for (const layer of this.layers) {
      if (layer.kind === 'search' && layer.result.searchedEntries < index.size) {
        const previous = layer.result
        layer.result = await index.search(previous.query, { resume: previous, signal })
        this.owned.add(layer.result)
        this.ready(signal)
      }
    }
    this.releaseUnused()
    const bits = this.membership(), size = index.size
    let transient = bits.length * 4
    try {
      let selected = 0
      for (const word of bits) selected += popcount(word)
      if (selected === size && (this.showBlanks || index.blankCount === 0)) {
        this.ready(signal)
        this.close()
        return { kind: 'committed', rule: null }
      }
      const include = size - selected > selected, count = include ? selected : size - selected
      this.reservation.charge(count * RUN_ELEMENT_BYTES)
      transient += count * RUN_ELEMENT_BYTES
      const kinds = new Uint8Array(count), numbers = new Float64Array(count), texts: (string | undefined)[] = new Array(count)
      let at = 0
      for (let ordinal = 0; ordinal < size; ordinal++) {
        if (hasBit(bits, ordinal) !== include) continue
        kinds[at] = index.entryKind(ordinal)
        numbers[at] = index.entryText(ordinal) === undefined ? index.entryNumber(ordinal) : 0
        texts[at++] = index.entryText(ordinal)
      }
      const ref = await this.store.put({ count, kinds, numbers, texts }, index.stamp, index.scope, signal)
      try { this.ready(signal) } catch (error) {
        await this.store.release(ref)
        throw error
      }
      this.close()
      const selection = include ? { mode: 'include' as const, keys: ref } : { mode: 'allExcept' as const, excluded: ref }
      return { kind: 'committed', rule: { kind: 'values', selection, showBlanks: this.showBlanks } }
    } finally {
      if (!this.closed) this.reservation.release(transient)
    }
  }
  private close(): void {
    if (this.closed) return
    this.closed = true
    this.index.pin(-1)
    for (const result of this.owned) this.index.release(result)
    this.owned.clear()
    this.current = undefined
    this.layers = []
    this.committedKeys = undefined
    this.reservation.close()
  }
  /** Cancel/Esc: discards the draft and pending searches; stores nothing. */
  cancel(): void {
    for (const controller of this.controllers) controller.abort()
    this.close()
  }
}

function committedRef(rule: ValuesFilter): PagedKeySetRef {
  if (!rule || rule.kind !== 'values' || typeof rule.showBlanks !== 'boolean') throw new ValueIndexError('invalid-rule', 'Not a values rule')
  const selection = rule.selection
  if (selection.mode === 'include') return selection.keys
  if (selection.mode === 'allExcept') return selection.excluded
  // Commit always compacts; uncompacted query patches need the worker index.
  throw new ValueIndexError('unsupported', 'Query-patch values rules must be compacted before evaluation')
}

export interface ValuesMembership {
  matches: (value: PreparedValue) => boolean
  /** Retained key-set charge (in options.cache when given). */
  bytes: number
  release(): void
}
/** Row-level evaluator for an applied values rule — the extension point for
 * FilterEngine (which still defers kind 'values' to this module). The rule is
 * valid only for the exact value domain of identity. */
export async function compileValuesMembership(
  rule: ValuesFilter, identity: ValueIndexIdentity, store: ValueRunStore, options: { cache?: ValueIndexCache; signal?: AbortSignal } = {},
): Promise<ValuesMembership> {
  const ref = committedRef(rule), classification = identity.classification
  const owner = { docId: identity.docId, sheetId: identity.sheetId, valueGen: identity.valueGen }
  if (!sameStamp(ref.owner, owner)) throw new ValueIndexError('stale-source', 'Values rule belongs to another source generation')
  if (classification.kind === 'loading') throw new FilterDiagnostic('incomplete', 'Filter incomplete: column classification is loading')
  const reservation = options.cache?.reserve()
  let keys: ValueKeySet
  try { keys = await loadKeySet(store, ref, valueScope(identity), options.signal, reservation) }
  catch (error) { reservation?.close(); throw error }
  const include = rule.selection.mode === 'include', showBlanks = rule.showBlanks
  const numberKind = classification.kind === 'date' ? KEY_DATE : KEY_NUMBER
  const matches = (value: PreparedValue): boolean => {
    let member: boolean
    switch (value.kind) {
      case 'unknown': case 'unavailable': throw new FilterDiagnostic('incomplete', 'Filter incomplete: authoritative value is unresolved')
      case 'blank': return showBlanks
      case 'number': member = Number.isFinite(value.value) ? keys.has(numberKind, value.value, undefined) : keys.has(KEY_ERROR, 0, '#NUM!'); break
      case 'boolean': member = keys.has(KEY_BOOLEAN, value.value ? 1 : 0, undefined); break
      case 'string': member = keys.has(KEY_STRING, 0, value.value); break
      case 'error': member = keys.has(KEY_ERROR, 0, value.value); break
    }
    return include ? member : !member
  }
  return { matches, bytes: keys.bytes, release: () => reservation?.close() }
}

/** Adapts SourceStore.queryColumn (cell-object pages) to typed lane pages until
 * the S1.2 typed-lanes queryColumn variant exists. One page per row band; the
 * owner is read live from the store. */
export function valueSourceFromStore(store: SourceStore, col: number, body: { r0: number; r1: number }, pageRows = 4096): ValueColumnSource {
  if (!Number.isSafeInteger(pageRows) || pageRows < 1 || body.r1 < body.r0 - 1) throw new RangeError('Invalid value source paging')
  const rows = body.r1 - body.r0 + 1
  return {
    get stamp() { return store.stamp },
    pageCount: Math.ceil(rows / pageRows),
    async readPage(page: number, signal: AbortSignal): Promise<ValuePageRead> {
      const stamp = store.stamp
      const r0 = body.r0 + page * pageRows, r1 = Math.min(body.r1, r0 + pageRows - 1), count = r1 - r0 + 1
      if (!Number.isSafeInteger(page) || page < 0 || count < 1) throw new ValueIndexError('invalid-cursor', 'Value page outside the source')
      const tags = new Uint8Array(count), numbers = new Float64Array(count), styles = new Uint32Array(count), texts: (string | undefined)[] = []
      const cursors = new Set<string>()
      let cursor: string | undefined, loading = false
      do {
        const result = await store.queryColumn({ col, rows: { kind: 'intervals', intervals: [{ r0, r1 }] }, limit: count, cursor }, { signal, stamp })
        checkSignal(signal, 'Value page read')
        if (!sameStamp(result.stamp, stamp) || !sameStamp(store.stamp, stamp)) throw new ValueIndexError('stale-source', 'Column page owner changed')
        if (result.error) return { kind: 'error', failure: { ...result.error } }
        // A covered page that is not complete is still not authoritative (§4.1).
        if (result.completeness === 'incomplete') return { kind: 'error', failure: { code: 'incomplete', retryable: true, message: 'Column page is incomplete' } }
        if (result.completeness !== 'complete') loading = true
        for (const item of result.items) {
          const outcome = place(item, r0, r1, col, tags, numbers, styles, texts)
          if (outcome) return outcome
        }
        cursor = result.nextCursor
        if (cursor !== undefined) {
          if (cursors.has(cursor)) throw new ValueIndexError('invalid-cursor', 'Column cursor made no progress')
          cursors.add(cursor)
        }
      } while (cursor !== undefined)
      if (loading) return { kind: 'unknown', pageId: `${stamp.sheetId}:${col}:${r0}` }
      if (tags.includes(0)) throw new ValueIndexError('page-error', 'Complete column page left rows uncovered')
      return { kind: 'values', pageId: `${stamp.sheetId}:${col}:${r0}`, count, tags, numbers, texts, styles }
    },
  }
}
function place(
  item: ColumnEntry, r0: number, r1: number, col: number,
  tags: Uint8Array, numbers: Float64Array, styles: Uint32Array, texts: (string | undefined)[],
): ValuePageRead | undefined {
  if (item.kind === 'unknown') return { kind: 'unknown', pageId: item.pageId }
  if (item.kind === 'error') return { kind: 'error', failure: { code: item.code, retryable: item.retryable, message: 'Column value source failed' } }
  const first = 'row' in item ? item.row : item.r0, last = 'row' in item ? item.row : item.r1
  if (!Number.isSafeInteger(first) || !Number.isSafeInteger(last) || first < r0 || last > r1 || last < first) {
    throw new ValueIndexError('page-error', 'Column entry outside the requested page')
  }
  for (let row = first; row <= last; row++) if (tags[row - r0]) throw new ValueIndexError('page-error', 'Overlapping column coverage')
  if (item.kind === 'blank') { tags.fill(SOURCE_BLANK, first - r0, last - r0 + 1); return }
  const at = first - r0, cell = item.cell
  if (cell.row !== first || cell.col !== col || cell.id !== first * 16384 + col) throw new ValueIndexError('page-error', 'Cell identity does not match its column entry')
  styles[at] = cell.styleIndex ?? 0
  if (item.kind === 'unavailable') { tags[at] = SOURCE_UNAVAILABLE; return }
  const value = cell.value
  if (cell.valueIsError) { tags[at] = SOURCE_ERROR; texts[at] = String(value) }
  else if (value === null || value === undefined) tags[at] = SOURCE_BLANK
  else if (typeof value === 'number') {
    if (Number.isFinite(value)) { tags[at] = SOURCE_NUMBER; numbers[at] = value }
    else { tags[at] = SOURCE_ERROR; texts[at] = '#NUM!' }
  } else if (typeof value === 'boolean') { tags[at] = SOURCE_BOOLEAN; numbers[at] = value ? 1 : 0 }
  else if (typeof value === 'string') { tags[at] = SOURCE_STRING; texts[at] = value }
  else throw new ValueIndexError('page-error', 'Unsupported authoritative value type')
}
