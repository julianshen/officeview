import type {
  SortSpec, SourceStore, RowOrder, RowVisibility, SourceStamp, SourceFailure, OperationRefusal, FilterRange, ColumnLanes,
} from './contracts'
import { createIdentityOrder, createRowOrder, pinRowPermutation, type CompactRowOrder } from './order'
import { MAX_ROWS, MAX_COLS, CoordinateError, validateCoordinate, validateSourceRect } from './coordinates'
import { CooperativeSlice } from './scheduler'
import { mergeIndexRuns } from './sortMerge'
import { numericTextKeys } from './sortText'
import { BOUNDED_SORT_LANES, type BoundedSortLaneStore } from './sortLanes'
import { buildSortPrefix, type SortPrefix } from './sortPrefix'
import { PREPARED_SORT_READER, type PreparedSortStore } from './sortPrepared'

declare const hiddenSourceBrand: unique symbol
export interface HiddenSourceRows {
  readonly [hiddenSourceBrand]: true
  hidden(sourceRow: number): boolean
}
/** Translate the CURRENT view's DISPLAY mask before a replacement sort starts
 * again from source order. Captured generations must remain immutable. */
export function hiddenSourceRowsFromDisplay(currentOrder: RowOrder, displayVisibility: RowVisibility): HiddenSourceRows {
  return Object.freeze({
    hidden: (sourceRow: number) => !displayVisibility.visible(currentOrder.sourceToDisplay(sourceRow)),
  }) as HiddenSourceRows
}

export type SortPhase = 'read' | 'dictionary' | 'candidates' | 'merge' | 'output' | 'inverse' | 'publish'
export interface SortOptions {
  signal?: AbortSignal
  stamp?: SourceStamp
  locale?: string
  hidden?: HiddenSourceRows
  /** Forbidden legacy DISPLAY mask, including structurally inferred options.
   * Translate the current view with hiddenSourceRowsFromDisplay instead. */
  visibility?: never
  now?: () => number
  yield?: (phase: SortPhase) => Promise<void>
  awaitYield?: (work: Promise<void>) => Promise<void>
  maxWorkerBytes?: number
  residentWorkerBytes?: number
  /** Ownership-aware accounting: the caller guarantees the store's raw TEXT
   * key strings stay retained and already charged in `residentWorkerBytes`
   * (pinned, non-evictable source) for the whole operation. The dictionary then
   * charges only its own reference slot, not a second copy of each string's
   * UTF-16/header bytes. Error texts and ordering are unaffected. */
  sourceRetainsStrings?: boolean
}
export interface SortLedger {
  tagsBytes: number
  numbersBytes: number
  idsBytes: number
  ranksBytes: number
  dictionaryBytes: number
  dictionaryMapPeakBytes: number
  pagePeakBytes: number
  constructionCopiesBytes: number
  fixedOverheadBytes: number
  indexBytes: number
  scratchBytes: number
  runBytes: number
  textKeyBytes: number
  prefixBytes: number
  orderBytes: number
  inverseBytes: number
  retainedBytes: number
  peakBytes: number
  residentWorkerBytes: number
  aggregatePeakBytes: number
}
/** Readonly consumer capability, including views; never expose a writable
 * backing buffer. Runtime storage belongs to the worker generation: consumers
 * must not cast/mutate it. No full-lane copies or per-element proxies are used. */
export interface ReadonlySortBuffer {
  readonly [index: number]: number
  readonly length: number
  readonly byteLength: number
  [Symbol.iterator](): IterableIterator<number>
  subarray(begin?: number, end?: number): ReadonlySortBuffer
}
export interface SortLane {
  readonly tags: ReadonlySortBuffer
  readonly numbers: ReadonlySortBuffer
  readonly ids?: ReadonlySortBuffer
  readonly strings: readonly string[]
  readonly errors: readonly string[]
  readonly stringRanks: ReadonlySortBuffer
  readonly errorRanks: ReadonlySortBuffer
}
export interface PreparedSortKeys {
  spec: SortSpec
  stamp: SourceStamp
  r0: number
  r1: number
  lanes: readonly SortLane[]
  ledger: SortLedger
  locale: string
}
export type SortFailure =
  | { kind: 'cancelled' }
  | { kind: 'incomplete' }
  | { kind: 'refused'; refusal: Extract<OperationRefusal, { reason: 'formula-value-unavailable' }> }
  | { kind: 'error'; error: SourceFailure }
  | { kind: 'resource-limit'; requiredBytes: number; limitBytes: number }
export type PreparedSortResult = SortFailure | { kind: 'complete'; prepared: PreparedSortKeys }
export type SortResult = SortFailure | {
  kind: 'complete'
  order: RowOrder
  metadata: { sort: SortSpec | null; stamp: SourceStamp; locale: string; structuralAdmission: 'caller-required' }
  ledger: SortLedger
}
const CHUNK = 4096
const MiB = 1024 * 1024
const NUMBER = 1, STRING = 2, BOOLEAN = 3, ERROR = 4, BLANK = 5, UNKNOWN = 6, UNAVAILABLE = 7
const COLUMN_TAGS = new Set([NUMBER, STRING, BOOLEAN, ERROR, BLANK, UNAVAILABLE])
const sameStamp = (a: SourceStamp, b: SourceStamp): boolean =>
  a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen
class Halt extends Error {
  constructor(readonly outcome: SortFailure) { super(outcome.kind) }
}
const sourceError = (code: SourceFailure['code'], message: string): Halt =>
  new Halt({ kind: 'error', error: { code, message, retryable: false } })
/** The forbidden DISPLAY-space `visibility` key is a type error, but casts and
 * plain-JS callers bypass types: a present key would otherwise be silently
 * ignored and the sort would run unpinned. Refuse it as an unsupported option
 * instead. An explicitly undefined key carries no mask and stays harmless. */
function rejectLegacyVisibility(options: SortOptions): void {
  if ((options as { visibility?: unknown }).visibility !== undefined) {
    throw sourceError('unsupported', 'Sort options.visibility was replaced by hidden: translate the current view with hiddenSourceRowsFromDisplay')
  }
}
function snapshotSpec(spec: SortSpec): SortSpec {
  validateSourceRect(spec.range)
  if (![0, 1].includes(spec.range.headerRows) || ![0, 1].includes(spec.range.totalsRows)
    || typeof spec.caseSensitive !== 'boolean' || spec.keys.length < 1 || spec.keys.length > 8) {
    throw sourceError('invalid-source-rect', 'Invalid sort specification')
  }
  const cols = allocate(() => new Set<number>())
  const keys = allocate(() => new Array<SortSpec['keys'][number]>())
  for (const key of spec.keys) {
    validateCoordinate(0, key.col)
    if (key.col < spec.range.c0 || key.col > spec.range.c1
      || cols.has(key.col) || !['asc', 'desc'].includes(key.order)) {
      throw sourceError('invalid-coordinate', 'Sort keys must be unique columns in the range with asc/desc direction')
    }
    const col = key.col, order = key.order
    // Read caller fields before classifying native owned-array growth failures.
    const ownedKey = Object.freeze({ col, order })
    allocate(() => cols.add(col))
    allocate(() => keys.push(ownedKey))
  }
  return Object.freeze({
    range: Object.freeze({ ...spec.range }),
    keys: Object.freeze(keys),
    caseSensitive: spec.caseSensitive,
  })
}
function emptyLedger(options: SortOptions): SortLedger {
  return {
    tagsBytes: 0, numbersBytes: 0, idsBytes: 0, ranksBytes: 0, dictionaryBytes: 0,
    dictionaryMapPeakBytes: 0, pagePeakBytes: 0, constructionCopiesBytes: 0, fixedOverheadBytes: 0,
    indexBytes: 0, scratchBytes: 0, runBytes: 0, textKeyBytes: 0, prefixBytes: 0, orderBytes: 0, inverseBytes: 0, retainedBytes: 0, peakBytes: 0,
    residentWorkerBytes: options.residentWorkerBytes ?? 0, aggregatePeakBytes: options.residentWorkerBytes ?? 0,
  }
}
class Budget {
  readonly ledger: SortLedger
  private live = 0
  readonly limit: number
  constructor(options: SortOptions) {
    this.ledger = emptyLedger(options)
    this.limit = options.maxWorkerBytes ?? 512 * MiB
    if (!Number.isFinite(this.limit) || this.limit < 0
      || !Number.isFinite(this.ledger.residentWorkerBytes) || this.ledger.residentWorkerBytes < 0) {
      throw sourceError('resource-limit', 'Invalid worker budget')
    }
    this.ledger.fixedOverheadBytes = 1024
    this.charge(1024)
  }
  charge(bytes: number): void {
    const aggregate = this.ledger.residentWorkerBytes + this.live + bytes
    if (aggregate > this.limit) throw new Halt({ kind: 'resource-limit', requiredBytes: aggregate, limitBytes: this.limit })
    this.live += bytes
    this.ledger.peakBytes = Math.max(this.ledger.peakBytes, this.live)
    this.ledger.aggregatePeakBytes = this.ledger.residentWorkerBytes + this.ledger.peakBytes
  }
  canCharge(bytes: number): boolean { return this.ledger.residentWorkerBytes + this.live + bytes <= this.limit }
  release(bytes: number): void { this.live -= bytes }
  finish(): SortLedger { this.ledger.retainedBytes = this.live; return Object.freeze({ ...this.ledger }) }
}
/** Shared S4.2 budget (scheduler.ts): yield after 4,096 work units or 10 ms,
 * whichever first. The host scheduler's MessageChannel ports are allocated
 * through `allocate` and closed in each public boundary's finally. */
class Cooperative extends CooperativeSlice<SortPhase> {
  constructor(readonly store: SourceStore, readonly stamp: SourceStamp, readonly options: SortOptions) {
    super({ ...options, checkEvery: 256 }, allocate)
  }
  override check(): void {
    if (this.options.signal?.aborted) throw new Halt({ kind: 'cancelled' })
    if (!sameStamp(this.store.stamp, this.stamp)) throw sourceError('stale-source', 'Source owner changed')
  }
}
function errorMessage(error: unknown): string {
  try { return error instanceof Error ? error.message : String(error) }
  catch { return 'Sort operation failed' }
}
function failure(error: unknown, options: SortOptions): SortFailure {
  let message: string | undefined
  try {
    if (options.signal?.aborted) return { kind: 'cancelled' }
    if (error instanceof Halt) return error.outcome
    if (error instanceof CoordinateError) return {
      kind: 'error', error: {
        code: error.code === 'invalid-source-rect' ? 'invalid-source-rect' : 'invalid-coordinate',
        retryable: false, message: error.message,
      },
    }
  } catch {
    // Untrusted thrown objects may reject even instanceof probes. Do not probe
    // that object again, including through error-message conversion.
    message = 'Sort operation failed'
  }
  return {
    kind: 'error', error: {
      code: 'page-error', retryable: false, message: message ?? errorMessage(error),
    },
  }
}
/** Only allocation sites may classify a RangeError as exhausted resources. */
function allocate<T>(create: () => T): T {
  try { return create() }
  catch (error) {
    if (error instanceof RangeError && !(error instanceof CoordinateError)) {
      throw sourceError('resource-limit', error.message)
    }
    throw error
  }
}
function sortCollator(locale: string, caseSensitive: boolean): Intl.Collator {
  try {
    return new Intl.Collator(locale, {
      numeric: true, sensitivity: caseSensitive ? 'variant' : 'accent', caseFirst: 'upper',
    })
  } catch (error) {
    // The frozen source codes have no locale-specific code; unsupported is the
    // closest nonretryable configuration outcome, not an allocation failure.
    throw sourceError('unsupported', error instanceof Error ? error.message : String(error))
  }
}
/** Typed natural-run boundaries are transient, including during dictionary ranking. */
async function mergeIndexes(
  index: Uint32Array, scratch: Uint32Array, count: number,
  compare: (a: number, b: number) => number, cooperative: Cooperative, phase: SortPhase, budget: Budget,
): Promise<Uint32Array> {
  if (count < 2) return index
  const bytes = (Math.ceil(count / 2) + 1) * 4
  budget.charge(bytes)
  budget.ledger.runBytes = Math.max(budget.ledger.runBytes, bytes)
  const starts = allocate(() => new Uint32Array(bytes / 4))
  try { return await mergeIndexRuns(index, scratch, starts, count, compare, cooperative, phase) }
  finally { budget.release(bytes) }
}
interface MutableLane {
  tags: Uint8Array
  numbers: Float64Array
  ids?: Uint32Array
  strings: string[]
  errors: string[]
  stringRanks: Uint32Array
  errorRanks: Uint32Array
}
async function rankDictionary(
  values: readonly string[],
  compare: (a: string, b: string) => number,
  budget: Budget,
  cooperative: Cooperative,
  collator?: Intl.Collator,
): Promise<Uint32Array> {
  const bytes = values.length * 4
  budget.charge(bytes * 3)
  budget.ledger.ranksBytes += bytes
  const ranks = allocate(() => new Uint32Array(values.length))
  const index = allocate(() => new Uint32Array(values.length)), scratch = allocate(() => new Uint32Array(values.length))
  for (let id = 0; id < values.length; id++) { index[id] = id; if (cooperative.due()) await cooperative.pause('dictionary') }
  const fast = collator && values.length >= 128 ? await numericTextKeys(values, collator, {
    charge: bytes => { budget.charge(bytes); budget.ledger.textKeyBytes = Math.max(budget.ledger.textKeyBytes, bytes) },
    release: bytes => budget.release(bytes), allocate,
  }, cooperative, 'dictionary') : undefined
  const compareIds = fast?.compare ?? ((a: number, b: number) => compare(values[a], values[b]))
  try {
    let sorted = await mergeIndexes(index, scratch, values.length, compareIds, cooperative, 'dictionary', budget)
    const assignRanks = async (): Promise<boolean> => {
      let rank = 0
      for (let at = 0; at < sorted.length; at++) {
        // Certify the COMPLETE observed domain, not just a finite alphabet.
        // Host-sorted adjacency proves total ordering; host equality defines
        // rank groups even if a candidate discriminator broke a tie.
        const comparison = at ? compare(values[sorted[at - 1]], values[sorted[at]]) : 0
        if (cooperative.due()) await cooperative.pause('dictionary')
        if (comparison > 0) return false
        if (comparison < 0) rank++
        ranks[sorted[at]] = rank
      }
      return true
    }
    if (!(await assignRanks())) {
      const other = sorted === index ? scratch : index
      sorted = await mergeIndexes(sorted, other, values.length, (a, b) => compare(values[a], values[b]), cooperative, 'dictionary', budget)
      if (!(await assignRanks())) throw sourceError('unsupported', 'Host collation did not define a consistent ordering')
    }
  } finally { fast?.dispose() }
  budget.release(bytes * 2)
  return ranks
}
/** Validate native view shape before charging/reading a typed source page. */
function sortLaneBytes(item: ColumnLanes): number {
  const views = [item.rows, item.tags, item.numbers, item.styles, item.text, item.textStart]
  const names = ['Uint32Array', 'Uint8Array', 'Float64Array', 'Uint32Array', 'Uint16Array', 'Uint32Array']
  for (let i = 0; i < views.length; i++) {
    if (!ArrayBuffer.isView(views[i]) || Object.prototype.toString.call(views[i]) !== `[object ${names[i]}]`) throw sourceError('page-error', 'Malformed column lane buffers')
  }
  const n = item.rows.length
  if (item.tags.length !== n || item.numbers.length !== n || item.styles.length !== n || item.textStart.length !== n + 1
    || item.textStart[0] !== 0 || item.textStart[n] !== item.text.length) throw sourceError('page-error', 'Malformed column lane lengths')
  // Count complete backing buffers, not only subviews that keep them alive.
  const buffers = new Set(views.map(view => view.buffer))
  let bytes = 0
  for (const buffer of buffers) bytes += buffer.byteLength
  return bytes
}
// Empty strings remain TEXT. Blank means an absent cell/run or a null value;
// treating formula/explicit empty text as absent would lose its authoritative type.
const rawStringBytes = (raw: unknown): number => typeof raw === 'string' ? raw.length * 2 + 24 : 0

async function prepare(
  store: SourceStore,
  spec: SortSpec,
  cooperative: Cooperative,
  budget: Budget,
  locale: string,
  forSort = false,
): Promise<PreparedSortKeys> {
  const r0 = spec.range.r0 + spec.range.headerRows, r1 = spec.range.r1 - spec.range.totalsRows
  const length = Math.max(0, r1 - r0 + 1), lanes: SortLane[] = []
  const collator = sortCollator(locale, spec.caseSensitive)
  const typed = (store as BoundedSortLaneStore)[BOUNDED_SORT_LANES] === true && !cooperative.options.sourceRetainsStrings
  let incomplete = false, unavailableCount = 0, firstAddress: { row: number; col: number } | undefined
  // Read hidden keys too: a filter may reveal them later. Unknown/unavailable
  // values therefore conservatively block the same authoritative body domain.
  for (const key of spec.keys) {
    budget.charge(length * 9 + 512)
    budget.ledger.fixedOverheadBytes += 512
    budget.ledger.tagsBytes += length
    budget.ledger.numbersBytes += length * 8
    const lane: MutableLane = {
      tags: allocate(() => new Uint8Array(length)), numbers: allocate(() => new Float64Array(length)),
      strings: [], errors: [],
      stringRanks: allocate(() => new Uint32Array(0)), errorRanks: allocate(() => new Uint32Array(0)),
    }
    let stringIds: Map<string, number> | undefined, errorIds: Map<string, number> | undefined, mapBytes = 0
    function intern(raw: string, error: boolean): number {
      if (!lane.ids) {
        budget.charge(length * 4)
        budget.ledger.idsBytes += length * 4
        lane.ids = allocate(() => new Uint32Array(length))
      }
      const map = error ? (errorIds ??= allocate(() => new Map())) : (stringIds ??= allocate(() => new Map()))
      const dictionary = error ? lane.errors : lane.strings
      const existing = map.get(raw)
      if (existing !== undefined) return existing
      // Conservative app ledger: raw UTF-16 + 24B string header + 16B array
      // capacity/reference + 128B Map entry/capacity. Host overhead is measured separately.
      // A source that already owns and charges the string leaves only the 16B slot.
      const retained = !error && cooperative.options.sourceRetainsStrings === true ? 16 : raw.length * 2 + 40
      budget.charge(retained + 128)
      budget.ledger.dictionaryBytes += retained
      mapBytes += 128
      budget.ledger.dictionaryMapPeakBytes = Math.max(budget.ledger.dictionaryMapPeakBytes, mapBytes)
      const id = dictionary.length
      allocate(() => dictionary.push(raw)); allocate(() => map.set(raw, id))
      return id
    }
    function unavailable(row: number): void {
      lane.tags[row - r0] = UNAVAILABLE; unavailableCount++
      if (!firstAddress || row < firstAddress.row || row === firstAddress.row && key.col < firstAddress.col) firstAddress = { row, col: key.col }
    }
    function put(row: number, value: string | number | boolean | null, error: boolean): void {
      const offset = row - r0
      if (error) { lane.tags[offset] = ERROR; const id = intern(String(value), true); lane.ids![offset] = id }
      else if (value === null) lane.tags[offset] = BLANK
      else if (typeof value === 'number') {
        if (Number.isFinite(value)) { lane.tags[offset] = NUMBER; lane.numbers[offset] = value }
        else { lane.tags[offset] = ERROR; const id = intern('#NUM!', true); lane.ids![offset] = id }
      } else if (typeof value === 'boolean') { lane.tags[offset] = BOOLEAN; lane.numbers[offset] = value ? 1 : 0 }
      else if (typeof value === 'string') { lane.tags[offset] = STRING; const id = intern(value, false); lane.ids![offset] = id }
      else throw sourceError('page-error', 'Unsupported authoritative key type')
    }
    const lease = (store as PreparedSortStore)[PREPARED_SORT_READER]?.(key.col, cooperative.stamp)
    if (lease) {
      try {
        cooperative.check()
        if (!sameStamp(lease.stamp, cooperative.stamp)) throw sourceError('stale-source', 'Prepared owner changed')
        // Completeness of the trusted column proves every absent row is blank.
        for (let at = 0; at < length; at++) {
          lane.tags[at] = BLANK
          if (cooperative.due()) await cooperative.pause('read')
        }
        let previous = r0 - 1
        const read = await lease.read(r0, r1, (row, value, error, missing) => {
          if (!Number.isSafeInteger(row) || row <= previous || row < r0 || row > r1) throw sourceError('page-error', 'Malformed prepared row order')
          previous = row
          if (missing) unavailable(row)
          else put(row, value, error)
        }, { due: () => cooperative.due(), pause: () => cooperative.pause('read') })
        if (read === 'stale') throw sourceError('stale-source', 'Prepared owner changed')
        cooperative.check()
      } finally { lease.release() }
    } else {
    for (let start = r0; start <= r1; start += CHUNK) {
      const end = Math.min(r1, start + CHUNK - 1)
      let cursor: string | undefined, pages = 0, cursorBytes = 0
      const cursors = allocate(() => new Set<string>())
      do {
        cooperative.check()
        // Reserve the bounded producer's JS staging, native output and text
        // before it allocates. A tight budget keeps the original cell path.
        const producerBytes = (end - start + 1) * 384 + 65536
        const producedTyped = typed && budget.canCharge(producerBytes)
        if (producedTyped) {
          budget.charge(producerBytes)
          budget.ledger.pagePeakBytes = Math.max(budget.ledger.pagePeakBytes, producerBytes)
        }
        let page: Awaited<ReturnType<SourceStore['queryColumn']>>
        try {
          page = await store.queryColumn({
            col: key.col, rows: { kind: 'intervals', intervals: [{ r0: start, r1: end }] }, limit: CHUNK, cursor, ...(producedTyped ? { encoding: 'lanes' as const } : {}),
          }, { stamp: cooperative.stamp, signal: cooperative.options.signal ?? allocate(() => new AbortController()).signal })
        } finally { if (producedTyped) budget.release(producerBytes) }
        cooperative.check()
        if (!sameStamp(page.stamp, cooperative.stamp)) throw sourceError('stale-source', 'Column page owner changed')
        if (page.error) throw new Halt({ kind: 'error', error: { ...page.error } })
        if (page.completeness !== 'complete') incomplete = true
        // A bounded source page can retain transient cell objects/raw strings even
        // when a repeated key does not enter the dictionary.
        let pageBytes = 256 * page.items.length
        budget.charge(pageBytes)
        for (const item of page.items) {
          if (item.kind === 'lanes') {
            if (!producedTyped) throw sourceError('page-error', 'Unrequested column lanes')
            const bytes = sortLaneBytes(item)
            pageBytes += bytes
            budget.charge(bytes)
          } else if (item.kind === 'cell' || item.kind === 'unavailable') {
            const cell = item.cell
            const bytes = rawStringBytes(cell.value) + rawStringBytes(cell.formula)
              + rawStringBytes(cell.arrayRef) + rawStringBytes(cell.formulaType)
              + rawStringBytes(cell.sharedFormula?.ref)
            pageBytes += bytes
            budget.charge(bytes)
          }
          if (cooperative.due()) await cooperative.pause('read')
        }
        budget.ledger.pagePeakBytes = Math.max(budget.ledger.pagePeakBytes, pageBytes)
        for (const item of page.items) {
          const first = 'row' in item ? item.row : item.r0, last = 'row' in item ? item.row : item.r1
          if (!Number.isSafeInteger(first) || !Number.isSafeInteger(last)
            || first < start || last > end || last < first) {
            throw sourceError('page-error', 'Column entry outside requested chunk')
          }
          if (item.kind === 'lanes') {
            // The run proves completeness over absent rows too. Populate blanks
            // first, then replace only validated, ordered native-lane records.
            for (let row = first; row <= last; row++) {
              if (lane.tags[row - r0]) throw sourceError('page-error', 'Overlapping column coverage')
              lane.tags[row - r0] = BLANK
              if (cooperative.due()) await cooperative.pause('read')
            }
            let previous = first - 1
            for (let at = 0; at < item.rows.length; at++) {
              const row = item.rows[at], tag = item.tags[at], number = item.numbers[at]
              const from = item.textStart[at], to = item.textStart[at + 1]
              if (row <= previous || row < first || row > last || from > to || to > item.text.length
                || !COLUMN_TAGS.has(tag)
                || tag === NUMBER && !Number.isFinite(number) || tag === BOOLEAN && number !== 0 && number !== 1) throw sourceError('page-error', 'Malformed column lane record')
              previous = row
              const offset = row - r0
              lane.tags[offset] = tag
              if (tag === NUMBER || tag === BOOLEAN) lane.numbers[offset] = number
              else if (tag === UNAVAILABLE) {
                unavailableCount++
                if (!firstAddress || row < firstAddress.row || row === firstAddress.row && key.col < firstAddress.col) firstAddress = { row, col: key.col }
              } else if (tag === STRING || tag === ERROR) {
                const bytes = (to - from) * 2 + 24
                budget.charge(bytes)
                let value = ''
                for (let pos = from; pos < to; pos += 256) {
                  const end = Math.min(to, pos + 256)
                  value = allocate(() => value + String.fromCharCode(...item.text.subarray(pos, end)))
                  // Explicit boundary around every bounded decode block.
                  cooperative.check()
                  if (cooperative.due()) await cooperative.pause('read')
                }
                const id = intern(value, tag === ERROR)
                lane.ids![offset] = id
                budget.release(bytes)
              }
              if (cooperative.due()) await cooperative.pause('read')
            }
            continue
          }
          for (let row = first; row <= last; row++) {
            const offset = row - r0
            if (lane.tags[offset]) throw sourceError('page-error', 'Overlapping column coverage')
            switch (item.kind) {
              case 'blank': lane.tags[offset] = BLANK; break
              case 'unknown': lane.tags[offset] = UNKNOWN; incomplete = true; break
              case 'error': throw new Halt({
                kind: 'error', error: { code: item.code, retryable: item.retryable, message: 'Column key source failed' },
              })
              case 'unavailable': {
                lane.tags[offset] = UNAVAILABLE; unavailableCount++
                if (!firstAddress || row < firstAddress.row
                  || row === firstAddress.row && key.col < firstAddress.col) firstAddress = { row, col: key.col }
                break
              }
              case 'cell': {
                const cell = item.cell, value = cell.value
                if (cell.row !== row || cell.col !== key.col || cell.id !== row * MAX_COLS + key.col) {
                  throw sourceError('page-error', 'Cell identity does not match its key entry')
                }
                put(row, value, Boolean(cell.valueIsError))
                break
              }
            }
            if (cooperative.due()) await cooperative.pause('read')
          }
        }
        budget.release(pageBytes)
        cursor = page.nextCursor
        if (cursor !== undefined) {
          if (cursors.has(cursor) || ++pages > CHUNK) throw sourceError('invalid-cursor', 'Column cursor made no progress')
          const bytes = cursor.length * 2 + 128
          budget.charge(bytes); cursorBytes += bytes
          const nextCursor = cursor
          allocate(() => cursors.add(nextCursor))
        }
        if (cooperative.due()) await cooperative.pause('read')
      } while (cursor !== undefined)
      budget.release(cursorBytes)
      for (let row = start; row <= end; row++) {
        if (!lane.tags[row - r0]) incomplete = true
        if (cooperative.due()) await cooperative.pause('read')
      }
    }
    }
    // Intern maps are no longer needed once the column is complete. Drop them
    // before rank/collation scratch so their payloads do not coexist at peak.
    stringIds = undefined; errorIds = undefined
    budget.release(mapBytes)
    lane.stringRanks = await rankDictionary(lane.strings, collator.compare, budget, cooperative, collator)
    lane.errorRanks = await rankDictionary(lane.errors, (a, b) => a < b ? -1 : a > b ? 1 : 0, budget, cooperative)
    if (forSort) {
      // The private operation lanes become direction-normalized keys. Public
      // prepareSortKeys still publishes raw values; no extra per-row copy is needed.
      const descending = key.order === 'desc'
      for (let at = 0; at < length; at++) {
        const tag = lane.tags[at]
        if (tag === STRING) lane.numbers[at] = lane.stringRanks[lane.ids![at]]
        else if (tag === ERROR) lane.numbers[at] = lane.errorRanks[lane.ids![at]]
        if (descending && tag < ERROR) {
          lane.tags[at] = 4 - tag
          lane.numbers[at] = -lane.numbers[at]
        }
        if (cooperative.due()) await cooperative.pause('read')
      }
    }
    const publishedLane = Object.freeze({ ...lane, strings: Object.freeze(lane.strings), errors: Object.freeze(lane.errors) })
    allocate(() => lanes.push(publishedLane))
  }
  if (firstAddress) throw new Halt({
    kind: 'refused', refusal: { reason: 'formula-value-unavailable', count: unavailableCount, firstAddress },
  })
  if (incomplete) throw new Halt({ kind: 'incomplete' })
  cooperative.check()
  return Object.freeze({ spec, stamp: cooperative.stamp, r0, r1, lanes: Object.freeze(lanes), ledger: budget.finish(), locale })
}
export async function prepareSortKeys(
  store: SourceStore,
  input: SortSpec,
  options: SortOptions = {},
): Promise<PreparedSortResult> {
  let cooperative: Cooperative | undefined
  try {
    options = Object.freeze({ ...options })
    rejectLegacyVisibility(options)
    const stamp = Object.freeze({ ...(options.stamp ?? store.stamp) }), spec = snapshotSpec(input)
    const budget = new Budget(options)
    cooperative = new Cooperative(store, stamp, options)
    cooperative.check()
    return { kind: 'complete', prepared: await prepare(store, spec, cooperative, budget, options.locale ?? 'en') }
  } catch (error) { return failure(error, options) }
  finally { cooperative?.dispose() }
}
function compareRows(prepared: PreparedSortKeys, a: number, b: number, prefix?: SortPrefix): number {
  const first = a - prepared.r0, second = b - prepared.r0
  if (prefix) {
    const comparison = prefix.codes[first] - prefix.codes[second]
    if (comparison) return comparison
  }
  for (let key = prefix?.keys ?? 0; key < prepared.lanes.length; key++) {
    const lane = prepared.lanes[key]
    const comparison = lane.tags[first] - lane.tags[second] || lane.numbers[first] - lane.numbers[second]
    if (comparison) return comparison
  }
  return a - b
}
/** Large bands own their buffers and construct/validate the inverse cooperatively.
 * S2.1 helpers synchronously copy/scan entire bands: keep them for bounded bands,
 * while this private builder preserves their frozen contract without shared edits. */
function ownedOrder(
  r0: number,
  r1: number,
  sourceAtOffset: Uint32Array,
  displayOffsetOfSource: Uint32Array,
): CompactRowOrder {
  const band = Object.freeze({ r0, r1, sourceAtOffset, displayOffsetOfSource })
  return Object.freeze({
    kind: 'bands', bands: Object.freeze([band]),
    diagnostics: Object.freeze({
      typedArrayBytes: sourceAtOffset.byteLength + displayOffsetOfSource.byteLength, bandCount: 1,
    }),
    displayToSource(row: number) {
      validateCoordinate(row, 0)
      return row >= r0 && row <= r1 ? sourceAtOffset[row - r0] : row
    },
    sourceToDisplay(row: number) {
      validateCoordinate(row, 0)
      return row >= r0 && row <= r1 ? r0 + displayOffsetOfSource[row - r0] : row
    },
  })
}
/** Pure draft only. S7.1/S3 must admit whole-row structural/outside-column effects
 * before a caller commits this object; this kernel cannot authorize those guards. */
export async function sortRowOrder(store: SourceStore, input: SortSpec | null, options: SortOptions = {}): Promise<SortResult> {
  let cooperative: Cooperative | undefined
  try {
    options = Object.freeze({ ...options })
    rejectLegacyVisibility(options)
    const stamp = Object.freeze({ ...(options.stamp ?? store.stamp) })
    const spec = input === null ? null : snapshotSpec(input), locale = options.locale ?? 'en'
    const budget = new Budget(options)
    cooperative = new Cooperative(store, stamp, options)
    cooperative.check()
    let order: RowOrder = createIdentityOrder(), emptyBody = false
    if (spec !== null) {
      const prepared = await prepare(store, spec, cooperative, budget, locale, true)
      const length = Math.max(0, prepared.r1 - prepared.r0 + 1)
      if (length) {
        const pinned = options.hidden ? (row: number) => options.hidden!.hidden(row) : () => false
        budget.charge(length * 8)
        budget.ledger.indexBytes = length * 4; budget.ledger.scratchBytes = length * 4
        const index = allocate(() => new Uint32Array(length)), scratch = allocate(() => new Uint32Array(length))
        let count = 0
        for (let row = prepared.r0; row <= prepared.r1; row++) {
          if (!pinned(row)) index[count++] = row
          if (cooperative.due()) await cooperative.pause('candidates')
        }
        let prefix = await buildSortPrefix(prepared.lanes, length, {
          canCharge: bytes => budget.canCharge(bytes), charge: bytes => budget.charge(bytes),
          release: bytes => budget.release(bytes), allocate,
        }, cooperative, 'read')
        budget.ledger.prefixBytes = prefix?.bytes ?? 0
        const sorted = await mergeIndexes(index, scratch, count, (a, b) => compareRows(prepared, a, b, prefix), cooperative, 'merge', budget)
        if (prefix) { budget.release(prefix.bytes); prefix = undefined }
        if (length <= CHUNK) {
          // pinRowPermutation peak: candidates + rows + copied order + inverse;
          // createRowOrder then copies order/inverse again while that band lives.
          budget.charge(length * 20)
          budget.ledger.constructionCopiesBytes = length * 12
          // Recheck external pins OUTSIDE allocation boundaries, using the freed
          // merge buffer for a stable snapshot. Helpers now receive only validated
          // worker-owned inputs/pins; their remaining RangeErrors are allocations.
          const pins = sorted === index ? scratch : index
          let movable = 0
          for (let offset = 0; offset < length; offset++) {
            pins[offset] = pinned(prepared.r0 + offset) ? 1 : 0
            if (!pins[offset]) movable++
            if (cooperative.due()) await cooperative.pause('output')
          }
          if (movable !== count) throw sourceError('page-error', 'Pinned population changed while sorting')
          for (let at = 0; at < count; at++) {
            if (pins[sorted[at] - prepared.r0]) throw sourceError('page-error', 'Pinned population changed while sorting')
          }
          const candidates = allocate(() => sorted.subarray(0, count))
          const band = allocate(() => pinRowPermutation(
            prepared.r0, prepared.r1, row => pins[row - prepared.r0] === 1, candidates,
          ))
          order = allocate(() => createRowOrder([band]))
          budget.release(length * 12)
        } else {
          budget.charge(length * 8)
          const rows = allocate(() => new Uint32Array(length)), inverse = allocate(() => new Uint32Array(length))
          let at = 0
          for (let offset = 0; offset < length; offset++) {
            const row = prepared.r0 + offset
            rows[offset] = pinned(row) ? row : sorted[at++]
            inverse[offset] = MAX_ROWS
            if (cooperative.due()) await cooperative.pause('output')
          }
          if (at !== count) throw sourceError('page-error', 'Pinned population changed while sorting')
          for (let offset = 0; offset < length; offset++) {
            const row = rows[offset]
            if (row < prepared.r0 || row > prepared.r1 || inverse[row - prepared.r0] !== MAX_ROWS
              || pinned(row) && row !== prepared.r0 + offset) {
              throw sourceError('page-error', 'Sort permutation is not a source bijection')
            }
            inverse[row - prepared.r0] = offset
            if (cooperative.due()) await cooperative.pause('inverse')
          }
          order = ownedOrder(prepared.r0, prepared.r1, rows, inverse)
        }
        budget.ledger.orderBytes = length * 4; budget.ledger.inverseBytes = length * 4
        budget.release(length * 8)
      } else emptyBody = true
      // Prepared lanes and fixed work reservations are private to this operation.
      const keyBytes = budget.ledger.tagsBytes + budget.ledger.numbersBytes
        + budget.ledger.idsBytes + budget.ledger.ranksBytes + budget.ledger.dictionaryBytes
      budget.release(keyBytes + budget.ledger.fixedOverheadBytes)
    }
    if (spec === null) budget.release(budget.ledger.fixedOverheadBytes)
    // An empty protected body returns the shared identity with no retained worker
    // work. Its small metadata object is caller bookkeeping, outside this pool.
    if (!emptyBody) budget.charge(256)
    await cooperative.pause('publish')
    return {
      kind: 'complete', order,
      metadata: Object.freeze({ sort: spec, stamp, locale, structuralAdmission: 'caller-required' }),
      ledger: budget.finish(),
    }
  } catch (error) { return failure(error, options) }
  finally { cooperative?.dispose() }
}
export type SortCommand =
  | { kind: 'cycle'; col: number; shift?: boolean; range: FilterRange }
  | { kind: 'toggle-case' }
  | { kind: 'clear' }
export type SortCommandNotice =
  | { reason: 'too-many-sort-keys'; maxKeys: 8 }
  | { reason: 'no-active-sort' }
  | { reason: 'invalid-sort-command'; error: SourceFailure }
/** Changed bounds start a fresh single ascending key even with Shift. Keep the
 * user's case preference; it is independent of range/key ownership. Refusals
 * preserve the exact previous state, including malformed external drafts. */
export function reduceSortCommand(sort: SortSpec | null, command: SortCommand): {
  sort: SortSpec | null; notice?: SortCommandNotice
} {
  try {
    if (command.kind === 'clear') return { sort: null }
    const previous = sort ? snapshotSpec(sort) : null
    if (command.kind === 'toggle-case') {
      return previous
        ? { sort: snapshotSpec({ ...previous, caseSensitive: !previous.caseSensitive }) }
        : { sort, notice: { reason: 'no-active-sort' } }
    }
    if (command.kind !== 'cycle') throw sourceError('invalid-source-rect', 'Invalid sort command')
    const next = snapshotSpec({
      range: command.range, keys: [{ col: command.col, order: 'asc' }],
      caseSensitive: previous?.caseSensitive ?? false,
    })
    const range = previous?.range
    const sameRange = range && range.r0 === next.range.r0 && range.r1 === next.range.r1
      && range.c0 === next.range.c0 && range.c1 === next.range.c1
      && range.headerRows === next.range.headerRows && range.totalsRows === next.range.totalsRows
    if (!sameRange) return { sort: next }
    const existing = previous!.keys.find(key => key.col === command.col)
    const direction = existing?.order === 'asc' ? 'desc' : existing?.order === 'desc' ? null : 'asc'
    const keys = command.shift ? [...previous!.keys] : []
    const at = keys.findIndex(key => key.col === command.col)
    if (direction === null) { if (at >= 0) keys.splice(at, 1) }
    else if (at >= 0) keys[at] = { col: command.col, order: direction }
    else {
      if (keys.length === 8) return { sort, notice: { reason: 'too-many-sort-keys', maxKeys: 8 } }
      keys.push({ col: command.col, order: direction })
    }
    return { sort: keys.length ? snapshotSpec({ ...next, keys }) : null }
  } catch (error) {
    let code: 'invalid-coordinate' | 'invalid-source-rect' = 'invalid-source-rect'
    // External getters can throw arbitrary values, even hostile conversion or
    // prototype hooks. Preserve typed internal validation without calling them.
    try {
      if (error instanceof CoordinateError && error.code === 'invalid-coordinate') code = 'invalid-coordinate'
      else if (error instanceof Halt && error.outcome.kind === 'error'
        && error.outcome.error.code === 'invalid-coordinate') code = 'invalid-coordinate'
    } catch { /* Keep the fixed shape diagnostic for hostile thrown objects. */ }
    return { sort, notice: { reason: 'invalid-sort-command', error: {
      code, retryable: false, message: 'Invalid sort command or draft shape',
    } } }
  }
}
// TODO(S7.3): exact external text runs/merge for huge unique domains. U raw
// strings cost sum(2*UTF16 length + 40B) retained + 128U Map construction bytes,
// 12U rank/index/scratch bytes, plus row lanes and order workspace. Refuse when
// their aggregate construction peak exceeds the caller's remaining worker pool.
