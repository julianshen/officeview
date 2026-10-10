import type { SortSpec, SourceStore, RowOrder, RowVisibility, SourceStamp, SourceFailure, OperationRefusal, FilterRange } from './contracts'
import { createIdentityOrder, createRowOrder, pinRowPermutation, hiddenSlotPins, type CompactRowOrder } from './order'
import { MAX_ROWS, CoordinateError, validateCoordinate, validateSourceRect } from './coordinates'

export type SortPhase = 'read' | 'dictionary' | 'candidates' | 'merge' | 'output' | 'inverse' | 'publish'
export interface SortOptions {
  signal?: AbortSignal
  stamp?: SourceStamp
  locale?: string
  visibility?: RowVisibility
  now?: () => number
  yield?: (phase: SortPhase) => Promise<void>
  maxWorkerBytes?: number
  residentWorkerBytes?: number
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
  orderBytes: number
  inverseBytes: number
  retainedBytes: number
  peakBytes: number
  residentWorkerBytes: number
  aggregatePeakBytes: number
}
export interface SortLane {
  tags: Uint8Array
  numbers: Float64Array
  ids?: Uint32Array
  strings: readonly string[]
  errors: readonly string[]
  stringRanks: Uint32Array
  errorRanks: Uint32Array
}
export interface PreparedSortKeys { spec: SortSpec; stamp: SourceStamp; r0: number; r1: number; lanes: readonly SortLane[]; ledger: SortLedger; locale: string }
export type SortFailure = { kind: 'cancelled' } | { kind: 'incomplete' } | { kind: 'refused'; refusal: Extract<OperationRefusal, { reason: 'formula-value-unavailable' }> } | { kind: 'error'; error: SourceFailure } | { kind: 'resource-limit'; requiredBytes: number; limitBytes: number }
export type PreparedSortResult = SortFailure | { kind: 'complete'; prepared: PreparedSortKeys }
export type SortResult = SortFailure | { kind: 'complete'; order: RowOrder; metadata: { sort: SortSpec | null; stamp: SourceStamp; locale: string; structuralAdmission: 'caller-required' }; ledger: SortLedger }
const CHUNK = 4096
const MiB = 1024 * 1024
const NUMBER = 1, STRING = 2, BOOLEAN = 3, ERROR = 4, BLANK = 5, UNKNOWN = 6, UNAVAILABLE = 7
const sameStamp = (a: SourceStamp, b: SourceStamp): boolean => a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen
class Halt extends Error {
  constructor(readonly outcome: SortFailure) { super(outcome.kind) }
}
const sourceError = (code: SourceFailure['code'], message: string): Halt => new Halt({ kind: 'error', error: { code, message, retryable: false } })
function snapshotSpec(spec: SortSpec): SortSpec {
  validateSourceRect(spec.range)
  if (![0, 1].includes(spec.range.headerRows) || ![0, 1].includes(spec.range.totalsRows)
    || typeof spec.caseSensitive !== 'boolean' || spec.keys.length < 1 || spec.keys.length > 8) throw sourceError('invalid-source-rect', 'Invalid sort specification')
  const cols = new Set<number>()
  for (const key of spec.keys) {
    validateCoordinate(0, key.col)
    if (key.col < spec.range.c0 || key.col > spec.range.c1 || cols.has(key.col) || !['asc', 'desc'].includes(key.order)) throw sourceError('invalid-coordinate', 'Sort keys must be unique columns in the range with asc/desc direction')
    cols.add(key.col)
  }
  return Object.freeze({ range: Object.freeze({ ...spec.range }), keys: Object.freeze(spec.keys.map(key => Object.freeze({ ...key }))), caseSensitive: spec.caseSensitive })
}
function emptyLedger(options: SortOptions): SortLedger {
  return { tagsBytes: 0, numbersBytes: 0, idsBytes: 0, ranksBytes: 0, dictionaryBytes: 0, dictionaryMapPeakBytes: 0, pagePeakBytes: 0, constructionCopiesBytes: 0, fixedOverheadBytes: 0, indexBytes: 0, scratchBytes: 0, orderBytes: 0, inverseBytes: 0, retainedBytes: 0, peakBytes: 0, residentWorkerBytes: options.residentWorkerBytes ?? 0, aggregatePeakBytes: options.residentWorkerBytes ?? 0 }
}
class Budget {
  readonly ledger: SortLedger
  private live = 0
  readonly limit: number
  constructor(options: SortOptions) {
    this.ledger = emptyLedger(options)
    this.limit = options.maxWorkerBytes ?? 512 * MiB
    if (!Number.isFinite(this.limit) || this.limit < 0 || !Number.isFinite(this.ledger.residentWorkerBytes) || this.ledger.residentWorkerBytes < 0) throw sourceError('resource-limit', 'Invalid worker budget')
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
  release(bytes: number): void { this.live -= bytes }
  finish(): SortLedger { this.ledger.retainedBytes = this.live; return Object.freeze({ ...this.ledger }) }
}
class Cooperative {
  private work = 0
  private last: number
  readonly now: () => number
  readonly schedule: (phase: SortPhase) => Promise<void>
  constructor(readonly store: SourceStore, readonly stamp: SourceStamp, readonly options: SortOptions) {
    this.now = options.now ?? (() => performance.now())
    this.schedule = options.yield ?? (() => new Promise(resolve => setTimeout(resolve, 0)))
    this.last = this.now()
  }
  check(): void {
    if (this.options.signal?.aborted) throw new Halt({ kind: 'cancelled' })
    if (!sameStamp(this.store.stamp, this.stamp)) throw sourceError('stale-source', 'Source owner changed')
  }
  due(): boolean { this.check(); return ++this.work >= CHUNK || this.now() - this.last >= 10 }
  async pause(phase: SortPhase): Promise<void> {
    this.check()
    await this.schedule(phase)
    this.check()
    this.work = 0
    this.last = this.now()
  }
}
function failure(error: unknown, options: SortOptions): SortFailure {
  if (options.signal?.aborted) return { kind: 'cancelled' }
  if (error instanceof Halt) return error.outcome
  if (error instanceof CoordinateError) return { kind: 'error', error: { code: error.code === 'invalid-source-rect' ? 'invalid-source-rect' : 'invalid-coordinate', retryable: false, message: error.message } }
  if (error instanceof RangeError) return { kind: 'error', error: { code: 'resource-limit', retryable: false, message: error.message } }
  return { kind: 'error', error: { code: 'page-error', retryable: true, message: error instanceof Error ? error.message : String(error) } }
}
/** Stable bottom-up index merge: both initialization and comparison/output work yield.
 * Native Array/TypedArray.sort cannot provide an abort boundary. */
async function mergeIndexes(index: Uint32Array, scratch: Uint32Array, count: number, compare: (a: number, b: number) => number, cooperative: Cooperative, phase: SortPhase): Promise<Uint32Array> {
  let input = index, output = scratch
  for (let width = 1; width < count; width *= 2) {
    for (let start = 0; start < count; start += width * 2) {
      const middle = Math.min(start + width, count), end = Math.min(start + width * 2, count)
      let left = start, right = middle
      for (let at = start; at < end; at++) {
        output[at] = left < middle && (right >= end || compare(input[left], input[right]) <= 0) ? input[left++] : input[right++]
        if (cooperative.due()) await cooperative.pause(phase)
      }
    }
    const swap = input; input = output; output = swap
  }
  return input
}
interface MutableLane extends SortLane { strings: string[]; errors: string[] }
async function rankDictionary(values: readonly string[], compare: (a: string, b: string) => number, budget: Budget, cooperative: Cooperative): Promise<Uint32Array> {
  const bytes = values.length * 4
  budget.charge(bytes * 3)
  budget.ledger.ranksBytes += bytes
  const ranks = new Uint32Array(values.length), index = new Uint32Array(values.length), scratch = new Uint32Array(values.length)
  for (let id = 0; id < values.length; id++) { index[id] = id; if (cooperative.due()) await cooperative.pause('dictionary') }
  const sorted = await mergeIndexes(index, scratch, values.length, (a, b) => compare(values[a], values[b]), cooperative, 'dictionary')
  let rank = 0
  for (let at = 0; at < sorted.length; at++) {
    if (at && compare(values[sorted[at - 1]], values[sorted[at]]) !== 0) rank++
    ranks[sorted[at]] = rank
    if (cooperative.due()) await cooperative.pause('dictionary')
  }
  budget.release(bytes * 2)
  return ranks
}
async function prepare(store: SourceStore, spec: SortSpec, cooperative: Cooperative, budget: Budget, locale: string): Promise<PreparedSortKeys> {
  const r0 = spec.range.r0 + spec.range.headerRows, r1 = spec.range.r1 - spec.range.totalsRows
  const length = Math.max(0, r1 - r0 + 1), lanes: SortLane[] = []
  const collator = new Intl.Collator(locale, { numeric: true, sensitivity: spec.caseSensitive ? 'variant' : 'accent', caseFirst: 'upper' })
  let incomplete = false, unavailableCount = 0, firstAddress: { row: number; col: number } | undefined
  for (const key of spec.keys) {
    budget.charge(length * 9 + 512)
    budget.ledger.fixedOverheadBytes += 512
    budget.ledger.tagsBytes += length
    budget.ledger.numbersBytes += length * 8
    const lane: MutableLane = { tags: new Uint8Array(length), numbers: new Float64Array(length), strings: [], errors: [], stringRanks: new Uint32Array(0), errorRanks: new Uint32Array(0) }
    let stringIds: Map<string, number> | undefined, errorIds: Map<string, number> | undefined, mapBytes = 0
    function intern(raw: string, error: boolean): number {
      if (!lane.ids) { budget.charge(length * 4); budget.ledger.idsBytes += length * 4; lane.ids = new Uint32Array(length) }
      const map = error ? (errorIds ??= new Map()) : (stringIds ??= new Map())
      const dictionary = error ? lane.errors : lane.strings
      const existing = map.get(raw)
      if (existing !== undefined) return existing
      // Conservative app ledger: raw UTF-16 + 24B string header + 16B array
      // capacity/reference + 128B Map entry/capacity. Host overhead is measured separately.
      const retained = raw.length * 2 + 40
      budget.charge(retained + 128)
      budget.ledger.dictionaryBytes += retained
      mapBytes += 128
      budget.ledger.dictionaryMapPeakBytes = Math.max(budget.ledger.dictionaryMapPeakBytes, mapBytes)
      const id = dictionary.length
      dictionary.push(raw); map.set(raw, id)
      return id
    }
    for (let start = r0; start <= r1; start += CHUNK) {
      const end = Math.min(r1, start + CHUNK - 1)
      let cursor: string | undefined, pages = 0, cursorBytes = 0
      const cursors = new Set<string>()
      do {
        cooperative.check()
        const page = await store.queryColumn({ col: key.col, rows: { kind: 'intervals', intervals: [{ r0: start, r1: end }] }, limit: CHUNK, cursor }, { stamp: cooperative.stamp, signal: cooperative.options.signal ?? new AbortController().signal })
        cooperative.check()
        if (!sameStamp(page.stamp, cooperative.stamp)) throw sourceError('stale-source', 'Column page owner changed')
        if (page.error) throw new Halt({ kind: 'error', error: { ...page.error } })
        if (page.completeness !== 'complete') incomplete = true
        // A bounded source page can retain transient cell objects/raw strings even
        // when a repeated key does not enter the dictionary.
        let pageBytes = 256 * page.items.length
        budget.charge(pageBytes)
        for (const item of page.items) {
          if (item.kind === 'cell' || item.kind === 'unavailable') {
            const strings = [item.cell.value, item.cell.formula, item.cell.arrayRef, item.cell.formulaType, item.cell.sharedFormula?.ref]
            for (const raw of strings) if (typeof raw === 'string') { const bytes = raw.length * 2 + 24; pageBytes += bytes; budget.charge(bytes) }
          }
          if (cooperative.due()) await cooperative.pause('read')
        }
        budget.ledger.pagePeakBytes = Math.max(budget.ledger.pagePeakBytes, pageBytes)
        for (const item of page.items) {
          const first = 'row' in item ? item.row : item.r0, last = 'row' in item ? item.row : item.r1
          if (!Number.isSafeInteger(first) || !Number.isSafeInteger(last) || first < start || last > end || last < first) throw sourceError('page-error', 'Column entry outside requested chunk')
          for (let row = first; row <= last; row++) {
            const offset = row - r0
            if (lane.tags[offset]) throw sourceError('page-error', 'Overlapping column coverage')
            switch (item.kind) {
              case 'blank': lane.tags[offset] = BLANK; break
              case 'unknown': lane.tags[offset] = UNKNOWN; incomplete = true; break
              case 'error': throw new Halt({ kind: 'error', error: { code: item.code, retryable: item.retryable, message: 'Column key source failed' } })
              case 'unavailable': {
                lane.tags[offset] = UNAVAILABLE; unavailableCount++
                if (!firstAddress || row < firstAddress.row || row === firstAddress.row && key.col < firstAddress.col) firstAddress = { row, col: key.col }
                break
              }
              case 'cell': {
                const cell = item.cell, value = cell.value
                if (cell.row !== row || cell.col !== key.col || cell.id !== row * 16384 + key.col) throw sourceError('page-error', 'Cell identity does not match its key entry')
                if (cell.valueIsError) { lane.tags[offset] = ERROR; const id = intern(String(value), true); lane.ids![offset] = id }
                else if (value === null) lane.tags[offset] = BLANK
                else if (typeof value === 'number' && Number.isFinite(value)) { lane.tags[offset] = NUMBER; lane.numbers[offset] = value }
                else if (typeof value === 'boolean') { lane.tags[offset] = BOOLEAN; lane.numbers[offset] = value ? 1 : 0 }
                else if (typeof value === 'string') { lane.tags[offset] = STRING; const id = intern(value, false); lane.ids![offset] = id }
                else throw sourceError('page-error', 'Nonfinite or unsupported authoritative key')
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
          cursors.add(cursor)
        }
        if (cooperative.due()) await cooperative.pause('read')
      } while (cursor !== undefined)
      budget.release(cursorBytes)
      for (let row = start; row <= end; row++) { if (!lane.tags[row - r0]) incomplete = true; if (cooperative.due()) await cooperative.pause('read') }
    }
    lane.stringRanks = await rankDictionary(lane.strings, collator.compare, budget, cooperative)
    lane.errorRanks = await rankDictionary(lane.errors, (a, b) => a < b ? -1 : a > b ? 1 : 0, budget, cooperative)
    stringIds = undefined; errorIds = undefined
    budget.release(mapBytes)
    lanes.push(Object.freeze({ ...lane, strings: Object.freeze(lane.strings), errors: Object.freeze(lane.errors) }))
  }
  if (firstAddress) throw new Halt({ kind: 'refused', refusal: { reason: 'formula-value-unavailable', count: unavailableCount, firstAddress } })
  if (incomplete) throw new Halt({ kind: 'incomplete' })
  cooperative.check()
  return Object.freeze({ spec, stamp: cooperative.stamp, r0, r1, lanes: Object.freeze(lanes), ledger: budget.finish(), locale })
}
export async function prepareSortKeys(store: SourceStore, input: SortSpec, options: SortOptions = {}): Promise<PreparedSortResult> {
  try {
    options = Object.freeze({ ...options })
    const stamp = Object.freeze({ ...(options.stamp ?? store.stamp) }), spec = snapshotSpec(input)
    const cooperative = new Cooperative(store, stamp, options), budget = new Budget(options)
    cooperative.check()
    return { kind: 'complete', prepared: await prepare(store, spec, cooperative, budget, options.locale ?? 'en') }
  } catch (error) { return failure(error, options) }
}
function compareRows(prepared: PreparedSortKeys, a: number, b: number): number {
  const first = a - prepared.r0, second = b - prepared.r0
  for (let key = 0; key < prepared.lanes.length; key++) {
    const lane = prepared.lanes[key], ta = lane.tags[first], tb = lane.tags[second]
    let comparison = 0
    if (ta !== tb) {
      comparison = ta - tb
      if (ta >= ERROR || tb >= ERROR) return comparison
    } else if (ta === STRING) comparison = lane.stringRanks[lane.ids![first]] - lane.stringRanks[lane.ids![second]]
    else if (ta === ERROR) comparison = lane.errorRanks[lane.ids![first]] - lane.errorRanks[lane.ids![second]]
    else if (ta !== BLANK) comparison = lane.numbers[first] - lane.numbers[second]
    if (comparison) return ta === ERROR || ta === BLANK ? comparison : prepared.spec.keys[key].order === 'asc' ? comparison : -comparison
  }
  return a - b
}
/** Large bands own their buffers and construct/validate the inverse cooperatively.
 * S2.1 helpers synchronously copy/scan entire bands: keep them for bounded bands,
 * while this private builder preserves their frozen contract without shared edits. */
function ownedOrder(r0: number, r1: number, sourceAtOffset: Uint32Array, displayOffsetOfSource: Uint32Array): CompactRowOrder {
  const band = Object.freeze({ r0, r1, sourceAtOffset, displayOffsetOfSource })
  return Object.freeze({ kind: 'bands', bands: Object.freeze([band]), diagnostics: Object.freeze({ typedArrayBytes: sourceAtOffset.byteLength + displayOffsetOfSource.byteLength, bandCount: 1 }),
    displayToSource(row: number) { validateCoordinate(row, 0); return row >= r0 && row <= r1 ? sourceAtOffset[row - r0] : row },
    sourceToDisplay(row: number) { validateCoordinate(row, 0); return row >= r0 && row <= r1 ? r0 + displayOffsetOfSource[row - r0] : row },
  })
}
/** Pure draft only. S7.1/S3 must admit whole-row structural/outside-column effects
 * before a caller commits this object; this kernel cannot authorize those guards. */
export async function sortRowOrder(store: SourceStore, input: SortSpec | null, options: SortOptions = {}): Promise<SortResult> {
  try {
    options = Object.freeze({ ...options })
    const stamp = Object.freeze({ ...(options.stamp ?? store.stamp) }), spec = input === null ? null : snapshotSpec(input), locale = options.locale ?? 'en'
    const cooperative = new Cooperative(store, stamp, options), budget = new Budget(options)
    cooperative.check()
    let order: RowOrder = createIdentityOrder()
    if (spec !== null) {
      const prepared = await prepare(store, spec, cooperative, budget, locale), length = Math.max(0, prepared.r1 - prepared.r0 + 1)
      if (length) {
        const pinned = options.visibility ? hiddenSlotPins(options.visibility) : () => false
        budget.charge(length * 8)
        budget.ledger.indexBytes = length * 4; budget.ledger.scratchBytes = length * 4
        const index = new Uint32Array(length), scratch = new Uint32Array(length)
        let count = 0
        for (let row = prepared.r0; row <= prepared.r1; row++) { if (!pinned(row)) index[count++] = row; if (cooperative.due()) await cooperative.pause('candidates') }
        const sorted = await mergeIndexes(index, scratch, count, (a, b) => compareRows(prepared, a, b), cooperative, 'merge')
        if (length <= CHUNK) {
          // pinRowPermutation peak: candidates + rows + copied order + inverse;
          // createRowOrder then copies order/inverse again while that band lives.
          budget.charge(length * 20)
          budget.ledger.constructionCopiesBytes = length * 12
          const band = pinRowPermutation(prepared.r0, prepared.r1, pinned, sorted.subarray(0, count))
          order = createRowOrder([band])
          budget.release(length * 12)
        } else {
          budget.charge(length * 8)
          const rows = new Uint32Array(length), inverse = new Uint32Array(length)
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
            if (row < prepared.r0 || row > prepared.r1 || inverse[row - prepared.r0] !== MAX_ROWS || pinned(row) && row !== prepared.r0 + offset) throw sourceError('page-error', 'Sort permutation is not a source bijection')
            inverse[row - prepared.r0] = offset
            if (cooperative.due()) await cooperative.pause('inverse')
          }
          order = ownedOrder(prepared.r0, prepared.r1, rows, inverse)
        }
        budget.ledger.orderBytes = length * 4; budget.ledger.inverseBytes = length * 4
        budget.release(length * 8)
        // Prepared lanes are private to this operation and not retained by output.
        const keyBytes = budget.ledger.tagsBytes + budget.ledger.numbersBytes + budget.ledger.idsBytes + budget.ledger.ranksBytes + budget.ledger.dictionaryBytes
        budget.release(keyBytes + budget.ledger.fixedOverheadBytes)
      } else {
        budget.release(budget.ledger.retainedBytes)
      }
    }
    if (spec === null) budget.release(budget.ledger.fixedOverheadBytes)
    budget.charge(256)
    await cooperative.pause('publish')
    return { kind: 'complete', order, metadata: Object.freeze({ sort: spec, stamp, locale, structuralAdmission: 'caller-required' }), ledger: budget.finish() }
  } catch (error) { return failure(error, options) }
}
export type SortCommand = { kind: 'cycle'; col: number; shift?: boolean; range: FilterRange } | { kind: 'toggle-case' } | { kind: 'clear' }
export function reduceSortCommand(sort: SortSpec | null, command: SortCommand): { sort: SortSpec | null; notice?: { reason: 'too-many-sort-keys'; maxKeys: 8 } } {
  if (command.kind === 'clear') return { sort: null }
  if (command.kind === 'toggle-case') return { sort: sort ? snapshotSpec({ ...sort, caseSensitive: !sort.caseSensitive }) : null }
  const existing = sort?.keys.find(key => key.col === command.col), direction = existing?.order === 'asc' ? 'desc' : existing?.order === 'desc' ? null : 'asc'
  const keys = command.shift ? [...sort?.keys ?? []] : []
  const at = keys.findIndex(key => key.col === command.col)
  if (direction === null) { if (at >= 0) keys.splice(at, 1) }
  else if (at >= 0) keys[at] = { col: command.col, order: direction }
  else {
    if (keys.length === 8) return { sort, notice: { reason: 'too-many-sort-keys', maxKeys: 8 } }
    keys.push({ col: command.col, order: direction })
  }
  return { sort: keys.length ? snapshotSpec({ range: command.range, keys, caseSensitive: sort?.caseSensitive ?? false }) : null }
}
// TODO(S7.3): exact external text runs/merge for huge unique domains. U raw
// strings cost sum(2*UTF16 length + 40B) retained + 128U Map construction bytes,
// 12U rank/index/scratch bytes, plus row lanes and order workspace. Refuse when
// their aggregate construction peak exceeds the caller's remaining worker pool.
