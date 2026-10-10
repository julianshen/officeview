import type { FilterSpec, IntervalSet, RowOrder, RowVisibility, SheetView, SourceRowInterval } from './contracts'
import { MAX_ROWS, validateCoordinate, validateSourceRect, type DisplaySlot, type SourceRow } from './coordinates'

/** Ordinal among visible rows, distinct from a worksheet display slot. */
export type VisibleRank = number

const WORDS_PER_BLOCK = 32
function popcount(word: number): number {
  word -= (word >>> 1) & 0x55555555
  word = (word & 0x33333333) + ((word >>> 2) & 0x33333333)
  return (((word + (word >>> 4)) & 0x0f0f0f0f) * 0x01010101) >>> 24
}
function checkCount(count: number): void {
  if (!Number.isSafeInteger(count) || Object.is(count, -0) || count < 0 || count > MAX_ROWS) throw new RangeError('Invalid row count')
}
function checkSlot(slot: number, count: number): void {
  validateCoordinate(slot, 0)
  if (slot >= count) throw new RangeError('Display slot outside mask')
}
function setBit(words: Uint32Array, slot: number, value: boolean): void {
  const word = Math.floor(slot / 32), bit = 1 << (slot % 32)
  if (value) words[word] |= bit
  else words[word] &= ~bit
}
function clearRange(words: Uint32Array, r0: number, r1: number): void {
  for (let word = Math.floor(r0 / 32); word <= Math.floor(r1 / 32); word++) {
    const first = Math.max(0, r0 - word * 32), last = Math.min(31, r1 - word * 32)
    const mask = (2 ** (last + 1) - 1) - (2 ** first - 1)
    words[word] &= ~mask
  }
}
function visibleWords(count: number): Uint32Array {
  checkCount(count)
  const words = new Uint32Array(Math.ceil(count / 32)).fill(0xffffffff)
  if (count % 32) words[words.length - 1] = 2 ** (count % 32) - 1
  return words
}
/** Two-level counts: binary-search blocks, then inspect at most 32 words.
 * Only mask-word operations use bitwise arithmetic; cell IDs never do (§3). */
export class CompactRowVisibility implements RowVisibility {
  readonly words: Uint32Array
  readonly summaries: Uint32Array
  constructor(readonly logicalCount: number, words: Uint32Array) {
    checkCount(logicalCount)
    if (!(words instanceof Uint32Array) || words.length !== Math.ceil(logicalCount / 32)) throw new RangeError('Invalid visibility mask length')
    this.words = words.slice()
    if (logicalCount % 32 && this.words.length) this.words[this.words.length - 1] &= 2 ** (logicalCount % 32) - 1
    this.summaries = new Uint32Array(Math.ceil(words.length / WORDS_PER_BLOCK) + 1)
    let sum = 0
    for (let word = 0; word < this.words.length; word++) {
      if (word % WORDS_PER_BLOCK === 0) this.summaries[Math.floor(word / WORDS_PER_BLOCK)] = sum
      sum += popcount(this.words[word])
    }
    this.summaries[this.summaries.length - 1] = sum
  }
  visible(displaySlot: DisplaySlot): boolean {
    checkSlot(displaySlot, this.logicalCount)
    return Boolean((this.words[Math.floor(displaySlot / 32)] >>> (displaySlot % 32)) & 1)
  }
  rank(endExclusive: DisplaySlot): VisibleRank {
    if (!Number.isSafeInteger(endExclusive) || endExclusive < 0 || endExclusive > this.logicalCount) throw new RangeError('Invalid rank endpoint')
    const wordEnd = Math.floor(endExclusive / 32), block = Math.floor(wordEnd / WORDS_PER_BLOCK)
    let sum = this.summaries[block]
    for (let word = block * WORDS_PER_BLOCK; word < wordEnd; word++) sum += popcount(this.words[word])
    if (endExclusive % 32) sum += popcount(this.words[wordEnd] & (2 ** (endExclusive % 32) - 1))
    return sum
  }
  select(visibleOrdinal: VisibleRank): DisplaySlot | null {
    if (!Number.isSafeInteger(visibleOrdinal) || visibleOrdinal < 0 || visibleOrdinal >= this.summaries[this.summaries.length - 1]) return null
    let lo = 0, hi = this.summaries.length - 1
    while (lo < hi) {
      const mid = Math.floor((lo + hi) / 2)
      if (this.summaries[mid + 1] <= visibleOrdinal) lo = mid + 1
      else hi = mid
    }
    let remaining = visibleOrdinal - this.summaries[lo]
    for (let word = lo * WORDS_PER_BLOCK; word < Math.min(this.words.length, (lo + 1) * WORDS_PER_BLOCK); word++) {
      const count = popcount(this.words[word])
      if (remaining >= count) { remaining -= count; continue }
      const bits = this.words[word]
      for (let bit = 0; bit < 32; bit++) {
        if ((bits >>> bit) & 1) {
          if (!remaining) return word * 32 + bit
          remaining--
        }
      }
    }
    return null
  }
  /** Strict next/previous, never the current slot. */
  next(displaySlot: DisplaySlot, direction: -1 | 1): DisplaySlot | null {
    checkSlot(displaySlot, this.logicalCount)
    if (direction !== 1 && direction !== -1) throw new RangeError('Invalid traversal direction')
    return this.select(direction === 1 ? this.rank(displaySlot + 1) : this.rank(displaySlot) - 1)
  }
}
/** Buffers are generation-owned and must not be mutated after publication. */
export function createRowVisibility(logicalCount = MAX_ROWS, predicate?: (displaySlot: DisplaySlot) => boolean): CompactRowVisibility {
  const words = visibleWords(logicalCount)
  if (predicate) for (let slot = 0; slot < logicalCount; slot++) setBit(words, slot, predicate(slot))
  return new CompactRowVisibility(logicalCount, words)
}
export function createIntervalSet(inputs: readonly SourceRowInterval[]): IntervalSet {
  const intervals: SourceRowInterval[] = []
  const sorted = inputs.map(interval => {
    validateCoordinate(interval.r0, 0); validateCoordinate(interval.r1, 0)
    if (interval.r0 > interval.r1) throw new RangeError('Reversed hidden interval')
    return { ...interval }
  }).sort((a, b) => a.r0 - b.r0)
  for (const interval of sorted) {
    const previous = intervals[intervals.length - 1]
    if (previous && previous.r1 + 1 >= interval.r0) previous.r1 = Math.max(previous.r1, interval.r1)
    else intervals.push(interval)
  }
  const snapshot = Object.freeze(intervals.map(interval => Object.freeze(interval)))
  return Object.freeze({ intervals: snapshot, contains(sourceRow: SourceRow) {
    validateCoordinate(sourceRow, 0)
    let lo = 0, hi = snapshot.length
    while (lo < hi) {
      const mid = Math.floor((lo + hi) / 2)
      if (snapshot[mid].r1 < sourceRow) lo = mid + 1
      else hi = mid
    }
    return lo < snapshot.length && snapshot[lo].r0 <= sourceRow
  } })
}
function snapshotWords(hidden: IntervalSet, count: number): Uint32Array {
  const words = visibleWords(count)
  for (const interval of hidden.intervals) {
    if (interval.r0 < count) clearRange(words, interval.r0, Math.min(count - 1, interval.r1))
  }
  return words
}
export function visibilityFromHidden(hidden: IntervalSet, count = MAX_ROWS): CompactRowVisibility {
  return new CompactRowVisibility(count, snapshotWords(hidden, count))
}
export function normalizeFilter(filter: FilterSpec | null): FilterSpec | null {
  if (!filter || !filter.columns.size) return null
  validateSourceRect(filter.range)
  if (![0, 1].includes(filter.range.headerRows) || ![0, 1].includes(filter.range.totalsRows)) throw new RangeError('Invalid protected row counts')
  return filter
}
/** One source-space predicate for all consumers; rule evaluation belongs to S8 (§5.2). */
export function isVisibleSource(sourceRow: SourceRow, view: Pick<SheetView, 'baseHidden' | 'filter'>, matches: (row: SourceRow, filter: FilterSpec) => boolean): boolean {
  validateCoordinate(sourceRow, 0)
  const filter = normalizeFilter(view.filter)
  if (filter && sourceRow >= filter.range.r0 + filter.range.headerRows && sourceRow <= filter.range.r1 - filter.range.totalsRows) return matches(sourceRow, filter)
  return !view.baseHidden.contains(sourceRow)
}
/** Rebuild from the immutable snapshot, never the previous mask. Visit only
 * affected source body rows/permutation bands, plus compact snapshot words (§5.2). */
export function projectVisibility(baseHidden: IntervalSet, input: FilterSpec | null, order: RowOrder, matches: (row: SourceRow, filter: FilterSpec) => boolean, count = MAX_ROWS): { filter: FilterSpec | null; visibility: CompactRowVisibility } {
  const filter = normalizeFilter(input), words = snapshotWords(baseHidden, count)
  for (const band of order.bands) {
    if (band.r1 >= count) throw new RangeError('Order exceeds visibility axis')
    for (let offset = 0; offset < band.sourceAtOffset.length; offset++) setBit(words, band.r0 + offset, !baseHidden.contains(band.sourceAtOffset[offset]))
  }
  if (filter) {
    if (filter.range.r1 >= count) throw new RangeError('Filter exceeds visibility axis')
    for (let sourceRow = filter.range.r0 + filter.range.headerRows; sourceRow <= filter.range.r1 - filter.range.totalsRows; sourceRow++) {
      setBit(words, order.sourceToDisplay(sourceRow), isVisibleSource(sourceRow, { baseHidden, filter }, matches))
    }
  }
  return { filter, visibility: new CompactRowVisibility(count, words) }
}
