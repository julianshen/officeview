import type { CompactSourceRowSet, GridSelection, GridSelectionPart, SourceRowInterval, FilterRange, LogicalCellBox, SheetView, AxisMetrics } from './contracts'
import { MAX_ROWS, formatA1, validateCoordinate, validateSourceRect, type SourceCell, type NavigatorRequest } from './coordinates'
import type { ProjectedSheetMetrics } from './projection'
import { guardSort, guardFilter, type GuardOutcome, type FilterGuardRequest, type SortGuardOptions } from './operationGuard'
import type { StructuralIndex, PopulatedCellSource } from './structure'
import { clampSheetZoom } from './sheetZoom'
import { retainActive, navigationTarget, hiddenReasonForCell, type ActiveNavigationContext, type HiddenTargetReason } from './scrollWindow'
import { createIntervalSet } from './visibility'

export const MAX_SELECTION_BYTES = 1024 * 1024
export const MAX_SELECTION_PARTS = 256
export const MAX_SELECTION_COL_INTERVALS = 1024
export class SelectionError extends Error {
  constructor(readonly code: 'unsupported' | 'resource-limit' | 'stale-source', message: string) { super(message); this.name = 'SelectionError' }
}
function opaqueRows(): never { throw new SelectionError('unsupported', 'Worker row refs require a membership resolver') }

export class SourceSelection {
  constructor(readonly ranges: readonly GridSelectionPart[]) { assertSelectionBound(ranges) }
  contains(at: SourceCell): boolean {
    validateCoordinate(at.row, at.col)
    return this.ranges.some(part => part.cols.some(c => at.col >= c.c0 && at.col <= c.c1) && rowContains(part.rows, at.row))
  }
  *sourceRows(): IterableIterator<number> {
    const rows = new RowBuilder()
    for (const part of this.ranges) rows.addSet(part.rows)
    yield* iterateRows(rows.finish())
  }
}
export function rowContains(rows: CompactSourceRowSet, row: number): boolean {
  if (rows.kind === 'intervals') return rows.intervals.some(r => row >= r.r0 && row <= r.r1)
  if (rows.kind === 'ref') return opaqueRows()
  if (rows.kind === 'bitset' && rows.space !== 'source') throw new RangeError('Selection requires source-space rows')
  if (row < rows.r0 || row > rows.r1) return false
  if (rows.kind === 'list') return rows.rows.includes(row)
  return Boolean(rows.words[Math.floor((row - rows.r0) / 32)] & (1 << ((row - rows.r0) % 32)))
}
function* iterateRows(rows: CompactSourceRowSet): IterableIterator<number> {
  if (rows.kind === 'ref') return opaqueRows()
  if (rows.kind === 'intervals') {
    for (const interval of rows.intervals) for (let row = interval.r0; row <= interval.r1; row++) yield row
  } else if (rows.kind === 'list') yield* rows.rows
  else for (let word = 0; word < rows.words.length; word++) {
    let bits = rows.words[word]
    while (bits) {
      const bit = 31 - Math.clz32(bits & -bits)
      yield rows.r0 + word * 32 + bit
      bits &= bits - 1
    }
  }
}
function popcount(word: number): number {
  word -= (word >>> 1) & 0x55555555
  word = (word & 0x33333333) + ((word >>> 2) & 0x33333333)
  return Math.imul((word + (word >>> 4)) & 0x0f0f0f0f, 0x01010101) >>> 24
}
function fillRows(bits: Uint32Array, r0: number, r1: number): void {
  for (let word = Math.floor(r0 / 32); word <= Math.floor(r1 / 32); word++) {
    const lo = Math.max(0, r0 - word * 32), hi = Math.min(31, r1 - word * 32)
    bits[word] |= (2 ** (hi + 1) - 1) - (2 ** lo - 1)
  }
}
function compactBits(bits: Uint32Array): CompactSourceRowSet {
  let first = 0, last = bits.length - 1, count = 0
  while (first < bits.length && !bits[first]) first++
  if (first === bits.length) return { kind: 'intervals', intervals: [] }
  while (!bits[last]) last--
  for (let word = first; word <= last; word++) count += popcount(bits[word])
  const low = first * 32 + (31 - Math.clz32(bits[first] & -bits[first])), high = last * 32 + 31 - Math.clz32(bits[last])
  if (count === high - low + 1) return { kind: 'intervals', intervals: [{ r0: low, r1: high }] }
  const intervals: SourceRowInterval[] = []
  for (let word = first; word <= last; word++) {
    let value = bits[word]
    while (value) {
      const row = word * 32 + 31 - Math.clz32(value & -value), previous = intervals[intervals.length - 1]
      if (previous && previous.r1 + 1 === row) previous.r1 = row
      else intervals.push({ r0: row, r1: row })
      if (intervals.length > 32) return { kind: 'bitset', space: 'source', r0: first * 32, r1: high, words: bits.slice(first, last + 1), count }
      value &= value - 1
    }
  }
  return { kind: 'intervals', intervals }
}
/** At most 32 intervals per row set, then a <=128 KiB mask. Snapshot admission
 * also bounds aggregate buffers, parts and column intervals. */
class RowBuilder {
  private intervals: SourceRowInterval[] = []
  private bits: Uint32Array | undefined
  add(r0: number, r1 = r0): void {
    if (this.bits) { fillRows(this.bits, r0, r1); return }
    this.intervals.push({ r0, r1 })
    if (this.intervals.length > 32) {
      this.bits = new Uint32Array(MAX_ROWS / 32)
      for (const interval of this.intervals) fillRows(this.bits, interval.r0, interval.r1)
      this.intervals = []
    }
  }
  addSet(rows: CompactSourceRowSet): void {
    if (rows.kind === 'bitset' && rows.r0 % 32 === 0) {
      if (rows.space !== 'source') throw new RangeError('Selection requires source-space rows')
      if (!this.bits) {
        this.bits = new Uint32Array(MAX_ROWS / 32)
        for (const interval of this.intervals) fillRows(this.bits, interval.r0, interval.r1)
        this.intervals = []
      }
      const start = rows.r0 / 32
      for (let i = 0; i < rows.words.length; i++) this.bits[start + i] |= rows.words[i]
      return
    }
    if (rows.kind === 'intervals') for (const interval of rows.intervals) this.add(interval.r0, interval.r1)
    else for (const row of iterateRows(rows)) this.add(row)
  }
  finish(): CompactSourceRowSet {
    if (this.bits) return compactBits(this.bits)
    const intervals: SourceRowInterval[] = []
    for (const next of this.intervals.sort((a, b) => a.r0 - b.r0)) {
      const previous = intervals[intervals.length - 1]
      if (previous && next.r0 <= previous.r1 + 1) previous.r1 = Math.max(previous.r1, next.r1)
      else intervals.push({ ...next })
    }
    return { kind: 'intervals', intervals }
  }
}
function validateCell(projection: ProjectedSheetMetrics, at: SourceCell): void {
  validateCoordinate(at.row, at.col)
  if (at.row >= projection.rows.logicalCount || at.col >= projection.cols.logicalCount) throw new RangeError('Cell outside projection capacity')
}
/** Both endpoints are SOURCE cells; display slots exist only inside the projection. */
export function extendDisplayRange(projection: ProjectedSheetMetrics, anchor: SourceCell, target: SourceCell): SourceSelection {
  validateCell(projection, anchor)
  validateCell(projection, target)
  if (!projection.sourceCellToLogical(anchor) || !projection.sourceCellToLogical(target)) throw new RangeError('Display extension requires visible endpoints')
  return new SourceSelection([{ rows: displayRowsBetween(projection, anchor.row, target.row), cols: displayColsBetween(projection, anchor.col, target.col) }])
}
function displayRowsBetween(projection: ProjectedSheetMetrics, anchor: number, target: number): CompactSourceRowSet {
  const a = projection.order.sourceToDisplay(anchor), b = projection.order.sourceToDisplay(target), lo = Math.min(a, b), hi = Math.max(a, b)
  const rows = new RowBuilder()
  for (const run of projection.rows.visibleBetween(projection.rows.start(lo), projection.rows.start(hi) + projection.rows.size(hi)).runs) {
    let next = Math.max(lo, run.start), end = Math.min(hi, run.end)
    for (const band of projection.order.bands) {
      if (band.r1 < next || band.r0 > end) continue
      if (next < band.r0) rows.add(next, band.r0 - 1)
      const first = Math.max(next, band.r0), last = Math.min(end, band.r1)
      if (first === band.r0 && last === band.r1) rows.add(band.r0, band.r1)
      else for (let slot = first; slot <= last; slot++) rows.add(projection.order.displayToSource(slot))
      next = last + 1
    }
    if (next <= end) rows.add(next, end)
  }
  return rows.finish()
}
function displayColsBetween(projection: ProjectedSheetMetrics, anchor: number, target: number): GridSelectionPart['cols'] {
  const c0 = Math.min(anchor, target), c1 = Math.max(anchor, target)
  return [...projection.cols.visibleBetween(projection.cols.start(c0), projection.cols.start(c1) + projection.cols.size(c1)).runs]
    .map(run => ({ c0: Math.max(c0, run.start), c1: Math.min(c1, run.end) }))
}

interface SelectionGesture { kind: 'cells' | 'row' | 'column'; anchor: SourceCell; prefix: readonly GridSelectionPart[] }
const gestures = new WeakMap<GridSelection, SelectionGesture>()
function retainedRanges(state: GridSelection): readonly GridSelectionPart[] { return [...state.ranges, ...(gestures.get(state)?.prefix ?? [])] }
function gestureSnapshot(active: SourceCell | null, anchor: SourceCell | null, ranges: readonly GridSelectionPart[], gesture: SelectionGesture): GridSelection {
  const state = snapshot(active, anchor, ranges), prefix = normalizeParts(gesture.prefix)
  assertSelectionBound([...state.ranges, ...prefix])
  gestures.set(state, { ...gesture, anchor: { ...gesture.anchor }, prefix: Object.freeze(prefix) })
  return state
}
function copyGesture(before: GridSelection, after: GridSelection): GridSelection {
  const gesture = gestures.get(before)
  if (gesture && after.anchor) gestures.set(after, { ...gesture, anchor: after.anchor })
  return after
}

function snapshot(active: SourceCell | null, anchor: SourceCell | null, ranges: readonly GridSelectionPart[]): GridSelection {
  const normalized = normalizeParts(ranges)
  assertSelectionBound(normalized)
  return Object.freeze({ active: active && Object.freeze({ ...active }), anchor: anchor && Object.freeze({ ...anchor }), ranges: Object.freeze(normalized) })
}
export function emptyGridSelection(): GridSelection { return snapshot(null, null, []) }
function nonempty(rows: CompactSourceRowSet): boolean {
  return rows.kind === 'intervals' ? rows.intervals.length > 0 : rows.kind === 'list' ? rows.rows.length > 0 : rows.count > 0
}
function bitsOf(rows: CompactSourceRowSet): Uint32Array {
  const bits = new Uint32Array(MAX_ROWS / 32)
  if (rows.kind === 'intervals') for (const interval of rows.intervals) fillRows(bits, interval.r0, interval.r1)
  else if (rows.kind === 'bitset' && rows.r0 % 32 === 0) bits.set(rows.words, rows.r0 / 32)
  else for (const row of iterateRows(rows)) fillRows(bits, row, row)
  return bits
}
function rowIntersection(a: CompactSourceRowSet, b: CompactSourceRowSet, difference: boolean): CompactSourceRowSet {
  if (a.kind === 'intervals' && b.kind === 'intervals') {
    const out = new RowBuilder()
    for (const interval of a.intervals) {
      let next = interval.r0
      for (const cut of b.intervals) {
        if (cut.r1 < next) continue
        if (cut.r0 > interval.r1) break
        if (difference && next < cut.r0) out.add(next, Math.min(interval.r1, cut.r0 - 1))
        if (!difference) out.add(Math.max(next, cut.r0), Math.min(interval.r1, cut.r1))
        next = Math.max(next, cut.r1 + 1)
        if (next > interval.r1) break
      }
      if (difference && next <= interval.r1) out.add(next, interval.r1)
    }
    return out.finish()
  }
  const left = bitsOf(a), right = bitsOf(b)
  let unchanged = true
  for (let word = 0; word < left.length; word++) {
    const next = left[word] & (difference ? ~right[word] : right[word])
    if (next !== left[word]) unchanged = false
    left[word] = next
  }
  if (unchanged) return a
  return compactBits(left)
}
function mergeCols(cols: GridSelectionPart['cols']): GridSelectionPart['cols'] {
  const out: { c0: number; c1: number }[] = []
  for (const col of [...cols].sort((a, b) => a.c0 - b.c0)) {
    const previous = out[out.length - 1]
    if (previous && col.c0 <= previous.c1 + 1) previous.c1 = Math.max(previous.c1, col.c1)
    else out.push({ ...col })
  }
  return out
}
function subtract(parts: readonly GridSelectionPart[], cut: GridSelectionPart): GridSelectionPart[] {
  return parts.flatMap(part => {
    const cols: { c0: number; c1: number }[] = []
    let overlaps = false
    for (const col of part.cols) {
      let next = col.c0
      for (const c of cut.cols) {
        if (c.c1 < next) continue
        if (c.c0 > col.c1) break
        overlaps = true
        if (next < c.c0) cols.push({ c0: next, c1: c.c0 - 1 })
        next = Math.max(next, c.c1 + 1)
        if (next > col.c1) break
      }
      if (next <= col.c1) cols.push({ c0: next, c1: col.c1 })
    }
    if (!overlaps) return [part]
    const common = rowIntersection(part.rows, cut.rows, false)
    if (!nonempty(common)) return [part]
    const remainder = rowIntersection(part.rows, cut.rows, true), out: GridSelectionPart[] = []
    if (nonempty(remainder)) out.push({ rows: remainder, cols: part.cols })
    if (cols.length) out.push({ rows: common, cols })
    return out
  })
}
function rowsKey(rows: CompactSourceRowSet): string {
  if (rows.kind === 'ref') return opaqueRows()
  if (rows.kind === 'intervals') return `i:${rows.intervals.map(r => `${r.r0}:${r.r1}`).join(',')}`
  let hash = 2166136261
  for (const value of rows.kind === 'list' ? rows.rows : rows.words) hash = Math.imul(hash ^ value, 16777619)
  return `${rows.kind}:${rows.r0}:${rows.r1}:${hash >>> 0}`
}
function sameRows(a: CompactSourceRowSet, b: CompactSourceRowSet): boolean {
  if (a === b) return true
  if (a.kind !== b.kind) return false
  if (a.kind === 'intervals' && b.kind === 'intervals') return a.intervals.length === b.intervals.length && a.intervals.every((r, i) => r.r0 === b.intervals[i].r0 && r.r1 === b.intervals[i].r1)
  if (a.kind === 'bitset' && b.kind === 'bitset') return a.r0 === b.r0 && a.r1 === b.r1 && a.count === b.count && a.words.length === b.words.length && a.words.every((word, i) => word === b.words[i])
  if (a.kind === 'list' && b.kind === 'list') return a.r0 === b.r0 && a.r1 === b.r1 && a.rows.length === b.rows.length && a.rows.every((row, i) => row === b.rows[i])
  return opaqueRows()
}
function normalizeParts(input: readonly GridSelectionPart[]): GridSelectionPart[] {
  let parts = [...input]
  while (true) {
    const buckets = new Map<string, GridSelectionPart[]>(), keys = new WeakMap<object, string>()
    for (const part of parts) {
      const key = keys.get(part.rows) ?? rowsKey(part.rows)
      keys.set(part.rows, key)
      const bucket = buckets.get(key) ?? []
      const same = bucket.find(p => sameRows(p.rows, part.rows))
      if (same) same.cols = mergeCols([...same.cols, ...part.cols])
      else bucket.push({ rows: part.rows, cols: mergeCols(part.cols) })
      buckets.set(key, bucket)
    }
    const columns = new Map<string, GridSelectionPart>()
    for (const part of [...buckets.values()].flat()) {
      const key = part.cols.map(c => `${c.c0}:${c.c1}`).join(',')
      const same = columns.get(key)
      if (same) { const rows = new RowBuilder(); rows.addSet(same.rows); rows.addSet(part.rows); same.rows = rows.finish() }
      else columns.set(key, part)
    }
    const next = [...columns.values()]
    if (next.length === parts.length) return next
    parts = next
  }
}
/** Exact typed bytes plus conservative metadata estimates; shared row buffers
 * are counted once across the supplied ranges (including multiple sheets). */
export function selectionBytes(ranges: readonly GridSelectionPart[]): number {
  const rows = new Set<object>(), buffers = new Set<ArrayBufferLike>()
  let bytes = 192
  for (const part of ranges) {
    bytes += 72 + part.cols.length * 40
    if (rows.has(part.rows)) continue
    rows.add(part.rows)
    bytes += 64
    if (part.rows.kind === 'intervals') bytes += part.rows.intervals.length * 40
    else if (part.rows.kind === 'ref') bytes += 128
    else {
      const buffer = part.rows.kind === 'bitset' ? part.rows.words.buffer : part.rows.rows.buffer
      if (!buffers.has(buffer)) { buffers.add(buffer); bytes += buffer.byteLength }
    }
  }
  return bytes
}
function assertSelectionBound(ranges: readonly GridSelectionPart[]): void {
  if (ranges.length > MAX_SELECTION_PARTS || ranges.reduce((count, part) => count + part.cols.length, 0) > MAX_SELECTION_COL_INTERVALS || selectionBytes(ranges) > MAX_SELECTION_BYTES) throw new SelectionError('resource-limit', 'Selection exceeds its aggregate budget; choose a simpler selection')
}
function togglePart(state: GridSelection, target: GridSelectionPart): readonly GridSelectionPart[] {
  let uncovered = [target]
  for (const part of state.ranges) uncovered = subtract(uncovered, part)
  const remaining = subtract(state.ranges, target)
  return uncovered.length ? [...remaining, target] : remaining
}
function extendedPart(projection: ProjectedSheetMetrics, gesture: SelectionGesture, target: SourceCell): GridSelectionPart {
  if (gesture.kind === 'row') return { rows: displayRowsBetween(projection, gesture.anchor.row, target.row), cols: [{ c0: 0, c1: projection.cols.logicalCount - 1 }] }
  if (gesture.kind === 'column') return { rows: { kind: 'intervals', intervals: [{ r0: 0, r1: projection.rows.logicalCount - 1 }] }, cols: displayColsBetween(projection, gesture.anchor.col, target.col) }
  return extendDisplayRange(projection, gesture.anchor, target).ranges[0]
}
function extendGesture(state: GridSelection, projection: ProjectedSheetMetrics, target: SourceCell, kind?: 'row' | 'column'): GridSelection {
  const previous = gestures.get(state)
  let anchor = previous?.anchor ?? state.anchor ?? state.active ?? target
  if (!previous || previous.kind === 'cells') {
    if (!projection.sourceCellToLogical(anchor)) anchor = state.active && projection.sourceCellToLogical(state.active) ? state.active : target
  }
  const gesture: SelectionGesture = { kind: kind ?? previous?.kind ?? 'cells', anchor, prefix: previous?.prefix ?? state.ranges.slice(0, -1) }
  const part = extendedPart(projection, gesture, target)
  return gestureSnapshot(target, anchor, [...subtract(gesture.prefix, part), part], gesture)
}
export function selectCell(state: GridSelection, projection: ProjectedSheetMetrics, at: SourceCell, options: { extend?: boolean; toggle?: boolean } = {}): GridSelection {
  validateCell(projection, at)
  if (!projection.sourceCellToLogical(at)) throw new RangeError('Hidden cells use the read-only inspector')
  if (options.extend) {
    return extendGesture(state, projection, at)
  }
  const part: GridSelectionPart = { rows: { kind: 'intervals', intervals: [{ r0: at.row, r1: at.row }] }, cols: [{ c0: at.col, c1: at.col }] }
  const prefix = options.toggle ? subtract(state.ranges, part) : []
  return gestureSnapshot(at, at, options.toggle ? togglePart(state, part) : [part], { kind: 'cells', anchor: at, prefix })
}
export function selectHeader(state: GridSelection, projection: ProjectedSheetMetrics, kind: 'row' | 'column', index: number, toggle = false, extend = false): GridSelection {
  const row = kind === 'row' ? index : state.active?.row ?? projection.visibleRankToSource(0) ?? 0
  const col = kind === 'column' ? index : state.active && projection.cols.size(state.active.col) > 0 ? state.active.col : axisCandidate(projection.cols, 0, 1) ?? 0
  const at = { row, col }
  validateCell(projection, at)
  if (extend && gestures.get(state)?.kind === kind && projection.sourceCellToLogical(at)) return extendGesture(state, projection, at, kind)
  const part: GridSelectionPart = kind === 'row'
    ? { rows: { kind: 'intervals', intervals: [{ r0: index, r1: index }] }, cols: [{ c0: 0, c1: projection.cols.logicalCount - 1 }] }
    : { rows: { kind: 'intervals', intervals: [{ r0: 0, r1: projection.rows.logicalCount - 1 }] }, cols: [{ c0: index, c1: index }] }
  // Header membership is full logical space (hidden members included); active must be visible.
  const active = projection.sourceCellToLogical(at) ? at : state.active && projection.sourceCellToLogical(state.active) ? state.active : null
  const prefix = toggle ? subtract(state.ranges, part) : []
  return gestureSnapshot(active, active, toggle ? togglePart(state, part) : [part], { kind, anchor: at, prefix })
}
export function selectionOperationRange(selection: GridSelection, explicit?: FilterRange): { ok: true; range: FilterRange } | Extract<GuardOutcome, { ok: false }> {
  if (explicit) {
    try {
      validateSourceRect(explicit)
      const active = selection.active
      if (!active || active.row < explicit.r0 || active.row > explicit.r1 || active.col < explicit.c0 || active.col > explicit.c1) return { ok: false, refusal: { reason: 'ambiguous-selection', message: 'The resolved operation range must contain the active source cell' } }
      return { ok: true, range: explicit }
    }
    catch (error) { return { ok: false, refusal: { reason: 'invalid-range', message: error instanceof Error ? error.message : 'Invalid operation range' } } }
  }
  return { ok: false, refusal: { reason: 'ambiguous-selection', message: 'Resolve the operation range through S7 table/current-region detection' } }
}
export async function guardSelectionSort(selection: GridSelection, structure: StructuralIndex, occupancy: PopulatedCellSource, explicit?: FilterRange, options?: SortGuardOptions): Promise<GuardOutcome> {
  const resolved = selectionOperationRange(selection, explicit)
  return resolved.ok ? guardSort(resolved.range, structure, occupancy, options) : resolved
}
export async function guardSelectionFilter(selection: GridSelection, request: FilterGuardRequest): Promise<GuardOutcome> {
  const resolved = selectionOperationRange(selection, request.range)
  return resolved.ok ? guardFilter({ ...request, range: resolved.range }) : resolved
}

export interface SelectionTransition { selection: GridSelection; announcement: string }
export type SelectionNavigationView = Pick<SheetView, 'visibility' | 'baseHidden' | 'filter'>
export interface SelectionNavigationResult extends SelectionTransition { kind: 'visible' | 'hidden'; target: SourceCell; hiddenReason?: HiddenTargetReason; logical?: LogicalCellBox }
function lowerVisibleRank(projection: ProjectedSheetMetrics, slot: number): number {
  let low = 0, high = projection.rows.visibleBetween(0, projection.rows.totalCss()).count
  while (low < high) {
    const mid = Math.floor((low + high) / 2)
    if (projection.visibleRankToDisplay(mid)! < slot) low = mid + 1
    else high = mid
  }
  return low
}
function bodyBounds(projection: ProjectedSheetMetrics, body: SourceRowInterval): { lo: number; hi: number } {
  validateCoordinate(body.r0, 0)
  validateCoordinate(body.r1, 0)
  if (body.r1 < body.r0 || body.r1 >= projection.rows.logicalCount) throw new RangeError('Invalid record body')
  let lo = body.r0, hi = body.r1
  for (const band of projection.order.bands) if (band.r1 >= body.r0 && band.r0 <= body.r1) {
    lo = Math.min(lo, band.r0)
    hi = Math.max(hi, band.r1)
  }
  return { lo, hi }
}
function eligibleRow(projection: ProjectedSheetMetrics, slot: number, direction: -1 | 1, body: SourceRowInterval | null): number | null {
  if (!body) return null
  const { lo, hi } = bodyBounds(projection, body), first = lowerVisibleRank(projection, lo), end = lowerVisibleRank(projection, hi + 1)
  let rank = direction > 0 ? Math.max(first, lowerVisibleRank(projection, slot)) : Math.min(end - 1, lowerVisibleRank(projection, slot + 1) - 1)
  for (; rank >= first && rank < end; rank += direction) {
    const row = projection.visibleRankToSource(rank)!
    if (row >= body.r0 && row <= body.r1) return row
  }
  return null
}
function axisCandidate(axis: AxisMetrics, index: number, direction: -1 | 1): number | null {
  if (!axis.logicalCount || !axis.totalCss()) return null
  if (direction > 0) {
    if (index >= axis.logicalCount) return null
    const start = Math.max(0, index)
    for (const run of axis.visibleBetween(axis.start(start), axis.totalCss()).runs) return Math.max(start, run.start)
  } else {
    if (index < 0) return null
    const end = Math.min(axis.logicalCount - 1, index)
    let last: number | null = null
    for (const run of axis.visibleBetween(0, axis.start(end) + axis.size(end)).runs) last = Math.min(end, run.end)
    return last
  }
  return null
}
function activeContext(state: GridSelection, projection: ProjectedSheetMetrics, view?: SelectionNavigationView): ActiveNavigationContext {
  const select = (rank: number) => rank < 0 ? null : projection.visibleRankToDisplay(rank)
  return { metrics: projection, view: { selection: state, ...(view ?? {
    filter: null, baseHidden: createIntervalSet([]), visibility: {
      visible: (slot: number) => projection.rows.size(slot) > 0,
      rank: (end: number) => lowerVisibleRank(projection, end), select,
      next: (slot: number, direction: -1 | 1) => select(lowerVisibleRank(projection, slot + (direction > 0 ? 1 : 0)) - (direction < 0 ? 1 : 0)),
    },
  }) } }
}
export function reconcileSelection(state: GridSelection, projection: ProjectedSheetMetrics, _body: SourceRowInterval | null, view?: SelectionNavigationView): SelectionTransition {
  if (!state.active) return { selection: state, announcement: '' }
  const at = retainActive(activeContext(state, projection, view), state.active)
  if (at && at.row === state.active.row && at.col === state.active.col) {
    if (state.anchor && !projection.sourceCellToLogical(state.anchor)) return { selection: copyGesture(state, Object.freeze({ active: state.active, anchor: state.active, ranges: state.ranges })), announcement: `Selection anchor moved to ${activeAnnouncement(projection, state.active)}` }
    return { selection: state, announcement: '' }
  }
  if (!at) return { selection: Object.freeze({ active: null, anchor: null, ranges: state.ranges }), announcement: 'No visible records.' }
  return { selection: copyGesture(state, Object.freeze({ active: Object.freeze(at), anchor: Object.freeze(at), ranges: state.ranges })), announcement: `Active moved to ${activeAnnouncement(projection, at)}` }
}
export function moveActive(state: GridSelection, projection: ProjectedSheetMetrics, direction: 'up' | 'down' | 'left' | 'right' | 'enter' | 'home' | 'end', body: SourceRowInterval | null, extend = false, view?: SelectionNavigationView): GridSelection {
  const current = reconcileSelection(state, projection, body, view).selection
  const full = projection.rows.logicalCount ? { r0: 0, r1: projection.rows.logicalCount - 1 } : null
  let row = current.active?.row ?? eligibleRow(projection, 0, 1, body ?? full), col = current.active?.col ?? axisCandidate(projection.cols, 0, 1)
  if (direction === 'home' || direction === 'end') {
    const forward = direction === 'home'
    row = eligibleRow(projection, forward ? 0 : projection.rows.logicalCount - 1, forward ? 1 : -1, full)
    col = axisCandidate(projection.cols, forward ? 0 : projection.cols.logicalCount - 1, forward ? 1 : -1)
  } else if (current.active) {
    const slot = projection.order.sourceToDisplay(current.active.row)
    if (direction === 'up' || direction === 'down' || direction === 'enter') row = eligibleRow(projection, slot + (direction === 'up' ? -1 : 1), direction === 'up' ? -1 : 1, full) ?? row
    else col = axisCandidate(projection.cols, current.active.col + (direction === 'left' ? -1 : 1), direction === 'left' ? -1 : 1) ?? col
  }
  return row === null || col === null ? current : selectCell(current, projection, { row, col }, { extend })
}
export function navigateSelection(state: GridSelection, projection: ProjectedSheetMetrics, request: NavigatorRequest, body: SourceRowInterval | null, view?: SelectionNavigationView): SelectionNavigationResult {
  const target = navigationTarget(projection, request)
  const logical = projection.sourceCellToLogical(target)
  if (logical) return { kind: 'visible', target, logical, selection: selectCell(state, projection, target), announcement: activeAnnouncement(projection, target) }
  const hiddenReason = hiddenReasonForCell(activeContext(state, projection, view), target)
  const current = reconcileSelection(state, projection, body, view)
  return { kind: 'hidden', target, selection: current.selection, hiddenReason, announcement: `${activeAnnouncement(projection, target)}; ${hiddenReason.replace('-', ' ')}. Read-only inspection.` }
}
/** Compatibility alias; the selection-specific API is navigateSelection. */
export { navigateSelection as navigateTo }
export function boundedSelectionText(text: string): string {
  if (text.length <= 256) return text
  const unit = text.charCodeAt(255)
  return text.slice(0, unit >= 0xd800 && unit <= 0xdbff ? 255 : 256)
}
export function activeAnnouncement(projection: ProjectedSheetMetrics, at: SourceCell, value?: string): string {
  validateCell(projection, at)
  const displayed = formatA1({ row: projection.order.sourceToDisplay(at.row), col: at.col }), original = formatA1(at)
  const bounded = value === undefined ? '' : `, value ${boundedSelectionText(value)}`
  return `${displayed}${bounded}${displayed === original ? '' : `; original ${original}`}`
}
export function recordCounts(projection: ProjectedSheetMetrics, body: SourceRowInterval | null): { visible: number; total: number } {
  if (!body) return { visible: 0, total: 0 }
  const { lo, hi } = bodyBounds(projection, body), first = lowerVisibleRank(projection, lo), end = lowerVisibleRank(projection, hi + 1)
  let visible = end - first
  if (lo !== body.r0 || hi !== body.r1) {
    visible = 0
    for (let rank = first; rank < end; rank++) {
      const row = projection.visibleRankToSource(rank)!
      if (row >= body.r0 && row <= body.r1) visible++
    }
  }
  return { visible, total: body.r1 - body.r0 + 1 }
}
/** Call with the resolved table/current-region/USED range, never worksheet capacity. */
export function recordBody(range: FilterRange | null): SourceRowInterval | null {
  if (!range) return null
  validateSourceRect(range)
  if ((range.headerRows !== 0 && range.headerRows !== 1) || (range.totalsRows !== 0 && range.totalsRows !== 1)) throw new RangeError('Invalid protected record rows')
  const r0 = range.r0 + range.headerRows, r1 = range.r1 - range.totalsRows
  return r0 > r1 ? null : { r0, r1 }
}
export interface RetainedSheetState { view: SheetView; projection: ProjectedSheetMetrics | null; body: SourceRowInterval | null; selection: GridSelection; zoom: number }
export class SheetNavigationSession {
  activeSheetId: string | null = null
  private readonly sheets = new Map<string, RetainedSheetState>()
  constructor(public docId: string) {}
  private admitSelection(id: string, selection: GridSelection): void {
    assertSelectionBound(selection.ranges)
    const selections = [...this.sheets].filter(([sheetId]) => sheetId !== id).map(([, state]) => state.selection)
    selections.push(selection)
    if (selectionBytes(selections.flatMap(state => [...retainedRanges(state)])) + (selections.length - 1) * 192 > MAX_SELECTION_BYTES) throw new SelectionError('resource-limit', 'Retained selections exceed the document selection budget')
  }
  register(id: string, state: Omit<RetainedSheetState, 'selection' | 'zoom'> & Partial<Pick<RetainedSheetState, 'selection' | 'zoom'>>): void {
    if (!id || state.view.sheetId !== id || this.sheets.has(id)) throw new RangeError('Invalid or duplicate sheet registration')
    const selection = state.projection ? reconcileSelection(state.selection ?? state.view.selection, state.projection, state.body, state.view).selection : state.selection ?? state.view.selection
    this.admitSelection(id, selection)
    this.sheets.set(id, Object.freeze({ ...state, view: Object.freeze({ ...state.view, selection }), selection, zoom: clampSheetZoom(state.zoom ?? 1) }))
    if (this.activeSheetId === null) this.activeSheetId = id
  }
  activate(id: string): RetainedSheetState { const sheet = this.get(id); this.activeSheetId = id; return sheet }
  get(id: string): RetainedSheetState { const state = this.sheets.get(id); if (!state) throw new RangeError('Unknown sheet'); return state }
  setSelection(id: string, selection: GridSelection): void { const state = this.get(id); this.admitSelection(id, selection); this.sheets.set(id, Object.freeze({ ...state, view: Object.freeze({ ...state.view, selection }), selection })) }
  setZoom(id: string, zoom: number): void { this.sheets.set(id, Object.freeze({ ...this.get(id), zoom: clampSheetZoom(zoom) })) }
  /** S5/S4 may release inactive geometry without losing source membership/view/zoom. */
  detachProjection(id: string): void { this.sheets.set(id, Object.freeze({ ...this.get(id), projection: null })) }
  updateProjection(id: string, projection: ProjectedSheetMetrics, body: SourceRowInterval | null, view = this.get(id).view): SelectionTransition {
    if (view.sheetId !== id) throw new RangeError('View belongs to another sheet')
    const state = this.get(id)
    if (!Number.isSafeInteger(view.revision) || view.revision < state.view.revision) throw new SelectionError('stale-source', 'Stale sheet view revision; retry with the current view')
    const result = reconcileSelection(state.selection, projection, body, view)
    this.admitSelection(id, result.selection)
    this.sheets.set(id, Object.freeze({ ...state, projection, body, view: Object.freeze({ ...view, selection: result.selection }), selection: result.selection }))
    return result
  }
  replaceDocument(docId: string): void { this.docId = docId; this.sheets.clear(); this.activeSheetId = null }
}
