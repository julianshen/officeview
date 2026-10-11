import type { CompactSourceRowSet, GridSelection, GridSelectionPart, SourceRowInterval, FilterRange, LogicalCellBox, SheetView, AxisMetrics } from './contracts'
import { MAX_ROWS, formatA1, validateCoordinate, validateSourceRect, type SourceCell, type NavigatorRequest } from './coordinates'
import type { ProjectedSheetMetrics } from './projection'
import { guardSort, guardFilter, type GuardOutcome, type FilterGuardRequest, type SortGuardOptions } from './operationGuard'
import type { StructuralIndex, PopulatedCellSource } from './structure'
import { clampSheetZoom } from './sheetZoom'

export class SourceSelection {
  constructor(readonly ranges: readonly GridSelectionPart[]) {}
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
  if (rows.kind === 'ref') throw new RangeError('Worker row refs require a membership resolver')
  if (rows.kind === 'bitset' && rows.space !== 'source') throw new RangeError('Selection requires source-space rows')
  if (row < rows.r0 || row > rows.r1) return false
  if (rows.kind === 'list') return rows.rows.includes(row)
  return Boolean(rows.words[Math.floor((row - rows.r0) / 32)] & (1 << ((row - rows.r0) % 32)))
}
function* iterateRows(rows: CompactSourceRowSet): IterableIterator<number> {
  if (rows.kind === 'ref') throw new RangeError('Worker row refs require a membership resolver')
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
      if (intervals.length > 32) return { kind: 'bitset', space: 'source', r0: first * 32, r1: last * 32 + 31, words: bits.slice(first, last + 1), count }
      value &= value - 1
    }
  }
  return { kind: 'intervals', intervals }
}
/** At most 32 interval objects, then a <=128 KiB source mask; never one object per row. */
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
  const a = projection.order.sourceToDisplay(anchor.row), b = projection.order.sourceToDisplay(target.row), lo = Math.min(a, b), hi = Math.max(a, b)
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
  const c0 = Math.min(anchor.col, target.col), c1 = Math.max(anchor.col, target.col)
  const cols = [...projection.cols.visibleBetween(projection.cols.start(c0), projection.cols.start(c1) + projection.cols.size(c1)).runs]
    .map(run => ({ c0: Math.max(c0, run.start), c1: Math.min(c1, run.end) }))
  return new SourceSelection([{ rows: rows.finish(), cols }])
}

function snapshot(active: SourceCell | null, anchor: SourceCell | null, ranges: readonly GridSelectionPart[]): GridSelection {
  return Object.freeze({ active: active && Object.freeze({ ...active }), anchor: anchor && Object.freeze({ ...anchor }), ranges: Object.freeze([...ranges]) })
}
export function emptyGridSelection(): GridSelection { return snapshot(null, null, []) }
function nonempty(rows: CompactSourceRowSet): boolean {
  return rows.kind === 'intervals' ? rows.intervals.length > 0 : rows.kind === 'list' ? rows.rows.length > 0 : rows.count > 0
}
function bitsOf(rows: CompactSourceRowSet): Uint32Array {
  const bits = new Uint32Array(MAX_ROWS / 32)
  if (rows.kind === 'intervals') for (const interval of rows.intervals) fillRows(bits, interval.r0, interval.r1)
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
  for (let word = 0; word < left.length; word++) left[word] &= difference ? ~right[word] : right[word]
  return compactBits(left)
}
function subtract(parts: readonly GridSelectionPart[], cut: GridSelectionPart): GridSelectionPart[] {
  let result = [...parts]
  for (const c of cut.cols) result = result.flatMap(part => part.cols.flatMap(col => {
    if (col.c1 < c.c0 || col.c0 > c.c1) return [{ rows: part.rows, cols: [col] }]
    const common = rowIntersection(part.rows, cut.rows, false)
    if (!nonempty(common)) return [{ rows: part.rows, cols: [col] }]
    const remainder = rowIntersection(part.rows, cut.rows, true), out: GridSelectionPart[] = []
    if (nonempty(remainder)) out.push({ rows: remainder, cols: [col] })
    if (col.c0 < c.c0) out.push({ rows: common, cols: [{ c0: col.c0, c1: c.c0 - 1 }] })
    if (col.c1 > c.c1) out.push({ rows: common, cols: [{ c0: c.c1 + 1, c1: col.c1 }] })
    return out
  }))
  return result
}
function togglePart(state: GridSelection, target: GridSelectionPart): readonly GridSelectionPart[] {
  let uncovered = [target]
  for (const part of state.ranges) uncovered = subtract(uncovered, part)
  const remaining = subtract(state.ranges, target)
  return uncovered.length ? [...remaining, target] : remaining
}
export function selectCell(state: GridSelection, projection: ProjectedSheetMetrics, at: SourceCell, options: { extend?: boolean; toggle?: boolean } = {}): GridSelection {
  validateCell(projection, at)
  if (!projection.sourceCellToLogical(at)) throw new RangeError('Hidden cells use the read-only inspector')
  if (options.extend) {
    const anchor = state.anchor && projection.sourceCellToLogical(state.anchor) ? state.anchor : state.active && projection.sourceCellToLogical(state.active) ? state.active : at
    return snapshot(at, anchor, extendDisplayRange(projection, anchor, at).ranges)
  }
  const part: GridSelectionPart = { rows: { kind: 'intervals', intervals: [{ r0: at.row, r1: at.row }] }, cols: [{ c0: at.col, c1: at.col }] }
  return snapshot(at, at, options.toggle ? togglePart(state, part) : [part])
}
export function selectHeader(state: GridSelection, projection: ProjectedSheetMetrics, kind: 'row' | 'column', index: number, toggle = false): GridSelection {
  const at = kind === 'row' ? { row: index, col: state.active?.col ?? 0 } : { row: state.active?.row ?? 0, col: index }
  validateCell(projection, at)
  const part: GridSelectionPart = kind === 'row'
    ? { rows: { kind: 'intervals', intervals: [{ r0: index, r1: index }] }, cols: [{ c0: 0, c1: projection.cols.logicalCount - 1 }] }
    : { rows: { kind: 'intervals', intervals: [{ r0: 0, r1: projection.rows.logicalCount - 1 }] }, cols: [{ c0: index, c1: index }] }
  // Header membership is full logical space (hidden members included); active must be visible.
  const active = projection.sourceCellToLogical(at) ? at : state.active
  return snapshot(active, active, toggle ? togglePart(state, part) : [part])
}
export function selectionOperationRange(selection: GridSelection, explicit?: FilterRange): { ok: true; range: FilterRange } | Extract<GuardOutcome, { ok: false }> {
  if (explicit) {
    try { validateSourceRect(explicit); return { ok: true, range: explicit } }
    catch (error) { return { ok: false, refusal: { reason: 'invalid-range', message: error instanceof Error ? error.message : 'Invalid operation range' } } }
  }
  const part = selection.ranges[0]
  if (selection.ranges.length === 1 && part.rows.kind === 'intervals' && part.rows.intervals.length === 1 && part.cols.length === 1) {
    const row = part.rows.intervals[0], col = part.cols[0]
    const range: FilterRange = { ...row, ...col, headerRows: 0, totalsRows: 0 }
    validateSourceRect(range)
    return { ok: true, range }
  }
  return { ok: false, refusal: { reason: 'ambiguous-selection', message: 'Choose an explicit operation range' } }
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
export interface NavigationResult extends SelectionTransition { kind: 'visible' | 'hidden'; target: SourceCell; hiddenReason?: string; logical?: LogicalCellBox }
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
function nearestColumn(axis: AxisMetrics, col: number): number | null {
  if (axis.size(col) > 0) return col
  const forward = axisCandidate(axis, col, 1), backward = axisCandidate(axis, col, -1)
  return forward !== null && (backward === null || forward - col <= col - backward) ? forward : backward
}
export function reconcileSelection(state: GridSelection, projection: ProjectedSheetMetrics, body: SourceRowInterval | null): SelectionTransition {
  if (!state.active) return { selection: state, announcement: '' }
  validateCell(projection, state.active)
  if (projection.sourceCellToLogical(state.active)) {
    if (state.anchor && !projection.sourceCellToLogical(state.anchor)) return { selection: Object.freeze({ active: state.active, anchor: state.active, ranges: state.ranges }), announcement: `Selection anchor moved to ${activeAnnouncement(projection, state.active)}` }
    return { selection: state, announcement: '' }
  }
  const slot = projection.order.sourceToDisplay(state.active.row)
  const eligible = body && state.active.row >= body.r0 && state.active.row <= body.r1 ? body : { r0: 0, r1: projection.rows.logicalCount - 1 }
  const forward = eligibleRow(projection, slot, 1, eligible), backward = eligibleRow(projection, slot, -1, eligible)
  const row = projection.rows.size(slot) > 0 ? state.active.row : forward !== null && (backward === null || projection.order.sourceToDisplay(forward) - slot <= slot - projection.order.sourceToDisplay(backward)) ? forward : backward
  const col = nearestColumn(projection.cols, state.active.col)
  if (row === null || col === null) return { selection: Object.freeze({ active: null, anchor: null, ranges: state.ranges }), announcement: 'No visible records.' }
  const at = { row, col }
  return { selection: Object.freeze({ active: Object.freeze(at), anchor: Object.freeze(at), ranges: state.ranges }), announcement: `Active moved to ${activeAnnouncement(projection, at)}` }
}
export function moveActive(state: GridSelection, projection: ProjectedSheetMetrics, direction: 'up' | 'down' | 'left' | 'right' | 'enter' | 'home' | 'end', body: SourceRowInterval | null, extend = false): GridSelection {
  const current = reconcileSelection(state, projection, body).selection
  const full = projection.rows.logicalCount ? { r0: 0, r1: projection.rows.logicalCount - 1 } : null
  let row = current.active?.row ?? eligibleRow(projection, 0, 1, body), col = current.active?.col ?? axisCandidate(projection.cols, 0, 1)
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
export function navigateTo(state: GridSelection, projection: ProjectedSheetMetrics, request: NavigatorRequest, body: SourceRowInterval | null): NavigationResult {
  const target = request.addressSpace === 'source' ? { ...request.at } : { row: projection.order.displayToSource(request.at.rowSlot), col: request.at.col }
  validateCell(projection, target)
  const logical = projection.sourceCellToLogical(target)
  if (logical) return { kind: 'visible', target, logical, selection: selectCell(state, projection, target), announcement: activeAnnouncement(projection, target) }
  const rowHidden = projection.rows.size(projection.order.sourceToDisplay(target.row)) === 0, colHidden = projection.cols.size(target.col) === 0
  const hiddenReason = rowHidden && colHidden ? 'Hidden row and column' : rowHidden ? 'Hidden row' : 'Hidden column'
  const current = reconcileSelection(state, projection, body)
  return { kind: 'hidden', target, selection: current.selection, hiddenReason, announcement: `${activeAnnouncement(projection, target)}; ${hiddenReason.toLowerCase()}. Read-only inspection.` }
}
export function activeAnnouncement(projection: ProjectedSheetMetrics, at: SourceCell, value?: string): string {
  validateCell(projection, at)
  const displayed = formatA1({ row: projection.order.sourceToDisplay(at.row), col: at.col }), original = formatA1(at)
  const bounded = value === undefined ? '' : `, value ${value.slice(0, 256)}`
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
  register(id: string, state: Omit<RetainedSheetState, 'selection' | 'zoom'> & Partial<Pick<RetainedSheetState, 'selection' | 'zoom'>>): void {
    if (!id || state.view.sheetId !== id || this.sheets.has(id)) throw new RangeError('Invalid or duplicate sheet registration')
    const selection = state.projection ? reconcileSelection(state.selection ?? state.view.selection, state.projection, state.body).selection : state.selection ?? state.view.selection
    this.sheets.set(id, Object.freeze({ ...state, selection, zoom: clampSheetZoom(state.zoom ?? 1) }))
    if (this.activeSheetId === null) this.activeSheetId = id
  }
  activate(id: string): RetainedSheetState { const sheet = this.get(id); this.activeSheetId = id; return sheet }
  get(id: string): RetainedSheetState { const state = this.sheets.get(id); if (!state) throw new RangeError('Unknown sheet'); return state }
  setSelection(id: string, selection: GridSelection): void { this.sheets.set(id, Object.freeze({ ...this.get(id), selection })) }
  setZoom(id: string, zoom: number): void { this.sheets.set(id, Object.freeze({ ...this.get(id), zoom: clampSheetZoom(zoom) })) }
  /** S5/S4 may release inactive geometry without losing source membership/view/zoom. */
  detachProjection(id: string): void { this.sheets.set(id, Object.freeze({ ...this.get(id), projection: null })) }
  updateProjection(id: string, projection: ProjectedSheetMetrics, body: SourceRowInterval | null, view = this.get(id).view): SelectionTransition {
    if (view.sheetId !== id) throw new RangeError('View belongs to another sheet')
    const state = this.get(id), result = reconcileSelection(state.selection, projection, body)
    this.sheets.set(id, Object.freeze({ ...state, projection, body, view, selection: result.selection }))
    return result
  }
  replaceDocument(docId: string): void { this.docId = docId; this.sheets.clear(); this.activeSheetId = null }
}
