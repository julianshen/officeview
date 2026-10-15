/** Request-local overflow proof. This private wire sidecar never certifies a
 * row-cache chunk or installs geometry outside the requested paint window. */
import type { AuthoritativeCell, CacheIdentity, Completeness, LogicalViewport, ViewportCell, ViewportResult } from './contracts'
import type { ProjectedSheetMetrics } from './projection'
import { cellBlocksOverflow } from './textOverflow'
import { FindError } from './find'
import { queryPayloadBytes } from './queryPayload'
import { sameIdentity } from './cacheIdentity'

export type { OverflowEnd, OverflowEntry, OverflowProof, OverflowViewportResult, OverflowCapture } from './overflowProof'
import type { OverflowEnd, OverflowEntry, OverflowProof, OverflowViewportResult, OverflowCapture, OverflowMerge } from './overflowProof'
interface Occupancy { row: number; col: number; blocking: boolean; start?: number; end?: number; left?: number; right?: number }
export class ViewportOverflowCollector implements OverflowCapture {
  private readonly occupancy: Occupancy[] = []
  private readonly runs: Array<{ r0: number; r1: number; c0: number; c1: number; complete: boolean }> = []
  private merges: readonly OverflowMerge[] = []
  private structureComplete = true
  private bytes = 256
  private outputBytes = 0
  private pageBytes = 0
  constructor(private readonly budget: number) {}
  private add(bytes: number): void {
    this.bytes += bytes
    this.check()
  }
  private check(): void {
    // The provider page, retained output and collector/index/proof coexist in
    // the same admitted 2Q owner. The largest page is intentionally counted
    // again inside outputBytes: a conservative peak bound while provider-page
    // storage and placed output may coexist. No fixed partition reduces caps.
    if (this.pageBytes + this.outputBytes + this.bytes > 2 * this.budget) throw new FindError('resource-limit', 'Viewport construction exceeds admitted workspace')
  }
  page(payload: unknown, records: number): void {
    const bytes = queryPayloadBytes(payload, this.budget)
    this.pageBytes = Math.max(this.pageBytes, bytes)
    // Payload ownership survives the page through placed cells; placements,
    // copied cell envelopes and readDetail's sets/sort share this upper bound.
    this.outputBytes += bytes + records * 384
    this.check()
  }
  structure(completeness: Completeness, merges: readonly OverflowMerge[]): void { this.structureComplete = completeness === 'complete'; this.add(merges.length * 224); this.merges = merges }
  cell(cell: AuthoritativeCell): void { this.add(96); this.occupancy.push({ row: cell.row, col: cell.col, blocking: cellBlocksOverflow(cell) }) }
  run(range: { r0: number; r1: number; c0: number; c1: number }, complete: boolean): void { this.add(160); this.runs.push({ ...range, complete }) }
  *finish(metrics: ProjectedSheetMetrics, cells: readonly ViewportCell[], owner: CacheIdentity, window: LogicalViewport, profile: 'D' | 'M', scope: 'frame' | 'tile'): Generator<number, OverflowProof> {
    const entries: OverflowEntry[] = [], pending: OverflowProof['pending'] = []
    // Admit the reference sort scratch before construction. Row descriptors are
    // sized after sorting/counting, before any Map/index/corridor allocation.
    this.add(128 + 16 * this.occupancy.length)
    const events = this.occupancy, scratch: Occupancy[] = new Array(events.length)
    for (let width = 1; width < events.length; width *= 2) {
      for (let from = 0; from < events.length; from += 2 * width) {
        const middle = Math.min(from + width, events.length), end = Math.min(from + 2 * width, events.length)
        let left = from, right = middle
        for (let at = from; at < end; at++) {
          const a = left < middle ? events[left] : undefined, b = right < end ? events[right] : undefined
          if (a && (!b || a.row < b.row || a.row === b.row && a.col <= b.col)) { scratch[at] = a; left++ }
          else { scratch[at] = b!; right++ }
          yield 4
        }
      }
      for (let at = 0; at < events.length; at++) { events[at] = scratch[at]; yield 1 }
    }
    let rowCount = 0, previousRow = -1
    for (const event of events) { if (event.row !== previousRow) { rowCount++; previousRow = event.row }; yield 1 }
    // Per row256B covers Map slot, descriptor and two array headers. Per cell
    // 64B covers four corridor fields and two16B index slots. Sort scratch can
    // still coexist: its16B/cell charge remains owned until finish drains.
    this.add(128 + 256 * rowCount + 64 * events.length)
    const rows = new Map<number, { positive: number[]; blockers: number[] }>()
    for (let at = 0; at < events.length; at++) {
      const event = events[at], start = metrics.cols.start(event.col), end = metrics.cols.start(event.col + 1)
      event.start = start; event.end = end; event.left = start; event.right = end
      let index = rows.get(event.row)
      if (!index) { index = { positive: [], blockers: [] }; rows.set(event.row, index) }
      if (end > start) { index.positive.push(at); if (event.blocking) index.blockers.push(at) }
      yield 2
    }
    // Contiguous known-empty corridors are propagated once in each direction.
    // Hidden evidence contributes no distance or blocker and is skipped.
    for (const index of rows.values()) {
      let previous: Occupancy | undefined
      for (const at of index.positive) {
        const event = events[at]
        if (event.end! > event.start!) {
          if (!event.blocking && previous && !previous.blocking && previous.end === event.start) event.left = previous.left
          previous = event
        }
        yield 2
      }
      previous = undefined
      for (let at = index.positive.length - 1; at >= 0; at--) {
        const event = events[index.positive[at]]
        if (event.end! > event.start!) {
          if (!event.blocking && previous && !previous.blocking && event.end === previous.start) event.right = previous.right
          previous = event
        }
        yield 2
      }
    }
    const lower = function* (from: number, to: number, col: number, indexes?: number[]): Generator<number, number> {
      while (from < to) {
        const middle = Math.floor((from + to) / 2), event = events[indexes ? indexes[middle] : middle]
        if (event.col < col) from = middle + 1
        else to = middle
        yield 1
      }
      return from
    }
    for (const item of cells) {
      if (item.kind !== 'cell' && item.kind !== 'unavailable') continue
      this.add(448)
      const { row, col } = item.source, box = item.bounds, index = rows.get(row)
      let from = 0, to = this.runs.length
      while (from < to) { const middle = Math.floor((from + to) / 2); if (this.runs[middle].r1 < row) from = middle + 1; else to = middle; yield 1 }
      const candidate = this.runs[from], run = candidate && candidate.r0 <= row ? candidate : undefined
      const h0 = run ? metrics.cols.start(run.c0) : box.x, h1 = run ? metrics.cols.start(run.c1 + 1) : box.x + box.width
      let x0 = h0, x1 = h1, leftEnd: OverflowEnd = 'paint-boundary', rightEnd: OverflowEnd = 'paint-boundary', merged = false
      // Structural rectangles remain intervals; never expand covered cells.
      for (const m of this.merges) {
        if (row >= m.range.minRow && row <= m.range.maxRow) {
          if (m.range.minCol <= col && col <= m.range.maxCol) merged = true
          else if (m.bounds.width > 0) {
            if (m.range.maxCol < col) x0 = Math.max(x0, m.bounds.x + m.bounds.width)
            if (m.range.minCol > col) x1 = Math.min(x1, m.bounds.x)
          }
        }
        yield 1
      }
      if (merged) { x0 = box.x; x1 = box.x + box.width; leftEnd = rightEnd = 'blocker' }
      else {
        if (index) {
          const at = yield* lower(0, index.blockers.length, col, index.blockers)
          if (at > 0) x0 = Math.max(x0, events[index.blockers[at - 1]].end!)
          const next = at < index.blockers.length && events[index.blockers[at]].col === col ? at + 1 : at
          if (next < index.blockers.length) x1 = Math.min(x1, events[index.blockers[next]].start!)
          yield 3
        }
        if (x0 > h0) leftEnd = 'blocker'
        if (x1 < h1) rightEnd = 'blocker'
        if (!run?.complete || !this.structureComplete) {
          const influence0 = x0, influence1 = x1
          let certified0 = box.x, certified1 = box.x + box.width
          if (this.structureComplete && index) {
            const at = yield* lower(0, index.positive.length, col, index.positive)
            // The origin may block; only the adjacent empty corridor extends it.
            const next = at < index.positive.length && events[index.positive[at]].col === col ? at + 1 : at
            const before = at > 0 ? events[index.positive[at - 1]] : undefined, after = next < index.positive.length ? events[index.positive[next]] : undefined
            if (before && !before.blocking && before.end === certified0) certified0 = before.left!
            if (after && !after.blocking && after.start === certified1) certified1 = after.right!
            yield 4
          }
          if (x0 < certified0) { x0 = certified0; leftEnd = 'unknown' }
          if (x1 > certified1) { x1 = certified1; rightEnd = 'unknown' }
          const p0 = window.box.x, p1 = p0 + window.box.width
          const influence = (from: number, to: number) => { const x = Math.max(p0, from), end = Math.min(p1, to); if (end > x) { this.add(256); pending.push({ bounds: { x, y: box.y, width: end - x, height: box.height }, error: false }) } }
          if (leftEnd === 'unknown') influence(influence0, x0)
          if (rightEnd === 'unknown') influence(x1, influence1)
          if (!this.structureComplete) influence(box.x, box.x + box.width)
        }
      }
      entries.push({ id: row * 16384 + col, row, col, left: Math.max(0, box.x - x0), right: Math.max(0, x1 - box.x - box.width), leftEnd, rightEnd, x0, x1, h0: Math.min(h0, box.x), h1: Math.max(h1, box.x + box.width), c0: run?.c0 ?? col, c1: run?.c1 ?? col })
      yield 1
    }
    return { version: 1, owner: { ...owner }, window: { ...window, box: { ...window.box } }, profile, scope, entries, pending }
  }
}

const finiteBox = (b: { x: number; y: number; width: number; height: number }) => [b.x, b.y, b.width, b.height].every(Number.isFinite) && b.width >= 0 && b.height >= 0
const coordinate = (row: number, col: number) => Number.isInteger(row) && row >= 0 && row < 1048576 && Number.isInteger(col) && col >= 0 && col < 16384
/** Validation allocates only inside the caller's joint reservation. */
export function validateOverflowProof(result: OverflowViewportResult, owner: CacheIdentity, window: LogicalViewport, profile: 'D' | 'M', scope: 'frame' | 'tile', columns: { start(col: number): number }): Map<number, OverflowEntry> {
  const proof = result.overflowProof
  const invalid = () => { throw new Error('Invalid or missing worker overflow evidence') }
  if (!proof || proof.version !== 1 || proof.profile !== profile || proof.scope !== scope || !Array.isArray(proof.entries) || !Array.isArray(proof.pending)) return invalid()
  if (!sameIdentity(proof.owner, owner)) return invalid()
  if (!proof.window || !finiteBox(proof.window.box) || !Number.isFinite(proof.window.zoom) || proof.window.zoom <= 0 || !Number.isFinite(proof.window.deviceScale) || proof.window.deviceScale <= 0 || proof.window.zoom !== window.zoom || proof.window.deviceScale !== window.deviceScale) return invalid()
  for (const field of ['x', 'y', 'width', 'height'] as const) if (proof.window.box[field] !== window.box[field]) return invalid()
  const byId = new Map<number, OverflowEntry>()
  for (const entry of proof.entries) {
    if (!coordinate(entry.row, entry.col) || entry.id !== entry.row * 16384 + entry.col || byId.has(entry.id) || ![entry.left, entry.right, entry.x0, entry.x1, entry.h0, entry.h1].every(Number.isFinite) || entry.left < 0 || entry.right < 0 || entry.x1 < entry.x0 || entry.x0 < entry.h0 || entry.x1 > entry.h1 || !coordinate(0, entry.c0) || !coordinate(0, entry.c1) || entry.c1 < entry.c0 || columns.start(entry.c0 + 1) < Math.max(0, window.box.x - 2048) || columns.start(entry.c1) > window.box.x + window.box.width + 2048 || !['blocker', 'paint-boundary', 'unknown'].includes(entry.leftEnd) || !['blocker', 'paint-boundary', 'unknown'].includes(entry.rightEnd)) return invalid()
    byId.set(entry.id, entry)
  }
  let count = 0
  if (result.content.kind === 'detail') for (const item of result.content.cells) {
    if (!coordinate(item.source.row, item.source.col) || !finiteBox(item.bounds)) return invalid()
    if (item.kind !== 'cell' && item.kind !== 'unavailable') { if (result.completeness === 'complete') return invalid(); continue }
    if (item.cell.row !== item.source.row || item.cell.col !== item.source.col || item.cell.id !== item.source.row * 16384 + item.source.col) return invalid()
    const entry = byId.get(item.source.row * 16384 + item.source.col)
    if (!entry || Math.abs(entry.x0 - (item.bounds.x - entry.left)) > 1e-7 || Math.abs(entry.x1 - (item.bounds.x + item.bounds.width + entry.right)) > 1e-7) return invalid()
    if (entry.h0 !== Math.min(columns.start(entry.c0), item.bounds.x) || entry.h1 !== Math.max(columns.start(entry.c1 + 1), item.bounds.x + item.bounds.width)) return invalid()
    if ((entry.leftEnd === 'unknown' && entry.x0 > window.box.x || entry.rightEnd === 'unknown' && entry.x1 < window.box.x + window.box.width) && !proof.pending.some(p => p.bounds.y < item.bounds.y + item.bounds.height && p.bounds.y + p.bounds.height > item.bounds.y)) return invalid()
    count++
  }
  for (const merge of result.merges ?? []) {
    const r = merge.range
    if (!coordinate(r.minRow, r.minCol) || !coordinate(r.maxRow, r.maxCol) || r.maxRow < r.minRow || r.maxCol < r.minCol || !finiteBox(merge.bounds)) return invalid()
  }
  if (count !== byId.size || proof.pending.some(p => !finiteBox(p.bounds) || typeof p.error !== 'boolean' || p.bounds.x < window.box.x || p.bounds.x + p.bounds.width > window.box.x + window.box.width) || result.completeness === 'complete' && proof.pending.length) return invalid()
  return byId
}

/** A bounded measurement stack is admitted before inspecting nested payload
 * shape. No traversal Map, Set, Object.values or payload-sized queue is built.
 * Aliases are counted repeatedly; depth/work limits reject cycles conservatively. */
export const VIEWPORT_MEASUREMENT_BYTES = 8192
export function measureViewportPayload(value: unknown, limit: number): number {
  let bytes = 0, nodes = 0
  const walk = (item: unknown, depth: number): void => {
    if (bytes > limit) return
    if (++nodes > 200000 || depth > 32) { bytes = limit + 1; return }
    if (typeof item === 'string') { bytes += 16 + 2 * item.length; return }
    if (typeof item !== 'object' || item === null) { bytes += 8; return }
    if (ArrayBuffer.isView(item)) { bytes += 24 + item.buffer.byteLength; return }
    if (Object.prototype.toString.call(item) === '[object ArrayBuffer]') { bytes += (item as ArrayBuffer).byteLength; return }
    if (Array.isArray(item)) {
      bytes += 24 + item.length * 16
      for (let i = 0; i < item.length && bytes <= limit; i++) walk(item[i], depth + 1)
      return
    }
    bytes += 24
    for (const field in item) {
      if (!Object.prototype.hasOwnProperty.call(item, field)) continue
      bytes += 16
      walk((item as Record<string, unknown>)[field], depth + 1)
      if (bytes > limit) break
    }
  }
  walk(value, 0)
  return bytes
}
/** Internal estimate, evaluated only under the admitted measurement owner. */
export function viewportPlanAdmission(reply: unknown, result: ViewportResult, styles: readonly import('../types').XlsxCellStyle[], adoption: number, rollback: number, columnMetadata: number, cap: number, oldPlanBytes = 0): { payload: number; planUpper: number; reserve: number; handoff: number; joint: number } {
  const payload = measureViewportPayload(reply, cap)
  if (payload > cap) return { payload, planUpper: 0, reserve: 0, handoff: 0, joint: 0 }
  const cells = result.content.kind === 'detail' ? result.content.cells : []
  const count = cells.length + (result.coarse?.length ?? 0) + (result.merges?.length ?? 0) + (result.rowLines?.length ?? 0) + (result.colLines?.length ?? 0)
  let styleBytes = 0
  // 128 direct-mapped keys use 1048 accounted bytes inside the existing
  // 8192-byte measurement owner. Collisions remeasure and overcharge safely;
  // document-sized style tables never allocate a proportional Set/bitset.
  const measured = new Float64Array(128).fill(NaN)
  for (const item of cells) {
    if (item.kind !== 'cell' && item.kind !== 'unavailable') continue
    const index = item.cell.styleIndex ?? 0
    if (Number.isInteger(index) && index >= 0) {
      const slot = index % measured.length
      if (measured[slot] === index) continue
      measured[slot] = index
    }
    const style = styles[index]
    if (style) styleBytes += measureViewportPayload(style, cap)
  }
  const planUpper = payload + count * 512 + styleBytes
  // A retired plan can remain reachable through rollback snapshots after its
  // named retained owner decreases. Keep that overlap in staging until commit.
  const retiredPlans = oldPlanBytes + planUpper > cap ? oldPlanBytes : 0
  const reserve = retiredPlans + 2 * payload + planUpper + 8192 + count * 512 + adoption + columnMetadata * 1024 + rollback
  const handoff = planUpper + adoption
  return { payload, planUpper, reserve, handoff, joint: VIEWPORT_MEASUREMENT_BYTES + reserve + handoff }
}
