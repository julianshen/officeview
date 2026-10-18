/// <reference lib="es2021.weakref" />
/** Document-owned occupancy: the LRU and pool hold weak owners only. */
import { MAX_COLS, isValidCoordinate } from './coordinates'
import type { MemoryPool, PoolEntry } from './tileCache'
import type { SheetGeometry } from './viewport'
import { cellBlocksOverflow, type RowOverflowInput } from './textOverflow'

/** Additional shared-document accounting. Reserve/release take byte deltas,
 * including construction scratch, before the corresponding allocations. */
export interface OverflowAccounting {
  reserve(bytes: number): boolean
  release(bytes: number): void
}
export class OverflowResourceError extends Error {
  constructor(reason: 'budget' | 'closed' = 'budget') {
    super(reason === 'closed' ? 'Text overflow occupancy owner is closed' : 'Text overflow occupancy exceeds the document memory budget')
    this.name = 'OverflowResourceError'
  }
}
interface Memo { key: string; snapshot?: WeakRef<readonly object[]>; input?: RowOverflowInput }
interface Slot { geometry: WeakRef<SheetGeometry>; rows: WeakRef<Map<number, Memo>>; row: number; memo: WeakRef<Memo>; entry: PoolEntry }
// Includes the input/closure, row-map slot, memo, weak references, LRU record,
// pool entry/callback/key and typed-array header, not just its backing buffer.
// Warmed Bun/JSC 1k/5k tiny-row probes measure about 1.7–1.9 KiB per
// retained entry including native map/finalizer bookkeeping. Round up to 4 KiB.
const ENTRY_BYTES = 4096
const MERGE_BYTES = 4
let nextCache = 0

export class OverflowCache {
  private readonly rows = new WeakMap<SheetGeometry, Map<number, Memo>>()
  private readonly slots = new Map<string, Slot>()
  private readonly finalized = new FinalizationRegistry<string>(key => this.drop(key))
  private readonly prefix = `overflow:${nextCache++}:`
  private nextRow = 0
  private held = 0
  private isClosed = false
  private readers = 0
  constructor(readonly pool: MemoryPool, readonly maxBytes = 2 * 1024 * 1024, readonly accounting?: OverflowAccounting) {}
  get closed(): boolean { return this.isClosed }
  get bytes(): number { return this.held }

  get(geometry: SheetGeometry, row: number): RowOverflowInput {
    if (this.isClosed || this.pool.closed) throw this.refusal()
    this.readers++
    try { return this.read(geometry, row) }
    finally { if (--this.readers === 0 && this.isClosed) this.finishClose() }
  }

  private read(geometry: SheetGeometry, row: number): RowOverflowInput {
    const cells = geometry.cellsOf(row)
    if (this.isClosed || this.pool.closed) throw this.refusal()
    const rows = this.rows.get(geometry)
    const hit = rows?.get(row)
    if (hit?.input && (hit.snapshot ? cells !== undefined && hit.snapshot.deref() === cells : cells === undefined)) {
      const slot = this.slots.get(hit.key)!
      slot.entry.touch(); this.slots.delete(hit.key); this.slots.set(hit.key, slot)
      return hit.input
    }
    if (hit) this.drop(hit.key)
    this.sweep()
    // Collected merge references coexist with the compact endpoint buffer.
    // Reserve both before collecting, as well as the full blocker
    // backing buffer (a subarray after deduplication still retains that buffer).
    const mergeCapacity = this.mergeCount(geometry, row)
    if (this.isClosed || this.pool.closed) throw this.refusal()
    const blockerCapacity = Math.min(cells?.length ?? 0, MAX_COLS)
    const dedupBytes = (cells?.length ?? 0) > MAX_COLS ? MAX_COLS : 0
    const retained = ENTRY_BYTES + blockerCapacity * 2 + mergeCapacity * MERGE_BYTES
    const scratch = 128 + blockerCapacity * 2 + mergeCapacity * 16 + dedupBytes
    const total = retained + scratch
    if (!this.reserve(total, retained)) throw this.refusal()
    if (this.isClosed) { this.uncharge(total); throw this.refusal() }
    const key = `${this.prefix}${this.nextRow++}`
    const memo: Memo = { key, ...(cells ? { snapshot: new WeakRef(cells) } : {}) }
    const weakMemo = new WeakRef(memo)
    const onRelease = this.releaseCallback(key, weakMemo)
    let entry = this.pool.retain(key, { tier: 'rebuildable', bytes: total, onRelease })
    if (!entry) { this.uncharge(total); throw this.refusal() }
    const unpin = entry.pin()
    const ownedRows = rows ?? new Map<number, Memo>()
    this.rows.set(geometry, ownedRows)
    this.slots.set(key, { geometry: new WeakRef(geometry), rows: new WeakRef(ownedRows), row, memo: weakMemo, entry })
    try {
      const metrics = geometry.metrics
      const colStart = (col: number) => metrics.cols.start(col)
      const blockers = new Uint16Array(blockerCapacity)
      const seen = dedupBytes ? new Uint8Array(dedupBytes) : undefined
      let count = 0
      for (const cell of cells ?? []) {
        if (!cellBlocksOverflow(cell) || colStart(cell.col + 1) <= colStart(cell.col)) continue
        if (seen?.[cell.col]) continue
        if (seen) seen[cell.col] = 1
        blockers[count++] = cell.col
      }
      const filled = blockers.subarray(0, count)
      filled.sort()
      let unique = 0
      for (const col of filled) if (unique === 0 || col !== blockers[unique - 1]) blockers[unique++] = col
      const collected = mergeCapacity ? geometry.merges.collect({ r0: row, r1: row, c0: 0, c1: MAX_COLS - 1 }) : []
      const mergeColumns = new Uint16Array(mergeCapacity * 2)
      let mergeCount = 0
      for (const merge of collected) {
        if (colStart(merge.maxCol + 1) <= colStart(merge.minCol)) continue
        mergeColumns[mergeCount++] = merge.minCol
        mergeColumns[mergeCount++] = merge.maxCol
      }
      if (this.isClosed || this.pool.closed) throw this.refusal()
      memo.input = { blockers: blockers.subarray(0, unique), merges: mergeColumns.subarray(0, mergeCount), colStart, lastCol: metrics.cols.logicalCount - 1 }
      entry = this.pool.retain(key, { tier: 'rebuildable', bytes: retained, onRelease })
      if (!entry) throw this.refusal()
      this.slots.get(key)!.entry = entry
      this.uncharge(scratch)
      ownedRows.set(row, memo)
      this.finalized.register(geometry, key, memo)
      return memo.input
    } catch (error) {
      this.drop(key)
      throw error
    } finally { unpin() }
  }

  dropGeometry(geometry: SheetGeometry): void {
    const rows = this.rows.get(geometry)
    if (rows) for (const memo of rows.values()) this.drop(memo.key)
    this.rows.delete(geometry)
  }

  /** Called on cache activity and document teardown; dead geometries release
   * charges even if the owner is still open. No row/input is rooted by the LRU. */
  sweep(): void {
    for (const [key, slot] of this.slots) if (!slot.geometry.deref() || !slot.memo.deref()) this.drop(key)
  }
  close(): void {
    if (this.isClosed) return
    this.isClosed = true
    if (!this.readers) this.finishClose()
  }
  private finishClose(): void {
    for (const key of this.slots.keys()) this.drop(key)
  }
  private mergeCount(geometry: SheetGeometry, row: number): number {
    if (!isValidCoordinate(row, 0) || !geometry.merges.size) return 0
    // walkFrom with a rejecting predicate allocates no hit array. Account its
    // bounded traversal/result objects before counting exact row intersections.
    if (!this.reserve(ENTRY_BYTES, 0)) throw this.refusal()
    const charge = this.pool.charge()
    if (!charge.set(ENTRY_BYTES)) { this.uncharge(ENTRY_BYTES); throw this.refusal() }
    try {
      let count = 0
      geometry.merges.walkFrom({ r0: row, r1: row, c0: 0, c1: MAX_COLS - 1 }, 0, 1, { pred: () => { count++; return false } })
      return count
    } finally { charge.release(); this.uncharge(ENTRY_BYTES) }
  }
  private refusal(): OverflowResourceError { return new OverflowResourceError(this.isClosed || this.pool.closed ? 'closed' : 'budget') }
  private reserve(bytes: number, retainedBytes: number): boolean {
    if (retainedBytes > this.maxBytes || !Number.isSafeInteger(bytes)) return false
    while (this.held + retainedBytes > this.maxBytes && this.slots.size) this.drop(this.slots.keys().next().value!)
    if (this.held + retainedBytes > this.maxBytes) return false
    while (this.accounting?.reserve(bytes) === false) {
      if (!this.slots.size) return false
      this.drop(this.slots.keys().next().value!)
    }
    this.held += bytes
    return true
  }
  private uncharge(bytes: number): void { this.held -= bytes; this.accounting?.release(bytes) }
  private released(key: string, memo: WeakRef<Memo>): void {
    const slot = this.slots.get(key)
    if (!slot) return
    const owned = memo.deref()
    if (owned) { this.finalized.unregister(owned); owned.input = undefined }
    const rows = slot.rows.deref()
    rows?.delete(slot.row)
    const geometry = slot.geometry.deref()
    if (geometry && !rows?.size) this.rows.delete(geometry)
    this.slots.delete(key)
    this.uncharge(slot.entry.bytes)
  }
  private drop(key: string): void {
    const slot = this.slots.get(key)
    if (!slot) return
    slot.entry.release()
    this.released(key, slot.memo)
  }
  // Create callbacks outside get()'s allocation scope: only weak memo ownership
  // and an opaque key can reach the pool, never cells, input or metrics.
  private releaseCallback(key: string, memo: WeakRef<Memo>): () => void {
    return () => this.released(key, memo)
  }
}

/** Private geometry binding; source/wire contracts stay unchanged. */
const bindings = new WeakMap<SheetGeometry, OverflowCache>()
export function bindOverflowCache(geometry: SheetGeometry, cache: OverflowCache): void {
  const previous = bindings.get(geometry)
  if (previous === cache) return
  previous?.dropGeometry(geometry)
  bindings.set(geometry, cache)
}
export function boundOverflowCache(geometry: SheetGeometry): OverflowCache | undefined { return bindings.get(geometry) }
