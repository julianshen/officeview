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
  constructor() { super('Text overflow occupancy exceeds the document memory budget'); this.name = 'OverflowResourceError' }
}
interface Memo { key: string; snapshot?: WeakRef<readonly object[]>; input?: RowOverflowInput }
interface Slot { geometry: WeakRef<SheetGeometry>; rows: WeakRef<Map<number, Memo>>; row: number; memo: WeakRef<Memo>; entry: PoolEntry }
// Includes the input/closure, row-map slot, memo, weak references, LRU record,
// pool entry/callback/key and typed-array header, not just its backing buffer.
// Warmed Bun/JSC 1k/5k tiny-row probes measure about 1.7–1.9 KiB per
// retained entry including native map/finalizer bookkeeping. Round up to 4 KiB.
const ENTRY_BYTES = 4096
const MERGE_BYTES = 96
let nextCache = 0

export class OverflowCache {
  private readonly rows = new WeakMap<SheetGeometry, Map<number, Memo>>()
  private readonly slots = new Map<string, Slot>()
  private readonly finalized = new FinalizationRegistry<string>(key => this.drop(key))
  private readonly prefix = `overflow:${nextCache++}:`
  private nextRow = 0
  private held = 0
  private closed = false
  constructor(readonly pool: MemoryPool, readonly maxBytes = 2 * 1024 * 1024, readonly accounting?: OverflowAccounting) {}
  get bytes(): number { return this.held }

  get(geometry: SheetGeometry, row: number): RowOverflowInput {
    if (this.closed || this.pool.closed) throw new OverflowResourceError()
    const cells = geometry.cellsOf(row)
    const rows = this.rows.get(geometry)
    const hit = rows?.get(row)
    if (hit?.input && (hit.snapshot ? cells !== undefined && hit.snapshot.deref() === cells : cells === undefined)) {
      const slot = this.slots.get(hit.key)!
      slot.entry.touch(); this.slots.delete(hit.key); this.slots.set(hit.key, slot)
      return hit.input
    }
    if (hit) this.drop(hit.key)
    this.sweep()
    // The merge collector and filtered merge array can coexist. Reserve their
    // conservative capacities before collecting, as well as the full blocker
    // backing buffer (a subarray after deduplication still retains that buffer).
    const mergeCapacity = this.mergeCount(geometry, row)
    const retained = ENTRY_BYTES + (cells?.length ?? 0) * 2 + mergeCapacity * MERGE_BYTES
    const scratch = 128 + (cells?.length ?? 0) * 2 + mergeCapacity * 16
    const total = retained + scratch
    if (!this.reserve(total)) throw new OverflowResourceError()
    const key = `${this.prefix}${this.nextRow++}`
    const memo: Memo = { key, ...(cells ? { snapshot: new WeakRef(cells) } : {}) }
    const weakMemo = new WeakRef(memo)
    const onRelease = this.releaseCallback(key, weakMemo)
    let entry = this.pool.retain(key, { tier: 'rebuildable', bytes: total, onRelease })
    if (!entry) { this.uncharge(total); throw new OverflowResourceError() }
    const ownedRows = rows ?? new Map<number, Memo>()
    if (!rows) this.rows.set(geometry, ownedRows)
    this.slots.set(key, { geometry: new WeakRef(geometry), rows: new WeakRef(ownedRows), row, memo: weakMemo, entry })
    try {
      const metrics = geometry.metrics
      const colStart = (col: number) => metrics.cols.start(col)
      const blockers = new Uint16Array(cells?.length ?? 0)
      let count = 0
      for (const cell of cells ?? []) {
        if (cellBlocksOverflow(cell) && colStart(cell.col + 1) > colStart(cell.col)) blockers[count++] = cell.col
      }
      const filled = blockers.subarray(0, count)
      filled.sort()
      let unique = 0
      for (const col of filled) if (unique === 0 || col !== blockers[unique - 1]) blockers[unique++] = col
      const collected = mergeCapacity ? geometry.merges.collect({ r0: row, r1: row, c0: 0, c1: MAX_COLS - 1 }) : []
      const merges = collected.filter(merge => colStart(merge.maxCol + 1) > colStart(merge.minCol))
      memo.input = { blockers: blockers.subarray(0, unique), merges, colStart, lastCol: metrics.cols.logicalCount - 1 }
      entry = this.pool.retain(key, { tier: 'rebuildable', bytes: retained, onRelease })
      if (!entry) throw new OverflowResourceError()
      this.slots.get(key)!.entry = entry
      this.uncharge(scratch)
      ownedRows.set(row, memo)
      this.finalized.register(geometry, key, memo)
      return memo.input
    } catch (error) {
      this.drop(key)
      throw error
    }
  }

  /** Called on cache activity and document teardown; dead geometries release
   * charges even if the owner is still open. No row/input is rooted by the LRU. */
  sweep(): void {
    for (const [key, slot] of this.slots) if (!slot.geometry.deref() || !slot.memo.deref()) this.drop(key)
  }
  close(): void {
    if (this.closed) return
    this.closed = true
    for (const key of this.slots.keys()) this.drop(key)
  }
  private mergeCount(geometry: SheetGeometry, row: number): number {
    if (!isValidCoordinate(row, 0) || !geometry.merges.size) return 0
    // walkFrom with a rejecting predicate allocates no hit array. Account its
    // bounded traversal/result objects before counting exact row intersections.
    if (!this.reserve(128)) throw new OverflowResourceError()
    const charge = this.pool.charge()
    if (!charge.set(128)) { this.uncharge(128); throw new OverflowResourceError() }
    try {
      let count = 0
      geometry.merges.walkFrom({ r0: row, r1: row, c0: 0, c1: MAX_COLS - 1 }, 0, 1, { pred: () => { count++; return false } })
      return count
    } finally { charge.release(); this.uncharge(128) }
  }
  private reserve(bytes: number): boolean {
    if (bytes > this.maxBytes || !Number.isSafeInteger(bytes)) return false
    while (this.held + bytes > this.maxBytes && this.slots.size) this.drop(this.slots.keys().next().value!)
    if (this.held + bytes > this.maxBytes) return false
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
