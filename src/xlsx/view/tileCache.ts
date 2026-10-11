/**
 * Resident memory pools and the bitmap tile cache (spec §§10.3, 12.4; plan S4.3).
 *
 * MemoryPool is the single app-accounted ledger of one pool for one document:
 * every retained cache and every transient buffer of that document charges it,
 * across all sheets and outstanding operations, so the budget is aggregate and
 * never a per-sheet allowance. Retained entries live in three LRU tiers that
 * evict in a fixed order once the pool crosses its high watermark:
 *
 * 1. `rebuildable` — prepared columns, structural handles, layout records and
 *    tiles: recomputed from resident or persisted data on demand.
 * 2. `source` — decoded source pages a persisted store can reload (S1.2 paged
 *    adapters register them; the interim object source has none).
 * 3. `view` — materialized inactive sheet views (order/mask buffers). Their
 *    compact logical state survives eviction and the owner rebuilds the buffers
 *    under the same identity; the active sheet's views are protected.
 *
 * Eviction calls the owner's `onRelease('evicted')` exactly once; an owner that
 * releases its own entry gets no callback. Closing the pool releases every entry
 * with `'closed'` and refuses all later retains/charges, so nothing revives.
 *
 * TileCache keeps Canvas2D base tiles (or another backend's surfaces) in the
 * bitmap pool and closes each surface (ImageBitmap.close) exactly once when it
 * is evicted, replaced, invalidated or disposed. S5.2 extends it with painting.
 */
import type { CacheIdentity, ProfileId } from './contracts'

const MiB = 1024 * 1024
/** §12.4 app-controlled resident budgets per pool. */
export const WORKER_BUDGETS: Readonly<Record<ProfileId, number>> = Object.freeze({ D: 512 * MiB, M: 256 * MiB })
export const MAIN_THREAD_BUDGETS: Readonly<Record<ProfileId, number>> = Object.freeze({ D: 64 * MiB, M: 32 * MiB })
export const BITMAP_BUDGETS: Readonly<Record<ProfileId, number>> = Object.freeze({ D: 128 * MiB, M: 64 * MiB })
/** Reference tile edge in device pixels; one RGBA8 tile is exactly 1 MiB. */
export const TILE_DEVICE_PX = 512
export const RGBA_BYTES = 4
/** Evict at 7/8 of the hard limit, down to 3/4, so admissions rarely wait. */
export const DEFAULT_HIGH_WATERMARK = 0.875
export const DEFAULT_LOW_WATERMARK = 0.75

export type PoolTier = 'rebuildable' | 'source' | 'view'
export const EVICTION_ORDER: readonly PoolTier[] = Object.freeze(['rebuildable', 'source', 'view'])
export type ReleaseReason = 'evicted' | 'closed'
export interface PoolEntryInit {
  tier: PoolTier
  bytes: number
  /** Owner scope (a sheet id): protect() shields an owner's views. */
  owner?: string
  /** False for entries that cannot be rebuilt (e.g. arbitrary committed drafts). */
  evictable?: boolean
  /** Called once when the pool evicts or closes the entry, never on release(). */
  onRelease?: (reason: ReleaseReason) => void
}
export interface PoolEntry {
  readonly key: string
  readonly tier: PoolTier
  readonly owner: string | undefined
  readonly bytes: number
  readonly live: boolean
  /** Marks the entry most recently used within its tier. */
  touch(): void
  /** Shields the entry from eviction until the returned (idempotent) unpin runs. */
  pin(): () => void
  /** The owner's own release: uncharges once, no callback. */
  release(): void
}
/** One operation's transient share; `set` is absolute, not additive. */
export interface PoolCharge {
  readonly bytes: number
  /** Resizes the charge, evicting retained data to make room; false (and
   * unchanged) when the bytes cannot fit under the hard limit. */
  set(bytes: number): boolean
  release(): void
}
export interface PoolStats {
  limitBytes: number
  highWatermarkBytes: number
  lowWatermarkBytes: number
  retainedBytes: number
  transientBytes: number
  totalBytes: number
  peakBytes: number
  entries: number
  byTier: Record<PoolTier, number>
  evictions: number
  evictedBytes: number
  closed: boolean
}
export interface MemoryPoolOptions {
  limitBytes: number
  /** Fractions of limitBytes. */
  highWatermark?: number
  lowWatermark?: number
}
interface Slot {
  entry: PoolEntry
  init: PoolEntryInit
  pins: number
  live: boolean
}
function validBytes(bytes: number): boolean { return Number.isSafeInteger(bytes) && bytes >= 0 }

export class MemoryPool {
  readonly limitBytes: number
  private readonly high: number
  private readonly low: number
  /** Map insertion order is recency order within the whole pool (oldest first). */
  private readonly slots = new Map<string, Slot>()
  private readonly tierBytes: Record<PoolTier, number> = { rebuildable: 0, source: 0, view: 0 }
  private retained = 0
  private transient = 0
  private peak = 0
  private evictions = 0
  private evictedBytes = 0
  private protectedOwner: string | null = null
  private isClosed = false

  constructor(options: MemoryPoolOptions) {
    const high = options.highWatermark ?? DEFAULT_HIGH_WATERMARK
    const low = options.lowWatermark ?? Math.min(high, DEFAULT_LOW_WATERMARK)
    if (!validBytes(options.limitBytes) || !(high > 0 && high <= 1) || !(low > 0 && low <= high)) {
      throw new RangeError('Invalid memory pool budget')
    }
    this.limitBytes = options.limitBytes
    this.high = Math.floor(options.limitBytes * high)
    this.low = Math.floor(options.limitBytes * low)
  }
  get closed(): boolean { return this.isClosed }
  get totalBytes(): number { return this.retained + this.transient }

  /** Admits a retained entry. Without `force` it first evicts other entries so
   * the total stays under the hard limit and returns null when it cannot; with
   * `force` (bytes an operation already admitted transiently) it always lands.
   * Either way the watermark pass runs afterwards, never evicting the new entry.
   * Retaining an existing key replaces (and releases) the old entry. */
  retain(key: string, init: PoolEntryInit, options: { force?: boolean } = {}): PoolEntry | null {
    if (this.isClosed || !validBytes(init.bytes)) return null
    const previous = this.slots.get(key)
    const reclaimable = previous ? previous.init.bytes : 0
    if (!options.force && !this.fit(init.bytes - reclaimable, previous)) return null
    previous?.entry.release()
    const pool = this
    const slot: Slot = {
      init: { ...init }, pins: 0, live: true,
      entry: {
        key, tier: init.tier, owner: init.owner, bytes: init.bytes,
        get live() { return slot.live },
        touch() {
          if (!slot.live) return
          pool.slots.delete(key)
          pool.slots.set(key, slot)
        },
        pin() {
          if (!slot.live) return () => {}
          slot.pins++
          let done = false
          return () => { if (!done) { done = true; slot.pins-- } }
        },
        release() { pool.drop(slot) },
      },
    }
    this.slots.set(key, slot)
    this.tierBytes[init.tier] += init.bytes
    this.retained += init.bytes
    this.notePeak()
    this.reclaim(slot)
    return slot.entry
  }
  get(key: string): PoolEntry | undefined { return this.slots.get(key)?.entry }
  charge(): PoolCharge {
    let held = 0
    const pool = this
    return {
      get bytes() { return held },
      set(bytes: number): boolean {
        if (pool.isClosed || !validBytes(bytes)) return false
        if (bytes > held && !pool.fit(bytes - held)) return false
        pool.transient += bytes - held
        held = bytes
        pool.notePeak()
        pool.reclaim(null)
        return true
      },
      release() { pool.transient -= held; held = 0 },
    }
  }
  /** Evicts until `bytes` more fit under the hard limit; false when they cannot. */
  makeRoom(bytes: number): boolean {
    return !this.isClosed && validBytes(bytes) && this.fit(bytes)
  }
  /** The active owner (sheet): its view entries are never evicted. */
  protect(owner: string | null): void { this.protectedOwner = owner }
  stats(): PoolStats {
    return {
      limitBytes: this.limitBytes, highWatermarkBytes: this.high, lowWatermarkBytes: this.low,
      retainedBytes: this.retained, transientBytes: this.transient, totalBytes: this.totalBytes, peakBytes: this.peak,
      entries: this.slots.size, byTier: { ...this.tierBytes }, evictions: this.evictions, evictedBytes: this.evictedBytes,
      closed: this.isClosed,
    }
  }
  /** Releases every entry ('closed') and refuses all later admissions. Live
   * transient charges stay counted until their operations release them. */
  close(): void {
    if (this.isClosed) return
    this.isClosed = true
    for (const slot of [...this.slots.values()]) {
      this.drop(slot)
      slot.init.onRelease?.('closed')
    }
  }

  private drop(slot: Slot): void {
    if (!slot.live) return
    slot.live = false
    if (this.slots.get(slot.entry.key) === slot) this.slots.delete(slot.entry.key)
    this.tierBytes[slot.init.tier] -= slot.init.bytes
    this.retained -= slot.init.bytes
  }
  private notePeak(): void { this.peak = Math.max(this.peak, this.totalBytes) }
  private evictable(slot: Slot, keep: Slot | null | undefined): boolean {
    if (slot === keep || slot.pins > 0 || slot.init.evictable === false) return false
    return !(slot.init.tier === 'view' && slot.init.owner !== undefined && slot.init.owner === this.protectedOwner)
  }
  /** Evicts in tier order, LRU within each tier, until `done` holds. */
  private evictUntil(done: () => boolean, keep: Slot | null | undefined): void {
    for (const tier of EVICTION_ORDER) {
      for (const slot of [...this.slots.values()]) {
        if (done()) return
        if (slot.init.tier !== tier || !this.evictable(slot, keep)) continue
        this.evictions++
        this.evictedBytes += slot.init.bytes
        this.drop(slot)
        slot.init.onRelease?.('evicted')
      }
    }
  }
  /** Room for `extra` more bytes under the hard limit (evicting if needed). */
  private fit(extra: number, keep?: Slot): boolean {
    if (this.totalBytes + extra <= this.limitBytes) return true
    // Never evict anything when even full eviction could not make the room.
    let reclaimable = 0
    for (const slot of this.slots.values()) if (this.evictable(slot, keep)) reclaimable += slot.init.bytes
    if (this.totalBytes - reclaimable + extra > this.limitBytes) return false
    this.evictUntil(() => this.totalBytes + extra <= this.limitBytes, keep)
    return true
  }
  /** High-watermark pass: down to the low watermark, sparing `keep`. */
  private reclaim(keep: Slot | null): void {
    if (this.totalBytes <= this.high) return
    this.evictUntil(() => this.totalBytes <= this.low, keep)
  }
}

export interface BitmapPoolPlan {
  devicePx: { width: number; height: number }
  /** One viewport-sized overlay surface. */
  overlayBytes: number
  /** Base tiles covering the viewport plus a one-tile overscan ring. */
  baseTiles: number
  baseBytes: number
  /** One in-flight replacement row or column of tiles, whichever is longer. */
  replacementBytes: number
  totalBytes: number
}
/** §12.4 Canvas2D reference arithmetic (RGBA8, 512×512 device-pixel tiles). */
export function bitmapPoolPlan(viewport: { cssWidth: number; cssHeight: number; dpr: number }): BitmapPoolPlan {
  const width = Math.ceil(viewport.cssWidth * viewport.dpr), height = Math.ceil(viewport.cssHeight * viewport.dpr)
  if (!(width > 0 && height > 0) || !Number.isSafeInteger(width * height * RGBA_BYTES)) throw new RangeError('Invalid viewport')
  const tileBytes = TILE_DEVICE_PX * TILE_DEVICE_PX * RGBA_BYTES
  const cols = Math.ceil((width + 2 * TILE_DEVICE_PX) / TILE_DEVICE_PX)
  const rows = Math.ceil((height + 2 * TILE_DEVICE_PX) / TILE_DEVICE_PX)
  const overlayBytes = width * height * RGBA_BYTES
  const baseBytes = cols * rows * tileBytes
  const replacementBytes = Math.max(cols, rows) * tileBytes
  return {
    devicePx: { width, height }, overlayBytes, baseTiles: cols * rows, baseBytes, replacementBytes,
    totalBytes: overlayBytes + baseBytes + replacementBytes,
  }
}

/** ImageBitmap, OffscreenCanvas or a GPU texture wrapper. */
export interface TileSurface { readonly width: number; readonly height: number; close?(): void }
/** Complete tile identity (§10.1): cache owner plus backend/zoom/DPR/grid cell. */
export interface TileKey {
  identity: CacheIdentity
  backend: string
  zoom: number
  dpr: number
  col: number
  row: number
}
export interface TileCacheStats { tiles: number; bytes: number; closed: number }
const identityKey = (identity: CacheIdentity): string =>
  JSON.stringify([identity.docId, identity.sheetId, identity.valueGen, identity.fontGen, identity.viewGen])
const tileKey = (key: TileKey): string =>
  JSON.stringify([identityKey(key.identity), key.backend, key.zoom, key.dpr, key.col, key.row])
function closeSurface(surface: TileSurface): void {
  try { surface.close?.() } catch { /* already closed or lost with its device */ }
}

interface TileRecord { key: TileKey; surface: TileSurface; bytes: number; entry: PoolEntry | null; alive: boolean }

export class TileCache {
  private readonly tiles = new Map<string, TileRecord>()
  private bytes = 0
  private closed = 0
  private disposed = false
  constructor(readonly pool: MemoryPool) {}

  get(key: TileKey): TileSurface | undefined {
    const found = this.tiles.get(tileKey(key))
    found?.entry?.touch()
    return found?.surface
  }
  /** Admits a surface into the bitmap pool (evicting LRU tiles as needed). On
   * refusal the caller still owns, and must close, the surface. */
  put(key: TileKey, surface: TileSurface): boolean {
    if (this.disposed) return false
    const id = tileKey(key)
    const bytes = Math.ceil(surface.width) * Math.ceil(surface.height) * RGBA_BYTES
    const previous = this.tiles.get(id)
    const record: TileRecord = { key: { ...key, identity: { ...key.identity } }, surface, bytes, entry: null, alive: true }
    const entry = this.pool.retain(`tile:${id}`, {
      tier: 'rebuildable', bytes, owner: key.identity.sheetId,
      onRelease: () => this.forget(id, record),
    })
    if (!entry) return false
    record.entry = entry
    // The pool released the replaced entry without a callback: close it here.
    if (previous) this.forget(id, previous)
    this.tiles.set(id, record)
    this.bytes += bytes
    return true
  }
  /** Keeps only tiles of `identity` (complete cache owner) and/or `backend`. */
  retainOnly(keep: { identity?: CacheIdentity; backend?: string }): number {
    const owner = keep.identity ? identityKey(keep.identity) : null
    return this.invalidate(key => (owner !== null && identityKey(key.identity) !== owner)
      || (keep.backend !== undefined && key.backend !== keep.backend))
  }
  invalidate(predicate: (key: TileKey) => boolean): number {
    let dropped = 0
    for (const [id, record] of [...this.tiles]) {
      if (!predicate(record.key)) continue
      record.entry?.release()
      this.forget(id, record)
      dropped++
    }
    return dropped
  }
  stats(): TileCacheStats { return { tiles: this.tiles.size, bytes: this.bytes, closed: this.closed } }
  dispose(): void {
    if (this.disposed) return
    this.invalidate(() => true)
    this.disposed = true
  }
  /** Single close point: each record is uncharged and closed at most once. */
  private forget(id: string, record: TileRecord): void {
    if (!record.alive) return
    record.alive = false
    if (this.tiles.get(id) === record) {
      this.tiles.delete(id)
      this.bytes -= record.bytes
    }
    this.closed++
    closeSurface(record.surface)
  }
}
