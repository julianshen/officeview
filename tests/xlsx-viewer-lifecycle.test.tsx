/**
 * S4.3 lifecycle: LRU/resource close and font/image leases (spec §§10.3, 12.4).
 *
 * Every memory assertion reads an app ledger (MemoryPool/TileCache/session
 * stats), never process heap figures: the ledger is deterministic, and browser
 * retained-memory verification belongs to S10. Interleavings are driven by
 * explicit gates; nothing depends on elapsed time.
 */
import { describe, expect, it, vi } from 'vitest'
import type { CacheIdentity } from '../src/xlsx/view/contracts'
import {
  BITMAP_BUDGETS, bitmapPoolPlan, EVICTION_ORDER, MemoryPool, TileCache, WORKER_BUDGETS,
  type TileKey, type TileSurface,
} from '../src/xlsx/view/tileCache'

const MiB = 1024 * 1024
const identity = (overrides: Partial<CacheIdentity> = {}): CacheIdentity =>
  ({ docId: 'doc', sheetId: 'A', valueGen: 0, fontGen: 0, viewGen: 0, ...overrides })

describe('MemoryPool: one accounted document-wide budget', () => {
  it('evicts rebuildable data first, then source pages, then inactive views, LRU within a tier', () => {
    const pool = new MemoryPool({ limitBytes: 1000, highWatermark: 0.9, lowWatermark: 0.5 })
    const released: string[] = []
    const add = (key: string, tier: 'rebuildable' | 'source' | 'view', bytes: number, owner?: string) =>
      pool.retain(key, { tier, bytes, owner, onRelease: reason => released.push(`${key}:${reason}`) })
    add('view-B', 'view', 200, 'B')
    add('page-1', 'source', 150)
    add('prep-1', 'rebuildable', 100)
    add('page-2', 'source', 150)
    add('prep-2', 'rebuildable', 100)
    pool.get('prep-1')!.touch()
    expect(pool.stats()).toMatchObject({ retainedBytes: 700, totalBytes: 700, entries: 5, evictions: 0 })
    expect(pool.stats().byTier).toEqual({ rebuildable: 200, source: 300, view: 200 })
    // 700 + 250 = 950 > 900 high watermark: evict down to the 500 low watermark,
    // oldest rebuildable (prep-2: prep-1 was touched) first, then source pages.
    add('view-A', 'view', 250, 'A')
    expect(released).toEqual(['prep-2:evicted', 'prep-1:evicted', 'page-1:evicted', 'page-2:evicted'])
    expect(pool.stats()).toMatchObject({ retainedBytes: 450, evictions: 4, evictedBytes: 500 })
    expect(EVICTION_ORDER).toEqual(['rebuildable', 'source', 'view'])
  })
  it('never evicts pinned, protected or non-evictable entries and refuses admission instead', () => {
    const pool = new MemoryPool({ limitBytes: 1000 })
    const onRelease = vi.fn()
    const active = pool.retain('view-A', { tier: 'view', bytes: 400, owner: 'A', onRelease })!
    pool.protect('A')
    const fixed = pool.retain('view-B', { tier: 'view', bytes: 300, owner: 'B', evictable: false, onRelease })!
    const prep = pool.retain('prep', { tier: 'rebuildable', bytes: 200, onRelease })!
    const unpin = prep.pin()
    // 900 unevictable bytes: a 200-byte admission cannot fit under the hard limit.
    expect(pool.retain('more', { tier: 'rebuildable', bytes: 200 })).toBeNull()
    expect(pool.makeRoom(200)).toBe(false)
    expect(onRelease).not.toHaveBeenCalled()
    unpin(); unpin()
    expect(pool.makeRoom(200)).toBe(true)
    expect(onRelease).toHaveBeenCalledTimes(1)
    expect(prep.live).toBe(false)
    expect(pool.stats().peakBytes).toBeLessThanOrEqual(1000)
    // The owner's own release never calls back and is idempotent.
    active.release(); active.release(); fixed.release()
    expect(onRelease).toHaveBeenCalledTimes(1)
    expect(pool.stats()).toMatchObject({ retainedBytes: 0, entries: 0 })
  })
  it('charges transient buffers to the same budget and returns to baseline on close', () => {
    const pool = new MemoryPool({ limitBytes: 1000 })
    const onRelease = vi.fn()
    pool.retain('prep', { tier: 'rebuildable', bytes: 600, onRelease })
    const draft = pool.charge()
    // Absolute set: growing a transient makes room by evicting rebuildable data.
    expect(draft.set(500)).toBe(true)
    expect(onRelease).toHaveBeenCalledWith('evicted')
    expect(pool.stats()).toMatchObject({ transientBytes: 500, retainedBytes: 0, totalBytes: 500 })
    expect(draft.set(1200)).toBe(false)
    expect(pool.stats().transientBytes).toBe(500)
    const view = vi.fn()
    pool.retain('view', { tier: 'view', bytes: 100, onRelease: view })
    pool.close()
    expect(view).toHaveBeenCalledWith('closed')
    draft.release(); draft.release()
    expect(pool.stats()).toMatchObject({ retainedBytes: 0, transientBytes: 0, totalBytes: 0, entries: 0, closed: true })
    // A closed pool never revives.
    expect(pool.retain('late', { tier: 'view', bytes: 1 })).toBeNull()
    expect(pool.charge().set(1)).toBe(false)
  })
})

/** Fake ImageBitmap: jsdom has none; close() is what eviction must call. */
function surface(width = 512, height = 512): TileSurface & { close: ReturnType<typeof vi.fn> } {
  return { width, height, close: vi.fn() }
}
const tile = (overrides: Partial<TileKey> = {}): TileKey =>
  ({ identity: identity(), backend: 'canvas2d', zoom: 1, dpr: 2, col: 0, row: 0, ...overrides })

describe('TileCache: bitmap pool with ImageBitmap close', () => {
  it('pins the §12.4 bitmap pool arithmetic for both profiles', () => {
    const desktop = bitmapPoolPlan({ cssWidth: 1440, cssHeight: 900, dpr: 2 })
    expect(desktop).toEqual({
      devicePx: { width: 2880, height: 1800 }, overlayBytes: 20_736_000, baseTiles: 48, baseBytes: 48 * MiB,
      replacementBytes: 8 * MiB, totalBytes: 20_736_000 + 56 * MiB,
    })
    expect(desktop.totalBytes).toBeLessThan(BITMAP_BUDGETS.D)
    const mobile = bitmapPoolPlan({ cssWidth: 1024, cssHeight: 768, dpr: 2 })
    expect(mobile).toMatchObject({ overlayBytes: 12 * MiB, baseTiles: 30, baseBytes: 30 * MiB, replacementBytes: 6 * MiB })
    expect(mobile.totalBytes).toBeLessThan(BITMAP_BUDGETS.M)
    expect(WORKER_BUDGETS).toEqual({ D: 512 * MiB, M: 256 * MiB })
  })
  it('charges width×height×4 per tile and closes each evicted surface exactly once', () => {
    const pool = new MemoryPool({ limitBytes: 4 * MiB, highWatermark: 1, lowWatermark: 1 })
    const cache = new TileCache(pool)
    const surfaces = [0, 1, 2, 3].map(() => surface())
    surfaces.forEach((s, col) => expect(cache.put(tile({ col }), s)).toBe(true))
    expect(cache.stats()).toEqual({ tiles: 4, bytes: 4 * MiB, closed: 0 })
    expect(cache.get(tile({ col: 0 }))).toBe(surfaces[0])
    // A fifth tile evicts the least recently used one (col 1; col 0 was read).
    const fifth = surface()
    expect(cache.put(tile({ col: 4 }), fifth)).toBe(true)
    expect(surfaces[1].close).toHaveBeenCalledTimes(1)
    expect(cache.get(tile({ col: 1 }))).toBeUndefined()
    expect(cache.stats()).toEqual({ tiles: 4, bytes: 4 * MiB, closed: 1 })
    // Replacing a tile under the same key closes the old surface once.
    const replacement = surface()
    cache.put(tile({ col: 0 }), replacement)
    expect(surfaces[0].close).toHaveBeenCalledTimes(1)
    cache.dispose()
    for (const s of [surfaces[2], surfaces[3], fifth, replacement]) expect(s.close).toHaveBeenCalledTimes(1)
    expect(cache.stats()).toEqual({ tiles: 0, bytes: 0, closed: 6 })
    expect(pool.stats()).toMatchObject({ retainedBytes: 0, entries: 0 })
    expect(cache.put(tile({ col: 9 }), surface())).toBe(false)
  })
  it('drops tiles of superseded font generations, other sheets and the previous backend', () => {
    const pool = new MemoryPool({ limitBytes: 64 * MiB })
    const cache = new TileCache(pool)
    const old = surface(), otherSheet = surface()
    cache.put(tile({ col: 0 }), old)
    cache.put(tile({ col: 0, identity: identity({ sheetId: 'B' }) }), otherSheet)
    // Font readiness advances fontGen: only the new identity survives.
    const fresh = identity({ fontGen: 1 })
    const current = surface()
    cache.put(tile({ col: 1, identity: fresh }), current)
    expect(cache.retainOnly({ identity: fresh })).toBe(2)
    expect(old.close).toHaveBeenCalledTimes(1)
    expect(otherSheet.close).toHaveBeenCalledTimes(1)
    // Backend switch: both backends' surfaces count at the switch peak, then
    // only the new backend's surfaces remain at steady state.
    const gpu = surface()
    cache.put(tile({ col: 1, identity: fresh, backend: 'webgpu' }), gpu)
    expect(cache.stats().bytes).toBe(2 * MiB)
    expect(pool.stats().peakBytes).toBeGreaterThanOrEqual(3 * MiB)
    expect(cache.retainOnly({ identity: fresh, backend: 'webgpu' })).toBe(1)
    expect(current.close).toHaveBeenCalledTimes(1)
    expect(gpu.close).not.toHaveBeenCalled()
    expect(cache.stats()).toEqual({ tiles: 1, bytes: MiB, closed: 3 })
    expect(cache.invalidate(key => key.backend === 'webgpu')).toBe(1)
    expect(cache.stats()).toEqual({ tiles: 0, bytes: 0, closed: 4 })
  })
})
