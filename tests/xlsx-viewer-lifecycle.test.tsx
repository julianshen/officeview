/**
 * S4.3 lifecycle: LRU/resource close and font/image leases (spec §§10.3, 12.4).
 *
 * Every memory assertion reads an app ledger (MemoryPool/TileCache/session
 * stats), never process heap figures: the ledger is deterministic, and browser
 * retained-memory verification belongs to S10. Interleavings are driven by
 * explicit gates; nothing depends on elapsed time.
 */
import { render, waitFor } from '@testing-library/react'
import { readFileSync } from 'node:fs'
import { describe, expect, it, vi } from 'vitest'
import type { CacheIdentity, FilterSpec, RowOrder, SortSpec, SourceStamp, SourceStore, WorkerReply } from '../src/xlsx/view/contracts'
import type { XlsxDocument } from '../src/xlsx/types'
import type { PptxDocument } from '../src/pptx/types'
import {
  BITMAP_BUDGETS, bitmapPoolPlan, EVICTION_ORDER, MemoryPool, TileCache, WORKER_BUDGETS,
  type TileKey, type TileSurface,
} from '../src/xlsx/view/tileCache'
import { formatA1, MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { createIdentityOrder, createRowOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { GUARD_SCRATCH_BYTES } from '../src/xlsx/view/operationGuard'
import {
  DocumentSession, MIN_WORKER_BUDGET_BYTES, ViewNotResidentError, ViewQuotaError, WorkerBudgetError, type FilterEvaluator, type OperationOutcome, type SessionOptions, type WorkbookSource,
} from '../src/worker/viewerSession'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { initialIdentity } from '../src/worker/viewerProtocol'
import { getPaintables } from '../src/render/paint'
import { acquireFonts } from '../src/core/fonts/register'
import { OfficeDoc } from '../src/components/OfficeDoc'

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
    // Memory pressure trims best-effort down to a target, views last.
    expect(pool.trim(250)).toBe(200)
    expect(released.at(-1)).toBe('view-B:evicted')
    expect(pool.stats()).toMatchObject({ retainedBytes: 250, entries: 1 })
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
    // No forced admission exists: an oversized entry is refused outright.
    expect(pool.retain('huge', { tier: 'view', bytes: 1001 })).toBeNull()
    expect(pool.stats().totalBytes).toBeLessThanOrEqual(1000)
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
function surface(width = 512, height = 512) {
  const close = vi.fn<() => void>()
  return { width, height, close } satisfies TileSurface
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

// ---------------------------------------------------------------------------
// Worker session: one document-wide worker budget, view eviction by recipe
// ---------------------------------------------------------------------------

type Value = number | string | null
interface CellInput { row: number; col: number; value: Value }
function sparseDoc(sheets: { sheetId: string; cells: CellInput[] }[]): XlsxDocument {
  return {
    sheets: sheets.map(sheet => {
      const byRow = new Map<number, CellInput[]>()
      for (const cell of sheet.cells) byRow.set(cell.row, [...(byRow.get(cell.row) ?? []), cell])
      return {
        name: `Sheet ${sheet.sheetId}`, sheetId: sheet.sheetId,
        rows: [...byRow.keys()].sort((a, b) => a - b).map(index => ({
          index,
          cells: byRow.get(index)!.sort((a, b) => a.col - b.col).filter(cell => cell.value !== null).map(cell => ({
            ref: formatA1({ row: cell.row, col: cell.col }), row: cell.row, col: cell.col, value: cell.value, styleIndex: 0, hasCachedValue: true,
          })),
        })),
        cols: [], merges: [], mergeRanges: [],
      }
    }),
    tables: [],
  }
}
/** Full-capacity sparse sheet: populated A1048576/XFD1048576 corners outside
 * the sort band plus a key column A over body rows 0..rows-1 (every 7th row). */
function fullCapacitySheet(sheetId: string, rows: number, seed: number): { sheetId: string; cells: CellInput[] } {
  const cells: CellInput[] = [
    { row: MAX_ROWS - 1, col: 0, value: 'A1048576' }, { row: MAX_ROWS - 1, col: MAX_COLS - 1, value: 'XFD1048576' },
    { row: rows, col: MAX_COLS - 1, value: 'below band' },
  ]
  for (let row = 0; row < rows; row += 7) cells.push({ row, col: 0, value: (row * 7919 + seed * 104729) % 1000 })
  return { sheetId, cells }
}
const band = (rows: number) => ({ r0: 0, r1: rows - 1, c0: 0, c1: 0, headerRows: 0 as const, totalsRows: 0 as const })
const sortOf = (rows: number, order: 'asc' | 'desc' = 'asc'): SortSpec => ({ range: band(rows), keys: [{ col: 0, order }], caseSensitive: false })
const filterOf = (rows: number): FilterSpec => ({ range: band(rows), columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) })
/** Deterministic S8 stand-in: hides every third body row. */
const thirds: FilterEvaluator = async () => ({ kind: 'complete', matches: row => row % 3 !== 0 })
const fast = { now: () => 0, yield: async () => {} }
const live = () => new AbortController().signal
/** Test-side digest of an order by source identity (bands + sourceAtOffset). */
function orderDigest(order: RowOrder): string {
  let hash = 0x811c9dc5
  const mix = (value: number) => { hash = Math.imul(hash ^ value, 0x01000193) >>> 0 }
  for (const entry of order.bands) {
    mix(entry.r0); mix(entry.r1)
    for (const row of entry.sourceAtOffset) mix(row)
  }
  return `${order.kind}:${order.bands.length}:${hash.toString(16)}`
}
async function committed(outcome: Promise<OperationOutcome>): Promise<void> {
  const result = await outcome
  if (result.kind !== 'committed') throw new Error(`expected commit, got ${JSON.stringify(result)}`)
}
const objectSession = (doc: XlsxDocument, options: SessionOptions = {}) =>
  new DocumentSession('doc', createObjectWorkbookSource(doc, 'doc'), { ...fast, filterEvaluator: thirds, ...options })

describe('DocumentSession lifecycle', () => {
  it('a released session never revives: no lazy sheet state, activation or publication', async () => {
    const session = objectSession(sparseDoc([fullCapacitySheet('A', 500, 1), fullCapacitySheet('B', 500, 2)]))
    await committed(session.sort(session.identity('A'), sortOf(500), live()))
    const owner = session.identity('A')
    expect(session.memoryStats().retainedBytes).toBeGreaterThan(0)
    await session.release()
    expect(() => session.view('B')).toThrow(/closed/i)
    expect(() => session.activate('B')).toThrow(/closed/i)
    expect(session.isCurrent(owner)).toBe(false)
    expect(session.commitDraft(owner, { order: createIdentityOrder(), visibility: visibilityFromHidden(createIntervalSet([])), sort: null, filter: null })).toBeNull()
    expect(await session.ensureView('A', live())).toEqual({ kind: 'stale' })
    expect(await session.sort(owner, sortOf(500), live())).toEqual({ kind: 'stale' })
    expect(session.memoryStats()).toMatchObject({ retainedBytes: 0, transientBytes: 0, viewBytes: 0, preparedBytes: 0 })
  })

  it('charges every sheet view, prepared column and transient draft to one document-wide budget', async () => {
    const rows = 65_536
    const ids = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
    const session = objectSession(sparseDoc(ids.map((id, i) => fullCapacitySheet(id, rows, i))), { workerBudgetBytes: 4 * MiB })
    const totals: number[] = []
    for (const id of ids) {
      session.activate(id)
      await committed(session.sort(session.identity(id), sortOf(rows), live()))
      const stats = session.memoryStats()
      totals.push(stats.totalBytes)
      expect(stats.limitBytes).toBe(4 * MiB)
      expect(stats.peakBytes).toBeLessThanOrEqual(4 * MiB)
      expect(session.isResident(id)).toBe(true)
    }
    const stats = session.memoryStats()
    // Eight resident sorted views (order + inverse + mask) would exceed the budget.
    const perView = rows * 8 + Math.ceil(MAX_ROWS / 32) * 4
    expect(perView * ids.length).toBeGreaterThan(4 * MiB)
    expect(stats.evictions).toBeGreaterThan(0)
    expect(stats.viewBytes).toBeLessThanOrEqual(stats.retainedBytes)
    expect(stats.transientBytes).toBe(0)
    expect(ids.filter(id => !session.isResident(id)).length).toBeGreaterThan(0)
    expect(Math.max(...totals)).toBeLessThanOrEqual(4 * MiB)
    await session.release()
  })

  it('an evicted view rebuilds by recipe under the same identity (filter, then sort with pinned rows)', async () => {
    const rows = 5000
    const session = objectSession(sparseDoc([fullCapacitySheet('A', rows, 3), fullCapacitySheet('B', rows, 4)]))
    await committed(session.filter(session.identity('A'), filterOf(rows), 0, live()))
    // The sort pins the filter-hidden rows at their source positions.
    await committed(session.sort(session.identity('A'), sortOf(rows, 'desc'), live()))
    const before = session.view('A')
    const digest = orderDigest(before.order)
    const generation = session.globalViewGen
    session.activate('B')
    // Memory pressure while B is active: everything evictable goes.
    evictAll(session)
    expect(session.isResident('A')).toBe(false)
    expect(() => session.view('A')).toThrow(/not resident/i)
    expect(session.isCurrent(before.identity)).toBe(true)
    const rebuilt = await session.ensureView('A', live())
    if (rebuilt.kind !== 'resident') throw new Error(`rebuild failed: ${JSON.stringify(rebuilt)}`)
    expect(orderDigest(rebuilt.view.order)).toBe(digest)
    expect(rebuilt.view.counts).toEqual(before.counts)
    expect(rebuilt.view.identity).toEqual(before.identity)
    expect(rebuilt.view.revision).toBe(before.revision)
    expect(session.globalViewGen).toBe(generation)
    expect(session.isCurrent(before.identity)).toBe(true)
    // Visible membership is identical slot by slot.
    for (let slot = 0; slot < rows; slot += 97) expect(rebuilt.view.visibility.visible(slot)).toBe(before.visibility.visible(slot))
    await session.release()
  })

  it('owner-free committed drafts are not rebuildable and are never evicted', async () => {
    const session = objectSession(sparseDoc([fullCapacitySheet('A', 50, 1), fullCapacitySheet('B', 50, 2)]))
    const order = createRowOrder([{ r0: 0, r1: 3, sourceAtOffset: new Uint32Array([3, 1, 0, 2]) }])
    await session.commitView('A', { order, visibility: visibilityFromHidden(createIntervalSet([])), sort: null, filter: null })
    session.activate('B')
    evictAll(session)
    expect(session.isResident('A')).toBe(true)
    expect([...session.view('A').order.bands[0].sourceAtOffset]).toEqual([3, 1, 0, 2])
    await session.release()
  })

  it('charges the guard scratch bitmap and projection metrics, released on every outcome', async () => {
    const rows = 5000
    const session = objectSession(sparseDoc([fullCapacitySheet('A', rows, 5)]))
    await committed(session.filter(session.identity('A'), filterOf(rows), 0, live()))
    const stats = session.memoryStats()
    const mask = Math.ceil(MAX_ROWS / 32) * 4
    // Draft mask + the guard's 128 KiB source bitmap are live at the same time.
    expect(stats.peakTransientBytes).toBeGreaterThanOrEqual(mask + GUARD_SCRATCH_BYTES)
    expect(GUARD_SCRATCH_BYTES).toBe(131_072)
    expect(stats.transientBytes).toBe(0)
    // A cancelled filter releases its charges too.
    const controller = new AbortController()
    controller.abort()
    expect(await session.filter(session.identity('A'), null, 0, controller.signal)).toMatchObject({ kind: 'cancelled' })
    expect(session.memoryStats().transientBytes).toBe(0)
    await session.release()
  })

  it('prepared columns and structural handles are rebuildable pool entries uncharged exactly once', async () => {
    const rows = 5000
    const session = objectSession(sparseDoc([fullCapacitySheet('A', rows, 6)]))
    expect(await session.prepareColumns(session.identity('A'), [0, MAX_COLS - 1], live())).toEqual({ kind: 'complete' })
    const prepared = session.preparedStats()
    expect(prepared.columns).toBe(2)
    expect(session.memoryStats().preparedBytes).toBe(prepared.bytes)
    expect(session.memory.stats().byTier.rebuildable).toBeGreaterThanOrEqual(prepared.bytes)
    await committed(session.sort(session.identity('A'), sortOf(rows), live()))
    expect(session.memoryStats().structureBytes).toBeGreaterThan(0)
    evictAll(session)
    expect(session.preparedStats()).toEqual({ columns: 0, bytes: 0 })
    expect(session.memoryStats()).toMatchObject({ preparedBytes: 0, structureBytes: 0 })
    expect(session.structureCacheSize()).toBe(0)
    // Rebuilt on demand, then dropped by the owner itself (no double uncharge).
    expect(await session.prepareColumns(session.identity('A'), [0], live())).toEqual({ kind: 'complete' })
    await session.release()
    expect(session.memoryStats()).toMatchObject({ preparedBytes: 0, retainedBytes: 0 })
  })

  it('source pages register in the source tier: evicted after rebuildable data, before views, closed on release', async () => {
    const doc = sparseDoc([fullCapacitySheet('A', 50, 1), fullCapacitySheet('B', 50, 2)])
    const inner = createObjectWorkbookSource(doc, 'doc')
    const pages: string[] = []
    let pool: MemoryPool | undefined
    const source: WorkbookSource = {
      valuePolicy: inner.valuePolicy, sheets: inner.sheets, store: id => inner.store(id), geometry: id => inner.geometry(id),
      release: () => inner.release(),
      bindMemory(memory) { pool = memory },
    }
    const session = new DocumentSession('doc', source, { ...fast, filterEvaluator: thirds, workerBudgetBytes: 2 * MiB })
    expect(pool).toBe(session.memory)
    for (const page of ['p1', 'p2']) pool!.retain(`page:${page}`, { tier: 'source', bytes: 256 * 1024, onRelease: reason => pages.push(`${page}:${reason}`) })
    await committed(session.sort(session.identity('A'), sortOf(50), live()))
    session.activate('B')
    evictAll(session)
    expect(pages).toEqual(['p1:evicted', 'p2:evicted'])
    pool!.retain('page:p3', { tier: 'source', bytes: 1024, onRelease: reason => pages.push(`p3:${reason}`) })
    await session.release()
    expect(pages).toEqual(['p1:evicted', 'p2:evicted', 'p3:closed'])
    expect(session.memory.stats()).toMatchObject({ retainedBytes: 0, closed: true })
  })
})

/** Memory pressure: evicts everything evictable (the active sheet's view and
 * pinned or non-rebuildable entries stay). */
function evictAll(session: DocumentSession): void { session.memory.trim(0) }

describe('F7 multi-sheet view stress', () => {
  it('aggregate memory plateaus while switching among many full-capacity sparse sheets with local sort/filter states', async () => {
    const rows = 65_536
    const ids = Array.from({ length: 12 }, (_v, i) => `S${i}`)
    const limit = 4 * MiB
    const session = objectSession(sparseDoc(ids.map((id, i) => fullCapacitySheet(id, rows, i))), { workerBudgetBytes: limit })
    const applied = new Map<string, { digest: string; view: ReturnType<DocumentSession['view']> }>()
    const rounds: number[][] = []
    for (let round = 0; round < 3; round++) {
      const totals: number[] = []
      for (const [index, id] of ids.entries()) {
        session.activate(id)
        if (round === 0) {
          // Alternate the operation history: sort→filter and filter→sort.
          if (index % 2) await committed(session.filter(session.identity(id), filterOf(rows), 0, live()))
          await committed(session.sort(session.identity(id), sortOf(rows, index % 3 ? 'asc' : 'desc'), live()))
          if (!(index % 2)) await committed(session.filter(session.identity(id), filterOf(rows), 0, live()))
          const view = session.view(id)
          applied.set(id, { digest: orderDigest(view.order), view })
        } else {
          const rebuilt = await session.ensureView(id, live())
          if (rebuilt.kind !== 'resident') throw new Error(`${id}: ${JSON.stringify(rebuilt)}`)
          const before = applied.get(id)!
          expect(orderDigest(rebuilt.view.order)).toBe(before.digest)
          expect(rebuilt.view.identity).toEqual(before.view.identity)
          expect(rebuilt.view.revision).toBe(before.view.revision)
          expect(rebuilt.view.counts).toEqual(before.view.counts)
        }
        const stats = session.memoryStats()
        expect(stats.totalBytes).toBeLessThanOrEqual(limit)
        expect(stats.transientBytes).toBe(0)
        totals.push(stats.retainedBytes)
      }
      rounds.push(totals)
    }
    const stats = session.memoryStats()
    expect(stats.peakBytes).toBeLessThanOrEqual(limit)
    expect(stats.evictions).toBeGreaterThan(ids.length)
    // Plateau: retained bytes stop growing with the number of visited sheets.
    // Steady-state rounds retain one constant ledger value under the high
    // watermark (a bounded set of materialized views), not 12 sheets' worth.
    const high = session.memory.stats().highWatermarkBytes
    for (const totals of rounds) expect(Math.max(...totals)).toBeLessThanOrEqual(high)
    expect(Math.max(...rounds[2])).toBeLessThanOrEqual(Math.max(...rounds[1]))
    expect(new Set(rounds[2]).size).toBe(1)
    const perView = rows * 8 + Math.ceil(MAX_ROWS / 32) * 4 + 4 * (Math.ceil(MAX_ROWS / 32 / 32) + 1)
    expect(rounds[2][0]).toBeLessThan(perView * ids.length / 2)
    expect(ids.filter(id => session.isResident(id)).length).toBeLessThan(ids.length / 2)
    await session.release()
    expect(session.memoryStats()).toMatchObject({ retainedBytes: 0, transientBytes: 0 })
  }, 60_000)
})

describe('ViewerKernel: repeated open/close', () => {
  it('returns every ledger to baseline after close with work in flight, and closed epochs never revive', async () => {
    const doc = sparseDoc([fullCapacitySheet('A', 5000, 1), fullCapacitySheet('B', 5000, 2)])
    let releases = 0
    const kernel = new ViewerKernel({
      parse: async () => doc,
      openSource: async (_handle, docId) => {
        const inner = createObjectWorkbookSource(doc, docId)
        return {
          valuePolicy: inner.valuePolicy, sheets: inner.sheets, store: id => inner.store(id), geometry: id => inner.geometry(id),
          release: async () => { releases++; await inner.release() },
        }
      },
    }, { ...fast, filterEvaluator: thirds })
    const replies: WorkerReply[] = []
    const sink = (out: WorkerReply) => { replies.push(out) }
    const closedSessions: DocumentSession[] = []
    const cycles = 10
    for (let cycle = 0; cycle < cycles; cycle++) {
      const docId = `doc-${cycle}`
      await kernel.handle({ kind: 'open', requestId: `o${cycle}`, docId, source: { kind: 'object', document: doc }, profile: 'D' }, sink)
      const session = kernel.session(docId)!
      const a = initialIdentity(docId, 'A')
      await kernel.handle({ kind: 'prepareColumns', requestId: `p${cycle}`, owner: a, cols: new Uint16Array([0]) }, sink)
      await kernel.handle({ kind: 'sort', requestId: `s${cycle}`, owner: a, spec: sortOf(5000) }, sink)
      await kernel.handle({ kind: 'activateSheet', requestId: `a${cycle}`, owner: session.workbookOwner(), sheetId: 'B' }, sink)
      // Leave a sort and a filter in flight, then close the document.
      const inflight = [
        kernel.handle({ kind: 'sort', requestId: `sb${cycle}`, owner: initialIdentity(docId, 'B'), spec: sortOf(5000, 'desc') }, sink),
        kernel.handle({ kind: 'filter', requestId: `fa${cycle}`, owner: session.identity('A'), spec: filterOf(5000), nowMs: 0 }, sink),
      ]
      await kernel.handle({ kind: 'close', requestId: `c${cycle}`, docId }, sink)
      await Promise.all(inflight)
      closedSessions.push(session)
      expect(kernel.lifecycleStats()).toEqual({ sessions: 0, inflight: 0, opening: 0, retainedBytes: 0, transientBytes: 0 })
      expect(session.memoryStats()).toMatchObject({ retainedBytes: 0, transientBytes: 0, viewBytes: 0, preparedBytes: 0 })
    }
    expect(releases).toBe(cycles)
    expect(kernel.stats.opens).toBe(cycles)
    const late = replies.filter(out => out.requestId.startsWith('sb') || out.requestId.startsWith('fa'))
    // In-flight work ends cancelled (document-closed) or stale, never as a completion.
    expect(late.filter(out => out.kind === 'complete')).toEqual([])
    for (const session of closedSessions) {
      expect(session.isClosed).toBe(true)
      expect(() => session.view('A')).toThrow(/closed/i)
      expect(session.memory.stats().closed).toBe(true)
    }
  })

  it('activateSheet rebuilds an evicted sheet view before replying', async () => {
    const doc = sparseDoc([fullCapacitySheet('A', 5000, 1), fullCapacitySheet('B', 5000, 2)])
    const kernel = new ViewerKernel({ parse: async () => doc }, { ...fast, filterEvaluator: thirds })
    const replies: WorkerReply[] = []
    const sink = (out: WorkerReply) => { replies.push(out) }
    await kernel.handle({ kind: 'open', requestId: 'o', docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }, sink)
    const session = kernel.session('doc')!
    await kernel.handle({ kind: 'sort', requestId: 's', owner: initialIdentity('doc', 'A'), spec: sortOf(5000, 'desc') }, sink)
    const digest = orderDigest(session.view('A').order)
    await kernel.handle({ kind: 'activateSheet', requestId: 'b', owner: session.workbookOwner(), sheetId: 'B' }, sink)
    session.memory.trim(0)
    expect(session.isResident('A')).toBe(false)
    await kernel.handle({ kind: 'activateSheet', requestId: 'a', owner: session.workbookOwner(), sheetId: 'A' }, sink)
    expect(replies.at(-1)).toMatchObject({ requestId: 'a', kind: 'complete' })
    expect(session.isResident('A')).toBe(true)
    expect(orderDigest(session.view('A').order)).toBe(digest)
  })
})

// ---------------------------------------------------------------------------
// Existing image/font leases: exactly-once retain/release
// ---------------------------------------------------------------------------

function fakeBitmap() {
  const close = vi.fn<() => void>()
  return { width: 10, height: 10, close }
}
const pngBytes = () => new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0, 0, 0, 0])

describe('image and font leases', () => {
  it('repeated extraction and double dispose close each decoded bitmap exactly once', async () => {
    const decoded: ReturnType<typeof fakeBitmap>[] = []
    const decodeImage = vi.fn(async () => { const bitmap = fakeBitmap(); decoded.push(bitmap); return bitmap as unknown as CanvasImageSource })
    const doc = { sheets: [{ name: 'S', rows: [], cols: [], merges: [], mergeRanges: [], maxRow: 1, maxCol: 1 }], images: [{ data: pngBytes(), mime: 'image/png' }, { data: pngBytes(), mime: 'image/png' }] } as unknown as XlsxDocument
    for (let i = 0; i < 5; i++) {
      const pages = await getPaintables(doc, { decodeImage })
      pages.dispose(); pages.dispose()
    }
    // Identical candidate pairs decode once per extraction.
    expect(decoded).toHaveLength(5)
    for (const bitmap of decoded) expect(bitmap.close).toHaveBeenCalledTimes(1)
  })

  it('a failure after decoding releases the decoded bitmaps and the font lease', async () => {
    const decoded: ReturnType<typeof fakeBitmap>[] = []
    const decodeImage = vi.fn(async () => { const bitmap = fakeBitmap(); decoded.push(bitmap); return bitmap as unknown as CanvasImageSource })
    const cleanup = vi.fn()
    const registerFont = vi.fn(async () => cleanup)
    // `cols` missing: sheet metrics throw after the images were decoded.
    const doc = {
      sheets: [{ name: 'Broken', rows: [], merges: [], mergeRanges: [] }], images: [{ data: pngBytes(), mime: 'image/png' }],
      embeddedFonts: [{ family: 'Liter', variant: 'regular', relationshipId: 'f', partPath: 'f', bytes: new Uint8Array(4) }],
    } as unknown as XlsxDocument
    await expect(getPaintables(doc, { decodeImage, registerFont })).rejects.toThrow()
    expect(decoded).toHaveLength(1)
    expect(decoded[0].close).toHaveBeenCalledTimes(1)
    expect(cleanup).toHaveBeenCalledTimes(1)
  })

  it('an offscreen consumer lease keeps fonts registered for active paint until the last release', async () => {
    const cleanup = vi.fn()
    const registerFont = vi.fn(async () => cleanup)
    const doc = { embeddedFonts: [{ family: 'Liter', variant: 'regular' as const, relationshipId: 'f', partPath: 'f', bytes: new Uint8Array(4) }] }
    const paint = await acquireFonts(doc, registerFont)
    const search = await acquireFonts(doc, registerFont)
    expect(registerFont).toHaveBeenCalledTimes(1)
    search.dispose(); search.dispose()
    expect(cleanup).not.toHaveBeenCalled()
    expect(paint.resolve('Liter')).toMatch(/^OfficeviewFont_/)
    paint.dispose(); paint.dispose()
    expect(cleanup).toHaveBeenCalledTimes(1)
    // A later consumer registers afresh: nothing revived from the released state.
    const again = await acquireFonts(doc, registerFont)
    expect(registerFont).toHaveBeenCalledTimes(2)
    again.dispose()
    expect(cleanup).toHaveBeenCalledTimes(2)
  })

  it('unmounting while font registration is in flight releases the face exactly once when it lands', async () => {
    const add = vi.fn(), remove = vi.fn()
    let land!: () => void
    const landed = new Promise<void>(resolve => { land = resolve })
    class Face { async load() { await landed; return this } }
    vi.stubGlobal('FontFace', Face)
    Object.defineProperty(document, 'fonts', { configurable: true, value: { add, delete: remove } })
    try {
      const doc: PptxDocument = {
        slideWidthEmu: 100000, slideHeightEmu: 100000, images: [], slides: [{ index: 0, widthEmu: 100000, heightEmu: 100000, shapes: [] }],
        embeddedFonts: [{ family: 'Liter', variant: 'regular', relationshipId: 'font', partPath: 'font', bytes: new Uint8Array(readFileSync('tests/fixtures/fonts/Liter-Regular.ttf')) }],
      }
      const view = render(<OfficeDoc document={doc} />)
      view.unmount()
      expect(add).not.toHaveBeenCalled()
      land()
      await waitFor(() => expect(remove).toHaveBeenCalledTimes(1))
      expect(add).toHaveBeenCalledTimes(1)
    } finally { vi.unstubAllGlobals() }
  })

  it('repeated OfficeDoc mount/unmount registers and removes each embedded face once', async () => {
    const add = vi.fn(), remove = vi.fn()
    class Face { async load() { return this } }
    vi.stubGlobal('FontFace', Face)
    Object.defineProperty(document, 'fonts', { configurable: true, value: { add, delete: remove } })
    try {
      const doc: PptxDocument = {
        slideWidthEmu: 100000, slideHeightEmu: 100000, images: [], slides: [{ index: 0, widthEmu: 100000, heightEmu: 100000, shapes: [] }],
        embeddedFonts: [{ family: 'Liter', variant: 'regular', relationshipId: 'font', partPath: 'font', bytes: new Uint8Array(readFileSync('tests/fixtures/fonts/Liter-Regular.ttf')) }],
      }
      for (let i = 0; i < 3; i++) {
        const view = render(<OfficeDoc document={doc} />)
        await waitFor(() => expect(view.container.querySelectorAll('canvas')).toHaveLength(1))
        view.unmount()
        await waitFor(() => expect(remove).toHaveBeenCalledTimes(i + 1))
      }
      expect(add).toHaveBeenCalledTimes(3)
    } finally { vi.unstubAllGlobals() }
  })
})

// ---------------------------------------------------------------------------
// Review fixes: hard-limit admission, value publication, activation, lazy masks
// ---------------------------------------------------------------------------

/** Workbook source whose values change on publish(): each generation reads a
 * different parsed document under one outer stamp sequence. */
function versionedSource(docs: XlsxDocument[]) {
  let gen = 0
  const inners = docs.map(doc => createObjectWorkbookSource(doc, 'doc'))
  const stores = new Map<string, SourceStore>()
  const source: WorkbookSource = {
    valuePolicy: inners[0].valuePolicy, sheets: inners[0].sheets, geometry: id => inners[0].geometry(id),
    release: async () => { for (const inner of inners) await inner.release() },
    store(sheetId) {
      let store = stores.get(sheetId)
      if (store) return store
      const inner = () => inners[gen].store(sheetId)
      const outer = (): SourceStamp => Object.freeze({ ...inner().stamp, valueGen: gen })
      const fix = <O extends { stamp: SourceStamp }>(options: O): O =>
        ({ ...options, stamp: options.stamp?.valueGen === gen ? inner().stamp : { ...inner().stamp, valueGen: -1 } })
      const out = <T extends { stamp: SourceStamp }>(page: T): T => ({ ...page, stamp: outer() })
      store = Object.create(inners[0].store(sheetId)) as SourceStore
      Object.defineProperty(store, 'stamp', { get: outer })
      store.queryColumn = (request, options) => inner().queryColumn(request, fix(options)).then(out)
      store.readRows = (range, cursor, options) => inner().readRows(range, cursor, fix(options)).then(out)
      store.queryStructural = (rect, options) => inner().queryStructural(rect, fix(options)).then(out)
      stores.set(sheetId, store)
      return store
    },
  }
  return {
    source,
    publish() { gen++ },
    /** A broken source: later reads see other values under the SAME stamp. */
    swapSilently() { inners.reverse() },
  }
}

describe('S4.3 review fixes', () => {
  it('an owner-free commit is admitted against the hard limit: evict first, then a typed quota with no state change', async () => {
    const limit = 400_000
    const session = objectSession(sparseDoc([fullCapacitySheet('A', 50, 1), fullCapacitySheet('B', 50, 2)]), { workerBudgetBytes: limit })
    const rows = 200_000
    const big = createRowOrder([{ r0: 0, r1: rows - 1, sourceAtOffset: Uint32Array.from({ length: rows }, (_v, i) => rows - 1 - i) }])
    const draft = { order: big, visibility: visibilityFromHidden(createIntervalSet([])), sort: null, filter: null }
    await expect(session.commitView('A', draft)).rejects.toBeInstanceOf(ViewQuotaError)
    await expect(session.commitView('A', draft)).rejects.toMatchObject({ requiredBytes: expect.any(Number), limitBytes: limit })
    expect(session.memoryStats().totalBytes).toBeLessThanOrEqual(limit)
    expect(session.memoryStats().peakBytes).toBeLessThanOrEqual(limit)
    expect(session.view('A')).toMatchObject({ revision: 0 })
    expect(session.globalViewGen).toBe(0)
    // A commit that fits evicts inactive views first and never exceeds the limit.
    const small = createRowOrder([{ r0: 0, r1: 9_999, sourceAtOffset: Uint32Array.from({ length: 10_000 }, (_v, i) => 9_999 - i) }])
    session.view('B')
    await session.commitView('A', { ...draft, order: small })
    expect(session.memoryStats().peakBytes).toBeLessThanOrEqual(limit)
    await session.release()
  })

  it('touching sheets charges nothing until an operation or view needs a mask; masks are admitted, never forced', async () => {
    const ids = ['A', 'B', 'C', 'D']
    const session = objectSession(sparseDoc(ids.map((id, i) => fullCapacitySheet(id, 50, i))), { workerBudgetBytes: 300_000 })
    for (const id of ids) { session.activate(id); session.identity(id) }
    session.workbookOwner()
    expect(session.memoryStats()).toMatchObject({ viewBytes: 0, retainedBytes: 0 })
    for (const id of ids) {
      session.activate(id)
      expect(session.view(id).revision).toBe(0)
      expect(session.memoryStats().totalBytes).toBeLessThanOrEqual(300_000)
    }
    expect(session.memoryStats().viewBytes).toBeGreaterThan(0)
    expect(session.memoryStats().evictions).toBeGreaterThan(0)
    // With the pool held by a non-evictable ballast, materializing refuses.
    session.activate('A')
    session.memory.trim(0)
    expect(session.isResident('B')).toBe(false)
    session.memory.retain('ballast', { tier: 'view', bytes: session.memory.limitBytes - session.memory.totalBytes, evictable: false })
    expect(() => session.view('B')).toThrow(ViewQuotaError)
    expect(session.isResident('B')).toBe(false)
    expect(session.memoryStats().totalBytes).toBeLessThanOrEqual(300_000)
    await session.release()
  })

  it('a value publication re-derives a sorted order identically whether the view stayed resident or was evicted', async () => {
    const rows = 500
    const docs = [
      sparseDoc([fullCapacitySheet('A', rows, 1), fullCapacitySheet('B', rows, 2)]),
      sparseDoc([fullCapacitySheet('A', rows, 9), fullCapacitySheet('B', rows, 2)]),
    ]
    const run = async (evict: boolean) => {
      const versioned = versionedSource(docs)
      const session = new DocumentSession('doc', versioned.source, { ...fast, filterEvaluator: thirds })
      await committed(session.filter(session.identity('A'), filterOf(rows), 0, live()))
      await committed(session.sort(session.identity('A'), sortOf(rows), live()))
      const before = session.view('A')
      if (evict) {
        session.activate('B')
        session.memory.trim(0)
        expect(session.isResident('A')).toBe(false)
      }
      versioned.publish()
      session.publishValues('A')
      // Same rule on both paths: the applied sort is re-derived from the new values.
      expect(session.isResident('A')).toBe(false)
      expect(() => session.view('A')).toThrow(ViewNotResidentError)
      const rebuilt = await session.ensureView('A', live())
      if (rebuilt.kind !== 'resident') throw new Error(JSON.stringify(rebuilt))
      expect(rebuilt.diagnostics?.map(d => d.code)).toEqual(['view-reapplied'])
      expect(rebuilt.view.revision).toBe(before.revision)
      expect(rebuilt.view.counts).toEqual(before.counts)
      await session.release()
      return { before: orderDigest(before.order), after: orderDigest(rebuilt.view.order) }
    }
    const resident = await run(false)
    const evicted = await run(true)
    expect(resident.after).toBe(evicted.after)
    expect(resident.after).not.toBe(resident.before)
    // Equal to a fresh filter→sort under the new values.
    const fresh = objectSession(docs[1])
    await committed(fresh.filter(fresh.identity('A'), filterOf(rows), 0, live()))
    await committed(fresh.sort(fresh.identity('A'), sortOf(rows), live()))
    expect(orderDigest(fresh.view('A').order)).toBe(resident.after)
    await fresh.release()
  })

  it('activateSheet switches the active sheet only after the view rebuild succeeds', async () => {
    const doc = sparseDoc([fullCapacitySheet('A', 5000, 1), fullCapacitySheet('B', 5000, 2)])
    const kernel = new ViewerKernel({ parse: async () => doc }, { ...fast, filterEvaluator: thirds })
    const replies: WorkerReply[] = []
    const sink = (out: WorkerReply) => { replies.push(out) }
    await kernel.handle({ kind: 'open', requestId: 'o', docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }, sink)
    const session = kernel.session('doc')!
    await kernel.handle({ kind: 'sort', requestId: 's', owner: initialIdentity('doc', 'A'), spec: sortOf(5000, 'desc') }, sink)
    await kernel.handle({ kind: 'activateSheet', requestId: 'b', owner: session.workbookOwner(), sheetId: 'B' }, sink)
    session.memory.trim(0)
    // Fill the pool with non-evictable ballast: A's rebuild cannot be admitted.
    session.memory.retain('ballast', { tier: 'view', bytes: session.memory.limitBytes - session.memory.totalBytes, evictable: false })
    await kernel.handle({ kind: 'activateSheet', requestId: 'a', owner: session.workbookOwner(), sheetId: 'A' }, sink)
    expect(replies.at(-1)).toMatchObject({ requestId: 'a', kind: 'quota', resource: 'memory' })
    expect(session.activeSheetId).toBe('B')
    expect(session.isResident('B')).toBe(true)
    expect(session.memoryStats().totalBytes).toBeLessThanOrEqual(session.memory.limitBytes)
    session.memory.get('ballast')!.release()
    await kernel.handle({ kind: 'activateSheet', requestId: 'a2', owner: session.workbookOwner(), sheetId: 'A' }, sink)
    expect(replies.at(-1)).toMatchObject({ requestId: 'a2', kind: 'complete' })
    expect(session.activeSheetId).toBe('A')
  })
})

describe('S4.3 review round 2', () => {
  it('every rebuild is verified against the committed order digest, not only the first', async () => {
    const rows = 500
    const docs = [
      sparseDoc([fullCapacitySheet('A', rows, 1), fullCapacitySheet('B', rows, 2)]),
      sparseDoc([fullCapacitySheet('A', rows, 9), fullCapacitySheet('B', rows, 2)]),
    ]
    const versioned = versionedSource(docs)
    const session = new DocumentSession('doc', versioned.source, { ...fast, filterEvaluator: thirds })
    await committed(session.sort(session.identity('A'), sortOf(rows), live()))
    session.activate('B')
    for (let i = 0; i < 3; i++) {
      session.memory.trim(0)
      expect((await session.ensureView('A', live())).kind).toBe('resident')
    }
    // Values change WITHOUT a publication (a broken source): the next rebuild
    // must be refused by the digest check, whatever number of rebuilds came first.
    versioned.swapSilently()
    session.memory.trim(0)
    const refused = await session.ensureView('A', live())
    expect(refused).toMatchObject({ kind: 'error', error: { code: 'page-error' } })
    expect(session.isResident('A')).toBe(false)
    await session.release()
  })

  it('refuses a worker budget below one full-axis sheet mask at construction', () => {
    const doc = sparseDoc([fullCapacitySheet('A', 50, 1)])
    expect(MIN_WORKER_BUDGET_BYTES).toBe(135_172)
    expect(() => objectSession(doc, { workerBudgetBytes: MIN_WORKER_BUDGET_BYTES - 1 })).toThrow(WorkerBudgetError)
    const session = objectSession(doc, { workerBudgetBytes: MIN_WORKER_BUDGET_BYTES })
    expect(session.view('A').revision).toBe(0)
    expect(session.memoryStats().totalBytes).toBeLessThanOrEqual(MIN_WORKER_BUDGET_BYTES)
  })

  it('the kernel never reads view() unguarded: activation after a value publication rebuilds via ensureView', async () => {
    expect(readFileSync('src/worker/viewerKernel.ts', 'utf8')).not.toMatch(/\.view\(/)
    const rows = 500
    const docs = [
      sparseDoc([fullCapacitySheet('A', rows, 1), fullCapacitySheet('B', rows, 2)]),
      sparseDoc([fullCapacitySheet('A', rows, 9), fullCapacitySheet('B', rows, 2)]),
    ]
    const versioned = versionedSource(docs)
    const kernel = new ViewerKernel({ parse: async () => docs[0], openSource: async () => versioned.source }, { ...fast, filterEvaluator: thirds })
    const replies: WorkerReply[] = []
    const sink = (out: WorkerReply) => { replies.push(out) }
    await kernel.handle({ kind: 'open', requestId: 'o', docId: 'doc', source: { kind: 'object', document: docs[0] }, profile: 'D' }, sink)
    const session = kernel.session('doc')!
    await kernel.handle({ kind: 'sort', requestId: 's', owner: initialIdentity('doc', 'A'), spec: sortOf(rows) }, sink)
    versioned.publish()
    session.publishValues('A')
    expect(() => session.view('A')).toThrow(ViewNotResidentError)
    await kernel.handle({ kind: 'activateSheet', requestId: 'a', owner: session.workbookOwner(), sheetId: 'A' }, sink)
    expect(replies.at(-1)).toMatchObject({ requestId: 'a', kind: 'complete', diagnostics: [{ code: 'view-reapplied' }] })
    expect(session.isResident('A')).toBe(true)
  })
})
