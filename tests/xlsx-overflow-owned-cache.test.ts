import { describe, expect, it, vi } from 'vitest'
import { spawnSync } from 'node:child_process'
import type { XlsxCell } from '../src/xlsx/types'
import { OverflowCache, OverflowResourceError } from '../src/xlsx/view/overflowCache'
import { MemoryPool } from '../src/xlsx/view/tileCache'
import { RectangleIndex } from '../src/xlsx/view/structure'
import { buildSheetGeometry, type SheetGeometry } from '../src/xlsx/view/viewport'
import { rowOverflowFromCells, textOverflowAllowance } from '../src/xlsx/view/textOverflow'

function geometry(cellsOf: (row: number) => readonly XlsxCell[] | undefined): SheetGeometry {
  const built = buildSheetGeometry({ name: 's', rows: [], cols: [], merges: [], mergeRanges: [] })
  if (built.kind !== 'ready') throw Error('fixture geometry')
  return { ...built, cellsOf }
}
function cells(count: number): XlsxCell[] {
  return Array.from({ length: count }, (_, col) => ({ row: 0, col, ref: '', styleIndex: 0, value: 1 }))
}
const poolOf = (limitBytes: number) => new MemoryPool({ limitBytes, highWatermark: 1, lowWatermark: 1 })

describe('document-owned overflow occupancy', () => {
  it('charges real aggregate pool pressure across two sheets and evicts exact rebuildable rows', () => {
    const pool = poolOf(12000), cache = new OverflowCache(pool, 12000)
    const row = cells(1024), a = geometry(() => row), b = geometry(() => row)
    const first = cache.get(a, 0)
    const allowance = textOverflowAllowance(first, 500)
    expect(pool.stats().retainedBytes).toBeGreaterThanOrEqual(2048)
    expect(first.blockers).toBeInstanceOf(Uint16Array)
    cache.get(b, 0)
    cache.get(b, 1)
    expect(pool.totalBytes).toBeLessThanOrEqual(12000)
    expect(cache.get(a, 0)).not.toBe(first)
    expect(textOverflowAllowance(cache.get(a, 0), 500)).toEqual(allowance)
  })

  it('checks and touches a stable snapshot in constant ownership work', () => {
    const cache = new OverflowCache(poolOf(200000))
    const row = cells(2), g = geometry(() => row)
    for (let r = 0; r < 100; r++) cache.get(g, r)
    const input = cache.get(g, 0)
    const deref = vi.spyOn(WeakRef.prototype, 'deref')
    deref.mockClear()
    expect(cache.get(g, 0)).toBe(input)
    expect(deref.mock.calls.length).toBeLessThanOrEqual(4)
    deref.mockRestore()
  })

  it('does not retain dead geometry or its real permutation metrics through pool callbacks', () => {
    const run = spawnSync('bun', ['tests/probes/xlsx-overflow-owner-probe.ts'], { encoding: 'utf8', timeout: 30000 })
    expect(run.status, run.stderr).toBe(0)
    expect(JSON.parse(run.stdout)).toEqual({ geometries: 0, metrics: 0, bytes: 0, pool: 0, ledger: 0 })
  })

  it.each([1000, 5000])('covers the actual retained ownership overhead of %i tiny rows', count => {
    const run = spawnSync('bun', ['tests/probes/xlsx-overflow-budget-probe.ts', String(count)], { encoding: 'utf8', timeout: 30000 })
    expect(run.status, run.stderr).toBe(0)
    const measured = JSON.parse(run.stdout) as { growth: number; charged: number }
    expect(measured.growth).toBeLessThanOrEqual(measured.charged)
  })

  it('covers retained heap and backing buffers under the full-width byte cap', () => {
    const run = spawnSync('bun', ['tests/probes/xlsx-overflow-budget-probe.ts', '512', '16384'], { encoding: 'utf8', timeout: 30000 })
    expect(run.status, run.stderr).toBe(0)
    const measured = JSON.parse(run.stdout) as { growth: number; charged: number }
    expect(measured.growth).toBeLessThanOrEqual(measured.charged)
    expect(measured.charged).toBeLessThanOrEqual(2 * 1024 * 1024)
  })

  it('does not reuse a dropped row after its weak snapshot has been collected', () => {
    const run = spawnSync('bun', ['tests/probes/xlsx-overflow-owner-probe.ts', 'drop'], { encoding: 'utf8', timeout: 30000 })
    expect(run.status, run.stderr).toBe(0)
    expect(JSON.parse(run.stdout)).toEqual({ blockers: 0 })
  })

  it('admits one intersecting merge without reserving every off-row merge in the document', () => {
    const cache = new OverflowCache(poolOf(5000))
    const g = { ...geometry(() => []), merges: RectangleIndex.build(Array.from({ length: 1000 }, (_, row) => ({ rect: { r0: row, r1: row, c0: 0, c1: 1 }, value: { minRow: row, maxRow: row, minCol: 0, maxCol: 1 } }))) }
    expect(textOverflowAllowance(cache.get(g, 0), 2).left).toBe(0)
  })

  it('rebuilds exact hidden, duplicate, edge-column, and merged allowances after eviction', () => {
    const row = [16383, 2, 2, 3, 5].map(col => ({ row: 0, col, ref: '', styleIndex: 0, value: '' }))
    const built = geometry(() => row)
    const colStart = (col: number) => col * 64 - (col > 3 ? 64 : 0)
    const ranges = [{ minRow: 0, maxRow: 0, minCol: 0, maxCol: 1 }, { minRow: 0, maxRow: 0, minCol: 3, maxCol: 3 }]
    const g = { ...built, metrics: { ...built.metrics, cols: { ...built.metrics.cols, start: colStart } }, merges: RectangleIndex.build(ranges.map(value => ({ rect: { r0: 0, r1: 0, c0: value.minCol, c1: value.maxCol }, value }))) }
    const cache = new OverflowCache(poolOf(8192)), first = cache.get(g, 0)
    expect(Array.from(first.blockers)).toEqual([2, 5, 16383])
    const expected = rowOverflowFromCells(row, ranges, colStart, 16383)
    cache.pool.trim(0)
    const rebuilt = cache.get(g, 0)
    expect(rebuilt).not.toBe(first)
    for (const col of [0, 1, 2, 3, 4, 5, 10, 16383]) expect(textOverflowAllowance(rebuilt, col)).toEqual(textOverflowAllowance(expected, col))
  })

  it('bounds all 512 full-width generated source rows across two sheet geometries', () => {
    const limit = 2 * 1024 * 1024, pool = poolOf(limit), cache = new OverflowCache(pool, limit)
    const generated = (row: number) => cells(16384).map(cell => ({ ...cell, row }))
    const sheets = [geometry(generated), geometry(generated)]
    for (let row = 0; row < 512; row++) {
      const input = cache.get(sheets[row % 2], row)
      expect(input.blockers.length).toBe(16384)
      expect(input.blockers[input.blockers.length - 1]).toBe(16383)
      expect(cache.bytes).toBe(pool.totalBytes)
      expect(pool.totalBytes).toBeLessThanOrEqual(limit)
    }
    expect(pool.stats().entries).toBeGreaterThan(0)
    expect(pool.stats().entries).toBeLessThan(512)
    cache.close()
    expect(pool.totalBytes).toBe(0)
  })

  it('invalidates stable snapshots on real pool eviction and close; an owner cannot revive', () => {
    const pool = poolOf(8192), cache = new OverflowCache(pool)
    const g = geometry(() => cells(16)), initial = cache.get(g, 0)
    pool.trim(0)
    expect(cache.bytes).toBe(0)
    expect(cache.get(g, 0)).not.toBe(initial)
    pool.close()
    expect(cache.bytes).toBe(0)
    expect(() => cache.get(g, 0)).toThrow(OverflowResourceError)
    cache.close()
    expect(pool.totalBytes).toBe(0)
    const own = new OverflowCache(poolOf(8192))
    own.get(g, 0)
    own.close(); own.close()
    expect(own.pool.totalBytes).toBe(0)
    expect(() => own.get(g, 0)).toThrow(OverflowResourceError)
  })

  it('reserves the retained backing buffer plus sort/merge coexistence before source reads', () => {
    const shared = poolOf(12000), charge = shared.charge(), pool = poolOf(12000)
    let held = 0, peakAtRead = 0
    const cache = new OverflowCache(pool, 12000, {
      reserve(bytes) { if (!charge.set(held + bytes)) return false; held += bytes; return true },
      release(bytes) { held -= bytes; charge.set(held) },
    })
    const row = new Proxy(cells(1024), { get(target, key, receiver) {
      if (typeof key === 'string' && /^\d+$/.test(key)) peakAtRead = shared.totalBytes
      return Reflect.get(target, key, receiver)
    } })
    const input = cache.get(geometry(() => row), 0)
    expect(peakAtRead).toBeGreaterThanOrEqual(input.blockers instanceof Uint16Array ? input.blockers.buffer.byteLength * 2 : Infinity)
    expect(shared.totalBytes).toBe(cache.bytes)
    expect(shared.stats().peakBytes).toBeGreaterThan(shared.totalBytes)
    cache.close()
    expect(shared.totalBytes).toBe(0)
  })

  it('uses the real shared accounting pool for external pressure and explicit refusal', () => {
    const shared = poolOf(16000), charge = shared.charge(), pool = poolOf(64000)
    let held = 0, released = 0
    const cache = new OverflowCache(pool, 64000, {
      reserve(bytes) { if (!charge.set(held + bytes)) return false; held += bytes; return true },
      release(bytes) { held -= bytes; charge.set(held) },
    })
    shared.retain('other-sheet', { tier: 'rebuildable', bytes: 12000, onRelease: () => { released++ } })
    const row = cells(1024), g = geometry(() => row)
    const first = cache.get(g, 0)
    expect(released).toBe(1)
    const fixed = shared.retain('fixed', { tier: 'view', bytes: 9000, evictable: false })!
    expect(fixed.live).toBe(true)
    expect(() => cache.get(g, 1)).toThrow(OverflowResourceError)
    expect(cache.bytes).toBe(0)
    expect(shared.totalBytes).toBe(9000)
    fixed.release()
    expect(cache.get(g, 0)).not.toBe(first)
    expect(textOverflowAllowance(cache.get(g, 0), 500)).toEqual(textOverflowAllowance(first, 500))
    cache.close()
    expect(shared.totalBytes).toBe(0)
  })

  it('treats the byte subcap as retained storage while construction uses admitted shared-pool room', () => {
    const pool = poolOf(12000), cache = new OverflowCache(pool, 4300)
    const input = cache.get(geometry(() => cells(64)), 0)
    expect(input.blockers.length).toBe(64)
    expect(cache.bytes).toBeLessThanOrEqual(4300)
    expect(pool.stats().peakBytes).toBeGreaterThan(4300)
    expect(pool.stats().peakBytes).toBeLessThanOrEqual(12000)
  })

  it('bounds blocker construction by source columns even for an overfull repair row of duplicates', () => {
    const row = Array.from({ length: 40000 }, () => ({ row: 0, col: 2, ref: 'C1', styleIndex: 0, value: 'text' }))
    const pool = poolOf(100000), cache = new OverflowCache(pool)
    const input = cache.get(geometry(() => row), 0)
    expect(Array.from(input.blockers)).toEqual([2])
    expect(textOverflowAllowance(input, 2).left).toBe(128)
    expect((input.blockers as Uint16Array).buffer.byteLength).toBeLessThanOrEqual(16384 * 2)
    expect(pool.stats().peakBytes).toBeLessThanOrEqual(100000)
  })

  it('keeps compact lossless merge endpoints and exact overlapping-merge semantics within the retained subcap', () => {
    const ranges = Array.from({ length: 20000 }, () => ({ minRow: 0, maxRow: 0, minCol: 100, maxCol: 200 }))
    const row = cells(4), built = geometry(() => row)
    const g = { ...built, merges: RectangleIndex.build(ranges.map(value => ({ rect: { r0: 0, r1: 0, c0: 100, c1: 200 }, value }))) }
    const pool = poolOf(1000000), cache = new OverflowCache(pool, 100000)
    const input = cache.get(g, 0)
    expect(input.merges).toBeInstanceOf(Uint16Array)
    expect(cache.bytes).toBeLessThanOrEqual(100000)
    const expected = rowOverflowFromCells(row, ranges, col => g.metrics.cols.start(col), 16383)
    for (const col of [0, 3, 50, 99, 100, 150, 200, 201, 250, 16383]) expect(textOverflowAllowance(input, col)).toEqual(textOverflowAllowance(expected, col))
    cache.pool.trim(0)
    expect(textOverflowAllowance(cache.get(g, 0), 250)).toEqual(textOverflowAllowance(expected, 250))
  })

  it('reuses the stable replacement snapshot after pressure evicts the geometry last row', () => {
    const row = cells(1024), g = geometry(() => row), cache = new OverflowCache(poolOf(12000), 12000)
    cache.get(g, 0)
    const input = cache.get(g, 1)
    expect(cache.get(g, 1)).toBe(input)
    expect(cache.pool.totalBytes).toBe(cache.bytes)
  })

  it('charges the full compact merge backing capacity, including invisible endpoints', () => {
    const row: XlsxCell[] = [], base = geometry(() => row)
    const colStart = (col: number) => col * 64 - (col > 2 ? 64 : 0)
    const plain = { ...base, metrics: { ...base.metrics, cols: { ...base.metrics.cols, start: colStart } } }
    const ranges = [0, 2, 5].map(col => ({ minRow: 0, maxRow: 0, minCol: col, maxCol: col }))
    const merged = { ...plain, merges: RectangleIndex.build(ranges.map(value => ({ rect: { r0: 0, r1: 0, c0: value.minCol, c1: value.maxCol }, value }))) }
    const cache = new OverflowCache(poolOf(20000))
    cache.get(plain, 0)
    const overhead = cache.bytes
    const input = cache.get(merged, 0)
    expect(input.merges).toBeInstanceOf(Uint16Array)
    const endpoints = input.merges as Uint16Array
    expect(endpoints.length).toBe(4)
    expect(endpoints.buffer.byteLength).toBe(3 * 2 * Uint16Array.BYTES_PER_ELEMENT)
    expect(cache.bytes - overhead * 2).toBe(endpoints.buffer.byteLength)
    expect(cache.pool.stats().retainedBytes).toBe(cache.bytes)
  })

  it('a real cache miss releases dead owners even when the finalizer supplies no cleanup', () => {
    const run = spawnSync('bun', ['tests/probes/xlsx-overflow-sweep-probe.ts'], { encoding: 'utf8', timeout: 30000 })
    expect(run.status, run.stderr).toBe(0)
    const measured = JSON.parse(run.stdout)
    expect(measured.deadGeometry).toBe(0)
    expect(measured.deadMetrics).toBe(0)
    expect(measured.before.entries).toBe(8)
    expect(measured.before.bytes).toBeGreaterThan(0)
    expect(measured.after.entries).toBe(1)
    expect(measured.after.bytes).toBeLessThan(measured.before.bytes)
    expect(measured.after.main).toBe(measured.after.bytes)
    expect(measured.after.pool).toBe(measured.after.bytes)
    expect(measured.closed).toEqual({ entries: 0, bytes: 0, pool: 0, main: 0 })
  })

  it('refuses before reading or allocating an oversized row or merge array', () => {
    const pool = poolOf(4300), cache = new OverflowCache(pool)
    let reads = 0
    const row = new Proxy(cells(16384), { get(target, key, receiver) {
      if (typeof key === 'string' && /^\d+$/.test(key)) reads++
      return Reflect.get(target, key, receiver)
    } })
    expect(() => cache.get(geometry(() => row), 0)).toThrow(OverflowResourceError)
    expect(reads).toBe(0)
    const g = { ...geometry(() => []), merges: RectangleIndex.build(Array.from({ length: 20 }, (_, col) => ({ rect: { r0: 0, r1: 0, c0: col, c1: col }, value: { minRow: 0, maxRow: 0, minCol: col, maxCol: col } }))) }
    expect(() => cache.get(g, 0)).toThrow(OverflowResourceError)
    expect(pool.totalBytes).toBe(0)
  })
})
