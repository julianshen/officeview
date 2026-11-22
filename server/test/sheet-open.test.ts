/** Stage 4a Task 3 — XLSX opens in tiles mode build no static paintables; sizes come from metrics (R5). */
import { afterAll, describe, expect, test } from 'bun:test'
import { readdirSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { DocumentHost, PaintablesUnbuilt, RenderRefusal, type OpenResult } from '../src/render'
import { buildXlsx } from '../../src/testdata/ooxml-builders'
import { Pool } from '../src/pool'
import { grid500x100 } from './sheet-fixtures'
import { hiddenFixtures } from '../../scripts/stream-s0/hidden-fixtures'

/** `fontDiagnostics` of a units-mode open of grid500x100, measured on the base ac749246 before this change (F11). */
const BASE_UNITS_FONT_DIAGNOSTICS = 0

const small = () => buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 0 }, { ref: 'B1', v: 42 }] }] }], ['small sheet'])

const hosts: DocumentHost[] = []
const fresh = () => { const h = new DocumentHost(); hosts.push(h); return h }
const pools: Pool[] = []
afterAll(async () => { await Promise.all(pools.map((p) => p.dispose())) })
afterAll(() => { for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k) })

type Disposable = { dispose(): void }
const entryOf = (h: DocumentHost, key: string) => (h as unknown as { docs: Map<string, { paintables?: Disposable }> }).docs.get(key)

describe('xlsx open modes (R5)', () => {
  test('tiles open builds no paintables and renders nothing', async () => {
    const bytes = await grid500x100()
    const h = fresh()
    const t = await h.open('grid', bytes, 'visible-only', { sheetMode: 'tiles' })
    expect(h.paintablesBuilt).toBe(0)
    expect(h.pages).toBe(0)
    const u = await fresh().open('grid', bytes, 'visible-only')
    expect(t.units).toEqual(u.units)
    expect(t.sheets).toEqual(['S'])
    expect(t.fontDiagnostics).toBe(0)
  }, 60_000)

  test('a units path builds paintables once', async () => {
    const job = { unitIndex: 0, scale: 1, projection: 'final-raster' as const, mark: { kind: 'fixed' as const } }
    // grid500x100 (6400 × 10000) is over the unit cap: its render is refused before anything is built.
    const g = fresh(), gt = await g.open('grid', await grid500x100(), 'visible-only', { sheetMode: 'tiles' })
    expect(() => g.render(gt.docKey, job)).toThrow(RenderRefusal)
    expect(g.paintablesBuilt).toBe(0)
    const h = fresh()
    const t = await h.open('small', await small(), 'visible-only', { sheetMode: 'tiles' })
    expect(() => h.render(t.docKey, job)).toThrow(PaintablesUnbuilt)
    expect(h.paintablesBuilt).toBe(0)
    // Two first uses at once share the entry's single build.
    await Promise.all([h.ensurePaintables(t.docKey), h.ensurePaintables(t.docKey)])
    const a = h.render(t.docKey, job)
    await h.ensurePaintables(t.docKey)
    const b = h.render(t.docKey, job)
    expect(a.kind).toBe('raster')
    expect(b.kind).toBe('raster')
    expect(a.w).toBe(t.units[0]!.w)
    expect(h.paintablesBuilt).toBe(1)
    expect(h.pages).toBe(2)
  }, 60_000)

  test('units open stays eager and unchanged', async () => {
    const h = fresh()
    const u = await h.open('grid', await grid500x100(), 'visible-only')
    expect(h.paintablesBuilt).toBe(1)
    expect(u.fontDiagnostics).toBe(BASE_UNITS_FONT_DIAGNOSTICS)
    expect(u.sheets).toEqual(['S'])
    expect(entryOf(h, u.docKey)!.paintables).toBeDefined()
  }, 60_000)

  test('units and tiles opens share one entry', async () => {
    const bytes = await grid500x100()
    const h = fresh()
    const t = await h.open('grid', bytes, 'visible-only', { sheetMode: 'tiles' })
    const u = await h.open('grid', bytes, 'visible-only', { sheetMode: 'units' })
    expect(u.docKey).toBe(t.docKey)
    expect(h.parses).toBe(1)
    expect(h.paintablesBuilt).toBe(1)
    expect(h.docCount).toBe(1)
    const p = entryOf(h, u.docKey)!.paintables!
    let disposed = 0
    const dispose = p.dispose.bind(p)
    p.dispose = () => { disposed++; dispose() }
    h.drop(u.docKey)
    h.drop(u.docKey)
    expect(disposed).toBe(1)
    expect(h.docCount).toBe(0)
  }, 60_000)

  test('hidden sheets are not in sheets', async () => {
    const fixtures = (await hiddenFixtures()).filter((f) => f.id === 'xlsx-hidden-sheet' || f.id === 'xlsx-veryhidden-sheet')
    expect(fixtures.length).toBe(2)
    for (const f of fixtures) {
      for (const sheetMode of ['tiles', 'units'] as const) {
        const o = await fresh().open(f.id, f.withHidden, 'visible-only', { sheetMode })
        expect(o.sheets).toEqual(['S'])
      }
    }
  }, 60_000)

  test('tiles-mode sizes equal units-mode sizes on every corpus xlsx', async () => {
    const dir = resolve(import.meta.dir, '../../corpus')
    const names = readdirSync(dir).filter((n) => n.toLowerCase().endsWith('.xlsx')).sort()
    expect(names.length).toBeGreaterThan(0)
    const outcome = async (o: Promise<OpenResult>) => { try { const r = await o; return { units: r.units, sheets: r.sheets } } catch (e) { return { error: (e as Error).message } } }
    let compared = 0
    for (const name of names) {
      const bytes = new Uint8Array(readFileSync(resolve(dir, name)))
      const th = new DocumentHost(), uh = new DocumentHost()
      const t = await outcome(th.open(name, bytes, 'visible-only', { sheetMode: 'tiles' }))
      const u = await outcome(uh.open(name, bytes, 'visible-only'))
      expect({ name, ...t }).toEqual({ name, ...u })
      if ('units' in t) compared++
      expect(th.paintablesBuilt).toBe(0)
      for (const h of [th, uh]) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
    }
    expect(compared).toBeGreaterThan(0)
  }, 600_000)

  test('pool: a tiles open, then units renders, and a recycled re-open builds lazily', async () => {
    const pool = new Pool({ workers: 1, recyclePages: 1 }); pools.push(pool)
    const bytes = await small()
    const t = await pool.open('small', bytes, 'visible-only', { sheetMode: 'tiles' })
    expect(t.sheets).toEqual(['S'])
    expect(t.fontDiagnostics).toBe(0)
    const job = { docRef: 'small', contentPolicy: 'visible-only' as const, unitIndex: 0, scale: 1, projection: 'final-raster' as const, mark: { kind: 'fixed' as const } }
    const a = await pool.render('s1', job, 1 << 20)
    expect(a.kind).toBe('raster')
    expect(a.w).toBe(t.units[0]!.w)
    expect(pool.stats().recycled).toBe(1)
    // The replacement worker re-opens in tiles mode; the render builds the paintables there.
    const b = await pool.render('s1', job, 1 << 20)
    expect(b.kind).toBe('raster')
    expect(b.w).toBe(a.w)
    expect(pool.stats().recycled).toBe(2)
  }, 60_000)
})
