/**
 * Stage 4a Task 7 — one 512 device-px sheet tile on the pool (S4 §6.5): planned and painted from the servable
 * snapshot only (F16, F19), with the visible mark laid out in global sheet device space (continuous across a seam),
 * full 512² surfaces at the edge (R15), no clock read, and the same admission as `render`.
 */
import { afterAll, describe, expect, spyOn, test } from 'bun:test'
import { DEFAULT_CAPS } from '../src/admission'
import { markOptions, paintMarkRegion } from '../src/mark/visible'
import { decodePng } from '../src/mark/png'
import { AdmissionRefused, Pool } from '../src/pool'
import { DocumentHost, RenderRefusal } from '../src/render'
import type { SheetClock } from '../src/sheet/clock'
import { TILE_BYTES_ESTIMATE, TILE_PALETTE_COLORS, tileBasisLabel, tileLimit, type TileJob } from '../src/sheet/tile'
import { THEME_DEFAULTS } from '../../src/styles/tokens'
import { TILE_ZOOM_STEPS } from '../../src/stream/sheet-wire'
import { tileRangeFor } from '../../src/xlsx/view/tileCache'
import { rejection } from './fixtures/rejection'
import { blankSheet, cfToday, grid500x100, millionRows } from './sheet-fixtures'

const hosts: DocumentHost[] = []
const pools: Pool[] = []
afterAll(async () => {
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
  await Promise.all(pools.map((p) => p.dispose()))
})

const UTC: SheetClock = { day: '2026-11-23', timeZone: 'UTC' }
const T = 512
const job = (o: Partial<TileJob> = {}): TileJob => ({ sheetIndex: 0, viewGen: 0, zoomStep: 5, dpr: 1, palette: 'light', col: 0, row: 0, mark: { kind: 'off' }, clock: UTC, ...o })
async function opened(bytes: Uint8Array) {
  const host = new DocumentHost(); hosts.push(host)
  const o = await host.open('ref', bytes, 'visible-only', { sheetMode: 'tiles' })
  return { host, docKey: o.docKey }
}
const sha = (b: Uint8Array) => new Bun.CryptoHasher('sha256').update(b).digest('hex')
/** Tiles (row-major, `cols` wide) composited into one RGBA surface. */
function composite(tiles: Uint8Array[], cols: number, rows: number): Uint8ClampedArray {
  const W = cols * T, out = new Uint8ClampedArray(W * rows * T * 4)
  tiles.forEach((png, i) => {
    const d = decodePng(png), c = i % cols, r = Math.floor(i / cols)
    expect([d.w, d.h]).toEqual([T, T])
    for (let y = 0; y < T; y++) out.set(d.rgba.subarray(y * T * 4, (y + 1) * T * 4), ((r * T + y) * W + c * T) * 4)
  })
  return out
}
/** Pixels of `a` and `b` (RGBA) that differ. */
function diffCount(a: Uint8ClampedArray, b: Uint8ClampedArray): number {
  expect(a.length).toBe(b.length)
  let n = 0
  for (let i = 0; i < a.length; i += 4) if (a[i] !== b[i] || a[i + 1] !== b[i + 1] || a[i + 2] !== b[i + 2] || a[i + 3] !== b[i + 3]) n++
  return n
}

describe('sheet tiles (S4 §6.5)', () => {
  test('a 512² tile, deterministic', async () => {
    const { host, docKey } = await opened(await grid500x100())
    const a = await host.renderTile(docKey, job())
    const b = await host.renderTile(docKey, job())
    expect([a.w, a.h]).toEqual([T, T])
    const d = decodePng(a.png)
    expect([d.w, d.h]).toEqual([T, T])
    expect(sha(a.png)).toBe(sha(b.png))
    // F24: the palette travels with the result.
    expect(a.palette).toBe('light')
    const blank = await opened(await blankSheet())
    const c = await blank.host.renderTile(blank.docKey, job())
    expect(sha(c.png)).not.toBe(sha(a.png))
  })

  const warmUp: Array<Partial<TileJob>> = [{ col: 1 }, {}, { zoomStep: 1, dpr: 2 }, { zoomStep: 8, col: 2, row: 3 }, { zoomStep: 0, mark: { kind: 'viewer', text: 'someone else' } }, { palette: 'dark' }, { zoomStep: 10, dpr: 2 }]
  const cfDay = { clock: { day: '2026-11-22', timeZone: 'UTC' } }
  // One test per workbook, each with the heavy-test budget (PR #104: 5.48 s combined, then 5.97 s for grid500x100 alone, on CI's Linux runner).
  for (const [name, make, targets] of [
    ['grid500x100', grid500x100, [{ mark: { kind: 'fixed' } }, { zoomStep: 1, dpr: 2, mark: { kind: 'fixed' } }, { zoomStep: 9, dpr: 2, col: 1 }]],
    ['cfToday', cfToday, [{ ...cfDay, mark: { kind: 'fixed' } }, { ...cfDay, zoomStep: 1, dpr: 2 }, { ...cfDay, zoomStep: 9 }]],
  ] as Array<[string, () => Promise<Uint8Array>, Array<Partial<TileJob>>]>) {
    test(`a tile's bytes do not depend on what its host painted before (shared layout cache, mark state): ${name}`, async () => {
      const bytes = await make()
      const warm = await opened(bytes)
      for (const o of warmUp) await warm.host.renderTile(warm.docKey, job({ ...o, ...('clock' in targets[0]! ? cfDay : {}) }))
      for (const t of targets) {
        // Cold: a fresh host (as a tracer rebuilds a reference); warm: after the other tiles above.
        const cold = await opened(bytes)
        const first = await cold.host.renderTile(cold.docKey, job(t))
        const later = await warm.host.renderTile(warm.docKey, job(t))
        expect([JSON.stringify(t), sha(later.png)]).toEqual([JSON.stringify(t), sha(first.png)])
      }
    }, 30_000)
  }

  test('the tile bound is tileRangeFor\'s last + 1 (F8)', () => {
    for (const [extent, zoom, dpr] of [[1_048_576, 0.33, 1], [1_048_576, 1, 1], [20_971_520, 4, 2], [100, 0.25, 1], [0, 1, 1]] as const) {
      const r = tileRangeFor({ origin: { x: 0, y: 0 }, viewport: { width: 1e12, height: 1 }, zoom, dpr, extent: { width: extent, height: 1 } })
      expect(tileLimit(extent, zoom, dpr)).toBe(r.c1 + 1)
    }
  })

  test('edge tiles are full surfaces', async () => {
    // R15.
    const { host, docKey } = await opened(await grid500x100())
    // F8: the last column tile at zoomStep 1 (0.33) is partial: 1,048,576 CSS px × 0.33 = 346,030.08 device px.
    const zoom = TILE_ZOOM_STEPS[1]!, extentDevice = 16_384 * 64 * zoom
    const last = tileLimit(16_384 * 64, zoom, 1) - 1
    expect(last).toBe(675)
    const inside = Math.floor(extentDevice - last * T)
    expect(inside).toBeGreaterThan(0)
    expect(inside).toBeLessThan(T - 2)
    const t = await host.renderTile(docKey, job({ zoomStep: 1, col: last }))
    const d = decodePng(t.png)
    expect([d.w, d.h]).toEqual([T, T])
    let beyond = 0, inked = 0
    for (let y = 0; y < T; y++) for (let x = 0; x < T; x++) {
      const i = (y * T + x) * 4, white = d.rgba[i] === 255 && d.rgba[i + 1] === 255 && d.rgba[i + 2] === 255 && d.rgba[i + 3] === 255
      if (x > inside + 1 && !white) beyond++
      if (x < inside - 1 && !white) inked++
    }
    expect(beyond).toBe(0)
    expect(inked).toBeGreaterThan(0)
    // One past the last column is refused.
    const past = await rejection(host.renderTile(docKey, job({ zoomStep: 1, col: last + 1 })))
    expect(past).toBeInstanceOf(RenderRefusal)
    expect((past as RenderRefusal).code).toBe('bad-job')
  })

  test('the visible mark is continuous across a seam', async () => {
    // F9: zoomStep 5, dpr 2. The reference is the unmarked blank tiles composited (a blank sheet has gridlines), then
    // one `paintMarkRegion` over the whole 1024×512 (or 512×1024) region.
    const { host, docKey } = await opened(await blankSheet())
    const scale = TILE_ZOOM_STEPS[5]! * 2
    const tile = async (col: number, row: number, mark: TileJob['mark']) => (await host.renderTile(docKey, job({ dpr: 2, col, row, mark }))).png
    const mark = markOptions({ kind: 'fixed' })!
    for (const [cols, rows, coords] of [[2, 1, [[0, 0], [1, 0]]], [1, 2, [[0, 0], [0, 1]]]] as const) {
      const marked = composite(await Promise.all(coords.map(([c, r]) => tile(c, r, { kind: 'fixed' }))), cols, rows)
      // The reference: the unmarked blank tiles on one surface, then one mark over the whole region.
      const base = composite(await Promise.all(coords.map(([c, r]) => tile(c, r, { kind: 'off' }))), cols, rows)
      const W = cols * T, H = rows * T
      const canvas = host.backend.createCanvas(W, H)
      const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
      const img = ctx.createImageData(W, H)
      img.data.set(base)
      ctx.putImageData(img, 0, 0)
      paintMarkRegion(ctx, mark, { x: 0, y: 0, w: W, h: H, scale })
      const reference = ctx.getImageData(0, 0, W, H).data
      expect(diffCount(base, reference)).toBeGreaterThan(1000)
      expect(diffCount(marked, reference)).toBe(0)
    }
  })

  test('bad jobs are refused', async () => {
    const { host, docKey } = await opened(await grid500x100())
    const rowsPast = tileLimit(1_048_576 * 20, 1, 1)
    for (const bad of [{ zoomStep: 11 }, { dpr: 3 as 1 }, { col: -1 }, { row: rowsPast }, { sheetIndex: 1 }, { viewGen: 1 }, { zoomStep: 1.5 }, { palette: 'sepia' as 'light' }]) {
      const e = await rejection(host.renderTile(docKey, job(bad)))
      expect(e).toBeInstanceOf(RenderRefusal)
      expect([JSON.stringify(bad), (e as RenderRefusal).code]).toEqual([JSON.stringify(bad), 'bad-job'])
    }
    // The last row tile is in range.
    expect((await host.renderTile(docKey, job({ row: rowsPast - 1 }))).w).toBe(T)
  })

  test('a tile reads no clock', async () => {
    const { host, docKey } = await opened(await cfToday())
    const spy = spyOn(Date, 'now').mockImplementation(() => { throw new Error('clock') })
    try {
      // A4 is 2026-11-22: filled under that day, unfilled on 2026-12-25.
      const today = await host.renderTile(docKey, job({ clock: { day: '2026-11-22', timeZone: 'UTC' } }))
      const other = await host.renderTile(docKey, job({ clock: { day: '2026-12-25', timeZone: 'UTC' } }))
      expect(sha(today.png)).not.toBe(sha(other.png))
    } finally { spy.mockRestore() }
  })

  test('the million-row last row renders', async () => {
    const { host, docKey } = await opened(await millionRows())
    // Rows 10–20 are 40 px, every other row 20 px.
    const y = 1_048_575 * 20 + 11 * 20
    const row = Math.floor(y / T), localY = y - row * T
    const t = await host.renderTile(docKey, job({ row }))
    const d = decodePng(t.png)
    expect([d.w, d.h]).toEqual([T, T])
    // Text ink (dark pixels) in the last row's band of column A, where `LAST-ROW` is painted.
    let dark = 0
    for (let yy = localY; yy < Math.min(T, localY + 20); yy++) for (let x = 0; x < 64; x++) { const i = (yy * T + x) * 4; if (d.rgba[i]! < 128) dark++ }
    expect(dark).toBeGreaterThan(10)
  })

  test('palettes come from the default theme tokens; the basis label is the Task 1 spike\'s', () => {
    expect(TILE_PALETTE_COLORS.light).toEqual({ pending: THEME_DEFAULTS.light.pending, 'pending-text': THEME_DEFAULTS.light['pending-text'] })
    expect(TILE_PALETTE_COLORS.dark).toEqual({ pending: THEME_DEFAULTS.dark.pending, 'pending-text': THEME_DEFAULTS.dark['pending-text'] })
    expect(tileBasisLabel('0', 1, 2, 675, 3)).toBe('sheet:0@0.33x2:675,3')
    expect(tileBasisLabel('2', 5, 1, 0, 0)).toBe('sheet:2@1x1:0,0')
    expect(TILE_BYTES_ESTIMATE).toBe(512 * 512 * 4 * 2)
  })

  test('pool.tile goes through admission', async () => {
    const pool = new Pool({ workers: 1, caps: { ...DEFAULT_CAPS, perSessionOutstanding: 1, perSessionQueue: 1 }, testHooks: true }); pools.push(pool)
    await pool.open('grid', await grid500x100(), 'visible-only', { sheetMode: 'tiles' })
    const tj = (o: Partial<TileJob> & { hangMs?: number } = {}) => ({ docRef: 'grid', contentPolicy: 'visible-only' as const, ...job(o) })
    const running = pool.tile('s1', tj({ hangMs: 600 }), TILE_BYTES_ESTIMATE)
    const queued = pool.tile('s1', tj(), TILE_BYTES_ESTIMATE)
    expect(pool.stats()).toMatchObject({ running: 1, queued: 1 })
    expect(await rejection(pool.tile('s1', tj(), TILE_BYTES_ESTIMATE))).toEqual(new AdmissionRefused('session-queue'))
    const [a, b] = await Promise.all([running, queued])
    expect([a.w, a.h, decodePng(a.png).w]).toEqual([T, T, T])
    expect(a.palette).toBe('light')
    expect(sha(a.png)).toBe(sha(b.png))
    expect(pool.stats()).toMatchObject({ running: 0, queued: 0, reserved: 0 })
    // A refused job reaches the caller as the worker's refusal.
    const e = await rejection(pool.tile('s1', tj({ viewGen: 1 }), TILE_BYTES_ESTIMATE))
    expect((e as { refusal?: string }).refusal).toBe('bad-job')
  }, 30_000)
})
