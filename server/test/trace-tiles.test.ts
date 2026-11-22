/**
 * Stage 4a Task 10 — `traceTiles` (R10): a pane capture of tiled sheet tiles traced through the R10 composite. The
 * named window's tiles are rebuilt through `DocumentHost.renderTile` under each sid's stored view and clock, hashed
 * against every marked `tile_deliveries` row, composited, registered once, and decided with τ for the sid count.
 *
 * Owner ruling (2026-11-23): blank tiles stay marked. R15's "flat tiles may abstain" becomes "a marked blank pane
 * never names a wrong session": the right sid or an abstain both pass.
 */
import { afterAll, describe, expect, setSystemTime, test } from 'bun:test'
import { Database } from 'bun:sqlite'
import { createCanvas } from '@napi-rs/canvas'
import { createHash } from 'node:crypto'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { TILE_PX } from '../../src/stream/sheet-wire'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import { decodeImage } from '../src/detect/image'
import { traceTiles, type TraceAnswer, type TraceObservation } from '../src/detect/trace'
import { deriveKeys, docRefOf, type Keys } from '../src/keys'
import { decodePng } from '../src/mark/png'
import type { Bitmap } from '../src/mark/types'
import { DocumentHost } from '../src/render'
import { Client, disposePool, start, type Harness } from './harness'
import { rejection } from './fixtures/rejection'
import { blankSheet, cfToday, grid500x100 } from './sheet-fixtures'

const dirs: string[] = []
afterAll(async () => { await disposePool(); for (const d of dirs.splice(0)) rmSync(d, { recursive: true, force: true }) })
const dbPath = () => { const d = mkdtempSync(join(tmpdir(), 'ov-trace-tiles-')); dirs.push(d); return join(d, 'audit.db') }

const master = new Uint8Array(32).fill(9)
const keys1 = await deriveKeys(master, 1)
const ring: ReadonlyMap<number, Keys> = new Map([[1, keys1]])
const WINDOW = Object.freeze({ c0: 0, c1: 1, r0: 0, r1: 1 })
const ZOOM_STEP = 5 // zoom 1
const sha = (b: Uint8Array) => createHash('sha256').update(b).digest('hex')

interface Served { h: Harness; store: AuditStore; path: string; sids: string[]; tiles: Map<string, Map<string, Uint8Array>> }
/** Three forensic sids (`fixed` visible mark) each fetch the 2×2 window at (zoom 1, dpr 1, light). */
async function serve(doc: Uint8Array, now?: () => number): Promise<Served> {
  const path = dbPath(), store = AuditStore.open(path)
  const h = await start({ doc, keys: keys1, policy: { visibleMark: 'fixed', forensicMark: true }, audit: new AuditTable({ store }), ...(now ? { now } : {}) })
  const sids: string[] = [], tiles = new Map<string, Map<string, Uint8Array>>()
  for (let i = 0; i < 3; i++) {
    const c = new Client(h.url)
    const m = await c.hello({ features: ['sheet-tiles@1'] }); await c.ready()
    const want = [{ col: 0, row: 0 }, { col: 1, row: 0 }, { col: 0, row: 1 }, { col: 1, row: 1 }]
    const r = await c.tiles(10, { sheetId: '0', viewGen: 0, zoomStep: ZOOM_STEP, dpr: 1, palette: 'light', tiles: want }, { ms: 60_000 })
    if (r.end.type !== 'tiles-done' || r.end.refused.length || r.tiles.length !== 4) throw new Error(`tiles: ${JSON.stringify(r.end)}`)
    const own = new Map<string, Uint8Array>()
    for (const t of r.tiles) own.set(`${t.col},${t.row}`, t.png)
    sids.push(m.session.sid); tiles.set(m.session.sid, own)
    c.close()
  }
  return { h, store, path, sids, tiles }
}

/** The sid's four delivered (marked) tiles composited, cropped to 900×900 at (40, 40). */
function paneCapture(own: Map<string, Uint8Array>): Bitmap {
  const W = 2 * TILE_PX, full = new Uint8ClampedArray(W * W * 4)
  for (let r = 0; r < 2; r++) for (let c = 0; c < 2; c++) {
    const d = decodePng(own.get(`${c},${r}`)!)
    for (let y = 0; y < TILE_PX; y++) full.set(d.rgba.subarray(y * TILE_PX * 4, (y + 1) * TILE_PX * 4), ((r * TILE_PX + y) * W + c * TILE_PX) * 4)
  }
  const x = 40, y = 40, w = 900, h = 900, data = new Uint8ClampedArray(w * h * 4)
  for (let row = 0; row < h; row++) data.set(full.subarray(((y + row) * W + x) * 4, ((y + row) * W + x + w) * 4), row * w * 4)
  return { width: w, height: h, data }
}
async function jpeg(b: Bitmap, q: number): Promise<Bitmap> {
  const cv = createCanvas(b.width, b.height), ctx = cv.getContext('2d'), img = ctx.createImageData(b.width, b.height)
  img.data.set(b.data); ctx.putImageData(img, 0, 0)
  return decodeImage(new Uint8Array(cv.encodeSync('jpeg', q)))
}
const host = new DocumentHost()
const traceOf = (s: Served, doc: Uint8Array, capture: Bitmap, observe?: (o: TraceObservation) => void): Promise<TraceAnswer> =>
  traceTiles({ capture, doc, sheetId: '0', zoomStep: ZOOM_STEP, dpr: 1, window: WINDOW, keys: ring, store: AuditStore.openReadOnlyForTrace(s.path), host, ...(observe ? { observe } : {}) })
const brief = (a: TraceAnswer) => a.kind === 'traced' ? { kind: a.kind, sid: a.sid, z: +a.z.toFixed(2), margin: +a.margin.toFixed(2) } : a

describe('traceTiles (R10)', () => {
  let grid: Uint8Array, s: Served
  test('a one-pane capture traces to its sid', async () => {
    grid = await grid500x100()
    s = await serve(grid)
    const B = s.sids[1]!
    const capture = paneCapture(s.tiles.get(B)!)
    let seen: TraceObservation | undefined
    const lossless = await traceOf(s, grid, capture, (o) => { seen = o })
    console.log('[task10] grid lossless', JSON.stringify(brief(lossless)))
    expect(lossless).toMatchObject({ kind: 'traced', sid: B, target: 'tiles' })
    // One candidate per sid (not per tile row): three scores, τ for three.
    expect(seen!.sids.slice().sort()).toEqual(s.sids.slice().sort())
    expect(seen!.tau.m).toBe(2 ** 16)
    const q90 = await traceOf(s, grid, await jpeg(capture, 90))
    console.log('[task10] grid jpeg-q90', JSON.stringify(brief(q90)))
    expect(q90).toMatchObject({ kind: 'traced', sid: B })
  }, 120_000)

  test('a palette with no deliveries has no candidates', async () => {
    const a = await traceTiles({ capture: paneCapture(s.tiles.get(s.sids[0]!)!), doc: grid, sheetId: '0', zoomStep: ZOOM_STEP, dpr: 1, palette: 'dark', window: WINDOW, keys: ring, store: AuditStore.openReadOnlyForTrace(s.path), host })
    expect(a).toEqual({ kind: 'no-signal', reason: 'no-candidates' })
  }, 60_000)

  test('a tampered reference fails closed', async () => {
    const db = new Database(s.path)
    db.query("UPDATE tile_deliveries SET ref_sha256 = ? WHERE rowid = (SELECT rowid FROM tile_deliveries WHERE sid = ? AND col = 1 AND row = 1)").run('0'.repeat(64), s.sids[2]!)
    db.close()
    expect(await traceOf(s, grid, paneCapture(s.tiles.get(s.sids[1]!)!))).toEqual({ kind: 'no-signal', reason: 'reference-mismatch' })
  }, 60_000)

  test('a marked row without its sheet view fails closed', async () => {
    const db = new Database(s.path)
    db.query('DELETE FROM sheet_views WHERE sid = ?').run(s.sids[0]!)
    db.close()
    const e = await rejection(traceOf(s, grid, paneCapture(s.tiles.get(s.sids[1]!)!)))
    expect(String(e)).toContain('sheet view')
    await s.h.stop()
  }, 60_000)

  test('a marked blank pane never names a wrong session', async () => {
    const blank = await blankSheet()
    const b = await serve(blank)
    const B = b.sids[1]!
    // Blank tiles are marked (owner ruling): the pane is not vacuously unattributable.
    for (const sid of b.sids) expect(b.store.tileDeliveries(sid).map((d) => d.status)).toEqual(['l1', 'l1', 'l1', 'l1'])
    const a = await traceOf(b, blank, paneCapture(b.tiles.get(B)!))
    console.log('[task10] blank lossless', JSON.stringify(brief(a)))
    const ok = (a.kind === 'traced' && a.sid === B) || (a.kind === 'no-signal' && (a.reason === 'flat' || a.reason === 'geometry'))
    expect([ok, a]).toEqual([true, a])
    await b.h.stop()
  }, 120_000)

  test('the rebuild uses the stored clock', async () => {
    // Day D = 2026-11-21 (cfToday's A3), not today's date; the tracer runs on D+1.
    const D = Date.UTC(2026, 10, 21, 12, 0, 0)
    const doc = await cfToday()
    const c = await serve(doc, () => D)
    expect(c.store.sheetView(c.sids[0]!, '0', 0)!.clock).toEqual({ day: '2026-11-21', timeZone: 'UTC' })
    const B = c.sids[1]!
    setSystemTime(new Date(D + 86_400_000))
    try {
      const a = await traceOf(c, doc, paneCapture(c.tiles.get(B)!))
      console.log('[task10] cfToday lossless (clock D, tracer D+1)', JSON.stringify(brief(a)))
      expect(a).toMatchObject({ kind: 'traced', sid: B })
    } finally { setSystemTime() }
    // Discriminating: the D+1 rebuild of tile (0,0) is not what was served.
    const { docKey } = await host.open(await docRefOf(keys1.kDoc, doc), doc, c.store.getRow(B)!.contentPolicy, { sheetMode: 'tiles' })
    const job = { sheetIndex: 0, viewGen: 0, zoomStep: ZOOM_STEP, dpr: 1 as const, palette: 'light' as const, col: 0, row: 0, mark: { kind: 'fixed' as const } }
    const served = c.store.tileDeliveries(B).find((d) => d.col === 0 && d.row === 0)!.refSha256
    expect(sha((await host.renderTile(docKey, { ...job, clock: { day: '2026-11-21', timeZone: 'UTC' } })).png)).toBe(served)
    expect(sha((await host.renderTile(docKey, { ...job, clock: { day: '2026-11-22', timeZone: 'UTC' } })).png)).not.toBe(served)
    await c.h.stop()
  }, 120_000)
})
