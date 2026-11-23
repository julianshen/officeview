/**
 * Stage 4a Task 9 — the `tiles` op (S4 §6.5, §6.8; R3, R4, R12, R13): one `XLSX_TILE` (kind 5) per delivered tile
 * under the request's reqId, then `tiles-done`; per-tile `'tile'` rate (refuse-only, never toward 4008); the
 * view-state cache key (the session clock keys canonical tiles); per-tile Layer-1 and `tile_deliveries`; the
 * `sheet_views` row; `sheet-cancel`; and a resume that stops the old socket's batch.
 */
import { afterAll, beforeAll, describe, expect, test } from 'bun:test'
import { InvalidOperand } from '../../src/stream/ops'
import { decodeTileBody, encodeTileBody, TILE_PX } from '../../src/stream/sheet-wire'
import { encodeUnit } from '../../src/stream/unit'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import type { Pool } from '../src/pool'
import { tileLimit } from '../src/sheet/tile'
import { Client, disposePool, pool, start } from './harness'
import { grid500x100 } from './sheet-fixtures'

afterAll(async () => { await disposePool() })

const TILES = ['sheet-tiles@1']
let grid: Uint8Array
beforeAll(async () => { grid = await grid500x100() })

const req = (o: Record<string, unknown> = {}) => ({ sheetId: '0', viewGen: 0, zoomStep: 5, dpr: 1, palette: 'light', tiles: [{ col: 0, row: 0 }], ...o })
/** `n` distinct coordinates, row-major over a 16-wide window from `(c0, r0)`. */
const coords = (n: number, c0 = 0, r0 = 0) => Array.from({ length: n }, (_, i) => ({ col: c0 + (i % 16), row: r0 + Math.floor(i / 16) }))
async function tilesClient(url: string) {
  const c = new Client(url)
  const m = await c.hello({ features: TILES })
  await c.ready()
  return { c, m }
}
function serverWith(o: Parameters<typeof start>[0] = {}) {
  const store = AuditStore.open(':memory:')
  return { store, h: start({ doc: grid, policy: { visibleMark: 'fixed', forensicMark: true }, audit: new AuditTable({ store }), ...o }) }
}

/** A minimal PNG header (signature + IHDR) claiming `w × h`: the codec reads the IHDR only. */
function pngOf(w: number, h: number): Uint8Array {
  const b = new Uint8Array(33), v = new DataView(b.buffer)
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]); v.setUint32(8, 13); b.set([73, 72, 68, 82], 12); v.setUint32(16, w); v.setUint32(20, h)
  return b
}
describe('XLSX_TILE body (R13)', () => {
  const t = { sheetIndex: 2, viewGen: 0, zoomStep: 10, dpr: 2 as const, col: 675, row: 0xffffffff, png: pngOf(TILE_PX, TILE_PX), w: TILE_PX, h: TILE_PX }
  test('round-trips, laid out as R13 under an encodeUnit header', () => {
    const bytes = encodeTileBody(t)
    const back = decodeTileBody(bytes)
    expect({ ...back, png: [...back.png] }).toEqual({ ...t, png: [...t.png] })
    const v = new DataView(bytes.buffer, bytes.byteOffset)
    expect([v.getUint32(0, true), bytes[12], v.getUint32(13, true), bytes[17], bytes[18], v.getUint16(19, true), v.getUint32(21, true), v.getUint32(25, true), v.getUint32(29, true), v.getUint32(33, true), v.getUint32(37, true)])
      .toEqual([2, 3, 0, 10, 2, 0, 675, 0xffffffff, TILE_PX, TILE_PX, 33])
  })
  test('anything else throws InvalidOperand', () => {
    const good = encodeTileBody(t)
    const patched = (at: number, byte: number) => { const b = good.slice(); b[at] = byte; return b }
    const bad: Array<[string, Uint8Array]> = [
      ['truncated', good.subarray(0, 20)],
      ['discriminator 2', patched(12, 2)],
      ['zoomStep 11', patched(17, 11)],
      ['dpr 3', patched(18, 3)],
      ['reserved', patched(19, 1)],
      ['not 512²', encodeUnit({ unitId: 0, discriminator: 3, body: (() => { const x = encodeTileBody(t).subarray(13).slice(); new DataView(x.buffer).setUint32(16, 256, true); x.set(pngOf(256, TILE_PX), 28); return x })() })],
    ]
    for (const [name, b] of bad) expect([name, (() => { try { decodeTileBody(b); return 'decoded' } catch (e) { return e instanceof InvalidOperand ? 'InvalidOperand' : String(e) } })()]).toEqual([name, 'InvalidOperand'])
    expect(() => encodeTileBody({ ...t, w: 256 })).toThrow(InvalidOperand)
    expect(() => encodeTileBody({ ...t, dpr: 3 as never })).toThrow(InvalidOperand)
  })
})

describe('tiles (S4 §6.5)', () => {
  test('happy path', async () => {
    const { store, h: hp } = serverWith(); const h = await hp
    const { c, m } = await tilesClient(h.url)
    const want = [{ col: 0, row: 0 }, { col: 1, row: 0 }, { col: 0, row: 1 }, { col: 1, row: 1 }]
    const r = await c.tiles(10, req({ tiles: want }))
    expect(r.end).toEqual({ reqId: 10, type: 'tiles-done', refused: [] })
    expect(r.tiles).toHaveLength(4)
    for (const t of r.tiles) {
      expect({ sheetIndex: t.sheetIndex, viewGen: t.viewGen, zoomStep: t.zoomStep, dpr: t.dpr, w: t.w, h: t.h }).toEqual({ sheetIndex: 0, viewGen: 0, zoomStep: 5, dpr: 1, w: TILE_PX, h: TILE_PX })
      const p = new DataView(t.png.buffer, t.png.byteOffset, t.png.byteLength)
      expect([p.getUint32(16), p.getUint32(20)]).toEqual([TILE_PX, TILE_PX])
    }
    const key = (t: { col: number; row: number }) => `${t.col},${t.row}`
    expect(r.tiles.map(key).sort()).toEqual(want.map(key).sort())
    const sid = m.session.sid
    const rows = store.tileDeliveries(sid)
    expect(rows).toHaveLength(4)
    expect(rows.every((d) => d.status === 'l1' && d.sheetId === '0' && d.viewGen === 0 && d.zoomStep === 5 && d.dpr === 1 && d.palette === 'light')).toBe(true)
    expect(rows.map(key).sort()).toEqual(want.map(key).sort())
    const view = store.sheetView(sid, '0', 0)!
    expect(view.spec).toEqual({ sort: null, filters: [], sizes: [] })
    expect(view.clock).toEqual(h.stream.sheetClockOf(sid)!)
    c.close(); await h.stop()
  }, 60_000)

  test('malformed batches are refused whole and charge nothing', async () => {
    const { store, h: hp } = serverWith(); const h = await hp
    const { c, m } = await tilesClient(h.url)
    // Sheet 0's whole domain at zoom 1, dpr 1 (16,384 default 64-px columns): the first column past the grid.
    const pastCol = tileLimit(16_384 * 64, 1, 1)
    const dup = [...coords(63), { col: 0, row: 0 }]
    const cases: Array<[string, Record<string, unknown>]> = [
      ['65 tiles', req({ tiles: coords(65) })],
      ['a duplicate', req({ tiles: dup })],
      ['zoomStep 11', req({ zoomStep: 11, tiles: coords(64) })],
      ['dpr 3', req({ dpr: 3, tiles: coords(64) })],
      ["palette 'blue'", req({ palette: 'blue', tiles: coords(64) })],
      ['col past the grid', req({ tiles: [...coords(63), { col: pastCol, row: 0 }] })],
      ["sheetId '9'", req({ sheetId: '9', tiles: coords(64) })],
      ['no token', req({ tiles: coords(64), token: undefined })],
    ]
    let reqId = 20
    for (const [name, body] of cases) {
      const r = await c.tiles(reqId, body)
      expect([name, r.tiles.length, r.end.type, r.end.code]).toEqual([name, 0, 'refused', 'policy-refused'])
      reqId++
    }
    const stale = await c.tiles(reqId++, req({ viewGen: 1, tiles: coords(64) }))
    expect([stale.tiles.length, stale.end.type, stale.end.code]).toEqual([0, 'refused', 'stale-view'])
    expect(store.tileDeliveries(m.session.sid)).toEqual([])
    expect(store.sheetView(m.session.sid, '0', 0)).toBeUndefined()
    // Nothing above was charged: a full batch is admitted whole under the real 120-per-second burst.
    const ok = await c.tiles(reqId, req({ tiles: coords(64) }), { ms: 60_000 })
    expect(ok.end).toEqual({ reqId, type: 'tiles-done', refused: [] })
    expect(ok.tiles).toHaveLength(64)
    c.close(); await h.stop()
  }, 120_000)

  test('tile rate refuses per tile, never closes', async () => {
    let t = Date.UTC(2026, 10, 23, 12, 0, 0)
    const { h: hp } = serverWith({ now: () => t }); const h = await hp
    const { c } = await tilesClient(h.url)
    const a = await c.tiles(30, req({ tiles: coords(64) }), { ms: 60_000 })
    expect(a.end.refused).toEqual([])
    const b = await c.tiles(31, req({ tiles: coords(64, 0, 4) }), { ms: 60_000 })
    expect(b.tiles).toHaveLength(56)
    expect(b.end.type).toBe('tiles-done')
    expect(b.end.refused).toHaveLength(8)
    expect(b.end.refused.every((x: { code: string }) => x.code === 'rate-limited')).toBe(true)
    expect(b.end.retryAfterMs).toBeGreaterThan(0)
    for (let i = 0; i < 10; i++) {
      const r = await c.tiles(40 + i, req({ tiles: coords(64, 0, 8) }))
      expect([r.tiles.length, r.end.type, r.end.refused.length]).toEqual([0, 'tiles-done', 64])
    }
    expect(c.ws.readyState).toBe(WebSocket.OPEN)
    t += b.end.retryAfterMs
    const refused = b.end.refused.map((x: { col: number; row: number }) => ({ col: x.col, row: x.row }))
    const again = await c.tiles(60, req({ tiles: refused }), { ms: 60_000 })
    expect(again.end.refused).toEqual([])
    expect(again.tiles).toHaveLength(8)
    expect(c.ws.readyState).toBe(WebSocket.OPEN)
    c.close(); await h.stop()
  }, 120_000)

  test('cancel stops the batch', async () => {
    const { h: hp } = serverWith(); const h = await hp
    const { c } = await tilesClient(h.url)
    c.send({ reqId: 70, type: 'tiles', token: c.manifest!.session.token, ...req({ tiles: coords(64) }) })
    c.send({ reqId: 71, type: 'sheet-cancel', target: 70 })
    const r = await c.tiles(70, {}, { send: false, ms: 60_000 })
    expect(r.end.type).toBe('tiles-done')
    expect(r.end.cancelled).toBe(true)
    expect(r.tiles.length).toBeLessThanOrEqual(8)
    // An unknown target is ignored; the session keeps working.
    c.send({ reqId: 72, type: 'sheet-cancel', target: 9999 })
    const ok = await c.tiles(73, req())
    expect(ok.end).toEqual({ reqId: 73, type: 'tiles-done', refused: [] })
    c.close(); await h.stop()
  }, 60_000)

  test('same clock shares canonical tiles; marked bytes differ', async () => {
    const base = pool()
    let calls = 0
    const counting = new Proxy(base, { get(target, key) { const v = Reflect.get(target, key, target); if (key === 'tile') return (...a: Parameters<Pool['tile']>) => { calls++; return target.tile(...a) }; return typeof v === 'function' ? v.bind(target) : v } })
    const { store, h: hp } = serverWith({ pool: counting }); const h = await hp
    const one = await tilesClient(h.url)
    const r1 = await one.c.tiles(80, req())
    const two = await tilesClient(h.url)
    const r2 = await two.c.tiles(81, req())
    expect([r1.tiles.length, r2.tiles.length]).toEqual([1, 1])
    expect(calls).toBe(1)
    expect(Buffer.compare(Buffer.from(r1.tiles[0]!.png), Buffer.from(r2.tiles[0]!.png))).not.toBe(0)
    const [d1, d2] = [store.tileDeliveries(one.m.session.sid), store.tileDeliveries(two.m.session.sid)]
    expect([d1.length, d2.length]).toEqual([1, 1])
    expect(d1[0]!.refSha256).toBe(d2[0]!.refSha256)
    one.c.close(); two.c.close(); await h.stop()
  }, 60_000)

  test('a different day or zone never shares', async () => {
    const base = pool()
    let calls = 0
    const counting = new Proxy(base, { get(target, key) { const v = Reflect.get(target, key, target); if (key === 'tile') return (...a: Parameters<Pool['tile']>) => { calls++; return target.tile(...a) }; return typeof v === 'function' ? v.bind(target) : v } })
    // 20:00 UTC: already the next calendar day in Tokyo.
    let t = Date.UTC(2026, 10, 23, 20, 0, 0)
    let zone: string | undefined
    const { h: hp } = serverWith({ pool: counting, now: () => t, authorize: async () => ({ viewerId: 'alice', policy: { visibleMark: 'fixed', forensicMark: true, ...(zone ? { timeZone: zone } : {}) } }) })
    const h = await hp
    const tileOf = async (reqId: number) => {
      const s = await tilesClient(h.url)
      const r = await s.c.tiles(reqId, req())
      expect(r.end).toEqual({ reqId, type: 'tiles-done', refused: [] })
      s.c.close()
      return h.stream.sheetClockOf(s.m.session.sid)!
    }
    expect(await tileOf(90)).toEqual({ day: '2026-11-23', timeZone: 'UTC' })
    expect(await tileOf(91)).toEqual({ day: '2026-11-23', timeZone: 'UTC' })
    expect(calls).toBe(1)
    t += 86_400_000
    expect(await tileOf(92)).toEqual({ day: '2026-11-24', timeZone: 'UTC' })
    expect(calls).toBe(2)
    zone = 'Asia/Tokyo'
    expect(await tileOf(93)).toEqual({ day: '2026-11-25', timeZone: 'Asia/Tokyo' })
    expect(calls).toBe(3)
    // November: London is on UTC, so the same day as the UTC sid above, in another zone.
    zone = 'Europe/London'
    expect(await tileOf(94)).toEqual({ day: '2026-11-24', timeZone: 'Europe/London' })
    expect(calls).toBe(4)
    await h.stop()
  }, 60_000)

  test("resume stops the old socket's batch", async () => {
    const sends: unknown[] = []
    const transport = {
      sendBinary(ws: unknown, f: Uint8Array) { sends.push(ws); return (ws as { sendBinary(f: Uint8Array): number }).sendBinary(f) },
      bufferedAmount(ws: unknown) { return (ws as { getBufferedAmount(): number }).getBufferedAmount() },
    }
    const { h: hp } = serverWith({ transport }); const h = await hp
    const a = await tilesClient(h.url)
    const oldWs = sends[0]
    a.c.send({ reqId: 100, type: 'tiles', token: a.m.session.token, ...req({ tiles: coords(64) }) })
    await a.c.next((m) => 'binary' in m && m.binary.kind === 5 && m.binary.reqId === 100, 60_000)
    const b = new Client(h.url)
    await b.hello({ features: TILES, resume: { sid: a.m.session.sid, resumeProof: a.m.session.token } })
    // From the new socket's manifest on (sent after the supersession), the old socket gets nothing.
    const mark = sends.findIndex((ws) => ws !== oldWs)
    expect(mark).toBeGreaterThan(0)
    await b.ready()
    // Long enough for several more tiles of the old batch, were it still running.
    await Bun.sleep(1_500)
    expect(sends.slice(mark).filter((ws) => ws === oldWs)).toEqual([])
    const ok = await b.tiles(101, req())
    expect(ok.end).toEqual({ reqId: 101, type: 'tiles-done', refused: [] })
    b.close(); await h.stop()
  }, 60_000)
})
