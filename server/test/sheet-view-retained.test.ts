/**
 * Stage 4c final review R2 — a derived view a worker retains (its `derivedViews` entry, the sized geometries and text
 * views built over it, the port's current state) is charged by a `DocumentLedger` hold the whole time, or gone: when a
 * view's last holder releases it (replaced, cleared, its session closed while another keeps the document), and when a
 * sort or filter job ran but its commit did not happen, its worker drops it. Through the real socket server and Pool.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { decodeSheetOrder, type SheetOrder, type WireRange } from '../../src/stream/sheet-wire'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import { Pool } from '../src/pool'
import { Client, start } from './harness'
import { SORT_GRID, sortGrid } from './sheet-fixtures'

const VIEW = ['sheet-view@1']
const T0 = Date.UTC(2026, 10, 28, 12)
const wire = (r0: number, r1: number, c0: number, c1: number, header: 0 | 1 = 1): WireRange => ({ r0, r1, c0, c1, header, totals: 0 })
const GRID = wire(0, SORT_GRID.rows - 1, 0, SORT_GRID.cols - 1)
const KEY_COLS = [0, 1, 2, 3]
const NOTHING = { derived: 0, bytes: 0, geometries: 0, texts: 0, ports: 0 }
type Text = Record<string, any>

const pools: Pool[] = []
afterAll(async () => { for (const p of pools) await p.dispose() })

/** A server over `pool` (testHooks), on an injected clock that steps 100 ms per request. */
async function serverOn(pool: Pool) {
  const clock = { t: T0 }
  const h = await start({ doc: await sortGrid(), pool, now: () => clock.t, tokenTtlSec: 86_400, audit: new AuditTable({ store: AuditStore.open(':memory:') }) })
  let rq = 1000
  async function open() {
    const c = new Client(h.url)
    const m = await c.hello({ features: VIEW })
    await c.ready()
    while (h.stream.table.get(m.session.sid)?.status !== 'stream') await Bun.sleep(5)
    const send = (type: string, body: Record<string, unknown>) => {
      clock.t += 100
      const reqId = ++rq
      c.send({ token: m.session.token, reqId, type, sheetId: '0', ...body })
      return reqId
    }
    async function answer(reqId: number): Promise<{ order?: SheetOrder; end?: Text }> {
      const parts: Text[] = []
      for (;;) {
        const x = await c.next((y) => 'text' in y && y.text.reqId === reqId, 20_000) as { text: Text }
        if (x.text.type !== 'sheet-order') return { end: x.text }
        parts.push(x.text)
        if (x.text.last === true) return { order: decodeSheetOrder(parts) }
      }
    }
    const sort = async (viewGen: number, keys: Array<{ col: number; order: 'asc' | 'desc' }>) => answer(send('sheet-sort', { viewGen, range: GRID, keys }))
    async function tile(viewGen: number) {
      clock.t += 100
      return c.tiles(++rq, { sheetId: '0', viewGen, zoomStep: 5, dpr: 1, palette: 'light', tiles: [{ col: 0, row: 0 }] })
    }
    return { c, sid: m.session.sid as string, send, answer, sort, tile }
  }
  const docRef = () => Object.keys(pool.stats().affinity)[0]!
  const retained = async () => (await pool.retainedViewsForTests(docRef(), 'visible-only'))!
  return { h, open, retained }
}

describe('derived views a worker retains are charged or gone (final review R2)', () => {
  test('replacement keeps only the current view; a clear leaves nothing; the ledger agrees', async () => {
    const pool = new Pool({ workers: 1, testHooks: true }); pools.push(pool)
    const s = await serverOn(pool)
    const a = await s.open()
    const base = s.h.stream.documentLedgerBytes()
    let viewGen = 0
    for (const col of KEY_COLS) for (const order of ['asc', 'desc'] as const) {
      const r = await a.sort(viewGen, [{ col, order }])
      expect({ col, order, end: r.end }).toEqual({ col, order, end: undefined })
      viewGen = r.order!.head.viewGen
      const t = await a.tile(viewGen)
      expect(t.tiles).toHaveLength(1)
      const kept = await s.retained()
      expect({ col, order, derived: kept.derived, ports: kept.ports, geometries: kept.geometries }).toEqual({ col, order, derived: 1, ports: 1, geometries: 1 })
      expect(kept.bytes).toBeGreaterThan(0)
    }
    const cleared = await a.sort(viewGen, [])
    expect(cleared.end).toBeUndefined()
    expect(s.h.stream.documentLedgerBytes()).toBe(base)
    expect(await s.retained()).toEqual(NOTHING)
    expect(pool.admission.reserved).toBe(0)
    a.c.close(); await s.h.stop()
  }, 30_000)

  test("the last sorted holder closing drops the view while another session keeps the document", async () => {
    const pool = new Pool({ workers: 1, testHooks: true }); pools.push(pool)
    const s = await serverOn(pool)
    const a = await s.open(), b = await s.open(), c = await s.open()
    const base = s.h.stream.documentLedgerBytes()
    expect((await a.sort(0, [{ col: 1, order: 'asc' }])).order?.head.viewGen).toBe(1)
    expect((await b.sort(0, [{ col: 1, order: 'asc' }])).order?.head.viewGen).toBe(1)
    expect((await s.retained()).derived).toBe(1)
    // b still holds the shared view: it stays.
    a.c.close()
    while (s.h.stream.table.get(a.sid)) await Bun.sleep(5)
    expect((await s.retained()).derived).toBe(1)
    b.c.close()
    while (s.h.stream.table.get(b.sid)) await Bun.sleep(5)
    expect(s.h.stream.documentLedgerBytes()).toBe(base)
    expect(await s.retained()).toEqual(NOTHING)
    c.c.close(); await s.h.stop()
  }, 30_000)

  test('a sort whose commit is cancelled after its job ran leaves nothing in the worker', async () => {
    const real = new Pool({ workers: 1, testHooks: true }); pools.push(real)
    let gate: Promise<void> | undefined, ran!: () => void
    const reached = new Promise<void>((r) => { ran = r })
    // The real job runs (and caches its view); its answer is held until the cancel has arrived.
    const proxy = new Proxy(real, {
      get(t, k) {
        if (k === 'sheetView') return async (...a: Parameters<Pool['sheetView']>) => {
          const r = await t.sheetView(...a)
          if (gate && a[1].op === 'sort') { ran(); await gate }
          return r
        }
        const v = Reflect.get(t, k, t)
        return typeof v === 'function' ? v.bind(t) : v
      },
    }) as Pool
    const s = await serverOn(proxy)
    const a = await s.open()
    const base = s.h.stream.documentLedgerBytes()
    let release!: () => void
    gate = new Promise<void>((r) => { release = r })
    const target = a.send('sheet-sort', { viewGen: 0, range: GRID, keys: [{ col: 2, order: 'desc' }] })
    await reached
    expect((await real.retainedViewsForTests(Object.keys(real.stats().affinity)[0]!, 'visible-only'))!.derived).toBe(1)
    a.send('sheet-cancel', { target })
    await Bun.sleep(50)
    release()
    expect((await a.answer(target)).end).toMatchObject({ type: 'refused', code: 'cancelled' })
    expect(s.h.stream.documentLedgerBytes()).toBe(base)
    expect(await real.retainedViewsForTests(Object.keys(real.stats().affinity)[0]!, 'visible-only')).toEqual(NOTHING)
    a.c.close(); await s.h.stop()
  }, 30_000)
})
