/**
 * Stage 4c Task 14a — a sheet's sort data (its key view and `sheet-view` port) is prepared by a pool job of its own,
 * under its own bounded deadline sized from the document's admitted bytes; the sort, filter, value list or detection
 * then runs as a second job under the 2 s render deadline. A cold first sort of a million-row sheet under
 * `visible-only` commits. A prep past its deadline answers `view-refused {reason: 'too-large'}`, never a bare
 * `policy-refused`; a cancel during prep commits nothing; a warm port is not built again.
 */
import { afterAll, beforeAll, describe, expect, test } from 'bun:test'
import { validateManifest } from '../../src/stream/manifest'
import { decodeSheetLayout, decodeSheetOrder, type SheetOrder, type StreamSheetLayout, type WireRange } from '../../src/stream/sheet-wire'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import { DEFAULT_CAPS } from '../src/admission'
import { Pool, PREP_DEADLINE_CAP_MS, PREP_MS_PER_MIB, sheetPrepDeadlineMs, ShowCancelled } from '../src/pool'
import { DocumentHost } from '../src/render'
import { rejection } from './fixtures/rejection'
import { Client, start } from './harness'
import { ONE_MILLION_ROWS, oneMillionSort, SORT_GRID, sortGrid } from './sheet-fixtures'

const VIEW = ['sheet-view@1']
const T0 = Date.UTC(2026, 10, 28, 12)
const MiB = 1024 * 1024
const wire = (r0: number, r1: number, c0: number, c1: number, header: 0 | 1 = 1): WireRange => ({ r0, r1, c0, c1, header, totals: 0 })
const GRID = wire(0, SORT_GRID.rows - 1, 0, SORT_GRID.cols - 1)
type Text = Record<string, any>

const pools: Pool[] = []
const hosts: DocumentHost[] = []
afterAll(async () => {
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
  for (const p of pools) await p.dispose()
})

/** `Client.hello` with a longer manifest wait: the document is parsed before the manifest is sent. */
async function slowHello(c: Client, ms: number) {
  await c.opened()
  c.send({ reqId: 1, type: 'hello', client: 'officeview/1.0', recVersions: [1], features: VIEW })
  const m = await c.next((x) => 'binary' in x && x.binary.kind === 1, ms) as { binary: { payload: Uint8Array } }
  const v = validateManifest(JSON.parse(new TextDecoder().decode(m.binary.payload)))
  if (!v.ok) throw new Error(`bad manifest field ${v.field}`)
  c.manifest = v.manifest
  return v.manifest
}

/** A server over its own pool, on an injected clock that steps 100 ms per request, with its own audit. */
async function serverOn(doc: Uint8Array, pool: Pool, o: Parameters<typeof start>[0] = {}) {
  pools.push(pool)
  const clock = { t: T0 }
  const store = AuditStore.open(':memory:')
  const h = await start({ doc, pool, now: () => clock.t, tokenTtlSec: 86_400, audit: new AuditTable({ store }), ...o })
  let rq = 1000
  async function open(helloMs?: number) {
    const c = new Client(h.url)
    const m = helloMs === undefined ? await c.hello({ features: VIEW }) : await slowHello(c, helloMs)
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
        const x = await c.next((y) => 'text' in y && y.text.reqId === reqId, 25_000) as { text: Text }
        if (x.text.type !== 'sheet-order') return { end: x.text }
        parts.push(x.text)
        if (x.text.last === true) return { order: decodeSheetOrder(parts) }
      }
    }
    const one = async (reqId: number) => (await c.next((y) => 'text' in y && y.text.reqId === reqId, 25_000) as { text: Text }).text
    async function layout(): Promise<StreamSheetLayout> {
      clock.t += 100
      const r = await c.sheetLayout(++rq, '0', { ms: 25_000 })
      if (r.end) throw new Error(`layout refused: ${JSON.stringify(r.end)}`)
      return decodeSheetLayout(r.parts.map((p) => p.text)) as StreamSheetLayout
    }
    return { c, sid: m.session.sid as string, send, answer, one, layout }
  }
  return { h, store, open }
}

describe('the prep deadline', () => {
  test('is the render deadline plus PREP_MS_PER_MIB per admitted MiB, capped', () => {
    expect(sheetPrepDeadlineMs(0, DEFAULT_CAPS.renderDeadlineMs)).toBe(DEFAULT_CAPS.renderDeadlineMs)
    expect(sheetPrepDeadlineMs(8 * MiB, DEFAULT_CAPS.renderDeadlineMs)).toBe(DEFAULT_CAPS.renderDeadlineMs + 8 * PREP_MS_PER_MIB)
    expect(sheetPrepDeadlineMs(MiB / 2, DEFAULT_CAPS.renderDeadlineMs)).toBe(DEFAULT_CAPS.renderDeadlineMs + Math.ceil(PREP_MS_PER_MIB / 2))
    expect(sheetPrepDeadlineMs(1024 * MiB, DEFAULT_CAPS.renderDeadlineMs)).toBe(PREP_DEADLINE_CAP_MS)
    expect(sheetPrepDeadlineMs(1024 * MiB, DEFAULT_CAPS.renderDeadlineMs, 50)).toBe(50)
    expect(DEFAULT_CAPS.renderDeadlineMs).toBe(2000)
  })
})

describe('a cold first sort of a million rows under visible-only', () => {
  let s: Awaited<ReturnType<typeof serverOn>>
  let a: Awaited<ReturnType<Awaited<ReturnType<typeof serverOn>>['open']>>
  // The open parses and projects the million rows (about 20 s measured); the sort test below has its own 30 s.
  beforeAll(async () => {
    s = await serverOn(await oneMillionSort(), new Pool({ workers: 1 }), { policy: { contentPolicy: 'visible-only' } })
    a = await s.open(29_000)
  }, 30_000)

  test('commits', async () => {
    const layout = await a.layout()
    expect(layout.sort).toBeNull()
    const range = wire(0, ONE_MILLION_ROWS - 1, 0, 0, 0)
    const r = await a.answer(a.send('sheet-sort', { viewGen: 0, range, keys: [{ col: 0, order: 'asc' }] }))
    expect(r.end).toBeUndefined()
    expect(r.order!.head.viewGen).toBe(1)
    expect(r.order!.bands[0]!.sourceAtOffset[0]).toBe(ONE_MILLION_ROWS - 1)
    expect(s.store.sheetOps(a.sid).map((o) => o.outcome)).toEqual(['ok'])
    a.c.close(); await s.h.stop()
  }, 30_000)
})

describe('a prep past its own deadline', () => {
  test('answers view-refused too-large for a sort, a value list and a detection; nothing commits', async () => {
    const pool = new Pool({ workers: 1, testHooks: true, prepDeadlineCapMsForTests: 100, prepHangMsForTests: 1000 })
    const s = await serverOn(await sortGrid(), pool)
    const a = await s.open()
    await a.layout()
    const sort = await a.answer(a.send('sheet-sort', { viewGen: 0, range: GRID, keys: [{ col: 1, order: 'asc' }] }))
    expect(sort.end).toEqual({ reqId: expect.any(Number), type: 'refused', code: 'view-refused', reason: 'too-large' })
    expect(s.h.stream.sheetRatesOf(a.sid)!.sheetBusy).toBe(false)
    const values = await a.one(a.send('filter-values', { viewGen: 0, range: GRID, col: 1 }))
    expect(values).toEqual({ reqId: expect.any(Number), type: 'refused', code: 'view-refused', reason: 'too-large' })
    const detect = await a.one(a.send('range-detect', { viewGen: 0, kind: 'sort', selection: { active: [2, 1], anchor: [2, 1], ranges: [] } }))
    expect(detect).toEqual({ reqId: expect.any(Number), type: 'refused', code: 'view-refused', reason: 'too-large' })
    expect(s.h.stream.sheetRatesOf(a.sid)!.sheetBusy).toBe(false)
    expect(s.store.sheetOps(a.sid).map((o) => [o.kind, o.outcome])).toEqual([['sort', 'view-refused'], ['filter-values', 'view-refused'], ['range-detect', 'view-refused']])
    expect((await a.layout()).sort).toBeNull()
    expect(pool.admission.reserved).toBe(0)
    a.c.close(); await s.h.stop()
  }, 30_000)

  test('a cancel during prep commits nothing; the next sort runs on the warm port', async () => {
    const pool = new Pool({ workers: 1, testHooks: true, prepHangMsForTests: 800 })
    const s = await serverOn(await sortGrid(), pool)
    const a = await s.open()
    await a.layout()
    const reqId = a.send('sheet-sort', { viewGen: 0, range: GRID, keys: [{ col: 1, order: 'asc' }] })
    await Bun.sleep(200)
    a.send('sheet-cancel', { target: reqId })
    const r = await a.answer(reqId)
    expect(r.end).toMatchObject({ type: 'refused', code: 'cancelled' })
    expect((await a.layout()).sort).toBeNull()
    const again = await a.answer(a.send('sheet-sort', { viewGen: 0, range: GRID, keys: [{ col: 1, order: 'asc' }] }))
    expect(again.order?.head.viewGen).toBe(1)
    expect(s.store.sheetOps(a.sid).map((o) => o.outcome)).toEqual(['cancelled', 'ok'])
    expect(pool.admission.reserved).toBe(0)
    a.c.close(); await s.h.stop()
  }, 30_000)
})

describe('an abort during prep', () => {
  test('runs no op: the pool job rejects ShowCancelled once the prep returns', async () => {
    const pool = new Pool({ workers: 1, testHooks: true, prepHangMsForTests: 400 }); pools.push(pool)
    await pool.open('doc', await sortGrid(), 'visible-only', { sheetMode: 'tiles' })
    const clock = { day: '2026-11-28', timeZone: 'UTC' }
    const ac = new AbortController()
    const job = pool.sheetView('s', { docRef: 'doc', contentPolicy: 'visible-only', clock, op: 'sort', sheetIndex: 0, base: { sort: null, sortFilter: null, filter: null }, sort: { range: GRID, keys: [{ col: 1, order: 'asc' }], caseSensitive: false } }, undefined, ac.signal)
    setTimeout(() => ac.abort(), 150)
    expect(await rejection(job)).toBeInstanceOf(ShowCancelled)
    expect(pool.admission.reserved).toBe(0)
  }, 30_000)
})

describe('prep on the document host', () => {
  test('builds the port once; a warm port is not built again', async () => {
    const host = new DocumentHost(); hosts.push(host)
    const { docKey } = await host.open('doc', await sortGrid(), 'full', { sheetMode: 'tiles' })
    const clock = { day: '2026-11-28', timeZone: 'UTC' }
    expect(await host.sheetViewPrep(docKey, clock, 'full', 0)).toEqual({ built: true })
    expect(await host.sheetViewPrep(docKey, clock, 'full', 0)).toEqual({ built: false })
    const ports = (host as unknown as { docs: Map<string, { viewPorts?: Map<string, unknown> }> }).docs.get(docKey)!.viewPorts!
    expect(ports.size).toBe(1)
    const r = await host.sheetView(docKey, clock, 'full', { op: 'sort', sheetIndex: 0, base: { sort: null, sortFilter: null, filter: null }, sort: { range: GRID, keys: [{ col: 1, order: 'asc' }], caseSensitive: false } })
    expect(r.op).toBe('view')
    expect(ports.size).toBe(1)
  })
})
