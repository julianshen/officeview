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
import { Pool, TILE_BYTES_ESTIMATE, PREP_DEADLINE_CAP_MS, PREP_MS_PER_MIB, PREP_TIMEOUT_MEMORY_MS, sheetPrepDeadlineMs, SheetPrepRefused, ShowCancelled } from '../src/pool'
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
    async function tiles(body: Record<string, unknown>) {
      clock.t += 100
      return c.tiles(++rq, body, { ms: 25_000 })
    }
    return { c, sid: m.session.sid as string, send, answer, one, layout, tiles }
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
  // The open parses and projects the million rows: about 20 s on macOS, about 50 s on the CI runner, pre-existing at
  // 4b fb624989 (docs/validation/streaming-s4c.md, "Linux CI and the million-row open"). So this pool's open deadline
  // is 120 s (the production default stays 30 s), the manifest wait 120 s and the hook 180 s; the sort test below has
  // its own 30 s (it parses nothing).
  beforeAll(async () => {
    s = await serverOn(await oneMillionSort(), new Pool({ workers: 1, openDeadlineMs: 120_000 }), { policy: { contentPolicy: 'visible-only' } })
    a = await s.open(120_000)
  }, 180_000)

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

describe('a sorted million-row view under visible-only after a worker recycle (fix round 1)', () => {
  const clock = { day: '2026-11-28', timeZone: 'UTC' }
  const range = wire(0, ONE_MILLION_ROWS - 1, 0, 0, 0)
  const view = { sort: { range, keys: [{ col: 0, order: 'asc' as const }], caseSensitive: false }, sortFilter: null, filter: null }
  const tileJob = { docRef: 'doc', contentPolicy: 'visible-only' as const, sheetIndex: 0, viewGen: 1, zoomStep: 5, dpr: 1 as const, palette: 'light' as const, col: 0, row: 0, mark: { kind: 'off' as const }, clock, view }
  let pool: Pool
  // A parse of the million rows costs about 20 s on macOS and about 50 s on the CI runner, pre-existing at 4b fb624989
  // (docs/validation/streaming-s4c.md, "Linux CI and the million-row open"). So the pool's open deadline is 120 s (the
  // production default stays 30 s), no hook holds two parses, and a hook that parses has 180 s. The test below
  // parses nothing and keeps its own 30 s.
  beforeAll(async () => {
    pool = new Pool({ workers: 1, recyclePages: 1, openDeadlineMs: 120_000 }); pools.push(pool)
    await pool.open('doc', await oneMillionSort(), 'visible-only', { sheetMode: 'tiles' })
    await pool.sheetLayout('s', { docRef: 'doc', contentPolicy: 'visible-only', sheetIndex: 0, clock })
  }, 180_000)
  // No parse: the sort and the first tile run on the open document; the tile's page then recycles the worker.
  beforeAll(async () => {
    const r = await pool.sheetView('s', { docRef: 'doc', contentPolicy: 'visible-only', clock, op: 'sort', sheetIndex: 0, base: { sort: null, sortFilter: null, filter: null }, sort: view.sort })
    expect(r.op).toBe('view')
    // The first tile paints on the warm port; the worker then recycles.
    expect((await pool.tile('s', tileJob, TILE_BYTES_ESTIMATE)).w).toBe(512)
    expect(pool.stats().recycled).toBe(1)
  }, 30_000)
  // One parse: the layout job on the recycled worker parses the document again (180 s, as above).
  beforeAll(async () => {
    await pool.sheetLayout('s', { docRef: 'doc', contentPolicy: 'visible-only', sheetIndex: 0, clock })
  }, 180_000)

  test('its later tiles paint: the cold port is prepared under the prep deadline', async () => {
    const preps = pool.stats().sheetPreps
    const t = await pool.tile('s', { ...tileJob, row: 1 }, TILE_BYTES_ESTIMATE)
    expect([t.w, t.h]).toEqual([512, 512])
    expect(pool.stats().sheetPreps).toBe(preps + 1)
    expect(pool.stats().recycled).toBe(2)
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
    // Fix round 1: the verdict is remembered, so the client's re-fetch (a derive) and every later op prep nothing.
    expect(pool.stats().sheetPreps).toBe(1)
    const refetch = await a.answer(a.send('sheet-order', {}))
    expect(refetch.end).toEqual({ reqId: expect.any(Number), type: 'refused', code: 'view-refused', reason: 'too-large' })
    expect(pool.stats().sheetPreps).toBe(1)
    expect(pool.stats().retired).toBe(1)
    const values = await a.one(a.send('filter-values', { viewGen: 0, range: GRID, col: 1 }))
    expect(values).toEqual({ reqId: expect.any(Number), type: 'refused', code: 'view-refused', reason: 'too-large' })
    const detect = await a.one(a.send('range-detect', { viewGen: 0, kind: 'sort', selection: { active: [2, 1], anchor: [2, 1], ranges: [] } }))
    expect(detect).toEqual({ reqId: expect.any(Number), type: 'refused', code: 'view-refused', reason: 'too-large' })
    expect(s.h.stream.sheetRatesOf(a.sid)!.sheetBusy).toBe(false)
    expect(s.store.sheetOps(a.sid).map((o) => [o.kind, o.outcome])).toEqual([['sort', 'view-refused'], ['filter-values', 'view-refused'], ['range-detect', 'view-refused']])
    expect([pool.stats().sheetPreps, pool.stats().retired]).toEqual([1, 1])
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

describe('tiles and text under a view prep a cold port (fix round 1)', () => {
  test('after a recycle, find, copy, a11y and tiles under the view answer too-large from one prep', async () => {
    const pool = new Pool({ workers: 1, testHooks: true, recyclePages: 1, prepDeadlineCapMsForTests: 100 })
    const s = await serverOn(await sortGrid(), pool, { policy: { copy: 'full', textA11y: true } })
    const a = await s.open()
    await a.layout()
    expect((await a.answer(a.send('sheet-sort', { viewGen: 0, range: GRID, keys: [{ col: 1, order: 'asc' }] }))).order?.head.viewGen).toBe(1)
    // A tile on the warm port paints; then the worker recycles (one page), and every cold port build is held 1 s.
    const tile = { sheetId: '0', viewGen: 1, zoomStep: 5, dpr: 1, palette: 'light' }
    expect((await a.tiles({ ...tile, tiles: [{ col: 0, row: 0 }] })).tiles).toHaveLength(1)
    expect(pool.stats().recycled).toBe(1)
    pool.prepHangForTests(1000)
    const preps = pool.stats().sheetPreps
    const find = await a.one(a.send('sheet-find', { query: 'apple' }))
    expect(find).toEqual({ reqId: expect.any(Number), type: 'refused', code: 'view-refused', reason: 'too-large' })
    expect(pool.stats().sheetPreps).toBe(preps + 1)
    const sel = { active: [1, 1], anchor: [1, 1], ranges: [{ rows: [[1, 2]], cols: [[1, 1]] }] }
    expect(await a.one(a.send('sheet-copy', { viewGen: 1, selection: sel }))).toEqual({ reqId: expect.any(Number), type: 'refused', code: 'view-refused', reason: 'too-large' })
    expect(await a.one(a.send('sheet-a11y', { viewGen: 1, rows: [[1, 2]], cols: [1, 1] }))).toEqual({ reqId: expect.any(Number), type: 'refused', code: 'view-refused', reason: 'too-large' })
    // A tile not in the render cache (row 1), so it reaches the pool.
    const tiles = await a.tiles({ ...tile, tiles: [{ col: 0, row: 1 }] })
    expect(tiles.tiles).toHaveLength(0)
    expect(tiles.end).toMatchObject({ type: 'tiles-done', refused: [{ col: 0, row: 1, code: 'view-refused', reason: 'too-large' }] })
    expect(pool.stats().sheetPreps).toBe(preps + 1)
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

describe('jobs queued before a prep refusal (final review R3)', () => {
  const clock = { day: '2026-11-28', timeZone: 'UTC' }
  const view = { sort: { range: GRID, keys: [{ col: 1, order: 'asc' as const }], caseSensitive: false }, sortFilter: null, filter: null }
  const route = { docRef: 'doc', contentPolicy: 'visible-only' as const, clock }
  async function pair(second: (pool: Pool) => Promise<unknown>) {
    const pool = new Pool({ workers: 1, testHooks: true, prepDeadlineCapMsForTests: 100, prepHangMsForTests: 1000 }); pools.push(pool)
    await pool.open('doc', await sortGrid(), 'visible-only', { sheetMode: 'tiles' })
    const first = pool.sheetView('s1', { ...route, op: 'derive', sheetIndex: 0, view })
    const queued = second(pool)
    const [a, b] = await Promise.all([rejection(first), rejection(queued)])
    return { pool, a, b }
  }
  test('a second derive on the same worker is refused from the remembered verdict: one prep, one retire', async () => {
    const { pool, a, b } = await pair((p) => p.sheetView('s2', { ...route, op: 'derive', sheetIndex: 0, view }))
    expect([a, b].map((e) => (e as Error).name)).toEqual(['SheetPrepRefused', 'SheetPrepRefused'])
    expect([pool.stats().sheetPreps, pool.stats().retired, pool.admission.reserved]).toEqual([1, 1, 0])
  }, 30_000)
  test('a queued find under the view is refused the same way', async () => {
    const { pool, a, b } = await pair((p) => p.sheetText('s2', { ...route, op: 'find', query: 'apple', matchCase: false, sheetIndex: 0, sizes: [], limit: 10, view }))
    expect([a, b].map((e) => (e as Error).name)).toEqual(['SheetPrepRefused', 'SheetPrepRefused'])
    expect([pool.stats().sheetPreps, pool.stats().retired, pool.admission.reserved]).toEqual([1, 1, 0])
  }, 30_000)
  test('a queued tile under the view is refused the same way', async () => {
    const tile = { ...route, sheetIndex: 0, viewGen: 1, zoomStep: 5, dpr: 1 as const, palette: 'light' as const, col: 0, row: 0, mark: { kind: 'off' as const }, view }
    const { pool, a, b } = await pair((p) => p.tile('s2', tile, TILE_BYTES_ESTIMATE))
    expect([a, b].map((e) => (e as Error).name)).toEqual(['SheetPrepRefused', 'SheetPrepRefused'])
    expect([pool.stats().sheetPreps, pool.stats().retired, pool.admission.reserved]).toEqual([1, 1, 0])
  }, 30_000)
})

describe('a remembered prep timeout expires (owner decision)', () => {
  const clock = { day: '2026-11-28', timeZone: 'UTC' }
  const view = { sort: { range: GRID, keys: [{ col: 1, order: 'asc' as const }], caseSensitive: false }, sortFilter: null, filter: null }
  const route = { docRef: 'doc', contentPolicy: 'visible-only' as const, clock }
  const derive = (pool: Pool, s: string) => pool.sheetView(s, { ...route, op: 'derive', sheetIndex: 0, view })
  const counts = (pool: Pool) => [pool.stats().sheetPreps, pool.stats().retired]
  /** A pool on a clock the test steps (`t.ms`), whose cold port builds hang past the prep deadline until released. */
  async function timedOut(o: Partial<ConstructorParameters<typeof Pool>[0]> = {}) {
    const t = { ms: 1_000_000 }
    const pool = new Pool({ workers: 1, testHooks: true, now: () => t.ms, prepDeadlineCapMsForTests: 100, prepHangMsForTests: 1000, ...o }); pools.push(pool)
    await pool.open('doc', await sortGrid(), 'visible-only', { sheetMode: 'tiles' })
    return { pool, t }
  }

  test('PREP_TIMEOUT_MEMORY_MS is 60 s', () => {
    expect(PREP_TIMEOUT_MEMORY_MS).toBe(60_000)
  })

  test('a job within 60 s of a prep timeout is refused too-large with no new prep and no new retire', async () => {
    const { pool, t } = await timedOut()
    const first = await rejection(derive(pool, 's1'))
    expect([(first as SheetPrepRefused).name, (first as SheetPrepRefused).cause]).toEqual(['SheetPrepRefused', 'timeout'])
    expect(counts(pool)).toEqual([1, 1])
    t.ms += PREP_TIMEOUT_MEMORY_MS - 1
    const again = await rejection(derive(pool, 's2'))
    expect([(again as SheetPrepRefused).name, (again as SheetPrepRefused).cause]).toEqual(['SheetPrepRefused', 'timeout'])
    expect(counts(pool)).toEqual([1, 1])
    expect(pool.admission.reserved).toBe(0)
  }, 30_000)

  test('past 60 s the remembered timeout is dropped: the same job preps again and commits', async () => {
    const { pool, t } = await timedOut()
    await rejection(derive(pool, 's1'))
    expect(counts(pool)).toEqual([1, 1])
    t.ms += PREP_TIMEOUT_MEMORY_MS + 1
    pool.prepHangForTests(0)
    const r = await derive(pool, 's2')
    expect(r.op).toBe('view')
    expect(counts(pool)).toEqual([2, 1])
    expect(pool.admission.reserved).toBe(0)
  }, 30_000)

  test('a new timeout re-arms the 60 s', async () => {
    const { pool, t } = await timedOut()
    await rejection(derive(pool, 's1'))
    t.ms += PREP_TIMEOUT_MEMORY_MS + 1
    // The hang is still in place: the prep runs again and times out again.
    await rejection(derive(pool, 's2'))
    expect(counts(pool)).toEqual([2, 2])
    t.ms += PREP_TIMEOUT_MEMORY_MS - 1
    const again = await rejection(derive(pool, 's3'))
    expect((again as SheetPrepRefused).cause).toBe('timeout')
    expect(counts(pool)).toEqual([2, 2])
  }, 30_000)

  test('a failed prep is still refused past 60 s', async () => {
    // A real worker whose `sheet-view-prep` jobs are answered with a failure reply (no refusal code) in its place.
    const factory = (url: string, options: WorkerOptions) => {
      const w = new Worker(url, options)
      let last = { parses: 0, pages: 0, docs: 0, probes: 0 }
      w.addEventListener('message', (e) => { const d = (e as MessageEvent).data; if (d && d.stats) last = d.stats })
      const post = w.postMessage.bind(w) as (m: unknown, t?: Transferable[]) => void
      w.postMessage = ((m: { type?: string; id?: number }, tr?: Transferable[]) => {
        if (m && m.type === 'sheet-view-prep') {
          queueMicrotask(() => w.onmessage?.({ data: { id: m.id, ok: false, error: 'forced prep failure', stats: last } } as MessageEvent))
          return
        }
        post(m, tr)
      }) as typeof w.postMessage
      return w
    }
    const { pool, t } = await timedOut({ workerFactoryForTests: factory, prepHangMsForTests: undefined, prepDeadlineCapMsForTests: undefined })
    const first = await rejection(derive(pool, 's1'))
    expect([(first as SheetPrepRefused).name, (first as SheetPrepRefused).cause]).toEqual(['SheetPrepRefused', 'failed'])
    expect(counts(pool)).toEqual([1, 0])
    t.ms += PREP_TIMEOUT_MEMORY_MS + 1
    const again = await rejection(derive(pool, 's2'))
    expect((again as SheetPrepRefused).cause).toBe('failed')
    expect(counts(pool)).toEqual([1, 0])
    expect(pool.admission.reserved).toBe(0)
  }, 30_000)
})
