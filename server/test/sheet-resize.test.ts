/**
 * Stage 4b Task 4 — per-session view state (S4 §6.3; plan P11–P14): `sheet-resize`, `viewGen`, the view-state caps,
 * `sheet-layout`'s session sizes, tiles painted at the session's sizes, and the canonical tier shared by equal views.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { hiddenFixtures } from '../../scripts/stream-s0/hidden-fixtures'
import { decodeSheetLayout, encodeSheetLayout, type StreamSheetLayout } from '../../src/stream/sheet-wire'
import { MAX_SESSION_SIZES } from '../../src/xlsx/view/axisSizes'
import { streamSheetGeometry } from '../../src/xlsx/view/streamGeometry'
import { renderSheetViewport } from '../../src/xlsx/view/renderViewport'
import { tileLogicalBox } from '../../src/xlsx/view/tileCache'
import { LayoutCache } from '../../src/xlsx/view/textLayout'
import { buildSheetGeometry, frameBudgetFor, planViewport, type SheetGeometry } from '../../src/xlsx/view/viewport'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import { decodePng } from '../src/mark/png'
import type { Pool } from '../src/pool'
import { DocumentHost } from '../src/render'
import type { SheetClock } from '../src/sheet/clock'
import { TILE_PALETTE_COLORS, tileLimit } from '../src/sheet/tile'
import { SessionViews, sessionSizesOf, ViewStateLedger, type WireSizes } from '../src/sheet/view-state'
import { Client, disposePool, pool, start } from './harness'
import { grid500x100, textGrid, unresolvedAnchor } from './sheet-fixtures'

const hosts: DocumentHost[] = []
afterAll(async () => {
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
  await disposePool()
})

const VIEW = ['sheet-view@1']
const TILES = ['sheet-tiles@1']
const T = 512
let rq = 100

async function open(url: string, features: readonly string[] = VIEW) {
  const c = new Client(url)
  const m = await c.hello({ features })
  await c.ready()
  return { c, sid: m.session.sid }
}
const resize = async (c: Client, body: Record<string, unknown>) => (await c.text({ reqId: ++rq, type: 'sheet-resize', sheetId: '0', viewGen: 0, ...body })).end
async function layoutOf(c: Client, sheetId = '0') {
  const r = await c.sheetLayout(++rq, sheetId)
  if (r.end) throw new Error(`layout refused: ${JSON.stringify(r.end)}`)
  return { parts: r.parts, l: decodeSheetLayout(r.parts.map((p) => p.text)) as StreamSheetLayout }
}
const tile00 = { zoomStep: 5, dpr: 1, palette: 'light', tiles: [{ col: 0, row: 0 }] }

/** A pool whose `sheetLayout` calls wait for the next `hold()` gate, in call order (calls with no gate left run at once). */
function gatedPool() {
  const gates: Array<{ wait: Promise<void>; open: () => void }> = []
  const hold = () => { let open!: () => void; const wait = new Promise<void>((r) => { open = r }); const g = { wait, open }; gates.push(g); return g }
  let next = 0
  const base = pool()
  const p = new Proxy(base, {
    get(t, k) {
      if (k === 'sheetLayout') return async (...a: Parameters<Pool['sheetLayout']>) => { const g = gates[next++]; if (g) await g.wait; return t.sheetLayout(...a) }
      const v = Reflect.get(t, k, t)
      return typeof v === 'function' ? v.bind(t) : v
    },
  })
  return { pool: p, hold }
}
/** A pool that counts its `tile` calls. */
function countingPool() {
  const n = { tiles: 0 }
  const base = pool()
  const p = new Proxy(base, {
    get(t, k) {
      if (k === 'tile') return (...a: Parameters<Pool['tile']>) => { n.tiles++; return t.tile(...a) }
      const v = Reflect.get(t, k, t)
      return typeof v === 'function' ? v.bind(t) : v
    },
  })
  return { pool: p, n }
}

const UTC: SheetClock = { day: new Date().toISOString().slice(0, 10), timeZone: 'UTC' }
async function snapshotOf(bytes: Uint8Array, clock: SheetClock = UTC) {
  const host = new DocumentHost(); hosts.push(host)
  const o = await host.open('ref', bytes, 'visible-only', { sheetMode: 'tiles' })
  return { host, docKey: o.docKey, snap: await host.sheetSnapshot(o.docKey, 0, clock) }
}

describe('sheet-resize (P11, P12)', () => {
  test('resize applies, clamps and advances viewGen', async () => {
    const h = await start({ doc: await grid500x100() })
    const { c } = await open(h.url)
    expect(await resize(c, { axis: 'col', index: 2, size: 2000 })).toEqual({ reqId: rq, type: 'sheet-view', sheetId: '0', viewGen: 1, change: { axis: 'col', index: 2, size: 1790 } })
    const a = (await layoutOf(c)).l
    expect(a.viewGen).toBe(1)
    expect(a.sessionCols).toEqual([[2, 1790]])
    expect(a.sessionRows ?? []).toEqual([])
    expect(await resize(c, { viewGen: 1, reset: true })).toEqual({ reqId: rq, type: 'sheet-view', sheetId: '0', viewGen: 2, change: { reset: true } })
    const b = (await layoutOf(c)).l
    expect(b.viewGen).toBe(2)
    expect(b.sessionCols ?? []).toEqual([])
    c.close(); await h.stop()
  }, 30_000)

  test('refusals never move viewGen', async () => {
    const viewGen0 = async (c: Client) => expect((await layoutOf(c)).l.viewGen).toBe(0)
    // stale-view
    const g = await start({ doc: await grid500x100() })
    const a = await open(g.url)
    expect(await resize(a.c, { viewGen: 5, axis: 'col', index: 2, size: 100 })).toMatchObject({ type: 'refused', code: 'stale-view' })
    await viewGen0(a.c)
    a.c.close(); await g.stop()
    // hidden-target: a hidden row under 'full'
    const rows = (await hiddenFixtures()).find((f) => f.id === 'xlsx-hidden-rows')!.withHidden
    const f = await start({ doc: rows, policy: { contentPolicy: 'full' } })
    const b = await open(f.url)
    expect(await resize(b.c, { axis: 'row', index: 1, size: 40 })).toMatchObject({ type: 'refused', code: 'hidden-target' })
    expect(await resize(b.c, { axis: 'row', index: 0, size: 40 })).toMatchObject({ type: 'sheet-view', viewGen: 1 })
    expect(await resize(b.c, { viewGen: 1, axis: 'row', index: 1, size: 40 })).toMatchObject({ type: 'refused', code: 'hidden-target' })
    expect((await layoutOf(b.c)).l.viewGen).toBe(1)
    b.c.close(); await f.stop()
    // coverage-unknown under visible-only
    const u = await start({ doc: await unresolvedAnchor() })
    const d = await open(u.url)
    expect(await resize(d.c, { axis: 'col', index: 0, size: 100 })).toMatchObject({ type: 'refused', code: 'coverage-unknown' })
    await viewGen0(d.c)
    d.c.close(); await u.stop()
  }, 30_000)

  test('the per-axis cap refuses the 4,097th column (SessionViews.apply)', () => {
    const ledger = new ViewStateLedger()
    const v = new SessionViews(ledger)
    for (let i = 0; i < MAX_SESSION_SIZES; i++) expect(v.apply(0, { axis: 'col', index: i, size: 100 })).toEqual({ viewGen: i + 1, applied: 100 })
    expect(MAX_SESSION_SIZES).toBe(4096)
    expect(v.apply(0, { axis: 'col', index: MAX_SESSION_SIZES, size: 100 })).toBe('view-state-limit')
    expect(v.viewGen(0)).toBe(4096)
    // An existing index still changes; rows and other sheets have their own caps.
    expect(v.apply(0, { axis: 'col', index: 7, size: 120 })).toEqual({ viewGen: 4097, applied: 120 })
    expect(v.apply(0, { axis: 'row', index: 0, size: 30 })).toEqual({ viewGen: 4098, applied: 30 })
    expect(v.apply(1, { axis: 'col', index: 0, size: 30 })).toEqual({ viewGen: 1, applied: 30 })
    expect(ledger.bytes).toBe((4096 + 2) * 128)
    v.dispose()
    expect(ledger.bytes).toBe(0)
  })

  test('the process cap refuses the 9th entry of a 1,024-byte ledger', async () => {
    const ledger = new ViewStateLedger(1024)
    const h = await start({ doc: await grid500x100(), viewStateLedger: ledger })
    const a = await open(h.url), b = await open(h.url)
    for (let i = 0; i < 5; i++) expect(await resize(a.c, { viewGen: i, axis: 'col', index: i, size: 100 })).toMatchObject({ type: 'sheet-view', viewGen: i + 1 })
    for (let i = 0; i < 3; i++) expect(await resize(b.c, { viewGen: i, axis: 'row', index: i, size: 30 })).toMatchObject({ type: 'sheet-view', viewGen: i + 1 })
    // An existing index takes no new entry.
    expect(await resize(b.c, { viewGen: 3, axis: 'row', index: 0, size: 31 })).toMatchObject({ type: 'sheet-view', viewGen: 4 })
    expect(await resize(b.c, { viewGen: 4, axis: 'row', index: 9, size: 30 })).toMatchObject({ type: 'refused', code: 'view-state-limit' })
    expect((await layoutOf(b.c)).l.viewGen).toBe(4)
    expect(ledger.bytes).toBe(1024)
    // A reset returns its entries; the teardown of a session returns the rest.
    expect(await resize(b.c, { viewGen: 4, reset: true })).toMatchObject({ type: 'sheet-view', viewGen: 5 })
    expect(ledger.bytes).toBe(5 * 128)
    a.c.close(); await a.c.closed
    for (let i = 0; i < 50 && ledger.bytes > 0; i++) await Bun.sleep(20)
    expect(ledger.bytes).toBe(0)
    b.c.close(); await h.stop()
  }, 30_000)

  test('three tabs may be resized', async () => {
    const h = await start({ doc: await textGrid() })
    const { c } = await open(h.url)
    for (const sheetId of ['0', '1', '2']) expect(await resize(c, { sheetId, axis: 'col', index: 0, size: 90 })).toEqual({ reqId: rq, type: 'sheet-view', sheetId, viewGen: 1, change: { axis: 'col', index: 0, size: 90 } })
    for (const sheetId of ['0', '1', '2']) expect((await layoutOf(c, sheetId)).l.sessionCols).toEqual([[0, 90]])
    c.close(); await h.stop()
  }, 30_000)

  test("client geometry equals the server's under sizes", async () => {
    const bytes = await grid500x100()
    const h = await start({ doc: bytes })
    const { c } = await open(h.url)
    expect(await resize(c, { viewGen: 0, axis: 'col', index: 1, size: 8 })).toMatchObject({ viewGen: 1 })
    expect(await resize(c, { viewGen: 1, axis: 'col', index: 3, size: 200 })).toMatchObject({ viewGen: 2 })
    expect(await resize(c, { viewGen: 2, axis: 'row', index: 4, size: 60 })).toMatchObject({ viewGen: 3 })
    const { l } = await layoutOf(c)
    expect([l.sessionCols, l.sessionRows]).toEqual([[[1, 8], [3, 200]], [[4, 60]]])
    const client = streamSheetGeometry(l).metrics
    const { snap } = await snapshotOf(bytes)
    const g = buildSheetGeometry(snap.sheet, { sizes: { cols: new Map([[1, 8], [3, 200]]), rows: new Map([[4, 60]]) } })
    if (g.kind !== 'ready') throw new Error('server geometry is a diagnostic')
    for (const j of [0, 1, 2, 3, 4, 16383]) expect([j, client.cols.start(j)]).toEqual([j, g.metrics.cols.start(j)])
    for (const i of [0, 4, 5, 499, 1048575]) expect([i, client.rows.start(i)]).toEqual([i, g.metrics.rows.start(i)])
    expect(client.cols.start(4) - client.cols.start(3)).toBe(200)
    c.close(); await h.stop()
  }, 30_000)

  test('a sheet-tiles@1-only layout carries no session sizes and is byte-identical to 4a', async () => {
    for (const bytes of [await grid500x100(), await unresolvedAnchor()]) {
      const h = await start({ doc: bytes })
      const t = await open(h.url, TILES), v = await open(h.url, VIEW)
      const tiles = (await layoutOf(t.c)).parts, view = (await layoutOf(v.c)).parts
      const { host, docKey } = await snapshotOf(bytes, h.stream.sheetClockOf(t.sid)!)
      const expected = encodeSheetLayout(await host.sheetLayout(docKey, 0, h.stream.sheetClockOf(t.sid)!))
      for (const parts of [tiles, view]) {
        expect(parts.map((p) => JSON.stringify(p.text))).toEqual(expected.map((p, i) => JSON.stringify({ reqId: parts[i]!.text.reqId, ...p })))
        for (const p of parts) for (const k of ['sessionRows', 'sessionCols', 'coverage', 'coverageUnknown']) expect([k, JSON.stringify(p.text).includes(k)]).toEqual([k, false])
      }
      t.c.close(); v.c.close(); await h.stop()
    }
  }, 30_000)
})

describe("the 'sheet' class (parent §14)", () => {
  test('the sheet class refuses at 31', async () => {
    let t = Date.now()
    const h = await start({ doc: await grid500x100(), now: () => t })
    const { c } = await open(h.url)
    for (let i = 0; i < 30; i++) {
      t += 60
      expect(await resize(c, { viewGen: i, axis: 'col', index: i, size: 100 })).toMatchObject({ type: 'sheet-view', viewGen: i + 1 })
    }
    t += 60
    expect(await resize(c, { viewGen: 30, axis: 'col', index: 40, size: 100 })).toMatchObject({ type: 'refused', code: 'rate-limited' })
    expect((await layoutOf(c)).l.viewGen).toBe(30)
    c.close(); await h.stop()
  }, 30_000)

  test('a second resize while one is in flight is refused', async () => {
    const gp = gatedPool()
    const h = await start({ doc: await grid500x100(), pool: gp.pool })
    const { c, sid } = await open(h.url)
    const gate = gp.hold()
    const r1 = ++rq, r2 = ++rq
    c.send({ reqId: r1, type: 'sheet-resize', token: c.manifest!.session.token, sheetId: '0', viewGen: 0, axis: 'col', index: 2, size: 100 })
    c.send({ reqId: r2, type: 'sheet-resize', token: c.manifest!.session.token, sheetId: '0', viewGen: 0, axis: 'col', index: 3, size: 100 })
    const second = (await c.next((m) => 'text' in m && m.text.reqId === r2) as { text: Record<string, unknown> }).text
    expect(second).toMatchObject({ type: 'refused', code: 'rate-limited' })
    expect(h.stream.sheetRatesOf(sid)!.sheetBusy).toBe(true)
    gate.open()
    const first = (await c.next((m) => 'text' in m && m.text.reqId === r1) as { text: Record<string, unknown> }).text
    expect(first).toMatchObject({ type: 'sheet-view', viewGen: 1 })
    expect(h.stream.sheetRatesOf(sid)!.sheetBusy).toBe(false)
    c.close(); await h.stop()
  }, 30_000)

  test("a superseded socket's late finally leaves the new op's busy flag", async () => {
    const gp = gatedPool()
    const h = await start({ doc: await textGrid(), pool: gp.pool })
    const a = new Client(h.url)
    const m = await a.hello({ features: VIEW }); await a.ready()
    const sid = m.session.sid
    const gateA = gp.hold(), gateB = gp.hold()
    const ra = ++rq
    a.send({ reqId: ra, type: 'sheet-resize', token: m.session.token, sheetId: '1', viewGen: 0, axis: 'col', index: 0, size: 100 })
    for (let i = 0; i < 100 && !h.stream.sheetRatesOf(sid)!.sheetBusy; i++) await Bun.sleep(10)
    expect(h.stream.sheetRatesOf(sid)!.sheetBusy).toBe(true)
    const b = new Client(h.url)
    await b.hello({ features: VIEW, resume: { sid, resumeProof: m.session.token } })
    expect((await a.closed).code).toBe(4409)
    await b.ready()
    const rb = ++rq
    b.send({ reqId: rb, type: 'sheet-resize', token: b.manifest!.session.token, sheetId: '2', viewGen: 0, axis: 'col', index: 0, size: 100 })
    for (let i = 0; i < 100 && !h.stream.sheetRatesOf(sid)!.sheetBusy; i++) await Bun.sleep(10)
    expect(h.stream.sheetRatesOf(sid)!.sheetBusy).toBe(true)
    // A's op ends (its socket no longer owns the sid): B's flag stays.
    gateA.open()
    await Bun.sleep(300)
    expect(h.stream.sheetRatesOf(sid)!.sheetBusy).toBe(true)
    gateB.open()
    expect((await b.next((x) => 'text' in x && x.text.reqId === rb) as { text: Record<string, unknown> }).text).toMatchObject({ type: 'sheet-view', sheetId: '2', viewGen: 1 })
    expect(h.stream.sheetRatesOf(sid)!.sheetBusy).toBe(false)
    // A's resize was never applied: its socket died before the apply.
    expect((await layoutOf(b, '1')).l.viewGen).toBe(0)
    b.close(); await h.stop()
  }, 30_000)
})

describe('sized tiles (P5, P14)', () => {
  const sizes: WireSizes = [{ axis: 'col', index: 2, px: 1790 }]
  test('sized tiles', async () => {
    const bytes = await grid500x100()
    const h = await start({ doc: bytes, policy: { visibleMark: 'off', forensicMark: false } })
    const { c, sid } = await open(h.url)
    // The 4a grid: 16,384 default columns are 2,048 tiles at zoom 1, dpr 1; column 2,048 is past it.
    expect(tileLimit(16_384 * 64, 1, 1)).toBe(2048)
    const past = { ...tile00, tiles: [{ col: 2048, row: 0 }] }
    expect((await c.tiles(++rq, { sheetId: '0', viewGen: 0, ...past })).end).toMatchObject({ type: 'refused', code: 'policy-refused' })
    const identity = (await c.tiles(++rq, { sheetId: '0', viewGen: 0, ...tile00 })).tiles[0]!
    expect(await resize(c, { axis: 'col', index: 2, size: 2000 })).toMatchObject({ viewGen: 1 })
    expect((await c.tiles(++rq, { sheetId: '0', viewGen: 0, ...tile00 })).end).toMatchObject({ type: 'refused', code: 'stale-view' })
    const r = await c.tiles(++rq, { sheetId: '0', viewGen: 1, ...tile00 })
    expect(r.end).toMatchObject({ type: 'tiles-done', refused: [] })
    expect(r.tiles.map((x) => [x.viewGen, x.col, x.row])).toEqual([[1, 0, 0]])
    // The reference: the snapshot's sheet painted over its geometry at the session's sizes.
    const { host, snap } = await snapshotOf(bytes, h.stream.sheetClockOf(sid)!)
    const g = buildSheetGeometry(snap.sheet, { sizes: sessionSizesOf(sizes) })
    if (g.kind !== 'ready') throw new Error('sized geometry is a diagnostic')
    const canvas = host.backend.createCanvas(T, T)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    const viewport = { box: tileLogicalBox(0, 0, 1, 1), zoom: 1, deviceScale: 1 }
    const frameBudget = frameBudgetFor('D')
    const plan = planViewport(g as SheetGeometry, viewport, 'D', { frameBudget })
    ctx.setTransform(1, 0, 0, 1, 0, 0)
    renderSheetViewport(snap.sheet, ctx, g, viewport, snap.assets, { colors: TILE_PALETTE_COLORS.light, preparedPlan: plan, profile: 'D', cache: new LayoutCache(), frameBudget })
    const reference = ctx.getImageData(0, 0, T, T).data
    const got = decodePng(r.tiles[0]!.png).rgba, before = decodePng(identity.png).rgba
    let diff = 0, changed = 0
    for (let i = 0; i < got.length; i++) { if (got[i] !== reference[i]) diff++; if (got[i] !== before[i]) changed++ }
    expect(diff).toBe(0)
    expect(changed).toBeGreaterThan(1000)
    // The extent grew by 1,726 px: column 2,048 is admitted now.
    const grown = await c.tiles(++rq, { sheetId: '0', viewGen: 1, ...past })
    expect(grown.end).toMatchObject({ type: 'tiles-done', refused: [] })
    expect(grown.tiles.map((x) => x.col)).toEqual([2048])
    c.close(); await h.stop()
  }, 30_000)

  test('same sizes share canonical tiles; different sizes never do', async () => {
    const cp = countingPool()
    const h = await start({ doc: await grid500x100(), pool: cp.pool, policy: { visibleMark: 'fixed', forensicMark: true } })
    const got: Uint8Array[] = []
    for (const [i, size] of [[0, 300], [1, 300], [2, 301]] as const) {
      const { c } = await open(h.url)
      expect(await resize(c, { axis: 'col', index: 2, size })).toMatchObject({ viewGen: 1 })
      const r = await c.tiles(++rq, { sheetId: '0', viewGen: 1, ...tile00 })
      expect(r.end).toMatchObject({ type: 'tiles-done', refused: [] })
      got.push(r.tiles[0]!.png)
      expect([i, cp.n.tiles]).toEqual([i, i === 0 ? 1 : i === 1 ? 1 : 2])
      c.close()
    }
    expect(Buffer.from(got[0]!).equals(Buffer.from(got[1]!))).toBe(false)
    await h.stop()
  }, 30_000)

  test('sheet_views records the sizes', async () => {
    const store = AuditStore.open(':memory:')
    const h = await start({ doc: await grid500x100(), audit: new AuditTable({ store }) })
    const { c, sid } = await open(h.url)
    expect((await c.tiles(++rq, { sheetId: '0', viewGen: 0, ...tile00 })).end).toMatchObject({ type: 'tiles-done' })
    expect(await resize(c, { axis: 'col', index: 2, size: 2000 })).toMatchObject({ viewGen: 1 })
    expect((await c.tiles(++rq, { sheetId: '0', viewGen: 1, ...tile00 })).end).toMatchObject({ type: 'tiles-done' })
    expect(store.sheetView(sid, '0', 0)!.spec.sizes).toEqual([])
    expect(store.sheetView(sid, '0', 1)!.spec.sizes).toEqual([{ axis: 'col', index: 2, px: 1790 }])
    c.close(); await h.stop()
  }, 30_000)

  test('a resize stops in-flight batches', async () => {
    const h = await start({ doc: await grid500x100() })
    const { c } = await open(h.url)
    await layoutOf(c)
    const tiles = []
    for (let row = 0; row < 8; row++) for (let col = 0; col < 8; col++) tiles.push({ col, row })
    const reqId = ++rq
    c.send({ reqId, type: 'tiles', token: c.manifest!.session.token, sheetId: '0', viewGen: 0, zoomStep: 5, dpr: 1, palette: 'light', tiles })
    await c.next((m) => 'binary' in m && m.binary.reqId === reqId, 20_000)
    expect(await resize(c, { axis: 'col', index: 2, size: 200 })).toMatchObject({ viewGen: 1 })
    const r = await c.tiles(reqId, {}, { send: false, ms: 20_000 })
    expect(r.end).toMatchObject({ type: 'tiles-done', cancelled: true })
    expect(r.tiles.length).toBeLessThan(63)
    c.close(); await h.stop()
  }, 30_000)
})
