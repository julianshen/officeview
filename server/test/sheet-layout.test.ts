/**
 * Stage 4a Task 6 — `sheet-layout` (S4 §6.4, R14): paint-facing geometry only, chunked under 64 KiB, capped, and
 * charged to `'unit'` once per request (S4 §6.8). The client rebuilds the server's axes from it
 * (`streamSheetGeometry`), so every position the tiles paint at agrees.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { hiddenFixtures } from '../../scripts/stream-s0/hidden-fixtures'
import { decodeSheetLayout, type StreamSheetLayout } from '../../src/stream/sheet-wire'
import { streamSheetGeometry } from '../../src/xlsx/view/streamGeometry'
import { buildSheetGeometry, type SheetGeometry } from '../../src/xlsx/view/viewport'
import { DocumentHost } from '../src/render'
import type { SheetClock } from '../src/sheet/clock'
import { augmentedFixtures } from './fixtures'
import { Client, disposePool, start } from './harness'
import { AUTO_FILTER, autoFilterCollapsed, FROZEN_MERGES, frozenMerges, grid500x100, insertSheetElement, manyMerges, millionRows, patchSheetXml, patchXlsx } from './sheet-fixtures'

afterAll(async () => {
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
  await disposePool()
})

const TILES = ['sheet-tiles@1']
const UTC: SheetClock = { day: '2026-11-23', timeZone: 'UTC' }
/** Calibrated (Task 6): this many merges, 64 per row (a static unit inside the manifest's 8192 px), encode to
 * exactly four parts. */
const FOUR_PART_MERGES = 6_000
const fourPartFixture = () => manyMerges(FOUR_PART_MERGES, { perRow: 64 })
type Policy = 'visible-only' | 'full'

const hosts: DocumentHost[] = []
/** The server's own geometry of sheet 0: `buildSheetGeometry` over the servable snapshot's sheet. */
async function serverGeometry(bytes: Uint8Array, policy: Policy = 'visible-only') {
  const host = new DocumentHost(); hosts.push(host)
  const o = await host.open('ref', bytes, policy, { sheetMode: 'tiles' })
  const snap = await host.sheetSnapshot(o.docKey, 0, UTC)
  const g = buildSheetGeometry(snap.sheet)
  if (g.kind !== 'ready') throw new Error('server geometry is a diagnostic')
  return { geometry: g as SheetGeometry, sheet: snap.sheet }
}

async function tilesClient(url: string) {
  const c = new Client(url)
  const m = await c.hello({ features: TILES })
  await c.ready()
  return { c, m }
}
/**
 * FINDING (pre-existing, Task 6 report): a sheet whose static unit is over 8192 px (`grid500x100` is 6400 × 10000)
 * gets a manifest `validateManifest` rejects (`units.list[0].h`, capped at 8192), in units and tiles mode alike.
 * This HELLO keeps the manifest unvalidated so the layout of such a sheet can still be tested; remove it (use
 * `tilesClient`) once the manifest is fixed.
 */
async function unvalidatedTilesClient(url: string) {
  const c = new Client(url)
  await c.opened()
  c.send({ reqId: 1, type: 'hello', client: 'officeview/1.0', recVersions: [1], features: TILES })
  const m = await c.next((x) => 'binary' in x && x.binary.kind === 1) as { binary: { payload: Uint8Array } }
  c.manifest = JSON.parse(new TextDecoder().decode(m.binary.payload))
  await c.ready()
  return { c, m: c.manifest! }
}
async function layoutOf(bytes: Uint8Array, o: { sheetId?: string; policy?: Policy; unvalidated?: boolean } = {}) {
  const sheetId = o.sheetId ?? '0'
  const h = await start({ doc: bytes, ...(o.policy === 'full' ? { policy: { contentPolicy: 'full' as const } } : {}) })
  const { c } = await (o.unvalidated ? unvalidatedTilesClient : tilesClient)(h.url)
  const r = await c.sheetLayout(10, sheetId)
  c.close(); await h.stop()
  if (r.end) throw new Error(`refused: ${JSON.stringify(r.end)}`)
  return r.parts
}
const decoded = (parts: Array<{ text: Record<string, any> }>): StreamSheetLayout => decodeSheetLayout(parts.map((p) => p.text))
const rect = (m: { minRow: number; maxRow: number; minCol: number; maxCol: number }) => ({ r0: m.minRow, r1: m.maxRow, c0: m.minCol, c1: m.maxCol })

describe('sheet-layout (S4 §6.4)', () => {
  test('layout round-trips and matches the server geometry', async () => {
    const rowsAt = [0, 1, 2, 9, 10, 20, 499, 65_535, 1_048_575], colsAt = [0, 1, 99, 16_383]
    // FINDING (pre-existing, Task 6 report): a frozen pane's `state="frozen"` fails the visible-only projection's
    // visibility scan (`unowned visibility flag pane state=frozen`), so the frozen fixture is served under 'full'.
    const cases: Array<[() => Promise<Uint8Array>, Policy, boolean]> = [[grid500x100, 'visible-only', true], [millionRows, 'visible-only', false], [frozenMerges, 'full', false]]
    for (const [make, policy, unvalidated] of cases) {
      const bytes = await make()
      const layout = decoded(await layoutOf(bytes, { policy, unvalidated }))
      const client = streamSheetGeometry(layout)
      const { geometry: server, sheet } = await serverGeometry(bytes, policy)
      for (const i of rowsAt) expect([make.name, i, client.metrics.rows.start(i)]).toEqual([make.name, i, server.metrics.rows.start(i)])
      for (const j of colsAt) expect([make.name, j, client.metrics.cols.start(j)]).toEqual([make.name, j, server.metrics.cols.start(j)])
      expect(client.metrics.rows.totalCss()).toBe(server.metrics.rows.totalCss())
      expect(client.metrics.cols.totalCss()).toBe(server.metrics.cols.totalCss())
      expect(layout.frozen).toEqual(sheet.frozen)
      expect(layout.merges).toEqual((sheet.mergeRanges ?? []).map(rect))
      expect(client.cellsOf(0)).toBeUndefined()
      expect(client.drawingsIn(0, 0, 1e6, 1e6, 1000)).toEqual([])
    }
    // The fixtures really exercise sizes, frozen panes and merges.
    const fm = decoded(await layoutOf(await frozenMerges(), { policy: 'full' }))
    expect(fm.frozen).toEqual({ rows: 3, cols: 2, topRow: 3, leftCol: 2 })
    expect(fm.merges).toHaveLength(FROZEN_MERGES.length)
    expect(fm.rows.length).toBeGreaterThan(0)
    expect(fm.cols.length).toBeGreaterThan(0)
    expect(decoded(await layoutOf(await millionRows())).rows).toEqual([{ start: 9, end: 19, sizeCss: 40 }])
  }, 120_000)

  test('hidden sizes never travel', async () => {
    const fx = await hiddenFixtures()
    const rowsFx = fx.find((f) => f.id === 'xlsx-hidden-rows')!, colsFx = fx.find((f) => f.id === 'xlsx-hidden-cols')!
    const tallHidden = await patchXlsx(rowsFx.withHidden, (zip) => patchSheetXml(zip, 0, (xml) => {
      if (!xml.includes('<row r="2" hidden="1">')) throw new Error('fixture: no hidden row 2')
      // A2:B2 lies wholly inside the hidden row (it paints nothing); C1:C3 spans it and a visible row.
      return insertSheetElement(xml.replace('<row r="2" hidden="1">', '<row r="2" hidden="1" ht="40" customHeight="1">'), 'mergeCells', '<mergeCells count="2"><mergeCell ref="A2:B2"/><mergeCell ref="C1:C3"/></mergeCells>')
    }))
    const wideHidden = await patchXlsx(colsFx.withHidden, (zip) => patchSheetXml(zip, 0, (xml) => {
      if (!xml.includes('width="10" hidden="1"')) throw new Error('fixture: no hidden column')
      return xml.replace('width="10" hidden="1"', 'width="30" hidden="1"')
    }))
    const covers = (list: Array<{ start: number; end: number }>, i: number) => list.some((x) => x.start <= i && i <= x.end)
    const r = decoded(await layoutOf(tallHidden))
    expect(covers(r.rows, 1)).toBe(false)
    expect(covers(r.collapsedRows, 1)).toBe(true)
    expect(r.collapsedRows).toEqual([{ start: 1, end: 1 }])
    // A merge wholly inside a collapsed band does not travel (builder's call); one that spans it does.
    expect(r.merges).toEqual([{ r0: 0, r1: 2, c0: 2, c1: 2 }])
    const c = decoded(await layoutOf(wideHidden))
    expect(covers(c.cols, 1)).toBe(false)
    expect(covers(c.collapsedCols, 1)).toBe(true)
    expect(c.collapsedCols).toEqual([{ start: 1, end: 1 }])
  }, 60_000)

  test('layout keys are on the allowlist', async () => {
    const allowed = new Set([
      // StreamSheetLayout and its nested field names.
      'sheetId', 'viewGen', 'defaults', 'rowPx', 'colPx', 'rows', 'cols', 'start', 'end', 'sizeCss', 'collapsedRows', 'collapsedCols',
      'frozen', 'topRow', 'leftCol', 'merges', 'r0', 'r1', 'c0', 'c1', 'used', 'filters', 'kind', 'range', 'header', 'totals',
      'sort', 'filterSummary', 'col', 'authorHidden', 'counts', 'visibleRows', 'totalRows',
      // The envelope.
      'reqId', 'type', 'seq', 'last', 'head',
    ])
    const keysOf = (v: unknown, out: Set<string>): Set<string> => {
      if (Array.isArray(v)) for (const x of v) keysOf(x, out)
      else if (v && typeof v === 'object') for (const [k, x] of Object.entries(v)) { out.add(k); keysOf(x, out) }
      return out
    }
    const xlsx = [...(await hiddenFixtures()), ...(await augmentedFixtures())].filter((f) => f.kind === 'xlsx')
    expect(xlsx.length).toBeGreaterThan(3)
    for (const f of xlsx) {
      const h = await start({ doc: f.withHidden })
      const { c, m } = await tilesClient(h.url)
      expect(m.sheetMode).toBe('tiles')
      let reqId = 20
      for (const { sheetId } of m.sheets!) {
        const r = await c.sheetLayout(++reqId, sheetId)
        expect(r.end).toBeUndefined()
        for (const p of r.parts) {
          const extra = [...keysOf(p.text, new Set())].filter((k) => !allowed.has(k))
          expect([f.id, sheetId, extra]).toEqual([f.id, sheetId, []])
          const json = JSON.stringify(p.text)
          expect([f.id, json.includes('MARKER-'), json.includes('VIS-')]).toEqual([f.id, false, false])
        }
      }
      c.close(); await h.stop()
    }
  }, 120_000)

  test('4a view fields', async () => {
    const l = decoded(await layoutOf(await autoFilterCollapsed()))
    const totalRows = AUTO_FILTER.lastRow
    expect(l.collapsedRows).toEqual([{ start: AUTO_FILTER.hiddenInside - 1, end: AUTO_FILTER.hiddenInside - 1 }, { start: AUTO_FILTER.hiddenOutside - 1, end: AUTO_FILTER.hiddenOutside - 1 }])
    expect(l.authorHidden).toBe(true)
    expect(l.counts).toEqual({ visibleRows: totalRows - 2, totalRows })
    expect(l.sort).toBeNull()
    expect(l.filterSummary).toEqual([])
    expect(l.viewGen).toBe(0)
    expect(l.filters).toEqual([{ kind: 'autoFilter', range: { r0: 0, r1: 7, c0: 0, c1: 1 }, header: 1, totals: 0 }])
    expect(l.used).toEqual({ r0: 0, r1: AUTO_FILTER.lastRow - 1, c0: 0, c1: 1 })
    // No filter range: nothing is author-hidden by a filter, whatever is collapsed.
    const plain = decoded(await layoutOf((await hiddenFixtures()).find((f) => f.id === 'xlsx-hidden-rows')!.withHidden))
    expect(plain.filters).toEqual([])
    expect(plain.authorHidden).toBe(false)
  }, 60_000)

  test('layout chunks under 64 KiB', async () => {
    const parts = await layoutOf(await manyMerges(20_000))
    expect(parts.length).toBeGreaterThanOrEqual(2)
    parts.forEach((p, i) => {
      expect(p.bytes).toBeLessThanOrEqual(65_536)
      expect(p.text.seq).toBe(i)
      expect(p.text.last).toBe(i === parts.length - 1)
      expect('head' in p.text).toBe(i === 0)
    })
    expect(parts.filter((p) => p.text.last === true)).toHaveLength(1)
    expect(decoded(parts).merges).toHaveLength(20_000)
  }, 60_000)

  test('over a cap is refused', async () => {
    const h = await start({ doc: await manyMerges(70_000) })
    const { c } = await tilesClient(h.url)
    const r = await c.sheetLayout(11, '0', { ms: 30_000 })
    expect(r.parts).toEqual([])
    expect(r.end).toMatchObject({ reqId: 11, type: 'refused', code: 'layout-too-large' })
    // The session survives the refusal.
    expect(c.ws.readyState).toBe(WebSocket.OPEN)
    c.close(); await h.stop()
  }, 60_000)

  test('layout is charged to unit once', async () => {
    const bytes = await fourPartFixture()
    // F13: every request also charges the 'message' rate (20/s burst; its breaches count toward 4008), so the injected
    // clock advances 100 ms per request: 61 requests span 6.1 s, inside one 'unit' minute.
    let t = Date.UTC(2026, 10, 23, 9, 0, 0)
    const h = await start({ doc: bytes, now: () => t })
    // 60 requests in a minute are answered; the 61st is refused.
    {
      const { c } = await tilesClient(h.url)
      for (let i = 0; i < 60; i++) {
        t += 100
        const r = await c.sheetLayout(100 + i, '0')
        expect([i, r.end]).toEqual([i, undefined])
        expect(r.parts.length).toBe(4)
      }
      t += 100
      const r = await c.sheetLayout(200, '0')
      expect(r.end).toMatchObject({ reqId: 200, type: 'refused', code: 'rate-limited' })
      await Bun.sleep(50)
      expect(c.ws.readyState).toBe(WebSocket.OPEN)
      c.close()
    }
    // Later parts charge nothing: 30 four-part answers, then 30 more requests, are all answered.
    {
      const { c } = await tilesClient(h.url)
      for (let i = 0; i < 30; i++) {
        t += 100
        const r = await c.sheetLayout(300 + i, '0')
        expect([i, r.end, r.parts.length]).toEqual([i, undefined, 4])
      }
      for (let i = 0; i < 30; i++) {
        t += 100
        const r = await c.sheetLayout(400 + i, '0')
        expect([i, r.end]).toEqual([i, undefined])
      }
      expect(c.ws.readyState).toBe(WebSocket.OPEN)
      c.close()
    }
    // An unknown or malformed sheetId is refused policy-refused.
    {
      const { c } = await tilesClient(h.url)
      let reqId = 500
      for (const sheetId of ['1', '01', '-1', 'x', '', 0, null, undefined, '4294967296']) {
        t += 100
        const r = await c.sheetLayout(++reqId, sheetId)
        expect([sheetId, r.end]).toMatchObject([sheetId, { type: 'refused', code: 'policy-refused' }])
      }
      c.close()
    }
    await h.stop()
  }, 180_000)

  test('units mode still refuses sheet-layout', async () => {
    const h = await start({ doc: await autoFilterCollapsed() })
    const c = new Client(h.url)
    await c.hello()
    await c.ready()
    const r = await c.sheetLayout(12, '0')
    expect(r.end).toMatchObject({ reqId: 12, type: 'refused', code: 'unsupported-op' })
    c.close(); await h.stop()
  }, 60_000)
})
