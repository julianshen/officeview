/**
 * Stage 4c Task 8 — `sheet-sort`, `sheet-filter`, `sheet-order` and cancel on the socket (S4 §6.3, §6.6; plan P2–P4,
 * P7, P8, P11–P13; parent §12 C1). A sort or filter is authorized against the authored layout and the sid's issued
 * lists before any engine run, runs as one pool job over the key view, commits synchronously (with its admission hold)
 * and is answered by `sheet-order` parts stamped with the session's `viewGen`. Tiles and text answers follow the
 * session's view; two sids with the same view share canonical tiles, two with different views never do (T6).
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { viewerControls } from '../../src/components/xlsx/viewerControlsClient'
import type { CacheIdentity, FilterRange, FilterSpec, SheetView, SortSpec } from '../../src/xlsx/view/contracts'
import { MainMemoryLedger } from '../../src/xlsx/view/mainLedger'
import { createIdentityOrder } from '../../src/xlsx/view/order'
import { emptyGridSelection } from '../../src/xlsx/view/selection'
import { createObjectWorkbookSource } from '../../src/xlsx/view/source'
import { visibilityFromHidden } from '../../src/xlsx/view/visibility'
import type { XlsxCell, XlsxDocument } from '../../src/xlsx/types'
import { buildXlsx, type XlsxCellSpec } from '../../src/testdata/ooxml-builders'
import { decodeSheetLayout, decodeSheetOrder, decodeTileBody, decodeValuesPage, LAYOUT_PART_BYTES, ORDER_LIMITS, TILE_ZOOM_STEPS, type SheetOrder, type StreamSheetLayout, type WireRange, type WireSortSpec } from '../../src/stream/sheet-wire'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import type { Pool, PoolSheetViewJob, PoolTileJob } from '../src/pool'
import { DocumentHost } from '../src/render'
import type { SheetViewReply } from '../src/sheet/view-ops'
import { orderBytes, type CanonicalView, type ResolvedFilter } from '../src/sheet/view-spec'
import { Client, disposePool, pool, start } from './harness'
import { streamSheetGeometry } from '../../src/xlsx/view/streamGeometry'
import { tileLimit } from '../src/sheet/tile'
import { FILTER_AUTO_HIDDEN, filterAutoHidden, patchSheetXml, patchXlsx, setRowAttrs, ONE_MILLION_ROWS, oneMillionSort, SORT_GRID, sortGrid, sortMerges, sortUnderShape, unresolvedAnchor } from './sheet-fixtures'

const hosts: DocumentHost[] = []
afterAll(async () => {
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
  await disposePool()
})

const VIEW = ['sheet-view@1']
const T0 = Date.UTC(2026, 10, 27, 12)
const wire = (r0: number, r1: number, c0: number, c1: number, header: 0 | 1 = 1): WireRange => ({ r0, r1, c0, c1, header, totals: 0 })
const local = (r: WireRange): FilterRange => ({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1, headerRows: r.header, totalsRows: r.totals })
const GRID = wire(0, SORT_GRID.rows - 1, 0, SORT_GRID.cols - 1)
const B_ASC: WireSortSpec['keys'] = [{ col: 1, order: 'asc' }]
const B_DESC: WireSortSpec['keys'] = [{ col: 1, order: 'desc' }]
const text = (op: string, first: string) => ({ kind: 'text', op, first, wildcards: false })
const tile00 = { zoomStep: 5, dpr: 1, palette: 'light', tiles: [{ col: 0, row: 0 }] }

type Part = { text: Record<string, any>; bytes: number }
type Answer = { parts: Part[]; order?: SheetOrder; end?: Record<string, any> }
const stripped = (parts: Part[]) => parts.map(({ text: { reqId: _r, ...t } }) => t)

/** The shared pool behind a seam: records `sheetView` and `tile` jobs, and can hold the next sort or filter job. */
function seamPool() {
  const base = pool()
  const seen = { views: [] as PoolSheetViewJob[], tiles: [] as PoolTileJob[] }
  let hold: { entered: () => void; gate: Promise<void> } | undefined
  const proxy = new Proxy(base, {
    get(t, k) {
      if (k === 'sheetView') return async (...a: Parameters<Pool['sheetView']>) => {
        seen.views.push(a[1])
        const h = hold
        if (h && (a[1].op === 'sort' || a[1].op === 'filter')) { hold = undefined; h.entered(); await h.gate }
        return t.sheetView(...a)
      }
      if (k === 'tile') return (...a: Parameters<Pool['tile']>) => { seen.tiles.push(a[1]); return t.tile(...a) }
      const v = Reflect.get(t, k, t)
      return typeof v === 'function' ? v.bind(t) : v
    },
  })
  /** Holds the next sort or filter job until `release`; `entered` resolves once it reached the pool. */
  function holdNext() {
    let release!: () => void, entered!: () => void
    const gate = new Promise<void>((r) => { release = r })
    const reached = new Promise<void>((r) => { entered = r })
    hold = { entered, gate }
    return { release, reached }
  }
  return { proxy, seen, holdNext, ops: () => seen.views.map((j) => j.op) }
}

/** A server on an injected clock (each request steps it 100 ms, so `'message'` never refuses), with its own audit. */
async function serverOn(doc: Uint8Array, o: Parameters<typeof start>[0] = {}) {
  const clock = { t: T0 }
  const store = AuditStore.open(':memory:')
  const h = await start({ doc, now: () => clock.t, tokenTtlSec: 86_400, audit: new AuditTable({ store }), ...o })
  let rq = 1000
  /** A new session, or (`resume`) the resume of `resume.sid` on a new socket. */
  async function open(resume?: { sid: string; token: string }) {
    const c = new Client(h.url)
    const m = await c.hello({ features: VIEW, ...(resume ? { resume: { sid: resume.sid, resumeProof: resume.token } } : {}) })
    await c.ready()
    while (h.stream.table.get(m.session.sid)?.status !== 'stream') await Bun.sleep(5)
    const sid = m.session.sid as string
    const send = (type: string, body: Record<string, unknown>) => {
      clock.t += 100
      const reqId = ++rq
      c.send({ token: m.session.token, reqId, type, sheetId: '0', ...body })
      return reqId
    }
    /** Every `sheet-order` part of `reqId` up to `last` (decoded), or the other answer as `end`. */
    async function answer(reqId: number): Promise<Answer> {
      const parts: Part[] = []
      for (;;) {
        const x = await c.next((y) => 'text' in y && y.text.reqId === reqId, 20_000) as Part
        if (x.text.type !== 'sheet-order') return { parts, end: x.text }
        parts.push(x)
        if (x.text.last === true) return { parts, order: decodeSheetOrder(parts.map((p) => p.text)) }
      }
    }
    const op = (type: string, body: Record<string, unknown>) => answer(send(type, body))
    const sort = (body: Record<string, unknown>) => op('sheet-sort', { viewGen: 0, ...body })
    const filter = (body: Record<string, unknown>) => op('sheet-filter', { viewGen: 0, ...body })
    const order = (sheetId = '0') => op('sheet-order', { sheetId })
    async function layout(sheetId = '0'): Promise<StreamSheetLayout> {
      clock.t += 100
      const r = await c.sheetLayout(++rq, sheetId)
      if (r.end) throw new Error(`layout refused: ${JSON.stringify(r.end)}`)
      return decodeSheetLayout(r.parts.map((p) => p.text)) as StreamSheetLayout
    }
    async function list(range: WireRange, col: number, viewGen = 0, sheetId = '0') {
      const reqId = send('filter-values', { sheetId, viewGen, range, col })
      const a = await c.next((y) => 'text' in y && y.text.reqId === reqId) as Part
      if (a.text.type !== 'filter-values-result') throw new Error(`values refused: ${JSON.stringify(a.text)}`)
      return decodeValuesPage(a.text)
    }
    async function tiles(viewGen: number, body: Record<string, unknown> = tile00) {
      clock.t += 100
      return c.tiles(++rq, { sheetId: '0', viewGen, ...body })
    }
    return { c, sid, token: m.session.token as string, send, answer, sort, filter, order, layout, list, tiles }
  }
  return { h, store, clock, open }
}

/** `[col, kind]` rows of sort or filter kind with outcome `ok`. */
const okRows = (rows: Array<{ kind: string; outcome: string; sheetId: string }>, sheetId = '0') => rows.filter((r) => (r.kind === 'sort' || r.kind === 'filter') && r.outcome === 'ok' && r.sheetId === sheetId)

// ---- the oracle: local viewerControls over the fixture parsed apart, its withheld cells deleted ----------------------

async function oracleDoc(bytes: Uint8Array, drop: (row: number, col: number) => boolean = () => false): Promise<XlsxDocument> {
  const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
  const sheet = doc.sheets[0]!
  const rows = sheet.rows.map((row) => ({ ...row, cells: row.cells.filter((c: XlsxCell) => !drop(c.row, c.col)) }))
  return { ...doc, sheets: [{ ...sheet, rows }, ...doc.sheets.slice(1)] }
}
/** `sortGrid` as the oracle sees it under `visible-only`. */
const gridOracle = (bytes: Uint8Array) => oracleDoc(bytes, (r, c) => c === SORT_GRID.zeroWidthCol || r === SORT_GRID.zeroHeightRow || r === SORT_GRID.hiddenRow)

async function oracle(doc: XlsxDocument) {
  const source = createObjectWorkbookSource(doc, 'oracle')
  const sheetId = source.sheets[0]!.sheetId
  const baseHidden = source.geometry(sheetId).baseHidden
  let view: SheetView = { sheetId, revision: 0, order: createIdentityOrder(), baseHidden, visibility: visibilityFromHidden(baseHidden), sort: null, filter: null, selection: emptyGridSelection() }
  let viewGen = 0
  const signal = () => new AbortController().signal
  const owner = (): CacheIdentity => ({ docId: 'oracle', sheetId, valueGen: source.store(sheetId).stamp.valueGen, fontGen: 0, viewGen })
  const snapshot = () => ({ owner: owner(), filter: view.filter, sort: view.sort, records: { visible: 0, total: 0, completeness: 'complete' as const } })
  const publish = (next: Pick<SheetView, 'order' | 'visibility' | 'filter' | 'sort'>) => { view = { ...view, ...next }; viewGen++; return snapshot() }
  const controls = await viewerControls({ doc, source, sheetId, snapshot, view: () => view, publish, ledger: new MainMemoryLedger() }, signal())
  return {
    async sort(spec: SortSpec) { return (await controls.client.sort(owner(), spec, signal())).kind },
    async filter(spec: FilterSpec | null) { return (await controls.client.filter(owner(), spec, signal())).kind },
    /** P7's counts over the body rows of `r`. */
    counts(r: WireRange) {
      let total = 0, visible = 0
      for (let d = r.r0 + r.header; d <= r.r1 - r.totals; d++) { total++; if (view.visibility.visible(d)) visible++ }
      return { visible, total }
    },
    shown(r0: number, r1: number) {
      const out = new Set<number>()
      for (let d = r0; d <= r1; d++) if (view.visibility.visible(d)) out.add(view.order.displayToSource(d))
      return out
    },
    dispose: async () => { controls.dispose(); await source.release() },
  }
}

/** Source rows of `[r0, r1]` an order shows (display slot through the band, then the mask). */
function shownSources(order: SheetOrder, r0: number, r1: number, authoredHidden: ReadonlySet<number> = new Set()): Set<number> {
  const out = new Set<number>()
  const m = order.mask
  for (let d = r0; d <= r1; d++) {
    const band = order.bands.find((b) => b.r0 <= d && d <= b.r1)
    const src = band ? band.sourceAtOffset[d - band.r0]! : d
    if (!m || d < m.r0 || d > m.r1) { if (!authoredHidden.has(src)) out.add(src); continue }
    const i = d - m.r0
    if ((m.words[i >>> 5]! >>> (i & 31)) & 1) out.add(src)
  }
  return out
}

async function hostOf(bytes: Uint8Array, policy: 'visible-only' | 'full') {
  const host = new DocumentHost(); hosts.push(host)
  const { docKey } = await host.open('doc', bytes, policy, { sheetMode: 'tiles' })
  return (clock: { day: string; timeZone: string }, view: CanonicalView) => host.sheetView(docKey, clock, policy, { op: 'derive', sheetIndex: 0, view })
}
const orderOf = (r: SheetViewReply): SheetOrder => {
  if (r.op !== 'view' || !('order' in r)) throw new Error(`no view: ${JSON.stringify(r)}`)
  return r.order
}

/** `HIDDEN_TAIL`: A1 `Key`, A2:A41 `keep-{r}`, A42:A61 `drop-{r}` in rows saved `hidden="1"` (no autoFilter). */
const HIDDEN_TAIL = Object.freeze({ shown: 41, hidden: Object.freeze(Array.from({ length: 20 }, (_, i) => 42 + i)) })
async function hiddenTail(): Promise<Uint8Array> {
  const texts = ['Key', ...Array.from({ length: 60 }, (_, i) => `${i + 2 > HIDDEN_TAIL.shown ? 'drop' : 'keep'}-${i + 2}`)]
  const rows = texts.map((_, i) => ({ r: i + 1, cells: [{ ref: `A${i + 1}`, t: 's', v: i } as XlsxCellSpec] }))
  return patchXlsx(await buildXlsx([{ name: 'S', rows }], texts), (zip) => patchSheetXml(zip, 0, (xml) => setRowAttrs(xml, HIDDEN_TAIL.hidden, { hidden: '1' })))
}

/** Three sheets `A`, `B`, `C`, each A1 = `H` then 1..5 down column A. */
function threeSheets(): Promise<Uint8Array> {
  const rows = [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 0 } as XlsxCellSpec] }, ...[5, 3, 1, 4, 2].map((v, i) => ({ r: i + 2, cells: [{ ref: `A${i + 2}`, v } as XlsxCellSpec] }))]
  return buildXlsx(['A', 'B', 'C'].map((name) => ({ name, rows })), ['H'])
}

describe('sheet-sort and sheet-filter (S4 §6.6; plan P4, P7)', () => {
  test('sort round trip', async () => {
    const bytes = await sortGrid()
    const s = await serverOn(bytes)
    const a = await s.open()
    const r = await a.sort({ range: GRID, keys: B_ASC })
    expect(r.end).toBeUndefined()
    const o = r.order!
    expect(o.head.viewGen).toBe(1)
    expect(o.head.sort).toEqual({ range: GRID, keys: B_ASC })
    // The band (and mask) equal Task 5's job over the same view; the counts equal the local engine's.
    const derive = await hostOf(bytes, 'visible-only')
    const want = orderOf(await derive(s.h.stream.sheetClockOf(a.sid)!, { sort: { range: GRID, keys: B_ASC, caseSensitive: false }, sortFilter: null, filter: null }))
    expect(o.bands.map((b) => ({ ...b, sourceAtOffset: Array.from(b.sourceAtOffset) }))).toEqual(want.bands.map((b) => ({ ...b, sourceAtOffset: Array.from(b.sourceAtOffset) })))
    expect(o.mask && { ...o.mask, words: Array.from(o.mask.words) }).toEqual(want.mask && { ...want.mask, words: Array.from(want.mask.words) })
    const or = await oracle(await gridOracle(bytes))
    try {
      expect(await or.sort({ range: local(GRID), keys: B_ASC, caseSensitive: false })).toBe('committed')
      expect(o.head.counts).toEqual(or.counts(GRID))
    } finally { await or.dispose() }
    for (const p of r.parts) expect(p.bytes).toBeLessThanOrEqual(LAYOUT_PART_BYTES)
    const l = await a.layout()
    expect(l.viewGen).toBe(1)
    expect(l.sort?.keys).toEqual(B_ASC)
    expect(l.filterSummary).toEqual([])
    const again = await a.order()
    expect(stripped(again.parts)).toEqual(stripped(r.parts))
    a.c.close(); await s.h.stop()
  }, 30_000)

  test('authorization refuses before anything moves', async () => {
    const bytes = await sortGrid()
    const seam = seamPool()
    const s = await serverOn(bytes, { pool: seam.proxy })
    const a = await s.open(), other = await s.open()
    const R3 = wire(0, 3, 0, SORT_GRID.cols - 1)
    const mine = await a.list(R3, 2), theirs = await other.list(R3, 2)
    expect(mine.total).toBe(3)
    const values = (list: string, ids: Array<[number, number]>) => ({ kind: 'values', list, mode: 'include', ids, blanks: false })
    const cases: Array<[string, () => Promise<Answer>, string]> = [
      ['old viewGen', () => a.sort({ viewGen: 5, range: GRID, keys: B_ASC }), 'stale-view'],
      ['key on E', () => a.sort({ range: GRID, keys: [{ col: SORT_GRID.zeroWidthCol, order: 'asc' }] }), 'hidden-target'],
      ['c1 = E', () => a.sort({ range: wire(0, 39, 0, SORT_GRID.zeroWidthCol), keys: B_ASC }), 'hidden-target'],
      ['filter column E', () => a.filter({ range: GRID, columns: [[SORT_GRID.zeroWidthCol, text('contains', '1')]] }), 'hidden-target'],
      ['number gt', () => a.filter({ range: GRID, columns: [[1, { kind: 'number', op: 'gt', first: 50 }]] }), 'filter-condition-refused'],
      ['compound(text, datePreset)', () => a.filter({ range: GRID, columns: [[3, { kind: 'compound', connective: 'and', first: text('contains', 'a'), second: { kind: 'datePreset', preset: 'today' } }]] }), 'filter-condition-refused'],
      ['date eq', () => a.filter({ range: GRID, columns: [[5, { kind: 'date', op: 'eq', first: { year: 2026, month: 11, day: 25 } }]] }), 'filter-condition-refused'],
      ["another sid's list", () => a.filter({ range: R3, columns: [[2, values(theirs.list, [[0, 0]])]] }), 'stale-filter-values'],
      ['a C list on D', () => a.filter({ range: R3, columns: [[3, values(mine.list, [[0, 0]])]] }), 'stale-filter-values'],
      ['position 99 of 3', () => a.filter({ range: R3, columns: [[2, values(mine.list, [[99, 99]])]] }), 'stale-filter-values'],
      ['committed with none', () => a.filter({ range: GRID, columns: [[1, { kind: 'committed' }]] }), 'policy-refused'],
    ]
    for (const [name, run, code] of cases) {
      const r = await run()
      expect({ name, end: r.end }).toMatchObject({ name, end: { type: 'refused', code } })
      const o = await a.order()
      expect({ name, viewGen: o.order?.head.viewGen }).toEqual({ name, viewGen: 0 })
    }
    expect(okRows(s.store.sheetOps(a.sid))).toEqual([])
    // P12: every refusal past the shape check has its row with the code.
    expect(s.store.sheetOps(a.sid).filter((r) => r.kind === 'sort' || r.kind === 'filter').map((r) => r.outcome)).toEqual(cases.map(([, , code]) => code))
    // P4: no engine ran a sort or filter for any of them.
    expect(seam.ops().filter((op) => op === 'sort' || op === 'filter')).toEqual([])

    // A list issued at viewGen 0 is stale at viewGen 1 (a resize moved the viewGen; nothing else moves).
    const b = await s.open()
    const listed = await b.list(R3, 2)
    const rz = b.send('sheet-resize', { viewGen: 0, axis: 'col', index: 0, size: 100 })
    expect((await b.c.next((y) => 'text' in y && y.text.reqId === rz) as Part).text).toMatchObject({ type: 'sheet-view', viewGen: 1 })
    expect((await b.filter({ viewGen: 1, range: R3, columns: [[2, values(listed.list, [[0, 0]])]] })).end).toMatchObject({ type: 'refused', code: 'stale-filter-values' })
    expect((await b.order()).order?.head.viewGen).toBe(1)
    expect(okRows(s.store.sheetOps(b.sid))).toEqual([])
    a.c.close(); other.c.close(); b.c.close(); await s.h.stop()

    // Unknown coverage under visible-only.
    const u = await serverOn(await unresolvedAnchor())
    const ua = await u.open()
    expect((await ua.sort({ range: wire(0, 1, 0, 1, 0), keys: [{ col: 0, order: 'asc' }] })).end).toMatchObject({ type: 'refused', code: 'coverage-unknown' })
    expect((await ua.order()).order?.head.viewGen).toBe(0)
    expect(okRows(u.store.sheetOps(ua.sid))).toEqual([])
    ua.c.close(); await u.h.stop()

    // P11: a third ordered sheet is view-state-limit, before any engine run.
    const tseam = seamPool()
    const t = await serverOn(await threeSheets(), { pool: tseam.proxy })
    const ta = await t.open()
    const col = wire(0, 5, 0, 0)
    for (const sheetId of ['0', '1']) expect((await ta.sort({ sheetId, range: col, keys: [{ col: 0, order: 'asc' }] })).order?.head.viewGen).toBe(1)
    const sorts = tseam.ops().filter((op) => op === 'sort').length
    expect((await ta.sort({ sheetId: '2', range: col, keys: [{ col: 0, order: 'asc' }] })).end).toMatchObject({ type: 'refused', code: 'view-state-limit' })
    expect(tseam.ops().filter((op) => op === 'sort').length).toBe(sorts)
    expect((await ta.order('2')).order?.head.viewGen).toBe(0)
    expect(okRows(t.store.sheetOps(ta.sid), '2')).toEqual([])
    expect(t.store.sheetOps(ta.sid).at(-1)).toMatchObject({ sheetId: '2', kind: 'sort', outcome: 'view-state-limit' })
    // Clearing an unsorted sheet's filter is the identity view: never ordered, never limited.
    expect((await ta.filter({ sheetId: '2', range: col, columns: [] })).order?.head.viewGen).toBe(0)
    ta.c.close(); await t.h.stop()
  }, 30_000)

  test('full allows typed conditions', async () => {
    const bytes = await sortGrid()
    const s = await serverOn(bytes, { policy: { contentPolicy: 'full' } })
    const a = await s.open()
    const r = await a.filter({ range: GRID, columns: [[1, { kind: 'number', op: 'gt', first: 50 }]] })
    expect(r.end).toBeUndefined()
    expect(r.order!.head).toMatchObject({ viewGen: 1, filterSummary: [{ col: 1, kind: 'number' }] })
    const or = await oracle(await oracleDoc(bytes))
    try {
      expect(await or.filter({ range: local(GRID), columns: new Map([[1, { kind: 'number', op: 'gt', first: 50 }]]) })).toBe('committed')
      const shown = shownSources(r.order!, 1, SORT_GRID.rows - 1)
      expect(shown.size).toBeGreaterThan(0)
      expect(shown).toEqual(or.shown(1, SORT_GRID.rows - 1))
      expect(r.order!.head.counts).toEqual(or.counts(GRID))
    } finally { await or.dispose() }
    const l = await a.layout()
    expect(l.filterSummary).toEqual([{ col: 1, kind: 'number' }])
    expect(l.sort).toBeNull()
    a.c.close(); await s.h.stop()
  }, 30_000)

  test('tautology filters', async () => {
    // S4 §6.6: `text neq ""` matches every row; under visible-only C1 still keeps the saved-autoFilter rows hidden.
    const bytes = await filterAutoHidden()
    const F = wire(0, 19, 0, 2)
    const hidden = FILTER_AUTO_HIDDEN.hidden.map((r) => r - 1)
    const s = await serverOn(bytes)
    const a = await s.open()
    const before = await a.layout()
    const r = await a.filter({ range: F, columns: [[0, text('neq', '')]] })
    expect(r.end).toBeUndefined()
    const shown = shownSources(r.order!, 1, 19)
    for (const row of hidden) expect({ row, shown: shown.has(row) }).toEqual({ row, shown: false })
    for (let row = 1; row < 20; row++) if (!hidden.includes(row)) expect({ row, shown: shown.has(row) }).toEqual({ row, shown: true })
    const after = await a.layout()
    expect(after.collapsedRows).toEqual(before.collapsedRows)
    expect(after.filterSummary).toEqual([{ col: 0, kind: 'text' }])
    expect((await a.filter({ viewGen: 1, range: F, columns: [[1, { kind: 'number', op: 'gte', first: -1e308 }]] })).end).toMatchObject({ type: 'refused', code: 'filter-condition-refused' })
    a.c.close(); await s.h.stop()

    // Under 'full' the filter decides alone, as locally.
    const f = await serverOn(bytes, { policy: { contentPolicy: 'full' } })
    const b = await f.open()
    const fr = await b.filter({ range: F, columns: [[0, text('neq', '')]] })
    expect(fr.end).toBeUndefined()
    const or = await oracle(await oracleDoc(bytes))
    try {
      expect(await or.filter({ range: local(F), columns: new Map([[0, { kind: 'text', op: 'neq', first: '', wildcards: false }]]) })).toBe('committed')
      expect(shownSources(fr.order!, 1, 19)).toEqual(or.shown(1, 19))
    } finally { await or.dispose() }
    b.c.close(); await f.h.stop()
  }, 30_000)

  test('committed rules', async () => {
    // Review Focus 3: a filter that keeps column B's committed values rule while editing column D applies both.
    const bytes = await sortGrid()
    const s = await serverOn(bytes)
    const a = await s.open()
    const b = await a.list(GRID, 1)
    const ids: Array<[number, number]> = [[0, 9]]
    const first = await a.filter({ range: GRID, columns: [[1, { kind: 'values', list: b.list, mode: 'include', ids, blanks: false }]] })
    expect(first.order?.head.viewGen).toBe(1)
    const both = await a.filter({ viewGen: 1, range: GRID, columns: [[1, { kind: 'committed' }], [3, text('contains', 'a')]] })
    expect(both.end).toBeUndefined()
    expect(both.order!.head).toMatchObject({ viewGen: 2, filterSummary: [{ col: 1, kind: 'values' }, { col: 3, kind: 'text' }] })
    const resolved: ResolvedFilter = { range: GRID, columns: [[1, { kind: 'values', list: { range: GRID, col: 1 }, positions: Array.from({ length: 10 }, (_, i) => i), mode: 'include', blanks: false }], [3, { kind: 'text', op: 'contains', first: 'a', wildcards: false }]] }
    const derive = await hostOf(bytes, 'visible-only')
    const want = orderOf(await derive(s.h.stream.sheetClockOf(a.sid)!, { sort: null, sortFilter: null, filter: resolved }))
    expect(shownSources(both.order!, 0, 39)).toEqual(shownSources(want, 0, 39))
    expect(both.order!.head.counts).toEqual(want.head.counts)
    // The row records the kind that ran, never `committed`.
    expect(s.store.sheetOps(a.sid).at(-1)).toMatchObject({ kind: 'filter', outcome: 'ok', rules: [[1, 'values'], [3, 'text']] })
    // A committed rule over another range has no committed rule there.
    expect((await a.filter({ viewGen: 2, range: wire(0, 20, 0, 5), columns: [[1, { kind: 'committed' }]] })).end).toMatchObject({ type: 'refused', code: 'policy-refused' })
    // Clearing the filter is the identity view again: the layout is 4b's.
    expect((await a.filter({ viewGen: 2, range: GRID, columns: [] })).order?.head).toMatchObject({ viewGen: 3, filterSummary: [], sort: null })
    const l = await a.layout()
    expect({ sort: l.sort, filterSummary: l.filterSummary, viewGen: l.viewGen }).toEqual({ sort: null, filterSummary: [], viewGen: 3 })
    a.c.close(); await s.h.stop()
  }, 30_000)
})

describe('the tile grid follows the view (P9)', () => {
  test('a full filter that shows authored-hidden rows gets its bottom tiles', async () => {
    const s = await serverOn(await hiddenTail(), { policy: { contentPolicy: 'full' } })
    const a = await s.open()
    const zoomStep = TILE_ZOOM_STEPS.length - 1, zoom = TILE_ZOOM_STEPS[zoomStep]!
    const identityRows = tileLimit(streamSheetGeometry(await a.layout()).metrics.rows.totalCss(), zoom, 1)
    const past = { zoomStep, dpr: 1, palette: 'light', tiles: [{ col: 0, row: identityRows }] }
    // Past the identity grid: no such tile.
    expect((await a.tiles(0, past)).end).toMatchObject({ type: 'refused', code: 'policy-refused' })
    const r = await a.filter({ range: wire(0, 60, 0, 0), columns: [[0, text('contains', '')]] })
    expect(r.order?.head.viewGen).toBe(1)
    const shown = shownSources(r.order!, 1, 60)
    for (const row of HIDDEN_TAIL.hidden) expect({ row, shown: shown.has(row - 1) }).toEqual({ row, shown: true })
    // The 20 shown rows grow the sheet: its bottom tile rows exist under the view.
    const got = await a.tiles(1, past)
    expect(got.end).toMatchObject({ type: 'tiles-done', refused: [] })
    expect(got.tiles.map((t) => [t.row, t.viewGen])).toEqual([[identityRows, 1]])
    // Clearing the filter is the identity grid again.
    expect((await a.filter({ viewGen: 1, range: wire(0, 60, 0, 0), columns: [] })).order?.head.viewGen).toBe(2)
    expect((await a.tiles(2, past)).end).toMatchObject({ type: 'refused', code: 'policy-refused' })
    a.c.close(); await s.h.stop()
  }, 30_000)
})

describe('engine refusals (P13)', () => {
  test('a local refusal is view-refused and moves nothing', async () => {
    // Review Focus 4: a sort over a merge or a cell-anchored drawing reaches the client with local's reason.
    for (const [bytes, refused] of [[await sortMerges(), { reason: 'merge', range: { r0: 4, r1: 5, c0: 1, c1: 1 } }], [await sortUnderShape(), { reason: 'drawing-impact' }]] as const) {
      const s = await serverOn(bytes)
      const a = await s.open()
      const { reqId: _r, ...end } = (await a.sort({ range: GRID, keys: B_ASC })).end!
      expect(end).toEqual({ type: 'refused', code: 'view-refused', ...refused })
      expect((await a.order()).order?.head.viewGen).toBe(0)
      expect((await a.layout()).sort).toBeNull()
      expect(s.store.sheetOps(a.sid)).toEqual([{ sid: a.sid, sheetId: '0', kind: 'sort', viewGen: 0, range: GRID, rules: [[1, 'asc']], outcome: 'view-refused', at: expect.any(Number) }])
      a.c.close(); await s.h.stop()
    }
  }, 30_000)
})

describe('tiles and deltas per sid (T6)', () => {
  test('T6 pair', async () => {
    const bytes = await sortGrid()
    const seam = seamPool()
    const s = await serverOn(bytes, { pool: seam.proxy })
    const x = await s.open(), y = await s.open(), z = await s.open()
    const DA = { range: GRID, columns: [[3, text('contains', 'a')]] }
    const x0 = await x.tiles(0)
    expect(seam.seen.tiles.length).toBe(1)
    const xf = await x.filter(DA)
    expect(xf.order?.head.viewGen).toBe(1)
    const x1 = await x.tiles(1)
    expect(seam.seen.tiles.length).toBe(2)
    expect(seam.seen.tiles[1]!.view).toBeDefined()
    // A second sid with the same view shares the canonical tile: no new pool.tile.
    expect((await z.filter(DA)).order?.head.viewGen).toBe(1)
    const z1 = await z.tiles(1)
    expect(seam.seen.tiles.length).toBe(2)
    // The unfiltered sid shares the identity tile (drawn with no view), never the filtered one.
    const y0 = await y.tiles(0)
    expect(seam.seen.tiles.length).toBe(2)
    expect(seam.seen.tiles[0]!.view).toBeUndefined()
    expect(seam.seen.tiles.map((t) => [t.col, t.row, t.zoomStep])).toEqual([[0, 0, 5], [0, 0, 5]])
    // The two canonical tiles differ: the filter hides rows of tile (0, 0).
    expect(x1.tiles[0]!.png).not.toEqual(x0.tiles[0]!.png)
    // Y's view never moved and it never asked for an order: no sheet-order part reaches it.
    expect((await y.layout()).viewGen).toBe(0)
    expect((await y.tiles(1)).end).toMatchObject({ type: 'refused', code: 'stale-view' })
    const y2 = await y.tiles(0)
    expect(seam.seen.tiles.length).toBe(2)
    // Each socket's frames carry only its own viewGen sequence (answers above, plus anything left unread).
    const gens = (c: Client) => c.msgs.flatMap((m) => 'binary' in m && m.binary.kind === 5 ? [decodeTileBody(m.binary.payload).viewGen] : 'text' in m && m.text.type === 'sheet-order' ? [m.text.head?.viewGen ?? -1] : [])
    expect([...x0.tiles, ...x1.tiles].map((t) => t.viewGen)).toEqual([0, 1])
    expect(xf.order!.head.viewGen).toBe(1)
    expect([...y0.tiles, ...y2.tiles].map((t) => t.viewGen)).toEqual([0, 0])
    expect(z1.tiles.map((t) => t.viewGen)).toEqual([1])
    for (const c of [x.c, y.c, z.c]) expect(gens(c)).toEqual([])
    x.c.close(); y.c.close(); z.c.close(); await s.h.stop()
  }, 30_000)

  test('text ops follow the view', async () => {
    // Open item 2: find skips rows hidden by the session's filter (Task 6 at the host, here over the socket).
    const s = await serverOn(await sortGrid())
    const a = await s.open()
    const find = async (query: string) => {
      const reqId = a.send('sheet-find', { query, scope: '0' })
      const hits: unknown[] = []
      for (;;) {
        const m = await a.c.next((y) => 'text' in y && y.text.reqId === reqId) as Part
        if (m.text.type === 'sheet-find-result') { hits.push(...m.text.hits); continue }
        return { hits, end: m.text }
      }
    }
    expect((await find('plum-3')).hits).toEqual([[0, 2, 3], [0, 34, 3]])
    expect((await a.filter({ range: GRID, columns: [[3, text('contains', 'a')]] })).order?.head.viewGen).toBe(1)
    expect((await find('plum-3')).hits).toEqual([])
    expect((await find('apple-8')).hits).toEqual([[0, 7, 3]])
    a.c.close(); await s.h.stop()
  }, 30_000)
})

describe('rates, cancel, admission and audit (P8, P11, P12)', () => {
  test('the sheet class', async () => {
    const seam = seamPool()
    const s = await serverOn(await sortGrid(), { pool: seam.proxy })
    const a = await s.open()
    let viewGen = 0
    for (let i = 1; i <= 30; i++) {
      const r = await a.sort({ viewGen, range: GRID, keys: i % 2 ? B_ASC : B_DESC })
      expect({ i, end: r.end }).toEqual({ i, end: undefined })
      viewGen = r.order!.head.viewGen
    }
    expect(viewGen).toBe(30)
    expect((await a.sort({ viewGen, range: GRID, keys: B_ASC })).end).toMatchObject({ type: 'refused', code: 'rate-limited' })
    expect(s.store.sheetOps(a.sid).at(-1)).toMatchObject({ kind: 'sort', outcome: 'rate-limited', rules: [[1, 'asc']] })

    // One in flight (its own session: a busy refusal is charged too).
    const b = await s.open()
    const held = seam.holdNext()
    const first = b.send('sheet-sort', { viewGen: 0, range: GRID, keys: B_ASC })
    await held.reached
    expect((await b.sort({ range: GRID, keys: B_DESC })).end).toMatchObject({ type: 'refused', code: 'rate-limited' })
    expect((await b.filter({ range: GRID, columns: [] })).end).toMatchObject({ type: 'refused', code: 'rate-limited' })
    held.release()
    expect((await b.answer(first)).order?.head.viewGen).toBe(1)
    expect((await b.sort({ viewGen: 1, range: GRID, keys: B_DESC })).order?.head.viewGen).toBe(2)
    for (const c of [a.c, b.c]) expect(c.ws.readyState).toBe(WebSocket.OPEN)
    a.c.close(); b.c.close(); await s.h.stop()
  }, 30_000)

  test('cancel', async () => {
    const seam = seamPool()
    const s = await serverOn(await sortGrid(), { pool: seam.proxy })
    const a = await s.open()
    const held = seam.holdNext()
    const target = a.send('sheet-sort', { viewGen: 0, range: GRID, keys: B_ASC })
    await held.reached
    a.send('sheet-cancel', { target })
    // The cancel is handled before this order request is answered.
    expect((await a.order()).order?.head.viewGen).toBe(0)
    held.release()
    expect((await a.answer(target)).end).toMatchObject({ type: 'refused', code: 'cancelled' })
    const after = await a.order()
    expect(after.order?.head.viewGen).toBe(0)
    expect(s.store.sheetOps(a.sid).at(-1)).toMatchObject({ kind: 'sort', outcome: 'cancelled' })
    // A cancel after the answer changes nothing.
    const done = a.send('sheet-sort', { viewGen: 0, range: GRID, keys: B_ASC })
    const answered = await a.answer(done)
    expect(answered.order?.head.viewGen).toBe(1)
    a.send('sheet-cancel', { target: done })
    const kept = await a.order()
    expect(stripped(kept.parts)).toEqual(stripped(answered.parts))
    expect((await a.layout()).viewGen).toBe(1)
    a.c.close(); await s.h.stop()
  }, 30_000)

  test('resume during an in-flight sort', async () => {
    // Review Focus 1: the old socket's sort never lands on the resumed session, its busy flag is gone, and a re-issued
    // sort is served.
    const seam = seamPool()
    const s = await serverOn(await sortGrid(), { pool: seam.proxy })
    const a = await s.open()
    const held = seam.holdNext()
    const lost = a.send('sheet-sort', { viewGen: 0, range: GRID, keys: B_ASC })
    await held.reached
    expect(s.h.stream.sheetRatesOf(a.sid)!.sheetBusy).toBe(true)
    const b = await s.open({ sid: a.sid, token: a.token })
    expect(b.sid).toBe(a.sid)
    expect((await a.c.closed).code).toBe(4409)
    expect(s.h.stream.sheetRatesOf(a.sid)!.sheetBusy).toBe(false)
    held.release()
    // The order job queues behind the old one on the sheet's port: by its answer the old job has ended, unapplied.
    expect((await b.order()).order?.head.viewGen).toBe(0)
    expect((await b.layout()).sort).toBeNull()
    expect(b.c.msgs.some((m) => 'text' in m && m.text.reqId === lost)).toBe(false)
    expect(s.h.stream.sheetRatesOf(a.sid)!.sheetBusy).toBe(false)
    const again = await b.sort({ range: GRID, keys: B_ASC })
    expect(again.order?.head.viewGen).toBe(1)
    expect(s.store.sheetOps(a.sid).filter((r) => r.outcome === 'ok')).toHaveLength(1)
    b.c.close(); await s.h.stop()
  }, 30_000)

  test('a million-row sort', async () => {
    const s = await serverOn(await oneMillionSort(), { policy: { contentPolicy: 'full' } })
    const a = await s.open(), b = await s.open()
    const held0 = s.h.stream.documentLedgerBytes()
    const range = wire(0, ONE_MILLION_ROWS - 1, 0, 0, 0)
    const view: CanonicalView = { sort: { range, keys: [{ col: 0, order: 'asc' }], caseSensitive: false }, sortFilter: null, filter: null }
    const r = await a.sort({ range, keys: [{ col: 0, order: 'asc' }] })
    expect(r.end).toBeUndefined()
    expect(r.parts.length).toBeLessThanOrEqual(ORDER_LIMITS.parts)
    for (const p of r.parts) expect(p.bytes).toBeLessThanOrEqual(LAYOUT_PART_BYTES)
    expect(r.order!.bands[0]!.sourceAtOffset[0]).toBe(ONE_MILLION_ROWS - 1)
    expect(s.h.stream.documentLedgerBytes() - held0).toBe(orderBytes(view))
    const rb = await b.sort({ range, keys: [{ col: 0, order: 'asc' }] })
    expect(rb.end).toBeUndefined()
    expect(rb.order!.head.viewGen).toBe(1)
    expect(s.h.stream.documentLedgerBytes() - held0).toBe(orderBytes(view))
    a.c.close()
    while (s.h.stream.table.get(a.sid)) await Bun.sleep(5)
    expect(s.h.stream.documentLedgerBytes() - held0).toBe(orderBytes(view))
    b.c.close()
    while (s.h.stream.table.get(b.sid)) await Bun.sleep(5)
    expect(s.h.stream.documentLedgerBytes()).toBe(held0)
    await s.h.stop()
  }, 30_000)

  test('audit', async () => {
    const s = await serverOn(await sortGrid())
    const a = await s.open()
    expect((await a.sort({ range: GRID, keys: B_ASC })).order?.head.viewGen).toBe(1)
    expect((await a.filter({ viewGen: 1, range: GRID, columns: [[1, { kind: 'number', op: 'gt', first: 57.25 }]] })).end).toMatchObject({ type: 'refused', code: 'filter-condition-refused' })
    expect((await a.filter({ viewGen: 1, range: GRID, columns: [[3, text('contains', 'OPERAND-SENTINEL')]] })).order?.head.viewGen).toBe(2)
    const rows = s.store.sheetOps(a.sid)
    expect(rows).toEqual([
      { sid: a.sid, sheetId: '0', kind: 'sort', viewGen: 0, range: GRID, rules: [[1, 'asc']], outcome: 'ok', at: expect.any(Number) },
      { sid: a.sid, sheetId: '0', kind: 'filter', viewGen: 1, range: GRID, rules: [[1, 'number']], outcome: 'filter-condition-refused', at: expect.any(Number) },
      { sid: a.sid, sheetId: '0', kind: 'filter', viewGen: 1, range: GRID, rules: [[3, 'text']], outcome: 'ok', at: expect.any(Number) },
    ])
    for (const operand of ['57.25', 'OPERAND-SENTINEL']) expect(JSON.stringify(rows)).not.toContain(operand)
    expect(s.store.sheetView(a.sid, '0', 1)?.spec).toEqual({ sort: { range: GRID, keys: B_ASC, caseSensitive: false }, sortFilter: null, filter: null, sizes: [] })
    // P12: typed operands are stored in the clear in sheet_views (never in sheet_ops).
    expect(s.store.sheetView(a.sid, '0', 2)?.spec).toMatchObject({ sort: { keys: B_ASC }, sortFilter: null, filter: { range: GRID, columns: [[3, { kind: 'text', op: 'contains', first: 'OPERAND-SENTINEL', wildcards: false }]] } })
    a.c.close(); await s.h.stop()
  }, 30_000)
})
