/**
 * Stage 4c Task 7 — `filter-values` and `range-detect` on the socket (S4 §6.6, §6.12; plan P4, P6, P12, P13). Value
 * lists are display text over the key view (a withheld cell lists only as "(Blanks)"), list ids are bound to their
 * sid, `(sheetId, range, col)` and `viewGen` (`ListLedger`), pages are ≤ 500 entries and ≤ 48 KiB, and the first page
 * of a request charges `'sheet'` while its cursor pages do not.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { viewerControls } from '../../src/components/xlsx/viewerControlsClient'
import type { CacheIdentity, FilterRange, SheetView } from '../../src/xlsx/view/contracts'
import { MainMemoryLedger } from '../../src/xlsx/view/mainLedger'
import { createIdentityOrder } from '../../src/xlsx/view/order'
import { emptyGridSelection } from '../../src/xlsx/view/selection'
import { createObjectWorkbookSource } from '../../src/xlsx/view/source'
import { visibilityFromHidden } from '../../src/xlsx/view/visibility'
import type { XlsxCell, XlsxDocument } from '../../src/xlsx/types'
import { buildXlsx, type XlsxCellSpec } from '../../src/testdata/ooxml-builders'
import { decodeDetection, decodeValuesPage, type ValuesPage, type WireRange } from '../../src/stream/sheet-wire'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import type { Pool } from '../src/pool'
import type { RateClass, SessionRates } from '../src/rate'
import { DocumentHost } from '../src/render'
import { ListLedger } from '../src/sheet/value-lists'
import { TEXT_MESSAGE_BYTES } from '../src/text/ops'
import { Client, disposePool, pool, start, type Harness } from './harness'
import { offViewportCovered, PAINTED_FORMATS, paintedFormats, rotatedBox, SORT_GRID, SORT_WITHHELD, sortWithheld, unresolvedAnchor } from './sheet-fixtures'

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
const rcOf = (ref: string): [number, number] => [Number(ref.slice(1)) - 1, ref.charCodeAt(0) - 65]

/** A server on an injected clock (each request steps it 100 ms, so `'message'` never refuses), with its own audit. */
async function serverOn(doc: Uint8Array, o: Parameters<typeof start>[0] = {}) {
  const clock = { t: T0 }
  const store = AuditStore.open(':memory:')
  const h = await start({ doc, now: () => clock.t, tokenTtlSec: 86_400, audit: new AuditTable({ store }), ...o })
  let rq = 1000
  async function open() {
    const c = new Client(h.url)
    const m = await c.hello({ features: VIEW })
    await c.ready()
    while (h.stream.table.get(m.session.sid)?.status !== 'stream') await Bun.sleep(5)
    const sid = m.session.sid as string
    /** Sends one op and returns its answer with the bytes it took on the wire. */
    const ask = async (type: string, body: Record<string, unknown>) => {
      clock.t += 100
      const reqId = ++rq
      c.send({ token: m.session.token, reqId, type, sheetId: '0', viewGen: 0, ...body })
      const a = await c.next((x) => 'text' in x && x.text.reqId === reqId) as { text: Record<string, any>; bytes: number }
      return { ...a.text, bytes: a.bytes } as Record<string, any> & { bytes: number }
    }
    const values = (body: Record<string, unknown>) => ask('filter-values', body)
    /** Every page of one list: the first page, then its cursors. */
    async function list(range: WireRange, col: number, search?: { text: string; wildcards: boolean }) {
      const pages: Array<ValuesPage & { bytes: number }> = []
      let a = await values({ range, col, ...(search ? { search } : {}) })
      for (;;) {
        if (a.type !== 'filter-values-result') throw new Error(`values refused: ${JSON.stringify(a)}`)
        const { bytes, ...msg } = a
        const p = decodeValuesPage(msg)
        pages.push({ ...p, bytes })
        if (!p.cursor) break
        a = await values({ cursor: p.cursor })
      }
      return { pages, items: pages.flatMap((p) => p.items), first: pages[0]! }
    }
    return { c, sid, ask, values, list }
  }
  return { h, store, clock, open }
}

/** Wraps the sid's `SessionRates.check` so each call is recorded by class. */
function chargesOf(h: Harness, sid: string): RateClass[] {
  const rates = h.stream.sheetRatesOf(sid) as SessionRates
  const seen: RateClass[] = []
  const check = rates.check.bind(rates)
  rates.check = (kind, now) => { seen.push(kind); return check(kind, now) }
  return seen
}

/** A one-sheet book: A1 = `H`, then `texts` down column A from A2. */
function textColumn(texts: readonly string[]): Promise<Uint8Array> {
  const strings = ['H', ...texts]
  const rows = strings.map((_, i) => ({ r: i + 1, cells: [{ ref: `A${i + 1}`, t: 's', v: i } as XlsxCellSpec] }))
  return buildXlsx([{ name: 'S', rows }], strings)
}

// ---- the oracle: local value labels over the fixture with its withheld cells deleted ----------------------------

async function oracleLabels(bytes: Uint8Array, range: WireRange, col: number, drop: (row: number, col: number) => boolean, replace: Record<string, string>): Promise<string[]> {
  const parsed = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
  const sheet = parsed.sheets[0]!
  const swap = new Map(Object.entries(replace).map(([ref, v]) => [rcOf(ref).join(':'), v]))
  const rows = sheet.rows.map((row) => ({ ...row, cells: row.cells.filter((c) => !drop(c.row, c.col)).map((c): XlsxCell => {
    const v = swap.get(`${c.row}:${c.col}`)
    return v === undefined ? c : { ref: c.ref, row: c.row, col: c.col, value: v, styleIndex: 0, style: { numFmtId: 0 } }
  }) }))
  const doc: XlsxDocument = { ...parsed, sheets: [{ ...sheet, rows }, ...parsed.sheets.slice(1)] }
  const source = createObjectWorkbookSource(doc, 'oracle')
  const sheetId = source.sheets[0]!.sheetId
  const baseHidden = source.geometry(sheetId).baseHidden
  const view: SheetView = { sheetId, revision: 0, order: createIdentityOrder(), baseHidden, visibility: visibilityFromHidden(baseHidden), sort: null, filter: null, selection: emptyGridSelection() }
  const owner = (): CacheIdentity => ({ docId: 'oracle', sheetId, valueGen: source.store(sheetId).stamp.valueGen, fontGen: 0, viewGen: 0 })
  const snapshot = () => ({ owner: owner(), filter: null, sort: null, records: { visible: 0, total: 0, completeness: 'complete' as const } })
  const signal = new AbortController().signal
  const controls = await viewerControls({ doc, source, sheetId, snapshot, view: () => view, publish: () => snapshot(), ledger: new MainMemoryLedger() }, signal)
  try {
    const port = await controls.client.values(owner(), local(range), col, null, signal)
    try {
      const labels = (await port.read({ text: '', offset: 0 }, signal)).entries.map((e) => e.label)
      return [...new Set(labels)]
    } finally { port.cancel() }
  } finally { controls.dispose(); await source.release() }
}
/** `sortWithheld` under `visible-only`, as Task 5's oracle deletes it. */
function withheldLabels(bytes: Uint8Array, range: WireRange, col: number) {
  const gone = new Set([SORT_WITHHELD.covered, SORT_WITHHELD.hiddenFormat, SORT_WITHHELD.cfHidden, 'C30', 'C31', 'C32', 'C33', 'C34'].map((r) => rcOf(r).join(':')))
  return oracleLabels(bytes, range, col, (r, c) => gone.has(`${r}:${c}`) || c === SORT_GRID.zeroWidthCol || r === SORT_GRID.zeroHeightRow || r === SORT_GRID.hiddenRow,
    { [SORT_WITHHELD.masked]: SORT_WITHHELD.maskedDisplay })
}

describe('filter-values (S4 §6.6; plan P6)', () => {
  test('lists are display text, withheld cells only as blanks', async () => {
    const bytes = await sortWithheld()
    const s = await serverOn(bytes)
    const a = await s.open()
    const b = await a.list(GRID, 1)
    const texts = b.items.map((i) => i.text)
    expect(texts).toEqual(await withheldLabels(bytes, GRID, 1))
    expect(texts.length).toBeGreaterThan(30)
    expect(b.first.blanks.present).toBe(true)
    for (const v of ['1', '5', '95']) expect(texts).not.toContain(v)
    const c = (await a.list(GRID, 2)).items.map((i) => i.text)
    expect(c.length).toBeGreaterThan(20)
    for (const r of [30, 31, 32, 33, 34]) for (const shown of [String(r * 1.25), (r * 1.25).toFixed(1)]) expect(c).not.toContain(shown)
    const d = (await a.list(GRID, 3)).items.map((i) => i.text)
    expect(d).toContain(SORT_WITHHELD.maskedDisplay)
    expect(d).not.toContain(SORT_WITHHELD.maskedText)
    // The first page names its list's classification; every page is the decoded wire shape.
    expect(b.first.classification).toBeDefined()
    a.c.close(); await s.h.stop()

    const p = await serverOn(await paintedFormats())
    const q = await p.open()
    expect((await q.list(wire(0, 3, 1, 1), 1)).items.map((i) => i.text)).toEqual([PAINTED_FORMATS.tablePaints])
    for (let col = 0; col <= 3; col++) {
      const all = (await q.list(wire(0, 5, 0, 3), col)).items.map((i) => i.text)
      for (const raw of [String(PAINTED_FORMATS.tableValue), String(PAINTED_FORMATS.cfValue)]) expect(all).not.toContain(raw)
    }
    q.c.close(); await p.h.stop()
  }, 30_000)

  test('rotated box and unknown coverage', async () => {
    // S4 §9 4c: the rotated box paints x 40..60, y −40..60: A1, A2 and A3 (x 0..64, y 0..60) meet x 40..60, y 30..50's
    // column, and A2 (y 20..40) and A3 (y 40..60) meet it outright.
    const r = await serverOn(await rotatedBox())
    const a = await r.open()
    const A = await a.list(wire(0, 5, 0, 2, 0), 0)
    const at = A.items.map((i) => i.text)
    for (const covered of ['ROT-2-1', 'ROT-3-1']) expect(at).not.toContain(covered)
    expect(at).toContain('ROT-4-1')
    expect(A.first.blanks.present).toBe(true)
    expect((await a.list(wire(0, 5, 0, 2, 0), 1)).items.map((i) => i.text)).toContain('ROT-2-2')
    a.c.close(); await r.h.stop()

    const u = await serverOn(await unresolvedAnchor())
    const b = await u.open()
    for (const col of [0, 1]) {
      const l = await b.list(wire(0, 1, 0, 1, 0), col)
      expect({ col, items: l.items, blanks: l.first.blanks.present }).toEqual({ col, items: [], blanks: true })
    }
    b.c.close(); await u.h.stop()
  }, 30_000)

  test('off-viewport covered cell', async () => {
    const n = { tiles: 0 }
    const base = pool()
    const counting = new Proxy(base, {
      get(t, k) {
        if (k === 'tile') return (...a: Parameters<Pool['tile']>) => { n.tiles++; return t.tile(...a) }
        const v = Reflect.get(t, k, t)
        return typeof v === 'function' ? v.bind(t) : v
      },
    })
    const s = await serverOn(await offViewportCovered(), { pool: counting })
    const a = await s.open()
    const z = await a.list(wire(0, 4999, 25, 25, 0), 25)
    expect(z.items.map((i) => i.text)).not.toContain('OFFVIEW-Z5000')
    expect(z.first.blanks.present).toBe(true)
    expect((await a.list(wire(0, 4999, 0, 0, 0), 0)).items.map((i) => i.text)).toContain('OFFVIEW-A1')
    expect(n.tiles).toBe(0)
    a.c.close(); await s.h.stop()
  }, 30_000)

  test('paging and search', async () => {
    const texts = Array.from({ length: 1200 }, (_, i) => `V-${i}`)
    const s = await serverOn(await textColumn(texts))
    const a = await s.open()
    const range = wire(0, 1200, 0, 0)
    const l = await a.list(range, 0)
    expect(l.pages.map((p) => p.items.length)).toEqual([500, 500, 200])
    for (const p of l.pages) {
      expect(p.bytes).toBeLessThanOrEqual(TEXT_MESSAGE_BYTES)
      expect(p.list).toBe(l.first.list)
      expect(p.total).toBe(1200)
      expect(p.truncated).toBe(false)
    }
    expect(l.pages.slice(1).every((p) => p.classification === undefined)).toBe(true)
    expect(l.items.map((i) => i.id)).toEqual(texts.map((_, i) => i))
    expect(l.items.map((i) => i.text)).toEqual(texts)
    // Stable: the same key gets the same list and ids.
    const again = await a.list(range, 0)
    expect(again.first.list).toBe(l.first.list)
    expect(again.items).toEqual(l.items)
    // A search answers matching entries with their full-list ids, over as many pages as it needs (309 match "7").
    const sevens = await a.list(range, 0, { text: '7', wildcards: false })
    const want = l.items.filter((i) => i.text.includes('7'))
    expect(want.length).toBe(309)
    expect(sevens.first.list).toBe(l.first.list)
    expect(sevens.items.map((i) => [i.id, i.text])).toEqual(want.map((i) => [i.id, i.text]))
    // A cursor of a superseded search is stale; the plain list's cursors still page.
    const s1 = decodeValuesPage((({ bytes: _, ...m }) => m)(await a.values({ range, col: 0, search: { text: 'V-', wildcards: false } })))
    expect(s1.cursor).toBeDefined()
    await a.values({ range, col: 0, search: { text: 'V-1', wildcards: false } })
    expect(await a.values({ cursor: s1.cursor })).toMatchObject({ type: 'refused', code: 'stale-filter-values' })
    expect(await a.values({ cursor: l.pages[0]!.cursor })).toMatchObject({ type: 'filter-values-result', list: l.first.list })
    for (const bad of [`${l.first.list}:0`, `${l.first.list}:1200`, `${l.first.list}:-5`, `${l.first.list}:05`, `${l.first.list}`, 'nope:500']) {
      expect({ bad, a: await a.values({ cursor: bad }) }).toMatchObject({ bad, a: { type: 'refused', code: 'stale-filter-values' } })
    }
    a.c.close(); await s.h.stop()

    // Long labels: a page over 48 KiB is split by entries, and its cursor continues from there.
    const long = Array.from({ length: 600 }, (_, i) => `L${String(i).padStart(4, '0')}-${'x'.repeat(195)}`)
    const ls = await serverOn(await textColumn(long))
    const la = await ls.open()
    const ll = await la.list(wire(0, 600, 0, 0), 0)
    expect(ll.first.items.length).toBeLessThan(500)
    expect(ll.pages.length).toBeGreaterThan(2)
    for (const p of ll.pages) expect(p.bytes).toBeLessThanOrEqual(TEXT_MESSAGE_BYTES)
    expect(ll.items.map((i) => [i.id, i.text])).toEqual(long.map((t, i) => [i, t]))
    la.c.close(); await ls.h.stop()

    const many = Array.from({ length: 12_000 }, (_, i) => `W-${i}`)
    const m = await serverOn(await textColumn(many))
    const b = await m.open()
    const big = await b.list(wire(0, 12_000, 0, 0), 0)
    expect(big.first.truncated).toBe(true)
    expect(big.items.length).toBe(10_000)
    expect(big.pages.length).toBe(20)
    expect(new Set(big.items.map((i) => i.id)).size).toBe(10_000)
    b.c.close(); await m.h.stop()
  }, 30_000)

  test('rates', async () => {
    const s = await serverOn(await textColumn(Array.from({ length: 1200 }, (_, i) => `V-${i}`)))
    const a = await s.open()
    const charges = chargesOf(s.h, a.sid)
    const range = wire(0, 1200, 0, 0)
    const first = await a.values({ range, col: 0 })
    expect(first.type).toBe('filter-values-result')
    expect(charges.splice(0)).toEqual(['message', 'sheet'])
    for (let i = 0; i < 30; i++) {
      expect(await a.values({ cursor: first.cursor })).toMatchObject({ type: 'filter-values-result', list: first.list })
    }
    expect(charges.splice(0)).toEqual(Array(30).fill('message'))
    for (let i = 2; i <= 30; i++) expect({ i, type: (await a.values({ range, col: 0 })).type }).toEqual({ i, type: 'filter-values-result' })
    expect(await a.values({ range, col: 0 })).toMatchObject({ type: 'refused', code: 'rate-limited' })
    // Cursor pages still answer: they charge only 'message'.
    expect(await a.values({ cursor: first.cursor })).toMatchObject({ type: 'filter-values-result' })
    // P12: one sheet_ops row per first page that passed its shape check (the refusal included), none per cursor page;
    // no search text in any row.
    const rows = s.store.sheetOps(a.sid)
    expect(rows.length).toBe(31)
    expect(rows.slice(0, 30).every((r) => r.kind === 'filter-values' && r.outcome === 'ok' && JSON.stringify(r.rules) === '[[0,"values"]]')).toBe(true)
    expect(rows[30]!.outcome).toBe('rate-limited')
    a.c.close(); await s.h.stop()
  }, 30_000)

  test('values on without copy', async () => {
    const s = await serverOn(await sortWithheld(), { policy: { copy: false } })
    const a = await s.open()
    expect((await a.list(GRID, 3)).items.map((i) => i.text)).toContain(SORT_WITHHELD.maskedDisplay)
    a.c.close(); await s.h.stop()
  }, 30_000)

  test('refusals: shape, stale-view, hidden-target', async () => {
    const s = await serverOn(await sortWithheld())
    const a = await s.open()
    const charges = chargesOf(s.h, a.sid)
    // Shape: uncharged beyond 'message', no row.
    expect(await a.values({ range: GRID, col: 9 })).toMatchObject({ type: 'refused', code: 'policy-refused' })
    expect(await a.values({ range: GRID, col: 1, cursor: 'x:1' })).toMatchObject({ type: 'refused', code: 'policy-refused' })
    expect(charges.splice(0)).toEqual(['message', 'message'])
    expect(s.store.sheetOps(a.sid)).toEqual([])
    expect(await a.values({ range: GRID, col: 1, viewGen: 3 })).toMatchObject({ type: 'refused', code: 'stale-view' })
    // Column E is authored width 0: a collapsed column is never listed.
    expect(await a.values({ range: GRID, col: SORT_GRID.zeroWidthCol })).toMatchObject({ type: 'refused', code: 'hidden-target' })
    expect(await a.values({ range: wire(0, 39, 4, 5), col: 5 })).toMatchObject({ type: 'refused', code: 'hidden-target' })
    expect(s.store.sheetOps(a.sid).map((r) => r.outcome)).toEqual(['stale-view', 'hidden-target', 'hidden-target'])
    a.c.close(); await s.h.stop()
  }, 30_000)

  test('ids are sid- and column-bound', async () => {
    const C = { sheetIndex: 0, range: GRID, col: 2, viewGen: 0 }
    const rule = (list: string, ids: Array<[number, number]>) => ({ kind: 'values' as const, list, mode: 'include' as const, ids, blanks: false })
    const mine = new ListLedger(), theirs = new ListLedger()
    const id = mine.issue(C, 3)
    expect(mine.issue({ ...C, range: { ...GRID } }, 3)).toBe(id)
    expect(mine.resolve(rule(id, [[0, 2]]), C)).toEqual({ kind: 'values', list: { range: GRID, col: 2 }, positions: [0, 1, 2], mode: 'include', blanks: false })
    expect(theirs.resolve(rule(id, [[0, 0]]), C)).toBe('stale-filter-values')
    expect(mine.resolve(rule(id, [[0, 0]]), { ...C, col: 3 })).toBe('stale-filter-values')
    expect(mine.resolve(rule(id, [[0, 0]]), { ...C, range: { ...GRID, header: 0 } })).toBe('stale-filter-values')
    expect(mine.resolve(rule(id, [[0, 0]]), { ...C, sheetIndex: 1 })).toBe('stale-filter-values')
    expect(mine.resolve(rule(id, [[99, 99]]), C)).toBe('stale-filter-values')
    expect(mine.resolve(rule(id, [[0, 0xffffffff]]), C)).toBe('stale-filter-values')
    expect(mine.resolve(rule(id, [[2, 1]]), C)).toBe('stale-filter-values')
    expect(mine.resolve(rule(id, [[1, 2], [0, 0]]), C)).toBe('stale-filter-values')
    expect(mine.resolve(rule(id, [[0, 0]]), { ...C, viewGen: 1 })).toBe('stale-filter-values')
    for (let col = 3; col < 3 + 15; col++) mine.issue({ ...C, col }, 3)
    expect(mine.get(id)).toBeDefined()
    mine.issue({ ...C, col: 40 }, 3)
    expect(mine.get(id)).toBeUndefined()
    expect(mine.resolve(rule(id, [[0, 0]]), C)).toBe('stale-filter-values')

    // On the socket: another sid's cursor is stale-filter-values.
    const s = await serverOn(await textColumn(Array.from({ length: 600 }, (_, i) => `V-${i}`)))
    const a = await s.open(), b = await s.open()
    const page = await a.values({ range: wire(0, 600, 0, 0), col: 0 })
    expect(page.type).toBe('filter-values-result')
    expect(await b.values({ cursor: page.cursor })).toMatchObject({ type: 'refused', code: 'stale-filter-values' })
    expect(await a.values({ cursor: page.cursor })).toMatchObject({ type: 'filter-values-result' })
    a.c.close(); b.c.close(); await s.h.stop()
  }, 30_000)
})

describe('range-detect (S4 §6.6; plan P13)', () => {
  test('range-detect', async () => {
    const bytes = await sortWithheld()
    const host = new DocumentHost(); hosts.push(host)
    const { docKey } = await host.open('doc', bytes, 'visible-only', { sheetMode: 'tiles' })
    const s = await serverOn(bytes)
    const a = await s.open()
    const selection = { active: rcOf('C31'), anchor: null, ranges: [] }
    const clock = s.h.stream.sheetClockOf(a.sid)!
    const want = await host.sheetView(docKey, clock, 'visible-only', { op: 'detect', sheetIndex: 0, kind: 'filter', selection })
    if (want.op !== 'detect' || !('detection' in want)) throw new Error(`detect: ${JSON.stringify(want)}`)
    const charges = chargesOf(s.h, a.sid)
    const got = await a.ask('range-detect', { kind: 'filter', selection })
    expect(got.type).toBe('range-detect-result')
    expect(decodeDetection((({ bytes: _, ...m }) => m)(got))).toEqual(want.detection)
    expect(charges.splice(0)).toEqual(['message', 'sheet'])
    expect(s.store.sheetOps(a.sid)).toEqual([{ sid: a.sid, sheetId: '0', kind: 'range-detect', viewGen: 0, range: want.detection.range, rules: [], outcome: 'ok', at: expect.any(Number) }])
    // Shape: uncharged.
    expect(await a.ask('range-detect', { kind: 'pivot', selection })).toMatchObject({ type: 'refused', code: 'policy-refused' })
    expect(charges.splice(0)).toEqual(['message'])
    // Task 1: the header sort's region excludes the active column (C30:C34 is withheld), so the local refusal travels.
    const sortWant = await host.sheetView(docKey, clock, 'visible-only', { op: 'detect', sheetIndex: 0, kind: 'sort', selection })
    if (sortWant.op !== 'detect' || !('refused' in sortWant)) throw new Error(`sort detect: ${JSON.stringify(sortWant)}`)
    expect(sortWant.refused.reason).toBe('outside-column-data')
    const { bytes: _b, reqId: _r, ...sorted } = await a.ask('range-detect', { kind: 'sort', selection })
    expect(sorted).toEqual({ type: 'refused', code: 'view-refused', ...sortWant.refused })
    expect(s.store.sheetOps(a.sid).at(-1)).toMatchObject({ kind: 'range-detect', outcome: 'view-refused', range: { ...sortWant.refused.range!, header: 0, totals: 0 } })
    // 28 more 'sheet' ops (a filter-values first page among them) are answered; the 31st is rate-limited.
    expect((await a.values({ range: GRID, col: 1 })).type).toBe('filter-values-result')
    for (let i = 4; i <= 30; i++) expect({ i, type: (await a.ask('range-detect', { kind: 'filter', selection })).type }).toEqual({ i, type: 'range-detect-result' })
    expect(await a.ask('range-detect', { kind: 'filter', selection })).toMatchObject({ type: 'refused', code: 'rate-limited' })
    expect(s.store.sheetOps(a.sid).at(-1)).toMatchObject({ kind: 'range-detect', outcome: 'rate-limited' })
    a.c.close(); await s.h.stop()
  }, 30_000)
})
