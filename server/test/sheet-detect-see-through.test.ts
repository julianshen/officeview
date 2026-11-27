/**
 * Stage 4c Task 11a — see-through range detection (owner 2026-11-27). Under `'visible-only'` the key view withholds
 * author-hidden lines (the geometry's collapsed rows and columns: `hidden="1"`, `ht="0"`, a saved autoFilter's rows,
 * `width="0"` or hidden columns). Detection steps over them: they neither end nor extend a region, and their contents
 * are never read. A rectangle selection whose only gaps are author-hidden lines is one rectangle. Cells withheld for
 * content reasons still read as blank and may end a region.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { buildXlsx, type XlsxCellSpec, type XlsxSheetSpec } from '../../src/testdata/ooxml-builders'
import { decodeDetection, decodeSheetOrder, decodeValuesPage, type SheetOrder, type WireDetection, type WireRange, type WireSelection } from '../../src/stream/sheet-wire'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import { DocumentHost } from '../src/render'
import type { SheetClock } from '../src/sheet/clock'
import type { SheetViewJob, SheetViewReply } from '../src/sheet/view-ops'
import type { CanonicalView } from '../src/sheet/view-spec'
import { Client, disposePool, start } from './harness'
import { colName, FILTER_AUTO_HIDDEN, filterAutoHidden, patchSheetXml, patchStylesXml, patchXlsx, setRowAttrs, SORT_GRID, sortGrid } from './sheet-fixtures'

type Policy = 'visible-only' | 'full'
const CLOCK: SheetClock = { day: '2026-11-27', timeZone: 'UTC' }
const IDENTITY: CanonicalView = { sort: null, sortFilter: null, filter: null }
const VIEW = ['sheet-view@1']
const T0 = Date.UTC(2026, 10, 27, 12)
const wire = (r0: number, r1: number, c0: number, c1: number, header: 0 | 1 = 1): WireRange => ({ r0, r1, c0, c1, header, totals: 0 })

const hosts: DocumentHost[] = []
afterAll(async () => {
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
  await disposePool()
})

async function serve(bytes: Uint8Array, policy: Policy) {
  const host = new DocumentHost()
  hosts.push(host)
  const { docKey } = await host.open('doc', bytes, policy, { sheetMode: 'tiles' })
  return (job: SheetViewJob) => host.sheetView(docKey, CLOCK, policy, job)
}
type Serve = Awaited<ReturnType<typeof serve>>

const at = (row: number, col: number, ranges: WireSelection['ranges'] = []): WireSelection => ({ active: [row, col], anchor: null, ranges })
async function detect(sv: Serve, kind: 'sort' | 'filter', selection: WireSelection): Promise<SheetViewReply> {
  return sv({ op: 'detect', sheetIndex: 0, kind, selection })
}
function detected(r: SheetViewReply): WireDetection {
  if (r.op !== 'detect' || !('detection' in r)) throw new Error(`detect: ${JSON.stringify(r)}`)
  return r.detection
}
const viewOf = (r: SheetViewReply): { view: CanonicalView; order: SheetOrder } => {
  if (r.op !== 'view' || !('order' in r)) throw new Error(`no view: ${JSON.stringify(r)}`)
  return r
}
/** The source row shown at display slot `d`, or null when the slot is hidden. */
function shownAt(order: SheetOrder, d: number, authoredHidden: ReadonlySet<number>): number | null {
  const band = order.bands.find((b) => b.r0 <= d && d <= b.r1)
  const src = band ? band.sourceAtOffset[d - band.r0]! : d
  const m = order.mask
  if (!m || d < m.r0 || d > m.r1) return authoredHidden.has(src) ? null : src
  const i = d - m.r0
  return (m.words[i >>> 5]! >>> (i & 31)) & 1 ? src : null
}
function shownSources(order: SheetOrder, r0: number, r1: number, authoredHidden: ReadonlySet<number>): Set<number> {
  const out = new Set<number>()
  for (let d = r0; d <= r1; d++) { const s = shownAt(order, d, authoredHidden); if (s !== null) out.add(s) }
  return out
}

const AUTO_HIDDEN = new Set(FILTER_AUTO_HIDDEN.hidden.map((r) => r - 1))
const GRID_HIDDEN = new Set([SORT_GRID.hiddenRow, SORT_GRID.zeroHeightRow])

// ---- fixtures for the no-leak pairs ---------------------------------------------------------------------------

/**
 * A1:C{rows} on sheet `S`: row 1 headers `Key`, `Num`, `Note`, then A `k-{r}`, B `r`, C `c-{r}`. `hidden` rows
 * (1-based) are saved `hidden="1"` and hold `fill(r)` (cells, possibly none) instead; `hiddenCols` (1-based) are
 * `hidden="1"` columns holding `colFill(r)` in every row.
 */
async function hiddenBook(o: { rows: number; hidden: number[]; fill: (r: number, s: (t: string) => number) => XlsxCellSpec[]; hiddenCols?: number[]; colFill?: (ref: string, s: (t: string) => number) => XlsxCellSpec | undefined }): Promise<Uint8Array> {
  const strings = ['Key', 'Num', 'Note']
  const s = (t: string) => { let i = strings.indexOf(t); if (i < 0) { i = strings.length; strings.push(t) } return i }
  const rows: XlsxSheetSpec['rows'] = Array.from({ length: o.rows }, (_, i) => {
    const r = i + 1
    const extra = (o.hiddenCols ?? []).flatMap((c) => { const cell = o.colFill?.(`${colName(c - 1)}${r}`, s); return cell ? [cell] : [] })
    if (r === 1) return { r, cells: [...['Key', 'Num', 'Note'].map((h, c): XlsxCellSpec => ({ ref: `${colName(c)}1`, t: 's', v: s(h) })), ...extra] }
    if (o.hidden.includes(r)) return { r, cells: [...o.fill(r, s), ...extra] }
    return { r, cells: [{ ref: `A${r}`, t: 's', v: s(`k-${r}`) }, { ref: `B${r}`, v: r }, { ref: `C${r}`, t: 's', v: s(`c-${r}`) }, ...extra] }
  })
  const cols = (o.hiddenCols ?? []).map((c) => `<col min="${c}" max="${c}" width="9" hidden="1" customWidth="1"/>`).join('')
  const bytes = await buildXlsx([{ name: 'S', rows, ...(cols ? { cols } : {}) }], strings)
  return patchXlsx(bytes, (zip) => patchSheetXml(zip, 0, (xml) => setRowAttrs(xml, o.hidden, { hidden: '1' })))
}
const SECRET = (r: number) => `SECRET-${r}`
/** Hidden-row contents: a text row, a number row, and nothing. */
const FILLS: Record<string, (r: number, s: (t: string) => number) => XlsxCellSpec[]> = {
  text: (r, s) => ['A', 'B', 'C'].map((c): XlsxCellSpec => ({ ref: `${c}${r}`, t: 's', v: s(SECRET(r)) })),
  number: (r) => [{ ref: `A${r}`, v: 900 + r }, { ref: `B${r}`, v: 900 + r }, { ref: `C${r}`, v: 900 + r }],
  empty: () => [],
}

// ---- detection parity ----------------------------------------------------------------------------------------

describe('see-through detection: visible-only matches full', () => {
  test('filterAutoHidden from B1: A1:C20 under both policies, for filter and sort', async () => {
    const bytes = await filterAutoHidden()
    const vo = await serve(bytes, 'visible-only'), full = await serve(bytes, 'full')
    for (const kind of ['filter', 'sort'] as const) {
      const want = detected(await detect(full, kind, at(0, 1)))
      expect(want.range).toEqual(wire(0, 19, 0, 2))
      expect({ kind, got: detected(await detect(vo, kind, at(0, 1))) }).toEqual({ kind, got: want })
    }
  }, 30_000)

  test('sortGrid from B1: the region steps over hidden row 9, ht=0 row 12 and width-0 column E', async () => {
    const bytes = await sortGrid()
    const vo = await serve(bytes, 'visible-only'), full = await serve(bytes, 'full')
    for (const kind of ['filter', 'sort'] as const) {
      const want = detected(await detect(full, kind, at(0, 1)))
      expect(want.range).toEqual(wire(0, SORT_GRID.rows - 1, 0, SORT_GRID.cols - 1))
      expect({ kind, got: detected(await detect(vo, kind, at(0, 1))) }).toEqual({ kind, got: want })
    }
    // Seeded below the hidden row, too.
    expect(detected(await detect(vo, 'filter', at(30, 3))).range).toEqual(wire(0, SORT_GRID.rows - 1, 0, SORT_GRID.cols - 1))
  }, 30_000)

  test('a header sort and a filter over the detected range commit; author-hidden rows stay hidden', async () => {
    const bytes = await filterAutoHidden()
    const vo = await serve(bytes, 'visible-only')
    const range = detected(await detect(vo, 'sort', at(0, 1))).range
    const sorted = viewOf(await vo({ op: 'sort', sheetIndex: 0, base: IDENTITY, sort: { range, keys: [{ col: 1, order: 'desc' }], caseSensitive: false } }))
    const shown = shownSources(sorted.order, 1, 19, AUTO_HIDDEN)
    for (const r of AUTO_HIDDEN) expect({ r, shown: shown.has(r) }).toEqual({ r, shown: false })
    // The first body row is the largest visible B: row 20 (source 19).
    expect(shownAt(sorted.order, 1, AUTO_HIDDEN)).toBe(19)

    const frange = detected(await detect(vo, 'filter', at(0, 1))).range
    const filtered = viewOf(await vo({ op: 'filter', sheetIndex: 0, base: IDENTITY, filter: { range: frange, columns: [[1, { kind: 'number', op: 'gt', first: 0 }]] } }))
    const kept = shownSources(filtered.order, 1, 19, AUTO_HIDDEN)
    for (const r of AUTO_HIDDEN) expect({ r, kept: kept.has(r) }).toEqual({ r, kept: false })
    expect(kept.size).toBe(19 - AUTO_HIDDEN.size)
  }, 30_000)

  test('content-withheld cells still end a region (Task 1 ruling stands for them)', async () => {
    // A `;;;` row splits the block: B4 under `;;;` (with A4, C4 empty) is withheld for content, not geometry.
    const strings = ['H1', 'H2']
    const rows: XlsxSheetSpec['rows'] = Array.from({ length: 8 }, (_, i) => {
      const r = i + 1
      if (r === 1) return { r, cells: [{ ref: 'A1', t: 's', v: 0 }, { ref: 'B1', t: 's', v: 1 }] }
      if (r === 4) return { r, cells: [{ ref: 'B4', v: 44, style: 3 }] }
      return { r, cells: [{ ref: `A${r}`, v: r }, { ref: `B${r}`, v: r * 2 }] }
    })
    const bytes = await patchXlsx(await buildXlsx([{ name: 'S', rows }], strings), (zip) => patchStylesXml(zip, (xml) => xml
      .replace('<fonts ', '<numFmts count="1"><numFmt numFmtId="166" formatCode=";;;"/></numFmts><fonts ')
      .replace('<cellXfs count="3">', '<cellXfs count="4">')
      .replace('</cellXfs>', '<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>')))
    const vo = await serve(bytes, 'visible-only'), full = await serve(bytes, 'full')
    expect(detected(await detect(full, 'filter', at(0, 0))).range).toEqual(wire(0, 7, 0, 1))
    expect(detected(await detect(vo, 'filter', at(0, 0))).range).toEqual(wire(0, 2, 0, 1))
  }, 30_000)
})

// ---- selection ------------------------------------------------------------------------------------------------

describe('see-through selection: a rectangle across author-hidden lines is one rectangle', () => {
  test('rows: [[0,2],[4,5],[7,19]] over hidden rows 3 and 6 is explicit A1:C20; a visible gap stays disjoint', async () => {
    const bytes = await filterAutoHidden()
    const vo = await serve(bytes, 'visible-only')
    const rect = detected(await detect(vo, 'filter', at(19, 1, [{ rows: [[0, 2], [4, 5], [7, 19]], cols: [[0, 2]] }])))
    expect(rect).toMatchObject({ range: wire(0, 19, 0, 2), source: 'explicit' })
    // A sub-rectangle B2:C12 across both hidden rows.
    const sub = detected(await detect(vo, 'filter', at(11, 2, [{ rows: [[1, 2], [4, 5], [7, 11]], cols: [[1, 2]] }])))
    expect(sub).toMatchObject({ range: { r0: 1, r1: 11, c0: 1, c1: 2 }, source: 'explicit' })
    // Row 5 (index 4) is visible and missing: disjoint, so the region from the active cell (as today).
    const gap = detected(await detect(vo, 'filter', at(11, 2, [{ rows: [[1, 2], [5, 5], [7, 11]], cols: [[1, 2]] }])))
    expect(gap).toMatchObject({ range: wire(0, 19, 0, 2), source: 'region' })
  }, 30_000)

  test('columns: cols [[1,3],[5,5]] over width-0 column E is one rectangle; rows over 9 and 12 too', async () => {
    const bytes = await sortGrid()
    const vo = await serve(bytes, 'visible-only')
    const rows: Array<[number, number]> = [[1, 7], [9, 10], [12, 20]]
    const r = detected(await detect(vo, 'filter', at(20, 5, [{ rows, cols: [[1, 3], [5, 5]] }])))
    expect(r).toMatchObject({ range: { r0: 1, r1: 20, c0: 1, c1: 5 }, source: 'explicit' })
    // Column D (3) missing from the selection is a visible gap.
    const gap = detected(await detect(vo, 'filter', at(20, 5, [{ rows, cols: [[1, 2], [5, 5]] }])))
    expect(gap.source).toBe('region')
  }, 30_000)

  test("'full' keeps the local rule: disjoint row runs seed the region", async () => {
    const full = await serve(await filterAutoHidden(), 'full')
    expect(detected(await detect(full, 'filter', at(11, 2, [{ rows: [[1, 2], [4, 5], [7, 11]], cols: [[1, 2]] }])))).toMatchObject({ range: wire(0, 19, 0, 2), source: 'region' })
  }, 30_000)
})

// ---- no oracle on hidden contents ------------------------------------------------------------------------------

describe('see-through never reads hidden contents', () => {
  /** Everything a visible-only client can learn from detection, the B value list and the views over the range. */
  async function answers(bytes: Uint8Array) {
    const vo = await serve(bytes, 'visible-only')
    const out: unknown[] = []
    for (const kind of ['filter', 'sort'] as const) {
      for (const sel of [at(0, 1), at(5, 2), at(2, 0, [{ rows: [[0, 2], [4, 5], [7, 11]], cols: [[0, 2]] }]), at(11, 0)]) out.push(await detect(vo, kind, sel))
    }
    const range = detected(await detect(vo, 'filter', at(0, 1))).range
    out.push(await vo({ op: 'values', sheetIndex: 0, view: IDENTITY, range, col: 1, offset: 0 }))
    out.push(await vo({ op: 'values', sheetIndex: 0, view: IDENTITY, range, col: 0, offset: 0 }))
    out.push(await vo({ op: 'sort', sheetIndex: 0, base: IDENTITY, sort: { range, keys: [{ col: 1, order: 'desc' }], caseSensitive: false } }))
    out.push(await vo({ op: 'filter', sheetIndex: 0, base: IDENTITY, filter: { range, columns: [[1, { kind: 'number', op: 'gt', first: 3 }]] } }))
    return JSON.stringify(out)
  }

  test('rows: text, number and empty hidden rows (inside and directly under the data) answer identically', async () => {
    const base = { rows: 12, hidden: [4, 7, 13] }
    const books = await Promise.all(Object.values(FILLS).map((fill) => hiddenBook({ ...base, rows: 13, fill })))
    const [text, number, empty] = await Promise.all(books.map(answers))
    expect(number).toBe(text!)
    expect(empty).toBe(text!)
    for (const r of [4, 7, 13]) expect(text).not.toContain(SECRET(r))
    // Non-vacuity: under 'full' detection is an oracle on the hidden rows (filled, the region reaches row 13; empty,
    // it stops above row 4); under 'visible-only' every variant answers A1:C12.
    const fullText = detected(await detect(await serve(books[0]!, 'full'), 'filter', at(0, 1)))
    const fullEmpty = detected(await detect(await serve(books[2]!, 'full'), 'filter', at(0, 1)))
    expect([fullText.range.r1, fullEmpty.range.r1]).toEqual([12, 2])
    expect(detected(JSON.parse(text!)[0] as SheetViewReply).range).toEqual(wire(0, 11, 0, 2))
  }, 30_000)

  test('columns: a hidden column D, populated or empty, answers identically', async () => {
    const base = { rows: 12, hidden: [4, 7], fill: FILLS.empty!, hiddenCols: [4] }
    const populated = await hiddenBook({ ...base, colFill: (ref, st) => ({ ref, t: 's', v: st(`SECRET-${ref}`) }) })
    const empty = await hiddenBook({ ...base, colFill: () => undefined })
    const [a, b] = await Promise.all([answers(populated), answers(empty)])
    expect(a).toBe(b)
    expect(a).not.toContain('SECRET')
    const fullPopulated = detected(await detect(await serve(populated, 'full'), 'filter', at(0, 1)))
    expect(fullPopulated.range.c1).toBe(3)
    expect(detected(JSON.parse(a)[0] as SheetViewReply).range).toEqual(wire(0, 11, 0, 2))
  }, 30_000)
})

// ---- on the socket --------------------------------------------------------------------------------------------

/** A server on an injected clock (each request steps it 100 ms), with its own audit; every text frame captured. */
async function serverOn(doc: Uint8Array) {
  const clock = { t: T0 }
  const store = AuditStore.open(':memory:')
  const h = await start({ doc, now: () => clock.t, tokenTtlSec: 86_400, audit: new AuditTable({ store }) })
  let rq = 1000
  const c = new Client(h.url)
  const m = await c.hello({ features: VIEW })
  await c.ready()
  while (h.stream.table.get(m.session.sid)?.status !== 'stream') await Bun.sleep(5)
  const texts: string[] = []
  const send = (type: string, body: Record<string, unknown>) => {
    clock.t += 100
    const reqId = ++rq
    c.send({ token: m.session.token, reqId, type, sheetId: '0', ...body })
    return reqId
  }
  const one = async (reqId: number) => {
    const a = await c.next((x) => 'text' in x && x.text.reqId === reqId) as { text: Record<string, any> }
    texts.push(JSON.stringify(a.text))
    return a.text
  }
  async function ordered(reqId: number): Promise<{ order?: SheetOrder; end?: Record<string, any> }> {
    const parts: Array<Record<string, any>> = []
    for (;;) {
      const x = await c.next((y) => 'text' in y && y.text.reqId === reqId, 20_000) as { text: Record<string, any> }
      texts.push(JSON.stringify(x.text))
      if (x.text.type !== 'sheet-order') return { end: x.text }
      parts.push(x.text)
      if (x.text.last === true) return { order: decodeSheetOrder(parts) }
    }
  }
  return {
    texts, store, sid: m.session.sid as string,
    detect: async (kind: 'sort' | 'filter', selection: WireSelection) => { const a = await one(send('range-detect', { viewGen: 0, kind, selection })); if (a.type !== 'range-detect-result') throw new Error(JSON.stringify(a)); return decodeDetection(a) },
    list: async (viewGen: number, range: WireRange, col: number) => { const a = await one(send('filter-values', { viewGen, range, col })); if (a.type !== 'filter-values-result') throw new Error(JSON.stringify(a)); return decodeValuesPage(a) },
    sort: (viewGen: number, body: Record<string, unknown>) => ordered(send('sheet-sort', { viewGen, ...body })),
    filter: (viewGen: number, body: Record<string, unknown>) => ordered(send('sheet-filter', { viewGen, ...body })),
    async stop() { c.close(); await h.stop() },
  }
}

describe('see-through on the socket (visible-only)', () => {
  test('filterAutoHidden: detect, sort descending, then filter by values; rows 4 and 7 stay hidden; nothing of them on the wire', async () => {
    const s = await serverOn(await filterAutoHidden())
    const d = await s.detect('sort', at(0, 1))
    expect(d.range).toEqual(wire(0, 19, 0, 2))
    // The shift-click rectangle as the client sends it.
    expect((await s.detect('filter', at(19, 1, [{ rows: [[0, 2], [4, 5], [7, 19]], cols: [[0, 2]] }]))).range).toEqual(wire(0, 19, 0, 2))
    const sorted = await s.sort(0, { range: d.range, keys: [{ col: 1, order: 'desc' }] })
    expect(sorted.end).toBeUndefined()
    expect(shownAt(sorted.order!, 1, AUTO_HIDDEN)).toBe(19)
    for (const r of AUTO_HIDDEN) expect({ r, shown: shownSources(sorted.order!, 1, 19, AUTO_HIDDEN).has(r) }).toEqual({ r, shown: false })
    const page = await s.list(1, d.range, 1)
    const labels = page.items.map((i) => i.text)
    expect(labels).not.toContain('4')
    expect(labels).not.toContain('7')
    const ids = page.items.filter((i) => i.text !== '10' && i.text !== '15').map((i): [number, number] => [i.id, i.id])
    const filtered = await s.filter(1, { range: d.range, columns: [[1, { kind: 'values', list: page.list, mode: 'include', ids, blanks: false }]] })
    expect(filtered.end).toBeUndefined()
    const kept = shownSources(filtered.order!, 1, 19, AUTO_HIDDEN)
    for (const r of [3, 6, 9, 14]) expect({ r, kept: kept.has(r) }).toEqual({ r, kept: false })
    expect(kept.size).toBe(19 - 4)
    // The wire audit for these answers: no hidden row's text.
    const all = s.texts.join('\n')
    for (const r of FILTER_AUTO_HIDDEN.hidden) expect(all).not.toContain(`drop-${r}`)
    expect(all).not.toContain('drop')
    await s.stop()
  }, 30_000)

  test('sortGrid: the B header sort is no longer refused outside-column-data', async () => {
    const s = await serverOn(await sortGrid())
    const d = await s.detect('sort', at(0, 1))
    expect(d.range).toEqual(wire(0, SORT_GRID.rows - 1, 0, SORT_GRID.cols - 1))
    const sorted = await s.sort(0, { range: d.range, keys: [{ col: 1, order: 'desc' }] })
    expect(sorted.end).toBeUndefined()
    // The `hidden="1"` row stays out of the mask. The `ht="0"` row is collapsed by its size, which travels with the
    // row (no filter, so no C1 on the mask), as locally.
    expect(shownSources(sorted.order!, 1, SORT_GRID.rows - 1, GRID_HIDDEN).has(SORT_GRID.hiddenRow)).toBe(false)
    const page = await s.list(1, d.range, 0)
    const all = s.texts.join('\n') + JSON.stringify(page.items)
    for (const r of GRID_HIDDEN) expect(all).not.toContain(`"N${r + 1}"`)
    await s.stop()
  }, 30_000)
})
