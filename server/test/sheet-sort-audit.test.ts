/**
 * Stage 4c Task 13 — the 4c wire audit, the pivot strip on the wire and the hostile battery (S4 §6.6, §6.9, §9 4c;
 * plan P4, P6, P7, P12, P13; parent §12 C1).
 *
 * **Wire audit.** A `sheet-view@1` session (visible-only, `copy: 'full'`) is captured byte for byte by a listener of
 * its own (text frames, raw binary frames, binary messages reassembled). Per visible sheet it range-detects from A1
 * (sort and filter), then over the detected range and the used range from A1 (headerless), and for `sortWithheld` also
 * over the band A21:F40 (every sort or filter holding row 20 is refused `drawing-impact`: B20 is under a cell-anchored
 * shape), for each column: lists the column's values (every page), filters by the list's first two entries, clears,
 * filters `text contains "e"`, clears, sorts ascending, clears. After the first committed sort and the first committed
 * filter of a session, and once at the identity view, it runs the 4b text answers under that view (find, copy of the
 * used range, a11y over the used range, fit of every used column). Each sheet ends with a descending sort of its last band, `sheet-order`, and every
 * tile of the used extent at (zoomStep 5, dpr 1) under that view.
 *
 * Fixtures: the brief's (`sentinels()`, every XLSX hidden fixture, `sortWithheld`, `paintedFormats`, `rotatedBox`,
 * `offViewportCovered`, `unresolvedAnchor`) plus one per remaining withholding class: `mergeTail` (merge tail),
 * `hiddenFormat` (`;;;`), `cfHiddenValue` (CF `showValue="0"`), `tableHiddenFormat` (table region style),
 * `zeroHeightRow` (an `ht="0"` row) and `filterAutoHidden` (author-hidden rows under see-through detection, and C1).
 * Short withheld values (`1`, `5`, `95`, `1234`, `5678`, `10`, `N9`, …) are checked by exact equality on decoded answer
 * fields (value-list items, copy fields, a11y texts), never by substring over raw frames.
 *
 * Rates: the server's clock is the test's own, stepped 100 ms per op (1.7 s per tile batch) and 61 s every 25 ops, so
 * the `'sheet'`, `'text'`, `'unit'`, `'tile'` and message windows admit the audit. The token lives a day.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { viewerControls } from '../../src/components/xlsx/viewerControlsClient'
import type { CacheIdentity, FilterRange, SheetView, SortSpec } from '../../src/xlsx/view/contracts'
import { MainMemoryLedger } from '../../src/xlsx/view/mainLedger'
import { createIdentityOrder } from '../../src/xlsx/view/order'
import { emptyGridSelection } from '../../src/xlsx/view/selection'
import { createObjectWorkbookSource } from '../../src/xlsx/view/source'
import { visibilityFromHidden } from '../../src/xlsx/view/visibility'
import type { XlsxCell, XlsxDocument } from '../../src/xlsx/types'
import { Reassembler } from '../../src/stream/frame'
import {
  decodeDetection, decodeSheetLayout, decodeSheetOrder, decodeValuesPage, TILE_ZOOM_STEPS, type SheetOrder, type StreamSheetLayout, type ValuesPage,
  type WireDetection, type WireRange, type WireSelection,
} from '../../src/stream/sheet-wire'
import { streamSheetGeometry } from '../../src/xlsx/view/streamGeometry'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import type { AuthDecision } from '../src/hooks'
import type { RateClass, SessionRates } from '../src/rate'
import { tileLimit } from '../src/sheet/tile'
import { augmentedFixtures } from './fixtures'
import { Client, disposePool, start, type Harness } from './harness'
import {
  cfHiddenValue, CF_HIDDEN_VALUES, FILTER_AUTO_HIDDEN, filterAutoHidden, hiddenFormat, HIDDEN_PIVOT, hiddenPivotCache, insertSheetElement, mergeTail,
  offViewportCovered, PAINTED_FORMATS, paintedFormats, patchSheetXml, patchXlsx, rotatedBox, SENTINELS, sentinels, SORT_GRID, SORT_WITHHELD,
  sortGrid, sortWithheld, TABLE_HIDDEN, tableHiddenFormat, unresolvedAnchor, ZERO_HEIGHT, zeroHeightRow,
} from './sheet-fixtures'

afterAll(async () => { await disposePool() })

const VIEW = ['sheet-view@1']
const T0 = Date.UTC(2026, 10, 27, 12)
const PANE = { x: 0, y: 0, width: 640, height: 400 }
const A11Y_CELLS = 2000
const CAPS: Record<number, [number, number]> = { 1: [256 << 10, 256 << 10], 2: [1 << 20, 8 << 20], 3: [8 << 20, 64 << 20], 4: [8 << 20, 8 << 20], 5: [8 << 20, 64 << 20], 6: [8 << 20, 32 << 20], 7: [8 << 20, 64 << 20], 8: [8 << 20, 64 << 20], 9: [4 << 20, 4 << 20] }
const latin1 = (b: Uint8Array) => Buffer.from(b.buffer, b.byteOffset, b.byteLength).toString('latin1')
const wire = (r0: number, r1: number, c0: number, c1: number, header: 0 | 1 = 1): WireRange => ({ r0, r1, c0, c1, header, totals: 0 })
const local = (r: WireRange): FilterRange => ({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1, headerRows: r.header, totalsRows: r.totals })
const rcOf = (ref: string): [number, number] => [Number(ref.slice(1)) - 1, ref.charCodeAt(0) - 65]
const sameBytes = (a: Uint8Array, b: Uint8Array) => Buffer.compare(Buffer.from(a), Buffer.from(b)) === 0
type Text = Record<string, any>
type Policy = AuthDecision['policy']

// ---- the server, its clock, a session ------------------------------------------------------------------------------

/** One server over `doc` on the test's clock (see the header), with its own audit store. */
async function server(doc: Uint8Array, policy: Policy = {}) {
  let skew = 0, ops = 0
  const store = AuditStore.open(':memory:')
  const h = await start({ doc, policy: { copy: 'full', visibleMark: 'off', forensicMark: false, ...policy }, now: () => T0 + skew, tokenTtlSec: 86_400, audit: new AuditTable({ store }) })
  const tick = (ms = 100) => { skew += ms; if (++ops % 25 === 0) skew += 61_000 }
  return { h, store, tick }
}
type Server = Awaited<ReturnType<typeof server>>

/** Everything the server sent one client: text frames, raw binary frames, and binary messages reassembled. */
interface Capture { texts: string[]; frames: Uint8Array[]; payloads: Uint8Array[]; drain(): Promise<void> }
function capture(c: Client): Capture {
  const cap: Capture = { texts: [], frames: [], payloads: [], drain: () => inbox }
  const r = new Reassembler({ acceptedVersions: new Set([1]), expects: () => true, decodedCap: (k) => CAPS[k]?.[1] ?? 0, encodedCap: (k) => CAPS[k]?.[0] ?? 0 })
  let inbox = Promise.resolve()
  c.ws.addEventListener('message', (e) => {
    if (typeof e.data === 'string') { cap.texts.push(e.data); return }
    const frame = new Uint8Array((e.data as ArrayBuffer).slice(0))
    cap.frames.push(frame)
    inbox = inbox.then(async () => {
      const res = await r.push(frame)
      if (res && !res.ok) throw new Error(`audit reassembler rejected a frame: ${res.reason}`)
      if (res?.ok) cap.payloads.push(res.payload)
    })
  })
  return cap
}

/** A `sheet-view@1` session on `s`, with the 4c and 4b ops as helpers (each steps the server clock). */
async function session(s: Server) {
  const c = new Client(s.h.url)
  const cap = capture(c)
  const m = await c.hello({ features: VIEW })
  await c.ready()
  while (s.h.stream.table.get(m.session.sid)?.status !== 'stream') await Bun.sleep(5)
  expect([m.sheetMode, m.sheetView, (m as Text).sheetSort]).toEqual(['tiles', 1, 1])
  const sid = m.session.sid as string
  let rq = 100
  /** Sends `body` and collects its reqId: every `partial` message until `last: true`, or the first other message. */
  async function ask(type: string, body: Text, partial?: string): Promise<{ parts: Text[]; end?: Text }> {
    s.tick()
    const reqId = ++rq
    c.send({ token: m.session.token, reqId, type, ...body })
    const parts: Text[] = []
    for (;;) {
      const x = await c.next((y) => 'text' in y && y.text.reqId === reqId, 30_000) as { text: Text }
      if (x.text.type !== partial) return { parts, end: x.text }
      parts.push(x.text)
      if (x.text.last === true) return { parts }
    }
  }
  const strip = ({ reqId: _r, ...t }: Text) => t
  return {
    c, cap, sid, manifest: m as Text,
    async layout(sheetId: string): Promise<{ layout: StreamSheetLayout; parts: Text[] }> {
      const r = await ask('sheet-layout', { sheetId }, 'sheet-layout')
      expect([sheetId, r.end]).toEqual([sheetId, undefined])
      return { layout: decodeSheetLayout(r.parts), parts: r.parts.map(strip) }
    },
    async detect(sheetId: string, viewGen: number, kind: 'sort' | 'filter', selection: WireSelection): Promise<WireDetection | Text> {
      const r = await ask('range-detect', { sheetId, viewGen, kind, selection })
      return r.end!.type === 'range-detect-result' ? decodeDetection(r.end) : r.end!
    },
    /** Every page of one list, or the refusal. */
    async list(sheetId: string, viewGen: number, range: WireRange, col: number): Promise<{ pages: ValuesPage[]; items: ValuesPage['items']; first?: ValuesPage; end?: Text }> {
      const pages: ValuesPage[] = []
      let a = (await ask('filter-values', { sheetId, viewGen, range, col })).end!
      for (;;) {
        if (a.type !== 'filter-values-result') return { pages, items: [], end: a }
        const p = decodeValuesPage(a)
        pages.push(p)
        if (!p.cursor) break
        a = (await ask('filter-values', { sheetId, viewGen, cursor: p.cursor })).end!
      }
      return { pages, items: pages.flatMap((p) => p.items), first: pages[0] }
    },
    async view(type: 'sheet-sort' | 'sheet-filter', body: Text): Promise<{ parts: Text[]; order?: SheetOrder; end?: Text }> {
      const r = await ask(type, body, 'sheet-order')
      return r.end ? r : { parts: r.parts, order: decodeSheetOrder(r.parts) }
    },
    async order(sheetId: string): Promise<{ parts: Text[]; order: SheetOrder }> {
      const r = await ask('sheet-order', { sheetId }, 'sheet-order')
      expect([sheetId, r.end]).toEqual([sheetId, undefined])
      return { parts: r.parts, order: decodeSheetOrder(r.parts) }
    },
    async find(query: string) {
      const r = await ask('sheet-find', { query }, 'sheet-find-result')
      expect([query, r.end?.type]).toEqual([query, 'sheet-find-done'])
      return { query, parts: r.parts, end: r.end!, hits: r.parts.flatMap((p) => p.hits as number[][]) }
    },
    async copy(sheetId: string, viewGen: number, r1: number, c1: number): Promise<string> {
      const r = await ask('sheet-copy', { sheetId, viewGen, selection: { active: [0, 0], anchor: [0, 0], ranges: [{ rows: [[0, r1]], cols: [[0, c1]] }] } }, 'copy-result')
      expect([sheetId, r.end]).toEqual([sheetId, undefined])
      return r.parts.map((p) => p.text as string).join('')
    },
    async a11y(sheetId: string, viewGen: number, r1: number, c1: number): Promise<Array<[number, number, string]>> {
      const out: Array<[number, number, string]> = []
      const colsPer = Math.min(c1 + 1, A11Y_CELLS)
      for (let c0 = 0; c0 <= c1; c0 += colsPer) {
        const cEnd = Math.min(c1, c0 + colsPer - 1), rowsPer = Math.floor(A11Y_CELLS / (cEnd - c0 + 1))
        for (let r0 = 0; r0 <= r1; r0 += rowsPer) {
          const r = await ask('sheet-a11y', { sheetId, viewGen, rows: [[r0, Math.min(r1, r0 + rowsPer - 1)]], cols: [c0, cEnd] }, 'sheet-a11y-result')
          expect([sheetId, r0, r.end]).toEqual([sheetId, r0, undefined])
          for (const p of r.parts) out.push(...(p.cells as Array<[number, number, string]>))
        }
      }
      return out
    },
    async fit(sheetId: string, viewGen: number, index: number): Promise<Text> {
      return (await ask('sheet-fit', { sheetId, viewGen, axis: 'col', index, pane: PANE })).end!
    },
    /** Every tile of `coords` at (zoomStep, dpr 1), in batches of 64; the PNGs by `col,row`. */
    async tiles(sheetId: string, viewGen: number, zoomStep: number, coords: Array<{ col: number; row: number }>): Promise<Map<string, Uint8Array>> {
      const out = new Map<string, Uint8Array>()
      for (let i = 0; i < coords.length; i += 64) {
        s.tick(1700)
        const reqId = ++rq
        const r = await c.tiles(reqId, { sheetId, viewGen, zoomStep, dpr: 1, palette: 'light', tiles: coords.slice(i, i + 64) }, { ms: 30_000 })
        expect([sheetId, r.end]).toEqual([sheetId, { reqId, type: 'tiles-done', refused: [] }])
        for (const t of r.tiles) out.set(`${t.col},${t.row}`, t.png)
      }
      expect(out.size).toBe(coords.length)
      return out
    },
  }
}
type Session = Awaited<ReturnType<typeof session>>

/** Every tile coordinate of `layout`'s used extent at `zoomStep`, dpr 1. */
function usedTiles(layout: StreamSheetLayout, zoomStep: number): Array<{ col: number; row: number }> {
  const mt = streamSheetGeometry(layout).metrics, u = layout.used!, z = TILE_ZOOM_STEPS[zoomStep]!
  const cols = tileLimit(mt.cols.start(u.c1) + mt.cols.size(u.c1), z, 1), rows = tileLimit(mt.rows.start(u.r1) + mt.rows.size(u.r1), z, 1)
  return Array.from({ length: cols * rows }, (_, i) => ({ col: i % cols, row: Math.floor(i / cols) }))
}

/** Source rows of `[r0, r1]` an order shows (display slot through the band, then the mask; a slot outside the mask
 * shows unless `authoredHidden` holds its source). */
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

// ---- the audit -----------------------------------------------------------------------------------------------------

interface ViewRec { sheetId: string; kind: 'sort' | 'filter'; rule: string; range: WireRange; col: number; parts: Text[]; order?: SheetOrder; end?: Text }
interface ListRec { sheetId: string; range: WireRange; col: number; items: ValuesPage['items']; blanks?: ValuesPage['blanks']; end?: Text }
interface Audited {
  name: string; manifest: Text; cap: Capture
  detections: Array<{ sheetId: string; kind: string; answer: WireDetection | Text }>
  lists: ListRec[]; views: ViewRec[]; orders: Text[][]
  finds: Array<{ query: string; hits: number[][]; parts: Text[]; end: Text }>
  copies: string[]; cells: Array<[number, number, string]>; fits: Text[]; tiles: number
}

/** The audit's ops over every visible sheet of `doc` (see the header); `bands` adds explicit bands for sheet 0. */
async function audited(name: string, doc: Uint8Array, queries: readonly string[], bands: readonly WireRange[] = []): Promise<Audited> {
  const s = await server(doc)
  const v = await session(s)
  const out: Audited = { name, manifest: v.manifest, cap: v.cap, detections: [], lists: [], views: [], orders: [], finds: [], copies: [], cells: [], fits: [], tiles: 0 }
  let textSorted = false, textFiltered = false
  try {
    for (const { sheetId } of v.manifest.sheets as Array<{ sheetId: string }>) {
      const { layout } = await v.layout(sheetId)
      if (!layout.used) continue
      const { r1, c1 } = layout.used
      let gen = layout.viewGen
      const textOps = async () => {
        for (const q of queries) out.finds.push(await v.find(q))
        out.copies.push(await v.copy(sheetId, gen, r1, c1))
        out.cells.push(...await v.a11y(sheetId, gen, r1, c1))
        for (let col = 0; col <= c1; col++) out.fits.push(await v.fit(sheetId, gen, col))
      }
      if (sheetId === '0') await textOps()
      const a1: WireSelection = { active: [0, 0], anchor: null, ranges: [] }
      let detected: WireRange | undefined
      for (const kind of ['sort', 'filter'] as const) {
        const answer = await v.detect(sheetId, gen, kind, a1)
        out.detections.push({ sheetId, kind, answer })
        if (kind === 'filter' && 'source' in answer) detected = answer.range as WireRange
      }
      // The detected range, then the used range from A1 (headerless, so every used column and row is listed and keyed).
      const used = wire(0, r1, 0, c1, 0)
      const same = (x: WireRange, y: WireRange) => JSON.stringify(x) === JSON.stringify(y)
      const sheetBands = [...(detected && !same(detected, used) ? [detected] : []), used, ...(sheetId === '0' ? bands : [])]
      /** Records the op; a commit moves `gen`, and is cleared again (`clear`) unless it is the sheet's last op. */
      const run = async (rec: Omit<ViewRec, 'parts' | 'order' | 'end'>, body: Text, clear: boolean) => {
        const r = await v.view(rec.kind === 'sort' ? 'sheet-sort' : 'sheet-filter', { sheetId, viewGen: gen, range: rec.range, ...body })
        out.views.push({ ...rec, ...r })
        if (!r.order) return false
        gen = r.order.head.viewGen
        if (rec.kind === 'sort' && !textSorted) { textSorted = true; await textOps() }
        if (rec.kind === 'filter' && !textFiltered) { textFiltered = true; await textOps() }
        if (clear) {
          const c = await v.view(rec.kind === 'sort' ? 'sheet-sort' : 'sheet-filter', { sheetId, viewGen: gen, range: rec.range, ...(rec.kind === 'sort' ? { keys: [] } : { columns: [] }) })
          expect([name, rec.rule, c.end]).toEqual([name, rec.rule, undefined])
          gen = c.order!.head.viewGen
        }
        return true
      }
      for (const band of sheetBands) {
        for (let col = band.c0; col <= band.c1; col++) {
          const l = await v.list(sheetId, gen, band, col)
          out.lists.push({ sheetId, range: band, col, items: l.items, ...(l.first ? { blanks: l.first.blanks } : {}), ...(l.end ? { end: l.end } : {}) })
          if (l.first) {
            const ids: Array<[number, number]> = l.items.slice(0, 2).map((i) => [i.id, i.id])
            await run({ sheetId, kind: 'filter', rule: 'values', range: band, col }, { columns: [[col, { kind: 'values', list: l.first.list, mode: 'include', ids, blanks: ids.length === 0 }]] }, true)
          }
          await run({ sheetId, kind: 'filter', rule: 'text', range: band, col }, { columns: [[col, { kind: 'text', op: 'contains', first: 'e', wildcards: false }]] }, true)
          await run({ sheetId, kind: 'sort', rule: 'asc', range: band, col }, { keys: [{ col, order: 'asc' }] }, true)
        }
      }
      // The sheet's last view: a descending sort of its last band, then its order and every tile of the used extent.
      const last = sheetBands.at(-1)!
      await run({ sheetId, kind: 'sort', rule: 'desc', range: last, col: last.c0 }, { keys: [{ col: last.c0, order: 'desc' }] }, false)
      out.orders.push((await v.order(sheetId)).parts)
      out.tiles += (await v.tiles(sheetId, gen, 5, usedTiles(layout, 5))).size
    }
    await v.cap.drain()
    return out
  } finally { v.c.close(); await s.h.stop() }
}

/** Every decoded text a session received: text frames, reassembled payloads and raw frames (latin1), and each copy. */
const decoded = (s: Audited): string[] => [...s.cap.texts, ...s.cap.payloads.map(latin1), ...s.cap.frames.map(latin1), ...s.copies]
/** Every string leaf (and key) of every JSON text frame and of the manifest. */
function stringLeaves(s: Audited): string[] {
  const out: string[] = []
  const walk = (v: unknown): void => {
    if (typeof v === 'string') out.push(v)
    else if (Array.isArray(v)) v.forEach(walk)
    else if (v && typeof v === 'object') for (const [k, x] of Object.entries(v)) { out.push(k); walk(x) }
  }
  for (const t of s.cap.texts) walk(JSON.parse(t))
  walk(s.manifest)
  return out
}
const leaked = (hay: readonly string[], needles: readonly string[]) => needles.filter((n) => hay.some((h) => h.includes(n)))
/** Every text field of every answer: value-list items, copy TSV fields, a11y texts. */
const answerFields = (s: Audited): string[] => [...s.lists.flatMap((l) => l.items.map((i) => i.text)), ...s.copies.flatMap((x) => x.split('\n').flatMap((l) => l.split('\t'))), ...s.cells.map(([, , t]) => t)]

/** The brief's needles (a hidden sheet's name and the markers are added per run), plus each added class's text. */
const NEEDLES = [
  SENTINELS.formula, SENTINELS.definedName, SENTINELS.number, SENTINELS.formulaText, SENTINELS.hiddenSheet, SENTINELS.hiddenCell, SENTINELS.comment, SENTINELS.drawingName,
  'COVER-B20', SORT_WITHHELD.maskedText, 'OFFVIEW-Z5000', 'ROT-1-1', 'ROT-2-1', 'ROT-3-1',
  'TAIL-B1', 'HIDDENFMT', 'UNRES-', ZERO_HEIGHT.sentinel, TABLE_HIDDEN.header, String(TABLE_HIDDEN.value), 'drop-4', 'drop-7',
  // sortGrid's author-hidden lines: D9 (row 9, hidden="1") and D12 (row 12, ht="0").
  `${SORT_GRID.words[9 % 8]}-9`, `${SORT_GRID.words[12 % 8]}-12`,
]
/** Per fixture: exact answer fields that must never occur, and visible ones that must (liveness). */
interface AuditFixture { name: string; doc: () => Promise<Uint8Array>; banned?: readonly string[]; controls: readonly string[]; bands?: readonly WireRange[] }
const BAND = wire(20, SORT_GRID.rows - 1, 0, SORT_GRID.cols - 1, 0)
const AUDIT_FIXTURES: readonly AuditFixture[] = [
  { name: 'sentinels', doc: sentinels, controls: [SENTINELS.visibleText] },
  // B20's 1, B21's 5, B22's 95; A9 and A12 (author-hidden rows); column E's r × 10 of rows 9 and 12.
  { name: 'sortWithheld', doc: sortWithheld, banned: ['1', '5', '95', 'N9', 'N12', '90', '120'], controls: [SORT_WITHHELD.maskedDisplay, 'N10', 'apple-8'], bands: [BAND] },
  // D2:D4 and A6/C6 sit under a column or row default style with no own `s`: the parser keeps no column or row style
  // for them, so they paint and are answered in both views (Task 1, accepted).
  { name: 'paintedFormats', doc: paintedFormats, banned: [String(PAINTED_FORMATS.tableValue), String(PAINTED_FORMATS.cfValue)], controls: [PAINTED_FORMATS.tablePaints, PAINTED_FORMATS.cfPaints] },
  { name: 'rotatedBox', doc: rotatedBox, controls: ['ROT-4-1', 'ROT-1-2'] },
  { name: 'offViewportCovered', doc: offViewportCovered, controls: ['OFFVIEW-A1'] },
  { name: 'unresolvedAnchor', doc: unresolvedAnchor, controls: [] },
  { name: 'mergeTail', doc: mergeTail, controls: ['MERGE-A1'] },
  { name: 'hiddenFormat', doc: hiddenFormat, controls: ['SHOWN-A1'] },
  { name: 'cfHiddenValue', doc: cfHiddenValue, banned: CF_HIDDEN_VALUES.map(String), controls: ['CFTEXT-B4'] },
  { name: 'tableHiddenFormat', doc: tableHiddenFormat, controls: [TABLE_HIDDEN.outside] },
  { name: 'zeroHeightRow', doc: zeroHeightRow, controls: [ZERO_HEIGHT.above, ZERO_HEIGHT.below] },
  { name: 'filterAutoHidden', doc: filterAutoHidden, banned: ['4', '7', 'C4', 'C7'], controls: ['keep-2', 'keep-20'] },
]

const runs = new Map<string, Audited>()
let markers: string[] = []
const queriesFor = (controls: readonly string[]) => [...NEEDLES, ...markers, 'MARKER-', 'H', ...controls]

describe('4c wire audit (visible-only, copy full)', () => {
  test('capture: every XLSX hidden fixture', async () => {
    const hidden = (await augmentedFixtures()).filter((f) => f.kind === 'xlsx')
    expect(hidden.map((f) => f.id).sort()).toEqual(['xlsx-comment', 'xlsx-hidden-cols', 'xlsx-hidden-rows', 'xlsx-hidden-sheet', 'xlsx-veryhidden-sheet'])
    markers = hidden.map((f) => f.marker)
    for (const f of hidden) runs.set(f.id, { ...await audited(f.id, f.withHidden, queriesFor([f.control])), controls: [f.control] } as Audited)
  }, 30_000)
  for (const f of AUDIT_FIXTURES) {
    test(`capture: ${f.name}`, async () => {
      runs.set(f.name, { ...await audited(f.name, await f.doc(), queriesFor(f.controls), f.bands), controls: f.controls } as Audited)
    }, 30_000)
  }

  test('no sentinel or withheld text in any 4c answer', () => {
    const all = [...runs.values()] as Array<Audited & { controls: readonly string[] }>
    expect(all.length).toBe(5 + AUDIT_FIXTURES.length)
    for (const s of all) {
      // Non-vacuity: every op class answered, and the fixture's visible text flows through lists, copy, a11y and find.
      expect([s.name, s.lists.length > 0, s.views.length > 0, s.copies.length > 0, s.fits.length > 0, s.orders.length > 0, s.tiles > 0]).toEqual([s.name, true, true, true, true, true, true])
      for (const ctl of s.controls) {
        expect([s.name, ctl, s.lists.some((l) => l.items.some((i) => i.text === ctl))]).toEqual([s.name, ctl, true])
        expect([s.name, ctl, s.copies.some((x) => x.includes(ctl)), s.cells.some(([, , t]) => t === ctl)]).toEqual([s.name, ctl, true, true])
        expect([s.name, ctl, s.finds.find((x) => x.query === ctl)!.hits.length > 0]).toEqual([s.name, ctl, true])
      }
      // The needles: no decoded message of any session holds one.
      expect([s.name, leaked(decoded(s), [...NEEDLES, ...markers, 'MARKER-'])]).toEqual([s.name, []])
      // A hidden sheet's name: no transmitted string is it.
      expect([s.name, stringLeaves(s).filter((x) => x === 'H' || x === SENTINELS.hiddenSheet)]).toEqual([s.name, []])
      // The exact withheld values: no value-list item, copy field or a11y text is one.
      const banned = AUDIT_FIXTURES.find((f) => f.name === s.name)?.banned ?? []
      expect([s.name, banned.filter((b) => answerFields(s).includes(b))]).toEqual([s.name, []])
    }
    const byName = (n: string) => runs.get(n)!
    // Committed views per fixture (so the audit ran answers under sorted and filtered views); unknown coverage commits
    // nothing: every sort and filter is refused coverage-unknown, and every list is blanks only.
    for (const s of all) {
      const committed = s.views.filter((x) => x.order).length
      if (s.name === 'unresolvedAnchor') {
        expect([...new Set(s.views.map((x) => x.end?.code))]).toEqual(['coverage-unknown'])
        expect(s.lists.map((l) => [l.items.length, l.blanks?.present])).toEqual(s.lists.map(() => [0, true]))
      } else if (s.name === 'mergeTail') {
        // Every sort holds the merge A1:C1 and is refused as locally; filters commit; B's list is blanks only.
        expect([...new Set(s.views.filter((x) => x.kind === 'sort').map((x) => `${x.end?.code}:${x.end?.reason}`))]).toEqual(['view-refused:merge'])
        expect(s.views.some((x) => x.kind === 'filter' && x.order)).toBe(true)
        expect(s.lists.filter((l) => l.col === 1).map((l) => [l.items.length, l.blanks?.present])).toEqual([[0, true]])
      } else expect([s.name, committed > 0, s.views.some((x) => x.kind === 'sort' && x.order), s.views.some((x) => x.kind === 'filter' && x.order)]).toEqual([s.name, true, true, true])
    }
    // sortWithheld: the band commits sorts and filters; its whole-grid sorts are refused drawing-impact (B20), and a
    // key on the zero-width column E hidden-target.
    const sw = byName('sortWithheld')
    expect(sw.views.filter((x) => x.range.r0 === BAND.r0 && x.order).length).toBeGreaterThanOrEqual(3 * 5)
    const gridSorts = sw.views.filter((x) => x.range.r0 === 0 && x.kind === 'sort')
    expect(gridSorts.map((x) => x.end?.reason ?? x.end?.code)).toEqual(gridSorts.map((x) => x.col === SORT_GRID.zeroWidthCol ? 'hidden-target' : 'drawing-impact'))
    expect(gridSorts.length).toBeGreaterThan(0)
    // Column E (zero width) is never listed: hidden-target.
    expect(sw.lists.filter((l) => l.col === SORT_GRID.zeroWidthCol).map((l) => l.end?.code)).toEqual(['hidden-target', 'hidden-target'])
    // B lists the withheld cells only as blanks.
    for (const l of sw.lists.filter((x) => x.col === 1)) expect(l.blanks?.present).toBe(true)
    // Masked D23 lists as its painted text; a text rule never matches the raw string.
    expect(sw.lists.filter((l) => l.col === 3).every((l) => l.items.some((i) => i.text === SORT_WITHHELD.maskedDisplay))).toBe(true)
    console.log(`4c wire audit: ${all.map((s) => `${s.name} ${s.lists.length} lists, ${s.views.filter((x) => x.order).length}/${s.views.length} views committed, ${s.finds.length} finds, ${s.cells.length} a11y cells, ${s.fits.length} fits, ${s.tiles} tiles, ${s.cap.texts.length + s.cap.frames.length} frames`).join('; ')}`)
    // filterAutoHidden: see-through detection spans the author-hidden rows (A1:C20).
    expect(byName('filterAutoHidden').detections.find((d) => d.kind === 'filter')!.answer).toMatchObject({ range: wire(0, 19, 0, 2) })
  })

  test('order answers carry positions only', () => {
    let parts = 0, heads = 0
    const HEAD_KEYS = ['sheetId', 'viewGen', 'counts', 'sort', 'filterSummary', 'caseSensitive', 'filterRange']
    for (const s of runs.values()) {
      const all = [...s.views.map((x) => x.parts), ...s.orders]
      for (const answer of all) for (const p of answer) {
        parts++
        expect([s.name, Object.keys(p).filter((k) => !['reqId', 'type', 'seq', 'last', 'head', 'band', 'mask'].includes(k))]).toEqual([s.name, []])
        expect(p.type).toBe('sheet-order')
        if (p.head) {
          heads++
          expect([s.name, Object.keys(p.head).filter((k) => !HEAD_KEYS.includes(k))]).toEqual([s.name, []])
          // The head's only strings are the sheet id, rule kinds and sort orders.
          const strings: string[] = []
          const walk = (v: unknown): void => { if (typeof v === 'string') strings.push(v); else if (Array.isArray(v)) v.forEach(walk); else if (v && typeof v === 'object') Object.values(v).forEach(walk) }
          walk(p.head)
          expect([s.name, strings.filter((x) => !/^(\d+|asc|desc|values|text|number|date|datePreset|compound)$/.test(x))]).toEqual([s.name, []])
        }
        for (const k of ['band', 'mask'] as const) if (p[k]) expect([s.name, Object.keys(p[k]).sort()]).toEqual([s.name, ['data', 'from', 'r0', 'r1']])
      }
    }
    expect(parts).toBeGreaterThan(0)
    expect(heads).toBeGreaterThan(0)
  })

  test('C1 and withheld keys in every order delta', async () => {
    // filterAutoHidden: no answer shows the author-hidden rows 4 and 7.
    const hidden = new Set(FILTER_AUTO_HIDDEN.hidden.map((r) => r - 1))
    const fah = runs.get('filterAutoHidden')!
    let checked = 0
    for (const x of [...fah.views.filter((y) => y.order).map((y) => y.order!), ...fah.orders.map((p) => decodeSheetOrder(p))]) {
      const shown = shownSources(x, 0, 19, hidden)
      expect([...hidden].filter((r) => shown.has(r))).toEqual([])
      checked++
    }
    expect(checked).toBeGreaterThan(0)
    // sortWithheld: every committed band answer equals the local engine over the fixture with its withheld cells deleted
    // (Task 5's oracle): sorts by order, filters by shown rows and counts.
    const bytes = await sortWithheld()
    const sw = runs.get('sortWithheld')!
    const recs = sw.views.filter((x) => x.order && x.range.r0 === BAND.r0)
    expect(recs.length).toBeGreaterThan(0)
    const lists = new Map(sw.lists.filter((l) => l.range.r0 === BAND.r0).map((l) => [l.col, l]))
    for (const x of recs) {
      const o = await oracle(await withheldOracle(bytes))
      try {
        if (x.kind === 'sort') {
          expect(await o.sort({ range: local(x.range), keys: [{ col: x.col, order: x.rule as 'asc' | 'desc' }], caseSensitive: false })).toBe('committed')
          expect({ col: x.col, rule: x.rule, order: Array.from(x.order!.bands[0]!.sourceAtOffset) }).toEqual({ col: x.col, rule: x.rule, order: o.order(x.range.r0, x.range.r1) })
        } else {
          const picked = new Set(lists.get(x.col)!.items.slice(0, 2).map((i) => i.text))
          const rule = x.rule === 'text' ? { kind: 'text' as const, op: 'contains' as const, first: 'e', wildcards: false }
            : await o.values(local(x.range), x.col, (l) => picked.has(l), picked.size === 0)
          const r = rule ? await o.filter({ range: local(x.range), columns: new Map([[x.col, rule]]) }) : 'unchanged'
          expect(r).toMatch(/committed|unchanged/)
          expect({ col: x.col, rule: x.rule, shown: shownSources(x.order!, x.range.r0, x.range.r1) }).toEqual({ col: x.col, rule: x.rule, shown: o.shown(x.range.r0, x.range.r1) })
          expect(x.order!.head.counts).toEqual(o.counts(x.range))
        }
      } finally { await o.dispose() }
    }
    // The masked string: `contains "e"` on D never matches D23 (its painted `****` has no "e"; its raw MASKED does).
    const dText = recs.find((x) => x.kind === 'filter' && x.rule === 'text' && x.col === 3)!
    expect(shownSources(dText.order!, BAND.r0, BAND.r1).has(rcOf(SORT_WITHHELD.masked)[0])).toBe(false)
    // B ascending puts B21 and B22 (withheld: ;;; and CF ;;;) among the blanks at the end.
    const bAsc = recs.find((x) => x.kind === 'sort' && x.col === 1)!
    expect(new Set(Array.from(bAsc.order!.bands[0]!.sourceAtOffset).slice(-2))).toEqual(new Set([rcOf(SORT_WITHHELD.hiddenFormat)[0], rcOf(SORT_WITHHELD.cfHidden)[0]]))
  }, 30_000)
})

// ---- the oracle: local viewerControls over the fixture parsed apart, its withheld cells deleted ----------------------

async function oracleDoc(bytes: Uint8Array, drop: (row: number, col: number) => boolean, replace: Record<string, string> = {}): Promise<XlsxDocument> {
  const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
  const sheet = doc.sheets[0]!
  const swap = new Map(Object.entries(replace).map(([ref, v]) => [rcOf(ref).join(':'), v]))
  const rows = sheet.rows.map((row) => ({ ...row, cells: row.cells.filter((c: XlsxCell) => !drop(c.row, c.col)).map((c): XlsxCell => {
    const v = swap.get(`${c.row}:${c.col}`)
    return v === undefined ? c : { ref: c.ref, row: c.row, col: c.col, value: v, styleIndex: 0, style: { numFmtId: 0 } }
  }) }))
  return { ...doc, sheets: [{ ...sheet, rows }, ...doc.sheets.slice(1)] }
}
/** `sortWithheld` as Task 5's oracle deletes it under `visible-only`. */
function withheldOracle(bytes: Uint8Array): Promise<XlsxDocument> {
  const gone = new Set([SORT_WITHHELD.covered, SORT_WITHHELD.hiddenFormat, SORT_WITHHELD.cfHidden, 'C30', 'C31', 'C32', 'C33', 'C34'].map((r) => rcOf(r).join(':')))
  return oracleDoc(bytes, (r, c) => gone.has(`${r}:${c}`) || c === SORT_GRID.zeroWidthCol || r === SORT_GRID.zeroHeightRow || r === SORT_GRID.hiddenRow,
    { [SORT_WITHHELD.masked]: SORT_WITHHELD.maskedDisplay })
}
async function oracle(doc: XlsxDocument) {
  const signal = () => new AbortController().signal
  const source = createObjectWorkbookSource(doc, 'oracle')
  const sheetId = source.sheets[0]!.sheetId
  const baseHidden = source.geometry(sheetId).baseHidden
  let view: SheetView = { sheetId, revision: 0, order: createIdentityOrder(), baseHidden, visibility: visibilityFromHidden(baseHidden), sort: null, filter: null, selection: emptyGridSelection() }
  let viewGen = 0
  const owner = (): CacheIdentity => ({ docId: 'oracle', sheetId, valueGen: source.store(sheetId).stamp.valueGen, fontGen: 0, viewGen })
  const snapshot = () => ({ owner: owner(), filter: view.filter, sort: view.sort, records: { visible: 0, total: 0, completeness: 'complete' as const } })
  const publish = (next: Pick<SheetView, 'order' | 'visibility' | 'filter' | 'sort'>) => { view = { ...view, ...next }; viewGen++; return snapshot() }
  const controls = await viewerControls({ doc, source, sheetId, snapshot, view: () => view, publish, ledger: new MainMemoryLedger() }, signal())
  const client = controls.client
  return {
    async sort(spec: SortSpec) { return (await client.sort(owner(), spec, signal())).kind },
    async filter(spec: Parameters<typeof client.filter>[1]) { return (await client.filter(owner(), spec, signal())).kind },
    async values(range: FilterRange, col: number, pick: (label: string) => boolean, blanks: boolean) {
      const port = await client.values(owner(), range, col, null, signal())
      const page = await port.read({ text: '', offset: 0 }, signal())
      const quiet = { text: '', offset: 10_000 }
      await port.change({ kind: 'all', selected: false, includeBlanks: false }, quiet, signal())
      for (const e of page.entries) if (pick(e.label)) await port.change({ kind: 'toggle', ordinal: e.ordinal, selected: true }, quiet, signal())
      await port.change({ kind: 'blanks', selected: blanks }, quiet, signal())
      const rule = await port.candidate(signal())
      port.transfer()
      return rule
    },
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
    order(r0: number, r1: number) { return Array.from({ length: r1 - r0 + 1 }, (_, i) => view.order.displayToSource(r0 + i)) },
    dispose: async () => { controls.dispose(); await source.release() },
  }
}

// ---- tiles under a view (the T2 method) ----------------------------------------------------------------------------

/** `bytes` with sheet 0's cells `refs` deleted and `revalue` cells' saved `<v>` changed (each must exist). */
function twinOf(bytes: Uint8Array, refs: readonly string[], revalue: Record<string, number> = {}): Promise<Uint8Array> {
  return patchXlsx(bytes, (zip) => patchSheetXml(zip, 0, (xml) => {
    for (const ref of refs) {
      const next = xml.replace(new RegExp(`<c r="${ref}"(?:[^>]*/>|[^>]*>.*?</c>)`), '')
      if (next === xml) throw new Error(`twin: no cell ${ref}`)
      xml = next
    }
    for (const [ref, v] of Object.entries(revalue)) {
      const next = xml.replace(new RegExp(`(<c r="${ref}"[^>]*>)<v>[^<]+</v>`), `$1<v>${v}</v>`)
      if (next === xml || !next.includes(`<v>${v}</v>`)) throw new Error(`twin: no valued cell ${ref}`)
      xml = next
    }
    return xml
  }))
}

/** Tiles of the band A21:F40 sorted by B ascending, every tile of the used extent at (zoomStep 5, dpr 1). */
async function sortedTiles(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const s = await server(bytes)
  const v = await session(s)
  try {
    const { layout } = await v.layout('0')
    const r = await v.view('sheet-sort', { sheetId: '0', viewGen: 0, range: BAND, keys: [{ col: 1, order: 'asc' }] })
    expect(r.end).toBeUndefined()
    return await v.tiles('0', r.order!.head.viewGen, 5, usedTiles(layout, 5))
  } finally { v.c.close(); await s.h.stop() }
}
const unequal = (a: Map<string, Uint8Array>, b: Map<string, Uint8Array>) => [...a.keys()].filter((k) => !b.has(k) || !sameBytes(a.get(k)!, b.get(k)!))

describe('tiles under a 4c view (T2 method)', () => {
  test('a sorted view paints no withheld cell', async () => {
    const bytes = await sortWithheld()
    // Column E's and row 12's cells, B21, B22 and C30:C34 deleted; B20 (under the box) holds another number, since an
    // empty B20 would let A20's text spill under the box (and A20 would be withheld); D23 is re-pointed at another
    // string (it keeps its `;;;"****"` style, so it still paints ****). The withheld rows sort as blanks either way.
    const eRefs = Array.from({ length: SORT_GRID.rows }, (_, i) => `E${i + 1}`)
    const row12 = Array.from({ length: SORT_GRID.cols }, (_, c) => `${String.fromCharCode(65 + c)}12`).filter((r) => r !== 'E12')
    const twin = await twinOf(bytes, [...eRefs, ...row12, 'B21', 'B22', 'C30', 'C31', 'C32', 'C33', 'C34'], { [SORT_WITHHELD.covered]: 777, [SORT_WITHHELD.masked]: 0 })
    const a = await sortedTiles(bytes), b = await sortedTiles(twin)
    expect(a.size).toBeGreaterThan(0)
    expect(unequal(a, b)).toEqual([])
    // Control: deleting a visible cell of the band (B25) changes the sorted tiles.
    const c = await sortedTiles(await twinOf(bytes, ['B25']))
    expect(unequal(a, c).length).toBeGreaterThan(0)
  }, 30_000)
})

// ---- the pivot strip on the wire -----------------------------------------------------------------------------------

/** `hiddenPivotCache` with PIVOT-SECRET and SLICER-SECRET physically removed from the cache parts. */
async function strippedPivotCache(): Promise<Uint8Array> {
  return patchXlsx(await hiddenPivotCache(), async (zip) => {
    const edit = async (path: string, f: (xml: string) => string) => {
      const xml = await zip.file(path)!.async('string'), out = f(xml)
      if (out === xml) throw new Error(`strip: ${path} not recognised`)
      zip.file(path, out)
    }
    await edit('xl/pivotCache/pivotCacheDefinition1.xml', (xml) => xml
      .replace(`<sharedItems count="5"><s v="South"/><s v="North"/><s v="West"/><s v="East"/><s v="${HIDDEN_PIVOT.pivotItem}"/></sharedItems>`, '<sharedItems count="4"><s v="South"/><s v="North"/><s v="West"/><s v="East"/></sharedItems>')
      .replace(`<sharedItems count="5"><s v="Bananas"/><s v="Dates"/><s v="Apples"/><s v="Cherries"/><s v="${HIDDEN_PIVOT.slicerItem}"/></sharedItems>`, '<sharedItems count="4"><s v="Bananas"/><s v="Dates"/><s v="Apples"/><s v="Cherries"/></sharedItems>'))
    await edit('xl/slicerCaches/slicerCache1.xml', (xml) => xml.replace('<items count="5"><i x="2" s="1"/><i x="1"/><i x="3" s="1"/><i x="0"/><i x="4"/></items>', '<items count="4"><i x="2" s="1"/><i x="1"/><i x="3" s="1"/><i x="0"/></items>'))
  })
}

/** Every frame and every tile (by `sheetId|col,row`) of a session that fetches each visible sheet's used extent. */
async function pivotTiles(bytes: Uint8Array, policy: Policy) {
  const s = await server(bytes, policy)
  const v = await session(s)
  const pngs = new Map<string, Uint8Array>()
  try {
    for (const { sheetId } of v.manifest.sheets as Array<{ sheetId: string }>) {
      const { layout } = await v.layout(sheetId)
      if (!layout.used) continue
      for (const [k, png] of await v.tiles(sheetId, 0, 5, usedTiles(layout, 5))) pngs.set(`${sheetId}|${k}`, png)
    }
    await v.cap.drain()
    return { pngs, frames: [...v.cap.texts, ...v.cap.payloads.map(latin1), ...v.cap.frames.map(latin1)] }
  } finally { v.c.close(); await s.h.stop() }
}

describe('the pivot strip on the wire (S4 §6.9)', () => {
  test('no frame carries a pivot or slicer cache item; every tile equals the stripped render', async () => {
    const bytes = await hiddenPivotCache(), stripped = await strippedPivotCache()
    const vis = await pivotTiles(bytes, {})
    expect(vis.pngs.size).toBeGreaterThan(0)
    expect(leaked(vis.frames, [HIDDEN_PIVOT.pivotItem, HIDDEN_PIVOT.slicerItem])).toEqual([])
    const visStripped = await pivotTiles(stripped, {})
    expect(unequal(vis.pngs, visStripped.pngs)).toEqual([])
    // Non-vacuity: under 'full' the slicer paints its cache's items, so its tiles differ from the stripped render.
    const full = await pivotTiles(bytes, { contentPolicy: 'full' }), fullStripped = await pivotTiles(stripped, { contentPolicy: 'full' })
    const differs = unequal(full.pngs, fullStripped.pngs)
    console.log(`pivot strip: ${vis.pngs.size} tiles; under 'full' ${differs.length} tile(s) differ from the stripped render (${differs.join(', ')})`)
    expect(differs.length).toBeGreaterThan(0)
  }, 30_000)
})

// ---- the hostile battery -------------------------------------------------------------------------------------------

/** `filterAutoHidden` with column B (`Num`) saved `width="0"`: a collapsed key column inside A1:C20. */
async function filterAutoHiddenColB(): Promise<Uint8Array> {
  return patchXlsx(await filterAutoHidden(), (zip) => patchSheetXml(zip, 0, (xml) => insertSheetElement(xml, 'cols', '<cols><col min="2" max="2" width="0" customWidth="1"/></cols>')))
}

/** Wraps the sid's `SessionRates.check` so each call is recorded by class. */
function chargesOf(h: Harness, sid: string): RateClass[] {
  const rates = h.stream.sheetRatesOf(sid) as SessionRates
  const seen: RateClass[] = []
  const check = rates.check.bind(rates)
  rates.check = (kind, now) => { seen.push(kind); return check(kind, now) }
  return seen
}

interface BatteryFixture { name: string; doc: () => Promise<Uint8Array>; range: WireRange; hiddenCol: number; textCol: number; listCol: number; otherCol: number; numCol: number; setup: { col: number; first: string } }
const BATTERY_FIXTURES: readonly BatteryFixture[] = [
  { name: 'sortGrid', doc: sortGrid, range: wire(0, SORT_GRID.rows - 1, 0, SORT_GRID.cols - 1), hiddenCol: SORT_GRID.zeroWidthCol, textCol: 3, listCol: 3, otherCol: 0, numCol: 1, setup: { col: 3, first: 'a' } },
  { name: 'filterAutoHidden', doc: filterAutoHiddenColB, range: wire(0, 19, 0, 2), hiddenCol: 1, textCol: 0, listCol: 0, otherCol: 2, numCol: 2, setup: { col: 2, first: '1' } },
]
const FORGED_LIST = '00000000-0000-4000-8000-000000000000'
const OUT_OF_GRID = 1_048_576

describe('the hostile battery (S4 §6.6, §9 4c; plan P4)', () => {
  for (const f of BATTERY_FIXTURES) {
    test(`${f.name}: each hostile op refuses with its P4 code and moves nothing`, async () => {
      const s = await server(await f.doc())
      const a = await session(s), other = await session(s), twin = await session(s)
      try {
        const text = (col: number, first: string) => [col, { kind: 'text', op: 'contains', first, wildcards: false }]
        // A non-identity starting view: a committed filter, so the counts and the probe tile are a filter's.
        const start = await a.view('sheet-filter', { sheetId: '0', viewGen: 0, range: f.range, columns: [text(f.setup.col, f.setup.first)] })
        expect(start.end).toBeUndefined()
        const gen = start.order!.head.viewGen
        const probe = async (v: Session, viewGen: number) => [...(await v.tiles('0', viewGen, 5, [{ col: 0, row: 0 }])).values()][0]!
        const snap = async () => ({ order: (await a.order('0')).parts.map(({ reqId: _r, ...p }) => p), layout: (await a.layout('0')).parts, tile: await probe(a, gen) })
        const before = await snap()
        expect(before.order[0]!.head.viewGen).toBe(gen)
        const mine = await a.list('0', gen, f.range, f.listCol), theirs = await other.list('0', 0, f.range, f.listCol)
        const otherCol = await a.list('0', gen, f.range, f.otherCol)
        expect([!!mine.first && mine.items.length > 0, !!theirs.first && !!otherCol.first]).toEqual([true, true])
        const values = (list: string) => ({ kind: 'values', list, mode: 'include', ids: [[mine.items[0]!.id, mine.items[0]!.id]], blanks: false })
        const keys65 = Array.from({ length: 65 }, () => ({ col: f.otherCol, order: 'asc' }))
        const wide = wire(f.range.r0, f.range.r1, 0, 64)
        const cols65 = Array.from({ length: 65 }, (_, c) => text(c, 'x'))
        type Hostile = { name: string; type: 'sheet-sort' | 'sheet-filter'; body: Text; code: string; row: boolean; twin: Text }
        const sort = (o: Text) => ({ range: f.range, keys: [{ col: f.otherCol, order: 'asc' }], ...o })
        const filter = (o: Text) => ({ range: f.range, columns: [text(f.textCol, 'e')], ...o })
        const cases: Hostile[] = [
          { name: 'forged list id', type: 'sheet-filter', body: filter({ columns: [[f.listCol, values(FORGED_LIST)]] }), code: 'stale-filter-values', row: true, twin: { columns: [[f.listCol, 'OWN']] } },
          { name: "another sid's id", type: 'sheet-filter', body: filter({ columns: [[f.listCol, values(theirs.first!.list)]] }), code: 'stale-filter-values', row: true, twin: { columns: [[f.listCol, 'OWN']] } },
          { name: "another column's id", type: 'sheet-filter', body: filter({ columns: [[f.listCol, values(otherCol.first!.list)]] }), code: 'stale-filter-values', row: true, twin: { columns: [[f.listCol, 'OWN']] } },
          { name: 'out-of-grid range (sort)', type: 'sheet-sort', body: sort({ range: { ...f.range, r1: OUT_OF_GRID } }), code: 'policy-refused', row: false, twin: sort({}) },
          { name: 'out-of-grid range (filter)', type: 'sheet-filter', body: filter({ range: { ...f.range, c1: 16_384 } }), code: 'policy-refused', row: false, twin: filter({}) },
          { name: 'hidden key column (sort)', type: 'sheet-sort', body: sort({ keys: [{ col: f.hiddenCol, order: 'asc' }] }), code: 'hidden-target', row: true, twin: sort({}) },
          { name: 'hidden key column (filter)', type: 'sheet-filter', body: filter({ columns: [text(f.hiddenCol, 'e')] }), code: 'hidden-target', row: true, twin: filter({}) },
          { name: 'stale viewGen (sort)', type: 'sheet-sort', body: sort({ viewGen: gen - 1 }), code: 'stale-view', row: true, twin: sort({}) },
          { name: 'stale viewGen (filter)', type: 'sheet-filter', body: filter({ viewGen: gen + 5 }), code: 'stale-view', row: true, twin: filter({}) },
          { name: 'typed rule under visible-only', type: 'sheet-filter', body: filter({ columns: [[f.numCol, { kind: 'number', op: 'gt', first: 3 }]] }), code: 'filter-condition-refused', row: true, twin: filter({ columns: [text(f.numCol, '1')] }) },
          { name: '65 rules (sort)', type: 'sheet-sort', body: sort({ keys: keys65 }), code: 'policy-refused', row: false, twin: sort({ keys: keys65.slice(0, 64) }) },
          { name: '65 rules (filter)', type: 'sheet-filter', body: filter({ range: wide, columns: cols65 }), code: 'policy-refused', row: false, twin: filter({ range: wide, columns: cols65.filter(([c]) => c !== f.hiddenCol).map(([c]) => text(c as number, '')) }) },
        ]
        const charges = chargesOf(s.h, a.sid)
        for (const k of cases) {
          const rows = s.store.sheetOps(a.sid).length
          charges.splice(0)
          const r = await a.view(k.type, { sheetId: '0', viewGen: gen, ...k.body })
          expect({ name: k.name, end: r.end }).toMatchObject({ name: k.name, end: { type: 'refused', code: k.code } })
          // P12: a refusal past the shape check has its row with the code; a shape refusal has none and is uncharged.
          const after = s.store.sheetOps(a.sid)
          if (k.row) expect({ name: k.name, n: after.length - rows, last: after.at(-1)?.outcome }).toEqual({ name: k.name, n: 1, last: k.code })
          else expect({ name: k.name, n: after.length - rows, charged: charges.includes('sheet') }).toEqual({ name: k.name, n: 0, charged: false })
          // viewGen, the layout bytes, the counts and the probe tile are as before.
          const now = await snap()
          expect({ name: k.name, order: now.order, layout: now.layout }).toEqual({ name: k.name, order: before.order, layout: before.layout })
          expect({ name: k.name, tile: sameBytes(now.tile, before.tile) }).toEqual({ name: k.name, tile: true })
        }
        // Each hostile op's valid twin commits on a session of its own (so each refusal is for the hostile part alone).
        let tg = 0
        for (const k of cases) {
          let body = k.twin
          if (JSON.stringify(body).includes('"OWN"')) {
            const own = await twin.list('0', tg, f.range, f.listCol)
            body = filter({ columns: [[f.listCol, { kind: 'values', list: own.first!.list, mode: 'include', ids: [[own.items[0]!.id, own.items[0]!.id]], blanks: false }]] })
          }
          const r = await twin.view(k.type, { sheetId: '0', viewGen: tg, ...body })
          if (k.name === '65 rules (sort)') {
            // 64 keys pass the shape check and P4 and reach the engine, which refuses them itself (P13): the local
            // engine sorts at most 8 keys, none repeated (local parity), so no 64-key sort can commit.
            expect({ name: k.name, end: r.end }).toMatchObject({ name: k.name, end: { type: 'refused', code: 'view-refused', reason: 'source-error' } })
            expect(s.store.sheetOps(twin.sid).at(-1)).toMatchObject({ kind: 'sort', outcome: 'view-refused', rules: keys65.slice(0, 64).map((x) => [x.col, 'asc']) })
            continue
          } else expect({ name: k.name, end: r.end }).toEqual({ name: k.name, end: undefined })
          if (r.order) tg = r.order.head.viewGen
          const clear = await twin.view(k.type, { sheetId: '0', viewGen: tg, range: body.range, ...(k.type === 'sheet-sort' ? { keys: [] } : { columns: [] }) })
          expect({ name: k.name, end: clear.end }).toEqual({ name: k.name, end: undefined })
          tg = clear.order!.head.viewGen
        }
        // The probe tile is the filter's, not the identity's (the snapshot compares a moved view).
        expect(sameBytes(before.tile, await probe(other, 0))).toBe(false)
      } finally { a.c.close(); other.c.close(); twin.c.close(); await s.h.stop() }
    }, 30_000)
  }
})
