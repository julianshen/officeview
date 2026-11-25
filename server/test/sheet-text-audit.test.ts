/**
 * Stage 4b Task 14 — the text wire audit and the withholding battery (S4 §9 4b; plan P2–P5).
 *
 * **Wire audit.** A `sheet-view@1` session (visible-only, `copy: 'full'`) over the whole server path is captured byte
 * for byte by a listener of its own on the socket (text frames, raw binary frames, and every binary message
 * reassembled). Per visible sheet it asks `sheet-layout`, a find for each sentinel, marker and withheld string (and a
 * visible control), a select-all copy, a unit-mode a11y over the used range in ≤ 2,000-cell windows, and a fit of
 * every used column. "Select-all" is the used range (A1 to the last used cell) in one selection: the client's
 * worksheet selection (every row and column) is refused `copy-too-large` before any cell is read, as the local engine
 * refuses it (Task 7 ruling), so it would carry no text at all. It runs for `sentinels()`, every XLSX hidden fixture,
 * and the withholding fixtures (Task 1's, `unfilledShapeOverText`, `rotatedBox`, `unresolvedAnchor`).
 *
 * **Battery.** {offViewportCovered, rotatedBox, unresolvedAnchor, cfHiddenFormat, absoluteOverB2 resized} ×
 * {find, copy, a11y}, each from three scroll histories of its own session: no tiles fetched, every tile of the used
 * extent fetched (zoomStep 0, dpr 1: about 50 tiles for the 5,000-row fixture, against about 780 at zoomStep 5), and
 * only the tiles that hold a withheld cell fetched (zoomStep 5, dpr 1). The answers are equal across the three, and
 * the withheld set is empty in each.
 *
 * Rates: the server's clock is the harness's own (`now`), advanced 100 ms per op and 61 s every 25 ops, so the
 * 60/min `'text'`, 30/min `'sheet'` and 20/s message windows admit an audit of about 70 a11y windows on one sid. The
 * token lives an hour so the advanced clock never expires it.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { Reassembler } from '../../src/stream/frame'
import { decodeSheetLayout, TILE_PX, TILE_ZOOM_STEPS, type StreamSheetLayout, type WireSelection } from '../../src/stream/sheet-wire'
import { streamSheetGeometry } from '../../src/xlsx/view/streamGeometry'
import { tileLimit } from '../src/sheet/tile'
import { augmentedFixtures } from './fixtures'
import { Client, disposePool, start, type Harness } from './harness'
import {
  absoluteOverB2, cfHiddenFormat, cfHiddenValue, CF_HIDDEN_FORMAT_VALUES, CF_HIDDEN_VALUES, hiddenFormat, mergeTail, offViewportCovered, rotatedBox,
  SENTINELS, sentinels, slidUnderAnchor, unfilledShapeOverText, unresolvedAnchor,
} from './sheet-fixtures'

afterAll(async () => { await disposePool() })

const VIEW = ['sheet-view@1']
const PANE = { x: 0, y: 0, width: 640, height: 400 }
const A11Y_CELLS = 2000
const CAPS: Record<number, [number, number]> = { 1: [256 << 10, 256 << 10], 2: [1 << 20, 8 << 20], 3: [8 << 20, 64 << 20], 4: [8 << 20, 8 << 20], 5: [8 << 20, 64 << 20], 6: [8 << 20, 32 << 20], 7: [8 << 20, 64 << 20], 8: [8 << 20, 64 << 20], 9: [4 << 20, 4 << 20] }
const latin1 = (b: Uint8Array) => Buffer.from(b.buffer, b.byteOffset, b.byteLength).toString('latin1')

// ---- the server, its clock, the capture --------------------------------------------------------------------------

/** One server over `doc` whose clock the test advances (see the header), so rate windows never refuse an audit op. */
async function server(doc: Uint8Array): Promise<{ h: Harness; tick(): void }> {
  let skew = 0, ops = 0
  const h = await start({ doc, policy: { copy: 'full' }, now: () => Date.now() + skew, tokenTtlSec: 3600 })
  return { h, tick: () => { skew += 100; if (++ops % 25 === 0) skew += 61_000 } }
}

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

// ---- the ops -------------------------------------------------------------------------------------------------------

type Text = Record<string, any>
/** Sends `msg` (with the session token) and collects its reqId: every `partial` message until one has `last: true`
 * (`done`), or the first other message (`end`). */
async function collect(c: Client, msg: Text, partial: string | undefined): Promise<{ parts: Text[]; end: Text }> {
  c.send({ token: c.manifest!.session.token, ...msg })
  const parts: Text[] = []
  for (;;) {
    const m = await c.next((m) => 'text' in m && m.text.reqId === msg.reqId, 30_000) as { text: Text }
    if (m.text.type !== partial) return { parts, end: m.text }
    parts.push(m.text)
    if (m.text.last === true) return { parts, end: m.text }
  }
}

interface View { c: Client; tick(): void; rq: number }
async function layoutOf(v: View, sheetId: string): Promise<StreamSheetLayout> {
  v.tick()
  const lr = await v.c.sheetLayout(++v.rq, sheetId)
  expect([sheetId, lr.end]).toEqual([sheetId, undefined])
  return decodeSheetLayout(lr.parts.map((p) => p.text))
}
async function find(v: View, query: string): Promise<{ hits: number[][]; parts: Text[]; end: Text }> {
  v.tick()
  const r = await collect(v.c, { reqId: ++v.rq, type: 'sheet-find', query }, 'sheet-find-result')
  expect([query, r.end.type]).toEqual([query, 'sheet-find-done'])
  return { ...r, hits: r.parts.flatMap((p) => p.hits as number[][]) }
}
async function copy(v: View, sheetId: string, viewGen: number, selection: WireSelection): Promise<string> {
  v.tick()
  const r = await collect(v.c, { reqId: ++v.rq, type: 'sheet-copy', sheetId, viewGen, selection }, 'copy-result')
  expect([sheetId, r.end.type, r.end.last]).toEqual([sheetId, 'copy-result', true])
  return r.parts.map((p) => p.text as string).join('')
}
/** A unit-mode a11y over rows 0..r1 × cols 0..c1, in windows of at most 2,000 cells; every served cell. */
async function a11y(v: View, sheetId: string, viewGen: number, r1: number, c1: number): Promise<Array<[number, number, string]>> {
  const out: Array<[number, number, string]> = []
  const colsPer = Math.min(c1 + 1, A11Y_CELLS)
  for (let c0 = 0; c0 <= c1; c0 += colsPer) {
    const cEnd = Math.min(c1, c0 + colsPer - 1), rowsPer = Math.floor(A11Y_CELLS / (cEnd - c0 + 1))
    for (let r0 = 0; r0 <= r1; r0 += rowsPer) {
      v.tick()
      const rows: Array<[number, number]> = [[r0, Math.min(r1, r0 + rowsPer - 1)]]
      const r = await collect(v.c, { reqId: ++v.rq, type: 'sheet-a11y', sheetId, viewGen, rows, cols: [c0, cEnd] }, 'sheet-a11y-result')
      expect([sheetId, r0, r.end.type, r.end.last]).toEqual([sheetId, r0, 'sheet-a11y-result', true])
      for (const p of r.parts) out.push(...(p.cells as Array<[number, number, string]>))
    }
  }
  return out
}
async function fit(v: View, sheetId: string, viewGen: number, index: number): Promise<number> {
  v.tick()
  const r = await collect(v.c, { reqId: ++v.rq, type: 'sheet-fit', sheetId, viewGen, axis: 'col', index, pane: PANE }, undefined)
  expect([sheetId, index, r.end.type]).toEqual([sheetId, index, 'sheet-fit-result'])
  return r.end.size as number
}
async function resize(v: View, sheetId: string, viewGen: number, r: Resize): Promise<number> {
  v.tick()
  const a = await collect(v.c, { reqId: ++v.rq, type: 'sheet-resize', sheetId, viewGen, ...r }, undefined)
  expect(a.end).toMatchObject({ type: 'sheet-view', sheetId, viewGen: viewGen + 1 })
  return a.end.viewGen as number
}
type Resize = { axis: 'row' | 'col'; index: number; size: number }
/** The used range from A1 as one selection (the audit's select-all; see the header). */
const usedSelection = (r1: number, c1: number): WireSelection => ({ active: [0, 0], anchor: [0, 0], ranges: [{ rows: [[0, r1]], cols: [[0, c1]] }] })

async function open(url: string, tick: () => void): Promise<View> {
  const c = new Client(url)
  const m = await c.hello({ features: VIEW })
  await c.ready()
  expect([m.sheetMode, m.sheetView, m.caps.a11y]).toEqual(['tiles', 1, 'unit'])
  return { c, tick, rq: 100 }
}

// ---- the audit -----------------------------------------------------------------------------------------------------

/** What one audited session sent, with the text answers in their own lists (copies joined per reqId). */
interface Audited {
  name: string; manifest: Text; cap: Capture
  /** Capture lengths (texts, frames, payloads) when the last phase began. */
  lastPhase: [number, number, number]
  finds: Array<{ phase: number; query: string; hits: number[][]; parts: Text[]; end: Text }>
  copies: Array<{ phase: number; sheetId: string; text: string }>
  cells: Array<{ phase: number; sheetId: string; cells: Array<[number, number, string]> }>
  fits: number
}
/** Runs the audit's ops over every visible sheet, once per phase; phase k > 0 first applies `resizes[k − 1]` to sheet 0. */
async function audited(name: string, doc: Uint8Array, queries: readonly string[], resizes: readonly Resize[] = []): Promise<Audited> {
  const { h, tick } = await server(doc)
  const v = await open(h.url, tick)
  const cap = capture(v.c)
  const out: Audited = { name, manifest: v.c.manifest as never, cap, lastPhase: [0, 0, 0], finds: [], copies: [], cells: [], fits: 0 }
  try {
    const viewGen = new Map<string, number>()
    for (let phase = 0; phase <= resizes.length; phase++) {
      if (phase === resizes.length) out.lastPhase = [cap.texts.length, cap.frames.length, cap.payloads.length]
      if (phase > 0) viewGen.set('0', await resize(v, '0', viewGen.get('0') ?? 0, resizes[phase - 1]!))
      for (const query of queries) out.finds.push({ phase, query, ...await find(v, query) })
      for (const { sheetId } of v.c.manifest!.sheets!) {
        const layout = await layoutOf(v, sheetId)
        const gen = viewGen.get(sheetId) ?? 0
        expect([name, sheetId, layout.viewGen]).toEqual([name, sheetId, gen])
        if (!layout.used) continue
        const { r1, c1 } = layout.used
        out.copies.push({ phase, sheetId, text: await copy(v, sheetId, gen, usedSelection(r1, c1)) })
        out.cells.push({ phase, sheetId, cells: await a11y(v, sheetId, gen, r1, c1) })
        for (let col = 0; col <= c1; col++) { await fit(v, sheetId, gen, col); out.fits++ }
      }
    }
    await cap.drain()
    return out
  } finally { v.c.close(); await h.stop() }
}

/** Every decoded text a session received, from `from` on: text frames, reassembled payloads and raw frames (latin1),
 * and each copy joined across its chunks (a needle may straddle a chunk). */
function decoded(s: Audited, from: [number, number, number] = [0, 0, 0], phase = 0): string[] {
  return [
    ...s.cap.texts.slice(from[0]), ...s.cap.payloads.slice(from[2]).map(latin1), ...s.cap.frames.slice(from[1]).map(latin1),
    ...s.copies.filter((x) => x.phase >= phase).map((x) => x.text),
  ]
}
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
/** The text fields of every text answer: each TSV field of each copy, and each a11y cell's text. */
const answerFields = (s: Audited): string[] => [...s.copies.flatMap((x) => x.text.split('\n').flatMap((l) => l.split('\t'))), ...s.cells.flatMap((x) => x.cells.map(([, , t]) => t))]

/** The withheld strings no text answer may carry (the brief's list, plus the other strings `sentinels()` hides). */
const WITHHELD = [
  SENTINELS.formula, SENTINELS.definedName, SENTINELS.number, SENTINELS.formulaText, SENTINELS.hiddenSheet, SENTINELS.hiddenCell, SENTINELS.comment, SENTINELS.drawingName,
  'COVERED-B2', 'ROT-1-1', 'ROT-2-1', 'ROT-3-1', 'OFFVIEW-Z5000', 'TAIL-B1', 'HIDDENFMT', 'ABS-B2', 'UNRES-',
]
/** Visible strings that must flow (liveness): each fixture's control text. */
const AUDIT_FIXTURES: ReadonlyArray<{ name: string; doc: () => Promise<Uint8Array>; controls: readonly string[]; resizes?: readonly Resize[] }> = [
  { name: 'sentinels', doc: sentinels, controls: [SENTINELS.visibleText] },
  { name: 'offViewportCovered', doc: offViewportCovered, controls: ['OFFVIEW-A1'] },
  { name: 'slidUnderAnchor', doc: slidUnderAnchor, controls: [], resizes: [{ axis: 'col', index: 1, size: 8 }] },
  { name: 'mergeTail', doc: mergeTail, controls: ['MERGE-A1'] },
  { name: 'hiddenFormat', doc: hiddenFormat, controls: ['SHOWN-A1'] },
  { name: 'cfHiddenValue', doc: cfHiddenValue, controls: ['CFTEXT-B4'] },
  { name: 'cfHiddenFormat', doc: cfHiddenFormat, controls: [] },
  { name: 'absoluteOverB2', doc: absoluteOverB2, controls: [], resizes: [{ axis: 'col', index: 0, size: 300 }] },
  { name: 'unfilledShapeOverText', doc: unfilledShapeOverText, controls: ['CELL-2-1', 'CELL-1-2'] },
  { name: 'rotatedBox', doc: rotatedBox, controls: ['ROT-4-1', 'ROT-1-2'] },
  { name: 'unresolvedAnchor', doc: unresolvedAnchor, controls: [] },
]

let runs: Promise<Audited[]> | undefined
/** Every audited session, run once (memoised across the audit's tests). */
function allAudits(): Promise<Audited[]> {
  return runs ??= (async () => {
    const out: Audited[] = []
    const hidden = (await augmentedFixtures()).filter((f) => f.kind === 'xlsx')
    expect(hidden.map((f) => f.id).sort()).toEqual(['xlsx-comment', 'xlsx-hidden-cols', 'xlsx-hidden-rows', 'xlsx-hidden-sheet', 'xlsx-veryhidden-sheet'])
    const markers = hidden.map((f) => f.marker)
    const queries = (controls: readonly string[]) => [...WITHHELD, ...markers, 'MARKER-', 'H', '7', '9', ...controls]
    for (const f of hidden) out.push(await audited(f.id, f.withHidden, queries([f.control])))
    for (const f of AUDIT_FIXTURES) out.push(await audited(f.name, await f.doc(), queries([...f.controls, 'SLID-E1']), f.resizes))
    return out
  })()
}

describe('4b text wire audit (visible-only, copy full)', () => {
  test('no sentinel or withheld text in any text answer', async () => {
    const all = await allAudits()
    const hidden = (await augmentedFixtures()).filter((f) => f.kind === 'xlsx')
    const markers = hidden.map((f) => f.marker)
    // Non-vacuity: every op answered, and visible text flows through find, copy and a11y.
    for (const s of all) {
      expect([s.name, s.copies.length > 0, s.cells.length > 0, s.fits > 0]).toEqual([s.name, true, true, true])
      const controls = [...(hidden.find((f) => f.id === s.name) ? [hidden.find((f) => f.id === s.name)!.control] : []), ...(AUDIT_FIXTURES.find((f) => f.name === s.name)?.controls ?? [])]
      for (const ctl of controls) {
        expect([s.name, ctl, s.copies.some((x) => x.text.includes(ctl)), s.cells.some((x) => x.cells.some(([, , t]) => t === ctl))]).toEqual([s.name, ctl, true, true])
        expect([s.name, ctl, s.finds.find((x) => x.query === ctl)!.hits.length > 0]).toEqual([s.name, ctl, true])
      }
    }
    const byName = (n: string) => all.find((s) => s.name === n)!
    // The sentinel workbook's visible cells: A1 paints its saved 10, B1 paints 1235 (number format 0).
    expect(byName('sentinels').copies.map((x) => x.text.split('\n')[0]!.split('\t').slice(0, 2))).toEqual([['10', '1235']])
    // The needles: no decoded message of any session holds one.
    for (const s of all) {
      expect([s.name, leaked(decoded(s), [...WITHHELD, ...markers, 'MARKER-'])]).toEqual([s.name, []])
      // A hidden sheet's name: the manifest lists the visible sheet only, and no transmitted string is it.
      expect([s.name, stringLeaves(s).filter((x) => x === 'H' || x === SENTINELS.hiddenSheet)]).toEqual([s.name, []])
    }
    for (const n of [...hidden.map((f) => f.id), 'sentinels']) expect([n, byName(n).manifest.sheets]).toEqual([n, [{ sheetId: '0', name: 'S' }]])
    // SLID-E1 answers before the shrink (the control) and nowhere after it.
    const slid = byName('slidUnderAnchor')
    expect(slid.copies.filter((x) => x.phase === 0).map((x) => x.text.includes('SLID-E1'))).toEqual([true])
    expect(slid.finds.find((x) => x.phase === 0 && x.query === 'SLID-E1')!.hits).toEqual([[0, 0, 4]])
    expect(leaked(decoded(slid, slid.lastPhase, 1), ['SLID-E1'])).toEqual([])
    expect(slid.finds.find((x) => x.phase === 1 && x.query === 'SLID-E1')!.hits).toEqual([])
    expect(slid.cells.filter((x) => x.phase === 1).flatMap((x) => x.cells)).toEqual([])
    // cfHiddenFormat's 7 and 9 (and cfHiddenValue's bar-only values) are in no text field and no find hit; B1's 4 is.
    for (const [n, hiddenValues, shown] of [['cfHiddenFormat', CF_HIDDEN_FORMAT_VALUES.slice(1), ['4']], ['cfHiddenValue', CF_HIDDEN_VALUES, ['CFTEXT-B4']]] as const) {
      const s = byName(n), fields = answerFields(s)
      expect([n, hiddenValues.map(String).filter((h) => fields.some((f) => f.includes(h)))]).toEqual([n, []])
      expect([n, shown.every((x) => fields.includes(x))]).toEqual([n, true])
    }
    const cf = byName('cfHiddenFormat')
    expect(['7', '9'].map((q) => cf.finds.find((x) => x.query === q)!.hits)).toEqual([[], []])
    expect(cf.copies.map((x) => x.text)).toEqual(['\t4\n\t\n\t'])
  }, 30_000)

  test('find answers carry positions only', async () => {
    const all = await allAudits()
    let results = 0, hits = 0
    for (const s of all) for (const f of s.finds) {
      for (const p of f.parts) {
        results++
        expect([s.name, f.query, Object.keys(p).sort()]).toEqual([s.name, f.query, ['hits', 'reqId', 'type']])
        expect(p.type).toBe('sheet-find-result')
        for (const hit of p.hits as unknown[]) {
          hits++
          expect([s.name, f.query, Array.isArray(hit) && (hit.length === 3 || hit.length === 4) && hit.every((x) => Number.isSafeInteger(x) && x >= 0)]).toEqual([s.name, f.query, true])
        }
      }
      expect([s.name, f.query, Object.keys(f.end).sort()]).toEqual([s.name, f.query, ['reqId', 'total', 'truncated', 'type']])
    }
    // Non-vacuity: the controls hit.
    expect(results).toBeGreaterThan(0)
    expect(hits).toBeGreaterThan(0)
  }, 30_000)
})

// ---- the battery ---------------------------------------------------------------------------------------------------

type History = 'none' | 'all' | 'cell'
interface Case { name: string; doc: () => Promise<Uint8Array>; withheld: ReadonlyArray<readonly [number, number, string]>; controls: ReadonlyArray<readonly [number, number, string]>; resize?: Resize }
const BATTERY: readonly Case[] = [
  { name: 'offViewportCovered', doc: offViewportCovered, withheld: [[4999, 25, 'OFFVIEW-Z5000']], controls: [[0, 0, 'OFFVIEW-A1']] },
  { name: 'rotatedBox', doc: rotatedBox, withheld: [[0, 0, 'ROT-1-1'], [1, 0, 'ROT-2-1'], [2, 0, 'ROT-3-1']], controls: [[3, 0, 'ROT-4-1'], [0, 1, 'ROT-1-2']] },
  { name: 'unresolvedAnchor', doc: unresolvedAnchor, withheld: [[0, 0, 'UNRES-1-1'], [0, 1, 'UNRES-1-2'], [1, 0, 'UNRES-2-1'], [1, 1, 'UNRES-2-2']], controls: [] },
  { name: 'cfHiddenFormat', doc: cfHiddenFormat, withheld: [[1, 1, '7'], [2, 1, '9']], controls: [[0, 1, '4']] },
  { name: 'absoluteOverB2 resized', doc: absoluteOverB2, withheld: [[1, 1, 'ABS-B2']], controls: [], resize: { axis: 'col', index: 0, size: 300 } },
]

/** Fetches the history's tiles of sheet 0 (all must arrive). */
async function scroll(v: View, history: History, layout: StreamSheetLayout, cells: ReadonlyArray<readonly [number, number, string]>): Promise<number> {
  if (history === 'none') return 0
  const mt = streamSheetGeometry(layout).metrics
  const zoomStep = history === 'all' ? 0 : 5, s = TILE_ZOOM_STEPS[zoomStep]!
  let coords: Array<{ col: number; row: number }>
  if (history === 'all') {
    const u = layout.used!
    const cols = tileLimit(mt.cols.start(u.c1) + mt.cols.size(u.c1), s, 1), rows = tileLimit(mt.rows.start(u.r1) + mt.rows.size(u.r1), s, 1)
    coords = Array.from({ length: cols * rows }, (_, i) => ({ col: i % cols, row: Math.floor(i / cols) }))
  } else {
    const seen = new Map<string, { col: number; row: number }>()
    for (const [r, c] of cells) {
      const x0 = mt.cols.start(c), x1 = x0 + mt.cols.size(c), y0 = mt.rows.start(r), y1 = y0 + mt.rows.size(r)
      for (let col = Math.floor(x0 * s / TILE_PX); col <= Math.floor((x1 * s - 1e-9) / TILE_PX); col++) {
        for (let row = Math.floor(y0 * s / TILE_PX); row <= Math.floor((y1 * s - 1e-9) / TILE_PX); row++) seen.set(`${col},${row}`, { col, row })
      }
    }
    coords = [...seen.values()]
  }
  let got = 0
  for (let i = 0; i < coords.length; i += 64) {
    v.tick()
    const r = await v.c.tiles(++v.rq, { sheetId: '0', viewGen: layout.viewGen, zoomStep, dpr: 1, palette: 'light', tiles: coords.slice(i, i + 64) }, { ms: 30_000 })
    expect(r.end).toEqual({ reqId: v.rq, type: 'tiles-done', refused: [] })
    got += r.tiles.length
  }
  expect(got).toBe(coords.length)
  return got
}

describe('4b withholding battery (S4 §9 4b)', () => {
  test('find, copy and a11y withhold the set from every scroll history', async () => {
    const report: string[] = []
    for (const k of BATTERY) {
      const { h, tick } = await server(await k.doc())
      try {
        const answers: Array<{ history: History; finds: number[][][]; copy: string; cells: Array<[number, number, string]> }> = []
        for (const history of ['none', 'all', 'cell'] as const) {
          const v = await open(h.url, tick)
          try {
            let layout = await layoutOf(v, '0')
            if (k.resize) { await resize(v, '0', 0, k.resize); layout = await layoutOf(v, '0') }
            const tiles = await scroll(v, history, layout, k.withheld)
            if (history !== 'none') expect([k.name, history, tiles > 0]).toEqual([k.name, history, true])
            const { r1, c1 } = layout.used!
            const queries = [...new Set([...k.withheld, ...k.controls].map(([, , t]) => t))]
            const finds: number[][][] = []
            for (const q of queries) finds.push((await find(v, q)).hits)
            answers.push({ history, finds, copy: await copy(v, '0', layout.viewGen, usedSelection(r1, c1)), cells: await a11y(v, '0', layout.viewGen, r1, c1) })
            report.push(`${k.name}/${history}: ${tiles} tiles`)
          } finally { v.c.close() }
        }
        for (const a of answers) {
          const tag = [k.name, a.history]
          const queries = [...new Set([...k.withheld, ...k.controls].map(([, , t]) => t))]
          const lines = a.copy.split('\n').map((l) => l.split('\t'))
          for (const [r, c, t] of k.withheld) {
            // find: no hit names a withheld cell, nor any cell for a withheld text; copy: its field is empty; a11y: unlisted.
            expect([...tag, t, a.finds.flat().filter(([s, hr, hc]) => s === 0 && hr === r && hc === c)]).toEqual([...tag, t, []])
            expect([...tag, t, a.finds[queries.indexOf(t)]]).toEqual([...tag, t, []])
            expect([...tag, t, lines[r]?.[c]]).toEqual([...tag, t, ''])
            expect([...tag, t, a.cells.filter(([ar, ac]) => ar === r && ac === c)]).toEqual([...tag, t, []])
          }
          for (const [r, c, t] of k.controls) {
            expect([...tag, t, lines[r]?.[c], a.cells.some(([ar, ac, at]) => ar === r && ac === c && at === t)]).toEqual([...tag, t, t, true])
            expect([...tag, t, a.finds[queries.indexOf(t)]!.some(([s, hr, hc]) => s === 0 && hr === r && hc === c)]).toEqual([...tag, t, true])
          }
          // The copy covers every withheld position (so an empty field is a withheld one, not a short TSV).
          for (const [r, c] of k.withheld) expect([...tag, lines.length > r && lines[r]!.length > c]).toEqual([...tag, true])
          if (!k.controls.length) expect([...tag, a.copy.replace(/[\t\n]/g, ''), a.cells]).toEqual([...tag, '', []])
        }
        // Same answers from every history.
        expect([k.name, answers[1]!.finds, answers[1]!.copy, answers[1]!.cells]).toEqual([k.name, answers[0]!.finds, answers[0]!.copy, answers[0]!.cells])
        expect([k.name, answers[2]!.finds, answers[2]!.copy, answers[2]!.cells]).toEqual([k.name, answers[0]!.finds, answers[0]!.copy, answers[0]!.cells])
      } finally { await h.stop() }
    }
    console.log(`withholding battery: ${report.join('; ')}`)
  }, 30_000)
})
