/**
 * Stage 4b Task 7 — `sheet-copy` (plan P2, P7, P8, P12, P13): local `serializeSelection` over the server text view, so
 * the TSV equals the local copy over the fixture with its withheld cells physically deleted (a withheld cell is an empty
 * field), under the Stage 3 copy caps and session quota, with one `copies` audit row per counted copy.
 */
import { Database } from 'bun:sqlite'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, describe, expect, test } from 'bun:test'
import { serializeSelection } from '../../src/xlsx/view/copy'
import { createObjectWorkbookSource } from '../../src/xlsx/view/source'
import { createIdentityOrder } from '../../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../../src/xlsx/view/visibility'
import { buildSheetGeometry } from '../../src/xlsx/view/viewport'
import { prepareCfTextScope } from '../../src/xlsx/cf/text'
import { cellDisplayText } from '../../src/xlsx/render'
import type { FormatBudget } from '../../src/xlsx/view/formatWorkspace'
import type { GridSelection } from '../../src/xlsx/view/contracts'
import type { WireSelection } from '../../src/stream/sheet-wire'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../../src/xlsx/types'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import type { AuthDecision } from '../src/hooks'
import { DocumentHost } from '../src/render'
import { Client, disposePool, start } from './harness'
import {
  bigText, BIG_TEXT, cfHiddenFormat, cfNumFmt, mergeTail, millionRows, offViewportCovered, rotatedBox, textGrid, TEXT_GRID,
  unfilledShapeOverText, unresolvedAnchor,
} from './sheet-fixtures'

const hosts: DocumentHost[] = []
afterAll(async () => {
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
  await disposePool()
})

const VIEW = ['sheet-view@1']
const COPY_LIMIT = 4 * 1024 * 1024
type Policy = AuthDecision['policy']
let rq = 700

/** The fixture's projected saved-values workbook (what the server's snapshots are built from). */
async function projectedOf(bytes: Uint8Array, policy: 'visible-only' | 'full' = 'visible-only'): Promise<XlsxDocument> {
  const host = new DocumentHost(); hosts.push(host)
  const o = await host.open('ref', bytes, policy, { sheetMode: 'tiles' })
  return (host as unknown as { docs: Map<string, { savedParse: XlsxDocument }> }).docs.get(o.docKey)!.savedParse
}
/** `doc` with every cell of sheet `sheetIndex` that `drop` names physically deleted. */
function without(doc: XlsxDocument, sheetIndex: number, drop: (cell: XlsxCell) => boolean): XlsxDocument {
  const del = (sheet: XlsxSheet): XlsxSheet => ({ ...sheet, rows: sheet.rows.map((row) => ({ ...row, cells: row.cells.filter((c) => !drop(c)) })) })
  return { ...doc, sheets: doc.sheets.map((s, i) => i === sheetIndex ? del(s) : s) }
}

/** A wire selection of `[r0, r1, c0, c1]` boxes (each one range of one row and one column interval). */
const boxes = (...b: Array<[number, number, number, number]>): WireSelection => ({ active: null, anchor: null, ranges: b.map(([r0, r1, c0, c1]) => ({ rows: [[r0, r1]], cols: [[c0, c1]] })) })
const toGrid = (s: WireSelection): GridSelection => ({
  active: s.active ? { row: s.active[0], col: s.active[1] } : null,
  anchor: s.anchor ? { row: s.anchor[0], col: s.anchor[1] } : null,
  ranges: s.ranges.map((r) => ({ rows: { kind: 'intervals' as const, intervals: r.rows.map(([r0, r1]) => ({ r0, r1 })) }, cols: r.cols.map(([c0, c1]) => ({ c0, c1 })) })),
})

const freeBudget: FormatBudget = { reserve() {}, release() {} }
/**
 * The local reference: `serializeSelection` over an object source of `doc`, with the inputs the resident viewer builds
 * (`InteractiveSheetDoc.tsx` ~2611–2659), written out here independently of the server's `textInputs`: the identity
 * order, the authored row visibility, columns hidden where the geometry's size is 0, the region-resolved style through
 * the CF scope, and the resident CF preparation (`prepareCfTextScope`) whenever the sheet has CF rules.
 */
async function localCopy(doc: XlsxDocument, sheetIndex: number, selection: WireSelection, o: { scope?: boolean } = {}) {
  const source = createObjectWorkbookSource(doc, 'local')
  const sheetId = source.sheets[sheetIndex]!.sheetId
  const sheet = doc.sheets[sheetIndex]!
  const geometry = buildSheetGeometry(sheet)
  if (geometry.kind !== 'ready') throw new Error('fixture geometry')
  const cols = geometry.metrics.cols
  const hidden: Array<{ r0: number; r1: number }> = []
  let at: number | null = null
  for (let col = 0; col < cols.logicalCount; col++) {
    if (cols.size(col) === 0) at ??= col
    else if (at !== null) { hidden.push({ r0: at, r1: col - 1 }); at = null }
  }
  if (at !== null) hidden.push({ r0: at, r1: cols.logicalCount - 1 })
  const store = source.store(sheetId)
  const identity = { ...store.stamp, fontGen: 0, viewGen: 0 }
  const hasCf = !!sheet.conditionalFormatting?.rules.length && o.scope !== false
  return serializeSelection({
    store, owner: identity, currentOwner: () => identity, allowed: () => true, signal: new AbortController().signal,
    selection: toGrid(selection), order: createIdentityOrder(),
    visibility: visibilityFromHidden(createIntervalSet(source.geometry(sheetId).baseHidden.intervals.map(({ r0, r1 }) => ({ r0, r1 })))),
    columns: visibilityFromHidden(createIntervalSet(hidden), cols.logicalCount),
    format: (cell, allocation, scope) => {
      const styled = geometry.cellsOf(cell.row)?.find((c) => c.col === cell.col)
      const base = sheet.regionStyles?.styleAt(cell.row, cell.col, styled?.style, 'screen') ?? styled?.style
      return cellDisplayText(sheet, { value: cell.value, style: scope?.effectiveStyle(cell.row, cell.col, base) ?? base }, allocation)
    },
    ...(hasCf ? { prepareScope: (rects: Parameters<typeof prepareCfTextScope>[1]['rects'], signal: AbortSignal) => prepareCfTextScope(sheet, { rects, signal, locale: 'en-US', budget: freeBudget }) } : {}),
  })
}
async function localText(doc: XlsxDocument, sheetIndex: number, selection: WireSelection, o: { scope?: boolean } = {}): Promise<string> {
  const r = await localCopy(doc, sheetIndex, selection, o)
  if (r.kind !== 'ready') throw new Error(`local copy ${r.kind}${'code' in r ? ` ${r.code}` : ''}`)
  return r.text
}
const shape = (t: string) => ({ lines: t.split('\n').length, tabs: t.split('\t').length - 1 })

async function open(url: string, extra: object = {}) {
  const c = new Client(url)
  const m = await c.hello({ features: VIEW, ...extra })
  await c.ready()
  return { c, m }
}
async function session(o: Parameters<typeof start>[0] = {}) {
  const store = AuditStore.open(':memory:')
  const h = await start({ audit: new AuditTable({ store }), ...o })
  const { c, m } = await open(h.url)
  return { h, c, m, store, stop: async () => { c.close(); await h.stop() } }
}

interface CopyAnswer { parts: Array<{ text: Record<string, any>; bytes: number }>; end: Record<string, any>; text?: string }
/** Sends a `sheet-copy` and collects its reqId's `copy-result` chunks until `last: true`, or the one other reply. */
async function copy(c: Client, body: Record<string, unknown>, reqId = ++rq): Promise<CopyAnswer> {
  c.send({ token: c.manifest!.session.token, reqId, type: 'sheet-copy', viewGen: 0, ...body })
  const parts: CopyAnswer['parts'] = []
  for (;;) {
    const m = await c.next((m) => 'text' in m && m.text.reqId === reqId, 60_000) as { text: Record<string, any>; bytes: number }
    if (m.text.type !== 'copy-result') return { parts, end: m.text }
    parts.push(m)
    if (m.text.last === true) return { parts, end: m.text, text: parts.map((p) => p.text.text as string).join('') }
  }
}
const copyText = async (c: Client, sheetId: string, selection: WireSelection): Promise<string> => {
  const a = await copy(c, { sheetId, selection })
  if (a.text === undefined) throw new Error(`copy refused: ${JSON.stringify(a.end)}`)
  return a.text
}

describe('sheet-copy (P8)', () => {
  test('copy TSV equals local on the deleted fixture', async () => {
    const bytes = await textGrid()
    const saved = await projectedOf(bytes)
    const [cr, cc] = TEXT_GRID.covered
    // Withheld in textGrid: the covered D20 and the tails of the merge A5:C6, on both grid sheets.
    const withheld = (c: XlsxCell) => (c.row === cr && c.col === cc) || (c.row >= 4 && c.row <= 5 && c.col <= 2 && !(c.row === 4 && c.col === 0))
    const deleted = without(without(saved, 0, withheld), 1, withheld)
    // Three ranges: A1:C8 (hidden row 7, the merge), B18:E22 (the covered D20), and E1:H3 + E29:H30 (hidden column F).
    const three: WireSelection = {
      active: [0, 0], anchor: [0, 0],
      ranges: [
        { rows: [[0, 7]], cols: [[0, 2]] },
        { rows: [[17, 21]], cols: [[1, 4]] },
        { rows: [[0, 2], [28, 29]], cols: [[4, 7]] },
      ],
    }
    const local = await localText(deleted, 0, three)
    // The comparison can fail: the undeleted copy carries NEEDLE-D20 and the merge tails.
    const authored = await localText(saved, 0, three)
    expect(authored).toContain('NEEDLE-D20')
    expect(authored).not.toBe(local)
    // Hidden row 7 and hidden column F are skipped locally.
    expect(local).not.toContain('HIDDEN')
    expect(local).not.toContain('NEEDLE-D20')

    const { c, stop } = await session({ doc: bytes, policy: { copy: 'full' } })
    try {
      expect(await copyText(c, '0', three)).toBe(local)
      expect(await copyText(c, '1', three)).toBe(await localText(deleted, 1, three))
      // Sheet Plain has nothing withheld: the server's copy is the authored local copy.
      const plain = boxes([0, 4, 0, 2])
      expect(await copyText(c, '2', plain)).toBe(await localText(saved, 2, plain))
      // Select-all (the client's worksheet selection, every row and column): the local engine refuses it before
      // reading a cell (its lower bound passes the clipboard cap), so the server answers copy-too-large, never text.
      const all: WireSelection = { active: [0, 0], anchor: [0, 0], ranges: [{ rows: [[0, 1_048_575]], cols: [[0, 16_383]] }] }
      expect(await localCopy(deleted, 0, all)).toMatchObject({ kind: 'refused', code: 'resource-limit' })
      const a = await copy(c, { sheetId: '0', selection: all })
      expect(a.parts).toEqual([])
      expect(a.end).toEqual({ reqId: a.end.reqId, type: 'refused', code: 'copy-too-large', limit: COPY_LIMIT })
    } finally { await stop() }
  }, 30_000)

  test('withheld cells are empty fields', async () => {
    const check = async (bytes: Uint8Array, selection: WireSelection, drop: (c: XlsxCell) => boolean) => {
      const saved = await projectedOf(bytes)
      const local = await localText(without(saved, 0, drop), 0, selection)
      // The comparison can fail: the authored copy differs.
      expect(await localText(saved, 0, selection)).not.toBe(local)
      const { c, stop } = await session({ doc: bytes, policy: { copy: 'full' } })
      try {
        const text = await copyText(c, '0', selection)
        expect(text).toBe(local)
        expect(shape(text)).toEqual(shape(await localText(saved, 0, selection)))
        return text
      } finally { await stop() }
    }
    const shapeText = await check(await unfilledShapeOverText(), boxes([0, 2, 1, 1]), (c) => c.row === 1 && c.col === 1)
    expect(shapeText.split('\n')).toEqual(['CELL-1-2', '', 'CELL-3-2'])

    const rot = await check(await rotatedBox(), boxes([0, 5, 0, 2]), (c) => c.col === 0 && c.row <= 2)
    expect(rot.split('\n').slice(0, 3).map((l) => l.split('\t')[0])).toEqual(['', '', ''])
    expect(rot.split('\n')[3]!.split('\t')[0]).toBe('ROT-4-1')

    const off = await check(await offViewportCovered(), boxes([4998, 5000, 25, 25]), (c) => c.row === 4999 && c.col === 25)
    expect(off).toBe('\n\n')

    const merge = await check(await mergeTail(), boxes([0, 0, 0, 2]), (c) => c.row === 0 && c.col === 1)
    expect(merge).toBe('MERGE-A1\t\t')

    const unres = await check(await unresolvedAnchor(), boxes([0, 1, 0, 1]), () => true)
    expect(unres).toBe('\t\n\t')
  }, 30_000)

  test('caps refuse at n+1 and allow at n, per copy', async () => {
    const { c, m, store, stop } = await session({ doc: await textGrid(), policy: { copy: { maxParagraphs: 3 } } })
    try {
      // Plain A1:C1: three text cells.
      expect(await copyText(c, '2', boxes([0, 0, 0, 2]))).toBe('P-1-1\tP-1-2\tP-1-3')
      // A1:A4: four, refused with nothing sent.
      const four = await copy(c, { sheetId: '2', selection: boxes([0, 3, 0, 0]) })
      expect(four.parts).toEqual([])
      expect(four.end).toEqual({ reqId: four.end.reqId, type: 'refused', code: 'copy-too-large', limit: 3, requested: 4 })
      expect(store.copies(m.session.sid).map((r) => r.paragraphs)).toEqual([3])
    } finally { await stop() }
    // Six cells of which three are withheld (rotatedBox A1:B3: A1..A3 covered) count 3.
    const r = await session({ doc: await rotatedBox(), policy: { copy: { maxParagraphs: 3 } } })
    try {
      const text = await copyText(r.c, '0', boxes([0, 2, 0, 1]))
      expect(text).toBe('\tROT-1-2\n\tROT-2-2\n\tROT-3-2')
      expect(r.store.copies(r.m.session.sid).map((x) => x.paragraphs)).toEqual([3])
    } finally { await r.stop() }
  }, 30_000)

  test('… and per session, measured against the resumed cap', async () => {
    let policy: Policy = { copy: { maxParagraphs: 10, maxSessionParagraphs: 5 } }
    const { h, c, m, store, stop } = await session({ doc: await textGrid(), authorize: async () => ({ viewerId: 'alice', policy }) })
    const sid = m.session.sid
    try {
      expect(await copyText(c, '2', boxes([0, 0, 0, 2]))).toBe('P-1-1\tP-1-2\tP-1-3')
      expect(await copyText(c, '2', boxes([1, 1, 0, 1]))).toBe('P-2-1\tneedle-plain')
      const over = await copy(c, { sheetId: '2', selection: boxes([2, 2, 0, 0]) })
      expect(over.parts).toEqual([])
      expect(over.end).toEqual({ reqId: over.end.reqId, type: 'refused', code: 'copy-quota-exhausted', limit: 5, requested: 1 })
      // A resume under a larger session cap: the used 5 is kept and measured against 6 (Review Focus 2).
      policy = { copy: { maxParagraphs: 10, maxSessionParagraphs: 6 } }
      const b = await open(h.url, { resume: { sid, resumeProof: m.session.token } })
      try {
        expect(b.m.session.sid).toBe(sid)
        expect(b.m.caps.copy).toEqual({ maxParagraphs: 10, maxSessionParagraphs: 6 })
        expect(await copyText(b.c, '2', boxes([2, 2, 0, 0]))).toBe('P-3-1')
        const again = await copy(b.c, { sheetId: '2', selection: boxes([3, 3, 0, 0]) })
        expect(again.end).toMatchObject({ type: 'refused', code: 'copy-quota-exhausted', limit: 6, requested: 1 })
        expect(store.copies(sid).map((r) => r.paragraphs)).toEqual([3, 2, 1])
      } finally { b.c.close() }
    } finally { await stop() }
  }, 30_000)

  test('copy off', async () => {
    const { c, m, store, stop } = await session({ doc: await textGrid(), policy: { copy: false } })
    try {
      const a = await copy(c, { sheetId: '2', selection: boxes([0, 0, 0, 0]) })
      expect(a.end).toEqual({ reqId: a.end.reqId, type: 'refused', code: 'text-unavailable' })
      expect(store.copies(m.session.sid)).toEqual([])
    } finally { await stop() }
  }, 30_000)

  test('stale-view: an old viewGen is refused, uncounted and unaudited', async () => {
    const { c, m, store, stop } = await session({ doc: await textGrid(), policy: { copy: { maxParagraphs: 10, maxSessionParagraphs: 1 } } })
    try {
      const r = await c.text({ reqId: ++rq, type: 'sheet-resize', sheetId: '2', viewGen: 0, axis: 'col', index: 0, size: 100 })
      expect(r.end).toMatchObject({ type: 'sheet-view', viewGen: 1 })
      const stale = await copy(c, { sheetId: '2', viewGen: 0, selection: boxes([0, 0, 0, 0]) })
      expect(stale.parts).toEqual([])
      expect(stale.end).toEqual({ reqId: stale.end.reqId, type: 'refused', code: 'stale-view' })
      expect(store.copies(m.session.sid)).toEqual([])
      // The quota of 1 is still whole: the current viewGen copies.
      expect((await copy(c, { sheetId: '2', viewGen: 1, selection: boxes([0, 0, 0, 0]) })).text).toBe('P-1-1')
      // Another sheet keeps its own viewGen 0.
      const other = await copy(c, { sheetId: '1', viewGen: 1, selection: boxes([0, 0, 0, 0]) })
      expect(other.end).toMatchObject({ type: 'refused', code: 'stale-view' })
    } finally { await stop() }
  }, 30_000)

  test('audit', async () => {
    const { c, m, store, stop } = await session({ doc: await textGrid(), policy: { copy: 'full' } })
    const sid = m.session.sid
    try {
      expect(await copyText(c, '0', boxes([0, 0, 0, 2]))).toBe(await localText(await projectedOf(await textGrid()), 0, boxes([0, 0, 0, 2])))
      // Two ranges: the rect is their bounding source rect.
      await copyText(c, '2', { active: null, anchor: null, ranges: [{ rows: [[1, 1]], cols: [[2, 2]] }, { rows: [[3, 4]], cols: [[0, 0]] }] })
      expect(store.copies(sid)).toEqual([
        { sid, kind: 'copy', paragraphs: 3, fromUnit: 0, toUnit: 0, sheetId: '0', rect: { r0: 0, c0: 0, r1: 0, c1: 2 }, at: expect.any(Number) },
        { sid, kind: 'copy', paragraphs: 3, fromUnit: 2, toUnit: 2, sheetId: '2', rect: { r0: 1, c0: 0, r1: 4, c1: 2 }, at: expect.any(Number) },
      ])
      expect(store.columnsOf('copies')).toEqual(['sid', 'kind', 'paragraphs', 'from_unit', 'to_unit', 'at', 'sheet_id', 'rect'])
    } finally { await stop() }

    // A store created before 4b (no sheet_id, rect) opens, is migrated, and keeps its old rows unchanged.
    const dir = mkdtempSync(join(tmpdir(), 'sheet-copy-audit-'))
    try {
      const path = join(dir, 'audit.db')
      const raw = new Database(path, { create: true, strict: true })
      raw.exec('CREATE TABLE copies (sid TEXT NOT NULL, kind TEXT NOT NULL, paragraphs INTEGER NOT NULL, from_unit INTEGER NOT NULL, to_unit INTEGER NOT NULL, at INTEGER NOT NULL)')
      raw.query('INSERT INTO copies (sid, kind, paragraphs, from_unit, to_unit, at) VALUES (?, ?, ?, ?, ?, ?)').run('s1', 'copy', 2, 0, 1, 5)
      raw.close()
      const s = AuditStore.open(path)
      expect(s.columnsOf('copies')).toEqual(['sid', 'kind', 'paragraphs', 'from_unit', 'to_unit', 'at', 'sheet_id', 'rect'])
      expect(s.copies('s1')).toEqual([{ sid: 's1', kind: 'copy', paragraphs: 2, fromUnit: 0, toUnit: 1, at: 5 }])
      s.close()
      AuditStore.open(path).close() // reopening is idempotent
    } finally { rmSync(dir, { recursive: true, force: true }) }
  }, 30_000)

  test('select-all on a million rows', async () => {
    const bytes = await millionRows()
    const saved = await projectedOf(bytes, 'full')
    // Review Focus 4's precondition, as measured: the client's select-all (every row and column) is refused by the
    // local engine before any cell is read, so the server answers copy-too-large {limit: 4 MiB}, and never text.
    const all: WireSelection = { active: [0, 0], anchor: [0, 0], ranges: [{ rows: [[0, 1_048_575]], cols: [[0, 16_383]] }] }
    expect(await localCopy(saved, 0, all)).toMatchObject({ kind: 'refused', code: 'resource-limit' })
    // A full-height column (A1:A1048576, every row of the sheet) is under 4 MiB locally.
    const column = boxes([0, 1_048_575, 0, 0])
    const local = await localText(saved, 0, column)
    expect(Buffer.byteLength(local, 'utf8')).toBeLessThanOrEqual(COPY_LIMIT)
    expect(local.split('\n').length).toBe(1_048_576)

    const { c, stop } = await session({ doc: bytes, policy: { copy: 'full' } })
    try {
      const a = await copy(c, { sheetId: '0', selection: all })
      expect(a.end).toEqual({ reqId: a.end.reqId, type: 'refused', code: 'copy-too-large', limit: COPY_LIMIT })
      const col = await copy(c, { sheetId: '0', selection: column })
      expect(col.text).toBe(local)
      expect(col.end.last).toBe(true)
      expect(c.ws.readyState).toBe(WebSocket.OPEN)
      for (const p of col.parts) expect(p.bytes).toBeLessThanOrEqual(48 * 1024)
    } finally { await stop() }

    // Over 4 MiB of UTF-8 (and under the local clipboard cap): copy-too-large with the measured bytes.
    const big = await bigText()
    const bigLocal = await localText(await projectedOf(big, 'full'), 0, boxes([0, BIG_TEXT.rows - 1, 0, 0]))
    const bigBytes = Buffer.byteLength(bigLocal, 'utf8')
    expect(bigBytes).toBeGreaterThan(5 * 1024 * 1024)
    const b = await session({ doc: big, policy: { copy: 'full' } })
    try {
      const a = await copy(b.c, { sheetId: '0', selection: boxes([0, BIG_TEXT.rows - 1, 0, 0]) })
      expect(a.parts).toEqual([])
      expect(a.end).toEqual({ reqId: a.end.reqId, type: 'refused', code: 'copy-too-large', limit: COPY_LIMIT, requested: bigBytes })
      expect(b.store.copies(b.m.session.sid)).toEqual([])
      expect(b.c.ws.readyState).toBe(WebSocket.OPEN)
    } finally { await b.stop() }
  }, 60_000)

  test('cfHiddenFormat copies as painted', async () => {
    const bytes = await cfHiddenFormat()
    const saved = await projectedOf(bytes)
    const sel = boxes([0, 2, 1, 1])
    // Local with the CF scope paints B2 and B3 nothing; without it they read 7 and 9 (so the test can fail).
    expect(await localText(saved, 0, sel)).toBe('4\n\n')
    expect(await localText(saved, 0, sel, { scope: false })).toBe('4\n7\n9')
    const { c, m, store, stop } = await session({ doc: bytes, policy: { copy: 'full' } })
    try {
      expect(await copyText(c, '0', sel)).toBe('4\n\n')
      // Twice on one session: the cached view's scope survives a copy (the copy never disposes it).
      expect(await copyText(c, '0', sel)).toBe('4\n\n')
      expect(store.copies(m.session.sid).map((r) => r.paragraphs)).toEqual([1, 1])
    } finally { await stop() }
    // The CF-effective display text: cfNumFmt's rule writes `0.0` on B2 (without the scope it reads 7).
    const num = await cfNumFmt()
    const numSaved = await projectedOf(num)
    const local = await localText(numSaved, 0, sel)
    expect(local).toBe('4\n7.0\n9')
    expect(await localText(numSaved, 0, sel, { scope: false })).toBe('4\n7\n9')
    const n = await session({ doc: num, policy: { copy: 'full' } })
    try {
      expect(await copyText(n.c, '0', sel)).toBe(local)
      // Twice on one session: the copy borrows the cached view's scope and never disposes it.
      expect(await copyText(n.c, '0', sel)).toBe(local)
    } finally { await n.stop() }
  }, 30_000)

  test('hostile selections', async () => {
    let t = Date.now()
    const { c, m, store, stop } = await session({ doc: await textGrid(), policy: { copy: { maxParagraphs: 10, maxSessionParagraphs: 1 } }, now: () => t })
    try {
      const ok = boxes([0, 0, 0, 0])
      const range = (rows: unknown, cols: unknown) => ({ active: null, anchor: null, ranges: [{ rows, cols }] })
      const hostile: Array<[string, Record<string, unknown>]> = [
        ['no selection', { sheetId: '2' }],
        ['no viewGen', { sheetId: '2', viewGen: undefined, selection: ok }],
        ['negative viewGen', { sheetId: '2', viewGen: -1, selection: ok }],
        ['sheet out of range', { sheetId: '3', selection: ok }],
        ['non-canonical sheetId', { sheetId: '02', selection: ok }],
        ['numeric sheetId', { sheetId: 2, selection: ok }],
        ['257 ranges', { sheetId: '2', selection: { active: null, anchor: null, ranges: Array.from({ length: 257 }, (_, i) => ({ rows: [[i, i]], cols: [[0, 0]] })) } }],
        ['descending interval', { sheetId: '2', selection: range([[3, 1]], [[0, 0]]) }],
        ['overlapping intervals', { sheetId: '2', selection: range([[0, 3], [2, 5]], [[0, 0]]) }],
        ['unordered intervals', { sheetId: '2', selection: range([[4, 5], [0, 1]], [[0, 0]]) }],
        ['row past the sheet', { sheetId: '2', selection: range([[0, 1_048_576]], [[0, 0]]) }],
        ['column past the sheet', { sheetId: '2', selection: range([[0, 0]], [[0, 16_384]]) }],
        ['fractional index', { sheetId: '2', selection: range([[0, 1.5]], [[0, 0]]) }],
        ['empty rows', { sheetId: '2', selection: range([], [[0, 0]]) }],
        ['a row-set object', { sheetId: '2', selection: range({ kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, [[0, 0]]) }],
        ['a ref row set', { sheetId: '2', selection: range({ kind: 'ref', refId: 'x', count: 1 }, [[0, 0]]) }],
        ['an extra range key', { sheetId: '2', selection: { active: null, anchor: null, ranges: [{ rows: [[0, 0]], cols: [[0, 0]], kind: 'x' }] } }],
        ['an extra selection key', { sheetId: '2', selection: { ...ok, gesture: {} } }],
        ['a bad active cell', { sheetId: '2', selection: { ...ok, active: [0] } }],
      ]
      for (const [name, body] of hostile) {
        // 100 ms apart on the injected clock, so the message class's 20-per-second burst is never the refusal.
        t += 100
        const a = await copy(c, body)
        expect([name, a.parts.length, a.end.type, a.end.code]).toEqual([name, 0, 'refused', 'policy-refused'])
      }
      expect(store.copies(m.session.sid)).toEqual([])
      // Nothing was charged: the session quota of 1 is whole.
      t += 100
      expect((await copy(c, { sheetId: '2', selection: ok })).text).toBe('P-1-1')
    } finally { await stop() }
  }, 60_000)
})
