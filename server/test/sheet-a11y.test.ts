/**
 * Stage 4b Task 8 — `sheet-a11y` (plan P2, P4, P10, P12, P13): screen-reader text for a window of source cells, local
 * `formatCellPresentation` over the server text view (CF descriptor omitted), so each served text equals the local
 * presentation over the fixture with its withheld cells physically deleted. `unit` serves the window uncounted with one
 * audit row; `focused` serves the active cell (1×1) and counts 1 against the shared copy quota; `off` is refused.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { formatCellPresentation } from '../../src/components/xlsx/formatPresentation'
import { prepareCfTextScope } from '../../src/xlsx/cf/text'
import { createObjectWorkbookSource } from '../../src/xlsx/view/source'
import { buildSheetGeometry } from '../../src/xlsx/view/viewport'
import type { FormatBudget } from '../../src/xlsx/view/formatWorkspace'
import type { WireSelection } from '../../src/stream/sheet-wire'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../../src/xlsx/types'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import type { AuthDecision } from '../src/hooks'
import { DocumentHost } from '../src/render'
import { Client, disposePool, start } from './harness'
import {
  a11yFocus, A11Y_FOCUS, bigText, BIG_TEXT, cfHiddenFormat, cfNumFmt, mergeTail, offViewportCovered, rotatedBox, textGrid, TEXT_GRID,
  unfilledShapeOverText, unresolvedAnchor,
} from './sheet-fixtures'

const hosts: DocumentHost[] = []
afterAll(async () => {
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
  await disposePool()
})

const VIEW = ['sheet-view@1']
type Policy = AuthDecision['policy']
type Cells = Array<[number, number, string]>
let rq = 800

async function projectedOf(bytes: Uint8Array): Promise<XlsxDocument> {
  const host = new DocumentHost(); hosts.push(host)
  const o = await host.open('ref', bytes, 'visible-only', { sheetMode: 'tiles' })
  return (host as unknown as { docs: Map<string, { savedParse: XlsxDocument }> }).docs.get(o.docKey)!.savedParse
}
function without(doc: XlsxDocument, sheetIndex: number, drop: (cell: XlsxCell) => boolean): XlsxDocument {
  const del = (sheet: XlsxSheet): XlsxSheet => ({ ...sheet, rows: sheet.rows.map((row) => ({ ...row, cells: row.cells.filter((c) => !drop(c)) })) })
  return { ...doc, sheets: doc.sheets.map((s, i) => i === sheetIndex ? del(s) : s) }
}

const freeBudget: FormatBudget = { reserve() {}, release() {} }
/**
 * The local reference, written independently of the server's `textInputs`: the cells of the window the resident grid
 * presents (authored-hidden rows and zero-width columns are not in the grid), each as `presentCell` gives it
 * (`InteractiveSheetDoc.tsx` ~274–282): with a CF scope, the effective style over the region-resolved style; without
 * one, the cell's own style. Only non-empty texts are listed, row by row.
 */
async function localA11y(doc: XlsxDocument, sheetIndex: number, rows: Array<[number, number]>, cols: [number, number], o: { scope?: boolean } = {}): Promise<Cells> {
  const source = createObjectWorkbookSource(doc, 'local')
  const sheetId = source.sheets[sheetIndex]!.sheetId
  const sheet = doc.sheets[sheetIndex]!
  const geometry = buildSheetGeometry(sheet)
  if (geometry.kind !== 'ready') throw new Error('fixture geometry')
  const hidden = source.geometry(sheetId).baseHidden.intervals
  const rects = rows.map(([r0, r1]) => ({ r0, r1, c0: cols[0], c1: cols[1] }))
  const scope = sheet.conditionalFormatting?.rules.length && o.scope !== false
    ? await prepareCfTextScope(sheet, { rects, signal: new AbortController().signal, locale: 'en-US', budget: freeBudget })
    : undefined
  const out: Cells = []
  for (const [r0, r1] of rows) for (let row = r0; row <= r1; row++) {
    if (hidden.some((h) => h.r0 <= row && row <= h.r1)) continue
    for (let col = cols[0]; col <= cols[1]; col++) {
      if (geometry.metrics.cols.size(col) === 0) continue
      const cell = geometry.cellsOf(row)?.find((c) => c.col === col)
      if (!cell) continue
      const base = sheet.regionStyles?.styleAt(row, col, cell.style, 'screen') ?? cell.style
      const cf = scope ? { gen: scope.gen, style: scope.effectiveStyle(row, col, base) } : undefined
      const { text } = formatCellPresentation(sheet, cell, freeBudget, false, cf)
      if (text !== '') out.push([row, col, text])
    }
  }
  scope?.dispose()
  return out
}

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

interface A11yAnswer { parts: Array<Record<string, any>>; end: Record<string, any>; cells?: Cells }
/** Sends a `sheet-a11y` and collects its reqId's `sheet-a11y-result` chunks until `last: true`, or the one other reply. */
async function a11y(c: Client, body: Record<string, unknown>, reqId = ++rq): Promise<A11yAnswer> {
  c.send({ token: c.manifest!.session.token, reqId, type: 'sheet-a11y', viewGen: 0, ...body })
  const parts: Array<Record<string, any>> = []
  for (;;) {
    const m = await c.next((m) => 'text' in m && m.text.reqId === reqId, 30_000) as { text: Record<string, any> }
    if (m.text.type !== 'sheet-a11y-result') return { parts, end: m.text }
    expect(m.text.seq).toBe(parts.length)
    parts.push(m.text)
    if (m.text.last === true) return { parts, end: m.text, cells: parts.flatMap((p) => p.cells as Cells) }
  }
}
const cellsOf = async (c: Client, sheetId: string, rows: Array<[number, number]>, cols: [number, number]): Promise<Cells> => {
  const a = await a11y(c, { sheetId, rows, cols })
  if (a.cells === undefined) throw new Error(`a11y refused: ${JSON.stringify(a.end)}`)
  return a.cells
}
const boxes = (...b: Array<[number, number, number, number]>): WireSelection => ({ active: null, anchor: null, ranges: b.map(([r0, r1, c0, c1]) => ({ rows: [[r0, r1]], cols: [[c0, c1]] })) })
async function copyText(c: Client, sheetId: string, selection: WireSelection): Promise<Record<string, any>> {
  const r = await c.text({ reqId: ++rq, type: 'sheet-copy', sheetId, viewGen: 0, selection }, { partial: [] })
  return r.end
}

describe('sheet-a11y (P10)', () => {
  test('unit mode serves the window, uncounted', async () => {
    const bytes = await textGrid()
    const saved = await projectedOf(bytes)
    const [cr, cc] = TEXT_GRID.covered
    const withheld = (c: XlsxCell) => (c.row === cr && c.col === cc) || (c.row >= 4 && c.row <= 5 && c.col <= 2 && !(c.row === 4 && c.col === 0))
    const deleted = without(saved, 0, withheld)
    const window: [Array<[number, number]>, [number, number]] = [[[0, 29]], [0, 7]]
    const local = await localA11y(deleted, 0, ...window)
    // The comparison can fail: the authored presentation carries the covered needle.
    expect((await localA11y(saved, 0, ...window)).some(([, , t]) => t === 'NEEDLE-D20')).toBe(true)
    expect(local.some(([, , t]) => t.includes('HIDDEN') || t === 'NEEDLE-D20')).toBe(false)
    expect(local).toContainEqual([4, 0, 'needle in merge'])

    let policy: Policy = { copy: 'full' }
    const { h, c, m, store, stop } = await session({ doc: bytes, authorize: async () => ({ viewerId: 'alice', policy }) })
    const sid = m.session.sid
    try {
      expect(m.can.textA11y).toBe(true)
      expect(m.caps.a11y).toBe('unit')
      const a = await a11y(c, { sheetId: '0', rows: window[0], cols: window[1] })
      expect(a.cells).toEqual(local)
      // A window that holds no text answers one empty, final chunk.
      const empty = await a11y(c, { sheetId: '2', rows: [[10, 12]], cols: [0, 2] })
      expect(empty.parts).toEqual([{ reqId: empty.end.reqId, type: 'sheet-a11y-result', seq: 0, cells: [], last: true }])
      // One audit row per served request: the window's bounding rect, the sheet index as the unit range.
      expect(store.copies(sid)).toEqual([
        { sid, kind: 'a11y', paragraphs: local.length, fromUnit: 0, toUnit: 0, sheetId: '0', rect: { r0: 0, c0: 0, r1: 29, c1: 7 }, at: expect.any(Number) },
      ])
      // Uncounted: under a capped resume the whole session quota of 3 is still free.
      policy = { copy: { maxParagraphs: 3, maxSessionParagraphs: 3 } }
      const b = await open(h.url, { resume: { sid, resumeProof: m.session.token } })
      try {
        expect(b.m.caps.copy).toEqual({ maxParagraphs: 3, maxSessionParagraphs: 3 })
        const r = await b.c.text({ reqId: ++rq, type: 'sheet-copy', sheetId: '2', viewGen: 0, selection: boxes([0, 0, 0, 2]) }, { partial: [] })
        expect(r.end).toMatchObject({ type: 'copy-result', text: 'P-1-1\tP-1-2\tP-1-3', last: true })
      } finally { b.c.close() }
    } finally { await stop() }

    // Under 'full' the view is the identity, so the hidden row 7 and column F still hold their text in the view: the
    // grid does not present them, and neither does the answer (as local).
    const full = await session({ doc: bytes, policy: { contentPolicy: 'full', copy: 'full' } })
    try {
      const cells = await cellsOf(full.c, '0', [[5, 7]], [4, 6])
      expect(cells).toEqual(await localA11y(saved, 0, [[5, 7]], [4, 6]))
      expect(cells).toEqual([[5, 4, 'E6'], [5, 6, '6'], [7, 4, 'E8'], [7, 6, '1']])
    } finally { await full.stop() }
  }, 30_000)

  test('withheld cells have no text', async () => {
    const check = async (bytes: Uint8Array, rows: Array<[number, number]>, cols: [number, number], drop: (c: XlsxCell) => boolean, o: { scope?: boolean } = {}) => {
      const saved = await projectedOf(bytes)
      const local = await localA11y(without(saved, 0, drop), 0, rows, cols)
      // The comparison can fail: the authored presentation differs (for a CF-undrawn cell, the one without the scope).
      expect(await localA11y(saved, 0, rows, cols, o)).not.toEqual(local)
      const { c, stop } = await session({ doc: bytes, policy: { copy: 'full' } })
      try {
        const cells = await cellsOf(c, '0', rows, cols)
        expect(cells).toEqual(local)
        return cells
      } finally { await stop() }
    }
    const at = (cells: Cells, row: number, col: number) => cells.find(([r, c]) => r === row && c === col)?.[2]

    const shape = await check(await unfilledShapeOverText(), [[0, 3]], [0, 3], (c) => c.row === 1 && c.col === 1)
    expect(at(shape, 1, 1)).toBeUndefined()
    expect(at(shape, 0, 0)).toBe('CELL-1-1')
    expect(shape.some(([, , t]) => t === 'COVERED-B2')).toBe(false)

    const rot = await check(await rotatedBox(), [[0, 5]], [0, 2], (c) => c.col === 0 && c.row <= 2)
    expect([0, 1, 2].map((r) => at(rot, r, 0))).toEqual([undefined, undefined, undefined])
    expect(at(rot, 3, 0)).toBe('ROT-4-1')
    expect(at(rot, 0, 1)).toBe('ROT-1-2')

    const off = await check(await offViewportCovered(), [[4989, 5009]], [25, 25], (c) => c.row === 4999 && c.col === 25)
    expect(off).toEqual([])

    const merge = await check(await mergeTail(), [[0, 0]], [0, 2], (c) => c.row === 0 && c.col === 1)
    expect(merge).toEqual([[0, 0, 'MERGE-A1']])

    const cf = await check(await cfHiddenFormat(), [[0, 2]], [1, 1], (c) => c.col === 1 && c.row >= 1, { scope: false })
    expect(cf).toEqual([[0, 1, '4']])

    const unres = await check(await unresolvedAnchor(), [[0, 1]], [0, 1], () => true)
    expect(unres).toEqual([])

    // The control: offViewportCovered's uncovered A1 is served.
    const { c, stop } = await session({ doc: await offViewportCovered(), policy: { copy: 'full' } })
    try { expect(await cellsOf(c, '0', [[0, 0]], [0, 0])).toEqual([[0, 0, 'OFFVIEW-A1']]) } finally { await stop() }
  }, 30_000)

  test('the text is the CF-effective presentation', async () => {
    const bytes = await cfNumFmt()
    const saved = await projectedOf(bytes)
    const local = await localA11y(saved, 0, [[0, 2]], [1, 1])
    expect(local).toEqual([[0, 1, '4'], [1, 1, '7.0'], [2, 1, '9']])
    // The comparison can fail: without the scope B2 is plain 7.
    expect(await localA11y(saved, 0, [[0, 2]], [1, 1], { scope: false })).toEqual([[0, 1, '4'], [1, 1, '7'], [2, 1, '9']])
    const { c, stop } = await session({ doc: bytes, policy: { copy: 'full' } })
    try {
      // Twice on one session: the view's scope is lent, never disposed by a job.
      expect(await cellsOf(c, '0', [[0, 2]], [1, 1])).toEqual(local)
      expect(await cellsOf(c, '0', [[0, 2]], [1, 1])).toEqual(local)
    } finally { await stop() }
  }, 30_000)

  test('focused counts 1 per non-empty cell', async () => {
    const { c, m, store, stop } = await session({ doc: await a11yFocus(), policy: { copy: { maxParagraphs: 3, maxSessionParagraphs: 2 } } })
    const sid = m.session.sid
    try {
      expect(m.caps.a11y).toBe('focused')
      expect(m.can.textA11y).toBe(true)
      const wide = await a11y(c, { sheetId: '0', rows: [[0, 1]], cols: [1, 1] })
      expect(wide.parts).toEqual([])
      expect(wide.end).toEqual({ reqId: wide.end.reqId, type: 'refused', code: 'policy-refused' })
      const two = await a11y(c, { sheetId: '0', rows: [[0, 0], [2, 2]], cols: [1, 1] })
      expect(two.end).toMatchObject({ type: 'refused', code: 'policy-refused' })
      expect(await cellsOf(c, '0', [[0, 0]], [1, 1])).toEqual([[...A11Y_FOCUS.b1]])
      // An empty cell counts 0 and writes no audit row.
      expect(await cellsOf(c, '0', [[1, 1]], [1, 1])).toEqual([])
      expect(await cellsOf(c, '0', [[2, 2]], [1, 1])).toEqual([[...A11Y_FOCUS.b3]])
      const over = await a11y(c, { sheetId: '0', rows: [[4, 4]], cols: [1, 1] })
      expect(over.parts).toEqual([])
      expect(over.end).toEqual({ reqId: over.end.reqId, type: 'refused', code: 'copy-quota-exhausted', limit: 2, requested: 1 })
      // An empty cell still answers once the quota is spent: it counts 0.
      expect(await cellsOf(c, '0', [[3, 3]], [1, 1])).toEqual([])
      expect(store.copies(sid)).toEqual([
        { sid, kind: 'a11y', paragraphs: 1, fromUnit: 0, toUnit: 0, sheetId: '0', rect: { r0: 0, c0: 1, r1: 0, c1: 1 }, at: expect.any(Number) },
        { sid, kind: 'a11y', paragraphs: 1, fromUnit: 0, toUnit: 0, sheetId: '0', rect: { r0: 2, c0: 1, r1: 2, c1: 1 }, at: expect.any(Number) },
      ])
    } finally { await stop() }
  }, 30_000)

  test('the quota is shared with copy', async () => {
    const { c, m, store, stop } = await session({ doc: await a11yFocus(), policy: { copy: { maxParagraphs: 3, maxSessionParagraphs: 2 } } })
    try {
      // B1:B3 holds two texts: the copy spends the whole session quota.
      expect(await copyText(c, '0', boxes([0, 2, 1, 1]))).toMatchObject({ type: 'copy-result', text: 'FOCUS-B1\n\nFOCUS-B3', last: true })
      const over = await a11y(c, { sheetId: '0', rows: [[4, 4]], cols: [1, 1] })
      expect(over.parts).toEqual([])
      expect(over.end).toEqual({ reqId: over.end.reqId, type: 'refused', code: 'copy-quota-exhausted', limit: 2, requested: 1 })
      expect(store.copies(m.session.sid).map((r) => [r.kind, r.paragraphs])).toEqual([['copy', 2]])
    } finally { await stop() }
  }, 30_000)

  test('off', async () => {
    let t = Date.now()
    const { c, m, store, stop } = await session({ doc: await textGrid(), policy: { copy: false }, now: () => t })
    try {
      expect(m.can.textA11y).toBe(false)
      expect(m.caps.a11y).toBe('off')
      for (let i = 0; i < 60; i++) {
        // 100 ms apart on the injected clock, so the message class's 20-per-second burst is never the refusal.
        t += 100
        const a = await a11y(c, { sheetId: '2', rows: [[0, 0]], cols: [0, 0] })
        expect([i, a.parts.length, a.end.code]).toEqual([i, 0, 'text-unavailable'])
      }
      // Each off-mode request was charged 'text' (60 per minute): the 61st is rate-limited.
      t += 100
      const a = await a11y(c, { sheetId: '2', rows: [[0, 0]], cols: [0, 0] })
      expect(a.end).toEqual({ reqId: a.end.reqId, type: 'refused', code: 'rate-limited' })
      expect(store.copies(m.session.sid)).toEqual([])
    } finally { await stop() }
  }, 30_000)

  test('stale-view: an old viewGen is refused, uncounted and unaudited', async () => {
    const { c, m, store, stop } = await session({ doc: await textGrid(), policy: { copy: { maxParagraphs: 3, maxSessionParagraphs: 1 } } })
    try {
      const r = await c.text({ reqId: ++rq, type: 'sheet-resize', sheetId: '2', viewGen: 0, axis: 'col', index: 0, size: 100 })
      expect(r.end).toMatchObject({ type: 'sheet-view', viewGen: 1 })
      const stale = await a11y(c, { sheetId: '2', viewGen: 0, rows: [[0, 0]], cols: [0, 0] })
      expect(stale.parts).toEqual([])
      expect(stale.end).toEqual({ reqId: stale.end.reqId, type: 'refused', code: 'stale-view' })
      expect(store.copies(m.session.sid)).toEqual([])
      // The quota of 1 is still whole: the current viewGen reads.
      expect((await a11y(c, { sheetId: '2', viewGen: 1, rows: [[0, 0]], cols: [0, 0] })).cells).toEqual([[0, 0, 'P-1-1']])
      // Another sheet keeps its own viewGen 0.
      expect((await a11y(c, { sheetId: '1', viewGen: 1, rows: [[0, 0]], cols: [0, 0] })).end).toMatchObject({ type: 'refused', code: 'stale-view' })
    } finally { await stop() }
  }, 30_000)

  test('a large answer is chunked at 48 KiB', async () => {
    const bytes = await bigText()
    const rows: Array<[number, number]> = [[0, BIG_TEXT.rows - 1]]
    const local = await localA11y(await projectedOf(bytes), 0, rows, [0, 0])
    // Each label is bounded at 256 UTF-16 units (mostly 3-byte U+6587): about 150 KB in all, over three messages.
    expect(local.length).toBe(BIG_TEXT.rows)
    expect(local.every(([, , t]) => t.length === 256)).toBe(true)
    const { c, stop } = await session({ doc: bytes, policy: { copy: 'full' } })
    try {
      const reqId = ++rq
      c.send({ token: c.manifest!.session.token, reqId, type: 'sheet-a11y', sheetId: '0', viewGen: 0, rows, cols: [0, 0] })
      const parts: Array<{ text: Record<string, any>; bytes: number }> = []
      for (;;) {
        const m = await c.next((m) => 'text' in m && m.text.reqId === reqId, 30_000) as { text: Record<string, any>; bytes: number }
        expect(m.text.type).toBe('sheet-a11y-result')
        parts.push(m)
        if (m.text.last === true) break
      }
      expect(parts.length).toBeGreaterThan(1)
      expect(parts.map((p) => p.text.seq)).toEqual(parts.map((_, i) => i))
      expect(parts.map((p) => p.text.last)).toEqual(parts.map((_, i) => i === parts.length - 1))
      for (const p of parts) expect(p.bytes).toBeLessThanOrEqual(48 * 1024)
      expect(parts.flatMap((p) => p.text.cells)).toEqual(local)
    } finally { await stop() }
  }, 30_000)

  test('limits', async () => {
    let t = Date.now()
    const { c, m, store, stop } = await session({ doc: await textGrid(), policy: { copy: 'full' }, now: () => t })
    try {
      const hostile: Array<[string, Record<string, unknown>]> = [
        ['2,001 cells', { sheetId: '2', rows: [[0, 2000]], cols: [0, 0] }],
        ['2,001 cells over two intervals', { sheetId: '2', rows: [[0, 1499], [2000, 2500]], cols: [0, 0] }],
        ['no rows', { sheetId: '2', cols: [0, 0] }],
        ['descending cols', { sheetId: '2', rows: [[0, 0]], cols: [2, 1] }],
        ['column past the sheet', { sheetId: '2', rows: [[0, 0]], cols: [0, 16_384] }],
        ['overlapping rows', { sheetId: '2', rows: [[0, 3], [2, 5]], cols: [0, 0] }],
        ['sheet out of range', { sheetId: '3', rows: [[0, 0]], cols: [0, 0] }],
        ['no viewGen', { sheetId: '2', viewGen: undefined, rows: [[0, 0]], cols: [0, 0] }],
      ]
      for (const [name, body] of hostile) {
        t += 100
        const a = await a11y(c, body)
        expect([name, a.parts.length, a.end.type, a.end.code]).toEqual([name, 0, 'refused', 'policy-refused'])
      }
      expect(store.copies(m.session.sid)).toEqual([])
      // 2,000 cells are served.
      t += 100
      const full = await a11y(c, { sheetId: '0', rows: [[0, 149], [150, 249]], cols: [0, 7] })
      expect(full.cells).toEqual(await localA11y(without(await projectedOf(await textGrid()), 0, (x) => (x.row === 19 && x.col === 3) || (x.row >= 4 && x.row <= 5 && x.col <= 2 && !(x.row === 4 && x.col === 0))), 0, [[0, 149], [150, 249]], [0, 7]))
    } finally { await stop() }
  }, 30_000)
})
