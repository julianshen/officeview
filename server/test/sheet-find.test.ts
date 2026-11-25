/**
 * Stage 4b Task 6 — `sheet-find` (plan P2, P6, P12, P13): local `WorkbookFind` over the server text view, batched one
 * `sheet-text` job per sheet in workbook order. Answers carry positions only (`[sheetIndex, row, col]`), so a find
 * equals the local find over the fixture with its withheld cells physically deleted, and no hit text crosses the wire.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { WorkbookFind, type FindSheetInput } from '../../src/xlsx/view/find'
import { createObjectWorkbookSource } from '../../src/xlsx/view/source'
import { initialWorkbookOwner } from '../../src/worker/viewerProtocol'
import { createIdentityOrder } from '../../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../../src/xlsx/view/visibility'
import { buildSheetGeometry } from '../../src/xlsx/view/viewport'
import { prepareCfTextScope } from '../../src/xlsx/cf/text'
import { cellDisplayText } from '../../src/xlsx/render'
import type { FormatBudget } from '../../src/xlsx/view/formatWorkspace'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../../src/xlsx/types'
import { DocumentHost } from '../src/render'
import type { Pool } from '../src/pool'
import { Client, disposePool, pool, start } from './harness'
import {
  cfHiddenFormat, cfNumFmt, hiddenFormat, manyHits, MANY_HITS, mergeTail, offViewportCovered, rotatedBox, slidUnderAnchor,
  textGrid, TEXT_GRID, unfilledShapeOverText, unresolvedAnchor,
} from './sheet-fixtures'

const hosts: DocumentHost[] = []
afterAll(async () => {
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
  await disposePool()
})

const VIEW = ['sheet-view@1']
let rq = 600

/** The fixture's projected saved-values workbook (what the server's snapshots are built from). */
async function projectedOf(bytes: Uint8Array, policy: 'visible-only' | 'full' = 'visible-only'): Promise<XlsxDocument> {
  const host = new DocumentHost(); hosts.push(host)
  const o = await host.open('ref', bytes, policy, { sheetMode: 'tiles' })
  return (host as unknown as { docs: Map<string, { savedParse: XlsxDocument }> }).docs.get(o.docKey)!.savedParse
}
/** `sheet` with every cell `drop` names physically deleted. */
const without = (sheet: XlsxSheet, drop: (cell: XlsxCell) => boolean): XlsxSheet => ({ ...sheet, rows: sheet.rows.map((row) => ({ ...row, cells: row.cells.filter((c) => !drop(c)) })) })

const freeBudget: FormatBudget = { reserve() {}, release() {} }
/**
 * The local reference: `WorkbookFind` over an object source of `doc`, with the inputs the resident viewer builds
 * (`InteractiveSheetDoc.tsx` ~1819–1857), written out here independently of the server's `textInputs`. `scope`: the
 * resident CF preparation (`prepareCfTextScope` per scanned batch).
 */
async function localFind(doc: XlsxDocument, query: string, o: { matchCase?: boolean; scope?: boolean } = {}): Promise<Array<[number, number, number]>> {
  const source = createObjectWorkbookSource(doc, 'local')
  const owner = initialWorkbookOwner('local', source.sheets)
  const entries = new Map(source.sheets.map(({ sheetId }, index) => {
    const sheet = doc.sheets[index]!
    const geometry = buildSheetGeometry(sheet)
    if (geometry.kind !== 'ready') throw new Error('fixture geometry')
    return [sheetId, { index, sheet, geometry }] as const
  }))
  const sheets: FindSheetInput[] = source.sheets.map(({ sheetId }) => {
    const v = owner.sheetVector.find((e) => e.sheetId === sheetId)!
    const { geometry } = entries.get(sheetId)!
    return {
      sheetId, store: source.store(sheetId),
      identity: { docId: 'local', sheetId, valueGen: v.valueGen, fontGen: owner.fontGen, viewGen: v.revision },
      order: createIdentityOrder(),
      visibility: visibilityFromHidden(createIntervalSet(source.geometry(sheetId).baseHidden.intervals.map(({ r0, r1 }) => ({ r0, r1 })))),
      isColumnVisible: (col: number) => geometry.metrics.cols.size(col) > 0,
    }
  })
  const find = new WorkbookFind({ owner, sheets }, { text: query, caseSensitive: o.matchCase === true }, {
    formatBudget: freeBudget,
    textOf: (cell, sheetId, allocation, scope) => {
      const e = entries.get(sheetId)!
      const styled = e.geometry.cellsOf(cell.row)?.find((c) => c.col === cell.col)
      const base = e.sheet.regionStyles?.styleAt(cell.row, cell.col, styled?.style, 'screen') ?? styled?.style
      return cellDisplayText(e.sheet, { value: (styled ?? cell).value, style: scope?.effectiveStyle(cell.row, cell.col, base) ?? base }, allocation)
    },
    ...(o.scope ? { prepareScope: (sheetId: string, rects: readonly import('../../src/xlsx/view/coordinates').SourceRect[], signal: AbortSignal) => prepareCfTextScope(entries.get(sheetId)!.sheet, { rects, signal, locale: 'en-US', budget: freeBudget }) } : {}),
  })
  const out: Array<[number, number, number]> = []
  try {
    for (;;) {
      const page = await find.nextPage()
      for (const h of page.hits) out.push([entries.get(h.sheetId)!.index, h.source.row, h.source.col])
      if (page.complete) break
    }
  } finally { find.close() }
  return out
}

async function open(url: string) {
  const c = new Client(url)
  await c.hello({ features: VIEW })
  await c.ready()
  return c
}
type Answer = { parts: Array<{ text: Record<string, any>; bytes: number }>; end: Record<string, any> }
const find = (c: Client, body: Record<string, unknown>, reqId = ++rq): Promise<Answer> =>
  c.text({ reqId, type: 'sheet-find', ...body }, { partial: ['sheet-find-result'] }) as Promise<Answer>
const hitsOf = (a: Answer): number[][] => a.parts.flatMap((p) => p.text.hits as number[][])
/** The `[row, col]` of every hit on sheet 0 whose cell text is `text`, against the projected workbook. */
const textsOf = (doc: XlsxDocument, a: Answer): string[] => hitsOf(a).map(([s, r, c]) => String(doc.sheets[s]!.rows.find((x) => x.index === r)?.cells.find((x) => x.col === c)?.value))

describe('sheet-find (P6)', () => {
  test('hit set and order equal local find on the deleted fixture', async () => {
    const bytes = await textGrid()
    const saved = await projectedOf(bytes)
    const [cr, cc] = TEXT_GRID.covered
    // Withheld in textGrid: the covered D20 and the tails of the merge A5:C6, on both grid sheets.
    const tail = (c: XlsxCell) => c.row >= 4 && c.row <= 5 && c.col <= 2 && !(c.row === 4 && c.col === 0)
    const deleted: XlsxDocument = { ...saved, sheets: saved.sheets.map((s, i) => i < 2 ? without(s, (c) => (c.row === cr && c.col === cc) || tail(c)) : s) }
    // The comparison can fail: the undeleted workbook finds NEEDLE-D20 on both grid sheets.
    const authored = await localFind(saved, 'needle')
    expect(authored).toContainEqual([0, cr, cc])
    expect(authored).toContainEqual([1, cr, cc])

    const h = await start({ doc: bytes })
    const c = await open(h.url)
    try {
      // '1,234.5' is the brief's query (it matches nothing in textGrid); '1.5' and '45.0' hit the `#,##0.0` column B.
      for (const query of ['needle', 'NEEDLE', '1,234.5', '2026', '1.5', '45.0']) {
        const local = await localFind(deleted, query)
        const a = await find(c, { query })
        expect([query, a.end.type]).toEqual([query, 'sheet-find-done'])
        expect([query, hitsOf(a)]).toEqual([query, local])
        expect(a.end.total).toBe(local.length)
        expect(a.end.truncated).toBe(false)
      }
      // 13 needle cells: 6 per grid sheet + 1 on Plain, less the covered D20 on each grid sheet.
      expect(hitsOf(await find(c, { query: 'needle' })).length).toBe(11)
      // matchCase: only the upper-case cells (NEEDLE ×2 per grid sheet; NEEDLE-D20 is withheld).
      const upper = await find(c, { query: 'NEEDLE', matchCase: true })
      expect(hitsOf(upper)).toEqual(await localFind(deleted, 'NEEDLE', { matchCase: true }))
      expect(hitsOf(upper).length).toBe(4)
      // A sheetId scope runs only that sheet.
      expect(hitsOf(await find(c, { query: 'needle', scope: '2' }))).toEqual([[2, 1, 1]])
      expect(hitsOf(await find(c, { query: 'needle', scope: '1' }))).toEqual((await localFind(deleted, 'needle')).filter(([s]) => s === 1))
    } finally { c.ws.close(); await h.stop() }
  }, 30_000)

  test('withheld cells never match', async () => {
    const run = async (bytes: Uint8Array, queries: string[]): Promise<Record<string, string[]>> => {
      const doc = await projectedOf(bytes, 'full')
      const h = await start({ doc: bytes })
      const c = await open(h.url)
      try {
        const out: Record<string, string[]> = {}
        for (const q of queries) {
          const a = await find(c, { query: q })
          expect(a.end.type).toBe('sheet-find-done')
          out[q] = textsOf(doc, a)
        }
        return out
      } finally { c.ws.close(); await h.stop() }
    }
    const shape = await run(await unfilledShapeOverText(), ['COVERED', 'CELL-2'])
    expect(shape.COVERED).toEqual([])
    expect(shape['CELL-2']).toEqual(['CELL-2-1', 'CELL-2-3', 'CELL-2-4'])

    const rot = (await run(await rotatedBox(), ['ROT-']))['ROT-']!
    for (const covered of ['ROT-1-1', 'ROT-2-1', 'ROT-3-1']) expect(rot).not.toContain(covered)
    expect(rot).toContain('ROT-4-1')
    expect(rot).toContain('ROT-1-2')
    expect(rot.length).toBe(15)

    const off = await run(await offViewportCovered(), ['OFFVIEW'])
    expect(off.OFFVIEW).toEqual(['OFFVIEW-A1'])

    const merge = await run(await mergeTail(), ['TAIL', 'MERGE'])
    expect(merge.TAIL).toEqual([])
    expect(merge.MERGE).toEqual(['MERGE-A1'])

    const fmt = await run(await hiddenFormat(), ['HIDDENFMT', 'SHOWN'])
    expect(fmt.HIDDENFMT).toEqual([])
    expect(fmt.SHOWN).toEqual(['SHOWN-A1'])

    const cf = await run(await cfHiddenFormat(), ['4', '7', '9'])
    expect(cf).toEqual({ 4: ['4'], 7: [], 9: [] })

    const unres = await run(await unresolvedAnchor(), ['UNRES', 'U', '1'])
    expect(unres).toEqual({ UNRES: [], U: [], 1: [] })

    // Current-view coverage: shrinking column B slides E1 under the anchored shape.
    const bytes = await slidUnderAnchor()
    const h = await start({ doc: bytes })
    const c = await open(h.url)
    try {
      expect(hitsOf(await find(c, { query: 'SLID' }))).toEqual([[0, 0, 4]])
      const r = await c.text({ reqId: ++rq, type: 'sheet-resize', sheetId: '0', viewGen: 0, axis: 'col', index: 1, size: 8 })
      expect(r.end.type).toBe('sheet-view')
      const after = await find(c, { query: 'SLID' })
      expect(after.end).toMatchObject({ type: 'sheet-find-done', total: 0, truncated: false })
      expect(hitsOf(after)).toEqual([])
    } finally { c.ws.close(); await h.stop() }
  }, 30_000)

  test('full finds as local', async () => {
    const bytes = await unfilledShapeOverText()
    const doc = await projectedOf(bytes, 'full')
    const h = await start({ doc: bytes, policy: { contentPolicy: 'full' } })
    const c = await open(h.url)
    try {
      const a = await find(c, { query: 'COVERED' })
      expect(hitsOf(a)).toEqual([[0, 1, 1]])
      expect(hitsOf(await find(c, { query: 'CELL' }))).toEqual(await localFind(doc, 'CELL'))
    } finally { c.ws.close(); await h.stop() }
  }, 30_000)

  test('CF-effective display text', async () => {
    const bytes = await cfNumFmt()
    const saved = await projectedOf(bytes)
    const local = await localFind(saved, '7.0', { scope: true })
    // The comparison can fail: without the CF scope B2 reads `7`, never `7.0`.
    expect(local).toEqual([[0, 1, 1]])
    expect(await localFind(saved, '7.0')).toEqual([])
    const h = await start({ doc: bytes })
    const c = await open(h.url)
    try {
      expect(hitsOf(await find(c, { query: '7.0' }))).toEqual(local)
      // Twice on one session: the cached view's scope survives a find (the find never disposes it).
      expect(hitsOf(await find(c, { query: '7.0' }))).toEqual(local)
      expect(hitsOf(await find(c, { query: '4' }))).toEqual(await localFind(saved, '4', { scope: true }))
    } finally { c.ws.close(); await h.stop() }
  }, 30_000)

  test('batched by sheet', async () => {
    const base = pool()
    const calls: Array<{ sheetIndex: number; limit: number }> = []
    let hold: Promise<void> | undefined
    const counting = new Proxy(base, {
      get(target, key) {
        if (key === 'sheetText') {
          return async (...args: Parameters<Pool['sheetText']>) => {
            const job = args[1] as { op: string; sheetIndex: number; limit: number }
            if (job.op === 'find') calls.push({ sheetIndex: job.sheetIndex, limit: job.limit })
            const r = await target.sheetText(...args)
            if (hold) await hold
            return r
          }
        }
        const v = Reflect.get(target, key, target)
        return typeof v === 'function' ? v.bind(target) : v
      },
    }) as Pool
    const h = await start({ doc: await textGrid(), pool: counting })
    try {
      const c = await open(h.url)
      const a = await find(c, { query: 'needle' })
      expect(a.end.type).toBe('sheet-find-done')
      // One job per sheet, in workbook order, each asking for what is left of the 5,000.
      expect(calls.map((x) => x.sheetIndex)).toEqual([0, 1, 2])
      expect(calls.map((x) => x.limit)).toEqual([5000, 4995, 4990])
      // The hits arrive in order, each batch as its own result message.
      expect(a.parts.map((p) => (p.text.hits as number[][]).map(([s]) => s)[0])).toEqual([0, 1, 2])

      // A socket closed during the first job: no second job runs.
      calls.length = 0
      let release!: () => void
      hold = new Promise<void>((r) => { release = r })
      c.send({ reqId: ++rq, type: 'sheet-find', token: c.manifest!.session.token, query: 'needle' })
      for (let i = 0; i < 200 && calls.length === 0; i++) await new Promise((r) => setTimeout(r, 25))
      expect(calls.length).toBe(1)
      c.ws.close()
      await c.closed
      await new Promise((r) => setTimeout(r, 100))
      release()
      await new Promise((r) => setTimeout(r, 500))
      expect(calls.length).toBe(1)
      hold = undefined

      // A newer find supersedes the running one: it is answered done with what it sent, `truncated: true`.
      const c2 = await open(h.url)
      calls.length = 0
      hold = new Promise<void>((r) => { release = r })
      const first = ++rq
      c2.send({ reqId: first, type: 'sheet-find', token: c2.manifest!.session.token, query: 'needle' })
      for (let i = 0; i < 200 && calls.length === 0; i++) await new Promise((r) => setTimeout(r, 25))
      const second = find(c2, { query: 'NEEDLE' })
      const superseded = await c2.next((m) => 'text' in m && m.text.reqId === first && m.text.type === 'sheet-find-done')
      expect((superseded as { text: Record<string, unknown> }).text).toMatchObject({ total: 0, truncated: true })
      hold = undefined
      release()
      const b = await second
      expect(b.end).toMatchObject({ type: 'sheet-find-done', truncated: false })
      // Nothing of the superseded find was sent after its done.
      expect(c2.msgs.filter((m) => 'text' in m && m.text.reqId === first)).toEqual([])
      c2.ws.close()
    } finally { await h.stop() }
  }, 30_000)

  test('no text on the wire', async () => {
    const h = await start({ doc: await textGrid() })
    const c = await open(h.url)
    try {
      const reqId = ++rq
      const a = await find(c, { query: 'needle' }, reqId)
      expect(hitsOf(a).length).toBe(11)
      for (const m of [...a.parts.map((p) => p.text), a.end]) {
        expect(JSON.stringify(m).toLowerCase()).not.toContain('needle')
        expect(Object.keys(m).sort()).toEqual(m.type === 'sheet-find-result' ? ['hits', 'reqId', 'type'] : ['reqId', 'total', 'truncated', 'type'])
      }
      for (const p of a.parts) for (const hit of p.text.hits as unknown[]) expect(Array.isArray(hit) && hit.length === 3 && hit.every((x) => Number.isInteger(x))).toBe(true)
      // Nothing else on the socket mentions it either.
      for (const m of c.msgs) if ('text' in m) expect(JSON.stringify(m.text).toLowerCase()).not.toContain('needle')
    } finally { c.ws.close(); await h.stop() }
  }, 30_000)

  test('limits', async () => {
    let t = Date.now()
    const h = await start({ doc: await manyHits(), now: () => t })
    const c = await open(h.url)
    try {
      const a = await find(c, { query: 'hit' })
      expect(a.end).toMatchObject({ type: 'sheet-find-done', total: 5000, truncated: true })
      const hits = hitsOf(a)
      expect(hits.length).toBe(5000)
      expect(hits[0]).toEqual([0, 0, 0])
      expect(hits[4999]).toEqual([0, 4999, 0])
      for (const p of a.parts) expect(p.bytes).toBeLessThanOrEqual(48 * 1024)
      expect(a.parts.length).toBeGreaterThan(1)
      // Below the cap, the answer is complete.
      const few = await find(c, { query: 'hit-59' })
      expect(few.end).toMatchObject({ total: 111, truncated: false }) // hit-59, hit-590..599, hit-5900..5999
      expect(MANY_HITS).toBe(6000)

      // 257 UTF-16 units: policy-refused (and not charged).
      const long = await find(c, { query: 'x'.repeat(257) })
      expect(long.end).toMatchObject({ type: 'refused', code: 'policy-refused' })
      expect((await find(c, { query: 'x'.repeat(256) })).end.type).toBe('sheet-find-done')

      // The 'text' class: 60 a minute, refuse-only. 3 finds were charged above.
      t += 61_000
      // 100 ms apart on the injected clock, so the message class's 20-per-second burst is never the refusal.
      for (let i = 0; i < 60; i++) { t += 100; expect([i, (await find(c, { query: 'zzz' })).end.type]).toEqual([i, 'sheet-find-done']) }
      t += 100
      const over = await find(c, { query: 'zzz' })
      expect(over.end).toMatchObject({ type: 'refused', code: 'rate-limited' })
      expect(c.ws.readyState).toBe(WebSocket.OPEN)
      t += 61_000
      expect((await find(c, { query: 'zzz' })).end.type).toBe('sheet-find-done')
    } finally { c.ws.close(); await h.stop() }
  }, 30_000)

  test('search works without copy', async () => {
    const h = await start({ doc: await textGrid(), policy: { copy: false } })
    const c = await open(h.url)
    try {
      const a = await find(c, { query: 'needle' })
      expect(a.end.type).toBe('sheet-find-done')
      expect(hitsOf(a).length).toBe(11)
    } finally { c.ws.close(); await h.stop() }
  }, 30_000)
})
