// @vitest-environment node
/**
 * Task 33 perf fix (perf-audit-s3.md findings 3–6): work bounds that do not grow with the number of pivots on a sheet.
 * Counter-based (`ownedRectsWork`: point queries and the index nodes they visit), never timing.
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { XlsxDocument } from '../src/xlsx/types'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { compositeOverlay, createResidentPivotClient } from '../src/components/xlsx/pivot/pivotClient'
import { ownedRectsWork } from '../src/xlsx/view/rectIndex'
import { assertOverlayCellsProtected, createObjectWorkbookSource } from '../src/xlsx/view/source'
import { combinePivotOverlays } from '../src/worker/viewerSession'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { MAX_COLS, MAX_ROWS, type SourceRect } from '../src/xlsx/view/coordinates'
import type { ValueOverlay, WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import type { PivotDelta } from '../src/xlsx/pivots/session/delta'

const FIXTURE = resolve(fileURLToPath(new URL('./fixtures/xlsx-pivots/engine/', import.meta.url)), 'layout-collapsed-item.xlsx')
const EXPAND: PivotDelta = { kind: 'collapse', field: 0, item: 0, collapsed: false }
const PIVOTS = 100
/** Work between two snapshots: queries, and index nodes visited per query. */
function work<T>(run: () => T): { result: T; queries: number; perQuery: number } {
  const before = ownedRectsWork()
  const result = run()
  const after = ownedRectsWork()
  const queries = after.queries - before.queries
  return { result, queries, perQuery: queries ? (after.visits - before.visits) / queries : 0 }
}
async function workAsync<T>(run: () => Promise<T>): Promise<{ result: T; queries: number; perQuery: number }> {
  const before = ownedRectsWork()
  const result = await run()
  const after = ownedRectsWork()
  const queries = after.queries - before.queries
  return { result, queries, perQuery: queries ? (after.visits - before.visits) / queries : 0 }
}
/** layout-collapsed-item (A3:I21; expanding Apples grows it to A3:I24) with 100 more pivots of its cache below it. */
async function manyPivots(): Promise<{ doc: XlsxDocument; sheet: number }> {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
  const sheet = doc.sheets.findIndex(candidate => candidate.pivotTables?.length), s = doc.sheets[sheet]!
  const def = s.pivotTables![0]!
  s.pivotTables = [def, ...Array.from({ length: PIVOTS }, (_, k) => ({ ...def, name: `Copy ${k}`, location: { ...def.location, ref: `A${30 + 25 * k}:I${48 + 25 * k}` } }))]
  return { doc, sheet }
}
// log2(101 rectangles) is under 7; an interval-tree point query visits a few nodes per level.
const PER_QUERY = 48

describe('finding 3: the overlap check does not scan every other pivot per cell', () => {
  it('resident: 100 other pivots, every queried cell answered by an index query of bounded work', async () => {
    const { doc, sheet } = await manyPivots()
    const client = createResidentPivotClient({ doc, ledger: new MainMemoryLedger(), install: () => {} })
    expect((await client.open(sheet, 0, new AbortController().signal)).ok).toBe(true)
    const { result, queries, perQuery } = await workAsync(() => client.apply(sheet, 0, EXPAND))
    expect(result).toMatchObject({ ok: true, def: { location: { ref: 'A3:I24' } } })
    expect(queries).toBeGreaterThanOrEqual(27)
    expect(perQuery).toBeLessThanOrEqual(PER_QUERY)
  })

  it('worker: the same, with the other pivots\' saved areas and commits in one per-sheet index', async () => {
    const { doc, sheet } = await manyPivots()
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) }, { yield: async () => {} })
    let serial = 0
    const id = () => `perf-${++serial}`
    const run = async (command: WorkerCommand) => { const replies: WorkerReply[] = []; await kernel.handle(command, reply => replies.push(reply)); return replies[0]! }
    const docId = id(), sheetId = doc.sheets[sheet]!.sheetId ?? String(sheet + 1)
    await run({ kind: 'open', requestId: id(), docId, source: { kind: 'object', document: doc }, profile: 'D' })
    expect(await run({ kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })).toMatchObject({ kind: 'complete' })
    const { result, queries, perQuery } = await workAsync(() => run({ kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 0, revision: 1, delta: EXPAND, grantBytes: 16 * 1024 * 1024 }))
    expect(result).toMatchObject({ kind: 'complete', result: { kind: 'pivotApply' } })
    expect(queries).toBeGreaterThanOrEqual(27)
    expect(perQuery).toBeLessThanOrEqual(PER_QUERY)
    await run({ kind: 'close', requestId: id(), docId })
  })
})

/** 100 pivot parts stacked down the sheet (rows 20k … 20k + 9, columns A–D), each cell holding its pivot's name. */
function parts(count = PIVOTS) {
  return Array.from({ length: count }, (_, k) => {
    const rect = { minRow: 20 * k, maxRow: 20 * k + 9, minCol: 0, maxCol: 3 }
    const cells = []
    for (let row = rect.minRow; row <= rect.maxRow; row++) for (let col = 0; col <= 3; col++) cells.push({ row, col, value: { kind: 'text' as const, value: `p${k}` } })
    const area = [{ r0: rect.minRow, r1: rect.maxRow, c0: 0, c1: 3 }]
    return { saved: area, current: area, layout: { rect, cells, merges: [] } as never }
  })
}

describe('finding 4: composite overlays index their pivots', () => {
  it('resident compositeOverlay: 100 pivots, 10,000 reads, bounded work per read; repeats deduplicated', () => {
    const list = parts()
    // A duplicate rectangle per part (saved = current): protectedRects keeps one of each.
    const { result: overlay } = work(() => compositeOverlay(list))
    expect(overlay.protectedRects).toHaveLength(PIVOTS)
    const { queries, perQuery } = work(() => {
      for (let n = 0; n < 10_000; n++) {
        const k = n % PIVOTS, row = 20 * k + (n % 10)
        expect(overlay.cell(row, n % 4)).toEqual({ value: `p${k}` })
      }
      expect(overlay.cell(20 * 7 + 15, 0)).toBeUndefined()
    })
    expect(queries).toBeGreaterThanOrEqual(10_000)
    expect(perQuery).toBeLessThanOrEqual(PER_QUERY)
  })

  it('resident compositeOverlay: the pivot whose current area holds a cell wins over one that only vacated it', () => {
    const [a, b] = parts(2)
    // B vacated rows 20–29 (its saved area) and A grew over row 20–21 (its current); A answers whatever the order.
    const grown = { ...a!, current: [{ r0: 0, r1: 21, c0: 0, c1: 3 }], layout: { ...(a!.layout as object), cells: [{ row: 21, col: 0, value: { kind: 'text', value: 'a-grown' } }], merges: [] } as never }
    const vacated = { ...b!, current: [{ r0: 40, r1: 40, c0: 0, c1: 0 }] }
    for (const order of [[vacated, grown], [grown, vacated]]) {
      const overlay = compositeOverlay(order)
      expect(overlay.cell(21, 0)).toEqual({ value: 'a-grown' })
      expect(overlay.cell(25, 0)).toBeNull()
    }
  })

  it('worker combinePivotOverlays: 100 committed overlays, bounded work per read, the later overlay winning', () => {
    const overlays: ValueOverlay[] = Array.from({ length: PIVOTS }, (_, k) => {
      const rect = { r0: 20 * k, r1: 20 * k + 9, c0: 0, c1: 3 }
      return { rects: [rect], protectedRects: [rect], merges: [], cell: () => ({ value: `p${k}` }) }
    })
    overlays.push({ rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }], protectedRects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }], merges: [], cell: () => ({ value: 'last' }) })
    const combined = combinePivotOverlays(overlays)!
    const { queries, perQuery } = work(() => {
      for (let n = 0; n < 10_000; n++) { const k = n % PIVOTS; expect(combined.cell(20 * k + 5, n % 4)).toEqual({ value: `p${k}` }) }
      expect(combined.cell(0, 0)).toEqual({ value: 'last' })
      expect(combined.cell(15, 0)).toBeUndefined()
    })
    expect(queries).toBeGreaterThanOrEqual(10_000)
    expect(perQuery).toBeLessThanOrEqual(PER_QUERY)
  })
})

describe('finding 5: the overlay assertion', () => {
  const doc = (): XlsxDocument => ({ sheets: [{ name: 'S', sheetId: 's', rows: [], cols: [], merges: [], mergeRanges: [] }] })
  it('setting the overlay already installed does no work (the identity check comes first)', () => {
    const store = createObjectWorkbookSource(doc(), 'd').store('s')
    let reads = 0
    // A rectangle outside protectedRects: the assertion reads its cells.
    const overlay: ValueOverlay = { rects: [{ r0: 0, r1: 9, c0: 0, c1: 9 }], protectedRects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }], merges: [], cell: () => { reads++; return null } }
    store.setValueOverlay!(overlay)
    const first = reads, gen = store.stamp.valueGen
    expect(first).toBeGreaterThan(0)
    store.setValueOverlay!(overlay)
    expect(reads).toBe(first)
    expect(store.stamp.valueGen).toBe(gen)
  })
  it('containment is indexed: 5,000 rectangles against 5,000 reverse-ordered bounds', () => {
    const rects: SourceRect[] = Array.from({ length: 5_000 }, (_, k) => ({ r0: 2 * k, r1: 2 * k, c0: 0, c1: 3 }))
    const overlay: ValueOverlay = { rects, protectedRects: [...rects].reverse().map(rect => ({ ...rect, c1: 5 })), merges: [], cell: () => ({ value: 1 }) }
    const { queries, perQuery } = work(() => assertOverlayCellsProtected(overlay))
    expect(queries).toBe(5_000)
    expect(perQuery).toBeLessThanOrEqual(PER_QUERY)
    // A value outside every bound still throws.
    expect(() => assertOverlayCellsProtected({ ...overlay, rects: [...rects, { r0: 1, r1: 1, c0: 0, c1: 0 }] })).toThrow(/outside protectedRects/)
  })
})

describe('finding 6: a resident sheet\'s saved pivot bounds are the file\'s, not the live definitions\'', () => {
  it('a commit that moves the live location leaves the guard on the saved rectangle', async () => {
    const parsed = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    const sheet = parsed.sheets.find(candidate => candidate.pivotTables?.length)!
    const store = createObjectWorkbookSource(parsed, 'd').store(sheet.sheetId ?? '1')
    const query = async () => (await store.queryStructural({ r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }, { signal: new AbortController().signal, stamp: store.stamp })).pivots ?? []
    const before = await query()
    expect(before).toContainEqual({ r0: 2, r1: 20, c0: 0, c1: 8 })
    // What installPivot does: the live definition and region follow the new location.
    sheet.pivotTables = sheet.pivotTables!.map(def => ({ ...def, location: { ...def.location, ref: 'K40:L41' } }))
    sheet.pivotRegions = sheet.pivotRegions!.map(region => ({ ...region, ref: 'K40:L41' }))
    const after = await query()
    expect(after).toContainEqual({ r0: 2, r1: 20, c0: 0, c1: 8 })
    expect(after).not.toContainEqual({ r0: 39, r1: 40, c0: 10, c1: 11 })
  })
})
