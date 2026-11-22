/**
 * Stage 4a Task 12 (S4 §6.1, R16): `StreamWorkbook`, the geometry-only viewer source of a tiles-mode stream. Its store
 * holds no cell data and every content method rejects `stream-no-cell-data`; it reports no paged content; it has no
 * `ViewerClient` and is not a `RemoteWorkbook`; its shell document is names only; its layout maps the wire's collapsed
 * runs; and the old `isRemoteWorkbook` guard is gone from `src/`.
 */
import { readdirSync, readFileSync, statSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, expectTypeOf, it, vi } from 'vitest'
import { StreamNoCellData, StreamSheetStore, StreamWorkbook, streamShellDoc } from '../src/components/xlsx/streamWorkbook'
import { sheetSourceOf } from '../src/components/xlsx/viewerSource'
import type { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import type { RemoteSheets } from '../src/stream/remote/remote-doc'
import { decodeSheetLayout, encodeSheetLayout, type StreamSheetLayout } from '../src/stream/sheet-wire'
import type { ReadOptions } from '../src/xlsx/view/contracts'

const LIST = [{ sheetId: '0', name: 'Data' }, { sheetId: '1', name: 'Summary' }]
const wire = (sheetId: string, over: Partial<StreamSheetLayout> = {}): StreamSheetLayout => ({
  sheetId, viewGen: 3, defaults: { rowPx: 20, colPx: 64 }, rows: [], cols: [], collapsedRows: [], collapsedCols: [],
  merges: [], filters: [], sort: null, filterSummary: [], authorHidden: false, counts: { visibleRows: 10, totalRows: 10 }, ...over,
})
function fakeSheets(layouts: Record<string, StreamSheetLayout> = {}): RemoteSheets & { layout: ReturnType<typeof vi.fn> } {
  return {
    list: LIST,
    layout: vi.fn(async (sheetId: string) => layouts[sheetId] ?? wire(sheetId)),
    tiles: vi.fn(async () => ({ refused: [] })),
  }
}
const opts = (store: StreamSheetStore): ReadOptions => ({ signal: new AbortController().signal, stamp: store.stamp })
const rect = { r0: 0, r1: 0, c0: 0, c1: 0 }

async function rejectsNoCellData(p: Promise<unknown>, method: string) {
  const error = await p.then(() => undefined, (e: unknown) => e)
  expect(error).toBeInstanceOf(StreamNoCellData)
  expect((error as StreamNoCellData).code).toBe('stream-no-cell-data')
  expect((error as StreamNoCellData).method).toBe(method)
}

describe('StreamWorkbook (S4 §6.1)', () => {
  it('every content method rejects stream-no-cell-data', async () => {
    const workbook = new StreamWorkbook(fakeSheets(), 'doc-ref')
    const store = workbook.store('0')
    expect(store).toBeInstanceOf(StreamSheetStore)
    const o = opts(store)
    await rejectsNoCellData(store.readCell({ row: 0, col: 0 }, o), 'readCell')
    await rejectsNoCellData(store.readRows(rect, undefined, o), 'readRows')
    await rejectsNoCellData(store.queryColumn({ col: 0, r0: 0, r1: 0 } as never, o), 'queryColumn')
    await rejectsNoCellData(store.scanPopulated({ rect } as never, o)[Symbol.asyncIterator]().next(), 'scanPopulated')
    await rejectsNoCellData(store.scanOccupied!({ rect } as never, o)[Symbol.asyncIterator]().next(), 'scanOccupied')
    await rejectsNoCellData(store.queryStructural(rect, o), 'queryStructural')
    await rejectsNoCellData(store.queryOverview({} as never, o), 'queryOverview')
    await rejectsNoCellData(store.ensureComplete(rect, o), 'ensureComplete')
  })

  it('stream reports no paged content', async () => {
    const workbook = new StreamWorkbook(fakeSheets(), 'doc-ref')
    for (const { sheetId } of LIST) {
      await workbook.activate(sheetId)
      expect(workbook.hasSparklines(sheetId)).toBe(false)
      expect(workbook.hasConditionalFormatting(sheetId)).toBe(false)
      expect(workbook.hasRegionStyles(sheetId)).toBe(false)
      expect(workbook.hasPivotDisplay(sheetId)).toBe(false)
      expect(workbook.pivotTableCount(sheetId)).toBe(0)
      const layout = workbook.layout(sheetId)
      expect(layout.hasSparklines).toBe(false)
      expect(layout.hasConditionalFormatting).toBe(false)
      expect(layout.hasRegionStyles).toBe(false)
      expect(layout.hasPivotDisplay).toBe(false)
      expect(layout.pivotTableCount).toBe(0)
      expect(workbook.store(sheetId).stamp.valueGen).toBe(0)
      expect(workbook.store(sheetId).completeness).toBe('complete')
    }
    const owner = workbook.identity('0'), signal = new AbortController().signal
    await rejectsNoCellData(workbook.sort(owner, null, signal), 'sort')
    await rejectsNoCellData(workbook.filter(owner, null, [], signal), 'filter')
  })

  it('no ViewerClient, not a RemoteWorkbook', () => {
    expectTypeOf<StreamWorkbook>().not.toHaveProperty('client')
    const workbook = new StreamWorkbook(fakeSheets(), 'doc-ref')
    // @ts-expect-error a StreamWorkbook is not a RemoteWorkbook (widened, never subclassed)
    const asPaged: RemoteWorkbook = workbook
    void asPaged
    expect(workbook.kind).toBe('stream')
    expect(Object.keys(workbook)).not.toContain('client')
    expect(sheetSourceOf(workbook)).toEqual({ paged: null, stream: workbook })
    expect(sheetSourceOf(undefined)).toEqual({ paged: null, stream: null })
  })

  it('zero-content shell', () => {
    const doc = streamShellDoc(LIST)
    expect(doc.sheets.map(s => [s.sheetId, s.name])).toEqual([['0', 'Data'], ['1', 'Summary']])
    for (const sheet of doc.sheets) {
      expect(sheet.rows).toEqual([])
      expect(sheet.rows.flatMap(row => row.cells)).toEqual([])
      expect(sheet.cols).toEqual([])
      expect(sheet.merges).toEqual([])
      expect(sheet.mergeRanges).toEqual([])
      expect(sheet.drawings ?? []).toEqual([])
      expect(sheet.regionStyles).toBeUndefined()
      expect(sheet.tables).toBeUndefined()
      expect(sheet.conditionalFormatting).toBeUndefined()
      expect(sheet.pivotTables).toBeUndefined()
    }
    expect(doc.images ?? []).toEqual([])
  })

  it('layout maps collapsed runs', async () => {
    const decoded = decodeSheetLayout(encodeSheetLayout(wire('0', { collapsedRows: [{ start: 1, end: 1 }], collapsedCols: [{ start: 2, end: 3 }], viewGen: 7 })))
    const sheets = fakeSheets({ '0': decoded })
    const workbook = new StreamWorkbook(sheets, 'doc-ref')
    expect(workbook.loadState('0').completeness).toBe('loading')
    const before = workbook.merges('0').gen
    let notified = 0
    workbook.subscribe(() => { notified++ })
    await workbook.activate('0')
    expect(sheets.layout).toHaveBeenCalledTimes(1)
    expect(sheets.layout.mock.calls[0][0]).toBe('0')
    expect(workbook.layout('0').hidden).toEqual([{ r0: 1, r1: 1 }])
    expect(workbook.layout('0').hiddenCols).toEqual([{ start: 2, end: 3 }])
    expect(workbook.loadState('0').completeness).toBe('complete')
    expect(workbook.merges('0').gen).toBeGreaterThan(before)
    expect(notified).toBeGreaterThan(0)
    expect(workbook.identity('0')).toEqual({ docId: 'doc-ref', sheetId: '0', valueGen: 0, fontGen: 0, viewGen: 7 })
    expect(workbook.view('0')).toMatchObject({ revision: 7, filter: null, sort: null })
    expect(workbook.tiles()).toBeTypeOf('function')
  })

  it('isRemoteWorkbook is gone', () => {
    const hits: string[] = []
    const walk = (dir: string) => {
      for (const name of readdirSync(dir)) {
        const path = join(dir, name)
        if (statSync(path).isDirectory()) walk(path)
        else if (/\.(ts|tsx)$/.test(name) && readFileSync(path, 'utf8').includes('isRemoteWorkbook')) hits.push(path)
      }
    }
    walk(join(__dirname, '..', 'src'))
    expect(hits).toEqual([])
  })
})
