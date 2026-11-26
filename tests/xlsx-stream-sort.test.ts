/**
 * Stage 4c Task 9: `StreamWorkbook`'s sort, filter, value lists, detection and the order delta (P7, P8). The order and
 * mask come from the server; the client builds its geometry from them and its layout, never from cell data.
 */
import { describe, expect, it, vi } from 'vitest'
import { StreamNoCellData, StreamWorkbook } from '../src/components/xlsx/streamWorkbook'
import type { RemoteSheets, RemoteSortOps } from '../src/stream/remote/remote-doc'
import { StreamSessionError } from '../src/stream/remote/errors'
import type { SheetOrder, StreamSheetLayout, WireRange } from '../src/stream/sheet-wire'
import type { CacheIdentity, OperationRefusal, SortSpec } from '../src/xlsx/view/contracts'
import type { XlsxSheet } from '../src/xlsx/types'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { createRowOrder } from '../src/xlsx/view/order'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { buildSheetGeometry, SHEET_DEFAULT_COL_CSS, SHEET_DEFAULT_ROW_CSS } from '../src/xlsx/view/viewport'
import { createRowVisibility } from '../src/xlsx/view/visibility'

const LIST = [{ sheetId: '0', name: 'Data' }, { sheetId: '1', name: 'Summary' }]
const RANGE: WireRange = { r0: 0, r1: 39, c0: 0, c1: 5, header: 0, totals: 0 }
const wire = (sheetId: string, over: Partial<StreamSheetLayout> = {}): StreamSheetLayout => ({
  sheetId, viewGen: 0, defaults: { rowPx: SHEET_DEFAULT_ROW_CSS, colPx: SHEET_DEFAULT_COL_CSS }, rows: [], cols: [], collapsedRows: [], collapsedCols: [],
  merges: [], filters: [], sort: null, filterSummary: [], authorHidden: false, counts: { visibleRows: 40, totalRows: 40 }, ...over,
})
const reversed = (r0: number, r1: number) => { const a = new Uint32Array(r1 - r0 + 1); for (let i = 0; i < a.length; i++) a[i] = r1 - i; return a }
/** The order the server sends: a reversed band over RANGE and a mask over it (bit i = display slot r0 + i). */
function orderAt(viewGen: number, visible: (slot: number) => boolean = () => true, over: Partial<SheetOrder['head']> = {}): SheetOrder {
  const words = new Uint32Array(Math.ceil((RANGE.r1 - RANGE.r0 + 1) / 32))
  for (let i = 0; i <= RANGE.r1 - RANGE.r0; i++) if (visible(RANGE.r0 + i)) words[i >>> 5] = (words[i >>> 5]! | (1 << (i & 31))) >>> 0
  return {
    head: { sheetId: '0', viewGen, counts: { visible: 40, total: 40 }, sort: { range: RANGE, keys: [{ col: 1, order: 'desc' }] }, filterSummary: [], ...over },
    bands: [{ r0: RANGE.r0, r1: RANGE.r1, sourceAtOffset: reversed(RANGE.r0, RANGE.r1) }],
    mask: { r0: RANGE.r0, r1: RANGE.r1, words },
  }
}
type SortOps = { -readonly [K in keyof RemoteSortOps]: RemoteSortOps[K] }
function fakeSheets(o: { layout?: (sheetId: string) => StreamSheetLayout; sortOps?: Partial<SortOps> | null } = {}) {
  const sortOps: SortOps = {
    typedFilters: false,
    sort: vi.fn(async (_s, viewGen) => orderAt(viewGen + 1)),
    filter: vi.fn(async (_s, viewGen) => orderAt(viewGen + 1)),
    values: vi.fn(async () => ({ list: 'L', items: [], blanks: { present: false, selected: true }, total: 0, truncated: false })),
    detect: vi.fn(async () => ({ range: RANGE, source: 'region' as const, header: 'heuristic' as const, overrideIgnored: false })),
    order: vi.fn(async () => orderAt(1)),
    ...(o.sortOps ?? {}),
  }
  const sheets = {
    list: LIST,
    layout: vi.fn(async (sheetId: string) => (o.layout ?? wire)(sheetId)),
    tiles: vi.fn(async () => ({ refused: [] })),
    view: { find: vi.fn(), copy: vi.fn(), a11y: vi.fn(), fit: vi.fn(), resize: vi.fn() },
    ...(o.sortOps === null ? {} : { sortOps }),
  }
  return { sheets: sheets as unknown as RemoteSheets, layout: sheets.layout, sortOps }
}
const owner = (wb: StreamWorkbook, sheetId = '0'): CacheIdentity => wb.identity(sheetId)
const SORT: SortSpec = { range: { r0: 0, r1: 39, c0: 0, c1: 5, headerRows: 0, totalsRows: 0 }, keys: [{ col: 1, order: 'desc' }], caseSensitive: false }
const never = new AbortController().signal

describe('StreamWorkbook order deltas (P7)', () => {
  it('applyOrder moves the view', async () => {
    const { sheets } = fakeSheets()
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    expect([wb.sortOps, wb.typedFilters]).toEqual([true, false])
    const notified = vi.fn()
    wb.subscribe(notified)
    wb.applyOrder(orderAt(1, (slot) => slot !== 9, { counts: { visible: 39, total: 40 }, filterSummary: [{ col: 2, kind: 'text' }] }))
    expect(wb.view('0').order.displayToSource(0)).toBe(39)
    expect(wb.view('0').revision).toBe(1)
    expect(wb.view('0').visibility!.visible(9)).toBe(false)
    expect(wb.identity('0').viewGen).toBe(1)
    expect(wb.wireLayout('0')).toMatchObject({ viewGen: 1, sort: { range: RANGE, keys: [{ col: 1, order: 'desc' }] }, filterSummary: [{ col: 2, kind: 'text' }], counts: { visibleRows: 39, totalRows: 40 } })
    expect(notified).toHaveBeenCalledTimes(1)
    // At or below the held viewGen: ignored.
    wb.applyOrder({ ...orderAt(1), bands: [] })
    expect(wb.view('0').order.displayToSource(0)).toBe(39)
    expect(notified).toHaveBeenCalledTimes(1)
  })

  it("client geometry equals the server's under the view", async () => {
    // Source row 3 is 40 px; source rows 12 and 45 are authored-hidden. Under the view the band reverses rows 0–39, a
    // 'full' filter reveals source row 12 and hides display slot 9; row 45 lies outside the band and stays hidden.
    const sheet = {
      name: 'Data', cols: [], merges: [], mergeRanges: [],
      rows: [{ index: 3, heightPt: 30, cells: [] }, { index: 12, hidden: true, cells: [] }, { index: 45, hidden: true, cells: [] }],
    } as unknown as XlsxSheet
    const layout = wire('0', { rows: [{ start: 3, end: 3, sizeCss: 40 }], collapsedRows: [{ start: 12, end: 12 }, { start: 45, end: 45 }] })
    const { sheets } = fakeSheets({ layout: () => layout })
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    const order = createRowOrder([{ r0: 0, r1: 39, sourceAtOffset: reversed(0, 39) }])
    const serverVisibility = createRowVisibility(MAX_ROWS, (slot) => slot !== 9 && order.displayToSource(slot) !== 45)
    wb.applyOrder(orderAt(1, (slot) => serverVisibility.visible(slot)))
    const view = wb.view('0')
    const rows = createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: layout.defaults.rowPx, overrides: layout.rows })
    const cols = createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: layout.defaults.colPx, overrides: layout.cols })
    const client = createSheetMetrics({ rows, cols, order: view.order, visibility: view.visibility })
    const server = buildSheetGeometry(sheet, { order, visibility: serverVisibility })
    if (server.kind !== 'ready') throw new Error('server geometry')
    for (const i of [0, 3, 9, 12, 27, 36, 39, 45, 46, 1_048_575]) expect([i, client.rows.start(i)]).toEqual([i, server.metrics.rows.start(i)])
    expect(client.rows.totalCss()).toBe(server.metrics.rows.totalCss())
  })

  it('a resize keeps the order', async () => {
    const { sheets } = fakeSheets()
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    wb.applyOrder(orderAt(1, (slot) => slot !== 9))
    const { order, visibility } = wb.view('0')
    wb.applyView({ sheetId: '0', viewGen: 2, change: { axis: 'col', index: 1, size: 80 } })
    expect(wb.view('0').revision).toBe(2)
    expect(wb.view('0').order).toBe(order)
    expect(wb.view('0').visibility).toBe(visibility)
    expect(wb.view('0').order.displayToSource(0)).toBe(39)
    expect(wb.wireLayout('0')).toMatchObject({ viewGen: 2, sort: { range: RANGE }, sessionCols: [[1, 80]] })
  })

  it('no sortOps, no sort', async () => {
    const { sheets } = fakeSheets({ sortOps: null })
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    expect([wb.sortOps, wb.typedFilters]).toEqual([false, false])
    const sort = await wb.sort(owner(wb), SORT, never).catch((e: unknown) => e)
    expect(sort).toBeInstanceOf(StreamNoCellData)
    expect((sort as StreamNoCellData).method).toBe('sort')
    const filter = await wb.filter(owner(wb), null, [], never).catch((e: unknown) => e)
    expect((filter as StreamNoCellData).method).toBe('filter')
    expect(((await wb.values('0', { range: RANGE, col: 1 }).catch((e: unknown) => e)) as StreamNoCellData).method).toBe('values')
    expect(((await wb.detect('0', 'sort', { active: [0, 0], anchor: [0, 0], ranges: [] }).catch((e: unknown) => e)) as StreamNoCellData).method).toBe('detect')
  })
})

describe('StreamWorkbook sort and filter (P8)', () => {
  it('a sort commits the order it answers, with progress 0 / 1 then 1 / 1', async () => {
    const { sheets, sortOps } = fakeSheets()
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    const progress = vi.fn()
    const result = await wb.sort(owner(wb), SORT, never, progress)
    expect(sortOps.sort).toHaveBeenCalledWith('0', 0, { range: RANGE, keys: [{ col: 1, order: 'desc' }], caseSensitive: false }, never)
    expect(progress.mock.calls).toEqual([[0, 1], [1, 1]])
    expect(result.kind).toBe('committed')
    expect(wb.view('0')).toMatchObject({ revision: 1, sort: SORT, filter: null })
    expect(wb.view('0').order.displayToSource(0)).toBe(39)
    // A filter of typed rules travels as its wire form; the committed sort stays in the view.
    const filter = await wb.filter(owner(wb), { range: SORT.range, columns: new Map([[2, { kind: 'text' as const, op: 'contains' as const, first: 'a', wildcards: false }]]) }, [], never)
    expect(filter.kind).toBe('committed')
    expect(sortOps.filter).toHaveBeenCalledWith('0', 1, { range: RANGE, columns: [[2, { kind: 'text', op: 'contains', first: 'a', wildcards: false }]] }, never)
    expect(wb.view('0')).toMatchObject({ revision: 2, sort: SORT })
    expect(wb.view('0').filter?.columns.get(2)).toEqual({ kind: 'text', op: 'contains', first: 'a', wildcards: false })
    // Clearing it sends the held range with no columns.
    await wb.filter(owner(wb), null, [], never)
    expect(sortOps.filter).toHaveBeenLastCalledWith('0', 2, { range: RANGE, columns: [] }, never)
  })

  it('a stale owner is stale without sending; a local values rule and a sort clear are refused without sending', async () => {
    const { sheets, sortOps } = fakeSheets()
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    expect((await wb.sort({ ...owner(wb), viewGen: 5 }, SORT, never)).kind).toBe('stale')
    const values = await wb.filter(owner(wb), { range: SORT.range, columns: new Map([[1, { kind: 'values' as const, selection: { mode: 'queryPatches' as const, base: 'all' as const, patches: [] }, showBlanks: true }]]) }, [], never)
    expect(values.kind).toBe('refused')
    expect((await wb.sort(owner(wb), null, never)).kind).toBe('refused')
    expect(sortOps.sort).not.toHaveBeenCalled()
    expect(sortOps.filter).not.toHaveBeenCalled()
  })

  it('a lost sort answer refetches', async () => {
    let layouts = 0
    const { sheets, layout, sortOps } = fakeSheets({
      layout: (id) => ++layouts === 1 ? wire(id) : wire(id, { viewGen: 1, sort: { range: RANGE, keys: [{ col: 1, order: 'desc' }] } }),
      sortOps: { sort: vi.fn(async () => { throw new StreamSessionError('session-rotated') }) },
    })
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    expect(sortOps.order).not.toHaveBeenCalled()
    const e = await wb.sort(owner(wb), SORT, never).catch((x: unknown) => x as StreamSessionError)
    expect((e as StreamSessionError).code).toBe('session-rotated')
    expect(layout).toHaveBeenCalledTimes(2)
    expect(sortOps.order).toHaveBeenCalledTimes(1)
    expect(layout.mock.invocationCallOrder[1]!).toBeLessThan((sortOps.order as ReturnType<typeof vi.fn>).mock.invocationCallOrder[0]!)
    expect(wb.identity('0').viewGen).toBe(1)
    expect(wb.view('0').order.displayToSource(0)).toBe(39)
    expect(wb.view('0').revision).toBe(1)
  })

  it('every refusal refetches; view-refused, stale-view and a cancel map to their outcomes', async () => {
    const merge: OperationRefusal = { reason: 'merge', range: { r0: 2, r1: 3, c0: 1, c1: 1 } }
    const errors: unknown[] = [
      Object.assign(new StreamSessionError('view-refused'), { refusal: merge }),
      new StreamSessionError('stale-view'),
      Object.assign(new Error('aborted'), { name: 'AbortError' }),
      new StreamSessionError('cancelled'),
      new StreamSessionError('policy-refused'),
    ]
    const { sheets, layout } = fakeSheets({ sortOps: { sort: vi.fn(async () => { throw errors.shift() }) } })
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    expect(await wb.sort(owner(wb), SORT, never)).toEqual({ kind: 'refused', refusal: merge })
    expect(await wb.sort(owner(wb), SORT, never)).toEqual({ kind: 'stale' })
    expect(await wb.sort(owner(wb), SORT, never)).toEqual({ kind: 'cancelled' })
    expect(await wb.sort(owner(wb), SORT, never)).toEqual({ kind: 'cancelled' })
    expect(((await wb.sort(owner(wb), SORT, never).catch((e: unknown) => e)) as StreamSessionError).code).toBe('policy-refused')
    expect(layout).toHaveBeenCalledTimes(6)
  })

  it('a layout with a view fetches its order and adopts both together', async () => {
    const { sheets, sortOps } = fakeSheets({ layout: (id) => wire(id, { viewGen: 1, sort: { range: RANGE, keys: [{ col: 1, order: 'desc' }] } }) })
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    expect(sortOps.order).toHaveBeenCalledTimes(1)
    expect([wb.identity('0').viewGen, wb.view('0').revision, wb.view('0').order.displayToSource(0)]).toEqual([1, 1, 39])
    expect(wb.view('0').sort).toEqual(SORT)
    await wb.refreshOrder('0')
    expect(sortOps.order).toHaveBeenCalledTimes(2)
  })

  it('values and detect ask at the held viewGen; stale-view refetches', async () => {
    const { sheets, sortOps, layout } = fakeSheets({ sortOps: { values: vi.fn(async () => { throw new StreamSessionError('stale-view') }) } })
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    wb.applyOrder(orderAt(3))
    const detected = await wb.detect('0', 'filter', { active: [1, 1], anchor: [1, 1], ranges: [] }, { header: 1 })
    expect(detected.source).toBe('region')
    expect(sortOps.detect).toHaveBeenCalledWith('0', 'filter', { active: [1, 1], anchor: [1, 1], ranges: [] }, { header: 1 }, undefined)
    expect(((await wb.values('0', { range: RANGE, col: 1 }).catch((e: unknown) => e)) as StreamSessionError).code).toBe('stale-view')
    expect(sortOps.values).toHaveBeenCalledWith({ sheetId: '0', viewGen: 3, range: RANGE, col: 1 }, undefined)
    expect(layout).toHaveBeenCalledTimes(2)
  })
})
