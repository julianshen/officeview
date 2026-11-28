/**
 * Stage 4c Task 10 (P15, P16): the sort and filter menus on a `sheetSort` stream. The viewer mounts the column header
 * buttons and the filter dropdown over `streamControls`, which speaks the server's `range-detect`, `filter-values`,
 * `sheet-sort` and `sheet-filter` through `StreamWorkbook`; the fake `RemoteSortOps` below answers as Task 5's server
 * does for the fixtures named in each test (a 40-row grid A1:F40 with a header row; `filterAutoHidden`).
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor, within } from '@testing-library/react'
import RemoteDocView from '../src/components/remote/RemoteDocView'
import { StreamSessionError } from '../src/stream/remote/errors'
import type { RemoteOfficeDoc, RemoteSheets, RemoteSortOps } from '../src/stream/remote/remote-doc'
import type { TilesRequest, TilesResult } from '../src/stream/remote/sheet-channel'
import type { SheetOrder, StreamSheetLayout, ValuesPage, WireRange } from '../src/stream/sheet-wire'
import type { OperationRefusal } from '../src/xlsx/view/contracts'
import { cellId, cellIdToCoordinate } from '../src/xlsx/view/coordinates'
import { failureText } from '../src/components/xlsx/controlUi'

const spies = vi.hoisted(() => ({ valueIndex: 0, valueIndexCache: 0, documentSession: 0 }))
vi.mock('../src/xlsx/view/valueIndex', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/xlsx/view/valueIndex')>()
  class ValueIndex extends actual.ValueIndex { constructor(...args: ConstructorParameters<typeof actual.ValueIndex>) { super(...args); spies.valueIndex++ } }
  class ValueIndexCache extends actual.ValueIndexCache { constructor(...args: ConstructorParameters<typeof actual.ValueIndexCache>) { super(...args); spies.valueIndexCache++ } }
  return { ...actual, ValueIndex, ValueIndexCache }
})
vi.mock('../src/worker/viewerSession', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/worker/viewerSession')>()
  class DocumentSession extends actual.DocumentSession { constructor(...args: ConstructorParameters<typeof actual.DocumentSession>) { super(...args); spies.documentSession++ } }
  return { ...actual, DocumentSession }
})

/** The grid's detected range: A1:F40, one header row. */
const RANGE: WireRange = { r0: 0, r1: 39, c0: 0, c1: 5, header: 1, totals: 0 }
/** `filterAutoHidden`: A1:C20 with header; rows 4 and 7 (1-based) saved hidden. */
const AUTO: WireRange = { r0: 0, r1: 19, c0: 0, c1: 2, header: 1, totals: 0 }

function wire(sheetId: string, over: Partial<StreamSheetLayout> = {}): StreamSheetLayout {
  return {
    sheetId, viewGen: 0, defaults: { rowPx: 20, colPx: 64 }, rows: [], cols: [], collapsedRows: [], collapsedCols: [],
    merges: [], filters: [], sort: null, filterSummary: [], authorHidden: false, counts: { visibleRows: 40, totalRows: 40 },
    used: { r0: 0, c0: 0, r1: 39, c1: 5 }, ...over,
  } as StreamSheetLayout
}
/** An order over `range`'s body: `source[i]` is the source row at display row `range.r0 + header + i`. */
function order(viewGen: number, range: WireRange, source: number[] | null, hidden: number[] = [], head: Partial<SheetOrder['head']> = {}): SheetOrder {
  const r0 = range.r0 + range.header, r1 = range.r1 - range.totals
  const words = new Uint32Array(Math.ceil((r1 - r0 + 1) / 32))
  for (let i = 0; i <= r1 - r0; i++) if (!hidden.includes(r0 + i)) words[i >>> 5] = (words[i >>> 5]! | (1 << (i & 31))) >>> 0
  return {
    head: { sheetId: '0', viewGen, counts: { visible: r1 - r0 + 1 - hidden.length, total: r1 - r0 + 1 }, sort: null, filterSummary: [], ...head },
    bands: source ? [{ r0, r1, sourceAtOffset: Uint32Array.from(source) }] : [],
    mask: { r0, r1, words },
  }
}
/** Ascending by B: the grid's B holds 40 − r, so source row 39 comes first. */
const ASC_B = Array.from({ length: 39 }, (_, i) => 39 - i)
const ITEMS = ['apple', 'banana', 'cherry', 'date']
const page = (over: Partial<ValuesPage> = {}): ValuesPage => ({
  list: 'L1', items: ITEMS.map((text, id) => ({ id, text, selected: true, error: false })), blanks: { present: true, selected: true }, total: ITEMS.length, truncated: false,
  classification: { kind: 'text', hasNumeric: false }, ...over,
})

type Ops = { -readonly [K in keyof RemoteSortOps]: RemoteSortOps[K] }
function fakeDoc(o: { layout?: (sheetId: string) => StreamSheetLayout; ops?: Partial<Ops>; detect?: WireRange } = {}) {
  const calls: string[] = []
  let inFlight = 0, maxInFlight = 0
  const track = <A extends unknown[], R>(name: string, run: (...args: A) => Promise<R>) => vi.fn(async (...args: A) => {
    calls.push(name); inFlight++; maxInFlight = Math.max(maxInFlight, inFlight)
    try { await new Promise(r => setTimeout(r, 5)); return await run(...args) } finally { inFlight-- }
  })
  const ops: Ops = {
    typedFilters: false,
    sort: track('sort', async (_s: string, viewGen: number) => order(viewGen + 1, RANGE, ASC_B, [], { sort: { range: RANGE, keys: [{ col: 1, order: 'asc' }] } })),
    filter: track('filter', async (_s: string, viewGen: number) => order(viewGen + 1, RANGE, null, [5], { filterSummary: [{ col: 3, kind: 'values' }], filterRange: RANGE })),
    values: track('values', async () => page()),
    detect: track('detect', async () => ({ range: o.detect ?? RANGE, source: 'region' as const, header: 'heuristic' as const, overrideIgnored: false })),
    order: track('order', async () => order(1, RANGE, null)),
    ...o.ops,
  }
  const list = [{ sheetId: '0', name: 'First' }, { sheetId: '1', name: 'Second' }]
  const manifest = {
    v: 1, docRef: 'doc-ref:sheet-sort', name: 'book.xlsx', kind: 'xlsx', units: { count: 2, list: [{ w: 300, h: 200 }, { w: 300, h: 200 }] },
    can: { search: true, copy: true, print: false, textA11y: false, present: false, notes: false, links: false },
    caps: { copy: null, a11y: 'off', print: { maxPages: null, maxSessionPages: null } },
    recorder: { scalesBuckets: [1, 2] }, notices: [],
    sheetMode: 'tiles', sheets: list, sheetView: 1, sheetSort: 1,
    tiles: { px: 512, zoomSteps: [0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4], dprs: [1, 2], palettes: ['light', 'dark'], maxBatch: 64 },
  }
  const sheets = {
    list,
    layout: vi.fn(async (sheetId: string) => (o.layout ?? wire)(sheetId)),
    tiles: vi.fn((_req: TilesRequest, _onTile: unknown, signal?: AbortSignal): Promise<TilesResult> => new Promise(resolve => signal?.addEventListener('abort', () => resolve({ refused: [], cancelled: true }), { once: true }))),
    view: { find: vi.fn(async () => ({ hits: [], total: 0, truncated: false })), copy: vi.fn(async () => ''), a11y: vi.fn(async () => []), fit: vi.fn(async () => 64), resize: vi.fn() },
    sortOps: ops,
  }
  const doc = {
    kind: 'remote', manifest, onManifestChange: vi.fn(() => () => {}),
    units: async () => [], paint: vi.fn(async () => {}), search: vi.fn(async () => ({ total: 0, truncated: false })), select: vi.fn(async () => ({ units: [], paragraphs: 0 })),
    copy: vi.fn(async () => ''), a11y: vi.fn(async () => []), print: vi.fn(async () => []), paintNotes: vi.fn(async () => {}), close: vi.fn(), ended: new Promise(() => {}),
    sheets: sheets as unknown as RemoteSheets,
  }
  return { doc, ops, calls, maxInFlight: () => maxInFlight }
}
const view = (doc: ReturnType<typeof fakeDoc>['doc']) => render(<RemoteDocView document={doc as unknown as RemoteOfficeDoc} />)
type View = ReturnType<typeof view>
const filterButton = (v: View, letter: string) => v.findByRole('button', { name: new RegExp(`^Filter column ${letter} `) }, { timeout: 20_000 })
/** Opens column `letter`'s filter dropdown and waits for its value list. */
async function openFilter(v: View, letter: string) {
  const button = await filterButton(v, letter)
  await waitFor(() => expect(v.container.querySelectorAll('[role="gridcell"]').length).toBeGreaterThan(10), { timeout: 10_000 })
  fireEvent.click(button)
  const dialog = await v.findByRole('dialog', { name: 'Column filter' }, { timeout: 10_000 })
  await waitFor(() => expect(within(dialog).getByRole('button', { name: 'Sort ascending' }).hasAttribute('disabled')).toBe(false), { timeout: 10_000 })
  return dialog
}
/** The source rows of the gridcells in the row at display slot `slot`. */
const sourceRowsAt = (v: View, slot: number) => [...v.container.querySelectorAll(`[role="row"][aria-rowindex="${slot + 2}"] [role="gridcell"]`)].map(cell => Number(cell.id.split('-cell-')[1]))
const statusText = (v: View) => v.container.textContent ?? ''
const occurrences = (text: string, needle: string) => text.split(needle).length - 1

beforeEach(() => { spies.valueIndex = 0; spies.valueIndexCache = 0; spies.documentSession = 0 })
afterEach(() => { cleanup(); vi.restoreAllMocks() })

describe('stream sort and filter menus (P16)', { timeout: 30_000 }, () => {
  it('sort from the column menu', async () => {
    const { doc, ops } = fakeDoc()
    const v = view(doc)
    const dialog = await openFilter(v, 'B')
    fireEvent.click(within(dialog).getByRole('button', { name: 'Sort ascending' }))
    await waitFor(() => expect(ops.sort).toHaveBeenCalledTimes(1))
    expect(vi.mocked(ops.sort).mock.calls[0]!.slice(0, 3)).toEqual(['0', 0, { range: RANGE, keys: [{ col: 1, order: 'asc' }], caseSensitive: false }])
    await waitFor(() => expect(sourceRowsAt(v, 1)[0]).toBe(cellId(39, 0)))
    expect(new Set(sourceRowsAt(v, 1).map(id => cellIdToCoordinate(id).row))).toEqual(new Set([39]))
    await waitFor(() => expect(statusText(v)).toContain('View-only sort applied; formulas unchanged.'))
  })

  it('filter by values', async () => {
    // A refetched view: the server holds a values filter on C, which the client knows by its summary only.
    const summary = [{ col: 2, kind: 'values' }]
    const { doc, ops } = fakeDoc({
      layout: (id) => wire(id, { viewGen: 1, filterSummary: summary }),
      ops: { order: vi.fn(async () => order(1, RANGE, null, [7], { filterSummary: summary, filterRange: RANGE })) },
    })
    const v = view(doc)
    const dialog = await openFilter(v, 'D')
    await waitFor(() => expect(within(dialog).getByRole('checkbox', { name: 'cherry' })).toBeTruthy())
    for (const text of ITEMS) expect(within(dialog).getByRole('checkbox', { name: text })).toBeTruthy()
    fireEvent.click(within(dialog).getByRole('checkbox', { name: 'banana' }))
    await waitFor(() => expect((within(dialog).getByRole('checkbox', { name: 'banana' }) as HTMLInputElement).checked).toBe(false))
    fireEvent.click(within(dialog).getByRole('checkbox', { name: 'cherry' }))
    await waitFor(() => expect((within(dialog).getByRole('checkbox', { name: 'cherry' }) as HTMLInputElement).checked).toBe(false))
    await waitFor(() => expect(within(dialog).getByRole('button', { name: 'OK' }).hasAttribute('disabled')).toBe(false))
    fireEvent.click(within(dialog).getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(ops.filter).toHaveBeenCalledTimes(1))
    expect(vi.mocked(ops.filter).mock.calls[0]!.slice(0, 3)).toEqual(['0', 1, {
      range: RANGE, columns: [[2, { kind: 'committed' }], [3, { kind: 'values', list: 'L1', mode: 'exclude', ids: [[1, 2]], blanks: true }]],
    }])
    await waitFor(() => expect(statusText(v)).toContain('Filter applied.'))
  })

  it('typed conditions greyed', async () => {
    const number = page({ classification: { kind: 'number', hasNumeric: true } })
    for (const typedFilters of [false, true]) {
      const { doc } = fakeDoc({ ops: { typedFilters, values: vi.fn(async () => number) } })
      const v = view(doc)
      const dialog = await openFilter(v, 'B')
      const option = await within(dialog).findByRole('option', { name: 'Number filters' })
      expect([typedFilters, option.hasAttribute('disabled')]).toEqual([typedFilters, !typedFilters])
      expect(option.getAttribute('title')).toBe(typedFilters ? null : "This filter isn't available in this viewing mode.")
      cleanup()
    }
    // A text/mixed column with numbers: the custom panel offers text conditions only.
    const mixed = page({ classification: { kind: 'text', hasNumeric: true } })
    for (const typedFilters of [false, true]) {
      const { doc } = fakeDoc({ ops: { typedFilters, values: vi.fn(async () => mixed) } })
      const v = view(doc)
      const dialog = await openFilter(v, 'B')
      fireEvent.change(within(dialog).getByRole('combobox', { name: 'Filter method' }), { target: { value: 'custom' } })
      const type = await within(dialog).findByRole('combobox', { name: 'Condition 1 type' })
      expect(within(type).queryAllByRole('option').map(option => option.textContent)).toEqual(typedFilters ? ['Text', 'Number'] : ['Text'])
      cleanup()
    }
  })

  it('refusals speak local strings', async () => {
    const merge: OperationRefusal = { reason: 'merge', range: { r0: 2, r1: 3, c0: 1, c1: 1 } }
    const outside: OperationRefusal = { reason: 'outside-column-data', range: { r0: 0, r1: 39, c0: 0, c1: 0 } }
    let detects = 0
    const { doc, ops } = fakeDoc({
      ops: {
        sort: vi.fn(async () => { throw new StreamSessionError('view-refused', 'view-refused: merge', { refusal: merge }) }),
        filter: vi.fn(async () => { throw new StreamSessionError('stale-filter-values') }),
        detect: vi.fn(async () => {
          // Task 1: the second header sort is refused outside-column-data (a sheet with withheld lines).
          if (++detects === 2) throw new StreamSessionError('view-refused', 'view-refused: outside-column-data', { refusal: outside })
          return { range: RANGE, source: 'region' as const, header: 'heuristic' as const, overrideIgnored: false }
        }),
      },
    })
    const v = view(doc)
    fireEvent.click(await v.findByRole('button', { name: /^Cycle sort column B/ }, { timeout: 20_000 }))
    await waitFor(() => expect(statusText(v)).toContain(failureText({ kind: 'refused', refusal: merge })))
    fireEvent.click(v.getByRole('button', { name: /^Cycle sort column B/ }))
    await waitFor(() => expect(statusText(v)).toContain(failureText({ kind: 'refused', refusal: outside })))
    expect(ops.sort).toHaveBeenCalledTimes(1)
    // A values filter whose list went stale: the list is fetched again and the dropdown says so.
    const dialog = await openFilter(v, 'D')
    await waitFor(() => expect(within(dialog).getByRole('checkbox', { name: 'banana' })).toBeTruthy())
    fireEvent.click(within(dialog).getByRole('checkbox', { name: 'banana' }))
    await waitFor(() => expect((within(dialog).getByRole('checkbox', { name: 'banana' }) as HTMLInputElement).checked).toBe(false))
    const listed = vi.mocked(ops.values).mock.calls.length
    fireEvent.click(within(dialog).getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(statusText(v)).toContain('Values changed; review the list and try again.'))
    expect(vi.mocked(ops.values).mock.calls.length).toBe(listed + 1)
    // The new list is shown with nothing deselected.
    await waitFor(() => expect((within(dialog).getByRole('checkbox', { name: 'banana' }) as HTMLInputElement).checked).toBe(true))
  })

  it('pacing', async () => {
    const f = fakeDoc()
    const v = view(f.doc)
    const dialog = await openFilter(v, 'B')
    await waitFor(() => expect(within(dialog).getByRole('checkbox', { name: 'apple' })).toBeTruthy())
    expect(f.calls).toEqual(['detect', 'values'])
    expect(f.maxInFlight()).toBe(1)
    const search = within(dialog).getByRole('textbox', { name: 'Search all values' })
    const before = vi.mocked(f.ops.values).mock.calls.length
    let lastKey = 0
    for (const text of ['a', 'ab', 'abc']) { fireEvent.change(search, { target: { value: text } }); lastKey = performance.now(); await act(async () => { await new Promise(r => setTimeout(r, 40)) }) }
    let sentAt = 0
    vi.mocked(f.ops.values).mockImplementation(async () => { sentAt = performance.now(); return page({ items: [{ id: 0, text: 'abc', selected: true, error: false }] }) })
    await waitFor(() => expect(vi.mocked(f.ops.values).mock.calls.length).toBe(before + 1), { timeout: 2_000 })
    await act(async () => { await new Promise(r => setTimeout(r, 400)) })
    const searches = vi.mocked(f.ops.values).mock.calls.slice(before)
    expect(searches).toHaveLength(1)
    expect(searches[0]![0]).toMatchObject({ sheetId: '0', viewGen: 0, range: RANGE, col: 1, search: { text: 'abc', wildcards: false } })
    expect(sentAt - lastKey).toBeGreaterThanOrEqual(245)
    expect(f.maxInFlight()).toBe(1)
  })

  it('author-hidden notice', async () => {
    const NOTICE = "Some rows are hidden by the author and can't be shown."
    const hidden = { collapsedRows: [{ start: 3, end: 3 }, { start: 6, end: 6 }], authorHidden: true, counts: { visibleRows: 18, totalRows: 20 }, used: { r0: 0, c0: 0, r1: 19, c1: 2 } }
    const { doc, ops } = fakeDoc({
      layout: (id) => wire(id, hidden), detect: AUTO,
      ops: {
        // C1 under visible-only: the authored-hidden rows 4 and 7 stay hidden under the filter, and once it is cleared.
        filter: vi.fn(async (_s: string, viewGen: number, spec) => spec!.columns.length
          ? order(viewGen + 1, AUTO, null, [3, 6, 10], { filterSummary: [{ col: 1, kind: 'values' }], filterRange: AUTO })
          : order(viewGen + 1, AUTO, null, [3, 6])),
      },
    })
    const v = view(doc)
    await filterButton(v, 'B')
    await waitFor(() => expect(occurrences(statusText(v), NOTICE)).toBe(1))
    let dialog = await openFilter(v, 'B')
    await waitFor(() => expect(within(dialog).getByRole('checkbox', { name: 'banana' })).toBeTruthy())
    fireEvent.click(within(dialog).getByRole('checkbox', { name: 'banana' }))
    await waitFor(() => expect((within(dialog).getByRole('checkbox', { name: 'banana' }) as HTMLInputElement).checked).toBe(false))
    fireEvent.click(within(dialog).getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(statusText(v)).toContain('Filter applied.'))
    const rendered = (row: number) => v.container.querySelector(`[id$="-cell-${cellId(row, 0)}"]`) !== null
    await waitFor(() => expect(rendered(10)).toBe(false))
    expect(occurrences(statusText(v), NOTICE)).toBe(1)
    dialog = await openFilter(v, 'B')
    fireEvent.click(within(dialog).getByRole('button', { name: 'Clear all filters' }))
    await waitFor(() => expect(ops.filter).toHaveBeenCalledTimes(2))
    expect(vi.mocked(ops.filter).mock.calls[1]![2]).toEqual({ range: AUTO, columns: [] })
    await waitFor(() => expect(rendered(10)).toBe(true))
    expect([rendered(2), rendered(3), rendered(4), rendered(6), rendered(7)]).toEqual([true, false, true, false, true])
    expect(occurrences(statusText(v), NOTICE)).toBe(1)
  })

  it("a 'full' filter that shows an authored-hidden row keeps it shown across a same-viewGen refetch", async () => {
    // Under 'full' the filter decides alone: source row 3 (collapsed in the layout) is shown by the server's mask. A
    // refused sort re-fetches the layout and the order at the same viewGen; the server's mask is kept as sent.
    const collapsed = { collapsedRows: [{ start: 3, end: 3 }] }
    let layout = wire('0', collapsed)
    const merge: OperationRefusal = { reason: 'merge', range: { r0: 2, r1: 3, c0: 1, c1: 1 } }
    const { doc, ops } = fakeDoc({
      layout: () => layout,
      ops: {
        typedFilters: true,
        filter: vi.fn(async (_s: string, viewGen: number) => {
          layout = wire('0', { ...collapsed, viewGen: viewGen + 1, filterSummary: [{ col: 1, kind: 'values' }] })
          return order(viewGen + 1, RANGE, null, [10], { filterSummary: [{ col: 1, kind: 'values' }], filterRange: RANGE })
        }),
        order: vi.fn(async () => order(1, RANGE, null, [10], { filterSummary: [{ col: 1, kind: 'values' }], filterRange: RANGE })),
        sort: vi.fn(async () => { throw new StreamSessionError('view-refused', 'view-refused: merge', { refusal: merge }) }),
      },
    })
    const v = view(doc)
    const rendered = (row: number) => v.container.querySelector(`[id$="-cell-${cellId(row, 0)}"]`) !== null
    const dialog = await openFilter(v, 'B')
    expect(rendered(3)).toBe(false)
    await waitFor(() => expect(within(dialog).getByRole('checkbox', { name: 'banana' })).toBeTruthy())
    fireEvent.click(within(dialog).getByRole('checkbox', { name: 'banana' }))
    await waitFor(() => expect((within(dialog).getByRole('checkbox', { name: 'banana' }) as HTMLInputElement).checked).toBe(false))
    fireEvent.click(within(dialog).getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(rendered(10)).toBe(false))
    expect(rendered(3)).toBe(true)
    fireEvent.click(v.getByRole('button', { name: /^Cycle sort column B/ }))
    await waitFor(() => expect(statusText(v)).toContain(failureText({ kind: 'refused', refusal: merge })))
    await waitFor(() => expect(ops.order).toHaveBeenCalledTimes(1))
    await act(async () => { await new Promise(r => setTimeout(r, 50)) })
    expect([rendered(3), rendered(10)]).toEqual([true, false])
  })

  it('no local value index', async () => {
    const { doc, ops } = fakeDoc()
    const v = view(doc)
    const dialog = await openFilter(v, 'B')
    await waitFor(() => expect(within(dialog).getByRole('checkbox', { name: 'apple' })).toBeTruthy())
    fireEvent.click(within(dialog).getByRole('checkbox', { name: 'apple' }))
    await waitFor(() => expect((within(dialog).getByRole('checkbox', { name: 'apple' }) as HTMLInputElement).checked).toBe(false))
    fireEvent.click(within(dialog).getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(ops.filter).toHaveBeenCalledTimes(1))
    expect(vi.mocked(ops.filter).mock.calls[0]![2]).toEqual({ range: RANGE, columns: [[1, { kind: 'values', list: 'L1', mode: 'exclude', ids: [[0, 0]], blanks: true }]] })
    expect(spies).toEqual({ valueIndex: 0, valueIndexCache: 0, documentSession: 0 })
  })

  it('the spies see a local controls port (positive control: a 4b stream still builds one)', async () => {
    const { doc } = fakeDoc()
    ;(doc.manifest as Record<string, unknown>).sheetSort = undefined
    delete (doc.sheets as { sortOps?: unknown }).sortOps
    const v = view(doc)
    await v.findByRole('grid', {}, { timeout: 20_000 })
    await waitFor(() => expect(spies.valueIndexCache).toBeGreaterThan(0))
    expect(v.queryByRole('button', { name: /^Filter column/ })).toBeNull()
  })
})
