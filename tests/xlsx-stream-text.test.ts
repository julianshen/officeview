/**
 * Stage 4b Task 9: `StreamWorkbook`'s session sizes, viewGen and text methods, and `streamText.ts` (the selection's
 * wire form and the stream find session). No local text index: every answer comes from `RemoteSheets.view`.
 */
import { describe, expect, it, vi } from 'vitest'
import { StreamWorkbook } from '../src/components/xlsx/streamWorkbook'
import { streamFindSession, wireSelection } from '../src/components/xlsx/streamText'
import { FindError } from '../src/xlsx/view/find'
import type { RemoteSheets } from '../src/stream/remote/remote-doc'
import { StreamSessionError } from '../src/stream/remote/errors'
import type { SheetFindHit, StreamSheetLayout } from '../src/stream/sheet-wire'
import type { GridSelection, WorkbookOwner } from '../src/xlsx/view/contracts'
import { streamSheetGeometry } from '../src/xlsx/view/streamGeometry'

const LIST = [{ sheetId: '0', name: 'Data' }, { sheetId: '1', name: 'Summary' }]
const wire = (sheetId: string, over: Partial<StreamSheetLayout> = {}): StreamSheetLayout => ({
  sheetId, viewGen: 0, defaults: { rowPx: 20, colPx: 64 }, rows: [], cols: [], collapsedRows: [], collapsedCols: [],
  merges: [], filters: [], sort: null, filterSummary: [], authorHidden: false, counts: { visibleRows: 10, totalRows: 10 }, ...over,
})
type View = NonNullable<RemoteSheets['view']>
function fakeSheets(o: { layout?: (sheetId: string) => StreamSheetLayout; view?: Partial<View> | null } = {}) {
  const view: View = {
    find: vi.fn(async () => ({ hits: [], total: 0, truncated: false })),
    copy: vi.fn(async () => 'tsv'),
    a11y: vi.fn(async () => []),
    fit: vi.fn(async () => 50),
    resize: vi.fn(async (r) => ({ sheetId: r.sheetId, viewGen: r.viewGen + 1, change: 'reset' in r ? { reset: true as const } : { axis: r.axis, index: r.index, size: r.size } })),
    ...(o.view ?? {}),
  }
  const sheets = {
    list: LIST,
    layout: vi.fn(async (sheetId: string) => (o.layout ?? wire)(sheetId)),
    tiles: vi.fn(async () => ({ refused: [] })),
    ...(o.view === null ? {} : { view }),
  }
  return { sheets: sheets as RemoteSheets, layout: sheets.layout, view }
}
const OWNER: WorkbookOwner = { docId: 'doc', globalViewGen: 0, workbookValueGen: 0, fontGen: 0, sheetVector: [] }
const errorOf = (p: Promise<unknown>) => p.then(() => { throw new Error('resolved') }, (e: unknown) => e as StreamSessionError)
const cellSelection = (ranges: GridSelection['ranges']): GridSelection => ({ active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 }, ranges })

describe('StreamWorkbook session sizes and viewGen (P11)', () => {
  it('applyView moves the geometry', async () => {
    const { sheets } = fakeSheets()
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    const notified = vi.fn()
    wb.subscribe(notified)
    wb.applyView({ sheetId: '0', viewGen: 1, change: { axis: 'col', index: 1, size: 8 } })
    expect(wb.geometry('0').cols.size(1)).toBe(8)
    expect(streamSheetGeometry(wb.wireLayout('0')).metrics.cols.size(1)).toBe(8)
    expect(wb.identity('0').viewGen).toBe(1)
    expect(wb.view('0').revision).toBe(1)
    expect(wb.sizes('0').cols.get(1)).toBe(8)
    expect(wb.wireLayout('0').sessionCols).toEqual([[1, 8]])
    expect(notified).toHaveBeenCalledTimes(1)
    // Stable between changes.
    expect(wb.sizes('0')).toBe(wb.sizes('0'))
    wb.applyView({ sheetId: '0', viewGen: 2, change: { axis: 'row', index: 4, size: 33 } })
    wb.applyView({ sheetId: '0', viewGen: 3, change: { axis: 'col', index: 0, size: 12 } })
    expect(wb.wireLayout('0').sessionCols).toEqual([[0, 12], [1, 8]])
    expect(wb.geometry('0').rows.size(4)).toBe(33)
    // An old change never rolls the view back.
    wb.applyView({ sheetId: '0', viewGen: 1, change: { axis: 'col', index: 1, size: 99 } })
    expect([wb.identity('0').viewGen, wb.geometry('0').cols.size(1)]).toEqual([3, 8])
    wb.applyView({ sheetId: '0', viewGen: 4, change: { reset: true } })
    expect(wb.geometry('0').cols.size(1)).toBe(64)
    expect(wb.wireLayout('0').sessionCols).toBeUndefined()
    expect(wb.wireLayout('0').sessionRows).toBeUndefined()
    expect(wb.sizes('0').cols.size + wb.sizes('0').rows.size).toBe(0)
    expect(wb.identity('0').viewGen).toBe(4)
  })

  it('the layout session sizes reach sizes() and geometry()', async () => {
    const { sheets } = fakeSheets({ layout: (id) => wire(id, { viewGen: 2, sessionRows: [[3, 40]] }) })
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    expect(wb.sizes('0').rows.get(3)).toBe(40)
    expect(wb.geometry('0').rows.size(3)).toBe(40)
  })

  it('textOps is true iff the view is present', () => {
    expect(new StreamWorkbook(fakeSheets().sheets, 'doc').textOps).toBe(true)
    expect(new StreamWorkbook(fakeSheets({ view: null }).sheets, 'doc').textOps).toBe(false)
  })
})

describe('StreamWorkbook text methods (P7, P9–P13)', () => {
  it('each method fills in the current viewGen; a11y passes its own through', async () => {
    const { sheets, view } = fakeSheets({ layout: (id) => wire(id, { viewGen: 7 }) })
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    expect(await wb.copy('0', cellSelection([{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 2 }] }, cols: [{ c0: 0, c1: 1 }] }]))).toBe('tsv')
    expect(view.copy).toHaveBeenCalledWith('0', 7, { active: [0, 0], anchor: [0, 0], ranges: [{ rows: [[0, 2]], cols: [[0, 1]] }] })
    expect(await wb.fit({ sheetId: '0', axis: 'col', index: 2, pane: { x: 0, y: 0, width: 10, height: 10 } })).toBe(50)
    expect(view.fit).toHaveBeenCalledWith({ sheetId: '0', viewGen: 7, axis: 'col', index: 2, pane: { x: 0, y: 0, width: 10, height: 10 } })
    await wb.a11y('0', 5, [[0, 0]], [0, 3])
    expect(view.a11y).toHaveBeenCalledWith('0', 5, [[0, 0]], [0, 3])
    const change = await wb.resize({ sheetId: '0', axis: 'col', index: 2, size: 90 })
    expect(view.resize).toHaveBeenCalledWith({ sheetId: '0', viewGen: 7, axis: 'col', index: 2, size: 90 })
    expect(change).toEqual({ sheetId: '0', viewGen: 8, change: { axis: 'col', index: 2, size: 90 } })
    // The applied change is held: the next request carries viewGen 8.
    expect([wb.identity('0').viewGen, wb.geometry('0').cols.size(2)]).toEqual([8, 90])
    await wb.resize({ sheetId: '0', reset: true })
    expect(view.resize).toHaveBeenLastCalledWith({ sheetId: '0', viewGen: 8, reset: true })
  })

  it('a selection over the wire caps rejects copy-too-large with no limit, without sending', async () => {
    const { sheets, view } = fakeSheets()
    const wb = new StreamWorkbook(sheets, 'doc')
    const parts = Array.from({ length: 257 }, (_, i) => ({ rows: { kind: 'intervals' as const, intervals: [{ r0: i * 2, r1: i * 2 }] }, cols: [{ c0: 0, c1: 0 }] }))
    const e = await errorOf(wb.copy('0', cellSelection(parts)))
    expect(e).toBeInstanceOf(StreamSessionError)
    expect([e.code, e.limit, e.requested]).toEqual(['copy-too-large', undefined, undefined])
    expect(view.copy).not.toHaveBeenCalled()
  })

  it('a rejected resize re-fetches sheet-layout', async () => {
    let gen = 0
    const { sheets, layout } = fakeSheets({
      layout: (id) => wire(id, { viewGen: gen, ...(gen ? { sessionCols: [[1, 77]] as Array<[number, number]> } : {}) }),
      view: { resize: vi.fn(async () => { throw new StreamSessionError('session-rotated') }) },
    })
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    expect(layout).toHaveBeenCalledTimes(1)
    gen = 5 // the server applied the lost resize
    const e = await errorOf(wb.resize({ sheetId: '0', axis: 'col', index: 1, size: 77 }))
    expect(e.code).toBe('session-rotated')
    expect(layout).toHaveBeenCalledTimes(2)
    expect(wb.identity('0').viewGen).toBe(5)
    expect(wb.geometry('0').cols.size(1)).toBe(77)
  })

  it('stale-view on copy, a11y or fit re-fetches sheet-layout and rejects stale-view', async () => {
    let gen = 1
    const stale = async () => { throw new StreamSessionError('stale-view') }
    const { sheets, layout } = fakeSheets({ layout: (id) => wire(id, { viewGen: gen }), view: { copy: vi.fn(stale), a11y: vi.fn(stale), fit: vi.fn(stale), resize: vi.fn(stale) } })
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    const one = cellSelection([{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, cols: [{ c0: 0, c1: 0 }] }])
    for (const [i, call] of ([
      () => wb.copy('0', one),
      () => wb.a11y('0', 1, [[0, 0]], [0, 0]),
      () => wb.fit({ sheetId: '0', axis: 'row', index: 0, pane: { x: 0, y: 0, width: 1, height: 1 } }),
      () => wb.resize({ sheetId: '0', axis: 'row', index: 0, size: 30 }),
    ] as const).entries()) {
      gen = 10 + i
      expect((await errorOf(call())).code).toBe('stale-view')
      expect(layout).toHaveBeenCalledTimes(2 + i)
      expect(wb.identity('0').viewGen).toBe(10 + i)
    }
  })

  it('a change at the held viewGen is a no-op (the server answers a no-op resize with the current viewGen)', async () => {
    const { sheets } = fakeSheets({ layout: (id) => wire(id, { viewGen: 4 }) })
    const wb = new StreamWorkbook(sheets, 'doc')
    await wb.activate('0')
    const notified = vi.fn()
    wb.subscribe(notified)
    wb.applyView({ sheetId: '0', viewGen: 4, change: { axis: 'col', index: 1, size: 8 } })
    expect([wb.sizes('0').cols.size, wb.geometry('0').cols.size(1), notified.mock.calls.length]).toEqual([0, 64, 0])
  })

  it('a layout older than the held view is not adopted', async () => {
    let release!: (l: StreamSheetLayout) => void
    const { sheets } = fakeSheets()
    ;(sheets.layout as ReturnType<typeof vi.fn>).mockImplementationOnce(() => new Promise<StreamSheetLayout>((r) => { release = r }))
    const wb = new StreamWorkbook(sheets, 'doc')
    const loading = wb.activate('0')
    wb.applyView({ sheetId: '0', viewGen: 2, change: { axis: 'col', index: 1, size: 8 } })
    release(wire('0', { viewGen: 1 }))
    await loading
    expect([wb.identity('0').viewGen, wb.geometry('0').cols.size(1)]).toEqual([2, 8])
  })

  it('no text methods without the view', async () => {
    const wb = new StreamWorkbook(fakeSheets({ view: null }).sheets, 'doc')
    const one = cellSelection([{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, cols: [{ c0: 0, c1: 0 }] }])
    expect((await errorOf(wb.copy('0', one))).code).toBe('unsupported-op')
    expect((await errorOf(wb.resize({ sheetId: '0', reset: true }))).code).toBe('unsupported-op')
  })
})

describe('wireSelection (P7)', () => {
  it('wireSelection', () => {
    const words = new Uint32Array(1)
    for (const r of [0, 1, 2, 5, 9]) words[0]! |= 1 << r
    const s = wireSelection({
      active: { row: 10, col: 3 }, anchor: { row: 10, col: 2 },
      ranges: [
        { rows: { kind: 'bitset', space: 'source', r0: 10, r1: 19, words, count: 5 }, cols: [{ c0: 4, c1: 6 }, { c0: 0, c1: 1 }] },
        { rows: { kind: 'list', r0: 0, r1: 9, rows: new Uint32Array([3, 4, 8]) }, cols: [{ c0: 2, c1: 2 }] },
      ],
    })
    expect(s).toEqual({
      active: [10, 3], anchor: [10, 2],
      ranges: [
        { rows: [[10, 12], [15, 15], [19, 19]], cols: [[0, 1], [4, 6]] },
        { rows: [[3, 4], [8, 8]], cols: [[2, 2]] },
      ],
    })
    const many = Array.from({ length: 257 }, (_, i) => ({ rows: { kind: 'intervals' as const, intervals: [{ r0: i, r1: i }] }, cols: [{ c0: 0, c1: 0 }] }))
    expect(wireSelection(cellSelection(many))).toBe('too-large')
    expect(wireSelection(cellSelection(many.slice(0, 256)))).not.toBe('too-large')
  })

  it('a ref part, or a selection past 64 KiB, is too large; an empty part is dropped', () => {
    expect(wireSelection(cellSelection([{ rows: { kind: 'ref', refId: 'r', count: 3, stamp: { docId: 'd', sheetId: '0', valueGen: 0 } }, cols: [{ c0: 0, c1: 0 }] }]))).toBe('too-large')
    const rows = new Uint32Array(Array.from({ length: 8000 }, (_, i) => i * 2))
    expect(wireSelection(cellSelection([{ rows: { kind: 'list', r0: 0, r1: 16000, rows }, cols: [{ c0: 0, c1: 0 }] }]))).toBe('too-large')
    expect(wireSelection({ active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [] }, cols: [{ c0: 0, c1: 0 }] }] })).toEqual({ active: null, anchor: null, ranges: [] })
  })
})

describe('streamFindSession (P6, F15)', () => {
  const hits = (n: number): SheetFindHit[] => Array.from({ length: n }, (_, i) => ({ sheetIndex: i < 300 ? 0 : 1, row: i, col: i % 7 }))

  it('stream find pages without text', async () => {
    const find = vi.fn(async (_q: { query: string; matchCase?: boolean }) => ({ hits: hits(450), total: 450, truncated: false }))
    const session = streamFindSession({ find } as never, OWNER, ['0', '1'], { text: 'q', caseSensitive: false })
    const pages = [await session.nextPage(), await session.nextPage(), await session.nextPage()]
    expect(find).toHaveBeenCalledTimes(1)
    expect(find.mock.calls[0]![0]).toEqual({ query: 'q', matchCase: false })
    expect(pages.map((p) => p.hits.length)).toEqual([200, 200, 50])
    expect(pages.map((p) => p.complete)).toEqual([false, false, true])
    expect(pages[2]!.totalKnown).toBe(450)
    expect(pages.every((p) => p.owner === OWNER && p.truncated === false)).toBe(true)
    const all = pages.flatMap((p) => p.hits)
    expect(all.every((h) => h.text === '' && h.start === 0 && h.end === 0)).toBe(true)
    expect(all[0]).toMatchObject({ sheetId: '0', source: { row: 0, col: 0 }, displaySlot: 0 })
    expect(all[449]).toMatchObject({ sheetId: '1', source: { row: 449, col: 449 % 7 }, displaySlot: 449 })
    session.close()
  })

  it('a truncated answer reaches the pages, with no exact total', async () => {
    const find = vi.fn(async (_q: { query: string; matchCase?: boolean }) => ({ hits: [{ sheetIndex: 1, row: 2, col: 3, hidden: 'filter' as const }], total: 1, truncated: true }))
    const session = streamFindSession({ find } as never, OWNER, ['0', '1'], { text: 'q', caseSensitive: true })
    const page = await session.nextPage()
    expect(find.mock.calls[0]![0]).toEqual({ query: 'q', matchCase: true })
    expect([page.complete, page.truncated, page.totalKnown]).toEqual([true, true, undefined])
    expect(page.hits[0]).toMatchObject({ sheetId: '1', hidden: 'filter' })
  })

  it('closing or aborting while the find is pending rejects FindError closed / cancelled', async () => {
    const pending = (_q: unknown, s?: AbortSignal) => new Promise<never>((_, reject) => s!.addEventListener('abort', () => reject(Object.assign(new Error('aborted'), { name: 'AbortError' }))))
    const a = streamFindSession({ find: pending } as never, OWNER, ['0'], { text: 'q', caseSensitive: false })
    const pa = a.nextPage().then(() => undefined, (e: unknown) => e as FindError)
    a.close()
    expect((await pa)).toBeInstanceOf(FindError)
    expect((await pa)!.code).toBe('closed')
    const outer = new AbortController()
    const b = streamFindSession({ find: pending } as never, OWNER, ['0'], { text: 'q', caseSensitive: false }, outer.signal)
    const pb = b.nextPage().then(() => undefined, (e: unknown) => e as FindError)
    outer.abort()
    expect((await pb)!.code).toBe('cancelled')
    // Superseded by a newer find (the session rejects `closed`): ended quietly as well.
    const c = streamFindSession({ find: async () => { throw new StreamSessionError('closed') } } as never, OWNER, ['0'], { text: 'q', caseSensitive: false })
    expect((await c.nextPage().then(() => undefined, (e: unknown) => e as FindError))!.code).toBe('closed')
    // Any other refusal is shown, as it came.
    const d = streamFindSession({ find: async () => { throw new StreamSessionError('rate-limited') } } as never, OWNER, ['0'], { text: 'q', caseSensitive: false })
    expect((await d.nextPage().then(() => undefined, (e: unknown) => e as StreamSessionError))!.code).toBe('rate-limited')
  })

  it('StreamWorkbook.findSession asks the view; closing aborts the find', async () => {
    let signal: AbortSignal | undefined
    const { sheets } = fakeSheets({ view: { find: vi.fn(async (_q, s?: AbortSignal) => { signal = s; return { hits: [], total: 0, truncated: false } }) } })
    const wb = new StreamWorkbook(sheets, 'doc')
    const session = wb.findSession(OWNER, { text: 'x', caseSensitive: false }, {})
    expect((await session.nextPage()).complete).toBe(true)
    session.close()
    expect(signal?.aborted).toBe(true)
  })
})
