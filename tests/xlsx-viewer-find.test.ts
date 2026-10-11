/**
 * S5.3 workbook find engine: global, display-ordered, visibility-honoring,
 * bounded and lazy (spec §11.4, AC-SEARCH-01). Owner/race/navigation cases
 * live in xlsx-viewer-find-races.test.tsx.
 */
import { describe, expect, it, vi } from 'vitest'
import type { AuthoritativeCell, Page, SourceStore } from '../src/xlsx/view/contracts'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { createRowOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { initialIdentity, initialWorkbookOwner } from '../src/worker/viewerProtocol'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { LayoutCache, type CellBodyLayout, type CellLayoutHost } from '../src/xlsx/view/textLayout'
import { FIND_PAGE_HITS, WorkbookFind, hitGeometry, matchTarget } from '../src/xlsx/view/find'
import { createFixture } from './helpers/xlsx-viewer-fixtures'
import { drain, findInputOf, where, workbookOf } from './helpers/xlsx-find'

const ci = (text: string) => ({ text, caseSensitive: false })

describe('workbook find: global source scan', () => {
  const doc = () => workbookOf([
    { sheetId: 'A', cells: [{ row: 0, col: 0, value: 'Alpha' }, { row: 1, col: 0, value: 'beta' }, { row: 2, col: 1, value: 42 }, { row: 4, col: 0, value: { error: '#DIV/0!' } }, { row: 900_000, col: 16_000, value: 'far alpha' }] },
    { sheetId: 'B', cells: [{ row: 0, col: 0, value: 'alphabet' }, { row: 1, col: 0, value: 'BETA' }] },
  ])

  it('finds offscreen and other-sheet matches in workbook then display order with source ids', async () => {
    const hits = await drain(new WorkbookFind(findInputOf(doc()).input, ci('alpha')))
    expect(where(hits)).toEqual([['A', 0, 0], ['A', 900_000, 16_000], ['B', 0, 0]])
    expect(hits[1].cellId).toBe(900_000 * MAX_COLS + 16_000)
    expect(hits[0]).toMatchObject({ text: 'Alpha', start: 0, end: 5, displaySlot: 0 })
    expect(hits[1]).toMatchObject({ text: 'far alpha', start: 4, end: 9 })
  })

  it('follows the existing logical search semantics: case folds by default, exact when caseSensitive, accents kept', async () => {
    const d = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'BETA' }, { row: 1, col: 0, value: 'beta' }, { row: 2, col: 0, value: 'béta' }] }])
    expect(where(await drain(new WorkbookFind(findInputOf(d).input, ci('beta'))))).toEqual([['A', 0, 0], ['A', 1, 0]])
    expect(where(await drain(new WorkbookFind(findInputOf(d).input, { text: 'beta', caseSensitive: true })))).toEqual([['A', 1, 0]])
    expect(where(await drain(new WorkbookFind(findInputOf(d).input, ci('BÉTA'))))).toEqual([['A', 2, 0]])
  })

  it('maps highlight offsets back to raw UTF-16 when lowercasing changes the length', async () => {
    // 'İ' lowercases to two code units; the hit still names raw offsets [2,4) of 'xİab'.
    const d = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'xİab' }] }])
    const [hit] = await drain(new WorkbookFind(findInputOf(d).input, ci('ab')))
    expect(hit).toMatchObject({ start: 2, end: 4 })
  })

  it('treats an empty query as no search term and matches error and boolean display text', async () => {
    const empty = await new WorkbookFind(findInputOf(doc()).input, ci('')).nextPage()
    expect(empty).toMatchObject({ complete: true, hits: [], totalKnown: 0 })
    expect(where(await drain(new WorkbookFind(findInputOf(doc()).input, ci('#div/0!'))))).toEqual([['A', 4, 0]])
    const bools = workbookOf([{ sheetId: 'A', cells: [{ row: 3, col: 2, value: true }] }])
    expect(where(await drain(new WorkbookFind(findInputOf(bools).input, ci('true'))))).toEqual([['A', 3, 2]])
  })

  it('excludes filter-hidden and authored-hidden rows through the single visibility mask', async () => {
    const d = workbookOf([{ sheetId: 'A', hiddenRows: [2], cells: [{ row: 0, col: 0, value: 'hit' }, { row: 1, col: 0, value: 'hit' }, { row: 2, col: 0, value: 'hit' }, { row: 3, col: 0, value: 'hit' }] }])
    const hits = await drain(new WorkbookFind(findInputOf(d, { filtered: { A: [0] } }).input, ci('hit')))
    expect(where(hits)).toEqual([['A', 1, 0], ['A', 3, 0]])
  })

  it('excludes hidden columns when the sheet supplies a column predicate', async () => {
    const d = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'hit' }, { row: 0, col: 1, value: 'hit' }] }])
    const { input } = findInputOf(d)
    const hits = await drain(new WorkbookFind({ ...input, sheets: [{ ...input.sheets[0], isColumnVisible: col => col !== 1 }] }, ci('hit')))
    expect(where(hits)).toEqual([['A', 0, 0]])
  })

  it('orders hits by current display slot after a sort, keeping source identity', async () => {
    const d = workbookOf([{ sheetId: 'A', cells: [0, 1, 2, 3, 5].map(row => ({ row, col: 0, value: `x${row}` })) }])
    // Display [3,0,1,2] over source rows 0..3; row 5 is outside the band.
    const order = createRowOrder([{ r0: 0, r1: 3, sourceAtOffset: Uint32Array.from([3, 0, 1, 2]) }])
    const hits = await drain(new WorkbookFind(findInputOf(d, { orders: { A: order } }).input, ci('x')))
    expect(hits.map(hit => hit.source.row)).toEqual([3, 0, 1, 2, 5])
    expect(hits.map(hit => hit.displaySlot)).toEqual([0, 1, 2, 3, 5])
    expect(matchTarget(hits[0])).toEqual({ sheetId: 'A', cell: { row: 3, col: 0 } })
  })

  it('matches caller-formatted display text when the host supplies it, raw text by default', async () => {
    const d = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 0.5 }] }])
    expect(await drain(new WorkbookFind(findInputOf(d).input, ci('50%')))).toHaveLength(0)
    const formatted = new WorkbookFind(findInputOf(d).input, ci('50%'), { textOf: cell => `${Number(cell.value) * 100}%` })
    expect((await drain(formatted)).map(hit => hit.text)).toEqual(['50%'])
    expect(formatted.stats.textsFormatted).toBe(1)
  })

  it('reports permanently unavailable formula values instead of skipping them silently', async () => {
    const d = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'hit' }, { row: 1, col: 0, value: { unavailable: true } }, { row: 2, col: 0, value: { unavailable: true } }] }])
    const { input } = findInputOf(d)
    const store = input.sheets[0].store
    const marked: SourceStore = Object.create(store, {
      scanPopulated: {
        value: async function* (request: Parameters<SourceStore['scanPopulated']>[0], options: Parameters<SourceStore['scanPopulated']>[1]): AsyncIterable<Page<AuthoritativeCell>> {
          for await (const page of store.scanPopulated(request, options)) {
            yield { ...page, items: page.items.map(cell => cell.formula && !cell.hasCachedValue ? { ...cell, unavailable: true as const } : cell) }
          }
        },
      },
    })
    const session = new WorkbookFind({ ...input, sheets: [{ ...input.sheets[0], store: marked }] }, ci('hit'))
    const page = await session.nextPage()
    expect(page).toMatchObject({ complete: true, totalKnown: 1, unavailableCount: 2 })
  })

  it('pages are bounded with a cursor; the total is known only at completion', async () => {
    const cells = Array.from({ length: 450 }, (_, row) => ({ row, col: 0, value: 'hit' }))
    const session = new WorkbookFind(findInputOf(workbookOf([{ sheetId: 'A', cells }])).input, ci('hit'))
    const first = await session.nextPage()
    expect(first.hits).toHaveLength(FIND_PAGE_HITS)
    expect(first.complete).toBe(false)
    expect(first.totalKnown).toBeUndefined()
    expect(typeof first.cursor).toBe('string')
    const second = await session.nextPage()
    expect(second.hits.map(hit => hit.source.row)[0]).toBe(200)
    const third = await session.nextPage()
    expect(third).toMatchObject({ complete: true, totalKnown: 450 })
    expect(third.hits).toHaveLength(50)
    // Scanning is lazy: retained hits never exceed one page plus one scan page.
    expect(session.stats.maxBufferedHits).toBeLessThanOrEqual(FIND_PAGE_HITS + 4096)
  })

  it('restricts the scan to the requested sheets', async () => {
    const hits = await drain(new WorkbookFind(findInputOf(doc()).input, { ...ci('alpha'), sheetIds: ['B'] }))
    expect(where(hits)).toEqual([['B', 0, 0]])
  })
})

describe('workbook find on F1 without painting or a full bitmap', () => {
  it('finds never-painted edge text visiting each populated cell once, with no coordinate walk', async () => {
    const fixture = createFixture('F1', { nnz: 100_000 })
    const probe = { cellAt: 0, isPopulated: 0, styleOf: 0 }
    const store = fixture.createStore({
      read: {
        cellAt: (row, col) => { probe.cellAt++; return fixture.cellAt(row, col) },
        isPopulated: (row, col) => { probe.isPopulated++; return fixture.isPopulated(row, col) },
        styleOf: (row, col) => { probe.styleOf++; return fixture.styleOf(row, col) },
      },
    })
    const { docId, sheetId, valueGen } = store.stamp
    const owner = { ...initialWorkbookOwner(docId, [{ sheetId }]), sheetVector: [{ sheetId, revision: 0, valueGen }] }
    const session = new WorkbookFind({
      owner,
      sheets: [{ sheetId, store, identity: { ...initialIdentity(docId, sheetId), valueGen }, order: createRowOrder([]), visibility: visibilityFromHidden(createIntervalSet([])) }],
    // A frozen clock: only the 4,096-unit work budget yields, so counts are deterministic.
    }, ci('A1048576'), { now: () => 0 })
    const hits = await drain(session)
    expect(where(hits)).toEqual([[store.stamp.sheetId, MAX_ROWS - 1, 0]])
    // Exactly one source read per populated cell, no coordinate walk, no layout.
    expect(probe.cellAt).toBe(fixture.nnz())
    expect(probe.isPopulated).toBe(0)
    expect(session.stats.cellsVisited).toBe(fixture.nnz())
    // 25 bounded 4,096-cell source pages, one cooperative yield between each.
    expect(session.stats.scanPages).toBe(Math.ceil(fixture.nnz() / 4096))
    expect(session.stats.yields).toBe(Math.ceil(fixture.nnz() / 4096) - 1)
    // One hit record retained, ever; no per-cell text/layout/bitmap allocation.
    expect(session.stats.hitsRecorded).toBe(1)
    expect(session.stats.maxBufferedHits).toBe(1)
    expect(session.stats.textsFormatted).toBe(0)
  })

  it('materializes canonical geometry only for the focused hit', async () => {
    const d = workbookOf([{ sheetId: 'A', cells: [{ row: 5, col: 3, value: 'needle in hay' }] }])
    const [hit] = await drain(new WorkbookFind(findInputOf(d).input, ci('in')))
    const metrics = createSheetMetrics({
      rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20 }),
      cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }),
    })
    const host: CellLayoutHost = {
      layoutBody: vi.fn((cell, w, h): CellBodyLayout => ({ x: 0, y: 0, w, h, formattedText: String(cell.value), text: { kind: 'plain', text: String(cell.value), font: '11px sans-serif', color: '#000', tx: 2, ty: 15, width: 130 } })),
    }
    const cache = new LayoutCache()
    const geometry = hitGeometry(hit, { ref: '', row: 5, col: 3, value: 'needle in hay', styleIndex: 0 }, metrics, initialIdentity('doc', 'A'), host, cache)!
    expect(host.layoutBody).toHaveBeenCalledTimes(1)
    expect(geometry.record.id).toBe(hit.cellId)
    expect(geometry.record.logicalBounds).toEqual({ x: 3 * 64, y: 5 * 20, width: 64, height: 20 })
    // 'in' is offsets [7,9) of 13 units across a 130px run starting at x=192+2.
    expect(geometry.rects).toEqual([{ x: 194 + 70, y: 100, width: 20, height: 20 }])
  })
})

describe('review: lowercase offset mapping (codex 7)', () => {
  const one = async (text: string, query: string) => {
    const d = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: text }] }])
    const [hit] = await drain(new WorkbookFind(findInputOf(d).input, ci(query)))
    return hit && [hit.start, hit.end]
  }
  it('covers the originating character of a lowercase expansion', async () => {
    expect(await one('İ', 'i')).toEqual([0, 1])
    expect(await one('xİy', 'i̇')).toEqual([1, 2])
  })
  it('keeps UTF-16 positions inside a surrogate pair after an expansion', async () => {
    expect(await one('İ😀', '\ud83d')).toEqual([1, 2])
    expect(await one('İ😀', '\ude00')).toEqual([2, 3])
    expect(await one('İ😀', '😀')).toEqual([1, 3])
  })
})

describe('review: typed source failures (codex 5)', () => {
  const failing = (error: { code: 'quota' | 'page-error'; retryable: boolean; message: string }) => {
    const d = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'hit' }, { row: 1, col: 0, value: 'hit' }] }])
    const { input } = findInputOf(d)
    const store = input.sheets[0].store
    const broken: SourceStore = Object.create(store, {
      scanPopulated: {
        value: async function* (request: Parameters<SourceStore['scanPopulated']>[0], options: Parameters<SourceStore['scanPopulated']>[1]): AsyncIterable<Page<AuthoritativeCell>> {
          for await (const page of store.scanPopulated(request, options)) yield { ...page, items: page.items.slice(0, 1), completeness: 'incomplete', error: { ...error } }
        },
      },
    })
    return new WorkbookFind({ ...input, sheets: [{ ...input.sheets[0], store: broken }] }, ci('hit'))
  }
  it.each([
    { code: 'quota' as const, retryable: true, message: 'storage quota exceeded' },
    { code: 'page-error' as const, retryable: false, message: 'corrupt page 7' },
  ])('keeps $code with its message and retryability next to the partial hits', async error => {
    const page = await failing(error).nextPage()
    expect(page.hits).toHaveLength(1)
    expect(page.totalKnown).toBeUndefined()
    expect(page.incomplete).toEqual(['A'])
    expect(page.failures).toEqual([{ sheetId: 'A', ...error, range: { r0: 0, r1: 1, c0: 0, c1: 0 } }])
  })
})
