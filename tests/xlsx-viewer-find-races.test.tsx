/**
 * S5.3 workbook find ownership, races, cancellation and navigation (spec
 * §§10.1, 11.4; AC-LIFE-01, AC-SEARCH-01). Every async publication captures
 * the aggregate workbook owner; a reply from an older owner is never applied.
 */
import { describe, expect, it } from 'vitest'
import type { AuthoritativeCell, Page, SourceStore, WorkbookOwner } from '../src/xlsx/view/contracts'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { createIdentityOrder, createRowOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createScrollWindow } from '../src/xlsx/view/scrollWindow'
import { SheetNavigationSession, emptyGridSelection, selectCell } from '../src/xlsx/view/selection'
import { initialIdentity } from '../src/worker/viewerProtocol'
import {
  FindController, FindError, WorkbookFind, findCacheKey, navigateToHit, workbookFindKey, type FindHit,
} from '../src/xlsx/view/find'
import { drain, findInputOf, where, workbookOf } from './helpers/xlsx-find'

const ci = (text: string) => ({ text, caseSensitive: false })
const owner = (globalViewGen: number, a: number, b: number, extra: Partial<WorkbookOwner> = {}): WorkbookOwner => ({
  docId: 'doc', globalViewGen, workbookValueGen: 0, fontGen: 0,
  sheetVector: [{ sheetId: 'A', revision: a, valueGen: 0 }, { sheetId: 'B', revision: b, valueGen: 0 }], ...extra,
})
const twoSheets = () => workbookOf([
  { sheetId: 'A', cells: [0, 1, 2, 3].map(row => ({ row, col: 0, value: `a${row} hit` })) },
  { sheetId: 'B', cells: [0, 1, 2, 3].map(row => ({ row, col: 0, value: `b${row} hit` })) },
])
const reversed = () => createRowOrder([{ r0: 0, r1: 3, sourceAtOffset: Uint32Array.from([3, 2, 1, 0]) }])
/** A clock that spends the 10 ms budget on every sample: a yield per batch. */
const ticking = () => { let t = 0; return () => (t += 100) }
async function rejection(promise: Promise<unknown>): Promise<FindError> {
  try { await promise } catch (error) { return error as FindError }
  throw new Error('expected a FindError rejection')
}

describe('aggregate workbook owner', () => {
  it('A revision 1 and B revision 1 never share an index key, even at an equal global generation', () => {
    const x = owner(1, 1, 0), y = owner(1, 0, 1)
    expect(workbookFindKey(x)).not.toBe(workbookFindKey(y))
    expect(findCacheKey(x, ci('hit'))).not.toBe(findCacheKey(y, ci('hit')))
    expect(workbookFindKey(owner(1, 0, 0, { workbookValueGen: 1 }))).not.toBe(workbookFindKey(owner(1, 0, 0)))
    expect(workbookFindKey(owner(1, 0, 0, { fontGen: 1 }))).not.toBe(workbookFindKey(owner(1, 0, 0)))
    expect(findCacheKey(x, ci('hit'))).not.toBe(findCacheKey(x, { text: 'hit', caseSensitive: true }))
    expect(findCacheKey(x, ci('hit'))).not.toBe(findCacheKey(x, { ...ci('hit'), sheetIds: ['A'] }))
  })

  it('a cached A=1 index is not reused for B=1: the B view is rescanned in its own display order', async () => {
    let current = owner(1, 1, 0)
    const controller = new FindController({ currentOwner: () => current })
    const d = twoSheets()
    const first = await controller.run(findInputOf(d, { owner: current, orders: { A: reversed() } }).input, ci('hit'))
    expect(where(first!.hits).slice(0, 4).map(([, row]) => row)).toEqual([3, 2, 1, 0])
    expect(controller.stats.cacheHits).toBe(0)
    // Same query, same owner: served from the aggregate cache without a scan.
    const again = await controller.run(findInputOf(d, { owner: current, orders: { A: reversed() } }).input, ci('hit'))
    expect(controller.stats.cacheHits).toBe(1)
    expect(where(again!.hits)).toEqual(where(first!.hits))
    current = owner(1, 0, 1)
    const swapped = await controller.run(findInputOf(d, { owner: current, orders: { B: reversed() } }).input, ci('hit'))
    expect(controller.stats.cacheHits).toBe(1)
    expect(where(swapped!.hits).map(([sheet, row]) => `${sheet}${row}`)).toEqual(['A0', 'A1', 'A2', 'A3', 'B3', 'B2', 'B1', 'B0'])
  })

  it('rejects a sheet identity that disagrees with the captured owner vector', async () => {
    const { input } = findInputOf(twoSheets())
    const bad = { ...input, sheets: [{ ...input.sheets[0], identity: { ...input.sheets[0].identity, viewGen: 7 } }, input.sheets[1]] }
    expect((await rejection(new WorkbookFind(bad, ci('hit')).nextPage())).code).toBe('stale')
    const otherDoc = { ...input, sheets: [{ ...input.sheets[0], identity: initialIdentity('other', 'A') }, input.sheets[1]] }
    expect((await rejection(new WorkbookFind(otherDoc, ci('hit')).nextPage())).code).toBe('stale')
  })
})

describe('stale async publication is never applied', () => {
  it.each([
    ['sort/filter view revision', (o: WorkbookOwner): WorkbookOwner => ({ ...o, globalViewGen: 1, sheetVector: [{ ...o.sheetVector[0], revision: 1 }, o.sheetVector[1]] })],
    ['value generation', (o: WorkbookOwner): WorkbookOwner => ({ ...o, workbookValueGen: 1, sheetVector: [{ ...o.sheetVector[0], valueGen: 1 }, o.sheetVector[1]] })],
    ['font generation', (o: WorkbookOwner): WorkbookOwner => ({ ...o, fontGen: 1 })],
  ])('an owner change (%s) during a scan yield rejects the page as stale', async (_label, bump) => {
    const cells = Array.from({ length: 10_000 }, (_, row) => ({ row, col: 0, value: 'hit' }))
    const { input } = findInputOf(workbookOf([{ sheetId: 'A', cells }, { sheetId: 'B', cells: [] }]))
    let current = input.owner, yields = 0
    const session = new WorkbookFind(input, ci('hit'), { pageSize: 50_000, currentOwner: () => current, yield: async () => { if (++yields === 1) current = bump(current) } })
    expect((await rejection(session.nextPage())).code).toBe('stale')
    expect(yields).toBe(1)
  })

  it('a source page published under an older value stamp is rejected as stale', async () => {
    const { input } = findInputOf(twoSheets())
    const store = input.sheets[0].store
    const moved: SourceStore = Object.create(store, {
      scanPopulated: {
        value: async function* (_request: unknown, options: { stamp: SourceStore['stamp'] }): AsyncIterable<Page<AuthoritativeCell>> {
          yield { items: [], completeness: 'incomplete', stamp: { ...options.stamp, valueGen: 9 }, error: { code: 'stale-source', retryable: true, message: 'value owner changed' } }
        },
      },
    })
    const session = new WorkbookFind({ ...input, sheets: [{ ...input.sheets[0], store: moved }, input.sheets[1]] }, ci('hit'))
    expect((await rejection(session.nextPage())).code).toBe('stale')
  })

  it('the controller drops a superseded run and a run whose owner went stale; the newer result stays published', async () => {
    let current = owner(0, 0, 0)
    const published: string[] = []
    const controller = new FindController({ currentOwner: () => current, onPublish: state => published.push(`${state.query.text}:${state.hits.length}:${state.complete}`) })
    const d = twoSheets()
    let release!: () => void
    const gate = new Promise<void>(resolve => { release = resolve })
    const slow = controller.run(findInputOf(d, { owner: current }).input, ci('a0'), { now: ticking(), yield: () => gate })
    const fast = await controller.run(findInputOf(d, { owner: current }).input, ci('b1'))
    release()
    expect(await slow).toBeNull()
    expect(where(fast!.hits)).toEqual([['B', 1, 0]])
    expect(controller.published?.query.text).toBe('b1')
    expect(published.every(entry => !entry.startsWith('a0:'))).toBe(true)

    // Owner flips while a run is in flight: its pages are never applied.
    let flip!: () => void
    const flipped = new Promise<void>(resolve => { flip = resolve })
    const stale = controller.run(findInputOf(d, { owner: current }).input, ci('hit'), { now: ticking(), yield: async () => { current = owner(1, 1, 0); flip() } })
    await flipped
    expect(await stale).toBeNull()
    expect(controller.published?.query.text).toBe('b1')
    expect(controller.stats.staleRejected).toBeGreaterThan(0)
  })
})

describe('cooperative, cancellable bounded batches', () => {
  it('yields on the S4.2 budget, cancels between batches and frees state on close', async () => {
    const cells = Array.from({ length: 50_000 }, (_, row) => ({ row, col: 0, value: `v${row}` }))
    const { input } = findInputOf(workbookOf([{ sheetId: 'A', cells }, { sheetId: 'B', cells: [] }]))
    const controller = new AbortController()
    let yields = 0
    const session = new WorkbookFind(input, ci('zzz'), { signal: controller.signal, yield: async () => { if (++yields === 3) controller.abort() } })
    expect((await rejection(session.nextPage())).code).toBe('cancelled')
    expect(yields).toBe(3)
    // At most one 4,096-unit budget of cells between yields.
    expect(session.stats.cellsVisited).toBeLessThanOrEqual(3 * 4096 + 4096)
    session.close()
    expect(session.stats.bufferedHits).toBe(0)
    expect((await rejection(session.nextPage())).code).toBe('closed')
  })

  it('a pre-aborted signal never scans', async () => {
    const { input } = findInputOf(twoSheets())
    const signal = AbortSignal.abort()
    const session = new WorkbookFind(input, ci('hit'), { signal })
    expect((await rejection(session.nextPage())).code).toBe('cancelled')
    expect(session.stats.cellsVisited).toBe(0)
  })
})

describe('navigation to a match', () => {
  function navigation() {
    const session = new SheetNavigationSession('doc')
    const viewport = { width: 800, height: 600 }
    const windows = new Map()
    for (const [sheetId, zoom] of [['A', 0.5], ['B', 2]] as const) {
      const order = createIdentityOrder()
      const baseHidden = createIntervalSet([])
      const visibility = visibilityFromHidden(createIntervalSet(sheetId === 'B' ? [{ r0: 10, r1: 10 }] : []))
      const projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }), order, visibility })
      const view = { sheetId, revision: 0, order, baseHidden, visibility, sort: null, filter: sheetId === 'B' ? { range: { r0: 0, r1: 20, c0: 0, c1: 0, headerRows: 0 as const, totalsRows: 0 as const }, columns: new Map([[0, { kind: 'text' as const, op: 'neq' as const, first: 'x', wildcards: false }]]) } : null, selection: emptyGridSelection() }
      session.register(sheetId, { view, projection, body: null, zoom })
      windows.set(sheetId, createScrollWindow(projection, viewport, { zoom }))
    }
    const a = session.get('A')
    session.setSelection('A', selectCell(a.selection, a.projection!, { row: 4, col: 2 }))
    return { session, viewport, windows }
  }
  const hitAt = (sheetId: string, row: number, col: number): FindHit => ({ sheetId, cellId: row * MAX_COLS + col, source: { row, col }, text: 'hit', start: 0, end: 3, displaySlot: row })

  it('activates the hit sheet, scrolls the offscreen cell into view and keeps every sheet zoom and the other selection', () => {
    const nav = navigation()
    const result = navigateToHit(nav, hitAt('B', MAX_ROWS - 1, MAX_COLS - 1))
    expect(result.kind).toBe('visible')
    expect(nav.session.activeSheetId).toBe('B')
    const b = nav.session.get('B')
    expect(b.zoom).toBe(2)
    expect(b.selection.active).toEqual({ row: MAX_ROWS - 1, col: MAX_COLS - 1 })
    const window = nav.windows.get('B')
    expect(window.zoom).toBe(2)
    // The corner's logical box lies inside the scrolled viewport (600x800 CSS at 2x zoom).
    // Sheet B hides display slot 10, so the corner sits one row higher than its source row.
    const box = b.projection!.sourceCellToLogical({ row: MAX_ROWS - 1, col: MAX_COLS - 1 })!
    expect(box.y).toBe((MAX_ROWS - 2) * 20)
    const { x, y } = window.logicalOrigin
    expect(y).toBeLessThanOrEqual(box.y)
    expect(y + 600 / 2).toBeGreaterThanOrEqual(box.y + box.height)
    expect(x).toBeLessThanOrEqual(box.x)
    expect(x + 800 / 2).toBeGreaterThanOrEqual(box.x + box.width)
    expect(nav.session.get('A').zoom).toBe(0.5)
    expect(nav.session.get('A').selection.active).toEqual({ row: 4, col: 2 })
    expect(nav.windows.get('A').zoom).toBe(0.5)
  })

  it('a hit hidden by a later filter opens read-only inspection without moving the view', () => {
    const nav = navigation()
    navigateToHit(nav, hitAt('B', 2, 0))
    const before = nav.windows.get('B')
    const result = navigateToHit(nav, hitAt('B', 10, 0))
    expect(result).toMatchObject({ kind: 'hidden', sheetId: 'B', hiddenReason: 'filter-hidden' })
    expect(nav.windows.get('B')).toBe(before)
    expect(nav.session.get('B').selection.active).toEqual({ row: 2, col: 0 })
  })

  it('an evicted projection is reported as a seam instead of guessing geometry', () => {
    const nav = navigation()
    nav.session.detachProjection('B')
    expect(navigateToHit(nav, hitAt('B', 1, 0))).toEqual({ kind: 'needs-projection', sheetId: 'B' })
    expect(nav.session.get('B').zoom).toBe(2)
  })

  it('end to end: a never-painted far-corner match on an inactive sheet is found and navigated to', async () => {
    const d = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'start' }] }, { sheetId: 'B', cells: [{ row: MAX_ROWS - 1, col: MAX_COLS - 1, value: 'corner needle' }] }])
    const hits = await drain(new WorkbookFind(findInputOf(d).input, ci('needle')))
    const nav = navigation()
    expect(navigateToHit(nav, hits[0]).kind).toBe('visible')
    expect(nav.session.get('B').selection.active).toEqual({ row: MAX_ROWS - 1, col: MAX_COLS - 1 })
    expect(nav.windows.get('B').zoom).toBe(2)
  })
})

describe('review: sheet coverage of the aggregate owner (codex 6)', () => {
  it('a requested sheet missing from the input is explicitly incomplete, never a final total', async () => {
    const { input } = findInputOf(twoSheets())
    const onlyA = { ...input, sheets: [input.sheets[0]] }
    const page = await new WorkbookFind(onlyA, ci('hit')).nextPage()
    expect(where(page.hits).map(([sheet]) => sheet)).toEqual(['A', 'A', 'A', 'A'])
    expect(page.complete).toBe(true)
    expect(page.totalKnown).toBeUndefined()
    expect(page.incomplete).toEqual(['B'])
    expect(page.failures).toMatchObject([{ sheetId: 'B', code: 'incomplete' }])
    // Not requested, not missing: a sheet-restricted query over A alone is final.
    expect((await new WorkbookFind(onlyA, { ...ci('hit'), sheetIds: ['A'] }).nextPage()).totalKnown).toBe(4)
  })

  it('scans in owner (workbook) order whatever order the sheets are supplied in', async () => {
    const { input } = findInputOf(twoSheets())
    const hits = await drain(new WorkbookFind({ ...input, sheets: [input.sheets[1], input.sheets[0]] }, ci('0 hit')))
    expect(where(hits)).toEqual([['A', 0, 0], ['B', 0, 0]])
  })

  it('a requested sheet unknown to the owner rejects as stale', async () => {
    const { input } = findInputOf(twoSheets())
    expect((await rejection(new WorkbookFind(input, { ...ci('hit'), sheetIds: ['Z'] }).nextPage())).code).toBe('stale')
  })

  it('a partial-coverage result is never cached for the full workbook request', async () => {
    const d = twoSheets()
    const { input } = findInputOf(d)
    const controller = new FindController({ currentOwner: () => input.owner })
    await controller.run({ ...input, sheets: [input.sheets[0]] }, ci('hit'))
    const full = await controller.run(findInputOf(d).input, ci('hit'))
    expect(controller.stats.cacheHits).toBe(0)
    expect(full!.hits.map(hit => hit.sheetId)).toContain('B')
    expect(full!.totalKnown).toBe(8)
  })
})

describe('review: settling a pending yield (codex 4)', () => {
  const settles = <T,>(promise: Promise<T>) => Promise.race([promise.then(() => 'resolved', (error: FindError) => error.code), new Promise(resolve => setTimeout(() => resolve('hung'), 200))])
  const pending = () => {
    const cells = Array.from({ length: 10_000 }, (_, row) => ({ row, col: 0, value: `v${row}` }))
    return findInputOf(workbookOf([{ sheetId: 'A', cells }, { sheetId: 'B', cells: [] }])).input
  }
  it('close() during a yield whose port was disposed rejects the page as closed instead of hanging', async () => {
    let entered!: () => void
    const inYield = new Promise<void>(resolve => { entered = resolve })
    // A disposed MessageChannel port never delivers: the yield never settles.
    const session = new WorkbookFind(pending(), ci('zzz'), { yield: () => { entered(); return new Promise<void>(() => {}) } })
    const page = settles(session.nextPage())
    await inYield
    session.close()
    expect(await page).toBe('closed')
  })
  it('an abort during a pending yield rejects as cancelled', async () => {
    let entered!: () => void
    const inYield = new Promise<void>(resolve => { entered = resolve })
    const abort = new AbortController()
    const session = new WorkbookFind(pending(), ci('zzz'), { signal: abort.signal, yield: () => { entered(); return new Promise<void>(() => {}) } })
    const page = settles(session.nextPage())
    await inYield
    abort.abort()
    expect(await page).toBe('cancelled')
  })
})
