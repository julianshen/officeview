/** S7.1 range detection, header policy and structural admission (spec §6.1). */
import { describe, expect, it } from 'vitest'
import { detectOperationRange, admitSortRange, type RangeDetection, type RangeDetectionRequest } from '../src/xlsx/view/rangeDetection'
import { ColumnChunkIndex, RowChunkIndex, StructuralIndex, currentRegion, fromStoreScan } from '../src/xlsx/view/structure'
import { guardSort } from '../src/xlsx/view/operationGuard'
import type {
  AuthoritativeCell,
  FilterRange,
  GridSelection,
  Page,
  SourceStamp,
  SourceStore,
  StructuralResult,
} from '../src/xlsx/view/contracts'
import { MAX_COLS, MAX_ROWS, type SourceRect } from '../src/xlsx/view/coordinates'
import { createFixture, readOptionsFor, type FixtureCell, type FixtureHandle } from './helpers/xlsx-viewer-fixtures'

const fullDomain: SourceRect = { r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }
const fixtureCell = (value: FixtureCell['value'], error = false): FixtureCell => Object.freeze({ value, valueIsError: error, hasCachedValue: true, styleIndex: 0 })
const noCache = (): FixtureCell => ({ value: null, formula: 'WEBSERVICE()', hasCachedValue: false, valueIsError: false, styleIndex: 0 })
function gridStore(
  rows: readonly (readonly (FixtureCell | undefined)[])[],
  rowCount?: number,
  extra?: { pending?: SourceRect[]; errors?: Array<{ range: SourceRect; code: 'page-error' | 'storage-error'; retryable: boolean }>; valuePolicy?: 'engine' | 'cached' },
): SourceStore {
  return createFixture('F2', { rowCount: rowCount ?? Math.max(rows.length, 3) }).createStore({
    valuePolicy: extra?.valuePolicy,
    pending: extra?.pending,
    errors: extra?.errors,
    read: {
      isPopulated: (row, col) => rows[row]?.[col] !== undefined,
      cellAt: (row, col) => rows[row]?.[col] ?? null,
    },
  })
}
function emptyIndex(store: SourceStore, stamp: SourceStamp = store.stamp): StructuralIndex {
  return new StructuralIndex({ stamp, completeness: 'complete', merges: [], tables: [], spills: [], drawings: [], missing: [] })
}
async function f5Setup(): Promise<{ handle: FixtureHandle; store: SourceStore; index: StructuralIndex }> {
  const handle = createFixture('F5')
  const store = handle.createStore()
  const result = await store.queryStructural(fullDomain, readOptionsFor(store))
  return { handle, store, index: new StructuralIndex(result) }
}
const sel = (row: number, col: number, ranges: GridSelection['ranges'] = []): GridSelection => ({
  active: { row, col }, anchor: { row, col }, ranges,
})
function request(store: SourceStore, structure: StructuralIndex, selection: GridSelection, extra?: Partial<RangeDetectionRequest>): RangeDetectionRequest {
  return { store, structure, selection, options: readOptionsFor(store), ...extra }
}
type Source = { kind: 'table'; id: string } | { kind: 'region' } | { kind: 'used' } | { kind: 'explicit' }
/** Expected complete envelope: range plus stamp and provenance (spec §§6.1, 6.4). */
function done(store: SourceStore, range: FilterRange, source: Source, header: 'metadata' | 'heuristic' | 'override', overrideIgnored = false): RangeDetection {
  return { kind: 'complete', range, stamp: store.stamp, source, header, overrideIgnored }
}
const headerGrid = (): (readonly (FixtureCell | undefined)[])[] => [
  [fixtureCell('Amount'), fixtureCell('Qty')],
  [fixtureCell(5), fixtureCell(10)],
  [fixtureCell(7), fixtureCell(3)],
  [fixtureCell(2), fixtureCell(8)],
]
const textGrid = (): (readonly (FixtureCell | undefined)[])[] => [
  [fixtureCell('a'), fixtureCell('b')],
  [fixtureCell('c'), fixtureCell('d')],
  [fixtureCell('e'), fixtureCell('f')],
]

/** Work counters for custom stores: every scan page request and every item served. */
interface ScanProbe { calls: number; cells: number }
type ProbedStore = SourceStore & { probe: ScanProbe }
const stampOf = (docId: string): SourceStamp => Object.freeze({ docId, sheetId: 's', valueGen: 0 })
function authoritative(row: number, col: number, value: string | number): AuthoritativeCell {
  return { id: row * MAX_COLS + col, row, col, value, hasCachedValue: true, valueIsError: false, styleIndex: 0 } as AuthoritativeCell
}
function unsupported(): never {
  throw new Error('not used by range detection')
}
function spin(ms: number): void {
  const end = performance.now() + ms
  while (performance.now() < end) { /* synchronous store work */ }
}
/** Base for hand-rolled stores whose scans are NOT clipped to the used hint. */
function customStore(
  docId: string, used: SourceRect | undefined,
  cellsIn: (rect: SourceRect, offset: number, limit: number) => { items: AuthoritativeCell[]; more: boolean },
  spinMs = 0,
): ProbedStore {
  const stamp = stampOf(docId)
  const probe: ScanProbe = { calls: 0, cells: 0 }
  return {
    stamp, valuePolicy: 'engine', probe,
    extent: used === undefined ? { rows: 1048576, cols: 16384 } : { rows: 1048576, cols: 16384, used },
    readCell: async () => unsupported(),
    readRows: async (range, _cursor, options) => {
      options.signal.throwIfAborted()
      const byRow = new Map<number, AuthoritativeCell[]>()
      for (let offset = 0; ; offset += 4096) {
        const page = cellsIn(range, offset, 4096)
        for (const cell of page.items) {
          const list = byRow.get(cell.row) ?? []
          list.push(cell)
          byRow.set(cell.row, list)
        }
        if (!page.more) break
      }
      const items = [...byRow.entries()].sort((a, b) => a[0] - b[0]).map(([row, cells]) => ({ row, cells }))
      return { items, completeness: 'complete', stamp }
    },
    queryColumn: async () => unsupported(),
    scanPopulated: async function* (query, options): AsyncIterable<Page<AuthoritativeCell>> {
      options.signal.throwIfAborted()
      probe.calls++
      if (spinMs > 0) spin(spinMs)
      const offset = query.cursor === undefined ? 0 : Number(query.cursor)
      const page = cellsIn(query.range, offset, query.limit)
      probe.cells += page.items.length
      yield { items: page.items, completeness: 'complete', stamp, nextCursor: page.more ? String(offset + page.items.length) : undefined }
    },
    queryStructural: async () => unsupported(),
    queryOverview: async () => unsupported(),
    ensureComplete: async () => unsupported(),
    release: async () => {},
  }
}
/** Column A populated on rows [0, rows); served in O(page) with offset cursors. */
function denseColumnStore(rows: number, spinMs = 0): ProbedStore {
  return customStore('dense-column', { r0: 0, r1: rows - 1, c0: 0, c1: 0 }, (rect, offset, limit) => {
    if (rect.c0 > 0) return { items: [], more: false }
    const r1 = Math.min(rect.r1, rows - 1)
    const first = rect.r0 + offset
    const items: AuthoritativeCell[] = []
    for (let row = first; row <= r1 && items.length < limit; row++) items.push(authoritative(row, 0, row))
    return { items, more: first + items.length <= r1 }
  }, spinMs)
}
/** Explicit sparse cell list. Scans serve row-major order by default; other
 * orders prove detection never depends on the (uncontracted) page order. */
function sparseStore(
  cells: ReadonlyArray<readonly [number, number, string | number]>, used?: SourceRect,
  order: 'row' | 'reverse' | 'column' = 'row', pageCap = Number.POSITIVE_INFINITY,
): ProbedStore {
  const sorted = [...cells].sort(order === 'row' ? (a, b) => a[0] - b[0] || a[1] - b[1]
    : order === 'column' ? (a, b) => a[1] - b[1] || a[0] - b[0] : (a, b) => b[0] - a[0] || b[1] - a[1])
  return customStore('sparse', used, (rect, offset, limit) => {
    const inside = sorted.filter(([row, col]) => row >= rect.r0 && row <= rect.r1 && col >= rect.c0 && col <= rect.c1)
    const items = inside.slice(offset, offset + Math.min(limit, pageCap)).map(([row, col, value]) => authoritative(row, col, value))
    return { items, more: offset + items.length < inside.length }
  })
}

describe('table precedence and metadata protection', () => {
  it('resolves the containing table with metadata header and totals rows', async () => {
    const { store, index } = await f5Setup()
    const result = await detectOperationRange(request(store, index, sel(10, 2)))
    expect(result).toEqual(done(store, { r0: 4, r1: 20, c0: 0, c1: 3, headerRows: 1, totalsRows: 1 }, { kind: 'table', id: 'table-0' }, 'metadata'))
    if (result.kind === 'complete') expect(Object.isFrozen(result.range)).toBe(true)
  })
  it('honors an explicit table headerRows of 0', async () => {
    const { store, index } = await f5Setup()
    const result = await detectOperationRange(request(store, index, sel(25, 0)))
    expect(result).toEqual(done(store, { r0: 24, r1: 30, c0: 0, c1: 2, headerRows: 0, totalsRows: 0 }, { kind: 'table', id: 'table-1' }, 'metadata'))
  })
  it('reports a dialog override on a table as ignored instead of dropping it silently', async () => {
    const { store, index } = await f5Setup()
    const result = await detectOperationRange(request(store, index, sel(10, 2), { override: { headerRows: 0 } }))
    expect(result).toEqual(done(store, { r0: 4, r1: 20, c0: 0, c1: 3, headerRows: 1, totalsRows: 1 }, { kind: 'table', id: 'table-0' }, 'metadata', true))
  })
  it('checks child-only slices over the body band, never over protected header rows', async () => {
    const rows = [
      [fixtureCell('T'), fixtureCell('H1'), fixtureCell('H2'), fixtureCell('H3')],
      [undefined, fixtureCell(1), fixtureCell(2), fixtureCell(3)],
      [undefined, fixtureCell(4), fixtureCell(5), fixtureCell(6)],
    ]
    const store = gridStore(rows)
    const range = { r0: 0, r1: 2, c0: 1, c1: 3 }
    const structure = new StructuralIndex({
      stamp: store.stamp, completeness: 'complete', spills: [], drawings: [], missing: [],
      merges: [{ id: 'm', range: { r0: 0, r1: 0, c0: 0, c1: 2 }, anchor: { row: 0, col: 0 } }],
      tables: [{ id: 't', range, headerRows: 1, totalsRows: 0 }],
    })
    const filterRange = { ...range, headerRows: 1 as const, totalsRows: 0 as const }
    expect(await guardSort(filterRange, structure, fromStoreScan(store, readOptionsFor(store)))).toEqual({ ok: true })
    expect(await detectOperationRange(request(store, structure, sel(1, 1)))).toEqual(done(store, filterRange, { kind: 'table', id: 't' }, 'metadata'))
    expect(await admitSortRange(request(store, structure, sel(1, 1)))).toEqual(done(store, filterRange, { kind: 'table', id: 't' }, 'metadata'))
  })
})

describe('non-table header policy', () => {
  it('treats a text header over numbers as a header row', async () => {
    const store = gridStore(headerGrid())
    const expected = done(store, { r0: 0, r1: 3, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 }, { kind: 'region' }, 'heuristic')
    expect(await detectOperationRange(request(store, emptyIndex(store), sel(1, 0)))).toEqual(expected)
    expect(await admitSortRange(request(store, emptyIndex(store), sel(1, 0)))).toEqual(expected)
  })
  it('treats an all-text first row as data, never an automatic header', async () => {
    const store = gridStore(textGrid())
    const result = await detectOperationRange(request(store, emptyIndex(store), sel(0, 0)))
    expect(result).toEqual(done(store, { r0: 0, r1: 2, c0: 0, c1: 1, headerRows: 0, totalsRows: 0 }, { kind: 'region' }, 'heuristic'))
    expect(await admitSortRange(request(store, emptyIndex(store), sel(0, 0)))).toEqual(result)
  })
  it('applies an explicit header override to the non-table heuristic', async () => {
    const store = gridStore(textGrid())
    const headed = await detectOperationRange(request(store, emptyIndex(store), sel(0, 0), { override: { headerRows: 1 } }))
    expect(headed).toEqual(done(store, { r0: 0, r1: 2, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 }, { kind: 'region' }, 'override'))
  })
  it('accepts only a header override: a non-table totals override is refused (spec §6.4)', async () => {
    const store = gridStore(textGrid())
    const tailed = { headerRows: 1, totalsRows: 1 } as unknown as { headerRows: 0 | 1 }
    expect(await detectOperationRange(request(store, emptyIndex(store), sel(0, 0), { override: tailed }))).toEqual({
      kind: 'refused', refusal: { reason: 'invalid-rule', message: expect.any(String) },
    })
  })
  it('refuses a malformed header override instead of guessing', async () => {
    const store = gridStore(textGrid())
    const bad = { headerRows: 2 } as unknown as { headerRows: 0 | 1 }
    expect(await detectOperationRange(request(store, emptyIndex(store), sel(0, 0), { override: bad }))).toEqual({
      kind: 'refused', refusal: { reason: 'invalid-rule', message: expect.any(String) },
    })
  })
  it('types cached formula values normally and never as blank', async () => {
    const rows = [ [fixtureCell('A'), fixtureCell('B')], [{ ...fixtureCell(5), formula: 'SUM(A2:A3)' }, fixtureCell(6)] ]
    const store = gridStore(rows)
    expect(await detectOperationRange(request(store, emptyIndex(store), sel(0, 0)))).toEqual(
      done(store, { r0: 0, r1: 1, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 }, { kind: 'region' }, 'heuristic'),
    )
  })
  it('refuses permanently unavailable formula values in the header probe instead of waiting forever', async () => {
    const rows = [
      [fixtureCell('A'), noCache(), noCache()],
      [fixtureCell(1), fixtureCell(2), noCache()],
    ]
    const store = gridStore(rows, undefined, { valuePolicy: 'cached' })
    const result = await detectOperationRange(request(store, emptyIndex(store), sel(0, 0)))
    expect(result).toEqual({
      kind: 'refused', refusal: { reason: 'formula-value-unavailable', count: 3, firstAddress: { row: 0, col: 1 } },
    })
    // Retrying cannot change a permanent state: the refusal is stable.
    expect(await detectOperationRange(request(store, emptyIndex(store), sel(0, 0)))).toEqual(result)
  })
  it('waits on a pending page in the header probe', async () => {
    const store = gridStore(headerGrid(), undefined, { pending: [{ r0: 1, r1: 1, c0: 1, c1: 1 }] })
    const result = await detectOperationRange(request(store, emptyIndex(store), sel(0, 0)))
    expect(result).toEqual({ kind: 'incomplete', missing: [{ r0: 0, r1: 1, c0: 0, c1: 1 }] })
  })
})

describe('ambiguous and invalid inputs', () => {
  it('refuses a selection with no active cell instead of guessing a seed', async () => {
    const { store, index } = await f5Setup()
    const selection: GridSelection = { active: null, anchor: null, ranges: [] }
    expect(await detectOperationRange(request(store, index, selection))).toEqual({
      kind: 'refused', refusal: { reason: 'ambiguous-selection', message: expect.any(String) },
    })
  })
  it('never coerces a disjoint selection into its bounding rectangle', async () => {
    const { store, index } = await f5Setup()
    const selection = sel(0, 3, [
      { rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 1 }] }, cols: [{ c0: 0, c1: 1 }] },
      { rows: { kind: 'intervals', intervals: [{ r0: 100, r1: 100 }] }, cols: [{ c0: 0, c1: 0 }] },
    ])
    const result = await detectOperationRange(request(store, index, selection))
    expect(result.kind).toBe('complete')
    if (result.kind === 'complete') {
      expect(result.range).toEqual({ r0: 0, r1: 9, c0: 0, c1: 9, headerRows: 0, totalsRows: 0 })
      expect(result.source).toEqual({ kind: 'region' })
      expect(result.range.r1).toBeLessThan(100)
    }
  })
  it('seeds region detection from the active cell for full-column selections', async () => {
    const { store, index } = await f5Setup()
    const selection = sel(0, 3, [
      { rows: { kind: 'intervals', intervals: [{ r0: 0, r1: MAX_ROWS - 1 }] }, cols: [{ c0: 0, c1: 0 }] },
    ])
    const result = await detectOperationRange(request(store, index, selection))
    expect(result.kind).toBe('complete')
    if (result.kind === 'complete') {
      expect(result.range).toMatchObject({ r0: 0, r1: 9 })
      expect(result.source).toEqual({ kind: 'region' })
    }
  })
  it('uses an explicit contiguous selection as the operation range', async () => {
    const store = gridStore(headerGrid())
    const selection = sel(1, 0, [{ rows: { kind: 'intervals', intervals: [{ r0: 1, r1: 2 }] }, cols: [{ c0: 0, c1: 1 }] }])
    expect(await detectOperationRange(request(store, emptyIndex(store), selection))).toEqual(
      done(store, { r0: 1, r1: 2, c0: 0, c1: 1, headerRows: 0, totalsRows: 0 }, { kind: 'explicit' }, 'heuristic'),
    )
  })
  it('refuses out-of-domain active cells and used ranges', async () => {
    const { store, index } = await f5Setup()
    const badCell = sel(-1, 0)
    expect(await detectOperationRange(request(store, index, badCell))).toEqual({
      kind: 'refused', refusal: { reason: 'invalid-range', message: expect.any(String) },
    })
    const badUsed = { r0: 5, r1: 3, c0: 0, c1: 0 }
    expect(await detectOperationRange(request(store, index, sel(50, 50), { usedRange: badUsed }))).toEqual({
      kind: 'refused', refusal: { reason: 'invalid-range', message: expect.any(String) },
    })
  })
})

describe('source identity', () => {
  it('refuses a structural index stamped for another document or value generation', async () => {
    const store = gridStore(headerGrid())
    const stale = emptyIndex(store, { ...store.stamp, docId: 'other-doc', valueGen: 99 })
    const tables = new StructuralIndex({
      stamp: { ...store.stamp, valueGen: 99 }, completeness: 'complete', merges: [], spills: [], drawings: [], missing: [],
      tables: [{ id: 't', range: { r0: 0, r1: 3, c0: 0, c1: 1 }, headerRows: 0, totalsRows: 1 }],
    })
    const refusal = { kind: 'refused', refusal: { reason: 'source-error', code: 'stale-source', retryable: true, message: expect.any(String) } }
    expect(await detectOperationRange(request(store, stale, sel(1, 0)))).toEqual(refusal)
    expect(await admitSortRange(request(store, tables, sel(1, 0)))).toEqual(refusal)
  })
  it('refuses read options stamped for another source', async () => {
    const store = gridStore(headerGrid())
    const options = { signal: new AbortController().signal, stamp: { ...store.stamp, valueGen: 1 } }
    expect(await detectOperationRange({ store, structure: emptyIndex(store), selection: sel(1, 0), options })).toEqual({
      kind: 'refused', refusal: { reason: 'source-error', code: 'stale-source', retryable: true, message: expect.any(String) },
    })
  })
})

describe('used-range fallback', () => {
  it('refuses a blank seed when the used range holds more than one current region', async () => {
    const rows = [
      [fixtureCell('K'), fixtureCell('V'), undefined, fixtureCell('K2'), fixtureCell('V2')],
      [fixtureCell(3), fixtureCell(30), undefined, fixtureCell(9), fixtureCell(90)],
      [fixtureCell(1), fixtureCell(10), undefined, fixtureCell(8), fixtureCell(80)],
    ]
    const store = gridStore(rows)
    const ambiguous = { kind: 'refused', refusal: { reason: 'ambiguous-selection', message: expect.any(String) } }
    expect(await detectOperationRange(request(store, emptyIndex(store), sel(1, 2)))).toEqual(ambiguous)
    expect(await admitSortRange(request(store, emptyIndex(store), sel(1, 2), { usedRange: { r0: 0, r1: 2, c0: 0, c1: 4 } }))).toEqual(ambiguous)
  })
  it('refuses F5 blank seeds inside and outside the used rows: its used range is not one region', async () => {
    const { store, index } = await f5Setup()
    const ambiguous = { kind: 'refused', refusal: { reason: 'ambiguous-selection', message: expect.any(String) } }
    expect(await detectOperationRange(request(store, index, sel(50, 50)))).toEqual(ambiguous)
    expect(await detectOperationRange(request(store, index, sel(500000, 0)))).toEqual(ambiguous)
    expect(await admitSortRange(request(store, index, sel(50, 50)))).toEqual(ambiguous)
  })
  it('resolves a blank seed to the sole current region of the used range and admits it', async () => {
    const store = gridStore(headerGrid())
    const expected = done(store, { r0: 0, r1: 3, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 }, { kind: 'used' }, 'heuristic')
    expect(await detectOperationRange(request(store, emptyIndex(store), sel(2, 5)))).toEqual(expected)
    expect(await admitSortRange(request(store, emptyIndex(store), sel(2, 5)))).toEqual(expected)
  })
  it('waits on an unknown used extent only for blank seeds', async () => {
    const store = gridStore(headerGrid())
    const noused = { ...store, extent: { rows: 1048576 as const, cols: 16384 as const } }
    expect(await detectOperationRange(request(noused, emptyIndex(noused), sel(0, 0)))).toEqual(
      done(noused, { r0: 0, r1: 3, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 }, { kind: 'region' }, 'heuristic'),
    )
    expect(await detectOperationRange(request(noused, emptyIndex(noused), sel(2, 5)))).toEqual({
      kind: 'incomplete', missing: [{ r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }],
    })
  })
  it('treats a narrow used extent as a hint and grows past it as the data requires', async () => {
    const store = gridStore(headerGrid())
    const expected = done(store, { r0: 0, r1: 3, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 }, { kind: 'region' }, 'heuristic')
    const hint = { r0: 0, r1: 1, c0: 0, c1: 1 }
    expect(await detectOperationRange(request(store, emptyIndex(store), sel(0, 0), { usedRange: hint }))).toEqual(expected)
    // An active row outside the hinted used rows still probes the data.
    expect(await detectOperationRange(request(store, emptyIndex(store), sel(3, 0), { usedRange: hint }))).toEqual(expected)
  })
})

describe('structural admission through the shared guard', () => {
  function tableIndex(store: SourceStore, tables: StructuralResult['tables']): StructuralIndex {
    return new StructuralIndex({ stamp: store.stamp, completeness: 'complete', merges: [], tables, spills: [], drawings: [], missing: [] })
  }
  it('refuses separated outside-column data across a blank gap atomically', async () => {
    const rows = [
      [fixtureCell('H0'), fixtureCell('H1')],
      [fixtureCell(1), fixtureCell(2), undefined, fixtureCell(9)],
      [fixtureCell(3), fixtureCell(4), undefined, fixtureCell(9)],
      [fixtureCell(5), fixtureCell(6), undefined, fixtureCell(9)],
      [fixtureCell(7), fixtureCell(8), undefined, fixtureCell(9)],
    ]
    const store = gridStore(rows)
    const structure = tableIndex(store, [{ id: 't', range: { r0: 0, r1: 4, c0: 0, c1: 1 }, headerRows: 1, totalsRows: 0 }])
    const selection = sel(2, 0)
    const before = JSON.stringify(selection)
    const outcome = await admitSortRange(request(store, structure, selection))
    expect(outcome).toEqual({ kind: 'refused', refusal: { reason: 'outside-column-data', range: { r0: 1, r1: 1, c0: 3, c1: 3 } } })
    expect('range' in outcome).toBe(false)
    expect(JSON.stringify(selection)).toBe(before)
  })
  it('waits on a pending outside-range page instead of a safe result', async () => {
    const rows = [
      [fixtureCell('H0'), fixtureCell('H1')],
      [fixtureCell(1), fixtureCell(2), undefined, fixtureCell(9)],
      [fixtureCell(3), fixtureCell(4), undefined, fixtureCell(9)],
      [fixtureCell(5), fixtureCell(6), undefined, fixtureCell(9)],
      [fixtureCell(7), fixtureCell(8), undefined, fixtureCell(9)],
    ]
    const store = gridStore(rows, undefined, { pending: [{ r0: 1, r1: 4, c0: 3, c1: 3 }] })
    const structure = tableIndex(store, [{ id: 't', range: { r0: 0, r1: 4, c0: 0, c1: 1 }, headerRows: 1, totalsRows: 0 }])
    const outcome = await admitSortRange(request(store, structure, sel(2, 0)))
    expect(outcome.kind).toBe('incomplete')
    if (outcome.kind === 'incomplete') {
      expect(outcome.missing).toEqual([{ r0: 1, r1: 4, c0: 2, c1: MAX_COLS - 1 }])
    }
  })
  it('refuses intersecting merges through the guard', async () => {
    const { store, index } = await f5Setup()
    expect(await admitSortRange(request(store, index, sel(0, 0)))).toEqual({
      kind: 'refused', refusal: { reason: 'merge', range: { r0: 0, r1: 1, c0: 0, c1: 1 } },
    })
  })
  it('refuses child-only spills at derivation time', async () => {
    const { store, index } = await f5Setup()
    expect(await detectOperationRange(request(store, index, sel(45, 3)))).toEqual({
      kind: 'refused', refusal: { reason: 'dynamic-spill', range: { r0: 40, r1: 50, c0: 3, c1: 3 } },
    })
  })
  it('refuses geometrically affected drawings without projections', async () => {
    const rows: (readonly (FixtureCell | undefined)[])[] = []
    rows[70] = [fixtureCell(1), fixtureCell(2)]
    rows[71] = [fixtureCell(3), fixtureCell(4)]
    const store = gridStore(rows, 72)
    const structure = new StructuralIndex({
      stamp: store.stamp, completeness: 'complete', merges: [], tables: [], spills: [],
      drawings: [{
        id: 'd70', kind: 'oneCell',
        from: { row: 70, col: 2, offsetXEmu: 0, offsetYEmu: 0 },
        extent: { x: 0, y: 0, width: 1, height: 1 },
        bounds: { r0: 70, r1: 74, c0: 2, c1: 4 },
      }],
      missing: [],
    })
    expect(await admitSortRange(request(store, structure, sel(70, 0)))).toEqual({
      kind: 'refused', refusal: { reason: 'drawing-impact', drawingId: 'd70' },
    })
  })
  it('refuses the F5 metadata table through the guard on far outside-column data', async () => {
    const { store, index } = await f5Setup()
    expect(await admitSortRange(request(store, index, sel(10, 2)))).toEqual({
      kind: 'refused', refusal: { reason: 'outside-column-data', range: { r0: 5, r1: 5, c0: 4, c1: 4 } },
    })
  })
  it('finds outside data at the far column 16383 beyond the used hint', async () => {
    const cells: Array<readonly [number, number, string | number]> = [
      [0, 0, 'K'], [0, 1, 'V'], [1, 0, 3], [1, 1, 30], [2, 0, 1], [2, 1, 10], [2, MAX_COLS - 1, 'far'],
    ]
    const store = sparseStore(cells, { r0: 0, r1: 2, c0: 0, c1: 1 })
    const detected = await detectOperationRange(request(store, emptyIndex(store), sel(1, 0)))
    expect(detected).toEqual(done(store, { r0: 0, r1: 2, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 }, { kind: 'region' }, 'heuristic'))
    expect(await admitSortRange(request(store, emptyIndex(store), sel(1, 0)))).toEqual({
      kind: 'refused', refusal: { reason: 'outside-column-data', range: { r0: 2, r1: 2, c0: MAX_COLS - 1, c1: MAX_COLS - 1 } },
    })
  })
  it('grows a region to the last column of the domain', async () => {
    const edge = MAX_COLS - 1
    const store = sparseStore([[0, edge - 1, 'a'], [0, edge, 'b'], [1, edge - 1, 1], [1, edge, 2]], { r0: 0, r1: 1, c0: edge - 1, c1: edge })
    expect(await detectOperationRange(request(store, emptyIndex(store), sel(1, edge)))).toEqual(
      done(store, { r0: 0, r1: 1, c0: edge - 1, c1: edge, headerRows: 1, totalsRows: 0 }, { kind: 'region' }, 'heuristic'),
    )
  })
})

describe('source failures and pending coverage', () => {
  const twoBlock = (): (readonly (FixtureCell | undefined)[])[] => [
    [fixtureCell('H'), fixtureCell('Q')], [fixtureCell(1), fixtureCell(2)], [fixtureCell(3), fixtureCell(4)],
  ]
  it('maps a retryable storage error in the region probe to a typed source-error refusal', async () => {
    const store = gridStore(twoBlock(), 3, { errors: [{ range: { r0: 0, r1: 2, c0: 2, c1: 2 }, code: 'storage-error', retryable: true }] })
    const refusal = { kind: 'refused', refusal: { reason: 'source-error', code: 'storage-error', retryable: true, message: expect.any(String) } }
    expect(await detectOperationRange(request(store, emptyIndex(store), sel(1, 0)))).toEqual(refusal)
    expect(await admitSortRange(request(store, emptyIndex(store), sel(1, 0)))).toEqual(refusal)
  })
  it('maps a storage error reached only by the guard on the table path to a typed refusal', async () => {
    const store = gridStore(twoBlock(), 3, { errors: [{ range: { r0: 0, r1: 2, c0: 5, c1: 5 }, code: 'storage-error', retryable: true }] })
    const structure = new StructuralIndex({
      stamp: store.stamp, completeness: 'complete', merges: [], spills: [], drawings: [], missing: [],
      tables: [{ id: 't', range: { r0: 0, r1: 2, c0: 0, c1: 1 }, headerRows: 1, totalsRows: 0 }],
    })
    expect(await admitSortRange(request(store, structure, sel(1, 0)))).toEqual({
      kind: 'refused', refusal: { reason: 'source-error', code: 'storage-error', retryable: true, message: expect.any(String) },
    })
  })
  it('waits on a pending page adjacent to a non-table current region', async () => {
    const store = gridStore(twoBlock(), 3, { pending: [{ r0: 0, r1: 2, c0: 2, c1: 2 }] })
    const result = await detectOperationRange(request(store, emptyIndex(store), sel(1, 0)))
    expect(result).toEqual({ kind: 'incomplete', missing: [{ r0: 0, r1: 2, c0: 2, c1: 2 }] })
    if (result.kind === 'incomplete') {
      expect(Object.isFrozen(result)).toBe(true)
      expect(Object.isFrozen(result.missing)).toBe(true)
      expect(Object.isFrozen(result.missing[0])).toBe(true)
    }
  })
  it('does not wait on a pending page that cannot touch the current region', async () => {
    const store = gridStore(twoBlock(), 3, { pending: [{ r0: 1, r1: 1, c0: 6, c1: 6 }] })
    expect(await detectOperationRange(request(store, emptyIndex(store), sel(1, 0)))).toEqual(
      done(store, { r0: 0, r1: 2, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 }, { kind: 'region' }, 'heuristic'),
    )
  })
})

describe('cancellation, bounds and immutability', () => {
  it('cancels instead of deciding when the signal is already aborted', async () => {
    const { store, index } = await f5Setup()
    const controller = new AbortController()
    controller.abort()
    const outcome = await detectOperationRange({
      store, structure: index, selection: sel(10, 2), options: { signal: controller.signal, stamp: { ...store.stamp } },
    })
    expect(outcome).toEqual({ kind: 'cancelled' })
  })
  it('lands a mid-run abort through cooperative yields', async () => {
    const store = denseColumnStore(65536, 1)
    const controller = new AbortController()
    const options = { signal: controller.signal, stamp: { ...store.stamp } }
    const started = performance.now()
    const pending = detectOperationRange({ store, structure: emptyIndex(store), selection: sel(32768, 0), options, pageLimit: 64 })
    setTimeout(() => controller.abort(), 5)
    expect(await pending).toEqual({ kind: 'cancelled' })
    expect(performance.now() - started).toBeLessThan(1000)
  })
  it('bounds current-region work on a full-capacity dense column', async () => {
    const store = denseColumnStore(MAX_ROWS)
    const result = await detectOperationRange(request(store, emptyIndex(store), sel(MAX_ROWS / 2, 0)))
    expect(result).toEqual(done(store, { r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, { kind: 'region' }, 'heuristic'))
    // Each region row is served about once and pages gallop: work is O(rows of
    // the region), never O(band x 16,384 columns) or repeated band rebuilds.
    expect(store.probe.cells).toBeLessThanOrEqual(MAX_ROWS + 4096)
    expect(store.probe.calls).toBeLessThanOrEqual(8192)
  })
  it('refuses with a typed resource-limit when the region probe exceeds its work budget', async () => {
    const store = denseColumnStore(MAX_ROWS)
    const result = await detectOperationRange(request(store, emptyIndex(store), sel(MAX_ROWS / 2, 0), { workBudget: 10_000 }))
    expect(result).toEqual({ kind: 'refused', refusal: { reason: 'resource-limit', message: expect.any(String) } })
    expect(store.probe.cells + store.probe.calls).toBeLessThanOrEqual(10_000 + 1024)
  })
  it('refuses malformed work budgets up front', async () => {
    const store = gridStore(headerGrid())
    for (const workBudget of [0, -1, 1.5, Number.NaN]) {
      expect(await detectOperationRange(request(store, emptyIndex(store), sel(1, 0), { workBudget }))).toEqual({
        kind: 'refused', refusal: { reason: 'invalid-range', message: expect.any(String) },
      })
    }
  })
  it('resolves single cells without probing past the band', async () => {
    const rows: (readonly (FixtureCell | undefined)[])[] = []
    rows[5] = [undefined, undefined, undefined, undefined, undefined, fixtureCell('x')]
    const store = gridStore(rows, 6)
    const result = await detectOperationRange(request(store, emptyIndex(store), sel(5, 5)))
    expect(result).toEqual(done(store, { r0: 5, r1: 5, c0: 5, c1: 5, headerRows: 0, totalsRows: 0 }, { kind: 'region' }, 'heuristic'))
    expect(await admitSortRange(request(store, emptyIndex(store), sel(5, 5)))).toEqual(result)
  })
  it('keeps source ranges immutable, freezes the envelope and leaves inputs untouched', async () => {
    const { handle, store, index } = await f5Setup()
    const selection = sel(10, 2)
    const before = JSON.stringify(selection)
    const samples = handle.canonicalSamples()
    const result = await detectOperationRange(request(store, index, selection))
    expect(result.kind).toBe('complete')
    if (result.kind === 'complete') {
      expect(Object.isFrozen(result)).toBe(true)
      expect(Object.isFrozen(result.range)).toBe(true)
      expect(Object.isFrozen(result.stamp)).toBe(true)
      expect(Object.isFrozen(result.source)).toBe(true)
      expect(result.range).toMatchObject({ r0: 4, r1: 20 })
    }
    expect(JSON.stringify(selection)).toBe(before)
    expect(handle.canonicalSamples()).toBe(samples)
  })
  it('pages tiny limits through every scan without changing the outcome', async () => {
    const store = gridStore(headerGrid())
    const specialized = request(store, emptyIndex(store), sel(1, 0), { pageLimit: 2 })
    const expected = done(store, { r0: 0, r1: 3, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 }, { kind: 'region' }, 'heuristic')
    expect(await detectOperationRange(specialized)).toEqual(expected)
    expect(await admitSortRange(specialized)).toEqual(expected)
  })
  it('refuses malformed page limits up front', async () => {
    const store = gridStore(headerGrid())
    for (const pageLimit of [0, -1, 1.5, Number.NaN]) {
      const outcome = await detectOperationRange(request(store, emptyIndex(store), sel(1, 0), { pageLimit }))
      expect(outcome).toEqual({ kind: 'refused', refusal: { reason: 'invalid-range', message: expect.any(String) } })
    }
  })
  it('matches the S3 reference current region on seeded random grids in any page order', async () => {
    let state = 20261011
    const random = () => (state = (state * 1103515245 + 12345) % 2147483648) / 2147483648
    const orders = ['row', 'reverse', 'column'] as const
    let checked = 0
    for (let trial = 0; trial < 60; trial++) {
      const rowCount = 4 + Math.floor(random() * 20), colCount = 4 + Math.floor(random() * 20)
      const density = 0.15 + random() * 0.5
      const cells: Array<readonly [number, number, number]> = []
      for (let row = 0; row < rowCount; row++) {
        for (let col = 0; col < colCount; col++) if (random() < density) cells.push([row, col, row])
      }
      if (cells.length === 0) continue
      const store = sparseStore(cells, { r0: 0, r1: rowCount - 1, c0: 0, c1: colCount - 1 }, orders[trial % 3], 1 + Math.floor(random() * 4))
      const reference = fromStoreScan(store, readOptionsFor(store))
      const band = { r0: 0, r1: MAX_ROWS - 1 }
      const rows = await RowChunkIndex.build(reference, band)
      const cols = await ColumnChunkIndex.build(reference, band)
      const [row, col] = cells[Math.floor(random() * cells.length)]
      const expected = currentRegion({ row, col }, rows.index, cols.index)
      const detected = await detectOperationRange(request(store, emptyIndex(store), sel(row, col), { pageLimit: 1 + Math.floor(random() * 8) }))
      expect(detected.kind).toBe('complete')
      if (detected.kind === 'complete') {
        const { r0, r1, c0, c1 } = detected.range
        expect({ r0, r1, c0, c1 }).toEqual(expected)
      }
      checked++
    }
    expect(checked).toBeGreaterThan(50)
  })
  it('grows a tall region completely past any seed window', async () => {
    const rows: (readonly (FixtureCell | undefined)[])[] = []
    for (let row = 0; row <= 6000; row++) rows.push([fixtureCell(row)])
    const store = gridStore(rows, 6001)
    const detected = await detectOperationRange(request(store, emptyIndex(store), sel(3000, 0)))
    expect(detected).toEqual(done(store, { r0: 0, r1: 6000, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, { kind: 'region' }, 'heuristic'))
    expect(await admitSortRange(request(store, emptyIndex(store), sel(3000, 0)))).toEqual(detected)
  })
  it('waits on structural coverage before trusting table metadata', async () => {
    const handle = createFixture('F5')
    const store = handle.createStore({ pending: [{ r0: 0, r1: 99, c0: 0, c1: 100 }] })
    const result = await store.queryStructural(fullDomain, readOptionsFor(store))
    const index = new StructuralIndex(result)
    expect(index.completeness).not.toBe('complete')
    const outcome = await detectOperationRange(request(store, index, sel(10, 2)))
    expect(outcome.kind).toBe('incomplete')
  })
})
