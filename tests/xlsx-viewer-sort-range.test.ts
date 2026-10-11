/** S7.1 range detection, header policy and structural admission (spec §6.1). */
import { describe, expect, it } from 'vitest'
import { detectOperationRange, admitSortRange, type RangeDetectionRequest } from '../src/xlsx/view/rangeDetection'
import { StructuralIndex } from '../src/xlsx/view/structure'
import type {
  GridSelection,
  SourceStore,
  StructuralResult,
} from '../src/xlsx/view/contracts'
import { MAX_COLS, MAX_ROWS, type SourceRect } from '../src/xlsx/view/coordinates'
import { createFixture, readOptionsFor, type FixtureCell, type FixtureHandle } from './helpers/xlsx-viewer-fixtures'

const fullDomain: SourceRect = { r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }
const fixtureCell = (value: FixtureCell['value'], error = false): FixtureCell => Object.freeze({ value, valueIsError: error, hasCachedValue: true, styleIndex: 0 })
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
function emptyIndex(store: SourceStore): StructuralIndex {
  return new StructuralIndex({ stamp: store.stamp, completeness: 'complete', merges: [], tables: [], spills: [], drawings: [], missing: [] })
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

describe('table precedence and metadata protection', () => {
  it('resolves the containing table with metadata header and totals rows', async () => {
    const { store, index } = await f5Setup()
    const result = await detectOperationRange(request(store, index, sel(10, 2)))
    expect(result).toEqual({ kind: 'complete', range: { r0: 4, r1: 20, c0: 0, c1: 3, headerRows: 1, totalsRows: 1 } })
    if (result.kind === 'complete') expect(Object.isFrozen(result.range)).toBe(true)
  })
  it('honors an explicit table headerRows of 0', async () => {
    const { store, index } = await f5Setup()
    const result = await detectOperationRange(request(store, index, sel(25, 0)))
    expect(result).toEqual({ kind: 'complete', range: { r0: 24, r1: 30, c0: 0, c1: 2, headerRows: 0, totalsRows: 0 } })
  })
  it('never overrides declared table metadata with a dialog override', async () => {
    const { store, index } = await f5Setup()
    const result = await detectOperationRange(request(store, index, sel(10, 2), { override: { headerRows: 0, totalsRows: 0 } }))
    expect(result).toEqual({ kind: 'complete', range: { r0: 4, r1: 20, c0: 0, c1: 3, headerRows: 1, totalsRows: 1 } })
  })
})

describe('non-table header policy', () => {
  it('treats a text header over numbers as a header row', async () => {
    const store = gridStore(headerGrid())
    const result = await detectOperationRange(request(store, emptyIndex(store), sel(1, 0)))
    expect(result).toEqual({ kind: 'complete', range: { r0: 0, r1: 3, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 } })
    const admitted = await admitSortRange(request(store, emptyIndex(store), sel(1, 0)))
    expect(admitted).toEqual({ kind: 'complete', range: { r0: 0, r1: 3, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 } })
  })
  it('treats an all-text first row as data, never an automatic header', async () => {
    const store = gridStore(textGrid())
    const result = await detectOperationRange(request(store, emptyIndex(store), sel(0, 0)))
    expect(result).toEqual({ kind: 'complete', range: { r0: 0, r1: 2, c0: 0, c1: 1, headerRows: 0, totalsRows: 0 } })
    expect(await admitSortRange(request(store, emptyIndex(store), sel(0, 0)))).toEqual(result)
  })
  it('applies an explicit override to the non-table heuristic', async () => {
    const store = gridStore(textGrid())
    const headed = await detectOperationRange(request(store, emptyIndex(store), sel(0, 0), { override: { headerRows: 1, totalsRows: 0 } }))
    expect(headed).toEqual({ kind: 'complete', range: { r0: 0, r1: 2, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 } })
    const tailed = await detectOperationRange(request(store, emptyIndex(store), sel(0, 0), { override: { headerRows: 1, totalsRows: 1 } }))
    expect(tailed).toEqual({ kind: 'complete', range: { r0: 0, r1: 2, c0: 0, c1: 1, headerRows: 1, totalsRows: 1 } })
  })
  it('refuses a malformed header override instead of guessing', async () => {
    const store = gridStore(textGrid())
    const bad = { headerRows: 2, totalsRows: 0 } as unknown as { headerRows: 0 | 1; totalsRows: 0 | 1 }
    expect(await detectOperationRange(request(store, emptyIndex(store), sel(0, 0), { override: bad }))).toEqual({
      kind: 'refused', refusal: { reason: 'invalid-rule', message: expect.any(String) },
    })
  })
  it('types cached formula values normally and never as blank', async () => {
    const rows = [ [fixtureCell('A'), fixtureCell('B')], [{ ...fixtureCell(5), formula: 'SUM(A2:A3)' }, fixtureCell(6)] ]
    const store = gridStore(rows)
    expect(await detectOperationRange(request(store, emptyIndex(store), sel(0, 0)))).toEqual({
      kind: 'complete', range: { r0: 0, r1: 1, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 },
    })
  })
  it('waits on cache-less formula values instead of typing them', async () => {
    const rows = [
      [fixtureCell('A'), { value: null, formula: 'WEBSERVICE()', hasCachedValue: false, valueIsError: false, styleIndex: 0 }],
      [fixtureCell(1), fixtureCell(2)],
    ]
    const store = gridStore(rows, undefined, { valuePolicy: 'cached' })
    const result = await detectOperationRange(request(store, emptyIndex(store), sel(0, 0)))
    expect(result.kind).toBe('incomplete')
    if (result.kind === 'incomplete') expect(result.missing).toEqual([{ r0: 0, r1: 0, c0: 1, c1: 1 }])
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
    if (result.kind === 'complete') expect(result.range).toMatchObject({ r0: 0, r1: 9 })
  })
  it('refuses out-of-domain active cells and used ranges', async () => {
    const { store, index } = await f5Setup()
    const badCell = sel(-1, 0)
    expect(await detectOperationRange(request(store, index, badCell))).toEqual({
      kind: 'refused', refusal: { reason: 'invalid-range', message: expect.any(String) },
    })
    const badUsed = { r0: 5, r1: 3, c0: 0, c1: 0 }
    expect(await detectOperationRange(request(store, index, sel(0, 0), { usedRange: badUsed }))).toEqual({
      kind: 'refused', refusal: { reason: 'invalid-range', message: expect.any(String) },
    })
  })
})

describe('used-range fallback', () => {
  it('falls back to the authoritative used extent for blank seeds', async () => {
    const { store, index } = await f5Setup()
    const result = await detectOperationRange(request(store, index, sel(50, 50)))
    expect(result).toEqual({ kind: 'complete', range: { r0: 0, r1: 199999, c0: 0, c1: 100, headerRows: 0, totalsRows: 0 } })
  })
  it('falls back to the used extent for seeds outside the used rows', async () => {
    const { store, index } = await f5Setup()
    const result = await detectOperationRange(request(store, index, sel(500000, 0)))
    expect(result.kind).toBe('complete')
    if (result.kind === 'complete') expect(result.range).toMatchObject({ r0: 0, r1: 199999 })
  })
  it('admits the used fallback through the shared guard', async () => {
    const { store, index } = await f5Setup()
    expect(await admitSortRange(request(store, index, sel(50, 50)))).toEqual({
      kind: 'refused', refusal: { reason: 'merge', range: { r0: 0, r1: 1, c0: 0, c1: 1 } },
    })
  })
  it('prefers a caller-supplied used extent and waits without one', async () => {
    const store = gridStore(headerGrid())
    const scoped = await detectOperationRange(request(store, emptyIndex(store), sel(0, 0), { usedRange: { r0: 0, r1: 3, c0: 0, c1: 1 } }))
    expect(scoped).toEqual({ kind: 'complete', range: { r0: 0, r1: 3, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 } })
    const noused = { ...store, extent: { rows: 1048576 as const, cols: 16384 as const } }
    const waiting = await detectOperationRange(request(noused, emptyIndex(noused), sel(0, 0)))
    expect(waiting.kind).toBe('incomplete')
    if (waiting.kind === 'incomplete') {
      expect(waiting.missing).toEqual([{ r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }])
    }
  })
  it('waits instead of truncating when data overruns a narrow used override', async () => {
    const store = gridStore(headerGrid())
    const result = await detectOperationRange(request(store, emptyIndex(store), sel(0, 0), { usedRange: { r0: 0, r1: 1, c0: 0, c1: 1 } }))
    expect(result.kind).toBe('incomplete')
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
  it('admits the F5 metadata table through the guard when far data refuses', async () => {
    const { store, index } = await f5Setup()
    expect(await admitSortRange(request(store, index, sel(10, 2)))).toEqual({
      kind: 'refused', refusal: { reason: 'outside-column-data', range: { r0: 5, r1: 5, c0: 4, c1: 4 } },
    })
  })
})

describe('cancellation, bounds and immutability', () => {  it('cancels instead of deciding when the signal is already aborted', async () => {
    const { store, index } = await f5Setup()
    const controller = new AbortController()
    controller.abort()
    const outcome = await detectOperationRange({
      store, structure: index, selection: sel(10, 2), options: { signal: controller.signal, stamp: { ...store.stamp } },
    })
    expect(outcome).toEqual({ kind: 'cancelled' })
  })
  it('resolves single cells without probing past the band', async () => {
    const rows: (readonly (FixtureCell | undefined)[])[] = []
    rows[5] = [undefined, undefined, undefined, undefined, undefined, fixtureCell('x')]
    const store = gridStore(rows, 6)
    const result = await detectOperationRange(request(store, emptyIndex(store), sel(5, 5)))
    expect(result).toEqual({ kind: 'complete', range: { r0: 5, r1: 5, c0: 5, c1: 5, headerRows: 0, totalsRows: 0 } })
    expect(await admitSortRange(request(store, emptyIndex(store), sel(5, 5)))).toEqual(result)
  })
  it('keeps source ranges immutable and inputs untouched', async () => {
    const { handle, store, index } = await f5Setup()
    const selection = sel(10, 2)
    const before = JSON.stringify(selection)
    const samples = handle.canonicalSamples()
    const result = await detectOperationRange(request(store, index, selection))
    expect(result.kind).toBe('complete')
    if (result.kind === 'complete') {
      expect(Object.isFrozen(result.range)).toBe(true)
      expect(result.range).toMatchObject({ r0: 4, r1: 20 })
    }
    expect(JSON.stringify(selection)).toBe(before)
    expect(handle.canonicalSamples()).toBe(samples)
  })
  it('pages tiny limits through every scan without changing the outcome', async () => {
    const store = gridStore(headerGrid())
    const specialized = request(store, emptyIndex(store), sel(1, 0), { pageLimit: 2 })
    expect(await detectOperationRange(specialized)).toEqual({
      kind: 'complete', range: { r0: 0, r1: 3, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 },
    })
    expect(await admitSortRange(specialized)).toEqual({
      kind: 'complete', range: { r0: 0, r1: 3, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 },
    })
  })
  it('refuses malformed page limits up front', async () => {
    const store = gridStore(headerGrid())
    for (const pageLimit of [0, -1, 1.5, Number.NaN]) {
      const outcome = await detectOperationRange(request(store, emptyIndex(store), sel(1, 0), { pageLimit }))
      expect(outcome).toEqual({ kind: 'refused', refusal: { reason: 'invalid-range', message: expect.any(String) } })
    }
  })
  it('widens seed windows until a tall region resolves completely', async () => {
    const rows: (readonly (FixtureCell | undefined)[])[] = []
    for (let row = 0; row <= 6000; row++) rows.push([fixtureCell(row)])
    const store = gridStore(rows, 6001)
    const detected = await detectOperationRange(request(store, emptyIndex(store), sel(3000, 0)))
    expect(detected).toEqual({ kind: 'complete', range: { r0: 0, r1: 6000, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 } })
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
