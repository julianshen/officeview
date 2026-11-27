/** Stage 4c Task 11a: range detection with see-through lines (`RangeDetectionRequest.seeThrough`, owner 2026-11-27).
 * See-through lines are held blank by the store; detection steps over them: they neither end nor extend a region, a
 * selection whose only gaps they are is one rectangle, and the header rule reads the next kept row. Absent, the local
 * rule is unchanged. */
import { describe, expect, it } from 'vitest'
import { detectOperationRange, type RangeDetectionRequest, type SeeThroughLines } from '../src/xlsx/view/rangeDetection'
import { StructuralIndex } from '../src/xlsx/view/structure'
import type { GridSelection, SourceStore } from '../src/xlsx/view/contracts'
import { MAX_ROWS } from '../src/xlsx/view/coordinates'
import { createFixture, readOptionsFor, type FixtureCell } from './helpers/xlsx-viewer-fixtures'

const cell = (value: FixtureCell['value']): FixtureCell => Object.freeze({ value, valueIsError: false, hasCachedValue: true, styleIndex: 0 })
type Grid = (readonly (FixtureCell | undefined)[])[]
function store(rows: Grid, rowCount = Math.max(rows.length, 3)): SourceStore {
  return createFixture('F2', { rowCount }).createStore({
    read: { isPopulated: (r, c) => rows[r]?.[c] !== undefined, cellAt: (r, c) => rows[r]?.[c] ?? null },
  })
}
const index = (s: SourceStore) => new StructuralIndex({ stamp: s.stamp, completeness: 'complete', merges: [], tables: [], spills: [], drawings: [], missing: [] })
const sel = (row: number, col: number, ranges: GridSelection['ranges'] = []): GridSelection => ({ active: { row, col }, anchor: { row, col }, ranges })
const req = (s: SourceStore, selection: GridSelection, seeThrough?: SeeThroughLines): RangeDetectionRequest =>
  ({ store: s, structure: index(s), selection, options: readOptionsFor(s), ...(seeThrough ? { seeThrough } : {}) })
const rows = (...runs: Array<[number, number]>): SeeThroughLines => ({ rows: runs.map(([start, end]) => ({ start, end })), cols: [] })
const cols = (...runs: Array<[number, number]>): SeeThroughLines => ({ rows: [], cols: runs.map(([start, end]) => ({ start, end })) })

/** Header `H`,`I` then numbers in A1:B{n}, with `blank` rows (0-based) held empty, as a key view holds hidden rows. */
function block(n: number, blank: readonly number[] = []): Grid {
  return Array.from({ length: n }, (_, r) => r === 0 ? [cell('H'), cell('I')] : blank.includes(r) ? [] : [cell(r), cell(r * 2)])
}

describe('see-through range detection', () => {
  it('a region steps over see-through rows; without them it stops (local rule)', async () => {
    const s = store(block(10, [3, 6]))
    expect(await detectOperationRange(req(s, sel(0, 0)))).toMatchObject({ kind: 'complete', range: { r0: 0, r1: 2, c0: 0, c1: 1 } })
    expect(await detectOperationRange(req(s, sel(0, 0), rows([3, 3], [6, 6])))).toMatchObject({ kind: 'complete', range: { r0: 0, r1: 9, c0: 0, c1: 1, headerRows: 1 }, source: { kind: 'region' } })
    // Seeded below a see-through row, too.
    expect(await detectOperationRange(req(s, sel(8, 1), rows([3, 3], [6, 6])))).toMatchObject({ range: { r0: 0, r1: 9 } })
  })

  it('see-through rows at the edges neither extend nor end the region; a blank kept row still ends it', async () => {
    const s = store(block(10, [7, 8]), 20)
    expect(await detectOperationRange(req(s, sel(0, 0), rows([7, 8])))).toMatchObject({ range: { r0: 0, r1: 9 } })
    expect(await detectOperationRange(req(s, sel(0, 0), rows([9, 12])))).toMatchObject({ range: { r0: 0, r1: 6 } })
    expect(await detectOperationRange(req(s, sel(0, 0), rows([7, 7])))).toMatchObject({ range: { r0: 0, r1: 6 } })
  })

  it('see-through columns are stepped over the same way', async () => {
    const g: Grid = Array.from({ length: 4 }, (_, r) => [cell(r === 0 ? 'A' : r), undefined, cell(r === 0 ? 'C' : r * 3)])
    const s = store(g)
    expect(await detectOperationRange(req(s, sel(0, 0)))).toMatchObject({ range: { c0: 0, c1: 0 } })
    expect(await detectOperationRange(req(s, sel(0, 0), cols([1, 1])))).toMatchObject({ range: { r0: 0, r1: 3, c0: 0, c1: 2 } })
  })

  it('the header rule reads the next kept row', async () => {
    // Row 1 is see-through (blank); row 2 holds numbers: a header. Without the lens, row 1 reads blank: no header.
    const s = store(block(6, [1]))
    const explicit = sel(4, 0, [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 5 }] }, cols: [{ c0: 0, c1: 1 }] }])
    expect(await detectOperationRange(req(s, explicit))).toMatchObject({ range: { headerRows: 0 } })
    expect(await detectOperationRange(req(s, explicit, rows([1, 1])))).toMatchObject({ range: { headerRows: 1 } })
  })

  it('a selection whose only gaps are see-through is one rectangle, in every row encoding; a kept gap is disjoint', async () => {
    const s = store(block(10, [3, 6]))
    const st = rows([3, 3], [6, 6])
    const encodings: GridSelection['ranges'][number]['rows'][] = [
      { kind: 'intervals', intervals: [{ r0: 1, r1: 2 }, { r0: 4, r1: 5 }, { r0: 7, r1: 8 }] },
      { kind: 'list', r0: 1, r1: 8, rows: Uint32Array.from([1, 2, 4, 5, 7, 8]) },
      { kind: 'bitset', space: 'source', r0: 1, r1: 8, count: 6, words: Uint32Array.from([0b11011011]) },
    ]
    for (const r of encodings) {
      const selection = sel(8, 1, [{ rows: r, cols: [{ c0: 0, c1: 1 }] }])
      expect(await detectOperationRange(req(s, selection))).toMatchObject({ source: { kind: 'region' } })
      expect({ kind: r.kind, got: await detectOperationRange(req(s, selection, st)) }).toMatchObject({ kind: r.kind, got: { range: { r0: 1, r1: 8, c0: 0, c1: 1 }, source: { kind: 'explicit' } } })
    }
    const gap = sel(8, 1, [{ rows: { kind: 'intervals', intervals: [{ r0: 1, r1: 2 }, { r0: 5, r1: 5 }, { r0: 7, r1: 8 }] }, cols: [{ c0: 0, c1: 1 }] }])
    expect(await detectOperationRange(req(s, gap, st))).toMatchObject({ source: { kind: 'region' }, range: { r0: 0, r1: 9 } })
    const twoCols = cols([1, 1])
    const colGap = sel(2, 2, [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 2 }] }, cols: [{ c0: 0, c1: 0 }, { c0: 2, c1: 2 }] }])
    expect(await detectOperationRange(req(store([[cell(1), undefined, cell(2)], [cell(3), undefined, cell(4)], [cell(5), undefined, cell(6)]]), colGap, twoCols))).toMatchObject({ range: { r0: 0, r1: 2, c0: 0, c1: 2 }, source: { kind: 'explicit' } })
  })

  it('a blank seed takes the used fallback through see-through rows; rows to the sheet end are see-through', async () => {
    const s = store(block(10, [3]), 12)
    expect(await detectOperationRange(req(s, sel(11, 5), rows([3, 3])))).toMatchObject({ kind: 'complete', range: { r0: 0, r1: 9 }, source: { kind: 'used' } })
    expect(await detectOperationRange(req(s, sel(11, 5)))).toMatchObject({ kind: 'refused', refusal: { reason: 'ambiguous-selection' } })
    const tail = store(block(5), 8)
    expect(await detectOperationRange(req(tail, sel(0, 0), rows([5, MAX_ROWS - 1])))).toMatchObject({ range: { r0: 0, r1: 4 } })
  })

  it('an empty see-through set is the local rule', async () => {
    const s = store(block(10, [3]))
    expect(await detectOperationRange(req(s, sel(0, 0), { rows: [], cols: [] }))).toEqual(await detectOperationRange(req(s, sel(0, 0))))
  })
})
