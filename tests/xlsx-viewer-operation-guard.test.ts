/** S3.2 operation guard and drawing prefix impact (spec sections 6.3, 8.1; AC-SEM-04, AC-SEM-05). */
import { describe, expect, it } from 'vitest'
import { MAX_COLS, MAX_ROWS, type SourceRect } from '../src/xlsx/view/coordinates'
import type {
  DrawingAnchor,
  FilterRange,
  OperationRefusal,
  RowVisibility,
  StructuralResult,
} from '../src/xlsx/view/contracts'
import { StructuralIndex, fromStoreScan, type PopulatedCellSource, type PopulatedPage } from '../src/xlsx/view/structure'
import {
  guardFilter,
  guardSort,
  isChildOnlyRange,
  refusalKey,
  selectionAllowed,
  type FilterGuardRequest,
} from '../src/xlsx/view/operationGuard'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createIdentityOrder, createRowOrder } from '../src/xlsx/view/order'
import { createRowVisibility } from '../src/xlsx/view/visibility'
import { createSheetMetrics, type ProjectedSheetMetrics } from '../src/xlsx/view/projection'
import { createFixture, readOptionsFor, type FixtureHandle } from './helpers/xlsx-viewer-fixtures'

const fullDomain: SourceRect = { r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }
function rectsIntersect(a: SourceRect, b: SourceRect): boolean {
  return a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1
}
function band(r0: number, r1: number, c0: number, c1: number, headerRows = 0, totalsRows = 0): FilterRange {
  return { r0, r1, c0, c1, headerRows: headerRows as 0 | 1, totalsRows: totalsRows as 0 | 1 }
}

async function f5Setup(): Promise<{ handle: FixtureHandle; index: StructuralIndex }> {
  const handle = createFixture('F5')
  const store = handle.createStore()
  const result = await store.queryStructural(fullDomain, readOptionsFor(store))
  return { handle, index: new StructuralIndex(result) }
}

function fixtureSource(handle: FixtureHandle, pending: readonly SourceRect[] = [], pageSize = 64): PopulatedCellSource {
  const keyOf = (range: SourceRect): string => `${range.r0},${range.r1},${range.c0},${range.c1}`
  return {
    async iterate(range: SourceRect, cursor?: string): Promise<PopulatedPage> {
      let offset = 0
      if (cursor !== undefined) {
        const parts = cursor.split(':')
        offset = Number(parts[2])
        if (parts.length !== 3 || parts[0] !== 'f5pop' || parts[1] !== keyOf(range)
          || !Number.isSafeInteger(offset) || offset < 0 || parts[2] !== String(offset)) {
          throw new RangeError('Invalid populated cursor')
        }
      }
      const cells = handle.scanCells(range, offset, pageSize, handle, [...pending])
      const missing = pending.filter(region => rectsIntersect(region, range))
      return {
        items: cells.map(({ row, col }) => ({ row, col })),
        completeness: missing.length > 0 ? 'loading' : 'complete',
        missing,
        nextCursor: cells.length === pageSize ? `f5pop:${keyOf(range)}:${offset + pageSize}` : undefined,
      }
    },
  }
}

const ROW_COUNT = 100
const ROW_PX = 20
function sheetProjection(visibility?: RowVisibility): ProjectedSheetMetrics {
  return createSheetMetrics({
    rows: createAxisMetrics({ logicalCount: ROW_COUNT, defaultSizeCss: ROW_PX }),
    cols: createAxisMetrics({ logicalCount: 8, defaultSizeCss: 64 }),
    order: createIdentityOrder(),
    visibility,
  })
}
function filterRequest(
  range: FilterRange,
  hidden: readonly number[],
  structure: StructuralIndex,
): FilterGuardRequest {
  const currentVisibility = createRowVisibility(ROW_COUNT)
  const candidateVisibility = createRowVisibility(ROW_COUNT, slot => !hidden.includes(slot))
  return {
    range,
    candidateVisibility,
    currentVisibility,
    projectionBefore: sheetProjection(currentVisibility),
    projectionAfter: sheetProjection(candidateVisibility),
    structure,
  }
}
function drawingBelowBand(): StructuralResult {
  const drawing: DrawingAnchor = {
    id: 'below-band',
    kind: 'twoCell',
    from: { row: 2, col: 0, offsetXEmu: 0, offsetYEmu: 0 },
    to: { row: 4, col: 2, offsetXEmu: 0, offsetYEmu: 0 },
    extent: { x: 0, y: 0, width: 1000, height: 1000 },
    bounds: { r0: 2, r1: 4, c0: 0, c1: 2 },
  }
  return {
    stamp: { docId: 'd', sheetId: 's', valueGen: 0 },
    completeness: 'complete', merges: [], tables: [], spills: [], drawings: [drawing], missing: [],
  }
}

describe('guardSort structural refusals', () => {
  it('refuses a sort intersecting a merge', async () => {
    const { handle, index } = await f5Setup()
    const outcome = await guardSort(band(0, 5, 0, 3), index, fixtureSource(handle))
    expect(outcome).toEqual({ ok: false, refusal: { reason: 'merge', range: { r0: 0, r1: 1, c0: 0, c1: 1 } } })
  })
  it('refuses a far-column blank merge outside the key columns but inside the moved band', async () => {
    const { handle, index } = await f5Setup()
    const outcome = await guardSort(band(50, 150, 0, 1), index, fixtureSource(handle))
    expect(outcome).toEqual({ ok: false, refusal: { reason: 'merge', range: { r0: 99, r1: 199999, c0: 1, c1: 100 } } })
  })
  it('refuses a fixed-array spill through a child-only extent with the anchor outside the range', async () => {
    const { handle, index } = await f5Setup()
    const outcome = await guardSort(band(32, 34, 2, 5), index, fixtureSource(handle))
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.refusal.reason).toBe('fixed-array')
  })
  it('refuses a dynamic spill through a child-only extent with the anchor outside the range', async () => {
    const { handle, index } = await f5Setup()
    const outcome = await guardSort(band(44, 46, 0, 2), index, fixtureSource(handle))
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.refusal).toMatchObject({ reason: 'dynamic-spill' })
  })
  it('refuses outside-column data separated by entirely blank columns', async () => {
    const empty = new StructuralIndex({
      stamp: { docId: 'd', sheetId: 's', valueGen: 0 },
      completeness: 'complete', merges: [], tables: [], spills: [], drawings: [], missing: [],
    })
    const source: PopulatedCellSource = {
      async iterate(range: SourceRect): Promise<PopulatedPage> {
        return {
          items: [{ row: 2, col: 3 }, { row: 4, col: 3 }].filter(
            cell => cell.row >= range.r0 && cell.row <= range.r1),
          completeness: 'complete',
          missing: [],
        }
      },
    }
    const outcome = await guardSort(band(0, 9, 0, 1), empty, source)
    expect(outcome).toEqual({ ok: false, refusal: { reason: 'outside-column-data', range: { r0: 2, r1: 2, c0: 3, c1: 3 } } })
  })
  it('accepts a clean band with no structure, drawings or outside data', async () => {
    const { handle, index } = await f5Setup()
    expect(await guardSort(band(60, 65, 0, 1), index, fixtureSource(handle))).toEqual({ ok: true })
  })
  it('queries only outside columns, never the key columns', async () => {
    const requested: SourceRect[] = []
    const source: PopulatedCellSource = {
      async iterate(range: SourceRect): Promise<PopulatedPage> {
        requested.push({ ...range })
        const items = range.r0 === 0 && range.c0 <= 100 && 100 <= range.c1 ? [{ row: 0, col: 100 }] : []
        return { items, completeness: 'complete', missing: [] }
      },
    }
    const empty = new StructuralIndex({
      stamp: { docId: 'd', sheetId: 's', valueGen: 0 },
      completeness: 'complete', merges: [], tables: [], spills: [], drawings: [], missing: [],
    })
    expect(await guardSort(band(0, 199999, 0, 3), empty, source)).toEqual({
      ok: false, refusal: { reason: 'outside-column-data', range: { r0: 0, r1: 0, c0: 100, c1: 100 } },
    })
    expect(requested.length).toBeGreaterThan(0)
    for (const rect of requested) expect(rect.c0 > 3 || rect.c1 < 0).toBe(true)
  })
  it('stops paging after the first outside cell', async () => {
    let calls = 0
    const source: PopulatedCellSource = {
      async iterate(_range: SourceRect, cursor?: string): Promise<PopulatedPage> {
        calls++
        if (cursor === undefined) {
          return { items: [{ row: 0, col: 100 }], completeness: 'complete', missing: [], nextCursor: 'more' }
        }
        return { items: [{ row: 50, col: 200 }], completeness: 'complete', missing: [] }
      },
    }
    const empty = new StructuralIndex({
      stamp: { docId: 'd', sheetId: 's', valueGen: 0 },
      completeness: 'complete', merges: [], tables: [], spills: [], drawings: [], missing: [],
    })
    expect(await guardSort(band(0, 199999, 0, 3), empty, source)).toEqual({
      ok: false, refusal: { reason: 'outside-column-data', range: { r0: 0, r1: 0, c0: 100, c1: 100 } },
    })
    expect(calls).toBe(1)
  })
  it('finds outside-column data on a later store page instead of truncating', async () => {
    const { handle } = await f5Setup()
    const empty = new StructuralIndex({
      stamp: { docId: 'd', sheetId: 's', valueGen: 0 },
      completeness: 'complete', merges: [], tables: [], spills: [], drawings: [], missing: [],
    })
    for (const pageLimit of [5, 2, 1]) {
      const store = handle.createStore()
      const source = fromStoreScan(store, readOptionsFor(store), pageLimit)
      expect(await guardSort(band(0, 99, 0, 3), empty, source)).toEqual({
        ok: false, refusal: { reason: 'outside-column-data', range: { r0: 0, r1: 0, c0: 6, c1: 6 } },
      })
    }
  })
  it('refuses a sort moving a cell-anchored drawing band', async () => {
    const { handle, index } = await f5Setup()
    const outcome = await guardSort(band(70, 74, 0, 1), index, fixtureSource(handle))
    expect(outcome).toEqual({ ok: false, refusal: { reason: 'drawing-impact', drawingId: 'drawing-1' } })
  })
  it('ignores absolute anchors and accepts an empty body band', async () => {
    const { handle } = await f5Setup()
    const absolute: DrawingAnchor = { id: 'abs', kind: 'absolute', extent: { x: 0, y: 0, width: 10, height: 10 } }
    const result: StructuralResult = {
      stamp: { docId: 'd', sheetId: 's', valueGen: 0 },
      completeness: 'complete', merges: [], tables: [], spills: [], drawings: [absolute], missing: [],
    }
    const empty: PopulatedCellSource = {
      async iterate(): Promise<PopulatedPage> { return { items: [], completeness: 'complete', missing: [] } },
    }
    expect(await guardSort(band(60, 65, 0, 1), new StructuralIndex(result), empty)).toEqual({ ok: true })
    const { index } = await f5Setup()
    expect(await guardSort(band(4, 4, 0, 3, 1, 0), index, fixtureSource(handle))).toEqual({ ok: true })
  })
})

describe('guardFilter merge/spill and drawing prefix impact', () => {
  it('refuses a filter intersecting a merge', async () => {
    const { index } = await f5Setup()
    const outcome = await guardFilter(filterRequest(band(0, 5, 0, 3), [2, 3], index))
    expect(outcome).toEqual({ ok: false, refusal: { reason: 'merge', range: { r0: 0, r1: 1, c0: 0, c1: 1 } } })
  })
  it('refuses a drawing below filtered body rows whose start moves 40 to 0', async () => {
    const index = new StructuralIndex(drawingBelowBand())
    const outcome = await guardFilter(filterRequest(band(0, 1, 0, 3), [0, 1], index))
    expect(outcome).toEqual({ ok: false, refusal: { reason: 'drawing-impact', drawingId: 'below-band' } })
  })
  it('refuses hiding rows within the occupied span', async () => {
    const index = new StructuralIndex(drawingBelowBand())
    const outcome = await guardFilter(filterRequest(band(0, 10, 0, 3), [3], index))
    expect(outcome).toEqual({ ok: false, refusal: { reason: 'drawing-impact', drawingId: 'below-band' } })
  })
  it('accepts below-anchor changes and absolute anchors', async () => {
    const below = new StructuralIndex(drawingBelowBand())
    expect(await guardFilter(filterRequest(band(5, 10, 0, 3), [8, 9], below))).toEqual({ ok: true })
    const absolute: DrawingAnchor = { id: 'abs', kind: 'absolute', extent: { x: 0, y: 0, width: 10, height: 10 } }
    const result: StructuralResult = {
      stamp: { docId: 'd', sheetId: 's', valueGen: 0 },
      completeness: 'complete', merges: [], tables: [], spills: [], drawings: [absolute], missing: [],
    }
    expect(await guardFilter(filterRequest(band(0, 1, 0, 3), [0, 1], new StructuralIndex(result)))).toEqual({ ok: true })
  })
  it('accepts a filter that changes no geometry', async () => {
    const index = new StructuralIndex(drawingBelowBand())
    expect(await guardFilter(filterRequest(band(0, 1, 0, 3), [], index))).toEqual({ ok: true })
  })
  it('returns a typed refusal for drawings anchored outside the projection', async () => {
    const drawing: DrawingAnchor = {
      id: 'past-end',
      kind: 'twoCell',
      from: { row: 5000, col: 0, offsetXEmu: 0, offsetYEmu: 0 },
      to: { row: 5005, col: 2, offsetXEmu: 0, offsetYEmu: 0 },
      extent: { x: 0, y: 0, width: 1000, height: 1000 },
      bounds: { r0: 5000, r1: 5005, c0: 0, c1: 2 },
    }
    const result: StructuralResult = {
      stamp: { docId: 'd', sheetId: 's', valueGen: 0 },
      completeness: 'complete', merges: [], tables: [], spills: [], drawings: [drawing], missing: [],
    }
    const outcome = await guardFilter(filterRequest(band(0, 1, 0, 3), [0], new StructuralIndex(result)))
    expect(outcome).toEqual({ ok: false, refusal: { reason: 'invalid-range', message: expect.any(String) } })
  })
  it('skips drawings beyond the projection when no changed row precedes them', async () => {
    const drawing: DrawingAnchor = {
      id: 'past-end',
      kind: 'oneCell',
      from: { row: 5000, col: 0, offsetXEmu: 0, offsetYEmu: 0 },
      extent: { x: 0, y: 0, width: 1000, height: 1000 },
      bounds: { r0: 5000, r1: 5005, c0: 0, c1: 2 },
    }
    const result: StructuralResult = {
      stamp: { docId: 'd', sheetId: 's', valueGen: 0 },
      completeness: 'complete', merges: [], tables: [], spills: [], drawings: [drawing], missing: [],
    }
    expect(await guardFilter(filterRequest(band(0, 1, 0, 3), [], new StructuralIndex(result)))).toEqual({ ok: true })
  })
  it('refuses a drawing beyond the projection only once a changed row precedes it', async () => {
    const beyond: DrawingAnchor = {
      id: 'past-end',
      kind: 'oneCell',
      from: { row: 5000, col: 0, offsetXEmu: 0, offsetYEmu: 0 },
      extent: { x: 0, y: 0, width: 1000, height: 1000 },
      bounds: { r0: 5000, r1: 5005, c0: 0, c1: 2 },
    }
    const base = drawingBelowBand()
    const result: StructuralResult = { ...base, drawings: [beyond, ...base.drawings] }
    expect(await guardFilter(filterRequest(band(0, 1, 0, 3), [], new StructuralIndex(result)))).toEqual({ ok: true })
    expect(await guardFilter(filterRequest(band(5, 10, 0, 3), [8, 9], new StructuralIndex(result)))).toEqual({
      ok: false, refusal: { reason: 'invalid-range', message: expect.any(String) },
    })
  })
  it('tracks changed rows through a band permutation in source order', async () => {
    const reversed = createRowOrder([{ r0: 0, r1: 9, sourceAtOffset: Uint32Array.from([9, 8, 7, 6, 5, 4, 3, 2, 1, 0]) }])
    const result: StructuralResult = {
      stamp: { docId: 'd', sheetId: 's', valueGen: 0 },
      completeness: 'complete', tables: [], spills: [], drawings: [], missing: [],
      merges: [{ id: 'row-8', range: { r0: 8, r1: 8, c0: 0, c1: 0 }, anchor: { row: 8, col: 0 } }],
    }
    const currentVisibility = createRowVisibility(ROW_COUNT)
    const candidateVisibility = createRowVisibility(ROW_COUNT, slot => slot > 1)
    const axes = () => ({
      rows: createAxisMetrics({ logicalCount: ROW_COUNT, defaultSizeCss: ROW_PX }),
      cols: createAxisMetrics({ logicalCount: 8, defaultSizeCss: 64 }),
    })
    const outcome = await guardFilter({
      range: band(0, 1, 0, 3),
      candidateVisibility,
      currentVisibility,
      projectionBefore: createSheetMetrics({ ...axes(), order: reversed, visibility: currentVisibility }),
      projectionAfter: createSheetMetrics({ ...axes(), order: reversed, visibility: candidateVisibility }),
      structure: new StructuralIndex(result),
    })
    expect(outcome).toEqual({ ok: false, refusal: { reason: 'merge', range: { r0: 8, r1: 8, c0: 0, c1: 0 } } })
  })
})

describe('incomplete coverage waits without authorizing', () => {
  it('waits for incomplete structural metadata', async () => {
    const { handle } = await f5Setup()
    const missing: SourceRect[] = [{ r0: 0, r1: 9, c0: 0, c1: 3 }]
    const store = handle.createStore({ pending: missing })
    const result = await store.queryStructural({ r0: 0, r1: 9, c0: 0, c1: MAX_COLS - 1 }, readOptionsFor(store))
    expect(result.completeness).toBe('loading')
    const outcome = await guardSort(band(0, 9, 0, 1), new StructuralIndex(result), fixtureSource(handle))
    expect(outcome).toEqual({ ok: false, refusal: { reason: 'incomplete', missing } })
  })
  it('waits for unknown occupancy pages instead of faking safety', async () => {
    const { handle, index } = await f5Setup()
    const pending: SourceRect[] = [{ r0: 60, r1: 65, c0: 0, c1: MAX_COLS - 1 }]
    const outcome = await guardSort(band(60, 65, 0, 1), index, fixtureSource(handle, pending))
    expect(outcome).toEqual({ ok: false, refusal: { reason: 'incomplete', missing: pending } })
  })
  it('treats style-only cells as blank through the real store contract', async () => {
    const handle = createFixture('F5')
    const store = handle.createStore({ read: {
      isPopulated: () => false,
      cellAt: () => null,
      styleOf: () => 1,
    } })
    expect(await store.readCell({ row: 4, col: 1 }, readOptionsFor(store))).toEqual({ kind: 'blank' })
    const empty = new StructuralIndex({
      stamp: { docId: 'd', sheetId: 's', valueGen: 0 },
      completeness: 'complete', merges: [], tables: [], spills: [], drawings: [], missing: [],
    })
    expect(await guardSort(band(0, 9, 0, 1), empty, fromStoreScan(store, readOptionsFor(store)))).toEqual({ ok: true })
  })
})

describe('atomic refusal and read-only selection', () => {
  it('changes no view state, order, visibility or stamp on refusal', async () => {
    const request = filterRequest(band(0, 1, 0, 3), [0, 1], new StructuralIndex(drawingBelowBand()))
    const index = request.structure
    const order = request.projectionBefore.order
    const totalCss = request.projectionBefore.rows.totalCss()
    const stampBefore = JSON.stringify(index.stamp)
    const rangeBefore = JSON.stringify(request.range)
    const outcome = await guardFilter(request)
    expect(outcome.ok).toBe(false)
    expect(request.projectionBefore.order).toBe(order)
    expect(request.projectionBefore.rows.totalCss()).toBe(totalCss)
    expect(JSON.stringify(index.stamp)).toBe(stampBefore)
    expect(JSON.stringify(request.range)).toBe(rangeBefore)
  })
  it('always allows read-only selection of spill and merge children', () => {
    expect(selectionAllowed({ row: 45, col: 3 })).toBe(true)
    expect(selectionAllowed({ row: 0, col: 1 })).toBe(true)
    expect(selectionAllowed(null)).toBe(true)
  })
  it('derives child-only range refusals from structural anchors and extents', async () => {
    const { index } = await f5Setup()
    expect(isChildOnlyRange({ r0: 33, r1: 33, c0: 1, c1: 1 }, index)).toMatchObject({ reason: 'fixed-array' })
    expect(isChildOnlyRange({ r0: 45, r1: 45, c0: 3, c1: 3 }, index)).toMatchObject({ reason: 'dynamic-spill' })
    expect(isChildOnlyRange({ r0: 32, r1: 32, c0: 0, c1: 0 }, index)).toBeNull()
    expect(isChildOnlyRange({ r0: 60, r1: 65, c0: 0, c1: 1 }, index)).toBeNull()
  })
  it('waits instead of reporting a safe null for loading structures', async () => {
    const { index } = await f5Setup()
    const missing: SourceRect[] = [{ r0: 0, r1: 9, c0: 0, c1: 3 }]
    const loading = new StructuralIndex({
      stamp: { docId: 'd', sheetId: 's', valueGen: 0 },
      completeness: 'loading', merges: [], tables: [], spills: [], drawings: [], missing,
    })
    expect(loading.completeness).toBe('loading')
    expect(isChildOnlyRange({ r0: 60, r1: 65, c0: 0, c1: 1 }, loading)).toEqual({
      reason: 'incomplete', missing,
    })
    expect(index.completeness).toBe('complete')
  })
})

describe('refusal identity and determinism', () => {
  it('carries a deduplicable key so the status line shows one notice', async () => {
    const { handle, index } = await f5Setup()
    const source = fixtureSource(handle)
    const first = await guardSort(band(0, 5, 0, 3), index, source)
    const second = await guardSort(band(0, 5, 0, 3), index, source)
    expect(first).toEqual(second)
    if (!first.ok && !second.ok) {
      const notices = new Set([refusalKey(first.refusal), refusalKey(second.refusal)])
      expect(notices.size).toBe(1)
      expect(refusalKey(first.refusal)).toBe('merge:0:1:0:1')
    } else {
      throw new Error('expected a merge refusal for the notices test')
    }
    const drawing = await guardSort(band(70, 74, 0, 1), index, source)
    if (!drawing.ok) {
      expect(refusalKey(drawing.refusal)).toBe('drawing-impact:drawing-1')
      expect(refusalKey(drawing.refusal)).not.toBe(refusalKey((first as { ok: false; refusal: OperationRefusal }).refusal))
    } else {
      throw new Error('expected a drawing refusal for the notices test')
    }
  })
  it('refuses an invalid pageLimit up front instead of throwing a raw RangeError', async () => {
    const { handle, index } = await f5Setup()
    const source = fixtureSource(handle)
    for (const pageLimit of [0, -1, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
      expect(await guardSort(band(60, 65, 0, 1), index, source, { pageLimit }))
        .toEqual({ ok: false, refusal: { reason: 'invalid-range', message: 'Invalid chunk page limit' } })
    }
    // Up front means before structural checks: a merge-intersecting band still reports the option error.
    expect(await guardSort(band(0, 5, 0, 3), index, source, { pageLimit: 0 }))
      .toEqual({ ok: false, refusal: { reason: 'invalid-range', message: 'Invalid chunk page limit' } })
    expect(await guardSort(band(60, 65, 0, 1), index, source, { pageLimit: 16 })).toEqual({ ok: true })
  })
  it('refuses invalid ranges instead of throwing', async () => {
    const { handle, index } = await f5Setup()
    const source = fixtureSource(handle)
    expect(await guardSort({ r0: 5, r1: 3, c0: 0, c1: 1, headerRows: 0, totalsRows: 0 }, index, source))
      .toEqual({ ok: false, refusal: { reason: 'invalid-range', message: 'Invalid inclusive source rectangle' } })
    expect(await guardSort({ r0: 0, r1: 5, c0: 0, c1: 1, headerRows: 2 as unknown as 0, totalsRows: 0 }, index, source))
      .toEqual({ ok: false, refusal: { reason: 'invalid-range', message: 'Invalid protected row counts' } })
  })
})
