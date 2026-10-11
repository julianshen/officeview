import { describe, expect, it } from 'vitest'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { createRowOrder } from '../src/xlsx/view/order'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { emptyGridSelection, extendDisplayRange, selectCell, selectHeader, SourceSelection, selectionOperationRange, guardSelectionSort, guardSelectionFilter, reconcileSelection, moveActive, navigateTo, activeAnnouncement, recordBody, recordCounts, SheetNavigationSession, rowContains } from '../src/xlsx/view/selection'
import { StructuralIndex } from '../src/xlsx/view/structure'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'

function fixtureProjection(rows: number[], hiddenRows: number[] = [], hiddenCols: number[] = []) {
  return createSheetMetrics({
    rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24, collapsed: hiddenRows.map(start => ({ start, end: start })) }),
    cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 80, collapsed: hiddenCols.map(start => ({ start, end: start })) }),
    order: createRowOrder([{ r0: 0, r1: rows.length - 1, sourceAtOffset: Uint32Array.from(rows) }]),
  })
}
describe('source membership through display extension (AC-SEL-01)', () => {
  it('preserves the plan counterexample verbatim', () => {
    const projection = fixtureProjection([0, 3, 1, 2])
    // extendDisplayRange(projection, anchor: SourceCell, target: SourceCell); both arguments are SOURCE cells
    const selection = extendDisplayRange(projection, { row: 3, col: 0 }, { row: 1, col: 0 }) // source 3 = slot 1, source 1 = slot 2
    expect([...selection.sourceRows()].sort((a, b) => a - b)).toEqual([1, 3])
    expect(selection.contains({ row: 2, col: 0 })).toBe(false)
    const wide = extendDisplayRange(projection, { row: 0, col: 0 }, { row: 1, col: 0 }) // source 0 = slot 0, source 1 = slot 2
    expect([...wide.sourceRows()].sort((a, b) => a - b)).toEqual([0, 1, 3])
    expect(wide.contains({ row: 2, col: 0 })).toBe(false)
  })
  it('excludes hidden rows and columns, and preserves source membership after a later sort', () => {
    const selection = extendDisplayRange(fixtureProjection([0, 3, 1, 2], [3], [1]), { row: 0, col: 0 }, { row: 2, col: 2 })
    expect([...selection.sourceRows()]).toEqual([0, 1, 2])
    expect(selection.contains({ row: 3, col: 0 })).toBe(false)
    expect(selection.contains({ row: 1, col: 1 })).toBe(false)
    fixtureProjection([0, 1, 2, 3])
    expect(selection.contains({ row: 2, col: 2 })).toBe(true)
  })
  it('retains full-axis identity ranges compactly and validates endpoints', () => {
    const projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 80 }) })
    const selected = extendDisplayRange(projection, { row: 0, col: 0 }, { row: MAX_ROWS - 1, col: MAX_COLS - 1 })
    expect(selected.ranges).toEqual([{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: MAX_ROWS - 1 }] }, cols: [{ c0: 0, c1: MAX_COLS - 1 }] }])
    expect(selected.contains({ row: MAX_ROWS - 1, col: MAX_COLS - 1 })).toBe(true)
    expect(() => extendDisplayRange(projection, { row: MAX_ROWS, col: 0 }, { row: 0, col: 0 })).toThrow()
  })
})

describe('selection boundary regressions', () => {
  it('matches an independent cell-set oracle over 200 toggle/Shift histories without overlap', () => {
    const order = [0, 3, 1, 2, 7, 6, 4, 5], projection = fixtureProjection(order)
    let state = emptyGridSelection(), anchor: { row: number; col: number } | null = null, expected = new Set<number>(), seed = 61
    const random = (n: number) => { seed = seed * 48271 % 2147483647; return seed % n }
    for (let action = 0; action < 200; action++) {
      const at = { row: random(8), col: random(4) }, mode = random(3)
      if (mode === 1) {
        const id = at.row * 4 + at.col
        if (expected.has(id)) expected.delete(id); else expected.add(id)
        state = selectCell(state, projection, at, { toggle: true })
        anchor = at
      } else if (mode === 2 && anchor) {
        expected = new Set()
        const first = order.indexOf(anchor.row), last = order.indexOf(at.row)
        for (let slot = Math.min(first, last); slot <= Math.max(first, last); slot++) {
          for (let col = Math.min(anchor.col, at.col); col <= Math.max(anchor.col, at.col); col++) expected.add(order[slot] * 4 + col)
        }
        state = selectCell(state, projection, at, { extend: true })
      } else {
        expected = new Set([at.row * 4 + at.col])
        state = selectCell(state, projection, at)
        anchor = at
      }
      const actual = new SourceSelection(state.ranges)
      for (let row = 0; row < 8; row++) for (let col = 0; col < 4; col++) {
        expect(actual.contains({ row, col })).toBe(expected.has(row * 4 + col))
        expect(state.ranges.filter(part => rowContains(part.rows, row) && part.cols.some(c => col >= c.c0 && col <= c.c1)).length).toBeLessThanOrEqual(1)
      }
    }
  })
  it('derives records from a resolved used-data body, excluding header/totals and implicit blanks', () => {
    const range = { r0: 0, r1: 3, c0: 0, c1: 0, headerRows: 1 as const, totalsRows: 1 as const }
    expect(recordBody(range)).toEqual({ r0: 1, r1: 2 })
    expect(recordCounts(fixtureProjection([0, 1, 2, 3]), recordBody(range))).toEqual({ visible: 2, total: 2 })
    expect(recordBody({ ...range, r1: 0, totalsRows: 0 })).toBeNull()
    expect(recordBody(null)).toBeNull()
  })
  it('preserves a legal outside-body row when only its column becomes hidden', () => {
    const before = fixtureProjection([0, 1, 2, 3]), after = fixtureProjection([0, 1, 2, 3], [], [1])
    const state = selectCell(emptyGridSelection(), before, { row: 1048575, col: 1 })
    expect(reconcileSelection(state, after, { r0: 1, r1: 3 }).selection.active).toEqual({ row: 1048575, col: 2 })
  })
  it('relocates an authored-hidden blank outside the body locally rather than to a distant record', () => {
    const before = fixtureProjection([0, 1, 2, 3]), after = fixtureProjection([0, 1, 2, 3], [1048575])
    const state = selectCell(emptyGridSelection(), before, { row: 1048575, col: 0 })
    expect(reconcileSelection(state, after, { r0: 1, r1: 3 }).selection.active).toEqual({ row: 1048574, col: 0 })
  })
  it('reanchors a hidden Shift anchor even when active remains visible', () => {
    const before = fixtureProjection([0, 3, 1, 2])
    let state = selectCell(emptyGridSelection(), before, { row: 3, col: 0 })
    state = selectCell(state, before, { row: 1, col: 0 }, { extend: true })
    const after = fixtureProjection([0, 3, 1, 2], [3])
    const reconciled = reconcileSelection(state, after, { r0: 1, r1: 3 })
    expect(reconciled.selection.anchor).toEqual(state.active)
    expect(reconciled.selection.ranges).toBe(state.ranges)
    const next = selectCell(state, after, { row: 2, col: 0 }, { extend: true })
    expect([...new SourceSelection(next.ranges).sourceRows()]).toEqual([1, 2])
  })
  it('keeps 100k fragmented rows in a bounded source-tagged bitset', () => {
    const rows = [...Array.from({ length: 100000 }, (_, i) => i * 2), ...Array.from({ length: 100000 }, (_, i) => i * 2 + 1)]
    const selected = extendDisplayRange(fixtureProjection(rows), { row: 0, col: 0 }, { row: 199998, col: 0 })
    expect(selected.ranges[0].rows.kind).toBe('bitset')
    const set = selected.ranges[0].rows
    if (set.kind !== 'bitset') throw new Error('Expected source bitset')
    expect(set.space).toBe('source')
    expect(set.count).toBe(100000)
    expect(set.words.byteLength).toBeLessThanOrEqual(131072)
    expect(selected.contains({ row: 199998, col: 0 })).toBe(true)
    expect(selected.contains({ row: 199999, col: 0 })).toBe(false)
  })
  it('returns a typed invalid-range refusal before operation guards can publish', async () => {
    const state = selectCell(emptyGridSelection(), fixtureProjection([0, 1, 2, 3]), { row: 1, col: 0 })
    const range = { r0: -1, r1: 2, c0: 0, c1: 1, headerRows: 0 as const, totalsRows: 0 as const }
    expect(selectionOperationRange(state, range)).toMatchObject({ ok: false, refusal: { reason: 'invalid-range' } })
  })
  it('delegates candidate filter admission to the same S3 guard', async () => {
    const before = fixtureProjection([0, 1, 2, 3]), after = fixtureProjection([0, 1, 2, 3], [2])
    const state = selectCell(emptyGridSelection(), before, { row: 2, col: 0 })
    const structure = new StructuralIndex({ stamp: { docId: 'd', sheetId: 's', valueGen: 0 }, completeness: 'complete', merges: [{ id: 'm', range: { r0: 1, r1: 2, c0: 0, c1: 1 }, anchor: { row: 1, col: 0 } }], spills: [], drawings: [], tables: [], missing: [] })
    const hidden = createIntervalSet([{ r0: 2, r1: 2 }])
    const result = await guardSelectionFilter(state, { range: { r0: 1, r1: 2, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, structure, currentVisibility: null, candidateVisibility: visibilityFromHidden(hidden), projectionBefore: before, projectionAfter: after })
    expect(result).toMatchObject({ ok: false, refusal: { reason: 'merge' } })
    expect(state.active).toEqual({ row: 2, col: 0 })
  })
})

describe('focus, navigation and retained sheet state (AC-SEL-02, AC-CAP-02)', () => {
  const body = { r0: 1, r1: 3 }
  it('moves arrows/Enter through visible display order, skipping hidden rows/columns', () => {
    const projection = fixtureProjection([0, 3, 1, 2], [1], [1])
    const state = selectCell(emptyGridSelection(), projection, { row: 3, col: 0 })
    const next = moveActive(state, projection, 'enter', body)
    expect(next.active).toEqual({ row: 2, col: 0 })
    expect(moveActive(next, projection, 'up', body).active).toEqual({ row: 3, col: 0 })
    expect(moveActive(state, projection, 'right', body).active).toEqual({ row: 3, col: 2 })
    const extended = moveActive(state, projection, 'down', body, true)
    expect([...new SourceSelection(extended.ranges).sourceRows()]).toEqual([2, 3])
  })
  it('preserves source active and selection after a later permutation', () => {
    const state = selectCell(emptyGridSelection(), fixtureProjection([0, 3, 1, 2]), { row: 3, col: 0 })
    const reconciled = reconcileSelection(state, fixtureProjection([0, 1, 2, 3]), body)
    expect(reconciled.selection).toBe(state)
    expect(reconciled.selection.active).toEqual({ row: 3, col: 0 })
  })
  it('relocates to the nearest eligible body cell, preferring forward on a tie', () => {
    const state = selectCell(emptyGridSelection(), fixtureProjection([0, 1, 2, 3]), { row: 2, col: 0 })
    const result = reconcileSelection(state, fixtureProjection([0, 1, 2, 3], [2]), body)
    expect(result.selection.active).toEqual({ row: 3, col: 0 })
    expect(result.selection.anchor).toEqual({ row: 3, col: 0 })
    expect(result.selection.ranges).toBe(state.ranges)
    expect(result.announcement).toContain('A4')
  })
  it('uses null active/anchor when every body record is hidden, and never resurrects on Clear', () => {
    const before = fixtureProjection([0, 3, 1, 2]), hidden = fixtureProjection([0, 3, 1, 2], [1, 2, 3])
    const state = selectCell(emptyGridSelection(), before, { row: 3, col: 0 })
    const result = reconcileSelection(state, hidden, body)
    expect(result.selection.active).toBeNull()
    expect(result.selection.anchor).toBeNull()
    expect(result.selection.ranges).toBe(state.ranges)
    expect(result.announcement).toBe('No visible records.')
    expect(reconcileSelection(result.selection, before, body).selection.active).toBeNull()
    expect(moveActive(result.selection, before, 'down', body).active).toEqual({ row: 3, col: 0 })
  })
  it('distinguishes Display A2 from Original A2 under [0,2,1]', () => {
    const projection = fixtureProjection([0, 2, 1])
    const displayed = navigateTo(emptyGridSelection(), projection, { addressSpace: 'display', at: { rowSlot: 1, col: 0 } }, { r0: 1, r1: 2 })
    expect(displayed.target).toEqual({ row: 2, col: 0 })
    expect(displayed.selection.active).toEqual({ row: 2, col: 0 })
    expect(activeAnnouncement(projection, displayed.target, '10')).toBe('A2, value 10; original A3')
    const original = navigateTo(displayed.selection, projection, { addressSpace: 'source', at: { row: 1, col: 0 } }, { r0: 1, r1: 2 })
    expect(original.announcement).toBe('A3; original A2')
  })
  it('returns an exact hidden inspection target without unhiding or inventing active ownership', () => {
    const projection = fixtureProjection([0, 2, 1], [2])
    const state = selectCell(emptyGridSelection(), projection, { row: 1, col: 0 })
    const result = navigateTo(state, projection, { addressSpace: 'source', at: { row: 2, col: 0 } }, { r0: 1, r1: 2 })
    expect(result.kind).toBe('hidden')
    expect(result.target).toEqual({ row: 2, col: 0 })
    expect(result.hiddenReason).toContain('row')
    expect(result.selection).toBe(state)
    expect(projection.sourceCellToLogical(result.target)).toBeNull()
    expect(navigateTo(emptyGridSelection(), projection, { addressSpace: 'source', at: result.target }, body).selection.active).toBeNull()
  })
  it('reaches legal blank edge coordinates outside a filtered body without overflow', () => {
    const projection = fixtureProjection([0, 2, 1]), target = { row: 1048575, col: 16383 }
    const result = navigateTo(emptyGridSelection(), projection, { addressSpace: 'display', at: { rowSlot: target.row, col: target.col } }, { r0: 1, r1: 2 })
    expect(result.kind).toBe('visible')
    expect(result.target).toEqual(target)
    expect(result.announcement).toBe('XFD1048576')
    expect(moveActive(result.selection, projection, 'home', body).active).toEqual({ row: 0, col: 0 })
    expect(moveActive(result.selection, projection, 'end', body).active).toEqual(target)
  })
  it('counts only the resolved body, including header-only and empty cases', () => {
    expect(recordCounts(fixtureProjection([0, 3, 1, 2], [1]), body)).toEqual({ total: 3, visible: 2 })
    expect(recordCounts(fixtureProjection([0, 1, 2, 3]), null)).toEqual({ total: 0, visible: 0 })
    expect(recordCounts(fixtureProjection([0, 1, 2, 3], [1, 2, 3]), body)).toEqual({ total: 3, visible: 0 })
  })
  it('retains selection, view and absolute zoom per sheet and resets them on document replacement', () => {
    const projection = fixtureProjection([0, 3, 1, 2]), session = new SheetNavigationSession('doc')
    const view = (sheetId: string) => ({ sheetId, revision: 1, order: projection.order, baseHidden: createIntervalSet([]), visibility: visibilityFromHidden(createIntervalSet([]), MAX_ROWS), sort: null, filter: null, selection: emptyGridSelection() })
    session.register('a', { projection, body, view: view('a') })
    session.register('b', { projection, body, view: view('b') })
    session.activate('a')
    session.setSelection('a', selectCell(emptyGridSelection(), projection, { row: 3, col: 0 }))
    session.setZoom('a', 2)
    session.activate('b')
    expect(session.get('b').zoom).toBe(1)
    const restored = session.activate('a')
    expect(restored.zoom).toBe(2)
    expect(restored.selection.active).toEqual({ row: 3, col: 0 })
    expect(restored.view.revision).toBe(1)
    session.detachProjection('a')
    expect(session.get('a').projection).toBeNull()
    expect(session.get('a').selection.active).toEqual({ row: 3, col: 0 })
    expect(session.get('a').zoom).toBe(2)
    session.updateProjection('a', projection, body)
    expect(session.get('a').projection).toBe(projection)
    session.replaceDocument('next')
    expect(session.activeSheetId).toBeNull()
    expect(() => session.get('a')).toThrow()
  })
})

describe('pointer, modifier and header transitions', () => {
  it('keeps active/anchor source-owned, extends through display order and toggles an exact cell', () => {
    const projection = fixtureProjection([0, 3, 1, 2])
    let state = selectCell(emptyGridSelection(), projection, { row: 3, col: 0 })
    expect(state.active).toEqual({ row: 3, col: 0 })
    state = selectCell(state, projection, { row: 1, col: 0 }, { extend: true })
    expect(state.anchor).toEqual({ row: 3, col: 0 })
    expect([...new SourceSelection(state.ranges).sourceRows()]).toEqual([1, 3])
    state = selectCell(state, projection, { row: 1, col: 0 }, { toggle: true })
    expect(new SourceSelection(state.ranges).contains({ row: 1, col: 0 })).toBe(false)
    expect(new SourceSelection(state.ranges).contains({ row: 3, col: 0 })).toBe(true)
    state = selectCell(state, projection, { row: 2, col: 1 }, { toggle: true })
    expect(new SourceSelection(state.ranges).contains({ row: 2, col: 1 })).toBe(true)
  })
  it('selects full logical headers compactly and removes only the toggled cell', () => {
    const projection = fixtureProjection([0, 3, 1, 2])
    let state = selectHeader(emptyGridSelection(), projection, 'column', 16383)
    expect(state.ranges).toEqual([{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 1048575 }] }, cols: [{ c0: 16383, c1: 16383 }] }])
    state = selectCell(state, projection, { row: 1048575, col: 16383 }, { toggle: true })
    expect(new SourceSelection(state.ranges).contains({ row: 1048574, col: 16383 })).toBe(true)
    expect(new SourceSelection(state.ranges).contains({ row: 1048575, col: 16383 })).toBe(false)
    const row = selectHeader(state, projection, 'row', 1048575)
    expect(row.ranges).toEqual([{ rows: { kind: 'intervals', intervals: [{ r0: 1048575, r1: 1048575 }] }, cols: [{ c0: 0, c1: 16383 }] }])
    expect(new SourceSelection(row.ranges).contains({ row: 1048575, col: 16383 })).toBe(true)
  })
  it('deduplicates intersecting header membership and toggles an entire selected header off', () => {
    const projection = fixtureProjection([0, 3, 1, 2])
    let state = selectHeader(emptyGridSelection(), projection, 'column', 0)
    state = selectHeader(state, projection, 'row', 1, true)
    state = selectHeader(state, projection, 'row', 1, true)
    expect(new SourceSelection(state.ranges).contains({ row: 1, col: 0 })).toBe(false)
    expect(new SourceSelection(state.ranges).contains({ row: 2, col: 0 })).toBe(true)
    expect(new SourceSelection(state.ranges).contains({ row: 1, col: 1 })).toBe(false)
  })
  it('refuses to coerce disjoint source membership into a bounding operation range', () => {
    const ranges = extendDisplayRange(fixtureProjection([0, 3, 1, 2]), { row: 3, col: 0 }, { row: 1, col: 0 }).ranges
    const state = { active: { row: 1, col: 0 }, anchor: { row: 3, col: 0 }, ranges }
    expect(selectionOperationRange(state)).toMatchObject({ ok: false, refusal: { reason: 'ambiguous-selection' } })
    const range = { r0: 1, r1: 3, c0: 0, c1: 0, headerRows: 0 as const, totalsRows: 0 as const }
    expect(selectionOperationRange(state, range)).toEqual({ ok: true, range })
  })
  it('allows highlighting merge children but delegates operation refusal to S3', async () => {
    const projection = fixtureProjection([0, 1, 2, 3]), state = selectCell(emptyGridSelection(), projection, { row: 2, col: 0 })
    expect(new SourceSelection(state.ranges).contains({ row: 2, col: 0 })).toBe(true)
    const structure = new StructuralIndex({ stamp: { docId: 'd', sheetId: 's', valueGen: 0 }, completeness: 'complete', merges: [{ id: 'm', range: { r0: 1, r1: 2, c0: 0, c1: 1 }, anchor: { row: 1, col: 0 } }], spills: [], drawings: [], tables: [], missing: [] })
    const outcome = await guardSelectionSort(state, structure, { async iterate() { return { items: [], completeness: 'complete', missing: [] } } })
    expect(outcome).toMatchObject({ ok: false, refusal: { reason: 'merge' } })
  })
})
