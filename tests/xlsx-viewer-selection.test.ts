import { describe, expect, it, vi } from 'vitest'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { createRowOrder, pinRowPermutation, hiddenSlotPins } from '../src/xlsx/view/order'
import { createSheetMetrics, type ProjectedSheetMetrics } from '../src/xlsx/view/projection'
import { emptyGridSelection, extendDisplayRange, selectCell, selectHeader, SourceSelection, selectionOperationRange, guardSelectionSort, guardSelectionFilter, reconcileSelection, moveActive, navigateTo, activeAnnouncement, recordBody, recordCounts, SheetNavigationSession, rowContains, selectionBytes, MAX_SELECTION_BYTES, MAX_SELECTION_PARTS } from '../src/xlsx/view/selection'
import { StructuralIndex } from '../src/xlsx/view/structure'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'

function authoredView(projection: ProjectedSheetMetrics) {
  const hidden = createIntervalSet(projection.rows.overrides.filter(run => run.sizeCss === 0).map(run => ({ r0: run.start, r1: run.end })))
  return { baseHidden: hidden, visibility: visibilityFromHidden(hidden, projection.rows.logicalCount), filter: null }
}

function filteredFixture(rows: number[], hiddenSlots: number[]) {
  const order = createRowOrder([{ r0: 0, r1: rows.length - 1, sourceAtOffset: Uint32Array.from(rows) }]), visibility = visibilityFromHidden(createIntervalSet(hiddenSlots.map(r0 => ({ r0, r1: r0 }))))
  const projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 80 }), order, visibility })
  const view = { visibility, baseHidden: createIntervalSet([]), filter: { range: { r0: 0, r1: rows.length - 1, c0: 0, c1: 0, headerRows: 1 as const, totalsRows: 0 as const }, columns: new Map([[0, { kind: 'number' as const, op: 'gt' as const, first: 0 }]]) } }
  return { projection, view }
}

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
    const sorted = fixtureProjection([0, 1, 2, 3])
    const state = { active: { row: 2, col: 2 }, anchor: { row: 0, col: 0 }, ranges: selection.ranges }
    const reconciled = reconcileSelection(state, sorted, authoredView(sorted)).selection
    expect([...new SourceSelection(reconciled.ranges).sourceRows()]).toEqual([0, 1, 2])
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
  it('replaces only the most recent Ctrl part during repeated Shift extension (review 5)', () => {
    const p = fixtureProjection([0, 1, 2, 3])
    let state = selectCell(emptyGridSelection(), p, { row: 0, col: 0 }, { toggle: true })
    state = selectCell(state, p, { row: 4, col: 2 }, { toggle: true })
    state = selectCell(state, p, { row: 0, col: 0 }, { extend: true })
    state = selectCell(state, p, { row: 7, col: 2 }, { extend: true })
    const selection = new SourceSelection(state.ranges)
    expect(selection.contains({ row: 0, col: 0 })).toBe(true)
    expect(selection.contains({ row: 0, col: 1 })).toBe(false)
    expect(selection.contains({ row: 7, col: 2 })).toBe(true)
    state = selectCell(state, p, { row: 1, col: 2 }, { extend: true })
    expect(new SourceSelection(state.ranges).contains({ row: 7, col: 2 })).toBe(false)
    expect(new SourceSelection(state.ranges).contains({ row: 0, col: 0 })).toBe(true)
  })
  it('preserves the Ctrl prefix after spread/structuredClone and merged rows (N2)', () => {
    const p = fixtureProjection([0, 1, 2, 3])
    let state = selectCell(emptyGridSelection(), p, { row: 0, col: 0 }, { toggle: true })
    state = selectCell(state, p, { row: 4, col: 0 }, { toggle: true })
    expect(state.ranges).toHaveLength(1) // equal columns merge the two row sets
    for (const copied of [{ ...state }, structuredClone(state)]) {
      const next = selectCell(copied, p, { row: 2, col: 0 }, { extend: true })
      expect([...new SourceSelection(next.ranges).sourceRows()]).toEqual([0, 2, 3, 4])
      const shrunk = selectCell(structuredClone(next), p, { row: 3, col: 0 }, { extend: true })
      expect([...new SourceSelection(shrunk.ranges).sourceRows()]).toEqual([0, 3, 4])
    }
  })
  it('serializes header mode and restores a temporarily overlapped Ctrl prefix (N2)', () => {
    const p = fixtureProjection([0, 1, 2, 3])
    let state = selectCell(emptyGridSelection(), p, { row: 0, col: 0 }, { toggle: true })
    state = selectHeader(state, p, 'row', 4, true)
    state = selectCell(structuredClone(state), p, { row: 0, col: 2 }, { extend: true })
    state = selectCell(structuredClone(state), p, { row: 3, col: 2 }, { extend: true })
    const actual = new SourceSelection(state.ranges)
    expect(actual.contains({ row: 0, col: 0 })).toBe(true)
    expect(actual.contains({ row: 0, col: 1 })).toBe(false)
    expect(actual.contains({ row: 3, col: MAX_COLS - 1 })).toBe(true)
  })
  it('preserves row header anchors when col0 is hidden and Shift-click extends full rows (review 5)', () => {
    const p = fixtureProjection([0, 1, 2, 3], [], [0])
    let state = selectHeader(emptyGridSelection(), p, 'row', 5)
    expect(state.anchor).toEqual({ row: 5, col: 1 })
    state = selectCell(state, p, { row: 9, col: 3 }, { extend: true })
    expect([...new SourceSelection(state.ranges).sourceRows()]).toEqual([5, 6, 7, 8, 9])
    expect(new SourceSelection(state.ranges).contains({ row: 7, col: MAX_COLS - 1 })).toBe(true)
  })
  it('supports Shift across row and column headers through the visible projection (review 5)', () => {
    const p = fixtureProjection([0, 3, 1, 2], [], [1])
    let rows = selectHeader(emptyGridSelection(), p, 'row', 3)
    rows = selectHeader(rows, p, 'row', 1, false, true)
    expect([...new SourceSelection(rows.ranges).sourceRows()]).toEqual([1, 3])
    expect(new SourceSelection(rows.ranges).contains({ row: 1, col: MAX_COLS - 1 })).toBe(true)
    let cols = selectHeader(emptyGridSelection(), p, 'column', 0)
    cols = selectHeader(cols, p, 'column', 3, false, true)
    expect(cols.anchor?.col).toBe(0)
    expect(new SourceSelection(cols.ranges).contains({ row: MAX_ROWS - 1, col: 2 })).toBe(true)
    expect(new SourceSelection(cols.ranges).contains({ row: MAX_ROWS - 1, col: 1 })).toBe(false)
  })
  it('retains one selection truth and rejects stale projection revisions atomically (review 8)', () => {
    const p = fixtureProjection([0, 1, 2, 3]), session = new SheetNavigationSession('doc')
    const view = { sheetId: 's', revision: 5, order: p.order, baseHidden: createIntervalSet([]), visibility: visibilityFromHidden(createIntervalSet([])), sort: null, filter: null, selection: emptyGridSelection() }
    session.register('s', { projection: p, body: null, view })
    const selected = selectCell(emptyGridSelection(), p, { row: 2, col: 0 })
    session.setSelection('s', selected)
    expect(session.get('s').view.selection).toBe(selected)
    const previous = session.get('s')
    expect(() => session.updateProjection('s', fixtureProjection([0, 3, 1, 2]), null, { ...view, revision: 4 })).toThrow(expect.objectContaining({ code: 'stale-source' }))
    expect(session.get('s')).toBe(previous)
    session.updateProjection('s', p, null, { ...view, revision: 6 })
    expect(session.get('s').view.selection).toBe(session.get('s').selection)
    expect(session.get('s').selection.active).toEqual({ row: 2, col: 0 })
  })
  it('uses S2 authored-hidden relocation rather than inferring a filter from used-data bounds (review 4)', () => {
    const state = selectCell(emptyGridSelection(), fixtureProjection([0, 1, 2, 3]), { row: 1, col: 0 })
    expect(reconcileSelection(state, fixtureProjection([0, 1, 2, 3], [1, 2, 3]), authoredView(fixtureProjection([0, 1, 2, 3], [1, 2, 3]))).selection.active).toEqual({ row: 0, col: 0 })
  })
  it('reports the typed S2 filter-hidden reason and relocates through the actual display mask (review 4/6)', () => {
    const f = filteredFixture([0, 3, 1, 2], [2])
    const state = selectCell(emptyGridSelection(), fixtureProjection([0, 3, 1, 2]), { row: 1, col: 0 })
    const result = navigateTo(state, f.projection, { addressSpace: 'source', at: { row: 1, col: 0 } }, f.view)
    expect(result.hiddenReason).toBe('filter-hidden')
    expect(result.selection.active).toEqual({ row: 2, col: 0 })
    expect(recordCounts(f.projection, { r0: 1, r1: 3 })).toEqual({ visible: 2, total: 3 })
    const selection = extendDisplayRange(f.projection, { row: 0, col: 0 }, { row: 2, col: 0 })
    expect([...selection.sourceRows()]).toEqual([0, 2, 3])
  })
  it('requires the actual filter view and clears focus when every body record is hidden (N3)', () => {
    const state = selectCell(emptyGridSelection(), fixtureProjection([0, 1, 2, 3]), { row: 2, col: 0 })
    const filtered = filteredFixture([0, 1, 2, 3], [1, 2, 3])
    const result = reconcileSelection(state, filtered.projection, filtered.view)
    expect(result.selection.active).toBeNull()
    expect(result.selection.anchor).toBeNull()
    expect(result.selection.ranges).toBe(state.ranges)
    expect(result.announcement).toBe('No visible records.')
  })
  it('traverses a genuine hidden-pinned permutation and compares membership to a display oracle (review 6)', () => {
    const visibility = visibilityFromHidden(createIntervalSet([{ r0: 1, r1: 1 }, { r0: 4, r1: 4 }]))
    const band = pinRowPermutation(0, 5, hiddenSlotPins(visibility), Uint32Array.from([5, 3, 2, 0]))
    const order = createRowOrder([{ r0: 0, r1: 5, sourceAtOffset: band.sourceAtOffset }])
    const p = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 80 }), order, visibility })
    const selected = extendDisplayRange(p, { row: 5, col: 0 }, { row: 2, col: 0 })
    const oracle = [0, 1, 2, 3].filter(slot => visibility.visible(slot)).map(slot => order.displayToSource(slot)).sort((a, b) => a - b)
    expect([...selected.sourceRows()]).toEqual(oracle)
    const state = selectCell(emptyGridSelection(), p, { row: 5, col: 0 })
    expect(moveActive(state, p, 'down', { r0: 0, r1: 5 }, authoredView(p)).active?.row).toBe(3)
  })
  it('seeds arrow navigation on an empty sheet with no resolved body (review 3)', () => {
    expect(moveActive(emptyGridSelection(), fixtureProjection([0, 1, 2, 3]), 'down', null, authoredView(fixtureProjection([0, 1, 2, 3]))).active).toEqual({ row: 0, col: 0 })
  })
  it('shares fragmented row membership across 200 Ctrl-column toggles (review 2)', () => {
    const half = MAX_ROWS / 2, order = new Uint32Array(MAX_ROWS)
    for (let i = 0; i < half; i++) { order[i] = i * 2; order[half + i] = i * 2 + 1 }
    const p = fixtureProjection(Array.from(order))
    let state = selectCell(emptyGridSelection(), p, { row: 0, col: 0 })
    state = selectCell(state, p, { row: MAX_ROWS - 2, col: 400 }, { extend: true })
    for (let col = 1; col < 400; col += 2) state = selectHeader(state, p, 'column', col, true)
    const buffers = new Set(state.ranges.flatMap(part => part.rows.kind === 'bitset' ? [part.rows.words.buffer] : []))
    expect(state.ranges.length).toBeLessThanOrEqual(3)
    expect([...buffers].reduce((sum, buffer) => sum + buffer.byteLength, 0)).toBeLessThanOrEqual(262144)
    const selection = new SourceSelection(state.ranges)
    expect(selection.contains({ row: 10, col: 2 })).toBe(true)
    expect(selection.contains({ row: 11, col: 2 })).toBe(false)
    expect(selection.contains({ row: 11, col: 3 })).toBe(true)
  }, 30000)
  it('admits 1000 scattered Ctrl cells within the per-sheet byte bound (N1)', () => {
    const p = fixtureProjection([0, 1, 2, 3])
    let state = emptyGridSelection()
    expect(() => {
      for (let i = 0; i < 1000; i++) state = selectCell(state, p, { row: i * 3, col: i * 7 }, { toggle: true })
    }).not.toThrow()
    const selected = new SourceSelection(state.ranges)
    for (let i = 0; i < 1000; i++) expect(selected.contains({ row: i * 3, col: i * 7 })).toBe(true)
    expect(selected.contains({ row: 3, col: 0 })).toBe(false)
    expect(selectionBytes(state.ranges)).toBeLessThan(MAX_SELECTION_BYTES)
  }, 30000)
  it('admits all 16384 Ctrl column-header toggles compactly (N1)', () => {
    const p = fixtureProjection([0, 1, 2, 3])
    let state = emptyGridSelection()
    // Visit alternate columns first to exercise thousands of disjoint intervals.
    expect(() => {
      for (let parity = 0; parity < 2; parity++) for (let col = parity; col < MAX_COLS; col += 2) state = selectHeader(state, p, 'column', col, true)
    }).not.toThrow()
    expect(state.ranges).toEqual([{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: MAX_ROWS - 1 }] }, cols: [{ c0: 0, c1: MAX_COLS - 1 }] }])
    expect(selectionBytes(state.ranges)).toBeLessThan(1024)
  }, 60000)
  it('retains 50 independently admitted fragmented Shift selections without a document cap (N1)', () => {
    const half = MAX_ROWS / 2, rows = new Uint32Array(MAX_ROWS)
    for (let i = 0; i < half; i++) { rows[i] = i * 2; rows[half + i] = i * 2 + 1 }
    const p = fixtureProjection(Array.from(rows)), session = new SheetNavigationSession('doc')
    let total = 0
    expect(() => {
      for (let i = 0; i < 50; i++) {
        const sheetId = String(i)
        let selection = selectCell(emptyGridSelection(), p, { row: 0, col: 0 })
        selection = selectCell(selection, p, { row: MAX_ROWS - 2, col: 0 }, { extend: true })
        const view = { sheetId, revision: 1, order: p.order, baseHidden: createIntervalSet([]), visibility: visibilityFromHidden(createIntervalSet([])), sort: null, filter: null, selection }
        session.register(sheetId, { view, projection: p, body: null })
        session.setSelection(sheetId, selection)
        session.updateProjection(sheetId, p, null)
        expect(session.get(sheetId).selection).toBe(selection)
        expect(selectionBytes(selection.ranges)).toBeLessThan(MAX_SELECTION_BYTES)
        total += selectionBytes(selection.ranges)
      }
    }).not.toThrow()
    expect(total).toBe(50 * 131440)
    expect(total).toBeLessThan(7 * 1024 * 1024)
  }, 30000)
  it('refuses only beyond the generous safety/byte bound without mutating membership (N1)', () => {
    const ranges = Array.from({ length: 4096 }, (_, i) => ({ rows: { kind: 'intervals' as const, intervals: [{ r0: i, r1: i }] }, cols: [{ c0: i, c1: i }] }))
    expect(MAX_SELECTION_PARTS).toBeGreaterThanOrEqual(4096)
    expect(() => new SourceSelection(ranges)).not.toThrow()
    expect(() => new SourceSelection([...ranges, ranges[0]])).toThrow(expect.objectContaining({ code: 'resource-limit' }))
    expect(ranges).toHaveLength(4096)
  })
  it('reuses the immutable fragmented-row fingerprint across Ctrl toggles (NIT)', () => {
    const p = fixtureProjection([0, 1, 2, 3]), words = new Uint32Array(MAX_ROWS / 32).fill(0x55555555)
    const iterations = vi.fn(() => Uint32Array.prototype[Symbol.iterator].call(words))
    Object.defineProperty(words, Symbol.iterator, { value: iterations })
    const rows = { kind: 'bitset' as const, space: 'source' as const, r0: 0, r1: MAX_ROWS - 2, count: MAX_ROWS / 2, words }
    let state: ReturnType<typeof emptyGridSelection> = { active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 }, ranges: [{ rows, cols: [{ c0: 0, c1: 0 }] }] }
    // Warm the cache without touching the large row mask's membership.
    state = selectCell(state, p, { row: 1, col: 1 }, { toggle: true })
    const warmed = iterations.mock.calls.length
    for (let col = 2; col < 102; col++) state = selectCell(state, p, { row: 1, col }, { toggle: true })
    expect(iterations.mock.calls.length).toBe(warmed)
    expect(state.ranges.find(part => part.rows === rows)?.rows).toBe(rows)
    const selected = new SourceSelection(state.ranges)
    expect(selected.contains({ row: 0, col: 0 })).toBe(true)
    expect(selected.contains({ row: 1, col: 0 })).toBe(false)
    expect(selected.contains({ row: 1, col: 101 })).toBe(true)
  })
  it('clamps fragmented bitset bounds to actual source membership on a 1000-row projection (review 11)', () => {
    const rows = Array.from({ length: 1000 }, (_, i) => i), random = (() => { let seed = 7; return (n: number) => { seed = seed * 48271 % 2147483647; return seed % n } })()
    for (let i = 999; i > 0; i--) { const j = random(i + 1); [rows[i], rows[j]] = [rows[j], rows[i]] }
    const p = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 1000, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: 10, defaultSizeCss: 80 }), order: createRowOrder([{ r0: 0, r1: 999, sourceAtOffset: Uint32Array.from(rows) }]) })
    const set = extendDisplayRange(p, { row: rows[0], col: 0 }, { row: rows[300], col: 0 }).ranges[0].rows
    expect(set.kind).toBe('bitset')
    if (set.kind !== 'bitset') throw new Error('Expected bitset')
    expect(set.r1).toBeLessThan(1000)
    expect(set.words.length).toBe(Math.ceil((set.r1 - set.r0 + 1) / 32))
  })
  it('returns a typed unsupported outcome for opaque row refs (review 10)', () => {
    const ranges = [{ rows: { kind: 'ref' as const, refId: 'worker', count: 1, stamp: { docId: 'd', sheetId: 's', valueGen: 1 } }, cols: [{ c0: 0, c1: 0 }] }]
    expect(() => new SourceSelection(ranges).contains({ row: 0, col: 0 })).toThrow(expect.objectContaining({ code: 'unsupported' }))
    expect(() => [...new SourceSelection(ranges).sourceRows()]).toThrow(expect.objectContaining({ code: 'unsupported' }))
    expect(() => selectCell({ active: null, anchor: null, ranges }, fixtureProjection([0, 1, 2, 3]), { row: 0, col: 0 }, { toggle: true })).toThrow(expect.objectContaining({ code: 'unsupported' }))
  })
  it('requires an S7-resolved operation range for single-cell and column-header selections (review 1)', () => {
    const p = fixtureProjection([0, 1, 2, 3])
    for (const selection of [selectCell(emptyGridSelection(), p, { row: 4, col: 2 }), selectHeader(emptyGridSelection(), p, 'column', 2)]) {
      expect(selectionOperationRange(selection)).toMatchObject({ ok: false, refusal: { reason: 'ambiguous-selection' } })
      const table = { r0: 0, r1: 8, c0: 2, c1: 2, headerRows: 1 as const, totalsRows: 1 as const }
      expect(selectionOperationRange(selection, table)).toEqual({ ok: true, range: table })
      expect(selectionOperationRange(selection, { ...table, c0: 3, c1: 3 })).toMatchObject({ ok: false, refusal: { reason: 'ambiguous-selection' } })
    }
  })
  it('matches an independent cell-set oracle over 200 toggle/Shift histories without overlap', () => {
    const order = [0, 3, 1, 2, 7, 6, 4, 5], projection = fixtureProjection(order)
    let state = emptyGridSelection(), anchor: { row: number; col: number } | null = null, expected = new Set<number>(), prefix = new Set<number>(), seed = 61
    const random = (n: number) => { seed = seed * 48271 % 2147483647; return seed % n }
    for (let action = 0; action < 200; action++) {
      const at = { row: random(8), col: random(4) }, mode = random(3)
      if (mode === 1) {
        const id = at.row * 4 + at.col
        prefix = new Set(expected)
        prefix.delete(id)
        if (expected.has(id)) expected.delete(id); else expected.add(id)
        state = selectCell(state, projection, at, { toggle: true })
        anchor = at
      } else if (mode === 2 && anchor) {
        expected = new Set(prefix)
        const first = order.indexOf(anchor.row), last = order.indexOf(at.row)
        for (let slot = Math.min(first, last); slot <= Math.max(first, last); slot++) {
          for (let col = Math.min(anchor.col, at.col); col <= Math.max(anchor.col, at.col); col++) expected.add(order[slot] * 4 + col)
        }
        state = selectCell(state, projection, at, { extend: true })
      } else {
        expected = new Set([at.row * 4 + at.col])
        prefix = new Set()
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
    expect(reconcileSelection(state, after, authoredView(after)).selection.active).toEqual({ row: 1048575, col: 2 })
  })
  it('relocates an authored-hidden blank outside the body locally rather than to a distant record', () => {
    const before = fixtureProjection([0, 1, 2, 3]), after = fixtureProjection([0, 1, 2, 3], [1048575])
    const state = selectCell(emptyGridSelection(), before, { row: 1048575, col: 0 })
    expect(reconcileSelection(state, after, authoredView(after)).selection.active).toEqual({ row: 1048574, col: 0 })
  })
  it('reanchors a hidden Shift anchor even when active remains visible', () => {
    const before = fixtureProjection([0, 3, 1, 2])
    let state = selectCell(emptyGridSelection(), before, { row: 3, col: 0 })
    state = selectCell(state, before, { row: 1, col: 0 }, { extend: true })
    const after = fixtureProjection([0, 3, 1, 2], [3])
    const reconciled = reconcileSelection(state, after, authoredView(after))
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
    const next = moveActive(state, projection, 'enter', body, authoredView(projection))
    expect(next.active).toEqual({ row: 2, col: 0 })
    expect(moveActive(next, projection, 'up', body, authoredView(projection)).active).toEqual({ row: 3, col: 0 })
    expect(moveActive(state, projection, 'right', body, authoredView(projection)).active).toEqual({ row: 3, col: 2 })
    const extended = moveActive(state, projection, 'down', body, authoredView(projection), true)
    expect([...new SourceSelection(extended.ranges).sourceRows()]).toEqual([2, 3])
  })
  it('preserves source active and selection after a later permutation', () => {
    const state = selectCell(emptyGridSelection(), fixtureProjection([0, 3, 1, 2]), { row: 3, col: 0 })
    const reconciled = reconcileSelection(state, fixtureProjection([0, 1, 2, 3]), authoredView(fixtureProjection([0, 1, 2, 3])))
    expect(reconciled.selection).toBe(state)
    expect(reconciled.selection.active).toEqual({ row: 3, col: 0 })
  })
  it('relocates to the nearest eligible body cell, preferring forward on a tie', () => {
    const state = selectCell(emptyGridSelection(), fixtureProjection([0, 1, 2, 3]), { row: 2, col: 0 })
    const result = reconcileSelection(state, fixtureProjection([0, 1, 2, 3], [2]), authoredView(fixtureProjection([0, 1, 2, 3], [2])))
    expect(result.selection.active).toEqual({ row: 3, col: 0 })
    expect(result.selection.anchor).toEqual({ row: 3, col: 0 })
    expect(result.selection.ranges).toBe(state.ranges)
    expect(result.announcement).toContain('A4')
  })
  it('uses null active/anchor when every body record is hidden, and never resurrects on Clear', () => {
    const before = fixtureProjection([0, 3, 1, 2]), filtered = filteredFixture([0, 3, 1, 2], [1, 2, 3]), hidden = filtered.projection
    const state = selectCell(emptyGridSelection(), before, { row: 3, col: 0 })
    const result = reconcileSelection(state, hidden, filtered.view)
    expect(result.selection.active).toBeNull()
    expect(result.selection.anchor).toBeNull()
    expect(result.selection.ranges).toBe(state.ranges)
    expect(result.announcement).toBe('No visible records.')
    expect(reconcileSelection(result.selection, before, authoredView(before)).selection.active).toBeNull()
    expect(moveActive(result.selection, before, 'down', body, authoredView(before)).active).toEqual({ row: 3, col: 0 })
  })
  it('distinguishes Display A2 from Original A2 under [0,2,1]', () => {
    const projection = fixtureProjection([0, 2, 1])
    const displayed = navigateTo(emptyGridSelection(), projection, { addressSpace: 'display', at: { rowSlot: 1, col: 0 } }, authoredView(projection))
    expect(displayed.target).toEqual({ row: 2, col: 0 })
    expect(displayed.selection.active).toEqual({ row: 2, col: 0 })
    expect(activeAnnouncement(projection, displayed.target, '10')).toBe('A2, value 10; original A3')
    const original = navigateTo(displayed.selection, projection, { addressSpace: 'source', at: { row: 1, col: 0 } }, authoredView(projection))
    expect(original.announcement).toBe('A3; original A2')
  })
  it('returns an exact hidden inspection target without unhiding or inventing active ownership', () => {
    const projection = fixtureProjection([0, 2, 1], [2])
    const state = selectCell(emptyGridSelection(), projection, { row: 1, col: 0 })
    const result = navigateTo(state, projection, { addressSpace: 'source', at: { row: 2, col: 0 } }, authoredView(projection))
    expect(result.kind).toBe('hidden')
    expect(result.target).toEqual({ row: 2, col: 0 })
    expect(result.hiddenReason).toBe('authored-hidden')
    expect(result.selection).toBe(state)
    expect(projection.sourceCellToLogical(result.target)).toBeNull()
    expect(navigateTo(emptyGridSelection(), projection, { addressSpace: 'source', at: result.target }, authoredView(projection)).selection.active).toBeNull()
  })
  it('reaches legal blank edge coordinates outside a filtered body without overflow', () => {
    const projection = fixtureProjection([0, 2, 1]), target = { row: 1048575, col: 16383 }
    const result = navigateTo(emptyGridSelection(), projection, { addressSpace: 'display', at: { rowSlot: target.row, col: target.col } }, authoredView(projection))
    expect(result.kind).toBe('visible')
    expect(result.target).toEqual(target)
    expect(result.announcement).toBe('XFD1048576')
    expect(moveActive(result.selection, projection, 'home', body, authoredView(projection)).active).toEqual({ row: 0, col: 0 })
    expect(moveActive(result.selection, projection, 'end', body, authoredView(projection)).active).toEqual(target)
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
    const outcome = await guardSelectionSort(state, structure, { async iterate() { return { items: [], completeness: 'complete', missing: [] } } }, { r0: 2, r1: 2, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 })
    expect(outcome).toMatchObject({ ok: false, refusal: { reason: 'merge' } })
  })
})
