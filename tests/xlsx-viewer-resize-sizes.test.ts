/**
 * View-only row/column resizing (spec §5.6 amendment): the session size
 * overlay on the full-capacity axes, its sparse storage, the drawing remap and
 * the auto-fit bound. Pure model tests; viewer wiring lives in
 * xlsx-viewer-resize.test.tsx.
 */
import { describe, expect, it } from 'vitest'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { createRowOrder } from '../src/xlsx/view/order'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import {
  AUTOFIT_MAX_CELLS, AUTOFIT_MAX_ROWS, MAX_SESSION_SIZES, RESIZE_LIMITS, SheetSizes, autoFitRowCandidates, clampSize, fitColumnWidth, fitRowHeight,
  mergeAxisOverrides, remapDrawings,
} from '../src/xlsx/view/axisSizes'
import type { XlsxDrawing, XlsxSheet } from '../src/xlsx/types'

const sheet = (extra: Partial<XlsxSheet> = {}): XlsxSheet => ({ name: 'S', rows: [], cols: [], merges: [], mergeRanges: [], ...extra })

describe('session size overlay: sparse merge into authored overrides', () => {
  it('splits an authored interval around a session index (no contradictory overlap)', () => {
    const merged = mergeAxisOverrides([{ start: 5, end: 10, sizeCss: 80 }], new Map([[7, 120]]))
    expect(merged).toEqual([{ start: 5, end: 6, sizeCss: 80 }, { start: 7, end: 7, sizeCss: 120 }, { start: 8, end: 10, sizeCss: 80 }])
    // The merged list is valid axis metadata (createAxisMetrics would throw on overlap).
    const axis = createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64, overrides: merged })
    expect(axis.size(6)).toBe(80)
    expect(axis.size(7)).toBe(120)
    expect(axis.size(8)).toBe(80)
  })

  it('places session indexes outside authored intervals, in order, at interval edges and at the axis ends', () => {
    const merged = mergeAxisOverrides([{ start: 2, end: 4, sizeCss: 30 }], new Map([[MAX_COLS - 1, 200], [4, 50], [0, 9], [2, 11]]))
    expect(merged).toEqual([
      { start: 0, end: 0, sizeCss: 9 }, { start: 2, end: 2, sizeCss: 11 }, { start: 3, end: 3, sizeCss: 30 },
      { start: 4, end: 4, sizeCss: 50 }, { start: MAX_COLS - 1, end: MAX_COLS - 1, sizeCss: 200 },
    ])
  })

  it('never resizes a zero-size (hidden) authored index: hidden stays hidden', () => {
    const merged = mergeAxisOverrides([{ start: 3, end: 5, sizeCss: 0 }], new Map([[4, 100]]))
    expect(merged).toEqual([{ start: 3, end: 5, sizeCss: 0 }])
    // Authored-hidden columns are `collapsed` in the axis and win over overrides.
    const axis = createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64, overrides: mergeAxisOverrides([], new Map([[4, 100]])), collapsed: [{ start: 4, end: 4 }] })
    expect(axis.size(4)).toBe(0)
  })

  it('works on the full-capacity axes without per-index arrays (segment count tracks session entries)', () => {
    const sizes = new Map([[MAX_ROWS - 1, 40], [0, 30], [524_288, 25]])
    const rows = createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20, overrides: mergeAxisOverrides([], sizes) })
    expect(rows.diagnostics.segmentCount).toBeLessThanOrEqual(7)
    expect(rows.totalCss()).toBe(20 * MAX_ROWS + 10 + 5 + 20)
    expect(rows.start(MAX_ROWS - 1)).toBe(20 * (MAX_ROWS - 1) + 10 + 5)
  })

  it('clamps to the documented limits and rounds to whole CSS pixels', () => {
    expect(clampSize('col', 1)).toBe(RESIZE_LIMITS.col.min)
    expect(clampSize('col', 99_999)).toBe(RESIZE_LIMITS.col.max)
    expect(clampSize('row', 0)).toBe(RESIZE_LIMITS.row.min)
    expect(clampSize('row', 10_000)).toBe(RESIZE_LIMITS.row.max)
    expect(clampSize('col', 72.4)).toBe(72)
    expect(clampSize('row', Number.NaN)).toBe(RESIZE_LIMITS.row.min)
  })
})

describe('SheetSizes: per-sheet session store', () => {
  it('bumps its generation on every change, resets to empty, and accounts its bytes', () => {
    const sizes = new SheetSizes()
    expect(sizes.empty).toBe(true)
    expect(sizes.set('col', 3, 100)).toBe(true)
    expect(sizes.gen).toBe(1)
    expect(sizes.set('row', 9, 33)).toBe(true)
    expect(sizes.gen).toBe(2)
    expect(sizes.get('col', 3)).toBe(100)
    expect(sizes.bytes).toBeGreaterThan(0)
    const snapshot = sizes.snapshot()
    sizes.set('col', 3, 120)
    // A published snapshot is immutable: geometry built from it never changes underneath.
    expect(snapshot.cols.get(3)).toBe(100)
    expect(sizes.reset()).toBe(true)
    expect(sizes.empty).toBe(true)
    expect(sizes.bytes).toBe(0)
    expect(sizes.reset()).toBe(false)
  })

  it('refuses a new index past the per-axis cap but still updates existing ones', () => {
    const sizes = new SheetSizes()
    for (let col = 0; col < MAX_SESSION_SIZES; col++) expect(sizes.set('col', col, 70)).toBe(true)
    expect(sizes.set('col', MAX_SESSION_SIZES, 70)).toBe(false)
    expect(sizes.set('col', 0, 90)).toBe(true)
    expect(sizes.set('row', 0, 30)).toBe(true)
  })
})

describe('row sizes follow the SOURCE row under sort and filter', () => {
  it('a session height keyed by source row moves with that row to its sorted display slot', () => {
    const rowsAxis = createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20, overrides: mergeAxisOverrides([], new Map([[1, 50]])) })
    const cols = createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 })
    // Band [0,2] reversed: source 2,1,0 at display 0,1,2 — source row 1 stays at slot 1;
    // swap 0 and 1 instead: source 1 is displayed first.
    const order = createRowOrder([{ r0: 0, r1: 2, sourceAtOffset: Uint32Array.from([1, 0, 2]) }])
    const metrics = createSheetMetrics({ rows: rowsAxis, cols, order })
    expect(metrics.rows.size(0)).toBe(50)
    expect(metrics.rows.size(1)).toBe(20)
    // A filter that hides source row 1 hides its session height too.
    const hidden = createSheetMetrics({ rows: rowsAxis, cols, order, visibility: visibilityFromHidden(createIntervalSet([{ r0: 0, r1: 0 }])) })
    expect(hidden.rows.size(0)).toBe(0)
  })

  it('buildSheetGeometry applies session sizes over authored ones and keeps authored-hidden collapsed', () => {
    const doc = sheet({
      rows: [{ index: 0, cells: [] }, { index: 2, hidden: true, cells: [] }, { index: 4, customHeight: true, heightPt: 30, cells: [] }],
      cols: [{ min: 1, max: 3, widthChars: 10 }, { min: 5, max: 5, hidden: true }],
    })
    const built = buildSheetGeometry(doc, { sizes: { rows: new Map([[4, 66], [2, 99]]), cols: new Map([[2, 150], [5, 150]]) } })
    if (built.kind !== 'ready') throw new Error('not ready')
    const { metrics } = built
    expect(metrics.cols.size(1)).toBe(75)
    expect(metrics.cols.size(2)).toBe(150)
    expect(metrics.cols.size(3)).toBe(75)
    expect(metrics.cols.size(5)).toBe(0)
    expect(metrics.rows.size(4)).toBe(66)
    expect(metrics.rows.size(2)).toBe(0)
  })
})

describe('drawings anchored to cells follow session sizes', () => {
  const node = (treePath: string, xEmu: number, yEmu: number, widthEmu: number, heightEmu: number): XlsxDrawing =>
    ({ xEmu, yEmu, widthEmu, heightEmu, source: { partPath: 'xl/drawings/drawing1.xml', treePath, element: 'sp', representation: 'native' } }) as XlsxDrawing
  const E = 9525
  const base = { rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }) }
  const sized = { rows: base.rows, cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64, overrides: [{ start: 1, end: 1, sizeCss: 128 }] }) }

  it('two-cell anchors move and stretch, one-cell anchors move, absolute anchors stay', () => {
    const two = node('twoCellAnchor[0]/sp[0]', 32 * E, 0, 96 * E, 20 * E) // x 32..128: spans col 0 (half) to col 2 (exact start)
    const one = node('oneCellAnchor[1]/sp[0]', 136 * E, 0, 10 * E, 10 * E) // col 2 offset 8
    const abs = node('absoluteAnchor[2]/sp[0]', 136 * E, 0, 10 * E, 10 * E)
    const [movedTwo, movedOne, kept] = remapDrawings([two, one, abs], base, sized)
    expect(movedTwo.xEmu).toBe(32 * E)
    expect(movedTwo.widthEmu).toBe((192 - 32) * E)
    expect(movedOne.xEmu).toBe(200 * E)
    expect(movedOne.widthEmu).toBe(10 * E)
    expect(kept).toBe(abs)
    // The source nodes are never mutated.
    expect(two.widthEmu).toBe(96 * E)
  })

  it('an offset past a narrowed cell clamps to that cell', () => {
    const narrow = { rows: base.rows, cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64, overrides: [{ start: 0, end: 0, sizeCss: 16 }] }) }
    const [moved] = remapDrawings([node('oneCellAnchor[0]/sp[0]', 40 * E, 0, 5 * E, 5 * E)], base, narrow)
    expect(moved.xEmu).toBe(16 * E)
  })
})

describe('auto-fit measurement and its bound', () => {
  const measure = (_font: string, text: string) => text.length * 7

  it('fits a column to its widest formatted line plus cell padding', () => {
    expect(fitColumnWidth([{ text: 'abc', font: 'f' }, { text: 'abcdefghij', font: 'f' }, { text: 'x\nlonger line', font: 'f' }], measure)).toBe(11 * 7 + 7)
    // Nothing measurable keeps the default width (Excel leaves an empty column alone).
    expect(fitColumnWidth([], measure)).toBe(64)
  })

  it('fits a row to its tallest cell, counting wrapped lines at the column width', () => {
    expect(fitRowHeight([], measure)).toBe(20)
    const plain = fitRowHeight([{ text: 'a', font: 'f', sizePx: 20, wrap: false, widthCss: 64 }], measure)
    expect(plain).toBe(Math.ceil(20 * 1.2 + 4))
    const wrapped = fitRowHeight([{ text: 'aaaa bbbb cccc', font: 'f', sizePx: 10, wrap: true, widthCss: 4 * 7 + 6 }], measure)
    expect(wrapped).toBe(Math.ceil(3 * 10 * 1.2 + 4))
  })

  it('candidate rows: visible pane rows first, then known rows, never more than AUTOFIT_MAX_ROWS', () => {
    const visible = [5, 6, 7]
    const known = function* () { for (let row = 0; row < MAX_ROWS; row++) yield row }
    const rows = autoFitRowCandidates(visible, known(), row => row !== 3)
    expect(rows.slice(0, 3)).toEqual([5, 6, 7])
    expect(rows).toHaveLength(AUTOFIT_MAX_ROWS)
    expect(rows).not.toContain(3)
    expect(new Set(rows).size).toBe(rows.length)
    expect(AUTOFIT_MAX_CELLS).toBeGreaterThan(0)
  })
})
