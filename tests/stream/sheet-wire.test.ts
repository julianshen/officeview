/** Stage 4a Task 6: the `sheet-layout` codec (S4 §6.4, R14) and the client geometry built from it. */
import { describe, expect, it } from 'vitest'
import {
  decodeSheetLayout, encodeSheetLayout, LAYOUT_LIMITS, LayoutInvalid, LayoutTooLarge, SHEET_MAX_COLS, SHEET_MAX_ROWS,
  type StreamSheetLayout,
} from '../../src/stream/sheet-wire'
import { MAX_COLS, MAX_ROWS } from '../../src/xlsx/view/coordinates'
import { streamSheetGeometry } from '../../src/xlsx/view/streamGeometry'

const merges = (n: number) => Array.from({ length: n }, (_, i) => ({ r0: i, r1: i, c0: 0, c1: 1 }))
const layout = (o: Partial<StreamSheetLayout> = {}): StreamSheetLayout => ({
  sheetId: '0', viewGen: 0, defaults: { rowPx: 20, colPx: 64 },
  rows: [{ start: 2, end: 4, sizeCss: 40 }], cols: [{ start: 1, end: 1, sizeCss: 150.5 }],
  collapsedRows: [{ start: 6, end: 6 }], collapsedCols: [{ start: 3, end: 5 }],
  frozen: { rows: 1, cols: 1, topRow: 1, leftCol: 1 },
  merges: merges(3), used: { r0: 0, r1: 9, c0: 0, c1: 4 },
  filters: [{ kind: 'autoFilter', range: { r0: 0, r1: 8, c0: 0, c1: 2 }, header: 1, totals: 0 }],
  sort: null, filterSummary: [], authorHidden: true, counts: { visibleRows: 9, totalRows: 10 },
  ...o,
})
/** The parts as the client receives them (ws.ts adds the reqId). */
const wire = (l: StreamSheetLayout) => encodeSheetLayout(l).map((p) => JSON.parse(JSON.stringify({ reqId: 7, ...p })) as Record<string, any>)
const rejects = (parts: unknown[]) => expect(() => decodeSheetLayout(parts)).toThrow(LayoutInvalid)

describe('sheet-layout codec', () => {
  it('pins the axis bounds to the engine (F5)', () => {
    expect(SHEET_MAX_ROWS).toBe(MAX_ROWS)
    expect(SHEET_MAX_COLS).toBe(MAX_COLS)
    expect(LAYOUT_LIMITS).toEqual({ intervals: 65_536, merges: 65_536, filters: 1_024, parts: 64 })
  })

  it('round-trips, one part and many', () => {
    const small = layout()
    expect(wire(small)).toHaveLength(1)
    expect(decodeSheetLayout(wire(small))).toEqual(small)
    const big = layout({ merges: merges(5_000) })
    const parts = wire(big)
    expect(parts.length).toBeGreaterThanOrEqual(3)
    for (const [i, p] of parts.entries()) {
      expect(new TextEncoder().encode(JSON.stringify(p)).length).toBeLessThanOrEqual(65_536)
      expect(p.seq).toBe(i)
      expect(p.last).toBe(i === parts.length - 1)
    }
    expect(decodeSheetLayout(parts)).toEqual(big)
    const { frozen: _f, used: _u, ...bare } = layout()
    expect(decodeSheetLayout(wire(bare))).toEqual(bare)
  })

  it('refuses a layout over a cap', () => {
    expect(() => encodeSheetLayout(layout({ merges: merges(LAYOUT_LIMITS.merges + 1) }))).toThrow(LayoutTooLarge)
    const rows = Array.from({ length: LAYOUT_LIMITS.intervals + 1 }, (_, i) => ({ start: 2 * i, end: 2 * i, sizeCss: 30 }))
    expect(() => encodeSheetLayout(layout({ rows }))).toThrow(LayoutTooLarge)
    const filters = Array.from({ length: LAYOUT_LIMITS.filters + 1 }, (_, i) => ({ kind: 'table' as const, range: { r0: i, r1: i, c0: 0, c1: 0 }, header: 1 as const, totals: 0 as const }))
    expect(() => encodeSheetLayout(layout({ filters }))).toThrow(LayoutTooLarge)
    // Every list under its cap, the message count over 64: too large as well.
    const max = Array.from({ length: LAYOUT_LIMITS.intervals }, (_, i) => ({ start: 2 * i, end: 2 * i, sizeCss: 30.123456789 }))
    expect(() => encodeSheetLayout(layout({ rows: max, merges: merges(LAYOUT_LIMITS.merges) }))).toThrow(LayoutTooLarge)
  })

  it('rejects a missing seq 1', () => {
    const parts = wire(layout({ merges: merges(5_000) }))
    rejects([parts[0], ...parts.slice(2)])
  })
  it('rejects a head on seq 1', () => {
    const parts = wire(layout({ merges: merges(5_000) }))
    parts[1]!.head = parts[0]!.head
    rejects(parts)
  })
  it('rejects an extra key values', () => {
    const top = wire(layout())
    top[0]!.values = ['x']
    rejects(top)
    const nested = wire(layout())
    nested[0]!.rows[0].values = ['x']
    rejects(nested)
    const head = wire(layout())
    head[0]!.head.counts.values = 1
    rejects(head)
  })
  it('rejects rows[0].start = -1', () => {
    const parts = wire(layout())
    parts[0]!.rows[0].start = -1
    rejects(parts)
  })
  it('rejects merges over the cap', () => {
    const over = merges(LAYOUT_LIMITS.merges + 1)
    // Spread over the parts a real answer could use, so only the cap is wrong.
    const head = wire(layout({ merges: [] }))[0]!
    const parts: Record<string, unknown>[] = []
    for (let i = 0; i < 40; i++) parts.push({ reqId: 7, type: 'sheet-layout', seq: i, last: false, ...(i === 0 ? { head: head.head, rows: head.rows, cols: head.cols, collapsedRows: head.collapsedRows, collapsedCols: head.collapsedCols, filters: head.filters } : {}), merges: over.slice(i * 1_639, (i + 1) * 1_639) })
    parts[39]!.last = true
    expect(parts.reduce((n, p) => n + (p.merges as unknown[]).length, 0)).toBe(LAYOUT_LIMITS.merges + 1)
    rejects(parts)
  })
  it('rejects a second last', () => {
    const parts = wire(layout({ merges: merges(5_000) }))
    parts[0]!.last = true
    rejects(parts)
    rejects([...parts.slice(0, -1), { ...parts.at(-1)!, last: false }])
  })
  it('rejects non-integers and indices out of the sheet', () => {
    const cases: Array<(p: Record<string, any>) => void> = [
      (p) => { p.rows[0].end = 2.5 },
      (p) => { p.collapsedRows[0].start = SHEET_MAX_ROWS },
      (p) => { p.cols[0].start = SHEET_MAX_COLS },
      (p) => { p.merges[0].c1 = SHEET_MAX_COLS },
      (p) => { p.merges[0].r0 = 2 },
      (p) => { p.head.viewGen = 0.5 },
      (p) => { p.head.counts.totalRows = SHEET_MAX_ROWS + 1 },
      (p) => { p.head.sort = { col: 1 } },
      (p) => { p.rows[0].sizeCss = -1 },
      (p) => { p.filters[0].header = 2 },
      (p) => { p.type = 'sheet-other' },
    ]
    for (const edit of cases) {
      const parts = wire(layout())
      edit(parts[0]!)
      rejects(parts)
    }
    rejects([])
  })
})

describe('streamSheetGeometry', () => {
  it('rebuilds the axes and merges, with no cells or drawings', () => {
    const g = streamSheetGeometry(layout())
    expect(g.kind).toBe('ready')
    expect(g.metrics.rows.start(5)).toBe(2 * 20 + 3 * 40)
    expect(g.metrics.rows.size(6)).toBe(0)
    expect(g.metrics.cols.start(2)).toBe(64 + 150.5)
    expect(g.metrics.cols.size(4)).toBe(0)
    expect(g.metrics.rows.logicalCount).toBe(MAX_ROWS)
    expect(g.metrics.cols.logicalCount).toBe(MAX_COLS)
    expect(g.cellsOf(0)).toBeUndefined()
    expect(g.drawingsIn(0, 0, 1e9, 1e9, 1_000)).toEqual([])
    expect(g.merges.collect({ r0: 1, r1: 1, c0: 0, c1: 5 })).toEqual([{ minRow: 1, maxRow: 1, minCol: 0, maxCol: 1 }])
  })
})

// ---- Stage 4b Task 2: the sheet-view codecs (append only) ----
import {
  A11Y_MAX_CELLS, a11yArgs, copyArgs, decodeA11yCells, decodeFindHits, decodeFitResult, decodeSheetView, decodeWireSelection,
  encodeWireSelection, FIND_MAX_HITS, findArgs, fitArgs, QUERY_MAX_UNITS, resizeArgs, SELECTION_LIMITS, SHEET_VIEW_FEATURE,
  type SheetResizeArgs, type SheetResizeRequest, type WireSelection,
} from '../../src/stream/sheet-wire'
import { InvalidOperand } from '../../src/stream/ops'

describe('sheet-view wire constants', () => {
  it('pins the 4b constants', () => {
    expect(SHEET_VIEW_FEATURE).toBe('sheet-view@1')
    expect(SELECTION_LIMITS).toEqual({ ranges: 256 })
    expect([A11Y_MAX_CELLS, FIND_MAX_HITS, QUERY_MAX_UNITS]).toEqual([2_000, 5_000, 256])
  })
})

const sel3 = (): WireSelection => ({
  active: [2, 1], anchor: [0, 0],
  ranges: [
    { rows: [[0, 3]], cols: [[0, 1]] },
    { rows: [[5, 5], [9, 12]], cols: [[4, 4]] },
    { rows: [[1_048_575, 1_048_575]], cols: [[0, 16_383]] },
  ],
})
const roundTrip = (v: unknown) => JSON.parse(JSON.stringify(v)) as unknown

describe('wire selection', () => {
  it('selection round-trips; hostile selections are refused', () => {
    const s = sel3()
    expect(decodeWireSelection(roundTrip(encodeWireSelection(s)))).toEqual(s)
    expect(decodeWireSelection(roundTrip(encodeWireSelection({ active: null, anchor: null, ranges: [] })))).toEqual({ active: null, anchor: null, ranges: [] })
    const bad = (mut: (x: any) => void) => { const x = roundTrip(encodeWireSelection(sel3())) as any; mut(x); return decodeWireSelection(x) }
    expect(bad((x) => { x.ranges = Array.from({ length: 257 }, (_, i) => ({ rows: [[i, i]], cols: [[0, 0]] })) })).toBeUndefined()
    const max = Array.from({ length: 256 }, (_, i) => ({ rows: [[i, i]], cols: [[0, 0]] }))
    expect(bad((x) => { x.ranges = max })).toBeDefined()
    expect(bad((x) => { x.ranges[0].rows = [[3, 2]] })).toBeUndefined()
    expect(bad((x) => { x.ranges[0].rows = [[0, 4], [3, 6]] })).toBeUndefined()
    expect(bad((x) => { x.ranges[0].rows = [[0, 4], [5, 6], [5, 6]] })).toBeUndefined()
    expect(bad((x) => { x.ranges[0].rows = [[5, 6], [0, 1]] })).toBeUndefined()
    expect(bad((x) => { x.ranges[0].rows = [[0, 1_048_576]] })).toBeUndefined()
    expect(bad((x) => { x.ranges[0].cols = [[0, 16_384]] })).toBeUndefined()
    expect(bad((x) => { x.active = [1_048_576, 0] })).toBeUndefined()
    expect(bad((x) => { x.anchor = [0, 16_384] })).toBeUndefined()
    expect(bad((x) => { x.kind = 'intervals' })).toBeUndefined()
    expect(bad((x) => { x.ranges[0].kind = 'intervals' })).toBeUndefined()
    expect(bad((x) => { x.ranges[0].rows = [[0.5, 1]] })).toBeUndefined()
    expect(bad((x) => { x.active = [1, 'a'] })).toBeUndefined()
    expect(bad((x) => { x.ranges[0].rows = [] })).toBeUndefined()
    expect(bad((x) => { delete x.anchor })).toBeUndefined()
    expect(bad((x) => { x.ranges = 'all' })).toBeUndefined()
    for (const v of [null, undefined, 3, 'x', [], [[0, 0]]]) expect(decodeWireSelection(v)).toBeUndefined()
  })
})

const REQ = { reqId: 3, type: 'sheet-find' }
describe('find args', () => {
  it('find args', () => {
    expect(findArgs({ ...REQ, query: 'x' }, 2)).toEqual({ query: 'x', matchCase: false, scope: 'workbook' })
    expect(findArgs({ ...REQ, query: 'x', matchCase: true, scope: '1' }, 2)).toEqual({ query: 'x', matchCase: true, scope: '1' })
    expect(findArgs({ ...REQ, query: 'x', wholeCell: false }, 2)).toBeDefined()
    expect(findArgs({ ...REQ, query: '' }, 2)).toBeUndefined()
    expect(findArgs({ ...REQ, query: 'a'.repeat(256) }, 2)).toBeDefined()
    expect(findArgs({ ...REQ, query: 'a'.repeat(257) }, 2)).toBeUndefined()
    expect(findArgs({ ...REQ, query: '😀'.repeat(129) }, 2)).toBeUndefined()
    expect(findArgs({ ...REQ, query: 'x', scope: '9' }, 2)).toBeUndefined()
    expect(findArgs({ ...REQ, query: 'x', scope: '2' }, 2)).toBeUndefined()
    expect(findArgs({ ...REQ, query: 'x', scope: '01' }, 2)).toBeUndefined()
    expect(findArgs({ ...REQ, query: 'x', scope: 1 }, 2)).toBeUndefined()
    expect(findArgs({ ...REQ, query: 'x', wholeCell: true }, 2)).toBeUndefined()
    expect(findArgs({ ...REQ, query: 'x', matchCase: 'yes' }, 2)).toBeUndefined()
    expect(findArgs({ ...REQ, query: 7 }, 2)).toBeUndefined()
    expect(findArgs({ ...REQ, query: 'x', cursor: 5 }, 2)).toBeDefined() // unknown keys are ignored, as the Stage 3 validators do
  })
})

describe('copy args', () => {
  it('accepts a well-formed copy and refuses hostile ones', () => {
    const ok = { reqId: 1, type: 'sheet-copy', sheetId: '1', viewGen: 4, selection: roundTrip(encodeWireSelection(sel3())) }
    expect(copyArgs(ok, 2)).toEqual({ sheetId: '1', sheetIndex: 1, viewGen: 4, selection: sel3() })
    expect(copyArgs({ ...ok, sheetId: '2' }, 2)).toBeUndefined()
    expect(copyArgs({ ...ok, sheetId: 1 }, 2)).toBeUndefined()
    expect(copyArgs({ ...ok, viewGen: -1 }, 2)).toBeUndefined()
    expect(copyArgs({ ...ok, viewGen: undefined }, 2)).toBeUndefined()
    expect(copyArgs({ ...ok, selection: { ...(ok.selection as object), kind: 'ref' } }, 2)).toBeUndefined()
    expect(copyArgs({ ...ok, selection: undefined }, 2)).toBeUndefined()
  })
})

describe('a11y args', () => {
  const base = { reqId: 1, type: 'sheet-a11y', sheetId: '0', viewGen: 0 }
  it('a11y args', () => {
    // 2,000 cells: 200 rows x 10 columns, and the same split over two row intervals.
    expect(a11yArgs({ ...base, rows: [[0, 199]], cols: [0, 9] }, 1)).toEqual({ sheetId: '0', sheetIndex: 0, viewGen: 0, rows: [[0, 199]], cols: [0, 9] })
    expect(a11yArgs({ ...base, rows: [[0, 99], [500, 599]], cols: [0, 9] }, 1)).toBeDefined()
    expect(a11yArgs({ ...base, rows: [[0, 200]], cols: [0, 9] }, 1)).toBeUndefined() // 2,010
    expect(a11yArgs({ ...base, rows: [[0, 99], [500, 600]], cols: [0, 9] }, 1)).toBeUndefined() // 2,010 over two intervals
    expect(a11yArgs({ ...base, rows: [[0, 1999]], cols: [3, 3] }, 1)).toBeDefined() // exactly 2,000
    expect(a11yArgs({ ...base, rows: [[0, 2000]], cols: [3, 3] }, 1)).toBeUndefined() // 2,001
    expect(a11yArgs({ ...base, rows: [[0, 1]], cols: [5, 2] }, 1)).toBeUndefined()
    expect(a11yArgs({ ...base, rows: [[0, 1]], cols: [0, 16_384] }, 1)).toBeUndefined()
    expect(a11yArgs({ ...base, rows: [[0, 1]], cols: [16_383, 16_383] }, 1)).toBeDefined()
    expect(a11yArgs({ ...base, rows: [[0, 1_048_576]], cols: [0, 0] }, 1)).toBeUndefined()
    expect(a11yArgs({ ...base, rows: [], cols: [0, 0] }, 1)).toBeUndefined()
    expect(a11yArgs({ ...base, rows: [[4, 6], [5, 8]], cols: [0, 0] }, 1)).toBeUndefined()
    expect(a11yArgs({ ...base, rows: [[4, 6]], cols: [0, 0, 1] }, 1)).toBeUndefined()
    expect(a11yArgs({ ...base, sheetId: '1', rows: [[0, 0]], cols: [0, 0] }, 1)).toBeUndefined()
  })
})

describe('fit and resize args', () => {
  const pane = { x: 0, y: 0, width: 640, height: 480 }
  const fit = { reqId: 1, type: 'sheet-fit', sheetId: '1', viewGen: 2, axis: 'col', index: 3, pane }
  it('fit and resize args', () => {
    expect(fitArgs(fit, 2)).toEqual({ sheetId: '1', sheetIndex: 1, viewGen: 2, axis: 'col', index: 3, pane })
    expect(fitArgs({ ...fit, axis: 'row', regions: [pane, pane, pane, pane] }, 2)).toMatchObject({ axis: 'row', regions: [pane, pane, pane, pane] })
    expect(fitArgs({ ...fit, regions: [pane, pane, pane, pane, pane] }, 2)).toBeUndefined()
    expect(fitArgs({ ...fit, pane: { ...pane, width: NaN } }, 2)).toBeUndefined()
    expect(fitArgs({ ...fit, pane: { ...pane, width: -1 } }, 2)).toBeUndefined()
    expect(fitArgs({ ...fit, pane: { ...pane, y: Infinity } }, 2)).toBeUndefined()
    expect(fitArgs({ ...fit, pane: { x: 0, y: 0, width: 1 } }, 2)).toBeUndefined()
    expect(fitArgs({ ...fit, regions: [{ ...pane, height: -2 }] }, 2)).toBeUndefined()
    expect(fitArgs({ ...fit, axis: 'z' }, 2)).toBeUndefined()
    expect(fitArgs({ ...fit, index: 1_048_576, axis: 'row' }, 2)).toBeUndefined()
    expect(fitArgs({ ...fit, index: 16_384 }, 2)).toBeUndefined()
    expect(fitArgs({ ...fit, sheetId: '5' }, 2)).toBeUndefined()

    const rs = { reqId: 1, type: 'sheet-resize', sheetId: '0', viewGen: 0, axis: 'row', index: 4, size: 30 }
    expect(resizeArgs(rs, 2)).toEqual({ sheetId: '0', sheetIndex: 0, viewGen: 0, axis: 'row', index: 4, size: 30 })
    expect(resizeArgs({ ...rs, size: 1e9 }, 2)).toMatchObject({ size: 1e9 }) // the server clamps
    expect(resizeArgs({ reqId: 1, type: 'sheet-resize', sheetId: '0', viewGen: 7, reset: true }, 2)).toEqual({ sheetId: '0', sheetIndex: 0, viewGen: 7, reset: true })
    expect(resizeArgs({ ...rs, reset: true }, 2)).toBeUndefined() // mixed forms
    expect(resizeArgs({ reqId: 1, type: 'sheet-resize', sheetId: '0', viewGen: 0, reset: true, axis: 'row' }, 2)).toBeUndefined()
    expect(resizeArgs({ reqId: 1, type: 'sheet-resize', sheetId: '0', viewGen: 0, reset: false }, 2)).toBeUndefined()
    expect(resizeArgs({ ...rs, axis: 'z' }, 2)).toBeUndefined()
    expect(resizeArgs({ ...rs, size: NaN }, 2)).toBeUndefined()
    expect(resizeArgs({ ...rs, size: -1 }, 2)).toBeUndefined()
    expect(resizeArgs({ ...rs, size: undefined }, 2)).toBeUndefined()
    expect(resizeArgs({ ...rs, index: 0.5 }, 2)).toBeUndefined()
    expect(resizeArgs({ ...rs, axis: 'col', index: 16_384 }, 2)).toBeUndefined()
    expect(resizeArgs({ ...rs, sheetId: '2' }, 2)).toBeUndefined()
  })

  it('SheetResizeRequest is SheetResizeArgs plus viewGen (F6)', () => {
    const a: SheetResizeArgs = { sheetId: '0', axis: 'col', index: 1, size: 10 }
    const b: SheetResizeArgs = { sheetId: '0', reset: true }
    const r1: SheetResizeRequest = { ...a, viewGen: 1 }, r2: SheetResizeRequest = { ...b, viewGen: 1 }
    // @ts-expect-error a resize needs a size
    const bad: SheetResizeArgs = { sheetId: '0', axis: 'col', index: 1 }
    expect([r1.viewGen, r2.viewGen, bad.sheetId]).toEqual([1, 1, '0'])
  })
})

describe('answer decoders', () => {
  it('hits decode', () => {
    expect(decodeFindHits({ hits: [[0, 4, 2], [1, 0, 0, 1]] })).toEqual([{ sheetIndex: 0, row: 4, col: 2 }, { sheetIndex: 1, row: 0, col: 0, hidden: 'filter' }])
    expect(decodeFindHits({ hits: [] })).toEqual([])
    for (const hits of [[[0, -1, 0]], [[0, 0, 16_384]], [[0, 1_048_576, 0]], [[0, 0]], [[0, 0, 0, 2]], [[0, 0, 0, 1, 1]], [[0.5, 0, 0]], [['0', 0, 0]], [[256, 0, 0]], 'x', [null]]) {
      expect(() => decodeFindHits({ hits })).toThrow(InvalidOperand)
    }
    expect(() => decodeFindHits({})).toThrow(InvalidOperand)
    expect(() => decodeFindHits({ hits: Array.from({ length: FIND_MAX_HITS + 1 }, () => [0, 0, 0]) })).toThrow(InvalidOperand)
  })

  it('decodes a11y cells and refuses hostile ones', () => {
    expect(decodeA11yCells({ cells: [[0, 1, 'a'], [5, 2, '']] })).toEqual([{ row: 0, col: 1, text: 'a' }, { row: 5, col: 2, text: '' }])
    for (const cells of [[[0, 1]], [[0, 1, 'a', 'b']], [[-1, 0, 'a']], [[0, 16_384, 'a']], [[0, 0, 7]], [[1_048_576, 0, 'a']], 'x']) {
      expect(() => decodeA11yCells({ cells })).toThrow(InvalidOperand)
    }
    expect(() => decodeA11yCells({ cells: Array.from({ length: A11Y_MAX_CELLS + 1 }, () => [0, 0, 'a']) })).toThrow(InvalidOperand)
    expect(() => decodeA11yCells({})).toThrow(InvalidOperand)
  })

  it('decodes a fit result and refuses hostile ones', () => {
    expect(decodeFitResult({ size: 123.5 })).toBe(123.5)
    expect(decodeFitResult({ size: 0 })).toBe(0)
    for (const size of [NaN, Infinity, -1, '5', null, undefined]) expect(() => decodeFitResult({ size })).toThrow(InvalidOperand)
    expect(() => decodeFitResult(7)).toThrow(InvalidOperand)
  })

  it('decodes a sheet view and refuses hostile ones', () => {
    const ok = { reqId: 1, type: 'sheet-view', sheetId: '2', viewGen: 5, change: { axis: 'col', index: 3, size: 91.5 } }
    expect(decodeSheetView(ok)).toEqual({ sheetId: '2', viewGen: 5, change: { axis: 'col', index: 3, size: 91.5 } })
    expect(decodeSheetView({ ...ok, change: { reset: true } })).toEqual({ sheetId: '2', viewGen: 5, change: { reset: true } })
    for (const bad of [
      { ...ok, sheetId: '02' }, { ...ok, viewGen: -1 }, { ...ok, change: undefined }, { ...ok, change: { axis: 'z', index: 0, size: 1 } },
      { ...ok, change: { axis: 'row', index: 1_048_576, size: 1 } }, { ...ok, change: { axis: 'col', index: 0, size: -1 } },
      { ...ok, change: { axis: 'col', index: 0, size: NaN } }, { ...ok, change: { reset: true, axis: 'row' } },
      { ...ok, change: { reset: false } }, { ...ok, change: { axis: 'col', index: 0 } }, { ...ok, extra: 1 }, null, 7,
    ]) expect(() => decodeSheetView(bad)).toThrow(InvalidOperand)
  })
})

describe('layout session sizes', () => {
  it('layout session sizes', () => {
    const cols = Array.from({ length: 5_000 }, (_, i): [number, number] => [i, 20.123456789 + (i % 50)])
    const big = layout({ sessionCols: cols, sessionRows: [[3, 41], [4, 18.25]] })
    const parts = wire(big)
    expect(parts.length).toBeGreaterThan(1)
    for (const p of parts) expect(new TextEncoder().encode(JSON.stringify(p)).length).toBeLessThanOrEqual(65_536)
    const back = decodeSheetLayout(parts)
    expect(back.sessionCols).toEqual(cols)
    expect(back.sessionRows).toEqual([[3, 41], [4, 18.25]])
    expect(back).toEqual(big)
  })

  it('a 4a-shaped layout decodes with empty session sizes, and an identity view encodes with neither key', () => {
    const parts = wire(layout())
    for (const p of parts) { expect(p).not.toHaveProperty('sessionRows'); expect(p).not.toHaveProperty('sessionCols') }
    const empty = wire(layout({ sessionRows: [], sessionCols: [] }))
    expect(empty).toEqual(parts)
    const back = decodeSheetLayout(parts)
    expect(back.sessionRows ?? []).toEqual([])
    expect(back.sessionCols ?? []).toEqual([])
    expect('sessionRows' in back).toBe(false)
    expect('sessionCols' in back).toBe(false)
  })

  it('refuses hostile session sizes', () => {
    const base = () => wire(layout({ sessionRows: [[1, 30]] }))
    const withP0 = (mut: (p: Record<string, any>) => void) => { const parts = base(); mut(parts[0]!); return parts }
    expect(() => decodeSheetLayout(withP0((p) => { p.sessionRows = [[-1, 30]] }))).toThrow(LayoutInvalid)
    expect(() => decodeSheetLayout(withP0((p) => { p.sessionRows = [[1_048_576, 30]] }))).toThrow(LayoutInvalid)
    expect(() => decodeSheetLayout(withP0((p) => { p.sessionCols = [[16_384, 30]] }))).toThrow(LayoutInvalid)
    expect(() => decodeSheetLayout(withP0((p) => { p.sessionRows = [[1, -30]] }))).toThrow(LayoutInvalid)
    expect(() => decodeSheetLayout(withP0((p) => { p.sessionRows = [[1, NaN]] }))).toThrow(LayoutInvalid)
    expect(() => decodeSheetLayout(withP0((p) => { p.sessionRows = [[1, 30, 9]] }))).toThrow(LayoutInvalid)
    expect(() => decodeSheetLayout(withP0((p) => { p.sessionRows = [{ index: 1, size: 3 }] }))).toThrow(LayoutInvalid)
    expect(() => decodeSheetLayout(withP0((p) => { p.sessionRows = 'x' }))).toThrow(LayoutInvalid)
    expect(() => decodeSheetLayout(withP0((p) => { p.bogus = [] }))).toThrow(LayoutInvalid)
    expect(() => encodeSheetLayout(layout({ sessionCols: Array.from({ length: LAYOUT_LIMITS.intervals + 1 }, (_, i): [number, number] => [i % 16_384, 20]) }))).toThrow(LayoutTooLarge)
  })
})

// ---- Stage 4c: sort, filter, value-list, range-detect and order codecs ----
import {
  decodeDetection, decodeSheetOrder, decodeValuesPage, detectArgs, encodeSheetOrder, filterArgs, FILTER_TEXT_UNITS, ORDER_LIMITS,
  orderArgs, SORT_FILTER_MAX_RULES, sortArgs, VALUE_LIST_LIMITS, valuesArgs, type SheetOrder, type WireRange,
} from '../../src/stream/sheet-wire'

const wr = (o: Partial<Record<keyof WireRange, number>> = {}): WireRange => ({ r0: 0, r1: 9, c0: 0, c1: 4, header: 1, totals: 0, ...o } as WireRange)
const threeKeys = (): Array<{ col: number; order: 'asc' | 'desc' }> => [{ col: 0, order: 'asc' }, { col: 2, order: 'desc' }, { col: 4, order: 'asc' }]
const sortMsg = (o: Record<string, unknown> = {}): Record<string, unknown> => ({ reqId: 3, type: 'sheet-sort', sheetId: '0', viewGen: 2, range: wr(), keys: threeKeys(), caseSensitive: true, ...o })
const filterMsg = (columns: unknown[], o: Record<string, unknown> = {}): Record<string, unknown> => ({ reqId: 4, type: 'sheet-filter', sheetId: '0', viewGen: 2, range: wr(), columns, ...o })
const text = { kind: 'text', op: 'contains', first: 'ab', wildcards: false }
const num = { kind: 'number', op: 'between', first: 1, second: 9.5 }
const values = (o: Record<string, unknown> = {}) => ({ kind: 'values', list: 'l-1', mode: 'include', ids: [[0, 3], [5, 5]], blanks: true, ...o })

describe('sort and filter args', () => {
  it('accepts a 3-key sort and a 4-column filter', () => {
    expect(sortArgs(sortMsg(), 1)).toEqual({ sheetId: '0', sheetIndex: 0, viewGen: 2, range: wr(), keys: threeKeys(), caseSensitive: true })
    expect(sortArgs(sortMsg({ caseSensitive: undefined }), 1)?.caseSensitive).toBe(false)
    const columns: Array<[number, unknown]> = [
      [0, text], [1, num], [2, { kind: 'compound', connective: 'or', first: text, second: { kind: 'number', op: 'gt', first: 3 } }], [4, values()],
    ]
    expect(filterArgs(filterMsg(columns), 1)).toEqual({ sheetId: '0', sheetIndex: 0, viewGen: 2, range: wr(), columns })
    const rest: Array<[number, unknown]> = [
      [0, { kind: 'date', op: 'between', first: { year: 2024, month: 2, day: 29 }, second: { year: 2025, month: 1, day: 1 } }],
      [1, { kind: 'datePreset', preset: 'thisWeek' }], [2, { kind: 'committed' }], [3, values({ mode: 'exclude', ids: [], blanks: false })],
    ]
    expect(filterArgs(filterMsg(rest), 1)?.columns).toEqual(rest)
    expect(filterArgs(filterMsg([]), 1)?.columns).toEqual([])
  })

  it('refuses every malformed sort', () => {
    expect(SORT_FILTER_MAX_RULES).toBe(64)
    const key = (col: number) => ({ col, order: 'asc' })
    const cases: Array<[string, Record<string, unknown>]> = [
      ['65 keys', sortMsg({ range: wr({ c1: 100 }), keys: Array.from({ length: 65 }, (_, i) => key(i)) })],
      ['no keys', sortMsg({ keys: [] })],
      ['key column outside the range', sortMsg({ keys: [key(5)] })],
      ['key column left of the range', sortMsg({ range: wr({ c0: 1 }), keys: [key(0)] })],
      ['header 2', sortMsg({ range: wr({ header: 2 }) })],
      ['totals 2', sortMsg({ range: wr({ totals: 2 }) })],
      ['body of 0 rows', sortMsg({ range: wr({ r0: 3, r1: 3, header: 1 }) })],
      ['body of 0 rows under header and totals', sortMsg({ range: wr({ r0: 3, r1: 4, header: 1, totals: 1 }) })],
      ['range past the grid', sortMsg({ range: wr({ r1: 1_048_576 }) })],
      ['columns past the grid', sortMsg({ range: wr({ c1: 16_384 }) })],
      ['inverted range', sortMsg({ range: wr({ r0: 5, r1: 2 }) })],
      ['unknown order', sortMsg({ keys: [{ col: 0, order: 'up' }] })],
      ['extra key field', sortMsg({ keys: [{ col: 0, order: 'asc', x: 1 }] })],
      ['non-boolean caseSensitive', sortMsg({ caseSensitive: 1 })],
      ['sheet out of range', sortMsg({ sheetId: '1' })],
      ['bad viewGen', sortMsg({ viewGen: -1 })],
      ['extra range field', sortMsg({ range: { ...wr(), x: 1 } })],
    ]
    for (const [name, m] of cases) expect(sortArgs(m, 1), name).toBeUndefined()
  })

  it('refuses every malformed filter', () => {
    const col = (c: number, rule: unknown): [number, unknown] => [c, rule]
    const many = Array.from({ length: 65 }, (_, i) => col(i, { kind: 'committed' }))
    const wide = wr({ c1: 100 })
    const cases: Array<[string, Record<string, unknown>]> = [
      ['65 columns', filterMsg(many, { range: wide })],
      ['column outside the range', filterMsg([col(5, text)])],
      ['header 2', filterMsg([col(0, text)], { range: wr({ header: 2 }) })],
      ['body of 0 rows', filterMsg([col(0, text)], { range: wr({ r0: 2, r1: 2 }) })],
      ['257-unit first', filterMsg([col(0, { ...text, first: 'a'.repeat(FILTER_TEXT_UNITS + 1) })])],
      ['ids interval i1 < i0', filterMsg([col(0, values({ ids: [[4, 3]] }))])],
      ['overlapping intervals', filterMsg([col(0, values({ ids: [[0, 3], [3, 5]] }))])],
      ['descending intervals', filterMsg([col(0, values({ ids: [[5, 6], [0, 1]] }))])],
      ['4,097 intervals', filterMsg([col(0, values({ ids: Array.from({ length: VALUE_LIST_LIMITS.intervals + 1 }, (_, i) => [i * 2, i * 2]) }))])],
      ['duplicate columns', filterMsg([col(1, text), col(1, num)])],
      ['descending columns', filterMsg([col(2, text), col(1, num)])],
      ['compound with a values child', filterMsg([col(0, { kind: 'compound', connective: 'and', first: text, second: values() })])],
      ['compound with a committed child', filterMsg([col(0, { kind: 'compound', connective: 'and', first: text, second: { kind: 'committed' } })])],
      ['nested compound', filterMsg([col(0, { kind: 'compound', connective: 'and', first: text, second: { kind: 'compound', connective: 'or', first: text, second: text } })])],
      ['unknown connective', filterMsg([col(0, { kind: 'compound', connective: 'xor', first: text, second: text })])],
      ['unknown kind', filterMsg([col(0, { kind: 'regex', first: 'a' })])],
      ['NaN first', filterMsg([col(0, { kind: 'number', op: 'gt', first: NaN })])],
      ['Infinity first', filterMsg([col(0, { kind: 'number', op: 'gt', first: Infinity })])],
      ['between without second', filterMsg([col(0, { kind: 'number', op: 'between', first: 1 })])],
      ['second on a single-operand op', filterMsg([col(0, { kind: 'number', op: 'gt', first: 1, second: 2 })])],
      ['unknown text op', filterMsg([col(0, { ...text, op: 'regex' })])],
      ['text wildcards not boolean', filterMsg([col(0, { ...text, wildcards: 'y' })])],
      ['impossible date', filterMsg([col(0, { kind: 'date', op: 'eq', first: { year: 2023, month: 2, day: 29 } })])],
      ['unknown preset', filterMsg([col(0, { kind: 'datePreset', preset: 'lastYear' })])],
      ['values without a list', filterMsg([col(0, values({ list: '' }))])],
      ['values mode', filterMsg([col(0, values({ mode: 'all' }))])],
      ['committed with a field', filterMsg([col(0, { kind: 'committed', x: 1 })])],
      ['column not a pair', filterMsg([{ col: 0, rule: text }])],
      ['columns not a list', filterMsg([], { columns: {} })],
    ]
    for (const [name, m] of cases) expect(filterArgs(m, 1), name).toBeUndefined()
  })
})

describe('values args', () => {
  it('accepts a first page and a cursor page', () => {
    const first = { reqId: 5, type: 'filter-values', sheetId: '0', viewGen: 1, range: wr(), col: 2 }
    expect(valuesArgs(first, 1)).toEqual({ sheetId: '0', sheetIndex: 0, viewGen: 1, range: wr(), col: 2 })
    expect(valuesArgs({ ...first, search: { text: 'a*', wildcards: true } }, 1)).toEqual({ sheetId: '0', sheetIndex: 0, viewGen: 1, range: wr(), col: 2, search: { text: 'a*', wildcards: true } })
    expect(valuesArgs({ reqId: 6, type: 'filter-values', sheetId: '0', viewGen: 1, cursor: 'c-1' }, 1)).toEqual({ sheetId: '0', sheetIndex: 0, viewGen: 1, cursor: 'c-1' })
  })

  it('refuses both forms at once, a long search and a bad column', () => {
    const first = { sheetId: '0', viewGen: 1, range: wr(), col: 2 }
    expect(valuesArgs({ ...first, cursor: 'c-1' }, 1)).toBeUndefined()
    expect(valuesArgs({ sheetId: '0', viewGen: 1, cursor: 'c-1', col: 2 }, 1)).toBeUndefined()
    expect(valuesArgs({ sheetId: '0', viewGen: 1, cursor: 'c-1', search: { text: 'a', wildcards: false } }, 1)).toBeUndefined()
    expect(valuesArgs({ ...first, search: { text: 'a'.repeat(FILTER_TEXT_UNITS + 1), wildcards: false } }, 1)).toBeUndefined()
    expect(valuesArgs({ ...first, search: { text: 'a' } }, 1)).toBeUndefined()
    expect(valuesArgs({ ...first, col: 5 }, 1)).toBeUndefined()
    expect(valuesArgs({ sheetId: '0', viewGen: 1, range: wr() }, 1)).toBeUndefined()
    expect(valuesArgs({ sheetId: '0', viewGen: 1 }, 1)).toBeUndefined()
    expect(valuesArgs({ sheetId: '0', viewGen: 1, cursor: '' }, 1)).toBeUndefined()
    expect(valuesArgs({ sheetId: '9', viewGen: 1, cursor: 'c-1' }, 1)).toBeUndefined()
    expect(VALUE_LIST_LIMITS).toEqual({ entries: 10_000, page: 500, intervals: 4_096 })
  })
})

describe('range-detect and order args', () => {
  const sel = { active: [1, 1], anchor: [1, 1], ranges: [{ rows: [[1, 1]], cols: [[1, 1]] }] }
  it('accepts and refuses range-detect', () => {
    const m = { sheetId: '0', viewGen: 1, kind: 'sort', selection: sel }
    expect(detectArgs(m, 1)).toEqual({ sheetId: '0', sheetIndex: 0, viewGen: 1, kind: 'sort', selection: sel })
    expect(detectArgs({ ...m, kind: 'filter', override: { header: 0 } }, 1)).toEqual({ sheetId: '0', sheetIndex: 0, viewGen: 1, kind: 'filter', selection: sel, override: { header: 0 } })
    for (const bad of [{ kind: 'pivot' }, { override: { header: 2 } }, { override: { header: 1, x: 1 } }, { override: 1 }, { selection: { ...sel, kind: 'x' } }, { selection: undefined }, { viewGen: 'a' }, { sheetId: '4' }]) {
      expect(detectArgs({ ...m, ...bad }, 1), JSON.stringify(bad)).toBeUndefined()
    }
  })
  it('accepts and refuses sheet-order', () => {
    expect(orderArgs({ sheetId: '1' }, 2)).toEqual({ sheetId: '1', sheetIndex: 1 })
    expect(orderArgs({ sheetId: '2' }, 2)).toBeUndefined()
    expect(orderArgs({ sheetId: 1 }, 2)).toBeUndefined()
  })
})

const orderHead = (o: Partial<SheetOrder['head']> = {}): SheetOrder['head'] => ({
  sheetId: '0', viewGen: 3, counts: { visible: 5, total: 8 }, sort: { range: wr({ r0: 0, r1: 7, header: 0 }), keys: [{ col: 1, order: 'desc' }] },
  filterSummary: [{ col: 2, kind: 'values' }], ...o,
})
const orderWire = (o: SheetOrder) => encodeSheetOrder(o).map((p) => JSON.parse(JSON.stringify({ reqId: 4_294_967_295, ...p })) as Record<string, any>)
const b64 = (words: number[]): string => {
  const bytes = new Uint8Array(words.length * 4), v = new DataView(bytes.buffer)
  words.forEach((w, i) => v.setUint32(i * 4, w, true))
  return btoa(String.fromCharCode(...bytes))
}
const smallOrder = (): SheetOrder => ({
  head: orderHead(), bands: [{ r0: 0, r1: 7, sourceAtOffset: Uint32Array.from([7, 6, 5, 4, 3, 2, 1, 0]) }], mask: { r0: 0, r1: 7, words: Uint32Array.of(0b10110101) },
})
const orderBad = (parts: unknown[]) => expect(() => decodeSheetOrder(parts)).toThrow(InvalidOperand)

describe('order parts', () => {
  it('pins the caps', () => expect(ORDER_LIMITS).toEqual({ parts: 160, bands: 1, rows: 1_048_576 }))

  it('round-trips a small order and an order with no band', () => {
    const o = smallOrder()
    const parts = orderWire(o)
    expect(parts[0]!.head).toBeDefined()
    expect(parts.at(-1)!.last).toBe(true)
    expect(decodeSheetOrder(parts)).toEqual(o)
    const bare: SheetOrder = { head: orderHead({ sort: null, filterSummary: [] }), bands: [], mask: null }
    expect(decodeSheetOrder(orderWire(bare))).toEqual(bare)
  })

  it('encodes a 1,048,576-row reverse band in at most 160 parts of at most 64 KiB, and decodes equal', () => {
    const n = 1_048_576
    const sourceAtOffset = new Uint32Array(n).map((_, i) => n - 1 - i)
    const words = new Uint32Array(Math.ceil(n / 32)).map((_, i) => (i * 2_654_435_761) >>> 0)
    const o: SheetOrder = { head: orderHead({ counts: { visible: n - 7, total: n }, sort: { range: wr({ r0: 0, r1: n - 1, header: 0 }), keys: [{ col: 0, order: 'desc' }] } }), bands: [{ r0: 0, r1: n - 1, sourceAtOffset }], mask: { r0: 0, r1: n - 1, words } }
    const parts = orderWire(o)
    expect(parts.length).toBeLessThanOrEqual(ORDER_LIMITS.parts)
    for (const p of parts) expect(new TextEncoder().encode(JSON.stringify(p)).length).toBeLessThanOrEqual(65_536)
    const back = decodeSheetOrder(parts)
    expect(back.head).toEqual(o.head)
    expect(back.bands).toHaveLength(1)
    expect(back.bands[0]!.sourceAtOffset).toEqual(sourceAtOffset)
    expect(back.mask!.words).toEqual(words)
  }, 30_000)

  it('throws for a dropped part, a duplicate entry, a gap, an overlap, a reorder and 161 parts', () => {
    const big = (): SheetOrder => {
      const n = 40_000
      return { head: orderHead({ counts: { visible: n, total: n }, sort: null }), bands: [{ r0: 0, r1: n - 1, sourceAtOffset: new Uint32Array(n).map((_, i) => n - 1 - i) }], mask: null }
    }
    const parts = orderWire(big())
    expect(parts.length).toBeGreaterThanOrEqual(4)
    expect(decodeSheetOrder(parts).bands[0]!.sourceAtOffset).toHaveLength(40_000)
    orderBad(parts.filter((_, i) => i !== 2))
    orderBad(parts.slice(0, -1))
    const swapped = [...parts]; [swapped[1], swapped[2]] = [swapped[2]!, swapped[1]!]
    orderBad(swapped)
    const gap = orderWire(big()); gap[2]!.band.from += 1; orderBad(gap)
    const overlap = orderWire(big()); overlap[2]!.band.from -= 1; orderBad(overlap)
    const lastLate = orderWire(big()); lastLate[1]!.last = true; orderBad(lastLate)
    const headTwice = orderWire(big()); headTwice[1]!.head = headTwice[0]!.head; orderBad(headTwice)
    const reqIdDrift = orderWire(big()); reqIdDrift[1]!.reqId = 9; orderBad(reqIdDrift)
    const dup = orderWire(smallOrder()); dup[1]!.band.data = b64([0, 1, 1, 3, 4, 5, 6, 7]); orderBad(dup)
    const outside = orderWire(smallOrder()); outside[1]!.band.data = b64([0, 1, 2, 3, 4, 5, 6, 8]); orderBad(outside)
    const short = orderWire(smallOrder()); short[1]!.band.data = b64([7, 6, 5, 4]); orderBad(short)
    const junk = orderWire(smallOrder()); junk[1]!.band.data = '!!!!'; orderBad(junk)
    const extra = orderWire(smallOrder()); extra[0]!.bogus = 1; orderBad(extra)
    orderBad(Array.from({ length: ORDER_LIMITS.parts + 1 }, (_, i) => ({ type: 'sheet-order', seq: i, last: i === ORDER_LIMITS.parts, ...(i === 0 ? { head: orderHead() } : {}) })))
    orderBad([])
  })

  it('refuses an over-cap order on encode and a bad head on decode', () => {
    const band = (r0: number, r1: number) => ({ r0, r1, sourceAtOffset: Uint32Array.from({ length: r1 - r0 + 1 }, (_, i) => r0 + i) })
    expect(() => encodeSheetOrder({ head: orderHead(), bands: [band(0, 1), band(2, 3)], mask: null })).toThrow(InvalidOperand)
    expect(() => encodeSheetOrder({ head: orderHead(), bands: [{ r0: 0, r1: 3, sourceAtOffset: Uint32Array.of(0, 1) }], mask: null })).toThrow(InvalidOperand)
    const heads: Array<(h: Record<string, any>) => void> = [
      (h) => { h.counts.visible = 99 }, (h) => { h.sheetId = '01' }, (h) => { h.sort.keys[0].col = 99 }, (h) => { h.filterSummary = [{ col: 0 }] }, (h) => { h.viewGen = -1 },
    ]
    for (const edit of heads) { const parts = orderWire(smallOrder()); edit(parts[0]!.head); orderBad(parts) }
  })
})

describe('values page', () => {
  const page = (o: Record<string, unknown> = {}) => ({
    reqId: 8, type: 'filter-values-result', list: 'l-1', items: [[0, 'a', 1], [1, 'b', 0], [2, '#N/A', 3]], blanks: { present: true, selected: false }, total: 3, truncated: false,
    cursor: 'c-2', classification: { kind: 'text', hasNumeric: true }, ...o,
  })
  it('decodes flags into selected and error', () => {
    const p = decodeValuesPage(page())
    expect(p.items).toEqual([{ id: 0, text: 'a', selected: true, error: false }, { id: 1, text: 'b', selected: false, error: false }, { id: 2, text: '#N/A', selected: true, error: true }])
    expect(p).toMatchObject({ list: 'l-1', total: 3, truncated: false, cursor: 'c-2', blanks: { present: true, selected: false }, classification: { kind: 'text', hasNumeric: true } })
    const bare = decodeValuesPage(page({ cursor: undefined, classification: undefined }))
    expect('cursor' in bare).toBe(false)
    expect('classification' in bare).toBe(false)
  })
  it('throws for an id outside [0, total) and for other off-contract pages', () => {
    const bad = (o: Record<string, unknown>) => expect(() => decodeValuesPage(page(o))).toThrow(InvalidOperand)
    bad({ items: [[3, 'a', 0]] })
    bad({ items: [[-1, 'a', 0]] })
    bad({ items: [[0, 'a', 4]] })
    bad({ items: [[0, 'a']] })
    bad({ items: [[0, 'a'.repeat(FILTER_TEXT_UNITS + 1), 0]] })
    bad({ items: Array.from({ length: VALUE_LIST_LIMITS.page + 1 }, (_, i) => [i, 'x', 0]), total: 600 })
    bad({ total: VALUE_LIST_LIMITS.entries + 1 })
    bad({ list: '' })
    bad({ truncated: 1 })
    bad({ blanks: { present: true } })
    bad({ classification: { kind: 'bool', hasNumeric: false } })
    bad({ bogus: 1 })
    expect(() => decodeValuesPage(null)).toThrow(InvalidOperand)
  })
})

describe('range detection answer', () => {
  const ans = (o: Record<string, unknown> = {}) => ({ reqId: 9, type: 'range-detect-result', range: wr(), source: 'region', header: 'heuristic', overrideIgnored: false, ...o })
  it('decodes and refuses', () => {
    expect(decodeDetection(ans())).toEqual({ range: wr(), source: 'region', header: 'heuristic', overrideIgnored: false })
    for (const bad of [{ source: 'guess' }, { header: 'auto' }, { overrideIgnored: 0 }, { range: wr({ header: 2 }) }, { bogus: 1 }]) {
      expect(() => decodeDetection(ans(bad)), JSON.stringify(bad)).toThrow(InvalidOperand)
    }
    expect(() => decodeDetection(undefined)).toThrow(InvalidOperand)
  })
})

describe('sheet-layout carries the sort (F1)', () => {
  const sorted = () => layout({ sort: { range: wr({ r0: 0, r1: 8 }), keys: [{ col: 1, order: 'desc' }, { col: 0, order: 'asc' }] }, filterSummary: [{ col: 2, kind: 'values' }] })
  it('round-trips a sort, and a null sort stays null', () => {
    const l = sorted()
    expect(decodeSheetLayout(wire(l))).toEqual(l)
    expect(wire(layout())[0]!.head.sort).toBeNull()
    expect(decodeSheetLayout(wire(layout())).sort).toBeNull()
  })
  it('refuses a malformed sort', () => {
    const edits: Array<(s: Record<string, any>) => void> = [
      (s) => { s.keys[0].col = 99 }, (s) => { s.keys = [] }, (s) => { s.range.header = 2 }, (s) => { s.x = 1 }, (s) => { s.keys[0].order = 'up' },
      (s) => { s.keys = Array.from({ length: 65 }, () => ({ col: 0, order: 'asc' })) },
    ]
    for (const edit of edits) {
      const parts = wire(sorted())
      edit(parts[0]!.head.sort)
      rejects(parts)
    }
  })
})
