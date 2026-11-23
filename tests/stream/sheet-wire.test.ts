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
