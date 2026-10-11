/** S5.2 review round 2: coarse selection overlay membership and output bounds (spec §§8.1, 11.1). */
import { describe, expect, it } from 'vitest'
import { coarseSelectionRecords, COARSE_SELECTION_RECORDS } from '../src/components/xlsx/selectionOverlay'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createRowOrder } from '../src/xlsx/view/order'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import type { GridSelection } from '../src/xlsx/view/contracts'

const metrics = (order?: ReturnType<typeof createRowOrder>) => createSheetMetrics({
  rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 1 }),
  cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 1 }),
  ...(order ? { order } : {}),
})
const reversed = (n: number) => createRowOrder([{ r0: 0, r1: n - 1, sourceAtOffset: Uint32Array.from({ length: n }, (_, i) => n - 1 - i) }])
const selection = (rows: GridSelection['ranges'][number]['rows'], c1 = 3): GridSelection => ({ active: null, anchor: null, ranges: [{ rows, cols: [{ c0: 0, c1 }] }] })

describe('coarse selection overlay', () => {
  it('never paints unselected display rows of a permuted band as selected', () => {
    const m = metrics(reversed(100_000))
    // Source rows 0..65,536 sit at display slots 34,463..99,999: none in the pane.
    const result = coarseSelectionRecords(selection({ kind: 'intervals', intervals: [{ r0: 0, r1: 65_536 }] }), m, { x: 0, y: 0, width: 400, height: 2000 }, 1)
    expect(result.records).toHaveLength(0)
  })

  it('is exact for visible selected slots inside a band and reports when the work budget truncates', () => {
    const m = metrics(reversed(100_000))
    // Source 99,990..99,999 = display 0..9.
    const exact = coarseSelectionRecords(selection({ kind: 'intervals', intervals: [{ r0: 99_990, r1: 99_999 }] }), m, { x: 0, y: 0, width: 400, height: 50 }, 1)
    expect(exact.records.map(r => [r.bounds.y, r.bounds.height])).toEqual([[0, 10]])
    expect(exact.incomplete).toBe(false)
    const huge = coarseSelectionRecords(selection({ kind: 'intervals', intervals: [{ r0: 0, r1: 99_999 }] }), metrics(reversed(100_000)), { x: 0, y: 0, width: 400, height: 100_000 }, 1)
    expect(huge.incomplete).toBe(true)
    // Truncated work skips slots; it never paints a slot it did not prove selected.
    for (const record of huge.records) expect(record.bounds.y + record.bounds.height).toBeLessThanOrEqual(100_000)
  })

  it('bounds the total output across both axes and every part', () => {
    const m = metrics()
    const rows = Uint32Array.from({ length: 65_536 }, (_, i) => i * 2)
    const cols = Array.from({ length: 256 }, (_, i) => ({ c0: i * 2, c1: i * 2 }))
    const fragmented: GridSelection = { active: null, anchor: null, ranges: [{ rows: { kind: 'list', r0: 0, r1: rows[rows.length - 1], rows }, cols }] }
    const result = coarseSelectionRecords(fragmented, m, { x: 0, y: 0, width: 512, height: 131_072 }, 1)
    expect(result.records.length).toBeLessThanOrEqual(COARSE_SELECTION_RECORDS)
    expect(result.records.length).toBeGreaterThan(0)
  })

  it('identity rows map exactly', () => {
    const result = coarseSelectionRecords(selection({ kind: 'intervals', intervals: [{ r0: 10, r1: 20 }] }), metrics(), { x: 0, y: 0, width: 100, height: 100 }, 1)
    expect(result.records.map(r => [r.bounds.y, r.bounds.height, r.bounds.x, r.bounds.width])).toEqual([[10, 11, 0, 4]])
  })
})
