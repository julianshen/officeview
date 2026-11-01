// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { computeAxes } from '../src/xlsx/pivots/engine/axes'
import { computePivot } from '../src/xlsx/pivots/engine/compute'
import { retainedLayoutBytes } from '../src/xlsx/pivots/session/session'
import type { PivotCacheDefinition, PivotTableDefinition } from '../src/xlsx/pivots/model'
import type { PivotViewState } from '../src/xlsx/pivots/engine/view'
import type { PivotLayout } from '../src/xlsx/pivots/engine/place'
import type { RecordStore } from '../src/xlsx/pivots/engine/records'

function showAllPivot(itemCount: number) {
  const cache = {
    source: { type: 'worksheet' }, olap: false,
    fields: [
      { name: 'Item', databaseField: true, sharedItems: { inline: false, items: Array.from({ length: itemCount }, (_, i) => ({ type: 'string', value: `I${i}` })) } },
      { name: 'Amount', databaseField: true, sharedItems: { inline: true, items: [] } },
    ],
  } as unknown as PivotCacheDefinition
  const def = {
    unsupported: [], location: { ref: 'A1' }, fields: [
      { items: Array.from({ length: itemCount }, (_, x) => ({ type: 'data', x })), defaultSubtotal: false },
      { items: [] },
    ], rowFields: [0], colFields: [], pageFields: [], dataFields: [], rowGrandTotals: true, colGrandTotals: true,
    compactData: true, outlineData: true,
  } as unknown as PivotTableDefinition
  const view = {
    rows: [0], cols: [], pages: [], filters: [], data: [{ field: 1, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832 }],
    fields: [0, 1].map((_, i) => ({ hidden: i === 0 ? Array(itemCount).fill(false) : [], collapsed: i === 0 ? Array(itemCount).fill(false) : [], sortType: 'manual',
      defaultSubtotal: false, subtotals: [], subtotalTop: false, compact: false, outline: false, insertBlankRow: false, fillDownLabels: false, showAll: i === 0 })),
  } as unknown as PivotViewState
  const store = { count: 1, bytes: 4, columns: [
    { kind: 'shared', index: Uint32Array.of(0) },
    { kind: 'inline', tag: Uint8Array.of(1), num: Float64Array.of(3), text: [], textIndex: Uint32Array.of(0) },
  ] } as unknown as RecordStore
  return { cache, def, view, store }
}

describe('Stage 4 pivot candidate memory admission', () => {
  it('reserves showAll axis output before computeAxes materializes its lines', () => {
    const input = showAllPivot(4096)
    const axes = computeAxes(input.store, input.cache, input.def, input.view)
    expect(axes.ok).toBe(true)
    if (!axes.ok) return
    const items = axes.value.rows.reduce((sum, line) => sum + line.items.length, 0)
    const minimumForOutput = items * 8 + axes.value.rows.length * 32
    let firstReservation = 0
    const result = computePivot(input.store, input.cache, input.def, input.view, {
      date1904: false,
      admitLayoutBytes: bytes => { firstReservation = bytes; return { bytes, release() {} } },
    })
    expect(result.ok).toBe(true)
    expect(firstReservation).toBeGreaterThanOrEqual(minimumForOutput)
  })

  it('retained layout estimate includes every row and column path item slot', () => {
    const slots = 16_384
    const layout = {
      rect: { minRow: 0, maxRow: 1, minCol: 0, maxCol: 0 },
      axes: { rows: [{ type: 'data', first: 0, items: Array.from({ length: slots }, (_, i) => i), data: 0 }], cols: [] },
      cells: [], merges: [], captionsUnverified: false,
    } as unknown as PivotLayout
    const bytes = retainedLayoutBytes({ count: 1 } as RecordStore, layout)
    expect(bytes).toBeGreaterThanOrEqual(slots * 8)
  })
})
