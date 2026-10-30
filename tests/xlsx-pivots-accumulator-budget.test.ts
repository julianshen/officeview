// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { computeAxes } from '../src/xlsx/pivots/engine/axes'
import { computeValues } from '../src/xlsx/pivots/engine/values'
import { VALUE_TAG, type RecordStore } from '../src/xlsx/pivots/engine/records'
import type { PivotCacheDefinition, PivotTableDefinition } from '../src/xlsx/pivots/model'
import type { FieldState, PivotViewState } from '../src/xlsx/pivots/engine/view'

const makeFieldState = (count: number, defaultSubtotal: boolean): FieldState => ({
  hidden: new Array(count).fill(false), collapsed: new Array(count).fill(false), sortType: 'manual',
  defaultSubtotal, subtotals: [], subtotalTop: false, compact: false, outline: true,
  insertBlankRow: false, fillDownLabels: false, showAll: false,
})

function seededValues(count: number, seed: number): { values: Float64Array; sum: number } {
  const values = new Float64Array(count)
  let state = seed >>> 0, sum = 0
  for (let i = 0; i < count; i++) {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    const value = (state % 10_000) / 100
    values[i] = value
    sum += value
  }
  return { values, sum }
}

describe('pivot accumulator budget', () => {
  it('computes the S2-P line/cardinality shape with a bounded sum accumulator', () => {
    // 100,000 records cover all 10 × 50 × 200 row tuples once; the 12 column
    // items are cycled across those tuples. This keeps test time reasonable
    // while preserving the benchmark's axes and grid cardinalities.
    const count = 100_000
    const cardinalities = [10, 50, 200, 12]
    const sharedItems = cardinalities.map(size => ({
      inline: false,
      items: Array.from({ length: size }, (_, i) => ({ type: 'string', value: String(i) })),
    }))
    const cache = {
      source: { type: 'worksheet' }, olap: false,
      fields: [...sharedItems.map(() => ({ databaseField: true })), { databaseField: true }, { databaseField: true }]
        .map((field, index) => ({ ...field, ...(index < 4 ? { sharedItems: sharedItems[index] } : { sharedItems: { inline: true, items: [] } }) })),
    } as unknown as PivotCacheDefinition
    const def = {
      fields: [...cardinalities.map(size => ({
        items: Array.from({ length: size }, (_, x) => ({ type: 'data', x, showDetails: true })),
        defaultSubtotal: true,
      })), { items: [] }, { items: [] }],
      rowGrandTotals: true, colGrandTotals: true,
    } as unknown as PivotTableDefinition
    const values1 = seededValues(count, 0x5eed1234), values2 = seededValues(count, 0xc0ffee)
    const fields = cardinalities.map((size, field) => makeFieldState(size, field < 3))
    fields.push(makeFieldState(0, false), makeFieldState(0, false))
    const view = {
      rows: [0, 1, 2], cols: [3], pages: [], filters: [], fields,
      data: [4, 5].map(field => ({ field, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: `Sum ${field}` })),
    } as unknown as PivotViewState
    const store: RecordStore = {
      count, bytes: 0,
      columns: [
        { kind: 'shared', index: Uint32Array.from({ length: count }, (_, i) => i % 10) },
        { kind: 'shared', index: Uint32Array.from({ length: count }, (_, i) => Math.floor(i / 10) % 50) },
        { kind: 'shared', index: Uint32Array.from({ length: count }, (_, i) => Math.floor(i / 500) % 200) },
        { kind: 'shared', index: Uint32Array.from({ length: count }, (_, i) => i % 12) },
        { kind: 'inline', tag: new Uint8Array(count).fill(VALUE_TAG.number), num: values1.values, text: [], textIndex: new Uint32Array(count) },
        { kind: 'inline', tag: new Uint8Array(count).fill(VALUE_TAG.number), num: values2.values, text: [], textIndex: new Uint32Array(count) },
      ],
    }

    const axes = computeAxes(store, cache, def, view)
    expect(axes.ok).toBe(true)
    if (!axes.ok) return
    const grid = computeValues(store, cache, def, view, axes.value)
    expect(grid.ok).toBe(true)
    if (!grid.ok) return

    const grandRow = axes.value.rows.findIndex(line => line.type === 'grand')
    const grandCol = axes.value.cols.findIndex(line => line.type === 'grand')
    expect(grandRow).toBeGreaterThanOrEqual(0)
    expect(grandCol).toBeGreaterThanOrEqual(0)
    expect(grid.value.get(grandRow, grandCol, 0)).toEqual({ kind: 'number', value: values1.sum })
    expect(grid.value.get(grandRow, grandCol, 1)).toEqual({ kind: 'number', value: values2.sum })
  })

  it('refuses a value grid whose estimated workspace exceeds the cap', () => {
    const line = { type: 'data' as const, first: 0, items: [], data: 0 }
    const rowCount = 20_000
    const dataCount = 200
    const data = Array.from({ length: dataCount }, (_, i) => ({
      field: 1, fn: 'sum' as const, showAs: 'normal' as const, baseField: -1, baseItem: 1048832, name: `Sum ${i}`,
    }))
    const axes = {
      rows: new Array(rowCount).fill(line), cols: [line], rowKey: new Uint32Array(), colKey: new Uint32Array(),
    }
    const cache = {
      source: { type: 'worksheet' }, olap: false, fields: [{}, { sharedItems: { inline: true, items: [] } }],
    } as unknown as PivotCacheDefinition
    const store: RecordStore = { count: 0, bytes: 0, columns: [undefined, { kind: 'inline', tag: new Uint8Array(), num: new Float64Array(), text: [], textIndex: new Uint32Array() }] }
    const view = { rows: [], cols: [], pages: [], filters: [], fields: [], data } as unknown as PivotViewState
    const result = computeValues(store, cache, {} as PivotTableDefinition, view, axes)
    expect(result.ok ? null : result.diagnostic.feature).toBe('pivot-values-limit')
  })
})
