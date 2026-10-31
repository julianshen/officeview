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
    // S2-AB: each leaf row sums its leaf cells across the columns, then the grand total sums those leaf-row
    // totals down the rows, both in display order (bit-exact on baseline R24C8), never the records directly.
    const cellSums = [new Float64Array(axes.value.rows.length * axes.value.cols.length),
      new Float64Array(axes.value.rows.length * axes.value.cols.length)]
    for (let record = 0; record < count; record++) {
      const cell = axes.value.rowKey[record]! * axes.value.cols.length + axes.value.colKey[record]!
      cellSums[0]![cell] += values1.values[record]!
      cellSums[1]![cell] += values2.values[record]!
    }
    const leafRows = axes.value.rows.map((line, index) => line.type === 'data' && line.items.length === 3 ? index : -1).filter(index => index >= 0)
    const leafCols = axes.value.cols.map((line, index) => line.type === 'data' && line.items.length === 1 ? index : -1).filter(index => index >= 0)
    for (let data = 0; data < 2; data++) {
      let expected = 0
      for (const row of leafRows) {
        let rowTotal = 0
        for (const col of leafCols) rowTotal += cellSums[data]![row * axes.value.cols.length + col]!
        expected += rowTotal
      }
      expect(grid.value.get(grandRow, grandCol, data)).toEqual({ kind: 'number', value: expected })
    }
  })

  it('gives bit-identical Sum cells on the fast path, the general path and an independent S2-AB reference', () => {
    // Random two-level row and one-level column pivots, with blanks and decimals whose sums depend on addition order.
    for (const [seed, subtotalTop] of [[1, false], [2, true], [3, false], [4, true], [5, false]] as const) {
      let state = seed * 2654435761 >>> 0
      const rand = () => (state = (Math.imul(state, 1664525) + 1013904223) >>> 0) / 4294967296
      const count = 400, cardinalities = [4, 5, 3]
      const sharedItems = cardinalities.map(size => ({ inline: false, items: Array.from({ length: size }, (_, i) => ({ type: 'string', value: String(i) })) }))
      const cache = {
        source: { type: 'worksheet' }, olap: false,
        fields: [...sharedItems.map(items => ({ databaseField: true, sharedItems: items })), { databaseField: true, sharedItems: { inline: true, items: [] } }],
      } as unknown as PivotCacheDefinition
      const def = {
        fields: [...cardinalities.map(size => ({ items: Array.from({ length: size }, (_, x) => ({ type: 'data', x })), defaultSubtotal: true })), { items: [] }],
        rowGrandTotals: true, colGrandTotals: true,
      } as unknown as PivotTableDefinition
      const fields = cardinalities.map((size, field) => ({ ...makeFieldState(size, field < 2), subtotalTop }))
      fields.push(makeFieldState(0, false))
      const index = cardinalities.map(size => Uint32Array.from({ length: count }, () => Math.floor(rand() * size)))
      const tag = Uint8Array.from({ length: count }, () => rand() < 0.1 ? VALUE_TAG.missing : VALUE_TAG.number)
      const num = Float64Array.from({ length: count }, () => Math.round((rand() - 0.3) * 1e6) / 100)
      const store: RecordStore = {
        count, bytes: 0,
        columns: [...index.map(column => ({ kind: 'shared' as const, index: column })), { kind: 'inline', tag, num, text: [], textIndex: new Uint32Array(count) }],
      }
      const sum = { field: 3, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: 'Sum' }
      const view = { rows: [0, 1], cols: [2], pages: [], filters: [], fields, data: [sum] } as unknown as PivotViewState
      const axes = computeAxes(store, cache, def, view)
      expect(axes.ok).toBe(true)
      if (!axes.ok) return
      const fast = computeValues(store, cache, def, view, axes.value)
      const general = computeValues(store, cache, def, { ...view, data: [sum, { ...sum, fn: 'count', name: 'Count' }] } as unknown as PivotViewState, axes.value)
      expect(fast.ok && general.ok).toBe(true)
      if (!fast.ok || !general.ok) return
      // Reference: leaf cells in record order; any total = Σ over its leaf rows of (Σ over its leaf columns).
      const { rows, cols, rowKey, colKey } = axes.value
      const leaf = new Map<string, number>()
      for (let record = 0; record < count; record++) {
        if (tag[record] !== VALUE_TAG.number) continue
        const key = `${rowKey[record]},${colKey[record]}`
        leaf.set(key, (leaf.get(key) ?? 0) + num[record]!)
      }
      const isLeaf = (lines: typeof rows, line: number) => lines[line]!.type === 'data'
        && !lines.some((other, k) => k !== line && other.items.length > lines[line]!.items.length && lines[line]!.items.every((item, n) => other.items[n] === item))
      const leavesOf = (lines: typeof rows, line: number) => lines.flatMap((other, k) => isLeaf(lines, k)
        && (lines[line]!.type === 'grand' || lines[line]!.items.every((item, n) => other.items[n] === item)) ? [k] : [])
      let compared = 0
      rows.forEach((_, row) => cols.forEach((__, col) => {
        let total = 0, present = false
        for (const leafRow of leavesOf(rows, row)) {
          let rowTotal = 0
          for (const leafCol of leavesOf(cols, col)) {
            const value = leaf.get(`${leafRow},${leafCol}`)
            if (value !== undefined) { rowTotal += value; present = true }
          }
          total += rowTotal
        }
        const fastCell = fast.value.get(row, col, 0)
        expect(general.value.get(row, col, 0), `seed ${seed} ${row},${col}`).toEqual(fastCell)
        if (fastCell.kind === 'empty') return
        compared++
        expect(fastCell, `seed ${seed} ${row},${col}`).toEqual(present ? { kind: 'number', value: total } : { kind: 'empty' })
      }))
      expect(compared).toBeGreaterThan(40)
    }
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
