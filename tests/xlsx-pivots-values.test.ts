// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { savedPivotLayout } from '../src/xlsx/pivots/savedLayout'
import { NO_LINE, type PivotAxes } from '../src/xlsx/pivots/engine/axes'
import { computeValues } from '../src/xlsx/pivots/engine/values'
import { VALUE_TAG, type RecordStore } from '../src/xlsx/pivots/engine/records'
import { loadAll, savedCell, type Loaded } from './helpers/xlsx-pivots-load'
import type { PivotCacheDefinition, PivotTableDefinition } from '../src/xlsx/pivots/model'
import type { DataFieldState, FieldState, PivotViewState } from '../src/xlsx/pivots/engine/view'
import { initialViewState } from '../src/xlsx/pivots/engine/view'
import { computePivot } from '../src/xlsx/pivots/engine/compute'
import type { PivotCellValue, PlaceOptions } from '../src/xlsx/pivots/engine/place'

const FIXTURES = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const MANIFEST = resolve(FIXTURES, 'engine/manifest.json')
// S2-AA: the only Task 18 refusals left. Their leaf cells match the textbook variance forms, but every tested
// accumulation of the leaf moments (S2-AB leaf rollup, records in display order, per-row or per-column partials,
// Welford with a parallel merge, two-pass, fused multiply-add) leaves 1-9 total cells off in the 15th digit.
const VARIANCE_REFUSAL = 'variance subtotal rounding differs from Excel: its total cells differ in the 15th digit'
const TASK18_UNMEASURED: Record<string, string> = {
  'values-fn-stdev': VARIANCE_REFUSAL,
  'values-fn-stdevp': VARIANCE_REFUSAL,
  'values-fn-var': VARIANCE_REFUSAL,
  'values-fn-varp': VARIANCE_REFUSAL,
}
const referenceNames = readdirSync(FIXTURES).filter(name => /^pivot-.*\.xlsx$/.test(name)).sort()
const loadedReferences = new Map<string, Promise<Loaded[]>>()
const loadReference = (name: string) => {
  let value = loadedReferences.get(name)
  if (!value) loadedReferences.set(name, value = loadAll(resolve(FIXTURES, name), 'skip'))
  return value
}
const optionsOf = (p: Loaded): PlaceOptions => ({ date1904: p.date1904 })
function expectSavedValues(p: Loaded, layout: { cells: Array<{ row: number; col: number; value: PivotCellValue }> }): void {
  const saved = savedPivotLayout(p.def)!
  const at = new Map(layout.cells.map(cell => [`${cell.row},${cell.col}`, cell.value]))
  for (let row = saved.dataRow0; row <= saved.rect.maxRow; row++) for (let col = saved.dataCol0; col <= saved.rect.maxCol; col++) {
    const actual = at.get(`${row},${col}`) ?? { kind: 'empty' as const }
    const expected = savedCell(p.sheet, row, col)
    if (actual.kind === 'number' && expected.kind === 'number') {
      expect(Number(actual.value.toPrecision(15)), `${p.file} ${p.def.name} R${row + 1}C${col + 1}`).toBe(Number(expected.value.toPrecision(15)))
    } else expect(actual, `${p.file} ${p.def.name} R${row + 1}C${col + 1}`).toEqual(expected)
  }
}

const dataFunctions = ['sum', 'count', 'countNums', 'average', 'max', 'min', 'product', 'stdDev', 'stdDevp', 'var', 'varp'] as const

const fieldState: FieldState = {
  hidden: [false, false, false, false], collapsed: [false, false, false, false], sortType: 'manual',
  defaultSubtotal: false, subtotals: [], subtotalTop: false, compact: true, outline: true,
  insertBlankRow: false, fillDownLabels: false, showAll: false,
}

function fixture(values: Array<number | string | null>, data: DataFieldState[] = dataFunctions.map(fn => ({ field: 1, fn, showAs: 'normal', baseField: -1, baseItem: 1048832, name: fn })), axesRows?: PivotAxes['rows'], rowKeys?: number[], axesCols?: PivotAxes['cols'], colKeys?: number[]) {
  const cache = {
    source: { type: 'worksheet' }, olap: false,
    fields: [
      { databaseField: true, sharedItems: { inline: false, items: ['A', 'B', 'C', 'D'].map(value => ({ type: 'string', value })) } },
      { databaseField: true, sharedItems: { inline: true, items: [] } },
    ],
  } as unknown as PivotCacheDefinition
  const def = { fields: [{ items: ['A', 'B', 'C', 'D'].map((_, x) => ({ type: 'data', x })) }, {}] } as unknown as PivotTableDefinition
  const view = {
    rows: [0], cols: [], pages: [], filters: [], fields: [fieldState, fieldState], data,
  } as PivotViewState
  const tags = values.map(value => value === null ? VALUE_TAG.missing : typeof value === 'number' ? VALUE_TAG.number : VALUE_TAG.string)
  const text = [...new Set(values.filter((value): value is string => typeof value === 'string'))]
  const store: RecordStore = {
    count: values.length,
    bytes: 0,
    columns: [
      { kind: 'shared', index: Uint32Array.from(rowKeys ?? values.map(() => 0)) },
      { kind: 'inline', tag: Uint8Array.from(tags), num: Float64Array.from(values, value => typeof value === 'number' ? value : 0), text, textIndex: Uint32Array.from(values, value => typeof value === 'string' ? text.indexOf(value) : 0) },
    ],
  }
  const line = { type: 'data' as const, first: 0, items: [0], data: 0 }
  const axes: PivotAxes = {
    rows: axesRows ?? [line], cols: axesCols ?? [{ ...line, items: [] }],
    rowKey: Uint32Array.from(rowKeys ?? values.map(() => 0)), colKey: Uint32Array.from(colKeys ?? values.map(() => 0)),
  }
  return { store, cache, def, view, axes }
}

describe('pivot value aggregation', () => {
  it('applies show-as to a named sum subtotal (values-showas-on-named-subtotals R10C4)', () => {
    const rows: PivotAxes['rows'] = [
      { type: 'data', first: 0, items: [0, 0], data: 0 },
      { type: 'data', first: 1, items: [0, 1], data: 0 },
      { type: 'sum', first: 0, items: [0], data: 0 },
      { type: 'data', first: 0, items: [1, 0], data: 0 },
      { type: 'data', first: 1, items: [1, 1], data: 0 },
      { type: 'sum', first: 0, items: [1], data: 0 },
    ]
    const f = fixture([1, 2, 3, 4], [
      { field: 1, fn: 'sum', showAs: 'percentOfTotal', baseField: -1, baseItem: 1048832, name: '' },
    ], rows, [0, 0, 1, 1])
    f.view.rows = [0, 0]
    f.view.fields = [{ ...fieldState, subtotals: ['sum'] }, fieldState]
    const axes = { ...f.axes, rowKey: Uint32Array.from([0, 1, 3, 4]) }
    const result = computeValues(f.store, f.cache, f.def, f.view, axes)
    expect(result.ok && [result.value.get(2, 0, 0), result.value.get(5, 0, 0)]).toEqual([
      { kind: 'number', value: 3 / 10 }, { kind: 'number', value: 7 / 10 },
    ])
  })

  it('fails closed for non-finite inline values that would collide with the NaN empty sentinel', () => {
    const f = fixture([Number.NaN], [{ field: 1, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: '' }])
    const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
    expect(result.ok ? null : result.diagnostic.feature).toBe('pivot-engine-unmeasured')
  })

  it('refuses non-finite values on the general function-bank path', () => {
    for (const value of [Number.NaN, Number.POSITIVE_INFINITY]) {
      const f = fixture([value, 2], [{ field: 1, fn: 'average', showAs: 'normal', baseField: -1, baseItem: 1048832, name: 'Average' }])
      const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
      expect(result.ok ? null : result.diagnostic.feature).toBe('pivot-engine-unmeasured')
    }
  })

  it('keeps normal all-sum cells identical when a twin show-as field selects the general path', () => {
    const rows: PivotAxes['rows'] = [
      { type: 'data', first: 0, items: [0], data: 0 },
      { type: 'data', first: 0, items: [1], data: 0 },
    ]
    const fast = fixture([10, 20, 30], [
      { field: 1, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: 'Fast' },
    ], rows, [0, 1, 0])
    const general = fixture([10, 20, 30], [
      { field: 1, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: 'Normal twin' },
      { field: 1, fn: 'sum', showAs: 'percentOfTotal', baseField: -1, baseItem: 1048832, name: 'General-path twin' },
    ], rows, [0, 1, 0])
    const fastResult = computeValues(fast.store, fast.cache, fast.def, fast.view, fast.axes)
    const generalResult = computeValues(general.store, general.cache, general.def, general.view, general.axes)
    expect(fastResult.ok).toBe(true)
    expect(generalResult.ok).toBe(true)
    if (!fastResult.ok || !generalResult.ok) return
    expect(rows.map((_, row) => fastResult.value.get(row, 0, 0))).toEqual(
      rows.map((_, row) => generalResult.value.get(row, 0, 0)),
    )
  })

  it('computes all 11 data functions over numbers, text and missing records', () => {
    const data = dataFunctions.map(fn => ({ field: 1, fn, showAs: 'normal' as const, baseField: -1, baseItem: 1048832, name: fn }))
    const { store, cache, def, view, axes } = fixture([2, 4, 'x', null, -3], data, [{ type: 'data', first: 0, items: [0], data: 0 }])
    const result = computeValues(store, cache, def, view, axes)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    const grid = result.value
    expect(dataFunctions.map((_, index) => grid.get(0, 0, index))).toEqual([
      { kind: 'number', value: 3 },
      { kind: 'number', value: 4 },
      { kind: 'number', value: 3 },
      { kind: 'number', value: 1 },
      { kind: 'number', value: 4 },
      { kind: 'number', value: -3 },
      { kind: 'number', value: -24 },
      { kind: 'number', value: Math.sqrt(13) },
      { kind: 'number', value: Math.sqrt(26 / 3) },
      { kind: 'number', value: 13 },
      { kind: 'number', value: 26 / 3 },
    ])
  })

  it('aggregates average subtotals and grand totals directly from their records', () => {
    const data = [{ field: 1, fn: 'average' as const, showAs: 'normal' as const, baseField: -1, baseItem: 1048832, name: 'Average' }]
    const rows: PivotAxes['rows'] = [
      { type: 'data', first: 0, items: [0], data: 0 },
      { type: 'data', first: 1, items: [0, 0], data: 0 },
      { type: 'data', first: 2, items: [0, 1], data: 0 },
      { type: 'default', first: 0, items: [0], data: 0 },
      { type: 'data', first: 0, items: [1], data: 0 },
      { type: 'data', first: 1, items: [1, 0], data: 0 },
      { type: 'default', first: 0, items: [1], data: 0 },
      { type: 'grand', first: 0, items: [], data: 0 },
    ]
    const { store, cache, def, view } = fixture([0, 10, 100], data, rows, [1, 2, 5])
    view.rows = [0, 0]
    const axes = { rows, cols: [{ type: 'data' as const, first: 0, items: [], data: 0 }], rowKey: Uint32Array.from([1, 2, 5]), colKey: new Uint32Array(3) }
    const result = computeValues(store, cache, def, view, axes)
    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.value.get(3, 0, 0)).toEqual({ kind: 'number', value: 5 })
      expect(result.value.get(6, 0, 0)).toEqual({ kind: 'number', value: 100 })
      expect(result.value.get(7, 0, 0)).toEqual({ kind: 'number', value: 110 / 3 })
    }
  })

  it('computes an average grand total from the source records', () => {
    const data = [{ field: 1, fn: 'average' as const, showAs: 'normal' as const, baseField: -1, baseItem: 1048832, name: 'Average' }]
    const rows: PivotAxes['rows'] = [
      { type: 'data', first: 0, items: [0], data: 0 },
      { type: 'data', first: 0, items: [1], data: 0 },
      { type: 'grand', first: 0, items: [], data: 0 },
    ]
    const f = fixture([10, 30, 100], data, rows, [0, 1, 1])
    const result = computeValues(f.store, f.cache, f.def, f.view, { ...f.axes, rowKey: Uint32Array.from([0, 1, 1]) })
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.value.get(2, 0, 0)).toEqual({ kind: 'number', value: 140 / 3 })
  })

  it('uses the function named by a subtotal line', () => {
    const rows: PivotAxes['rows'] = [
      { type: 'data', first: 1, items: [0, 0], data: 0 },
      { type: 'data', first: 1, items: [0, 1], data: 0 },
      { type: 'avg', first: 0, items: [0], data: 0 },
    ]
    const f = fixture([10, 30], [{ field: 1, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: 'Sum' }], rows, [0, 1])
    const view = { ...f.view, rows: [0, 0] }
    const result = computeValues(f.store, f.cache, f.def, view, f.axes)
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.value.get(2, 0, 0)).toEqual({ kind: 'number', value: 20 })
  })

  it('matches average to record-order sum divided by count at 15 significant digits', () => {
    const values = Array.from({ length: 200 }, (_, i) => Math.sin(i * 12.345) * 1e8)
    const expected = values.reduce((sum, value) => sum + value, 0) / values.length
    const data = [{ field: 1, fn: 'average' as const, showAs: 'normal' as const, baseField: -1, baseItem: 1048832, name: 'Average' }]
    const f = fixture(values, data, [{ type: 'data', first: 0, items: [0], data: 0 }])
    const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
    expect(result.ok).toBe(true)
    if (result.ok) {
      const actual = result.value.get(0, 0, 0)
      expect(actual.kind).toBe('number')
      if (actual.kind === 'number') expect(Number(actual.value.toPrecision(15))).toBe(Number(expected.toPrecision(15)))
    }
  })

  it('returns blank when a record group has no numeric values for numeric aggregates', () => {
    for (const fn of ['sum', 'average', 'max', 'min', 'product', 'stdDevp', 'varp'] as const) {
      const f = fixture(['text', null], [{ field: 1, fn, showAs: 'normal', baseField: -1, baseItem: 1048832, name: fn }])
      const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
      expect(result.ok && result.value.get(0, 0, 0), fn).toEqual({ kind: 'empty' })
    }
  })

  it('excludes filtered records from data, subtotal and grand totals without losing the 1M scan path', () => {
    const rows: PivotAxes['rows'] = [
      { type: 'data', first: 0, items: [0], data: 0 },
      { type: 'data', first: 0, items: [1], data: 0 },
      { type: 'grand', first: 0, items: [], data: 0 },
    ]
    const f = fixture([10, 20, 30], [{ field: 1, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: '' }], rows, [0, 1, 0])
    const result = computeValues(f.store, f.cache, f.def, f.view, { ...f.axes, rowKey: Uint32Array.from([0, NO_LINE, 0]) })
    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.value.get(0, 0, 0)).toEqual({ kind: 'number', value: 40 })
      expect(result.value.get(1, 0, 0)).toEqual({ kind: 'empty' })
      expect(result.value.get(2, 0, 0)).toEqual({ kind: 'number', value: 40 })
    }
  })

  it('computes difference from previous and next base items', () => {
    const baseRows: PivotAxes['rows'] = [0, 1, 2].map(x => ({ type: 'data', first: 0, items: [x], data: 0 }))
    const compute = (baseItem: number) => {
      const data = [{ field: 1, fn: 'sum' as const, showAs: 'difference' as const, baseField: 0, baseItem, name: '' }]
      const f = fixture([10, 30, 60], data, baseRows, [0, 1, 2])
      return computeValues(f.store, f.cache, f.def, f.view, { ...f.axes, rowKey: Uint32Array.from([0, 1, 2]) })
    }
    const previous = compute(1048828), next = compute(1048829)
    expect(previous.ok && [0, 1, 2].map(row => previous.value.get(row, 0, 0))).toEqual([
      { kind: 'empty' }, { kind: 'number', value: 20 }, { kind: 'number', value: 30 },
    ])
    expect(next.ok && [0, 1, 2].map(row => next.value.get(row, 0, 0))).toEqual([
      { kind: 'number', value: -20 }, { kind: 'number', value: -30 }, { kind: 'empty' },
    ])
  })

  it('computes percent of parent total from direct child records', () => {
    const rows: PivotAxes['rows'] = [
      { type: 'data', first: 1, items: [0, 0], data: 0 },
      { type: 'data', first: 1, items: [0, 1], data: 0 },
      { type: 'default', first: 0, items: [0], data: 0 },
      { type: 'data', first: 1, items: [1, 0], data: 0 },
      { type: 'default', first: 0, items: [1], data: 0 },
    ]
    const data = [{ field: 1, fn: 'sum' as const, showAs: 'percentOfParent' as const, baseField: 0, baseItem: 1048832, name: '' }]
    const f = fixture([20, 30, 50], data, rows, [0, 1, 3])
    f.view.rows = [0, 0]
    const axes = { ...f.axes, rowKey: Uint32Array.from([0, 1, 3]) }
    const result = computeValues(f.store, f.cache, f.def, f.view, axes)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.get(0, 0, 0)).toEqual({ kind: 'number', value: 0.4 })
    expect(result.value.get(1, 0, 0)).toEqual({ kind: 'number', value: 0.6 })
    expect(result.value.get(3, 0, 0)).toEqual({ kind: 'number', value: 1 })
  })

  it('computes the x14 parent-row and parent-column totals', () => {
    const hierarchy: PivotAxes['rows'] = [
      { type: 'data', first: 1, items: [0, 0], data: 0 },
      { type: 'data', first: 1, items: [0, 1], data: 0 },
      { type: 'default', first: 0, items: [0], data: 0 },
      { type: 'grand', first: 0, items: [], data: 0 },
    ]
    const parentRow = fixture([20, 30], [{ field: 1, fn: 'sum', showAs: 'percentOfParentRow', baseField: -1, baseItem: 1048832, name: '' }], hierarchy, [0, 1])
    const parentRowView = { ...parentRow.view, rows: [0, 0] }
    const rowResult = computeValues(parentRow.store, parentRow.cache, parentRow.def, parentRowView, parentRow.axes)
    expect(rowResult.ok).toBe(true)
    if (!rowResult.ok) return
    expect([rowResult.value.get(0, 0, 0), rowResult.value.get(1, 0, 0)]).toEqual([
      { kind: 'number', value: 0.4 }, { kind: 'number', value: 0.6 },
    ])

    const parentCol = fixture([20, 30], [{ field: 1, fn: 'sum', showAs: 'percentOfParentCol', baseField: -1, baseItem: 1048832, name: '' }])
    const parentColView = { ...parentCol.view, rows: [], cols: [0, 0] }
    const colAxes: PivotAxes = {
      rows: [{ type: 'data', first: 0, items: [], data: 0 }], cols: hierarchy,
      rowKey: new Uint32Array([0, 0]), colKey: new Uint32Array([0, 1]),
    }
    const colResult = computeValues(parentCol.store, parentCol.cache, parentCol.def, parentColView, colAxes)
    expect(colResult.ok).toBe(true)
    if (!colResult.ok) return
    expect([colResult.value.get(0, 0, 0), colResult.value.get(0, 1, 0)]).toEqual([
      { kind: 'number', value: 0.4 }, { kind: 'number', value: 0.6 },
    ])
  })

  it('preserves record-order Float64 addition for the 15-significant-digit rule', () => {
    const data = [{ field: 1, fn: 'sum' as const, showAs: 'normal' as const, baseField: -1, baseItem: 1048832, name: '' }]
    const f = fixture([1e16, 1, -1e16], data, [{ type: 'data', first: 0, items: [0], data: 0 }])
    const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.value.get(0, 0, 0)).toEqual({ kind: 'number', value: 0 })
  })

  it('rolls sum subtotals up from leaf cells in display order (S2-AB)', () => {
    const rows: PivotAxes['rows'] = [
      { type: 'data', first: 1, items: [0, 0], data: 0 },
      { type: 'data', first: 1, items: [0, 1], data: 0 },
      { type: 'default', first: 0, items: [0], data: 0 },
      { type: 'grand', first: 0, items: [], data: 0 },
    ]
    // Record order: A(1e16), B(1), A(-1e16). Leaf A sums to 0, then B to 1.
    const f = fixture([1e16, 1, -1e16], [{ field: 1, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: '' }], rows, [0, 1, 0])
    f.view.rows = [0, 0]
    const axes = { ...f.axes, rowKey: Uint32Array.from([0, 1, 0]) }
    const result = computeValues(f.store, f.cache, f.def, f.view, axes)
    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.value.get(0, 0, 0)).toEqual({ kind: 'number', value: 0 })
      expect(result.value.get(1, 0, 0)).toEqual({ kind: 'number', value: 1 })
      expect(result.value.get(2, 0, 0)).toEqual({ kind: 'number', value: 1 })
      expect(result.value.get(3, 0, 0)).toEqual({ kind: 'number', value: 1 })
    }
  })

  it('rolls sum totals as the sum over leaf rows of each leaf row\'s sum over leaf columns (S2-AB, bit-exact)', () => {
    // Leaf cells a b / c d = 0.1 0.1 / 0.2 0.7: (a+b)+(c+d) = 1.0999999999999999, while a row-major running sum
    // and a column-first sum both give 1.1. Excel's baseline R24C8 and R6C8 follow the row-partial order.
    const rows: PivotAxes['rows'] = [
      { type: 'data', first: 0, items: [0], data: 0 }, { type: 'data', first: 0, items: [1], data: 0 },
      { type: 'grand', first: 0, items: [], data: 0 },
    ]
    const cols: PivotAxes['cols'] = [
      { type: 'data', first: 0, items: [0], data: 0 }, { type: 'data', first: 0, items: [1], data: 0 },
      { type: 'grand', first: 0, items: [], data: 0 },
    ]
    for (const showAs of ['normal', 'percentOfTotal'] as const) {
      const f = fixture([0.1, 0.1, 0.2, 0.7], [{ field: 1, fn: 'sum', showAs, baseField: -1, baseItem: 1048832, name: '' }], rows, [0, 0, 1, 1], cols, [0, 1, 0, 1])
      f.view.cols = [0]
      const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
      expect(result.ok, showAs).toBe(true)
      if (!result.ok) continue
      if (showAs === 'normal') {
        expect(result.value.get(2, 2, 0)).toEqual({ kind: 'number', value: (0.1 + 0.1) + (0.2 + 0.7) })
        expect(result.value.get(2, 2, 0)).not.toEqual({ kind: 'number', value: 1.1 })
      } else expect(result.value.get(1, 1, 0)).toEqual({ kind: 'number', value: 0.7 / ((0.1 + 0.1) + (0.2 + 0.7)) })
    }
  })

  it('sums an average total row record by record, leaf row after leaf row (values-average-hidden-item R20C6)', () => {
    const rows: PivotAxes['rows'] = [
      { type: 'data', first: 0, items: [0], data: 0 }, { type: 'data', first: 0, items: [1], data: 0 },
      { type: 'grand', first: 0, items: [], data: 0 },
    ]
    // Records A 0.1, B 0.1, B 1: (0.1 + 0.1) + 1 = 1.2, while leaf sums 0.1 + (0.1 + 1) = 1.2000000000000002.
    const f = fixture([0.1, 0.1, 1], [{ field: 1, fn: 'average', showAs: 'normal', baseField: -1, baseItem: 1048832, name: '' }], rows, [0, 1, 1])
    const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.value.get(2, 0, 0)).toEqual({ kind: 'number', value: ((0.1 + 0.1) + 1) / 3 })
  })

  it('shows Count Numbers as 0 over text and blank without values; Count is blank without values', () => {
    const rows: PivotAxes['rows'] = [
      { type: 'data', first: 0, items: [0], data: 0 }, { type: 'data', first: 0, items: [1], data: 0 },
    ]
    for (const fn of ['countNums', 'count'] as const) {
      const f = fixture(['text', null], [{ field: 1, fn, showAs: 'normal', baseField: -1, baseItem: 1048832, name: fn }], rows, [0, 1])
      const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
      expect(result.ok && [result.value.get(0, 0, 0), result.value.get(1, 0, 0)], fn).toEqual([
        { kind: 'number', value: fn === 'countNums' ? 0 : 1 }, { kind: 'empty' },
      ])
    }
  })

  it('applies the measured difference, % of and % difference rules to blank, own and missing base cells', () => {
    // Rows: item 0 (base), item 1 (records), item 2 (no records); explicit base item 0.
    const rows: PivotAxes['rows'] = [0, 1, 2].map(x => ({ type: 'data' as const, first: 0, items: [x], data: 0 }))
    const run = (showAs: DataFieldState['showAs'], values: Array<number | null>, keys: number[], baseItem = 0) => {
      const f = fixture(values, [{ field: 1, fn: 'sum', showAs, baseField: 0, baseItem, name: '' }], rows, keys)
      const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
      return result.ok ? [0, 1, 2].map(row => result.value.get(row, 0, 0)) : result.diagnostic.message
    }
    // Difference: own line blank; blank current counts as 0.
    expect(run('difference', [10, 30], [0, 1])).toEqual([{ kind: 'empty' }, { kind: 'number', value: 20 }, { kind: 'number', value: -10 }])
    // % of: own line 1; blank current with a base value is #NULL!.
    expect(run('percent', [10, 30], [0, 1])).toEqual([{ kind: 'number', value: 1 }, { kind: 'number', value: 3 }, { kind: 'error', value: '#NULL!' }])
    // % difference: own line blank; a zero base is #DIV/0!; a blank base leaves the cell blank.
    expect(run('percentDiff', [0, 30, 5], [0, 1, 2])).toEqual([{ kind: 'empty' }, { kind: 'error', value: '#DIV/0!' }, { kind: 'error', value: '#DIV/0!' }])
    // Unmeasured, so refused: % difference of a blank cell from a base value, and % of a blank base cell.
    expect(run('percentDiff', [10, 30], [0, 1])).toBe('% difference of a blank cell from its base item is not measured')
    expect(run('percent', [30], [1])).toBe('% of a base item with a blank base cell is not measured')
    expect(run('percentDiff', [30, 40], [1, 2])).toEqual([{ kind: 'empty' }, { kind: 'empty' }, { kind: 'empty' }])
    // Difference from previous: a blank base counts as 0 (difference-previous R22C4).
    expect(run('difference', [30], [1], 1048828)).toEqual([{ kind: 'empty' }, { kind: 'number', value: 30 }, { kind: 'number', value: -30 }])
    // An explicit base item without a line is #N/A (difference-item R23).
    expect(run('difference', [30], [1], 3)).toEqual([{ kind: 'error', value: '#N/A' }, { kind: 'error', value: '#N/A' }, { kind: 'error', value: '#N/A' }])
  })

  it('computes index as value × grand ÷ row total ÷ column total, and 0 for a cell without records', () => {
    const rows: PivotAxes['rows'] = [0, 1].map(x => ({ type: 'data' as const, first: 0, items: [x], data: 0 }))
    const cols: PivotAxes['cols'] = [0, 1].map(x => ({ type: 'data' as const, first: 0, items: [x], data: 0 }))
    const f = fixture([10, 30, 20], [{ field: 1, fn: 'sum', showAs: 'index', baseField: -1, baseItem: 1048832, name: '' }], rows, [0, 0, 1], cols, [0, 1, 0])
    f.view.cols = [0]
    const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.get(0, 0, 0)).toEqual({ kind: 'number', value: 10 * 60 / 40 / 30 })
    expect(result.value.get(1, 1, 0)).toEqual({ kind: 'number', value: 0 })
  })

  it('leaves % of parent total with a base field blank on the grand line (percent-of-parent-region R60C2)', () => {
    const rows: PivotAxes['rows'] = [
      { type: 'data', first: 0, items: [0], data: 0 }, { type: 'data', first: 0, items: [1], data: 0 },
      { type: 'grand', first: 0, items: [], data: 0 },
    ]
    const f = fixture([10, 30], [{ field: 1, fn: 'sum', showAs: 'percentOfParent', baseField: 0, baseItem: 1048832, name: '' }], rows, [0, 1])
    const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
    expect(result.ok && result.value.get(2, 0, 0)).toEqual({ kind: 'empty' })
  })

  it('shows #DIV/0! for a sample statistic over one number and uses the measured textbook variance forms', () => {
    for (const fn of ['stdDev', 'var'] as const) {
      const f = fixture([10], [{ field: 1, fn, showAs: 'normal', baseField: -1, baseItem: 1048832, name: fn }])
      const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
      expect(result.ok && result.value.get(0, 0, 0), fn).toEqual({ kind: 'error', value: '#DIV/0!' })
    }
    // values-fn-stdevp R12C2: Excel 27.65999999999757 for a two-record leaf, the textbook n·Σx² − (Σx)² form.
    const xs = [1e6 + 0.1, 1e6 + 55.42]
    const n = 2, s = xs[0]! + xs[1]!, q = xs[0]! * xs[0]! + xs[1]! * xs[1]!
    for (const [fn, expected] of [['varp', (n * q - s * s) / (n * n)], ['var', (q - s * s / n) / (n - 1)]] as const) {
      const f = fixture(xs, [{ field: 1, fn, showAs: 'normal', baseField: -1, baseItem: 1048832, name: fn }])
      const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
      expect(result.ok && result.value.get(0, 0, 0), fn).toEqual({ kind: 'number', value: expected })
    }
  })

  it('fails closed when the aggregate or the default show-as base overflows or is unmeasured', () => {
    const sum = fixture([Number.MAX_VALUE, Number.MAX_VALUE], [{ field: 1, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: '' }])
    const overflow = computeValues(sum.store, sum.cache, sum.def, sum.view, sum.axes)
    expect(overflow.ok ? null : overflow.diagnostic.feature).toBe('pivot-engine-unmeasured')

    const previous = fixture([10, 20], [{ field: 1, fn: 'sum', showAs: 'difference', baseField: 0, baseItem: 1048832, name: '' }], [
      { type: 'data', first: 0, items: [0], data: 0 }, { type: 'data', first: 0, items: [1], data: 0 },
    ], [0, 1])
    const defaultBase = computeValues(previous.store, previous.cache, previous.def, previous.view, previous.axes)
    expect(defaultBase.ok ? null : defaultBase.diagnostic.feature).toBe('pivot-engine-unmeasured')
  })

  it('computes percent totals, base-item percent and running total in the selected field order', () => {
    const rows: PivotAxes['rows'] = [0, 1, 2].map(x => ({ type: 'data', first: 0, items: [x], data: 0 }))
    const modes: DataFieldState[] = [
      { field: 1, fn: 'sum', showAs: 'percentOfTotal', baseField: -1, baseItem: 1048832, name: '' },
      { field: 1, fn: 'sum', showAs: 'percent', baseField: 0, baseItem: 0, name: '' },
      { field: 1, fn: 'sum', showAs: 'runTotal', baseField: 0, baseItem: 1048832, name: '' },
      { field: 1, fn: 'sum', showAs: 'percentOfRunningTotal', baseField: 0, baseItem: 1048832, name: '' },
    ]
    const f = fixture([10, 30, 60], modes, rows, [0, 1, 2])
    const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect([0, 1, 2].map(row => result.value.get(row, 0, 0))).toEqual([
      { kind: 'number', value: 0.1 }, { kind: 'number', value: 0.3 }, { kind: 'number', value: 0.6 },
    ])
    expect([0, 1, 2].map(row => result.value.get(row, 0, 1))).toEqual([
      { kind: 'number', value: 1 }, { kind: 'number', value: 3 }, { kind: 'number', value: 6 },
    ])
    expect([0, 1, 2].map(row => result.value.get(row, 0, 2))).toEqual([
      { kind: 'number', value: 10 }, { kind: 'number', value: 40 }, { kind: 'number', value: 100 },
    ])
    expect([0, 1, 2].map(row => result.value.get(row, 0, 3))).toEqual([
      { kind: 'number', value: 0.1 }, { kind: 'number', value: 0.4 }, { kind: 'number', value: 1 },
    ])
  })

  it('uses the grand total as the parent of top-level items and subtotal lines', () => {
    const rows: PivotAxes['rows'] = [
      { type: 'data', first: 1, items: [0, 0], data: 0 },
      { type: 'data', first: 1, items: [0, 1], data: 0 },
      { type: 'default', first: 0, items: [0], data: 0 },
      { type: 'data', first: 1, items: [1, 0], data: 0 },
      { type: 'data', first: 1, items: [1, 1], data: 0 },
      { type: 'default', first: 0, items: [1], data: 0 },
      { type: 'grand', first: 0, items: [], data: 0 },
    ]
    const data = [{ field: 1, fn: 'sum' as const, showAs: 'percentOfParentRow' as const, baseField: -1, baseItem: 1048832, name: '' }]
    const f = fixture([10, 30, 20, 40], data, rows, [0, 1, 3, 4])
    const view = { ...f.view, rows: [0, 0] }
    const result = computeValues(f.store, f.cache, f.def, view, f.axes)
    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.value.get(2, 0, 0)).toEqual({ kind: 'number', value: 0.4 })
      expect(result.value.get(5, 0, 0)).toEqual({ kind: 'number', value: 0.6 })
    }
  })

  it('shows 100 percent for an item at the selected parent field level', () => {
    const rows: PivotAxes['rows'] = [
      { type: 'data', first: 0, items: [0], data: 0 },
      { type: 'data', first: 0, items: [1], data: 0 },
      { type: 'grand', first: 0, items: [], data: 0 },
    ]
    const data = [{ field: 1, fn: 'sum' as const, showAs: 'percentOfParent' as const, baseField: 0, baseItem: 1048832, name: '' }]
    const f = fixture([10, 30], data, rows, [0, 1])
    const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
    expect(result.ok).toBe(true)
    if (result.ok) expect([result.value.get(0, 0, 0), result.value.get(1, 0, 0)]).toEqual([
      { kind: 'number', value: 1 }, { kind: 'number', value: 1 },
    ])
  })

  it('leaves the selected base item\'s own difference cells blank (values-showas-difference-item rows 10-13)', () => {
    const data = [{ field: 1, fn: 'sum' as const, showAs: 'difference' as const, baseField: 0, baseItem: 0, name: '' }]
    const rows: PivotAxes['rows'] = [0, 1].map(x => ({ type: 'data', first: 0, items: [x], data: 0 }))
    const f = fixture([10, 30], data, rows, [0, 1])
    const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.value.get(0, 0, 0)).toEqual({ kind: 'empty' })
  })

  it('computes percent of row, percent of column, and index from record-based totals', () => {
    const rows: PivotAxes['rows'] = [0, 1].map(x => ({ type: 'data', first: 0, items: [x], data: 0 }))
    const cols: PivotAxes['cols'] = [0, 1].map(x => ({ type: 'data', first: 0, items: [x], data: 0 }))
    const state = (showAs: DataFieldState['showAs']): DataFieldState => ({ field: 1, fn: 'sum', showAs, baseField: -1, baseItem: 1048832, name: '' })
    const modes = [state('percentOfRow'), state('percentOfCol'), state('index')]
    const f = fixture([10, 30, 20, 40], modes, rows, [0, 0, 1, 1], cols, [0, 1, 0, 1])
    const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect([result.value.get(0, 0, 0), result.value.get(0, 1, 0)]).toEqual([
      { kind: 'number', value: 0.25 }, { kind: 'number', value: 0.75 },
    ])
    expect([result.value.get(0, 0, 1), result.value.get(1, 0, 1)]).toEqual([
      { kind: 'number', value: 1 / 3 }, { kind: 'number', value: 2 / 3 },
    ])
    expect(result.value.get(0, 0, 2)).toEqual({ kind: 'number', value: 10 * 100 / 40 / 30 })
  })

  it('computes ascending and descending ranks when the compared values have no ties', () => {
    const rows: PivotAxes['rows'] = [0, 1, 2].map(x => ({ type: 'data', first: 0, items: [x], data: 0 }))
    const state = (showAs: DataFieldState['showAs']): DataFieldState => ({ field: 1, fn: 'sum', showAs, baseField: 0, baseItem: 1048832, name: '' })
    const f = fixture([10, 30, 60], [state('rankAscending'), state('rankDescending')], rows, [0, 1, 2])
    const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect([0, 1, 2].map(row => result.value.get(row, 0, 0))).toEqual([
      { kind: 'number', value: 1 }, { kind: 'number', value: 2 }, { kind: 'number', value: 3 },
    ])
    expect([0, 1, 2].map(row => result.value.get(row, 0, 1))).toEqual([
      { kind: 'number', value: 3 }, { kind: 'number', value: 2 }, { kind: 'number', value: 1 },
    ])
  })

  it('uses Excel rank ties, error values and visible records, and exposes divide-by-zero errors', () => {
    const ranked = fixture([1, 1], [{ field: 1, fn: 'sum', showAs: 'rankAscending', baseField: 0, baseItem: 1048832, name: '' }], [
      { type: 'data', first: 0, items: [0], data: 0 }, { type: 'data', first: 0, items: [1], data: 0 },
    ], [0, 1])
    const rank = computeValues(ranked.store, ranked.cache, ranked.def, ranked.view, ranked.axes)
    expect(rank.ok).toBe(true)
    if (rank.ok) expect([rank.value.get(0, 0, 0), rank.value.get(1, 0, 0)]).toEqual([
      { kind: 'number', value: 1 }, { kind: 'number', value: 1 },
    ])

    const zero = fixture([0, 0], [{ field: 1, fn: 'sum', showAs: 'percentOfTotal', baseField: -1, baseItem: 1048832, name: '' }])
    const divideByZero = computeValues(zero.store, zero.cache, zero.def, zero.view, zero.axes)
    expect(divideByZero.ok).toBe(true)
    if (divideByZero.ok) expect(divideByZero.value.get(0, 0, 0)).toEqual({ kind: 'error', value: '#DIV/0!' })

    const error = fixture(['#N/A'], [{ field: 1, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: '' }])
    const errorColumn = error.store.columns[1]!
    if (errorColumn.kind !== 'inline') throw new Error('inline value field expected')
    errorColumn.tag[0] = VALUE_TAG.error
    const errorResult = computeValues(error.store, error.cache, error.def, error.view, error.axes)
    expect(errorResult.ok).toBe(true)
    if (errorResult.ok) expect(errorResult.value.get(0, 0, 0)).toEqual({ kind: 'error', value: '#N/A' })

    const filtered = fixture([1, 2], [{ field: 1, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: '' }])
    filtered.axes.rowKey[1] = 0xffffffff
    const filteredResult = computeValues(filtered.store, filtered.cache, filtered.def, filtered.view, filtered.axes)
    expect(filteredResult.ok).toBe(true)
  })
})

describe('computePivot composition and Excel saved values', () => {
  it.each(referenceNames)('%s: computed value cells match the saved cells to 15 significant digits', async name => {
    const cases = await loadReference(name)
    for (const p of cases) {
      const view = initialViewState(p.def, p.cache)
      const result = computePivot(p.store, p.cache, p.def, view, optionsOf(p))
      expect(result.ok, `${name} ${p.def.name}: ${result.ok ? '' : result.diagnostic.feature + ': ' + result.diagnostic.message}`).toBe(true)
      if (!result.ok) continue
      expectSavedValues(p, result.value)
    }
  })

  if (!existsSync(MANIFEST)) {
    it.skip('SKIPPED: Task 15 engine/manifest.json is absent; task:18 oracle cases not available', () => {})
  } else {
    it('runs every non-pending Task 15 manifest case for Task 18', async () => {
      const raw = JSON.parse(readFileSync(MANIFEST, 'utf8')) as { cases?: Array<{ task: number; status?: string; case: string; file: string; pivot: string }> } | Array<{ task: number; status?: string; case: string; file: string; pivot: string }>
      const cases = Array.isArray(raw) ? raw : raw.cases ?? []
      for (const entry of cases.filter(item => item.task === 18 && item.status !== 'pending')) {
        const loaded = await loadAll(resolve(process.cwd(), entry.file), 'skip')
        const p = loaded.find(candidate => candidate.def.name === entry.pivot)
        if (!p) throw new Error(`${entry.file}: no pivot named ${entry.pivot}`)
        const result = computePivot(p.store, p.cache, p.def, initialViewState(p.def, p.cache), optionsOf(p))
        const refusal = TASK18_UNMEASURED[entry.case]
        if (refusal) {
          expect(result.ok, `${entry.case} should refuse its exact unmeasured construct`).toBe(false)
          if (!result.ok) {
            expect(result.diagnostic.feature).toBe('pivot-engine-unmeasured')
            expect(result.diagnostic.message).toContain(refusal)
          }
          continue
        }
        expect(result.ok, `${entry.file} ${entry.pivot}: ${result.ok ? '' : result.diagnostic.message}`).toBe(true)
        if (result.ok) expectSavedValues(p, result.value)
      }
    })
  }
})
