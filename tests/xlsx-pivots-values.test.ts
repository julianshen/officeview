// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { loadPivotWorkbook } from '../src/xlsx/pivots/load'
import { savedPivotLayout } from '../src/xlsx/pivots/savedLayout'
import type { PivotAxes } from '../src/xlsx/pivots/engine/axes'
import { computeValues } from '../src/xlsx/pivots/engine/values'
import { decodeRecords, RECORD_LIMITS, VALUE_TAG, type RecordStore } from '../src/xlsx/pivots/engine/records'
import type { PivotCacheDefinition, PivotTableDefinition } from '../src/xlsx/pivots/model'
import type { DataFieldState, FieldState, PivotViewState } from '../src/xlsx/pivots/engine/view'
import { initialViewState } from '../src/xlsx/pivots/engine/view'
import { computePivot } from '../src/xlsx/pivots/engine/compute'
import type { PivotCellValue, PlaceOptions } from '../src/xlsx/pivots/engine/place'

const FIXTURES = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const MANIFEST = resolve(FIXTURES, 'engine/manifest.json')
interface Loaded { file: string; sheet: { rows: Array<{ index: number; cells: Array<{ col: number; value: unknown; valueIsError?: boolean }> }> }; def: PivotTableDefinition; cache: PivotCacheDefinition; store: RecordStore; date1904: boolean }
async function* once(bytes: Uint8Array): AsyncIterable<Uint8Array> { yield bytes }
async function loadAll(path: string): Promise<Loaded[]> {
  const pkg = await OfficePackage.load(new Uint8Array(readFileSync(path)))
  const doc = await parseXlsx(pkg)
  const regions = new Map(doc.sheets.filter(sheet => sheet.pivotRegions?.length).map(sheet => [sheet.name, sheet.pivotRegions!]))
  const workbook = await loadPivotWorkbook(pkg, regions)
  const date1904 = doc.sheets[0]?.semantics?.dateSystem === '1904'
  const out: Loaded[] = []
  for (const def of workbook.tables) {
    const cache = workbook.caches.get(def.cacheId)
    if (!cache?.recordsPartPath) continue
    const bytes = await pkg.bytes(cache.recordsPartPath)
    const result = await decodeRecords(once(bytes!), cache, RECORD_LIMITS.D, date1904)
    if (!result.ok) throw new Error(`${path}: ${result.diagnostic.message}`)
    out.push({ file: path, sheet: doc.sheets.find(sheet => sheet.name === def.sheet)!, def, cache, store: result.store, date1904 })
  }
  return out
}
const referenceNames = readdirSync(FIXTURES).filter(name => /^pivot-.*\.xlsx$/.test(name)).sort()
const loadedReferences = new Map<string, Promise<Loaded[]>>()
const loadReference = (name: string) => {
  let value = loadedReferences.get(name)
  if (!value) loadedReferences.set(name, value = loadAll(resolve(FIXTURES, name)))
  return value
}
function savedCell(sheet: Loaded['sheet'], row: number, col: number): PivotCellValue {
  const cell = sheet.rows.find(entry => entry.index === row)?.cells.find(entry => entry.col === col)
  if (!cell || cell.value === null || cell.value === '') return { kind: 'empty' }
  if (cell.valueIsError) return { kind: 'error', value: String(cell.value) }
  if (typeof cell.value === 'number') return { kind: 'number', value: cell.value }
  return { kind: 'text', value: String(cell.value) }
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

  it('fails closed when a record group has no numeric values for numeric aggregates', () => {
    for (const fn of ['average', 'max', 'min', 'product', 'stdDevp', 'varp'] as const) {
      const f = fixture(['text', null], [{ field: 1, fn, showAs: 'normal', baseField: -1, baseItem: 1048832, name: fn }])
      const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
      expect(result.ok ? null : result.diagnostic.feature, fn).toBe('pivot-engine-unmeasured')
    }
    for (const fn of ['stdDev', 'var'] as const) {
      const f = fixture([10], [{ field: 1, fn, showAs: 'normal', baseField: -1, baseItem: 1048832, name: fn }])
      const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
      expect(result.ok ? null : result.diagnostic.feature, fn).toBe('pivot-engine-unmeasured')
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

  it('fails closed on difference-from cells for the selected base item', () => {
    const data = [{ field: 1, fn: 'sum' as const, showAs: 'difference' as const, baseField: 0, baseItem: 0, name: '' }]
    const rows: PivotAxes['rows'] = [0, 1].map(x => ({ type: 'data', first: 0, items: [x], data: 0 }))
    const f = fixture([10, 30], data, rows, [0, 1])
    const result = computeValues(f.store, f.cache, f.def, f.view, f.axes)
    expect(result.ok ? null : result.diagnostic.feature).toBe('pivot-engine-unmeasured')
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
    expect(result.value.get(0, 0, 2)).toEqual({ kind: 'number', value: 1000 / 1200 })
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

  it('fails closed for rank ties, a division-by-zero show-as denominator, error values and filtered records', () => {
    const ranked = fixture([1, 1], [{ field: 1, fn: 'sum', showAs: 'rankAscending', baseField: 0, baseItem: 1048832, name: '' }], [
      { type: 'data', first: 0, items: [0], data: 0 }, { type: 'data', first: 0, items: [1], data: 0 },
    ], [0, 1])
    const rank = computeValues(ranked.store, ranked.cache, ranked.def, ranked.view, ranked.axes)
    expect(rank.ok ? null : rank.diagnostic.feature).toBe('pivot-engine-unmeasured')

    const zero = fixture([0, 0], [{ field: 1, fn: 'sum', showAs: 'percentOfTotal', baseField: -1, baseItem: 1048832, name: '' }])
    const divideByZero = computeValues(zero.store, zero.cache, zero.def, zero.view, zero.axes)
    expect(divideByZero.ok ? null : divideByZero.diagnostic.feature).toBe('pivot-engine-unmeasured')

    const error = fixture(['#N/A'], [{ field: 1, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: '' }])
    const errorColumn = error.store.columns[1]!
    if (errorColumn.kind !== 'inline') throw new Error('inline value field expected')
    errorColumn.tag[0] = VALUE_TAG.error
    const errorResult = computeValues(error.store, error.cache, error.def, error.view, error.axes)
    expect(errorResult.ok ? null : errorResult.diagnostic.feature).toBe('pivot-engine-unmeasured')

    const filtered = fixture([1, 2], [{ field: 1, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: '' }])
    filtered.axes.rowKey[1] = 0xffffffff
    const filteredResult = computeValues(filtered.store, filtered.cache, filtered.def, filtered.view, filtered.axes)
    expect(filteredResult.ok ? null : filteredResult.diagnostic.feature).toBe('pivot-engine-unmeasured')
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
      const raw = JSON.parse(readFileSync(MANIFEST, 'utf8')) as { cases?: Array<{ task: number; status?: string; file: string; pivot: string }> } | Array<{ task: number; status?: string; file: string; pivot: string }>
      const cases = Array.isArray(raw) ? raw : raw.cases ?? []
      for (const entry of cases.filter(item => item.task === 18 && item.status !== 'pending')) {
        const loaded = await loadAll(resolve(FIXTURES, entry.file))
        const p = loaded.find(candidate => candidate.def.name === entry.pivot)
        if (!p) throw new Error(`${entry.file}: no pivot named ${entry.pivot}`)
        const result = computePivot(p.store, p.cache, p.def, initialViewState(p.def, p.cache), optionsOf(p))
        expect(result.ok, `${entry.file} ${entry.pivot}: ${result.ok ? '' : result.diagnostic.message}`).toBe(true)
        if (result.ok) expectSavedValues(p, result.value)
      }
    })
  }
})
