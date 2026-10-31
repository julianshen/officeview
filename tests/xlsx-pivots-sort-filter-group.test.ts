// @vitest-environment node
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { computeAxes, NO_LINE } from '../src/xlsx/pivots/engine/axes'
import { groupColumn } from '../src/xlsx/pivots/engine/groups'
import { initialViewState } from '../src/xlsx/pivots/engine/view'
import type { FieldState } from '../src/xlsx/pivots/engine/view'
import type { PivotCacheDefinition, PivotFilter, PivotTableDefinition, PivotValue } from '../src/xlsx/pivots/model'
import { VALUE_TAG, type RecordStore } from '../src/xlsx/pivots/engine/records'
import { loadAll } from './helpers/xlsx-pivots-load'
import { checkSavedLayout, sheetCellReader } from '../src/xlsx/pivots/engine/faithful'

const field = (count: number): FieldState => ({
  hidden: Array(count).fill(false), collapsed: Array(count).fill(false), sortType: 'manual',
  defaultSubtotal: false, subtotals: [], subtotalTop: false, compact: false, outline: true,
  insertBlankRow: false, fillDownLabels: false, showAll: false,
})

function fixture(items: PivotValue[], recordItems: number[], measures: number[], filters: PivotFilter[] = []) {
  const cache = {
    source: { type: 'worksheet' }, olap: false,
    fields: [
      { databaseField: true, sharedItems: { inline: false, items } },
      { databaseField: true, sharedItems: { inline: true, items: [] } },
    ],
  } as unknown as PivotCacheDefinition
  const def = {
    rowGrandTotals: false, colGrandTotals: false, fields: [
      { items: items.map((_, x) => ({ type: 'data', x, hidden: false, showDetails: true })) }, {},
    ],
  } as unknown as PivotTableDefinition
  const store: RecordStore = {
    count: recordItems.length, bytes: 0,
    columns: [
      { kind: 'shared', index: Uint32Array.from(recordItems) },
      { kind: 'inline', tag: new Uint8Array(measures.length).fill(VALUE_TAG.number), num: Float64Array.from(measures), text: [], textIndex: new Uint32Array(measures.length) },
    ],
  }
  const view = {
    rows: [0], cols: [], pages: [], filters, data: [{ field: 1, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: 'Sum' }],
    fields: [field(items.length), field(0)],
  } as ReturnType<typeof initialViewState>
  return { cache, def, store, view }
}
const result = (f: ReturnType<typeof fixture>) => computeAxes(f.store, f.cache, f.def, f.view)
const lines = (value: ReturnType<typeof result>) => value.ok ? value.value.rows.filter(line => line.items.length === 1).map(line => line.items[0]) : []
const filter = (patch: Partial<PivotFilter>): PivotFilter => ({
  fld: 0, type: 'unknown', evalOrder: 0, id: 0, autoFilter: { columns: [] }, ...patch,
})

describe('Task 21 sort, filters and group mapping', () => {
  it('sorts homogeneous numeric labels ascending and descending, preserving manual order', () => {
    const base = fixture([{ type: 'number', value: 30 }, { type: 'number', value: 10 }, { type: 'number', value: 20 }], [0, 1, 2], [1, 1, 1])
    expect(lines(result(base))).toEqual([0, 1, 2])
    base.view.fields = [{ ...base.view.fields[0]!, sortType: 'ascending' }, base.view.fields[1]!]
    expect(lines(result(base))).toEqual([1, 2, 0])
    base.view.fields = [{ ...base.view.fields[0]!, sortType: 'descending' }, base.view.fields[1]!]
    expect(lines(result(base))).toEqual([0, 2, 1])
  })

  it('sorts by Values and preserves source order for tied ranks', () => {
    const p = fixture(['A', 'B', 'C'].map(value => ({ type: 'string', value })), [0, 0, 1, 1, 2], [2, 3, 9, 1, 4])
    p.view.fields = [{ ...p.view.fields[0]!, sortType: 'descending', autoSortScope: {
      type: 'normal', dataOnly: true, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false,
      outline: true, collapsedLevelsAreSubtotals: false, references: [{ field: -2, x: [0], selected: true, byPosition: false, relative: false, defaultSubtotal: false, subtotals: [] }],
    } }, p.view.fields[1]!]
    expect(lines(result(p))).toEqual([1, 0, 2])
    p.store = { ...p.store, columns: [p.store.columns[0]!, { kind: 'inline', tag: new Uint8Array(5).fill(VALUE_TAG.number), num: Float64Array.from([2, 3, 9, 1, 5]), text: [], textIndex: new Uint32Array(5) }] }
    expect(lines(result(p))).toEqual([1, 0, 2])
    const noReference = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0, 1], [1, 2])
    noReference.view.fields = [{ ...noReference.view.fields[0]!, sortType: 'ascending', autoSortScope: {
      type: 'normal', dataOnly: true, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false,
      outline: true, collapsedLevelsAreSubtotals: false, references: [],
    } }, noReference.view.fields[1]!]
    expect(result(noReference)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
  })

  it('top Items, Percent and Sum filters use their measured cutoffs', () => {
    const p = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0, 0, 0, 1], [1, 1, 1, 100], [filter({
      type: 'count', iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: false, val: 1 } }] },
    })])
    const items = result(p)
    expect(items.ok).toBe(true)
    if (items.ok) expect([...items.value.rowKey]).toEqual([NO_LINE, NO_LINE, NO_LINE, 0])
    for (const mode of ['percent', 'sum'] as const) {
      p.view.filters = [filter({ type: mode, iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: mode === 'percent', val: mode === 'percent' ? 50 : 75 } }] } })]
      const selected = result(p)
      expect(selected.ok).toBe(true)
      if (selected.ok) expect([...selected.value.rowKey]).toEqual([NO_LINE, NO_LINE, NO_LINE, 0])
    }
    p.view.filters = [filter({ type: 'count', iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true,
      top10: { top: true, percent: false, val: 1, filterVal: 99 } }] } })]
    expect(result(p).ok).toBe(true) // filterVal is descriptive; count selects by aggregate rank.
  })

  it('sorts mixed values and blanks, and applies Excel custom-list month order', () => {
    const mixed = fixture([{ type: 'number', value: 1 }, { type: 'string', value: 'A' }, { type: 'missing' }], [0, 1, 2], [1, 1, 1])
    mixed.view.fields = [{ ...mixed.view.fields[0]!, sortType: 'ascending' }, mixed.view.fields[1]!]
    expect(lines(result(mixed))).toEqual([0, 1, 2])
    mixed.view.fields = [{ ...mixed.view.fields[0]!, sortType: 'descending' }, mixed.view.fields[1]!]
    expect(lines(result(mixed))).toEqual([2, 1, 0])
    const months = fixture(['Mar', 'Jan', 'Feb'].map(value => ({ type: 'string', value })), [0, 1, 2], [1, 1, 1])
    months.view.fields = [{ ...months.view.fields[0]!, sortType: 'ascending' }, months.view.fields[1]!]
    expect(lines(result(months))).toEqual([1, 2, 0])
    months.def.customListSort = false
    expect(lines(result(months))).toEqual([2, 1, 0])
    const grouped = fixture([{ type: 'number', value: 1 }, { type: 'number', value: 11 }], [0, 1], [1, 1])
    grouped.cache.fields.push({ databaseField: false, fieldGroup: { base: 0, rangePr: { groupBy: 'range', groupInterval: 10, startNum: 1, endNum: 20 },
      groupItems: [{ type: 'string', value: '1-10' }, { type: 'string', value: '11-20' }] } } as never)
    grouped.def.fields.push({ items: [{ type: 'data', x: 0 }, { type: 'data', x: 1 }] } as never)
    grouped.view.rows = [2]
    grouped.view.fields = [...grouped.view.fields, { ...field(2), sortType: 'ascending' }]
    expect(result(grouped)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
  })

  it('applies case-insensitive wildcard label filtering', () => {
    const f = fixture(['Apple', 'apricot', 'Berry'].map(value => ({ type: 'string', value })), [0, 1, 2], [1, 1, 1], [filter({ type: 'captionBeginsWith', stringValue1: 'AP*' })])
    const axes = result(f)
    expect(axes.ok).toBe(true)
    if (axes.ok) expect([...axes.value.rowKey]).toEqual([0, 1, NO_LINE])
  })

  it('sorts measured Code punctuation and space pairs', () => {
    const sorted = fixture(['a-z', 'ab', '[x', '1x'].map(value => ({ type: 'string', value })), [0, 1, 2, 3], [1, 1, 1, 1])
    sorted.view.fields = [{ ...sorted.view.fields[0]!, sortType: 'ascending' }, sorted.view.fields[1]!]
    expect(result(sorted)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })

    const allowed = fixture(['b', 'A', '2', '1', 'a1'].map(value => ({ type: 'string', value })), [0, 1, 2, 3, 4], [1, 1, 1, 1, 1])
    allowed.view.fields = [{ ...allowed.view.fields[0]!, sortType: 'ascending' }, allowed.view.fields[1]!]
    expect(lines(result(allowed))).toEqual([3, 2, 1, 4, 0])

    const ambiguous = fixture(['A 1', 'A1'].map(value => ({ type: 'string', value })), [0, 1], [1, 1])
    ambiguous.view.fields = [{ ...ambiguous.view.fields[0]!, sortType: 'ascending' }, ambiguous.view.fields[1]!]
    expect(lines(result(ambiguous))).toEqual([0, 1])

    const separated = fixture(['North Europe', 'North Asia'].map(value => ({ type: 'string', value })), [0, 1], [1, 1])
    separated.view.fields = [{ ...separated.view.fields[0]!, sortType: 'ascending' }, separated.view.fields[1]!]
    expect(lines(result(separated))).toEqual([1, 0])

    const f = fixture(['Apple', 'apricot', 'Berry'].map(value => ({ type: 'string', value })), [0, 1, 2], [1, 1, 1], [filter({
      type: 'captionEqual', stringValue1: 'A?ple', autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true,
        customFilters: { and: false, filters: [{ operator: 'equal', val: 'A?ple' }, { operator: 'equal', val: 'berry' }] } }] },
    })])
    const axes = result(f)
    expect(axes.ok).toBe(true)
    if (axes.ok) expect([...axes.value.rowKey]).toEqual([0, NO_LINE, 1])
    f.view.filters = [filter({ type: 'captionEqual', stringValue1: 'A*', autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true,
      customFilters: { and: true, filters: [{ operator: 'greaterThanOrEqual', val: 'Apple' }, { operator: 'lessThan', val: 'Berry' }] } }] } })]
    const andAxes = result(f)
    expect(andAxes.ok).toBe(true)
    if (andAxes.ok) expect([...andAxes.value.rowKey]).toEqual([0, 1, NO_LINE])
    f.view.filters = [filter({ type: 'captionBetween', stringValue1: 'Apple', stringValue2: 'Citrus' })]
    const between = result(f)
    expect(between.ok).toBe(true)
    if (between.ok) expect([...between.value.rowKey]).toEqual([0, 1, 2])
    f.view.filters = [filter({ type: 'captionGreaterThan', stringValue1: 'a-z' })]
    expect(result(f).ok).toBe(true)
    const escaped = fixture([{ type: 'string', value: 'A*star' }, { type: 'string', value: 'Astar' }], [0, 1], [1, 1], [filter({
      type: 'captionEqual', stringValue1: 'A~*star', autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true,
        customFilters: { and: true, filters: [{ operator: 'equal', val: 'A~*star' }] } }] },
    })])
    const escapedAxes = result(escaped)
    expect(escapedAxes.ok).toBe(true)
    if (escapedAxes.ok) expect([...escapedAxes.value.rowKey]).toEqual([0, NO_LINE])
  })

  it('compares caption spaces using measured Excel ordering', () => {
    const ambiguous = fixture(['A 1', 'A1'].map(value => ({ type: 'string', value })), [0, 1], [1, 1], [filter({
      type: 'captionGreaterThan', stringValue1: 'A1',
    })])
    const spaceOrder = result(ambiguous)
    expect(spaceOrder.ok).toBe(true)
    if (spaceOrder.ok) expect([...spaceOrder.value.rowKey]).toEqual([NO_LINE, NO_LINE])

    const separated = fixture(['North Asia', 'North Europe'].map(value => ({ type: 'string', value })), [0, 1], [1, 1], [filter({
      type: 'captionLessThan', stringValue1: 'North Europe',
    })])
    const axes = result(separated)
    expect(axes.ok).toBe(true)
    if (axes.ok) expect([...axes.value.rowKey]).toEqual([0, NO_LINE])
  })

  it('filters numeric and renamed captions while dates remain outside measured label semantics', () => {
    for (const value of [{ type: 'number', value: 10 }, { type: 'missing' }] as PivotValue[]) {
      const p = fixture([value], [0], [1], [filter({ type: 'captionNotEqual', stringValue1: 'A' })])
      if (value.type === 'missing') expect(result(p)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
      else expect(result(p).ok).toBe(true)
    }
    const dateCaption = fixture([{ type: 'date', value: '2024-01-15T00:00:00' }], [0], [1], [filter({ type: 'captionNotEqual', stringValue1: 'A' })])
    expect(result(dateCaption)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
    const renamed = fixture([{ type: 'string', value: 'Apple' }], [0], [1], [filter({ type: 'captionEqual', stringValue1: 'Apple' })])
    renamed.def.fields[0] = { ...renamed.def.fields[0]!, items: [{ type: 'data', x: 0, name: 'Favorite' } as never] }
    const renamedResult = result(renamed)
    expect(renamedResult.ok).toBe(true)
    if (renamedResult.ok) expect([...renamedResult.value.rowKey]).toEqual([NO_LINE])
    const time = fixture([{ type: 'date', value: '2024-01-15T12:00:00' }], [0], [1], [filter({ type: 'dateEqual', stringValue1: '2024-01-15T00:00:00' })])
    expect(result(time)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
  })

  it('applies date filters and cross-field showAll visibility', () => {
    const date = fixture([
      { type: 'date', value: '2024-01-01T00:00:00' }, { type: 'date', value: '2024-02-01T00:00:00' },
    ], [0, 1], [1, 1], [filter({ type: 'dateNewerThanOrEqual', stringValue1: '2024-02-01T00:00:00' })])
    const dates = result(date)
    expect(dates.ok).toBe(true)
    if (dates.ok) expect([...dates.value.rowKey]).toEqual([NO_LINE, 0])

    const cross = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0, 1], [1, 1], [filter({ type: 'captionEqual', stringValue1: 'A' })])
    cross.cache.fields.push({ databaseField: true, sharedItems: { inline: false, items: [{ type: 'string', value: 'X' }, { type: 'string', value: 'Y' }] } } as never)
    cross.def.fields.push({ items: [{ type: 'data', x: 0 }, { type: 'data', x: 1 }] } as never)
    cross.store = { ...cross.store, columns: [...cross.store.columns, { kind: 'shared', index: Uint32Array.from([0, 1]) }] }
    cross.view.rows = [0]; cross.view.cols = [2]
    cross.view.fields = [{ ...cross.view.fields[0]!, showAll: true }, cross.view.fields[1]!, { ...field(2), showAll: true }]
    expect(result(cross).ok).toBe(true)
    cross.view.filters = [filter({ type: 'valueGreaterThan', iMeasureFld: 0, stringValue1: '0' })]
    expect(result(cross).ok).toBe(true)
  })

  it('applies value and top-Items filters independently within each parent group', () => {
    const p = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0, 0, 1, 1], [10, 1, 1, 20])
    p.cache.fields[0] = { databaseField: true, sharedItems: { inline: false, items: [{ type: 'string', value: 'X' }, { type: 'string', value: 'Y' }] } } as never
    p.cache.fields.push({ databaseField: true, sharedItems: { inline: false, items: [{ type: 'string', value: 'A' }, { type: 'string', value: 'B' }] } } as never)
    p.def.fields[0] = { items: [{ type: 'data', x: 0 } as never, { type: 'data', x: 1 } as never] } as never
    p.def.fields.push({ items: [{ type: 'data', x: 0 }, { type: 'data', x: 1 }] } as never)
    p.store = { ...p.store, columns: [
      { kind: 'shared', index: Uint32Array.from([0, 0, 1, 1]) },
      p.store.columns[1]!,
      { kind: 'shared', index: Uint32Array.from([0, 1, 0, 1]) },
    ] }
    p.view.rows = [0, 2]
    p.view.fields = [field(2), field(2), field(2)]
    for (const filterValue of [
      filter({ fld: 2, type: 'valueGreaterThan', iMeasureFld: 0, stringValue1: '5' }),
      filter({ fld: 2, type: 'count', iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: false, val: 1 } }] } }),
    ]) {
      p.view.filters = [filterValue]
      const nested = result(p)
      expect(nested, JSON.stringify(nested)).toMatchObject({ ok: true })
      if (nested.ok) expect([...nested.value.rowKey]).toEqual([1, NO_LINE, NO_LINE, 3])
    }
  })

  it('excludes same-field filtered items from showAll while retaining allowed no-data items', () => {
    const f = fixture(['Apple', 'Berry', 'Citrus'].map(value => ({ type: 'string', value })), [0], [1], [filter({ type: 'captionEqual', stringValue1: 'Apple' })])
    f.view.fields = [{ ...f.view.fields[0]!, showAll: true }, f.view.fields[1]!]
    const axes = result(f)
    expect(axes.ok).toBe(true)
    if (axes.ok) expect(axes.value.rows.map(line => line.items[0])).toEqual([0])
    for (const filterValue of [
      filter({ type: 'valueNotEqual', iMeasureFld: 0, stringValue1: '15' }),
      filter({ type: 'count', iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: false, val: 1 } }] } }),
    ]) {
      f.view.filters = [filterValue]
      expect(result(f)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
    }
  })

  it('refuses classic gridDropZones geometry without rejecting unrelated hidden sorted axes', () => {
    const classic = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0, 1], [1, 2])
    classic.def.gridDropZones = true
    classic.def.dataOnRows = false
    // One data field with row fields is measured (fix round B1, Task 15 layout-classic-single-data); other
    // shapes without a column field stay refused, e.g. two data fields (Values on rows).
    expect(result(classic).ok).toBe(true)
    classic.view.data = [classic.view.data[0]!, classic.view.data[0]!]
    classic.view.rows = [...classic.view.rows, -2]
    expect(result(classic)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured', message: 'classic pivot layout (gridDropZones) without a column field is measured for one data field with row fields only' } })

    const ordinary = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0, 0], [1, 2])
    ordinary.cache.fields.push({ databaseField: true, sharedItems: { inline: false, items: [{ type: 'string', value: 'X' }, { type: 'string', value: 'Y' }] } } as never)
    ordinary.def.fields.push({ items: [{ type: 'data', x: 0 }, { type: 'data', x: 1 }] } as never)
    ordinary.store = { ...ordinary.store, columns: [...ordinary.store.columns, { kind: 'shared', index: Uint32Array.from([0, 0]) }] }
    ordinary.view.rows = [0]
    ordinary.view.cols = [2]
    ordinary.view.fields = [
      { ...field(2), hidden: [false, true], sortType: 'ascending' },
      field(0),
      { ...field(2), hidden: [false, true] },
    ]
    expect(result(ordinary).ok).toBe(true)

    const gridWithoutRows = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0, 1], [1, 2])
    gridWithoutRows.def.gridDropZones = true
    gridWithoutRows.view.cols = [-2]
    expect(result(gridWithoutRows).ok).toBe(true)
  })

  it('applies value filters using the selected data field aggregate', () => {
    const f = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0, 0, 1, 1], [2, 3, 9, 11], [filter({ type: 'valueGreaterThan', iMeasureFld: 0, stringValue1: '10' })])
    const axes = result(f)
    expect(axes.ok).toBe(true)
    if (axes.ok) expect([...axes.value.rowKey]).toEqual([NO_LINE, NO_LINE, 0, 0])
  })

  it('includes top-10 cutoff ties and retains dynamic-date refusal', () => {
    const tied = fixture(['A', 'B', 'C'].map(value => ({ type: 'string', value })), [0, 1, 2], [10, 10, 1], [filter({
      type: 'count', iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: false, val: 1 } }] },
    })])
    expect(lines(result(tied))).toEqual([0, 1])
    const unique = fixture(['A', 'B', 'C'].map(value => ({ type: 'string', value })), [0, 1, 2], [10, 5, 1], [filter({
      type: 'percent', iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: true, val: 34 } }] },
    })])
    expect(lines(result(unique))).toEqual([0])
    const dynamic = fixture([{ type: 'date', value: '2026-10-30T00:00:00' }], [0], [1], [filter({ type: 'today' })])
    expect(result(dynamic)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-filter-dynamic-date' } })
  })

  it('maps localized date group captions and out-of-range buckets with blanks', () => {
    const cache = {
      source: { type: 'worksheet' }, olap: false,
      fields: [
        { databaseField: true, sharedItems: { inline: false, items: [
          { type: 'date', value: '2020-06-01T00:00:00' }, { type: 'date', value: '2021-03-02T00:00:00' }, { type: 'missing' },
        ] } },
        { databaseField: false, fieldGroup: { base: 0, rangePr: { autoStart: false, autoEnd: false, groupBy: 'years', groupInterval: 1,
          startDate: '2020-01-01T00:00:00', endDate: '2021-12-31T00:00:00' }, groupItems: [
      { type: 'string', value: '(blank)' }, { type: 'string', value: '2020年' }, { type: 'string', value: '2021年' },
        ] } },
      ],
    } as unknown as PivotCacheDefinition
    const def = { rowGrandTotals: false, colGrandTotals: false, fields: [{ items: [0, 1, 2].map(x => ({ type: 'data', x })) },
      { items: [0, 1, 2].map(x => ({ type: 'data', x })) }] } as unknown as PivotTableDefinition
    const store: RecordStore = { count: 4, bytes: 0, columns: [{ kind: 'shared', index: Uint32Array.from([0, 1, 2, 0]) }] }
    expect(groupColumn(store, cache, 1)).toMatchObject({ ok: true, value: Uint32Array.from([1, 2, 0, 1]) })
    const view = { rows: [1, 0], cols: [], pages: [], filters: [], data: [], fields: [field(3), field(3)] } as ReturnType<typeof initialViewState>
    const axes = computeAxes(store, cache, def, view)
    expect(axes.ok).toBe(true)
    if (axes.ok) expect(axes.value.rows.map(line => line.items)).toEqual([[0], [0, 2], [1], [1, 0], [2], [2, 1]])
    const bounded = { ...cache, fields: [{ ...cache.fields[0]!, sharedItems: { inline: false, items: [
      { type: 'date', value: '2019-01-01T00:00:00' }, { type: 'date', value: '2022-01-01T00:00:00' },
    ] } }, { ...cache.fields[1]!, fieldGroup: { ...cache.fields[1]!.fieldGroup!, groupItems: [
      { type: 'string', value: '<2020/1/1' }, { type: 'string', value: '2020年' }, { type: 'string', value: '2021年' }, { type: 'string', value: '>2021/12/31' },
    ] } }] } as unknown as PivotCacheDefinition
    expect(groupColumn({ ...store, count: 2, columns: [{ kind: 'shared', index: Uint32Array.of(0, 1) }] }, bounded, 1))
      .toMatchObject({ ok: true, value: Uint32Array.from([0, 3]) })
  })

  it('maps discrete groups through their base item indices', () => {
    const cache = { fields: [
      { databaseField: true, sharedItems: { inline: false, items: [{ type: 'string', value: 'A' }, { type: 'string', value: 'B' }, { type: 'string', value: 'C' }] } },
      { databaseField: false, fieldGroup: { base: 0, discretePr: [1, 0, 1], groupItems: [{ type: 'string', value: 'Other' }, { type: 'string', value: 'Selected' }] } },
    ] } as unknown as PivotCacheDefinition
    const store: RecordStore = { count: 3, bytes: 0, columns: [{ kind: 'shared', index: Uint32Array.from([0, 1, 2]) }] }
    expect(groupColumn(store, cache, 1)).toMatchObject({ ok: true, value: Uint32Array.from([1, 0, 1]) })
    const withBlank = { fields: [
      { databaseField: true, sharedItems: { inline: false, items: [{ type: 'string', value: 'A' }, { type: 'missing' }, { type: 'string', value: 'B' }] } },
      { databaseField: false, fieldGroup: { base: 0, discretePr: [0, 1, 0], groupItems: [{ type: 'string', value: 'Other' }, { type: 'string', value: 'No value' }] } },
    ] } as unknown as PivotCacheDefinition
    const blankStore: RecordStore = { count: 3, bytes: 0, columns: [{ kind: 'shared', index: Uint32Array.from([0, 1, 2]) }] }
    expect(groupColumn(blankStore, withBlank, 1)).toMatchObject({ ok: true, value: Uint32Array.from([0, 1, 0]) })
  })

  it('maps fractional numeric ranges and out-of-range values', () => {
    const cache = { fields: [
      { databaseField: true, sharedItems: { inline: false, items: [-1, 0.5, 0.2, 2, 3].map(value => ({ type: 'number', value })) } },
      { databaseField: false, fieldGroup: { base: 0, rangePr: { autoStart: false, autoEnd: false, groupBy: 'range', groupInterval: 0.5, startNum: 0, endNum: 2 },
        groupItems: ['<0', '0-0.5', '0.5-1', '1-1.5', '1.5-2', '>2'].map(value => ({ type: 'string', value })) } },
    ] } as unknown as PivotCacheDefinition
    const store: RecordStore = { count: 5, bytes: 0, columns: [{ kind: 'shared', index: Uint32Array.from([0, 1, 2, 3, 4]) }] }
    expect(groupColumn(store, cache, 1)).toMatchObject({ ok: true, value: Uint32Array.from([0, 2, 1, 4, 5]) })

    const integerRange = { ...cache, fields: [cache.fields[0]!, { ...cache.fields[1]!, fieldGroup: {
      ...cache.fields[1]!.fieldGroup!, rangePr: { ...cache.fields[1]!.fieldGroup!.rangePr!, groupInterval: 1 },
    } }] } as unknown as PivotCacheDefinition
    expect(groupColumn({ ...store, count: 1, columns: [{ kind: 'shared', index: Uint32Array.of(0) }] }, integerRange, 1))
      .toMatchObject({ ok: true, value: Uint32Array.from([0]) })
    expect(groupColumn({ ...store, count: 1, columns: [{ kind: 'shared', index: Uint32Array.of(4) }] }, integerRange, 1))
      .toMatchObject({ ok: true, value: Uint32Array.from([5]) })
  })

  it('maps decimal values on exact interval boundaries without a floating-point off-by-one', () => {
    const values = Array.from({ length: 11 }, (_, index) => index / 10)
    const cache = { fields: [
      { databaseField: true, sharedItems: { inline: false, items: [{ type: 'number', value: 0.3 }, { type: 'number', value: 0.6 }, { type: 'number', value: 0.7 }] } },
      { databaseField: false, fieldGroup: { base: 0, rangePr: { groupBy: 'range', groupInterval: 0.1, startNum: 0, endNum: 1 },
        groupItems: values.slice(0, -1).map((start, index) => ({ type: 'string', value: `${start.toFixed(1)}-${values[index + 1]!.toFixed(1)}` })) } },
    ] } as unknown as PivotCacheDefinition
    const store: RecordStore = { count: 3, bytes: 0, columns: [{ kind: 'shared', index: Uint32Array.from([0, 1, 2]) }] }
    expect(groupColumn(store, cache, 1)).toMatchObject({ ok: true, value: Uint32Array.from([3, 6, 7]) })

    const fractionalStart = { ...cache, fields: [cache.fields[0]!, { ...cache.fields[1]!, fieldGroup: {
      ...cache.fields[1]!.fieldGroup!, rangePr: { ...cache.fields[1]!.fieldGroup!.rangePr!, groupInterval: 1, startNum: 0.5, endNum: 1.5 },
      groupItems: ['<0.5', '0.5-1.5', '>1.5'].map(value => ({ type: 'string', value })),
    } }] } as unknown as PivotCacheDefinition
    expect(groupColumn({ ...store, count: 1, columns: [{ kind: 'shared', index: Uint32Array.of(0) }] }, fractionalStart, 1))
      .toMatchObject({ ok: true, value: Uint32Array.from([0]) })
  })
})

describe('Task 21 Excel manifest cases', () => {
  const fixtures = resolve(process.cwd(), 'tests/fixtures/xlsx-pivots')
  const manifestPath = resolve(fixtures, 'engine/manifest.json')
  if (!existsSync(manifestPath)) {
    it.skip('SKIPPED: tests/fixtures/xlsx-pivots/engine/manifest.json is absent; Task 21 Excel oracle cases not run', () => {})
    return
  }
  const raw = JSON.parse(readFileSync(manifestPath, 'utf8')) as { cases?: Array<{ task: number; status?: string; case: string; file: string; pivot: string }> } | Array<{ task: number; status?: string; case: string; file: string; pivot: string }>
  const cases = (Array.isArray(raw) ? raw : raw.cases ?? []).filter(entry => entry.task === 21 && entry.status !== 'pending')
  const exactRefusals: Record<string, string> = {
    'sort-error-item': 'error-valued pivot item field 0 is saved by Excel as text but place.ts currently emits a worksheet error cell',
    'sort-value-showall-no-data': "showAll value sort for field 1 has multiple no-data items under parent 4, whose tie order is not measured",
  }
  if (!cases.length) it.skip('SKIPPED: Task 21 has no runnable Excel manifest cases', () => {})
  else it.each(cases.map(entry => [entry.case, entry] as const))('manifest %s', async (_name, entry) => {
    const local = resolve(fixtures, 'engine', entry.file)
    const loaded = await loadAll(existsSync(local) ? local : resolve(process.cwd(), entry.file))
    const pivot = loaded.find(candidate => candidate.def.name === entry.pivot)
    if (!pivot) throw new Error(`${entry.case}: pivot ${entry.pivot} not found`)
    const faithful = checkSavedLayout(pivot.store, pivot.cache, pivot.def, sheetCellReader(pivot.sheet), { date1904: pivot.date1904 })
    const expectedRefusal = exactRefusals[entry.case]
    if (expectedRefusal) {
      expect(faithful.status, entry.case).toBe('unverified')
      const axes = computeAxes(pivot.store, pivot.cache, pivot.def, initialViewState(pivot.def, pivot.cache))
      expect(axes.ok ? '' : axes.diagnostic.message, entry.case).toBe(expectedRefusal)
    } else expect(faithful.status, entry.case).toBe('ok')
  })
})

describe('Task 21 oracle fix round 1 counterexamples', () => {
  it('uses the oracle top-percent share threshold from filter-top-percent-30', async () => {
    const file = resolve(process.cwd(), 'tests/fixtures/xlsx-pivots/engine/filter-top-percent-30.xlsx')
    const p = (await loadAll(file))[0]!
    const computed = computeAxes(p.store, p.cache, p.def, initialViewState(p.def, p.cache))
    expect(computed.ok, computed.ok ? '' : computed.diagnostic.message).toBe(true)
    if (!computed.ok) return
    const items = p.def.fields[0]!.items!
    const shared = p.cache.fields[0]!.sharedItems!.items
    const labels = computed.value.rows.filter(line => line.type === 'data' && line.items.length === 1)
      .map(line => {
        const item = shared[items[line.items[0]!]!.x!]!
        return item.type === 'string' ? item.value : undefined
      })
    expect(labels).toEqual(['South', 'West'])
  })

  it('uses Excel apostrophe and hyphen collation, position-independent measured symbols, and refuses unmeasured positions', () => {
    for (const [input, expected] of [
      [["O'Brien", 'Olsen'], [0, 1]], [['T-shirt', 'Tu'], [0, 1]], [['-a', 'b'], [0, 1]],
      [['ab', "a'b", 'a-b'], [0, 1, 2]],
    ] as const) {
      const p = fixture(input.map(value => ({ type: 'string', value })), input.map((_, i) => i), input.map(() => 1))
      p.view.fields = [{ ...p.view.fields[0]!, sortType: 'ascending' }, p.view.fields[1]!]
      expect(lines(result(p))).toEqual(expected)
    }
    const interiorSymbol = fixture(['x#1', 'xa'].map(value => ({ type: 'string', value })), [0, 1], [1, 1])
    interiorSymbol.view.fields = [{ ...interiorSymbol.view.fields[0]!, sortType: 'ascending' }, interiorSymbol.view.fields[1]!]
    expect(result(interiorSymbol)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
  })

  it('refuses sorted group fields and unmeasured custom-list mixtures', () => {
    const p = fixture([{ type: 'number', value: 1 }, { type: 'number', value: 11 }], [0, 1], [1, 1])
    p.cache.fields.push({ databaseField: false, fieldGroup: { base: 0, rangePr: { groupBy: 'range', groupInterval: 10, startNum: 1, endNum: 20 },
      groupItems: [{ type: 'string', value: '1-10' }, { type: 'string', value: '11-20' }] } } as never)
    p.def.fields.push({ items: [{ type: 'data', x: 0 }, { type: 'data', x: 1 }] } as never)
    p.view.rows = [2]
    p.view.fields = [...p.view.fields, { ...field(2), sortType: 'ascending' }]
    expect(result(p)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })

    const mixedList = fixture(['Jan', 'Other'].map(value => ({ type: 'string', value })), [0, 1], [1, 1])
    mixedList.view.fields = [{ ...mixedList.view.fields[0]!, sortType: 'ascending' }, mixedList.view.fields[1]!]
    expect(result(mixedList)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
  })

  it('places a fractional value in the Excel-saved final bucket beyond endNum', () => {
    const labels = ['<0.5', ...Array.from({ length: 40 }, (_, index) => {
      const start = 0.5 + index * 2.5, end = start + 2.5
      return `${Number(start.toFixed(1))}-${Number(end.toFixed(1))}`
    }), '>100.5']
    const cache = { fields: [
      { databaseField: true, sharedItems: { inline: true, items: [] } },
      { databaseField: false, fieldGroup: { base: 0, rangePr: { groupBy: 'range', groupInterval: 2.5, startNum: 0.5, endNum: 100 },
        groupItems: labels.map(value => ({ type: 'string', value })) } },
    ] } as unknown as PivotCacheDefinition
    const store: RecordStore = { count: 1, bytes: 0, columns: [
      { kind: 'inline', tag: Uint8Array.of(VALUE_TAG.number), num: Float64Array.of(100.25), text: [], textIndex: Uint32Array.of(0) },
    ] }
    expect(groupColumn(store, cache, 1)).toMatchObject({ ok: true, value: Uint32Array.of(40) })
  })

  it('refuses booleans and mixed date sorts', () => {
    for (const items of [
      [{ type: 'boolean', value: true }, { type: 'string', value: 'A' }],
      [{ type: 'date', value: '2024-01-01T00:00:00' }, { type: 'number', value: 1 }],
    ] as PivotValue[][]) {
      const p = fixture(items, [0, 1], [1, 1])
      p.view.fields = [{ ...p.view.fields[0]!, sortType: 'ascending' }, p.view.fields[1]!]
      expect(result(p)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
    }
  })

  it('refuses unmeasured showAll value-sort empty-item placements', () => {
    const nested = fixture(['A', 'B', 'C', 'D'].map(value => ({ type: 'string', value })), [0, 1, 2, 3], [10, 5, 1, 1])
    nested.cache.fields[0] = { databaseField: true, sharedItems: { inline: false, items: ['A', 'B', 'C', 'D'].map(value => ({ type: 'string', value })) } } as never
    nested.def.fields[0] = { items: ['A', 'B', 'C', 'D'].map((_, x) => ({ type: 'data', x })) } as never
    nested.cache.fields.push({ databaseField: true, sharedItems: { inline: false, items: [{ type: 'string', value: 'P' }] } } as never)
    nested.def.fields.push({ items: [{ type: 'data', x: 0 }] } as never)
    nested.store = { ...nested.store, columns: [nested.store.columns[0]!, nested.store.columns[1]!, { kind: 'shared', index: Uint32Array.from([0, 0, 1, 1]) }] }
    nested.view.rows = [2, 0]
    nested.view.fields = [ { ...field(4), showAll: true, sortType: 'descending', autoSortScope: { type: 'normal', dataOnly: true, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false, outline: true, collapsedLevelsAreSubtotals: false, references: [{ field: -2, x: [0], selected: true, byPosition: false, relative: false, defaultSubtotal: false, subtotals: [] }] } }, nested.view.fields[1]!, { ...field(1) } ]
    expect(result(nested)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })

    const ascending = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0], [10])
    ascending.view.fields = [{ ...ascending.view.fields[0]!, showAll: true, sortType: 'ascending', autoSortScope: { type: 'normal', dataOnly: true, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false, outline: true, collapsedLevelsAreSubtotals: false, references: [{ field: -2, x: [0], selected: true, byPosition: false, relative: false, defaultSubtotal: false, subtotals: [] }] } }, ascending.view.fields[1]!]
    expect(result(ascending)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })

    const noRecordsParent = fixture(['A', 'B', 'C'].map(value => ({ type: 'string', value })), [0, 1], [10, 5])
    noRecordsParent.cache.fields.push({ databaseField: true, sharedItems: { inline: false, items: [{ type: 'string', value: 'P' }, { type: 'string', value: 'Q' }] } } as never)
    noRecordsParent.def.fields.push({ items: [{ type: 'data', x: 0 }, { type: 'data', x: 1 }] } as never)
    noRecordsParent.store = { ...noRecordsParent.store, columns: [noRecordsParent.store.columns[0]!, noRecordsParent.store.columns[1]!, { kind: 'shared', index: Uint32Array.of(0, 0) }] }
    noRecordsParent.view.rows = [2, 0]
    noRecordsParent.view.fields = [
      { ...field(3), showAll: true, sortType: 'descending', autoSortScope: { type: 'normal', dataOnly: true, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false, outline: true, collapsedLevelsAreSubtotals: false, references: [{ field: -2, x: [0], selected: true, byPosition: false, relative: false, defaultSubtotal: false, subtotals: [] }] } },
      noRecordsParent.view.fields[1]!, { ...field(2), showAll: true },
    ]
    expect(result(noRecordsParent)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
  })

  it('refuses value filters, top-10 and value sorts when records have a blank aggregate', () => {
    const withBlankSum = (filters: PivotFilter[] = []) => {
      const p = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0, 1], [1, 1], filters)
      p.store = { ...p.store, columns: [p.store.columns[0]!, {
        kind: 'inline', tag: Uint8Array.of(VALUE_TAG.missing, VALUE_TAG.number), num: Float64Array.of(0, 10), text: [], textIndex: Uint32Array.of(0, 0),
      }] }
      return p
    }
    const valueFilter = withBlankSum([filter({ type: 'valueNotEqual', iMeasureFld: 0, stringValue1: '5' })])
    expect(result(valueFilter)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured', message: expect.stringContaining('blank aggregate') } })

    const top10 = withBlankSum([filter({ type: 'percent', iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true,
      top10: { top: true, percent: true, val: 50 } }] } })])
    expect(result(top10)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured', message: expect.stringContaining('blank aggregate') } })

    const valueSort = withBlankSum()
    valueSort.view.fields = [{ ...valueSort.view.fields[0]!, sortType: 'descending', autoSortScope: {
      type: 'normal', dataOnly: true, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false,
      outline: true, collapsedLevelsAreSubtotals: false, references: [{ field: -2, x: [0], selected: true, byPosition: false, relative: false, defaultSubtotal: false, subtotals: [] }],
    } }, valueSort.view.fields[1]!]
    expect(result(valueSort)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured', message: expect.stringContaining('blank aggregate') } })
  })

  it('handles Values before and between filtered or sorted row fields without throwing', () => {
    const before = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0, 1], [1, 5], [filter({ type: 'valueGreaterThan', iMeasureFld: 0, stringValue1: '2' })])
    before.view.rows = [-2, 0]
    expect(() => result(before)).not.toThrow()
    expect(result(before)).toMatchObject({ ok: true })
    before.view.filters = []
    before.view.fields = [{ ...before.view.fields[0]!, sortType: 'descending', autoSortScope: { type: 'normal', dataOnly: true, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false, outline: true, collapsedLevelsAreSubtotals: false, references: [{ field: -2, x: [0], selected: true, byPosition: false, relative: false, defaultSubtotal: false, subtotals: [] }] } }, before.view.fields[1]!]
    expect(() => result(before)).not.toThrow()
    expect(result(before)).toMatchObject({ ok: true })

    const between = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0, 1], [1, 5], [filter({ type: 'valueGreaterThan', iMeasureFld: 0, stringValue1: '2' })])
    between.cache.fields.push({ databaseField: true, sharedItems: { inline: false, items: [{ type: 'string', value: 'P' }, { type: 'string', value: 'Q' }] } } as never)
    between.def.fields.push({ items: [{ type: 'data', x: 0 }, { type: 'data', x: 1 }] } as never)
    between.store = { ...between.store, columns: [between.store.columns[0]!, between.store.columns[1]!, { kind: 'shared', index: Uint32Array.of(0, 1) }] }
    between.view.rows = [2, -2, 0]
    between.view.fields = [between.view.fields[0]!, between.view.fields[1]!, field(2)]
    expect(() => result(between)).not.toThrow()
    expect(result(between)).toMatchObject({ ok: true })
    between.view.filters = []
    between.view.fields = [{ ...between.view.fields[0]!, sortType: 'descending', autoSortScope: { type: 'normal', dataOnly: true, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false, outline: true, collapsedLevelsAreSubtotals: false, references: [{ field: -2, x: [0], selected: true, byPosition: false, relative: false, defaultSubtotal: false, subtotals: [] }] } }, between.view.fields[1]!, between.view.fields[2]!]
    expect(() => result(between)).not.toThrow()
    expect(result(between)).toMatchObject({ ok: true })
  })

  it('refuses blank negated label filters, number-group blanks and 1900 phantom-era date groups', () => {
    const blankFilter = fixture([{ type: 'missing' }], [0], [1], [filter({ type: 'captionNotEqual', stringValue1: 'A' })])
    expect(result(blankFilter)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })

    const numberGroup = { fields: [
      { databaseField: true, sharedItems: { inline: false, items: [{ type: 'missing' }] } },
      { databaseField: false, fieldGroup: { base: 0, rangePr: { groupBy: 'range', groupInterval: 1, startNum: 0, endNum: 2 }, groupItems: [{ type: 'string', value: '<0' }, { type: 'string', value: '0-1' }, { type: 'string', value: '1-2' }, { type: 'string', value: '>2' }] } },
    ] } as unknown as PivotCacheDefinition
    const blankStore: RecordStore = { count: 1, bytes: 0, columns: [{ kind: 'shared', index: Uint32Array.of(0) }] }
    expect(groupColumn(blankStore, numberGroup, 1)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })

    const oldDate = { fields: [
      { databaseField: true, sharedItems: { inline: false, items: [{ type: 'date', value: '1900-02-01T00:00:00' }] } },
      { databaseField: false, fieldGroup: { base: 0, rangePr: { groupBy: 'months', startDate: '1900-01-01T00:00:00', endDate: '1900-12-31T00:00:00' }, groupItems: [{ type: 'string', value: '1月' }, { type: 'string', value: '2月' }] } },
    ] } as unknown as PivotCacheDefinition
    expect(groupColumn(blankStore, oldDate, 1)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })

    const boundaryTime = { fields: [
      { databaseField: true, sharedItems: { inline: false, items: [{ type: 'date', value: '2020-01-01T15:00:00' }] } },
      { databaseField: false, fieldGroup: { base: 0, rangePr: { groupBy: 'years', startDate: '2020-01-01T00:00:00', endDate: '2020-12-31T00:00:00' },
        groupItems: [{ type: 'string', value: '2020年' }] } },
    ] } as unknown as PivotCacheDefinition
    expect(groupColumn(blankStore, boundaryTime, 1)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
  })
})

describe('Task 21 oracle fix round 2 counterexamples', () => {
  const nestedFixture = (filterValue?: PivotFilter, valueSort = false) => {
    const p = fixture(['A', 'B', 'C'].map(value => ({ type: 'string', value })), [0, 0, 0, 1, 1, 1], [100, 2, 8, 1, 30, 4])
    p.cache.fields[0] = { databaseField: true, sharedItems: { inline: false, items: ['X', 'Y'].map(value => ({ type: 'string', value })) } } as never
    p.cache.fields.push({ databaseField: true, sharedItems: { inline: false, items: ['P1', 'P2', 'P3'].map(value => ({ type: 'string', value })) } } as never)
    p.def.fields[0] = { items: ['X', 'Y'].map((_, x) => ({ type: 'data', x })) } as never
    p.def.fields.push({ items: ['P1', 'P2', 'P3'].map((_, x) => ({ type: 'data', x })) } as never)
    p.store = { ...p.store, count: 6, columns: [p.store.columns[0]!, p.store.columns[1]!, { kind: 'shared', index: Uint32Array.from([0, 1, 2, 0, 1, 2]) }] }
    p.view.rows = [0, 2]
    p.view.fields = [field(2), field(0), { ...field(2), hidden: [true, false] }]
    p.view.filters = filterValue ? [filterValue] : []
    if (valueSort) p.view.fields = [p.view.fields[0]!, p.view.fields[1]!, { ...p.view.fields[2]!, sortType: 'descending', autoSortScope: {
      type: 'normal', dataOnly: true, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false,
      outline: true, collapsedLevelsAreSubtotals: false,
      references: [{ field: -2, x: [0], selected: true, byPosition: false, relative: false, defaultSubtotal: false, subtotals: [] }],
    } }]
    return p
  }

  it('builds nested filter and sort item presence only from records still eligible', () => {
    const scenarios: Array<[string, ReturnType<typeof nestedFixture>, number[], number[]]> = [
      ['value-filter', nestedFixture(filter({ fld: 2, type: 'valueGreaterThan', iMeasureFld: 0, stringValue1: '5' })), [NO_LINE, NO_LINE, 1, NO_LINE, 3, NO_LINE], [2, 1]],
      ['top-items-filter', nestedFixture(filter({
        fld: 2, type: 'count', iMeasureFld: 0,
        autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: false, val: 1 } }] },
      })), [NO_LINE, NO_LINE, 1, NO_LINE, 3, NO_LINE], [2, 1]],
      ['value-sort', nestedFixture(undefined, true), [NO_LINE, 2, 1, NO_LINE, 4, 5], [2, 1, 1, 2]],
    ]
    for (const [name, p, expectedKeys, expectedChildren] of scenarios) {
      const computed = result(p)
      expect(computed.ok, `${name}: ${computed.ok ? '' : computed.diagnostic.message}`).toBe(true)
      if (computed.ok) {
        expect([...computed.value.rowKey], `${name} record selection/order`).toEqual(expectedKeys)
        expect(computed.value.rows.filter(line => line.items.length === 2).map(line => line.items[1]), `${name} nested order`).toEqual(expectedChildren)
      }
    }
  })

  it('refuses unmeasured equal-weight label orderings', () => {
    for (const labels of [[' x', '#x'], [' b', '#a'], ['a-bc', 'ab-c'], ["a'bc", "ab'c"], ["a'b-c", "a-b'c"]]) {
      const p = fixture(labels.map(value => ({ type: 'string', value })), [0, 1], [1, 1])
      p.view.fields = [{ ...p.view.fields[0]!, sortType: 'ascending' }, p.view.fields[1]!]
      expect(result(p)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
    }
    const measuredPunctuation = fixture(['ab', "a'b", 'a-b'].map(value => ({ type: 'string', value })), [0, 1, 2], [1, 1, 1])
    measuredPunctuation.view.fields = [{ ...measuredPunctuation.view.fields[0]!, sortType: 'ascending' }, measuredPunctuation.view.fields[1]!]
    expect(lines(result(measuredPunctuation))).toEqual([0, 1, 2])
  })

  it('refuses bottom-N items, percent and sum filters', () => {
    for (const type of ['count', 'percent', 'sum'] as const) {
      const p = fixture(['A', 'B', 'C'].map(value => ({ type: 'string', value })), [0, 1, 2], [10, 5, 1], [filter({
        type, iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true,
          top10: { top: false, percent: type === 'percent', val: type === 'percent' ? 50 : 1 } }] },
      })])
      expect(result(p), type).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
    }
  })

  it('refuses nested Top Percent until the per-parent cutoff is measured', () => {
    const p = nestedFixture(filter({ fld: 2, type: 'percent', iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true,
      top10: { top: true, percent: true, val: 50 } }] } }))
    p.view.fields = [p.view.fields[0]!, p.view.fields[1]!, { ...p.view.fields[2]!, hidden: [false, false] }]
    expect(result(p)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
  })
})
