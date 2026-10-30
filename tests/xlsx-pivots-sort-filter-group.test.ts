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

  it('sorts by the Values reference in autoSortScope and refuses tied ranks', () => {
    const p = fixture(['A', 'B', 'C'].map(value => ({ type: 'string', value })), [0, 0, 1, 1, 2], [2, 3, 9, 1, 4])
    p.view.fields = [{ ...p.view.fields[0]!, sortType: 'descending', autoSortScope: {
      type: 'normal', dataOnly: true, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false,
      outline: true, collapsedLevelsAreSubtotals: false, references: [{ field: -2, x: [0], selected: true, byPosition: false, relative: false, defaultSubtotal: false, subtotals: [] }],
    } }, p.view.fields[1]!]
    expect(lines(result(p))).toEqual([1, 0, 2])
    p.store = { ...p.store, columns: [p.store.columns[0]!, { kind: 'inline', tag: new Uint8Array(5).fill(VALUE_TAG.number), num: Float64Array.from([2, 3, 9, 1, 5]), text: [], textIndex: new Uint32Array(5) }] }
    expect(result(p)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
    const noReference = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0, 1], [1, 2])
    noReference.view.fields = [{ ...noReference.view.fields[0]!, sortType: 'ascending', autoSortScope: {
      type: 'normal', dataOnly: true, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false,
      outline: true, collapsedLevelsAreSubtotals: false, references: [],
    } }, noReference.view.fields[1]!]
    expect(result(noReference)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
  })

  it('top Items ranks by the selected measure aggregate, while Percent and Sum modes refuse', () => {
    const p = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0, 0, 0, 1], [1, 1, 1, 100], [filter({
      type: 'count', iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: false, val: 1 } }] },
    })])
    const items = result(p)
    expect(items.ok).toBe(true)
    if (items.ok) expect([...items.value.rowKey]).toEqual([NO_LINE, NO_LINE, NO_LINE, 0])
    for (const mode of ['percent', 'sum'] as const) {
      p.view.filters = [filter({ type: mode, iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: false, val: 50 } }] } })]
      expect(result(p)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
    }
    p.view.filters = [filter({ type: 'count', iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true,
      top10: { top: true, percent: false, val: 1, filterVal: 99 } }] } })]
    expect(result(p)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
  })

  it('refuses unmeasured mixed-type and built-in custom-list sort order', () => {
    const mixed = fixture([{ type: 'number', value: 1 }, { type: 'string', value: 'A' }, { type: 'missing' }], [0, 1, 2], [1, 1, 1])
    mixed.view.fields = [{ ...mixed.view.fields[0]!, sortType: 'ascending' }, mixed.view.fields[1]!]
    expect(result(mixed)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
    const months = fixture(['Jan', 'Feb', 'Mar'].map(value => ({ type: 'string', value })), [0, 1, 2], [1, 1, 1])
    months.view.fields = [{ ...months.view.fields[0]!, sortType: 'ascending' }, months.view.fields[1]!]
    expect(result(months)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
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

  it('uses English word collation for ASCII punctuation and custom label filters', () => {
    const sorted = fixture(['a-z', 'ab', '[x', '1x'].map(value => ({ type: 'string', value })), [0, 1, 2, 3], [1, 1, 1, 1])
    sorted.view.fields = [{ ...sorted.view.fields[0]!, sortType: 'ascending' }, sorted.view.fields[1]!]
    expect(lines(result(sorted))).toEqual([2, 3, 1, 0])

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
    const escaped = fixture([{ type: 'string', value: 'A*star' }, { type: 'string', value: 'Astar' }], [0, 1], [1, 1], [filter({
      type: 'captionEqual', stringValue1: 'A~*star', autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true,
        customFilters: { and: true, filters: [{ operator: 'equal', val: 'A~*star' }] } }] },
    })])
    const escapedAxes = result(escaped)
    expect(escapedAxes.ok).toBe(true)
    if (escapedAxes.ok) expect([...escapedAxes.value.rowKey]).toEqual([0, NO_LINE])
  })

  it('refuses label filters on unmeasured captions and filters renamed items by no guessed caption', () => {
    for (const value of [{ type: 'number', value: 10 }, { type: 'date', value: '2024-01-15T00:00:00' }, { type: 'missing' }] as PivotValue[]) {
      const p = fixture([value], [0], [1], [filter({ type: 'captionNotEqual', stringValue1: 'A' })])
      expect(result(p)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
    }
    const renamed = fixture([{ type: 'string', value: 'Apple' }], [0], [1], [filter({ type: 'captionEqual', stringValue1: 'Apple' })])
    renamed.def.fields[0] = { ...renamed.def.fields[0]!, items: [{ type: 'data', x: 0, name: 'Favorite' } as never] }
    expect(result(renamed)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
    const time = fixture([{ type: 'date', value: '2024-01-15T12:00:00' }], [0], [1], [filter({ type: 'dateEqual', stringValue1: '2024-01-15T00:00:00' })])
    expect(result(time)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
  })

  it('applies a fixed date filter and refuses unmeasured cross-field no-data visibility', () => {
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
    expect(result(cross)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
    cross.view.filters = [filter({ type: 'valueGreaterThan', iMeasureFld: 0, stringValue1: '0' })]
    expect(result(cross)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
  })

  it('refuses value and top-Items filters on a nested field until parent-group semantics are measured', () => {
    const p = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0, 0, 1, 1], [10, 1, 1, 20])
    p.cache.fields[0] = { databaseField: true, sharedItems: { inline: false, items: [{ type: 'string', value: 'X' }, { type: 'string', value: 'Y' }] } } as never
    p.cache.fields.push({ databaseField: true, sharedItems: { inline: false, items: [{ type: 'string', value: 'A' }, { type: 'string', value: 'B' }] } } as never)
    p.def.fields[0] = { items: [{ type: 'data', x: 0 } as never, { type: 'data', x: 1 } as never] } as never
    p.def.fields.push({ items: [{ type: 'data', x: 0 }, { type: 'data', x: 1 }] } as never)
    p.store = { ...p.store, columns: [
      { kind: 'shared', index: Uint32Array.from([0, 0, 1, 1]) },
      { kind: 'shared', index: Uint32Array.from([0, 1, 0, 1]) },
      p.store.columns[1]!,
    ] }
    p.view.rows = [0, 2]
    p.view.fields = [field(2), field(2), field(2)]
    for (const filterValue of [
      filter({ fld: 2, type: 'valueGreaterThan', iMeasureFld: 0, stringValue1: '5' }),
      filter({ fld: 2, type: 'count', iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: false, val: 1 } }] } }),
    ]) {
      p.view.filters = [filterValue]
      expect(result(p)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
    }
  })

  it('excludes same-field filtered items from showAll while retaining allowed no-data items', () => {
    const f = fixture(['Apple', 'Berry', 'Citrus'].map(value => ({ type: 'string', value })), [0], [1], [filter({ type: 'captionEqual', stringValue1: 'Apple' })])
    f.view.fields = [{ ...f.view.fields[0]!, showAll: true }, f.view.fields[1]!]
    const axes = result(f)
    expect(axes.ok).toBe(true)
    if (axes.ok) expect(axes.value.rows.map(line => line.items[0])).toEqual([0])
  })

  it('applies value filters using the selected data field aggregate', () => {
    const f = fixture(['A', 'B'].map(value => ({ type: 'string', value })), [0, 0, 1, 1], [2, 3, 9, 11], [filter({ type: 'valueGreaterThan', iMeasureFld: 0, stringValue1: '10' })])
    const axes = result(f)
    expect(axes.ok).toBe(true)
    if (axes.ok) expect([...axes.value.rowKey]).toEqual([NO_LINE, NO_LINE, 0, 0])
  })

  it('refuses top-10 threshold ties and dynamic date filters', () => {
    const tied = fixture(['A', 'B', 'C'].map(value => ({ type: 'string', value })), [0, 1, 2], [10, 10, 1], [filter({
      type: 'sum', iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: false, val: 1 } }] },
    })])
    expect(result(tied)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
    const unique = fixture(['A', 'B', 'C'].map(value => ({ type: 'string', value })), [0, 1, 2], [10, 5, 1], [filter({
      type: 'sum', iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: true, val: 34 } }] },
    })])
    expect(result(unique)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
    const dynamic = fixture([{ type: 'date', value: '2026-10-30T00:00:00' }], [0], [1], [filter({ type: 'today' })])
    expect(result(dynamic)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-filter-dynamic-date' } })
  })

  it('maps date groups with blanks alongside their base field and refuses unmeasured out-of-range buckets', () => {
    const cache = {
      source: { type: 'worksheet' }, olap: false,
      fields: [
        { databaseField: true, sharedItems: { inline: false, items: [
          { type: 'date', value: '2020-06-01T00:00:00' }, { type: 'date', value: '2021-03-02T00:00:00' }, { type: 'missing' },
        ] } },
        { databaseField: false, fieldGroup: { base: 0, rangePr: { autoStart: false, autoEnd: false, groupBy: 'years', groupInterval: 1,
          startDate: '2020-01-01T00:00:00', endDate: '2021-12-31T00:00:00' }, groupItems: [
          { type: 'string', value: '(blank)' }, { type: 'string', value: '2020' }, { type: 'string', value: '2021' },
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
    const outside = { ...cache, fields: [{ ...cache.fields[0]!, sharedItems: { inline: false, items: [
      { type: 'date', value: '2019-01-01T00:00:00' },
    ] } }, cache.fields[1]!] } as unknown as PivotCacheDefinition
    expect(groupColumn({ ...store, count: 1, columns: [{ kind: 'shared', index: Uint32Array.of(0) }] }, outside, 1))
      .toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
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

  it('maps numeric range boundaries arithmetically, including decimal intervals', () => {
    const cache = { fields: [
      { databaseField: true, sharedItems: { inline: false, items: [-1, 0.5, 0.2, 2, 3].map(value => ({ type: 'number', value })) } },
      { databaseField: false, fieldGroup: { base: 0, rangePr: { autoStart: false, autoEnd: false, groupBy: 'range', groupInterval: 0.5, startNum: 0, endNum: 2 },
        groupItems: ['<0', '0-0.5', '0.5-1', '1-1.5', '1.5-2', '>2'].map(value => ({ type: 'string', value })) } },
    ] } as unknown as PivotCacheDefinition
    const store: RecordStore = { count: 5, bytes: 0, columns: [{ kind: 'shared', index: Uint32Array.from([0, 1, 2, 3, 4]) }] }
    expect(groupColumn(store, cache, 1)).toMatchObject({ ok: true, value: Uint32Array.from([0, 2, 1, 4, 5]) })
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
  if (!cases.length) it.skip('SKIPPED: Task 21 has no runnable Excel manifest cases', () => {})
  else it.each(cases.map(entry => [entry.case, entry] as const))('manifest %s', async (_name, entry) => {
    const local = resolve(fixtures, 'engine', entry.file)
    const loaded = await loadAll(existsSync(local) ? local : resolve(process.cwd(), entry.file))
    const pivot = loaded.find(candidate => candidate.def.name === entry.pivot)
    if (!pivot) throw new Error(`${entry.case}: pivot ${entry.pivot} not found`)
    expect(checkSavedLayout(pivot.store, pivot.cache, pivot.def, sheetCellReader(pivot.sheet), { date1904: pivot.date1904 }).status, entry.case).toBe('ok')
  })
})
