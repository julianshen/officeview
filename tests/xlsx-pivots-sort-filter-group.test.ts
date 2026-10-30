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

  it('sorts a top-level field by its rankBy data aggregate and refuses tied ranks', () => {
    const p = fixture(['A', 'B', 'C'].map(value => ({ type: 'string', value })), [0, 0, 1, 1, 2], [2, 3, 9, 1, 4])
    p.def.fields[0] = { ...p.def.fields[0]!, rankBy: 0 }
    p.view.fields = [{ ...p.view.fields[0]!, sortType: 'descending', autoSortScope: {
      type: 'normal', dataOnly: true, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false,
      outline: true, collapsedLevelsAreSubtotals: false, references: [],
    } }, p.view.fields[1]!]
    expect(lines(result(p))).toEqual([1, 0, 2])
    p.store = { ...p.store, columns: [p.store.columns[0]!, { kind: 'inline', tag: new Uint8Array(5).fill(VALUE_TAG.number), num: Float64Array.from([2, 3, 9, 1, 5]), text: [], textIndex: new Uint32Array(5) }] }
    expect(result(p)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
  })

  it('refuses unmeasured mixed-type and built-in custom-list sort order', () => {
    const mixed = fixture([{ type: 'number', value: 1 }, { type: 'string', value: 'A' }, { type: 'missing' }], [0, 1, 2], [1, 1, 1])
    mixed.view.fields = [{ ...mixed.view.fields[0]!, sortType: 'ascending' }, mixed.view.fields[1]!]
    expect(result(mixed)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
    const months = fixture(['Jan', 'Feb', 'Mar'].map(value => ({ type: 'string', value })), [0, 1, 2], [1, 1, 1])
    months.view.fields = [{ ...months.view.fields[0]!, sortType: 'ascending' }, months.view.fields[1]!]
    expect(result(months)).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-engine-unmeasured' } })
  })

  it('applies case-insensitive wildcard label filtering', () => {
    const f = fixture(['Apple', 'apricot', 'Berry'].map(value => ({ type: 'string', value })), [0, 1, 2], [1, 1, 1], [filter({ type: 'captionBeginsWith', stringValue1: 'AP*' })])
    const axes = result(f)
    expect(axes.ok).toBe(true)
    if (axes.ok) expect([...axes.value.rowKey]).toEqual([0, 1, NO_LINE])
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
    expect(result(unique).ok).toBe(true)
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
    expect(computeAxes(store, cache, def, view).ok).toBe(true)
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
