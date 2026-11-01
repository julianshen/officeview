// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { loadAll } from './helpers/xlsx-pivots-load'
import { groupColumn } from '../src/xlsx/pivots/engine/groups'
import { PIVOT_DELTA_KINDS, type PivotDelta, type PivotDeltaContext } from '../src/xlsx/pivots/session/delta'
import { openPivotSession } from '../src/xlsx/pivots/session/session'
import { sheetCellReader } from '../src/xlsx/pivots/engine/faithful'
import { initialViewState } from '../src/xlsx/pivots/engine/view'
import { applyGrouping } from '../src/xlsx/pivots/session/grouping'

const ROOT = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const interaction = (name: string) => resolve(ROOT, `interaction/${name}.xlsx`)
const pivot = (loaded: Awaited<ReturnType<typeof loadAll>>[number], verified = new Set(PIVOT_DELTA_KINDS)) => {
  const saved = sheetCellReader(loaded.sheet)
  const opened = openPivotSession({ def: loaded.def, cache: loaded.cache, store: loaded.store, date1904: loaded.date1904,
    saved, occupied: () => false, linkedCharts: [], verifiedKinds: verified })
  if (!opened.ok) throw new Error(`${opened.diagnostic.feature}: ${opened.diagnostic.message}`)
  return opened.value
}
const groupingContext = (def: PivotDeltaContext['def'], cache: PivotDeltaContext['cache'], store: PivotDeltaContext['store'], view: PivotDeltaContext['view'], date1904: boolean): PivotDeltaContext => ({
  def, cache, store, view, date1904, captionRule: { locale: 'zh-TW', caption: () => undefined },
})

describe('Task 37 grouping', () => {
  it('groups days, N-day intervals, hours, minutes and seconds including serial 60', async () => {
    for (const [file, groupBy] of [
      ['s4-group-date-days-single', 'days'], ['s4-group-date-days-7', 'days'],
      ['s4-group-date-time-hms', 'hours'], ['s4-group-date-time-hms', 'minutes'], ['s4-group-date-time-hms', 'seconds'],
    ] as const) {
      const loaded = (await loadAll(interaction(file)))[0]!
      const field = loaded.cache.fields.findIndex(item => item.fieldGroup?.rangePr?.groupBy === groupBy)
      expect(field, `${file} should contain the saved grouped cache field`).toBeGreaterThanOrEqual(0)
      const grouped = groupColumn(loaded.store, loaded.cache, field)
      expect(grouped.ok, grouped.ok ? undefined : `${grouped.diagnostic.feature}: ${grouped.diagnostic.message}`).toBe(true)
    }
  })

  it('maps a missing date record into the measured lower bucket for a day group', async () => {
    const loaded = (await loadAll(interaction('s4-group-date-days-single')))[0]!
    const field = loaded.cache.fields.findIndex(item => item.fieldGroup?.rangePr?.groupBy === 'days')
    const base = loaded.cache.fields[field]!.fieldGroup!.base!
    const original = loaded.cache.fields[base]!.sharedItems!
    const missingIndex = original.items.length
    const cache = { ...loaded.cache, fields: loaded.cache.fields.map((item, index) => index === base
      ? { ...item, sharedItems: { ...original, containsBlank: true, items: [...original.items, { type: 'missing' as const }] } } : item) }
    const column = loaded.store.columns[base]!
    if (column.kind !== 'shared') throw new Error('Expected a shared date column')
    const indices = new Uint32Array(column.index)
    indices[0] = missingIndex
    const store = { ...loaded.store, columns: loaded.store.columns.map((item, index) => index === base ? { ...column, index: indices } : item) }
    const grouped = groupColumn(store, cache, field)
    expect(grouped.ok, grouped.ok ? undefined : `${grouped.diagnostic.feature}: ${grouped.diagnostic.message}`).toBe(true)
    if (!grouped.ok) return
    const lower = loaded.cache.fields[field]!.fieldGroup!.groupItems!.findIndex(item => item.type === 'string' && item.value.startsWith('<'))
    expect(lower).toBeGreaterThanOrEqual(0)
    expect(grouped.value[0]).toBe(lower)
  })

  it('groups number ranges with blank, negative, below-start and above-end source records', async () => {
    const loaded = (await loadAll(interaction('s4-group-number-sales-blanks-manual')))[0]!
    const field = loaded.cache.fields.findIndex(item => item.fieldGroup?.rangePr?.groupBy === 'range')
    expect(field).toBeGreaterThanOrEqual(0)
    const grouped = groupColumn(loaded.store, loaded.cache, field)
    expect(grouped.ok, grouped.ok ? undefined : `${grouped.diagnostic.feature}: ${grouped.diagnostic.message}`).toBe(true)
  })

  it('applies a multi-level date group as a pure candidate, then commits its cache and axis hierarchy', async () => {
    const seed = (await loadAll(interaction('interaction-filter-date-base')))[0]!
    const session = pivot(seed)
    const oldCache = session.cache, oldRows = session.view.rows
    const delta: PivotDelta = { kind: 'groupDate', field: 5, periods: ['months', 'quarters', 'years'] }
    const candidate = session.apply(delta)
    expect(candidate.ok, candidate.ok ? undefined : `${candidate.diagnostic.feature}: ${candidate.diagnostic.message}`).toBe(true)
    if (!candidate.ok) return
    expect(session.cache).toBe(oldCache)
    expect(session.view.rows).toBe(oldRows)
    expect(candidate.value.view.rows).toEqual(expect.arrayContaining([expect.any(Number)]))
    expect(candidate.value.cache.fields.length).toBeGreaterThan(oldCache.fields.length)
    expect(session.commit(candidate.value)).toBe(true)
    expect(session.cache).toBe(candidate.value.cache)
    expect(session.view.rows.slice(-3)).toEqual([14, 13, 12])
    const regroup = session.apply({ kind: 'groupDate', field: 14, periods: ['months', 'years'] })
    expect(regroup.ok, regroup.ok ? undefined : `${regroup.diagnostic.feature}: ${regroup.diagnostic.message}`).toBe(true)
    if (!regroup.ok) return
    expect(session.commit(regroup.value)).toBe(true)
    expect(session.view.rows.slice(-2)).toEqual([13, 12])
    const ungrouped = session.apply({ kind: 'ungroup', field: 5 })
    expect(ungrouped.ok, ungrouped.ok ? undefined : `${ungrouped.diagnostic.feature}: ${ungrouped.diagnostic.message}`).toBe(true)
    if (!ungrouped.ok) return
    expect(session.commit(ungrouped.value)).toBe(true)
    expect(session.view.rows).toContain(5)
    expect(session.cache.fields).toHaveLength(oldCache.fields.length)
  })

  it('points the base date group at its coarsest parent', async () => {
    const seed = (await loadAll(interaction('interaction-filter-date-base')))[0]!
    const session = pivot(seed)
    const grouped = session.apply({ kind: 'groupDate', field: 5, periods: ['months', 'quarters', 'years'] })
    expect(grouped.ok).toBe(true)
    if (!grouped.ok) return
    expect(grouped.value.cache.fields[5]?.fieldGroup?.par).toBe(14)
    expect(grouped.value.cache.fields[14]?.fieldGroup?.rangePr?.groupBy).toBe('years')
  })

  it('ungroups a date field at index zero', async () => {
    const seed = (await loadAll(interaction('interaction-filter-date-base')))[0]!
    const fields = [...seed.cache.fields], defs = [...seed.def.fields], columns = [...seed.store.columns]
    ;[fields[0], fields[5]] = [fields[5]!, fields[0]!]
    ;[defs[0], defs[5]] = [defs[5]!, defs[0]!]
    ;[columns[0], columns[5]] = [columns[5]!, columns[0]!]
    const remap = (field: number) => field === 0 ? 5 : field === 5 ? 0 : field
    const loaded = { ...seed, cache: { ...seed.cache, fields }, def: { ...seed.def, fields: defs,
      rowFields: seed.def.rowFields.map(remap), colFields: seed.def.colFields.map(remap), pageFields: seed.def.pageFields.map(item => ({ ...item, fld: remap(item.fld) })) },
      store: { ...seed.store, columns } }
    const session = pivot(loaded)
    const grouped = session.apply({ kind: 'groupDate', field: 0, periods: ['months'] })
    expect(grouped.ok).toBe(true)
    if (!grouped.ok) return
    expect(session.commit(grouped.value)).toBe(true)
    const ungrouped = session.apply({ kind: 'ungroup', field: 0 })
    expect(ungrouped.ok, ungrouped.ok ? undefined : `${ungrouped.diagnostic.feature}: ${ungrouped.diagnostic.message}`).toBe(true)

    const numericSeed = (await loadAll(interaction('s4-group-number-units-auto')))[0]!
    const numericFields = [...numericSeed.cache.fields], numericDefs = [...numericSeed.def.fields], numericColumns = [...numericSeed.store.columns]
    ;[numericFields[0], numericFields[3]] = [numericFields[3]!, numericFields[0]!]
    ;[numericDefs[0], numericDefs[3]] = [numericDefs[3]!, numericDefs[0]!]
    ;[numericColumns[0], numericColumns[3]] = [numericColumns[3]!, numericColumns[0]!]
    const numericCache = { ...numericSeed.cache, fields: numericFields }, numericDef = { ...numericSeed.def, fields: numericDefs, rowFields: [0], colFields: [] }
    const numericContext = groupingContext(numericDef, numericCache, { ...numericSeed.store, columns: numericColumns }, initialViewState(numericDef, numericCache), numericSeed.date1904)
    const numberGroup = applyGrouping(numericContext, { kind: 'groupNumber', field: 0, start: 0, end: 5000, by: 1000 })
    expect(numberGroup.ok).toBe(true)
    if (!numberGroup.ok) return
    const numberUngroup = applyGrouping(groupingContext(numberGroup.value.def, numberGroup.value.cache, numberGroup.value.store, numberGroup.value.view, numericSeed.date1904), { kind: 'ungroup', field: 0 })
    expect(numberUngroup.ok, numberUngroup.ok ? undefined : `${numberUngroup.diagnostic.feature}: ${numberUngroup.diagnostic.message}`).toBe(true)
  })

  it('keeps surviving grouped field indexes valid when regrouping an earlier hierarchy', async () => {
    const seed = (await loadAll(interaction('interaction-filter-date-base')))[0]!
    const cacheFields = [...seed.cache.fields], defFields = [...seed.def.fields], columns = [...seed.store.columns]
    const base = 5, second = cacheFields.length
    cacheFields.push({ ...cacheFields[base]!, name: 'Second date' })
    defFields.push({ ...defFields[base]!, name: 'Second date', items: [] })
    columns.push(columns[base]!)
    const def = { ...seed.def, fields: defFields, rowFields: [...seed.def.rowFields, second] }
    const cache = { ...seed.cache, fields: cacheFields }, store = { ...seed.store, columns }
    let context: PivotDeltaContext = groupingContext(def, cache, store, initialViewState(def, cache), seed.date1904)
    const date1 = applyGrouping(context, { kind: 'groupDate', field: base, periods: ['months', 'years'] })
    expect(date1.ok).toBe(true)
    if (!date1.ok) return
    context = { ...context, ...date1.value }
    const date2 = applyGrouping(context, { kind: 'groupDate', field: second, periods: ['months'] })
    expect(date2.ok).toBe(true)
    if (!date2.ok) return
    context = { ...context, ...date2.value }
    const regroup = applyGrouping(context, { kind: 'groupDate', field: base, periods: ['years'] })
    expect(regroup.ok).toBe(true)
    if (!regroup.ok) return
    const refs = [...regroup.value.view.rows, ...regroup.value.view.cols, ...regroup.value.view.pages.map(p => p.field), ...regroup.value.view.filters.map(f => f.fld),
      ...regroup.value.def.rowFields, ...regroup.value.def.colFields, ...regroup.value.def.pageFields.map(p => p.fld)]
    expect(refs.every(field => field < regroup.value.cache.fields.length)).toBe(true)
    expect(regroup.value.def.fields).toHaveLength(regroup.value.cache.fields.length)
    const surviving = regroup.value.cache.fields.findIndex(field => field.fieldGroup?.base === second)
    expect(surviving).toBeGreaterThanOrEqual(0)
    expect(regroup.value.view.rows).toContain(surviving)
    expect(regroup.value.cache.fields[surviving]?.fieldGroup?.base).toBe(second)
  })

  it('does not create an inverted final N-day bucket', async () => {
    const loaded = (await loadAll(interaction('s4-group-date-days-7')))[0]!
    const field = loaded.cache.fields.findIndex(item => item.fieldGroup?.rangePr?.groupBy === 'days')
    const base = loaded.cache.fields[field]!.fieldGroup!.base!
    const session = pivot(loaded)
    const grouped = session.apply({ kind: 'groupDate', field: base, periods: ['days'], daysInterval: 7 })
    expect(grouped.ok).toBe(true)
    if (!grouped.ok) return
    const labels = grouped.value.cache.fields.at(-1)?.fieldGroup?.groupItems?.map(item => item.type === 'string' ? item.value : '') ?? []
    expect(labels).toHaveLength(312)
    expect(labels).not.toContain('2024/12/24 - 2024/12/21')
  })

  it('reserves date captions before allocating them', async () => {
    const seed = (await loadAll(interaction('interaction-filter-date-base')))[0]!
    const session = pivot(seed)
    let reserved = 0, arrayAllocations = 0
    const context: PivotDeltaContext = { ...groupingContext(session.def, session.cache, seed.store, session.view, seed.date1904), admitCandidateBytes: (bytes: number) => { reserved = bytes; return undefined } }
    const originalFrom = Array.from
    Array.from = function (...args: Parameters<typeof Array.from>) { arrayAllocations++; return originalFrom(...args) } as typeof Array.from
    let result
    try { result = applyGrouping(context, { kind: 'groupDate', field: 5, periods: ['years'] }) }
    finally { Array.from = originalFrom }
    expect(reserved).toBeGreaterThan(0)
    expect(arrayAllocations).toBe(0)
    expect(result?.ok).toBe(false)
  })

  it('reserves the field-remap candidate before grouping a large cache', async () => {
    const seed = (await loadAll(interaction('interaction-filter-date-base')))[0]!
    const count = 1024
    const cacheFields = Array.from({ length: count }, (_, index) => ({ ...seed.cache.fields[index % seed.cache.fields.length]!, name: `Field ${index}`, fieldGroup: undefined }))
    const defFields = Array.from({ length: count }, (_, index) => ({ ...seed.def.fields[index % seed.def.fields.length]!, name: `Field ${index}` }))
    const cache = { ...seed.cache, fields: cacheFields }
    const def = { ...seed.def, fields: defFields }
    const view = initialViewState(def, cache)
    let reserved = 0
    const result = applyGrouping({ ...groupingContext(def, cache, seed.store, view, seed.date1904), admitCandidateBytes: bytes => {
      reserved = bytes
      return { bytes, release() {} }
    } }, { kind: 'groupDate', field: 5, periods: ['months'] })
    expect(result.ok).toBe(true)
    expect(reserved).toBeGreaterThanOrEqual(count * 200)
  })

  it('uses discrete inclusive labels for a non-negative integer number field', async () => {
    const seed = (await loadAll(interaction('s4-group-number-units-auto')))[0]!
    const cache = { ...seed.cache, fields: seed.cache.fields.map((field, index) => index === 4 ? { ...field, sharedItems: { ...field.sharedItems!, containsInteger: true, minValue: 0, maxValue: 50 } } : field) }
    const context = groupingContext(seed.def, cache, seed.store, initialViewState(seed.def, cache), seed.date1904)
    const grouped = applyGrouping(context, { kind: 'groupNumber', field: 4, start: 0, end: 50, by: 25 })
    expect(grouped.ok).toBe(true)
    if (!grouped.ok) return
    expect(grouped.value.cache.fields[4]?.fieldGroup?.groupItems?.map(item => item.type === 'string' ? item.value : '')).toEqual(['<0', '0-24', '25-49', '>50'])
  })

  it('ungroups a grouped inline base only after reserving its shared-item materialization', async () => {
    const seed = (await loadAll(interaction('s4-group-number-units-auto')))[0]!
    const field = 3
    expect(seed.cache.fields[field]?.sharedItems?.inline).toBe(true)
    const def = { ...seed.def, rowFields: [field], colFields: [] }
    const context = groupingContext(def, seed.cache, seed.store, initialViewState(def, seed.cache), seed.date1904)
    const grouped = applyGrouping(context,
      { kind: 'groupNumber', field, start: 0, end: 5000, by: 1000 })
    expect(grouped.ok).toBe(true)
    if (!grouped.ok) return
    let reserved = 0
    const refused = applyGrouping({ ...groupingContext(grouped.value.def, grouped.value.cache, grouped.value.store, grouped.value.view, seed.date1904), admitCandidateBytes: bytes => { reserved = bytes; return undefined } }, { kind: 'ungroup', field })
    expect(refused.ok).toBe(false)
    expect(reserved).toBeGreaterThan(seed.store.count * 4)
    const ungrouped = applyGrouping({ ...groupingContext(grouped.value.def, grouped.value.cache, grouped.value.store, grouped.value.view, seed.date1904), admitCandidateBytes: bytes => ({ bytes, release() {} }) }, { kind: 'ungroup', field })
    expect(ungrouped.ok).toBe(true)
    if (ungrouped.ok) expect(ungrouped.value.cache.fields[field]?.sharedItems?.inline).toBe(false)
  })

  it('ungroups a saved discrete group and restores the unchanged base items', async () => {
    const seed = (await loadAll(interaction('s4-ungroup-manual-discrete')))[0]!
    const base = 1, virtual = seed.cache.fields.length
    const labels = [{ type: 'string' as const, value: 'Fruit' }, { type: 'string' as const, value: 'Other' }]
    const cache = { ...seed.cache, fields: [...seed.cache.fields, { ...seed.cache.fields[base]!, name: 'Fruit group', databaseField: false,
      fieldGroup: { base, discretePr: [0, 0, 1, 1], groupItems: labels }, sharedItems: { ...seed.cache.fields[base]!.sharedItems!, items: labels } }] }
    const def = { ...seed.def, fields: [...seed.def.fields, { ...seed.def.fields[base]!, name: 'Fruit group', items: [] }], rowFields: [...seed.def.rowFields, virtual] }
    const context = groupingContext(def, cache, seed.store, initialViewState(def, cache), seed.date1904)
    const ungrouped = applyGrouping(context, { kind: 'ungroup', field: virtual })
    expect(ungrouped.ok).toBe(true)
    if (!ungrouped.ok) return
    expect(ungrouped.value.cache.fields).toHaveLength(seed.cache.fields.length)
    expect(ungrouped.value.cache.fields[base]?.sharedItems?.items).toEqual(seed.cache.fields[base]?.sharedItems?.items)
  })

  it('regrouping removes obsolete fields from filters and sort-area references', async () => {
    const seed = (await loadAll(interaction('interaction-filter-date-base')))[0]!
    const session = pivot(seed)
    const grouped = session.apply({ kind: 'groupDate', field: 5, periods: ['months', 'years'] })
    expect(grouped.ok).toBe(true)
    if (!grouped.ok) return
    const obsolete = 12
    const filter = { fld: obsolete, type: 'captionEqual' as const, evalOrder: 0, id: 99, autoFilter: { columns: [] } }
    const ref = { field: obsolete, selected: true, byPosition: false, relative: false, defaultSubtotal: false, subtotals: [], x: [] }
    const area = { type: 'normal' as const, dataOnly: false, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false, outline: false, collapsedLevelsAreSubtotals: false, references: [ref] }
    const context: PivotDeltaContext = { def: grouped.value.def, cache: grouped.value.cache, store: grouped.value.store,
      view: { ...grouped.value.view, filters: [filter], fields: grouped.value.view.fields.map((state, index) => index === 0 ? { ...state, autoSortScope: area } : state) },
      date1904: seed.date1904, captionRule: { locale: 'zh-TW', caption: () => undefined } }
    const regrouped = applyGrouping(context, { kind: 'groupDate', field: 5, periods: ['years'] })
    expect(regrouped.ok).toBe(true)
    if (!regrouped.ok) return
    expect(regrouped.value.view.filters).toEqual([])
    expect(regrouped.value.view.fields[0]?.autoSortScope).toBeUndefined()
  })

  it('reset after multi-level grouping restores the saved cache, definition and view', async () => {
    const seed = (await loadAll(interaction('interaction-filter-date-base')))[0]!
    const session = pivot(seed), original = { cache: session.cache, def: session.def, view: session.view }
    const grouped = session.apply({ kind: 'groupDate', field: 5, periods: ['months', 'quarters', 'years'] })
    expect(grouped.ok).toBe(true)
    if (!grouped.ok) return
    expect(session.commit(grouped.value)).toBe(true)
    const reset = session.reset()
    expect(reset.ok).toBe(true)
    if (!reset.ok) return
    expect(reset.value.cache).toBe(original.cache)
    expect(reset.value.def).toBe(original.def)
    expect(reset.value.view).toEqual(original.view)
  })

  it('groups and ungroups a number field without mutating the opened cache', async () => {
    const loaded = (await loadAll(interaction('s4-move-sales-only-rows')))[0]!
    const session = pivot(loaded)
    const oldCache = session.cache
    const grouped = session.apply({ kind: 'groupNumber', field: 3, start: 0, end: 5000, by: 1000 })
    expect(grouped.ok, grouped.ok ? undefined : `${grouped.diagnostic.feature}: ${grouped.diagnostic.message}`).toBe(true)
    if (!grouped.ok) return
    expect(grouped.value.cache.fields[3]?.fieldGroup?.groupItems).toHaveLength(7)
    expect(session.cache).toBe(oldCache)
    expect(session.commit(grouped.value)).toBe(true)
    const ungrouped = session.apply({ kind: 'ungroup', field: 3 })
    expect(ungrouped.ok, ungrouped.ok ? undefined : `${ungrouped.diagnostic.feature}: ${ungrouped.diagnostic.message}`).toBe(true)
    if (!ungrouped.ok) return
    expect(ungrouped.value.cache.fields[3]?.fieldGroup).toBeUndefined()
    expect(session.commit(ungrouped.value)).toBe(true)
  })

  it('admits cloned auto-sort reference arrays before ungrouping', async () => {
    const loaded = (await loadAll(interaction('s4-move-sales-only-rows')))[0]!
    const groupedField = 3
    const base = groupingContext(loaded.def, loaded.cache, loaded.store, initialViewState(loaded.def, loaded.cache), loaded.date1904)
    const grouped = applyGrouping(base, { kind: 'groupNumber', field: groupedField, start: 0, end: 5000, by: 1000 })
    expect(grouped.ok).toBe(true)
    if (!grouped.ok) return
    const retainedField = 0
    const referenceCount = 20_000
    const references = Array.from({ length: referenceCount }, (_, index) => ({
      field: undefined, selected: true, byPosition: false, relative: false, defaultSubtotal: false, subtotals: [], x: [index],
    }))
    const def = { ...grouped.value.def, fields: grouped.value.def.fields.map((field, index) => index === retainedField
      ? { ...field, autoSortScope: { type: 'normal' as const, dataOnly: false, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false,
        outline: false, collapsedLevelsAreSubtotals: false, references } } : field) }
    const ctx = groupingContext(def, grouped.value.cache, grouped.value.store, grouped.value.view, loaded.date1904)
    let admittedBytes = 0
    const result = applyGrouping({ ...ctx, admitCandidateBytes(bytes) { admittedBytes = bytes; return { bytes, release() {} } } }, { kind: 'ungroup', field: groupedField })
    expect(result.ok).toBe(true)
    expect(admittedBytes).toBeGreaterThanOrEqual(referenceCount * 80)
    if (result.ok) expect(result.value.def.fields[retainedField]?.autoSortScope?.references).toHaveLength(referenceCount)
  })
})
