// @vitest-environment node
/**
 * The pivot session (plan Task 25, spec §6.4, §7.1, §7.6, §7.7): the pure view deltas, the caption rule, open / apply /
 * commit / reset, and the resident runtime that keeps caches and streams records on demand.
 */
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { applyDelta, type PivotDelta } from '../src/xlsx/pivots/session/delta'
import { computePivot } from '../src/xlsx/pivots/engine/compute'
import { initialViewState, type PivotViewState } from '../src/xlsx/pivots/engine/view'
import type { PivotFilter } from '../src/xlsx/pivots/model'
import { loadAll, type Loaded } from './helpers/xlsx-pivots-load'

const FIXTURES = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const ENGINE = resolve(FIXTURES, 'engine')
const cached = new Map<string, Promise<Loaded[]>>()
const load = (name: string) => { const path = resolve(ENGINE, `${name}.xlsx`); let p = cached.get(path); if (!p) cached.set(path, p = loadAll(path)); return p }
const baseline = async () => (await load('baseline'))[0]!
/** Field indices of the engine seed (scripts/xlsx-refs/engine-cases.ts `F`). */
const REGION = 0, PRODUCT = 1, YEAR = 2

const labelFilter = (fld: number, value: string, id = 1): PivotFilter => ({
  fld, type: 'captionBeginsWith', evalOrder: -1, id, stringValue1: value,
  autoFilter: { ref: 'A1', columns: [{ colId: 0, hiddenButton: false, showButton: true, customFilters: { and: false, filters: [{ operator: 'equal', val: `${value}*` }] } }] },
})
const valueFilter = (fld: number, above: number, id = 2): PivotFilter => ({
  fld, type: 'valueGreaterThan', evalOrder: -1, id, iMeasureFld: 0,
  autoFilter: { ref: 'A1', columns: [{ colId: 0, hiddenButton: false, showButton: true, customFilters: { and: false, filters: [{ operator: 'greaterThan', val: String(above) }] } }] },
})
const topFilter = (fld: number, count: number, id = 3): PivotFilter => ({
  fld, type: 'count', evalOrder: -1, id, iMeasureFld: 0,
  autoFilter: { ref: 'A1', columns: [{ colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: false, val: count, filterVal: count } }] },
})
const apply = (p: Loaded, view: PivotViewState, delta: PivotDelta) => applyDelta(view, delta, p.def, p.cache)
function ok<T>(result: { ok: true; value: T } | { ok: false; diagnostic: { feature: string; message: string } }): T {
  if (!result.ok) throw new Error(`refused ${result.diagnostic.feature}: ${result.diagnostic.message}`)
  return result.value
}
const refusal = (result: { ok: boolean; diagnostic?: { feature: string } }) => result.ok ? 'ok' : result.diagnostic!.feature

describe('applyDelta maps each kind onto the view', () => {
  it('collapse marks collapsed[item]; collapseField and collapseAll mark every item', async () => {
    const p = await baseline()
    const view = initialViewState(p.def, p.cache)
    const one = ok(apply(p, view, { kind: 'collapse', field: REGION, item: 0, collapsed: true }))
    expect(one.fields[REGION]!.collapsed).toEqual([true, false, false, false, false, false])
    // Pure: the input view is untouched.
    expect(view.fields[REGION]!.collapsed).toEqual([false, false, false, false, false, false])
    expect(ok(apply(p, one, { kind: 'collapse', field: REGION, item: 0, collapsed: false })).fields[REGION]!.collapsed).toEqual(view.fields[REGION]!.collapsed)
    // A field's data items only: the default-subtotal item keeps its state.
    expect(ok(apply(p, view, { kind: 'collapseField', field: REGION, collapsed: true })).fields[REGION]!.collapsed).toEqual([true, true, true, true, true, false])
    const all = ok(apply(p, view, { kind: 'collapseAll', collapsed: true }))
    expect(all.fields[REGION]!.collapsed).toEqual([true, true, true, true, true, false])
    expect(all.fields[PRODUCT]!.collapsed).toEqual([true, true, true, true, false])
    expect(all.fields[YEAR]!.collapsed).toEqual([true, true, true, false])
    // Fields on no axis are left alone.
    expect(all.fields[3]).toBe(view.fields[3])
  })

  it('collapse refuses an item that is not a data item of the field', async () => {
    const p = await baseline()
    const view = initialViewState(p.def, p.cache)
    expect(refusal(apply(p, view, { kind: 'collapse', field: REGION, item: 5, collapsed: true }))).toBe('pivot-delta-invalid')
    expect(refusal(apply(p, view, { kind: 'collapse', field: REGION, item: 9, collapsed: true }))).toBe('pivot-delta-invalid')
    expect(refusal(apply(p, view, { kind: 'collapse', field: 99, item: 0, collapsed: true }))).toBe('pivot-delta-invalid')
  })

  it('items sets hidden, and refuses hiding every item', async () => {
    const p = await baseline()
    const view = initialViewState(p.def, p.cache)
    const next = ok(apply(p, view, { kind: 'items', field: REGION, hidden: [3, 4] }))
    expect(next.fields[REGION]!.hidden).toEqual([false, false, false, true, true, false])
    expect(ok(apply(p, next, { kind: 'items', field: REGION, hidden: [] })).fields[REGION]!.hidden).toEqual(view.fields[REGION]!.hidden)
    expect(refusal(apply(p, view, { kind: 'items', field: REGION, hidden: [0, 1, 2, 3, 4] }))).toBe('pivot-delta-invalid')
    expect(refusal(apply(p, view, { kind: 'items', field: REGION, hidden: [5] }))).toBe('pivot-delta-invalid')
  })

  it('page sets pages, only on a page field', async () => {
    const p = (await load('quirk-page-single'))[0]!
    const view = initialViewState(p.def, p.cache)
    expect(view.pages).toEqual([{ field: PRODUCT, selected: [0] }])
    expect(ok(apply(p, view, { kind: 'page', field: PRODUCT, selected: 'all' })).pages).toEqual([{ field: PRODUCT, selected: 'all' }])
    expect(ok(apply(p, view, { kind: 'page', field: PRODUCT, selected: [1, 2] })).pages).toEqual([{ field: PRODUCT, selected: [1, 2] }])
    expect(refusal(apply(p, view, { kind: 'page', field: PRODUCT, selected: [] }))).toBe('pivot-delta-invalid')
    expect(refusal(apply(p, view, { kind: 'page', field: PRODUCT, selected: [4] }))).toBe('pivot-delta-invalid')
    expect(refusal(apply(p, view, { kind: 'page', field: REGION, selected: 'all' }))).toBe('pivot-delta-invalid')
  })

  it('filter replaces that field\'s filter of the same family; another family refuses as the engine does', async () => {
    const p = await baseline()
    const view = initialViewState(p.def, p.cache)
    const first = ok(apply(p, view, { kind: 'filter', field: PRODUCT, filter: labelFilter(PRODUCT, 'B') }))
    expect(first.filters).toEqual([labelFilter(PRODUCT, 'B')])
    // Same family (label or date; value or top-10): replaced in place.
    const replaced = ok(apply(p, first, { kind: 'filter', field: PRODUCT, filter: labelFilter(PRODUCT, 'C', 7) }))
    expect(replaced.filters).toEqual([labelFilter(PRODUCT, 'C', 7)])
    const value = ok(apply(p, view, { kind: 'filter', field: REGION, filter: valueFilter(REGION, 50000) }))
    expect(ok(apply(p, value, { kind: 'filter', field: REGION, filter: topFilter(REGION, 2) })).filters).toEqual([topFilter(REGION, 2)])
    // Other fields keep their filters.
    expect(ok(apply(p, first, { kind: 'filter', field: REGION, filter: valueFilter(REGION, 50000) })).filters).toEqual([labelFilter(PRODUCT, 'B'), valueFilter(REGION, 50000)])
    // The seed saves multipleFieldFilters="0": a second filter of the other family on one field is the engine's
    // "more than one filter" refusal (pivot-engine-unmeasured), never a silent replace.
    expect(p.def.multipleFieldFilters).toBe(false)
    expect(refusal(apply(p, first, { kind: 'filter', field: PRODUCT, filter: valueFilter(PRODUCT, 1000) }))).toBe('pivot-engine-unmeasured')
    // With multiple filters per field allowed, the engine computes one label and one value filter on a field: both kept.
    const multi = { ...p, def: { ...p.def, multipleFieldFilters: true } }
    const both = ok(apply(multi, first, { kind: 'filter', field: PRODUCT, filter: valueFilter(PRODUCT, 1000) }))
    expect(both.filters).toEqual([labelFilter(PRODUCT, 'B'), valueFilter(PRODUCT, 1000)])
    expect(computePivot(p.store, p.cache, multi.def, both, { date1904: false }).ok).toBe(true)
    // The filter must name the delta's field.
    expect(refusal(apply(p, view, { kind: 'filter', field: REGION, filter: labelFilter(PRODUCT, 'B') }))).toBe('pivot-delta-invalid')
  })

  it('clearFilters removes the field\'s filters and its hidden items', async () => {
    const p = await baseline()
    const view = initialViewState(p.def, p.cache)
    const filtered = ok(apply(p, ok(apply(p, ok(apply(p, view, { kind: 'filter', field: PRODUCT, filter: labelFilter(PRODUCT, 'B') })),
      { kind: 'filter', field: REGION, filter: valueFilter(REGION, 50000) })), { kind: 'items', field: PRODUCT, hidden: [3] }))
    const cleared = ok(apply(p, filtered, { kind: 'clearFilters', field: PRODUCT }))
    expect(cleared.filters).toEqual([valueFilter(REGION, 50000)])
    expect(cleared.fields[PRODUCT]!.hidden).toEqual(view.fields[PRODUCT]!.hidden)
  })

  it('sort sets sortType and autoSortScope', async () => {
    const p = await baseline()
    const view = initialViewState(p.def, p.cache)
    const label = ok(apply(p, view, { kind: 'sort', field: PRODUCT, order: 'descending' }))
    expect(label.fields[PRODUCT]!.sortType).toBe('descending')
    expect(label.fields[PRODUCT]!.autoSortScope).toBeUndefined()
    const byValue = ok(apply(p, view, { kind: 'sort', field: REGION, order: 'descending', byData: 0 }))
    expect(byValue.fields[REGION]!.sortType).toBe('descending')
    // The scope Excel saves for "sort by Sum of Sales" (sort-value-top-level).
    const saved = (await load('sort-value-top-level'))[0]!.def.fields[REGION]!.autoSortScope
    expect(byValue.fields[REGION]!.autoSortScope).toEqual(saved)
    // A label sort after a value sort drops the scope.
    expect(ok(apply(p, byValue, { kind: 'sort', field: REGION, order: 'ascending' })).fields[REGION]!.autoSortScope).toBeUndefined()
    expect(refusal(apply(p, view, { kind: 'sort', field: REGION, order: 'descending', byData: 2 }))).toBe('pivot-delta-invalid')
    expect(refusal(apply(p, view, { kind: 'sort', field: 3, order: 'descending' }))).toBe('pivot-delta-invalid')
  })
})
