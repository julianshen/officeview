/**
 * Stage 3 view deltas (plan Task 25, spec §6.7, §12): one UI interaction applied to a pivot's view state. Pure: the
 * input view is never mutated, unchanged fields keep their objects. Lazy chunk `pivots`.
 *
 * `PivotDelta` is the plan's declared union verbatim (Ruling S3-A: Task 24's manifest JSON uses this shape; any change
 * goes through a ruling). Item indices are pivot-item indices (the field's `items` list; the cache's shared-item order
 * when the definition lists none), as in `FieldState`.
 *
 * Rulings recorded here (Task 25):
 * - Only a field's DATA items can be hidden, collapsed or selected; an index past the field's items, or naming a
 *   subtotal item, is `pivot-delta-invalid`. Hiding every data item is refused the same way (Excel's filter list
 *   cannot be confirmed with nothing selected).
 * - `collapseField` and `collapseAll` mark data items only; `collapseAll` covers every row and column field.
 * - `filter`: a filter of the same family on the field (label or date; value or top-10, the engine's two kinds,
 *   axes.ts) is replaced in place. A filter of the other family on a field that already has one refuses with the
 *   engine's "more than one filter" feature, `pivot-engine-unmeasured`, whatever `multipleFieldFilters` says: no
 *   oracle case combines the two families on one field, so neither Excel's replace (one filter per field) nor its
 *   combined evaluation (multiple filters allowed) is measured (§9).
 * - `clearFilters` drops the field's pivot filters and shows its hidden items, and a page field goes back to (All):
 *   Excel's "Clear Filter From" (Task 24 measures it; the kind stays refused until then, session.ts).
 * - `page` sets the page field's selection only; its hidden items are an `items` delta.
 * - `sort` sets `sortType`; `byData` (a data-field index) sets the `autoSortScope` Excel saves for "sort by a value
 *   column" (Task 15 sort-value-top-level), and a sort without `byData` drops it.
 */
import {
  PIVOT_VALUES_FIELD,
  type PivotArea, type PivotCacheDefinition, type PivotDiagnostic, type PivotFilter, type PivotTableDefinition,
} from '../model'
import type { EngineResult } from '../engine/axes'
import { isDateFilter, isLabelFilter } from '../engine/filters'
import type { FieldState, PivotViewState } from '../engine/view'

export type PivotDelta =
  | { kind: 'collapse'; field: number; item: number; collapsed: boolean }
  | { kind: 'collapseField'; field: number; collapsed: boolean }
  | { kind: 'collapseAll'; collapsed: boolean }
  | { kind: 'items'; field: number; hidden: readonly number[] }          // pivot-item indices hidden
  | { kind: 'page'; field: number; selected: readonly number[] | 'all' }
  | { kind: 'filter'; field: number; filter: PivotFilter }               // label | value | date | top-10
  | { kind: 'clearFilters'; field: number }
  | { kind: 'sort'; field: number; order: 'ascending' | 'descending'; byData?: number }
export type PivotDeltaKind = PivotDelta['kind']
export const PIVOT_DELTA_KINDS: readonly PivotDeltaKind[] = Object.freeze(['collapse', 'collapseField', 'collapseAll', 'items', 'page', 'filter', 'clearFilters', 'sort'])

const refuse = (feature: string, message: string): EngineResult<never> => ({ ok: false, diagnostic: { kind: 'unsupported', feature, message } satisfies PivotDiagnostic })
const invalid = (message: string) => refuse('pivot-delta-invalid', message)

/** The data-item indices of a field (its `items` list's `t="data"` entries; every shared item without a list). */
function dataItems(def: PivotTableDefinition, view: PivotViewState, field: number): number[] {
  const items = def.fields[field]?.items
  if (items) return items.flatMap((item, k) => item.type === 'data' ? [k] : [])
  return (view.fields[field]?.hidden ?? []).map((_, k) => k)
}
const withField = (view: PivotViewState, field: number, change: Partial<FieldState>): PivotViewState =>
  ({ ...view, fields: view.fields.map((state, k) => k === field ? { ...state, ...change } : state) })
const setFlags = (flags: readonly boolean[], items: readonly number[], value: boolean) => {
  const out = [...flags]
  for (const k of items) out[k] = value
  return out
}
/** The engine's two filter kinds per field (axes.ts): label or date, value or top-10. */
const family = (filter: PivotFilter) => isLabelFilter(filter) || isDateFilter(filter) ? 'label' : 'value'

/** The `autoSortScope` Excel saves for "sort by data field `data`" (sort-value-top-level). */
function valueScope(data: number): PivotArea {
  return {
    type: 'normal', dataOnly: false, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false, outline: false,
    collapsedLevelsAreSubtotals: false, fieldPosition: 0,
    references: [{ field: PIVOT_VALUES_FIELD, selected: false, byPosition: false, relative: false, defaultSubtotal: false, subtotals: [], x: [data] }],
  }
}

export function applyDelta(view: PivotViewState, delta: PivotDelta, def: PivotTableDefinition, _cache: PivotCacheDefinition): EngineResult<PivotViewState> {
  if (delta.kind === 'collapseAll') {
    const axis = new Set([...view.rows, ...view.cols].filter(field => field !== PIVOT_VALUES_FIELD))
    return { ok: true, value: { ...view, fields: view.fields.map((state, field) => axis.has(field) ? { ...state, collapsed: setFlags(state.collapsed, dataItems(def, view, field), delta.collapsed) } : state) } }
  }
  const field = delta.field
  const state = view.fields[field]
  if (!Number.isInteger(field) || field < 0 || !state) return invalid(`field ${field} is not a pivot field`)
  const data = dataItems(def, view, field)
  const isData = new Set(data)
  const badItem = (items: readonly number[]) => items.find(item => !isData.has(item))
  const onAxis = view.rows.includes(field) || view.cols.includes(field)
  switch (delta.kind) {
    case 'collapse':
      if (!isData.has(delta.item)) return invalid(`item ${delta.item} is not a data item of field ${field}`)
      return { ok: true, value: withField(view, field, { collapsed: setFlags(state.collapsed, [delta.item], delta.collapsed) }) }
    case 'collapseField':
      return { ok: true, value: withField(view, field, { collapsed: setFlags(state.collapsed, data, delta.collapsed) }) }
    case 'items': {
      const bad = badItem(delta.hidden)
      if (bad !== undefined) return invalid(`item ${bad} is not a data item of field ${field}`)
      const hidden = new Set(delta.hidden)
      if (data.length && data.every(item => hidden.has(item))) return invalid(`every item of field ${field} would be hidden`)
      return { ok: true, value: withField(view, field, { hidden: state.hidden.map((_, k) => isData.has(k) ? hidden.has(k) : state.hidden[k]!) }) }
    }
    case 'page': {
      const at = view.pages.findIndex(page => page.field === field)
      if (at < 0) return invalid(`field ${field} is not a page field`)
      if (delta.selected !== 'all') {
        if (!delta.selected.length) return invalid(`page field ${field} would have no item selected`)
        const bad = badItem(delta.selected)
        if (bad !== undefined) return invalid(`item ${bad} is not a data item of field ${field}`)
      }
      const selected = delta.selected === 'all' ? 'all' as const : [...new Set(delta.selected)]
      return { ok: true, value: { ...view, pages: view.pages.map((page, k) => k === at ? { ...page, selected } : page) } }
    }
    case 'filter': {
      if (delta.filter.fld !== field) return invalid(`the filter names field ${delta.filter.fld}, not ${field}`)
      const kind = family(delta.filter)
      const same = view.filters.findIndex(filter => filter.fld === field && family(filter) === kind)
      if (same >= 0) return { ok: true, value: { ...view, filters: view.filters.map((filter, k) => k === same ? delta.filter : filter) } }
      if (view.filters.some(filter => filter.fld === field))
        return refuse('pivot-engine-unmeasured', `field ${field} would have a label or date filter and a value or top-10 filter; combining them is not measured`)
      return { ok: true, value: { ...view, filters: [...view.filters, delta.filter] } }
    }
    case 'clearFilters': {
      const filters = view.filters.filter(filter => filter.fld !== field)
      const cleared = withField({ ...view, filters }, field, { hidden: state.hidden.map((flag, k) => isData.has(k) ? false : flag) })
      return { ok: true, value: { ...cleared, pages: view.pages.map(page => page.field === field ? { ...page, selected: 'all' as const } : page) } }
    }
    case 'sort': {
      if (!onAxis) return invalid(`field ${field} is not on the row or column axis`)
      if (delta.byData !== undefined && !(Number.isInteger(delta.byData) && delta.byData >= 0 && delta.byData < view.data.length))
        return invalid(`data field ${delta.byData} does not exist`)
      const { autoSortScope: _scope, ...rest } = state
      const next: FieldState = { ...rest, sortType: delta.order, ...(delta.byData !== undefined ? { autoSortScope: valueScope(delta.byData) } : {}) }
      return { ok: true, value: { ...view, fields: view.fields.map((s, k) => k === field ? next : s) } }
    }
  }
}
