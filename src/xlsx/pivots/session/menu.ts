/**
 * The pivot menus' model (plan Task 28): a field's items with their labels and state, the data fields, the next filter
 * id. Pure and serialisable, so both paths build it where the cache lives: the resident client on the main thread,
 * the worker for a paged or worker-parsed workbook (it rides on the `pivotOpen`/`pivotApply` replies). A number or
 * date item carries its value and number format (`number`); the main thread formats its label (`label` is then the
 * value as plain text until it does). Lazy chunk `pivots`.
 */
import { PIVOT_VALUES_FIELD, type PivotCacheDefinition, type PivotFieldSortType, type PivotFilter, type PivotTableDefinition } from '../model'
import { pivotItemLabel } from '../engine/place'
import type { PivotViewState } from '../engine/view'

/** One item of a field, by its pivot-item index (the index a delta names). */
export interface PivotMenuItem { index: number; label: string; hidden: boolean; collapsed: boolean; number?: { value: number; numFmtId: number } }
export interface PivotMenuField {
  index: number
  name: string
  axis: 'row' | 'col' | 'page' | 'none'
  /** The field's data items (subtotal items are not listed). */
  items: PivotMenuItem[]
  sort: PivotFieldSortType
  /** The data field a value sort follows. */
  byData?: number
  filters: readonly PivotFilter[]
  /** What the label/date filter compares: the cache field's item types. */
  values: 'text' | 'number' | 'date'
  /** A level lies below this field on its axis (collapse entire field applies). */
  expandable: boolean
  /** A page field's selection. */
  page?: readonly number[] | 'all'
}
export interface PivotMenuModel {
  name: string
  fields: PivotMenuField[]
  data: Array<{ index: number; name: string }>
  /** The id a new filter takes (Excel's filter ids are unique in the pivot). */
  nextFilterId: number
}

const FUNCTION_NAMES: Record<string, string> = {
  sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count', stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp',
}
export function pivotFieldName(def: PivotTableDefinition, cache: PivotCacheDefinition | undefined, field: number): string {
  if (field === PIVOT_VALUES_FIELD) return def.dataCaption || 'Values'
  return def.fields[field]?.name ?? cache?.fields[field]?.name ?? `Field ${field + 1}`
}
/** An item's label as text, or its number with the format to show it in (a custom format id falls back: dates to the
 * built-in short date, numbers to General, since no format code travels with the cache). */
export function pivotMenuItemLabel(def: PivotTableDefinition, cache: PivotCacheDefinition, field: number, item: number, date1904: boolean): Pick<PivotMenuItem, 'label' | 'number'> {
  const label = pivotItemLabel(def, cache, field, item, date1904)
  if (label.value.kind === 'empty') return { label: '' }
  if (label.value.kind !== 'number') return { label: label.value.value }
  const shared = cache.fields[field]?.sharedItems
  const dates = Boolean(shared?.containsDate && !shared.containsString && !shared.containsNumber)
  const numFmtId = label.numFmtId !== undefined && label.numFmtId < 164 ? label.numFmtId : dates ? 14 : 0
  return { label: String(label.value.value), number: { value: label.value.value, numFmtId } }
}
function valuesOf(cache: PivotCacheDefinition, field: number): PivotMenuField['values'] {
  const shared = cache.fields[field]?.sharedItems
  if (!shared || shared.containsString || shared.containsMixedTypes) return 'text'
  if (shared.containsDate && !shared.containsNumber) return 'date'
  if (shared.containsNumber && !shared.containsDate) return 'number'
  return 'text'
}
export function pivotMenuModel(def: PivotTableDefinition, cache: PivotCacheDefinition, view: PivotViewState, date1904: boolean): PivotMenuModel {
  const axisOf = (field: number): PivotMenuField['axis'] => view.rows.includes(field) ? 'row' : view.cols.includes(field) ? 'col' : view.pages.some(page => page.field === field) ? 'page' : 'none'
  const below = (fields: readonly number[], field: number) => { const at = fields.indexOf(field); return at >= 0 && fields.slice(at + 1).length > 0 }
  const used = [...new Set([...view.rows, ...view.cols, ...view.pages.map(page => page.field)])].filter(field => field !== PIVOT_VALUES_FIELD)
  const fields = used.map((field): PivotMenuField => {
    const state = view.fields[field]
    const list = def.fields[field]?.items
    const indices = list ? list.flatMap((item, k) => item.type === 'data' ? [k] : []) : (state?.hidden ?? []).map((_, k) => k)
    const scope = state?.autoSortScope?.references.find(ref => ref.field === PIVOT_VALUES_FIELD)?.x[0]
    const page = view.pages.find(entry => entry.field === field)
    return {
      index: field, name: pivotFieldName(def, cache, field), axis: axisOf(field),
      items: indices.map(k => ({ index: k, ...pivotMenuItemLabel(def, cache, field, k, date1904), hidden: state?.hidden[k] ?? false, collapsed: state?.collapsed[k] ?? false })),
      sort: state?.sortType ?? 'manual', ...(scope !== undefined ? { byData: scope } : {}),
      filters: view.filters.filter(filter => filter.fld === field), values: valuesOf(cache, field),
      expandable: below(view.rows, field) || below(view.cols, field),
      ...(page ? { page: page.selected } : {}),
    }
  })
  const data = view.data.map((entry, index) => ({ index, name: entry.name || `${FUNCTION_NAMES[entry.fn] ?? entry.fn} of ${pivotFieldName(def, cache, entry.field)}` }))
  const nextFilterId = view.filters.reduce((max, filter) => Math.max(max, filter.id), 0) + 1
  return { name: String(def.name ?? ''), fields, data, nextFilterId }
}
