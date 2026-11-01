/**
 * The pivot menus' model (plan Task 28): a field's items with their labels and state, the data fields, the next filter
 * id. Pure and serialisable, so both paths build it where the cache lives: the resident client on the main thread,
 * the worker for a paged or worker-parsed workbook (it rides on the `pivotOpen`/`pivotApply` replies). A number or
 * date item carries its value and number format (`number`); the main thread formats its label (`label` is then the
 * value as plain text until it does). Lazy chunk `pivots`.
 */
import { PIVOT_VALUES_FIELD, type PivotCacheDefinition, type PivotDataFunction, type PivotFieldSortType, type PivotFilter, type PivotShowAsX14, type PivotShowDataAs, type PivotTableDefinition } from '../model'
import { pivotItemLabel } from '../engine/place'
import type { PivotViewState } from '../engine/view'
import { PIVOT_FUNCTION_NAMES } from './functionNames'
import { pivotMenuGroupInfo, pivotMenuGroupPeriods } from './grouping'
import type { PivotDateGroupPeriod } from './delta'

/** One item of a field, by its pivot-item index (the index a delta names). */
export interface PivotMenuItem { index: number; label: string; hidden: boolean; collapsed: boolean; number?: { value: number; numFmtId: number } }
export interface PivotMenuGroupInfo {
  kind?: 'date' | 'number'
  grouped: boolean
  baseField: number
  bounds?: { min: number; max: number }
  date?: { periods: readonly PivotDateGroupPeriod[]; startDate?: number; endDate?: number; autoStart: boolean; autoEnd: boolean; daysInterval: number }
  number?: { start?: number; end?: number; by: number; autoStart: boolean; autoEnd: boolean }
}
export interface PivotMenuField {
  index: number
  sourceOrder: number
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
  /** A page field's effective selection: 'all' only when no data item is hidden. A saved multi-item report filter is
   * (All) plus hidden items (interaction-page-multi); it reads as the items left shown (PR #57 review). */
  page?: readonly number[] | 'all'
  /** A page field saved with "Select multiple items" on (`multipleItemSelectionAllowed`). */
  multiple?: boolean
  calculated?: boolean
  virtualGroup?: boolean
  group?: PivotMenuGroupInfo
  defaultDataFn: 'sum' | 'count'
  permissions: { dragToRow: boolean; dragToCol: boolean; dragToPage: boolean; dragToData: boolean; dragOff: boolean }
}
export interface PivotMenuModel {
  name: string
  date1904?: boolean
  fields: PivotMenuField[]
  rows: readonly number[]
  cols: readonly number[]
  pages: readonly number[]
  valuesCaption: string
  disableFieldList?: boolean
  fieldListSortAscending?: boolean
  canAutoCaptionDataField?: boolean
  data: Array<{ index: number; field: number; name: string; fn: PivotDataFunction; showAs: PivotShowDataAs | PivotShowAsX14; baseField: number; baseItem: number; numFmtId?: number }>
  /** The id a new filter takes (Excel's filter ids are unique in the pivot). */
  nextFilterId: number
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
  const groupPeriods = pivotMenuGroupPeriods(cache)
  const rowPositions = new Map<number, number>(), colPositions = new Map<number, number>()
  view.rows.forEach((field, index) => { if (!rowPositions.has(field)) rowPositions.set(field, index) })
  view.cols.forEach((field, index) => { if (!colPositions.has(field)) colPositions.set(field, index) })
  const pageByField = new Map<number, PivotViewState['pages'][number]>()
  for (const page of view.pages) if (!pageByField.has(page.field)) pageByField.set(page.field, page)
  const axisOf = (field: number): PivotMenuField['axis'] => rowPositions.has(field) ? 'row' : colPositions.has(field) ? 'col' : pageByField.has(field) ? 'page' : 'none'
  const below = (positions: ReadonlyMap<number, number>, fields: readonly number[], field: number) => {
    const at = positions.get(field)
    return at !== undefined && at < fields.length - 1
  }
  const used = [...new Set([...view.rows, ...view.cols, ...view.pages.map(page => page.field)])].filter(field => field !== PIVOT_VALUES_FIELD)
  const usedSet = new Set(used)
  const filtersByField = new Map<number, PivotViewState['filters'][number][]>()
  for (const filter of view.filters) {
    const list = filtersByField.get(filter.fld) ?? []
    list.push(filter)
    filtersByField.set(filter.fld, list)
  }
  const order = [...used, ...def.fields.map((_, index) => index).filter(field => !usedSet.has(field))]
  const fields = order.map((field): PivotMenuField => {
    const state = view.fields[field]
    const list = def.fields[field]?.items
    const indices = list ? list.flatMap((item, k) => item.type === 'data' ? [k] : []) : (state?.hidden ?? []).map((_, k) => k)
    const scope = state?.autoSortScope?.references.find(ref => ref.field === PIVOT_VALUES_FIELD)?.x[0]
    const page = pageByField.get(field)
    return {
      index: field, sourceOrder: field, name: pivotFieldName(def, cache, field), axis: axisOf(field),
      items: indices.map(k => ({ index: k, ...pivotMenuItemLabel(def, cache, field, k, date1904), hidden: state?.hidden[k] ?? false, collapsed: state?.collapsed[k] ?? false })),
      sort: state?.sortType ?? 'manual', ...(scope !== undefined ? { byData: scope } : {}),
      filters: filtersByField.get(field) ?? [], values: valuesOf(cache, field),
      expandable: below(rowPositions, view.rows, field) || below(colPositions, view.cols, field),
      ...(page ? { page: page.selected === 'all' && indices.some(k => state?.hidden[k]) ? indices.filter(k => !state?.hidden[k]) : page.selected } : {}),
      ...(page && def.fields[field]?.multipleItemSelectionAllowed ? { multiple: true } : {}),
      ...(cache.fields[field]?.formula !== undefined ? { calculated: true } : {}),
      ...(cache.fields[field]?.fieldGroup ? { virtualGroup: true } : {}),
      group: pivotMenuGroupInfo(cache, field, date1904, groupPeriods),
      defaultDataFn: cache.fields[field]?.formula !== undefined || (cache.fields[field]?.sharedItems?.containsNumber && !cache.fields[field]?.sharedItems?.containsString && !cache.fields[field]?.sharedItems?.containsDate) ? 'sum' : 'count',
      permissions: { dragToRow: Boolean(def.fields[field]?.dragToRow), dragToCol: Boolean(def.fields[field]?.dragToCol), dragToPage: Boolean(def.fields[field]?.dragToPage), dragToData: Boolean(def.fields[field]?.dragToData), dragOff: Boolean(def.fields[field]?.dragOff) },
    }
  })
  const data = view.data.map((entry, index) => ({ index, field: entry.field, name: entry.name || `${PIVOT_FUNCTION_NAMES[entry.fn] ?? entry.fn} of ${pivotFieldName(def, cache, entry.field)}`, fn: entry.fn, showAs: entry.showAs, baseField: entry.baseField, baseItem: entry.baseItem, ...(entry.numFmtId !== undefined ? { numFmtId: entry.numFmtId } : {}) }))
  const nextFilterId = view.filters.reduce((max, filter) => Math.max(max, filter.id), 0) + 1
  return { name: String(def.name ?? ''), date1904, fields, rows: view.rows, cols: view.cols, pages: view.pages.map(page => page.field), valuesCaption: def.dataCaption || 'Values', ...(def.disableFieldList ? { disableFieldList: true } : {}), ...(def.fieldListSortAscending !== undefined ? { fieldListSortAscending: def.fieldListSortAscending } : {}), canAutoCaptionDataField: true, data, nextFilterId }
}
