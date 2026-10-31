/**
 * The pivot view state (spec §5.2, §6.7): what a user can change about a pivot,
 * separated from the saved definition. The engine lays out a view; the saved
 * state is `initialViewState(def, cache)`. Lazy chunk `pivots` (Stage 2: library only).
 */
import {
  dataFieldSource,
  type PivotArea, type PivotCacheDefinition, type PivotDataFunction, type PivotFieldSortType, type PivotFilter,
  type PivotShowAsX14, type PivotShowDataAs, type PivotSubtotalFunction, type PivotTableDefinition,
} from '../model'

/** One pivot field's state. `hidden` and `collapsed` are indexed by pivot item (the field's `items` list; the
 * cache's shared-item order when the definition lists no items). */
export interface FieldState { hidden: readonly boolean[]; collapsed: readonly boolean[]; sortType: PivotFieldSortType; autoSortScope?: PivotArea;
  defaultSubtotal: boolean; subtotals: readonly PivotSubtotalFunction[]; subtotalTop: boolean; compact: boolean; outline: boolean;
  insertBlankRow: boolean; fillDownLabels: boolean; showAll: boolean }
/** A page (report filter) field: one selected pivot item, or every item not hidden in its FieldState. */
export interface PageState { field: number; selected: readonly number[] | 'all' }
/** A data field. `name` is the definition's `dataField@name`; '' when absent (Excel's default name, an implicit caption). */
export interface DataFieldState { field: number; fn: PivotDataFunction; showAs: PivotShowDataAs | PivotShowAsX14; baseField: number; baseItem: number; name: string; numFmtId?: number }
export interface PivotViewState { rows: readonly number[]; cols: readonly number[]; pages: readonly PageState[]; data: readonly DataFieldState[]; fields: readonly FieldState[]; filters: readonly PivotFilter[] }

/** The view the saved definition describes. */
export function initialViewState(def: PivotTableDefinition, cache: PivotCacheDefinition): PivotViewState {
  const fields = def.fields.map((field, index): FieldState => {
    // A field without an items list (other writers) shows its cache items in shared-item order, none hidden.
    const count = field.items ? field.items.length : cache.fields[index]?.sharedItems?.items.length ?? 0
    const items = field.items
    return {
      hidden: Array.from({ length: count }, (_, k) => items?.[k]?.hidden ?? false),
      collapsed: Array.from({ length: count }, (_, k) => items?.[k] ? !items[k]!.showDetails : false),
      sortType: field.sortType,
      ...(field.autoSortScope ? { autoSortScope: field.autoSortScope } : {}),
      defaultSubtotal: field.defaultSubtotal,
      subtotals: [...field.subtotals],
      subtotalTop: field.subtotalTop,
      compact: field.compact,
      outline: field.outline,
      insertBlankRow: field.insertBlankRow,
      fillDownLabels: field.x14?.fillDownLabels ?? def.x14?.fillDownLabelsDefault ?? false,
      showAll: field.showAll,
    }
  })
  return {
    rows: [...def.rowFields],
    cols: [...def.colFields],
    pages: def.pageFields.map(page => ({ field: page.fld, selected: page.item !== undefined ? [page.item] : 'all' })),
    data: def.dataFields.map(data => ({
      field: dataFieldSource(data),
      fn: data.subtotal,
      showAs: data.x14?.pivotShowAs ?? data.showDataAs,
      baseField: data.baseField,
      baseItem: data.baseItem,
      name: data.name ?? '',
      ...(data.numFmtId !== undefined ? { numFmtId: data.numFmtId } : {}),
    })),
    fields,
    filters: [...def.filters],
  }
}
