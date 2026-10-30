/**
 * Pivot axes (spec §5.2): the row and column lines of a view, in the shape of the saved `rowItems`/`colItems`
 * (`PivotLine`), and each record's leaf line. Lazy chunk `pivots` (Stage 2: library only).
 *
 * What the lines follow, read from Excel's saved layouts (Stage 1 references tests/fixtures/xlsx-pivots/pivot-*.xlsx,
 * Excel 16.114; Excel-made corpus files):
 * - items keep the field's `items` order (manual sort); an item shows when it is not hidden and has records left
 *   after every field's hidden items and page selections, or when its field has `showAll` (corpus
 *   poi-ExcelPivotTableSample Tabelle3: a page selection drops a row item it leaves without records);
 * - a row field with `outline` gives each parent item its own line (compact, outline); without it (tabular) the
 *   parent shares its first child's line (pivot-layout-tabular); column parents never have their own line;
 * - subtotals follow each item of a field that has a deeper non-Values field: on the parent's own line when the
 *   field is outline with `subtotalTop` and only the default subtotal (pivot-layout-compact), else after the
 *   children (pivot-subtotals-bottom, pivot-layout-tabular, lo-calc-Pivot3_Column_Grand_Subtotals), one line per
 *   data field when the Values field lies deeper on the same axis (pivot-values-on-rows);
 * - a collapsed item (`sd="0"`) with a deeper non-Values field is one line and hides its children (pivot-collapsed-item);
 *   with the Values field deeper (last on rows, outline form) it is its own line, then one line per data field with
 *   the levels between left unfilled (-1) and `r` at the Values level, and no subtotal lines (Task 15
 *   layout-collapsed-item-above-values: <i><x/></i>, <i r="2"><x/></i>, <i r="2" i="1"><x v="1"/></i>);
 * - an item shown only by its field's `showAll` (no records under it) lists the next field's visible items under it
 *   (Task 15 quirk-cross-hidden-item-showall: Central, whose records are all a hidden product, keeps Apples, Bananas
 *   and Cherries), refused when one of those items has no passing record anywhere (not measured);
 * - the Values field is an axis level whose items are the data fields; grand totals are one line per data field
 *   when it is on the axis (pivot-values-on-rows), and need a non-Values field on the axis (lo-calc-two-data-fields);
 *   `rowGrandTotals` gates the row axis's grand line and `colGrandTotals` the column axis's (lo-calc-Pivot2_Row_Compact,
 *   lo-calc-Pivot3_Column_Grand_Subtotals);
 * - each line's `first` (`r`) is the levels it shares with the previous non-grand line (an unfilled level counts as
 *   shared), at most all but its last.
 * - another field's hidden item or page selection drops an item it leaves without records (Task 15
 *   quirk-cross-hidden-item, quirk-page-single, quirk-page-multiple), unless its field has `showAll`
 *   (quirk-page-single-showall, quirk-cross-hidden-item-showall).
 * Unmeasured (Task 15 cases decide): named subtotals (one line per function in schema order, always at the bottom),
 * blank rows (one `blank` line after each item block of a field with a deeper field), a showAll item with no records
 * whose next field has no visible item (its line alone in tabular form).
 * Refused as `pivot-engine-unmeasured` until measured: `mergeItem` outside tabular fields with totals only after
 * outer items, classic layout without a column field beyond one data field with row fields, a hidden item on a field
 * on no axis or page, a collapsed item above Values on columns, in tabular form or above a Values level that is not
 * the last, mixed-type/custom-list/non-ASCII sort order, nested/value-sort ties, cross-field filter visibility with
 * `showAll` and out-of-range group captions. Dynamic date filters use the dedicated feature refusal.
 */
import { PIVOT_VALUES_FIELD, type PivotCacheDefinition, type PivotDiagnostic, type PivotItemType, type PivotSubtotalFunction, type PivotTableDefinition, type PivotValue } from '../model'
import type { PivotLine } from '../savedLayout'
import type { RecordStore } from './records'
import type { FieldState, PivotViewState } from './view'
import { sortItemOrder, itemValuesByPivotIndex } from './sort'
import { groupColumn, NO_GROUP } from './groups'
import { isDateFilter, isDynamicDateFilter, isLabelFilter, isTop10Filter, isValueFilter, matchesDateFilter, matchesLabelFilter, matchesValueFilter } from './filters'
import { aggregateByKey } from './values'

export interface PivotAxes { rows: PivotLine[]; cols: PivotLine[]; rowKey: Uint32Array; colKey: Uint32Array }   // per record: its leaf line, or NO_LINE
export const NO_LINE = 0xffffffff
export type EngineResult<T> = { ok: true; value: T } | { ok: false; diagnostic: PivotDiagnostic }

const fail = (kind: PivotDiagnostic['kind'], feature: string, message: string): { ok: false; diagnostic: PivotDiagnostic } => ({ ok: false, diagnostic: { kind, feature, message } })
/** Constructs no task has implemented yet. */
const pending = (message: string) => fail('unsupported', 'pivot-engine-pending', message)
const invalid = (message: string) => fail('parse', 'pivot-definition-invalid', message)
/** Behaviour Excel has not been measured on yet (Task 15 adds the cases); the pivot stays display-only. */
const unmeasured = (message: string) => fail('unsupported', 'pivot-engine-unmeasured', message)

/** The subtotal lines a field's items get: its named functions, else the default subtotal, else none. */
export function subtotalTypes(field: FieldState): PivotItemType[] {
  if (field.subtotals.length) return [...field.subtotals] as PivotSubtotalFunction[]
  return field.defaultSubtotal ? ['default'] : []
}

/** Whether a field's item subtotal sits on the item's own line (compact or outline form, subtotalTop, default subtotal
 * only, and no Values field deeper on the axis). */
function subtotalOnTop(fields: readonly number[], level: number, field: FieldState): boolean {
  const values = fields.indexOf(PIVOT_VALUES_FIELD)
  return field.outline && field.subtotalTop && field.subtotals.length === 0 && field.defaultSubtotal && !(values > level)
}
const hasFieldAfter = (fields: readonly number[], level: number) => fields.slice(level + 1).some(field => field !== PIVOT_VALUES_FIELD)

/** Whether a line's cells hold values: leaf, subtotal and grand lines do; a parent's own line only when it is
 * collapsed or carries its subtotal on top; blank lines and lines above the Values level never. */
export function lineCarriesValues(view: PivotViewState, axis: 'rows' | 'cols', line: PivotLine): boolean {
  if (line.type === 'blank') return false
  if (line.type !== 'data') return true
  const fields = axis === 'rows' ? view.rows : view.cols
  if (line.items.length >= fields.length) return true
  const level = line.items.length - 1, fieldIndex = fields[level]!
  if (fieldIndex === PIVOT_VALUES_FIELD || fields.indexOf(PIVOT_VALUES_FIELD) > level) return false
  const field = view.fields[fieldIndex]!
  if (!hasFieldAfter(fields, level) || !field.outline) return true
  return field.collapsed[line.items[level]!] === true || subtotalOnTop(fields, level, field)
}

interface Node { children: Map<number, Node>; leaf: number }
const newNode = (): Node => ({ children: new Map(), leaf: NO_LINE })

/** Per used field: cache shared index → pivot item index (-1 for none) and the pivot items in display order. */
interface ItemMap { toPivot: Int32Array; order: number[]; size: number; recordItems?: Uint32Array; grouped?: boolean }
function itemMap(def: PivotTableDefinition, cache: PivotCacheDefinition, field: number): ItemMap | string {
  const shared = cache.fields[field]?.sharedItems
  if (!shared || shared.inline) return `field ${field} has no shared items`
  const items = def.fields[field]?.items
  const toPivot = new Int32Array(shared.items.length).fill(-1)
  const order: number[] = []
  if (!items) shared.items.forEach((_, k) => { toPivot[k] = k; order.push(k) })
  else for (let k = 0; k < items.length; k++) {
    const item = items[k]!
    if (item.type !== 'data') continue
    if (item.x === undefined || item.x < 0 || item.x >= shared.items.length) return `field ${field} item ${k} points outside its shared items`
    toPivot[item.x] = k
    order.push(k)
  }
  return { toPivot, order, size: items ? items.length : shared.items.length }
}

export function computeAxes(store: RecordStore, cache: PivotCacheDefinition, def: PivotTableDefinition, view: PivotViewState): EngineResult<PivotAxes> {
  if (view.filters.some(isDynamicDateFilter)) return fail('unsupported', 'pivot-filter-dynamic-date', 'dynamic date filters depend on the current date and are not implemented')
  // Classic layout with no column-axis field is measured for one data field with row fields
  // (layout-classic-single-data; placeCells' geometry). A Values field on columns keeps the
  // observed two-data-field geometry (lo-calc-two-data-fields, layout-classic-two-data).
  if (def.gridDropZones && view.cols.length === 0 && !(view.data.length === 1 && view.rows.length > 0))
    return unmeasured('classic pivot layout (gridDropZones) without a column field is measured for one data field with row fields only')
  // A calculated field is evaluated as data only (values.ts, calculated.ts): Excel does not place one on an axis or page.
  for (const field of [...view.rows, ...view.cols, ...view.pages.map(page => page.field)]) {
    if (cache.fields[field]?.formula !== undefined) return fail('unsupported', 'pivot-calculated-field', `calculated field ${field} on an axis or page is not supported`)
  }
  for (const field of [...view.rows, ...view.cols, ...view.pages.map(page => page.field), ...view.data.map(data => data.field)]) {
    const cacheField = cache.fields[field]
    if (cacheField && cacheField.formula === undefined && !cacheField.databaseField && !cacheField.fieldGroup) return pending(`field ${field} has no records and no formula`)
  }
  const axisFields = [...view.rows, ...view.cols].filter(field => field !== PIVOT_VALUES_FIELD)
  if ([...view.rows, ...view.cols].filter(field => field === PIVOT_VALUES_FIELD).length > 1) return invalid('the Values field is on more than one axis level')
  const pageFields = view.pages.map(page => page.field)
  const axisUsed = new Set([...axisFields, ...pageFields])
  if (axisUsed.size !== axisFields.length + pageFields.length) return invalid('a field is used on more than one axis')
  // Whether a hidden item on a field on no axis or page filters records is unmeasured (a slicer on a field outside the
  // pivot may store its selection that way): fail closed until Task 15 has an Excel case.
  const filterFields = new Set(view.filters.map(filter => filter.fld))
  if (view.filters.some(filter => axisFields.some(target => target !== filter.fld && view.fields[target]?.showAll)))
    return unmeasured('cross-field filter visibility with showAll is not measured')
  const used = new Set([...axisUsed, ...filterFields])
  const offAxis = view.fields.findIndex((field, index) => field.hidden.includes(true) && !axisUsed.has(index) && !filterFields.has(index))
  if (offAxis >= 0) return unmeasured(`field ${offAxis} is on no axis but has hidden items`)
  for (const field of used) {
    const state = view.fields[field], cacheField = cache.fields[field]
    if (!state || !cacheField || !def.fields[field]) return invalid(`field ${field} does not exist`)
    if (def.fields[field]!.items?.some(item => item.calculated) || cacheField.sharedItems?.items.some(item => item.calculated)) return fail('unsupported', 'pivot-calculated-item', `field ${field} has calculated items (no Excel oracle, spec §5.6)`)
    // Sorting changes the layout only on rows and columns (a page field's order is its dropdown's).
    if (!axisFields.includes(field)) continue
  }

  // Record filter: the hidden items of every row, column and page field, and each page's selection.
  const maps = new Map<number, ItemMap>()
  const filters: Array<{ field: number; map: ItemMap; selected: ReadonlySet<number> | 'all'; state: FieldState }> = []
  for (const field of used) {
    const column = store.columns[field]
    const map = fieldItemMap(store, def, cache, field)
    if (!map.ok) return map
    if (!column && !cache.fields[field]?.fieldGroup) return invalid(`field ${field} has no record column`)
    maps.set(field, map.value)
    const state = view.fields[field]!
    const selected = view.pages.find(page => page.field === field)?.selected ?? 'all'
    if (selected === 'all' && !state.hidden.includes(true)) continue
    if (column && column.kind !== 'shared' && !cache.fields[field]?.fieldGroup) return unmeasured(`inline field ${field} on an axis needs a measured item map`)
    filters.push({ field, map: map.value, selected: selected === 'all' ? 'all' : new Set(selected), state })
  }
  const passes = new Uint8Array(store.count).fill(1)
  const allowedByField = new Map<number, Set<number>>()
  const constrainItems = (field: number, next: Set<number>) => {
    const previous = allowedByField.get(field)
    if (previous) {
      for (const item of previous) if (!next.has(item)) previous.delete(item)
    } else allowedByField.set(field, next)
  }
  for (const entry of filters) for (let r = 0; r < store.count; r++) {
    const item = pivotItemAt(entry.map, r)
    if (item === undefined || item < 0 || entry.state.hidden[item] || (entry.selected !== 'all' && !entry.selected.has(item))) passes[r] = 0
  }

  // Apply pivot filters in their declared evaluation order. Each accepted filter narrows
  // the input to the next filter, while the field map provides the aggregate bucket.
  for (const filter of [...view.filters].sort((a, b) => a.evalOrder - b.evalOrder)) {
    const map = maps.get(filter.fld)
    if (!map) return invalid(`filter field ${filter.fld} does not exist`)
    if (isLabelFilter(filter) && def.fields[filter.fld]?.items?.some(item => item.name !== undefined))
      return unmeasured(`renamed item captions for label filter field ${filter.fld} are not measured`)
    const labels = itemValuesByPivotIndex(fieldValues(cache, filter.fld), def.fields[filter.fld]?.items)
    if (isLabelFilter(filter) || isDateFilter(filter)) {
      const allowed = new Set<number>()
      for (const item of map.order) {
        const decision = isLabelFilter(filter) ? matchesLabelFilter(filter, labels[item]) : matchesDateFilter(filter, labels[item])
        if (decision === undefined) return unmeasured(`${filter.type} criteria or item type for field ${filter.fld} are not measured`)
        if (decision) allowed.add(item)
      }
      constrainItems(filter.fld, allowed)
      for (let r = 0; r < store.count; r++) if (passes[r]) {
        const item = pivotItemAt(map, r)
        if (item === undefined) return invalid(`record ${r} has no item for filter field ${filter.fld}`)
        if (!allowed.has(item)) passes[r] = 0
      }
      continue
    }
    if (isValueFilter(filter) || isTop10Filter(filter)) {
      const dataIndex = filter.iMeasureFld
      if (dataIndex === undefined || !view.data[dataIndex]) return invalid(`filter ${filter.id} names a missing data field`)
      const axis = view.rows.includes(filter.fld) ? view.rows : view.cols.includes(filter.fld) ? view.cols : undefined
      if (axis && axis.indexOf(filter.fld) > 0) return unmeasured(`nested ${filter.type} filter on field ${filter.fld} needs a parent-group oracle`)
      const keys = new Uint32Array(store.count)
      for (let r = 0; r < store.count; r++) {
        const item = pivotItemAt(map, r)
        keys[r] = item === undefined || item < 0 ? NO_LINE : item
      }
      if (isTop10Filter(filter) && filter.type !== 'count') return unmeasured(`top-10 ${filter.type} thresholds are not measured`)
      const aggregate = aggregateByKey(store, cache, view, dataIndex, keys, map.size, passes)
      if (!aggregate.ok) return aggregate
      if (axisFields.includes(filter.fld) && view.fields[filter.fld]?.showAll &&
        map.order.some(item => !view.fields[filter.fld]!.hidden[item] && aggregate.value[item] === undefined))
        return unmeasured(`showAll with ${filter.type} filter on field ${filter.fld} has visible items without aggregates`)
      let allowed: Set<number>
      if (isValueFilter(filter)) {
        allowed = new Set()
        for (let item = 0; item < map.size; item++) {
          const decision = matchesValueFilter(filter, aggregate.value[item])
          if (decision === undefined) return unmeasured(`value filter ${filter.type} criteria are not measured`)
          if (decision) allowed.add(item)
        }
      } else {
        const column = filter.autoFilter.columns[0], top10 = column?.top10
        if (!top10) return unmeasured(`missing top-10 criteria for filter ${filter.id} are not measured`)
        if (top10.percent) return unmeasured(`top-10 percent mode for filter ${filter.id} is not measured`)
        const ranked = map.order.map(item => ({ item, value: aggregate.value[item] })).filter((entry): entry is { item: number; value: number } => entry.value !== undefined)
          .sort((a, b) => (top10.top ? b.value - a.value : a.value - b.value) || a.item - b.item)
        if (!Number.isSafeInteger(top10.val) || top10.val <= 0) return invalid(`top-10 item count for filter ${filter.id} is outside its valid range`)
        const count = Math.min(ranked.length, top10.val)
        if (ranked.length > count && ranked[count - 1]!.value === ranked[count]!.value)
          return unmeasured(`top-10 threshold ties for filter ${filter.id} are not measured`)
        const threshold = ranked[count - 1]?.value
        if (top10.filterVal !== undefined && threshold !== undefined && Math.abs(top10.filterVal - threshold) > Math.max(1, Math.abs(threshold)) * 1e-12)
          return unmeasured(`top-10 saved threshold for filter ${filter.id} disagrees with its item aggregates`)
        allowed = new Set(ranked.filter(entry => threshold !== undefined && (top10.top ? entry.value >= threshold : entry.value <= threshold)).map(entry => entry.item))
      }
      constrainItems(filter.fld, allowed)
      for (let r = 0; r < store.count; r++) if (passes[r] && !allowed.has(keys[r]!)) passes[r] = 0
      continue
    }
    return unmeasured(`filter ${filter.type} is not implemented`)
  }

  const sortedOrders = new Map<number, readonly number[]>()
  for (const field of axisFields) {
    const state = view.fields[field]!
    const map = maps.get(field)!
    if ((state.sortType !== 'manual' || state.autoSortScope) && cache.fields[field]?.fieldGroup)
      return unmeasured(`sorted group field ${field} has no measured group-order rule`)
    if (state.autoSortScope) {
      const fields = view.rows.includes(field) ? view.rows : view.cols
      const level = fields.indexOf(field)
      if (level !== 0) return unmeasured(`autoSortScope for nested field ${field} needs an Excel parent-group oracle`)
      const references = state.autoSortScope.references
      const values = references.filter(reference => reference.field === PIVOT_VALUES_FIELD)
      if (references.length !== 1 || values.length !== 1 || values[0]!.x.length !== 1)
        return unmeasured(`autoSortScope for field ${field} has ambiguous or non-Values pivotArea references`)
      const dataIndex = values[0]!.x[0]!
      if (!view.data[dataIndex]) return unmeasured(`autoSortScope for field ${field} references a missing data field`)
      if (state.sortType === 'manual') return unmeasured(`autoSortScope for manually sorted field ${field} is not measured`)
      const keys = new Uint32Array(store.count)
      for (let record = 0; record < store.count; record++) {
        const raw = map.recordItems![record]!
        const item = map.grouped && raw === NO_GROUP ? -1 : map.toPivot[raw]
        if (passes[record] && (item === undefined || item < 0)) return unmeasured(`autoSortScope field ${field} has an unmapped record item`)
        keys[record] = item === undefined || item < 0 ? NO_LINE : item
      }
      const aggregate = aggregateByKey(store, cache, view, dataIndex, keys, map.size, passes)
      if (!aggregate.ok) return aggregate
      const scored = map.order.map(item => ({ item, value: aggregate.value[item] }))
        .filter((entry): entry is { item: number; value: number } => entry.value !== undefined)
        .sort((a, b) => (state.sortType === 'descending' ? b.value - a.value : a.value - b.value) || a.item - b.item)
      for (let i = 1; i < scored.length; i++) if (scored[i - 1]!.value === scored[i]!.value)
        return unmeasured(`autoSortScope ties for field ${field} are not measured`)
      if (view.fields[field]!.showAll && scored.length < map.order.filter(item => !view.fields[field]!.hidden[item]).length)
        return unmeasured(`autoSortScope placement of no-data items for field ${field} is not measured`)
      sortedOrders.set(field, scored.map(entry => entry.item))
      continue
    }
    if (state.sortType === 'manual') continue
    const sorted = sortItemOrder(map.order, itemValuesByPivotIndex(fieldValues(cache, field), def.fields[field]?.items), state.sortType, field)
    if (!sorted.ok) return sorted
    sortedOrders.set(field, sorted.value)
  }

  const rows = axis(store, def, view, 'rows', maps, passes, sortedOrders, allowedByField)
  if (!rows.ok) return rows
  const cols = axis(store, def, view, 'cols', maps, passes, sortedOrders, allowedByField)
  if (!cols.ok) return cols
  if (def.mergeItem && !mergeMeasured(view, rows.value.lines, cols.value.lines))
    return unmeasured('merged item labels (mergeItem) are measured for tabular axis fields with totals only after outer items')
  return { ok: true, value: { rows: rows.value.lines, cols: cols.value.lines, rowKey: rows.value.key, colKey: cols.value.key } }
}

/** The `mergeItem` shape Task 15 layout-merge-labels measures: tabular row and column fields with no repeated labels,
 * blank rows or collapsed items, the Values field (if any) last on columns, row totals only after the outermost
 * items (default subtotal) and column totals only as the grand total. */
function mergeMeasured(view: PivotViewState, rows: readonly PivotLine[], cols: readonly PivotLine[]): boolean {
  if (view.rows.includes(PIVOT_VALUES_FIELD)) return false
  if (view.cols.includes(PIVOT_VALUES_FIELD) && view.cols.indexOf(PIVOT_VALUES_FIELD) !== view.cols.length - 1) return false
  for (const field of [...view.rows, ...view.cols]) {
    if (field === PIVOT_VALUES_FIELD) continue
    const state = view.fields[field]!
    if (state.outline || state.fillDownLabels || state.insertBlankRow || state.collapsed.includes(true)) return false
  }
  return rows.every(line => line.type === 'data' || line.type === 'grand' || (line.type === 'default' && line.items.length === 1))
    && cols.every(line => line.type === 'data' || line.type === 'grand')
}

function fieldValues(cache: PivotCacheDefinition, field: number): PivotValue[] {
  return cache.fields[field]?.fieldGroup?.groupItems ?? cache.fields[field]?.sharedItems?.items ?? []
}

function fieldItemMap(store: RecordStore, def: PivotTableDefinition, cache: PivotCacheDefinition, field: number): EngineResult<ItemMap> {
  const grouped = cache.fields[field]?.fieldGroup
  if (!grouped) {
    const map = itemMap(def, cache, field)
    const column = store.columns[field]
    if (typeof map === 'string') return invalid(map)
    if (!column || column.kind !== 'shared') return unmeasured(`inline field ${field} needs a measured item map`)
    return { ok: true, value: { ...map, recordItems: column.index } }
  }
  const groupItems = grouped.groupItems ?? []
  if (!groupItems.length) return invalid(`group field ${field} has no groupItems`)
  const generated = groupColumn(store, cache, field)
  if (!generated.ok) return generated
  const items = def.fields[field]?.items
  const toPivot = new Int32Array(groupItems.length).fill(-1), order: number[] = []
  if (!items) groupItems.forEach((_, index) => { toPivot[index] = index; order.push(index) })
  else for (let index = 0; index < items.length; index++) {
    const item = items[index]!
    if (item.type !== 'data') continue
    if (item.x === undefined || item.x < 0 || item.x >= groupItems.length) return invalid(`group field ${field} item ${index} points outside groupItems`)
    toPivot[item.x] = index; order.push(index)
  }
  return { ok: true, value: { toPivot, order, size: items?.length ?? groupItems.length, recordItems: generated.value, grouped: true } }
}

function pivotItemAt(map: ItemMap, record: number): number | undefined {
  if (!map.recordItems) return undefined
  const raw = map.recordItems[record]
  return raw === NO_GROUP ? -1 : map.toPivot[raw]
}

function axis(store: RecordStore, def: PivotTableDefinition, view: PivotViewState, which: 'rows' | 'cols', maps: Map<number, ItemMap>, passes: Uint8Array, sortedOrders: Map<number, readonly number[]>, allowedByField: Map<number, Set<number>>): EngineResult<{ lines: PivotLine[]; key: Uint32Array }> {
  const fields = which === 'rows' ? view.rows : view.cols
  const valuesLevel = fields.indexOf(PIVOT_VALUES_FIELD)
  const dataCount = view.data.length
  // Real (non-Values) levels, and for each whether a collapsed item there truncates the tuple.
  const real = fields.map((field, level) => ({ field, level })).filter(entry => entry.field !== PIVOT_VALUES_FIELD)
  const sizes = real.map(({ field }) => maps.get(field)!.size)
  const recordItems = real.map(({ field }) => maps.get(field)!.recordItems!)
  const toPivot = real.map(({ field }) => maps.get(field)!.toPivot)
  const groupedLevels = real.map(({ field }) => maps.get(field)!.grouped === true)
  const collapses = real.map(({ field, level }) => hasFieldAfter(fields, level) && view.fields[field]!.collapsed.includes(true) ? view.fields[field]!.collapsed : undefined)
  // A collapsed item above the Values field: measured on rows, outline form, Values last (layout-collapsed-item-above-values).
  for (let k = 0; k < real.length; k++) {
    const { field, level } = real[k]!
    if (!(level < valuesLevel && collapses[k])) continue
    if (which !== 'rows') return unmeasured('a collapsed item above the Values field on columns is not measured')
    if (valuesLevel !== fields.length - 1) return unmeasured('a collapsed item above the Values field is measured with Values as the last row level only')
    if (!view.fields[field]!.outline) return unmeasured('a collapsed item above the Values field is measured in outline form only')
  }

  // One pass over the records: each passing record's tuple of pivot item indices (cut after a collapsed item) is
  // hashed to a number in mixed radix (item + 1 per level, 0 past the cut), or to a string past 2^53. A small
  // key space is a direct table instead of a Map.
  let span = 1
  const radix = sizes.map(size => { const r = span; span *= size + 1; return r })
  const numeric = span <= Number.MAX_SAFE_INTEGER
  const table = span <= Math.max(1 << 16, 2 * store.count) && span <= 1 << 22 ? new Int32Array(span).fill(-1) : undefined
  const root = newNode()
  const nodes: Node[] = []
  const nodeOf = new Map<number | string, number>()
  // Reuse the output key buffer to hold node ids during grouping, then rewrite
  // ids to leaf line numbers in place after the line traversal.
  const recordKey = new Uint32Array(store.count).fill(NO_LINE)
  const tuple: number[] = [], levels = real.length
  for (let r = 0; r < store.count; r++) {
    if (!passes[r]) continue
    tuple.length = 0
    let hash = 0
    for (let k = 0; k < levels; k++) {
      const raw = recordItems[k]![r]!
      const item = groupedLevels[k] && raw === NO_GROUP ? -1 : toPivot[k]![raw]
      if (item === undefined || item < 0) return invalid(`a record of field ${real[k]!.field} has no pivot item`)
      tuple.push(item)
      hash += (item + 1) * radix[k]!
      if (collapses[k] && collapses[k]![item]) break
    }
    const key = numeric ? hash : tuple.join(',')
    let id = table ? table[hash]! : nodeOf.get(key)
    if (id === undefined || id < 0) {
      let at: Node = root
      for (const item of tuple) {
        let child: Node | undefined = at.children.get(item)
        if (!child) at.children.set(item, child = newNode())
        at = child
      }
      id = nodes.push(at) - 1
      if (table) table[hash] = id
      else nodeOf.set(key, id)
    }
    recordKey[r] = id
  }

  const lines: PivotLine[] = []
  let previous: number[] = []
  const push = (type: PivotItemType, items: number[], data: number) => {
    let shared = 0
    // An unfilled level (-1, under a collapsed item above Values) counts as shared: Excel writes r past it.
    while (shared < previous.length && shared < items.length && previous[shared] === items[shared]) shared++
    if (shared === previous.length) while (shared < items.length && items[shared] === -1) shared++
    lines.push({ type, first: type === 'grand' ? 0 : Math.min(shared, Math.max(0, items.length - 1)), items: type === 'grand' ? [] : items, data: Math.max(0, data) })
    if (type !== 'grand') previous = items
    return lines.length - 1
  }
  const ownLines = (field: number) => which === 'rows' && (field === PIVOT_VALUES_FIELD ? def.outlineData : view.fields[field]!.outline)
  const ranks = new Map<number, Map<number, number>>()
  for (const [field, order] of sortedOrders) ranks.set(field, new Map(order.map((item, index) => [item, index])))

  const visit = (level: number, path: number[], node: Node | undefined, data: number): void => {
    if (level === fields.length) {
      const line = push('data', path, data)
      if (node && node.leaf === NO_LINE) node.leaf = line
      return
    }
    const field = fields[level]!
    if (field === PIVOT_VALUES_FIELD) {
      for (let d = 0; d < dataCount; d++) {
        const items = [...path, d]
        if (level < fields.length - 1 && ownLines(field)) push('data', items, d)
        visit(level + 1, items, node, d)
      }
      return
    }
    const state = view.fields[field]!
    const deeper = hasFieldAfter(fields, level), last = level === fields.length - 1
    const rank = sortedOrders.get(field)
    const rankOf = ranks.get(field)
    // With no node (an item shown by showAll with no records above), the field's visible items show
    // (quirk-cross-hidden-item-showall), provided each has passing records somewhere (else unmeasured).
    let children: number[]
    if (node && !state.showAll) children = [...node.children.keys()].sort((a, b) => rankOf ? (rankOf.get(a) ?? a) - (rankOf.get(b) ?? b) : a - b)
    else {
      children = (rank ?? maps.get(field)!.order).filter(item => !state.hidden[item] && (allowedByField.get(field)?.has(item) ?? true))
      if (!node && !state.showAll) {
        const present = itemsWithRecords(field)
        const missing = children.find(item => !present.has(item))
        if (missing !== undefined) { refusal ??= `an item without records under showAll lists field ${field} item ${missing}, which no passing record has`; return }
      }
    }
    for (const item of children) {
      const child = node?.children.get(item)
      const items = [...path, item]
      if (last) { visit(level + 1, items, child, data); continue }
      if (deeper && state.collapsed[item] && valuesLevel > level) {
        // Collapsed above Values (gated above: rows, outline, Values last): its own line, then one line per data
        // field with the levels between unfilled; no subtotal lines (layout-collapsed-item-above-values).
        push('data', items, data)
        const padded = [...items]
        while (padded.length < valuesLevel) padded.push(-1)
        let leaf = NO_LINE
        for (let d = 0; d < dataCount; d++) { const line = push('data', [...padded, d], d); if (d === 0) leaf = line }
        if (child && child.leaf === NO_LINE) child.leaf = leaf
      } else if (deeper && state.collapsed[item]) {
        const line = push('data', items, data)
        if (child && child.leaf === NO_LINE) child.leaf = line
      } else {
        const own = ownLines(field), start = lines.length
        if (own) push('data', items, data)
        visit(level + 1, items, child, data)
        // A tabular item with nothing under it (showAll) still has its line.
        if (!own && lines.length === start) push('data', items, data)
        if (deeper && !subtotalOnTop(fields, level, state)) for (const type of subtotalTypes(state)) {
          if (data < 0 && valuesLevel > level) for (let d = 0; d < dataCount; d++) push(type, items, d)
          else push(type, items, data)
        }
      }
      if (deeper && which === 'rows' && state.insertBlankRow) push('blank', items, data)
    }
  }
  // Pivot items with at least one passing record, per field (built on first use: only under a record-less showAll item).
  let refusal: string | undefined
  const withRecords = new Map<number, Set<number>>()
  const itemsWithRecords = (field: number) => {
    let present = withRecords.get(field)
    if (!present) {
      present = new Set()
      const map = maps.get(field)!
      for (let r = 0; r < store.count; r++) if (passes[r]) { const item = pivotItemAt(map, r); if (item !== undefined && item >= 0) present.add(item) }
      withRecords.set(field, present)
    }
    return present
  }
  visit(0, [], root, -1)
  if (refusal !== undefined) return unmeasured(refusal)
  if (real.length && (which === 'rows' ? def.rowGrandTotals : def.colGrandTotals)) {
    if (valuesLevel >= 0) for (let d = 0; d < dataCount; d++) push('grand', [], d)
    else push('grand', [], 0)
  }

  for (let r = 0; r < store.count; r++) if (recordKey[r] !== NO_LINE) recordKey[r] = nodes[recordKey[r]!]!.leaf
  return { ok: true, value: { lines, key: recordKey } }
}
