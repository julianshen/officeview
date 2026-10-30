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
 * - the Values field is an axis level whose items are the data fields; grand totals are one line per data field
 *   when it is on the axis (pivot-values-on-rows), and need a non-Values field on the axis (lo-calc-two-data-fields);
 *   `rowGrandTotals` gates the row axis's grand line and `colGrandTotals` the column axis's (lo-calc-Pivot2_Row_Compact,
 *   lo-calc-Pivot3_Column_Grand_Subtotals);
 * - each line's `first` (`r`) is the levels it shares with the previous non-grand line, at most all but its last.
 * Unmeasured (Task 15 cases decide): named subtotals (one line per function in schema order, always at the bottom),
 * blank rows (one `blank` line after each item block of a field with a deeper field), an item with no records under
 * `showAll` (its line alone in tabular form), whether another field's hidden item drops an item (taken as above).
 * Refused as `pivot-engine-unmeasured` until measured: `mergeItem`, a hidden item on a field on no axis or page, and
 * a collapsed item above the Values field.
 */
import { PIVOT_VALUES_FIELD, type PivotCacheDefinition, type PivotDiagnostic, type PivotItemType, type PivotSubtotalFunction, type PivotTableDefinition } from '../model'
import type { PivotLine } from '../savedLayout'
import type { RecordStore } from './records'
import type { FieldState, PivotViewState } from './view'

export interface PivotAxes { rows: PivotLine[]; cols: PivotLine[]; rowKey: Uint32Array; colKey: Uint32Array }   // per record: its leaf line, or NO_LINE
export const NO_LINE = 0xffffffff
export type EngineResult<T> = { ok: true; value: T } | { ok: false; diagnostic: PivotDiagnostic }

const fail = (kind: PivotDiagnostic['kind'], feature: string, message: string): { ok: false; diagnostic: PivotDiagnostic } => ({ ok: false, diagnostic: { kind, feature, message } })
/** Sorting, label/value/date/top-10 filters and grouping are Task 21; calculated items are decided by Task 22. */
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
interface ItemMap { toPivot: Int32Array; order: number[]; size: number }
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
  if (view.filters.length) return pending('pivot filters (label, value, date, top 10) are not computed yet')
  // Merged labels are unmeasured: fail closed until Task 15 has an Excel case.
  if (def.mergeItem) return unmeasured('merged item labels (mergeItem) are not measured yet')
  // Calculated fields are Task 22's, on an axis, a page or as data.
  for (const field of [...view.rows, ...view.cols, ...view.pages.map(page => page.field), ...view.data.map(data => data.field)]) {
    const cacheField = cache.fields[field]
    if (cacheField && (cacheField.formula !== undefined || (!cacheField.databaseField && !cacheField.fieldGroup))) return pending(`field ${field} is a calculated field`)
  }
  const axisFields = [...view.rows, ...view.cols].filter(field => field !== PIVOT_VALUES_FIELD)
  if ([...view.rows, ...view.cols].filter(field => field === PIVOT_VALUES_FIELD).length > 1) return invalid('the Values field is on more than one axis level')
  const pageFields = view.pages.map(page => page.field)
  const used = new Set([...axisFields, ...pageFields])
  if (used.size !== axisFields.length + pageFields.length) return invalid('a field is used on more than one axis')
  // Whether a hidden item on a field on no axis or page filters records is unmeasured (a slicer on a field outside the
  // pivot may store its selection that way): fail closed until Task 15 has an Excel case.
  const offAxis = view.fields.findIndex((field, index) => field.hidden.includes(true) && !used.has(index))
  if (offAxis >= 0) return unmeasured(`field ${offAxis} is on no axis but has hidden items`)
  for (const field of used) {
    const state = view.fields[field], cacheField = cache.fields[field]
    if (!state || !cacheField || !def.fields[field]) return invalid(`field ${field} does not exist`)
    if (cacheField.fieldGroup) return pending(`field ${field} is grouped`)
    if (def.fields[field]!.items?.some(item => item.calculated) || cacheField.sharedItems?.items.some(item => item.calculated)) return pending(`field ${field} has calculated items`)
    // Sorting changes the layout only on rows and columns (a page field's order is its dropdown's).
    if (!axisFields.includes(field)) continue
    if (state.sortType !== 'manual' || state.autoSortScope) return pending(`field ${field} is sorted`)
  }

  // Record filter: the hidden items of every row, column and page field, and each page's selection.
  const maps = new Map<number, ItemMap>()
  const filters: Array<{ index: Uint32Array; allow: Uint8Array }> = []
  for (const field of used) {
    const column = store.columns[field]
    const map = itemMap(def, cache, field)
    if (typeof map === 'string') return invalid(map)
    if (!column || column.kind !== 'shared') return invalid(`field ${field} has no shared record column`)
    maps.set(field, map)
    const state = view.fields[field]!
    const selected = view.pages.find(page => page.field === field)?.selected ?? 'all'
    if (selected === 'all' && !state.hidden.includes(true)) continue
    const allow = new Uint8Array(map.toPivot.length)
    map.toPivot.forEach((item, x) => { allow[x] = +(item >= 0 && !state.hidden[item] && (selected === 'all' || selected.includes(item))) })
    filters.push({ index: column.index, allow })
  }
  const passes = new Uint8Array(store.count).fill(1)
  for (const { index, allow } of filters) for (let r = 0; r < store.count; r++) if (!allow[index[r]!]) passes[r] = 0

  const rows = axis(store, def, view, 'rows', maps, passes)
  if (!rows.ok) return rows
  const cols = axis(store, def, view, 'cols', maps, passes)
  if (!cols.ok) return cols
  return { ok: true, value: { rows: rows.value.lines, cols: cols.value.lines, rowKey: rows.value.key, colKey: cols.value.key } }
}

function axis(store: RecordStore, def: PivotTableDefinition, view: PivotViewState, which: 'rows' | 'cols', maps: Map<number, ItemMap>, passes: Uint8Array): EngineResult<{ lines: PivotLine[]; key: Uint32Array }> {
  const fields = which === 'rows' ? view.rows : view.cols
  const valuesLevel = fields.indexOf(PIVOT_VALUES_FIELD)
  const dataCount = view.data.length
  // Real (non-Values) levels, and for each whether a collapsed item there truncates the tuple.
  const real = fields.map((field, level) => ({ field, level })).filter(entry => entry.field !== PIVOT_VALUES_FIELD)
  const columns = real.map(({ field }) => (store.columns[field] as { index: Uint32Array }).index)
  const toPivot = real.map(({ field }) => maps.get(field)!.toPivot)
  const sizes = real.map(({ field }) => maps.get(field)!.size)
  const collapses = real.map(({ field, level }) => hasFieldAfter(fields, level) && view.fields[field]!.collapsed.includes(true) ? view.fields[field]!.collapsed : undefined)
  if (real.some(({ level }, k) => level < valuesLevel && collapses[k]))
    return unmeasured('a collapsed item above the Values field is not measured yet')

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
  const recordNode = new Uint32Array(store.count).fill(NO_LINE)
  const tuple: number[] = [], levels = real.length
  for (let r = 0; r < store.count; r++) {
    if (!passes[r]) continue
    tuple.length = 0
    let hash = 0
    for (let k = 0; k < levels; k++) {
      const item = toPivot[k]![columns[k]![r]!]
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
    recordNode[r] = id
  }

  const lines: PivotLine[] = []
  let previous: number[] = []
  const push = (type: PivotItemType, items: number[], data: number) => {
    let shared = 0
    while (shared < previous.length && shared < items.length && previous[shared] === items[shared]) shared++
    lines.push({ type, first: type === 'grand' ? 0 : Math.min(shared, Math.max(0, items.length - 1)), items: type === 'grand' ? [] : items, data: Math.max(0, data) })
    if (type !== 'grand') previous = items
    return lines.length - 1
  }
  const ownLines = (field: number) => which === 'rows' && (field === PIVOT_VALUES_FIELD ? def.outlineData : view.fields[field]!.outline)

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
    const children = state.showAll
      ? maps.get(field)!.order.filter(item => !state.hidden[item])
      : [...(node?.children.keys() ?? [])].sort((a, b) => a - b)
    for (const item of children) {
      const child = node?.children.get(item)
      const items = [...path, item]
      if (last) { visit(level + 1, items, child, data); continue }
      if (deeper && state.collapsed[item]) {
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
  visit(0, [], root, -1)
  if (real.length && (which === 'rows' ? def.rowGrandTotals : def.colGrandTotals)) {
    if (valuesLevel >= 0) for (let d = 0; d < dataCount; d++) push('grand', [], d)
    else push('grand', [], 0)
  }

  const key = new Uint32Array(store.count).fill(NO_LINE)
  for (let r = 0; r < store.count; r++) if (recordNode[r] !== NO_LINE) key[r] = nodes[recordNode[r]!]!.leaf
  return { ok: true, value: { lines, key } }
}
