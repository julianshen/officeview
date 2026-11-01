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
 * - `collapse` and `collapseField` refuse a field on no row or column axis. They and `collapseAll` mark data items only. `collapseAll` is Excel's "collapse entire field" on every
 *   field with a level below it on its axis (Task 24 interaction-collapse-all-*: Region on rows, Year above Values on
 *   columns); the innermost field of each axis is left alone.
 * - `filter`: a filter of the same family on the field (label or date; value or top-10, the engine's two kinds,
 *   axes.ts) is replaced in place. A filter of the other family on a field that already has one replaces it when
 *   the pivot allows one filter per field (`multipleFieldFilters="0"`, Ruling S3-Y: Excel's filter panel resets the
 *   other family, owner-filter-two-families); with several filters allowed it still refuses with the engine's "more
 *   than one filter" feature, `pivot-engine-unmeasured` (the combined evaluation is not measured, §9).
 * - `clearFilters` drops the field's pivot filters and shows its hidden items, and a page field goes back to (All):
 *   Excel's "Clear Filter From" (Task 24 interaction-clear-*, -clear-filters-hidden-items, -filter-date-clear).
 * - Item indices stay those of the start file's `items` list across a delta sequence, a sort included: the view never
 *   reorders a field's items (Excel re-saves a sorted field's list in the new order; interaction-sort-then-collapse).
 * - `page` sets the page field's selection; selecting (All) also shows every item of the field (Task 24
 *   interaction-page-multi-to-all), a single or multiple selection leaves its hidden items as they are.
 * - `sort` sets `sortType`; `byData` (a data-field index) sets the `autoSortScope` Excel saves for "sort by a value
 *   column" (Task 15 sort-value-top-level), and a sort without `byData` drops it.
 */
import {
  PIVOT_VALUES_FIELD,
  type PivotArea, type PivotCacheDefinition, type PivotDataFunction, type PivotDiagnostic, type PivotFilter,
  type PivotShowAsX14, type PivotShowDataAs, type PivotTableDefinition,
} from '../model'
import type { EngineResult } from '../engine/axes'
import { isDateFilter, isLabelFilter } from '../engine/filters'
import type { FieldState, PivotViewState } from '../engine/view'
import type { RecordColumn, RecordStore } from '../engine/records'
import { VALUE_TAG } from '../engine/records'
import type { PivotFieldItem, PivotCacheItem } from '../model'
import { pageFieldsFit } from '../engine/place'
import { sortItemOrder } from '../engine/sort'
import { ZH_TW_CAPTIONS } from './captions'
import { decodeSerial } from '../../formula/serial'
import { applyGrouping } from './grouping'

export type { PivotDataFunction } from '../model'
export interface PivotMemoryHold { readonly bytes: number; release(): void }
export interface PivotDeltaContext {
  def: PivotTableDefinition; cache: PivotCacheDefinition; store: RecordStore; view: PivotViewState; date1904: boolean
  captionRule?: { locale?: 'zh-TW' | 'file'; caption(kind: 'dataField', args: { fn: string; field?: string }): string | undefined }
  admitCandidateBytes?: (extraBytes: number) => PivotMemoryHold | undefined
}
export interface PivotDeltaStateResult { def: PivotTableDefinition; cache: PivotCacheDefinition; store: RecordStore; view: PivotViewState; extraBytes: number; hold?: PivotMemoryHold }

const asciiLower = (s: string) => s.replace(/[A-Z]+/g, run => run.toLowerCase())
function inlineValue(column: Extract<RecordColumn, { kind: 'inline' }>, row: number, date1904: boolean): PivotCacheItem {
  const tag = column.tag[row]!
  if (tag === VALUE_TAG.missing) return { type: 'missing' }
  if (tag === VALUE_TAG.number) return { type: 'number', value: column.num[row]! }
  if (tag === VALUE_TAG.boolean) return { type: 'boolean', value: column.num[row] !== 0 }
  const text = column.text[column.textIndex[row]!] ?? ''
  if (tag === VALUE_TAG.error) return { type: 'error', value: text }
  if (tag === VALUE_TAG.date) {
    const decoded = decodeSerial(column.num[row]!, date1904 ? '1904' : '1900')
    if ('civil' in decoded) {
      const { year, month, day } = decoded.civil
      const date = `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`
      if (decoded.fraction === 0) return { type: 'date', value: date }
      const seconds = decoded.fraction * 86_400
      const wholeSeconds = Math.floor(seconds)
      const decimal = (seconds - wholeSeconds).toFixed(12).slice(1).replace(/0+$/, '')
      const fraction = decimal.length > 1 ? decimal : ''
      const hours = Math.floor(wholeSeconds / 3_600), minutes = Math.floor(wholeSeconds / 60) % 60, second = wholeSeconds % 60
      return { type: 'date', value: `${date}T${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(second).padStart(2, '0')}${fraction}` }
    }
    return { type: 'error', value: '#VALUE!' }
  }
  return { type: 'string', value: text }
}
const valueKey = (v: PivotCacheItem) => v.type === 'missing' ? 'missing' : `${v.type}:${v.type === 'string' ? asciiLower(v.value) : String(v.value)}`
function fieldStateForMaterialized(old: FieldState, count: number): FieldState {
  return { ...old, hidden: Array.from({ length: count }, (_, i) => old.hidden[i] ?? false), collapsed: Array.from({ length: count }, (_, i) => old.collapsed[i] ?? false) }
}
function dataFunction(field: number, cache: PivotCacheDefinition): PivotDataFunction {
  const cf = cache.fields[field]
  return cf?.formula !== undefined || (cf?.sharedItems?.containsNumber && !cf.sharedItems.containsString && !cf.sharedItems.containsDate) ? 'sum' : 'count'
}
function remapDataReferences(view: PivotViewState, map: readonly (number | undefined)[]): PivotViewState {
  return {
    ...view,
    filters: view.filters.flatMap(filter => {
      const index = filter.iMeasureFld
      if (index === undefined) return [filter]
      const next = map[index]
      return next === undefined ? [] : [{ ...filter, iMeasureFld: next }]
    }),
    fields: view.fields.map(state => {
      const scope = state.autoSortScope
      if (!scope) return state
      const references = scope.references.flatMap(ref => {
        if (ref.field !== PIVOT_VALUES_FIELD) return [ref]
        const next = ref.x.map(index => map[index])
        return next.some(index => index === undefined) ? [] : [{ ...ref, x: next as number[] }]
      })
      return { ...state, ...(references.length ? { autoSortScope: { ...scope, references } } : { autoSortScope: undefined }) }
    }),
  }
}

/** Task 35 candidate transition. Materialization is reserved before allocating the candidate index or item arrays. */
export function applyDeltaState(ctx: PivotDeltaContext, delta: PivotDelta): EngineResult<PivotDeltaStateResult> {
  if (delta.kind === 'groupDate' || delta.kind === 'groupNumber' || delta.kind === 'ungroup') return applyGrouping(ctx, delta)
  if (!['moveField', 'addDataField', 'removeDataField', 'moveDataField', 'valueSettings'].includes(delta.kind)) {
    const view = applyDelta(ctx.view, delta, ctx.def, ctx.cache)
    return view.ok ? { ok: true, value: { def: ctx.def, cache: ctx.cache, store: ctx.store, view: view.value, extraBytes: 0 } } : view
  }
  const { def, cache, store, view } = ctx
  let nextDef = def, nextCache = cache, nextStore = store, nextView = view, extraBytes = 0
  let hold: PivotMemoryHold | undefined
  const failWith = (result: EngineResult<never>): EngineResult<PivotDeltaStateResult> => { hold?.release(); return result }
  const reject = (feature: string, message: string) => failWith(refuse(feature, message))
  if (delta.kind !== 'valueSettings' && def.disableFieldList) return reject('pivot-field-list-disabled', `The field list is disabled for ${def.name}.`)
  if (delta.kind === 'moveField') {
    const field = delta.field
    if (field === PIVOT_VALUES_FIELD) {
      if (delta.axis !== 'row' && delta.axis !== 'col') return reject('pivot-delta-invalid', 'Values can be placed on Rows or Columns only.')
      if (view.data.length < 2) return reject('pivot-delta-invalid', 'Values is available on an axis only when there are at least two data fields.')
      const source = view.rows.includes(field) ? 'row' : view.cols.includes(field) ? 'col' : 'none'
      const rows = view.rows.filter(f => f !== field), cols = view.cols.filter(f => f !== field)
      const target = delta.axis === 'row' ? rows : cols
      const at = delta.index ?? target.length
      if (!Number.isInteger(at) || at < 0 || at > target.length) return reject('pivot-delta-invalid', 'The Values position is invalid.')
      target.splice(at, 0, field)
      nextView = { ...view, rows, cols }
      void source
    } else {
      if (!Number.isInteger(field) || field < 0 || field >= def.fields.length || !def.fields[field]) return reject('pivot-delta-invalid', `Field ${field} does not exist.`)
      const pf = def.fields[field]!
      const cacheField = cache.fields[field]
      const current: 'row' | 'col' | 'page' | 'none' = view.rows.includes(field) ? 'row' : view.cols.includes(field) ? 'col' : view.pages.some(p => p.field === field) ? 'page' : 'none'
      if (cacheField?.formula !== undefined && delta.axis !== 'none') return reject('pivot-calculated-field', `Calculated field ${field} can be used in Values only.`)
      const permission = delta.axis === 'row' ? pf.dragToRow : delta.axis === 'col' ? pf.dragToCol : delta.axis === 'page' ? pf.dragToPage : pf.dragOff
      if (!permission) return reject('pivot-delta-invalid', `Field ${field} cannot be moved to ${delta.axis}.`)
      if (current !== delta.axis && current !== 'none' && !pf.dragOff) return reject('pivot-delta-invalid', `Field ${field} cannot be removed from its current area.`)
      const rows = [...view.rows], cols = [...view.cols], pages = [...view.pages]
      const removeAt = <T,>(array: T[], item: T) => { const i = array.indexOf(item); if (i >= 0) array.splice(i, 1) }
      removeAt(rows, field); removeAt(cols, field)
      const pageAt = pages.findIndex(p => p.field === field); if (pageAt >= 0) pages.splice(pageAt, 1)
      const target = delta.axis === 'row' ? rows : delta.axis === 'col' ? cols : undefined
      if (target) {
        const at = delta.index ?? target.length
        if (!Number.isInteger(at) || at < 0 || at > target.length) return reject('pivot-delta-invalid', `The ${delta.axis} position is invalid.`)
        target.splice(at, 0, field)
      } else if (delta.axis === 'page') {
        const at = delta.index ?? pages.length
        if (!Number.isInteger(at) || at < 0 || at > pages.length) return reject('pivot-delta-invalid', 'The Filters position is invalid.')
        const selected = current === 'page' && pageAt >= 0 ? view.pages[pageAt]!.selected : 'all'
        pages.splice(at, 0, { field, selected })
      }
      nextView = { ...view, rows, cols, pages, ...(delta.axis === 'page' ? { filters: view.filters.filter(filter => filter.fld !== field) } : {}) }
      if (delta.axis === 'none') {
        nextView = { ...nextView, filters: view.filters.filter(f => f.fld !== field), pages: pages.filter(p => p.field !== field), fields: view.fields.map((f, i) => i === field ? { ...f, hidden: f.hidden.map(() => false), collapsed: f.collapsed.map(() => false), sortType: 'manual', autoSortScope: undefined } : f) }
      }
      if (delta.axis === 'page' && !pageFieldsFit(def, nextView)) return reject('pivot-page-fields-fit', `Field ${field} cannot fit in the page-field area.`)

      const cf = cache.fields[field], column = store.columns[field]
      if ((delta.axis === 'row' || delta.axis === 'col' || delta.axis === 'page') && cf?.sharedItems?.inline) {
        if (!column || column.kind !== 'inline') return reject('pivot-records-absent', `Field ${field} has no inline record column to materialize.`)
        // A size upper bound derived from source cardinality. The reservation is taken before item/index arrays exist.
        const maximum = store.count * (4 + 96 + 8) + 256
        const admitted = ctx.admitCandidateBytes?.(maximum)
        if (ctx.admitCandidateBytes && !admitted) return reject('pivot-records-limit', `Materializing field ${field} exceeds the memory budget.`)
        hold = admitted; extraBytes = maximum
        const values: PivotCacheItem[] = [], indexByKey = new Map<string, number>(), rowIndex = new Uint32Array(store.count)
        const seenLower = new Map<string, string>()
        for (let row = 0; row < store.count; row++) {
          const value = inlineValue(column, row, ctx.date1904), key = valueKey(value)
          if (value.type === 'string') {
            const lower = value.value.toLowerCase(), folded = asciiLower(value.value), previous = seenLower.get(lower)
            if (previous !== undefined && previous !== folded) return reject('pivot-engine-unmeasured', `Field ${field} contains non-ASCII case-fold collisions.`)
            seenLower.set(lower, folded)
          }
          let index = indexByKey.get(key)
          if (index === undefined) { index = values.length; indexByKey.set(key, index); values.push(value) }
          rowIndex[row] = index
        }
        const old = cf.sharedItems
        nextCache = { ...cache, fields: cache.fields.map((f, i) => i === field ? { ...f, sharedItems: {
          ...(old ?? { containsSemiMixedTypes: false, containsNonDate: false, containsDate: false, containsString: false, containsBlank: false, containsMixedTypes: false, containsNumber: false, containsInteger: false, longText: false, inline: true, items: [] }),
          inline: false, items: values,
        } } : f) }
        nextStore = { ...store, columns: store.columns.map((c, i) => i === field ? { kind: 'shared' as const, index: rowIndex } : c), bytes: store.bytes + rowIndex.byteLength }
        // The oracle measures date items together with blanks: dates sort chronologically and (blank) follows.
        const hasDates = values.some(value => value.type === 'date')
        const mixedDateBlank = hasDates && values.every(value => value.type === 'date' || value.type === 'missing')
        const order = mixedDateBlank
          ? { ok: true as const, value: values.map((_, x) => x).sort((a, b) => {
              const left = values[a]!, right = values[b]!
              if (left.type === 'missing') return right.type === 'missing' ? a - b : 1
              if (right.type === 'missing') return -1
              if (left.type !== 'date' || right.type !== 'date') return a - b
              return left.value < right.value ? -1 : left.value > right.value ? 1 : a - b
            }) }
          : sortItemOrder(values.map((_, x) => x), values, 'ascending', field, def.customListSort !== false)
        if (!order.ok) return failWith(order)
        const items: PivotFieldItem[] = order.value.map(x => ({ type: 'data', x, hidden: false, char: false, showDetails: true, calculated: false, missing: false, childItems: false, expanded: false, drillAcross: false }))
        nextDef = { ...def, fields: def.fields.map((f, i) => i === field ? { ...f, axis: delta.axis === 'row' ? 'axisRow' : delta.axis === 'col' ? 'axisCol' : 'axisPage', items: f.items?.length ? f.items : items } : f) }
        nextView = { ...nextView, fields: nextView.fields.map((f, i) => i === field ? fieldStateForMaterialized(f, values.length) : f) }
      } else {
        const axis = delta.axis === 'row' ? 'axisRow' : delta.axis === 'col' ? 'axisCol' : delta.axis === 'page' ? 'axisPage' : undefined
        nextDef = { ...def, fields: def.fields.map((f, i) => i === field ? { ...f, ...(axis ? { axis } : { axis: undefined }) } : f) }
      }
      if (delta.axis !== 'none') nextDef = { ...nextDef, fields: nextDef.fields.map((f, i) => i === field ? { ...f, axis: delta.axis === 'row' ? 'axisRow' : delta.axis === 'col' ? 'axisCol' : 'axisPage' } : f) }
    }
  } else if (delta.kind === 'addDataField') {
    const field = delta.field
    if (!Number.isInteger(field) || field < 0 || field >= def.fields.length) return reject('pivot-delta-invalid', `Field ${field} does not exist.`)
    const pf = def.fields[field]!
    if (!pf.dragToData) return reject('pivot-delta-invalid', `Field ${field} cannot be moved to Values.`)
    if (!cache.fields[field]?.databaseField && cache.fields[field]?.formula === undefined) return reject('pivot-engine-unmeasured', `Field ${field} has no value column.`)
    const at = delta.index ?? view.data.length
    if (!Number.isInteger(at) || at < 0 || at > view.data.length) return reject('pivot-delta-invalid', 'The Values position is invalid.')
    const fn = dataFunction(field, cache)
    const fieldName = def.fields[field]?.name ?? cache.fields[field]?.name ?? `Field ${field + 1}`
    const table = ZH_TW_CAPTIONS[`dataField:${fn}`]
    const fromMeasuredTable = table?.replace('{field}', fieldName)
    const baseName = ctx.captionRule?.caption('dataField', { fn, field: fieldName }) ?? fromMeasuredTable
    if (!baseName) return reject('pivot-caption-locale', `A default caption for ${fn} is not measured; provide a custom name.`)
    let name = baseName, suffix = 2
    const names = new Set(view.data.map(d => d.name).filter(Boolean))
    while (names.has(name)) name = `${baseName}${suffix++}`
    const data = [...view.data]; data.splice(at, 0, { field, fn, showAs: 'normal', baseField: -1, baseItem: 1048832, name })
    const map = Array.from({ length: view.data.length }, (_, i) => i < at ? i : i + 1)
    nextView = remapDataReferences({ ...view, data }, map)
    if (view.data.length === 1) {
      const valuesAxis = def.dataOnRows ? 'rows' : 'cols'
      nextView = { ...nextView, [valuesAxis]: [...nextView[valuesAxis], PIVOT_VALUES_FIELD] }
    }
    nextDef = { ...def, fields: def.fields.map((f, i) => i === field ? { ...f, dataField: true } : f) }
  } else if (delta.kind === 'removeDataField') {
    if (!Number.isInteger(delta.dataIndex) || delta.dataIndex < 0 || delta.dataIndex >= view.data.length) return reject('pivot-delta-invalid', `Data field ${delta.dataIndex} does not exist.`)
    const data = [...view.data]; data.splice(delta.dataIndex, 1)
    const map = Array.from({ length: view.data.length }, (_, i) => i === delta.dataIndex ? undefined : i > delta.dataIndex ? i - 1 : i)
    nextView = remapDataReferences({ ...view, data }, map)
    if (data.length < 2) nextView = { ...nextView, rows: nextView.rows.filter(f => f !== PIVOT_VALUES_FIELD), cols: nextView.cols.filter(f => f !== PIVOT_VALUES_FIELD) }
    const remaining = new Set(data.map(d => d.field))
    nextDef = { ...def, fields: def.fields.map((f, i) => i >= 0 && i < def.fields.length ? { ...f, dataField: remaining.has(i) } : f) }
  } else if (delta.kind === 'moveDataField') {
    if (!Number.isInteger(delta.dataIndex) || !Number.isInteger(delta.toIndex) || delta.dataIndex < 0 || delta.dataIndex >= view.data.length || delta.toIndex < 0 || delta.toIndex >= view.data.length) return reject('pivot-delta-invalid', 'The Values reorder index is invalid.')
    const data = [...view.data], [item] = data.splice(delta.dataIndex, 1); data.splice(delta.toIndex, 0, item!)
    const map = Array.from({ length: view.data.length }, (_, old) => data.findIndex(entry => entry === view.data[old]))
    nextView = remapDataReferences({ ...view, data }, map)
  } else if (delta.kind === 'valueSettings') {
    const { dataIndex } = delta
    if (!Number.isInteger(dataIndex) || dataIndex < 0 || dataIndex >= view.data.length) {
      return reject('pivot-delta-invalid', `Data field ${dataIndex} does not exist.`)
    }
    const current = view.data[dataIndex]!
    const sourceField = current.field
    const isCalculated = cache.fields[sourceField]?.formula !== undefined

    const VALID_FNS = new Set<PivotDataFunction>(['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp'])
    if (delta.fn !== undefined && !VALID_FNS.has(delta.fn)) return reject('pivot-delta-invalid', `Summary function ${delta.fn} is not supported.`)
    const nextFn = delta.fn ?? current.fn
    if (isCalculated && nextFn !== 'sum') return reject('pivot-delta-invalid', 'Calculated fields only support Sum.')

    const VALID_SHOW_AS = new Set<PivotValueShowAs>([
      'normal', 'difference', 'percent', 'percentDiff', 'runTotal',
      'percentOfRow', 'percentOfCol', 'percentOfTotal', 'index',
      'percentOfParent', 'percentOfParentRow', 'percentOfParentCol',
      'percentOfRunningTotal', 'rankAscending', 'rankDescending',
    ])
    if (delta.showAs !== undefined && !VALID_SHOW_AS.has(delta.showAs)) return reject('pivot-delta-invalid', `Show values as mode ${delta.showAs} is not supported.`)
    const nextShowAs = delta.showAs ?? current.showAs
    if (isCalculated && nextShowAs !== 'normal' && nextShowAs !== 'percentOfTotal') return reject('pivot-delta-invalid', 'Calculated fields only support No Calculation or % of Grand Total.')

    const needsBaseField = ['percentOfParent', 'difference', 'percent', 'percentDiff', 'runTotal', 'percentOfRunningTotal', 'rankAscending', 'rankDescending'].includes(nextShowAs)
    const needsBaseItem = ['difference', 'percent', 'percentDiff'].includes(nextShowAs)

    let baseField = delta.baseField !== undefined ? delta.baseField : current.baseField
    let baseItem = delta.baseItem !== undefined ? delta.baseItem : current.baseItem

    if (needsBaseField) {
      if (baseField === undefined || !Number.isInteger(baseField) || baseField < 0 || baseField >= def.fields.length) {
        return reject('pivot-delta-invalid', `Base field is required for ${nextShowAs}.`)
      }
      const onRows = view.rows.includes(baseField)
      const onCols = view.cols.includes(baseField)
      if (nextShowAs === 'percentOfParent') {
        if (!onRows) return reject('pivot-delta-invalid', '% of Parent Total requires a row base field.')
      } else {
        if (!onRows && !onCols) return reject('pivot-delta-invalid', `Base field ${baseField} must be on rows or columns.`)
      }
    } else {
      if (delta.showAs !== undefined && delta.baseField === undefined) {
        baseField = -1
      }
    }

    if (needsBaseItem) {
      if (baseItem === undefined || !Number.isInteger(baseItem)) {
        return reject('pivot-delta-invalid', `Base item is required for ${nextShowAs}.`)
      }
      const isPrevNext = baseItem === 1048828 || baseItem === 1048829
      const bf = def.fields[baseField]
      const itemCount = bf?.items?.length ?? cache.fields[baseField]?.sharedItems?.items.length ?? 0
      const isItemIndex = baseItem >= 0 && baseItem < itemCount
      if (!isPrevNext && !isItemIndex) {
        return reject('pivot-delta-invalid', `Base item ${baseItem} is invalid for base field ${baseField}.`)
      }
    } else {
      if (delta.showAs !== undefined && delta.baseItem === undefined) {
        baseItem = 1048832
      }
    }

    const PRESET_NUMFMT_IDS = new Set([0, 1, 2, 3, 4, 9, 10, 11, 14, 37, 38, 39, 40, 44, 49])
    let nextNumFmtId: number | undefined
    if (delta.numFmtId === null) {
      nextNumFmtId = undefined
    } else if (delta.numFmtId !== undefined) {
      if (!Number.isInteger(delta.numFmtId) || delta.numFmtId < 0) {
        return reject('pivot-delta-invalid', `Number format ID ${delta.numFmtId} is invalid.`)
      }
      if (PRESET_NUMFMT_IDS.has(delta.numFmtId)) {
        nextNumFmtId = delta.numFmtId
      } else if (delta.numFmtId >= 164 && delta.numFmtId === current.numFmtId) {
        nextNumFmtId = delta.numFmtId
      } else {
        return reject('pivot-delta-invalid', `Number format ID ${delta.numFmtId} is not supported.`)
      }
    } else {
      nextNumFmtId = current.numFmtId
    }

    let nextName: string
    if (delta.name !== undefined) {
      const trimmed = delta.name.trim()
      if (!trimmed) return reject('pivot-delta-invalid', 'The custom name cannot be blank.')
      const otherNames = view.data.filter((_, i) => i !== dataIndex).map(d => d.name).filter(Boolean)
      if (otherNames.some(n => n.toLowerCase() === trimmed.toLowerCase())) {
        return reject('pivot-delta-invalid', `A data field named "${trimmed}" already exists.`)
      }
      nextName = trimmed
    } else if (delta.fn !== undefined) {
      if (ctx.captionRule?.locale === 'file' || !ctx.captionRule) {
        return reject('pivot-caption-locale', `A default caption for ${nextFn} is not measured; provide a custom name.`)
      }
      const fieldName = def.fields[sourceField]?.name ?? cache.fields[sourceField]?.name ?? `Field ${sourceField + 1}`
      const table = ZH_TW_CAPTIONS[`dataField:${nextFn}`]
      const fromMeasuredTable = table?.replace('{field}', fieldName)
      const baseName = ctx.captionRule.caption('dataField', { fn: nextFn, field: fieldName }) ?? fromMeasuredTable
      if (!baseName) return reject('pivot-caption-locale', `A default caption for ${nextFn} is not measured; provide a custom name.`)
      let disambiguated = baseName, suffix = 2
      const otherNames = new Set(view.data.filter((_, i) => i !== dataIndex).map(d => d.name).filter(Boolean))
      while (otherNames.has(disambiguated)) disambiguated = `${baseName}${suffix++}`
      nextName = disambiguated
    } else {
      nextName = current.name
    }

    const nextData = view.data.map((entry, i) => {
      if (i !== dataIndex) return entry
      const updated = {
        ...entry,
        fn: nextFn,
        showAs: nextShowAs,
        baseField,
        baseItem,
        name: nextName,
        ...(nextNumFmtId !== undefined ? { numFmtId: nextNumFmtId } : {}),
      }
      if (nextNumFmtId === undefined) delete (updated as any).numFmtId
      return updated
    })
    nextView = { ...view, data: nextData }

    const X14_SHOW_AS = new Set<string>(['percentOfParent', 'percentOfParentRow', 'percentOfParentCol', 'percentOfRunningTotal', 'rankAscending', 'rankDescending'])
    if (def.dataFields && def.dataFields[dataIndex]) {
      nextDef = {
        ...def,
        dataFields: def.dataFields.map((df, i) => {
          if (i !== dataIndex) return df
          const updatedDf = {
            ...df,
            subtotal: nextFn,
            showDataAs: (X14_SHOW_AS.has(nextShowAs) ? 'normal' : nextShowAs) as any,
            baseField,
            baseItem,
            name: nextName,
            ...(nextNumFmtId !== undefined ? { numFmtId: nextNumFmtId } : {}),
            ...(X14_SHOW_AS.has(nextShowAs) ? { x14: { ...df.x14, pivotShowAs: nextShowAs as any } } : {}),
          }
          if (nextNumFmtId === undefined) delete (updatedDf as any).numFmtId
          return updatedDf
        }),
      }
    }
  }
  return { ok: true, value: { def: nextDef, cache: nextCache, store: nextStore, view: nextView, extraBytes, ...(hold ? { hold } : {}) } }
}

export type PivotDateGroupPeriod = 'seconds' | 'minutes' | 'hours' | 'days' | 'months' | 'quarters' | 'years'
export type PivotValueShowAs = PivotShowDataAs | PivotShowAsX14

export type PivotDelta =
  // Stage 3 kinds (unchanged)
  | { kind: 'collapse'; field: number; item: number; collapsed: boolean }
  | { kind: 'collapseField'; field: number; collapsed: boolean }
  | { kind: 'collapseAll'; collapsed: boolean }
  | { kind: 'items'; field: number; hidden: readonly number[] }          // pivot-item indices hidden
  | { kind: 'page'; field: number; selected: readonly number[] | 'all' }
  | { kind: 'filter'; field: number; filter: PivotFilter }               // label | value | date | top-10
  | { kind: 'clearFilters'; field: number }
  | { kind: 'sort'; field: number; order: 'ascending' | 'descending'; byData?: number }
  // Stage 4 kinds (Task 35: field list; Task 36: value settings; Task 37: grouping)
  | { kind: 'moveField'; field: number; axis: 'row' | 'col' | 'page' | 'none'; index?: number }
  | { kind: 'addDataField'; field: number; index?: number }
  | { kind: 'removeDataField'; dataIndex: number }
  | { kind: 'moveDataField'; dataIndex: number; toIndex: number }
  | { kind: 'valueSettings'; dataIndex: number; fn?: PivotDataFunction; showAs?: PivotValueShowAs; baseField?: number; baseItem?: number; name?: string; numFmtId?: number | null }
  | { kind: 'groupDate'; field: number; periods: readonly PivotDateGroupPeriod[]; startDate?: number; endDate?: number; daysInterval?: number }
  | { kind: 'groupNumber'; field: number; start?: number; end?: number; by: number }
  | { kind: 'ungroup'; field: number }
export type PivotDeltaKind = PivotDelta['kind']
export const PIVOT_DELTA_KINDS: readonly PivotDeltaKind[] = Object.freeze([
  'collapse', 'collapseField', 'collapseAll', 'items', 'page', 'filter', 'clearFilters', 'sort',
  'moveField', 'addDataField', 'removeDataField', 'moveDataField', 'valueSettings', 'groupDate', 'groupNumber', 'ungroup',
])

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
  switch (delta.kind) {
    case 'moveField':
    case 'addDataField':
    case 'removeDataField':
    case 'moveDataField':
    case 'valueSettings':
    case 'groupDate':
    case 'groupNumber':
    case 'ungroup':
      return refuse('pivot-interaction-unverified', `the ${delta.kind} interaction is not implemented yet`)
  }
  if (delta.kind === 'collapseAll') {
    // Every field with a level below it on its axis (the Values field counts as a level).
    const axis = new Set([...view.rows.slice(0, -1), ...view.cols.slice(0, -1)].filter(field => field !== PIVOT_VALUES_FIELD))
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
      if (!onAxis) return invalid(`field ${field} is not on the row or column axis`)
      if (!isData.has(delta.item)) return invalid(`item ${delta.item} is not a data item of field ${field}`)
      return { ok: true, value: withField(view, field, { collapsed: setFlags(state.collapsed, [delta.item], delta.collapsed) }) }
    case 'collapseField':
      if (!onAxis) return invalid(`field ${field} is not on the row or column axis`)
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
      const next = { ...view, pages: view.pages.map((page, k) => k === at ? { ...page, selected } : page) }
      // (All) shows every item; a list shows each item it names, even one a saved multi-item filter ((All) plus hidden
      // items) had hidden (PR #57 review). Other items' hidden flags stay as they are.
      const shown = selected === 'all' ? isData : new Set(selected)
      return { ok: true, value: withField(next, field, { hidden: state.hidden.map((flag, k) => shown.has(k) && isData.has(k) ? false : flag) }) }
    }
    case 'filter': {
      if (delta.filter.fld !== field) return invalid(`the filter names field ${delta.filter.fld}, not ${field}`)
      const kind = family(delta.filter)
      const same = view.filters.findIndex(filter => filter.fld === field && family(filter) === kind)
      if (same >= 0) return { ok: true, value: { ...view, filters: view.filters.map((filter, k) => k === same ? delta.filter : filter) } }
      if (view.filters.some(filter => filter.fld === field)) {
        // S3-Y: one filter per field replaces the other family's, keeping its place in the evaluation order.
        if (!def.multipleFieldFilters) return { ok: true, value: { ...view, filters: view.filters.flatMap(filter => filter.fld !== field ? [filter] : filter === view.filters.find(other => other.fld === field) ? [delta.filter] : []) } }
        return refuse('pivot-engine-unmeasured', `field ${field} would have a label or date filter and a value or top-10 filter; combining them is not measured`)
      }
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
