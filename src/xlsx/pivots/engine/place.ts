/**
 * Layout to cells (spec §5.7): the axes of a view placed on the sheet from the pivot's anchor, with labels,
 * captions, page fields and value cells, and the definition as the live view would save it. Lazy chunk `pivots`
 * (Stage 2: library only).
 *
 * Placement read from Excel's saved cells (Stage 1 references tests/fixtures/xlsx-pivots/pivot-*.xlsx, Excel
 * 16.114 zh-TW; Excel-made corpus files):
 * - the anchor is `location`'s top-left; one header row holds the column header (the "Column Labels" caption in
 *   compact form, else each column field's name over its column), then one row per column field holds its item
 *   labels; with no column field there is a single header row, and with only the Values field on columns and
 *   `hideValuesRow` its item row is the single header row, in compact (lo-calc-shared-dategroup), outline and tabular
 *   form alike (Task 15 layout-outline-three-rows, layout-tabular-three-rows: A3:E60, firstHeaderRow 0), but not in
 *   classic layout (`gridDropZones`, lo-calc-two-data-fields keeps the row);
 * - classic layout with no column field (one data field, row fields; Task 15 layout-classic-single-data) has two
 *   header rows: the data field's name top-left, then the row header with Excel's "Total" over the value column
 *   (saved "合計" at B4), firstHeaderRow = firstDataRow = 2;
 * - the last header row holds the row header over each label column: the "Row Labels" caption over a compact
 *   column, else the field's name (pivot-layout-outline);
 * - with one data field and no Values field, its name goes in the top-left cell when both axes have fields
 *   (poi-ExcelPivotTableSample), over the value column with no column field (lo-calc-Pivot2_Row_Compact), and in the
 *   label column with no row field (lo-calc-Pivot3_Column_Grand_Subtotals);
 * - a line writes its items from its `first` level (repeated labels write the levels above too,
 *   pivot-repeat-labels); a subtotal's caption is "<item> Total", or "<item> <data field>" with the Values field
 *   deeper (pivot-values-on-rows); a grand total's caption is `grandTotalCaption`, or "Total <data field>" on the
 *   axis carrying the Values field (pivot-layout-compact);
 * - page fields sit above the location, one blank row between (savedPivotLayout's placement).
 * Captions follow the Stage 2 caption ruling: an explicit name in the definition, else the saved text for the slot
 * (`savedCaption`), else Excel's English default marked `caption: 'implicit'` with `captionsUnverified` set.
 * - merged labels (`mergeItem`, Task 15 layout-merge-labels, tabular fields): an outer row label is merged down over
 *   its item's lines (A6:A8), a row subtotal or grand caption across the label columns from its own (A9:B9, A24:B24);
 *   an outer column label across its item's columns (C4:D4), a column grand caption down to the last header row
 *   (I4:I5); a one-cell span is not merged. computeAxes refuses the shapes this case does not cover.
 * Unmeasured: named-subtotal captions, `showHeaders` off, an empty pivot.
 */
import type { SheetRect } from '../../tableStyles/model'
import {
  dataFieldSource, PIVOT_VALUES_FIELD,
  type PivotAxisItem, type PivotCacheDefinition, type PivotDataField, type PivotField, type PivotFieldItem,
  type PivotPageField, type PivotShowAsX14, type PivotTableDefinition,
} from '../model'
import type { PivotLine } from '../savedLayout'
import { isoToSerial } from './records'
import { lineCarriesValues, subtotalTypes, type PivotAxes } from './axes'
import type { PivotViewState } from './view'

export type PivotCellValue = { kind: 'empty' } | { kind: 'number'; value: number } | { kind: 'text'; value: string } | { kind: 'error'; value: string }
export interface PivotValueGrid { get(rowLine: number, colLine: number, data: number): PivotCellValue }
export interface PivotLayoutCell { row: number; col: number; value: PivotCellValue; numFmtId?: number; caption?: 'implicit' }
export interface PivotLayout { rect: SheetRect; axes: PivotAxes; cells: PivotLayoutCell[]; merges: SheetRect[]; captionsUnverified: boolean }
/** Ruling S2-J: the workbook's date system is explicit (the cache does not know it). `savedCaption` is the saved text
 * of a caption slot, for the saved state only (Stage 2 caption ruling). */
export interface PlaceOptions { date1904: boolean; savedCaption?: (row: number, col: number) => string | undefined }

const FUNCTION_NAMES: Record<string, string> = {
  sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count',
  stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp',
  // Named subtotal item types.
  countA: 'Count', avg: 'Average', stdDevP: 'StdDevp', varP: 'Varp',
}
const X14_SHOW_AS = new Set<string>(['percentOfParent', 'percentOfParentRow', 'percentOfParentCol', 'percentOfRunningTotal', 'rankAscending', 'rankDescending'])

const CELL = /^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})/
function anchorOf(ref: string): { row: number; col: number } {
  const m = CELL.exec(ref)
  if (!m) return { row: 0, col: 0 }
  return { row: Number(m[2]) - 1, col: [...m[1]!].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1 }
}
const columnName = (col: number) => { let s = ''; for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s }
const refOf = (rect: SheetRect) => `${columnName(rect.minCol)}${rect.minRow + 1}:${columnName(rect.maxCol)}${rect.maxRow + 1}`

interface Geometry {
  rect: SheetRect
  firstHeaderRow: number; firstDataRow: number; firstDataCol: number; rowPageCount: number; colPageCount: number
  /** Sheet rows/cols: first column-item row, first data row and column; each row level's label column. */
  itemRow0: number; dataRow0: number; dataCol0: number; labelCol: number[]
  pages: Array<{ row: number; col: number; index: number }>
}

/** Whether the field after `field` on the row axis shares its label column (compact form). */
const shares = (def: PivotTableDefinition, view: PivotViewState, field: number) =>
  field === PIVOT_VALUES_FIELD ? def.compactData && def.outlineData : Boolean(view.fields[field]?.compact && view.fields[field]?.outline)

function geometry(def: PivotTableDefinition, view: PivotViewState, axes: PivotAxes): Geometry {
  const anchor = anchorOf(def.location.ref)
  const labelCol: number[] = []
  view.rows.forEach((_, level) => labelCol.push(level === 0 ? anchor.col : labelCol[level - 1]! + (shares(def, view, view.rows[level - 1]!) ? 0 : 1)))
  const firstDataCol = view.rows.length ? labelCol[labelCol.length - 1]! - anchor.col + 1 : 1
  const columnFields = view.cols.length
  // hideValuesRow holds in every form but classic (see the header).
  const valuesRowHidden = columnFields === 1 && view.cols[0] === PIVOT_VALUES_FIELD && Boolean(def.x14?.hideValuesRow) && !def.gridDropZones
  const classic = classicNoColumns(def, view)
  const firstHeaderRow = valuesRowHidden ? 0 : classic ? 2 : 1
  const firstDataRow = classic ? 2 : columnFields === 0 || valuesRowHidden ? 1 : 1 + columnFields
  // An axis with nothing on it keeps one line; it has a body cell only when there is a value to show.
  const bodyRows = view.rows.length === 0 && !view.data.length ? 0 : axes.rows.length
  const bodyCols = view.cols.length === 0 && !view.data.length ? 0 : axes.cols.length
  const rect = { minRow: anchor.row, minCol: anchor.col, maxRow: anchor.row + firstDataRow + bodyRows - 1, maxCol: anchor.col + firstDataCol + bodyCols - 1 }
  // Page fields: "down, then over" unless pageOverThenDown, wrapping at pageWrap; a blank row above the location.
  const n = view.pages.length, wrap = def.pageWrap > 0 ? def.pageWrap : n
  let rowPageCount = 0, colPageCount = 0
  if (n) {
    if (def.pageOverThenDown) { colPageCount = Math.min(n, wrap); rowPageCount = Math.ceil(n / colPageCount) }
    else { rowPageCount = Math.min(n, wrap); colPageCount = Math.ceil(n / rowPageCount) }
  }
  const pages: Geometry['pages'] = []
  view.pages.forEach((_, k) => {
    const i = def.pageOverThenDown ? Math.floor(k / colPageCount) : k % rowPageCount
    const j = def.pageOverThenDown ? k % colPageCount : Math.floor(k / rowPageCount)
    const row = anchor.row - 1 - rowPageCount + i
    if (row >= 0) pages.push({ row, col: anchor.col + 3 * j, index: k })
  })
  return {
    rect, firstHeaderRow, firstDataRow, firstDataCol, rowPageCount, colPageCount,
    itemRow0: anchor.row + firstHeaderRow, dataRow0: anchor.row + firstDataRow, dataCol0: anchor.col + firstDataCol, labelCol, pages,
  }
}

/** Classic layout (`gridDropZones`) with no column field: two header rows (see the header). */
const classicNoColumns = (def: PivotTableDefinition, view: PivotViewState) => def.gridDropZones && view.cols.length === 0

interface Label { value: PivotCellValue; numFmtId?: number; implicit?: string }
const text = (value: string): Label => ({ value: { kind: 'text', value } })
const implicit = (english: string): Label => ({ value: { kind: 'text', value: english }, implicit: english })
/** A label's text for composing a caption ("<item> Total"). */
const labelText = (label: Label) => label.value.kind === 'empty' ? '' : String(label.value.value)

export function placeCells(def: PivotTableDefinition, view: PivotViewState, cache: PivotCacheDefinition, axes: PivotAxes, values: PivotValueGrid,
  options: PlaceOptions): PivotLayout {
  const { date1904, savedCaption } = options
  const geo = geometry(def, view, axes)
  const { rect, dataRow0, dataCol0, itemRow0, labelCol } = geo
  const cells = new Map<string, PivotLayoutCell>()
  const put = (row: number, col: number, label: Label) => {
    let cell: PivotLayoutCell = { row, col, value: label.value, ...(label.numFmtId !== undefined ? { numFmtId: label.numFmtId } : {}) }
    if (label.implicit !== undefined) {
      const saved = savedCaption?.(row, col)
      cell = saved !== undefined ? { row, col, value: { kind: 'text', value: saved } } : { ...cell, caption: 'implicit' }
    }
    cells.set(`${row},${col}`, cell)
  }

  const fieldName = (field: number) => field === PIVOT_VALUES_FIELD ? def.dataCaption : def.fields[field]?.name ?? cache.fields[field]?.name ?? ''
  const dataName = (d: number): Label => {
    const data = view.data[d]
    if (!data) return text('')
    return data.name ? text(data.name) : implicit(`${FUNCTION_NAMES[data.fn] ?? data.fn} of ${fieldName(data.field)}`)
  }
  const itemLabel = (field: number, item: number): Label => {
    if (field === PIVOT_VALUES_FIELD) return dataName(item)
    const pivotItem = def.fields[field]?.items?.[item]
    if (pivotItem?.name !== undefined) return text(pivotItem.name)
    const cacheField = cache.fields[field]
    const itemValues = cacheField?.fieldGroup?.groupItems ?? cacheField?.sharedItems?.items
    const value = itemValues?.[pivotItem ? pivotItem.x ?? -1 : item]
    if (!value) return text('')
    if (value.caption !== undefined) return text(value.caption)
    const numFmtId = def.fields[field]?.numFmtId ?? cacheField?.numFmtId
    switch (value.type) {
      case 'missing': return implicit('(blank)')
      case 'number': return { value: { kind: 'number', value: value.value }, ...(numFmtId !== undefined ? { numFmtId } : {}) }
      // Date items show as serials in the field's number format (lo-calc-pivot_dark1), in the workbook's date system.
      case 'date': return { value: { kind: 'number', value: isoToSerial(value.value, date1904) }, ...(numFmtId !== undefined ? { numFmtId } : {}) }
      case 'boolean': return text(value.value ? 'TRUE' : 'FALSE')
      // Excel saves an error-valued item's label as text (Task 15 sort-error-item R24C1: the text #N/A).
      case 'error': return text(value.value)
      case 'string': return text(value.value)
    }
  }
  const explicitOr = (explicit: string | undefined, english: string): Label => explicit !== undefined ? text(explicit) : implicit(english)
  /** A subtotal or grand line's caption. */
  const totalCaption = (fields: readonly number[], line: PivotLine): Label => {
    const valuesLevel = fields.indexOf(PIVOT_VALUES_FIELD)
    if (line.type === 'grand') {
      if (valuesLevel < 0) return explicitOr(def.grandTotalCaption, 'Grand Total')
      const name = dataName(line.data)
      return implicit(`Total ${labelText(name)}`)
    }
    const level = line.items.length - 1, field = fields[level]!
    const item = itemLabel(field, line.items[level]!)
    const custom = def.fields[field]?.subtotalCaption
    if (custom !== undefined) return text(custom)
    if (line.type === 'default') {
      if (valuesLevel <= level) return implicit(`${labelText(item)} Total`)
      // With the Values field deeper, the caption names the data field (pivot-values-on-rows: "East Sum of Sales").
      const name = dataName(line.data)
      const caption = `${labelText(item)} ${labelText(name)}`
      return item.implicit === undefined && name.implicit === undefined && item.value.kind === 'text' ? text(caption) : implicit(caption)
    }
    return implicit(`${labelText(item)} ${FUNCTION_NAMES[line.type] ?? line.type}`)
  }

  const valuesOnRows = view.rows.includes(PIVOT_VALUES_FIELD), valuesOnCols = view.cols.includes(PIVOT_VALUES_FIELD)
  // Header: the single data field's name.
  if (view.data.length === 1 && !valuesOnRows && !valuesOnCols) {
    const name = dataName(0)
    if ((view.rows.length && view.cols.length) || classicNoColumns(def, view)) put(rect.minRow, rect.minCol, name)
    else if (view.rows.length) put(dataRow0 - 1, dataCol0, name)
    else if (view.cols.length) put(dataRow0, rect.minCol, name)
    else put(dataRow0 - 1, dataCol0, name)
  }
  if (def.showHeaders) {
    // The column header row: "Column Labels" in compact form, else each column field's name.
    if (view.cols.length && geo.firstHeaderRow > 0) {
      if (shares(def, view, view.cols[0]!)) put(rect.minRow, dataCol0, explicitOr(def.colHeaderCaption, 'Column Labels'))
      else view.cols.forEach((field, k) => { if (dataCol0 + k <= rect.maxCol) put(rect.minRow, dataCol0 + k, text(fieldName(field))) })
    }
    // The row header over each label column.
    labelCol.forEach((col, level) => {
      if (labelCol.indexOf(col) !== level) return
      const field = view.rows[level]!
      put(dataRow0 - 1, col, shares(def, view, field) ? explicitOr(def.rowHeaderCaption, 'Row Labels') : text(fieldName(field)))
    })
  }
  // Classic layout with no column field: Excel's "Total" over the value column (layout-classic-single-data B4).
  if (classicNoColumns(def, view) && view.data.length) put(dataRow0 - 1, dataCol0, implicit('Total'))
  // Column item labels.
  axes.cols.forEach((line, m) => {
    const col = dataCol0 + m
    if (col > rect.maxCol) return
    if (line.type === 'data') {
      for (let level = line.first; level < line.items.length; level++) if (itemRow0 + level < dataRow0) put(itemRow0 + level, col, itemLabel(view.cols[level]!, line.items[level]!))
    } else if (line.type !== 'blank') {
      const level = line.type === 'grand' ? 0 : line.items.length - 1
      if (itemRow0 + level < dataRow0) put(itemRow0 + level, col, totalCaption(view.cols, line))
    }
  })
  // Row labels.
  axes.rows.forEach((line, n) => {
    const row = dataRow0 + n
    if (row > rect.maxRow) return
    if (line.type === 'data') {
      for (let level = 0; level < line.items.length; level++) {
        const field = view.rows[level]!
        const repeat = field !== PIVOT_VALUES_FIELD && view.fields[field]?.fillDownLabels && !shares(def, view, field)
        if (level >= line.first || repeat) put(row, labelCol[level]!, itemLabel(field, line.items[level]!))
      }
    } else if (line.type === 'grand') put(row, rect.minCol, totalCaption(view.rows, line))
    else if (line.type !== 'blank') put(row, labelCol[line.items.length - 1] ?? rect.minCol, totalCaption(view.rows, line))
  })
  // Page fields: the field's name, then its selection.
  for (const { row, col, index } of geo.pages) {
    const page = view.pages[index]!
    put(row, col, text(fieldName(page.field)))
    const state = view.fields[page.field]
    const order = def.fields[page.field]?.items
      ? def.fields[page.field]!.items!.flatMap((item, k) => item.type === 'data' ? [k] : [])
      : (cache.fields[page.field]?.sharedItems?.items ?? []).map((_, k) => k)
    const shown = page.selected === 'all' ? order.filter(k => !state?.hidden[k]) : page.selected
    put(row, col + 1, shown.length === 1 ? itemLabel(page.field, shown[0]!) : page.selected === 'all' && shown.length === order.length ? implicit('(All)') : implicit('(Multiple Items)'))
  }
  // Value cells.
  if (view.data.length) {
    const rowCarries = axes.rows.map(line => lineCarriesValues(view, 'rows', line))
    const colCarries = axes.cols.map(line => lineCarriesValues(view, 'cols', line))
    for (let n = 0; n < axes.rows.length && dataRow0 + n <= rect.maxRow; n++) {
      if (!rowCarries[n]) continue
      for (let m = 0; m < axes.cols.length && dataCol0 + m <= rect.maxCol; m++) {
        if (!colCarries[m]) continue
        const data = valuesOnRows ? axes.rows[n]!.data : valuesOnCols ? axes.cols[m]!.data : 0
        const value = values.get(n, m, data)
        if (value.kind === 'empty') continue
        const numFmtId = view.data[data]?.numFmtId
        cells.set(`${dataRow0 + n},${dataCol0 + m}`, { row: dataRow0 + n, col: dataCol0 + m, value, ...(numFmtId !== undefined ? { numFmtId } : {}) })
      }
    }
  }

  const list = [...cells.values()].sort((a, b) => a.row - b.row || a.col - b.col)
  return { rect, axes, cells: list, merges: def.mergeItem ? mergedLabels(axes, geo) : [], captionsUnverified: list.some(cell => cell.caption === 'implicit') }
}

/** The merged label ranges of a `mergeItem` pivot (see the header; layout-merge-labels). Only the shapes computeAxes
 * admits reach here: tabular fields, no Values field on rows, totals only after outer items. */
function mergedLabels(axes: PivotAxes, geo: Geometry): SheetRect[] {
  const merges: SheetRect[] = []
  const { rect, dataRow0, dataCol0, itemRow0, labelCol } = geo
  const add = (minRow: number, minCol: number, maxRow: number, maxCol: number) => { if (maxRow > minRow || maxCol > minCol) merges.push({ minRow, minCol, maxRow, maxCol }) }
  /** The last line of the block a level-`level` label on line `n` heads: the following data lines under the same item. */
  const blockEnd = (lines: PivotLine[], n: number, level: number) => {
    const head = lines[n]!
    let end = n
    while (end + 1 < lines.length) {
      const next = lines[end + 1]!
      if (next.type !== 'data' || next.items.length <= level || next.items.slice(0, level + 1).some((item, k) => item !== head.items[k])) break
      end++
    }
    return end
  }
  const lastLabel = labelCol[labelCol.length - 1] ?? rect.minCol
  axes.rows.forEach((line, n) => {
    if (line.type === 'data') {
      for (let level = line.first; level < line.items.length - 1; level++) add(dataRow0 + n, labelCol[level]!, dataRow0 + blockEnd(axes.rows, n, level), labelCol[level]!)
    } else if (line.type !== 'blank') add(dataRow0 + n, line.type === 'grand' ? rect.minCol : labelCol[line.items.length - 1]!, dataRow0 + n, lastLabel)
  })
  axes.cols.forEach((line, m) => {
    if (line.type === 'data') {
      for (let level = line.first; level < line.items.length - 1; level++) add(itemRow0 + level, dataCol0 + m, itemRow0 + level, dataCol0 + blockEnd(axes.cols, m, level))
    } else if (line.type === 'grand') add(itemRow0, dataCol0 + m, dataRow0 - 1, dataCol0 + m)
  })
  return merges
}

const fieldItem = (type: PivotFieldItem['type'], x?: number): PivotFieldItem => ({
  type, ...(x !== undefined ? { x } : {}), hidden: false, char: false, showDetails: true, calculated: false, missing: false, childItems: false, expanded: false, drillAcross: false,
})
const sameList = <T>(a: readonly T[], b: readonly T[]) => a.length === b.length && a.every((v, k) => v === b[k])
const encode = (lines: PivotLine[]): PivotAxisItem[] => lines.map(line => ({
  type: line.type, r: line.first, i: line.data, x: line.type === 'grand' ? [0] : line.items.slice(line.first),
}))

/** The definition as the live view would save it: fields' axis, items (h, sd), subtotal flags, rowFields/colFields/pageFields/dataFields,
 * rowItems/colItems from `axes`, and `location` resized from the layout. Stage 1's `savedPivotLayout`, `savedPivotRoles`, `pivotButtons`
 * and `formats` matching then run unchanged on it: this is the bridge for style roles and button placements (Task 26 consumes it). */
export function liveDefinition(def: PivotTableDefinition, view: PivotViewState, axes: PivotAxes, layout: PivotLayout): PivotTableDefinition {
  const geo = geometry(def, view, axes)
  // Perf audit item 9: per-field lookups through Sets and Maps built once, not array scans per field or item.
  const rowSet = new Set(view.rows), colSet = new Set(view.cols), dataSet = new Set(view.data.map(data => data.field))
  const pageOf = new Map<number, PivotViewState['pages'][number]>()
  for (const page of view.pages) if (!pageOf.has(page.field)) pageOf.set(page.field, page)
  const axisOf = (field: number): PivotField['axis'] =>
    rowSet.has(field) ? 'axisRow' : colSet.has(field) ? 'axisCol' : pageOf.has(field) ? 'axisPage' : undefined
  const fillDefault = def.x14?.fillDownLabelsDefault ?? false
  const fields = def.fields.map((field, index): PivotField => {
    const state = view.fields[index]
    if (!state) return field
    const page = pageOf.get(index)
    const multi = page && page.selected !== 'all' && page.selected.length > 1 ? new Set(page.selected) : undefined
    const subtotalsChanged = state.defaultSubtotal !== field.defaultSubtotal || !sameList(state.subtotals, field.subtotals)
    // A field saved without an items list (other writers) has the cache's shared-item order as its pivot items, one
    // state slot per shared item (initialViewState); an edit to it materializes that list so the edit is saved.
    const edited = Boolean(multi) || state.hidden.includes(true) || state.collapsed.includes(true)
    const source = field.items ?? (edited ? [
      ...state.hidden.map((_, x) => fieldItem('data', x)),
      ...subtotalTypes(state).map(type => fieldItem(type)),
    ] : undefined)
    let items = source?.map((item, k) => ({ ...item, hidden: multi ? item.type === 'data' && !multi.has(k) : state.hidden[k] ?? item.hidden, showDetails: !(state.collapsed[k] ?? !item.showDetails) }))
    if (items && field.items && subtotalsChanged) items = [...items.filter(item => item.type === 'data'), ...subtotalTypes(state).map(type => fieldItem(type))]
    const { axis: _axis, x14: _x14, autoSortScope: _scope, ...rest } = field
    const axis = axisOf(index) ?? (field.axis === 'axisValues' ? field.axis : undefined)
    const x14 = field.x14 || state.fillDownLabels !== fillDefault ? { fillDownLabels: state.fillDownLabels, ignore: field.x14?.ignore ?? false } : undefined
    return {
      ...rest,
      ...(axis ? { axis } : {}),
      dataField: dataSet.has(index),
      ...(items ? { items } : {}),
      sortType: state.sortType,
      ...(state.autoSortScope ? { autoSortScope: state.autoSortScope } : {}),
      defaultSubtotal: state.defaultSubtotal,
      subtotals: [...state.subtotals],
      subtotalTop: state.subtotalTop,
      compact: state.compact,
      outline: state.outline,
      insertBlankRow: state.insertBlankRow,
      showAll: state.showAll,
      ...(x14 ? { x14 } : {}),
    }
  })
  const dataFields = view.data.map((data, k): PivotDataField => {
    const old = def.dataFields[k]
    if (old && dataFieldSource(old) === data.field && old.subtotal === data.fn && (old.x14?.pivotShowAs ?? old.showDataAs) === data.showAs
      && old.baseField === data.baseField && old.baseItem === data.baseItem && (old.name ?? '') === data.name && old.numFmtId === data.numFmtId) return old
    const x14 = X14_SHOW_AS.has(data.showAs)
    return {
      ...(data.name ? { name: data.name } : {}),
      fld: data.field, subtotal: data.fn,
      showDataAs: x14 ? 'normal' : data.showAs as PivotDataField['showDataAs'],
      baseField: data.baseField, baseItem: data.baseItem,
      ...(data.numFmtId !== undefined ? { numFmtId: data.numFmtId } : {}),
      ...(x14 ? { x14: { pivotShowAs: data.showAs as PivotShowAsX14 } } : {}),
    }
  })
  const savedPages = new Map<number, PivotPageField>()
  for (const candidate of def.pageFields) if (!savedPages.has(candidate.fld)) savedPages.set(candidate.fld, candidate)
  const pageFields = view.pages.map((page): PivotPageField => {
    const old = savedPages.get(page.field)
    const item = page.selected !== 'all' && page.selected.length === 1 ? page.selected[0] : undefined
    const { item: _item, ...rest } = old ?? { fld: page.field }
    return { ...rest, fld: page.field, ...(item !== undefined ? { item } : {}) }
  })
  const valuesAxis = view.rows.includes(PIVOT_VALUES_FIELD) ? view.rows : view.cols.includes(PIVOT_VALUES_FIELD) ? view.cols : undefined
  const savedAxis = def.rowFields.includes(PIVOT_VALUES_FIELD) ? def.rowFields : def.colFields.includes(PIVOT_VALUES_FIELD) ? def.colFields : undefined
  const position = valuesAxis?.indexOf(PIVOT_VALUES_FIELD) ?? -1
  const samePosition = valuesAxis !== undefined && (valuesAxis === view.rows) === (savedAxis === def.rowFields) && savedAxis?.indexOf(PIVOT_VALUES_FIELD) === position
  const dataPosition = !valuesAxis || samePosition ? def.dataPosition : position === valuesAxis.length - 1 ? undefined : position
  const { dataPosition: _position, ...base } = def
  return {
    ...base,
    ...(dataPosition !== undefined ? { dataPosition } : {}),
    dataOnRows: valuesAxis ? valuesAxis === view.rows : def.dataOnRows,
    location: {
      ref: refOf(layout.rect),
      firstHeaderRow: geo.firstHeaderRow, firstDataRow: geo.firstDataRow, firstDataCol: geo.firstDataCol,
      rowPageCount: geo.rowPageCount, colPageCount: geo.colPageCount,
    },
    fields,
    rowFields: [...view.rows],
    colFields: [...view.cols],
    rowItems: encode(axes.rows),
    colItems: encode(axes.cols),
    pageFields,
    dataFields,
    filters: [...view.filters],
  }
}
