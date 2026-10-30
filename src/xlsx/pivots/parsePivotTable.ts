/**
 * `pivotTableDefinition` → PivotTableDefinition (spec §3). Lazy chunk `pivots`.
 *
 * Failure policy: a part without a readable `location` or `cacheId` is not a
 * pivot we can place or bind, so it returns undefined with a `parse`
 * diagnostic and the table is skipped (its saved cells still paint as plain
 * cells). Anything else unknown or malformed is recorded in `unsupported`,
 * the schema default stands, and the table is kept: display needs only the
 * location, saved items and style, and `unsupported` makes it display-only.
 */
import { attrs, getChild, getChildren, type XmlNode } from '../../core/xml'
import type { PivotRegionRef } from './stub'
import {
  dataFieldSource, PIVOT_BASE_ITEM_NEXT, PIVOT_BASE_ITEM_PREVIOUS, PIVOT_VALUES_FIELD,
  type PivotArea, type PivotAreaReference, type PivotAxis, type PivotAxisItem, type PivotCacheDefinition, type PivotDataField, type PivotDiagnostic,
  type PivotField, type PivotFieldItem, type PivotFilter, type PivotFilterColumn, type PivotItemType, type PivotSubtotalFunction, type PivotTableDefinition,
} from './model'
import {
  bool, checkChildren, checkCount, countDiagnostics, double, int, oneOf, optBool, optDouble, optInt, optOneOf, optUint, pushDiagnostic,
  readAttrs, readExtensions, readIndices, readList, str, uint, Unsupported,
} from './parseCacheDefinition'

export const X14_PIVOT_TABLE_EXT = '{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}'
export const X14_PIVOT_FIELD_EXT = '{2946ED86-A175-432a-8AC1-64E0C546D7DE}'
export const X14_DATA_FIELD_EXT = '{E15A36E0-9728-4e99-A89B-3F7291B0FE68}'
export const XPDL_PIVOT_TABLE_EXT = '{747A6164-185A-40DC-8AA5-F01512510D54}'
/** The Values field as an unsigned pivot-area reference field. */
const VALUES_FIELD_UNSIGNED = 4294967294

const AXES = ['axisRow', 'axisCol', 'axisPage', 'axisValues'] as const satisfies readonly PivotAxis[]
const ITEM_TYPES = ['data', 'default', 'sum', 'countA', 'avg', 'max', 'min', 'product', 'count', 'stdDev', 'stdDevP', 'var', 'varP', 'grand', 'blank'] as const satisfies readonly PivotItemType[]
const SUBTOTALS = ['sum', 'countA', 'avg', 'max', 'min', 'product', 'count', 'stdDev', 'stdDevP', 'var', 'varP'] as const satisfies readonly PivotSubtotalFunction[]
const SUBTOTAL_ATTRS = SUBTOTALS.map(name => `${name}Subtotal`)
const FILTER_TYPES = [
  'unknown', 'count', 'percent', 'sum', 'captionEqual', 'captionNotEqual', 'captionBeginsWith', 'captionNotBeginsWith', 'captionEndsWith',
  'captionNotEndsWith', 'captionContains', 'captionNotContains', 'captionGreaterThan', 'captionGreaterThanOrEqual', 'captionLessThan',
  'captionLessThanOrEqual', 'captionBetween', 'captionNotBetween', 'valueEqual', 'valueNotEqual', 'valueGreaterThan', 'valueGreaterThanOrEqual',
  'valueLessThan', 'valueLessThanOrEqual', 'valueBetween', 'valueNotBetween', 'dateEqual', 'dateNotEqual', 'dateOlderThan', 'dateOlderThanOrEqual',
  'dateNewerThan', 'dateNewerThanOrEqual', 'dateBetween', 'dateNotBetween', 'tomorrow', 'today', 'yesterday', 'nextWeek', 'thisWeek', 'lastWeek',
  'nextMonth', 'thisMonth', 'lastMonth', 'nextQuarter', 'thisQuarter', 'lastQuarter', 'nextYear', 'thisYear', 'lastYear', 'yearToDate',
  'Q1', 'Q2', 'Q3', 'Q4', 'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M11', 'M12',
] as const satisfies readonly PivotFilter['type'][]
const AREA_REF = /^\$?[A-Z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Z]{1,3}\$?[1-9]\d{0,6})?$/

/** The named subtotal flags set on a field or reference, in schema order. */
function subtotalsOf(node: XmlNode, el: string, u: Unsupported): PivotSubtotalFunction[] {
  const a = attrs(node)
  return SUBTOTALS.filter((_, i) => bool(false)(a[SUBTOTAL_ATTRS[i]], () => u.add(`${el}@${SUBTOTAL_ATTRS[i]}=${a[SUBTOTAL_ATTRS[i]]}`)))
}

function readArea(node: XmlNode, u: Unsupported): PivotArea {
  const a = readAttrs(node, 'pivotArea', {
    field: optInt, type: oneOf(['none', 'normal', 'data', 'all', 'origin', 'button', 'topEnd', 'topRight'] as const, 'normal'),
    dataOnly: bool(true), labelOnly: bool(false), grandRow: bool(false), grandCol: bool(false), cacheIndex: bool(false),
    outline: bool(true), offset: str, collapsedLevelsAreSubtotals: bool(false), axis: optOneOf(AXES), fieldPosition: optUint,
  }, u)
  checkChildren(node, 'pivotArea', ['references'], u)
  readExtensions(node, {}, [], u)
  const refs = getChild(node, 'references')
  if (refs) {
    readList(refs, 'references', getChildren(refs, 'reference').length, u)
    checkChildren(refs, 'references', ['reference'], u)
  }
  const references = getChildren(refs, 'reference').map((ref): PivotAreaReference => {
    const r = readAttrs(ref, 'reference', {
      field: optUint, selected: bool(true), byPosition: bool(false), relative: bool(false), defaultSubtotal: bool(false),
    }, u, ['count', ...SUBTOTAL_ATTRS])
    checkChildren(ref, 'reference', ['x'], u)
    readExtensions(ref, {}, [], u)
    const field = r.field === VALUES_FIELD_UNSIGNED ? -2 : r.field
    const x = readIndices(ref, 'reference', u)
    checkCount(ref, 'reference', x.length, u)
    return { ...r, ...(field !== undefined ? { field } : {}), subtotals: subtotalsOf(ref, 'reference', u), x }
  })
  return { ...a, references }
}

/** A required `pivotArea` child; a missing one is recorded and read as the all-defaults area. */
function requiredArea(node: XmlNode, el: string, u: Unsupported): PivotArea {
  const area = getChild(node, 'pivotArea')
  if (!area) u.add(`${el}/pivotArea(missing)`)
  return readArea(area ?? {}, u)
}

function readFieldItem(node: XmlNode, u: Unsupported): PivotFieldItem {
  const a = readAttrs(node, 'item', {
    n: str, t: oneOf(ITEM_TYPES, 'data'), h: bool(false), s: bool(false), sd: bool(true), f: bool(false), m: bool(false),
    c: bool(false), x: optUint, d: bool(false), e: bool(true),
  }, u)
  checkChildren(node, 'item', [], u)
  return {
    type: a.t, ...(a.x !== undefined ? { x: a.x } : {}), ...(a.n !== undefined ? { name: a.n } : {}),
    hidden: a.h, char: a.s, showDetails: a.sd, calculated: a.f, missing: a.m, childItems: a.c, expanded: a.d, drillAcross: a.e,
  }
}

function readField(node: XmlNode, u: Unsupported): PivotField {
  const a = readAttrs(node, 'pivotField', {
    name: str, axis: optOneOf(AXES), dataField: bool(false), subtotalCaption: str, showDropDowns: bool(true), hiddenLevel: bool(false),
    uniqueMemberProperty: str, compact: bool(true), allDrilled: bool(false), numFmtId: optUint, outline: bool(true), subtotalTop: bool(true),
    dragToRow: bool(true), dragToCol: bool(true), multipleItemSelectionAllowed: bool(false), dragToPage: bool(true), dragToData: bool(true),
    dragOff: bool(true), showAll: bool(true), insertBlankRow: bool(false), serverField: bool(false), insertPageBreak: bool(false),
    autoShow: bool(false), topAutoShow: bool(true), hideNewItems: bool(false), measureFilter: bool(false), includeNewItemsInFilter: bool(false),
    itemPageCount: uint(10), sortType: oneOf(['manual', 'ascending', 'descending'] as const, 'manual'), dataSourceSort: optBool,
    nonAutoSortDefault: bool(false), rankBy: optUint, defaultSubtotal: bool(true), showPropCell: bool(false), showPropTip: bool(false),
    showPropAsCaption: bool(false), defaultAttributeDrillState: bool(false),
  }, u, SUBTOTAL_ATTRS)
  checkChildren(node, 'pivotField', ['items', 'autoSortScope'], u)
  const field: PivotField = { ...a, subtotals: subtotalsOf(node, 'pivotField', u) }
  const items = getChild(node, 'items')
  if (items) {
    checkChildren(items, 'items', ['item'], u)
    field.items = getChildren(items, 'item').map(item => readFieldItem(item, u))
    readList(items, 'items', field.items.length, u)
  }
  const scope = getChild(getChild(node, 'autoSortScope'), 'pivotArea')
  if (scope) field.autoSortScope = readArea(scope, u)
  readExtensions(node, {
    [X14_PIVOT_FIELD_EXT]: ext => {
      const x14 = getChild(ext, 'pivotField')
      if (x14) field.x14 = readAttrs(x14, 'x14:pivotField', { fillDownLabels: bool(false), ignore: bool(false) }, u)
    },
  }, [], u)
  return field
}

/** `rowFields`/`colFields`: signed field indices (-2 is the Values field). An unreadable entry is recorded and dropped. */
function readAxisFields(node: XmlNode | undefined, el: string, u: Unsupported): number[] {
  if (!node) return []
  readList(node, el, getChildren(node, 'field').length, u)
  checkChildren(node, el, ['field'], u)
  return getChildren(node, 'field').flatMap(f => {
    const x = readAttrs(f, 'field', { x: optInt }, u).x
    if (x === undefined) u.add('field@x(missing)')
    return x === undefined ? [] : [x]
  })
}

function readAxisItems(node: XmlNode | undefined, el: string, u: Unsupported): PivotAxisItem[] | undefined {
  if (!node) return undefined
  readList(node, el, getChildren(node, 'i').length, u)
  checkChildren(node, el, ['i'], u)
  return getChildren(node, 'i').map(i => {
    const a = readAttrs(i, 'i', { t: oneOf(ITEM_TYPES, 'data'), r: uint(0), i: uint(0) }, u)
    checkChildren(i, 'i', ['x'], u)
    return { type: a.t, r: a.r, i: a.i, x: readIndices(i, 'i', u, 'x') }
  })
}

function readDataField(node: XmlNode, u: Unsupported): PivotDataField {
  const a = readAttrs(node, 'dataField', {
    name: str, fld: optUint,
    subtotal: oneOf(['average', 'count', 'countNums', 'max', 'min', 'product', 'stdDev', 'stdDevp', 'sum', 'var', 'varp'] as const, 'sum'),
    showDataAs: oneOf(['normal', 'difference', 'percent', 'percentDiff', 'runTotal', 'percentOfRow', 'percentOfCol', 'percentOfTotal', 'index'] as const, 'normal'),
    baseField: int(-1), baseItem: uint(1048832), numFmtId: optUint,
  }, u)
  checkChildren(node, 'dataField', [], u)
  if (a.fld === undefined) u.add('dataField@fld(missing)')
  const field: PivotDataField = { ...a, fld: a.fld ?? 0 }
  readExtensions(node, {
    [X14_DATA_FIELD_EXT]: ext => {
      const x14 = getChild(ext, 'dataField')
      if (x14) field.x14 = readAttrs(x14, 'x14:dataField', {
        pivotShowAs: optOneOf(['percentOfParent', 'percentOfParentRow', 'percentOfParentCol', 'percentOfRunningTotal', 'rankAscending', 'rankDescending'] as const),
        sourceField: optUint,
      }, u)
    },
  }, [], u)
  return field
}

function readFilterColumn(node: XmlNode, u: Unsupported): PivotFilterColumn {
  const a = readAttrs(node, 'filterColumn', { colId: uint(0), hiddenButton: bool(false), showButton: bool(true) }, u)
  // colorFilter and iconFilter cannot occur on a pivot and are not modelled.
  checkChildren(node, 'filterColumn', ['top10', 'customFilters', 'dynamicFilter', 'filters'], u)
  readExtensions(node, {}, [], u)
  const column: PivotFilterColumn = { ...a }
  const top10 = getChild(node, 'top10')
  if (top10) column.top10 = readAttrs(top10, 'top10', { top: bool(true), percent: bool(false), val: double(0), filterVal: optDouble }, u)
  const custom = getChild(node, 'customFilters')
  if (custom) column.customFilters = {
    ...readAttrs(custom, 'customFilters', { and: bool(false) }, u),
    filters: getChildren(custom, 'customFilter').map(f => readAttrs(f, 'customFilter', {
      operator: oneOf(['equal', 'lessThan', 'lessThanOrEqual', 'notEqual', 'greaterThanOrEqual', 'greaterThan'] as const, 'equal'), val: str,
    }, u)),
  }
  const dynamic = getChild(node, 'dynamicFilter')
  if (dynamic) {
    const d = readAttrs(dynamic, 'dynamicFilter', { type: str, val: optDouble, valIso: str, maxVal: optDouble, maxValIso: str }, u)
    column.dynamicFilter = { ...d, type: d.type ?? 'null' }
  }
  const filters = getChild(node, 'filters')
  if (filters) {
    // dateGroupItem and calendarType are not modelled.
    checkChildren(filters, 'filters', ['filter'], u)
    column.filters = {
      ...readAttrs(filters, 'filters', { blank: bool(false) }, u),
      values: getChildren(filters, 'filter').map(f => readAttrs(f, 'filter', { val: str }, u).val ?? ''),
    }
  }
  return column
}

function readFilter(node: XmlNode, u: Unsupported): PivotFilter {
  const a = readAttrs(node, 'filter', {
    fld: uint(0), mpFld: optUint, type: oneOf(FILTER_TYPES, 'unknown'), evalOrder: int(0), id: uint(0), iMeasureHier: optUint,
    iMeasureFld: optUint, name: str, description: str, stringValue1: str, stringValue2: str,
  }, u)
  checkChildren(node, 'filter', ['autoFilter'], u)
  readExtensions(node, {}, [], u)
  const auto = getChild(node, 'autoFilter')
  if (auto) {
    checkChildren(auto, 'autoFilter', ['filterColumn'], u)
    readExtensions(auto, {}, [], u)
  }
  return {
    ...a,
    autoFilter: { ...(auto ? readAttrs(auto, 'autoFilter', { ref: str }, u) : {}), columns: getChildren(auto, 'filterColumn').map(c => readFilterColumn(c, u)) },
  }
}

/** Parses one definition, or returns undefined (with a diagnostic) when it has no usable `location` or `cacheId`. */
export function parsePivotTable(root: XmlNode, ctx: { sheet: string; region: PivotRegionRef }, diagnostics: PivotDiagnostic[]): PivotTableDefinition | undefined {
  const partPath = ctx.region.partPath
  const u = new Unsupported()
  const a = readAttrs(root, 'pivotTableDefinition', {
    name: str, cacheId: optUint, dataOnRows: bool(false), dataPosition: optUint, autoFormatId: optUint,
    applyNumberFormats: bool(false), applyBorderFormats: bool(false), applyFontFormats: bool(false), applyPatternFormats: bool(false),
    applyAlignmentFormats: bool(false), applyWidthHeightFormats: bool(false),
    dataCaption: str, grandTotalCaption: str, errorCaption: str, showError: bool(false), missingCaption: str, showMissing: bool(true),
    pageStyle: str, pivotTableStyle: str, vacatedStyle: str, tag: str, updatedVersion: uint(0), minRefreshableVersion: uint(0),
    asteriskTotals: bool(false), showItems: bool(true), editData: bool(false), disableFieldList: bool(false), showCalcMbrs: bool(true),
    visualTotals: bool(true), showMultipleLabel: bool(true), showDataDropDown: bool(true), showDrill: bool(true), printDrill: bool(false),
    showMemberPropertyTips: bool(true), showDataTips: bool(true), enableWizard: bool(true), enableDrill: bool(true),
    enableFieldProperties: bool(true), preserveFormatting: bool(true), useAutoFormatting: bool(false), pageWrap: uint(0),
    pageOverThenDown: bool(false), subtotalHiddenItems: bool(false), rowGrandTotals: bool(true), colGrandTotals: bool(true),
    fieldPrintTitles: bool(false), itemPrintTitles: bool(false), mergeItem: bool(false), showDropZones: bool(true), createdVersion: uint(0),
    indent: uint(1), showEmptyRow: bool(false), showEmptyCol: bool(false), showHeaders: bool(true), compact: bool(true), outline: bool(false),
    outlineData: bool(false), compactData: bool(true), published: bool(false), gridDropZones: bool(false), immersive: bool(true),
    multipleFieldFilters: bool(true), chartFormat: uint(0), rowHeaderCaption: str, colHeaderCaption: str, fieldListSortAscending: bool(false),
    mdxSubqueries: bool(false), customListSort: bool(true),
  }, u)
  const locationNode = getChild(root, 'location')
  const location = locationNode && readAttrs(locationNode, 'location', {
    ref: str, firstHeaderRow: optUint, firstDataRow: optUint, firstDataCol: optUint, rowPageCount: uint(0), colPageCount: uint(0),
  }, u)
  if (!location || !location.ref || !AREA_REF.test(location.ref) || location.firstHeaderRow === undefined || location.firstDataRow === undefined || location.firstDataCol === undefined || a.cacheId === undefined) {
    const what = a.cacheId === undefined ? 'cacheId' : 'location'
    pushDiagnostic(diagnostics, { kind: 'parse', feature: 'pivot-definition-unreadable', message: `Pivot table ${what} could not be read from ${partPath}`, partPath })
    return undefined
  }
  // pivotHierarchies, rowHierarchiesUsage and colHierarchiesUsage (OLAP) are not modelled.
  checkChildren(root, 'pivotTableDefinition', [
    'location', 'pivotFields', 'rowFields', 'rowItems', 'colFields', 'colItems', 'pageFields', 'dataFields', 'formats',
    'conditionalFormats', 'chartFormats', 'pivotTableStyleInfo', 'filters',
  ], u)
  const list = (name: string, child: string) => {
    const node = getChild(root, name)
    if (node) {
      readList(node, name, getChildren(node, child).length, u)
      checkChildren(node, name, [child], u)
    }
    return getChildren(node, child)
  }
  const { cacheId, name, dataCaption, ...rest } = a
  const table: PivotTableDefinition = {
    sheet: ctx.sheet, partPath, region: ctx.region, cacheId, name: name ?? '', dataCaption: dataCaption ?? '', ...rest,
    location: {
      ref: location.ref, firstHeaderRow: location.firstHeaderRow, firstDataRow: location.firstDataRow, firstDataCol: location.firstDataCol,
      rowPageCount: location.rowPageCount, colPageCount: location.colPageCount,
    },
    fields: list('pivotFields', 'pivotField').map(node => readField(node, u)),
    rowFields: readAxisFields(getChild(root, 'rowFields'), 'rowFields', u),
    colFields: readAxisFields(getChild(root, 'colFields'), 'colFields', u),
    pageFields: list('pageFields', 'pageField').map(node => {
      checkChildren(node, 'pageField', [], u)
      readExtensions(node, {}, [], u)
      const p = readAttrs(node, 'pageField', { fld: optInt, item: optUint, hier: optInt, name: str, cap: str }, u)
      if (p.fld === undefined) u.add('pageField@fld(missing)')
      return { ...p, fld: p.fld ?? 0 }
    }),
    dataFields: list('dataFields', 'dataField').map(node => readDataField(node, u)),
    formats: list('formats', 'format').map(node => {
      const f = readAttrs(node, 'format', { action: oneOf(['blank', 'formatting', 'drill', 'formula'] as const, 'formatting'), dxfId: optUint }, u)
      checkChildren(node, 'format', ['pivotArea'], u)
      readExtensions(node, {}, [], u)
      return { ...f, pivotArea: requiredArea(node, 'format', u) }
    }),
    conditionalFormats: list('conditionalFormats', 'conditionalFormat').map(node => {
      const c = readAttrs(node, 'conditionalFormat', {
        scope: oneOf(['selection', 'data', 'field'] as const, 'selection'), type: oneOf(['none', 'all', 'row', 'column'] as const, 'none'), priority: uint(0),
      }, u)
      checkChildren(node, 'conditionalFormat', ['pivotAreas'], u)
      readExtensions(node, {}, [], u)
      const areas = getChild(node, 'pivotAreas')
      if (areas) {
        readList(areas, 'pivotAreas', getChildren(areas, 'pivotArea').length, u)
        checkChildren(areas, 'pivotAreas', ['pivotArea'], u)
      }
      return { ...c, pivotAreas: getChildren(areas, 'pivotArea').map(area => readArea(area, u)) }
    }),
    chartFormats: list('chartFormats', 'chartFormat').map(node => {
      const c = readAttrs(node, 'chartFormat', { chart: uint(0), format: uint(0), series: bool(false) }, u)
      checkChildren(node, 'chartFormat', ['pivotArea'], u)
      return { ...c, pivotArea: requiredArea(node, 'chartFormat', u) }
    }),
    filters: list('filters', 'filter').map(node => readFilter(node, u)),
    unsupported: [],
  }
  const rowItems = readAxisItems(getChild(root, 'rowItems'), 'rowItems', u)
  const colItems = readAxisItems(getChild(root, 'colItems'), 'colItems', u)
  if (rowItems) table.rowItems = rowItems
  if (colItems) table.colItems = colItems
  const style = getChild(root, 'pivotTableStyleInfo')
  if (style) table.styleInfo = readAttrs(style, 'pivotTableStyleInfo', {
    name: str, showRowHeaders: bool(false), showColHeaders: bool(false), showRowStripes: bool(false), showColStripes: bool(false), showLastColumn: bool(false),
  }, u)
  readExtensions(root, {
    [X14_PIVOT_TABLE_EXT]: ext => {
      const node = getChild(ext, 'pivotTableDefinition')
      if (!node) return
      // x14 pivotEdits, pivotChanges and conditionalFormats are not modelled.
      checkChildren(node, 'x14:pivotTableDefinition', [], u)
      table.x14 = readAttrs(node, 'x14:pivotTableDefinition', {
        fillDownLabelsDefault: bool(false), visualTotalsForSets: bool(false), calculatedMembersInFilters: bool(false), altText: str,
        altTextSummary: str, enableEdit: bool(false), autoApply: bool(false),
        allocationMethod: oneOf(['equalAllocation', 'equalIncrement', 'weightedAllocation', 'weightedIncrement'] as const, 'equalAllocation'),
        weightExpression: str, hideValuesRow: bool(false),
      }, u)
    },
    [XPDL_PIVOT_TABLE_EXT]: ext => {
      // Defaults for fields added later; attributes beyond these two are ignored, not recorded (no effect on the saved pivot).
      const a = attrs(getChild(ext, 'pivotTableDefinition16'))
      const ignore = () => {}
      table.defaultLayout = {}
      const enabled = optBool(a.EnabledSubtotalsDefault, ignore), top = optBool(a.SubtotalsOnTopDefault, ignore)
      if (enabled !== undefined) table.defaultLayout.enabledSubtotalsDefault = enabled
      if (top !== undefined) table.defaultLayout.subtotalsOnTopDefault = top
    },
  }, [], u)
  table.unsupported = u.list()
  countDiagnostics(u, partPath, diagnostics)
  return table
}


/** Invalid-index entries kept per table; the first one is named in the diagnostic. */
const MAX_INVALID_INDICES = 32
/** showDataAs modes that read `baseField`/`baseItem` (the rest leave them inert, often written as 0/0 against an item-less field). */
const USES_BASE_ITEM: readonly PivotDataField['showDataAs'][] = ['difference', 'percent', 'percentDiff']
const USES_BASE_FIELD: readonly PivotDataField['showDataAs'][] = [...USES_BASE_ITEM, 'runTotal']
const X14_USES_BASE_FIELD: readonly string[] = ['percentOfParent', 'percentOfRunningTotal', 'rankAscending', 'rankDescending']

/**
 * Index references checked against their field's domain once the cache is resolved (ruling AC). Returns
 * `invalid-index:<where>` entries; any entry makes the table display-only (§7.6), never laid out.
 * - pivotField item `x`: the cache field's group items, else its shared items;
 * - `rowFields`/`colFields`: a pivot field, each at most once across both axes, or -2 (Values) at most once and only
 *   with data fields;
 * - `rowItems`/`colItems` `x` at level `r + j`: that axis field's `items` (data fields for -2); `i`: a data field.
 *   `r + x.length` never passes the axis depth (one level on an empty axis). A grand line's `x` is a placeholder
 *   and only has to be non-negative;
 * - `pageField`: `fld` a pivot field, `item` one of its items;
 * - `dataField`: the effective field (`dataFieldSource`: x14 `sourceField`, else `fld`) a pivot field; `baseField`/`baseItem` where showDataAs uses them,
 *   keeping the previous/next sentinels;
 * - every pivot area (formats, conditional and chart formats, `autoSortScope`): `pivotArea@field` and
 *   `reference@field` a pivot field, or -2 (Values) when the table has data fields. A reference's `x` indices are deferred: with `byPosition`
 *   they count visible positions and with `relative` they are offsets, so their domain depends on the layout,
 *   and a Values reference's `x` names data fields; Stage 2's area matching resolves them in context;
 * - cache `fieldGroup`: `base`/`par` cache fields, `discretePr` group items.
 */
export function validateIndices(table: PivotTableDefinition, cache: PivotCacheDefinition | undefined): string[] {
  const out: string[] = []
  const bad = (where: string) => { if (out.length < MAX_INVALID_INDICES) out.push(`invalid-index:${where}`) }
  const nFields = table.fields.length
  const isField = (f: number) => Number.isInteger(f) && f >= 0 && f < nFields
  const itemCount = (f: number) => table.fields[f]?.items?.length ?? 0
  if (cache && cache.fields.length !== nFields) bad('pivotFields')
  if (cache) {
    cache.fields.forEach((cf, f) => {
      const group = cf.fieldGroup
      if (!group) return
      if (group.base !== undefined && group.base >= cache.fields.length) bad(`cache.fieldGroup[${f}].base`)
      if (group.par !== undefined && group.par >= cache.fields.length) bad(`cache.fieldGroup[${f}].par`)
      const groups = group.groupItems?.length ?? 0
      group.discretePr?.forEach((x, k) => { if (x >= groups) bad(`cache.fieldGroup[${f}].discretePr[${k}]`) })
    })
    table.fields.forEach((field, f) => {
      const cf = cache.fields[f]
      if (!cf || !field.items) return
      const domain = cf.fieldGroup?.groupItems?.length ?? cf.sharedItems?.items.length ?? 0
      field.items.forEach((item, k) => { if (item.x !== undefined && item.x >= domain) bad(`pivotField[${f}].item[${k}].x`) })
    })
  }
  let values = 0
  // A field sits on the axes at most once (Excel moves, never copies, a field): a repeat would let a
  // line's item vector grow with the axis list instead of the field count.
  const onAxis = new Set<number>()
  const axisFields = (name: string, list: readonly number[]) => list.forEach((x, j) => {
    if (x === PIVOT_VALUES_FIELD) {
      if (++values > 1 || table.dataFields.length === 0) bad(`${name}[${j}]`)
    } else if (!isField(x) || onAxis.has(x)) bad(`${name}[${j}]`)
    else onAxis.add(x)
  })
  axisFields('rowFields', table.rowFields)
  axisFields('colFields', table.colFields)
  const dataCount = Math.max(1, table.dataFields.length)
  const axisItems = (name: string, items: readonly PivotAxisItem[] | undefined, axis: readonly number[]) => items?.forEach((line, n) => {
    if (line.i >= dataCount) bad(`${name}[${n}].i`)
    // A line writes levels r … r + x.length - 1 of the axis; a grand line's placeholder `<x/>` is legal on an
    // empty axis. Each `x` of any other line past the depth is reported below; a grand line, or a bare `r`, here.
    if (line.r + line.x.length > Math.max(1, axis.length) && (line.type === 'grand' || !line.x.length)) return bad(`${name}[${n}].x`)
    line.x.forEach((x, j) => {
      const level = line.r + j
      if (x < 0) return bad(`${name}[${n}].x[${j}]`)
      if (line.type === 'grand') return
      const f = axis[level]
      const domain = f === undefined ? 0 : f === PIVOT_VALUES_FIELD ? table.dataFields.length : itemCount(f)
      if (x >= domain) bad(`${name}[${n}].x[${j}]`)
    })
  })
  axisItems('rowItems', table.rowItems, table.rowFields)
  axisItems('colItems', table.colItems, table.colFields)
  table.pageFields.forEach((page, n) => {
    if (!isField(page.fld)) bad(`pageField[${n}].fld`)
    else if (page.item !== undefined && page.item >= itemCount(page.fld)) bad(`pageField[${n}].item`)
  })
  // Values (-2, written 4294967294 on references) is a legal area field only when the table has data fields.
  const areaField = (f: number) => f === PIVOT_VALUES_FIELD ? table.dataFields.length > 0 : isField(f)
  const area = (where: string, a: PivotArea | undefined) => {
    if (!a) return
    if (a.field !== undefined && !areaField(a.field)) bad(`${where}.field`)
    a.references.forEach((ref, j) => { if (ref.field !== undefined && !areaField(ref.field)) bad(`${where}.reference[${j}].field`) })
  }
  table.fields.forEach((field, f) => area(`pivotField[${f}].autoSortScope`, field.autoSortScope))
  table.formats.forEach((format, n) => area(`format[${n}].pivotArea`, format.pivotArea))
  table.conditionalFormats.forEach((format, n) => format.pivotAreas.forEach((a, k) => area(`conditionalFormat[${n}].pivotArea[${k}]`, a)))
  table.chartFormats.forEach((format, n) => area(`chartFormat[${n}].pivotArea`, format.pivotArea))
  table.dataFields.forEach((data, n) => {
    // The effective field: a present x14 sourceField overrides fld (MS-XLSX), so fld is checked only without one.
    if (!isField(dataFieldSource(data))) bad(data.x14?.sourceField !== undefined ? `dataField[${n}].x14.sourceField` : `dataField[${n}].fld`)
    const showAs = data.x14?.pivotShowAs
    if (!USES_BASE_FIELD.includes(data.showDataAs) && !(showAs && X14_USES_BASE_FIELD.includes(showAs))) return
    if (!isField(data.baseField)) return bad(`dataField[${n}].baseField`)
    if (!USES_BASE_ITEM.includes(data.showDataAs)) return
    const sentinel = data.baseItem === PIVOT_BASE_ITEM_PREVIOUS || data.baseItem === PIVOT_BASE_ITEM_NEXT
    if (!sentinel && data.baseItem >= itemCount(data.baseField)) bad(`dataField[${n}].baseItem`)
  })
  return out
}
