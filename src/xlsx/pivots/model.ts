/**
 * Typed pivot model (spec §3): cache definitions and pivot table definitions as
 * written, before any layout. Lazy chunk `pivots`: no eager module may import
 * this file except as `import type`.
 *
 * Attribute defaults are ECMA-376 Part 1 §18.10 schema defaults (the XSD's
 * `default=`), never values a particular writer happens to emit. Where the
 * schema declares no default the field is optional, or the comment says what
 * an absent attribute means.
 */
import type { PivotRegionRef } from './stub'

/** `{ kind, feature: 'pivot-…', message }`, the shape `XlsxSheet.diagnostics` carries. */
export interface PivotDiagnostic { kind: 'unsupported' | 'parse'; feature: string; message: string; partPath?: string }

/** The Values pseudo-field in `rowFields`/`colFields` (`x="-2"`) and in pivot-area references (written unsigned as 4294967294). */
export const PIVOT_VALUES_FIELD = -2
/** `dataField@baseItem` sentinels: previous and next item ("difference from (previous)"), and the schema default. */
export const PIVOT_BASE_ITEM_PREVIOUS = 1048828
export const PIVOT_BASE_ITEM_NEXT = 1048829
export const PIVOT_BASE_ITEM_DEFAULT = 1048832

// ---------------------------------------------------------------- values

/** A cache value (`m` `n` `b` `e` `s` `d`). Dates stay as the ISO dateTime text written. */
export type PivotValue =
  | { type: 'missing' }
  | { type: 'number'; value: number }
  | { type: 'boolean'; value: boolean }
  | { type: 'error'; value: string }
  | { type: 'string'; value: string }
  | { type: 'date'; value: string }

/** A shared or group item: its value plus the CT_* item flags `u` (unused), `f` (calculated) and `c` (caption). */
export type PivotCacheItem = PivotValue & { unused?: true; calculated?: true; caption?: string }

// ---------------------------------------------------------------- cache definition

export type PivotCacheSourceType = 'worksheet' | 'external' | 'consolidation' | 'scenario'
export interface PivotCacheSource {
  type: PivotCacheSourceType
  connectionId: number
  /** `worksheetSource`: a range (`ref` + `sheet`) or a defined name / table (`name`); `relId` points at an external workbook. */
  worksheet?: { ref?: string; name?: string; sheet?: string; relId?: string }
  consolidation?: { autoPage: boolean }
}

export interface PivotSharedItems {
  containsSemiMixedTypes: boolean
  containsNonDate: boolean
  containsDate: boolean
  containsString: boolean
  containsBlank: boolean
  containsMixedTypes: boolean
  containsNumber: boolean
  containsInteger: boolean
  longText: boolean
  minValue?: number
  maxValue?: number
  minDate?: string
  maxDate?: string
  /** No item list: the field's records carry values inline (`n`, `s`, `d`, …) instead of `x` indices into `items`. */
  inline: boolean
  /** Document order, which is the index space of record `x` values. */
  items: PivotCacheItem[]
}

export type PivotGroupBy = 'range' | 'seconds' | 'minutes' | 'hours' | 'days' | 'months' | 'quarters' | 'years'
export interface PivotRangeGrouping {
  autoStart: boolean
  autoEnd: boolean
  groupBy: PivotGroupBy
  startNum?: number
  endNum?: number
  startDate?: string
  endDate?: string
  groupInterval: number
}
export interface PivotFieldGroup {
  /** Parent group field index. */
  par?: number
  /** Base field index the groups derive from. */
  base?: number
  rangePr?: PivotRangeGrouping
  /** `discretePr`: for each base item, the index of its group item. */
  discretePr?: number[]
  groupItems?: PivotCacheItem[]
}

export interface PivotCacheField {
  name: string
  caption?: string
  propertyName?: string
  serverField: boolean
  uniqueList: boolean
  numFmtId?: number
  /** A calculated field's formula (with `databaseField="0"`). */
  formula?: string
  sqlType: number
  hierarchy: number
  level: number
  databaseField: boolean
  mappingCount?: number
  memberPropertyField: boolean
  sharedItems?: PivotSharedItems
  fieldGroup?: PivotFieldGroup
}

/** x14 `pivotCacheDefinition` extension. `pivotCacheId` is what slicer caches' `tabular@pivotCacheId` name. */
export interface PivotCacheX14 {
  pivotCacheId?: number
  slicerData: boolean
  supportSubqueryNonVisual: boolean
  supportSubqueryCalcMem: boolean
  supportAddCalcMems: boolean
}

export interface PivotCacheDefinition {
  cacheId: number
  partPath: string
  /** Resolved from `r:id`. Stage 1 never reads it; its absence matters for §7.6 and §7.7. */
  recordsPartPath?: string
  source: PivotCacheSource
  fields: PivotCacheField[]
  invalid: boolean
  saveData: boolean
  refreshOnLoad: boolean
  optimizeMemory: boolean
  enableRefresh: boolean
  refreshedBy?: string
  refreshedDate?: number
  refreshedDateIso?: string
  backgroundQuery: boolean
  missingItemsLimit?: number
  createdVersion: number
  refreshedVersion: number
  minRefreshableVersion: number
  recordCount?: number
  upgradeOnRefresh: boolean
  supportSubquery: boolean
  supportAdvancedDrill: boolean
  /** An OLAP / data-model cache: `cacheHierarchies`, `dimensions`, `measureGroups`, `kpis`, `tupleCache` or `calculatedMembers`, or `@tupleCache`. */
  olap: boolean
  x14?: PivotCacheX14
  /** `calculatedItems` (§5.6): parsed for the engine's `pivot-calculated-item` refusal; not evaluated. */
  calculatedItems?: PivotCalculatedItem[]
  /** Constructs this parser does not model (see `PivotTableDefinition.unsupported`). */
  unsupported: string[]
}

/** A calculated item (CT_CalculatedItem): its formula and the pivot area that names its field and item. */
export interface PivotCalculatedItem { field?: number; formula: string; area: PivotArea }

// ---------------------------------------------------------------- pivot table definition

export type PivotAxis = 'axisRow' | 'axisCol' | 'axisPage' | 'axisValues'
export type PivotItemType = 'data' | 'default' | 'sum' | 'countA' | 'avg' | 'max' | 'min' | 'product' | 'count' | 'stdDev' | 'stdDevP' | 'var' | 'varP' | 'grand' | 'blank'
/** Named subtotal functions: `sumSubtotal`… on a field or pivot-area reference. */
export type PivotSubtotalFunction = 'sum' | 'countA' | 'avg' | 'max' | 'min' | 'product' | 'count' | 'stdDev' | 'stdDevP' | 'var' | 'varP'
export type PivotDataFunction = 'average' | 'count' | 'countNums' | 'max' | 'min' | 'product' | 'stdDev' | 'stdDevp' | 'sum' | 'var' | 'varp'
export type PivotShowDataAs = 'normal' | 'difference' | 'percent' | 'percentDiff' | 'runTotal' | 'percentOfRow' | 'percentOfCol' | 'percentOfTotal' | 'index'
/** x14 `dataField@pivotShowAs`, which supersedes `showDataAs` when present. */
export type PivotShowAsX14 = 'percentOfParent' | 'percentOfParentRow' | 'percentOfParentCol' | 'percentOfRunningTotal' | 'rankAscending' | 'rankDescending'
export type PivotFieldSortType = 'manual' | 'ascending' | 'descending'

export interface PivotLocation {
  ref: string
  firstHeaderRow: number
  firstDataRow: number
  firstDataCol: number
  rowPageCount: number
  colPageCount: number
}

/** `pivotField/items/item`. */
export interface PivotFieldItem {
  type: PivotItemType
  /** Shared-item index (absent on subtotal/grand/blank items). */
  x?: number
  /** `n`: the user caption. */
  name?: string
  hidden: boolean
  /** `s`: the item has a character value. */
  char: boolean
  /** `sd`: details shown; false is a collapsed item. */
  showDetails: boolean
  calculated: boolean
  missing: boolean
  /** `c`: the item has children. */
  childItems: boolean
  /** `d`: expanded (drilled across). */
  expanded: boolean
  /** `e`: drill-across attributes. */
  drillAcross: boolean
}

export interface PivotField {
  name?: string
  axis?: PivotAxis
  dataField: boolean
  subtotalCaption?: string
  showDropDowns: boolean
  hiddenLevel: boolean
  uniqueMemberProperty?: string
  compact: boolean
  allDrilled: boolean
  numFmtId?: number
  outline: boolean
  subtotalTop: boolean
  dragToRow: boolean
  dragToCol: boolean
  multipleItemSelectionAllowed: boolean
  dragToPage: boolean
  dragToData: boolean
  dragOff: boolean
  showAll: boolean
  insertBlankRow: boolean
  serverField: boolean
  insertPageBreak: boolean
  autoShow: boolean
  topAutoShow: boolean
  hideNewItems: boolean
  measureFilter: boolean
  includeNewItemsInFilter: boolean
  itemPageCount: number
  sortType: PivotFieldSortType
  dataSourceSort?: boolean
  nonAutoSortDefault: boolean
  rankBy?: number
  defaultSubtotal: boolean
  /** The named subtotal flags that are set (`sumSubtotal="1"` → 'sum'), in schema order. */
  subtotals: PivotSubtotalFunction[]
  showPropCell: boolean
  showPropTip: boolean
  showPropAsCaption: boolean
  defaultAttributeDrillState: boolean
  items?: PivotFieldItem[]
  autoSortScope?: PivotArea
  x14?: { fillDownLabels: boolean; ignore: boolean }
}

/** `rowItems/i` and `colItems/i`: `r` leading levels repeat the previous line, then `x` item indices; `i` is the data-field index. */
export interface PivotAxisItem { type: PivotItemType; r: number; i: number; x: number[] }

export interface PivotPageField { fld: number; item?: number; hier?: number; name?: string; cap?: string }

export interface PivotDataField {
  name?: string
  fld: number
  subtotal: PivotDataFunction
  showDataAs: PivotShowDataAs
  baseField: number
  /** Kept as written, including the previous/next sentinels 1048828 and 1048829. */
  baseItem: number
  numFmtId?: number
  /** x14 `dataField`. `sourceField`, when present, overrides `fld` (MS-XLSX); read the field through `dataFieldSource`. */
  x14?: { pivotShowAs?: PivotShowAsX14; sourceField?: number }
}

/** The pivot field a data field aggregates: x14 `sourceField` when present, else `fld`. Consumers use this, never `fld` directly. */
export function dataFieldSource(field: PivotDataField): number {
  return field.x14?.sourceField ?? field.fld
}

export type PivotAreaType = 'none' | 'normal' | 'data' | 'all' | 'origin' | 'button' | 'topEnd' | 'topRight'
export interface PivotAreaReference {
  /** Field index; the Values field is normalised from 4294967294 to -2. */
  field?: number
  selected: boolean
  byPosition: boolean
  relative: boolean
  defaultSubtotal: boolean
  subtotals: PivotSubtotalFunction[]
  x: number[]
}
export interface PivotArea {
  field?: number
  type: PivotAreaType
  dataOnly: boolean
  labelOnly: boolean
  grandRow: boolean
  grandCol: boolean
  cacheIndex: boolean
  outline: boolean
  offset?: string
  collapsedLevelsAreSubtotals: boolean
  axis?: PivotAxis
  fieldPosition?: number
  references: PivotAreaReference[]
}

export interface PivotFormat { action: 'blank' | 'formatting' | 'drill' | 'formula'; dxfId?: number; pivotArea: PivotArea }
export interface PivotConditionalFormat { scope: 'selection' | 'data' | 'field'; type: 'none' | 'all' | 'row' | 'column'; priority: number; pivotAreas: PivotArea[] }
export interface PivotChartFormat { chart: number; format: number; series: boolean; pivotArea: PivotArea }

export type PivotFilterType =
  | 'unknown' | 'count' | 'percent' | 'sum'
  | 'captionEqual' | 'captionNotEqual' | 'captionBeginsWith' | 'captionNotBeginsWith' | 'captionEndsWith' | 'captionNotEndsWith'
  | 'captionContains' | 'captionNotContains' | 'captionGreaterThan' | 'captionGreaterThanOrEqual' | 'captionLessThan'
  | 'captionLessThanOrEqual' | 'captionBetween' | 'captionNotBetween'
  | 'valueEqual' | 'valueNotEqual' | 'valueGreaterThan' | 'valueGreaterThanOrEqual' | 'valueLessThan' | 'valueLessThanOrEqual'
  | 'valueBetween' | 'valueNotBetween'
  | 'dateEqual' | 'dateNotEqual' | 'dateOlderThan' | 'dateOlderThanOrEqual' | 'dateNewerThan' | 'dateNewerThanOrEqual'
  | 'dateBetween' | 'dateNotBetween'
  | 'tomorrow' | 'today' | 'yesterday' | 'nextWeek' | 'thisWeek' | 'lastWeek' | 'nextMonth' | 'thisMonth' | 'lastMonth'
  | 'nextQuarter' | 'thisQuarter' | 'lastQuarter' | 'nextYear' | 'thisYear' | 'lastYear' | 'yearToDate'
  | 'Q1' | 'Q2' | 'Q3' | 'Q4' | 'M1' | 'M2' | 'M3' | 'M4' | 'M5' | 'M6' | 'M7' | 'M8' | 'M9' | 'M10' | 'M11' | 'M12'

/** The `autoFilter` a pivot filter carries: one column holding its criteria. */
export interface PivotFilterColumn {
  colId: number
  hiddenButton: boolean
  showButton: boolean
  top10?: { top: boolean; percent: boolean; val: number; filterVal?: number }
  customFilters?: { and: boolean; filters: Array<{ operator: 'equal' | 'lessThan' | 'lessThanOrEqual' | 'notEqual' | 'greaterThanOrEqual' | 'greaterThan'; val?: string }> }
  dynamicFilter?: { type: string; val?: number; valIso?: string; maxVal?: number; maxValIso?: string }
  filters?: { blank: boolean; values: string[] }
}
export interface PivotFilter {
  fld: number
  mpFld?: number
  type: PivotFilterType
  evalOrder: number
  id: number
  iMeasureHier?: number
  iMeasureFld?: number
  name?: string
  description?: string
  stringValue1?: string
  stringValue2?: string
  autoFilter: { ref?: string; columns: PivotFilterColumn[] }
}

/** `pivotTableStyleInfo`. ECMA declares no defaults for these flags; an absent flag is off. */
export interface PivotTableStyleInfo {
  name?: string
  showRowHeaders: boolean
  showColHeaders: boolean
  showRowStripes: boolean
  showColStripes: boolean
  showLastColumn: boolean
}

/** x14 `pivotTableDefinition` extension (MS-XLSX CT_PivotTableDefinition). */
export interface PivotTableX14 {
  fillDownLabelsDefault: boolean
  visualTotalsForSets: boolean
  calculatedMembersInFilters: boolean
  altText?: string
  altTextSummary?: string
  enableEdit: boolean
  autoApply: boolean
  allocationMethod: 'equalAllocation' | 'equalIncrement' | 'weightedAllocation' | 'weightedIncrement'
  weightExpression?: string
  hideValuesRow: boolean
}

export interface PivotTableDefinition {
  /** The sheet the region map listed this table under. */
  sheet: string
  partPath: string
  region: PivotRegionRef
  cacheId: number
  /** The cache part, from the workbook's `pivotCaches` (or, failing that, the definition's own relationship). */
  cachePartPath?: string

  name: string
  dataOnRows: boolean
  dataPosition?: number
  autoFormatId?: number
  /** AG_AutoFormat flags: optional with no schema default; absent is false. */
  applyNumberFormats: boolean
  applyBorderFormats: boolean
  applyFontFormats: boolean
  applyPatternFormats: boolean
  applyAlignmentFormats: boolean
  applyWidthHeightFormats: boolean
  dataCaption: string
  grandTotalCaption?: string
  errorCaption?: string
  showError: boolean
  missingCaption?: string
  showMissing: boolean
  pageStyle?: string
  pivotTableStyle?: string
  vacatedStyle?: string
  tag?: string
  updatedVersion: number
  minRefreshableVersion: number
  asteriskTotals: boolean
  showItems: boolean
  editData: boolean
  disableFieldList: boolean
  showCalcMbrs: boolean
  visualTotals: boolean
  showMultipleLabel: boolean
  showDataDropDown: boolean
  showDrill: boolean
  printDrill: boolean
  showMemberPropertyTips: boolean
  showDataTips: boolean
  enableWizard: boolean
  enableDrill: boolean
  enableFieldProperties: boolean
  preserveFormatting: boolean
  useAutoFormatting: boolean
  pageWrap: number
  pageOverThenDown: boolean
  subtotalHiddenItems: boolean
  rowGrandTotals: boolean
  colGrandTotals: boolean
  fieldPrintTitles: boolean
  itemPrintTitles: boolean
  mergeItem: boolean
  showDropZones: boolean
  createdVersion: number
  indent: number
  showEmptyRow: boolean
  showEmptyCol: boolean
  showHeaders: boolean
  /** Table-level layout defaults: `compact` true but `outline` false (fields default both true). */
  compact: boolean
  outline: boolean
  outlineData: boolean
  compactData: boolean
  published: boolean
  gridDropZones: boolean
  immersive: boolean
  multipleFieldFilters: boolean
  chartFormat: number
  rowHeaderCaption?: string
  colHeaderCaption?: string
  fieldListSortAscending: boolean
  mdxSubqueries: boolean
  customListSort: boolean

  location: PivotLocation
  fields: PivotField[]
  /** Field indices; `PIVOT_VALUES_FIELD` (-2) is the Values pseudo-field. */
  rowFields: number[]
  colFields: number[]
  /** The saved layout. Undefined when the part has no `rowItems`/`colItems` (other writers omit them). */
  rowItems?: PivotAxisItem[]
  colItems?: PivotAxisItem[]
  pageFields: PivotPageField[]
  dataFields: PivotDataField[]
  formats: PivotFormat[]
  conditionalFormats: PivotConditionalFormat[]
  chartFormats: PivotChartFormat[]
  filters: PivotFilter[]
  styleInfo?: PivotTableStyleInfo
  x14?: PivotTableX14
  /** xpdl `pivotTableDefinition16`: defaults Excel applies to fields added later (Stage 4); no effect on the saved layout. */
  defaultLayout?: { enabledSubtotalsDefault?: boolean; subtotalsOnTopDefault?: boolean }

  /**
   * Every construct the parser does not model, as a short path (`pivotField@foo`,
   * `pivotTableDefinition/pivotHierarchies`, `ext{uri}`, `dataField@subtotal=bogus`),
   * deduplicated. Includes the cache's own list, prefixed `cache:`. Non-empty makes
   * the pivot display-only (§7.6); display never depends on it.
   */
  unsupported: string[]
}

export interface PivotWorkbook {
  caches: Map<number, PivotCacheDefinition>
  tables: PivotTableDefinition[]
  diagnostics: PivotDiagnostic[]
}

// ---------------------------------------------------------------- pivot charts

/** `c:pivotSource`: the pivot a chart is bound to. */
export interface PivotChartSource { name: string; fmtId: number }
/** x14 `pivotOptions` on a pivot chart: which field-button drop zones show. */
export interface PivotChartOptions {
  dropZoneFilter: boolean
  dropZoneCategories: boolean
  dropZoneData: boolean
  dropZoneSeries: boolean
  dropZonesVisible: boolean
}
