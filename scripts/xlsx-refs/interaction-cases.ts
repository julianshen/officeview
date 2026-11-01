/**
 * Stage 3 Excel interaction oracle (pivots plan Task 24): the seeds and cases that `build-engine.ts --interaction`
 * turns into Excel-saved workbooks under tests/fixtures/xlsx-pivots/interaction/, each with the `PivotDelta` list the
 * pivot session (Task 25) applies to the case's starting state to reproduce Excel's saved result.
 *
 * Pure module (no I/O, no Node APIs): the tests import it, and `tsconfig.json` type-checks it through them.
 *
 * A case starts from its seed's file, or from `base`: an Excel-saved file of the same seed family (a Stage 2 case
 * under engine/, an earlier interaction case, or a seed variant under recipes/). It then takes one of the Stage 2
 * routes (engine-cases.ts): 'applescript' (`lines`, then `refresh table`, save) or 'xml' (`xmlPrep`, saved; the part
 * edit `xmlEdit`; Excel opens it, runs `post`, `refresh table`, saves). Markers prove Excel kept the change.
 *
 * `delta` is written by name (`DeltaSpec`: field and item captions) and resolved to engine indices against the start
 * file when the manifest is written (`resolveDelta`): a pivot-item index is the position in the field's `<items>`
 * list (else the cache's shared-item order), which on the chart seed is not the shared index (Region lists x=3,1,0,2).
 *
 * A case with `setup` made a change no `PivotDelta` expresses (a field move, a layout option, removed explicit
 * captions, column widths). Its saved file is a measurement, or the `base` of later cases; it is not a delta oracle.
 */
import {
  PIVOT_BASE_ITEM_NEXT, PIVOT_BASE_ITEM_PREVIOUS, PIVOT_VALUES_FIELD,
  type PivotDataFunction, type PivotFilter,
} from '../../src/xlsx/pivots/model'
import type { PivotDateGroupPeriod, PivotDelta, PivotValueShowAs } from '../../src/xlsx/pivots/session/delta'
import {
  F, SALES, SEED_PIVOT_MARKERS, UNITS, addFilter, attrOf, autoSort, childrenOf, colField, custom, dataField, editItem, editPivotField, editRoot,
  fieldAttr, fieldHas, fieldItem, fieldItems, find, form, hide, hideItem, labelFilter, must, onlyRows, rowField, setFieldItems, sharedItemValues,
  splice, str, subtotalsAt, toCols, toPage, toRows, valueFilter,
  type PivotParts, type Target, field as pivotField,
} from './engine-cases'

// ---------------------------------------------------------------------------------------------------------------
// PivotDelta (Tasks 25 & 35–37)
// ---------------------------------------------------------------------------------------------------------------

/** Task 25 & Stage 4's delta type (src/xlsx/pivots/session/delta.ts), the manifest's `delta` entries (ruling S3-A). */
export type { PivotDelta }

/** An item named by its caption (a number item by its value, e.g. Year 2023). */
export type ItemName = string | number
/** A `PivotDelta` with fields and items named; `resolveDelta` turns it into engine indices. */
export type DeltaSpec =
  | { kind: 'collapse'; field: string; item: ItemName; collapsed: boolean }
  | { kind: 'collapseField'; field: string; collapsed: boolean }
  | { kind: 'collapseAll'; collapsed: boolean }
  | { kind: 'items'; field: string; hidden: readonly ItemName[] }
  | { kind: 'page'; field: string; selected: readonly ItemName[] | 'all' }
  | { kind: 'filter'; field: string; filter: Omit<PivotFilter, 'fld'> }
  | { kind: 'clearFilters'; field: string }
  | { kind: 'sort'; field: string; order: 'ascending' | 'descending'; byData?: number }
  | { kind: 'moveField'; field: string; axis: 'row' | 'col' | 'page' | 'none'; index?: number }
  | { kind: 'addDataField'; field: string; index?: number }
  | { kind: 'removeDataField'; dataIndex: number }
  | { kind: 'moveDataField'; dataIndex: number; toIndex: number }
  | {
      kind: 'valueSettings'
      dataIndex: number
      fn?: PivotDataFunction
      showAs?: PivotValueShowAs
      baseField?: string
      baseItem?: '(previous)' | '(next)' | ItemName
      name?: string
      numFmtId?: number | null
    }
  | { kind: 'groupDate'; field: string; periods: readonly PivotDateGroupPeriod[]; startDate?: number; endDate?: number; daysInterval?: number }
  | { kind: 'groupNumber'; field: string; start?: number; end?: number; by: number }
  | { kind: 'ungroup'; field: string }

/** What resolution needs from the start file: cache field names, and each field's pivot items as captions ('' for a
 * subtotal or blank item; undefined when the field has no `<items>` list, so the shared items stand in). */
export interface StartShape { fields: readonly string[]; items: ReadonlyArray<readonly string[] | undefined>; shared: ReadonlyArray<readonly string[]> }

export function resolveDelta(spec: DeltaSpec, start: StartShape): PivotDelta {
  const fieldOf = (name: string) => {
    const index = start.fields.indexOf(name)
    if (index < 0) throw new Error(`delta: field ${name} not in the start file`)
    return index
  }
  const fieldOrValuesOf = (name: string) => name === 'Values' ? PIVOT_VALUES_FIELD : fieldOf(name)
  const itemOf = (field: number, name: ItemName) => {
    const list = start.items[field] ?? start.shared[field] ?? []
    const index = list.indexOf(String(name))
    if (index < 0) throw new Error(`delta: item ${name} of ${start.fields[field]} not in the start file`)
    return index
  }
  switch (spec.kind) {
    case 'collapse': { const f = fieldOf(spec.field); return { kind: 'collapse', field: f, item: itemOf(f, spec.item), collapsed: spec.collapsed } }
    case 'collapseField': return { kind: 'collapseField', field: fieldOf(spec.field), collapsed: spec.collapsed }
    case 'collapseAll': return { kind: 'collapseAll', collapsed: spec.collapsed }
    case 'items': { const f = fieldOf(spec.field); return { kind: 'items', field: f, hidden: spec.hidden.map(i => itemOf(f, i)).sort((a, b) => a - b) } }
    case 'page': { const f = fieldOf(spec.field); return { kind: 'page', field: f, selected: spec.selected === 'all' ? 'all' : spec.selected.map(i => itemOf(f, i)).sort((a, b) => a - b) } }
    case 'filter': { const f = fieldOf(spec.field); return { kind: 'filter', field: f, filter: { fld: f, ...spec.filter } } }
    case 'clearFilters': return { kind: 'clearFilters', field: fieldOf(spec.field) }
    case 'sort': return { kind: 'sort', field: fieldOf(spec.field), order: spec.order, ...(spec.byData !== undefined ? { byData: spec.byData } : {}) }
    case 'moveField':
      return {
        kind: 'moveField',
        field: fieldOrValuesOf(spec.field),
        axis: spec.axis,
        ...(spec.index !== undefined ? { index: spec.index } : {}),
      }
    case 'addDataField':
      return { kind: 'addDataField', field: fieldOf(spec.field), ...(spec.index !== undefined ? { index: spec.index } : {}) }
    case 'removeDataField':
      return { kind: 'removeDataField', dataIndex: spec.dataIndex }
    case 'moveDataField':
      return { kind: 'moveDataField', dataIndex: spec.dataIndex, toIndex: spec.toIndex }
    case 'valueSettings': {
      const baseField = spec.baseField !== undefined ? fieldOf(spec.baseField) : undefined
      const baseItem = spec.baseItem === undefined ? undefined
        : spec.baseItem === '(previous)' ? PIVOT_BASE_ITEM_PREVIOUS
          : spec.baseItem === '(next)' ? PIVOT_BASE_ITEM_NEXT
            : (() => {
                if (baseField === undefined) throw new Error('delta: baseItem requires baseField')
                return itemOf(baseField, spec.baseItem)
              })()
      return {
        kind: 'valueSettings',
        dataIndex: spec.dataIndex,
        ...(spec.fn !== undefined ? { fn: spec.fn } : {}),
        ...(spec.showAs !== undefined ? { showAs: spec.showAs } : {}),
        ...(baseField !== undefined ? { baseField } : {}),
        ...(baseItem !== undefined ? { baseItem } : {}),
        ...(spec.name !== undefined ? { name: spec.name } : {}),
        ...(spec.numFmtId !== undefined ? { numFmtId: spec.numFmtId } : {}),
      }
    }
    case 'groupDate':
      return {
        kind: 'groupDate',
        field: fieldOf(spec.field),
        periods: [...spec.periods],
        ...(spec.startDate !== undefined ? { startDate: spec.startDate } : {}),
        ...(spec.endDate !== undefined ? { endDate: spec.endDate } : {}),
        ...(spec.daysInterval !== undefined ? { daysInterval: spec.daysInterval } : {}),
      }
    case 'groupNumber':
      return {
        kind: 'groupNumber',
        field: fieldOf(spec.field),
        ...(spec.start !== undefined ? { start: spec.start } : {}),
        ...(spec.end !== undefined ? { end: spec.end } : {}),
        by: spec.by,
      }
    case 'ungroup':
      return { kind: 'ungroup', field: fieldOf(spec.field) }
  }
}

/** The start shape read off saved parts (pivot and cache definitions), for the tests and the dry run. */
export function startShapeOf(parts: PivotParts): StartShape {
  const fieldsEl = must(find(parts.cache, ['cacheFields']), 'cacheFields')
  const fields = childrenOf(parts.cache, fieldsEl, 'cacheField').map(f => attrOf(parts.cache, f, 'name') ?? '')
  const shared = fields.map((_, i) => sharedItemValues(parts.cache, i).map(v => v ?? ''))
  const groupLabels = (i: number) => {
    const g = find(parts.cache, ['cacheFields', ['cacheField', i], 'fieldGroup', 'groupItems'])
    return g ? childrenOf(parts.cache, g).map(item => item.name === 'm' ? '' : attrOf(parts.cache, item, 'v') ?? '') : undefined
  }
  const items = fields.map((_, i) => {
    if (!find(parts.pivot, ['pivotFields', ['pivotField', i], 'items'])) return undefined
    const labels = groupLabels(i) ?? shared[i]!
    return fieldItems(parts.pivot, i).map(item => {
      if (/\st="(?!data)/.test(item)) return ''
      const n = /\sn="([^"]*)"/.exec(item)
      if (n) return n[1]!
      const x = /\sx="(\d+)"/.exec(item)
      return x ? labels[Number(x[1])] ?? '' : ''
    })
  })
  return { fields, items, shared }
}

// ---------------------------------------------------------------------------------------------------------------
// Seeds
// ---------------------------------------------------------------------------------------------------------------

export type InteractionSeed = 'engine-seed' | 'chart-seed' | 'getpivotdata-seed' | 'refresh-seed'
/** Each seed's file, relative to tests/fixtures/xlsx-pivots (without .xlsx). */
export const SEED_FILES: Record<InteractionSeed, string> = {
  'engine-seed': 'recipes/engine-seed',
  'chart-seed': 'recipes/pivot-chart-seed',
  'getpivotdata-seed': 'recipes/getpivotdata-seed',
  'refresh-seed': 'recipes/refresh-seed',
}
/** Sheet-level proof: the named sheet's XML matches (or, with `absent`, must not match) `re`. */
export interface SheetMarker { sheet: string; re: RegExp; absent?: boolean }

/** A seed `build-engine.ts --interaction` makes from the engine seed. */
export interface SeedBuild {
  id: 'getpivotdata-seed' | 'refresh-seed' | 'refresh-seed-nodata'
  /** Output, relative to tests/fixtures/xlsx-pivots (without .xlsx). */
  file: string
  lines: (t: Target) => string[]
  /** false: save without `refresh table`, so the records stay as they were (the refresh seeds' stale records). */
  refresh: boolean
  markers: RegExp[]
  cacheMarkers?: RegExp[]
  recordMarkers?: RegExp[]
  recordAbsent?: RegExp[]
  sheetMarkers?: SheetMarker[]
  /** Records part must be absent (saveData="0"). */
  noRecords?: boolean
  /** When the AppleScript route fails: `fallbackLines` (the reachable part), then this cache edit; `dropRecords`
   * also removes the records part, its relationship and content-type override. */
  fallback?: { lines: (t: Target) => string[]; cache: Record<string, string>; dropRecords?: boolean }
  probe: string
  measures: string
}

/** GETPIVOTDATA formulas on sheet "Formulas" (column A the label, column B the formula; the pivot anchor is Pivot!A3). */
export const GPD_FORMULAS: ReadonlyArray<readonly [string, string]> = [
  ['total', '=GETPIVOTDATA("Sum of Sales",Pivot!$A$3)'],
  ['item', '=GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region","East")'],
  ['two fields', '=GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region","East","Product","Apples")'],
  ['column item', '=GETPIVOTDATA("Sum of Units",Pivot!$A$3,"Year",2023)'],
  ['row and column items', '=GETPIVOTDATA("Sum of Units",Pivot!$A$3,"Region","North","Year",2023)'],
  ['item hidden by gpd-hide-item (West)', '=GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region","West")'],
  ['detail under the item gpd-collapse-item collapses (East/Bananas)', '=GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region","East","Product","Bananas")'],
  ['item gpd-filter-out filters out (South/Cherries)', '=GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region","South","Product","Cherries")'],
  ['inner item alone, no subtotal line (Cherries)', '=GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Product","Cherries")'],
  ['item a Product filter leaves without records (Central)', '=GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region","Central")'],
  ['cell reference into the pivot', '=Pivot!B6'],
  ['range over pivot cells', '=SUM(Pivot!B6:B8)'],
]
export const GPD_SHEET = 'Formulas'
/** Refresh seeds: Data!A2 East → North and Data!D2 1234.5 → 1500.25 after the records were saved, so the source
 * differs from the records (East/Apples/2022 loses its only record on refresh). */
export const REFRESH_EDITS = { regionCell: 'A2', region: 'North', salesCell: 'D2', sales: 1500.25 } as const
const refreshDataLines = (t: Target) => [
  `set value of range "${REFRESH_EDITS.regionCell}" of ${t.data} to ${str(REFRESH_EDITS.region)}`,
  `set value of range "${REFRESH_EDITS.salesCell}" of ${t.data} to ${REFRESH_EDITS.sales}`,
]
const refreshOnOpen = (t: Target) => `set refresh on file open of (pivot cache of ${t.p}) to true`
const DATA_CELL_MARKERS: SheetMarker[] = [
  { sheet: 'Data', re: new RegExp(String.raw`<c r="${REFRESH_EDITS.salesCell}"[^>]*>(?:<f>[^<]*<\/f>)?<v>${REFRESH_EDITS.sales}<\/v>`) },
]

export const SEED_BUILDS: readonly SeedBuild[] = [
  {
    id: 'getpivotdata-seed', file: SEED_FILES['getpivotdata-seed'], refresh: true, probe: 'make-worksheet',
    lines: t => [
      `set gpdSheet to make new worksheet at end of ${t.wb}`,
      `set name of gpdSheet to ${str(GPD_SHEET)}`,
      ...GPD_FORMULAS.flatMap(([label, formula], i) => [
        `set value of range "A${i + 1}" of worksheet ${str(GPD_SHEET)} of ${t.wb} to ${str(label)}`,
        `set formula of range "B${i + 1}" of worksheet ${str(GPD_SHEET)} of ${t.wb} to ${str(formula)}`,
      ]),
    ],
    markers: [...SEED_PIVOT_MARKERS],
    sheetMarkers: [{ sheet: GPD_SHEET, re: /<f>GETPIVOTDATA\(/ }, { sheet: GPD_SHEET, re: /<c r="B1"[^>]*><f>[^<]*<\/f><v>/ }],
    measures: 'the engine seed plus GETPIVOTDATA formulas (and plain references into the pivot) on sheet Formulas (§7.5)',
  },
  {
    id: 'refresh-seed', file: SEED_FILES['refresh-seed'], refresh: false, probe: 'refresh-on-file-open',
    lines: t => [...refreshDataLines(t), refreshOnOpen(t)],
    markers: [...SEED_PIVOT_MARKERS], cacheMarkers: [/<pivotCacheDefinition [^>]*refreshOnLoad="1"/],
    recordAbsent: [/<n v="1500\.25"\/>/], sheetMarkers: DATA_CELL_MARKERS,
    fallback: { lines: refreshDataLines, cache: { refreshOnLoad: '1' } },
    measures: 'refreshOnLoad="1" with records older than the source (§7.7): the viewer shows the saved cells, a session refreshes from Data',
  },
  {
    id: 'refresh-seed-nodata', file: 'recipes/refresh-seed-nodata', refresh: false, probe: 'save-data',
    lines: t => [...refreshDataLines(t), refreshOnOpen(t), `set save data of ${t.p} to false`],
    markers: [...SEED_PIVOT_MARKERS], cacheMarkers: [/<pivotCacheDefinition [^>]*refreshOnLoad="1"/, /<pivotCacheDefinition [^>]*saveData="0"/],
    noRecords: true, sheetMarkers: DATA_CELL_MARKERS,
    fallback: { lines: refreshDataLines, cache: { refreshOnLoad: '1', saveData: '0' }, dropRecords: true },
    measures: 'refreshOnLoad="1" with saveData="0" (no records part, §7.7 CODEX-R2-2): only the source can rebuild the pivot',
  },
]

// ---------------------------------------------------------------------------------------------------------------
// Cases
// ---------------------------------------------------------------------------------------------------------------

export type InteractionTask = 25 | 29 | 30 | 31 | 32 | 35 | 36 | 37
export type InteractionRoute = 'applescript' | 'xml' | 'recipe'
export interface InteractionCase {
  id: string
  task: InteractionTask
  seed: InteractionSeed
  /** The start file when it is not the seed's own (relative to tests/fixtures/xlsx-pivots, without .xlsx). */
  base?: string
  /** A change no PivotDelta expresses, made before `delta` (see the header): not a delta oracle. */
  setup?: string
  route: InteractionRoute
  lines?: (t: Target) => string[]
  xmlPrep?: (t: Target) => string[]
  xmlEdit?: (parts: PivotParts) => PivotParts
  /** xml route: AppleScript run on the edited file before `refresh table`. */
  post?: (t: Target) => string[]
  fallback?: 'xml'
  /** UI clicks for docs/validation/xlsx-pivots-recipes-stage4.md when neither AppleScript nor verified XML reaches the case. */
  recipe?: string
  markers: RegExp[]
  cacheMarkers?: RegExp[]
  absent?: RegExp[]
  cacheAbsent?: RegExp[]
  recordMarkers?: RegExp[]
  sheetMarkers?: SheetMarker[]
  /** The saved file keeps a pivot chart (c:pivotSource) — the chart cases. */
  chart?: boolean
  delta: readonly DeltaSpec[]
  probe?: string
  measures: string
}

const CASES: InteractionCase[] = []
type Spec = Omit<InteractionCase, 'id' | 'task' | 'seed'> & { seed?: InteractionSeed }
function add(id: string, task: InteractionTask, spec: Spec): void {
  if (CASES.some(c => c.id === id)) throw new Error(`duplicate interaction case ${id}`)
  CASES.push({ id, task, seed: spec.seed ?? 'engine-seed', ...spec })
}

// Marker building blocks.
const COMPACT_OFF = /<pivotField [^>]*compact="0"/, SUBTOTAL_BOTTOM = /<pivotField [^>]*subtotalTop="0"/
const collapsedItem = (index: number) => fieldItem(index, String.raw`[^>]*\bsd="0"`)
/** An item of the field still expanded (a data item, with `x`, lacking sd="0"). */
const expandedItem = (index: number) => fieldItem(index, String.raw`(?![^>]*\bsd="0")[^>]*\bx="`)
const hiddenItem = (index: number) => fieldItem(index, String.raw`[^>]*\bh="1"`)
const ANY_FILTER = /<filter /
const AXES_KEPT = rowField(F.Region)

// AppleScript building blocks.
const showDetail = (t: Target, name: string, item: string, on: boolean) => `set show detail of pivot item ${str(item)} of ${pivotField(t, name)} to ${on}`
/** Field-level `show detail` (Excel.sdef pivot field property: "Collapse/Expand Entire Field"). */
const fieldDetail = (t: Target, name: string, on: boolean) => `set show detail of ${pivotField(t, name)} to ${on}`
const showItem = (t: Target, name: string, item: string) => `set visible of pivot item ${str(item)} of ${pivotField(t, name)} to true`
const clearAll = (t: Target, name: string) => `clear all filters (${pivotField(t, name)})`
const currentPage = (t: Target, name: string, item: string) => `set current page of ${pivotField(t, name)} to ${str(item)}`

// XML building blocks.
/** sd="0" on every item of a field, the default subtotal item included (or, with `on`, sd removed from every item):
 * Excel's own field-level collapse writes `<item sd="0" x="0"/>` … `<item t="default" sd="0"/>`
 * (interaction/interaction-collapse-field-compact-top.xlsx). */
function setFieldDetail(pivot: string, index: number, on: boolean): string {
  return setFieldItems(pivot, index, fieldItems(pivot, index).map(item => {
    const bare = item.replace(/\ssd="[01]"/, '')
    if (on) return bare
    return /\sx="/.test(bare) ? bare.replace('<item ', '<item sd="0" ') : bare.replace(/\s*\/>$/, ' sd="0"/>')
  }))
}
/** sd="0" on the item of `fld` that points at shared item `value`. */
function collapseItemXml(parts: PivotParts, fld: number, value: string): PivotParts {
  const x = sharedItemValues(parts.cache, fld).indexOf(value)
  if (x < 0) throw new Error(`xml edit: item ${value} of field ${fld} not found`)
  return { ...parts, pivot: editItem(parts.pivot, fld, x, { sd: '0' }) }
}
const unhideField = (pivot: string, index: number) => setFieldItems(pivot, index, fieldItems(pivot, index).map(item => item.replace(/\sh="1"/, '')))
function dropFilters(pivot: string, fld: number): string {
  const filters = find(pivot, ['filters'])
  if (!filters) return pivot
  const keep = childrenOf(pivot, filters, 'filter').filter(f => attrOf(pivot, f, 'fld') !== String(fld)).map(f => pivot.slice(f.start, f.end))
  const block = keep.length ? `<filters count="${keep.length}">${keep.join('')}</filters>` : ''
  return editPivotField(splice(pivot, filters.start, filters.end, block), fld, { measureFilter: null })
}
/** The page field's single selection removed: every item shown, "(All)". */
const pageAll = (pivot: string, fld: number) => pivot.replace(new RegExp(`(<pageField fld="${fld}"[^>]*?) item="\\d+"`), '$1')
/** Multiple items allowed on a page field, `hidden` hidden, the single selection removed. */
function pageMulti(parts: PivotParts, fld: number, hidden: readonly string[]): PivotParts {
  const shared = sharedItemValues(parts.cache, fld)
  const xs = new Set(hidden.map(v => { const x = shared.indexOf(v); if (x < 0) throw new Error(`xml edit: item ${v} of field ${fld} not found`); return x }))
  const items = fieldItems(parts.pivot, fld).map(item => {
    const x = /\sx="(\d+)"/.exec(item)
    return x && xs.has(Number(x[1])) ? item.replace('<item ', '<item h="1" ') : item
  })
  return { ...parts, pivot: pageAll(setFieldItems(editPivotField(parts.pivot, fld, { multipleItemSelectionAllowed: '1' }), fld, items), fld) }
}

// Filter literals for the delta (the manifest substitutes Excel's saved filter for the field when the case has one).
const filterLiteral = (type: PivotFilter['type'], body: Omit<PivotFilter['autoFilter']['columns'][number], 'colId' | 'hiddenButton' | 'showButton'>,
  extra: Partial<PivotFilter> = {}): Omit<PivotFilter, 'fld'> => ({
  type, evalOrder: -1, id: 0, ...extra, autoFilter: { ref: 'A1', columns: [{ colId: 0, hiddenButton: false, showButton: true, ...body }] },
})
/** Chart seed (Region totals of Sales: South 58440.44, East 36597.71, North 31451.53, West 29762.62): keeps South and East. */
export const CHART_VALUE_THRESHOLD = 34000
/** Engine seed, Date on rows: dates 2020-01-01..2021-12-31 (as Stage 2 filter-date-between). */
const DATE_FROM = 43831, DATE_TO = 44561

// ---- Expand and collapse (Task 25) -------------------------------------------------------------------------
/** Each form with subtotals at the top and at the bottom, and the Excel-saved file that has it. */
const FORMS = [
  ['compact', 'top', undefined],
  ['compact', 'bottom', 'engine/layout-subtotals-bottom-compact'],
  ['outline', 'top', 'engine/layout-subtotals-top-outline'],
  ['outline', 'bottom', 'engine/layout-subtotals-bottom-outline'],
  ['tabular', 'top', 'engine/layout-subtotals-top-tabular'],
  ['tabular', 'bottom', 'interaction/interaction-base-tabular-bottom'],
] as const
add('interaction-base-tabular-bottom', 25, {
  setup: 'tabular form, subtotals at the bottom (Stage 2 has no such file)', route: 'applescript',
  lines: t => [form(t, 'tabular'), subtotalsAt(t, 'bottom')], markers: [COMPACT_OFF, SUBTOTAL_BOTTOM], delta: [],
  measures: 'the tabular/bottom start state for the collapse cases',
})
for (const [f, where, base] of FORMS) {
  const tag = `${f}-${where}`
  add(`interaction-collapse-item-${tag}`, 25, {
    base, route: 'applescript', lines: t => [showDetail(t, 'Region', 'East', false)],
    markers: [collapsedItem(F.Region)], delta: [{ kind: 'collapse', field: 'Region', item: 'East', collapsed: true }],
    fallback: 'xml', xmlEdit: p => collapseItemXml(p, F.Region, 'East'),
    measures: `collapse one outer row item in ${f} form with subtotals at the ${where}: the collapsed line, its subtotal, the vacated rows (§10.1)`,
  })
  add(`interaction-collapse-field-${tag}`, 25, {
    base, route: 'applescript', lines: t => [fieldDetail(t, 'Region', false)],
    markers: [collapsedItem(F.Region)], absent: [expandedItem(F.Region)], delta: [{ kind: 'collapseField', field: 'Region', collapsed: true }],
    fallback: 'xml', xmlEdit: p => ({ ...p, pivot: setFieldDetail(p.pivot, F.Region, false) }), probe: 'field-show-detail',
    measures: `collapse the entire outer row field in ${f} form, subtotals at the ${where}`,
  })
}
for (const [f, where, base] of [FORMS[0], FORMS[3]] as const) {
  add(`interaction-collapse-all-${f}-${where}`, 25, {
    base, route: 'applescript', lines: t => [fieldDetail(t, 'Region', false), fieldDetail(t, 'Year', false)],
    markers: [collapsedItem(F.Region), collapsedItem(F.Year)], absent: [expandedItem(F.Region), expandedItem(F.Year)],
    delta: [{ kind: 'collapseAll', collapsed: true }],
    fallback: 'xml', xmlEdit: p => ({ ...p, pivot: setFieldDetail(setFieldDetail(p.pivot, F.Region, false), F.Year, false) }), probe: 'field-show-detail',
    measures: `collapse everything in ${f} form (subtotals at the ${where}): every item of each field that has a level below it on its axis ` +
      '(Region on rows; Year on columns, above Values); the innermost row field Product is not collapsible in the UI',
  })
}
add('interaction-expand-item', 25, {
  base: 'engine/layout-collapsed-item', route: 'applescript', lines: t => [showDetail(t, 'Region', 'East', true)],
  markers: [AXES_KEPT], absent: [/\bsd="0"/], delta: [{ kind: 'collapse', field: 'Region', item: 'East', collapsed: false }],
  measures: 'expand a collapsed item: back to the seed (sd removed, rows regrow)',
})
add('interaction-expand-field', 25, {
  base: 'interaction/interaction-collapse-field-compact-top', route: 'applescript', lines: t => [fieldDetail(t, 'Region', true)],
  markers: [AXES_KEPT], absent: [/\bsd="0"/], delta: [{ kind: 'collapseField', field: 'Region', collapsed: false }], probe: 'field-show-detail',
  fallback: 'xml', xmlEdit: p => ({ ...p, pivot: setFieldDetail(p.pivot, F.Region, true) }),
  measures: 'expand the entire field after collapsing it',
})
add('interaction-expand-all', 25, {
  base: 'interaction/interaction-collapse-all-compact-top', route: 'applescript', lines: t => [fieldDetail(t, 'Region', true), fieldDetail(t, 'Year', true)],
  markers: [AXES_KEPT], absent: [/\bsd="0"/], delta: [{ kind: 'collapseAll', collapsed: false }], probe: 'field-show-detail',
  fallback: 'xml', xmlEdit: p => ({ ...p, pivot: setFieldDetail(setFieldDetail(p.pivot, F.Region, true), F.Year, true) }),
  measures: 'expand everything after collapsing everything: back to the seed, and whether column widths grow back (§7.8)',
})

// ---- Item visibility (Task 25) -----------------------------------------------------------------------------
add('interaction-filter-all-hidden', 25, {
  route: 'applescript', lines: t => ['East', 'South', 'West', 'Central'].map(i => hideItem(t, 'Region', i)),
  markers: [hiddenItem(F.Region)], delta: [{ kind: 'items', field: 'Region', hidden: ['East', 'South', 'West', 'Central'] }],
  measures: 'every Region item hidden but North (§10.4): the one remaining group, subtotals and grand totals; whether column A shrinks (§7.8)',
})
add('interaction-items-inner-all-but-one', 25, {
  route: 'applescript', lines: t => ['Apples', 'Bananas', 'Elderberries'].map(i => hideItem(t, 'Product', i)),
  markers: [hiddenItem(F.Product)], delta: [{ kind: 'items', field: 'Product', hidden: ['Apples', 'Bananas', 'Elderberries'] }],
  measures: 'every inner item hidden but Cherries: Central (Elderberries only) drops; outer lines keep one child each',
})
add('interaction-items-show-again', 25, {
  base: 'engine/quirk-hidden-row-item', route: 'applescript', lines: t => [showItem(t, 'Region', 'West')],
  markers: [AXES_KEPT], absent: [hiddenItem(F.Region)], delta: [{ kind: 'items', field: 'Region', hidden: [] }],
  fallback: 'xml', xmlEdit: p => ({ ...p, pivot: unhideField(p.pivot, F.Region) }),
  measures: 'select all after one item was hidden: back to the seed',
})
add('interaction-clear-filters-hidden-items', 25, {
  base: 'engine/quirk-hidden-row-item', route: 'applescript', lines: t => [clearAll(t, 'Region')],
  markers: [AXES_KEPT], absent: [hiddenItem(F.Region)], delta: [{ kind: 'clearFilters', field: 'Region' }], probe: 'clear-all-filters',
  measures: 'Clear Filter From "Region" with only a manual (hidden-item) filter: does it show the hidden items again? (no xml fallback: the fallback would presuppose the answer)',
})

// ---- Page fields (Task 25) ---------------------------------------------------------------------------------
add('interaction-page-select-other', 25, {
  base: 'engine/quirk-page-single', route: 'applescript', lines: t => [currentPage(t, 'Product', 'Cherries')],
  markers: [/<pageField [^>]*item="\d+"/], delta: [{ kind: 'page', field: 'Product', selected: ['Cherries'] }],
  measures: 'a single page selection changed (Apples → Cherries): Central drops again',
})
add('interaction-page-all', 25, {
  base: 'engine/quirk-page-single', route: 'xml', xmlEdit: p => ({ ...p, pivot: pageAll(p.pivot, F.Product) }),
  markers: [/<pageFields count="1">/], absent: [/<pageField [^>]*item=/], delta: [{ kind: 'page', field: 'Product', selected: 'all' }],
  measures: 'page selection back to (All): the page cell caption "(All)" in zh-TW, Central returns',
})
add('interaction-page-multi', 25, {
  base: 'engine/quirk-page-single', route: 'xml', xmlEdit: p => pageMulti(p, F.Product, ['Bananas', 'Elderberries']),
  markers: [fieldAttr(F.Product, 'multipleItemSelectionAllowed', '1'), hiddenItem(F.Product)], absent: [/<pageField [^>]*item=/],
  delta: [{ kind: 'page', field: 'Product', selected: ['Apples', 'Cherries'] }],
  measures: 'single → multiple page selection (Apples and Cherries; `enable multiple page items` is refused, so the xml route)',
})
add('interaction-page-multi-to-all', 25, {
  base: 'engine/quirk-page-multiple', route: 'applescript', lines: t => [showItem(t, 'Product', 'Bananas'), showItem(t, 'Product', 'Elderberries')],
  markers: [/<pageFields count="1">/], absent: [hiddenItem(F.Product), /<pageField [^>]*item=/], delta: [{ kind: 'page', field: 'Product', selected: 'all' }],
  fallback: 'xml', xmlEdit: p => ({ ...p, pivot: unhideField(p.pivot, F.Product) }),
  measures: 'a multiple page selection toggled back to all: "(All)" again; whether multipleItemSelectionAllowed stays (applyDelta(page, all) must clear the hidden items)',
})

// ---- Filters (Task 25) -------------------------------------------------------------------------------------
for (const [id, base, name, fld, what] of [
  ['interaction-clear-label-filter', 'engine/filter-label-begins-with', 'Product', F.Product, 'a label filter'],
  ['interaction-clear-value-filter', 'engine/filter-value-greater-than', 'Region', F.Region, 'a value filter'],
  ['interaction-clear-top10-filter', 'engine/filter-top-count', 'Region', F.Region, 'a top-10 filter'],
] as const) {
  add(id, 25, {
    base, route: 'applescript', lines: t => [clearAll(t, name)], markers: [AXES_KEPT], absent: [ANY_FILTER], delta: [{ kind: 'clearFilters', field: name }],
    fallback: 'xml', xmlEdit: p => ({ ...p, pivot: dropFilters(p.pivot, fld) }), probe: 'clear-all-filters',
    measures: `clear ${what}: back to the seed (and whether measureFilter is dropped)`,
  })
}
add('interaction-filter-date-base', 25, {
  setup: 'Date alone on rows (Region and Product off the axis)', route: 'applescript', lines: t => onlyRows(t, 'Date'),
  markers: [rowField(F.Date), /<rowFields count="1">/], delta: [], measures: 'the start state of the date-filter cases (Date items in the cache)',
})
add('interaction-filter-date-between', 25, {
  base: 'interaction/interaction-filter-date-base', route: 'xml',
  xmlEdit: p => ({ ...p, pivot: addFilter(p.pivot, F.Date, 'dateBetween', custom(['greaterThanOrEqual', String(DATE_FROM)], ['lessThanOrEqual', String(DATE_TO)])) }),
  markers: [/<filter [^>]*type="dateBetween"/],
  delta: [{ kind: 'filter', field: 'Date', filter: filterLiteral('dateBetween', { customFilters: { and: true, filters: [{ operator: 'greaterThanOrEqual', val: String(DATE_FROM) }, { operator: 'lessThanOrEqual', val: String(DATE_TO) }] } }) }],
  measures: 'a date filter as a delta on Date-on-rows (Stage 2 filter-date-between changes the cache, so it is not same-cache): 2020-01-01..2021-12-31, blanks out?',
})
add('interaction-filter-date-clear', 25, {
  base: 'engine/filter-date-between', route: 'applescript', lines: t => [clearAll(t, 'Date')], markers: [rowField(F.Date)], absent: [ANY_FILTER],
  delta: [{ kind: 'clearFilters', field: 'Date' }], fallback: 'xml', xmlEdit: p => ({ ...p, pivot: dropFilters(p.pivot, F.Date) }), probe: 'clear-all-filters',
  measures: 'clear a date filter',
})

// ---- Sort (Task 25; Stage 2 covers label and value sorts descending) ----------------------------------------
add('interaction-sort-value-ascending-ties', 25, {
  route: 'applescript', lines: t => [autoSort(t, 'Region', 'ascending', UNITS)],
  markers: [fieldAttr(F.Region, 'sortType', 'ascending'), fieldHas(F.Region, String.raw`[^>]*(?<!\/)>(?:(?!<\/pivotField>)[\s\S])*?<autoSortScope>`)],
  delta: [{ kind: 'sort', field: 'Region', order: 'ascending', byData: 1 }],
  measures: 'sort ascending by a value column with ties (North = South on Sum of Units): tie order ascending',
})
add('interaction-sort-then-collapse', 25, {
  route: 'applescript', lines: t => [autoSort(t, 'Region', 'descending', 'Region'), showDetail(t, 'Region', 'West', false)],
  markers: [fieldAttr(F.Region, 'sortType', 'descending'), collapsedItem(F.Region)],
  delta: [{ kind: 'sort', field: 'Region', order: 'descending' }, { kind: 'collapse', field: 'Region', item: 'West', collapsed: true }],
  measures: 'two deltas in order: a label sort, then a collapse addressed by pivot-item index (indices are the start file\'s, not the sorted order)',
})

// ---- Captions (Task 25) ------------------------------------------------------------------------------------
const NO_CAPTIONS = { grandTotalCaption: null, rowHeaderCaption: null, colHeaderCaption: null }
const CAPTIONS_GONE = [/ grandTotalCaption=/, / rowHeaderCaption=/, / colHeaderCaption=/]
add('caption-defaults', 25, {
  setup: 'the seed\'s explicit grandTotalCaption, rowHeaderCaption and colHeaderCaption removed', route: 'xml',
  xmlEdit: p => ({ ...p, pivot: editRoot(p.pivot, NO_CAPTIONS) }), markers: [AXES_KEPT], absent: CAPTIONS_GONE, delta: [],
  measures: 'zh-TW defaults for Row Labels, Column Labels and Grand Total (the seed spells them out in English); Values stays explicit: dataCaption is required and Excel always writes it',
})
add('caption-defaults-single-data', 25, {
  setup: 'Sum of Units removed from Values; explicit captions removed', route: 'xml',
  xmlPrep: t => [`set pivot field orientation of ${dataField(t, UNITS)} to orient as hidden`], xmlEdit: p => ({ ...p, pivot: editRoot(p.pivot, NO_CAPTIONS) }),
  markers: [/<dataFields count="1">/], absent: CAPTIONS_GONE, delta: [],
  measures: 'one data field: its name in the top-left cell, both grand totals as the implicit Grand Total caption',
})
add('caption-defaults-outline-bottom', 25, {
  setup: 'outline form, subtotals at the bottom, explicit captions removed', route: 'xml',
  xmlPrep: t => [form(t, 'outline'), subtotalsAt(t, 'bottom')], xmlEdit: p => ({ ...p, pivot: editRoot(p.pivot, NO_CAPTIONS) }),
  markers: [COMPACT_OFF, SUBTOTAL_BOTTOM], absent: CAPTIONS_GONE, delta: [],
  measures: 'field-name headers over each column (Values field = dataCaption), bottom subtotal captions and Grand Total, all in zh-TW',
})
add('caption-defaults-collapsed-bottom', 25, {
  base: 'interaction/caption-defaults-outline-bottom', route: 'applescript', lines: t => [showDetail(t, 'Region', 'East', false)],
  markers: [collapsedItem(F.Region)], delta: [{ kind: 'collapse', field: 'Region', item: 'East', collapsed: true }],
  measures: 'a collapse on a pivot whose captions are all implicit: the caption rule must fill the new subtotal-less line (the oracle for resolveCaptions)',
})

// ---- Autofit (Task 29) -------------------------------------------------------------------------------------
const widths = (t: Target, w: number) => `set column width of range "A:K" of ${t.ws} to ${w}`
add('autofit-refresh-narrow', 29, {
  setup: 'columns A:K set to width 4 before the refresh', route: 'applescript', lines: t => [widths(t, 4)], markers: [AXES_KEPT], delta: [],
  measures: 'does a plain refresh autofit the pivot columns (useAutoFormatting="1")? widths before 4, after?',
})
add('autofit-collapse-narrow', 29, {
  setup: 'columns A:K set to width 4 before the collapse', route: 'applescript', lines: t => [widths(t, 4), showDetail(t, 'Region', 'East', false)],
  markers: [collapsedItem(F.Region)], delta: [{ kind: 'collapse', field: 'Region', item: 'East', collapsed: true }],
  measures: 'an update after narrow columns: does Excel grow them to fit, and with what padding over the text width?',
})
add('autofit-collapse-wide', 29, {
  setup: 'columns A:K set to width 40 before the collapse', route: 'applescript', lines: t => [widths(t, 40), showDetail(t, 'Region', 'East', false)],
  markers: [collapsedItem(F.Region)], delta: [{ kind: 'collapse', field: 'Region', item: 'East', collapsed: true }],
  measures: 'grow only? wide columns after an update: shrunk to fit or kept',
})
add('autofit-hide-longest', 29, {
  route: 'applescript', lines: t => [hideItem(t, 'Product', 'Elderberries'), hideItem(t, 'Region', 'Central')],
  markers: [hiddenItem(F.Product), hiddenItem(F.Region)],
  delta: [{ kind: 'items', field: 'Region', hidden: ['Central'] }, { kind: 'items', field: 'Product', hidden: ['Elderberries'] }],
  measures: 'the longest row label (Elderberries) leaves: does column A shrink with no manual width?',
})

// ---- Pivot chart (Task 30) ---------------------------------------------------------------------------------
const chart = (id: string, spec: Omit<Spec, 'seed' | 'chart'>) => add(id, 30, { seed: 'chart-seed', chart: true, ...spec })
chart('chart-baseline', { route: 'applescript', lines: () => [], markers: [AXES_KEPT, /<chartFormats count="\d+">/], delta: [], measures: 'the chart seed refreshed: series, categories and pivotFmts as saved' })
chart('chart-hide-item', {
  route: 'applescript', lines: t => [hideItem(t, 'Region', 'West')], markers: [hiddenItem(F.Region)], delta: [{ kind: 'items', field: 'Region', hidden: ['West'] }],
  measures: 'a hidden category group: category ref shrinks, series count unchanged',
})
chart('chart-collapse-item', {
  route: 'applescript', lines: t => [showDetail(t, 'Region', 'East', false)], markers: [collapsedItem(F.Region)],
  delta: [{ kind: 'collapse', field: 'Region', item: 'East', collapsed: true }], measures: 'a collapsed category: one category for East, multi-level categories',
})
chart('chart-value-filter', {
  route: 'xml', xmlEdit: p => ({ ...p, pivot: valueFilter(p.pivot, F.Region, 0, 'valueGreaterThan', CHART_VALUE_THRESHOLD) }), markers: [/<filter [^>]*type="valueGreaterThan"/],
  delta: [{ kind: 'filter', field: 'Region', filter: filterLiteral('valueGreaterThan', { customFilters: { and: false, filters: [{ operator: 'greaterThan', val: String(CHART_VALUE_THRESHOLD) }] } }, { iMeasureFld: 0 }) }],
  measures: 'a value filter that drops two category groups (§10.7)',
})
chart('chart-sort', {
  route: 'applescript', lines: t => [autoSort(t, 'Region', 'descending', 'Region')], markers: [fieldAttr(F.Region, 'sortType', 'descending')],
  fallback: 'xml', xmlEdit: p => ({ ...p, pivot: editPivotField(p.pivot, F.Region, { sortType: 'descending' }) }),
  delta: [{ kind: 'sort', field: 'Region', order: 'descending' }], measures: 'category order follows a label sort',
})
chart('chart-values-on-rows', {
  setup: 'Values moved to rows (a field move)', route: 'applescript',
  lines: t => [`set pivot field orientation of data pivot field of ${t.p} to orient as row field`], markers: [/ dataOnRows="1"/, rowField(-2)], delta: [],
  measures: 'series and categories with Values on rows (§6.6, CODEX-3): what the chart gate must map',
})
chart('chart-page-base', {
  setup: 'Product moved to a page field (a field move)', route: 'applescript', lines: t => [toPage(t, 'Product')], markers: [/<pageFields count="1">/], delta: [],
  measures: 'a pivot chart over a pivot with a page field: the start state of chart-page-select',
})
chart('chart-page-select', {
  base: 'interaction/chart-page-base', route: 'applescript', lines: t => [currentPage(t, 'Product', 'Apples')], markers: [/<pageField [^>]*item="\d+"/],
  delta: [{ kind: 'page', field: 'Product', selected: ['Apples'] }], measures: 'a page selection under a pivot chart: values change, refs do not',
})

// ---- GETPIVOTDATA (Task 31) --------------------------------------------------------------------------------
const gpd = (id: string, spec: Omit<Spec, 'seed' | 'sheetMarkers'>) =>
  add(id, 31, { seed: 'getpivotdata-seed', sheetMarkers: [{ sheet: GPD_SHEET, re: /<f>GETPIVOTDATA\(/ }], ...spec })
gpd('gpd-baseline', { route: 'applescript', lines: () => [], markers: [...SEED_PIVOT_MARKERS], delta: [], measures: 'the saved GETPIVOTDATA values on the unchanged pivot' })
gpd('gpd-hide-item', {
  route: 'applescript', lines: t => [hideItem(t, 'Region', 'West')], markers: [hiddenItem(F.Region)], delta: [{ kind: 'items', field: 'Region', hidden: ['West'] }],
  measures: 'GETPIVOTDATA naming a hidden item (§10.5: #REF! expected) and the grand total without it',
})
gpd('gpd-collapse-item', {
  route: 'applescript', lines: t => [showDetail(t, 'Region', 'East', false)], markers: [collapsedItem(F.Region)],
  delta: [{ kind: 'collapse', field: 'Region', item: 'East', collapsed: true }], measures: 'GETPIVOTDATA naming detail under a collapsed item (East/Bananas)',
})
gpd('gpd-filter-out', {
  route: 'xml', xmlEdit: p => ({ ...p, pivot: labelFilter(p.pivot, F.Product, 'captionBeginsWith', 'B') }), markers: [/<filter [^>]*type="captionBeginsWith"/],
  delta: [{ kind: 'filter', field: 'Product', filter: filterLiteral('captionBeginsWith', { customFilters: { and: false, filters: [{ operator: 'equal', val: 'B*' }] } }, { stringValue1: 'B' }) }],
  measures: 'GETPIVOTDATA naming an item a label filter removes (South/Cherries) and Central, left without records',
})

// ---- refreshOnLoad (Task 32) -------------------------------------------------------------------------------
const refreshed = [/<n v="1500\.25"\/>/]
add('refresh-on-load', 32, {
  seed: 'refresh-seed', route: 'applescript', lines: () => [], markers: [...SEED_PIVOT_MARKERS], cacheMarkers: [/refreshOnLoad="1"/], recordMarkers: refreshed, delta: [],
  measures: 'the refreshed output of a refreshOnLoad pivot whose records were older than Data (§7.7): East/Apples/2022 empties, North gains a record',
})
add('refresh-on-load-collapse', 32, {
  seed: 'refresh-seed', route: 'applescript', lines: t => [showDetail(t, 'Region', 'East', false)], markers: [collapsedItem(F.Region)], recordMarkers: refreshed,
  delta: [{ kind: 'collapse', field: 'Region', item: 'East', collapsed: true }],
  measures: 'the first interaction on a refreshOnLoad pivot: refresh from the source, then the change',
})
add('refresh-on-load-nodata', 32, {
  seed: 'refresh-seed', base: 'recipes/refresh-seed-nodata', route: 'applescript', lines: () => [], markers: [...SEED_PIVOT_MARKERS],
  cacheMarkers: [/saveData="0"/], delta: [], measures: 'refreshOnLoad with saveData="0": the cells Excel computes from the source alone',
})

// ---- Stage 4: Field moves, data fields & inline materialization (Task 35) ----------------------------------
const toData = (t: Target, name: string) => `set pivot field orientation of ${pivotField(t, name)} to orient as data field`
const hideData = (t: Target, name: string) => `set pivot field orientation of ${dataField(t, name)} to orient as hidden`
const hideDataAt = (t: Target, index: number) => `set pivot field orientation of data field ${index} of ${t.p} to orient as hidden`
const groupCell = (t: Target, name: string) => `(label range of pivot item 1 of ${pivotField(t, name)})`
const firstRowGroupCell = (t: Target) => `(label range of pivot item 1 of row field 1 of ${t.p})`
/** A number group's item 1 is the empty `<start` bucket (no label range), so Excel is pointed at the first data bucket. */
const numberGroupCell = (t: Target, name: string) => `(label range of pivot item 2 of ${pivotField(t, name)})`
const ALL_DATE_PERIODS: readonly PivotDateGroupPeriod[] = ['seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years']
const allPeriods = (...on: PivotDateGroupPeriod[]) => `{${ALL_DATE_PERIODS.map(p => on.includes(p)).join(', ')}}`

add('s4-move-inline-text-to-row', 35, {
  route: 'applescript', lines: t => [toRows(t, 'Channel')],
  markers: [rowField(F.Channel), /<rowFields count="3">/], cacheMarkers: [/<s v="Online"\/>/, /<n v="42"\/>/],
  delta: [{ kind: 'moveField', field: 'Channel', axis: 'row' }],
  measures: 'inline mixed text/number/blank field (Channel) moved to Rows: materialises sharedItems, <items> and default field flags',
})
add('s4-move-inline-date-to-col', 35, {
  route: 'applescript', lines: t => [toCols(t, 'Date')],
  markers: [colField(F.Date), /<colFields count="3">/], cacheMarkers: [/<d v="2019-01-15T00:00:00"\/>/],
  delta: [{ kind: 'moveField', field: 'Date', axis: 'col' }],
  measures: 'inline date field with blanks (Date) moved to Columns: materialises sharedItems and pins default insertion index relative to Values (-2)',
})
add('s4-move-inline-to-page', 35, {
  route: 'applescript', lines: t => [toPage(t, 'Month')],
  markers: [/<pageFields count="1"><pageField fld="9"/], cacheMarkers: [/<s v="Jan"\/>/],
  delta: [{ kind: 'moveField', field: 'Month', axis: 'page' }],
  measures: 'inline text field (Month) moved to Filters/Page: materialises sharedItems/<items> and defaults selection to (All)',
})
add('s4-move-reorder-rows', 35, {
  route: 'applescript', lines: t => [`set position of ${pivotField(t, 'Product')} to 1`], probe: 'field-position',
  markers: [/<rowFields count="2"><field x="1"\/><field x="0"\/><\/rowFields>/],
  delta: [{ kind: 'moveField', field: 'Product', axis: 'row', index: 0 }],
  measures: 'reorder within Rows: Product moved above Region ([1, 0])',
})
add('s4-move-values-to-rows-first', 35, {
  route: 'applescript', probe: 'field-position',
  lines: t => [`set pivot field orientation of data pivot field of ${t.p} to orient as row field`, `set position of data pivot field of ${t.p} to 1`],
  markers: [/ dataOnRows="1"/, /<rowFields count="3"><field x="-2"\/><field x="0"\/><field x="1"\/><\/rowFields>/],
  delta: [{ kind: 'moveField', field: 'Values', axis: 'row', index: 0 }],
  measures: 'Values (-2) moved to Rows at index 0 (dataOnRows="1" dataPosition="0")',
})
add('s4-filtered-row-base', 35, {
  route: 'xml',
  xmlPrep: t => [hideItem(t, 'Region', 'West'), showDetail(t, 'Region', 'East', false), autoSort(t, 'Region', 'descending', 'Region')],
  xmlEdit: p => ({ ...p, pivot: labelFilter(p.pivot, F.Region, 'captionGreaterThan', 'D') }),
  markers: [hiddenItem(F.Region), collapsedItem(F.Region), fieldAttr(F.Region, 'sortType', 'descending'), /<filter [^>]*type="captionGreaterThan"/],
  delta: [
    { kind: 'items', field: 'Region', hidden: ['West'] },
    { kind: 'collapse', field: 'Region', item: 'East', collapsed: true },
    { kind: 'sort', field: 'Region', order: 'descending' },
    { kind: 'filter', field: 'Region', filter: filterLiteral('captionGreaterThan', { customFilters: { and: false, filters: [{ operator: 'greaterThan', val: 'D' }] } }, { stringValue1: 'D' }) },
  ],
  measures: 'Region on Rows carrying hidden West, collapsed East, descending sort and label filter >"D" (excludes Central): start state for cross-axis retention tests',
})
add('s4-move-row-to-col', 35, {
  base: 'interaction/s4-filtered-row-base', route: 'applescript', lines: t => [toCols(t, 'Region')],
  markers: [colField(F.Region)],
  delta: [{ kind: 'moveField', field: 'Region', axis: 'col' }],
  measures: 'Row → Column move of a field with hidden item, collapsed item, label filter and descending sort: pins cross-axis retention (S4-E)',
})
add('s4-move-row-to-col-and-back', 35, {
  base: 'interaction/s4-filtered-row-base', route: 'applescript', probe: 'field-position',
  lines: t => [toCols(t, 'Region'), toRows(t, 'Region'), `set position of ${pivotField(t, 'Region')} to 1`],
  markers: [/<rowFields count="2"><field x="0"\/><field x="1"\/><\/rowFields>/],
  delta: [{ kind: 'moveField', field: 'Region', axis: 'col' }, { kind: 'moveField', field: 'Region', axis: 'row', index: 0 }],
  measures: 'Row → Column → Row round-trip of a filtered/hidden/collapsed/sorted field: pins S4-E preservation vs reset',
})
add('s4-move-filtered-field-to-page', 35, {
  base: 'interaction/s4-filtered-row-base', route: 'applescript', lines: t => [toPage(t, 'Region')],
  markers: [/<pageFields count="1"><pageField fld="0"/],
  delta: [{ kind: 'moveField', field: 'Region', axis: 'page' }],
  measures: 'Row → Page move of a field carrying hidden item, collapsed item, label filter and descending sort',
})
add('s4-move-remove-filtered-field', 35, {
  base: 'interaction/s4-filtered-row-base', route: 'applescript', lines: t => [hide(t, 'Region')],
  markers: [/<rowFields count="1"><field x="1"\/><\/rowFields>/], absent: [rowField(F.Region), colField(F.Region)],
  delta: [{ kind: 'moveField', field: 'Region', axis: 'none' }],
  measures: 'Row → none removal of a field carrying hidden item, collapsed item, label filter and descending sort: pins what Excel clears on removal',
})
add('s4-data-remove-to-single', 35, {
  route: 'applescript', lines: t => [hideData(t, UNITS)],
  markers: [/<dataFields count="1">/, /<colFields count="1"><field x="2"\/><\/colFields>/], absent: [colField(-2), rowField(-2)],
  delta: [{ kind: 'removeDataField', dataIndex: 1 }],
  measures: '2 → 1 data fields removes Values (-2) from Columns',
})
add('s4-data-add-second-default-axis', 35, {
  base: 'interaction/s4-data-remove-to-single', route: 'applescript', lines: t => [toData(t, 'Units')],
  markers: [/<dataFields count="2">/, colField(-2)],
  delta: [{ kind: 'addDataField', field: 'Units' }],
  measures: '1 → 2 data fields from a single-data pivot: Values (-2) axis & default position and zh-TW default data-field caption',
})
add('s4-data-add-second-after-values-on-rows', 35, {
  base: 'engine/layout-values-on-rows-two-data', route: 'applescript',
  lines: t => [hideData(t, UNITS), toData(t, 'Units')],
  markers: [/<dataFields count="2">/],
  delta: [{ kind: 'removeDataField', dataIndex: 1 }, { kind: 'addDataField', field: 'Units' }],
  measures: '2 → 1 → 2 data fields when Values was previously on Rows (dataOnRows="1"): pins whether 1 → 2 restores Values to Rows or defaults to Columns',
})
add('s4-data-add-numeric-vs-text', 35, {
  route: 'applescript', lines: t => [toData(t, 'Net Amount'), toData(t, 'Channel')],
  markers: [/<dataFields count="4">/, /<dataField [^>]*fld="7"/, /<dataField [^>]*fld="6"[^>]*subtotal="count"|<dataField [^>]*subtotal="count"[^>]*fld="6"/],
  delta: [{ kind: 'addDataField', field: 'Net Amount' }, { kind: 'addDataField', field: 'Channel' }],
  measures: 'default summary function rule: Net Amount (numbers + blanks) → sum ("加總 - Net Amount"); Channel (mixed text/number/blank) → count ("計數 - Channel")',
})
add('s4-data-add-duplicate-field', 35, {
  route: 'applescript', probe: 'add-duplicate-data-field',
  lines: t => [toData(t, 'Sales'), toData(t, 'Sales')],
  markers: [/<dataFields count="4">/, /<dataField [^>]*name="[^"]*Sales2"/],
  delta: [{ kind: 'addDataField', field: 'Sales' }, { kind: 'addDataField', field: 'Sales' }],
  measures: 'adding Sales when "Sum of Sales" exists ("加總 - Sales") and again when "加總 - Sales" exists ("加總 - Sales2"): pins duplicate suffix rule',
})
add('s4-data-reorder', 35, {
  route: 'applescript', probe: 'data-field-position',
  lines: t => [`set position of ${dataField(t, SALES)} to 2`],
  markers: [/<dataFields count="2"><dataField [^>]*fld="4"[^>]*\/><dataField [^>]*fld="3"/],
  delta: [{ kind: 'moveDataField', dataIndex: 0, toIndex: 1 }],
  measures: 'reorder data fields (Sum of Sales ↔ Sum of Units)',
})
add('s4-data-reorder-remap-measure', 35, {
  base: 'engine/filter-top-count-ties', route: 'applescript', probe: 'data-field-position',
  lines: t => [autoSort(t, 'Product', 'descending', UNITS), `set position of ${dataField(t, UNITS)} to 1`],
  markers: [/<dataFields count="2"><dataField [^>]*fld="4"[^>]*\/><dataField [^>]*fld="3"/, /<filter [^>]*iMeasureFld="0"/],
  delta: [{ kind: 'sort', field: 'Product', order: 'descending', byData: 1 }, { kind: 'moveDataField', dataIndex: 1, toIndex: 0 }],
  measures: 'moveDataField(1 → 0) with an active top-10 filter and value sort on Sum of Units: pins iMeasureFld and autoSortScope remapping',
})
add('s4-data-remove-all', 35, {
  route: 'applescript', lines: t => [hideData(t, UNITS), hideData(t, SALES)],
  markers: [rowField(F.Region), colField(F.Year)], absent: [/<dataFields /],
  delta: [{ kind: 'removeDataField', dataIndex: 1 }, { kind: 'removeDataField', dataIndex: 0 }],
  measures: 'remove both data fields (0 data fields, Row and Column fields kept): pins 0-data layout cells and location',
})
add('s4-empty-pivot-all-removed', 35, {
  route: 'applescript', lines: t => [hideData(t, UNITS), hideData(t, SALES), hide(t, 'Year'), hide(t, 'Product'), hide(t, 'Region')],
  markers: [/<location /], absent: [/<rowFields /, /<colFields /, /<pageFields /, /<dataFields /],
  delta: [
    { kind: 'removeDataField', dataIndex: 1 },
    { kind: 'removeDataField', dataIndex: 0 },
    { kind: 'moveField', field: 'Year', axis: 'none' },
    { kind: 'moveField', field: 'Product', axis: 'none' },
    { kind: 'moveField', field: 'Region', axis: 'none' },
  ],
  measures: 'all row, column, page and data fields removed (0 fields on all axes): pins empty-pivot cells and location (place.ts:35)',
})
add('s4-move-units-only-rows', 35, {
  route: 'applescript', lines: t => onlyRows(t, 'Units'),
  markers: [/<rowFields count="1"><field x="4"\/><\/rowFields>/, /<dataFields count="2">/],
  delta: [
    { kind: 'moveField', field: 'Product', axis: 'none' },
    { kind: 'moveField', field: 'Region', axis: 'none' },
    { kind: 'moveField', field: 'Units', axis: 'row' },
  ],
  measures: 'inline numeric field Units moved onto Rows while Sum of Units remains in Values (also start base for Units number grouping)',
})
add('s4-move-sales-only-rows', 35, {
  route: 'applescript', lines: t => onlyRows(t, 'Sales'),
  markers: [/<rowFields count="1"><field x="3"\/><\/rowFields>/, /<dataFields count="2">/],
  delta: [
    { kind: 'moveField', field: 'Product', axis: 'none' },
    { kind: 'moveField', field: 'Region', axis: 'none' },
    { kind: 'moveField', field: 'Sales', axis: 'row' },
  ],
  measures: 'inline numeric field Sales (with 3 blanks and 3 negatives) moved onto Rows while Sum of Sales remains in Values (base for Sales number grouping)',
})
add('s4-chart-values-to-rows', 35, {
  seed: 'chart-seed', chart: true, route: 'applescript',
  lines: t => [`set pivot field orientation of data pivot field of ${t.p} to orient as row field`],
  markers: [/ dataOnRows="1"/, rowField(-2)],
  delta: [{ kind: 'moveField', field: 'Values', axis: 'row' }],
  measures: 'chart-linked pivot after moving Values (-2) to Rows as a Stage 4 delta',
})
add('s4-chart-remove-second-data', 35, {
  seed: 'chart-seed', chart: true, route: 'applescript',
  lines: t => [hideDataAt(t, 2)],
  markers: [/<dataFields count="1">/], absent: [colField(-2)],
  delta: [{ kind: 'removeDataField', dataIndex: 1 }],
  measures: 'chart-linked pivot reduced from 2 data fields to 1 (Values -2 removed from Columns, single series)',
})

// ---- Stage 4: Value field settings & GETPIVOTDATA follow (Task 36) ------------------------------------------
add('s4-values-no-totals-base', 36, {
  setup: 'Product removed from rows and rowGrandTotals="0" colGrandTotals="0" so hasTotals is false (S2-AA)',
  route: 'applescript',
  lines: t => [hide(t, 'Product'), `set column grand of ${t.p} to false`, `set row grand of ${t.p} to false`],
  markers: [/<rowFields count="1"><field x="0"\/><\/rowFields>/, / rowGrandTotals="0"/, / colGrandTotals="0"/],
  delta: [],
  measures: 'single-row-field total-less base (hasTotals === false) for S2-AA stdDev/stdDevp/var/varp valueSettings oracles',
})
for (const [slug, fnName, asFn] of [
  ['stdev', 'stdDev', 'standard deviation'],
  ['stdevp', 'stdDevp', 'standard deviation p'],
  ['var', 'var', 'var'],
  ['varp', 'varp', 'var p'],
] as const) {
  add(`s4-values-fn-${slug}-no-totals`, 36, {
    base: 'interaction/s4-values-no-totals-base', route: 'applescript',
    lines: t => [`set function of ${dataField(t, SALES)} to do ${asFn}`],
    markers: [new RegExp(`<dataField [^>]*subtotal="${fnName}"`)],
    delta: [{ kind: 'valueSettings', dataIndex: 0, fn: fnName }],
    measures: `${fnName} on a total-less pivot (hasTotals === false, allowed under S2-AA) with zh-TW default data-field caption`,
  })
}
add('s4-values-custom-name', 36, {
  route: 'applescript',
  lines: t => [`set function of ${dataField(t, SALES)} to do average`, `set name of data field 1 of ${t.p} to "Revenue"`],
  markers: [/<dataField [^>]*name="Revenue"/, /<dataField [^>]*subtotal="average"/],
  delta: [{ kind: 'valueSettings', dataIndex: 0, fn: 'average', name: 'Revenue' }],
  measures: 'valueSettings changing summary function to average with explicit custom name "Revenue"',
})
add('s4-values-numfmt-currency-percent', 36, {
  route: 'applescript', probe: 'data-field-number-format',
  lines: t => [
    `set number format of ${dataField(t, SALES)} to "#,##0.00"`,
    `set calculation of ${dataField(t, UNITS)} to pivot field calculation percent of total`,
    `set number format of ${dataField(t, UNITS)} to "0.00%"`,
  ],
  markers: [/<dataField [^>]*numFmtId="4"/, /<dataField [^>]*showDataAs="percentOfTotal"/, /<dataField [^>]*numFmtId="10"/],
  delta: [
    { kind: 'valueSettings', dataIndex: 0, numFmtId: 4 },
    { kind: 'valueSettings', dataIndex: 1, showAs: 'percentOfTotal', numFmtId: 10 },
  ],
  measures: 'built-in number formats numFmtId=4 (#,##0.00) on Sum of Sales and numFmtId=10 (0.00%) with percentOfTotal on Sum of Units',
})
add('s4-values-showas-previous-next', 36, {
  route: 'applescript',
  lines: t => [
    `set calculation of ${dataField(t, SALES)} to pivot field calculation difference from`,
    `set base field of ${dataField(t, SALES)} to "Year"`,
    `set base item of ${dataField(t, SALES)} to "(previous)"`,
    `set calculation of ${dataField(t, UNITS)} to pivot field calculation percent difference from`,
    `set base field of ${dataField(t, UNITS)} to "Year"`,
    `set base item of ${dataField(t, UNITS)} to "(next)"`,
  ],
  markers: [/<dataField [^>]*baseItem="1048828"/, /<dataField [^>]*baseItem="1048829"/],
  delta: [
    { kind: 'valueSettings', dataIndex: 0, showAs: 'difference', baseField: 'Year', baseItem: '(previous)' },
    { kind: 'valueSettings', dataIndex: 1, showAs: 'percentDiff', baseField: 'Year', baseItem: '(next)' },
  ],
  measures: 'difference from (previous) (1048828) and % difference from (next) (1048829) applied together across the two data fields',
})
add('s4-gpd-rename-and-move', 36, {
  seed: 'getpivotdata-seed', route: 'applescript',
  lines: t => [toRows(t, 'Year'), `set function of ${dataField(t, UNITS)} to do average`, `set name of data field 2 of ${t.p} to "Avg Units"`],
  markers: [rowField(F.Year), /<dataField [^>]*name="Avg Units"/],
  sheetMarkers: [{ sheet: GPD_SHEET, re: /<f>GETPIVOTDATA\(/ }],
  delta: [
    { kind: 'moveField', field: 'Year', axis: 'row' },
    { kind: 'valueSettings', dataIndex: 1, fn: 'average', name: 'Avg Units' },
  ],
  measures: 'GETPIVOTDATA formulas after Year moves from Columns to Rows and "Sum of Units" is renamed to "Avg Units" (average)',
})

// ---- Stage 4: Grouping, regrouping & ungrouping (Task 37) ---------------------------------------------------
add('s4-group-date-single-months', 37, {
  base: 'interaction/interaction-filter-date-base', route: 'applescript',
  lines: t => [`group ${groupCell(t, 'Date')} periods ${allPeriods('months')}`],
  markers: [rowField(12)], cacheMarkers: [/<rangePr [^>]*groupBy="months"/, /<cacheField [^>]*databaseField="0"/],
  delta: [{ kind: 'groupDate', field: 'Date', periods: ['months'] }],
  measures: 'single-period date grouping (months) on Date-on-rows: creates 1 virtual cacheField ("月 (Date)", x="12") with <fieldGroup base="5"> and blank bucket',
})
add('s4-group-date-days-single', 37, {
  base: 'interaction/interaction-filter-date-base', route: 'applescript', probe: 'group-days-interval',
  lines: t => [`group ${groupCell(t, 'Date')} periods ${allPeriods('days')}`],
  markers: [rowField(12)], cacheMarkers: [/<rangePr [^>]*groupBy="days"/, /<cacheField [^>]*databaseField="0"/],
  delta: [{ kind: 'groupDate', field: 'Date', periods: ['days'] }],
  measures: '1-day date grouping (groupBy="days", groupInterval=1, virtual field "天 (Date)"): leap-year 366-day calendar table and zh-TW day captions',
})
add('s4-group-date-days-7', 37, {
  base: 'interaction/interaction-filter-date-base', route: 'applescript', probe: 'group-days-interval',
  lines: t => [`group ${groupCell(t, 'Date')} by 7 periods ${allPeriods('days')}`],
  markers: [rowField(12)],
  cacheMarkers: [/<rangePr [^>]*groupBy="days"[^>]*groupInterval="7"|<rangePr [^>]*groupInterval="7"[^>]*groupBy="days"/],
  delta: [{ kind: 'groupDate', field: 'Date', periods: ['days'], daysInterval: 7 }],
  measures: '7-day interval grouping (groupBy="days" groupInterval="7"): exact 7-day range captions, end clamping and blank bucket',
})
add('s4-group-date-time-hms', 37, {
  base: 'engine/quirk-dates-1900-rows', route: 'applescript', probe: 'group-time-hms',
  lines: t => [`group ${groupCell(t, 'Stamp')} periods ${allPeriods('seconds', 'minutes', 'hours')}`],
  markers: [/<rowFields count="3">/],
  cacheMarkers: [/<rangePr [^>]*groupBy="hours"/, /<rangePr [^>]*groupBy="minutes"/, /<rangePr [^>]*groupBy="seconds"/],
  delta: [{ kind: 'groupDate', field: 'Stamp', periods: ['seconds', 'minutes', 'hours'] }],
  measures: 'Stamp grouped by hours/minutes/seconds: pins serials 0.5, 1, 59 (1900-02-28), 60 (fictitious 1900-02-29), 61, 45000.25 and blanks under S2-H',
})
add('s4-group-date-multi-yqm', 37, {
  base: 'interaction/interaction-filter-date-base', route: 'applescript',
  lines: t => [`group ${groupCell(t, 'Date')} periods ${allPeriods('months', 'quarters', 'years')}`],
  markers: [/<rowFields count="3">/],
  cacheMarkers: [/<rangePr [^>]*groupBy="years"/, /<rangePr [^>]*groupBy="quarters"/, /<rangePr [^>]*groupBy="months"/],
  delta: [{ kind: 'groupDate', field: 'Date', periods: ['months', 'quarters', 'years'] }],
  measures: 'Date grouped from ungrouped items into Years › Quarters › Months: pins virtual cacheField/pivotField creation, zh-TW names and par/base links',
})
add('s4-group-date-manual-bounds', 37, {
  base: 'interaction/interaction-filter-date-base', route: 'applescript',
  lines: t => [`group ${groupCell(t, 'Date')} start 43831 end 45291 periods ${allPeriods('quarters', 'years')}`],
  markers: [/<rowFields count="2">/],
  cacheMarkers: [/<rangePr [^>]*autoStart="0"/, /<rangePr [^>]*autoEnd="0"/, /<rangePr [^>]*groupBy="years"/, /<rangePr [^>]*groupBy="quarters"/],
  delta: [{ kind: 'groupDate', field: 'Date', periods: ['quarters', 'years'], startDate: 43831, endDate: 45291 }],
  measures: 'multi-level date grouping (Years › Quarters) with manual bounds 2020-01-01..2023-12-31 inside 2019..2024: pins <start, >end and blank buckets',
})
add('s4-regroup-date-yqm-to-ym', 37, {
  base: 'interaction/s4-group-date-multi-yqm', route: 'applescript',
  lines: t => [`group ${firstRowGroupCell(t)} periods ${allPeriods('months', 'years')}`],
  markers: [/<rowFields count="2">/],
  cacheMarkers: [/<rangePr [^>]*groupBy="years"/, /<rangePr [^>]*groupBy="months"/],
  cacheAbsent: [/<rangePr [^>]*groupBy="quarters"/],
  delta: [{ kind: 'groupDate', field: 'Date', periods: ['months', 'years'] }],
  measures: 'regrouping Years › Quarters › Months down to Years › Months: pins removal of the Quarters virtual cacheField and pivotField',
})
add('s4-ungroup-date-multi', 37, {
  base: 'interaction/s4-group-date-multi-yqm', route: 'applescript', probe: 'ungroup',
  lines: t => [`ungroup ${firstRowGroupCell(t)}`],
  markers: [/<rowFields count="1"><field x="5"\/><\/rowFields>/],
  cacheAbsent: [/<fieldGroup /],
  delta: [{ kind: 'ungroup', field: 'Date' }],
  measures: 'ungrouping a 3-level date hierarchy back to ungrouped Date items: pins removal of virtual fields and restoration of Date on Rows',
})
add('s4-ungroup-date-single', 37, {
  base: 'interaction/s4-group-date-single-months', route: 'applescript', probe: 'ungroup',
  lines: t => [`ungroup ${firstRowGroupCell(t)}`],
  markers: [/<rowFields count="1"><field x="5"\/><\/rowFields>/],
  cacheAbsent: [/<fieldGroup /],
  delta: [{ kind: 'ungroup', field: 'Date' }],
  measures: 'ungrouping a 1-level date group (months, virtual field "月 (Date)") back to ungrouped Date items',
})
add('s4-group-number-units-auto', 37, {
  base: 'interaction/s4-move-units-only-rows', route: 'applescript', probe: 'group-number-auto',
  lines: t => [`group ${groupCell(t, 'Units')} by 25`],
  markers: [rowField(F.Units)], cacheMarkers: [/<rangePr [^>]*groupInterval="25"/],
  delta: [{ kind: 'groupNumber', field: 'Units', by: 25 }],
  measures: 'Units grouped by 25 with automatic min/max bounds (-5..233): pins autoStart/autoEnd startNum/endNum and bucket labels',
})
add('s4-group-number-sales-blanks-manual', 37, {
  base: 'interaction/s4-move-sales-only-rows', route: 'applescript',
  lines: t => [`group ${groupCell(t, 'Sales')} start 0 end 5000 by 1000`],
  markers: [rowField(F.Sales)], cacheMarkers: [/<rangePr [^>]*groupInterval="1000"/],
  delta: [{ kind: 'groupNumber', field: 'Sales', start: 0, end: 5000, by: 1000 }],
  measures: 'Sales (3 blanks, 3 negatives, values >5000) grouped with manual 0..5000 by 1000: pins <0, >5000 and blank bucket placement (groups.ts:175)',
})
add('s4-regroup-number-units', 37, {
  base: 'interaction/s4-group-number-units-auto', route: 'applescript',
  lines: t => [`group ${numberGroupCell(t, 'Units')} by 50`],
  markers: [rowField(F.Units)], cacheMarkers: [/<rangePr [^>]*groupInterval="50"/],
  delta: [{ kind: 'groupNumber', field: 'Units', by: 50 }],
  measures: 'regrouping Units from interval 25 to interval 50',
})
add('s4-ungroup-number-units', 37, {
  base: 'interaction/s4-group-number-units-auto', route: 'applescript', probe: 'ungroup',
  lines: t => [`ungroup ${numberGroupCell(t, 'Units')}`],
  markers: [rowField(F.Units)], cacheAbsent: [/<fieldGroup /],
  delta: [{ kind: 'ungroup', field: 'Units' }],
  measures: 'ungrouping numeric range group on Units back to distinct numeric items',
})
add('s4-ungroup-manual-discrete', 37, {
  base: 'engine/group-manual', route: 'applescript', probe: 'ungroup',
  lines: t => [`ungroup ${firstRowGroupCell(t)}`],
  markers: [/<rowFields count="1"><field x="1"\/><\/rowFields>/], cacheAbsent: [/<discretePr /],
  delta: [{ kind: 'ungroup', field: 'Product2' }],
  measures: 'ungrouping saved manual discrete group (Product2 over Product) back to Product alone on Rows',
})

export const INTERACTION_CASES: readonly InteractionCase[] = CASES

/** Probes for the incantations Stage 2 did not confirm (`build-engine.ts --interaction --probes`). */
export const INTERACTION_PROBES: Record<string, string> = {
  'field-show-detail': '`set show detail of <pivot field> to false|true` (Excel.sdef pivot field property; Stage 2 used the item property only)',
  'clear-all-filters': '`clear all filters (<pivot field>)` (Excel.sdef command, direct parameter a pivot table or field)',
  'make-worksheet': '`set s to make new worksheet at end of <workbook>` then `set name of s to …` (GETPIVOTDATA seed)',
  'refresh-on-file-open': '`set refresh on file open of (pivot cache of <pivot>) to true` (Excel.sdef pivot cache property)',
  'save-data': '`set save data of <pivot> to false` (Excel.sdef pivot table property)',
  'field-position': '`set position of <pivot field | data pivot field> to 1` (1-based axis reorder)',
  'data-field-position': '`set position of <data field> to 1|2` (1-based Values reorder)',
  'add-duplicate-data-field': '`set pivot field orientation of <pivot field> to orient as data field` when the source column is already in Values',
  'data-field-number-format': '`set number format of <data field> to "#,##0.00"` (built-in numFmtId on <dataField>)',
  'group-days-interval': '`group <cell> [by 7] periods {false, false, false, true, false, false, false}` (1-day and 7-day intervals)',
  'group-time-hms': '`group <cell> periods {true, true, true, false, false, false, false}` (hours, minutes, seconds)',
  'group-number-auto': '`group <cell> by 25` (numeric grouping with automatic min/max bounds)',
  'ungroup': '`ungroup <cell>` on single-level date, multi-level date, number and manual discrete groups',
}

/** The file a case starts from, relative to tests/fixtures/xlsx-pivots (without .xlsx). */
export const startOf = (c: Pick<InteractionCase, 'seed' | 'base'>) => c.base ?? SEED_FILES[c.seed]
