import { attrs, elementChildren, getChildren, type XmlNode } from '../../core/xml'

/** Declared file filter/sort state as diagnostics only (spec §7.5, §13.3 items
 * 8-9). The small parser builds this from part XML; the large adapter produces
 * the identical structure later from its own traversal, reusing the kinds, the
 * classifiers and the ref parser below. Nothing here ever becomes a viewer
 * rule: supported file filters persist as hidden rows, Clear restores the
 * load snapshot, and stored sortState never reorders source rows. */
export interface DeclaredRange {
  minRow: number
  minCol: number
  maxRow: number
  maxCol: number
}
export type DeclaredFilterRule =
  | { kind: 'values'; valueCount: number; hasBlank: boolean; hasDateGroups: boolean }
  | { kind: 'custom'; operator: 'and' | 'or'; conditions: { operator: string; value: string }[]; totalCount: number }
  | { kind: 'top10'; top?: boolean; percent?: boolean; val?: string; filterVal?: string }
  | { kind: 'dynamic'; filterType: string; value?: string; maxValue?: string }
  | { kind: 'color'; dxfId?: string; cellColor?: string }
  | { kind: 'unknown'; element: string }
export interface DeclaredFilterColumn {
  index: number
  colId?: number
  hiddenButton?: boolean
  showButton?: boolean
  rule?: DeclaredFilterRule
  malformed?: string
}
export interface DeclaredAutoFilter {
  refVerbatim?: string
  range?: DeclaredRange
  columns: DeclaredFilterColumn[]
  malformed?: string
}
export type DeclaredSortBy = 'value' | 'cellColor' | 'fontColor' | 'icon' | 'unknown'
export interface DeclaredSortCondition {
  refVerbatim?: string
  range?: DeclaredRange
  sortBy: DeclaredSortBy
  sortByVerbatim?: string
  descending?: boolean
  dxfId?: string
  customListCount?: number
  malformed?: string
}
export interface DeclaredSortState {
  refVerbatim?: string
  range?: DeclaredRange
  caseSensitive?: boolean
  conditions: DeclaredSortCondition[]
  malformed?: string
}
export interface DeclaredTableRef {
  displayName?: string
  partPath: string
}
export interface DeclaredTableFilters {
  table: DeclaredTableRef
  autoFilter: DeclaredAutoFilter
  sortState?: DeclaredSortState
}
export interface DeclaredSheetFilters {
  sheetId: string
  name: string
  workbookIndex: number
  partPath: string
  worksheetAutoFilter?: DeclaredAutoFilter
  worksheetSortState?: DeclaredSortState
  tableAutoFilters: DeclaredTableFilters[]
}
export type DeclaredFilterRuleTag = 'filters' | 'top10' | 'customFilters' | 'dynamicFilter' | 'colorFilter' | 'unknown'

const MAX_ROW = 1048576
const MAX_COL = 16384
const MAX_RECORDED_CONDITIONS = 8

/** Element name of a filterColumn rule child, for either adapter. */
export function classifyFilterRuleTag(tag: string): DeclaredFilterRuleTag {
  switch (tag) {
    case 'filters': case 'top10': case 'customFilters': case 'dynamicFilter': case 'colorFilter': return tag
    default: return 'unknown'
  }
}
/** sortBy spelling; absent means value per the schema default. */
export function sortByOf(value: string | undefined): DeclaredSortBy {
  switch (value) {
    case undefined: case 'value': return 'value'
    case 'cellColor': case 'fontColor': case 'icon': return value
    default: return 'unknown'
  }
}
/** Strict A1 ref: uppercase letters, row within grid, normalized corners.
 * Anything else is not a range (the verbatim spelling is kept by callers). */
export function parseDeclaredRef(verbatim: string): DeclaredRange | undefined {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(verbatim)
  if (!m) return undefined
  const col = (letters: string): number => {
    let n = 0
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64)
    return n - 1
  }
  const c0 = col(m[1]), r0 = parseInt(m[2], 10) - 1
  const c1 = m[3] !== undefined ? col(m[3]) : c0, r1 = m[4] !== undefined ? parseInt(m[4], 10) - 1 : r0
  if (![c0, r0, c1, r1].every(Number.isSafeInteger)) return undefined
  if (r0 < 0 || c0 < 0 || r1 >= MAX_ROW || c1 >= MAX_COL || r0 > r1 || c0 > c1) return undefined
  return { minRow: r0, minCol: c0, maxRow: r1, maxCol: c1 }
}
function textAttr(value: unknown): string | undefined {
  return typeof value === 'string' && value !== '' ? value : undefined
}
function boolAttr(value: unknown): boolean | undefined {
  if (value === '1' || value === 'true') return true
  if (value === '0' || value === 'false') return false
  return undefined
}
function intAttr(value: unknown): number | undefined {
  if (typeof value !== 'string' || value === '') return undefined
  const n = parseInt(value, 10)
  return Number.isSafeInteger(n) && n >= 0 ? n : undefined
}
function ruleOf(tag: string, node: XmlNode): DeclaredFilterRule {
  const a = attrs(node)
  switch (classifyFilterRuleTag(tag)) {
    case 'filters': {
      const valueCount = getChildren(node, 'filter').length + getChildren(node, 'dateGroupItem').length
      const hasDateGroups = getChildren(node, 'dateGroupItem').length > 0
      return { kind: 'values', valueCount, hasBlank: a.blank === '1' || a.blank === 'true', hasDateGroups }
    }
    case 'top10':
      return {
        kind: 'top10',
        ...(boolAttr(a.top) !== undefined ? { top: boolAttr(a.top) } : {}),
        ...(boolAttr(a.percent) !== undefined ? { percent: boolAttr(a.percent) } : {}),
        ...(textAttr(a.val) !== undefined ? { val: textAttr(a.val) } : {}),
        ...(textAttr(a.filterVal) !== undefined ? { filterVal: textAttr(a.filterVal) } : {}),
      }
    case 'customFilters': {
      const customs = getChildren(node, 'customFilter')
      const conditions = customs.slice(0, MAX_RECORDED_CONDITIONS).map(entry => {
        const ca = attrs(entry)
        return { operator: textAttr(ca.operator) ?? 'equal', value: textAttr(ca.val) ?? '' }
      })
      return { kind: 'custom', operator: a.and === '1' || a.and === 'true' ? 'and' : 'or', conditions, totalCount: customs.length }
    }
    case 'dynamicFilter':
      return {
        kind: 'dynamic',
        filterType: textAttr(a.type) ?? '',
        ...(textAttr(a.val) !== undefined ? { value: textAttr(a.val) } : {}),
        ...(textAttr(a.maxVal) !== undefined ? { maxValue: textAttr(a.maxVal) } : {}),
      }
    case 'colorFilter':
      return {
        kind: 'color',
        ...(textAttr(a.dxfId) !== undefined ? { dxfId: textAttr(a.dxfId) } : {}),
        ...(textAttr(a.cellColor) !== undefined ? { cellColor: textAttr(a.cellColor) } : {}),
      }
    default: return { kind: 'unknown', element: tag }
  }
}
/** Parses one <autoFilter> element; undefined when absent. Malformed input is
 * recorded, never thrown and never repaired into a range. */
export function parseDeclaredAutoFilter(node: XmlNode | undefined): DeclaredAutoFilter | undefined {
  if (!node) return undefined
  const a = attrs(node)
  const refVerbatim = textAttr(a.ref)
  const out: DeclaredAutoFilter = { columns: [] }
  if (refVerbatim === undefined) out.malformed = 'missing-ref'
  else {
    out.refVerbatim = refVerbatim
    const range = parseDeclaredRef(refVerbatim)
    if (!range) out.malformed = 'invalid-ref'
    else out.range = range
  }
  getChildren(node, 'filterColumn').forEach((child, index) => {
    const ca = attrs(child)
    const colId = intAttr(ca.colId)
    const column: DeclaredFilterColumn = { index, colId: undefined }
    if (colId !== undefined) column.colId = colId
    else column.malformed = 'invalid-colId'
    const hidden = boolAttr(ca.hiddenButton), shown = boolAttr(ca.showButton)
    if (hidden !== undefined) column.hiddenButton = hidden
    if (shown !== undefined) column.showButton = shown
    const rules = elementChildren(child)
    if (rules.length > 1) column.malformed = 'multiple-rule-elements'
    const first = rules[0]
    if (first) column.rule = ruleOf(first[0], first[1])
    out.columns.push(column)
  })
  return out
}
function sortConditionOf(node: XmlNode): DeclaredSortCondition {
  const a = attrs(node)
  const out: DeclaredSortCondition = { sortBy: sortByOf(textAttr(a.sortBy)) }
  const refVerbatim = textAttr(a.ref)
  if (refVerbatim === undefined) out.malformed = 'missing-ref'
  else {
    out.refVerbatim = refVerbatim
    const range = parseDeclaredRef(refVerbatim)
    if (!range) out.malformed = 'invalid-ref'
    else out.range = range
  }
  if (out.sortBy === 'unknown') out.sortByVerbatim = a.sortBy === undefined ? undefined : String(a.sortBy)
  const descending = boolAttr(a.descending)
  if (descending !== undefined) out.descending = descending
  const dxfId = textAttr(a.dxfId)
  if (dxfId !== undefined) out.dxfId = dxfId
  const customList = textAttr(a.customList)
  if (customList !== undefined) out.customListCount = customList.split(',').length
  return out
}
/** Parses one <sortState> element; undefined when absent. Stored order is
 * never touched: this is a description of what the file declares. */
export function parseDeclaredSortState(node: XmlNode | undefined): DeclaredSortState | undefined {
  if (!node) return undefined
  const a = attrs(node)
  const refVerbatim = textAttr(a.ref)
  const out: DeclaredSortState = { conditions: getChildren(node, 'sortCondition').map(sortConditionOf) }
  if (refVerbatim === undefined) out.malformed = 'missing-ref'
  else {
    out.refVerbatim = refVerbatim
    const range = parseDeclaredRef(refVerbatim)
    if (!range) out.malformed = 'invalid-ref'
    else out.range = range
  }
  const caseSensitive = boolAttr(a.caseSensitive)
  if (caseSensitive !== undefined) out.caseSensitive = caseSensitive
  return out
}
/** Parses a <table> part's <autoFilter>/<sortState>; undefined when neither is
 * present. Attribution is verbatim (displayName as authored, part path), so a
 * table whose metadata is invalid still reports what it declares. */
export function parseDeclaredTable(node: XmlNode | undefined, partPath: string): DeclaredTableFilters | undefined {
  if (!node) return undefined
  const autoFilter = parseDeclaredAutoFilter(getChildren(node, 'autoFilter')[0])
  if (!autoFilter) return undefined
  const displayName = textAttr(attrs(node).displayName)
  const out: DeclaredTableFilters = { table: { partPath, ...(displayName !== undefined ? { displayName } : {}) }, autoFilter }
  const sortState = parseDeclaredSortState(getChildren(node, 'sortState')[0])
  if (sortState) out.sortState = sortState
  return out
}
/** Whether the sheet declares any file filter/sort state (well-formed or not). */
export function hasDeclaredFiltering(entry: DeclaredSheetFilters | undefined): boolean {
  if (!entry) return false
  return entry.worksheetAutoFilter !== undefined || entry.worksheetSortState !== undefined || entry.tableAutoFilters.length > 0
}
function ruleSummary(rule: DeclaredFilterRule): string {
  switch (rule.kind) {
    case 'values': return `values (${rule.valueCount} listed${rule.hasBlank ? ', blanks' : ''})`
    case 'custom': return `custom (${rule.totalCount} conditions)`
    case 'top10': return 'Top 10'
    case 'dynamic': return `dynamic (${rule.filterType || 'untyped'})`
    case 'color': return 'color'
    default: return 'unsupported'
  }
}
/** One-line status copy for Clear/snapshot explanations (spec §7.5): the file
 * filter effects are the load snapshot, not editable viewer rules. */
export function summarizeDeclaredFiltering(entry: DeclaredSheetFilters): string {
  const parts: string[] = []
  const auto = entry.worksheetAutoFilter
  if (auto?.refVerbatim) {
    const kinds = auto.columns.map(column => column.rule ? ruleSummary(column.rule) : 'unfiltered').join(', ')
    parts.push(`a file filter on ${auto.refVerbatim} (${auto.columns.length} filtered columns: ${kinds})`)
  } else if (auto) parts.push('a file filter declaration the viewer could not read')
  if (entry.worksheetSortState?.refVerbatim) parts.push(`a stored sort on ${entry.worksheetSortState.refVerbatim}`)
  for (const table of entry.tableAutoFilters) {
    const label = table.table.displayName ?? table.table.partPath
    if (table.autoFilter.refVerbatim) parts.push(`table "${label}" declares a filter on ${table.autoFilter.refVerbatim}`)
    else parts.push(`table "${label}" declares a filter the viewer could not read`)
  }
  const declared = parts.length > 0 ? parts.join('; ') : 'file filter declarations'
  return `Sheet "${entry.name}" has ${declared}; hidden rows are part of the file snapshot and Clear restores this snapshot without creating editable filter rules.`
}
