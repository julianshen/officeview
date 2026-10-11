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
  /** Columns beyond MAX_DECLARED_ITEMS are counted, not retained. */
  omittedColumns?: number
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
  /** Direct worksheet/table child, or nested inside <autoFilter> (both valid). */
  placement?: 'direct' | 'nested'
  conditions: DeclaredSortCondition[]
  /** Conditions beyond MAX_DECLARED_ITEMS are counted, not retained. */
  omittedConditions?: number
  malformed?: string
}
export interface DeclaredTableRef {
  displayName?: string
  partPath: string
}
export interface DeclaredTableFilters {
  table: DeclaredTableRef
  /** Absent for sort-only tables (autoFilter is optional per the schema). */
  autoFilter?: DeclaredAutoFilter
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
/** Retained records per autoFilter/sortState; the rest are counted (§12.4). */
export const MAX_DECLARED_ITEMS = 1024
/** Column/sort kinds listed in status copy before "and N more". */
const MAX_SUMMARY_ITEMS = 8

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
  // Complete spelling only: parseInt would silently repair "0junk" to 0 and
  // "1.5" to 1. xsd:unsignedInt allows leading zeros, so digits-only is exact.
  if (typeof value !== 'string' || !/^\d+$/.test(value)) return undefined
  const n = parseInt(value, 10)
  return Number.isSafeInteger(n) && n >= 0 ? n : undefined
}
/** Attribute record as either adapter sees it (tree attrs or SAX attributes). */
export type DeclaredAttrs = Readonly<Record<string, unknown>>
/** One filterColumn rule child, reduced to what the kind needs: the tag, its
 * attributes and child counts. customFilters holds at most
 * MAX_RECORDED_CONDITIONS records, so a huge rule costs a counter, not an array. */
export interface DeclaredRuleInput {
  tag: string
  attrs: DeclaredAttrs
  filterCount: number
  dateGroupItemCount: number
  customFilters: DeclaredAttrs[]
  customFilterCount: number
}
/** One filterColumn: its attributes, the number of rule alternatives (extLst
 * excluded) and the first one. */
export interface DeclaredColumnInput {
  attrs: DeclaredAttrs
  ruleCount: number
  firstRule?: DeclaredRuleInput
}
export function declaredRuleOf(input: DeclaredRuleInput): DeclaredFilterRule {
  const a = input.attrs
  switch (classifyFilterRuleTag(input.tag)) {
    case 'filters':
      return {
        kind: 'values',
        valueCount: input.filterCount + input.dateGroupItemCount,
        hasBlank: a.blank === '1' || a.blank === 'true',
        hasDateGroups: input.dateGroupItemCount > 0,
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
      const conditions = input.customFilters.slice(0, MAX_RECORDED_CONDITIONS).map(ca => (
        { operator: textAttr(ca.operator) ?? 'equal', value: textAttr(ca.val) ?? '' }
      ))
      return { kind: 'custom', operator: a.and === '1' || a.and === 'true' ? 'and' : 'or', conditions, totalCount: input.customFilterCount }
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
    default: return { kind: 'unknown', element: input.tag }
  }
}
function refInto(out: { refVerbatim?: string; range?: DeclaredRange; malformed?: string }, value: unknown): void {
  const refVerbatim = textAttr(value)
  if (refVerbatim === undefined) {
    out.malformed = 'missing-ref'
    return
  }
  out.refVerbatim = refVerbatim
  const range = parseDeclaredRef(refVerbatim)
  if (!range) out.malformed = 'invalid-ref'
  else out.range = range
}
export interface DeclaredAutoFilterBuilder {
  addColumn(input: DeclaredColumnInput): void
  finish(): DeclaredAutoFilter
}
/** Adapter-neutral accumulator for one <autoFilter>: bounding and validation
 * live here so the tree and streaming adapters produce identical diagnostics.
 * Malformed input is recorded, never thrown and never repaired into a range. */
export function declaredAutoFilterBuilder(attrs: DeclaredAttrs): DeclaredAutoFilterBuilder {
  const out: DeclaredAutoFilter = { columns: [] }
  refInto(out, attrs.ref)
  let index = 0
  return {
    addColumn(input) {
      if (index >= MAX_DECLARED_ITEMS) {
        index++
        out.omittedColumns = (out.omittedColumns ?? 0) + 1
        return
      }
      const ca = input.attrs
      const colId = intAttr(ca.colId)
      const column: DeclaredFilterColumn = { index: index++, colId: undefined }
      if (colId !== undefined) column.colId = colId
      else column.malformed = 'invalid-colId'
      const hidden = boolAttr(ca.hiddenButton), shown = boolAttr(ca.showButton)
      if (hidden !== undefined) column.hiddenButton = hidden
      if (shown !== undefined) column.showButton = shown
      if (input.ruleCount > 1) column.malformed = 'multiple-rule-elements'
      if (input.firstRule) column.rule = declaredRuleOf(input.firstRule)
      out.columns.push(column)
    },
    finish() {
      // Second pass over retained columns: colId is a 0-based offset into the
      // declared range, so it is only meaningful against a known range. The
      // first flag wins; evidence is kept, never repaired into a valid index.
      const width = out.range ? out.range.maxCol - out.range.minCol + 1 : undefined
      const seen = new Set<number>()
      for (const column of out.columns) {
        if (column.colId === undefined || column.malformed !== undefined) {
          if (column.colId !== undefined) seen.add(column.colId)
          continue
        }
        if (width !== undefined && column.colId >= width) column.malformed = 'colId-out-of-range'
        else if (seen.has(column.colId)) column.malformed = 'duplicate-colId'
        else seen.add(column.colId)
      }
      return out
    },
  }
}
function sortConditionOf(a: DeclaredAttrs): DeclaredSortCondition {
  const out: DeclaredSortCondition = { sortBy: sortByOf(textAttr(a.sortBy)) }
  refInto(out, a.ref)
  if (out.sortBy === 'unknown') out.sortByVerbatim = a.sortBy === undefined ? undefined : String(a.sortBy)
  const descending = boolAttr(a.descending)
  if (descending !== undefined) out.descending = descending
  const dxfId = textAttr(a.dxfId)
  if (dxfId !== undefined) out.dxfId = dxfId
  const customList = textAttr(a.customList)
  if (customList !== undefined) out.customListCount = customList.split(',').length
  return out
}
export interface DeclaredSortStateBuilder {
  addCondition(attrs: DeclaredAttrs): void
  finish(): DeclaredSortState
}
/** Adapter-neutral accumulator for one <sortState>. Stored order is never
 * touched: this describes what the file declares. The caller attributes
 * placement (direct child vs nested inside <autoFilter>). */
export function declaredSortStateBuilder(attrs: DeclaredAttrs, placement: 'direct' | 'nested'): DeclaredSortStateBuilder {
  const out: DeclaredSortState = { placement, conditions: [] }
  return {
    addCondition(condition) {
      if (out.conditions.length >= MAX_DECLARED_ITEMS) {
        out.omittedConditions = (out.omittedConditions ?? 0) + 1
        return
      }
      out.conditions.push(sortConditionOf(condition))
    },
    finish() {
      refInto(out, attrs.ref)
      const caseSensitive = boolAttr(attrs.caseSensitive)
      if (caseSensitive !== undefined) out.caseSensitive = caseSensitive
      return out
    },
  }
}
/** Table attribution is verbatim (displayName as authored, part path), so a
 * table whose metadata is invalid still reports what it declares; undefined
 * when the table declares neither. */
export function declaredTableOf(
  attrs: DeclaredAttrs, partPath: string, autoFilter: DeclaredAutoFilter | undefined, sortState: DeclaredSortState | undefined,
): DeclaredTableFilters | undefined {
  if (!autoFilter && !sortState) return undefined
  const displayName = textAttr(attrs.displayName)
  const out: DeclaredTableFilters = { table: { partPath, ...(displayName !== undefined ? { displayName } : {}) } }
  if (autoFilter) out.autoFilter = autoFilter
  if (sortState) out.sortState = sortState
  return out
}

// Tree adapter: walks parsed part XML into the builders above.

function ruleInputOf(tag: string, node: XmlNode): DeclaredRuleInput {
  const customs = getChildren(node, 'customFilter')
  return {
    tag,
    attrs: attrs(node),
    filterCount: getChildren(node, 'filter').length,
    dateGroupItemCount: getChildren(node, 'dateGroupItem').length,
    customFilters: customs.slice(0, MAX_RECORDED_CONDITIONS).map(entry => attrs(entry)),
    customFilterCount: customs.length,
  }
}
/** Parses one <autoFilter> element; undefined when absent. */
export function parseDeclaredAutoFilter(node: XmlNode | undefined): DeclaredAutoFilter | undefined {
  if (!node) return undefined
  const builder = declaredAutoFilterBuilder(attrs(node))
  for (const child of getChildren(node, 'filterColumn')) {
    // extLst is extension metadata, not a filter criterion: it is neither a
    // rule alternative nor evidence of multiple rules.
    const rules = elementChildren(child).filter(([tag]) => tag !== 'extLst')
    const first = rules[0]
    builder.addColumn({ attrs: attrs(child), ruleCount: rules.length, ...(first ? { firstRule: ruleInputOf(first[0], first[1]) } : {}) })
  }
  return builder.finish()
}
/** Parses one <sortState> element; undefined when absent. */
export function parseDeclaredSortState(node: XmlNode | undefined, placement: 'direct' | 'nested'): DeclaredSortState | undefined {
  if (!node) return undefined
  const builder = declaredSortStateBuilder(attrs(node), placement)
  for (const condition of getChildren(node, 'sortCondition')) builder.addCondition(attrs(condition))
  return builder.finish()
}
/** Parses a <table> part's <autoFilter>/<sortState>; undefined when neither is
 * present. */
export function parseDeclaredTable(node: XmlNode | undefined, partPath: string): DeclaredTableFilters | undefined {
  if (!node) return undefined
  const autoNode = getChildren(node, 'autoFilter')[0]
  const autoFilter = parseDeclaredAutoFilter(autoNode)
  const sortState = pickDeclaredSortState(
    parseDeclaredSortState(getChildren(node, 'sortState')[0], 'direct'),
    parseDeclaredSortState(autoNode ? getChildren(autoNode, 'sortState')[0] : undefined, 'nested'),
  )
  return declaredTableOf(attrs(node), partPath, autoFilter, sortState)
}
/** Direct child wins over a nested <autoFilter>/<sortState>; the conflict is
 * flagged on the kept declaration (first flag wins, evidence never merged). */
export function pickDeclaredSortState(
  direct: DeclaredSortState | undefined, nested: DeclaredSortState | undefined,
): DeclaredSortState | undefined {
  if (!direct) return nested
  if (nested && direct.malformed === undefined) direct.malformed = 'duplicate-sortState'
  return direct
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
    const shown = auto.columns.slice(0, MAX_SUMMARY_ITEMS)
    const listed = shown.map(column => column.rule ? ruleSummary(column.rule) : 'unfiltered').join(', ')
    const hidden = auto.columns.length - shown.length + (auto.omittedColumns ?? 0)
    parts.push(`a file filter on ${auto.refVerbatim} (${auto.columns.length} filtered columns: ${listed}${hidden > 0 ? `; and ${hidden} more` : ''})`)
  } else if (auto) parts.push('a file filter declaration the viewer could not read')
  if (entry.worksheetSortState?.refVerbatim) parts.push(`a stored sort on ${entry.worksheetSortState.refVerbatim}`)
  entry.tableAutoFilters.slice(0, MAX_SUMMARY_ITEMS).forEach(table => {
    const label = table.table.displayName ?? table.table.partPath
    if (table.autoFilter?.refVerbatim) parts.push(`table "${label}" declares a filter on ${table.autoFilter.refVerbatim}`)
    else parts.push(`table "${label}" declares a filter the viewer could not read`)
  })
  const hiddenTables = entry.tableAutoFilters.length - Math.min(entry.tableAutoFilters.length, MAX_SUMMARY_ITEMS)
  if (hiddenTables > 0) parts.push(`and ${hiddenTables} more table declarations`)
  const declared = parts.length > 0 ? parts.join('; ') : 'file filter declarations'
  return `Sheet "${entry.name}" has ${declared}; hidden rows are part of the file snapshot and Clear restores this snapshot without creating editable filter rules.`
}
