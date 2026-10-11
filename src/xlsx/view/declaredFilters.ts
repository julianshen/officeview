import { attrs, elementChildren, getChildren, type XmlNode } from '../../core/xml'

/** Declared file filter/sort state as diagnostics only (spec §7.5, §13.3 items
 * 8-9). The small parser builds this from part XML; the large adapter produces
 * the identical structure later from its own traversal by feeding the same
 * adapter-neutral builders below. Nothing here ever becomes a viewer rule:
 * supported file filters persist as hidden rows, Clear restores the load
 * snapshot, and stored sortState never reorders source rows. */
export interface DeclaredRange {
  minRow: number
  minCol: number
  maxRow: number
  maxCol: number
}
/** Records that copy file text set `truncated` when any retained string was
 * cut to MAX_DECLARED_TEXT. */
export type DeclaredFilterRule =
  | { kind: 'values'; valueCount: number; hasBlank: boolean; hasDateGroups: boolean }
  | { kind: 'custom'; operator: 'and' | 'or'; conditions: { operator: string; value: string }[]; totalCount: number; truncated?: true }
  | { kind: 'top10'; top?: boolean; percent?: boolean; val?: string; filterVal?: string; truncated?: true }
  | { kind: 'dynamic'; filterType: string; value?: string; maxValue?: string; truncated?: true }
  | { kind: 'color'; dxfId?: string; cellColor?: string; truncated?: true }
  | { kind: 'unknown'; element: string; truncated?: true }
/** `malformed` is the first flag raised; `flags` lists every distinct flag in
 * the order raised, so a later problem is never lost behind an earlier one. */
interface Flagged {
  malformed?: string
  flags?: string[]
}
export interface DeclaredFilterColumn extends Flagged {
  index: number
  colId?: number
  hiddenButton?: boolean
  showButton?: boolean
  rule?: DeclaredFilterRule
}
export interface DeclaredAutoFilter extends Flagged {
  refVerbatim?: string
  range?: DeclaredRange
  columns: DeclaredFilterColumn[]
  /** Columns beyond MAX_DECLARED_ITEMS or the document byte budget are counted, not retained. */
  omittedColumns?: number
  truncated?: true
}
export type DeclaredSortBy = 'value' | 'cellColor' | 'fontColor' | 'icon' | 'unknown'
export interface DeclaredSortCondition extends Flagged {
  refVerbatim?: string
  range?: DeclaredRange
  sortBy: DeclaredSortBy
  sortByVerbatim?: string
  descending?: boolean
  dxfId?: string
  customListCount?: number
  truncated?: true
}
export interface DeclaredSortState extends Flagged {
  refVerbatim?: string
  range?: DeclaredRange
  caseSensitive?: boolean
  /** columnSort as authored; true is a left-to-right sort. */
  columnSort?: boolean
  sortMethod?: string
  /** Sort kinds the viewer cannot describe as a row order. */
  unsupported?: 'left-to-right'
  /** Direct worksheet/table child, or nested inside <autoFilter> (both valid). */
  placement?: 'direct' | 'nested'
  /** Valid sortStates in the other placement that were not retained (a direct
   * one wins over a nested one; having both is not malformed). */
  omittedSortStates?: number
  conditions: DeclaredSortCondition[]
  /** Conditions beyond MAX_DECLARED_ITEMS or the document byte budget are counted, not retained. */
  omittedConditions?: number
  truncated?: true
}
export interface DeclaredTableRef {
  displayName?: string
  partPath: string
  truncated?: true
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
  /** Declaring tables beyond MAX_DECLARED_TABLES or the document byte budget, counted not retained. */
  omittedTables?: number
  /** Repeated tablePart references to a part already seen on this sheet. */
  duplicateTableParts?: number
}
export type DeclaredFilterRuleTag = 'filters' | 'top10' | 'customFilters' | 'dynamicFilter' | 'colorFilter' | 'unknown'

const MAX_ROW = 1048576
const MAX_COL = 16384
const MAX_RECORDED_CONDITIONS = 8
/** Retained records per autoFilter/sortState; the rest are counted (§12.4). */
export const MAX_DECLARED_ITEMS = 1024
/** Retained declaring tables per sheet; the rest are counted (§12.4). */
export const MAX_DECLARED_TABLES = 64
/** Serialized size of retained column, sort-condition and table records per
 * document across all sheets and tables, so many declaring parts or long text
 * cannot multiply the per-item caps (§12.4 aggregate budgets). The records only
 * feed bounded status copy, so a small cap keeps the main-thread document lean. */
export const MAX_DECLARED_BYTES = 256 * 1024
/** Longest file string copied into a diagnostic record. */
export const MAX_DECLARED_TEXT = 256
/** Column/sort kinds listed in status copy before "and N more". */
const MAX_SUMMARY_ITEMS = 8
const MAX_SUMMARY_LABEL = 64
const MAX_SUMMARY_DECLARED = 1536

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
function flag(owner: Flagged, reason: string): void {
  if (owner.malformed === undefined) owner.malformed = reason
  const flags = owner.flags ?? (owner.flags = [])
  if (!flags.includes(reason)) flags.push(reason)
}
/** Copies file text into a record, cut to MAX_DECLARED_TEXT (never inside a
 * surrogate pair) with the owner flagged. */
function keep(text: string, owner: { truncated?: true }): string {
  if (text.length <= MAX_DECLARED_TEXT) return text
  owner.truncated = true
  return cut(text, MAX_DECLARED_TEXT)
}
function cut(text: string, length: number): string {
  const last = text.charCodeAt(length - 1)
  return text.slice(0, last >= 0xd800 && last <= 0xdbff ? length - 1 : length)
}
/** Remaining serialized bytes for one document's retained records; shared by
 * every builder. A record is admitted while any budget remains and charged its
 * final JSON size (finish-time flags reserved up front), so the overshoot is at
 * most one record (each record's text is already capped) plus the small
 * per-sheet wrapper and autoFilter/sortState shells built after exhaustion.
 * Both adapters build identical records, so they spend the budget identically. */
export interface DeclaredBudget {
  bytes: number
}
export function createDeclaredBudget(): DeclaredBudget {
  return { bytes: MAX_DECLARED_BYTES }
}
function hasRoom(budget: DeclaredBudget | undefined): boolean {
  return budget === undefined || budget.bytes > 0
}
function charge(budget: DeclaredBudget | undefined, record: unknown): void {
  if (budget) budget.bytes -= JSON.stringify(record).length
}
/** Worst-case growth of a column record from the finish() colId pass, which
 * adds at most one flag (out-of-range or duplicate) after admission. Reserved
 * when the column is charged so the budget covers the final record. It is not
 * refunded: a refund would reopen an exhausted budget after each autoFilter
 * and let every later one overshoot by a record. */
const FINISH_FLAG_RESERVE = JSON.stringify({ malformed: 'colId-out-of-range', flags: ['colId-out-of-range'] }).length

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
    case 'top10': {
      const rule: DeclaredFilterRule = { kind: 'top10' }
      const top = boolAttr(a.top), percent = boolAttr(a.percent), val = textAttr(a.val), filterVal = textAttr(a.filterVal)
      if (top !== undefined) rule.top = top
      if (percent !== undefined) rule.percent = percent
      if (val !== undefined) rule.val = keep(val, rule)
      if (filterVal !== undefined) rule.filterVal = keep(filterVal, rule)
      return rule
    }
    case 'customFilters': {
      const rule: DeclaredFilterRule = {
        kind: 'custom', operator: a.and === '1' || a.and === 'true' ? 'and' : 'or', conditions: [], totalCount: input.customFilterCount,
      }
      for (const ca of input.customFilters.slice(0, MAX_RECORDED_CONDITIONS)) {
        rule.conditions.push({ operator: keep(textAttr(ca.operator) ?? 'equal', rule), value: keep(textAttr(ca.val) ?? '', rule) })
      }
      return rule
    }
    case 'dynamicFilter': {
      const rule: DeclaredFilterRule = { kind: 'dynamic', filterType: '' }
      rule.filterType = keep(textAttr(a.type) ?? '', rule)
      const value = textAttr(a.val), maxValue = textAttr(a.maxVal)
      if (value !== undefined) rule.value = keep(value, rule)
      if (maxValue !== undefined) rule.maxValue = keep(maxValue, rule)
      return rule
    }
    case 'colorFilter': {
      const rule: DeclaredFilterRule = { kind: 'color' }
      const dxfId = textAttr(a.dxfId), cellColor = textAttr(a.cellColor)
      if (dxfId !== undefined) rule.dxfId = keep(dxfId, rule)
      if (cellColor !== undefined) rule.cellColor = keep(cellColor, rule)
      return rule
    }
    default: {
      const rule: DeclaredFilterRule = { kind: 'unknown', element: '' }
      rule.element = keep(input.tag, rule)
      return rule
    }
  }
}
/** The range is parsed from the full spelling; only the stored copy is cut. */
function refInto(out: Flagged & { refVerbatim?: string; range?: DeclaredRange; truncated?: true }, value: unknown): void {
  const refVerbatim = textAttr(value)
  if (refVerbatim === undefined) {
    flag(out, 'missing-ref')
    return
  }
  out.refVerbatim = keep(refVerbatim, out)
  const range = parseDeclaredRef(refVerbatim)
  if (!range) flag(out, 'invalid-ref')
  else out.range = range
}
export interface DeclaredAutoFilterBuilder {
  addColumn(input: DeclaredColumnInput): void
  finish(): DeclaredAutoFilter
}
/** Adapter-neutral accumulator for one <autoFilter>: bounding and validation
 * live here so the tree and streaming adapters produce identical diagnostics.
 * Malformed input is recorded, never thrown and never repaired into a range. */
export function declaredAutoFilterBuilder(attrs: DeclaredAttrs, budget?: DeclaredBudget): DeclaredAutoFilterBuilder {
  const out: DeclaredAutoFilter = { columns: [] }
  refInto(out, attrs.ref)
  let index = 0
  return {
    addColumn(input) {
      if (index >= MAX_DECLARED_ITEMS || !hasRoom(budget)) {
        index++
        out.omittedColumns = (out.omittedColumns ?? 0) + 1
        return
      }
      const ca = input.attrs
      const colId = intAttr(ca.colId)
      const column: DeclaredFilterColumn = { index: index++, colId: undefined }
      if (colId !== undefined) column.colId = colId
      else flag(column, 'invalid-colId')
      const hidden = boolAttr(ca.hiddenButton), shown = boolAttr(ca.showButton)
      if (hidden !== undefined) column.hiddenButton = hidden
      if (shown !== undefined) column.showButton = shown
      if (input.ruleCount > 1) flag(column, 'multiple-rule-elements')
      if (input.firstRule) column.rule = declaredRuleOf(input.firstRule)
      charge(budget, column)
      if (budget && column.colId !== undefined) budget.bytes -= FINISH_FLAG_RESERVE
      out.columns.push(column)
    },
    finish() {
      // Second pass over retained columns: colId is a 0-based offset into the
      // declared range, so it is only meaningful against a known range.
      // Evidence is kept, never repaired into a valid index.
      const width = out.range ? out.range.maxCol - out.range.minCol + 1 : undefined
      const seen = new Set<number>()
      for (const column of out.columns) {
        if (column.colId === undefined) continue
        if (width !== undefined && column.colId >= width) flag(column, 'colId-out-of-range')
        else if (seen.has(column.colId)) flag(column, 'duplicate-colId')
        else seen.add(column.colId)
      }
      return out
    },
  }
}
function sortConditionOf(a: DeclaredAttrs): DeclaredSortCondition {
  // An empty spelling is not ST_SortBy: only an absent attribute defaults.
  const sortBy = a.sortBy === undefined ? undefined : String(a.sortBy)
  const out: DeclaredSortCondition = { sortBy: sortByOf(sortBy) }
  refInto(out, a.ref)
  if (out.sortBy === 'unknown') {
    out.sortByVerbatim = keep(sortBy ?? '', out)
    flag(out, 'invalid-sortBy')
  }
  const descending = boolAttr(a.descending)
  if (descending !== undefined) out.descending = descending
  const dxfId = textAttr(a.dxfId)
  if (dxfId !== undefined) out.dxfId = keep(dxfId, out)
  // CT_SortCondition is a leaf element: customList is an ST_Xstring attribute
  // (ISO/IEC 29500-1 sortCondition), never a child element.
  const customList = textAttr(a.customList)
  if (customList !== undefined) {
    // Counted in place: splitting a long list would copy every entry.
    let count = 1
    for (let i = 0; i < customList.length; i++) if (customList.charCodeAt(i) === 44) count++
    out.customListCount = count
  }
  return out
}
export interface DeclaredSortStateBuilder {
  addCondition(attrs: DeclaredAttrs): void
  finish(): DeclaredSortState
}
/** Adapter-neutral accumulator for one <sortState>. Stored order is never
 * touched: this describes what the file declares. The caller attributes
 * placement (direct child vs nested inside <autoFilter>). */
export function declaredSortStateBuilder(
  attrs: DeclaredAttrs, placement: 'direct' | 'nested', budget?: DeclaredBudget,
): DeclaredSortStateBuilder {
  const out: DeclaredSortState = { placement, conditions: [] }
  return {
    addCondition(condition) {
      if (out.conditions.length >= MAX_DECLARED_ITEMS || !hasRoom(budget)) {
        out.omittedConditions = (out.omittedConditions ?? 0) + 1
        return
      }
      const record = sortConditionOf(condition)
      charge(budget, record)
      out.conditions.push(record)
    },
    finish() {
      refInto(out, attrs.ref)
      const caseSensitive = boolAttr(attrs.caseSensitive)
      if (caseSensitive !== undefined) out.caseSensitive = caseSensitive
      const columnSort = boolAttr(attrs.columnSort)
      if (columnSort !== undefined) out.columnSort = columnSort
      if (columnSort) out.unsupported = 'left-to-right'
      const sortMethod = textAttr(attrs.sortMethod)
      if (sortMethod !== undefined) out.sortMethod = keep(sortMethod, out)
      return out
    },
  }
}
export interface DeclaredContainerState {
  autoFilter?: DeclaredAutoFilter
  sortState?: DeclaredSortState
}
export interface DeclaredContainerBuilder {
  /** Builder for the first <autoFilter>; undefined (and flagged) for later ones. */
  autoFilter(attrs: DeclaredAttrs): DeclaredAutoFilterBuilder | undefined
  /** Builder for the first sortState of a placement; undefined (and flagged)
   * for later ones. Nested ones belong to the first <autoFilter> only. */
  sortState(attrs: DeclaredAttrs, placement: 'direct' | 'nested'): DeclaredSortStateBuilder | undefined
  finish(): DeclaredContainerState
}
/** One worksheet or table part's declarations, fed in document order. Only the
 * first declaration of each kind is kept; true repeats (same placement) are
 * flagged on it. A direct sortState wins over one nested in <autoFilter>; both
 * placements are schema-valid, so the dropped one is counted, not flagged. */
export function declaredContainerBuilder(budget?: DeclaredBudget): DeclaredContainerBuilder {
  let auto: DeclaredAutoFilterBuilder | undefined
  let autoRepeats = 0
  const sorts: Record<'direct' | 'nested', { builder?: DeclaredSortStateBuilder; repeats: number }> = {
    direct: { repeats: 0 }, nested: { repeats: 0 },
  }
  return {
    autoFilter(attrs) {
      if (auto) {
        autoRepeats++
        return undefined
      }
      return auto = declaredAutoFilterBuilder(attrs, budget)
    },
    sortState(attrs, placement) {
      const slot = sorts[placement]
      if (slot.builder) {
        slot.repeats++
        return undefined
      }
      return slot.builder = declaredSortStateBuilder(attrs, placement, budget)
    },
    finish() {
      const out: DeclaredContainerState = {}
      const autoFilter = auto?.finish()
      if (autoFilter) {
        if (autoRepeats > 0) flag(autoFilter, 'duplicate-autoFilter')
        out.autoFilter = autoFilter
      }
      const direct = sorts.direct.builder?.finish(), nested = sorts.nested.builder?.finish()
      if (direct && sorts.direct.repeats > 0) flag(direct, 'duplicate-sortState')
      if (nested && sorts.nested.repeats > 0) flag(nested, 'duplicate-sortState')
      if (direct && nested) direct.omittedSortStates = 1
      // Records were charged as built; charge the final shells (refs, flags)
      // once every container-level flag is in place.
      if (out.autoFilter) charge(budget, { ...out.autoFilter, columns: [] })
      if (out.sortState) charge(budget, { ...out.sortState, conditions: [] })
      const sortState = direct ?? nested
      if (sortState) out.sortState = sortState
      return out
    },
  }
}
/** Table attribution is verbatim (displayName as authored, part path), so a
 * table whose metadata is invalid still reports what it declares; undefined
 * when the table declares neither. */
export function declaredTableOf(attrs: DeclaredAttrs, partPath: string, state: DeclaredContainerState): DeclaredTableFilters | undefined {
  if (!state.autoFilter && !state.sortState) return undefined
  const table: DeclaredTableRef = { partPath }
  const displayName = textAttr(attrs.displayName)
  if (displayName !== undefined) table.displayName = keep(displayName, table)
  const out: DeclaredTableFilters = { table }
  if (state.autoFilter) out.autoFilter = state.autoFilter
  if (state.sortState) out.sortState = state.sortState
  return out
}
export interface DeclaredSheetBuilder {
  readonly entry: DeclaredSheetFilters
  readonly budget: DeclaredBudget
  /** Offers one resolved tablePart. `declares` says whether the part has any
   * autoFilter/sortState; `build` runs only for a new part with room left. */
  addTable(partPath: string, declares: boolean, build: () => DeclaredTableFilters | undefined): void
}
/** Per-sheet table accumulator: repeated references to one part are counted
 * once, and declaring tables beyond MAX_DECLARED_TABLES or the document byte
 * budget are counted, not built. */
export function declaredSheetBuilder(entry: DeclaredSheetFilters, budget: DeclaredBudget): DeclaredSheetBuilder {
  const seen = new Set<string>()
  return {
    entry,
    budget,
    addTable(partPath, declares, build) {
      if (seen.has(partPath)) {
        entry.duplicateTableParts = (entry.duplicateTableParts ?? 0) + 1
        return
      }
      seen.add(partPath)
      if (!declares) return
      if (entry.tableAutoFilters.length >= MAX_DECLARED_TABLES || !hasRoom(budget)) {
        entry.omittedTables = (entry.omittedTables ?? 0) + 1
        return
      }
      const table = build()
      if (!table) return
      // The container charged its records and shells; charge the attribution.
      charge(budget, { table: table.table })
      entry.tableAutoFilters.push(table)
    },
  }
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
function feedSortState(builder: DeclaredSortStateBuilder | undefined, node: XmlNode): void {
  if (!builder) return
  for (const condition of getChildren(node, 'sortCondition')) builder.addCondition(attrs(condition))
}
/** Parses a worksheet's or table's <autoFilter>/<sortState> children in
 * schema order (filterColumns, then the nested sortState, then direct ones). */
export function parseDeclaredContainer(node: XmlNode, budget?: DeclaredBudget): DeclaredContainerState {
  const container = declaredContainerBuilder(budget)
  for (const autoNode of getChildren(node, 'autoFilter')) {
    const builder = container.autoFilter(attrs(autoNode))
    if (!builder) continue
    for (const child of getChildren(autoNode, 'filterColumn')) {
      // extLst is extension metadata, not a filter criterion: it is neither a
      // rule alternative nor evidence of multiple rules.
      const rules = elementChildren(child).filter(([tag]) => tag !== 'extLst')
      const first = rules[0]
      builder.addColumn({ attrs: attrs(child), ruleCount: rules.length, ...(first ? { firstRule: ruleInputOf(first[0], first[1]) } : {}) })
    }
    for (const sortNode of getChildren(autoNode, 'sortState')) feedSortState(container.sortState(attrs(sortNode), 'nested'), sortNode)
  }
  for (const sortNode of getChildren(node, 'sortState')) feedSortState(container.sortState(attrs(sortNode), 'direct'), sortNode)
  return container.finish()
}
/** Offers one parsed <table> part to the sheet accumulator. */
export function collectDeclaredTable(sheet: DeclaredSheetBuilder, node: XmlNode, partPath: string): void {
  const declares = getChildren(node, 'autoFilter').length > 0 || getChildren(node, 'sortState').length > 0
  sheet.addTable(partPath, declares, () => declaredTableOf(attrs(node), partPath, parseDeclaredContainer(node, sheet.budget)))
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
    case 'dynamic': return `dynamic (${label(rule.filterType) || 'untyped'})`
    case 'color': return 'color'
    default: return 'unsupported'
  }
}
/** File text in status copy is cut short; the record keeps the longer copy. */
function label(text: string): string {
  return text.length > MAX_SUMMARY_LABEL ? `${cut(text, MAX_SUMMARY_LABEL - 1)}…` : text
}
function filterSummary(auto: DeclaredAutoFilter): string {
  if (!auto.range) {
    return auto.refVerbatim !== undefined
      ? `a file filter on an unreadable range "${label(auto.refVerbatim)}"`
      : 'a file filter the viewer could not read'
  }
  // Only columns carrying a rule filter anything; a bare filterColumn (e.g.
  // hiddenButton only) is button metadata.
  const ruled = auto.columns.filter(column => column.rule !== undefined)
  const shown = ruled.slice(0, MAX_SUMMARY_ITEMS).map(column => ruleSummary(column.rule!))
  const more = ruled.length - shown.length
  const details = ruled.length === 0
    ? 'no filtered columns'
    : `${ruled.length} filtered column${ruled.length === 1 ? '' : 's'}: ${shown.join(', ')}${more > 0 ? `, and ${more} more` : ''}`
  const unread = auto.omittedColumns ? `; ${auto.omittedColumns} more column declarations not read` : ''
  return `a file filter on ${label(auto.refVerbatim ?? '')} (${details}${unread})`
}
function sortSummary(sort: DeclaredSortState): string {
  const kind = sort.unsupported === 'left-to-right' ? 'a stored left-to-right sort' : 'a stored sort'
  const suffix = sort.unsupported ? ' (unsupported)' : ''
  if (sort.range) return `${kind} on ${label(sort.refVerbatim ?? '')}${suffix}`
  if (sort.refVerbatim !== undefined) return `${kind} on an unreadable range "${label(sort.refVerbatim)}"${suffix}`
  return `${kind} the viewer could not read${suffix}`
}
/** One-line status copy for Clear/snapshot explanations (spec §7.5): the file
 * filter effects are the load snapshot, not editable viewer rules. Bounded
 * regardless of how much the file declares. */
export function summarizeDeclaredFiltering(entry: DeclaredSheetFilters): string {
  const parts: string[] = []
  if (entry.worksheetAutoFilter) parts.push(filterSummary(entry.worksheetAutoFilter))
  if (entry.worksheetSortState) parts.push(sortSummary(entry.worksheetSortState))
  const shownTables = entry.tableAutoFilters.slice(0, MAX_SUMMARY_ITEMS)
  for (const table of shownTables) {
    const what: string[] = []
    if (table.autoFilter) what.push(filterSummary(table.autoFilter))
    if (table.sortState) what.push(sortSummary(table.sortState))
    parts.push(`table "${label(table.table.displayName ?? table.table.partPath)}" has ${what.join(' and ')}`)
  }
  const hiddenTables = entry.tableAutoFilters.length - shownTables.length + (entry.omittedTables ?? 0)
  if (hiddenTables > 0) parts.push(`and ${hiddenTables} more table declarations`)
  let declared = parts.length > 0 ? parts.join('; ') : 'file filter declarations'
  if (declared.length > MAX_SUMMARY_DECLARED) declared = `${cut(declared, MAX_SUMMARY_DECLARED - 1)}…`
  return `Sheet "${label(entry.name)}": ${declared}. Hidden rows are part of the file snapshot and Clear restores this snapshot without creating editable filter rules.`
}
