import type { ColumnRule, FilterSpec, IntervalSet, RowOrder, TextOp } from './contracts'
import type { ResolvedSemantics } from '../formula/types'
import { lowerDateCondition, matchesDateInterval, type DateContext } from './datePredicates'
import { formatCommon } from '../formula/format'
import { validateCoordinate, validateSourceRect } from './coordinates'
import { FilterDiagnostic, TextFilterMatcher, WorkBudget, type FilterLimits } from './textFilter'
import { projectVisibility, normalizeFilter } from './visibility'

export { FilterDiagnostic } from './textFilter'
export type { FilterLimits, FilterWork } from './textFilter'
export type { NumberOp, TextOp, DateOp, CalendarDay, DatePreset, TypedCondition, CompoundFilter, ValuesFilter, ValueSelection, PagedKeySetRef, ColumnRule, FilterSpec } from './contracts'
/** Prepared authority from a complete body-domain scan. Display text is a
 * caller-owned, bounded style/value-generation lane; raw values stay exact. */
export type PreparedValue =
  | { kind: 'number'; value: number; displayText?: string; numFmtId?: number; formatCode?: string }
  | { kind: 'string'; value: string }
  | { kind: 'boolean'; value: boolean; displayText?: string }
  | { kind: 'error'; value: string }
  | { kind: 'blank' | 'unknown' | 'unavailable' }
export interface ColumnClassification { kind: 'number' | 'date' | 'text' | 'empty' | 'loading'; hasNumeric: boolean }
export interface FilterContext extends DateContext { locale: string }
export type CompiledRule = (value: PreparedValue, limits?: FilterLimits) => boolean
function resolved(value: PreparedValue): void {
  if (value.kind === 'unknown' || value.kind === 'unavailable') throw new FilterDiagnostic('incomplete', 'Filter incomplete: authoritative value is unresolved')
  if (value.kind === 'number' && !Number.isFinite(value.value)) throw new RangeError('Non-finite prepared numeric value')
}
/** formatCommon owns grammar validation. Its date-token classifier is private,
 * so this adapter only reads validated active-section tokens (no new grammar).
 * Fixed en-US probes avoid confusing its date-label locale limit with type. */
function dateFormatted(value: Extract<PreparedValue, { kind: 'number' }>, semantics: ResolvedSemantics, budget: WorkBudget): boolean {
  if (!value.formatCode) return value.numFmtId !== undefined && value.numFmtId >= 14 && value.numFmtId <= 22
  budget.text(value.formatCode)
  const probe = formatCommon(value.value, value.formatCode, { ...semantics, locale: 'en-US' })
  if (probe.kind !== 'text') return false
  const active = value.value < 0 ? 1 : value.value === 0 ? 2 : 0
  let section = 0, quoted = false, escaped = false, found = false
  const sections: boolean[] = []
  for (const point of value.formatCode) {
    if (quoted) {
      if (point === '"') quoted = false
      continue
    }
    if (escaped) { escaped = false; continue }
    if (point === '\\') { escaped = true; continue }
    if (point === '"') { quoted = true; continue }
    if (point === ';') { sections[section++] = found; found = false; continue }
    if (/[ymdhs]/i.test(point)) found = true
  }
  sections[section] = found
  return sections[active < sections.length ? active : 0] ?? false
}
export function classifyColumn(values: Iterable<PreparedValue>, complete: boolean, semantics: ResolvedSemantics, limits?: FilterLimits): ColumnClassification {
  const budget = new WorkBudget(limits)
  let hasNumeric = false, hasOther = false, hasDate = false, missing = !complete
  for (const value of values) {
    budget.charge(1)
    if (value.kind === 'unknown' || value.kind === 'unavailable') { missing = true; continue }
    resolved(value)
    if (value.kind === 'blank' || value.kind === 'error') continue
    if (value.kind === 'number') { hasNumeric = true; hasDate ||= dateFormatted(value, semantics, budget) }
    else hasOther = true
  }
  return { kind: missing ? 'loading' : hasOther ? 'text' : hasNumeric ? hasDate ? 'date' : 'number' : 'empty', hasNumeric }
}
export function filterMenus(classification: ColumnClassification): readonly string[] {
  switch (classification.kind) {
    case 'loading': return []
    case 'empty': return ['values', 'blanks']
    case 'number': return ['values', 'blanks', 'number']
    case 'date': return ['values', 'blanks', 'date']
    case 'text': return classification.hasNumeric ? ['values', 'blanks', 'text', 'number'] : ['values', 'blanks', 'text']
  }
}
export function rawValueKey(value: PreparedValue, classification: ColumnClassification): string | null {
  resolved(value)
  if (classification.kind === 'loading') throw new FilterDiagnostic('incomplete', 'Filter incomplete: column classification is loading')
  switch (value.kind) {
    case 'number': return `${classification.kind === 'date' ? 'd' : 'n'}:${value.value === 0 ? 0 : value.value}`
    case 'string': return `s:${value.value}`
    case 'boolean': return value.value ? 'b:1' : 'b:0'
    case 'error': return `e:${value.value}`
    default: return null
  }
}
export class FilterEngine {
  private readonly text: TextFilterMatcher
  readonly context: Readonly<FilterContext>
  constructor(context: FilterContext) {
    this.context = { ...context }
    this.text = new TextFilterMatcher(context.locale)
  }
  matchText(op: TextOp, pattern: string, wildcards: boolean, text: string, limits?: FilterLimits): boolean {
    return this.text.match(op, pattern, wildcards, text, limits)
  }
  compileRule(rule: ColumnRule, classification: ColumnClassification): CompiledRule {
    const columnKind = classification.kind
    if (columnKind === 'loading') throw new FilterDiagnostic('incomplete', 'Filter incomplete: column classification is loading')
    if (rule.kind === 'values') throw new FilterDiagnostic('unsupported-values', 'Values membership belongs to S8.2')
    if (rule.kind === 'compound') {
      if (rule.connective !== 'and' && rule.connective !== 'or') throw new RangeError('Invalid compound connective')
      const connective = rule.connective
      const first = this.compileRule(rule.first, classification), second = this.compileRule(rule.second, classification)
      return (value, limits) => {
        // Evaluate both lanes so a successful branch cannot disguise missing
        // authoritative values or an unprepared formatted-text lane.
        const budget = new WorkBudget(limits)
        const shared = { ...limits, work: budget.work }
        const a = first(value, shared), b = second(value, shared)
        return connective === 'and' ? a && b : a || b
      }
    }
    if (rule.kind === 'number') {
      if (!['eq', 'neq', 'gt', 'gte', 'lt', 'lte', 'between'].includes(rule.op)
        || !Number.isFinite(rule.first) || (rule.second !== undefined && !Number.isFinite(rule.second))) throw new RangeError('Invalid numeric condition')
      if (rule.op === 'between' && (rule.second === undefined || rule.first > rule.second)) throw new RangeError('Missing or reversed numeric bounds')
      const { op, first, second } = rule
      return (value, limits) => {
        new WorkBudget(limits)
        resolved(value)
        if (value.kind !== 'number') return false
        switch (op) {
          case 'eq': return value.value === first
          case 'neq': return value.value !== first
          case 'gt': return value.value > first
          case 'gte': return value.value >= first
          case 'lt': return value.value < first
          case 'lte': return value.value <= first
          case 'between': return value.value >= first && value.value <= second!
        }
      }
    }
    if (rule.kind === 'date' || rule.kind === 'datePreset') {
      const interval = lowerDateCondition(rule, this.context)
      return (value, limits) => {
        new WorkBudget(limits)
        resolved(value)
        return value.kind === 'number' && columnKind === 'date' && matchesDateInterval(value.value, interval)
      }
    }
    if (rule.kind !== 'text') throw new RangeError('Invalid condition kind')
    this.text.validate(rule.op, rule.first, rule.wildcards)
    const { op, first, wildcards } = rule
    return (value, limits) => {
      new WorkBudget(limits)
      resolved(value)
      if (value.kind === 'blank' || value.kind === 'error') return false
      if (value.kind === 'string') return this.matchText(op, first, wildcards, value.value, limits)
      if ((value.kind === 'number' || value.kind === 'boolean') && columnKind === 'text') {
        if (value.displayText === undefined) throw new FilterDiagnostic('missing-display-text', 'Missing formatted display text lane')
        return this.matchText(op, first, wildcards, value.displayText, limits)
      }
      return false
    }
  }
  compileFilter(spec: FilterSpec, columns: ReadonlyMap<number, ColumnClassification>, read: (row: number, col: number) => PreparedValue, limits?: FilterLimits): (row: number) => boolean {
    validateSourceRect(spec.range)
    normalizeFilter(spec)
    const budget = new WorkBudget(limits)
    const shared = { ...limits, maxWork: budget.maxWork, maxTextLength: budget.maxTextLength, work: budget.work }
    const compiled: { col: number; match: CompiledRule }[] = []
    for (const [col, rule] of spec.columns) {
      validateCoordinate(0, col)
      if (col < spec.range.c0 || col > spec.range.c1) throw new RangeError('Filter column outside range')
      const classification = columns.get(col)
      if (!classification) throw new FilterDiagnostic('incomplete', 'Filter incomplete: missing column classification')
      compiled.push({ col, match: this.compileRule(rule, classification) })
    }
    return row => {
      validateCoordinate(row, 0)
      budget.charge(1)
      let matched = true
      for (const entry of compiled) {
        budget.charge(1)
        const result = entry.match(read(row, entry.col), shared)
        matched = result && matched
      }
      return matched
    }
  }
  projectFilter(spec: FilterSpec | null, baseHidden: IntervalSet, order: RowOrder, columns: ReadonlyMap<number, ColumnClassification>, read: (row: number, col: number) => PreparedValue, count?: number, limits?: FilterLimits) {
    const normalized = normalizeFilter(spec)
    const matches = normalized ? this.compileFilter(normalized, columns, read, limits) : () => true
    return projectVisibility(baseHidden, normalized, order, matches, count)
  }
}
