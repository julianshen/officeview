import type { ColumnRule, FilterSpec, IntervalSet, RowOrder, TextOp } from './contracts'
import { lowerDateCondition, matchesDateInterval, type DateContext } from './datePredicates'
import { validateCoordinate, validateSourceRect } from './coordinates'
import { FilterDiagnostic, TextFilterMatcher, WorkBudget, type FilterLimits } from './textFilter'
import { projectVisibility, normalizeFilter } from './visibility'

export { FilterDiagnostic, WorkBudget } from './textFilter'
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
/** Classification uses the positive (first) section regardless of value sign.
 * Brackets and authored literals do not contribute date tokens. */
function dateFormatted(value: Extract<PreparedValue, { kind: 'number' }>, budget: WorkBudget): boolean {
  if (value.formatCode === undefined) {
    const id = value.numFmtId
    return id !== undefined && ((id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 50 && id <= 58))
  }
  budget.text(value.formatCode)
  let active = '', quoted = false, escaped = false, bracket = '', bracketed = false, elapsed = false
  for (const point of value.formatCode) {
    if (bracketed) {
      if (point === ']') { elapsed ||= /^(h{1,2}|m{1,2}|s{1,2})$/i.test(bracket); bracketed = false }
      else bracket += point
      continue
    }
    if (quoted) { if (point === '"') quoted = false; continue }
    if (escaped) { escaped = false; continue }
    if (point === '\\' || point === '_' || point === '*') { escaped = true; continue }
    if (point === '"') { quoted = true; continue }
    if (point === '[') { bracketed = true; bracket = ''; continue }
    if (point === ';') break
    active += point
  }
  if (elapsed || quoted || bracketed || escaped) return false
  // Exponent E is numeric syntax; era e is a date token otherwise.
  active = active.replace(/[eE][+-]0+/g, '').replace(/General/gi, '')
  // Definitive date tokens classify even combinations the renderer cannot format.
  if (/[ydgebh]/i.test(active) || /am\/pm|a\/p/i.test(active)) return true
  if (/[0#?@]/.test(active.replace(/s+\.0+/gi, 's'))) return false
  return /m/i.test(active) && !/s/i.test(active)
}
export function classifyColumn(values: Iterable<PreparedValue>, complete: boolean, limits?: FilterLimits): ColumnClassification {
  const budget = new WorkBudget(limits)
  let hasNumeric = false, hasOther = false, hasDate = false, missing = !complete
  for (const value of values) {
    budget.resetAllowance()
    budget.charge(1)
    if (value.kind === 'unknown' || value.kind === 'unavailable') { missing = true; continue }
    resolved(value)
    if (value.kind === 'blank' || value.kind === 'error') continue
    if (value.kind === 'number') { hasNumeric = true; hasDate ||= dateFormatted(value, budget) }
    else hasOther = true
  }
  return { kind: missing ? 'loading' : hasOther ? 'text' : hasNumeric ? hasDate ? 'date' : 'number' : 'empty', hasNumeric }
}
export type FilterMenu = 'values' | 'blanks' | 'text' | 'number' | 'date'
export function filterMenus(classification: ColumnClassification): readonly FilterMenu[] {
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
  compileRule(rule: ColumnRule, classification: ColumnClassification, limits?: FilterLimits): CompiledRule {
    const compilation = new WorkBudget(limits)
    const evaluate = this.prepareRule(rule, classification, compilation)
    return (value, override = limits) => evaluate(value, new WorkBudget(override))
  }
  private prepareRule(rule: ColumnRule, classification: ColumnClassification, compilation: WorkBudget): (value: PreparedValue, budget: WorkBudget) => boolean {
    compilation.charge(1)
    if (!rule || typeof rule !== 'object' || typeof rule.kind !== 'string') throw new RangeError('Invalid condition kind')
    const columnKind = classification.kind
    if (columnKind === 'loading') throw new FilterDiagnostic('incomplete', 'Filter incomplete: column classification is loading')
    if (rule.kind === 'values') throw new FilterDiagnostic('unsupported-values', 'Values membership belongs to S8.2')
    if (rule.kind === 'compound') {
      for (const child of [rule.first, rule.second]) {
        if (!child || typeof child !== 'object' || !['number', 'text', 'date', 'datePreset'].includes(child.kind)) throw new RangeError('Compound children must be typed conditions')
      }
      if (rule.connective !== 'and' && rule.connective !== 'or') throw new RangeError('Invalid compound connective')
      const connective = rule.connective
      const first = this.prepareRule(rule.first, classification, compilation), second = this.prepareRule(rule.second, classification, compilation)
      return (value, budget) => {
        // Evaluate both lanes so a successful branch cannot disguise missing
        // authoritative values or an unprepared formatted-text lane.
        budget.charge(0)
        const a = first(value, budget), b = second(value, budget)
        return connective === 'and' ? a && b : a || b
      }
    }
    if (rule.kind === 'number') {
      if (!['eq', 'neq', 'gt', 'gte', 'lt', 'lte', 'between'].includes(rule.op)
        || !Number.isFinite(rule.first) || (rule.second !== undefined && !Number.isFinite(rule.second))) throw new RangeError('Invalid numeric condition')
      if (rule.op === 'between' && (rule.second === undefined || rule.first > rule.second)) throw new RangeError('Missing or reversed numeric bounds')
      const { op, first, second } = rule
      return (value, budget) => {
        budget.charge(0)
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
      return (value, budget) => {
        budget.charge(0)
        resolved(value)
        return value.kind === 'number' && columnKind === 'date' && matchesDateInterval(value.value, interval)
      }
    }
    if (rule.kind !== 'text') throw new RangeError('Invalid condition kind')
    const match = this.text.prepare(rule.op, rule.first, rule.wildcards, compilation)
    return (value, budget) => {
      budget.charge(0)
      resolved(value)
      if (value.kind === 'blank' || value.kind === 'error') return false
      if (value.kind === 'string') return match(value.value, budget)
      if ((value.kind === 'number' || value.kind === 'boolean') && columnKind === 'text') {
        if (value.displayText === undefined) throw new FilterDiagnostic('missing-display-text', 'Missing formatted display text lane')
        return match(value.displayText, budget)
      }
      return false
    }
  }
  compileFilter(spec: FilterSpec, columns: ReadonlyMap<number, ColumnClassification>, read: (row: number, col: number) => PreparedValue, limits?: FilterLimits): (row: number) => boolean {
    validateSourceRect(spec.range)
    normalizeFilter(spec)
    const budget = new WorkBudget(limits)
    const compiled: { col: number; match: (value: PreparedValue, budget: WorkBudget) => boolean }[] = []
    for (const [col, rule] of spec.columns) {
      budget.charge(1)
      validateCoordinate(0, col)
      if (col < spec.range.c0 || col > spec.range.c1) throw new RangeError('Filter column outside range')
      const classification = columns.get(col)
      if (!classification) throw new FilterDiagnostic('incomplete', 'Filter incomplete: missing column classification')
      compiled.push({ col, match: this.prepareRule(rule, classification, budget) })
    }
    return row => {
      validateCoordinate(row, 0)
      new WorkBudget(limits).charge(0)
      let matched = true
      for (const entry of compiled) {
        // Each source column cell owns its allowance; compound children share it.
        const cell = new WorkBudget(limits)
        cell.charge(1)
        const result = entry.match(read(row, entry.col), cell)
        matched = result && matched
      }
      return matched
    }
  }
  projectFilter(spec: FilterSpec | null, baseHidden: IntervalSet, order: RowOrder, columns: ReadonlyMap<number, ColumnClassification>, read: (row: number, col: number) => PreparedValue, count?: number, limits?: FilterLimits) {
    const budget = new WorkBudget(limits)
    const normalized = normalizeFilter(spec)
    const matches = normalized ? this.compileFilter(normalized, columns, read, limits) : () => true
    const result = projectVisibility(baseHidden, normalized, order, matches, count)
    budget.charge(0)
    return result
  }
}
