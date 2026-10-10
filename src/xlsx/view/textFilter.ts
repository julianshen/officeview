import type { TextOp } from './contracts'
export interface FilterWork { used: number }
export interface FilterLimits { maxWork?: number; maxTextLength?: number; signal?: AbortSignal; work?: FilterWork }
export class FilterDiagnostic extends Error {
  constructor(readonly code: 'resource-limit' | 'cancelled' | 'incomplete' | 'missing-display-text' | 'unsupported-values', message: string) {
    super(message)
    this.name = 'FilterDiagnostic'
  }
}
export const DEFAULT_FILTER_WORK = 1_000_000
export const DEFAULT_FILTER_TEXT_LENGTH = 65_536
export class WorkBudget {
  readonly work: FilterWork
  readonly maxWork: number
  readonly maxTextLength: number
  constructor(readonly limits: FilterLimits = {}) {
    this.maxWork = limits.maxWork === undefined ? DEFAULT_FILTER_WORK : limits.maxWork
    this.maxTextLength = limits.maxTextLength === undefined ? DEFAULT_FILTER_TEXT_LENGTH : limits.maxTextLength
    for (const value of [this.maxWork, this.maxTextLength, limits.work?.used ?? 0]) {
      if (!Number.isSafeInteger(value) || value < 0) throw new RangeError('Invalid filter resource policy')
    }
    this.work = limits.work ?? { used: 0 }
    this.charge(0)
  }
  charge(amount: number): void {
    if (this.limits.signal?.aborted) throw new FilterDiagnostic('cancelled', 'Filter cancelled')
    if (amount > this.maxWork - this.work.used) throw new FilterDiagnostic('resource-limit', 'Filter resource work limit exceeded')
    this.work.used += amount
  }
  text(value: string): void {
    if (value.length > this.maxTextLength) throw new FilterDiagnostic('resource-limit', 'Filter resource text length exceeded')
    this.charge(value.length)
  }
}
type Token = { kind: 'literal'; text: string } | { kind: 'star' } | { kind: 'one' }
const TEXT_OPS: readonly TextOp[] = ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains']
/** The formula wildcard helper deliberately rejects Unicode and compares ASCII
 * characters individually. This contract needs host-collated literal runs of
 * differing lengths and code-point wildcard elements, so it cannot reuse it. */
function tokens(pattern: string, budget: WorkBudget): Token[] {
  budget.text(pattern)
  const result: Token[] = []
  let literal = '', escaped = false
  const flush = () => { if (literal) { result.push({ kind: 'literal', text: literal }); literal = '' } }
  for (const point of pattern) {
    if (escaped) {
      if (point !== '*' && point !== '?' && point !== '~') throw new RangeError('Wildcard escape must precede a metacharacter')
      literal += point; escaped = false
    }
    else if (point === '~') escaped = true
    else if (point === '*' || point === '?') { flush(); result.push({ kind: point === '*' ? 'star' : 'one' }) }
    else literal += point
  }
  if (escaped) throw new RangeError('Malformed wildcard trailing escape')
  flush()
  return result
}
export class TextFilterMatcher {
  readonly collator: Intl.Collator
  constructor(locale: string) {
    this.collator = new Intl.Collator(locale, { sensitivity: 'accent', numeric: false, usage: 'search' })
  }
  validate(op: TextOp, pattern: string, wildcards: boolean): void {
    if (!TEXT_OPS.includes(op) || typeof pattern !== 'string' || typeof wildcards !== 'boolean') throw new RangeError('Invalid text condition')
    const budget = new WorkBudget()
    if (wildcards) tokens(pattern, budget)
    else budget.text(pattern)
  }
  match(op: TextOp, pattern: string, wildcards: boolean, text: string, limits?: FilterLimits): boolean {
    if (!TEXT_OPS.includes(op) || typeof pattern !== 'string' || typeof text !== 'string' || typeof wildcards !== 'boolean') throw new RangeError('Invalid text condition')
    const budget = new WorkBudget(limits)
    budget.text(text)
    const negative = op === 'neq' || op === 'not-contains'
    const positive = op === 'neq' ? 'eq' : op === 'not-contains' ? 'contains' : op
    const points: number[] = [0]
    for (const point of text) points.push(points[points.length - 1] + point.length)
    const count = points.length - 1
    const equal = (literal: string, start: number, end: number) => {
      // Charge inspected UTF-16 length as well as candidate count, not merely
      // DP states: collation and slicing long candidates also consume work.
      budget.charge(1 + literal.length + points[end] - points[start])
      return this.collator.compare(literal, text.slice(points[start], points[end])) === 0
    }
    let matched = false
    if (!wildcards) {
      budget.text(pattern)
      if (positive === 'eq') matched = equal(pattern, 0, count)
      else {
        outer: for (let start = 0; start <= count; start++) {
          if (positive === 'begins' && start > 0) break
          const endStart = positive === 'ends' ? count : start
          for (let end = endStart; end <= count; end++) {
            if (equal(pattern, start, end)) { matched = true; break outer }
          }
        }
      }
    } else {
      const parsed = tokens(pattern, budget)
      if (positive === 'ends' || positive === 'contains') parsed.unshift({ kind: 'star' })
      if (positive === 'begins' || positive === 'contains') parsed.push({ kind: 'star' })
      let row = new Uint8Array(count + 1)
      row[0] = 1
      for (const token of parsed) {
        budget.charge(count + 1)
        const next = new Uint8Array(count + 1)
        if (token.kind === 'star') {
          next[0] = row[0]
          for (let end = 1; end <= count; end++) next[end] = row[end] || next[end - 1]
        } else if (token.kind === 'one') {
          for (let start = 0; start < count; start++) if (row[start]) next[start + 1] = 1
        } else {
          for (let start = 0; start <= count; start++) {
            if (!row[start]) continue
            // Empty slices matter for host-collation ignorables; the run may
            // also consume a ligature or decomposed spelling of another length.
            for (let end = start; end <= count; end++) if (equal(token.text, start, end)) next[end] = 1
          }
        }
        row = next
      }
      matched = Boolean(row[count])
    }
    return negative ? !matched : matched
  }
}
