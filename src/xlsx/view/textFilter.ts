import type { TextOp } from './contracts'
export interface FilterWork { used: number }
export interface FilterLimits {
  /** Per compilation call / source cell; compound children share the allowance. */
  maxWork?: number
  maxTextLength?: number
  signal?: AbortSignal
  /** Instrumentation only: accumulated observations never impose a scan quota. */
  work?: FilterWork
  /** Explicit whole-scan quota, sized and owned by the worker/caller. */
  budget?: WorkBudget
}
export class FilterDiagnostic extends Error {
  constructor(readonly code: 'resource-limit' | 'cancelled' | 'incomplete' | 'missing-display-text' | 'unsupported-values', message: string) {
    super(message)
    this.name = 'FilterDiagnostic'
  }
}
/** Per-cell allowance, never an implicit limit on the complete row domain. */
export const DEFAULT_FILTER_WORK = 10_000_000
export const DEFAULT_FILTER_TEXT_LENGTH = 65_536
export class WorkBudget {
  readonly work: FilterWork
  private used = 0
  readonly maxWork: number
  readonly maxTextLength: number
  constructor(readonly limits: FilterLimits = {}) {
    this.maxWork = limits.maxWork === undefined ? DEFAULT_FILTER_WORK : limits.maxWork
    this.maxTextLength = limits.maxTextLength === undefined ? DEFAULT_FILTER_TEXT_LENGTH : limits.maxTextLength
    if (limits.work !== undefined && (limits.work === null || typeof limits.work !== 'object')) throw new RangeError('Invalid filter work counter')
    const used = limits.work === undefined ? 0 : limits.work.used
    for (const value of [this.maxWork, this.maxTextLength, used]) {
      if (!Number.isSafeInteger(value) || value < 0) throw new RangeError('Invalid filter resource policy')
    }
    if (limits.budget !== undefined && !(limits.budget instanceof WorkBudget)) throw new RangeError('Invalid cumulative filter budget')
    this.work = limits.work === undefined ? { used: 0 } : limits.work
    this.charge(0)
  }
  charge(amount: number): void {
    if (this.limits.signal?.aborted) throw new FilterDiagnostic('cancelled', 'Filter cancelled')
    if (amount > this.maxWork - this.used) throw new FilterDiagnostic('resource-limit', 'Filter resource work limit exceeded')
    this.limits.budget?.charge(amount)
    this.used += amount
    this.work.used += amount
  }
  text(value: string): void {
    if (value.length > this.maxTextLength) throw new FilterDiagnostic('resource-limit', 'Filter resource text length exceeded')
    this.charge(value.length)
  }
}
type Token = Readonly<{ kind: 'literal'; text: string } | { kind: 'star' } | { kind: 'one' }>
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
      // Strict ~ escapes match src/xlsx/formula/functions/wildcard.ts.
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
export class ReferenceTextFilterMatcher {
  readonly collator: Intl.Collator
  constructor(locale: string) {
    this.collator = new Intl.Collator(locale, { sensitivity: 'accent', numeric: false, usage: 'search' })
  }
  validate(op: TextOp, pattern: string, wildcards: boolean, limits?: FilterLimits): void {
    this.prepare(op, pattern, wildcards, new WorkBudget(limits))
  }
  prepare(op: TextOp, pattern: string, wildcards: boolean, budget: WorkBudget): (text: string, budget: WorkBudget) => boolean {
    if (!TEXT_OPS.includes(op) || typeof pattern !== 'string' || typeof wildcards !== 'boolean') throw new RangeError('Invalid text condition')
    const positive = op === 'neq' ? 'eq' : op === 'not-contains' ? 'contains' : op
    let parsed: readonly Token[] | undefined
    if (wildcards) {
      const list = tokens(pattern, budget)
      if (positive === 'ends' || positive === 'contains') list.unshift({ kind: 'star' })
      if (positive === 'begins' || positive === 'contains') list.push({ kind: 'star' })
      parsed = Object.freeze(list.map(token => Object.freeze(token)))
    } else budget.text(pattern)
    return this.prepared(op, pattern, parsed, budget)
  }
  protected prepared(op: TextOp, pattern: string, parsed: readonly Token[] | undefined, _budget: WorkBudget): (text: string, budget: WorkBudget) => boolean {
    return (text, budget) => this.evaluate(op, pattern, parsed, text, budget)
  }
  match(op: TextOp, pattern: string, wildcards: boolean, text: string, limits?: FilterLimits): boolean {
    const budget = new WorkBudget(limits)
    return this.prepare(op, pattern, wildcards, budget)(text, budget)
  }
  protected evaluate(op: TextOp, pattern: string, parsed: readonly Token[] | undefined, text: string, budget: WorkBudget, period?: number): boolean {
    if (typeof text !== 'string') throw new RangeError('Invalid text condition')
    budget.text(text)
    const negative = op === 'neq' || op === 'not-contains'
    const positive = op === 'neq' ? 'eq' : op === 'not-contains' ? 'contains' : op
    const points: number[] = [0]
    for (const point of text) points.push(points[points.length - 1] + point.length)
    const count = points.length - 1
    const comparisons = new Map<string, Map<number, boolean>>()
    const equal = (literal: string, start: number, end: number) => {
      if (period !== undefined && parsed) {
        const key = (start % period) * (count + 1) + end - start
        let run = comparisons.get(literal)
        if (!run) { run = new Map(); comparisons.set(literal, run) }
        budget.charge(1)
        const cached = run.get(key)
        if (cached !== undefined) return cached
        budget.charge(literal.length + points[end] - points[start])
        const matched = this.collator.compare(literal, text.slice(points[start], points[end])) === 0
        run.set(key, matched)
        return matched
      }
      // Charge inspected UTF-16 length as well as candidate count, not merely
      // DP states: collation and slicing long candidates also consume work.
      budget.charge(1 + literal.length + points[end] - points[start])
      return this.collator.compare(literal, text.slice(points[start], points[end])) === 0
    }
    let matched = false
    if (!parsed) {
      if (positive === 'eq') matched = equal(pattern, 0, count)
      else {
        outer: for (let start = 0; start <= count; start++) {
          if ((positive === 'begins' && start > 0) || (positive === 'contains' && period !== undefined && start >= period)) break
          const endStart = positive === 'ends' ? count : start
          for (let end = endStart; end <= count; end++) {
            if (equal(pattern, start, end)) { matched = true; break outer }
          }
        }
      }
    } else {
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

// Certificate boundary: the released ICU root/search tables have one context-
// free primary weight per printable ASCII character, common secondary weight,
// A-Z/a-z equality, and zero weights for U+200B. No all-alphabet contractions.
// https://github.com/unicode-org/icu/blob/release-74-2/icu4c/source/data/unidata/FractionalUCA.txt
// https://github.com/unicode-org/icu/blob/release-78.2/icu4c/source/data/unidata/FractionalUCA.txt
// Both releases: icu4c/source/data/coll/{root,en,en_US}.txt (root search tailoring).
// Custom ICU data, other versions/locales, and browser hosts retain reference semantics.
function certifiedAsciiBackend(collator: Intl.Collator): boolean {
  if (typeof process === 'undefined' || !process.versions || process.env.NODE_ICU_DATA !== undefined || process.env.ICU_DATA !== undefined
    || process.env.NODE_OPTIONS?.includes('--icu-data-dir')
    || process.execArgv.some(arg => arg.startsWith('--icu-data-dir'))) return false
  const versions = process.versions
  // This guard assumes the reported runtime uses its released, unmodified ICU data.
  const backend = (versions.bun === '1.3.10' && versions.icu === '74.2' && versions.unicode === '15.1')
    || (!versions.bun && versions.node === '22.22.3' && versions.icu === '78.2' && versions.cldr === '48.0' && versions.unicode === '17.0')
  const native = (value: unknown) => typeof value === 'function' && Function.prototype.toString.call(value).includes('[native code]')
  if (!backend || !native(Intl.Collator) || !native(collator.compare) || !native(collator.resolvedOptions)
    || !native(Object.getOwnPropertyDescriptor(Intl.Collator.prototype, 'compare')?.get)) return false
  const options = collator.resolvedOptions()
  return (options.locale === 'en' || options.locale === 'en-US') && options.usage === 'search' && options.sensitivity === 'accent'
    && !options.numeric && !options.ignorePunctuation && options.collation === 'default' && options.caseFirst === 'false'
}
const CERTIFIED_ASCII = /^[\x20-\x7e\u200b]*$/
const foldCertified = (value: string) => value.replace(/\u200b/g, '').toLowerCase()
interface LiteralSearch { readonly text: string; readonly failure: readonly number[] }
function prepareLiteral(text: string, budget: WorkBudget): LiteralSearch {
  const failure = Array<number>(text.length).fill(0)
  for (let i = 1, length = 0; i < text.length; i++) {
    budget.charge(1)
    while (length > 0 && text[i] !== text[length]) { budget.charge(1); length = failure[length - 1] }
    if (text[i] === text[length]) length++
    failure[i] = length
  }
  return Object.freeze({ text, failure: Object.freeze(failure) })
}
function literalOccurrences(text: string, literal: LiteralSearch, budget: WorkBudget): Uint8Array {
  const result = new Uint8Array(text.length + 1), needle = literal.text
  if (!needle.length) { budget.charge(result.length); result.fill(1); return result }
  for (let i = 0, length = 0; i < text.length; i++) {
    budget.charge(1)
    while (length > 0 && text[i] !== needle[length]) { budget.charge(1); length = literal.failure[length - 1] }
    if (text[i] === needle[length]) length++
    if (length === needle.length) { result[i + 1 - length] = 1; length = literal.failure[length - 1] }
  }
  return result
}

export class TextFilterMatcher extends ReferenceTextFilterMatcher {
  readonly certifiedAsciiProfile: boolean
  constructor(locale: string) {
    super(locale)
    this.certifiedAsciiProfile = certifiedAsciiBackend(this.collator)
  }
  protected prepared(op: TextOp, pattern: string, parsed: readonly Token[] | undefined, budget: WorkBudget): (text: string, budget: WorkBudget) => boolean {
    const positive = op === 'neq' ? 'eq' : op === 'not-contains' ? 'contains' : op
    if (parsed) {
      const start = positive === 'ends' || positive === 'contains' ? 1 : 0
      const end = parsed.length - (positive === 'begins' || positive === 'contains' ? 1 : 0)
      if (end - start === 1 && parsed[start].kind === 'literal') {
        pattern = (parsed[start] as Extract<Token, { kind: 'literal' }>).text
        parsed = undefined
      }
    }
    if (!this.certifiedAsciiProfile || (!parsed ? !CERTIFIED_ASCII.test(pattern) : parsed.some(token => token.kind === 'literal' && !CERTIFIED_ASCII.test(token.text)))) {
      return super.prepared(op, pattern, parsed, budget)
    }
    const literals = new Map<string, LiteralSearch>()
    for (const raw of parsed ? parsed.filter((token): token is Extract<Token, { kind: 'literal' }> => token.kind === 'literal').map(token => token.text) : [pattern]) {
      if (!literals.has(raw)) literals.set(raw, prepareLiteral(foldCertified(raw), budget))
    }
    return (text, budget) => {
      budget.charge(0)
      if (typeof text === 'string' && text.length > budget.maxTextLength) budget.text(text)
      if (typeof text !== 'string' || !CERTIFIED_ASCII.test(text)) return this.evaluate(op, pattern, parsed, text, budget)
      budget.text(text)
      const folded = foldCertified(text), negative = op === 'neq' || op === 'not-contains'
      const positive = op === 'neq' ? 'eq' : op === 'not-contains' ? 'contains' : op
      let matched: boolean
      if (!parsed) {
        const literal = literals.get(pattern)!
        budget.charge(literal.text.length + 1)
        if (positive === 'eq') matched = folded === literal.text
        else if (positive === 'begins') matched = folded.startsWith(literal.text)
        else if (positive === 'ends') matched = folded.endsWith(literal.text)
        else {
          const hits = literalOccurrences(folded, literal, budget)
          budget.charge(hits.length)
          matched = hits.some(value => value === 1)
        }
      } else {
        // Prefix counts preserve every raw code-point boundary in a plateau of
        // ignored characters. Literal transitions union all endpoint plateaus;
        // ? still consumes exactly one raw element, including an ignored one.
        const weights = new Uint32Array(text.length + 1)
        for (let i = 0; i < text.length; i++) weights[i + 1] = weights[i] + (text[i] === '\u200b' ? 0 : 1)
        let row = new Uint8Array(text.length + 1)
        row[0] = 1
        const occurrences = new Map<string, Uint8Array>()
        for (const token of parsed) {
          budget.charge(text.length + 1)
          const next = new Uint8Array(text.length + 1)
          if (token.kind === 'star') {
            next[0] = row[0]
            for (let end = 1; end <= text.length; end++) next[end] = row[end] || next[end - 1]
          } else if (token.kind === 'one') {
            for (let start = 0; start < text.length; start++) if (row[start]) next[start + 1] = 1
          } else {
            const literal = literals.get(token.text)!
            let hits = occurrences.get(token.text)
            if (!hits) { hits = literalOccurrences(folded, literal, budget); occurrences.set(token.text, hits) }
            const targets = new Uint8Array(folded.length + 1)
            budget.charge((text.length + 1) * 2)
            for (let start = 0; start <= text.length; start++) {
              const target = weights[start] + literal.text.length
              if (row[start] && hits[weights[start]] && target <= folded.length) targets[target] = 1
            }
            for (let end = 0; end <= text.length; end++) next[end] = targets[weights[end]]
          }
          row = next
        }
        matched = Boolean(row[text.length])
      }
      return negative ? !matched : matched
    }
  }
  protected evaluate(op: TextOp, pattern: string, parsed: readonly Token[] | undefined, text: string, budget: WorkBudget): boolean {
    if (typeof text !== 'string') throw new RangeError('Invalid text condition')
    const positive = op === 'neq' ? 'eq' : op === 'not-contains' ? 'contains' : op
    const negative = op === 'neq' || op === 'not-contains'
    if (parsed) {
      const start = positive === 'ends' || positive === 'contains' ? 1 : 0
      const end = parsed.length - (positive === 'begins' || positive === 'contains' ? 1 : 0)
      const core = parsed.slice(start, end)
      if (core.length === 1 && core[0].kind === 'literal') {
        pattern = core[0].text
        parsed = undefined
      }
    }
    if (!parsed && positive === 'eq') return super.evaluate(op, pattern, parsed, text, budget)
    let period: number | undefined
    if (/^[\x20-\x7e]*$/.test(text)) {
      budget.charge(text.length)
      if (!parsed && /^[\x20-\x7e]*$/.test(pattern) && positive !== 'eq') {
        const folded = text.toLowerCase(), needle = pattern.toLowerCase()
        const at = positive === 'begins' ? (folded.startsWith(needle) ? 0 : -1)
          : positive === 'ends' ? (folded.endsWith(needle) ? text.length - pattern.length : -1) : folded.indexOf(needle)
        if (at >= 0) {
          budget.text(text)
          budget.charge(1 + pattern.length * 2)
          // A host-verified actual slice is an exact witness, including tailored locales.
          if (this.collator.compare(pattern, text.slice(at, at + pattern.length)) === 0) return !negative
        }
      }
      // Raw periodicity, not a guessed collation profile: every candidate at
      // later starts is byte-identical to one at start % period. No Intl
      // contraction/ignorable assumptions or Unicode endpoint bounds apply.
      for (let candidate = 1; (parsed || positive === 'contains') && candidate <= Math.min(16, text.length); candidate++) {
        let repeats = true
        for (let i = candidate; i < text.length; i++) {
          budget.charge(1)
          if (text[i] !== text[i % candidate]) { repeats = false; break }
        }
        if (repeats) { period = candidate; break }
      }
    }
    return super.evaluate(op, pattern, parsed, text, budget, period)
  }
}
