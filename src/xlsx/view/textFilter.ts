import type { TextOp } from './contracts'
import { canonicalize, graphemeEnds, isTurkicLocale, keyBegins, keyContains, keyEnds, type FoldVariant } from './canonicalText'
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
  /** Start another source cell without resetting observations or the caller's scan quota. */
  resetAllowance(): void {
    this.used = 0
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
    this.admit(value.length)
    this.charge(value.length)
  }
  /** Raw-length admission without charging: the single text-length policy,
   * applied before any canonicalization or search-key computation. */
  admit(length: number): void {
    if (!Number.isSafeInteger(length) || length < 0) throw new RangeError('Invalid text length')
    this.charge(0)
    if (length > this.maxTextLength) throw new FilterDiagnostic('resource-limit', 'Filter resource text length exceeded')
  }
}
type Token = Readonly<{ kind: 'literal'; text: string } | { kind: 'star' } | { kind: 'one' }>
const TEXT_OPS: readonly TextOp[] = ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains']
/** The formula wildcard helper deliberately rejects Unicode and compares ASCII
 * characters individually. This contract needs canonical literal runs and
 * grapheme wildcard units, so it cannot reuse it. Tokenizing precedes
 * canonicalization so that escapes are judged on the raw pattern. */
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

/** Grapheme units of one compiled pattern, interned to positive IDs. A text
 * grapheme absent from the pattern maps to 0: it can never equal a pattern
 * unit and `?` accepts any unit, so the table stays bounded by the pattern
 * instead of growing with every scanned value. */
class GraphemeIds {
  private readonly ascii = new Int32Array(128)
  private readonly points = new Map<number, number>()
  private readonly clusters = new Map<string, number>()
  private next = 1
  private find(text: string, start: number, end: number, add: boolean): number {
    const point = text.codePointAt(start)!
    if (point < 128 && end - start === 1) return this.ascii[point] || (add ? this.ascii[point] = this.next++ : 0)
    if (end - start === (point > 0xffff ? 2 : 1)) {
      let id = this.points.get(point)
      if (id === undefined && add) this.points.set(point, id = this.next++)
      return id ?? 0
    }
    const cluster = text.slice(start, end)
    let id = this.clusters.get(cluster)
    if (id === undefined && add) this.clusters.set(cluster, id = this.next++)
    return id ?? 0
  }
  intern(canonical: string, into: number[]): void {
    const ends = new Int32Array(canonical.length)
    for (let k = 0, count = graphemeEnds(canonical, ends), start = 0; k < count; start = ends[k++]) into.push(this.find(canonical, start, ends[k], true))
  }
  /** Grapheme IDs of a canonical value, written over its end offsets. */
  lookup(canonical: string, budget: WorkBudget): Int32Array {
    budget.charge(canonical.length)
    const ids = new Int32Array(canonical.length), count = graphemeEnds(canonical, ids)
    for (let k = 0, start = 0; k < count; k++) {
      const end = ids[k]
      ids[k] = this.find(canonical, start, end, false)
      start = end
    }
    return ids.subarray(0, count)
  }
}
const ONE = -1
/** Pattern units between stars; `?` units are ONE. */
interface Segment { readonly ids: Int32Array; readonly failure?: Int32Array }
function segment(ids: readonly number[], budget: WorkBudget): Segment {
  const array = Int32Array.from(ids)
  if (array.includes(ONE)) return { ids: array }
  // KMP failure table for linear leftmost search over grapheme IDs.
  const failure = new Int32Array(array.length)
  for (let i = 1, length = 0; i < array.length; i++) {
    budget.charge(1)
    while (length > 0 && array[i] !== array[length]) { budget.charge(1); length = failure[length - 1] }
    if (array[i] === array[length]) length++
    failure[i] = length
  }
  return { ids: array, failure }
}
function matchesAt(text: Int32Array, pattern: Int32Array, at: number, budget: WorkBudget): boolean {
  budget.charge(pattern.length + 1)
  for (let k = 0; k < pattern.length; k++) if (pattern[k] !== ONE && pattern[k] !== text[at + k]) return false
  return true
}
/** Leftmost start in [from, limit - length] where the segment matches, or -1. */
function leftmost(text: Int32Array, part: Segment, from: number, limit: number, budget: WorkBudget): number {
  const { ids: needle, failure } = part
  if (!failure) {
    // `?` segments: direct comparison per candidate, O(text x segment).
    for (let at = from; at + needle.length <= limit; at++) if (matchesAt(text, needle, at, budget)) return at
    return -1
  }
  for (let i = from, length = 0; i < limit; i++) {
    budget.charge(1)
    while (length > 0 && text[i] !== needle[length]) { budget.charge(1); length = failure[length - 1] }
    if (text[i] === needle[length]) length++
    if (length === needle.length) return i + 1 - length
  }
  return -1
}
/** A compiled text condition. Calling it canonicalizes a raw value; `key`
 * accepts a value's canonical search key (canonicalSearchKey with the same
 * locale), e.g. one supplied by a cache through the S8.2 searchKeyOf port. */
export interface PreparedTextPredicate {
  (text: string, budget: WorkBudget): boolean
  /** Same admission and charges as the raw call: `rawLength` is the raw value's
   * UTF-16 length, admitted against maxTextLength and charged before the key. */
  key(canonical: string, rawLength: number, budget: WorkBudget): boolean
}

/** Canonical text matching (spec §7.1): literal conditions compare canonical
 * strings with endpoints on extended-grapheme boundaries; wildcard conditions
 * match `?` to one canonical grapheme and `*` to any run. Literal operations
 * and literal-only wildcard segments are linear (string prefilter, then KMP
 * over interned grapheme IDs); `*` segments are placed greedily leftmost,
 * which is exact for glob patterns, so there is no backtracking. */
export class TextFilterMatcher {
  readonly locale: string
  private readonly turkic: boolean
  constructor(locale: string) {
    this.turkic = isTurkicLocale(locale)
    this.locale = locale
  }
  validate(op: TextOp, pattern: string, wildcards: boolean, limits?: FilterLimits): void {
    this.prepare(op, pattern, wildcards, new WorkBudget(limits))
  }
  match(op: TextOp, pattern: string, wildcards: boolean, text: string, limits?: FilterLimits): boolean {
    const budget = new WorkBudget(limits)
    return this.prepare(op, pattern, wildcards, budget)(text, budget)
  }
  prepare(op: TextOp, pattern: string, wildcards: boolean, budget: WorkBudget): PreparedTextPredicate {
    if (!TEXT_OPS.includes(op) || typeof pattern !== 'string' || typeof wildcards !== 'boolean') throw new RangeError('Invalid text condition')
    const negative = op === 'neq' || op === 'not-contains'
    const positive = op === 'neq' ? 'eq' : op === 'not-contains' ? 'contains' : op
    const parsed = wildcards ? tokens(pattern, budget) : undefined
    if (!parsed) budget.text(pattern)
    const canon = (text: string) => {
      const result = canonicalize(text, this.turkic)
      budget.charge(result.length)
      return result
    }
    const interned = new GraphemeIds()
    let evaluate: (canonical: string, budget: WorkBudget) => boolean
    // A query whose canonical form is empty is no search term (spec §7.1).
    let inert = false
    if (!parsed || parsed.every(token => token.kind === 'literal')) {
      // Wildcard-free patterns use the literal matcher; tokens merged escapes.
      const needle = canon(parsed ? parsed.map(token => (token as Extract<Token, { kind: 'literal' }>).text).join('') : pattern)
      inert = !needle
      evaluate = (canonical, budget) => {
        budget.charge(1 + Math.min(canonical.length, needle.length))
        if (positive === 'eq') return canonical === needle
        // Boundary verification segments the value once: linear work.
        budget.charge(canonical.length)
        if (positive === 'begins') return keyBegins(canonical, needle)
        if (positive === 'ends') return keyEnds(canonical, needle)
        budget.charge(canonical.length)
        return keyContains(canonical, needle)
      }
    } else {
      const items: number[] = []
      if (positive === 'ends' || positive === 'contains') items.push(0)
      for (const token of parsed) {
        if (token.kind === 'literal') interned.intern(canon(token.text), items)
        else if (token.kind === 'one') items.push(ONE)
        else if (items[items.length - 1] !== 0) items.push(0)
      }
      if ((positive === 'begins' || positive === 'contains') && items[items.length - 1] !== 0) items.push(0)
      // Star (0) separated segments; adjacent stars were collapsed above.
      const parts: number[][] = [[]]
      for (const item of items) {
        budget.charge(1)
        if (item === 0) parts.push([])
        else parts[parts.length - 1].push(item)
      }
      const segments = parts.map(part => segment(part, budget))
      const first = segments[0], last = segments[segments.length - 1], middle = segments.slice(1, -1)
      const minimum = segments.reduce((sum, part) => sum + part.ids.length, 0)
      evaluate = (canonical, budget) => {
        const ids = interned.lookup(canonical, budget), count = ids.length
        if (count < minimum) return false
        if (segments.length === 1) return count === first.ids.length && matchesAt(ids, first.ids, 0, budget)
        if (!matchesAt(ids, first.ids, 0, budget)) return false
        const limit = count - last.ids.length
        let at = first.ids.length
        for (const part of middle) {
          const found = leftmost(ids, part, at, limit, budget)
          if (found < 0) return false
          at = found + part.ids.length
        }
        return at <= limit && matchesAt(ids, last.ids, limit, budget)
      }
    }
    const predicate = ((text: string, budget: WorkBudget) => {
      budget.charge(0)
      if (typeof text !== 'string') throw new RangeError('Invalid text condition')
      // Preflight raw length before normalization makes any copies.
      budget.text(text)
      if (inert) return true
      const canonical = canonicalize(text, this.turkic)
      budget.charge(canonical.length)
      return evaluate(canonical, budget) !== negative
    }) as PreparedTextPredicate
    predicate.key = (canonical, rawLength, budget) => {
      budget.charge(0)
      if (typeof canonical !== 'string') throw new RangeError('Invalid text condition')
      budget.admit(rawLength)
      budget.charge(rawLength)
      if (inert) return true
      budget.charge(canonical.length)
      return evaluate(canonical, budget) !== negative
    }
    return predicate
  }
}

/** Key-level matcher for the S8.2 value search: evaluates a text condition
 * against canonical search keys (canonicalSearchKey with the same fold
 * variant), applying the grapheme-boundary and wildcard rules of spec §7.1.
 * `rawLength` (the raw value's UTF-16 length) is admitted and charged exactly
 * as on the raw path; each key gets a fresh WorkBudget with `limits`. */
export function compileKeyMatcher(op: TextOp, pattern: string, wildcards: boolean, variant: FoldVariant, limits?: FilterLimits): (key: string, rawLength: number) => boolean {
  if (variant !== 'default' && variant !== 'turkic') throw new RangeError('Invalid fold variant')
  const predicate = new TextFilterMatcher(variant === 'turkic' ? 'tr' : 'und').prepare(op, pattern, wildcards, new WorkBudget(limits))
  return (key, rawLength) => predicate.key(key, rawLength, new WorkBudget(limits))
}
