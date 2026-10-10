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
    budget.charge(0)
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

/** Certificates accept the empirical, context-independent equality model for
 * the actual alphabet, not a universal theorem about Intl. All raw pairs and
 * relevant triples (observed adjacency, plus exhaustive triples for tiny
 * alphabets) must agree with class concatenation and have no cross-length
 * equality collisions. Canonical/compatibility expansions and combining-mark
 * contexts are inconclusive. Only certified alphabets admit class-ID KMP/DP;
 * other inputs retain the exact raw-slice reference and its explicit budget.
 * An observed expansion factor alone cannot bound every future candidate end.
 * In particular, ignored plateaus have unbounded raw length, and ? consumes raw
 * code points. We therefore do not infer bounded windows for failed profiles.
 */
const COLLATOR_CONSTRUCTOR = Intl.Collator
const COLLATOR_COMPARE_GETTER = Object.getOwnPropertyDescriptor(Intl.Collator.prototype, 'compare')?.get
const MAX_ALPHABET = 40
const MAX_PROBE_WORK = 90_000
const MAX_PROBE_CACHE = 8192
const MAX_PROFILES = 128
const MAX_PROFILE_KEY = 8192
interface AlphabetProfile { readonly classes: ReadonlyMap<string, number> }
interface ProbeCache {
  readonly comparisons: Map<string, number>
  readonly profiles: Map<string, AlphabetProfile | null>
}
const probeCaches = new Map<string, ProbeCache>()
class InconclusiveProbe extends Error {}
function retain<K, V>(map: Map<K, V>, key: K, value: V, maximum: number): void {
  if (map.size >= maximum) map.delete(map.keys().next().value!)
  map.set(key, value)
}
interface LiteralSearch { readonly ids: readonly number[]; readonly failure: readonly number[] }
function prepareLiteral(ids: readonly number[], budget: WorkBudget): LiteralSearch {
  const failure = Array<number>(ids.length).fill(0)
  for (let i = 1, length = 0; i < ids.length; i++) {
    budget.charge(1)
    while (length > 0 && ids[i] !== ids[length]) { budget.charge(1); length = failure[length - 1] }
    if (ids[i] === ids[length]) length++
    failure[i] = length
  }
  return { ids, failure }
}
function literalOccurrences(text: readonly number[], literal: LiteralSearch, budget: WorkBudget): Uint8Array {
  const result = new Uint8Array(text.length + 1), needle = literal.ids
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
  /** Compatibility name: reports the LAST actual alphabet, including Unicode.
   * False before the first evaluation; no speculative constructor certificate. */
  certifiedAsciiProfile = false
  readonly certification = { comparisons: 0, cacheHits: 0, chargedWork: 0 }
  private readonly ownerCompare: Intl.Collator['compare']
  private readonly ownerOptions: Intl.Collator['resolvedOptions']
  private readonly nativeOwner: boolean
  private readonly config: string
  constructor(locale: string) {
    super(locale)
    this.ownerCompare = this.collator.compare
    this.ownerOptions = this.collator.resolvedOptions
    const native = (value: unknown) => typeof value === 'function' && Function.prototype.toString.call(value).includes('[native code]')
    this.nativeOwner = Intl.Collator === COLLATOR_CONSTRUCTOR && native(Intl.Collator)
      && Object.getOwnPropertyDescriptor(Intl.Collator.prototype, 'compare')?.get === COLLATOR_COMPARE_GETTER
      && native(COLLATOR_COMPARE_GETTER)
      && native(this.ownerCompare) && native(this.ownerOptions)
    this.config = JSON.stringify(this.collator.resolvedOptions())
  }
  private profile(runs: readonly string[], text: string, budget: WorkBudget): AlphabetProfile | undefined {
    this.certifiedAsciiProfile = false
    this.certification.comparisons = this.certification.cacheHits = this.certification.chargedWork = 0
    if (!this.nativeOwner || Intl.Collator !== COLLATOR_CONSTRUCTOR
      || Object.getOwnPropertyDescriptor(Intl.Collator.prototype, 'compare')?.get !== COLLATOR_COMPARE_GETTER
      || this.collator.compare !== this.ownerCompare
      || this.collator.resolvedOptions !== this.ownerOptions) return undefined
    const alphabet = new Set<string>(), triples = new Set<string>()
    for (const run of [...runs, text]) {
      let previous = '', before = ''
      for (const point of run) {
        budget.charge(1)
        alphabet.add(point)
        if (alphabet.size > MAX_ALPHABET) return undefined
        if (before) triples.add(before + previous + point)
        before = previous; previous = point
      }
    }
    const points = [...alphabet].sort()
    // Combining marks/ZWJ attach across raw boundaries; expansion characters
    // must fail even when their counterpart is absent from this alphabet.
    if (points.some(point => /\p{M}|\u200d/u.test(point)
      || [...point.normalize('NFKD')].filter(value => !/\p{M}/u.test(value)).length > 1
      || [...point.toUpperCase()].length > 1)) return undefined
    if (points.length <= 5) for (const a of points) for (const b of points) for (const c of points) triples.add(a + b + c)
    const key = JSON.stringify([points, [...triples].sort()])
    let cache = probeCaches.get(this.config)
    if (!cache) {
      cache = { comparisons: new Map(), profiles: new Map() }
      retain(probeCaches, this.config, cache, 8)
    }
    if (cache.profiles.has(key)) {
      budget.charge(1)
      this.certification.cacheHits++
      this.certification.chargedWork++
      const found = cache.profiles.get(key)
      this.certifiedAsciiProfile = Boolean(found)
      return found ?? undefined
    }
    const save = (profile: AlphabetProfile | null) => {
      // Entry counts alone would retain megabytes of unique triple keys.
      if (key.length <= MAX_PROFILE_KEY) retain(cache.profiles, key, profile, MAX_PROFILES)
    }
    let probeWork = 0
    const compare = (a: string, b: string): number => {
      const probeKey = JSON.stringify([a, b]), found = cache.comparisons.get(probeKey)
      const amount = found === undefined ? 1 + a.length + b.length : 1
      if (probeWork + amount > MAX_PROBE_WORK) throw new InconclusiveProbe()
      budget.charge(amount)
      probeWork += amount
      this.certification.chargedWork += amount
      if (found !== undefined) { this.certification.cacheHits++; return found }
      const value = this.ownerCompare(a, b)
      this.certification.comparisons++
      retain(cache.comparisons, probeKey, value, MAX_PROBE_CACHE)
      return value
    }
    try {
      const classes = new Map<string, number>()
      const sorted = points.slice().sort(compare)
      let id = 0, last: string | undefined
      for (const point of sorted) {
        if (compare(point, '') === 0) classes.set(point, 0)
        else {
          if (last === undefined || compare(last, point) !== 0) id++
          classes.set(point, id); last = point
        }
      }
      const representatives = new Map<number, string>()
      for (const point of points) { const id = classes.get(point)!; if (id && !representatives.has(id)) representatives.set(id, point) }
      const signature = (value: string) => [...value].map(point => classes.get(point)!).filter(id => id !== 0).join(',')
      const candidates = new Set(['', ...points, ...triples])
      for (const a of points) for (const b of points) candidates.add(a + b)
      // A pair/triple equality must coincide exactly with its class sequence.
      // Sorting makes every equality collision adjacent, without quadratic
      // comparison of the already quadratic pair domain.
      for (const candidate of candidates) {
        const canonical = [...candidate].map(point => representatives.get(classes.get(point)!) ?? '').join('')
        if (compare(candidate, canonical) !== 0) {
          save(null)
          return undefined
        }
      }
      const ordered = [...candidates].sort(compare)
      for (let i = 1; i < ordered.length; i++) {
        if ((compare(ordered[i - 1], ordered[i]) === 0) !== (signature(ordered[i - 1]) === signature(ordered[i]))) {
          save(null)
          return undefined
        }
      }
      const result = { classes }
      save(result)
      this.certifiedAsciiProfile = true
      return result
    } catch (error) {
      if (!(error instanceof InconclusiveProbe)) throw error
      save(null)
      return undefined
    }
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
    budget.charge(pattern.length)
    const runs = parsed ? parsed.filter((token): token is Extract<Token, { kind: 'literal' }> => token.kind === 'literal').map(token => token.text) : [pattern]
    return (text, budget) => {
      budget.charge(0)
      if (typeof text !== 'string') throw new RangeError('Invalid text condition')
      if (text.length > budget.maxTextLength) budget.text(text)
      if (!parsed && positive === 'eq') {
        this.certifiedAsciiProfile = false
        this.certification.comparisons = this.certification.cacheHits = this.certification.chargedWork = 0
        // Full-string equality already needs just one actual comparison.
        budget.charge(1 + pattern.length + text.length)
        const matched = this.collator.compare(pattern, text) === 0
        return op === 'neq' ? !matched : matched
      }
      if (!runs.length) {
        // Literal-free wildcards use raw boundaries without a certificate.
        this.certifiedAsciiProfile = false
        this.certification.comparisons = this.certification.cacheHits = this.certification.chargedWork = 0
      }
      const profile = runs.length ? this.profile(runs, text, budget) : { classes: new Map<string, number>() }
      if (!profile) return this.evaluate(op, pattern, parsed, text, budget)
      budget.text(text)
      const raw = Array.from(text), weights = new Uint32Array(raw.length + 1), folded: number[] = []
      for (let i = 0; i < raw.length; i++) {
        const id = profile.classes.get(raw[i])!
        if (id) folded.push(id)
        weights[i + 1] = folded.length
      }
      const literals = new Map<string, LiteralSearch>()
      for (const run of runs) if (!literals.has(run)) {
        const ids = Array.from(run, point => profile.classes.get(point)!).filter(id => id !== 0)
        literals.set(run, prepareLiteral(ids, budget))
      }
      const negative = op === 'neq' || op === 'not-contains'
      let matched: boolean
      if (!parsed) {
        const literal = literals.get(pattern)!, needle = literal.ids
        budget.charge(needle.length + 1)
        if (positive === 'eq') matched = folded.length === needle.length && needle.every((id, index) => id === folded[index])
        else if (positive === 'begins') matched = needle.length <= folded.length && needle.every((id, index) => id === folded[index])
        else if (positive === 'ends') matched = needle.length <= folded.length && needle.every((id, index) => id === folded[folded.length - needle.length + index])
        else {
          const hits = literalOccurrences(folded, literal, budget)
          budget.charge(hits.length)
          matched = hits.some(value => value === 1)
        }
      } else {
        let row = new Uint8Array(raw.length + 1)
        row[0] = 1
        const occurrences = new Map<string, Uint8Array>()
        for (const token of parsed) {
          budget.charge(raw.length + 1)
          const next = new Uint8Array(raw.length + 1)
          if (token.kind === 'star') {
            next[0] = row[0]
            for (let end = 1; end <= raw.length; end++) next[end] = row[end] || next[end - 1]
          } else if (token.kind === 'one') {
            for (let start = 0; start < raw.length; start++) if (row[start]) next[start + 1] = 1
          } else {
            const literal = literals.get(token.text)!
            if (!literal.ids.length) {
              budget.charge(raw.length + 1)
              let reachable = false
              for (let end = 0; end <= raw.length; end++) {
                if (end === 0 || weights[end] !== weights[end - 1]) reachable = false
                reachable ||= Boolean(row[end])
                next[end] = reachable ? 1 : 0
              }
            } else {
              let hits = occurrences.get(token.text)
              if (!hits) { hits = literalOccurrences(folded, literal, budget); occurrences.set(token.text, hits) }
              const targets = new Uint8Array(folded.length + 1)
              budget.charge((raw.length + 1) * 2)
              for (let start = 0; start <= raw.length; start++) {
                const target = weights[start] + literal.ids.length
                if (row[start] && hits[weights[start]] && target <= folded.length) targets[target] = 1
              }
              for (let end = 0; end <= raw.length; end++) next[end] = targets[weights[end]]
            }
          }
          row = next
        }
        matched = Boolean(row[raw.length])
      }
      return negative ? !matched : matched
    }
  }
  protected evaluate(op: TextOp, pattern: string, parsed: readonly Token[] | undefined, text: string, budget: WorkBudget): boolean {
    if (typeof text !== 'string') throw new RangeError('Invalid text condition')
    budget.charge(0)
    if (text.length > budget.maxTextLength) budget.text(text)
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
    budget.charge(text.length)
    if (!parsed) {
      // A positive actual slice is a witness regardless of locale, expansions
      // or context. UTF-16 search must not introduce a boundary inside a pair.
      const boundary = (at: number) => at === 0 || at === text.length
        || !(text.charCodeAt(at - 1) >= 0xd800 && text.charCodeAt(at - 1) <= 0xdbff
          && text.charCodeAt(at) >= 0xdc00 && text.charCodeAt(at) <= 0xdfff)
      let at = positive === 'begins' ? (text.startsWith(pattern) ? 0 : -1)
        : positive === 'ends' ? (text.endsWith(pattern) ? text.length - pattern.length : -1) : text.indexOf(pattern)
      while (at >= 0) {
        budget.charge(1 + pattern.length * 2)
        if (boundary(at) && boundary(at + pattern.length)
          && this.collator.compare(pattern, text.slice(at, at + pattern.length)) === 0) {
          budget.text(text)
          return !negative
        }
        if (positive !== 'contains') break
        at = text.indexOf(pattern, at + 1)
      }
      if (/^[\x20-\x7e]*$/.test(text) && /^[\x20-\x7e]*$/.test(pattern)) {
        const folded = text.toLowerCase(), needle = pattern.toLowerCase()
        const at = positive === 'begins' ? (folded.startsWith(needle) ? 0 : -1)
          : positive === 'ends' ? (folded.endsWith(needle) ? text.length - pattern.length : -1) : folded.indexOf(needle)
        if (at >= 0) {
          budget.text(text)
          budget.charge(1 + pattern.length * 2)
          if (this.collator.compare(pattern, text.slice(at, at + pattern.length)) === 0) return !negative
        }
      }
    }
    if (parsed || positive === 'contains') {
      // For every phase and element length a raw periodic string has exactly
      // the same UTF-16 slice. Cache only these byte-identical comparisons;
      // periods count code points so emoji and lone surrogates stay separate.
      const elements = Array.from(text)
      for (let candidate = 1; candidate <= Math.min(16, elements.length); candidate++) {
        let repeats = true
        for (let i = candidate; i < elements.length; i++) {
          budget.charge(1)
          if (elements[i] !== elements[i % candidate]) { repeats = false; break }
        }
        if (repeats) { period = candidate; break }
      }
    }
    return super.evaluate(op, pattern, parsed, text, budget, period)
  }
}
