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
 * the code points actually observed, not a universal theorem about Intl.
 * Every code point gets a class (0 = equal to the empty string, i.e. completely
 * ignorable; otherwise equal code points share an id), cached per collator
 * across cells. Cost is bounded by the text, not by alphabet size: only raw
 * windows of up to three ADJACENT code points actually present in the text or
 * a literal are probed, and each must equal the concatenation of its class
 * representatives with no cross-signature equality collision among the
 * windows. Canonical/compatibility expansions and combining-mark contexts are
 * inconclusive, and so is any run of two or more adjacent ignorables: native
 * context rules (e.g. Catalan l + U+00B7) look through a single ignorable but
 * not through a run, and a window of three cannot see across a run. Only
 * certified inputs admit class-ID KMP/DP; everything else retains the exact
 * raw-slice reference and its explicit budget. An observed expansion factor
 * alone cannot bound every future candidate end, and ignored plateaus have
 * unbounded raw length, so no windows are inferred for failed profiles.
 */
const COLLATOR_CONSTRUCTOR = Intl.Collator
const COLLATOR_COMPARE_GETTER = Object.getOwnPropertyDescriptor(Intl.Collator.prototype, 'compare')?.get
const UNSAFE = -1
const MAX_CLASS_TABLE = 32768
const MAX_PROBE_WORK = 90_000
/** Window and signature work grows with the text, so the probe cap does too. */
const PROBE_WORK_PER_POINT = 400
const MAX_PROBE_CACHE = 65536
interface ClassTable {
  /** Point to class id: 0 ignorable, UNSAFE for expansion/mark contexts. */
  readonly ids: Map<string, number>
  /** Representative per class id (index 0 is the empty string). */
  readonly reps: string[]
  /** Class ids ordered by collation, for binary-search insertion. */
  readonly order: number[]
}
interface AlphabetProfile { readonly classes: ReadonlyMap<string, number> }
interface ProbeCache {
  readonly comparisons: Map<string, number>
  table: ClassTable
}
const probeCaches = new Map<string, ProbeCache>()
class InconclusiveProbe extends Error {}
function retain<K, V>(map: Map<K, V>, key: K, value: V, maximum: number): void {
  if (map.size >= maximum) map.delete(map.keys().next().value!)
  map.set(key, value)
}
const newTable = (): ClassTable => ({ ids: new Map(), reps: [''], order: [] })
function unsafePoint(point: string): boolean {
  // Combining marks/ZWJ attach across raw boundaries; expansion characters
  // must fail even when their counterpart is absent from this text.
  return /\p{M}|\u200d/u.test(point)
    || [...point.normalize('NFKD')].filter(value => !/\p{M}/u.test(value)).length > 1
    || [...point.toUpperCase()].length > 1
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
    const strings = [...runs, text], arrays: string[][] = []
    let total = 0
    for (const value of strings) {
      const points = Array.from(value)
      budget.charge(points.length)
      total += points.length
      arrays.push(points)
    }
    let cache = probeCaches.get(this.config)
    if (!cache) {
      cache = { comparisons: new Map(), table: newTable() }
      retain(probeCaches, this.config, cache, 8)
    }
    if (cache.table.ids.size > MAX_CLASS_TABLE) cache.table = newTable()
    const table = cache.table, limit = MAX_PROBE_WORK + PROBE_WORK_PER_POINT * total
    let probeWork = 0
    const compare = (a: string, b: string): number => {
      const probeKey = `${a.length}|${a}${b}`, found = cache.comparisons.get(probeKey)
      const amount = found === undefined ? 1 + a.length + b.length : 1
      if (probeWork + amount > limit) throw new InconclusiveProbe()
      budget.charge(amount)
      probeWork += amount
      this.certification.chargedWork += amount
      if (found !== undefined) { this.certification.cacheHits++; return found }
      const value = this.ownerCompare(a, b)
      this.certification.comparisons++
      retain(cache.comparisons, probeKey, value, MAX_PROBE_CACHE)
      return value
    }
    const classify = (point: string): number => {
      const known = table.ids.get(point)
      if (known !== undefined) return known
      let id: number
      if (unsafePoint(point)) id = UNSAFE
      else if (compare(point, '') === 0) id = 0
      else {
        let low = 0, high = table.order.length, found = 0
        while (low < high && !found) {
          const middle = Math.floor((low + high) / 2), order = compare(point, table.reps[table.order[middle]])
          if (order === 0) found = table.order[middle]
          else if (order < 0) high = middle
          else low = middle + 1
        }
        if (found) id = found
        else {
          id = table.reps.length
          table.reps.push(point)
          table.order.splice(low, 0, id)
        }
      }
      table.ids.set(point, id)
      return id
    }
    try {
      // Pass 1: classes, expansion/mark contexts, and ignorable runs. A run of
      // two or more ignorables defeats the three-point windows below.
      const identifiers: number[][] = []
      for (const points of arrays) {
        const ids: number[] = []
        for (let i = 0; i < points.length; i++) {
          const id = classify(points[i])
          if (id === UNSAFE || (id === 0 && i > 0 && ids[i - 1] === 0)) return undefined
          ids.push(id)
        }
        identifiers.push(ids)
      }
      if (table.ids.size > MAX_CLASS_TABLE * 2) return undefined
      // Pass 2: windows of up to three adjacent code points that actually occur.
      const windows = new Map<string, string>(), signatures = new Map<string, string>()
      for (let which = 0; which < arrays.length; which++) {
        const points = arrays[which], ids = identifiers[which]
        for (let i = 0; i < points.length; i++) {
          let window = '', signature = '', canonical = ''
          for (let length = 1; length <= 3 && i + length <= points.length; length++) {
            budget.charge(1)
            window += points[i + length - 1]
            const id = ids[i + length - 1]
            if (id) { signature += ',' + id; canonical += table.reps[id] }
            if (!windows.has(window)) windows.set(window, canonical)
            if (!signatures.has(signature)) signatures.set(signature, canonical)
          }
        }
      }
      // A window equality must coincide exactly with its class sequence.
      for (const [window, canonical] of windows) {
        if (window !== canonical && compare(window, canonical) !== 0) return undefined
      }
      // Distinct class sequences must never collate equal across lengths
      // (e.g. Danish aa/å). Sorting makes every collision adjacent.
      const ordered = [...signatures.values()].sort(compare)
      for (let i = 1; i < ordered.length; i++) if (compare(ordered[i - 1], ordered[i]) === 0) return undefined
      this.certifiedAsciiProfile = true
      return { classes: table.ids }
    } catch (error) {
      if (!(error instanceof InconclusiveProbe)) throw error
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
