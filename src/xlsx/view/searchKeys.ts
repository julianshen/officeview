import { sha256HexJs } from '../../debug/harness/sha256'
import { TextFilterMatcher } from './textFilter'
import { ValueIndexError, type LaneAccount } from './valueRuns'

/** Optional precomputed search keys for value-list search (spec §7.4; value-
 * search spike decisions 2026-10-11). The key function is injected; S8.5
 * supplies the canonical-form key. Until then the default strategy returns
 * null for every value, so every search uses the exact S8.1 matcher.
 *
 * Keys are a necessary-condition filter only: a key miss proves "no match"
 * and every key hit is confirmed by the exact matcher. There is deliberately
 * no "equivalent" mode until S8.5 provides boundary-aware key matchers (under
 * canonical matching key('é') includes 'e' although contains('é', 'e') fails). */
export type SearchKeyOf = (text: string) => string | null
export type FoldVariant = 'default' | 'turkic'
export interface SearchKeyStrategySpec {
  canonicalizerVersion: string
  unicodeVersion: string
  foldVariant: FoldVariant
  /** null means "no key; use the exact matcher". */
  searchKeyOf: SearchKeyOf
  /** Necessary condition for an exact contains match; false is decisive.
   * Default: valueKey.includes(queryKey). */
  matchKey?: (valueKey: string, queryKey: string) => boolean
}
/** Frozen, versioned and certified by createSearchKeyStrategy only. */
export interface SearchKeyStrategy {
  readonly id: string
  readonly foldVariant: FoldVariant
  readonly searchKeyOf: SearchKeyOf
  readonly matchKey: (valueKey: string, queryKey: string) => boolean
}
const certified = new WeakSet<SearchKeyStrategy>()
export function isCertifiedSearchKeyStrategy(strategy: unknown): strategy is SearchKeyStrategy {
  return typeof strategy === 'object' && strategy !== null && certified.has(strategy as SearchKeyStrategy)
}
/** The fold variant a query locale needs; keys of another variant are unused. */
export function foldVariantOf(locale: string): FoldVariant {
  const language = String(locale).split(/[-_]/)[0].toLowerCase()
  return language === 'tr' || language === 'az' ? 'turkic' : 'default'
}
const PROBE_VALUES = [
  'alpha', 'ALPHA', 'beta', 'Straße', 'STRASSE', 'strasse', 'é', 'é', 'café', 'CAFE', 'ﬁle', 'file', 'Ångström', 'aa', 'å',
  '😀x', '\ud800a', 'İstanbul', 'istanbul', 'ISTANBUL', 'item-12345', '12', '012', 'a b', 'A-B', '#N/A', 'n/a', 'l·l', 'x​y',
]
function probeQueries(): string[] {
  const queries = new Set<string>()
  for (const value of PROBE_VALUES) {
    const points = Array.from(value)
    queries.add(value).add(value.toUpperCase()).add(value.toLowerCase())
    for (let length = 1; length <= 3; length++) queries.add(points.slice(0, length).join('')).add(points.slice(-length).join(''))
  }
  return [...queries]
}
/** Rejects a key function whose decisive misses contradict the exact matcher
 * on a fixed hostile probe set. A guard, not a proof: S8.5 must run the full
 * differential suite for the real canonical key. */
function probe(strategy: SearchKeyStrategy): void {
  const matcher = new TextFilterMatcher(strategy.foldVariant === 'turkic' ? 'tr-TR' : 'en-US')
  const key = (text: string): string | null => {
    const value = strategy.searchKeyOf(text)
    if (value !== null && typeof value !== 'string') throw new ValueIndexError('invalid-rule', 'Search key function returned a non-string key')
    return value
  }
  const valueKeys = PROBE_VALUES.map(key)
  for (const query of probeQueries()) {
    const queryKey = key(query)
    if (queryKey === null) continue
    for (let at = 0; at < PROBE_VALUES.length; at++) {
      const valueKey = valueKeys[at]
      if (valueKey === null || strategy.matchKey(valueKey, queryKey)) continue
      if (matcher.match('contains', query, false, PROBE_VALUES[at])) {
        throw new ValueIndexError('invalid-rule', `Search key strategy ${strategy.id} misses an exact match (${JSON.stringify(query)} in ${JSON.stringify(PROBE_VALUES[at])})`)
      }
    }
  }
}
export function createSearchKeyStrategy(spec: SearchKeyStrategySpec): SearchKeyStrategy {
  const { canonicalizerVersion, unicodeVersion, foldVariant, searchKeyOf } = spec ?? {} as SearchKeyStrategySpec
  if (typeof canonicalizerVersion !== 'string' || !canonicalizerVersion || typeof unicodeVersion !== 'string' || !unicodeVersion
    || (foldVariant !== 'default' && foldVariant !== 'turkic') || typeof searchKeyOf !== 'function'
    || (spec.matchKey !== undefined && typeof spec.matchKey !== 'function')) {
    throw new ValueIndexError('invalid-rule', 'Invalid search key strategy specification')
  }
  const matchKey = spec.matchKey ?? ((valueKey: string, queryKey: string) => valueKey.includes(queryKey))
  const strategy: SearchKeyStrategy = Object.freeze({ id: `${canonicalizerVersion}|unicode-${unicodeVersion}|${foldVariant}`, foldVariant, searchKeyOf, matchKey })
  probe(strategy)
  certified.add(strategy)
  return strategy
}
export const EXACT_SEARCH: SearchKeyStrategy = Object.freeze({ id: 'exact', foldVariant: 'default' as const, searchKeyOf: () => null, matchKey: () => true })
certified.add(EXACT_SEARCH)

/** Keys are valid only for this exact value domain and key strategy. */
export interface SearchKeyOwner {
  docId: string
  sheetId: string
  valueGen: number
  col: number
  classificationGen: number
  /** Full value scope (range/column/classification), see valueScope. */
  scope: string
  strategyId: string
}
/** Keys aligned with the index's raw text dictionary; hash is the content
 * hash of those texts (textChunkHash/combineTextHash) for exact alignment. */
export interface SearchKeyRun {
  readonly count: number
  readonly hash: string
  readonly keys: readonly (string | null)[]
}
export const SEARCH_HASH_CHUNK = 4096
/** Upper bound of one hash scratch chunk (bytes of length-prefixed UTF-16). */
export const SEARCH_HASH_CHUNK_BYTES = 1024 * 1024
/** SHA-256 of length-prefixed UTF-16 code units; lone surrogates are hashed
 * losslessly (no text codec). */
export function textChunkHash(texts: readonly string[], start: number, end: number): string {
  let size = 0
  for (let i = start; i < end; i++) size += 4 + texts[i].length * 2
  const bytes = new Uint8Array(size), view = new DataView(bytes.buffer)
  let at = 0
  for (let i = start; i < end; i++) {
    const text = texts[i]
    view.setUint32(at, text.length, true); at += 4
    for (let unit = 0; unit < text.length; unit++) { view.setUint16(at, text.charCodeAt(unit), true); at += 2 }
  }
  return sha256HexJs(bytes)
}
export const combineTextHash = (count: number, chunks: readonly string[]): string => sha256HexJs(`${count}:${chunks.join(',')}`)
export interface SearchKeyStore {
  /** null when absent; typed storage-error on failure. */
  load(owner: SearchKeyOwner, signal?: AbortSignal): Promise<SearchKeyRun | null>
  /** Typed quota or storage-error on failure. */
  save(owner: SearchKeyOwner, run: SearchKeyRun, signal?: AbortSignal): Promise<void>
  delete(owner: SearchKeyOwner): Promise<void>
  /** RAM-backed stores charge their retained bytes to the owning cache pool. */
  bindAccount?(account: LaneAccount): void
  /** Frees at least bytes of the oldest entries under pool pressure; returns bytes freed. */
  shed?(bytes: number): number
}
export const searchKeyOwnerId = (owner: SearchKeyOwner): string => `${owner.scope}\u0000${owner.strategyId}`
const runBytes = (run: SearchKeyRun): number => {
  let bytes = 64 + run.hash.length * 2
  for (const key of run.keys) bytes += key === null ? 8 : key.length * 2 + 48
  return bytes
}
function validateRun(run: SearchKeyRun): void {
  if (!run || typeof run.hash !== 'string' || !Array.isArray(run.keys) || run.count !== run.keys.length
    || run.keys.some(key => key !== null && typeof key !== 'string')) {
    throw new ValueIndexError('invalid-rule', 'Malformed search key run')
  }
}
/** The v1 backend: RAM only, charged to the cache pool it is bound to, with
 * oldest-first shedding. No durable persistence in v1. */
export class MemorySearchKeyStore implements SearchKeyStore {
  readonly limitBytes: number
  usedBytes = 0
  saves = 0
  loads = 0
  private account: LaneAccount | undefined
  private readonly runs = new Map<string, { run: SearchKeyRun; bytes: number }>()
  constructor(options: { limitBytes?: number; account?: LaneAccount } = {}) {
    this.limitBytes = options.limitBytes ?? Infinity
    if (!(this.limitBytes >= 0)) throw new RangeError('Invalid search key store limit')
    this.account = options.account
  }
  bindAccount(account: LaneAccount): void {
    if (this.account) return
    account.charge(this.usedBytes)
    this.account = account
  }
  async load(owner: SearchKeyOwner, signal?: AbortSignal): Promise<SearchKeyRun | null> {
    if (signal?.aborted) throw new ValueIndexError('cancelled', 'Search key load cancelled')
    this.loads++
    const id = searchKeyOwnerId(owner), stored = this.runs.get(id)
    if (!stored) return null
    this.runs.delete(id)
    this.runs.set(id, stored)
    return { count: stored.run.count, hash: stored.run.hash, keys: stored.run.keys.slice() }
  }
  async save(owner: SearchKeyOwner, run: SearchKeyRun, signal?: AbortSignal): Promise<void> {
    if (signal?.aborted) throw new ValueIndexError('cancelled', 'Search key save cancelled')
    validateRun(run)
    const id = searchKeyOwnerId(owner), previous = this.runs.get(id)?.bytes ?? 0, bytes = runBytes(run)
    if (this.usedBytes - previous + bytes > this.limitBytes) {
      throw new ValueIndexError('quota', 'Search key storage quota exceeded', {
        resource: 'storage', requiredBytes: this.usedBytes - previous + bytes, limitBytes: this.limitBytes, retryable: false,
      })
    }
    this.account?.charge(bytes)
    const copy = Object.freeze({ count: run.count, hash: run.hash, keys: Object.freeze(run.keys.slice()) })
    if (previous) await this.delete(owner)
    this.runs.set(id, { run: copy, bytes })
    this.usedBytes += bytes
    this.saves++
  }
  async delete(owner: SearchKeyOwner): Promise<void> { this.remove(searchKeyOwnerId(owner)) }
  private remove(id: string): number {
    const stored = this.runs.get(id)
    if (!stored) return 0
    this.runs.delete(id)
    this.usedBytes -= stored.bytes
    this.account?.release(stored.bytes)
    return stored.bytes
  }
  shed(bytes: number): number {
    let freed = 0
    for (const id of [...this.runs.keys()]) {
      if (freed >= bytes) break
      freed += this.remove(id)
    }
    return freed
  }
}
