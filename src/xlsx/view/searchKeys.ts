import { ValueIndexError } from './valueRuns'

/** Optional precomputed search keys for value-list search (spec §7.4, user
 * decision 2026-10-11). The key function is injected; S8.5 supplies the
 * canonical-form key once the text-matching amendment lands. Until then the
 * default strategy returns null for every value, so every search uses the
 * exact S8.1 matcher. A null key always means "use the exact matcher". */
export type SearchKeyOf = (text: string) => string | null
export type FoldVariant = 'default' | 'turkic'
export interface SearchKeyStrategySpec {
  canonicalizerVersion: string
  unicodeVersion: string
  foldVariant: FoldVariant
  searchKeyOf: SearchKeyOf
  matchKey?: (valueKey: string, queryKey: string) => boolean
}
export interface SearchKeyStrategy {
  readonly id: string
  readonly foldVariant: FoldVariant
  readonly searchKeyOf: SearchKeyOf
  readonly matchKey: (valueKey: string, queryKey: string) => boolean
}
export function foldVariantOf(_locale: string): FoldVariant { return 'default' }
export function createSearchKeyStrategy(spec: SearchKeyStrategySpec): SearchKeyStrategy {
  return { id: spec.canonicalizerVersion, foldVariant: spec.foldVariant, searchKeyOf: spec.searchKeyOf, matchKey: spec.matchKey ?? ((a, b) => a.includes(b)) }
}
export const EXACT_SEARCH: SearchKeyStrategy = Object.freeze({ id: 'exact', foldVariant: 'default' as const, searchKeyOf: () => null, matchKey: () => true })
/** Keys are valid only for this exact value domain and key function. */
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
/** Keys aligned with the index's raw text dictionary; texts are stored so a
 * loader can verify alignment exactly instead of trusting ordinals. */
export interface SearchKeyRun {
  readonly count: number
  readonly hash: string
  readonly keys: readonly (string | null)[]
}
export interface SearchKeyStore {
  /** null when absent; typed storage-error on failure. */
  load(owner: SearchKeyOwner, signal?: AbortSignal): Promise<SearchKeyRun | null>
  /** Typed quota (resource storage) or storage-error on failure. */
  save(owner: SearchKeyOwner, run: SearchKeyRun, signal?: AbortSignal): Promise<void>
  delete(owner: SearchKeyOwner): Promise<void>
}
export const searchKeyOwnerId = (owner: SearchKeyOwner): string => `${owner.scope}\u0000${owner.strategyId}`
const runBytes = (run: SearchKeyRun): number => {
  let bytes = 64 + run.hash.length * 2
  for (const key of run.keys) bytes += key === null ? 8 : key.length * 2 + 40
  return bytes
}
function validateRun(run: SearchKeyRun): void {
  if (!run || typeof run.hash !== 'string' || !Array.isArray(run.keys) || run.count !== run.keys.length
    || run.keys.some(key => key !== null && typeof key !== 'string')) {
    throw new ValueIndexError('invalid-rule', 'Malformed search key run')
  }
}
/** Reference backend for tests and small sessions; the durable backend is a
 * later decision (a persistence spike is pending). */
export class MemorySearchKeyStore implements SearchKeyStore {
  readonly limitBytes: number
  usedBytes = 0
  saves = 0
  loads = 0
  private readonly runs = new Map<string, { run: SearchKeyRun; bytes: number }>()
  constructor(options: { limitBytes?: number; account?: unknown } = {}) {
    this.limitBytes = options.limitBytes ?? Infinity
    if (!(this.limitBytes >= 0)) throw new RangeError('Invalid search key store limit')
  }
  async load(owner: SearchKeyOwner, signal?: AbortSignal): Promise<SearchKeyRun | null> {
    if (signal?.aborted) throw new ValueIndexError('cancelled', 'Search key load cancelled')
    this.loads++
    const stored = this.runs.get(searchKeyOwnerId(owner))
    return stored ? { count: stored.run.count, hash: stored.run.hash, keys: stored.run.keys.slice() } : null
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
    this.runs.set(id, { run: Object.freeze({ count: run.count, hash: run.hash, keys: Object.freeze(run.keys.slice()) }), bytes })
    this.usedBytes += bytes - previous
    this.saves++
  }
  async delete(owner: SearchKeyOwner): Promise<void> {
    const id = searchKeyOwnerId(owner), stored = this.runs.get(id)
    if (!stored) return
    this.runs.delete(id)
    this.usedBytes -= stored.bytes
  }
}
