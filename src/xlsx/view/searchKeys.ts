import { ValueIndexError } from './valueRuns'

/** Optional precomputed search keys for value-list search (spec §7.4, user
 * decision 2026-10-11). The key function is injected; S8.5 supplies the
 * canonical-form key once the text-matching amendment lands. Until then the
 * default strategy returns null for every value, so every search uses the
 * exact S8.1 matcher. A null key always means "use the exact matcher". */
export type SearchKeyOf = (text: string) => string | null
export interface SearchKeyStrategy {
  /** Version identity of the key function; persisted keys are keyed by it. */
  readonly id: string
  readonly searchKeyOf: SearchKeyOf
  /** 'equivalent': for non-null keys, contains(query, text) holds exactly when
   * key(text) includes key(query). 'candidate': inclusion is only necessary,
   * so every key hit is confirmed by the exact matcher. */
  readonly exactness: 'equivalent' | 'candidate'
}
export const EXACT_SEARCH: SearchKeyStrategy = Object.freeze({ id: 'exact', searchKeyOf: () => null, exactness: 'candidate' as const })
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
  readonly texts: readonly string[]
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
export class MemorySearchKeyStore implements SearchKeyStore {
  readonly limitBytes: number
  usedBytes = 0
  saves = 0
  loads = 0
  constructor(options: { limitBytes?: number } = {}) { this.limitBytes = options.limitBytes ?? Infinity }
  async load(_owner: SearchKeyOwner): Promise<SearchKeyRun | null> { this.loads++; return null }
  async save(_owner: SearchKeyOwner, _run: SearchKeyRun): Promise<void> {}
  async delete(_owner: SearchKeyOwner): Promise<void> {}
}
void ValueIndexError
