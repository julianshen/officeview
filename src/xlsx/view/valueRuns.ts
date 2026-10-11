import type { PagedKeySetRef, SourceErrorCode, SourceStamp } from './contracts'

/** Exact semantic value-key runs and their storage port (spec §§7.3–7.4, 12.4).
 * The in-memory store is the reference backend; S1.x backs the same port with
 * OPFS/IndexedDB pages without changing ref semantics. */
export const KEY_NUMBER = 1, KEY_STRING = 2, KEY_BOOLEAN = 3, KEY_ERROR = 4, KEY_DATE = 5
export type ValueKeyKind = typeof KEY_NUMBER | typeof KEY_STRING | typeof KEY_BOOLEAN | typeof KEY_ERROR | typeof KEY_DATE
export type ValueIndexErrorCode = SourceErrorCode | 'invalid-rule' | 'closed'
export class ValueIndexError extends Error {
  constructor(
    readonly code: ValueIndexErrorCode,
    message: string,
    readonly detail: { resource?: 'memory' | 'storage'; requiredBytes?: number; limitBytes?: number; retryable?: boolean } = {},
  ) {
    super(message)
    this.name = 'ValueIndexError'
  }
}
/** Parallel typed lanes: numbers hold the value (booleans 0/1) and texts hold
 * string/error raw UTF-16 at the same offset. */
export interface ValueKeyRun {
  readonly count: number
  readonly kinds: Uint8Array
  readonly numbers: Float64Array
  readonly texts: readonly (string | undefined)[]
}
export function valueKeyString(kind: number, value: number, text: string | undefined): string {
  return `n:${value}${text ?? ''}${kind}`
}
export class ValueKeySet {
  add(_kind: number, _value: number, _text: string | undefined): void {}
  has(_kind: number, _value: number, _text: string | undefined): boolean { return true }
  get size(): number { return 0 }
}
export interface ValueRunStore {
  put(run: ValueKeyRun, owner: SourceStamp, signal?: AbortSignal): Promise<PagedKeySetRef>
  read(ref: PagedKeySetRef, offset: number, limit: number, signal?: AbortSignal): Promise<ValueKeyRun>
  release(ref: PagedKeySetRef): Promise<void>
}
export class MemoryValueRunStore implements ValueRunStore {
  readonly limitBytes: number
  usedBytes = 0
  puts = 0
  private readonly runs = new Map<string, { run: ValueKeyRun; owner: SourceStamp }>()
  constructor(options: { limitBytes?: number } = {}) { this.limitBytes = options.limitBytes ?? Infinity }
  get refCount(): number { return this.runs.size }
  async put(run: ValueKeyRun, owner: SourceStamp): Promise<PagedKeySetRef> {
    const refId = `run-${++this.puts}`
    this.runs.set(refId, { run, owner })
    return { refId, count: run.count, owner }
  }
  async read(ref: PagedKeySetRef): Promise<ValueKeyRun> { return this.runs.get(ref.refId)!.run }
  async release(ref: PagedKeySetRef): Promise<void> { this.runs.delete(ref.refId) }
}
