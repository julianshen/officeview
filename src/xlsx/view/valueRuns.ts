import { allocationTargets, type AllocationContext } from './allocationAccount'
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
export const sameStamp = (a: SourceStamp, b: SourceStamp): boolean =>
  a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen
/** Parallel typed lanes: numbers hold the value (booleans 0/1) and texts hold
 * string/error raw UTF-16 at the same offset. */
export interface ValueKeyRun {
  readonly count: number
  readonly kinds: Uint8Array
  readonly numbers: Float64Array
  readonly texts: readonly (string | undefined)[]
}
const textual = (kind: number): boolean => kind === KEY_STRING || kind === KEY_ERROR
/** Same encoding as filter.ts rawValueKey; -0 joins 0 and strings keep raw code units. */
export function valueKeyString(kind: number, value: number, text: string | undefined): string {
  switch (kind) {
    case KEY_NUMBER: return `n:${value === 0 ? 0 : value}`
    case KEY_DATE: return `d:${value === 0 ? 0 : value}`
    case KEY_BOOLEAN: return value ? 'b:1' : 'b:0'
    case KEY_STRING: return `s:${text}`
    case KEY_ERROR: return `e:${text}`
    default: throw new ValueIndexError('invalid-rule', 'Unknown value key kind')
  }
}
/** Approximate retained bytes of one key-run element, for store admission. */
export const runElementBytes = (kind: number, text: string | undefined): number =>
  9 + (textual(kind) ? (text?.length ?? 0) * 2 + 40 : 0)

const hashView = new Float64Array(1), hashWords = new Uint32Array(hashView.buffer)
/** Mixes the IEEE-754 words of a finite, -0-normalized number. */
function hashNumber(value: number): number {
  hashView[0] = value
  let hash = Math.imul(hashWords[0] ^ 0x9e3779b9, 0x85ebca6b) ^ Math.imul(hashWords[1] ^ 0x7f4a7c15, 0xc2b2ae35)
  hash ^= hash >>> 15
  hash = Math.imul(hash, 0x27d4eb2f)
  return (hash ^ (hash >>> 13)) >>> 0
}
/** Receives every typed allocation so a pool can admit it before it exists. */
export interface LaneAccount {
  charge(bytes: number): void
  release(bytes: number): void
}
/** Open-addressing number → id table in typed lanes; a Map would box doubles. */
export class NumberTable {
  private keys: Float64Array
  private ids: Int32Array
  private mask: number
  size = 0
  constructor(private readonly account?: LaneAccount, capacity = 64) {
    account?.charge(capacity * 12)
    try {
      this.keys = new Float64Array(capacity)
      this.ids = new Int32Array(capacity).fill(-1)
    } catch (error) { account?.release(capacity * 12); throw error }
    this.mask = capacity - 1
  }
  get bytes(): number { return this.keys.length * 12 }
  get(value: number): number {
    for (let slot = hashNumber(value) & this.mask; ; slot = (slot + 1) & this.mask) {
      const id = this.ids[slot]
      if (id < 0) return -1
      if (this.keys[slot] === value) return id
    }
  }
  /** The caller guarantees value is absent. */
  set(value: number, id: number): void {
    if ((this.size + 1) * 2 > this.keys.length) this.grow()
    let slot = hashNumber(value) & this.mask
    while (this.ids[slot] >= 0) slot = (slot + 1) & this.mask
    this.keys[slot] = value
    this.ids[slot] = id
    this.size++
  }
  private grow(): void {
    const capacity = this.keys.length * 2, oldBytes = this.bytes
    this.account?.charge(capacity * 12)
    const keys = this.keys, ids = this.ids
    let nextKeys: Float64Array, nextIds: Int32Array
    try { nextKeys = new Float64Array(capacity); nextIds = new Int32Array(capacity).fill(-1) }
    catch (error) { this.account?.release(capacity * 12); throw error }
    this.keys = nextKeys
    this.ids = nextIds
    this.mask = capacity - 1
    this.size = 0
    for (let slot = 0; slot < ids.length; slot++) if (ids[slot] >= 0) this.set(keys[slot], ids[slot])
    this.account?.release(oldBytes)
  }
}
/** Estimated retained bytes per text key (raw UTF-16, header, Set entry).
 * Numeric keys are charged as their actual typed hash-table allocations. */
export const keySetElementBytes = (kind: number, text: string | undefined): number =>
  textual(kind) ? (text?.length ?? 0) * 2 + 40 + 48 : 0
/** Exact key membership used by committed rules and draft bases. bytes is
 * the charged total, including hash-table capacity and growth. */
export class ValueKeySet {
  private readonly numbers: NumberTable
  private readonly dates: NumberTable
  private readonly booleans: Uint8Array
  private readonly strings: Set<string>
  private readonly errors: Set<string>
  private readonly ledger: LaneAccount
  size = 0
  bytes = 0
  constructor(account?: LaneAccount) {
    account?.charge(640)
    this.bytes = 640
    this.ledger = {
      charge: bytes => { account?.charge(bytes); this.bytes += bytes },
      release: bytes => { account?.release(bytes); this.bytes -= bytes },
    }
    try {
      this.strings = new Set<string>()
      this.errors = new Set<string>()
      this.ledger.charge(2)
      this.booleans = new Uint8Array(2)
      this.numbers = new NumberTable(this.ledger)
      this.dates = new NumberTable(this.ledger)
    } catch (error) { this.ledger.release(this.bytes); throw error }
  }
  add(kind: number, value: number, text: string | undefined): void {
    if (this.has(kind, value, text)) return
    const bytes = keySetElementBytes(kind, text)
    if (bytes) this.ledger.charge(bytes)
    switch (kind) {
      case KEY_NUMBER: this.numbers.set(value === 0 ? 0 : value, 0); break
      case KEY_DATE: this.dates.set(value === 0 ? 0 : value, 0); break
      case KEY_BOOLEAN: this.booleans[value ? 1 : 0] = 1; break
      case KEY_STRING: this.strings.add(text!); break
      case KEY_ERROR: this.errors.add(text!); break
    }
    this.size++
  }
  has(kind: number, value: number, text: string | undefined): boolean {
    switch (kind) {
      case KEY_NUMBER: return this.numbers.get(value === 0 ? 0 : value) >= 0
      case KEY_DATE: return this.dates.get(value === 0 ? 0 : value) >= 0
      case KEY_BOOLEAN: return this.booleans[value ? 1 : 0] === 1
      case KEY_STRING: return this.strings.has(text!)
      case KEY_ERROR: return this.errors.has(text!)
      default: throw new ValueIndexError('invalid-rule', 'Unknown value key kind')
    }
  }
}
export function validateRun(run: ValueKeyRun): void {
  if (!Number.isSafeInteger(run.count) || run.count < 0 || run.kinds.length < run.count || run.numbers.length < run.count) {
    throw new ValueIndexError('invalid-rule', 'Malformed value key run')
  }
  for (let i = 0; i < run.count; i++) {
    const kind = run.kinds[i]
    if (kind < KEY_NUMBER || kind > KEY_DATE || (textual(kind) ? typeof run.texts[i] !== 'string' : !Number.isFinite(run.numbers[i]))) {
      throw new ValueIndexError('invalid-rule', 'Malformed value key run element')
    }
  }
}
export interface ValueRunStore {
  /** Stores an immutable copy; refusal is a typed storage quota error. */
  /** scope names the exact value domain (range/column/classification, see
   * valueScope); key runs are meaningless in any other column. */
  put(run: ValueKeyRun, owner: SourceStamp, scope: string, signal?: AbortSignal): Promise<PagedKeySetRef>
  /** Reads [offset, offset+limit); another owner or scope is stale. */
  read(ref: PagedKeySetRef, offset: number, limit: number, scope: string, signal?: AbortSignal): Promise<ValueKeyRun>
  release(ref: PagedKeySetRef): Promise<void>
}
export const VALUE_RUN_PAGE = 65_536
/** Charges each loaded key to account (when given) before retaining it. */
export async function loadKeySet(store: ValueRunStore, ref: PagedKeySetRef, scope: string, signal?: AbortSignal, account?: LaneAccount): Promise<ValueKeySet> {
  const set = new ValueKeySet(account)
  let offset = 0
  while (offset < ref.count) {
    const run = await store.read(ref, offset, Math.min(VALUE_RUN_PAGE, ref.count - offset), scope, signal)
    if (signal?.aborted) throw new ValueIndexError('cancelled', 'Value membership cancelled')
    if (run.count === 0) throw new ValueIndexError('page-error', 'Value key run ended early')
    validateRun(run)
    for (let i = 0; i < run.count; i++) set.add(run.kinds[i], run.numbers[i], run.texts[i])
    offset += run.count
  }
  if (set.size !== ref.count) throw new ValueIndexError('page-error', 'Value key run contains duplicate keys')
  return set
}
interface StoredRun { run: ValueKeyRun; owner: SourceStamp; scope: string; bytes: number; readers: number; retired: boolean }
interface ReaderToken { stored: StoredRun; remaining: number }
export class MemoryValueRunStore implements ValueRunStore {
  readonly limitBytes: number
  usedBytes = 0
  puts = 0
  private readonly runs = new Map<string, StoredRun>()
  /** RAM-backed: retained bytes are charged to account (a worker pool) when given. */
  private readonly account: LaneAccount | undefined
  private readonly readerAllocation: AllocationContext | undefined
  private readonly readerWatcher: { register(target: object, token: ReaderToken): void } | undefined
  constructor(options: { limitBytes?: number; account?: LaneAccount; readerAllocation?: AllocationContext } = {}) {
    this.limitBytes = options.limitBytes ?? Infinity
    if (!(this.limitBytes >= 0)) throw new RangeError('Invalid value run store limit')
    this.account = options.account
    this.readerAllocation = options.readerAllocation
    const Registry = (globalThis as unknown as { FinalizationRegistry?: new (cleanup: (token: ReaderToken) => void) => { register(target: object, token: ReaderToken): void } }).FinalizationRegistry
    if (options.readerAllocation && Registry) this.readerWatcher = new Registry(token => {
      if (--token.remaining > 0) return
      token.stored.readers--
      if (token.stored.retired && token.stored.readers === 0) this.drop(token.stored)
    })
  }
  get refCount(): number { return this.runs.size }
  async put(run: ValueKeyRun, owner: SourceStamp, scope: string, signal?: AbortSignal): Promise<PagedKeySetRef> {
    if (signal?.aborted) throw new ValueIndexError('cancelled', 'Value run write cancelled')
    validateRun(run)
    if (typeof scope !== 'string') throw new ValueIndexError('invalid-rule', 'Value run scope is required')
    let bytes = 512 + scope.length * 2 + run.count * 8
    for (let i = 0; i < run.count; i++) bytes += runElementBytes(run.kinds[i], run.texts[i])
    if (this.usedBytes + bytes > this.limitBytes) {
      throw new ValueIndexError('quota', 'Value run storage quota exceeded', {
        resource: 'storage', requiredBytes: this.usedBytes + bytes, limitBytes: this.limitBytes, retryable: false,
      })
    }
    this.account?.charge(bytes)
    let copy: ValueKeyRun
    try { copy = Object.freeze({
      count: run.count, kinds: run.kinds.slice(0, run.count), numbers: run.numbers.slice(0, run.count),
      texts: Object.freeze(Array.from({ length: run.count }, (_, i) => run.texts[i])),
    }) } catch (error) { this.account?.release(bytes); throw error }
    const refId = `values-${++this.puts}`
    this.runs.set(refId, { run: copy, owner: { ...owner }, scope, bytes, readers: 0, retired: false })
    this.usedBytes += bytes
    return Object.freeze({ refId, count: run.count, owner: Object.freeze({ ...owner }) })
  }
  async read(ref: PagedKeySetRef, offset: number, limit: number, scope: string, signal?: AbortSignal): Promise<ValueKeyRun> {
    if (signal?.aborted) throw new ValueIndexError('cancelled', 'Value run read cancelled')
    const stored = this.runs.get(ref.refId)
    if (!stored || !sameStamp(stored.owner, ref.owner) || stored.run.count !== ref.count || stored.scope !== scope) {
      throw new ValueIndexError('stale-source', 'Value run ref is missing or owned by another source generation or column')
    }
    if (!Number.isSafeInteger(offset) || !Number.isSafeInteger(limit) || offset < 0 || limit < 1 || limit > VALUE_RUN_PAGE) {
      throw new ValueIndexError('invalid-cursor', 'Invalid value run page')
    }
    const end = Math.min(stored.run.count, offset + limit)
    const count = Math.max(0, end - offset), account = this.readerAllocation?.reserve()
    account?.charge(count * 17 + 256)
    try {
      const page = { count, kinds: stored.run.kinds.slice(offset, end), numbers: stored.run.numbers.slice(offset, end), texts: stored.run.texts.slice(offset, end) }
      const heldBytes = account?.bytes ?? 0
      const targets = account ? allocationTargets(page, [], account) : []
      account?.adopt(targets, heldBytes + targets.length * 64 + 128)
      if (account && page.texts.some(text => text !== undefined)) {
        // String aliases keep the original stored owner until every escaped
        // page/container/lane target dies; the run remains charged once.
        stored.readers++
        const token = { stored, remaining: targets.length }
        for (const target of targets) this.readerWatcher?.register(target, token)
      }
      return page
    } finally { account?.release() }
  }
  async release(ref: PagedKeySetRef): Promise<void> {
    const stored = this.runs.get(ref.refId)
    if (!stored) return
    this.runs.delete(ref.refId)
    stored.retired = true
    if (!stored.readers) this.drop(stored)
  }
  private drop(stored: StoredRun): void {
    this.usedBytes -= stored.bytes
    this.account?.release(stored.bytes)
  }
}
