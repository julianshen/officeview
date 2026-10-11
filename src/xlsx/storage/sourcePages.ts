/**
 * Sparse source pages for the paged adapter (spec §§4.1-4.3, plan S1.2).
 *
 * A page is one band of whole rows of one sheet, stored as structure-of-arrays
 * lanes: row (Uint32) and column (Uint16) stay separate lanes, never a packed
 * cell id (rule 3); text keeps raw UTF-16 code units, so lone surrogates
 * survive the binary round trip. Only stored records are encoded: blank rows
 * and cells cost nothing. Shared strings are stored the same way, in chunks
 * addressed by integer id, and decoded through the same bounded cache.
 *
 * Persistence goes through a PageBackend (OPFS, IndexedDB, or memory for small
 * admitted documents); decoded pages live in the document's MemoryPool
 * `source` tier, so the aggregate worker budget evicts them and a later read
 * reloads the same bytes from the backend under the same source identity.
 */
import { PackageError, type PackageErrorCode } from '../../core/package-error'
import type { MemoryPool, PoolEntry } from '../view/tileCache'
import type { QuotaLedger, Reservation } from './quota'
import type { ImportedPage, PageSink } from './importSession'
import { addRecord, emptyAggregate, type RectAggregate } from './overviewIndex'

// ---------------------------------------------------------------------------
// Payload

/** Value kinds of a stored record; mirrors parse.ts cell normalization. */
export const K_EMPTY = 0 // no value (style-only, uncached formula, empty numeric <v/>, t="e" without <v>)
export const K_NUMBER = 1
export const K_SST = 2 // nums holds the shared-string id
export const K_TEXT = 3 // text lane (inline string, t="str", numeric cell with non-numeric text)
export const K_BOOL = 4 // nums 1/0
export const K_ERROR = 5 // text lane, valueIsError
export const F_CACHED = 1
export const F_META = 2
/** t="e": valueIsError, also when the error cell has no <v> (value null). */
export const F_ERROR = 4

/** Formula metadata of one record; strings round-trip through JSON (lone surrogates are escaped). */
export interface FormulaMeta {
  formula?: string
  sharedFormula?: { si: number; ref?: string }
  arrayRef?: string
  formulaType?: string
  dynamicArray?: true
  ca?: true
}
export interface CellPagePayload {
  rows: Uint32Array
  cols: Uint16Array
  kinds: Uint8Array
  flags: Uint8Array
  styles: Uint32Array
  nums: Float64Array
  text: Uint16Array
  /** length = records + 1 */
  textStart: Uint32Array
  /** Record index -> formula metadata, only for records with F_META. */
  meta: ReadonlyMap<number, FormulaMeta>
}
export interface PageSummary {
  sheetId: string
  pageId: string
  r0: number
  r1: number
  /** Column bounds of the stored records, or c0 > c1 when the page holds none. */
  c0: number
  c1: number
  aggregate: RectAggregate
  bytes: number
}
export type StoredPage = ImportedPage & { payload: CellPagePayload }

export const recordPopulated = (p: CellPagePayload, i: number): boolean => p.kinds[i] !== K_EMPTY || (p.meta.get(i)?.formula !== undefined)

/** Builds one page; rows must be appended in ascending order. */
export class PageBuilder {
  private rows: number[] = []
  private cols: number[] = []
  private kinds: number[] = []
  private flags: number[] = []
  private styles: number[] = []
  private nums: number[] = []
  private texts: string[] = []
  private meta = new Map<number, FormulaMeta>()
  private units = 0
  get records(): number { return this.rows.length }
  /** Encoded size estimate used for the target page size. */
  get bytes(): number { return this.rows.length * 20 + this.units * 2 + this.meta.size * 64 }
  add(row: number, col: number, kind: number, styleIndex: number, num: number, text: string, cached: boolean, meta: FormulaMeta | undefined, error = kind === K_ERROR): void {
    const i = this.rows.length
    this.rows.push(row); this.cols.push(col); this.kinds.push(kind); this.styles.push(styleIndex); this.nums.push(num)
    this.texts.push(text)
    this.units += text.length
    this.flags.push((cached ? F_CACHED : 0) | (meta ? F_META : 0) | (error ? F_ERROR : 0))
    if (meta) this.meta.set(i, meta)
  }
  build(): CellPagePayload {
    const n = this.rows.length
    const text = new Uint16Array(this.units), textStart = new Uint32Array(n + 1)
    let at = 0
    for (let i = 0; i < n; i++) {
      textStart[i] = at
      const t = this.texts[i]
      for (let k = 0; k < t.length; k++) text[at++] = t.charCodeAt(k)
    }
    textStart[n] = at
    return {
      rows: Uint32Array.from(this.rows), cols: Uint16Array.from(this.cols), kinds: Uint8Array.from(this.kinds),
      flags: Uint8Array.from(this.flags), styles: Uint32Array.from(this.styles), nums: Float64Array.from(this.nums),
      text, textStart, meta: this.meta,
    }
  }
}

export function textAt(p: { text: Uint16Array; textStart: Uint32Array }, i: number): string {
  const start = p.textStart[i], end = p.textStart[i + 1]
  let out = ''
  for (let at = start; at < end; at += 8192) out += String.fromCharCode(...p.text.subarray(at, Math.min(end, at + 8192)))
  return out
}

export function summarize(sheetId: string, pageId: string, r0: number, r1: number, p: CellPagePayload, bytes: number): PageSummary {
  const aggregate = emptyAggregate()
  let c0 = Number.MAX_SAFE_INTEGER, c1 = -1
  for (let i = 0; i < p.rows.length; i++) {
    addRecord(aggregate, p.styles[i], recordPopulated(p, i))
    if (p.cols[i] < c0) c0 = p.cols[i]
    if (p.cols[i] > c1) c1 = p.cols[i]
  }
  return { sheetId, pageId, r0, r1, c0: c1 < 0 ? 1 : c0, c1: c1 < 0 ? 0 : c1, aggregate, bytes }
}

// Binary codec: 8 u32 header words, then 8-byte aligned lanes, then the meta JSON as UTF-16.
const MAGIC = 0x4f565031 // 'OVP1'
const align8 = (n: number) => (n + 7) & ~7

export function encodePage(p: CellPagePayload): Uint8Array {
  const n = p.rows.length
  const metaJson = p.meta.size ? JSON.stringify([...p.meta]) : ''
  const sizes = [8 * n, 4 * n, 4 * n, 4 * (n + 1), 2 * n, 2 * p.text.length, n, n, 2 * metaJson.length]
  const total = 32 + sizes.reduce((sum, s) => sum + align8(s), 0)
  const buffer = new ArrayBuffer(total)
  const head = new Uint32Array(buffer, 0, 8)
  head[0] = MAGIC; head[1] = n; head[2] = p.text.length; head[3] = metaJson.length
  let at = 32
  const put = (source: ArrayLike<number>, View: new (b: ArrayBuffer, o: number, l: number) => { set(a: ArrayLike<number>): void }, bytesPer: number, size: number) => {
    new View(buffer, at, size / bytesPer).set(source)
    at += align8(size)
  }
  put(p.nums, Float64Array, 8, sizes[0])
  put(p.rows, Uint32Array, 4, sizes[1])
  put(p.styles, Uint32Array, 4, sizes[2])
  put(p.textStart, Uint32Array, 4, sizes[3])
  put(p.cols, Uint16Array, 2, sizes[4])
  put(p.text, Uint16Array, 2, sizes[5])
  put(p.kinds, Uint8Array, 1, sizes[6])
  put(p.flags, Uint8Array, 1, sizes[7])
  const metaUnits = new Uint16Array(buffer, at, metaJson.length)
  for (let k = 0; k < metaJson.length; k++) metaUnits[k] = metaJson.charCodeAt(k)
  return new Uint8Array(buffer)
}

export function decodePage(bytes: Uint8Array): CellPagePayload {
  // Copy into an aligned buffer: backends may hand out unaligned views.
  const buffer = bytes.byteOffset % 8 === 0 && bytes.byteLength === bytes.buffer.byteLength ? bytes.buffer as ArrayBuffer : bytes.slice().buffer as ArrayBuffer
  if (buffer.byteLength < 32) throw new PackageError('parse-error', 'source page too short')
  const head = new Uint32Array(buffer, 0, 8)
  if (head[0] !== MAGIC) throw new PackageError('parse-error', 'source page has a bad header')
  const n = head[1], units = head[2], metaLength = head[3]
  const sizes = [8 * n, 4 * n, 4 * n, 4 * (n + 1), 2 * n, 2 * units, n, n, 2 * metaLength]
  if (32 + sizes.reduce((sum, s) => sum + align8(s), 0) !== buffer.byteLength) throw new PackageError('parse-error', 'source page size mismatch')
  let at = 32
  const take = <T>(View: new (b: ArrayBuffer, o: number, l: number) => T, bytesPer: number, size: number): T => {
    const view = new View(buffer, at, size / bytesPer)
    at += align8(size)
    return view
  }
  const nums = take(Float64Array, 8, sizes[0])
  const rows = take(Uint32Array, 4, sizes[1])
  const styles = take(Uint32Array, 4, sizes[2])
  const textStart = take(Uint32Array, 4, sizes[3])
  const cols = take(Uint16Array, 2, sizes[4])
  const text = take(Uint16Array, 2, sizes[5])
  const kinds = take(Uint8Array, 1, sizes[6])
  const flags = take(Uint8Array, 1, sizes[7])
  const metaUnits = take(Uint16Array, 2, sizes[8])
  let meta: Map<number, FormulaMeta> = new Map()
  if (metaLength) {
    let json = ''
    for (let k = 0; k < metaLength; k += 8192) json += String.fromCharCode(...metaUnits.subarray(k, Math.min(metaLength, k + 8192)))
    meta = new Map(JSON.parse(json) as [number, FormulaMeta][])
  }
  return { rows, cols, kinds, flags, styles, nums, text, textStart, meta }
}

/** Shared strings: a chunk is a text lane plus offsets. */
export interface StringChunk { text: Uint16Array; textStart: Uint32Array }
export function encodeStrings(strings: readonly string[]): Uint8Array {
  let units = 0
  for (const s of strings) units += s.length
  const buffer = new ArrayBuffer(8 + align8(4 * (strings.length + 1)) + 2 * units)
  const head = new Uint32Array(buffer, 0, 2)
  head[0] = strings.length; head[1] = units
  const starts = new Uint32Array(buffer, 8, strings.length + 1)
  const text = new Uint16Array(buffer, 8 + align8(4 * (strings.length + 1)), units)
  let at = 0
  strings.forEach((s, i) => {
    starts[i] = at
    for (let k = 0; k < s.length; k++) text[at++] = s.charCodeAt(k)
  })
  starts[strings.length] = at
  return new Uint8Array(buffer)
}
export function decodeStrings(bytes: Uint8Array): StringChunk {
  const buffer = bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength ? bytes.buffer as ArrayBuffer : bytes.slice().buffer as ArrayBuffer
  const head = new Uint32Array(buffer, 0, 2)
  const count = head[0], units = head[1]
  if (8 + align8(4 * (count + 1)) + 2 * units !== buffer.byteLength) throw new PackageError('parse-error', 'shared-string chunk size mismatch')
  return { textStart: new Uint32Array(buffer, 8, count + 1), text: new Uint16Array(buffer, 8 + align8(4 * (count + 1)), units) }
}

// ---------------------------------------------------------------------------
// Backends

export type PageBackendKind = 'memory' | 'opfs' | 'indexeddb'
export interface PageBackend {
  readonly kind: PageBackendKind
  /** Resolves once the bytes are durable (transaction complete / file closed). */
  put(key: string, bytes: Uint8Array): Promise<void>
  get(key: string): Promise<Uint8Array | undefined>
  /** Removes every stored page of this backend's owner. Idempotent. */
  removeAll(): Promise<void>
}

function storageError(error: unknown, action: string): PackageError {
  if (error instanceof PackageError) return error
  const name = (error as { name?: unknown } | null)?.name
  if (name === 'QuotaExceededError') return new PackageError('quota', `browser storage quota exceeded while ${action}`)
  return new PackageError('storage-error', `browser storage failed while ${action}${error instanceof Error ? `: ${error.message}` : ''}`)
}

export function memoryPageBackend(): PageBackend & { readonly size: number } {
  const pages = new Map<string, Uint8Array>()
  return {
    kind: 'memory',
    get size() { return pages.size },
    async put(key, bytes) { pages.set(key, bytes.slice()) },
    async get(key) { return pages.get(key) },
    async removeAll() { pages.clear() },
  }
}

interface OpfsDir {
  getFileHandle(name: string, opts?: { create?: boolean }): Promise<unknown>
  removeEntry(name: string): Promise<void>
  keys?(): AsyncIterableIterator<string>
}
interface OpfsFile {
  createWritable(): Promise<{ write(chunk: Uint8Array): Promise<void>; close(): Promise<void>; abort?(): Promise<void> }>
  getFile(): Promise<Blob>
}
export const PAGE_PREFIX = 'officeview-pages-'
/** One flat file per page, named `${PAGE_PREFIX}${owner}-${key}` so removeAll can find them. */
export function opfsPageBackend(dir: OpfsDir, owner: string): PageBackend {
  const prefix = `${PAGE_PREFIX}${owner}-`
  const written = new Set<string>()
  const fileName = (key: string) => `${prefix}${key.replace(/[^A-Za-z0-9_.-]/g, '_')}`
  return {
    kind: 'opfs',
    async put(key, bytes) {
      const name = fileName(key)
      try {
        const handle = await dir.getFileHandle(name, { create: true }) as OpfsFile
        written.add(name)
        const writable = await handle.createWritable()
        try {
          await writable.write(bytes)
          await writable.close()
        } catch (error) {
          try { await writable.abort?.() } catch { /* already failed */ }
          throw error
        }
      } catch (error) {
        throw storageError(error, 'writing a source page')
      }
    },
    async get(key) {
      try {
        const handle = await dir.getFileHandle(fileName(key)) as OpfsFile
        return new Uint8Array(await (await handle.getFile()).arrayBuffer())
      } catch (error) {
        if ((error as { name?: unknown } | null)?.name === 'NotFoundError') return undefined
        throw storageError(error, 'reading a source page')
      }
    },
    async removeAll() {
      const names = new Set(written)
      if (dir.keys) for await (const name of dir.keys()) if (name.startsWith(prefix)) names.add(name)
      for (const name of names) {
        try { await dir.removeEntry(name) } catch (error) {
          if ((error as { name?: unknown } | null)?.name !== 'NotFoundError') throw storageError(error, 'removing source pages')
        }
        written.delete(name)
      }
    },
  }
}

const PAGE_STORE = 'pages'
/** One database per owner; each put is its own transaction, awaited to completion. */
export function indexedDbPageBackend(factory: IDBFactory, owner: string): PageBackend {
  const name = `${PAGE_PREFIX}${owner}`
  let db: Promise<IDBDatabase> | undefined
  const open = () => db ??= new Promise<IDBDatabase>((resolve, reject) => {
    const req = factory.open(name, 1)
    req.onupgradeneeded = () => {
      const d = req.result
      if (!d.objectStoreNames.contains(PAGE_STORE)) d.createObjectStore(PAGE_STORE)
    }
    req.onsuccess = () => resolve(req.result)
    req.onerror = () => reject(req.error)
  })
  return {
    kind: 'indexeddb',
    async put(key, bytes) {
      try {
        const d = await open()
        const tx = d.transaction(PAGE_STORE, 'readwrite')
        const done = new Promise<void>((resolve, reject) => {
          tx.oncomplete = () => resolve()
          tx.onerror = () => reject(tx.error)
          tx.onabort = () => reject(tx.error ?? new DOMException('transaction aborted', 'AbortError'))
        })
        tx.objectStore(PAGE_STORE).put(new Blob([bytes as BlobPart]), key)
        await done
      } catch (error) {
        throw storageError(error, 'writing a source page')
      }
    },
    async get(key) {
      try {
        const d = await open()
        const tx = d.transaction(PAGE_STORE, 'readonly')
        const req = tx.objectStore(PAGE_STORE).get(key) as IDBRequest<Blob | undefined>
        const blob = await new Promise<Blob | undefined>((resolve, reject) => {
          req.onsuccess = () => resolve(req.result)
          req.onerror = () => reject(req.error)
        })
        return blob === undefined ? undefined : new Uint8Array(await blob.arrayBuffer())
      } catch (error) {
        throw storageError(error, 'reading a source page')
      }
    },
    async removeAll() {
      if (db) {
        try { (await db).close() } catch { /* never opened */ }
        db = undefined
      }
      await new Promise<void>((resolve, reject) => {
        const req = factory.deleteDatabase(name)
        req.onsuccess = () => resolve()
        req.onerror = () => reject(storageError(req.error, 'removing source pages'))
      })
    },
  }
}

export interface PageStorageEnv {
  opfs?: () => Promise<OpfsDir>
  indexedDb?: IDBFactory
  estimate?(): Promise<{ quota: number; usage: number }>
}
/** Run-time probe; nothing is assumed about a browser. */
export function browserPageStorageEnv(): PageStorageEnv {
  const nav = typeof navigator === 'undefined' ? undefined : navigator as Navigator & { storage?: StorageManager }
  const storage = nav?.storage
  const env: PageStorageEnv = {}
  if (storage && typeof storage.getDirectory === 'function' && typeof FileSystemFileHandle !== 'undefined' && 'createWritable' in FileSystemFileHandle.prototype) {
    env.opfs = () => storage.getDirectory() as unknown as Promise<OpfsDir>
  }
  if (typeof indexedDB !== 'undefined') env.indexedDb = indexedDB
  if (storage && typeof storage.estimate === 'function') {
    env.estimate = async () => {
      const e = await storage.estimate()
      return { quota: e.quota ?? 0, usage: e.usage ?? 0 }
    }
  }
  return env
}

/**
 * Picks where a document's pages live. Persistent storage is required for a
 * paged document; memory is a deterministic fallback only for documents whose
 * expected page bytes fit `memoryLimitBytes` (small admitted documents). No
 * storage and too large, or too little quota, is a typed refusal, never a
 * silent truncation or an unbounded in-memory store.
 */
export async function choosePageBackend(env: PageStorageEnv, owner: string, needBytes: number, memoryLimitBytes: number): Promise<PageBackend> {
  if (needBytes <= memoryLimitBytes) return memoryPageBackend()
  if (env.estimate) {
    let estimate: { quota: number; usage: number } | undefined
    try { estimate = await env.estimate() } catch { estimate = undefined }
    if (estimate && estimate.quota - estimate.usage < needBytes) {
      throw new PackageError('quota', `browser storage has ${Math.max(0, estimate.quota - estimate.usage)} bytes free; about ${needBytes} are needed to open this spreadsheet. Free some site storage and retry.`)
    }
  }
  if (env.opfs) {
    try { return opfsPageBackend(await env.opfs(), owner) } catch { /* fall through to IndexedDB */ }
  }
  if (env.indexedDb) return indexedDbPageBackend(env.indexedDb, owner)
  throw new PackageError('storage-error', 'This spreadsheet needs browser storage (OPFS or IndexedDB) to open, and none is available in this browser or mode.')
}

// ---------------------------------------------------------------------------
// Store: PageSink for the importer, page index and the decoded-page cache

export interface SourcePageStoreOptions {
  docId: string
  backend: PageBackend
  ledger?: QuotaLedger
  /** Decoded-page budget used until a MemoryPool is bound (default 32 MiB). */
  cacheBytes?: number
}
export interface PageStoreStats {
  committed: number
  storedBytes: number
  loads: number
  hits: number
  coalesced: number
  evictions: number
  residentBytes: number
  peakResidentBytes: number
}
interface Resident<T> { value: T; bytes: number; entry?: PoolEntry }

/** Decoded bytes of a payload (the typed arrays plus a per-meta allowance). */
export const payloadBytes = (p: CellPagePayload): number =>
  p.nums.byteLength + p.rows.byteLength + p.styles.byteLength + p.textStart.byteLength + p.cols.byteLength + p.text.byteLength + p.kinds.byteLength + p.flags.byteLength + p.meta.size * 96

export class SourcePageStore implements PageSink {
  readonly docId: string
  private readonly backend: PageBackend
  private readonly ledger?: QuotaLedger
  private readonly index = new Map<string, PageSummary[]>()
  private readonly reservations: Reservation[] = []
  private readonly resident = new Map<string, Resident<CellPagePayload | StringChunk>>()
  private readonly loading = new Map<string, Promise<CellPagePayload | StringChunk>>()
  private pool: MemoryPool | undefined
  private cacheBytes: number
  private stringChunkSize = 0
  private stringCount = 0
  private closed = false
  private abandoned = false
  private readonly counters: PageStoreStats = { committed: 0, storedBytes: 0, loads: 0, hits: 0, coalesced: 0, evictions: 0, residentBytes: 0, peakResidentBytes: 0 }

  constructor(options: SourcePageStoreOptions) {
    this.docId = options.docId
    this.backend = options.backend
    this.ledger = options.ledger
    this.cacheBytes = options.cacheBytes ?? 32 * 1024 * 1024
  }
  get backendKind(): PageBackendKind { return this.backend.kind }
  get stats(): PageStoreStats { return { ...this.counters } }
  get live(): boolean { return !this.closed && !this.abandoned }

  private check(docId: string): void {
    if (this.abandoned || this.closed) throw new PackageError('cancelled', 'source page store is closed')
    if (docId !== this.docId) throw new PackageError('cancelled', `page for document ${docId} rejected by the store of ${this.docId}`)
  }
  private charge(pool: 'pages' | 'sharedStrings', bytes: number): void {
    if (this.ledger) this.reservations.push(this.ledger.reserve(pool, bytes))
  }

  async commit(page: ImportedPage & { payload?: CellPagePayload }, docId = this.docId): Promise<void> {
    this.check(docId)
    if (!page.payload) throw new PackageError('parse-error', `page ${page.pageId} has no payload`)
    const bytes = encodePage(page.payload)
    this.charge('pages', bytes.byteLength)
    try {
      await this.backend.put(`${page.sheetId}/${page.pageId}`, bytes)
    } catch (error) {
      throw storageError(error, 'writing a source page')
    }
    // A store abandoned while the write was in flight never indexes the page.
    this.check(docId)
    const summary = summarize(page.sheetId, page.pageId, page.r0, page.r1, page.payload, bytes.byteLength)
    const list = this.index.get(page.sheetId) ?? []
    if (list.length && list[list.length - 1].r1 >= page.r0) throw new PackageError('parse-error', `page ${page.pageId} overlaps the previous page`)
    list.push(summary)
    this.index.set(page.sheetId, list)
    this.counters.committed++
    this.counters.storedBytes += bytes.byteLength
  }
  /** Shared strings, in fixed-size chunks addressed by id. */
  async commitStrings(chunk: number, strings: readonly string[], chunkSize: number, docId = this.docId): Promise<void> {
    this.check(docId)
    const bytes = encodeStrings(strings)
    this.charge('sharedStrings', bytes.byteLength)
    await this.backend.put(`sst/${chunk}`, bytes)
    this.check(docId)
    this.stringChunkSize = chunkSize
    this.stringCount = Math.max(this.stringCount, chunk * chunkSize + strings.length)
    this.counters.storedBytes += bytes.byteLength
  }
  get sharedStringCount(): number { return this.stringCount }

  async abandon(docId: string): Promise<void> {
    if (docId !== this.docId) return
    this.abandoned = true
    await this.dispose()
  }
  async close(): Promise<void> {
    this.closed = true
    await this.dispose()
  }
  private async dispose(): Promise<void> {
    for (const r of this.reservations.splice(0)) r.release()
    for (const res of this.resident.values()) res.entry?.release()
    this.resident.clear()
    this.loading.clear()
    this.index.clear()
    this.counters.residentBytes = 0
    await this.backend.removeAll()
  }

  pages(sheetId: string): readonly PageSummary[] { return this.index.get(sheetId) ?? [] }

  /** Registers decoded pages in the document pool's `source` tier from now on. */
  bindMemory(pool: MemoryPool): void {
    this.pool = pool
    for (const [key, res] of [...this.resident]) {
      if (!res.entry) res.entry = this.retain(key, res.bytes)
      if (!res.entry) this.drop(key, false)
    }
  }
  private retain(key: string, bytes: number): PoolEntry | undefined {
    if (!this.pool || this.pool.closed) return undefined
    try {
      return this.pool.retain(`source:${this.docId}:${key}`, {
        tier: 'source', bytes, onRelease: () => this.drop(key, true),
      }) ?? undefined
    } catch {
      return undefined
    }
  }
  private drop(key: string, evicted: boolean): void {
    const res = this.resident.get(key)
    if (!res) return
    this.resident.delete(key)
    this.counters.residentBytes -= res.bytes
    if (evicted) this.counters.evictions++
  }
  /** Local LRU used only while no pool is bound. */
  private trimLocal(): void {
    if (this.pool) return
    for (const key of this.resident.keys()) {
      if (this.counters.residentBytes <= this.cacheBytes) break
      this.drop(key, true)
    }
  }
  private async load<T extends CellPagePayload | StringChunk>(key: string, decode: (bytes: Uint8Array) => T, size: (value: T) => number): Promise<T> {
    if (!this.live) throw new PackageError('cancelled', 'source page store is closed')
    const hit = this.resident.get(key)
    if (hit) {
      this.counters.hits++
      this.resident.delete(key)
      this.resident.set(key, hit)
      hit.entry?.touch()
      return hit.value as T
    }
    const pending = this.loading.get(key)
    if (pending) { this.counters.coalesced++; return pending as Promise<T> }
    const promise = (async () => {
      this.counters.loads++
      const bytes = await this.backend.get(key)
      if (bytes === undefined) throw new PackageError('storage-error', `source page ${key} is missing from browser storage`)
      const value = decode(bytes)
      if (this.live && !this.resident.has(key)) {
        const bytesHeld = size(value)
        this.resident.set(key, { value, bytes: bytesHeld })
        this.counters.residentBytes += bytesHeld
        if (this.counters.residentBytes > this.counters.peakResidentBytes) this.counters.peakResidentBytes = this.counters.residentBytes
        const res = this.resident.get(key)!
        res.entry = this.retain(key, bytesHeld)
        // A bound pool that cannot admit the page serves it once, unretained.
        if (this.pool && !res.entry) this.drop(key, false)
        this.trimLocal()
      }
      return value
    })()
    this.loading.set(key, promise)
    try { return await promise } finally { this.loading.delete(key) }
  }
  page(sheetId: string, pageId: string): Promise<CellPagePayload> {
    return this.load(`${sheetId}/${pageId}`, decodePage, payloadBytes)
  }
  async sharedString(id: number): Promise<string | undefined> {
    if (!Number.isSafeInteger(id) || id < 0 || id >= this.stringCount || this.stringChunkSize === 0) return undefined
    const chunk = Math.floor(id / this.stringChunkSize)
    const strings = await this.load(`sst/${chunk}`, decodeStrings, (c) => c.text.byteLength + c.textStart.byteLength)
    return textAt(strings, id - chunk * this.stringChunkSize)
  }
}

export function failureCode(error: unknown): PackageErrorCode {
  return error instanceof PackageError ? error.code : 'storage-error'
}
