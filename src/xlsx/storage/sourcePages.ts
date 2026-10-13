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
import { MemoryPool, WORKER_BUDGETS, type PoolEntry } from '../view/tileCache'
import { BUILDER_RECORD_BYTES, ConstructionMemory, META_OBJECT_BYTES, PAGE_SUMMARY_BYTES, TEXT_STRING_BYTES, type ConstructionShare } from './constructionMemory'
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
/** The value is a blob record (spec §4.3.4): nums holds the document's blob
 * id and the text lane holds nothing for this record. */
export const F_BLOB = 8
/** §4.3.4 initial profile: 64 KiB inline-string soft threshold, the same
 * constant as ImportLimits.inlineStringSoftBytes. Measured in stored UTF-16
 * bytes (2 per code unit), which is what the value would add to the text lane.
 * admission.ts checkCaps compares the same field to a character count to route
 * object vs paged; that routing is unchanged and independent of the spill. */
export const INLINE_TEXT_SOFT_BYTES = 64 * 1024

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
/** A built page and the oversized values its F_BLOB records reference by
 * ordinal (nums), in record order; commit stores them as blob records first. */
export type StoredPage = ImportedPage & { payload: CellPagePayload; blobs?: readonly string[] }

/**
 * Length of JSON.stringify(value) for the JSON subset formula metadata uses
 * (strings, finite numbers, booleans, null, arrays, plain objects with
 * undefined members omitted), computed without allocating the JSON, so the
 * encode buffer and meta string can be charged before they exist.
 */
export function jsonLength(value: unknown): number {
  switch (typeof value) {
    case 'string': {
      let n = 2
      for (let k = 0; k < value.length; k++) {
        const c = value.charCodeAt(k)
        if (c === 0x22 || c === 0x5c || c === 0x08 || c === 0x0c || c === 0x0a || c === 0x0d || c === 0x09) n += 2
        else if (c < 0x20) n += 6
        else if (c >= 0xd800 && c <= 0xdbff) {
          const next = value.charCodeAt(k + 1)
          if (next >= 0xdc00 && next <= 0xdfff) { n += 2; k++ } else n += 6
        } else if (c >= 0xdc00 && c <= 0xdfff) n += 6
        else n += 1
      }
      return n
    }
    case 'number': return Number.isFinite(value) ? String(value).length : 4
    case 'boolean': return value ? 4 : 5
    case 'object': {
      if (value === null) return 4
      if (Array.isArray(value)) {
        let n = 2 + Math.max(0, value.length - 1)
        for (const item of value) n += item === undefined || typeof item === 'function' ? 4 : jsonLength(item)
        return n
      }
      let n = 2, members = 0
      for (const [key, item] of Object.entries(value)) {
        if (item === undefined || typeof item === 'function') continue
        n += jsonLength(key) + 1 + jsonLength(item)
        members++
      }
      return n + Math.max(0, members - 1)
    }
    default: return 4
  }
}
const metaArrayLength = (meta: ReadonlyMap<number, FormulaMeta>): number => {
  if (!meta.size) return 0
  let n = 2 + meta.size - 1
  for (const entry of meta) n += jsonLength(entry)
  return n
}

export const recordPopulated = (p: CellPagePayload, i: number): boolean => p.kinds[i] !== K_EMPTY || (p.meta.get(i)?.formula !== undefined)

/** Builds one page; rows must be appended in ascending order. A text value
 * over the inline soft threshold is spilled: the record keeps its kind and
 * error flag, gets F_BLOB and its ordinal in `blobs`, and adds no text units. */
export class PageBuilder {
  private readonly inlineTextBytes: number
  private spilled: string[] = []
  private spilledUnits = 0
  private rows: number[] = []
  private cols: number[] = []
  private kinds: number[] = []
  private flags: number[] = []
  private styles: number[] = []
  private nums: number[] = []
  private texts: string[] = []
  private meta = new Map<number, FormulaMeta>()
  private units = 0
  /** Non-empty inline and spilled strings held (each costs a string header). */
  private strings = 0
  /** Sum of JSON.stringify([index, meta]) lengths, for the exact meta lane size. */
  private metaChars = 0
  constructor(options: { inlineTextBytes?: number } = {}) {
    this.inlineTextBytes = options.inlineTextBytes ?? INLINE_TEXT_SOFT_BYTES
  }
  get records(): number { return this.rows.length }
  /** Oversized values held for this page, in F_BLOB ordinal order. */
  get blobs(): readonly string[] { return this.spilled }
  /** Stored bytes of the spilled values (2 per UTF-16 code unit). */
  get spilledBytes(): number { return 2 * this.spilledUnits }
  /** Estimated heap held by this builder, for construction accounting (constructionMemory.ts). */
  get memoryBytes(): number {
    return this.rows.length * BUILDER_RECORD_BYTES + this.strings * TEXT_STRING_BYTES + 2 * (this.units + this.spilledUnits + this.metaChars) + this.meta.size * META_OBJECT_BYTES
  }
  /** What add(…, text, …, meta) will add to memoryBytes, measured without allocating (charge it first). */
  costOf(text: string, meta: FormulaMeta | undefined): number {
    return BUILDER_RECORD_BYTES + (text.length ? TEXT_STRING_BYTES + 2 * text.length : 0)
      + (meta ? META_OBJECT_BYTES + 2 * jsonLength([this.rows.length, meta]) : 0)
  }
  /** Exact encodePage() size of what has been added (the page target is charged on it). */
  get bytes(): number {
    const n = this.rows.length
    const metaLength = this.meta.size ? 2 + this.metaChars + (this.meta.size - 1) : 0
    return encodedSize(n, this.units, metaLength)
  }
  add(row: number, col: number, kind: number, styleIndex: number, num: number, text: string, cached: boolean, meta: FormulaMeta | undefined, error = kind === K_ERROR): void {
    const i = this.rows.length
    const spill = 2 * text.length > this.inlineTextBytes
    if (spill) {
      num = this.spilled.length
      this.spilled.push(text)
      this.spilledUnits += text.length
      text = ''
    }
    this.rows.push(row); this.cols.push(col); this.kinds.push(kind); this.styles.push(styleIndex); this.nums.push(num)
    this.texts.push(text)
    this.units += text.length
    if (text.length || spill) this.strings++
    this.flags.push((cached ? F_CACHED : 0) | (meta ? F_META : 0) | (error ? F_ERROR : 0) | (spill ? F_BLOB : 0))
    if (meta) {
      this.meta.set(i, meta)
      this.metaChars += jsonLength([i, meta])
    }
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

function encodedSize(n: number, units: number, metaLength: number): number {
  return 32 + [8 * n, 4 * n, 4 * n, 4 * (n + 1), 2 * n, 2 * units, n, n, 2 * metaLength].reduce((sum, size) => sum + align8(size), 0)
}

/** `reserve` is told the scratch bytes (encoded buffer plus meta JSON) before
 * the buffer is allocated, so the caller can charge or refuse them first. */
export function encodePage(p: CellPagePayload, reserve?: (scratchBytes: number) => void): Uint8Array {
  const n = p.rows.length
  // Both the buffer and the meta JSON are sized and charged before either is
  // allocated, from the metadata as it is now (FormulaMeta objects are mutable,
  // so a cached length could be stale).
  const metaLength = metaArrayLength(p.meta)
  const scratch = (length: number) => 32 + [8 * n, 4 * n, 4 * n, 4 * (n + 1), 2 * n, 2 * p.text.length, n, n, 2 * length].reduce((sum, s) => sum + align8(s), 0) + 2 * length
  reserve?.(scratch(metaLength))
  const metaJson = p.meta.size ? JSON.stringify([...p.meta]) : ''
  if (metaJson.length !== metaLength) throw new PackageError('parse-error', 'formula metadata changed while its page was encoded')
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
export function encodeStrings(strings: readonly string[], reserve?: (scratchBytes: number) => void): Uint8Array {
  let units = 0
  for (const s of strings) units += s.length
  const total = 8 + align8(4 * (strings.length + 1)) + 2 * units
  reserve?.(total)
  const buffer = new ArrayBuffer(total)
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
  /** Resolves once the bytes are durable (transaction complete / file closed).
   * Ownership of `bytes` passes to the backend; callers must not reuse them. */
  put(key: string, bytes: Uint8Array): Promise<void>
  get(key: string): Promise<Uint8Array | undefined>
  /** Removes every stored page of this backend's owner. Idempotent. */
  /** Removes one record (a failed page's blob records). Idempotent: a missing key is not an error. */
  delete(key: string): Promise<void>
  /** Keys this owner holds, examining at most `max` entries of the underlying
   * store (not `max` matches), stopping early when `signal` aborts. Optional for test fakes. */
  list?(max: number, signal?: AbortSignal): Promise<string[]>
  removeAll(): Promise<void>
}

function storageError(error: unknown, action: string): PackageError {
  if (error instanceof PackageError) return error
  const name = (error as { name?: unknown } | null)?.name
  if (name === 'QuotaExceededError') return new PackageError('quota', `browser storage quota exceeded while ${action}`)
  return new PackageError('storage-error', `browser storage failed while ${action}${error instanceof Error ? `: ${error.message}` : ''}`)
}

/** Takes ownership of the bytes it is given (the store hands it a fresh
 * encodePage buffer and never reuses it), so held bytes equal charged bytes. */
export function memoryPageBackend(): PageBackend & { readonly size: number; readonly heldBytes: number } {
  const pages = new Map<string, Uint8Array>()
  return {
    kind: 'memory',
    get size() { return pages.size },
    get heldBytes() { let n = 0; for (const b of pages.values()) n += b.byteLength; return n },
    async put(key, bytes) { pages.set(key, bytes) },
    async get(key) { return pages.get(key) },
    async delete(key) { pages.delete(key) },
    async list(max, signal) {
      const out: string[] = []
      for (const key of pages.keys()) {
        if (out.length >= max || signal?.aborted) break
        out.push(key)
      }
      return out
    },
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
/**
 * Visits at most `max` directory entries, checking the bound and `signal`
 * before each next() (never pulling an entry it will not examine), and closes
 * the iterator when it stops early. `visit` returning false stops it too.
 */
async function eachEntry(names: AsyncIterable<string>, max: number, signal: AbortSignal | undefined, visit: (name: string) => boolean | void): Promise<void> {
  const it = names[Symbol.asyncIterator]()
  let exhausted = false
  try {
    for (let examined = 0; examined < max && !signal?.aborted; examined++) {
      const next = await it.next()
      if (next.done) { exhausted = true; return }
      if (visit(next.value) === false) return
    }
  } finally {
    if (!exhausted) try { await it.return?.() } catch { /* closing is best effort */ }
  }
}
const escapeName = (text: string): string => text.replace(/[^A-Za-z0-9.]/g, (c) => `_${c.charCodeAt(0).toString(16).padStart(4, '0')}`)
const unescapeName = (text: string): string => text.replace(/_([0-9a-f]{4})/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)))
/** One flat file per page, named `${PAGE_PREFIX}${owner}-${key}` so removeAll can find them. */
export function opfsPageBackend(dir: OpfsDir, owner: string): PageBackend {
  // Injective names: every character outside [A-Za-z0-9.] becomes _hhhh (its
  // UTF-16 code unit), so distinct keys never share a file, and the encoded
  // owner contains no '-', so one owner's prefix never matches another's files.
  const prefix = `${PAGE_PREFIX}${escapeName(owner)}-`
  const written = new Set<string>()
  const fileName = (key: string) => `${prefix}${escapeName(key)}`
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
    async delete(key) {
      const name = fileName(key)
      try { await dir.removeEntry(name) } catch (error) {
        if ((error as { name?: unknown } | null)?.name !== 'NotFoundError') throw storageError(error, 'removing a source record')
      }
      written.delete(name)
    },
    async list(max, signal) {
      const out: string[] = []
      // The OPFS root is shared with other owners and the spool: every entry examined counts.
      const names = dir.keys ? dir.keys() : (async function* () { yield* written })()
      await eachEntry(names, max, signal, (name) => {
        if (name.startsWith(prefix)) out.push(unescapeName(name.slice(prefix.length)))
      })
      return out
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
    async delete(key) {
      try {
        const d = await open()
        const tx = d.transaction(PAGE_STORE, 'readwrite')
        const done = new Promise<void>((resolve, reject) => {
          tx.oncomplete = () => resolve()
          tx.onerror = () => reject(tx.error)
          tx.onabort = () => reject(tx.error ?? new DOMException('transaction aborted', 'AbortError'))
        })
        tx.objectStore(PAGE_STORE).delete(key)
        await done
      } catch (error) {
        throw storageError(error, 'removing a source record')
      }
    },
    async list(max) {
      try {
        const d = await open()
        const tx = d.transaction(PAGE_STORE, 'readonly')
        const req = tx.objectStore(PAGE_STORE).getAllKeys(null, max)
        const keys = await new Promise<IDBValidKey[]>((resolve, reject) => {
          req.onsuccess = () => resolve(req.result)
          req.onerror = () => reject(req.error)
        })
        return keys.map(String)
      } catch (error) {
        throw storageError(error, 'listing source records')
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

/** Same age rule as the spool sweep (core/spool.ts): a leftover older than this belongs to a dead session. */
export const PAGE_OWNER_MAX_AGE_MS = 24 * 3600_000
/** Entries a sweep examines per backend at most (bounded work at session start). */
export const SWEEP_MAX_ENTRIES = 4096
/** A page-storage owner: `${docId}-${creation time in base 36}.${64 random bits in hex}`.
 * The random part makes owners collision-resistant (two opens of one docId in
 * the same millisecond never share storage); the sweep reads the age back. */
export function pageOwner(docId: string, now = Date.now(), random: () => Uint8Array = () => crypto.getRandomValues(new Uint8Array(8))): string {
  const hex = [...random()].map((b) => b.toString(16).padStart(2, '0')).join('')
  return `${docId}-${now.toString(36)}.${hex}`
}
function ownerCreated(owner: string): number | undefined {
  const match = /-([0-9a-z]+)\.[0-9a-f]{16}$/.exec(owner)
  if (!match) return undefined
  const created = parseInt(match[1], 36)
  return Number.isSafeInteger(created) ? created : undefined
}
/** The Web Locks name an owner is held under for its whole lifetime. */
export const ownerLockName = (owner: string): string => `${PAGE_PREFIX}lock-${owner}`

/** The subset of the Web Locks API (navigator.locks) used here; origin-wide, so shared by every tab and worker. */
export interface LockManagerLike {
  request<T>(name: string, options: { mode?: 'exclusive' | 'shared'; ifAvailable?: boolean }, callback: (lock: unknown) => Promise<T> | T): Promise<T>
}
/**
 * Cross-realm owner exclusion. An owner's lock is held from before its first
 * write until its storage is removed; the sweep deletes another owner's
 * storage only while it holds that owner's lock itself (ifAvailable), which a
 * live tab or worker never lets it have.
 */
export interface OwnerLocks {
  /** Acquires the owner's exclusive lock (refusing if anyone holds it); resolves with its release. */
  hold(owner: string): Promise<() => void>
  /** Runs `fn` holding the owner's lock if nobody else holds it; false (fn not run) otherwise. */
  tryWith(owner: string, fn: () => Promise<void>): Promise<boolean>
}
export function ownerLocks(manager: LockManagerLike): OwnerLocks {
  return {
    hold(owner) {
      return new Promise((resolve, reject) => {
        let release!: () => void
        const held = new Promise<void>((r) => { release = r })
        manager.request(ownerLockName(owner), { mode: 'exclusive', ifAvailable: true }, async (lock) => {
          if (!lock) { reject(new PackageError('storage-error', `page storage owner ${owner} is already in use`)); return }
          resolve(release)
          await held
        }).catch(reject)
      })
    },
    tryWith(owner, fn) {
      return manager.request(ownerLockName(owner), { mode: 'exclusive', ifAvailable: true }, async (lock) => {
        if (!lock) return false
        await fn()
        return true
      })
    },
  }
}
/** Heap model of the sweep's working set (Map entry, names array and owner
 * string per owner; array slot and string per name), plus 2 B per code unit.
 * Pinned by scripts/construction-memory-probe.ts ('leftover-sweep'). */
export const SWEEP_OWNER_BYTES = 256
export const SWEEP_NAME_BYTES = 48
/**
 * Removes page storage left by crashed or killed sessions: OPFS files and
 * IndexedDB databases of owners older than `maxAgeMs` (the spool's 24 h rule),
 * never an owner this process holds. Bounded (at most SWEEP_MAX_ENTRIES
 * examined per backend), cancellable between entries, best effort (a failed
 * removal is left for the next sweep).
 *
 * The names it groups by owner stay reachable while it awaits each owner's
 * lock, so `reserve(totalBytes)` (a construction share's need) is called with
 * the running total before each is retained. A refusal stops collecting: the
 * owners already charged are still swept, the rest wait for a later sweep. A
 * refusal while sweeping (the share's store was disposed) stops the sweep: it
 * is checked again after every await, before each lock, removal and deletion,
 * so only a removal or deletion already started may still land.
 */
export async function sweepLeftoverPages(env: PageStorageEnv, options: { now?: number; maxAgeMs?: number; signal?: AbortSignal; maxEntries?: number; reserve?: (totalBytes: number) => void } = {}): Promise<number> {
  const now = options.now ?? Date.now()
  const maxAge = options.maxAgeMs ?? PAGE_OWNER_MAX_AGE_MS
  const maxEntries = options.maxEntries ?? SWEEP_MAX_ENTRIES
  let charged = 0
  /** False (nothing retained) when the working set cannot grow by `bytes`. */
  const admit = (bytes: number): boolean => {
    try { options.reserve?.(charged + bytes) } catch { return false }
    charged += bytes
    return true
  }
  /** Whether what is retained is still charged (a released share refuses). */
  const stillCharged = (): boolean => admit(0)
  /** Checked after every await, before the next lock, removal or deletion. */
  const stop = (): boolean => options.signal?.aborted === true || !stillCharged()
  // Without Web Locks no owner can be proven dead: other owners are never swept.
  const locks = env.locks
  if (!locks) return 0
  // Age is an extra guard; the lock is the proof that no tab or worker holds the owner.
  const old = (owner: string) => {
    const created = ownerCreated(owner)
    return created !== undefined && created <= now && now - created > maxAge
  }
  let removed = 0
  if (env.opfs) {
    try {
      const dir = await env.opfs()
      if (dir.keys) {
        const byOwner = new Map<string, string[]>()
        await eachEntry(dir.keys(), maxEntries, options.signal, (name) => {
          if (!name.startsWith(PAGE_PREFIX)) return
          const rest = name.slice(PAGE_PREFIX.length)
          const dash = rest.indexOf('-')
          if (dash <= 0) return
          const owner = unescapeName(rest.slice(0, dash))
          if (!old(owner)) return
          let names = byOwner.get(owner)
          if (!admit((names ? 0 : SWEEP_OWNER_BYTES + 2 * owner.length) + SWEEP_NAME_BYTES + 2 * name.length)) return false
          if (!names) byOwner.set(owner, names = [])
          names.push(name)
        })
        for (const [owner, names] of byOwner) {
          if (stop()) break
          await locks.tryWith(owner, async () => {
            for (const name of names) {
              // Also after the lock wait and each removal: disposal may have come meanwhile.
              if (stop()) break
              try { await dir.removeEntry(name); removed++ } catch { /* already gone */ }
            }
          }).catch(() => false)
        }
      }
    } catch { /* no OPFS this time */ }
  }
  const factory = env.indexedDb as (IDBFactory & { databases?(): Promise<{ name?: string }[]> }) | undefined
  if (factory?.databases && !options.signal?.aborted) {
    try {
      let dbs = (await factory.databases()).slice(0, maxEntries)
      // The listing is the browser's; the retained slice is charged like the OPFS grouping.
      if (!admit(dbs.reduce((n, { name }) => n + SWEEP_NAME_BYTES + 2 * (name?.length ?? 0), 0))) dbs = []
      for (const { name } of dbs) {
        if (stop()) break
        if (!name?.startsWith(PAGE_PREFIX) || name.startsWith(`${PAGE_PREFIX}lock-`)) continue
        const owner = name.slice(PAGE_PREFIX.length)
        if (!old(owner)) continue
        await locks.tryWith(owner, async () => {
          if (stop()) return
          await new Promise<void>((resolve, reject) => {
            const req = factory.deleteDatabase(name)
            req.onsuccess = () => { removed++; resolve() }
            req.onerror = () => reject(req.error)
            // Still open somewhere: never wait on it; the next sweep retries.
            req.onblocked = () => reject(new DOMException('blocked', 'InvalidStateError'))
          })
        }).catch(() => false)
      }
    } catch { /* databases() unavailable */ }
  }
  return removed
}

export interface PageStorageEnv {
  opfs?: () => Promise<OpfsDir>
  indexedDb?: IDBFactory
  estimate?(): Promise<{ quota: number; usage: number }>
  /** Web Locks (navigator.locks); without them no other owner is ever swept. */
  locks?: OwnerLocks
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
  const locks = (nav as (Navigator & { locks?: LockManagerLike }) | undefined)?.locks
  if (locks && typeof locks.request === 'function') env.locks = ownerLocks(locks)
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
  /** Provisional construction budget until the document pool is bound (default: the desktop worker budget). */
  memoryBudgetBytes?: number
  /** Release of the owner's Web Lock, held by the caller for this store's lifetime
   * (released once the store's storage is removed and no write it started can
   * still land; see `ownerReleased`). Required for an OPFS/IndexedDB store to
   * sweep within its owner. */
  ownerLock?: () => void
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
  /** Oversized-value blob records written (spec §4.3.4) and their stored bytes. */
  blobs: number
  blobBytes: number
}
interface Resident<T> { value: T; bytes: number; entry?: PoolEntry }
interface WrittenBlob { key: string; record: number; ordinal: number; bytes: number; counted: boolean; reservation: Reservation | undefined; ram: () => void }
/** `/` and `%` escaped, so a segment never contains a raw separator (injective for any code units). */
const keySegment = (text: string): string => text.replace(/[%/]/g, (c) => (c === '%' ? '%25' : '%2F'))
const unkeySegment = (text: string): string => text.replace(/%2F|%25/g, (c) => (c === '%2F' ? '/' : '%'))
/** Record keys live in disjoint namespaces: `page/<sheet>/<page>`, `blob/<id>`, `sst/<chunk>`.
 * Any sheet or page id (e.g. a sheet called "sst" or "blob") maps to its own key. */
export const pageRecordKey = (sheetId: string, pageId: string): string => `page/${keySegment(sheetId)}/${keySegment(pageId)}`
/** Decoded-object overhead of one cached page/chunk (payload object, typed-array views,
 * meta Map, resident record, pool entry), charged even when its backing is shared.
 * Pinned by scripts/construction-memory-probe.ts ('decoded-pages'). */
export const DECODED_PAGE_OVERHEAD_BYTES = 2048
export const DECODED_CHUNK_OVERHEAD_BYTES = 1280
/** Persistent pool entry for the records a RAM-backed (memory) backend retains. */
const RAM_KEY = 'memory-backend'

/** Typed-lane bytes of a payload. */
const laneBytes = (p: CellPagePayload): number =>
  p.nums.byteLength + p.rows.byteLength + p.styles.byteLength + p.textStart.byteLength + p.cols.byteLength + p.text.byteLength + p.kinds.byteLength + p.flags.byteLength
/** Formula metadata a payload owns: per object plus 2 B per JSON character (at least its strings). */
const metaBytes = (p: CellPagePayload): number => p.meta.size * META_OBJECT_BYTES + 2 * metaArrayLength(p.meta)
/** Heap held by a payload: its typed lanes and its formula metadata. */
export const payloadBytes = (p: CellPagePayload): number => laneBytes(p) + metaBytes(p)

export class SourcePageStore implements PageSink {
  readonly docId: string
  private readonly backend: PageBackend
  private chargedTo?: QuotaLedger
  private readonly index = new Map<string, PageSummary[]>()
  private readonly reservations: Reservation[] = []
  private readonly resident = new Map<string, Resident<CellPagePayload | StringChunk>>()
  private readonly loading = new Map<string, Promise<CellPagePayload | StringChunk>>()
  /** Import/index construction charges (§12.4); see constructionMemory.ts. */
  readonly construction: ConstructionMemory
  private pool: MemoryPool | undefined
  private cacheBytes: number
  private stringChunkSize = 0
  private stringCount = 0
  private blobCount = 0
  /** Bytes the memory backend retains for this store (charged to the pool, see holdRam). */
  private ramBytes = 0
  private readonly committedChunks = new Set<number>()
  /** Keys whose write has started but whose page/chunk is not yet published or rolled back. */
  private readonly inFlight = new Set<string>()
  /** Blob ids a rollback deleted (ids are never reused; everything else below blobCount is referenced). */
  private readonly discardedBlobs = new Set<number>()
  /** Sheets with a commit in flight: one at a time per sheet, so validation precedes every write. */
  private readonly committing = new Set<string>()
  private closed = false
  private abandoned = false
  private ownerLock: (() => void) | undefined
  /** Started commits (write, publication or rollback), settled or not; see dispose. */
  private readonly operations = new Set<Promise<unknown>>()
  private disposing: Promise<void> | undefined
  private released!: () => void
  /** Resolves once the owner's lock release was handed to the lock manager (see dispose); never rejects. */
  readonly ownerReleased: Promise<void> = new Promise<void>((resolve) => { this.released = resolve })
  private readonly counters: PageStoreStats = { committed: 0, storedBytes: 0, loads: 0, hits: 0, coalesced: 0, evictions: 0, residentBytes: 0, peakResidentBytes: 0, blobs: 0, blobBytes: 0 }

  constructor(options: SourcePageStoreOptions) {
    this.docId = options.docId
    this.backend = options.backend
    this.chargedTo = options.ledger
    this.cacheBytes = options.cacheBytes ?? 32 * 1024 * 1024
    this.ownerLock = options.ownerLock
    this.construction = new ConstructionMemory(new MemoryPool({ limitBytes: options.memoryBudgetBytes ?? WORKER_BUDGETS.D }), options.docId)
  }
  get backendKind(): PageBackendKind { return this.backend.kind }
  get ledger(): QuotaLedger | undefined { return this.chargedTo }
  /** The import session's ledger, for a store created without one; never replaces an existing ledger. */
  attachLedger(ledger: QuotaLedger): void { this.chargedTo ??= ledger }
  get stats(): PageStoreStats { return { ...this.counters } }
  get live(): boolean { return !this.closed && !this.abandoned }

  private check(docId: string): void {
    if (this.abandoned || this.closed) throw new PackageError('cancelled', 'source page store is closed')
    if (docId !== this.docId) throw new PackageError('cancelled', `page for document ${docId} rejected by the store of ${this.docId}`)
  }
  private charge(pool: 'pages' | 'sharedStrings' | 'blobRecords', bytes: number): Reservation | undefined {
    if (!this.chargedTo) return undefined
    const reservation = this.chargedTo.reserve(pool, bytes)
    this.reservations.push(reservation)
    return reservation
  }
  private uncharge(reservation: Reservation | undefined): void {
    if (!reservation) return
    reservation.release()
    const at = this.reservations.indexOf(reservation)
    if (at >= 0) this.reservations.splice(at, 1)
  }
  /**
   * A memory backend keeps the encoded buffer it is given in RAM for the
   * store's lifetime: charge it to the document pool (persistent, non-evictable)
   * before handing it over, refusing typed when the pool cannot hold it. Other
   * backends persist to browser storage and hold nothing. Returns the undo for
   * a record that is deleted again (idempotent; a disposed store released all).
   * While the write is in flight the buffer is also in the encode share, a
   * brief conservative double count.
   */
  private holdRam(bytes: number): () => void {
    if (this.backend.kind !== 'memory') return () => {}
    this.construction.persist(RAM_KEY, this.ramBytes + bytes, 'source pages held in memory')
    this.ramBytes += bytes
    let undone = false
    return () => {
      if (undone || !this.live) return
      undone = true
      this.ramBytes -= bytes
      try { this.construction.persist(RAM_KEY, this.ramBytes, 'source pages held in memory') } catch { /* pool closed */ }
    }
  }

  async commit(page: ImportedPage & { payload?: CellPagePayload; blobs?: readonly string[] }, docId = this.docId): Promise<void> {
    this.check(docId)
    if (!page.payload) throw new PackageError('parse-error', `page ${page.pageId} has no payload`)
    const list = this.index.get(page.sheetId) ?? []
    // Every rejection happens before the first write, so a rejected page can
    // never overwrite or orphan what an earlier page stored.
    if (this.committing.has(page.sheetId)) throw new PackageError('parse-error', `page ${page.pageId} arrived while another page of sheet ${page.sheetId} is being committed`)
    if (!(Number.isSafeInteger(page.r0) && Number.isSafeInteger(page.r1) && page.r0 >= 0 && page.r1 >= page.r0)) throw new PackageError('parse-error', `page ${page.pageId} has an invalid row range`)
    if (list.length && list[list.length - 1].r1 >= page.r0) throw new PackageError('parse-error', `page ${page.pageId} overlaps the previous page`)
    if (list.some((summary) => summary.pageId === page.pageId)) throw new PackageError('parse-error', `page ${page.pageId} is already committed`)
    this.committing.add(page.sheetId)
    try {
      await this.track(this.commitChecked(page as ImportedPage & { payload: CellPagePayload; blobs?: readonly string[] }, list, docId))
    } finally {
      this.committing.delete(page.sheetId)
    }
  }
  /** Registers a started storage mutation until it settled, rollback included. */
  private track<T>(operation: Promise<T>): Promise<T> {
    const settled = operation.then(() => {}, () => {})
    this.operations.add(settled)
    void settled.then(() => this.operations.delete(settled))
    return operation
  }
  private async commitChecked(page: ImportedPage & { payload: CellPagePayload; blobs?: readonly string[] }, list: PageSummary[], docId: string): Promise<void> {
    // The summary is persistent index memory: charged before anything is written.
    const indexKey = `index:${page.sheetId}`
    this.construction.persist(indexKey, (list.length + 1) * PAGE_SUMMARY_BYTES, 'the page index')
    const scratch = this.construction.share('page encoding')
    const written: WrittenBlob[] = []
    const pageKey = pageRecordKey(page.sheetId, page.pageId)
    let pageCharge: Reservation | undefined
    let pageRam: () => void = () => {}
    let pageWritten = false
    let bytes: Uint8Array
    try {
      await this.commitBlobs(page.payload, page.blobs ?? [], page.pageId, docId, scratch, written)
      bytes = encodePage(page.payload, (n) => scratch.need(n))
      pageCharge = this.charge('pages', bytes.byteLength)
      pageRam = this.holdRam(bytes.byteLength)
      pageWritten = true
      this.inFlight.add(pageKey)
      try {
        await this.backend.put(pageKey, bytes)
      } catch (error) {
        throw storageError(error, 'writing a source page')
      }
      // A store abandoned while the write was in flight never indexes the page.
      this.check(docId)
      // Publication is the last step inside the rollback boundary.
      list.push(summarize(page.sheetId, page.pageId, page.r0, page.r1, page.payload, bytes.byteLength))
      this.index.set(page.sheetId, list)
    } catch (error) {
      if (this.live) {
        // The page was not indexed: storage, ledger and pool return to their pre-page state.
        // A record the backend cannot delete keeps its charges: they match what is still held.
        const removed = !pageWritten || await this.backend.delete(pageKey).then(() => true, () => false)
        if (removed) { this.uncharge(pageCharge); pageRam() }
        await this.discardBlobs(page.payload, written)
        try { this.construction.persist(indexKey, list.length * PAGE_SUMMARY_BYTES, 'the page index') } catch { /* pool refused or closed */ }
      } else {
        // Abandoned/closed meanwhile: dispose already released every charge and
        // removed every record, but a write still in flight lands afterwards.
        // The owner stays locked until this rollback settled (see dispose).
        for (const blob of written) {
          page.payload.nums[blob.record] = blob.ordinal
          if (blob.counted) { this.counters.blobs--; this.counters.blobBytes -= blob.bytes }
          await this.backend.delete(blob.key).catch(() => {})
        }
        if (pageWritten) await this.backend.delete(pageKey).catch(() => {})
      }
      throw error
    } finally {
      scratch.release()
      this.inFlight.delete(pageKey)
      for (const blob of written) this.inFlight.delete(blob.key)
    }
    this.counters.committed++
    this.counters.storedBytes += bytes.byteLength
  }
  /**
   * Stores the page's oversized values as blob records before the page itself,
   * so a committed page never references a missing record, and rewrites each
   * F_BLOB record's ordinal to its document-wide blob id. Each record holds one
   * value, bounded by the tokenizer's maxTokenChars, and is charged to the
   * storage-backed blobRecords pool; abandon/close remove them with the pages.
   */
  private async commitBlobs(p: CellPagePayload, blobs: readonly string[], pageId: string, docId: string, scratch: ConstructionShare, written: WrittenBlob[]): Promise<void> {
    const ids = new Map<number, number>()
    for (let i = 0; i < p.flags.length; i++) {
      if ((p.flags[i] & F_BLOB) === 0) continue
      const ordinal = p.nums[i]
      const text = blobs[ordinal]
      if (text === undefined || ids.has(ordinal)) throw new PackageError('parse-error', `page ${pageId} references a missing oversized value`)
      const id = this.blobCount++
      const bytes = encodeStrings([text], (n) => scratch.need(n))
      // Tracked before the write: a failed put can still leave a partial record (an OPFS file).
      const blob: WrittenBlob = { key: `blob/${id}`, record: i, ordinal, bytes: bytes.byteLength, counted: false, reservation: undefined, ram: () => {} }
      written.push(blob)
      this.inFlight.add(blob.key)
      blob.reservation = this.charge('blobRecords', bytes.byteLength)
      blob.ram = this.holdRam(bytes.byteLength)
      try {
        await this.backend.put(blob.key, bytes)
      } catch (error) {
        throw storageError(error, 'writing an oversized cell value')
      }
      this.check(docId)
      this.counters.blobs++
      this.counters.blobBytes += blob.bytes
      blob.counted = true
      ids.set(ordinal, id)
      p.nums[i] = id
    }
    if (ids.size !== blobs.length) throw new PackageError('parse-error', `page ${pageId} carries unreferenced oversized values`)
  }
  /** Removes a failed page's blob records and releases their charges. A record
   * the backend cannot delete keeps its charge: the ledger matches storage. */
  private async discardBlobs(p: CellPagePayload, written: readonly WrittenBlob[]): Promise<void> {
    for (const blob of written) {
      p.nums[blob.record] = blob.ordinal
      try { await this.backend.delete(blob.key) } catch { continue }
      this.discardedBlobs.add(Number(blob.key.slice(5)))
      this.uncharge(blob.reservation)
      blob.ram()
      if (blob.counted) {
        this.counters.blobs--
        this.counters.blobBytes -= blob.bytes
      }
    }
  }
  /** Shared strings, in fixed-size chunks addressed by id. */
  async commitStrings(chunk: number, strings: readonly string[], chunkSize: number, docId = this.docId): Promise<void> {
    this.check(docId)
    // Like pages: a chunk is never overwritten, so a rejected one cannot corrupt a stored one.
    if (!Number.isSafeInteger(chunk) || chunk < 0 || this.committedChunks.has(chunk)) throw new PackageError('parse-error', `shared-string chunk ${chunk} is invalid or already committed`)
    this.committedChunks.add(chunk)
    await this.track(this.commitChunk(chunk, strings, chunkSize, docId))
  }
  private async commitChunk(chunk: number, strings: readonly string[], chunkSize: number, docId: string): Promise<void> {
    const key = `sst/${chunk}`
    const scratch = this.construction.share('shared-string encoding')
    let stored = 0
    let charge: Reservation | undefined
    let ram: () => void = () => {}
    let attempted = false
    try {
      const bytes = encodeStrings(strings, (n) => scratch.need(n))
      charge = this.charge('sharedStrings', bytes.byteLength)
      ram = this.holdRam(bytes.byteLength)
      attempted = true
      this.inFlight.add(key)
      await this.backend.put(key, bytes)
      stored = bytes.byteLength
      this.check(docId)
    } catch (error) {
      if (this.live) {
        const removed = !attempted || await this.backend.delete(key).then(() => true, () => false)
        if (removed) { this.uncharge(charge); ram(); this.committedChunks.delete(chunk) }
      } else if (attempted) await this.backend.delete(key).catch(() => {})
      throw error
    } finally {
      scratch.release()
      this.inFlight.delete(key)
    }
    this.counters.storedBytes += stored
    this.stringChunkSize = chunkSize
    this.stringCount = Math.max(this.stringCount, chunk * chunkSize + strings.length)
  }
  get sharedStringCount(): number { return this.stringCount }

  /** Whether `key` belongs to this store's committed index or to a write in flight. */
  private referenced(key: string): boolean {
    if (this.inFlight.has(key)) return true
    const parts = key.split('/')
    if (parts.length === 3 && parts[0] === 'page') {
      const sheetId = unkeySegment(parts[1]), pageId = unkeySegment(parts[2])
      return pageRecordKey(sheetId, pageId) === key && (this.index.get(sheetId) ?? []).some((summary) => summary.pageId === pageId)
    }
    if (parts.length === 2 && (parts[0] === 'blob' || parts[0] === 'sst') && /^(0|[1-9][0-9]*)$/.test(parts[1])) {
      const n = Number(parts[1])
      return parts[0] === 'sst' ? this.committedChunks.has(n) : n < this.blobCount && !this.discardedBlobs.has(n)
    }
    return false
  }
  /**
   * Store-open orphan sweep: deletes records this owner holds that the
   * committed index does not reference. Only within an owner this store
   * exclusively holds: its Web Lock for OPFS/IndexedDB, or a memory backend
   * (private to this realm); otherwise another writer could share the owner.
   * Bounded (SWEEP_MAX_ENTRIES examined), cancellable, best effort: a listing
   * failure skips the sweep (`skipped`), never failing the import. The listing
   * is a transient charge and nothing removed was charged, so no charge remains.
   */
  sweepUnreferenced(signal?: AbortSignal): Promise<{ removed: number; skipped?: string }> {
    if (!this.backend.list || !this.live) return Promise.resolve({ removed: 0 })
    if (this.backend.kind !== 'memory' && !this.ownerLock) return Promise.resolve({ removed: 0, skipped: 'the store does not hold its owner lock' })
    // Its deletes are by key too: the owner stays locked until they settled (see dispose).
    return this.track(this.sweepOwned(this.backend.list.bind(this.backend), signal))
  }
  private async sweepOwned(list: NonNullable<PageBackend['list']>, signal?: AbortSignal): Promise<{ removed: number; skipped?: string }> {
    const share = this.construction.share('orphan sweep listing')
    let removed = 0
    try {
      try { share.need(SWEEP_MAX_ENTRIES * 128) } catch { return { removed: 0, skipped: 'no memory for the listing' } }
      let keys: string[]
      try {
        keys = await list(SWEEP_MAX_ENTRIES, signal)
      } catch (error) {
        return { removed: 0, skipped: `listing failed: ${error instanceof Error ? error.message : String(error)}` }
      }
      for (const key of keys) {
        if (signal?.aborted || !this.live) break
        if (this.referenced(key)) continue
        try { await this.backend.delete(key); removed++ } catch { /* left for the next sweep */ }
      }
    } finally {
      share.release()
    }
    return { removed }
  }

  /**
   * The once-per-process leftover sweep (sweepLeftoverPages), detached from
   * the import: its owner grouping is construction memory of this store (§12.4).
   * Closing/abandoning the store stops it at its next check (before each lock,
   * removal and deletion) but its charge lasts until it returned, since the
   * grouping stays reachable until then. Never rejects.
   */
  sweepLeftovers(env: PageStorageEnv, options: { now?: number; maxAgeMs?: number } = {}): Promise<number> {
    const share = this.construction.share('leftover page sweep', { untilReleased: true })
    return sweepLeftoverPages(env, { ...options, reserve: (bytes) => share.need(bytes) }).catch(() => 0).finally(() => share.release())
  }
  async abandon(docId: string): Promise<void> {
    if (docId !== this.docId) return
    this.abandoned = true
    await this.dispose()
  }
  async close(): Promise<void> {
    this.closed = true
    await this.dispose()
  }
  /** Once per store: overlapping and repeated close/abandon share one disposal. */
  private dispose(): Promise<void> {
    return this.disposing ??= this.disposeOnce()
  }
  /**
   * Releases every charge and removes the owner's storage. A write started
   * before (an OPFS/IndexedDB put cannot be withdrawn) may still land, and its
   * rollback deletes the record by key: until both settled the owner must not
   * pass to anyone else, or a recovering holder of the same owner could lose a
   * new record to the old rollback. So close/abandon return once storage is
   * removed (cancellation never waits on writes already in flight), but the owner's lock is
   * kept until every started commit settled and storage is removed once more;
   * `ownerReleased` resolves then. A backend that never completes keeps the
   * owner locked until its realm ends, as a crash would.
   */
  private async disposeOnce(): Promise<void> {
    for (const r of this.reservations.splice(0)) r.release()
    this.construction.releaseAll()
    this.ramBytes = 0
    for (const res of this.resident.values()) res.entry?.release()
    this.resident.clear()
    this.loading.clear()
    this.index.clear()
    this.counters.residentBytes = 0
    const release = () => {
      // Only now may a sweep in another realm (or a recovering store) take the owner.
      const unlock = this.ownerLock
      this.ownerLock = undefined
      unlock?.()
      this.released()
    }
    // No new commit starts once closed/abandoned, so this set only shrinks.
    const started = [...this.operations]
    try {
      await this.backend.removeAll()
    } finally {
      if (started.length === 0) release()
      else void Promise.all(started).then(() => this.backend.removeAll()).catch(() => {}).finally(release)
    }
  }

  pages(sheetId: string): readonly PageSummary[] { return this.index.get(sheetId) ?? [] }

  /** Registers decoded pages in the document pool's `source` tier from now on. */
  bindMemory(pool: MemoryPool): void {
    // Atomic: throws the typed memory-budget refusal and changes nothing when
    // the pool cannot hold the construction charges and memory-held records.
    this.construction.rebind(pool)
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
  /**
   * `size` is told whether the decoded value views the very buffer a memory
   * backend retains (already charged by holdRam), so only the rest is charged.
   */
  private async load<T extends CellPagePayload | StringChunk>(key: string, decode: (bytes: Uint8Array) => T, size: (value: T, sharesHeldBuffer: boolean) => number): Promise<T> {
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
        const view = 'rows' in value ? value.rows : value.text
        const bytesHeld = size(value, this.backend.kind === 'memory' && view.buffer === bytes.buffer)
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
    return this.load(pageRecordKey(sheetId, pageId), decodePage, (p, shared) => DECODED_PAGE_OVERHEAD_BYTES + (shared ? metaBytes(p) : payloadBytes(p)))
  }
  /** An oversized value by blob id, through the same bounded decoded cache as pages. */
  async blob(id: number): Promise<string> {
    if (!Number.isSafeInteger(id) || id < 0 || id >= this.blobCount) throw new PackageError('storage-error', `oversized cell value ${id} is not stored`)
    const record = await this.load(`blob/${id}`, decodeStrings, (c, shared) => DECODED_CHUNK_OVERHEAD_BYTES + (shared ? 0 : c.text.byteLength + c.textStart.byteLength))
    if (record.textStart.length !== 2) throw new PackageError('parse-error', `oversized cell value ${id} is malformed`)
    return textAt(record, 0)
  }
  async sharedString(id: number): Promise<string | undefined> {
    if (!Number.isSafeInteger(id) || id < 0 || id >= this.stringCount || this.stringChunkSize === 0) return undefined
    const chunk = Math.floor(id / this.stringChunkSize)
    const strings = await this.load(`sst/${chunk}`, decodeStrings, (c, shared) => DECODED_CHUNK_OVERHEAD_BYTES + (shared ? 0 : c.text.byteLength + c.textStart.byteLength))
    return textAt(strings, id - chunk * this.stringChunkSize)
  }
}

export function failureCode(error: unknown): PackageErrorCode {
  return error instanceof PackageError ? error.code : 'storage-error'
}
