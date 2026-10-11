/**
 * Bounded spooling of sequential (network/stream) spreadsheet input to a
 * seekable Blob (design note §2.8, §4.2.1).
 *
 * An OOXML central directory sits at the end of the archive, so a download
 * must land somewhere seekable before it can be admitted. Backends, in
 * order: memory under a threshold, OPFS, IndexedDB chunk blobs, memory up to
 * a hard cap, otherwise a typed failure. The advertised length is only a hint
 * for the starting backend: every append re-checks the actual running total.
 * No truncated Blob is ever published and there is no unbounded RAM fallback.
 */
import { PackageError, throwIfAborted, type PackageErrorCode } from './package-error'

export type SpoolKind = 'memory' | 'opfs' | 'indexeddb'

/** The spooled bytes plus explicit cleanup ownership. */
export interface SpooledSource {
  readonly blob: Blob
  /** Actual bytes accepted, never the advertised length. */
  readonly size: number
  readonly kind: SpoolKind
  /** Idempotent; removes the OPFS file / IDB database / memory chunks. */
  release(): Promise<void>
}

export interface SpoolSink {
  readonly kind: SpoolKind
  /** Resolves when the chunk is durably accepted: the backpressure point. */
  append(chunk: Uint8Array, signal?: AbortSignal): Promise<void>
  /** Only on a complete download; the returned handle owns cleanup from here on. */
  finish(): Promise<SpooledSource>
  /** Failure/cancel before finish(): removes every temporary byte. Idempotent. */
  abort(): Promise<void>
}

export interface SpoolBackend {
  available(): boolean
  create(name: string): Promise<SpoolSink>
  /** Entry names, for the leftover sweep. */
  list?(): Promise<string[]>
  remove?(name: string): Promise<void>
}

/** Injectable, so tests never touch a real browser. */
export interface StorageEnv {
  /** Proposed 32 MiB D / 16 MiB M (§4.3). */
  memoryLimitBytes: number
  /** Default 2 x memoryLimitBytes. */
  hardMemoryCapBytes?: number
  opfs?: SpoolBackend
  indexedDb?: SpoolBackend
  estimate?(): Promise<{ quota: number; usage: number }>
  now?(): number
}

export interface SpoolOptions {
  /** The paged-path compressed cap (256 MiB D / 128 MiB M). */
  maxBytes: number
  /** Content-Length or a remote Blob size: an untrusted hint. */
  advertisedBytes?: number
}

export const SPOOL_PREFIX = 'officeview-spool-'
const MiB = 1024 * 1024
const MIN_FREE_BYTES = 64 * MiB
const LEFTOVER_AGE_MS = 24 * 3600_000

function storageFailure(error: unknown, action: string): PackageError {
  if (error instanceof PackageError) return error
  const name = (error as { name?: unknown } | null)?.name
  if (name === 'QuotaExceededError') return new PackageError('quota', `browser storage quota exceeded while ${action}`)
  const detail = error instanceof Error ? `: ${error.message}` : ''
  return new PackageError('storage-error', `browser storage failed while ${action}${detail}`)
}

function spoolName(env: StorageEnv): string {
  const now = env.now?.() ?? Date.now()
  return `${SPOOL_PREFIX}${now}-${Math.random().toString(36).slice(2, 10)}`
}

type Pick = { sink: SpoolSink } | { code: Extract<PackageErrorCode, 'quota' | 'storage-error'>; reason: string }

async function pickPersistent(env: StorageEnv, needBytes: number): Promise<Pick> {
  const backends = [env.opfs, env.indexedDb].filter((b): b is SpoolBackend => b !== undefined && safeAvailable(b))
  if (backends.length === 0) return { code: 'storage-error', reason: 'no browser storage (OPFS or IndexedDB) is available' }
  if (env.estimate) {
    let estimate: { quota: number; usage: number } | undefined
    try { estimate = await env.estimate() } catch { estimate = undefined }
    // OPFS and IndexedDB share the origin quota, so one check covers both.
    if (estimate && estimate.quota - estimate.usage < needBytes) {
      return { code: 'quota', reason: `browser storage has ${Math.max(0, estimate.quota - estimate.usage)} bytes free, ${needBytes} needed` }
    }
  }
  let lastError: unknown
  for (const backend of backends) {
    try {
      return { sink: await backend.create(spoolName(env)) }
    } catch (error) {
      lastError = error
    }
  }
  return { code: 'storage-error', reason: storageFailure(lastError, 'creating a spool').message }
}

function safeAvailable(backend: SpoolBackend): boolean {
  try { return backend.available() } catch { return false }
}

function checkLimit(value: number, label: string): void {
  if (!Number.isSafeInteger(value) || value < 0) throw new RangeError(`${label} must be a non-negative safe integer`)
}

/** Open a spool for one download. The returned sink owns every temporary byte until finish(). */
export async function openSpool(env: StorageEnv, opts: SpoolOptions, signal: AbortSignal): Promise<SpoolSink> {
  checkLimit(opts.maxBytes, 'maxBytes')
  checkLimit(env.memoryLimitBytes, 'memoryLimitBytes')
  const hardCap = env.hardMemoryCapBytes ?? 2 * env.memoryLimitBytes
  throwIfAborted(signal)
  let backend: SpoolSink | undefined
  // Set once no persistent backend can be used: memory continues up to the hard cap.
  let memoryOnly: { code: 'quota' | 'storage-error'; reason: string } | undefined
  const advertised = opts.advertisedBytes !== undefined && Number.isSafeInteger(opts.advertisedBytes) && opts.advertisedBytes >= 0
    ? opts.advertisedBytes
    : undefined
  if (advertised !== undefined && advertised > env.memoryLimitBytes) {
    const pick = await pickPersistent(env, Math.max(2 * advertised, MIN_FREE_BYTES))
    if ('sink' in pick) backend = pick.sink
    else memoryOnly = pick
  }
  throwIfAborted(signal)

  let chunks: Uint8Array[] = []
  let total = 0
  let failed: PackageError | undefined
  let finished = false
  let aborted = false

  const discard = async (): Promise<void> => {
    chunks = []
    if (aborted) return
    aborted = true
    if (backend) {
      try { await backend.abort() } catch { /* best effort: the sweep removes leftovers */ }
    }
  }
  const fail = async (error: PackageError): Promise<never> => {
    failed ??= error
    await discard()
    throw failed
  }

  const sink: SpoolSink = {
    get kind() { return backend?.kind ?? 'memory' },
    async append(chunk, appendSignal) {
      if (failed) throw failed
      if (finished) throw new PackageError('parse-error', 'append after finish')
      if (signal.aborted || appendSignal?.aborted) await fail(new PackageError('cancelled', 'download cancelled'))
      const next = total + chunk.length
      if (next > opts.maxBytes) {
        await fail(new PackageError('resource-limit', `download exceeds the ${opts.maxBytes}-byte limit for spreadsheets`))
      }
      if (!backend && memoryOnly === undefined && next > env.memoryLimitBytes) {
        // Promote once; the copy is at most memoryLimitBytes.
        const pick = await pickPersistent(env, Math.max(2 * Math.max(advertised ?? 0, next), MIN_FREE_BYTES))
        if ('sink' in pick) {
          backend = pick.sink
          const held = chunks
          chunks = []
          try {
            for (const c of held) await backend.append(c)
          } catch (error) {
            await fail(storageFailure(error, 'writing the spool'))
          }
        } else {
          memoryOnly = pick
        }
      }
      if (backend) {
        try {
          await backend.append(chunk)
        } catch (error) {
          await fail(storageFailure(error, 'writing the spool'))
        }
      } else {
        if (next > hardCap) {
          const why = memoryOnly ?? { code: 'storage-error' as const, reason: 'no browser storage is available' }
          await fail(new PackageError(why.code, `the download is larger than ${hardCap} bytes and cannot be stored: ${why.reason}`))
        }
        chunks.push(chunk.slice())
      }
      total = next
      if (signal.aborted || appendSignal?.aborted) await fail(new PackageError('cancelled', 'download cancelled'))
    },
    async finish() {
      if (failed) throw failed
      if (finished) throw new PackageError('parse-error', 'spool already finished')
      if (signal.aborted) await fail(new PackageError('cancelled', 'download cancelled'))
      finished = true
      if (!backend) {
        const blob = new Blob(chunks as BlobPart[])
        chunks = []
        let held: Blob | undefined = blob
        return {
          get blob() {
            if (!held) throw new PackageError('cancelled', 'spooled source already released')
            return held
          },
          size: total,
          kind: 'memory',
          async release() { held = undefined },
        }
      }
      let source: SpooledSource
      try {
        source = await backend.finish()
      } catch (error) {
        finished = false
        return fail(storageFailure(error, 'finishing the spool'))
      }
      if (source.blob.size !== total) {
        await source.release().catch(() => {})
        failed = new PackageError('storage-error', `spool returned ${source.blob.size} of ${total} bytes`)
        throw failed
      }
      let released = false
      return {
        blob: source.blob,
        size: total,
        kind: source.kind,
        async release() {
          if (released) return
          // marked only after success, so a failed removal can be retried
          await source.release()
          released = true
        },
      }
    },
    async abort() {
      if (finished) return
      failed ??= new PackageError('cancelled', 'spool aborted')
      await discard()
    },
  }
  return sink
}

/** Best-effort removal of spools left by a crashed or killed session. */
export async function sweepLeftoverSpools(env: StorageEnv, maxAgeMs = LEFTOVER_AGE_MS): Promise<number> {
  const now = env.now?.() ?? Date.now()
  let removed = 0
  for (const backend of [env.opfs, env.indexedDb]) {
    if (!backend?.list || !backend.remove || !safeAvailable(backend)) continue
    let names: string[]
    try { names = await backend.list() } catch { continue }
    for (const name of names) {
      if (!name.startsWith(SPOOL_PREFIX)) continue
      const created = Number(name.slice(SPOOL_PREFIX.length).split('-')[0])
      if (!Number.isFinite(created) || now - created <= maxAgeMs) continue
      try {
        await backend.remove(name)
        removed++
      } catch { /* still in use or already gone */ }
    }
  }
  return removed
}

interface WritableLike { write(data: Uint8Array): Promise<void>; close(): Promise<void>; abort?(): Promise<void> }
interface FileHandleLike { createWritable(): Promise<WritableLike>; getFile(): Promise<Blob> }
interface DirectoryLike {
  getFileHandle(name: string, options?: { create?: boolean }): Promise<FileHandleLike>
  removeEntry(name: string): Promise<void>
  keys?(): AsyncIterable<string>
}

/** OPFS through FileSystemWritableFileStream. getFile() is a file-backed File. */
export function createOpfsBackend(getDirectory: () => Promise<FileSystemDirectoryHandle>, available: () => boolean = () => true): SpoolBackend {
  const dir = async () => (await getDirectory()) as unknown as DirectoryLike
  const remove = async (name: string): Promise<void> => {
    try {
      await (await dir()).removeEntry(name)
    } catch (error) {
      if ((error as { name?: unknown } | null)?.name !== 'NotFoundError') throw error
    }
  }
  return {
    available,
    async create(name) {
      const handle = await (await dir()).getFileHandle(name, { create: true })
      let writable: WritableLike
      try {
        writable = await handle.createWritable()
      } catch (error) {
        await remove(name).catch(() => {})
        throw error
      }
      let closed = false
      return {
        kind: 'opfs',
        append: (chunk) => writable.write(chunk),
        async finish() {
          await writable.close()
          closed = true
          const blob = await handle.getFile()
          return { blob, size: blob.size, kind: 'opfs', release: () => remove(name) }
        },
        async abort() {
          if (!closed) {
            closed = true
            try { await writable.abort?.() } catch { /* already failed */ }
          }
          await remove(name)
        },
      }
    },
    async list() {
      const names: string[] = []
      const d = await dir()
      if (!d.keys) return names
      for await (const name of d.keys()) names.push(name)
      return names
    },
    remove,
  }
}

function settle<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error)
  })
}

function completion(tx: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve()
    tx.onerror = () => reject(tx.error)
    tx.onabort = () => reject(tx.error ?? new DOMException('transaction aborted', 'AbortError'))
  })
}

const STORE = 'chunks'
const IDB_BATCH_BYTES = 8 * MiB

/**
 * IndexedDB chunk blobs. Chunks are collected (at most batchBytes) and then
 * written with every put() issued synchronously inside ONE transaction:
 * awaiting a stream read between put() calls lets the transaction auto-commit
 * (measured in Chromium, §2.8).
 */
export function createIndexedDbBackend(factory: IDBFactory, options: { batchBytes?: number; available?: () => boolean } = {}): SpoolBackend {
  const batchBytes = options.batchBytes ?? IDB_BATCH_BYTES
  const deleteDb = async (name: string): Promise<void> => {
    await new Promise<void>((resolve, reject) => {
      const req = factory.deleteDatabase(name)
      req.onsuccess = () => resolve()
      req.onerror = () => reject(req.error)
    })
  }
  return {
    available: options.available ?? (() => true),
    async create(name) {
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const req = factory.open(name, 1)
        req.onupgradeneeded = () => {
          const d = req.result
          if (!d.objectStoreNames.contains(STORE)) d.createObjectStore(STORE)
        }
        req.onsuccess = () => resolve(req.result)
        req.onerror = () => reject(req.error)
      })
      let pending: Uint8Array[] = []
      let pendingBytes = 0
      let seq = 0
      let closed = false
      const close = (): void => {
        if (closed) return
        closed = true
        db.close()
      }
      const flush = async (): Promise<void> => {
        if (pending.length === 0) return
        const batch = pending
        pending = []
        pendingBytes = 0
        const tx = db.transaction(STORE, 'readwrite')
        const done = completion(tx)
        const store = tx.objectStore(STORE)
        for (const chunk of batch) store.put(new Blob([chunk as BlobPart]), seq++)
        await done
      }
      return {
        kind: 'indexeddb',
        async append(chunk) {
          pending.push(chunk.slice())
          pendingBytes += chunk.length
          if (pendingBytes >= batchBytes) await flush()
        },
        async finish() {
          await flush()
          const tx = db.transaction(STORE, 'readonly')
          const parts = await settle(tx.objectStore(STORE).getAll() as IDBRequest<Blob[]>)
          close()
          const blob = new Blob(parts)
          let released = false
          return {
            blob,
            size: blob.size,
            kind: 'indexeddb',
            async release() {
              if (released) return
              await deleteDb(name)
              released = true
            },
          }
        },
        async abort() {
          pending = []
          close()
          await deleteDb(name)
        },
      }
    },
    async list() {
      const listing = (factory as IDBFactory & { databases?: () => Promise<Array<{ name?: string }>> }).databases
      if (typeof listing !== 'function') return []
      return (await listing.call(factory)).map((d) => d.name ?? '').filter((n) => n !== '')
    },
    remove: deleteDb,
  }
}

/** Run-time feature probe; never assumes a backend (Safari/Firefox unmeasured, §2.8). */
export function browserStorageEnv(memoryLimitBytes: number): StorageEnv {
  const nav = typeof navigator === 'undefined' ? undefined : navigator as Navigator & { storage?: StorageManager }
  const storage = nav?.storage
  const env: StorageEnv = { memoryLimitBytes }
  if (storage && typeof storage.getDirectory === 'function') {
    const writable = typeof FileSystemFileHandle !== 'undefined' && 'createWritable' in FileSystemFileHandle.prototype
    env.opfs = createOpfsBackend(() => storage.getDirectory(), () => writable)
  }
  if (typeof indexedDB !== 'undefined') env.indexedDb = createIndexedDbBackend(indexedDB)
  if (storage && typeof storage.estimate === 'function') {
    env.estimate = async () => {
      const e = await storage.estimate()
      return { quota: e.quota ?? 0, usage: e.usage ?? 0 }
    }
  }
  return env
}
