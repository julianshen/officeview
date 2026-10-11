/** Fake StorageEnv backends for the spool tests (jsdom has no OPFS or
 * IndexedDB). Each fake records what is stored so a test can assert that no
 * temporary byte survives a failure, cancellation or release. */
import type { SpoolBackend, SpoolSink, SpooledSource, StorageEnv } from '../../src/core/spool'

export interface FakeBackendLog {
  created: string[]
  removed: string[]
  /** Names whose bytes are still stored. */
  live: Map<string, Uint8Array[]>
  writes: number
}

export interface FakeBackendOptions {
  available?: boolean
  /** Throw this from the n-th write (1-based). */
  failOnWrite?: { at: number; error: unknown }
  /** Called before each write resolves; await it to hold the write open. */
  gate?: () => Promise<void>
  createError?: unknown
  /** Names already present (leftovers from an earlier session). */
  existing?: string[]
}

export function fakeBackend(kind: 'opfs' | 'indexeddb', options: FakeBackendOptions = {}): SpoolBackend & { log: FakeBackendLog } {
  const log: FakeBackendLog = { created: [], removed: [], live: new Map(), writes: 0 }
  for (const name of options.existing ?? []) log.live.set(name, [])
  const remove = (name: string): void => {
    if (log.live.delete(name)) log.removed.push(name)
  }
  return {
    log,
    available: () => options.available !== false,
    async create(name) {
      if (options.createError !== undefined) throw options.createError
      log.created.push(name)
      const chunks: Uint8Array[] = []
      log.live.set(name, chunks)
      let size = 0
      const sink: SpoolSink = {
        kind,
        async append(chunk) {
          log.writes++
          if (options.failOnWrite && log.writes === options.failOnWrite.at) throw options.failOnWrite.error
          if (options.gate) await options.gate()
          chunks.push(chunk.slice())
          size += chunk.length
        },
        async finish(): Promise<SpooledSource> {
          const blob = new Blob(chunks as BlobPart[])
          return { blob, size, kind, release: async () => remove(name) }
        },
        async abort() {
          remove(name)
        },
      }
      return sink
    },
    async list() {
      return [...log.live.keys()]
    },
    async remove(name) {
      remove(name)
    },
  }
}

export function fakeEnv(overrides: Partial<StorageEnv> = {}): StorageEnv {
  return { memoryLimitBytes: 1024, ...overrides }
}

/**
 * Minimal IndexedDB double with auto-commit: a transaction is active only
 * until the current macrotask ends (as a browser transaction is once its
 * creating task returns to the event loop with no request pending), so a
 * sink that awaits a stream read between put() calls hits
 * TransactionInactiveError exactly as Chromium did in the §2.8 measurement.
 */
export function fakeIndexedDb(): IDBFactory & { stores: Map<string, Map<number, Blob>>; transactions: Array<{ mode: string; puts: number }>; deleted: string[] } {
  const stores = new Map<string, Map<number, Blob>>()
  const transactions: Array<{ mode: string; puts: number }> = []
  const deleted: string[] = []
  const later = (fn: () => void) => setTimeout(fn, 0)
  function request<T>(run: () => T): IDBRequest<T> {
    const req = { result: undefined as T | undefined, error: null, onsuccess: null as null | ((e: unknown) => void), onerror: null as null | ((e: unknown) => void) }
    later(() => {
      try {
        req.result = run()
        req.onsuccess?.({ target: req })
      } catch (error) {
        ;(req as { error: unknown }).error = error
        req.onerror?.({ target: req })
      }
    })
    return req as unknown as IDBRequest<T>
  }
  function database(name: string): IDBDatabase {
    const data = stores.get(name)!
    let closed = false
    return {
      name,
      objectStoreNames: { contains: (s: string) => s === 'chunks' },
      createObjectStore: () => ({}),
      close: () => { closed = true },
      transaction(_store: string, mode: IDBTransactionMode = 'readonly') {
        if (closed) throw new DOMException('closed', 'InvalidStateError')
        const record = { mode, puts: 0 }
        transactions.push(record)
        let active = true
        const tx = {
          error: null,
          oncomplete: null as null | (() => void),
          onerror: null as null | (() => void),
          onabort: null as null | (() => void),
          objectStore: () => ({
            put(value: Blob, key: number) {
              if (!active) throw new DOMException('transaction committed', 'TransactionInactiveError')
              record.puts++
              data.set(key, value)
              return request(() => key)
            },
            getAll() {
              if (!active) throw new DOMException('transaction committed', 'TransactionInactiveError')
              return request(() => [...data.keys()].sort((a, b) => a - b).map((k) => data.get(k)!))
            },
          }),
        }
        later(() => {
          active = false
          later(() => tx.oncomplete?.())
        })
        return tx as unknown as IDBTransaction
      },
    } as unknown as IDBDatabase
  }
  const factory = {
    stores,
    transactions,
    deleted,
    open(name: string) {
      const req = {
        result: undefined as IDBDatabase | undefined, error: null,
        onsuccess: null as null | (() => void), onerror: null as null | (() => void),
        onupgradeneeded: null as null | (() => void), onblocked: null,
      }
      later(() => {
        const fresh = !stores.has(name)
        if (fresh) stores.set(name, new Map())
        req.result = database(name)
        if (fresh) req.onupgradeneeded?.()
        req.onsuccess?.()
      })
      return req as unknown as IDBOpenDBRequest
    },
    deleteDatabase(name: string) {
      const req = { onsuccess: null as null | (() => void), onerror: null, onblocked: null }
      later(() => {
        if (stores.delete(name)) deleted.push(name)
        req.onsuccess?.()
      })
      return req as unknown as IDBOpenDBRequest
    },
    async databases() {
      return [...stores.keys()].map((name) => ({ name, version: 1 }))
    },
    cmp: () => 0,
  }
  return factory as unknown as ReturnType<typeof fakeIndexedDb>
}

export interface FakeOpfsDir {
  files: Map<string, Uint8Array[]>
  removed: string[]
  getFileHandle(name: string, opts?: { create?: boolean }): Promise<unknown>
  removeEntry(name: string): Promise<void>
  keys(): AsyncIterableIterator<string>
}

export function fakeOpfsDirectory(options: { failOnWrite?: { at: number; error: unknown } } = {}): FakeOpfsDir {
  const files = new Map<string, Uint8Array[]>()
  const removed: string[] = []
  let writes = 0
  return {
    files,
    removed,
    async getFileHandle(name, opts) {
      if (!files.has(name)) {
        if (!opts?.create) throw new DOMException('missing', 'NotFoundError')
        files.set(name, [])
      }
      return {
        async createWritable() {
          const staged: Uint8Array[] = []
          return {
            async write(chunk: Uint8Array) {
              writes++
              if (options.failOnWrite && writes === options.failOnWrite.at) throw options.failOnWrite.error
              staged.push(chunk.slice())
            },
            async close() { files.set(name, staged) },
            async abort() { staged.length = 0 },
          }
        },
        async getFile() { return new Blob((files.get(name) ?? []) as BlobPart[]) },
      }
    },
    async removeEntry(name) {
      if (!files.delete(name)) throw new DOMException('missing', 'NotFoundError')
      removed.push(name)
    },
    async *keys() {
      for (const name of [...files.keys()]) yield name
    },
  }
}
