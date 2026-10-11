/** S1.1b network spool (design note §2.8, §4.2.1; RED tests 15, 15a, 16, 17). */
import { describe, expect, it } from 'vitest'
import {
  createIndexedDbBackend, createOpfsBackend, openSpool, sweepLeftoverSpools, SPOOL_PREFIX,
  type SpoolSink, type SpooledSource, type StorageEnv,
} from '../src/core/spool'
import { PackageError } from '../src/core/package-error'
import { fakeBackend, fakeEnv, fakeIndexedDb, fakeOpfsDirectory } from './helpers/fake-storage'

const KiB = 1024
const MiB = 1024 * KiB
const signal = () => new AbortController().signal

function pattern(n: number, seed = 1): Uint8Array {
  const out = new Uint8Array(n)
  for (let i = 0, x = seed; i < n; i++) { x = (x * 1103515245 + 12345) >>> 0; out[i] = x >>> 24 }
  return out
}

async function feed(sink: SpoolSink, bytes: Uint8Array, chunk: number): Promise<void> {
  for (let at = 0; at < bytes.length; at += chunk) await sink.append(bytes.subarray(at, at + chunk))
}

async function blobBytes(source: SpooledSource): Promise<Uint8Array> {
  return new Uint8Array(await source.blob.arrayBuffer())
}

function same(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false
  return true
}

async function failure(p: Promise<unknown>): Promise<PackageError> {
  try { await p } catch (e) { if (e instanceof PackageError) return e; throw e }
  throw new Error('expected a typed failure')
}

describe('backend selection (RED test 15)', () => {
  it('a known small length stays in memory and the bytes come back losslessly', async () => {
    const opfs = fakeBackend('opfs')
    const sink = await openSpool(fakeEnv({ memoryLimitBytes: 4 * KiB, opfs }), { maxBytes: MiB, advertisedBytes: 3000 }, signal())
    const bytes = pattern(3000)
    await feed(sink, bytes, 700)
    const src = await sink.finish()
    expect(src.kind).toBe('memory')
    expect(src.size).toBe(3000)
    expect(same(await blobBytes(src), bytes)).toBe(true)
    expect(opfs.log.created).toEqual([])
  })

  it('an unknown length promotes to OPFS once at the memory limit, copying at most the limit', async () => {
    const opfs = fakeBackend('opfs')
    const sink = await openSpool(fakeEnv({ memoryLimitBytes: 4 * KiB, opfs, estimate: async () => ({ quota: GiB(1), usage: 0 }) }), { maxBytes: MiB }, signal())
    const bytes = pattern(20 * KiB, 3)
    await feed(sink, bytes, 1000)
    const src = await sink.finish()
    expect(src.kind).toBe('opfs')
    expect(opfs.log.created).toHaveLength(1)
    expect(same(await blobBytes(src), bytes)).toBe(true)
  })

  it('without OPFS the persistent backend is IndexedDB', async () => {
    const indexedDb = fakeBackend('indexeddb')
    const env = fakeEnv({ memoryLimitBytes: 4 * KiB, opfs: fakeBackend('opfs', { available: false }), indexedDb })
    const sink = await openSpool(env, { maxBytes: MiB }, signal())
    const bytes = pattern(10 * KiB, 5)
    await feed(sink, bytes, 1024)
    const src = await sink.finish()
    expect(src.kind).toBe('indexeddb')
    expect(same(await blobBytes(src), bytes)).toBe(true)
  })

  it('an advertised length above the memory limit starts in the persistent backend', async () => {
    const opfs = fakeBackend('opfs')
    const sink = await openSpool(fakeEnv({ memoryLimitBytes: 4 * KiB, opfs }), { maxBytes: MiB, advertisedBytes: 100 * KiB }, signal())
    expect(sink.kind).toBe('opfs')
    await sink.append(pattern(10))
    expect(opfs.log.writes).toBe(1)
  })

  it('no persistent backend: memory up to the hard cap, then a typed storage-error and nothing kept', async () => {
    const env = fakeEnv({ memoryLimitBytes: 4 * KiB })
    const ok = await openSpool(env, { maxBytes: MiB }, signal())
    await feed(ok, pattern(8 * KiB), 1024) // exactly the default hard cap (2 x limit)
    expect((await ok.finish()).size).toBe(8 * KiB)

    const sink = await openSpool(env, { maxBytes: MiB }, signal())
    const err = await failure(feed(sink, pattern(9 * KiB), 1024))
    expect(err.code).toBe('storage-error')
    expect(await failure(sink.finish())).toBe(err)
  })

  it('insufficient estimate() with backends present gives a typed quota failure above the hard cap', async () => {
    const opfs = fakeBackend('opfs')
    const indexedDb = fakeBackend('indexeddb')
    const env = fakeEnv({ memoryLimitBytes: 4 * KiB, opfs, indexedDb, estimate: async () => ({ quota: 64 * MiB, usage: 63 * MiB }) })
    const sink = await openSpool(env, { maxBytes: MiB }, signal())
    const err = await failure(feed(sink, pattern(9 * KiB), 1024))
    expect(err.code).toBe('quota')
    expect(opfs.log.created).toEqual([])
    expect(indexedDb.log.created).toEqual([])
  })

  it('a backend that cannot be created falls through to the next one', async () => {
    const opfs = fakeBackend('opfs', { createError: new DOMException('denied', 'SecurityError') })
    const indexedDb = fakeBackend('indexeddb')
    const sink = await openSpool(fakeEnv({ memoryLimitBytes: KiB, opfs, indexedDb }), { maxBytes: MiB }, signal())
    await feed(sink, pattern(4 * KiB), 512)
    expect((await sink.finish()).kind).toBe('indexeddb')
  })

  it('lying length: 1 KiB advertised, 40 MiB delivered, promoted by the actual count on the crossing append', async () => {
    const opfs = fakeBackend('opfs')
    const env = fakeEnv({ memoryLimitBytes: 4 * MiB, opfs })
    const sink = await openSpool(env, { maxBytes: 64 * MiB, advertisedBytes: KiB }, signal())
    expect(sink.kind).toBe('memory')
    const chunk = pattern(256 * KiB, 9)
    let appended = 0
    while (appended + chunk.length <= 4 * MiB) { await sink.append(chunk); appended += chunk.length }
    expect(opfs.log.created).toHaveLength(0)
    await sink.append(chunk)
    expect(opfs.log.created).toHaveLength(1)
    for (appended += chunk.length; appended < 40 * MiB; appended += chunk.length) await sink.append(chunk)
    const src = await sink.finish()
    expect(src.size).toBe(40 * MiB)
    expect(src.kind).toBe('opfs')
  })

  it('maxBytes is enforced on every append whatever the advertised length said, before accepting the chunk', async () => {
    const opfs = fakeBackend('opfs')
    const sink = await openSpool(fakeEnv({ memoryLimitBytes: 4 * MiB, opfs }), { maxBytes: 16 * MiB, advertisedBytes: KiB }, signal())
    const chunk = pattern(MiB, 11)
    let accepted = 0
    const err = await failure((async () => { for (let i = 0; i < 64; i++) { await sink.append(chunk); accepted += chunk.length } })())
    expect(err.code).toBe('resource-limit')
    expect(accepted).toBe(16 * MiB)
    expect(opfs.log.live.size).toBe(0)
    expect(opfs.log.removed).toHaveLength(1)
  })
})

describe('ownership and cleanup (RED test 15a)', () => {
  it('download failure before finish(): abort() removes every temporary byte', async () => {
    const opfs = fakeBackend('opfs')
    const sink = await openSpool(fakeEnv({ memoryLimitBytes: KiB, opfs }), { maxBytes: MiB }, signal())
    await feed(sink, pattern(5 * KiB), 1024)
    expect(opfs.log.live.size).toBe(1)
    await sink.abort()
    await sink.abort()
    expect(opfs.log.live.size).toBe(0)
    expect(opfs.log.removed).toHaveLength(1)
  })

  it('cancellation during append surfaces cancelled and removes the partial spool', async () => {
    const opfs = fakeBackend('opfs')
    const ac = new AbortController()
    const sink = await openSpool(fakeEnv({ memoryLimitBytes: KiB, opfs }), { maxBytes: MiB }, ac.signal)
    await feed(sink, pattern(3 * KiB), 1024)
    ac.abort()
    const err = await failure(sink.append(pattern(10)))
    expect(err.code).toBe('cancelled')
    expect(opfs.log.live.size).toBe(0)
  })

  it('finish() hands ownership to the SpooledSource; release() is idempotent', async () => {
    const opfs = fakeBackend('opfs')
    const sink = await openSpool(fakeEnv({ memoryLimitBytes: KiB, opfs }), { maxBytes: MiB }, signal())
    await feed(sink, pattern(4 * KiB), 1024)
    const src = await sink.finish()
    expect(opfs.log.live.size).toBe(1)
    await src.release()
    await src.release()
    expect(opfs.log.live.size).toBe(0)
    expect(opfs.log.removed).toHaveLength(1)
  })

  it('leftover spools older than 24 h are swept; fresh and foreign entries stay', async () => {
    const now = Date.UTC(2026, 9, 11)
    const old = `${SPOOL_PREFIX}${now - 25 * 3600_000}-a`
    const fresh = `${SPOOL_PREFIX}${now - 3600_000}-b`
    const opfs = fakeBackend('opfs', { existing: [old, fresh, 'unrelated.bin'] })
    const indexedDb = fakeBackend('indexeddb', { existing: [`${SPOOL_PREFIX}${now - 48 * 3600_000}-c`] })
    const removed = await sweepLeftoverSpools(fakeEnv({ opfs, indexedDb, now: () => now }))
    expect(removed).toBe(2)
    expect([...opfs.log.live.keys()].sort()).toEqual([fresh, 'unrelated.bin'].sort())
    expect(indexedDb.log.live.size).toBe(0)
  })
})

describe('backpressure (RED test 16)', () => {
  it('append resolves only once the backend accepted the chunk', async () => {
    let open: () => void = () => {}
    const opfs = fakeBackend('opfs', { gate: () => new Promise<void>((r) => { open = r }) })
    const sink = await openSpool(fakeEnv({ memoryLimitBytes: KiB, opfs }), { maxBytes: MiB, advertisedBytes: MiB }, signal())
    let done = false
    const pending = sink.append(pattern(100)).then(() => { done = true })
    await new Promise((r) => setTimeout(r, 5))
    expect(done).toBe(false)
    open()
    await pending
    expect(done).toBe(true)
  })

  it('IndexedDB chunks are batched into one transaction per 8 MiB, never awaited between put() calls', async () => {
    const idb = fakeIndexedDb()
    const backend = createIndexedDbBackend(idb)
    const sink = await backend.create(`${SPOOL_PREFIX}1-x`)
    const bytes = pattern(20 * MiB, 13)
    for (let at = 0; at < bytes.length; at += 256 * KiB) {
      await sink.append(bytes.subarray(at, at + 256 * KiB))
      await new Promise((r) => setTimeout(r, 0)) // a stream read between appends ends the task
    }
    const src = await sink.finish()
    expect(same(await blobBytes(src), bytes)).toBe(true)
    const writes = idb.transactions.filter((t) => t.mode === 'readwrite')
    expect(writes.map((t) => t.puts)).toEqual([32, 32, 16])
    await src.release()
    expect(idb.stores.size).toBe(0)
  })

  it('the OPFS backend writes through a writable and release() removes the file', async () => {
    const dir = fakeOpfsDirectory()
    const backend = createOpfsBackend(async () => dir as unknown as FileSystemDirectoryHandle)
    const name = `${SPOOL_PREFIX}1-y`
    const sink = await backend.create(name)
    const bytes = pattern(300 * KiB, 17)
    await feed(sink, bytes, 64 * KiB)
    const src = await sink.finish()
    expect(same(await blobBytes(src), bytes)).toBe(true)
    await src.release()
    await src.release()
    expect(dir.files.has(name)).toBe(false)
    expect(await backend.list!()).toEqual([])
  })
})

describe('storage failure mid-write (RED test 17)', () => {
  it('QuotaExceededError from the backend is a typed quota failure and the partial file is removed', async () => {
    const opfs = fakeBackend('opfs', { failOnWrite: { at: 3, error: new DOMException('full', 'QuotaExceededError') } })
    const sink = await openSpool(fakeEnv({ memoryLimitBytes: KiB, opfs }), { maxBytes: MiB, advertisedBytes: MiB }, signal())
    const err = await failure(feed(sink, pattern(8 * KiB), 1024))
    expect(err.code).toBe('quota')
    expect(opfs.log.live.size).toBe(0)
    expect(await failure(sink.append(pattern(1)))).toBe(err)
  })

  it('any other backend failure is a typed storage-error, also through the real OPFS adapter', async () => {
    const dir = fakeOpfsDirectory({ failOnWrite: { at: 2, error: new Error('disk gone') } })
    const opfs = createOpfsBackend(async () => dir as unknown as FileSystemDirectoryHandle)
    const env: StorageEnv = { memoryLimitBytes: KiB, opfs }
    const sink = await openSpool(env, { maxBytes: MiB, advertisedBytes: MiB }, signal())
    const err = await failure(feed(sink, pattern(4 * KiB), 1024))
    expect(err.code).toBe('storage-error')
    expect(dir.files.size).toBe(0)
  })
})

function GiB(n: number): number {
  return n * 1024 * MiB
}
