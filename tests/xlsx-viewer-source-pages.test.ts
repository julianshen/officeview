/**
 * S1.2 sparse source pages: payload codec, OPFS/IndexedDB/memory backends with
 * typed storage/quota refusal, the PageSink store (stale owners, abandon and
 * close cleanup, ledger charges) and the decoded-page cache under the S4.3
 * MemoryPool (coalescing, eviction, reload under the same identity).
 */
import { describe, expect, it } from 'vitest'
import {
  K_BOOL, K_EMPTY, K_ERROR, K_NUMBER, K_SST, K_TEXT, PageBuilder, SourcePageStore, choosePageBackend, decodePage, encodePage,
  indexedDbPageBackend, memoryPageBackend, opfsPageBackend, textAt, type PageBackend,
} from '../src/xlsx/storage/sourcePages'
import { createQuotaLedger, quotaLimitsFor } from '../src/xlsx/storage/quota'
import { MemoryPool } from '../src/xlsx/view/tileCache'
import { PackageError } from '../src/core/package-error'
import { fakeIndexedDb, fakeOpfsDirectory } from './helpers/fake-storage'

function samplePage() {
  const b = new PageBuilder()
  b.add(0, 0, K_NUMBER, 1, 1.5, '', true, undefined)
  b.add(0, 16383, K_SST, 0, 7, '', true, undefined)
  b.add(262_144, 3, K_TEXT, 2, 0, 'a\ud800b', true, undefined)
  b.add(1_048_575, 16383, K_ERROR, 0, 0, '#N/A', true, { formula: 'NA()', ca: true })
  b.add(1_048_575, 16383 - 1, K_EMPTY, 3, 0, '', false, { formula: 'X\udc00', sharedFormula: { si: 2, ref: 'A1:A9' } })
  b.add(1_048_575, 16383 - 2, K_BOOL, 0, 1, '', true, undefined)
  return b.build()
}

describe('source page payload codec', () => {
  it('round-trips rows, columns, kinds, raw UTF-16 text and formula metadata', () => {
    const page = samplePage()
    const back = decodePage(encodePage(page))
    expect([...back.rows]).toEqual([0, 0, 262_144, 1_048_575, 1_048_575, 1_048_575])
    expect([...back.cols]).toEqual([0, 16383, 3, 16383, 16382, 16381])
    expect([...back.kinds]).toEqual([K_NUMBER, K_SST, K_TEXT, K_ERROR, K_EMPTY, K_BOOL])
    expect(back.nums[0]).toBe(1.5)
    expect(back.nums[1]).toBe(7)
    expect(textAt(back, 2)).toBe('a\ud800b')
    expect(textAt(back, 3)).toBe('#N/A')
    expect(back.meta.get(4)).toEqual({ formula: 'X\udc00', sharedFormula: { si: 2, ref: 'A1:A9' } })
    expect(back.meta.get(3)).toEqual({ formula: 'NA()', ca: true })
    // Row and column lanes are separate: row 262,144 never aliases row 0 (rule 3).
    expect(back.rows.BYTES_PER_ELEMENT).toBe(4)
    expect(back.cols.BYTES_PER_ELEMENT).toBe(2)
  })

  it('rejects a truncated or foreign page with a typed parse-error', () => {
    const bytes = encodePage(samplePage())
    expect(() => decodePage(bytes.subarray(0, bytes.length - 8))).toThrow(PackageError)
    const foreign = bytes.slice()
    foreign[0] ^= 0xff
    expect(() => decodePage(foreign)).toThrow(/bad header/)
  })
})

describe('page backends', () => {
  const roundTrip = async (backend: PageBackend) => {
    await backend.put('1/p0', new Uint8Array([1, 2, 3]))
    await backend.put('sst/0', new Uint8Array([9]))
    expect([...(await backend.get('1/p0'))!]).toEqual([1, 2, 3])
    expect(await backend.get('1/missing')).toBeUndefined()
    await backend.removeAll()
    expect(await backend.get('1/p0')).toBeUndefined()
    await backend.removeAll()
  }
  it('memory backend stores, reads and removes', async () => { await roundTrip(memoryPageBackend()) })
  it('OPFS backend writes one file per page and removes every file on cleanup', async () => {
    const dir = fakeOpfsDirectory()
    const backend = opfsPageBackend(dir, 'doc-1')
    await backend.put('1/p0', new Uint8Array([1]))
    expect([...dir.files.keys()].every((name) => name.startsWith('officeview-pages-doc-1-'))).toBe(true)
    await roundTrip(backend)
    expect(dir.files.size).toBe(0)
  })
  it('IndexedDB backend awaits each transaction and deletes its database on cleanup', async () => {
    const idb = fakeIndexedDb()
    const backend = indexedDbPageBackend(idb, 'doc-1')
    await roundTrip(backend)
    expect(idb.deleted).toContain('officeview-pages-doc-1')
    expect(idb.transactions.filter((t) => t.mode === 'readwrite').every((t) => t.puts === 1)).toBe(true)
  })
  it('a storage write failure is a typed storage-error and quota exhaustion a typed quota', async () => {
    const failing = opfsPageBackend(fakeOpfsDirectory({ failOnWrite: { at: 1, error: new DOMException('disk', 'UnknownError') } }), 'd')
    await expect(failing.put('1/p0', new Uint8Array([1]))).rejects.toMatchObject({ code: 'storage-error' })
    const full = opfsPageBackend(fakeOpfsDirectory({ failOnWrite: { at: 1, error: new DOMException('full', 'QuotaExceededError') } }), 'd')
    await expect(full.put('1/p0', new Uint8Array([1]))).rejects.toMatchObject({ code: 'quota' })
  })
})

describe('choosePageBackend (quota admission, no silent RAM fallback)', () => {
  it('keeps a small admitted document in memory', async () => {
    expect((await choosePageBackend({}, 'd', 1024, 4096)).kind).toBe('memory')
  })
  it('prefers OPFS, then IndexedDB, for a large document', async () => {
    expect((await choosePageBackend({ opfs: async () => fakeOpfsDirectory(), indexedDb: fakeIndexedDb() }, 'd', 1 << 20, 4096)).kind).toBe('opfs')
    expect((await choosePageBackend({ indexedDb: fakeIndexedDb() }, 'd', 1 << 20, 4096)).kind).toBe('indexeddb')
  })
  it('refuses a large document with no storage (storage-error) or too little quota (quota)', async () => {
    await expect(choosePageBackend({}, 'd', 1 << 20, 4096)).rejects.toMatchObject({ code: 'storage-error' })
    await expect(choosePageBackend({ indexedDb: fakeIndexedDb(), estimate: async () => ({ quota: 1000, usage: 900 }) }, 'd', 1 << 20, 4096))
      .rejects.toMatchObject({ code: 'quota' })
  })
})

describe('SourcePageStore as the import PageSink', () => {
  const imported = (pageId: string, r0: number, r1: number) => {
    const b = new PageBuilder()
    b.add(r0, 0, K_NUMBER, 1, r0, '', true, undefined)
    b.add(r1, 2, K_NUMBER, 1, r1, '', true, undefined)
    return { sheetId: '1', pageId, r0, r1, cellCount: 2, bytes: b.bytes, payload: b.build() }
  }

  it('indexes committed pages with occupancy summaries and charges the ledger pages pool', async () => {
    const ledger = createQuotaLedger(quotaLimitsFor('D'), { maxTotalExpandedBytes: Number.MAX_SAFE_INTEGER })
    const store = new SourcePageStore({ docId: 'd', backend: memoryPageBackend(), ledger })
    await store.commit(imported('p0', 0, 9))
    await store.commit(imported('p1', 10, 19))
    expect(store.pages('1').map((p) => [p.pageId, p.r0, p.r1, p.c0, p.c1, p.aggregate.records])).toEqual([['p0', 0, 9, 0, 2, 2], ['p1', 10, 19, 0, 2, 2]])
    expect(ledger.snapshot().current.pages).toBe(store.stats.storedBytes)
    await store.close()
    expect(ledger.snapshot().current.pages).toBe(0)
  })

  it('rejects a commit for another document and any commit after abandon; abandon removes stored pages', async () => {
    const backend = memoryPageBackend()
    const store = new SourcePageStore({ docId: 'd', backend })
    await store.commit(imported('p0', 0, 9))
    await expect(store.commit(imported('p1', 10, 19), 'stale-doc')).rejects.toMatchObject({ code: 'cancelled' })
    await store.abandon('d')
    expect(backend.size).toBe(0)
    await expect(store.commit(imported('p1', 10, 19))).rejects.toMatchObject({ code: 'cancelled' })
    expect(store.pages('1')).toEqual([])
  })

  it('a page written while the store is abandoned is never indexed', async () => {
    let release!: () => void
    const gate = new Promise<void>((resolve) => { release = resolve })
    const inner = memoryPageBackend()
    const slow: PageBackend = { kind: 'memory', put: async (k, b) => { await gate; await inner.put(k, b) }, get: (k) => inner.get(k), removeAll: () => inner.removeAll() }
    const store = new SourcePageStore({ docId: 'd', backend: slow })
    const pending = store.commit(imported('p0', 0, 9))
    await store.abandon('d')
    release()
    await expect(pending).rejects.toMatchObject({ code: 'cancelled' })
    expect(store.pages('1')).toEqual([])
  })

  it('coalesces concurrent loads, evicts under the MemoryPool and reloads the same bytes', async () => {
    const store = new SourcePageStore({ docId: 'd', backend: memoryPageBackend() })
    for (let i = 0; i < 4; i++) await store.commit(imported(`p${i}`, i * 10, i * 10 + 9))
    const pool = new MemoryPool({ limitBytes: 120, highWatermark: 1, lowWatermark: 1 })
    store.bindMemory(pool)
    const [a, b] = await Promise.all([store.page('1', 'p0'), store.page('1', 'p0')])
    expect(a).toBe(b)
    expect(store.stats).toMatchObject({ loads: 1, coalesced: 1 })
    for (let i = 1; i < 4; i++) await store.page('1', `p${i}`)
    expect(pool.stats().byTier.source).toBeLessThanOrEqual(120)
    expect(store.stats.evictions).toBeGreaterThan(0)
    const again = await store.page('1', 'p0')
    expect([...again.rows]).toEqual([...a.rows])
    expect(store.stats.loads).toBeGreaterThan(4)
  })

  it('serves shared strings by id from stored chunks', async () => {
    const store = new SourcePageStore({ docId: 'd', backend: memoryPageBackend() })
    await store.commitStrings(0, ['a', 'b\ud800'], 2)
    await store.commitStrings(1, ['c'], 2)
    expect(await store.sharedString(1)).toBe('b\ud800')
    expect(await store.sharedString(2)).toBe('c')
    expect(await store.sharedString(3)).toBeUndefined()
  })
})
