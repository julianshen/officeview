/**
 * S13.4 gap B (spec §12.4): ingestion and index construction memory is charged
 * to the document MemoryPool. PageBuilder buffers, pending pages and encode
 * scratch are transient shares released on completion/cancel; page summaries
 * are persistent. A pool that cannot hold the construction peak refuses the
 * import with a typed memory diagnostic instead of over-allocating.
 */
import { describe, expect, it } from 'vitest'
import { PAGE_SUMMARY_BYTES } from '../src/xlsx/storage/constructionMemory'
import { SourcePageStore, memoryPageBackend, payloadBytes, type PageBackend } from '../src/xlsx/storage/sourcePages'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { MemoryPool } from '../src/xlsx/view/tileCache'
import { colName, repeat, xlsxBytes } from './helpers/xlsx-package'

const MiB = 1024 * 1024
const rowsXml = (n: number, width = 3) => repeat(n, (r) => `<row r="${r + 1}">${repeat(width, (c) => `<c r="${colName(c)}${r + 1}"><v>${r * width + c}</v></c>`)}</row>`)
const blobOf = (bytes: Uint8Array) => ({ kind: 'blob' as const, blob: new Blob([bytes as BlobPart]) })
function gatedBackend(openAfter: number) {
  const inner = memoryPageBackend()
  let puts = 0
  let release!: () => void
  const gate = new Promise<void>((resolve) => { release = resolve })
  let blocked!: () => void
  const reached = new Promise<void>((resolve) => { blocked = resolve })
  const backend: PageBackend = {
    kind: 'memory',
    async put(key, bytes) { if (++puts > openAfter) { blocked(); await gate } await inner.put(key, bytes) },
    get: (key) => inner.get(key),
    removeAll: () => inner.removeAll(),
  }
  return { backend, release: () => release(), reached, inner }
}

describe('construction memory is charged to the document pool (spec §12.4)', () => {
  it('while a page write is in flight the pool holds its pending payload, its encode buffer and the page summaries', async () => {
    const pool = new MemoryPool({ limitBytes: 256 * MiB })
    const gated = gatedBackend(1)
    const store = new SourcePageStore({ docId: 'd', backend: gated.backend })
    store.bindMemory(pool)
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rowsXml(3000) })), { docId: 'd', profile: 'D', store, pageRows: 500 })
    await gated.reached
    const during = pool.stats()
    gated.release()
    expect(await session.done).toMatchObject({ kind: 'complete' })
    const [, p1] = store.pages('1')
    const pending = payloadBytes(await store.page('1', 'p1'))
    // p1 is being written: its built payload, its encoded copy and the summaries of p0 and p1.
    expect(during.transientBytes).toBeGreaterThanOrEqual(pending + p1.bytes)
    expect(during.retainedBytes).toBe(2 * PAGE_SUMMARY_BYTES)
    expect(during.peakBytes).toBeGreaterThanOrEqual(pending + p1.bytes + 2 * PAGE_SUMMARY_BYTES)
    await session.release()
  })

  it('returns to steady state after the import: no transient share, only the persistent summaries; release frees them', async () => {
    const pool = new MemoryPool({ limitBytes: 256 * MiB })
    const store = new SourcePageStore({ docId: 'd', backend: memoryPageBackend() })
    store.bindMemory(pool)
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rowsXml(3000), sharedStrings: Array.from({ length: 5000 }, (_, i) => `string ${i}`) })), { docId: 'd', profile: 'D', store, pageRows: 500 })
    expect(await session.done).toMatchObject({ kind: 'complete' })
    const pages = store.pages('1')
    const maxPage = Math.max(...pages.map((p) => p.bytes))
    const after = pool.stats()
    expect(after.transientBytes).toBe(0)
    expect(after.retainedBytes).toBe(pages.length * PAGE_SUMMARY_BYTES)
    // The peak saw at least one pending page plus its encoded copy.
    expect(after.peakBytes).toBeGreaterThanOrEqual(2 * maxPage)
    await session.release()
    expect(pool.stats().totalBytes).toBe(0)
  })

  it('a constrained pool refuses the import with a typed memory diagnostic and never over-allocates', async () => {
    const limitBytes = 32 * 1024
    const pool = new MemoryPool({ limitBytes })
    const store = new SourcePageStore({ docId: 'd', backend: memoryPageBackend() })
    store.bindMemory(pool)
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rowsXml(3000) })), { docId: 'd', profile: 'D', store, pageRows: 500 })
    const outcome = await session.done
    expect(outcome).toMatchObject({ kind: 'incomplete', failure: { code: 'resource-limit' } })
    expect(session.diagnostics.map((d) => d.code)).toContain('import-memory-budget')
    expect(pool.stats().peakBytes).toBeLessThanOrEqual(limitBytes)
    expect(pool.stats().transientBytes).toBe(0)
    await session.release()
    expect(pool.stats().totalBytes).toBe(0)
  })

  it('pending shared strings are charged too: a pool too small for one chunk refuses typed', async () => {
    const pool = new MemoryPool({ limitBytes: 64 * 1024 })
    const store = new SourcePageStore({ docId: 'd', backend: memoryPageBackend() })
    store.bindMemory(pool)
    const strings = Array.from({ length: 4096 }, (_, i) => `shared string number ${i} with some padding`)
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rowsXml(3), sharedStrings: strings })), { docId: 'd', profile: 'D', store })
    expect(await session.done).toMatchObject({ kind: 'incomplete', failure: { code: 'resource-limit' } })
    expect(session.diagnostics.map((d) => d.code)).toContain('import-memory-budget')
    expect(pool.stats().peakBytes).toBeLessThanOrEqual(64 * 1024)
    expect(pool.stats().transientBytes).toBe(0)
    await session.release()
  })

  it('cancellation releases every transient share and the summaries', async () => {
    const pool = new MemoryPool({ limitBytes: 256 * MiB })
    const gated = gatedBackend(2)
    const store = new SourcePageStore({ docId: 'd', backend: gated.backend })
    store.bindMemory(pool)
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rowsXml(3000) })), { docId: 'd', profile: 'D', store, pageRows: 500 })
    await gated.reached
    expect(pool.stats().totalBytes).toBeGreaterThan(0)
    const cancelled = session.cancel()
    gated.release()
    await cancelled
    expect(await session.done).toEqual({ kind: 'cancelled' })
    expect(pool.stats().transientBytes).toBe(0)
    expect(pool.stats().totalBytes).toBe(0)
  })

  it('charges made before the document pool is bound move into it on bind; a pool too small for them ends the import typed', async () => {
    const gated = gatedBackend(1)
    const store = new SourcePageStore({ docId: 'd', backend: gated.backend })
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rowsXml(3000) })), { docId: 'd', profile: 'D', store, pageRows: 500 })
    await gated.reached
    const pool = new MemoryPool({ limitBytes: 256 * MiB })
    store.bindMemory(pool)
    expect(pool.stats().transientBytes).toBeGreaterThan(0)
    expect(pool.stats().retainedBytes).toBe(2 * PAGE_SUMMARY_BYTES)
    gated.release()
    expect(await session.done).toMatchObject({ kind: 'complete' })
    expect(pool.stats().transientBytes).toBe(0)
    await session.release()

    const gated2 = gatedBackend(1)
    const store2 = new SourcePageStore({ docId: 'd', backend: gated2.backend })
    const session2 = await openImportSession(blobOf(xlsxBytes({ sheetData: rowsXml(3000) })), { docId: 'd', profile: 'D', store: store2, pageRows: 500 })
    await gated2.reached
    const tiny = new MemoryPool({ limitBytes: 1024 })
    expect(() => store2.bindMemory(tiny)).not.toThrow()
    gated2.release()
    expect(await session2.done).toMatchObject({ kind: 'incomplete', failure: { code: 'resource-limit' } })
    expect(tiny.stats().peakBytes).toBeLessThanOrEqual(1024)
    await session2.release()
  })
})
