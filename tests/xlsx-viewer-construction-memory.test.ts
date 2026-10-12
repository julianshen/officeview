/**
 * S13.4 gap B (spec §12.4): ingestion and index construction memory is charged
 * to the document MemoryPool. PageBuilder buffers, pending pages and encode
 * scratch are transient shares released on completion/cancel; page summaries
 * are persistent. A pool that cannot hold the construction peak refuses the
 * import with a typed memory diagnostic instead of over-allocating.
 */
import { describe, expect, it } from 'vitest'
import { spawnSync } from 'node:child_process'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { PAGE_SUMMARY_BYTES } from '../src/xlsx/storage/constructionMemory'
import { PageBuilder, K_EMPTY, SourcePageStore, encodePage, jsonLength, memoryPageBackend, opfsPageBackend, payloadBytes, type PageBackend } from '../src/xlsx/storage/sourcePages'
import { fakeOpfsDirectory } from './helpers/fake-storage'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { DocumentSession, MIN_WORKER_BUDGET_BYTES } from '../src/worker/viewerSession'
import { ViewerKernel } from '../src/worker/viewerKernel'
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
    delete: (k) => inner.delete(k), removeAll: () => inner.removeAll(),
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
    const [p0] = store.pages('1')
    // p1 is being written: its built payload, its encoded copy and the summaries of p0 and p1;
    // the memory backend retains p0 and (from the moment it is handed over) p1.
    expect(during.transientBytes).toBeGreaterThanOrEqual(pending + p1.bytes)
    expect(during.retainedBytes).toBe(2 * PAGE_SUMMARY_BYTES + p0.bytes + p1.bytes)
    expect(during.peakBytes).toBeGreaterThanOrEqual(pending + p1.bytes + 2 * PAGE_SUMMARY_BYTES)
    await session.release()
  })

  it('returns to steady state after the import: no transient share, only the summaries and the memory-held records; release frees them', async () => {
    const pool = new MemoryPool({ limitBytes: 256 * MiB })
    const backend = memoryPageBackend()
    const store = new SourcePageStore({ docId: 'd', backend })
    store.bindMemory(pool)
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rowsXml(3000), sharedStrings: Array.from({ length: 5000 }, (_, i) => `string ${i}`) })), { docId: 'd', profile: 'D', store, pageRows: 500 })
    expect(await session.done).toMatchObject({ kind: 'complete' })
    const pages = store.pages('1')
    const maxPage = Math.max(...pages.map((p) => p.bytes))
    const after = pool.stats()
    expect(after.transientBytes).toBe(0)
    // The memory backend's retained buffers (pages and shared-string chunks) are worker RAM.
    expect(after.retainedBytes).toBe(pages.length * PAGE_SUMMARY_BYTES + backend.heldBytes)
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

  it('charges made before the document pool is bound move into it on bind; a pool too small for them refuses the bind typed', async () => {
    const gated = gatedBackend(1)
    const store = new SourcePageStore({ docId: 'd', backend: gated.backend })
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rowsXml(3000) })), { docId: 'd', profile: 'D', store, pageRows: 500 })
    await gated.reached
    const pool = new MemoryPool({ limitBytes: 256 * MiB })
    store.bindMemory(pool)
    expect(pool.stats().transientBytes).toBeGreaterThan(0)
    // Summaries of p0 and p1 plus the buffers the memory backend holds (p0 stored, p1 handed over).
    expect(pool.stats().retainedBytes).toBe(store.construction.persistentBytes)
    expect(store.construction.persistentBytes).toBeGreaterThan(2 * PAGE_SUMMARY_BYTES)
    gated.release()
    expect(await session.done).toMatchObject({ kind: 'complete' })
    expect(pool.stats().transientBytes).toBe(0)
    await session.release()

    // A pool too small for the live charges: bind refuses typed and moves nothing.
    const gated2 = gatedBackend(1)
    const store2 = new SourcePageStore({ docId: 'd', backend: gated2.backend })
    const session2 = await openImportSession(blobOf(xlsxBytes({ sheetData: rowsXml(3000) })), { docId: 'd', profile: 'D', store: store2, pageRows: 500 })
    await gated2.reached
    const held = store2.construction.persistentBytes
    const tiny = new MemoryPool({ limitBytes: 1024 })
    expect(() => store2.bindMemory(tiny)).toThrow(expect.objectContaining({ code: 'resource-limit', reason: 'memory-budget' }))
    expect(tiny.stats().totalBytes).toBe(0)
    expect(store2.construction.persistentBytes).toBe(held)
    expect(store2.construction.transientBytes).toBeGreaterThan(0)
    gated2.release()
    expect(await session2.done).toMatchObject({ kind: 'complete' })
    await session2.release()
  })
})

describe('review round 1 (codex) construction-memory findings', () => {
  const big = (i: number) => `${String.fromCharCode(65 + i)}`.repeat(50_000)
  const bigRows = (n: number) => repeat(n, (r) => `<row r="${r + 1}"><c r="A${r + 1}" t="inlineStr"><is><t>${big(r)}</t></is></c></row>`)

  it('MAJOR 1: records a memory backend retains are charged to the pool for their lifetime (blobs included)', async () => {
    // Twelve 50,000-unit values: 1.2 MB of blob records retained in RAM.
    const refused = new SourcePageStore({ docId: 'd', backend: memoryPageBackend() })
    refused.bindMemory(new MemoryPool({ limitBytes: 1024 * 1024 }))
    const a = await openImportSession(blobOf(xlsxBytes({ sheetData: bigRows(12) })), { docId: 'd', profile: 'D', store: refused, pageRows: 1 })
    expect(await a.done).toMatchObject({ kind: 'incomplete', failure: { code: 'resource-limit' } })
    expect(a.diagnostics.map((d) => d.code)).toContain('import-memory-budget')
    await a.release()

    const pool = new MemoryPool({ limitBytes: 8 * MiB })
    const backend = memoryPageBackend()
    const store = new SourcePageStore({ docId: 'd', backend })
    store.bindMemory(pool)
    const b = await openImportSession(blobOf(xlsxBytes({ sheetData: bigRows(12) })), { docId: 'd', profile: 'D', store, pageRows: 1 })
    expect(await b.done).toMatchObject({ kind: 'complete' })
    expect(backend.heldBytes).toBeGreaterThan(1_200_000)
    expect(pool.stats().retainedBytes).toBe(store.pages('1').length * PAGE_SUMMARY_BYTES + backend.heldBytes)
    // Decoded views share the held buffers: reading them charges no second copy.
    const sheet = createPagedWorkbookSource(b).store('1')
    expect(await sheet.readCell({ row: 3, col: 0 }, { signal: new AbortController().signal, stamp: { ...sheet.stamp } })).toMatchObject({ cell: { value: big(3) } })
    expect(pool.stats().byTier.source - (store.pages('1').length * PAGE_SUMMARY_BYTES + backend.heldBytes)).toBeLessThan(1024)
    await b.release()
    expect(pool.stats().totalBytes).toBe(0)
  })

  it('MAJOR 2: binding a completed store to a pool too small for its live charges refuses typed and drops nothing', async () => {
    const backend = memoryPageBackend()
    const store = new SourcePageStore({ docId: 'd', backend })
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rowsXml(10) })), { docId: 'd', profile: 'D', store })
    expect(await session.done).toMatchObject({ kind: 'complete' })
    const held = store.construction.persistentBytes
    expect(held).toBe(store.pages('1').length * PAGE_SUMMARY_BYTES + backend.heldBytes)
    const small = new MemoryPool({ limitBytes: 100 })
    expect(() => store.bindMemory(small)).toThrow(expect.objectContaining({ code: 'resource-limit', reason: 'memory-budget' }))
    expect(small.stats().totalBytes).toBe(0)
    expect(store.construction.persistentBytes).toBe(held)
    const pool = new MemoryPool({ limitBytes: 8 * MiB })
    store.bindMemory(pool)
    expect(pool.stats().retainedBytes).toBe(held)
    await session.release()
  })

  it('MAJOR 2: the viewer session and the worker kernel surface the bind refusal typed', async () => {
    const store = new SourcePageStore({ docId: 'd', backend: memoryPageBackend() })
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rowsXml(20_000) })), { docId: 'd', profile: 'D', store })
    await session.done
    const source = createPagedWorkbookSource(session)
    expect(() => new DocumentSession('d', source, { workerBudgetBytes: MIN_WORKER_BUDGET_BYTES })).toThrow(expect.objectContaining({ reason: 'memory-budget' }))
    await source.release()

    const kernel = new ViewerKernel({ parse: async () => { throw new Error('no parse expected') }, importLimits: { object: { rows: 1 } } }, { yield: async () => {}, workerBudgetBytes: MIN_WORKER_BUDGET_BYTES })
    const replies: unknown[] = []
    await kernel.handle({ kind: 'open', requestId: 'r1', docId: 'k', source: { kind: 'blob', blob: new Blob([xlsxBytes({ sheetData: rowsXml(20_000) }) as BlobPart]) }, profile: 'D' }, (out) => { replies.push(out) })
    const last = replies[replies.length - 1] as { kind: string; code?: string; diagnostics?: { code: string }[] }
    expect(last.kind).toBe('error')
    expect(JSON.stringify(last)).toContain('resource-limit')
    expect(JSON.stringify(last)).toContain('import-memory-budget')
  })

  it('MAJOR 3: a pending page keeps its formula-metadata charge; JSON and buffer are charged before allocation', async () => {
    // 80 long formulas with CJK literals: ~1.28 MB of formula text in one pending page.
    const formula = (i: number) => `CONCAT("${'漢'.repeat(8000)}",${i})`
    const rows = repeat(80, (r) => `<row r="${r + 1}"><c r="A${r + 1}"><f>${formula(r)}</f><v>${r}</v></c></row>`)
    const pool = new MemoryPool({ limitBytes: 256 * MiB })
    const store = new SourcePageStore({ docId: 'd', backend: opfsPageBackend(fakeOpfsDirectory(), 'd') })
    store.bindMemory(pool)
    let atCommit = 0
    const commit = store.commit.bind(store)
    store.commit = async (page, docId) => { atCommit = Math.max(atCommit, pool.stats().transientBytes); return commit(page, docId) }
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rows })), { docId: 'd', profile: 'D', store })
    expect(await session.done).toMatchObject({ kind: 'complete' })
    expect(atCommit).toBeGreaterThanOrEqual(2 * 80 * 8000)
    // A pool that cannot hold the pending metadata refuses typed instead.
    const tight = new SourcePageStore({ docId: 't', backend: opfsPageBackend(fakeOpfsDirectory(), 't') })
    tight.bindMemory(new MemoryPool({ limitBytes: 1024 * 1024 }))
    const refused = await openImportSession(blobOf(xlsxBytes({ sheetData: rows })), { docId: 't', profile: 'D', store: tight })
    expect(await refused.done).toMatchObject({ kind: 'incomplete', failure: { code: 'resource-limit' } })
    await refused.release()
    await session.release()
  })

  it('MAJOR 3: encodePage asks for the exact scratch (buffer + meta JSON) before allocating either', () => {
    const b = new PageBuilder()
    b.add(0, 0, K_EMPTY, 0, 0, '', false, { formula: `"\ud800"&CHAR(10)&"${'漢'.repeat(50)}"` })
    b.add(1, 0, K_EMPTY, 0, 0, '', false, { formula: 'A1', sharedFormula: { si: 3, ref: 'A1:A9' }, ca: true })
    const page = b.build()
    let asked = -1
    const bytes = encodePage(page, (n) => { if (asked < 0) asked = n })
    const json = JSON.stringify([...page.meta])
    expect(asked).toBe(bytes.length + 2 * json.length)
    expect(page.metaJsonLength).toBe(json.length)
  })

  it('jsonLength equals JSON.stringify for the formula-metadata subset (escapes, lone surrogates, pairs, numbers)', () => {
    const samples: unknown[] = [
      'plain', 'q"b\\s', 'ctl\u0001\u001f\b\f\n\r\t', '\ud800', '\udc00x', '😀', '\ud83d\ude00\ud800', '\u2028\u2029',
      0, -0, 1.5, 1e21, 123456789, true, false, null, [1, 'a', [2]], { a: 1, b: undefined, c: { si: 2, ref: 'A1' } },
      [7, { formula: 'X\udc00', sharedFormula: { si: 2, ref: 'A1:A9' }, dynamicArray: true }],
    ]
    for (const value of samples) expect(jsonLength(value), JSON.stringify(value)).toBe(JSON.stringify(value).length)
  })
})

describe('review round 1 (codex) MAJOR 4: the estimate model is conservative against a Bun heap probe', () => {
  it('each modelled structure charges at least its measured live heap (and at most twice it), across capacity transitions', () => {
    const probe = spawnSync('bun', [join(dirname(fileURLToPath(import.meta.url)), '..', 'scripts', 'construction-memory-probe.ts')], { encoding: 'utf8', timeout: 120_000 })
    expect(probe.status, probe.stderr).toBe(0)
    const results = JSON.parse(probe.stdout.trim().split('\n').at(-1)!) as { name: string; count: number; observed: number; model: number }[]
    expect(new Set(results.map((r) => r.name))).toEqual(new Set(['builder-numeric', 'builder-text', 'builder-formula', 'summaries', 'pending-strings']))
    for (const r of results) {
      expect(r.model, `${r.name} x${r.count}`).toBeGreaterThanOrEqual(r.observed)
      expect(r.model, `${r.name} x${r.count}`).toBeLessThanOrEqual(2 * r.observed)
    }
  }, 150_000)
})
