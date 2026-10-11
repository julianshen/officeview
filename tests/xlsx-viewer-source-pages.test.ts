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

// ---------------------------------------------------------------------------
// Paged import session: RED 21 (progressive commits through the sink), 22
// (expanded-byte accounting and progress from completed batches), 23
// (incomplete/cancelled/stale never publish a complete blank source) and 14a
// (paged queryStructural incompleteness).

import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { readZipInventory, bytesRangeSource } from '../src/core/zip-stream'
import type { ImportProgress } from '../src/xlsx/storage/importSession'
import type { SourceStore } from '../src/xlsx/view/contracts'
import { colName, repeat, xlsxBytes } from './helpers/xlsx-package'

const rowsXml = (n: number, width = 3) => repeat(n, (r) => `<row r="${r + 1}">${repeat(width, (c) => `<c r="${colName(c)}${r + 1}"><v>${r * width + c}</v></c>`)}</row>`)
const blobOf = (bytes: Uint8Array) => ({ kind: 'blob' as const, blob: new Blob([bytes as BlobPart]) })
const ro = (store: SourceStore) => ({ signal: new AbortController().signal, stamp: { ...store.stamp } })
function gatedBackend(openAfter: number) {
  const inner = memoryPageBackend()
  let puts = 0
  let release!: () => void
  const gate = new Promise<void>((resolve) => { release = resolve })
  let removed = 0
  const backend: PageBackend = {
    kind: 'memory',
    async put(key, bytes) { if (++puts > openAfter) await gate; await inner.put(key, bytes) },
    get: (key) => inner.get(key),
    async removeAll() { removed++; await inner.removeAll() },
  }
  return { backend, release: () => release(), get removed() { return removed }, get size() { return inner.size } }
}
const settle = () => new Promise((resolve) => setTimeout(resolve, 20))

describe('paged import session', () => {
  it('RED 21: the head page is committed and readable while the import is still loading; complete only after the last commit', async () => {
    const bytes = xlsxBytes({ sheetData: rowsXml(3000) })
    const sheetSize = (await readZipInventory(bytesRangeSource(bytes), { maxEntries: 100, maxCentralDirectoryBytes: 1 << 20, maxNameBytes: 1024, maxPartExpandedBytes: 1 << 30, maxTotalExpandedBytes: 1 << 30, maxRatio: 1e9 })).find('xl/worksheets/sheet1.xml')!.uncompressedSize
    const gated = gatedBackend(1)
    const store = new SourcePageStore({ docId: 'd', backend: gated.backend })
    const commits: number[] = []
    const session = await openImportSession(blobOf(bytes), { docId: 'd', profile: 'D', store, pageRows: 100, onCommit: (p) => commits.push(p.r1) })
    await settle()
    expect(commits).toEqual([99])
    expect(session.completeness).toBe('loading')
    expect(session.ledger.snapshot().totalExpanded).toBeLessThan(sheetSize)
    const source = createPagedWorkbookSource(session)
    const sheet = source.store('1')
    expect(await sheet.readCell({ row: 5, col: 1 }, ro(sheet))).toMatchObject({ kind: 'cell', cell: { value: 16 } })
    // Uncommitted coordinates are unknown, never blank (incl. rows that will turn out empty).
    expect(await sheet.readCell({ row: 2500, col: 1 }, ro(sheet))).toMatchObject({ kind: 'unknown' })
    expect(await sheet.readCell({ row: 900_000, col: 0 }, ro(sheet))).toMatchObject({ kind: 'unknown' })
    const head = await sheet.readRows({ r0: 0, r1: 200, c0: 0, c1: 2 }, undefined, ro(sheet))
    expect(head.completeness).toBe('loading')
    expect(head.items.map((r) => r.row)).toEqual(Array.from({ length: 100 }, (_, i) => i))
    const overview = await sheet.queryOverview({ maxGroups: 4, groups: [{ groupId: 'tail', rows: { kind: 'intervals', intervals: [{ r0: 2000, r1: 2100 }] }, cols: [{ c0: 0, c1: 2 }] }] }, ro(sheet))
    expect(overview.items[0]).toMatchObject({ content: 'unknown' })
    expect(overview.completeness).toBe('loading')
    gated.release()
    expect(await session.done).toMatchObject({ kind: 'complete' })
    expect(session.completeness).toBe('complete')
    expect(commits.at(-1)).toBe(1_048_575)
    expect(commits.length).toBe(31)
    expect(await sheet.readCell({ row: 900_000, col: 0 }, ro(sheet))).toEqual({ kind: 'blank' })
    await source.release()
    expect(gated.removed).toBeGreaterThan(0)
  })

  it('RED 21: a sink that rejects a commit ends the session incomplete with a typed failure and abandons every page', async () => {
    const inner = memoryPageBackend()
    let puts = 0
    let removed = false
    const backend: PageBackend = {
      kind: 'memory',
      async put(key, bytes) { if (++puts === 3) throw new DOMException('full', 'QuotaExceededError'); await inner.put(key, bytes) },
      get: (key) => inner.get(key),
      async removeAll() { removed = true; await inner.removeAll() },
    }
    const store = new SourcePageStore({ docId: 'd', backend })
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rowsXml(1000) })), { docId: 'd', profile: 'D', store, pageRows: 100 })
    const outcome = await session.done
    expect(outcome).toMatchObject({ kind: 'incomplete', failure: { code: 'quota' } })
    expect(session.completeness).toBe('incomplete')
    expect(removed).toBe(true)
    expect(inner.size).toBe(0)
    const sheet = createPagedWorkbookSource(session).store('1')
    expect((await sheet.readCell({ row: 999, col: 0 }, ro(sheet))).kind).toBe('error')
  })

  it('RED 22: expanded bytes equal the inflated part sizes; progress comes from completed batches with known totals', async () => {
    const bytes = xlsxBytes({ sheetData: rowsXml(800), sharedStrings: ['a', 'b'] })
    const inventory = await readZipInventory(bytesRangeSource(bytes), { maxEntries: 100, maxCentralDirectoryBytes: 1 << 20, maxNameBytes: 1024, maxPartExpandedBytes: 1 << 30, maxTotalExpandedBytes: 1 << 30, maxRatio: 1e9 })
    const read = ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/sharedStrings.xml', 'xl/worksheets/sheet1.xml']
    const expected = read.reduce((n, name) => n + inventory.find(name)!.uncompressedSize, 0)
    const events: ImportProgress[] = []
    const store = new SourcePageStore({ docId: 'd', backend: memoryPageBackend() })
    const session = await openImportSession(blobOf(bytes), { docId: 'd', profile: 'D', store, pageRows: 100, onProgress: (p) => events.push({ ...p }) })
    await session.done
    expect(session.ledger.snapshot().totalExpanded).toBe(expected)
    const last = events.at(-1)!
    expect(last).toMatchObject({ phase: 'done', expandedBytes: expected, expandedTotal: expected, populatedCells: 2400, compressedTotal: bytes.length })
    const sheetEvents = events.filter((e) => e.phase === 'sheet')
    expect(sheetEvents.length).toBe(9)
    // populatedCells grows only by whole committed pages (300 cells per 100-row page).
    expect(sheetEvents.map((e) => e.populatedCells)).toEqual([300, 600, 900, 1200, 1500, 1800, 2100, 2400, 2400])
    for (let i = 1; i < events.length; i++) {
      expect(events[i].expandedBytes).toBeGreaterThanOrEqual(events[i - 1].expandedBytes)
      expect(events[i].compressedRead).toBeGreaterThanOrEqual(events[i - 1].compressedRead)
    }
    await session.release()
  })

  it('RED 22: worksheet rels, table and drawing parts join the expanded total before they are read', async () => {
    const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    const bytes = xlsxBytes({
      sheetData: rowsXml(30),
      afterSheetData: `<drawing r:id="rId2" ${R}/><tableParts count="2"><tablePart r:id="rId1" ${R}/><tablePart r:id="rId1" ${R}/></tableParts>`,
      extra: [
        { name: 'xl/worksheets/_rels/sheet1.xml.rels', data: '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>' },
        { name: 'xl/tables/table1.xml', data: '<?xml version="1.0"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:C30"><autoFilter ref="A1:C30"/><tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/><tableColumn id="3" name="c"/></tableColumns></table>' },
        { name: 'xl/drawings/drawing1.xml', data: '<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>' },
      ],
    })
    const inventory = await readZipInventory(bytesRangeSource(bytes), { maxEntries: 100, maxCentralDirectoryBytes: 1 << 20, maxNameBytes: 1024, maxPartExpandedBytes: 1 << 30, maxTotalExpandedBytes: 1 << 30, maxRatio: 1e9 })
    const read = ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/worksheets/sheet1.xml', 'xl/worksheets/_rels/sheet1.xml.rels', 'xl/tables/table1.xml', 'xl/drawings/drawing1.xml']
    const expected = read.reduce((n, name) => n + inventory.find(name)!.uncompressedSize, 0)
    const events: ImportProgress[] = []
    const store = new SourcePageStore({ docId: 'd', backend: memoryPageBackend() })
    const session = await openImportSession(blobOf(bytes), { docId: 'd', profile: 'D', store, onProgress: (p) => events.push({ ...p }) })
    expect(await session.done).toMatchObject({ kind: 'complete' })
    for (const e of events) expect(e.expandedBytes).toBeLessThanOrEqual(e.expandedTotal as number)
    expect(session.ledger.snapshot().totalExpanded).toBe(expected)
    expect(events.at(-1)).toMatchObject({ phase: 'done', expandedBytes: expected, expandedTotal: expected })
    // A repeated tablePart resolves twice, exactly as parse.ts does; its part is inflated once.
    expect(session.workbook.sheets[0].tables?.map((t) => t.id)).toEqual(['1', '1'])
    await session.release()
  })

  it('RED 23: AlternateContent-wrapped cell data, malformed XML and a cell limit end incomplete with a typed failure, never complete blank', async () => {
    const cases: Array<[string, Uint8Array, string, Parameters<typeof openImportSession>[1]['limits']?]> = [
      ['alternate content', xlsxBytes({ sheetData: '', afterSheetData: '', beforeSheetData: '', extra: [] }).length ? xlsxBytes({ sheetData: rowsXml(3).replace('<row r="2">', '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x"><row r="2">').replace('</row><row r="3">', '</row></mc:Choice></mc:AlternateContent><row r="3">') }) : new Uint8Array(), 'unsupported'],
      ['malformed', xlsxBytes({ sheetData: rowsXml(3) + '<row r="9"><c r="A9"><v>1</c></row>' }), 'parse-error'],
      ['cell limit', xlsxBytes({ sheetData: rowsXml(50) }), 'resource-limit', { paged: { cells: 100 } }],
    ]
    for (const [name, bytes, code, limits] of cases) {
      const store = new SourcePageStore({ docId: 'd', backend: memoryPageBackend() })
      const session = await openImportSession(blobOf(bytes), { docId: 'd', profile: 'D', store, pageRows: 1, ...(limits ? { limits } : {}) })
      const outcome = await session.done
      expect(outcome, name).toMatchObject({ kind: 'incomplete', failure: { code } })
      expect(session.completeness, name).toBe('incomplete')
      const sheet = createPagedWorkbookSource(session).store('1')
      const far = await sheet.readCell({ row: 1_048_575, col: 0 }, ro(sheet))
      expect(far.kind, name).toBe('error')
      const structure = await sheet.queryStructural({ r0: 0, r1: 1_048_575, c0: 0, c1: 16_383 }, ro(sheet))
      expect(structure.completeness, name).toBe('incomplete')
      expect(structure.error?.code, name).toBe(code)
      const done = await sheet.ensureComplete({ r0: 0, r1: 1_048_575, c0: 0, c1: 0 }, ro(sheet))
      expect(done.completeness, name).toBe('incomplete')
      await session.release()
    }
  })

  it('RED 23: cancellation releases every reservation and removes stored pages; a stale owner cannot publish', async () => {
    const gated = gatedBackend(2)
    const ledger = createQuotaLedger(quotaLimitsFor('D'), { maxTotalExpandedBytes: 1 << 30 })
    const store = new SourcePageStore({ docId: 'd', backend: gated.backend, ledger })
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rowsXml(2000) })), { docId: 'd', profile: 'D', store, ledger, pageRows: 100 })
    await settle()
    expect(session.completeness).toBe('loading')
    const cancelled = session.cancel()
    gated.release()
    await cancelled
    expect(await session.done).toEqual({ kind: 'cancelled' })
    expect(Object.values(ledger.snapshot().current).every((v) => v === 0)).toBe(true)
    expect(gated.size).toBe(0)
    expect(store.pages('1')).toEqual([])
    // The cancelled owner's late page is refused, never indexed.
    await expect(store.commit({ sheetId: '1', pageId: 'late', r0: 0, r1: 0, cellCount: 0, bytes: 0, payload: new PageBuilder().build() })).rejects.toMatchObject({ code: 'cancelled' })
    const sheet = createPagedWorkbookSource(session).store('1')
    expect((await sheet.readCell({ row: 0, col: 0 }, ro(sheet))).kind).not.toBe('blank')
  })

  it('RED 14a: AlternateContent-wrapped mergeCells, tablePart or drawing make queryStructural incomplete with a typed error', async () => {
    const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
    for (const [domain, xml] of [
      ['merges', `<mc:AlternateContent ${MC}><mc:Choice Requires="x"><mergeCells><mergeCell ref="A1:B2"/></mergeCells></mc:Choice></mc:AlternateContent>`],
      ['tables', `<mc:AlternateContent ${MC}><mc:Choice Requires="x"><tableParts count="1"><tablePart r:id="rId9" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></tableParts></mc:Choice></mc:AlternateContent>`],
      ['drawings', `<mc:AlternateContent ${MC}><mc:Choice Requires="x"><drawing r:id="rId9" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></mc:Choice></mc:AlternateContent>`],
    ] as const) {
      const store = new SourcePageStore({ docId: 'd', backend: memoryPageBackend() })
      const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rowsXml(3), afterSheetData: xml })), { docId: 'd', profile: 'D', store })
      expect(await session.done, domain).toMatchObject({ kind: 'complete' })
      const sheet = createPagedWorkbookSource(session).store('1')
      const rect = { r0: 0, r1: 9, c0: 0, c1: 5 }
      const result = await sheet.queryStructural(rect, ro(sheet))
      expect(result.completeness, domain).toBe('incomplete')
      expect(result.missing, domain).toEqual([rect])
      expect(result.error?.code, domain).toBe('unsupported')
      expect(result.error?.message, domain).toContain(domain)
      expect(session.diagnostics.some((d) => d.code === 'structure-incomplete'), domain).toBe(true)
      // Cells stay readable: selection/inspection is unaffected.
      expect(await sheet.readCell({ row: 0, col: 0 }, ro(sheet))).toMatchObject({ kind: 'cell' })
      await session.release()
    }
  })
})

import { importLimitsFor } from '../src/xlsx/storage/importSession'
describe('page storage budget (measured: F2 full stores 8,238,808 cells in 211.0 MiB of pages, 26.9 B/cell)', () => {
  it('the profile page budget covers the paged cell cap at the measured density, so F2 is not refused by quota', () => {
    for (const profile of ['D', 'M'] as const) {
      expect(quotaLimitsFor(profile).pages).toBeGreaterThanOrEqual(importLimitsFor(profile).paged.cells * 27)
    }
    expect(quotaLimitsFor('D').pages).toBeGreaterThan(211 * 1024 * 1024)
  })
})
