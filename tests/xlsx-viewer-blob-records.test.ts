/**
 * S13.4 gap A (spec §4.3.4): an oversized individual cell value is stored as a
 * bounded blob record in the page backend and referenced by its page, never
 * inline in the page's UTF-16 text lane. Every reader resolves the reference
 * to the identical raw UTF-16 value; cancellation and quota failures while a
 * blob is written remove every partial record with a typed outcome.
 */
import { describe, expect, it } from 'vitest'
import {
  F_BLOB, K_NUMBER, K_TEXT, PageBuilder, SourcePageStore, encodePage, indexedDbPageBackend, memoryPageBackend, opfsPageBackend, ownerLocks, pageOwner, sweepLeftoverPages, type PageBackend,
} from '../src/xlsx/storage/sourcePages'
import { fakeIndexedDb, fakeLockManager, fakeOpfsDirectory } from './helpers/fake-storage'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { importLimitsFor } from '../src/xlsx/storage/importSession'
import { createQuotaLedger, quotaLimitsFor } from '../src/xlsx/storage/quota'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { MemoryPool } from '../src/xlsx/view/tileCache'
import { WorkbookFind } from '../src/xlsx/view/find'
import { serializeSelection } from '../src/xlsx/view/copy'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { MAX_COLS } from '../src/xlsx/view/coordinates'
import type { SourceStore } from '../src/xlsx/view/contracts'
import { drain, findInputOf, workbookOf } from './helpers/xlsx-find'
import { xlsxBytes } from './helpers/xlsx-package'

const KiB = 1024
/** 200 KiB of UTF-16 text with a lone surrogate in the middle and a needle at the very end. */
const BIG_XML = `${'x'.repeat(100 * KiB)}_xD800_${'y'.repeat(50 * KiB - 1)}NEEDLE`
const BIG = `${'x'.repeat(100 * KiB)}\ud800${'y'.repeat(50 * KiB - 1)}NEEDLE`
const BIG_BYTES = 2 * BIG.length
const sheetData = [
  '<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="inlineStr"><is><t>small</t></is></c></row>',
  `<row r="2"><c r="A2"><v>2</v></c><c r="B2" t="inlineStr"><is><t>${BIG_XML}</t></is></c></row>`,
  '<row r="3"><c r="A3"><v>3</v></c></row>',
].join('')
const blobOf = (bytes: Uint8Array) => ({ kind: 'blob' as const, blob: new Blob([bytes as BlobPart]) })
const ro = (store: SourceStore) => ({ signal: new AbortController().signal, stamp: { ...store.stamp } })

function recordingBackend(hooks: { beforePut?: (key: string) => Promise<void> | void } = {}) {
  const inner = memoryPageBackend()
  const keys: string[] = []
  const backend: PageBackend = {
    kind: 'memory',
    async put(key, bytes) { await hooks.beforePut?.(key); keys.push(key); await inner.put(key, bytes) },
    get: (key) => inner.get(key),
    delete: (k) => inner.delete(k), removeAll: () => inner.removeAll(),
  }
  return { backend, keys, inner }
}

describe('oversized values become referenced blob records (spec §4.3.4)', () => {
  it('a 200 KiB single-cell string is written as one blob record; its page stays bounded and references it', async () => {
    expect(BIG.length).toBeGreaterThan(200 * KiB / 2)
    const rec = recordingBackend()
    const store = new SourcePageStore({ docId: 'd', backend: rec.backend })
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData })), { docId: 'd', profile: 'D', store })
    expect(await session.done).toMatchObject({ kind: 'complete' })
    const blobKeys = rec.keys.filter((k) => k.startsWith('blob/'))
    expect(blobKeys).toHaveLength(1)
    // No page carries the value inline: every page is far below the value's own size
    // and within the 64 KiB inline soft threshold.
    const pages = store.pages('1')
    expect(Math.max(...pages.map((p) => p.bytes))).toBeLessThan(importLimitsFor('D').inlineStringSoftBytes)
    expect(store.stats.blobs).toBe(1)
    expect(store.stats.blobBytes).toBeGreaterThanOrEqual(BIG_BYTES)
    // Charged to the storage ledger like pages, and released on close.
    expect(session.ledger.snapshot().current.blobRecords).toBe(store.stats.blobBytes)
    const sheet = createPagedWorkbookSource(session).store('1')
    expect(await sheet.readCell({ row: 1, col: 1 }, ro(sheet))).toMatchObject({ kind: 'cell', cell: { value: BIG } })
    expect(await sheet.readCell({ row: 0, col: 1 }, ro(sheet))).toMatchObject({ kind: 'cell', cell: { value: 'small' } })
    await session.release()
    expect(rec.inner.size).toBe(0)
    expect(session.ledger.snapshot().current.blobRecords).toBe(0)
  })

  it('values at the soft threshold stay inline; one code unit over spills', async () => {
    const units = importLimitsFor('D').inlineStringSoftBytes / 2
    const rows = `<row r="1"><c r="A1" t="inlineStr"><is><t>${'a'.repeat(units)}</t></is></c><c r="B1" t="inlineStr"><is><t>${'b'.repeat(units + 1)}</t></is></c></row>`
    const rec = recordingBackend()
    const store = new SourcePageStore({ docId: 'd', backend: rec.backend })
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rows })), { docId: 'd', profile: 'D', store })
    expect(await session.done).toMatchObject({ kind: 'complete' })
    expect(rec.keys.filter((k) => k.startsWith('blob/'))).toHaveLength(1)
    const sheet = createPagedWorkbookSource(session).store('1')
    expect(await sheet.readCell({ row: 0, col: 0 }, ro(sheet))).toMatchObject({ cell: { value: 'a'.repeat(units) } })
    expect(await sheet.readCell({ row: 0, col: 1 }, ro(sheet))).toMatchObject({ cell: { value: 'b'.repeat(units + 1) } })
    await session.release()
  })

  it('a blob is resolved lazily through the decoded-page cache (coalesced, evictable, reloaded identically)', async () => {
    const rec = recordingBackend()
    const store = new SourcePageStore({ docId: 'd', backend: rec.backend })
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData })), { docId: 'd', profile: 'D', store })
    await session.done
    const sheet = createPagedWorkbookSource(session).store('1')
    const before = store.stats.loads
    // Reading only small cells never loads the blob.
    await sheet.readCell({ row: 0, col: 1 }, ro(sheet))
    const pageOnly = store.stats.loads - before
    const [a, b] = await Promise.all([sheet.readCell({ row: 1, col: 1 }, ro(sheet)), sheet.readCell({ row: 1, col: 1 }, ro(sheet))])
    expect(a).toEqual(b)
    expect(store.stats.loads - before).toBe(pageOnly + 1)
    await session.release()
  })

  it('error cells and t="str" values over the threshold keep their kind and error flag', async () => {
    const long = 'z'.repeat(40 * KiB)
    const rows = `<row r="1"><c r="A1" t="str"><f>REPT("z",40960)</f><v>${long}</v></c><c r="B1" t="e"><v>${long}</v></c><c r="C1"><v>${long}</v></c></row>`
    const store = new SourcePageStore({ docId: 'd', backend: memoryPageBackend() })
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: rows })), { docId: 'd', profile: 'D', store })
    expect(await session.done).toMatchObject({ kind: 'complete' })
    expect(store.stats.blobs).toBe(3)
    const sheet = createPagedWorkbookSource(session).store('1')
    expect(await sheet.readCell({ row: 0, col: 0 }, ro(sheet))).toMatchObject({ cell: { value: long, formula: 'REPT("z",40960)', valueIsError: false } })
    expect(await sheet.readCell({ row: 0, col: 1 }, ro(sheet))).toMatchObject({ cell: { value: long, valueIsError: true } })
    expect(await sheet.readCell({ row: 0, col: 2 }, ro(sheet))).toMatchObject({ cell: { value: long } })
    await session.release()
  })
})

describe('cancellation and quota failure while a blob is written remove partial data', () => {
  it('cancellation while the blob write is in flight: cancelled, no record left, every reservation released', async () => {
    let release!: () => void
    const gate = new Promise<void>((resolve) => { release = resolve })
    let blobPutStarted!: () => void
    const started = new Promise<void>((resolve) => { blobPutStarted = resolve })
    const rec = recordingBackend({ async beforePut(key) { if (key.startsWith('blob/')) { blobPutStarted(); await gate } } })
    const ledger = createQuotaLedger(quotaLimitsFor('D'), { maxTotalExpandedBytes: 1 << 30 })
    const store = new SourcePageStore({ docId: 'd', backend: rec.backend, ledger })
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData })), { docId: 'd', profile: 'D', store, ledger, pageRows: 1 })
    await started
    const cancelled = session.cancel()
    release()
    await cancelled
    expect(await session.done).toEqual({ kind: 'cancelled' })
    expect(rec.inner.size).toBe(0)
    expect(Object.values(ledger.snapshot().current).every((v) => v === 0)).toBe(true)
    expect(store.pages('1')).toEqual([])
  })

  it('browser quota exhaustion on the blob write: typed quota failure and every stored record removed', async () => {
    const rec = recordingBackend({ beforePut(key) { if (key.startsWith('blob/')) throw new DOMException('full', 'QuotaExceededError') } })
    const store = new SourcePageStore({ docId: 'd', backend: rec.backend })
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData })), { docId: 'd', profile: 'D', store, pageRows: 1 })
    expect(await session.done).toMatchObject({ kind: 'incomplete', failure: { code: 'quota' } })
    expect(rec.inner.size).toBe(0)
    const sheet = createPagedWorkbookSource(session).store('1')
    expect((await sheet.readCell({ row: 1, col: 1 }, ro(sheet))).kind).toBe('error')
  })

  it('the blobRecords storage budget refuses with a typed quota failure and removes every stored record', async () => {
    const ledger = createQuotaLedger({ ...quotaLimitsFor('D'), blobRecords: 64 * KiB }, { maxTotalExpandedBytes: 1 << 30 })
    const rec = recordingBackend()
    const store = new SourcePageStore({ docId: 'd', backend: rec.backend, ledger })
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData })), { docId: 'd', profile: 'D', store, ledger, pageRows: 1 })
    expect(await session.done).toMatchObject({ kind: 'incomplete', failure: { code: 'quota' } })
    expect(rec.keys.some((k) => k.startsWith('page/1/'))).toBe(true) // a head page had landed
    expect(rec.inner.size).toBe(0)
    expect(Object.values(ledger.snapshot().current).every((v) => v === 0)).toBe(true)
  })
})

describe('find and copy read the full blob value', () => {
  it('find hits a needle past the inline threshold; copy returns the whole raw UTF-16 value', async () => {
    const store = new SourcePageStore({ docId: 'doc', backend: memoryPageBackend() })
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData })), { docId: 'doc', profile: 'D', store })
    await session.done
    const source = createPagedWorkbookSource(session)
    const { input } = findInputOf(workbookOf([]), { docId: 'doc', source })
    const hits = await drain(new WorkbookFind(input, { text: 'needle', caseSensitive: false }))
    expect(hits.map((h) => [h.source.row, h.source.col, h.start, h.end])).toEqual([[1, 1, BIG.length - 6, BIG.length]])

    const sheet = source.store('1')
    const owner = { ...sheet.stamp, fontGen: 0, viewGen: 0 }
    const out = await serializeSelection({
      store: sheet, owner, currentOwner: () => owner,
      selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 1, r1: 1 }] }, cols: [{ c0: 1, c1: 1 }] }] },
      order: createIdentityOrder(), visibility: visibilityFromHidden(createIntervalSet([])), columns: visibilityFromHidden(createIntervalSet([]), MAX_COLS),
      allowed: () => true, signal: new AbortController().signal,
    })
    expect(out).toMatchObject({ kind: 'ready' })
    expect(out.kind === 'ready' && out.text === BIG).toBe(true)
    await source.release()
  })
})

describe('blob records in the browser page backends', () => {
  it('OPFS and IndexedDB store the blob next to the pages and remove it on release', async () => {
    const dir = fakeOpfsDirectory()
    const opfsStore = new SourcePageStore({ docId: 'd', backend: opfsPageBackend(dir, 'd') })
    const a = await openImportSession(blobOf(xlsxBytes({ sheetData })), { docId: 'd', profile: 'D', store: opfsStore })
    expect(await a.done).toMatchObject({ kind: 'complete' })
    expect([...dir.files.keys()].some((name) => name.includes('blob'))).toBe(true)
    const sheetA = createPagedWorkbookSource(a).store('1')
    expect(await sheetA.readCell({ row: 1, col: 1 }, ro(sheetA))).toMatchObject({ cell: { value: BIG } })
    await a.release()
    expect(dir.files.size).toBe(0)

    const idb = fakeIndexedDb()
    const idbStore = new SourcePageStore({ docId: 'd', backend: indexedDbPageBackend(idb, 'd') })
    const b = await openImportSession(blobOf(xlsxBytes({ sheetData })), { docId: 'd', profile: 'D', store: idbStore })
    expect(await b.done).toMatchObject({ kind: 'complete' })
    const sheetB = createPagedWorkbookSource(b).store('1')
    expect(await sheetB.readCell({ row: 1, col: 1 }, ro(sheetB))).toMatchObject({ cell: { value: BIG } })
    await b.release()
    expect(idb.deleted).toContain('officeview-pages-d')
  })
})

describe('PageBuilder spill', () => {
  it('keeps the record kind, sets F_BLOB with an ordinal, adds no text units and keeps bytes exact', () => {
    const b = new PageBuilder({ inlineTextBytes: 8 })
    b.add(0, 0, K_TEXT, 0, 0, 'abcd', true, undefined)
    b.add(0, 1, K_TEXT, 0, 0, 'abcde\ud800', true, undefined)
    expect(b.blobs).toEqual(['abcde\ud800'])
    expect(b.spilledBytes).toBe(12)
    const page = b.build()
    expect(page.text.length).toBe(4)
    expect(page.flags[1] & F_BLOB).toBe(F_BLOB)
    expect(page.flags[0] & F_BLOB).toBe(0)
    expect(page.kinds[1]).toBe(K_TEXT)
    expect(page.nums[1]).toBe(0)
    expect(b.bytes).toBe(encodePage(page).length)
  })
})

describe('a failed page removes the blob records it already wrote (storage, ledger and pool return to the pre-page state)', () => {
  const big = (fill: string) => fill.repeat(40 * KiB)
  /** A page of `numbers` numeric records followed by the given oversized texts. */
  function pageOf(pageId: string, r0: number, numbers: number, texts: string[]) {
    const b = new PageBuilder()
    for (let i = 0; i < numbers; i++) b.add(r0, i % 16_000, K_NUMBER, 0, i, '', true, undefined)
    texts.forEach((t, i) => b.add(r0 + 1, i, K_TEXT, 0, 0, t, true, undefined))
    return { sheetId: '1', pageId, r0, r1: r0 + 1, cellCount: b.records, bytes: b.bytes, payload: b.build(), blobs: [...b.blobs] }
  }
  async function setup(backend: PageBackend, poolBytes = 256 * 1024 * 1024) {
    const ledger = createQuotaLedger(quotaLimitsFor('D'), { maxTotalExpandedBytes: 1 << 30 })
    const pool = new MemoryPool({ limitBytes: poolBytes })
    const store = new SourcePageStore({ docId: 'd', backend, ledger })
    store.bindMemory(pool)
    await store.commit(pageOf('p0', 0, 4, [big('h')]))
    const snapshot = () => ({ current: ledger.snapshot().current, stats: { ...store.stats, loads: 0 }, pool: { retained: pool.stats().retainedBytes, transient: pool.stats().transientBytes }, pages: store.pages('1').length })
    return { store, ledger, pool, before: snapshot(), snapshot }
  }

  it('memory refusal while encoding the page after its blobs were written', async () => {
    const rec = recordingBackend()
    // Room for the head page and the blob scratch, not for the large page's encode buffer.
    const { store, before, snapshot } = await setup(rec.backend, 512 * 1024)
    const sizeBefore = rec.inner.size
    const page = pageOf('p1', 10, 40_000, [big('a'), big('b')])
    const ordinals = [...page.payload.nums.subarray(40_000)]
    await expect(store.commit(page)).rejects.toMatchObject({ code: 'resource-limit', reason: 'memory-budget' })
    expect(rec.keys.filter((k) => k.startsWith('blob/'))).toHaveLength(3)
    expect(rec.inner.size).toBe(sizeBefore)
    expect(snapshot()).toEqual(before)
    // The payload is restored too: its F_BLOB records hold their ordinals again.
    expect([...page.payload.nums.subarray(40_000)]).toEqual(ordinals)
  })

  it('parse errors: a missing referenced value, or an unreferenced one, after earlier blobs were written', async () => {
    const rec = recordingBackend()
    const { store, before, snapshot } = await setup(rec.backend)
    const sizeBefore = rec.inner.size
    const missing = pageOf('p1', 10, 2, [big('a'), big('b')])
    await expect(store.commit({ ...missing, blobs: missing.blobs.slice(0, 1) })).rejects.toMatchObject({ code: 'parse-error' })
    expect(rec.inner.size).toBe(sizeBefore)
    expect(snapshot()).toEqual(before)
    const extra = pageOf('p1', 10, 2, [big('c')])
    await expect(store.commit({ ...extra, blobs: [...extra.blobs, big('d')] })).rejects.toMatchObject({ code: 'parse-error' })
    expect(rec.inner.size).toBe(sizeBefore)
    expect(snapshot()).toEqual(before)
  })

  it('a storage failure on the page write after its blobs: the blobs and the page charge are removed', async () => {
    const rec = recordingBackend({ beforePut(key) { if (key === 'page/1/p1') throw new DOMException('disk', 'UnknownError') } })
    const { store, before, snapshot } = await setup(rec.backend)
    const sizeBefore = rec.inner.size
    await expect(store.commit(pageOf('p1', 10, 2, [big('a'), big('b')]))).rejects.toMatchObject({ code: 'storage-error' })
    expect(rec.inner.size).toBe(sizeBefore)
    expect(snapshot()).toEqual(before)
  })

  it('a partial record left by a failed blob write is deleted and uncharged', async () => {
    let n = 0
    const rec = recordingBackend()
    const partial: PageBackend = {
      ...rec.backend,
      async put(key, bytes) { if (key.startsWith('blob/') && ++n === 3) { await rec.inner.put(key, bytes); throw new DOMException('disk', 'UnknownError') } await rec.backend.put(key, bytes) },
    }
    const { store, before, snapshot } = await setup(partial)
    const sizeBefore = rec.inner.size
    await expect(store.commit(pageOf('p1', 10, 2, [big('a'), big('b')]))).rejects.toMatchObject({ code: 'storage-error' })
    expect(rec.inner.size).toBe(sizeBefore)
    expect(snapshot()).toEqual(before)
  })

  it('cancellation mid-page (abandon during a blob write) removes every record and releases every charge', async () => {
    let release!: () => void
    const gate = new Promise<void>((resolve) => { release = resolve })
    let reached!: () => void
    const started = new Promise<void>((resolve) => { reached = resolve })
    let n = 0
    const rec = recordingBackend({ async beforePut(key) { if (key.startsWith('blob/') && ++n === 3) { reached(); await gate } } })
    const { store, ledger, pool } = await setup(rec.backend)
    const pending = store.commit(pageOf('p1', 10, 2, [big('a'), big('b')]))
    await started
    await store.abandon('d')
    release()
    await expect(pending).rejects.toMatchObject({ code: 'cancelled' })
    expect(rec.inner.size).toBe(0)
    expect(Object.values(ledger.snapshot().current).every((v) => v === 0)).toBe(true)
    expect(pool.stats().totalBytes).toBe(0)
  })

  it('OPFS and IndexedDB delete single records', async () => {
    const dir = fakeOpfsDirectory()
    const opfs = opfsPageBackend(dir, 'd')
    await opfs.put('blob/0', new Uint8Array([1]))
    await opfs.put('1/p0', new Uint8Array([2]))
    await opfs.delete('blob/0')
    await opfs.delete('blob/0')
    expect(await opfs.get('blob/0')).toBeUndefined()
    expect(dir.files.size).toBe(1)
    const idb = indexedDbPageBackend(fakeIndexedDb(), 'd')
    await idb.put('blob/0', new Uint8Array([1]))
    await idb.put('1/p0', new Uint8Array([2]))
    await idb.delete('blob/0')
    expect(await idb.get('blob/0')).toBeUndefined()
    expect([...(await idb.get('1/p0'))!]).toEqual([2])
  })
})

describe('review round 1 (codex) MAJOR 5: rejected commits are validated before any write', () => {
  const textPage = (pageId: string, r0: number, r1: number, text: string) => {
    const b = new PageBuilder()
    b.add(r0, 0, K_TEXT, 0, 0, text, true, undefined)
    return { sheetId: '1', pageId, r0, r1, cellCount: 1, bytes: b.bytes, payload: b.build(), blobs: [...b.blobs] }
  }
  it('an overlapping page (same or new id) and a repeated shared-string chunk leave stored bytes, summaries, charges and blobs intact', async () => {
    const rec = recordingBackend()
    const ledger = createQuotaLedger(quotaLimitsFor('D'), { maxTotalExpandedBytes: 1 << 30 })
    const pool = new MemoryPool({ limitBytes: 64 * 1024 * 1024 })
    const store = new SourcePageStore({ docId: 'd', backend: rec.backend, ledger })
    store.bindMemory(pool)
    await store.commit(textPage('p0', 0, 0, 'a'.repeat(40_000)))
    await store.commitStrings(0, ['s0'], 4)
    const snapshot = () => ({ size: rec.inner.size, writes: rec.keys.length, current: ledger.snapshot().current, blobs: store.stats.blobs, pages: store.pages('1').map((p) => ({ ...p })), retained: pool.stats().retainedBytes })
    const before = snapshot()
    await expect(store.commit(textPage('p0', 0, 0, 'b'.repeat(40_000)))).rejects.toThrow(/overlaps/)
    await expect(store.commit(textPage('p1', 0, 3, 'c'.repeat(40_000)))).rejects.toThrow(/overlaps/)
    await expect(store.commit(textPage('p0', 5, 6, 'd'.repeat(40_000)))).rejects.toThrow(/already committed/)
    await expect(store.commitStrings(0, ['other'], 4)).rejects.toThrow(/already committed/)
    expect(snapshot()).toEqual(before)
    expect(await store.blob(0)).toBe('a'.repeat(40_000))
    expect(await store.sharedString(0)).toBe('s0')
    const sheetPage = await store.page('1', 'p0')
    expect(sheetPage.nums[0]).toBe(0)
  })
})

describe('crash-orphan sweep (Web Locks ownership, codex round 2)', () => {
  const HOUR = 3600_000
  const esc = (t: string) => t.replace(/[^A-Za-z0-9.]/g, (c) => `_${c.charCodeAt(0).toString(16).padStart(4, '0')}`)
  const opfsName = (owner: string, key: string) => `officeview-pages-${esc(owner)}-${esc(key)}`
  async function writeRaw(dir: ReturnType<typeof fakeOpfsDirectory>, name: string) {
    const handle = await dir.getFileHandle(name, { create: true }) as { createWritable(): Promise<{ write(b: Uint8Array): Promise<void>; close(): Promise<void> }> }
    const w = await handle.createWritable()
    await w.write(new Uint8Array([1, 2, 3]))
    await w.close()
  }
  const textPage = (sheetId: string, pageId: string, r0: number, text: string) => {
    const b = new PageBuilder()
    b.add(r0, 0, K_TEXT, 0, 0, text, true, undefined)
    return { sheetId, pageId, r0, r1: r0, cellCount: 1, bytes: b.bytes, payload: b.build(), blobs: [...b.blobs] }
  }

  it('owners are collision-resistant and held exclusively across realms', async () => {
    const now = Date.UTC(2026, 9, 13)
    expect(pageOwner('doc', now)).not.toBe(pageOwner('doc', now))
    const mgr = fakeLockManager()
    const realmA = ownerLocks(mgr), realmB = ownerLocks(mgr)
    const owner = pageOwner('doc', now)
    const release = await realmA.hold(owner)
    await expect(realmB.hold(owner)).rejects.toMatchObject({ code: 'storage-error' })
    expect(await realmB.tryWith(owner, async () => {})).toBe(false)
    release()
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(await realmB.tryWith(owner, async () => {})).toBe(true)
  })

  it("session start: realm B never deletes realm A's live storage (even 25 h old); dead owners go; no Web Locks means no sweep", async () => {
    const now = Date.UTC(2026, 9, 13)
    const mgr = fakeLockManager()
    const dir = fakeOpfsDirectory()
    const idb = fakeIndexedDb()
    // Realm A: a live session whose owner looks 25 h old (long-lived tab or a forward clock jump).
    const liveOld = pageOwner('doc-a', now - 25 * HOUR)
    const releaseA = await ownerLocks(mgr).hold(liveOld)
    const a = new SourcePageStore({ docId: 'a', backend: opfsPageBackend(dir, liveOld), ownerLock: releaseA })
    await a.commit(textPage('1', 'p0', 0, 'live'))
    // Leftovers of a dead session, and a fresh owner of another live tab.
    const dead = pageOwner('doc-b', now - 26 * HOUR), fresh = pageOwner('doc-c', now - HOUR)
    await writeRaw(dir, opfsName(dead, 'blob/0'))
    await writeRaw(dir, opfsName(dead, 'page/1/p0'))
    await writeRaw(dir, opfsName(fresh, 'page/1/p0'))
    await writeRaw(dir, 'unrelated.bin')
    for (const owner of [dead, fresh]) await new Promise<void>((resolve) => { const req = idb.open(`officeview-pages-${owner}`, 1); req.onsuccess = () => resolve() })

    expect(await sweepLeftoverPages({ opfs: async () => dir, indexedDb: idb }, { now })).toBe(0) // no Web Locks: nothing is provably dead
    expect(await sweepLeftoverPages({ opfs: async () => dir, indexedDb: idb, locks: ownerLocks(mgr) }, { now })).toBe(3)
    expect(idb.deleted).toEqual([`officeview-pages-${dead}`])
    expect([...dir.files.keys()].some((n) => n.includes(esc(dead)))).toBe(false)
    expect([...dir.files.keys()].some((n) => n.includes(esc(fresh)))).toBe(true)
    expect(dir.files.has('unrelated.bin')).toBe(true)
    // A's page survived realm B's sweep.
    expect(textAtPage(await a.page('1', 'p0'))).toBe('live')
    // Cancelled and bounded sweeps remove nothing.
    await writeRaw(dir, opfsName(dead, 'blob/1'))
    const aborted = new AbortController()
    aborted.abort()
    expect(await sweepLeftoverPages({ opfs: async () => dir, locks: ownerLocks(mgr) }, { now, signal: aborted.signal })).toBe(0)
    expect(await sweepLeftoverPages({ opfs: async () => dir, locks: ownerLocks(mgr) }, { now, maxEntries: 1 })).toBe(0)
    await a.close()
    expect(mgr.held.size).toBe(0)
  })

  it("store-open cleanup never touches another writer's records: without the owner lock it skips", async () => {
    const mgr = fakeLockManager()
    const dir = fakeOpfsDirectory()
    const owner = pageOwner('doc', Date.now())
    let release!: () => void
    const gate = new Promise<void>((resolve) => { release = resolve })
    const inner = opfsPageBackend(dir, owner)
    const slowPage: PageBackend = { ...inner, put: async (key, bytes) => { if (key.startsWith('page/')) await gate; await inner.put(key, bytes) } }
    const a = new SourcePageStore({ docId: 'a', backend: slowPage, ownerLock: await ownerLocks(mgr).hold(owner) })
    const pending = a.commit(textPage('1', 'p0', 0, 'x'.repeat(40_000)))
    await new Promise((resolve) => setTimeout(resolve, 10))
    // A second store on the same owner (the collision codex reproduced) cannot get the lock, so it cannot sweep.
    await expect(ownerLocks(mgr).hold(owner)).rejects.toMatchObject({ code: 'storage-error' })
    const b = new SourcePageStore({ docId: 'b', backend: opfsPageBackend(dir, owner) })
    expect(await b.sweepUnreferenced()).toMatchObject({ removed: 0, skipped: expect.stringContaining('owner lock') })
    release()
    await pending
    expect(await a.blob(0)).toBe('x'.repeat(40_000))
    await a.close()
  })

  it('simulated crash (blob written, page not): once the dead realm\'s lock is gone the next holder removes the orphan; ledger matches storage', async () => {
    const mgr = fakeLockManager()
    const dir = fakeOpfsDirectory()
    const owner = pageOwner('doc', Date.now())
    const crashing = opfsPageBackend(dir, owner)
    const hung: PageBackend = { ...crashing, put: (key, bytes) => key === 'page/1/p1' ? new Promise<void>(() => {}) : crashing.put(key, bytes) }
    const releaseA = await ownerLocks(mgr).hold(owner)
    const a = new SourcePageStore({ docId: 'd', backend: hung, ownerLock: releaseA })
    await a.commit(textPage('1', 'p0', 0, 'a'.repeat(40_000)))
    void a.commit(textPage('1', 'p1', 1, 'b'.repeat(40_000)))
    await new Promise((resolve) => setTimeout(resolve, 20))
    expect(dir.files.size).toBe(3)
    // The process dies: the browser drops its locks; nothing ran its cleanup.
    releaseA()
    await new Promise((resolve) => setTimeout(resolve, 0))
    const ledger = createQuotaLedger(quotaLimitsFor('D'), { maxTotalExpandedBytes: 1 << 30 })
    const pool = new MemoryPool({ limitBytes: 64 * 1024 * 1024 })
    const b = new SourcePageStore({ docId: 'd', backend: opfsPageBackend(dir, owner), ledger, ownerLock: await ownerLocks(mgr).hold(owner) })
    b.bindMemory(pool)
    expect(await b.sweepUnreferenced()).toEqual({ removed: 3 })
    expect(dir.files.size).toBe(0)
    expect(Object.values(ledger.snapshot().current).every((v) => v === 0)).toBe(true)
    expect(pool.stats().totalBytes).toBe(0)
    // Then an import through the store: storage holds exactly what it committed and charged.
    await writeRaw(dir, opfsName(owner, 'blob/7'))
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData })), { docId: 'd', profile: 'D', store: b, ledger })
    expect(await session.done).toMatchObject({ kind: 'complete' })
    expect([...dir.files.keys()].some((name) => name.endsWith('blob_002f7'))).toBe(false)
    expect(dir.files.size).toBe(b.pages('1').length + b.stats.blobs)
    expect(ledger.snapshot().current.pages + ledger.snapshot().current.blobRecords).toBe(b.stats.storedBytes + b.stats.blobBytes)
    expect(await b.sweepUnreferenced()).toEqual({ removed: 0 })
    await session.release()
    expect(mgr.held.size).toBe(0)
  })

  it('pages of sheets named like reserved namespaces ("sst", "blob", ids with "/") are never classified as orphans', async () => {
    const store = new SourcePageStore({ docId: 'd', backend: memoryPageBackend() })
    const num = (sheetId: string, pageId: string, value: number) => {
      const b = new PageBuilder()
      b.add(0, 0, K_NUMBER, 0, value, '', true, undefined)
      return { sheetId, pageId, r0: 0, r1: 0, cellCount: 1, bytes: b.bytes, payload: b.build() }
    }
    await store.commit(num('sst', 'p0', 42))
    await store.commit(num('blob', '0', 43))
    await store.commit(num('a/b', 'c', 44))
    await store.commit(num('a', 'b/c', 45))
    expect(await store.sweepUnreferenced()).toEqual({ removed: 0 })
    expect((await store.page('sst', 'p0')).nums[0]).toBe(42)
    expect((await store.page('blob', '0')).nums[0]).toBe(43)
    expect((await store.page('a/b', 'c')).nums[0]).toBe(44)
    expect((await store.page('a', 'b/c')).nums[0]).toBe(45)
  })

  it('listing is bounded by entries examined (not owner matches) and cancellable; a listing failure skips cleanup without failing the import', async () => {
    const dir = fakeOpfsDirectory()
    for (let i = 0; i < 5000; i++) await writeRaw(dir, `foreign-${i}`)
    const backend = opfsPageBackend(dir, 'owner-x')
    await backend.put('blob/0', new Uint8Array([1]))
    expect(await backend.list!(4096)).toEqual([])
    expect(await backend.list!(10_000)).toEqual(['blob/0'])
    const aborted = new AbortController()
    aborted.abort()
    expect(await backend.list!(10_000, aborted.signal)).toEqual([])

    const inner = memoryPageBackend()
    const flaky: PageBackend = { ...inner, list: async () => { throw new Error('temporary enumeration failure') } }
    const store = new SourcePageStore({ docId: 'd', backend: flaky })
    const session = await openImportSession(blobOf(xlsxBytes({ sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>' })), { docId: 'd', profile: 'D', store })
    expect(await session.done).toMatchObject({ kind: 'complete' })
    expect(session.diagnostics.map((d) => d.code)).toContain('storage-cleanup-skipped')
    await session.release()
  })
})
const textAtPage = (p: { text: Uint16Array; textStart: Uint32Array }) => String.fromCharCode(...p.text.subarray(p.textStart[0], p.textStart[1]))
