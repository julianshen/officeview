/**
 * S13.4 gap A (spec §4.3.4): an oversized individual cell value is stored as a
 * bounded blob record in the page backend and referenced by its page, never
 * inline in the page's UTF-16 text lane. Every reader resolves the reference
 * to the identical raw UTF-16 value; cancellation and quota failures while a
 * blob is written remove every partial record with a typed outcome.
 */
import { describe, expect, it } from 'vitest'
import { SourcePageStore, memoryPageBackend, type PageBackend } from '../src/xlsx/storage/sourcePages'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { importLimitsFor } from '../src/xlsx/storage/importSession'
import { createQuotaLedger, quotaLimitsFor } from '../src/xlsx/storage/quota'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
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
    removeAll: () => inner.removeAll(),
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
    expect(rec.keys.some((k) => k.startsWith('1/'))).toBe(true) // a head page had landed
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
