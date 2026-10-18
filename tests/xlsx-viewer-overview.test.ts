/**
 * S1.2 overview and scan RED cases on the paged adapter: empty wide bands and
 * whole-page groups are answered from page summaries without loading pages,
 * a cut page is loaded and counted exactly, injected storage errors surface as
 * typed failures (never blank), stale stamps are refused, and a 1M-row-shaped
 * 8-column scan runs in bounded batches with bounded page residency.
 */
import { describe, expect, it } from 'vitest'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { SourcePageStore, memoryPageBackend, opfsPageBackend, type PageBackend } from '../src/xlsx/storage/sourcePages'
import { fakeOpfsDirectory } from './helpers/fake-storage'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import type { SourceStore } from '../src/xlsx/view/contracts'
import { colName, repeat, xlsxBytes } from './helpers/xlsx-package'

const ro = (store: SourceStore) => ({ signal: new AbortController().signal, stamp: { ...store.stamp } })
const grid = (rows: number, width: number, style = 0) => repeat(rows, (r) => `<row r="${r + 1}">${repeat(width, (c) => `<c r="${colName(c)}${r + 1}"${style ? ` s="${style}"` : ''}><v>${r + c}</v></c>`)}</row>`)

async function paged(bytes: Uint8Array, backend: PageBackend = memoryPageBackend(), pageRows = 100, cacheBytes?: number) {
  const store = new SourcePageStore({ docId: 'd', backend, cacheBytes })
  const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes as BlobPart]) }, { docId: 'd', profile: 'D', store, pageRows })
  expect(await session.done).toMatchObject({ kind: 'complete' })
  const source = createPagedWorkbookSource(session)
  return { store, session, source, sheet: source.store('1') }
}

describe('paged queryOverview', () => {
  it('answers empty wide bands and whole-page groups from summaries without loading a page', async () => {
    const { store, sheet, source } = await paged(xlsxBytes({ sheetData: grid(400, 4, 1), styles: '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="0"/></cellXfs>' }))
    const before = store.stats.loads
    const page = await sheet.queryOverview({
      maxGroups: 8,
      groups: [
        { groupId: 'wide-empty', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 1_048_575 }] }, cols: [{ c0: 10, c1: 16_383 }] },
        { groupId: 'below', rows: { kind: 'intervals', intervals: [{ r0: 400, r1: 1_048_575 }] }, cols: [{ c0: 0, c1: 3 }] },
        { groupId: 'pages-0-1', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 199 }] }, cols: [{ c0: 0, c1: 3 }] },
      ],
    }, ro(sheet))
    expect(store.stats.loads).toBe(before)
    expect(page.completeness).toBe('complete')
    expect(page.items).toEqual([
      { groupId: 'wide-empty', populatedCount: 0, countKind: 'exact', style: { kind: 'uniform', styleIndex: 0 }, content: 'blank' },
      { groupId: 'below', populatedCount: 0, countKind: 'exact', style: { kind: 'uniform', styleIndex: 0 }, content: 'blank' },
      { groupId: 'pages-0-1', populatedCount: 800, countKind: 'exact', style: { kind: 'uniform', styleIndex: 1 }, content: 'populated' },
    ])
    await source.release()
  })

  it('loads only the cut page and counts it exactly, including noncontiguous rows and sparse column chunks', async () => {
    const { store, sheet, source } = await paged(xlsxBytes({ sheetData: grid(400, 4) }))
    const before = store.stats.loads
    const page = await sheet.queryOverview({
      maxGroups: 8,
      groups: [
        { groupId: 'cut', rows: { kind: 'list', r0: 150, r1: 152, rows: new Uint32Array([150, 152]) }, cols: [{ c0: 1, c1: 1 }, { c0: 3, c1: 5 }] },
      ],
    }, ro(sheet))
    expect(store.stats.loads - before).toBe(1)
    expect(page.items).toEqual([{ groupId: 'cut', populatedCount: 4, countKind: 'exact', style: { kind: 'uniform', styleIndex: 0 }, content: 'mixed' }])
    await source.release()
  })

  it('an injected storage error is a typed failure for cells, columns and overview; a stale stamp is refused', async () => {
    const inner = memoryPageBackend()
    let failReads = false
    const backend: PageBackend = {
      kind: 'memory', put: (k, b) => inner.put(k, b), delete: (k) => inner.delete(k), removeAll: () => inner.removeAll(),
      async get(key) { if (failReads) throw Object.assign(new Error('disk gone'), { code: 'storage-error' }); return inner.get(key) },
    }
    const { sheet, source } = await paged(xlsxBytes({ sheetData: grid(300, 2) }), backend)
    failReads = true
    expect(await sheet.readCell({ row: 250, col: 1 }, ro(sheet))).toEqual({ kind: 'error', code: 'storage-error', retryable: true })
    const col = await sheet.queryColumn({ col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 299 }] }, limit: 10 }, ro(sheet))
    expect(col).toMatchObject({ completeness: 'incomplete', error: { code: 'storage-error' } })
    const ov = await sheet.queryOverview({ maxGroups: 2, groups: [{ groupId: 'g', rows: { kind: 'intervals', intervals: [{ r0: 10, r1: 20 }] }, cols: [{ c0: 0, c1: 0 }] }] }, ro(sheet))
    expect(ov).toMatchObject({ completeness: 'incomplete', error: { code: 'storage-error' } })
    const stale = await sheet.queryOverview({ maxGroups: 2, groups: [] }, { signal: new AbortController().signal, stamp: { ...sheet.stamp, valueGen: 7 } })
    expect(stale.error?.code).toBe('stale-source')
    await source.release()
  })
})

describe('paged scan batches (1M×8 shape, reduced rows in CI)', () => {
  it('scans 8 columns in bounded batches with bounded decoded-page residency and exact totals', async () => {
    const rows = 20_000
    // A persistent backend: decoded pages are fresh copies the LRU must bound
    // (a memory backend's decoded views share its already-charged buffers).
    const { store, sheet, source } = await paged(xlsxBytes({ sheetData: grid(rows, 8) }), opfsPageBackend(fakeOpfsDirectory(), 'd'), 2_000, 256 * 1024)
    // ten 2,000-row pages plus the empty tail page that proves rows 20,000+ blank
    expect(store.pages('1').map((p) => p.aggregate.records)).toEqual([...Array(10).fill(16_000), 0])
    let batches = 0, cells = 0, maxBatch = 0, sum = 0
    for await (const page of sheet.scanPopulated({ range: { r0: 0, r1: 1_048_575, c0: 0, c1: 7 }, limit: 10_000 }, ro(sheet))) {
      expect(page.error).toBeUndefined()
      batches++
      cells += page.items.length
      maxBatch = Math.max(maxBatch, page.items.length)
      for (const cell of page.items) sum += cell.value as number
    }
    expect(cells).toBe(rows * 8)
    expect(maxBatch).toBeLessThanOrEqual(10_000)
    expect(batches).toBe(17)
    // sum over r,c of (r + c) = 8 * sum(r) + rows * sum(c)
    expect(sum).toBe(8 * (rows * (rows - 1)) / 2 + rows * 28)
    // The local LRU kept decoded pages near its budget while all 10 pages were read.
    expect(store.stats.peakResidentBytes).toBeLessThanOrEqual(256 * 1024 + 300_000)
    expect(store.stats.evictions).toBeGreaterThan(0)
    await source.release()
  })
})

describe('review MINOR 3: paged queries stop at the next page once cancelled', () => {
  it('readRows, queryColumn, overview and scanPopulated abort after the page in flight', async () => {
    const inner = memoryPageBackend()
    let onGet: (() => void) | undefined
    const backend: PageBackend = { kind: 'memory', put: (k, b) => inner.put(k, b), delete: (k) => inner.delete(k), removeAll: () => inner.removeAll(), async get(k) { onGet?.(); return inner.get(k) } }
    const { store, sheet, source } = await paged(xlsxBytes({ sheetData: grid(1000, 2) }), backend, 100, 1)
    const run = async (label: string, query: (signal: AbortSignal) => Promise<{ error?: { code: string } }>) => {
      const controller = new AbortController()
      const before = store.stats.loads
      onGet = () => controller.abort()
      const page = await query(controller.signal)
      onGet = undefined
      expect(page.error?.code, label).toBe('cancelled')
      expect(store.stats.loads - before, label).toBe(1)
    }
    const o = (signal: AbortSignal) => ({ signal, stamp: { ...sheet.stamp } })
    await run('readRows', (signal) => sheet.readRows({ r0: 0, r1: 999, c0: 0, c1: 1 }, undefined, o(signal)))
    await run('queryColumn', (signal) => sheet.queryColumn({ col: 1, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 999 }] }, limit: 100_000, encoding: 'lanes' }, o(signal)))
    await run('overview', (signal) => sheet.queryOverview({ maxGroups: 1, groups: [{ groupId: 'g', rows: { kind: 'list', r0: 5, r1: 995, rows: Uint32Array.from({ length: 100 }, (_, i) => i * 10 + 5) }, cols: [{ c0: 0, c1: 0 }] }] }, o(signal)))
    await run('scanPopulated', async (signal) => {
      let last: { error?: { code: string } } = {}
      for await (const page of sheet.scanPopulated({ range: { r0: 0, r1: 999, c0: 0, c1: 1 }, limit: 100_000 }, o(signal))) last = page
      return last
    })
    await source.release()
  })
})
