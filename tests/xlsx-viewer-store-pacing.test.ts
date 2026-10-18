/** Shipped SourceStores decode one request in steps (decision 10; spec §10.2):
 * at most 4,096 cells of store work between real yields, with signal and stamp
 * checkpoints. A request's limit and the call count are unchanged, so a
 * 65,536-cell strip is still one request. */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { PagedXlsxSheetAdapter, createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { SourcePageStore, memoryPageBackend } from '../src/xlsx/storage/sourcePages'
import { DocumentSession } from '../src/worker/viewerSession'
import { semanticsFor } from '../src/xlsx/view/semanticsFor'
import { buildZip } from './helpers/zip-builder'
import { xlsxEntries } from './helpers/xlsx-package'
import type { SourceStore } from '../src/xlsx/view/contracts'
import type { XlsxDocument } from '../src/xlsx/types'

const HEADER = { r0: 0, r1: 1, c0: 0, c1: 8191 } // codex's wide header read: 16,384 cells
const STRIP = { r0: 0, r1: 8191, c0: 0, c1: 7 } // one 65,536-cell strip request
type Pacing = { now: () => number; yield: () => Promise<void> }

/** Object document whose cells count store work: a cell counts once per slice
 * in which the store reads its row or column. */
function objectDoc(rows: number, cols: number) {
  const work: { cells: number; slice: number; onTouch?: () => void } = { cells: 0, slice: 0 }
  const docRows = Array.from({ length: rows }, (_, row) => ({ index: row, cells: Array.from({ length: cols }, (_, col) => {
    let seen = -1
    const touch = () => { if (seen !== work.slice) { seen = work.slice; work.cells++; work.onTouch?.() } }
    return Object.defineProperties({ ref: '', value: row * cols + col, styleIndex: 0 } as Record<string, unknown>, {
      row: { enumerable: true, get() { touch(); return row } },
      col: { enumerable: true, get() { touch(); return col } },
    })
  }) }))
  const doc = { sheets: [{ name: 's', sheetId: 's', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], semantics: semanticsFor('1900'), rows: docRows }], tables: [] } as unknown as XlsxDocument
  return { doc, work }
}
function frozen(work: { cells: number; slice: number }) {
  const slices: number[] = []
  let since = 0
  const pacing: Pacing = { now: () => 0, yield: async () => { slices.push(work.cells - since); since = work.cells; work.slice++ } }
  return { pacing, slices, done: () => { slices.push(work.cells - since); return slices } }
}
const options = (store: SourceStore, signal = new AbortController().signal) => ({ signal, stamp: { ...store.stamp } })
async function firstPage<T>(pages: AsyncIterable<T>): Promise<T> {
  for await (const page of pages) return page
  throw new Error('no page')
}
const cellsOf = (page: { items: readonly { cells: readonly unknown[] }[] }) => page.items.reduce((sum, item) => sum + item.cells.length, 0)

describe('object store decodes a request in bounded steps', () => {
  it('splits codex\'s 16,384-cell header read', async () => {
    const { doc, work } = objectDoc(2, 8192), probe = frozen(work)
    const store = createObjectWorkbookSource(doc, 'd', { pacing: probe.pacing }).store('s')
    work.cells = 0; work.slice++
    const page = await store.readRows(HEADER, undefined, options(store))
    expect(page.error).toBeUndefined()
    expect(cellsOf(page)).toBe(16_384)
    const slices = probe.done()
    expect(slices.length - 1).toBeGreaterThanOrEqual(3)
    expect(Math.max(...slices)).toBeLessThanOrEqual(4096)
  })
  it('splits one 65,536-cell strip without splitting the request', async () => {
    const { doc, work } = objectDoc(8192, 8), probe = frozen(work)
    const store = createObjectWorkbookSource(doc, 'd', { pacing: probe.pacing }).store('s')
    work.cells = 0; work.slice++
    const page = await firstPage(store.scanPopulated({ range: STRIP, limit: 65_536 }, options(store)))
    expect(page.items.length).toBe(65_536)
    const slices = probe.done()
    expect(slices.length - 1).toBeGreaterThanOrEqual(15)
    expect(Math.max(...slices)).toBeLessThanOrEqual(4096)
  })
  it.each(['rows', 'scan'] as const)('services a queued macrotask abort mid-request (%s)', async kind => {
    const { doc, work } = kind === 'rows' ? objectDoc(2, 8192) : objectDoc(8192, 8)
    const store = createObjectWorkbookSource(doc, 'd').store('s'), controller = new AbortController()
    work.cells = 0; work.slice++
    setImmediate(() => controller.abort())
    const page = kind === 'rows' ? await store.readRows(HEADER, undefined, options(store, controller.signal))
      : await firstPage(store.scanPopulated({ range: STRIP, limit: 65_536 }, options(store, controller.signal)))
    expect(page.error?.code).toBe('cancelled')
    expect(work.cells).toBeLessThanOrEqual(4096 + 256)
  })
  it('returns stale-source when the owner is released mid-request, at the next checkpoint', async () => {
    const { doc, work } = objectDoc(2, 8192)
    const store = createObjectWorkbookSource(doc, 'd').store('s')
    work.cells = 0; work.slice++
    work.onTouch = () => { if (work.cells === 100) void store.release() }
    const page = await store.readRows(HEADER, undefined, options(store))
    expect(page.error?.code).toBe('stale-source')
    expect(work.cells).toBeLessThanOrEqual(256)
  })
})

/** Paged import of a dense block; its decode work is counted per decoded cell. */
async function pagedStore(rows: number, cols: number, pacing?: Pacing) {
  const letters = (col: number): string => { let out = ''; for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) out = String.fromCharCode(65 + (n - 1) % 26) + out; return out }
  let sheetData = ''
  for (let row = 0; row < rows; row++) {
    sheetData += `<row r="${row + 1}">`
    for (let col = 0; col < cols; col++) sheetData += `<c r="${letters(col)}${row + 1}"><v>${row * cols + col}</v></c>`
    sheetData += '</row>'
  }
  const bytes = buildZip(xlsxEntries({ sheetData })).bytes
  const pages = new SourcePageStore({ docId: 'p', backend: memoryPageBackend(), memoryBudgetBytes: 256 * 1024 * 1024 })
  const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes as BlobPart]) }, { docId: 'p', profile: 'D', store: pages })
  await session.done
  const source = createPagedWorkbookSource(session, pacing ? { pacing } : {})
  return { source, store: source.store('1') }
}
describe('paged store decodes a request in bounded steps', () => {
  afterEach(() => { vi.restoreAllMocks() })
  function counting() {
    const work = { cells: 0, slice: 0 }
    const proto = PagedXlsxSheetAdapter.prototype as unknown as { rowCell: (...args: unknown[]) => unknown }
    const rowCell = proto.rowCell
    vi.spyOn(proto, 'rowCell').mockImplementation(function (this: unknown, ...args: unknown[]) { work.cells++; return rowCell.apply(this, args) })
    return work
  }
  it('splits codex\'s 16,384-cell header read', async () => {
    const work = counting(), probe = frozen(work)
    const { source, store } = await pagedStore(2, 8192, probe.pacing)
    work.cells = 0; work.slice++
    const page = await store.readRows(HEADER, undefined, options(store))
    expect(cellsOf(page)).toBe(16_384)
    const slices = probe.done()
    expect(slices.length - 1).toBeGreaterThanOrEqual(3)
    expect(Math.max(...slices)).toBeLessThanOrEqual(4096)
    await source.release()
  }, 60_000)
  it('splits one 65,536-cell strip without splitting the request', async () => {
    const work = counting(), probe = frozen(work)
    const { source, store } = await pagedStore(8192, 8, probe.pacing)
    work.cells = 0; work.slice++
    const page = await firstPage(store.scanPopulated({ range: STRIP, limit: 65_536 }, options(store)))
    expect(page.items.length).toBe(65_536)
    const slices = probe.done()
    expect(slices.length - 1).toBeGreaterThanOrEqual(15)
    expect(Math.max(...slices)).toBeLessThanOrEqual(4096)
    await source.release()
  }, 60_000)
  it('steps a 65,536-cell scanOccupied strip (coordinates only, same loop)', async () => {
    const work = { cells: 0, slice: 0 }, probe = frozen(work)
    const { source, store } = await pagedStore(8192, 8, probe.pacing)
    // Coordinates only: no payload decode to count, so units are the yields themselves.
    const page = await firstPage(store.scanOccupied!({ range: STRIP, limit: 65_536 }, options(store)))
    expect(page.items.length).toBe(65_536)
    expect(probe.done().length - 1).toBeGreaterThanOrEqual(15)
    await source.release()
  }, 60_000)
  it.each(['rows', 'scan'] as const)('services a queued macrotask abort mid-request (%s)', async kind => {
    const work = counting()
    const { source, store } = kind === 'rows' ? await pagedStore(2, 8192) : await pagedStore(8192, 8)
    const controller = new AbortController()
    work.cells = 0; work.slice++
    setImmediate(() => controller.abort())
    const page = kind === 'rows' ? await store.readRows(HEADER, undefined, options(store, controller.signal))
      : await firstPage(store.scanPopulated({ range: STRIP, limit: 65_536 }, options(store, controller.signal)))
    expect(page.error?.code).toBe('cancelled')
    expect(work.cells).toBeLessThanOrEqual(4096 + 256)
    await source.release()
  }, 60_000)
})

describe('paged scan: no page escapes before its checkpoint (codex review of 6ee6282)', () => {
  afterEach(() => { vi.restoreAllMocks() })
  // A page that fills within a batch is charged and checked before it is
  // returned: an abort during its decoding returns no page.
  it.each([[100, 1], [256, 1], [4096, 3841]])('a %i-record page with an abort at record %i is cancelled, not returned', async (limit, at) => {
    const { source, store } = await pagedStore(limit, 1)
    const controller = new AbortController()
    const proto = PagedXlsxSheetAdapter.prototype as unknown as { rowCell: (...args: unknown[]) => unknown }
    const rowCell = proto.rowCell
    let calls = 0
    vi.spyOn(proto, 'rowCell').mockImplementation(function (this: unknown, ...args: unknown[]) { if (++calls === at) controller.abort(); return rowCell.apply(this, args) })
    try {
      const page = await firstPage(store.scanPopulated({ range: { r0: 0, r1: limit - 1, c0: 0, c1: 0 }, limit }, options(store, controller.signal)))
      expect(page.error?.code).toBe('cancelled')
      expect(page.items).toHaveLength(0)
      expect(calls - at).toBeLessThanOrEqual(256)
    } finally { await source.release() }
  }, 60_000)
})

describe('paged readRows counts visited records, not just decoded cells (codex D3)', () => {
  it('checks while finding the end of a wide row for one requested column, and never succeeds after a cancel', async () => {
    const { source, store } = await pagedStore(1, 8192)
    const controller = new AbortController()
    const adapter = store as unknown as { load: (...args: unknown[]) => Promise<{ rows: Uint32Array }> }
    const load = adapter.load.bind(store)
    let visited = 0
    const spy = vi.spyOn(adapter, 'load').mockImplementation(async (...args) => {
      const p = await load(...args)
      return { ...p, rows: new Proxy(p.rows, { get(target, prop) {
        if (typeof prop === 'string' && /^\d+$/.test(prop) && ++visited === 100) controller.abort()
        const value = Reflect.get(target, prop, target)
        return typeof value === 'function' ? value.bind(target) : value
      } }) }
    })
    try {
      const page = await store.readRows({ r0: 0, r1: 0, c0: 0, c1: 0 }, undefined, options(store, controller.signal))
      expect(controller.signal.aborted).toBe(true)
      expect(page.error?.code).toBe('cancelled')
      expect(visited).toBeLessThanOrEqual(356)
    } finally { spy.mockRestore(); await source.release() }
  }, 60_000)
})

describe('paged readRows final check (codex D3)', () => {
  it('a cancel after the last checkpoint (a read under 256 units) never ends in success', async () => {
    const { source, store } = await pagedStore(1, 100)
    const controller = new AbortController()
    const adapter = store as unknown as { load: (...args: unknown[]) => Promise<{ rows: Uint32Array }> }
    const load = adapter.load.bind(store)
    let visited = 0
    const spy = vi.spyOn(adapter, 'load').mockImplementation(async (...args) => {
      const p = await load(...args)
      return { ...p, rows: new Proxy(p.rows, { get(target, prop) {
        if (typeof prop === 'string' && /^\d+$/.test(prop) && ++visited === 10) controller.abort()
        const value = Reflect.get(target, prop, target)
        return typeof value === 'function' ? value.bind(target) : value
      } }) }
    })
    try {
      const page = await store.readRows({ r0: 0, r1: 0, c0: 0, c1: 0 }, undefined, options(store, controller.signal))
      expect(controller.signal.aborted).toBe(true)
      expect(page.error?.code).toBe('cancelled')
    } finally { spy.mockRestore(); await source.release() }
  }, 60_000)
})

describe('remote store: the worker serves a read in bounded steps', () => {
  it('services a queued macrotask abort while serving codex\'s header read', async () => {
    const { doc, work } = objectDoc(2, 8192)
    const source = createObjectWorkbookSource(doc, 'd')
    source.store('s') // built (and its cells indexed) before counting
    const session = new DocumentSession('d', source, { profile: 'D' })
    const owner = session.identity('s'), controller = new AbortController()
    work.cells = 0; work.slice++
    setImmediate(() => controller.abort())
    const out = await session.read(owner, { op: 'rows', range: HEADER }, controller.signal)
    expect(out).toEqual({ kind: 'cancelled' })
    expect(work.cells).toBeLessThanOrEqual(4096 + 256)
    await session.release()
  })
})
