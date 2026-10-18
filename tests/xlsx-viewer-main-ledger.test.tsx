/**
 * S10.2 / spec §12.4: the document-scoped main-thread application ledger.
 * Every main-owned source charges and releases; close and replacement return
 * to zero; the high-water mark is at least any concurrent sum (a transferred
 * reply and the cache decoded from it both count while both are live).
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { useEffect, useState, type JSX } from 'react'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { OfficeFile } from '../src/components/OfficeFile'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import { MainMemoryLedger, SHAPE_ENTRY_BYTES, estimatePayloadBytes, leaseTargets, shapeCacheBytes } from '../src/xlsx/view/mainLedger'
import { openPagedWorkbook } from '../src/components/xlsx/pagedWorkbook'
import { pagedWorkbookOf } from '../src/components/pagedDocument'
import { pagedSheetGeometry } from '../src/components/xlsx/pagedGeometry'
import type { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import v8 from 'node:v8'
import vm from 'node:vm'

// Real collection: lifetime charges end only when V8 has collected every target.
v8.setFlagsFromString('--expose_gc')
const gc = vm.runInNewContext('gc') as () => void
async function collect(): Promise<void> {
  for (let i = 0; i < 6; i++) { gc(); await new Promise(resolve => setTimeout(resolve, 10)) }
}
/** React's DEVELOPMENT build keeps owner stacks (_debugStack) on fibers, and
 * the last event's target fiber (return_targetInst) and the last root with
 * passive updates stay referenced until the next event/commit; through the
 * owner stack's closures they reach the unmounted viewer (heap snapshots).
 * The production build detaches those fibers. These dev-build tests flush that
 * dev-only retention; the production behaviour (collection with no further
 * render or event) is tested in xlsx-viewer-unmount-gc.test.tsx. */
function PassiveUpdate(): JSX.Element {
  const [n, setN] = useState(0)
  useEffect(() => { if (n === 0) setN(1) }, [n])
  return <div>{n}</div>
}
async function flushReactRetention(): Promise<void> {
  cleanup()
  // A passive effect that schedules an update moves React's module variable to this root.
  const blank = render(<PassiveUpdate />)
  await tick()
  // React's event system also keeps the last event's target fiber
  // (return_targetInst): one event in the blank root replaces it.
  fireEvent.click(blank.container.firstChild as Element)
  blank.unmount()
  cleanup()
}
import { MergeQueryCache } from '../src/xlsx/view/viewport'
import { ValueIndexCache, valueSourceFromStore } from '../src/xlsx/view/valueIndex'
import { CELL_INDEX_ROW_BYTES, CELL_INDEX_SLOT_BYTES, buildSheetGeometry } from '../src/xlsx/view/viewport'
import { RectangleIndex } from '../src/xlsx/view/structure'
import { domClassBytes, domClassCounts } from '../src/components/xlsx/domCost'
import { DOM_CLASS_BYTES } from '../src/xlsx/view/mainLedger'
import type { MainPool } from '../src/xlsx/view/mainLedger'
import { workbookOf } from './helpers/xlsx-find'
import { xlsxBytes } from './helpers/xlsx-package'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'

vi.mock('../src/xlsx/storage/importSession', async original => {
  const actual = await original<typeof import('../src/xlsx/storage/importSession')>()
  return { ...actual, importLimitsFor: (profile: 'D' | 'M', overrides?: Parameters<typeof actual.importLimitsFor>[1]) => actual.importLimitsFor(profile, { ...overrides, object: { ...overrides?.object, rows: 1 } }) }
})
afterEach(() => { cleanup(); vi.restoreAllMocks() })

const quiet = { yield: async () => {} }
const rows = Array.from({ length: 40 }, (_, r) => `<row r="${r + 1}"><c r="A${r + 1}" t="inlineStr"><is><t>name ${r}</t></is></c><c r="B${r + 1}"><v>${(r * 7) % 11}</v></c></row>`).join('')
const workbook = () => new Blob([xlsxBytes({ sheetData: rows }) as BlobPart])
const tick = () => act(async () => { await new Promise(resolve => setTimeout(resolve, 5)) })

function workerHarness() {
  const parse = vi.fn(async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)))
  const kernel = new ViewerKernel({ parse }, quiet)
  const listeners = new Set<(event: MessageEvent) => void>()
  const posted: WorkerCommand[] = []
  const held: Array<() => void> = []
  let hold: ((command: WorkerCommand) => boolean) | null = null
  const port = {
    postMessage(message: unknown, transfer: Transferable[]) {
      // A Blob crosses by handle (jsdom's structuredClone cannot clone it).
      const open = message as { kind: string; source?: { kind: string; blob?: Blob } }
      const blob = open.kind === 'open' && open.source?.kind === 'blob' ? open.source.blob : undefined
      const command = structuredClone(blob ? { ...open, source: { kind: 'blob' } } : message, { transfer: transfer as ArrayBuffer[] }) as WorkerCommand
      if (blob && command.kind === 'open') command.source = { kind: 'blob', blob }
      posted.push(command)
      void kernel.handle(command, (out: WorkerReply, replyTransfer) => {
        const deliver = () => {
          const data = structuredClone(out, { transfer: replyTransfer })
          for (const listener of [...listeners]) listener({ data } as MessageEvent)
        }
        if (hold?.(command)) held.push(deliver)
        else deliver()
      })
    },
    addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
    removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
    terminate: vi.fn(),
  }
  return {
    port, kernel, parse, posted,
    reads: (op?: string) => posted.filter(c => c.kind === 'read' && (!op || c.query.op === op)).length,
    hold(predicate: (command: WorkerCommand) => boolean) { hold = predicate },
    release() { hold = null; for (const deliver of held.splice(0)) deliver() },
  }
}

describe('MainMemoryLedger', () => {
  it('charges absolutely per owner, releases, and keeps a high-water mark with its pool split', () => {
    const ledger = new MainMemoryLedger()
    ledger.set('replies', 'r1', 1000)
    ledger.set('rowCache', 'rows', 700)          // the decoded copy, while the reply is still live
    expect(ledger.snapshot()).toMatchObject({ bytes: 1700, peakBytes: 1700 })
    ledger.release('replies', 'r1')
    ledger.set('rowCache', 'rows', 900)          // re-measure replaces, never adds
    expect(ledger.snapshot()).toMatchObject({ bytes: 900, peakBytes: 1700, atPeak: { replies: 1000, rowCache: 700 } })
    ledger.set('dom', 'grid', 50); ledger.set('copy', 'session', 10)
    ledger.releaseAll('dom')
    expect(ledger.snapshot().byPool).toMatchObject({ dom: 0, copy: 10, rowCache: 900 })
    ledger.releaseAll()
    expect(ledger.snapshot()).toMatchObject({ bytes: 0, peakBytes: 1700 })
    expect(() => ledger.set('layout', 'x', -1)).toThrow(RangeError)
  })
  it('estimates payloads: each buffer once, strings by code units', () => {
    const buffer = new Uint32Array(1000)
    const bytes = estimatePayloadBytes({ a: buffer, b: buffer.subarray(10), s: 'x'.repeat(100) })
    expect(bytes).toBeGreaterThanOrEqual(4000 + 200)
    expect(bytes).toBeLessThan(4000 + 200 + 400)
  })
  it('the merge query cache accounts its retained pages and evictions', () => {
    const cache = new MergeQueryCache(2)
    const geometry = { merges: { query: () => ({ items: [{}, {}], complete: true }) } } as never
    const work = { mergeQueries: 0 } as never
    cache.query(geometry, { r0: 0, r1: 1, c0: 0, c1: 1 }, 10, work)
    const one = cache.bytes
    expect(one).toBeGreaterThan(0)
    cache.query(geometry, { r0: 2, r1: 3, c0: 0, c1: 1 }, 10, work)
    cache.query(geometry, { r0: 4, r1: 5, c0: 0, c1: 1 }, 10, work)
    expect(cache.bytes).toBe(one * 2) // the oldest page was evicted
  })
})

describe('controls value cache', () => {
  it('reports its used bytes on every charge and release (allocation-time ledger charge)', () => {
    const seen: number[] = []
    const cache = new ValueIndexCache({ limitBytes: 1024, onUsage: bytes => seen.push(bytes) })
    const reservation = cache.reserve()
    reservation.charge(100)
    reservation.charge(50)
    reservation.release()
    expect(seen).toEqual([100, 150, 0])
    expect(cache.usedBytes).toBe(0)
  })
})

describe('lifetime charges (review round 3, findings 1-3)', () => {
  it('a charge lasts while any target is reachable and ends only after collection, never on a timer', async () => {
    const ledger = new MainMemoryLedger()
    const holder: { reply: unknown; page: unknown } = { reply: { result: { result: { page: { items: [{ row: 1 }] } } } }, page: null }
    ledger.chargeWhileReachable('replies', leaseTargets(holder.reply), 2_000_000)
    // The consumer keeps only the nested page (a scan generator across tasks).
    holder.page = (holder.reply as { result: { result: { page: unknown } } }).result.result.page
    holder.reply = null
    await new Promise(resolve => setTimeout(resolve, 20))
    await collect()
    expect(ledger.snapshot().byPool.replies).toBe(2_000_000)
    holder.page = null
    await collect()
    expect(ledger.snapshot().byPool.replies).toBe(0)
    expect(ledger.snapshot().peakBytes).toBe(2_000_000)
  })
})

describe('review round 4 (opencode): honest structure charges', () => {
  it('shaping cache estimate: per entry plus a hash table of at least 64 buckets, grown at 3/4 load', () => {
    expect(shapeCacheBytes(0, 0)).toBe(0)
    expect(shapeCacheBytes(1, 5)).toBe(1360 + 40 + 64 * 218)
    expect(shapeCacheBytes(48, 0)).toBe(48 * 1360 + 64 * 218)
    expect(shapeCacheBytes(49, 0)).toBe(49 * 1360 + 128 * 218)
    // The F2 snapshot: 111 caches of ~80 entries (~8 units each) -> within 5 % of 14.9 MiB.
    const f2 = 111 * shapeCacheBytes(80, 80 * 8) / 1048576
    expect(f2).toBeGreaterThan(14.9 * 0.95)
    expect(f2).toBeLessThan(14.9 * 1.10)
  })
  it('DOM elements are counted by class: inside gridcells, inside row/column headers, and the rest', () => {
    const root = document.createElement('div')
    root.innerHTML = '<div role="grid"><div role="row"><div role="columnheader"><button><span>A</span></button><button>f</button></div>'
      + '<div role="rowheader"><button>1</button></div><div role="gridcell"><span>1</span></div><div role="gridcell"><span>2</span></div></div></div><p role="status">x</p>'
    const counts = domClassCounts(root)
    expect(counts).toEqual({ cell: 4, header: 6, other: 3 })
    expect(domClassBytes(counts)).toBe(4 * DOM_CLASS_BYTES.cell + 6 * DOM_CLASS_BYTES.header + 3 * DOM_CLASS_BYTES.other)
    expect(domClassCounts(null)).toEqual({ cell: 0, header: 0, other: 0 })
  })
  it('a lifetime charge needs at least one target (nothing could ever end it)', () => {
    expect(() => new MainMemoryLedger().chargeWhileReachable('replies', [], 10)).toThrow(RangeError)
  })
  it('resident geometry charges its per-row cell index, not 200 B per authored row', () => {
    const rows = Array.from({ length: 1000 }, (_, r) => ({ index: r * 3, cells: [{ col: 0, value: r }, { col: 2, value: 'x' }] }))
    const sheet = { name: 'S', rows, cols: [], merges: [], mergeRanges: [] } as never
    const built = buildSheetGeometry(sheet)
    if (built.kind !== 'ready') throw new Error('geometry')
    expect(built.cellIndexBytes).toBe(1000 * CELL_INDEX_ROW_BYTES + 2000 * CELL_INDEX_SLOT_BYTES)
  })
  it('the merge index charges ~150 B measured per entry (160 B), plus its interval tree', () => {
    const index = RectangleIndex.build(Array.from({ length: 100 }, (_, i) => ({ rect: { r0: i * 2, r1: i * 2 + 1, c0: 0, c1: 1 }, value: i })))
    expect(index.estimatedBytes).toBeGreaterThanOrEqual(100 * 160)
  })
})

describe('reply charges settle when the payload is adopted (round-4 follow-up 1)', () => {
  it('settle ends a lifetime charge at once; a later collection is a no-op; unknown targets do nothing', async () => {
    const ledger = new MainMemoryLedger()
    let reply: { result: { page: { items: number[] } } } | null = { result: { page: { items: [1, 2, 3] } } }
    ledger.chargeWhileReachable('replies', leaseTargets(reply), 5000)
    ledger.chargeWhileReachable('replies', [{}], 7)
    ledger.set('rowCache', 'rows', 4000)          // the owner charges first...
    ledger.settle(reply.result.page)              // ...then the payload's charge ends
    expect(ledger.snapshot().byPool.replies).toBe(7)
    ledger.settle({})                             // not a registered target
    ledger.settle(reply.result.page)              // settling twice is harmless
    reply = null
    await collect()
    expect(ledger.snapshot().byPool).toMatchObject({ replies: 0, rowCache: 4000 })
    expect(ledger.snapshot().peakBytes).toBe(9007)
  })
})

/** Records every ledger change with the pool totals right after it. */
function timeline(ledger: MainMemoryLedger): Array<{ pool: MainPool; delta: number; byPool: Record<MainPool, number> }> {
  const events: Array<{ pool: MainPool; delta: number; byPool: Record<MainPool, number> }> = []
  const set = ledger.set.bind(ledger)
  ledger.set = (pool, owner, bytes) => {
    const before = ledger.snapshot().byPool[pool]
    set(pool, owner, bytes)
    const after = ledger.snapshot().byPool
    if (after[pool] !== before) events.push({ pool, delta: after[pool] - before, byPool: after })
  }
  return events
}

/** Opens a paged workbook over the in-process kernel and returns its RemoteWorkbook. */
async function openRemote(): Promise<{ remote: RemoteWorkbook; release: () => Promise<void> }> {
  const h = workerHarness()
  const doc = await openPagedWorkbook(workbook(), { port: h.port, profile: 'D', signal: new AbortController().signal })
  const handle = pagedWorkbookOf(doc)!
  return { remote: handle.source as RemoteWorkbook, release: () => handle.release() }
}

describe('worker-backed document ledger', () => {
  it('charges geometry and projection buffers for their lifetime; a closed workbook drops them even while it is still held', async () => {
    const { remote, release } = await openRemote()
    const ledger = remote.mainLedger
    let geometry: unknown = pagedSheetGeometry(remote, '1', createIdentityOrder(), visibilityFromHidden(createIntervalSet([])))
    expect(ledger.snapshot().byPool.geometry).toBeGreaterThan(0)
    expect(remote.retainedResources().geometries).toBe(1)
    geometry = null
    await release()
    // Still holding `remote` (a host that keeps a closed document): it must hold nothing.
    expect(remote.retainedResources()).toEqual({ stores: 0, views: 0, geometries: 0, layouts: 0, structures: 0 })
    await collect()
    expect(ledger.snapshot().bytes).toBe(0)
    void geometry
  })

  it('a scan page kept by its consumer keeps its reply charged across tasks', async () => {
    const { remote, release } = await openRemote()
    const ledger = remote.mainLedger
    await collect()
    const before = ledger.snapshot().byPool.replies
    let iterator: AsyncIterator<{ readonly items: readonly unknown[] }> | null = remote.store('1').scanPopulated({ range: { r0: 0, r1: 39, c0: 0, c1: 1 }, limit: 64 }, { signal: new AbortController().signal, stamp: { ...remote.store('1').stamp } })[Symbol.asyncIterator]()
    let page: { readonly items: readonly unknown[] } | null = (await iterator!.next()).value
    expect(page!.items.length).toBeGreaterThan(0)
    await new Promise(resolve => setTimeout(resolve, 20))
    await collect()
    expect(ledger.snapshot().byPool.replies).toBeGreaterThan(before)
    await iterator!.return?.()
    page = null; iterator = null
    await collect()
    expect(ledger.snapshot().byPool.replies).toBe(before)
    await release()
  })

  it('in the viewer: row cache, replies, view buffers, DOM; unmount releases the viewer pools, collection the rest', FIRST_PAINT_TEST, async () => {
    // The render result lives only inside this function: a test frame holding
    // it would itself keep the tree (and its buffers) reachable.
    const ledger = await (async () => {
      const h = workerHarness()
      let found: MainMemoryLedger | null = null
      const internal = { onMainLedger: (l: MainMemoryLedger) => { found = l } } as object
      const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} {...internal} />)
      await waitFor(() => expect(ui.getByRole('gridcell', { name: /value name 1\b/ })).toBeTruthy(), FIRST_PAINT_WAIT)
      await tick()
      const loaded = found!.snapshot()
      expect(loaded.byPool.rowCache).toBeGreaterThan(0)
      expect(loaded.byPool.geometry).toBeGreaterThan(0)
      expect(loaded.byPool.dom).toBeGreaterThan(0)
      // Charged by class over everything the viewer renders.
      const root = ui.container.querySelector('[data-officeview-sheet-screen]')!.parentElement!
      expect(loaded.byPool.dom).toBe(domClassBytes(domClassCounts(root)))
      // Retained tiles drew text: the browser's per-canvas shaping cache is charged.
      expect(loaded.byPool.canvasText).toBeGreaterThanOrEqual(SHAPE_ENTRY_BYTES)
      fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' }))
      await waitFor(() => expect(ui.getByRole('columnheader', { name: /Column B; sort priority 1, ascending/ })).toBeTruthy())
      await tick()
      const sorted = found!.snapshot()
      expect(sorted.byPool.viewBuffers).toBeGreaterThan(0)
      expect(sorted.peakBytes).toBeGreaterThanOrEqual(Object.values(sorted.byPool).reduce((a, b) => a + b, 0))
      ui.unmount()
      await waitFor(() => expect(found!.snapshot().byPool.dom).toBe(0))
      expect(found!.snapshot().byPool.canvasText).toBe(0)
      return found!
    })()
    await flushReactRetention()
    await collect()
    expect(ledger.snapshot().bytes).toBe(0)
  })
})

describe('worker-backed adoption settles replies without waiting for GC (round-4 follow-up 1)', () => {
  it('value list (F4-like): each value reply is settled after the index charged its page; the peak stays near the retained index', async () => {
    const { remote, release } = await openRemote()
    const ledger = remote.mainLedger
    const store = remote.store('1')
    await collect()
    const base = ledger.snapshot()
    const events = timeline(ledger)
    const cache = new ValueIndexCache({ limitBytes: 16 * 1024 * 1024, onUsage: bytes => ledger.set('controls', 'cache', bytes) })
    const identity = { ...store.stamp, range: { r0: 0, r1: 39, c0: 0, c1: 1, headerRows: 0 as const, totalsRows: 0 as const }, col: 0, classification: { kind: 'text' as const, hasNumeric: false }, classificationGen: 0 }
    const index = cache.open(identity, valueSourceFromStore(store, 0, { r0: 0, r1: 39 }, 8))
    expect(await index.build()).toBe('complete')
    const replies = events.filter(e => e.pool === 'replies')
    const charged = replies.filter(e => e.delta > 0), settled = replies.filter(e => e.delta < 0)
    expect(charged.length).toBe(5)                         // one valueQuery reply per 8-row page
    expect(settled.length).toBe(5)                         // every one settled, no GC involved
    // No window where neither is charged: each settle happens with the index's charge in place.
    for (const e of settled) expect(e.byPool.controls).toBeGreaterThan(0)
    // Without GC, the reply pool is back at its baseline (or lower, if old replies were collected).
    expect(ledger.snapshot().byPool.replies).toBeLessThanOrEqual(base.byPool.replies)
    // At most one page's reply is ever charged next to the index: the peak stays
    // within one reply of the retained index instead of growing with every page.
    const maxReply = Math.max(...charged.map(e => e.delta))
    const peak = Math.max(...events.map(e => e.byPool.controls + e.byPool.replies - base.byPool.replies))
    expect(peak).toBeLessThanOrEqual(ledger.snapshot().byPool.controls + maxReply)
    await release()
  })

  it('row pages: the row cache charges its copies, then the page reply is settled', async () => {
    const { remote, release } = await openRemote()
    const ledger = remote.mainLedger
    await collect()
    const base = ledger.snapshot().byPool.replies
    const events = timeline(ledger)
    await remote.store('1').load(Array.from({ length: 40 }, (_, r) => r), 0, 1)
    const settled = events.filter(e => e.pool === 'replies' && e.delta < 0)
    expect(settled.length).toBeGreaterThan(0)
    for (const e of settled) expect(e.byPool.rowCache).toBeGreaterThan(0)
    expect(ledger.snapshot().byPool.replies).toBeLessThanOrEqual(base)
    await release()
  })

  it('a committed sort view: the adopted view charges first, then its reply is settled', async () => {
    const { remote, release } = await openRemote()
    const ledger = remote.mainLedger
    await collect()
    const base = ledger.snapshot().byPool.replies
    const events = timeline(ledger)
    const spec = { range: { r0: 0, r1: 39, c0: 0, c1: 1, headerRows: 0 as const, totalsRows: 0 as const }, keys: [{ col: 1, order: 'asc' as const }], caseSensitive: false }
    const commit = await remote.sort(remote.identity('1'), spec, new AbortController().signal)
    expect(commit.kind).toBe('committed')
    const settled = events.filter(e => e.pool === 'replies' && e.delta < 0)
    expect(settled.length).toBeGreaterThan(0)
    for (const e of settled) expect(e.byPool.viewBuffers).toBeGreaterThan(0)
    // prepareColumns' reply carries nothing worth keeping; the view reply is settled.
    expect(ledger.snapshot().byPool.replies - base).toBeLessThan(2048)
    await release()
  })
})

describe('DOM a child commits on its own is charged when it lands (codex round 5, finding 1)', () => {
  it('the filter popup\'s value list, rendered by the popup\'s own state, is in the dom pool while idle', async () => {
    const cells = [{ row: 0, col: 0, value: 'Name' }, ...Array.from({ length: 310 }, (_, i) => ({ row: i + 1, col: 0, value: `value ${String(i).padStart(3, '0')}` }))]
    const ledger = await (async () => {
      let found: MainMemoryLedger | null = null
      const ui = render(<InteractiveSheetDoc doc={workbookOf([{ sheetId: 'A', cells }])} {...({ onMainLedger: (l: MainMemoryLedger) => { found = l } } as object)} />)
      await waitFor(() => expect(found?.snapshot().byPool.dom ?? 0).toBeGreaterThan(0))
      fireEvent.click(ui.getByRole('button', { name: /^Filter column A/ }))
      await waitFor(() => expect(ui.container.querySelectorAll('[data-value-ordinal]').length).toBeGreaterThanOrEqual(200), { timeout: 20_000 })
      await new Promise(resolve => setTimeout(resolve, 250))
      const root = ui.container.querySelector('[data-officeview-sheet-screen]')!.parentElement!
      // No parent render since the values landed: the charge still matches the rendered tree.
      expect(found!.snapshot().byPool.dom).toBe(domClassBytes(domClassCounts(root)))
      expect(domClassCounts(root).other).toBeGreaterThan(600)
      ui.unmount()
      return found!
    })()
    expect(ledger.snapshot().byPool.dom).toBe(0)
  }, 30_000)
})

describe('resident document ledger', () => {
  const doc = () => workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'Name' }, { row: 1, col: 0, value: 'beta' }, { row: 2, col: 0, value: 'alpha' }] }])
  it('charges the document, geometry, DOM and a copy session while it runs; unmount and replacement end the charges', async () => {
    Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: vi.fn(async () => {}) } })
    const ledgers = await (async () => {
      const found: MainMemoryLedger[] = []
      const copies: number[] = []
      const ui = render(<InteractiveSheetDoc doc={doc()} {...({ onMainLedger: (l: MainMemoryLedger) => { found.push(l); const set = l.set.bind(l); l.set = (pool, owner, bytes) => { if (pool === 'copy') copies.push(bytes); set(pool, owner, bytes) } } } as object)} />)
      await waitFor(() => expect(found.at(-1)?.snapshot().byPool.dom ?? 0).toBeGreaterThan(0))
      expect(found.at(-1)!.snapshot().byPool.document).toBeGreaterThan(0)
      expect(found.at(-1)!.snapshot().byPool.geometry).toBeGreaterThan(0)
      // A resident sort publishes main-built order buffers: charged while the view holds them.
      fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column A' }))
      await waitFor(() => expect(found.at(-1)!.snapshot().byPool.viewBuffers).toBeGreaterThan(0))
      fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: 'A2' } })
      fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
      fireEvent.click(ui.getByRole('button', { name: 'Copy selection' }))
      await waitFor(() => expect(copies.at(-1)).toBe(0))
      expect(Math.max(...copies)).toBeGreaterThan(0)
      const first = found.at(-1)!
      ui.rerender(<InteractiveSheetDoc doc={doc()} {...({ onMainLedger: (l: MainMemoryLedger) => { found.push(l) } } as object)} />)
      await waitFor(() => expect(found.at(-1)).not.toBe(first))
      expect(first.snapshot().byPool.dom).toBe(0)
      ui.unmount()
      return found
    })()
    await flushReactRetention()
    await collect()
    for (const ledger of ledgers) expect(ledger.snapshot().bytes).toBe(0)
  })
})

it('atomically adopts copy reservations into document reachability without another peak or refund gap', () => {
  const registrations: { target: object; token: string }[] = []
  const ledger = new MainMemoryLedger({ register: (target, token) => registrations.push({ target, token }) })
  const sheet = {}, diagnostic = {}
  ledger.set('copy', 'job', 1000)
  let atCommit = 0
  const token = ledger.adoptWhileReachable('copy', 'job', 'document', [sheet, diagnostic], 300, () => { atCommit = ledger.bytes })
  expect(ledger.snapshot().byPool).toMatchObject({ copy: 700, document: 300 })
  expect(ledger.bytes).toBe(1000)
  expect(ledger.peakBytes).toBe(1000)
  expect(atCommit).toBe(1000)
  ledger.release('copy', 'job')
  expect(ledger.bytes).toBe(300)
  expect(registrations.map(entry => entry.token)).toEqual([token, token])
  ledger.collected(token)
  expect(ledger.bytes).toBe(300)
  ledger.collected(token)
  expect(ledger.bytes).toBe(0)
})

it('leaves reserved ownership untouched if lifetime registration or publication throws', () => {
  const ledger = new MainMemoryLedger({ register: () => {} })
  ledger.set('copy', 'job', 1000)
  expect(() => ledger.adoptWhileReachable('copy', 'job', 'document', [{}], 300, () => { throw new Error('publish failed') })).toThrow('publish failed')
  expect(ledger.snapshot().byPool).toMatchObject({ copy: 1000, document: 0 })
  expect(ledger.peakBytes).toBe(1000)
  expect(() => ledger.adoptWhileReachable('copy', 'job', 'document', [{}], 1001)).toThrow(RangeError)
  expect(ledger.bytes).toBe(1000)
  const throwing = new MainMemoryLedger({ register: () => { throw new Error('registration failed') } })
  throwing.set('copy', 'job', 1000)
  expect(() => throwing.adoptWhileReachable('copy', 'job', 'document', [{}], 300)).toThrow('registration failed')
  expect(throwing.snapshot().byPool).toMatchObject({ copy: 1000, document: 0 })
})

it('real collection ends an adopted diagnostic only after both document and detached diagnostic stop reaching it', async () => {
  const ledger = new MainMemoryLedger()
  const owners: { sheet?: object; diagnostic?: object } = {}
  const make = () => {
    const diagnostic = { message: 'owned canonical diagnostic' }, sheet = { diagnostics: [diagnostic] }
    owners.sheet = sheet; owners.diagnostic = diagnostic
    ledger.set('copy', 'job', 400)
    ledger.adoptWhileReachable('copy', 'job', 'document', [sheet, diagnostic], 300)
    ledger.release('copy', 'job')
  }
  make()
  await collect()
  expect(ledger.snapshot().byPool.document).toBe(300)
  delete owners.sheet
  await collect()
  expect(ledger.snapshot().byPool.document).toBe(300)
  delete owners.diagnostic
  await collect()
  expect(ledger.snapshot().byPool.document).toBe(0)
})
