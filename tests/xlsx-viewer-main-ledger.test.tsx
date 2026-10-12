/**
 * S10.2 / spec §12.4: the document-scoped main-thread application ledger.
 * Every main-owned source charges and releases; close and replacement return
 * to zero; the high-water mark is at least any concurrent sum (a transferred
 * reply and the cache decoded from it both count while both are live).
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { OfficeFile } from '../src/components/OfficeFile'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import { DOM_ELEMENT_BYTES, MainMemoryLedger, estimatePayloadBytes } from '../src/xlsx/view/mainLedger'
import { MergeQueryCache } from '../src/xlsx/view/viewport'
import { ValueIndexCache } from '../src/xlsx/view/valueIndex'
import { workbookOf } from './helpers/xlsx-find'
import { xlsxBytes } from './helpers/xlsx-package'

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

describe('worker-backed document ledger', () => {
  it('charges row cache, replies (released after their task) and adopted view buffers; close returns to zero', async () => {
    const h = workerHarness()
    let ledger: MainMemoryLedger | null = null
    const internal = { onMainLedger: (l: MainMemoryLedger) => { ledger = l } } as object
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} {...internal} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /value name 1\b/ })).toBeTruthy())
    await tick()
    const loaded = ledger!.snapshot()
    expect(loaded.byPool.rowCache).toBeGreaterThan(0)
    expect(loaded.byPool.dom).toBeGreaterThan(0)
    expect(loaded.byPool.dom % DOM_ELEMENT_BYTES).toBe(0)
    expect(loaded.byPool.replies).toBe(0)          // every reply released after its task
    expect(loaded.peakBytes).toBeGreaterThan(loaded.bytes - loaded.byPool.dom) // replies + decoded rows overlapped
    fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' }))
    await waitFor(() => expect(ui.getByRole('columnheader', { name: /Column B; sort priority 1, ascending/ })).toBeTruthy())
    await tick()
    const sorted = ledger!.snapshot()
    expect(sorted.byPool.viewBuffers).toBeGreaterThan(0)
    expect(sorted.byPool.layout).toBeGreaterThanOrEqual(0)
    const peak = sorted.peakBytes
    expect(peak).toBeGreaterThanOrEqual(Object.values(sorted.byPool).reduce((a, b) => a + b, 0))
    ui.unmount()
    await waitFor(() => expect(ledger!.snapshot().bytes).toBe(0))
    expect(ledger!.snapshot().peakBytes).toBeGreaterThanOrEqual(peak) // the high-water mark survives close
  })
})

describe('resident document ledger', () => {
  const doc = () => workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'Name' }, { row: 1, col: 0, value: 'beta' }, { row: 2, col: 0, value: 'alpha' }] }])
  it('charges the owned DOM and layout caches, and a copy session only while it runs; unmount and replacement return to zero', async () => {
    Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: vi.fn(async () => {}) } })
    const ledgers: MainMemoryLedger[] = []
    const copies: number[] = []
    const ui = render(<InteractiveSheetDoc doc={doc()} {...({ onMainLedger: (l: MainMemoryLedger) => { ledgers.push(l); const set = l.set.bind(l); l.set = (pool, owner, bytes) => { if (pool === 'copy') copies.push(bytes); set(pool, owner, bytes) } } } as object)} />)
    await waitFor(() => expect(ledgers.at(-1)?.snapshot().byPool.dom ?? 0).toBeGreaterThan(0))
    fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: 'A2' } })
    fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
    fireEvent.click(ui.getByRole('button', { name: 'Copy selection' }))
    await waitFor(() => expect(copies.at(-1)).toBe(0))
    expect(Math.max(...copies)).toBeGreaterThan(0)
    const first = ledgers.at(-1)!
    ui.rerender(<InteractiveSheetDoc doc={doc()} {...({ onMainLedger: (l: MainMemoryLedger) => { ledgers.push(l) } } as object)} />)
    await waitFor(() => expect(ledgers.at(-1)).not.toBe(first))
    expect(first.snapshot().bytes).toBe(0)            // the replaced document's ledger is released
    ui.unmount()
    expect(ledgers.at(-1)!.snapshot().bytes).toBe(0)
  })
})
