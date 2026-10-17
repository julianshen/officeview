/**
 * Stream F item 4 (round 2): the main ledger must stay within the M budget at
 * a worker-backed sort by construction.
 * - Zero-copy adoption: the reply's band buffers (sourceAtOffset and its
 *   inverse) become the adopted view's buffers and their charge moves from the
 *   reply to the view (adoptWhileReachable), so adopting adds only the local
 *   visibility words, never a second copy of the bands. Before: wire copy +
 *   order copy + rebuilt inverse, charged as staging + view on top of the reply.
 * - Admission: before the sort command is sent, the incoming reply (plus the
 *   local visibility) is reserved against the main budget; if it does not fit,
 *   the sort is refused honestly ("not enough memory to apply the sort on this
 *   device"), the current view stays, and the ledger never exceeds the budget.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { OfficeFile } from '../src/components/OfficeFile'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import type { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { xlsxBytes } from './helpers/xlsx-package'
import { MAIN_THREAD_BUDGETS } from '../src/xlsx/view/tileCache'
import { viewReplyBytes } from '../src/worker/viewerRemoteSource'

// Worker-backed (paged) even for this small workbook.
vi.mock('../src/xlsx/storage/importSession', async original => {
  const actual = await original<typeof import('../src/xlsx/storage/importSession')>()
  return { ...actual, importLimitsFor: (profile: 'D' | 'M', overrides?: Parameters<typeof actual.importLimitsFor>[1]) => actual.importLimitsFor(profile, { ...overrides, object: { ...overrides?.object, rows: 1 } }) }
})

const RealRegistry = globalThis.FinalizationRegistry
beforeEach(() => {
  // No collection callback ever runs: lifetime charges end only when released explicitly.
  globalThis.FinalizationRegistry = class { register() {} unregister() { return false } } as unknown as typeof FinalizationRegistry
})
afterEach(() => { cleanup(); globalThis.FinalizationRegistry = RealRegistry })

const quiet = { yield: async () => {} }
const ROWS = Number(globalThis.process?.env?.ADOPT_ROWS ?? 40000)
const rows = Array.from({ length: ROWS }, (_, r) => `<row r="${r + 1}"><c r="A${r + 1}" t="inlineStr"><is><t>name ${r}</t></is></c><c r="B${r + 1}"><v>${(r * 7) % 11}</v></c></row>`).join('')
const workbook = () => new Blob([xlsxBytes({ sheetData: rows }) as BlobPart])

function workerPort() {
  const kernel = new ViewerKernel({ parse: async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)) }, quiet)
  const listeners = new Set<(event: MessageEvent) => void>()
  return {
    postMessage(message: unknown, transfer: Transferable[]) {
      const open = message as { kind: string; source?: { kind: string; blob?: Blob } }
      const blob = open.kind === 'open' && open.source?.kind === 'blob' ? open.source.blob : undefined
      const command = structuredClone(blob ? { ...open, source: { kind: 'blob' } } : message, { transfer: transfer as ArrayBuffer[] }) as WorkerCommand
      if (blob && command.kind === 'open') command.source = { kind: 'blob', blob }
      void kernel.handle(command, (out: WorkerReply, replyTransfer) => {
        const data = structuredClone(out, { transfer: replyTransfer })
        for (const listener of [...listeners]) listener({ data } as MessageEvent)
      })
    },
    addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
    removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
    terminate: vi.fn(),
  }
}

const MAX_ROWS_VISIBILITY = Math.ceil(1048576 / 32) * 4
function mount(onLedger: (ledger: MainMemoryLedger) => void) {
  const port = workerPort()
  return render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => port} {...({ onMainLedger: onLedger } as object)} />)
}
function track(ledger: MainMemoryLedger) {
  // max: the whole ledger; view: the reply and view-buffer pools (the adoption
  // itself, apart from rows and layout the re-render then loads).
  const state = { max: 0, view: 0 }
  const set = ledger.set.bind(ledger)
  ledger.set = ((...args: Parameters<MainMemoryLedger['set']>) => {
    set(...args)
    const pools = ledger.snapshot().byPool
    state.max = Math.max(state.max, ledger.bytes)
    state.view = Math.max(state.view, pools.replies + pools.viewBuffers)
  }) as MainMemoryLedger['set']
  return state
}

describe('worker-backed sort: bounded main-ledger adoption', () => {
  it('adopting the sorted view adds at most the reply payload plus the local visibility (no band copies)', async () => {
    let ledger: MainMemoryLedger | null = null, state = { max: 0, view: 0 }
    const ui = mount(l => { if (ledger !== l) { ledger = l; state = track(l) } })
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /value name 1\b/ })).toBeTruthy(), { timeout: 60_000 })
    const header = () => ui.getByRole('columnheader', { name: /^Column B/ })
    await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(new RegExp(`complete body ${ROWS}`)), { timeout: 60_000 })
    const pools = ledger!.snapshot().byPool, before = pools.replies + pools.viewBuffers
    state.view = before
    fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' }))
    await waitFor(() => expect(header().getAttribute('aria-sort')).toBe('ascending'), { timeout: 60_000 })
    // The reservation admitted before the command (one band over the body rows:
    // sourceAtOffset + displayOffsetOfSource adopted in place, the wire and local
    // visibility words) bounds the whole adoption; it is far below reply + copies.
    const bound = viewReplyBytes(ROWS) + 16 * 1024
    expect(viewReplyBytes(ROWS)).toBeLessThanOrEqual(2 * 4 * ROWS + 2 * MAX_ROWS_VISIBILITY + 80 * 1024)
    expect(state.view - before, `adoption peak growth ${state.view - before} B over ${before} B`).toBeLessThanOrEqual(bound)
    ui.unmount()
  }, 120_000)

  it('refuses a sort whose reply would not fit the main budget, keeps the view, and never exceeds the budget', async () => {
    let ledger: MainMemoryLedger | null = null, state = { max: 0, view: 0 }
    const ui = mount(l => { if (ledger !== l) { ledger = l; state = track(l) } })
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /value name 1\b/ })).toBeTruthy(), { timeout: 60_000 })
    await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(new RegExp(`complete body ${ROWS}`)), { timeout: 60_000 })
    const budget = MAIN_THREAD_BUDGETS.D
    // Something else already holds nearly all of the main budget.
    ledger!.set('document', 'test-pressure', budget - ledger!.bytes - 64 * 1024)
    state.max = ledger!.bytes
    fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' }))
    await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(/not enough memory to apply the sort on this device/i), { timeout: 60_000 })
    expect(ui.getByRole('columnheader', { name: /^Column B/ }).getAttribute('aria-sort')).toBeNull()
    expect(state.max).toBeLessThanOrEqual(budget)
    // With the pressure gone the same sort applies (main and worker stayed in step).
    ledger!.release('document', 'test-pressure')
    fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' }))
    await waitFor(() => expect(ui.getByRole('columnheader', { name: /^Column B/ }).getAttribute('aria-sort')).toBe('ascending'), { timeout: 60_000 })
    ui.unmount()
  }, 120_000)
})
