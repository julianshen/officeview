/**
 * Release benchmark regression (F2 header sort, 1,048,576 rows): every
 * worker-backed sort adopts a new view whose order bands (~8 MiB at F2) were
 * charged to the main ledger for as long as they were reachable, i.e. until a
 * garbage collection ran their FinalizationRegistry callbacks. Superseded views
 * nothing held any more stayed charged until then, the ledger climbed past the
 * 64 MiB desktop budget after ~9 sorts, and the next sort's 256 B control
 * admission was refused ("Main document memory quota exceeded").
 *
 * A superseded view's charge must end deterministically once its last owner
 * (the workbook, the committed display, a copy) releases it, never on
 * collection timing. This file disables collection callbacks entirely
 * (FinalizationRegistry never fires), so only explicit releases can keep the
 * pool bounded.
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
const rows = Array.from({ length: 40 }, (_, r) => `<row r="${r + 1}"><c r="A${r + 1}" t="inlineStr"><is><t>name ${r}</t></is></c><c r="B${r + 1}"><v>${(r * 7) % 11}</v></c></row>`).join('')
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

describe('worker-backed sort views: superseded view buffers are released without garbage collection', () => {
  it('repeated header sorts keep the view-buffer pool at the live views', async () => {
    let ledger: MainMemoryLedger | null = null
    const port = workerPort()
    const internal = { onMainLedger: (l: MainMemoryLedger) => { ledger = l } } as object
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => port} {...internal} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /value name 1\b/ })).toBeTruthy())
    const header = () => ui.getByRole('columnheader', { name: /^Column B/ })
    const cycle = async () => {
      const before = header().getAttribute('aria-sort')
      fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' }))
      await waitFor(() => expect(header().getAttribute('aria-sort')).not.toBe(before))
    }
    // ascending, descending, clear: one full cycle, then three more.
    const pools: number[] = []
    for (let i = 0; i < 12; i++) { await cycle(); pools.push(ledger!.snapshot().byPool.viewBuffers) }
    expect(pools[0]).toBeGreaterThan(0)
    // Only live views are charged, so the pool is the same at the same point of
    // every cycle. On 485a57c every superseded view stayed charged and the pool
    // grew by one view (~135 KiB here, ~8 MiB at F2) per sort.
    for (let i = 3; i < pools.length; i++) expect(pools[i], `sort ${i + 1}: ${pools.join(' ')}`).toBeLessThanOrEqual(pools[i % 3])
    ui.unmount()
  }, 30_000)
})
