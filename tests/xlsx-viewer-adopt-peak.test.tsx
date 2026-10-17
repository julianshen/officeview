/**
 * Stream F item 4 (iOS Simulator / Chromium M-profile run): the main ledger
 * peaked over the 32 MiB M budget at the F2 header sort (viewBuffers 12.9 MB +
 * replies 8.4 MB at the peak). Part of that peak was a double count: adopt()
 * charged the new view (whose order already holds the local band copies) while
 * the staging charge for those same local copies was still set. The swap must
 * never count the local bands twice: the view-buffer pool's high-water mark
 * during one sort is at most the previous view plus the adopted view.
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
import { FIRST_PAINT_WAIT } from './helpers/viewer-timing'

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

describe('worker-backed sort adoption does not double count the staged bands', () => {
  it('the view-buffer pool never exceeds the adopted view during the swap', async () => {
    let ledger: MainMemoryLedger | null = null
    let maxViewBuffers = 0
    const port = workerPort()
    const internal = { onMainLedger: (l: MainMemoryLedger) => {
      if (ledger === l) return
      ledger = l
      const set = l.set.bind(l)
      l.set = ((...args: Parameters<MainMemoryLedger['set']>) => { const out = set(...args); maxViewBuffers = Math.max(maxViewBuffers, l.snapshot().byPool.viewBuffers); return out }) as MainMemoryLedger['set']
    } } as object
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => port} {...internal} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /value name 1\b/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    const header = () => ui.getByRole('columnheader', { name: /^Column B/ })
    const before = header().getAttribute('aria-sort')
    const previous = maxViewBuffers = ledger!.snapshot().byPool.viewBuffers
    fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' }))
    await waitFor(() => expect(header().getAttribute('aria-sort')).not.toBe(before))
    const held = ledger!.snapshot().byPool.viewBuffers
    expect(held).toBeGreaterThan(0)
    // The previous view stays charged while its display still holds it (a real
    // overlap, by design); the staged local bands are already in the new view.
    expect(maxViewBuffers).toBeLessThanOrEqual(previous + held)
  })
})
