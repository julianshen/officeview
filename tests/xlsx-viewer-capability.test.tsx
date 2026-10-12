/**
 * S10.3 capability failures for the worker-backed interactive viewer (spec
 * §§4.2, 10.3, 12.4; LIFE-03, AC-CAP-03). Both regressions were first
 * reproduced in Chromium on the e2e page (tests/browser/interactive-viewer-
 * capability.spec.ts):
 *
 * - Worker unavailable: a host `viewerWorker` factory that throws (`new
 *   Worker` with no Worker global) failed the whole document with "Worker is
 *   not defined", although a dead worker, or no worker at all, keeps
 *   compatibility mode with its notice.
 * - Storage quota denied while the worker imports pages: every read answered
 *   `incomplete` + `quota`, yet the status claimed "Visible body 0 / complete
 *   body 0" and cells stayed "Loading…" with no diagnosis — an incomplete
 *   source masquerading as an empty, complete one.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { OfficeFile } from '../src/components/OfficeFile'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import { xlsxBytes } from './helpers/xlsx-package'

// Admission's object path allows one row, so these small fixtures are 'paged'.
vi.mock('../src/xlsx/storage/importSession', async original => {
  const actual = await original<typeof import('../src/xlsx/storage/importSession')>()
  return { ...actual, importLimitsFor: (profile: 'D' | 'M', overrides?: Parameters<typeof actual.importLimitsFor>[1]) => actual.importLimitsFor(profile, { ...overrides, object: { ...overrides?.object, rows: 1 } }) }
})

afterEach(() => { cleanup(); vi.restoreAllMocks() })

const SHEET = '<row r="1"><c r="A1" t="inlineStr"><is><t>alpha</t></is></c><c r="B1"><v>3</v></c></row>'
  + '<row r="2"><c r="A2" t="inlineStr"><is><t>beta</t></is></c><c r="B2"><v>1</v></c></row>'
const workbook = () => new Blob([xlsxBytes({ sheetData: SHEET }) as BlobPart])

/** A real kernel behind a port with worker message semantics; `rewrite` may
 * replace a reply before delivery (to model a worker-side failure). */
function workerPort(rewrite?: (command: WorkerCommand, reply: WorkerReply) => WorkerReply, hold?: { on: boolean }) {
  const kernel = new ViewerKernel({ parse: async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)) }, { yield: async () => {} })
  const listeners = new Set<(event: MessageEvent) => void>()
  return {
    postMessage(message: unknown, transfer: Transferable[]) {
      const open = message as { kind: string; source?: { kind: string; blob?: Blob } }
      const blob = open.kind === 'open' && open.source?.kind === 'blob' ? open.source.blob : undefined
      const command = structuredClone(blob ? { ...open, source: { kind: 'blob' } } : message, { transfer: transfer as ArrayBuffer[] }) as WorkerCommand
      if (blob && command.kind === 'open') command.source = { kind: 'blob', blob }
      // A held read never answers (a slow worker): its caller stays pending.
      if (hold?.on && command.kind === 'read') return
      void kernel.handle(command, (out: WorkerReply, replyTransfer) => {
        const data = structuredClone(rewrite ? rewrite(command, out) : out, { transfer: replyTransfer })
        for (const listener of [...listeners]) listener({ data } as MessageEvent)
      })
    },
    addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
    removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
    terminate: vi.fn(),
  }
}

describe('Worker unavailable', () => {
  it('a viewerWorker factory that throws keeps compatibility mode with its notice, not a failed document', async () => {
    const factory = vi.fn(() => { throw new ReferenceError('Worker is not defined') })
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={factory} error={message => <div role="alert">{message}</div>} />)
    await waitFor(() => expect(ui.container.textContent).toMatch(/Large file — opened in compatibility mode/))
    expect(factory).toHaveBeenCalledTimes(1)
    expect(ui.queryByRole('alert')).toBeNull()
    expect(ui.getByRole('gridcell', { name: /^A1, value alpha/ })).toBeTruthy()
  })
})

describe('storage quota denied during the worker import', () => {
  // The exact reply shape Chromium's worker produced with IndexedDB puts failing
  // with QuotaExceededError (OPFS denied), for every read after open.
  const QUOTA = { code: 'quota', retryable: true, message: 'browser storage quota exceeded while writing a source page' }
  const quotaFailure = (command: WorkerCommand, reply: WorkerReply): WorkerReply => {
    if (command.kind !== 'read' || reply.kind !== 'complete' || reply.result.kind !== 'read') return reply
    const read = reply.result.result as { op: string; result?: Record<string, unknown>; page?: Record<string, unknown> }
    const stamp = (read.result ?? read.page)?.stamp
    if (read.op === 'structural') return { ...reply, result: { kind: 'read', result: { op: 'structural', result: { stamp, merges: [], tables: [], spills: [], drawings: [], completeness: 'incomplete', missing: [{ r0: 0, r1: 1048575, c0: 0, c1: 16383 }], error: QUOTA } } } } as unknown as WorkerReply
    if (read.op === 'complete') return { ...reply, result: { kind: 'read', result: { op: 'complete', result: { ...read.result, completeness: 'incomplete', missing: [read.result!.range], error: QUOTA } } } } as unknown as WorkerReply
    if (read.op === 'rows') return { ...reply, result: { kind: 'read', result: { op: 'rows', page: { items: [], completeness: 'incomplete', stamp, error: QUOTA } } } } as unknown as WorkerReply
    return reply
  }

  it('states the quota failure and an incomplete body; never "complete body 0"', async () => {
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => workerPort(quotaFailure)} />)
    const status = () => ui.getAllByRole('status').map(node => node.textContent ?? '').join(' ')
    await waitFor(() => expect(status()).toMatch(/browser storage quota exceeded while writing a source page/), { timeout: 4000 })
    expect(status()).toMatch(/\(incomplete\)/)
    expect(status()).not.toMatch(/complete body/)
  }, 10_000)
})

describe('Esc cancels a pending copy (spec §9.3 fall-through)', () => {
  // Reproduced by reading the handler: it compared the lower-cased key with
  // 'Escape', so Esc outside the grid never cancelled a copy waiting on values.
  it('Esc on the Cancel copy control stops a copy that waits for worker values', async () => {
    Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: vi.fn(async () => {}) } })
    const hold = { on: false }
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => workerPort(undefined, hold)} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /^A1, value alpha/ })).toBeTruthy())
    hold.on = true
    fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: 'A900000' } })
    fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
    fireEvent.click(ui.getByRole('button', { name: 'Copy selection' }))
    const cancel = await waitFor(() => ui.getByRole('button', { name: 'Cancel copy' }))
    act(() => cancel.focus())
    fireEvent.keyDown(cancel, { key: 'Escape' })
    await waitFor(() => expect(ui.queryByRole('button', { name: 'Cancel copy' })).toBeNull(), { timeout: 2000 })
    expect(ui.container.textContent).toMatch(/Copy cancelled\./)
  })
})
