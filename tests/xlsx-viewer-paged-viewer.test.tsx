/**
 * S5.4 worker-backed interactive viewer for paged workbooks (spec §§4.1, 10.1–
 * 10.4, 11). With `interactiveSheets` and a viewer worker, a paged admission
 * opens in the worker and the interactive viewer reads bounded windows through
 * worker commands; with the flag off nothing changes. The worker here is the
 * real resident kernel behind a port with worker message semantics.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { OfficeFile } from '../src/components/OfficeFile'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import type { SheetFrameStats } from '../src/components/xlsx/InteractiveSheetDoc'
import { xlsxBytes } from './helpers/xlsx-package'

// Admission's object path allows one row, so these small fixtures are 'paged'
// on the main thread and in the worker alike (the same override both sides).
vi.mock('../src/xlsx/storage/importSession', async original => {
  const actual = await original<typeof import('../src/xlsx/storage/importSession')>()
  return { ...actual, importLimitsFor: (profile: 'D' | 'M', overrides?: Parameters<typeof actual.importLimitsFor>[1]) => actual.importLimitsFor(profile, { ...overrides, object: { ...overrides?.object, rows: 1 } }) }
})

afterEach(() => { cleanup(); vi.restoreAllMocks() })

const quiet = { yield: async () => {} }
const CORNER = '<row r="1"><c r="A1" t="inlineStr"><is><t>alpha</t></is></c><c r="B1"><v>3</v></c></row>'
  + '<row r="2"><c r="A2" t="inlineStr"><is><t>beta</t></is></c><c r="B2"><v>1</v></c></row>'
  + '<row r="3"><c r="A3" t="inlineStr"><is><t>gamma</t></is></c><c r="B3"><v>2</v></c></row>'
  + '<row r="1048576"><c r="XFD1048576" t="inlineStr"><is><t>corner</t></is></c></row>'
const workbook = (sheetData = CORNER) => new Blob([xlsxBytes({ sheetData }) as BlobPart])

/** A port with dedicated-worker semantics over a real kernel: commands and
 * replies are structured-cloned; replies can be held to model a slow worker. */
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
    addEventListener(_type: 'message', listener: (event: MessageEvent) => void) { listeners.add(listener) },
    removeEventListener(_type: 'message', listener: (event: MessageEvent) => void) { listeners.delete(listener) },
    terminate: vi.fn(),
  }
  return {
    port, kernel, parse, posted,
    reads: (op?: string) => posted.filter(c => c.kind === 'read' && (!op || c.query.op === op)).length,
    hold(predicate: (command: WorkerCommand) => boolean) { hold = predicate },
    release() { hold = null; for (const deliver of held.splice(0)) deliver() },
  }
}
function frames() {
  const all: SheetFrameStats[] = []
  return { all, onFrame: (stats: SheetFrameStats) => { all.push(stats) }, last: () => all[all.length - 1] }
}
function clipboard() { const write = vi.fn(async () => {}); Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: write } }); return write }
function go(ui: ReturnType<typeof render>, at: string) {
  fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: at } })
  fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
  const grid = ui.getByRole('grid')
  act(() => grid.focus())
  return grid
}
const internal = (props: Record<string, unknown>) => props as object

describe('S5.4 routing: interactiveSheets mounts the worker-backed viewer for paged admissions', () => {
  it('opens a paged admission in the worker (no main-thread parse, shell document) and shows notices once', async () => {
    const h = workerHarness()
    const factory = vi.fn(() => h.port)
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={factory} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy())
    expect(factory).toHaveBeenCalledTimes(1)
    const open = h.posted.find(command => command.kind === 'open')
    expect(open?.kind === 'open' && open.source.kind).toBe('blob')
    expect(h.parse).not.toHaveBeenCalled()
    // Cells came from bounded worker row reads, not from a main-thread document.
    expect(h.reads('rows')).toBeGreaterThan(0)
    expect(ui.container.textContent).not.toMatch(/compatibility mode/)
    await waitFor(() => expect(ui.container.textContent).toMatch(/Showing saved values/))
    expect(ui.container.textContent?.match(/Showing saved values/g)).toHaveLength(1)
  })

  it('flag off (or no worker): the paged admission keeps compatibility mode and never starts a worker', async () => {
    const factory = vi.fn(() => workerHarness().port)
    const off = render(<OfficeFile data={workbook()} viewerWorker={factory} />)
    await waitFor(() => expect(off.container.textContent).toMatch(/Large file — opened in compatibility mode/))
    expect(off.container.textContent?.match(/opened in compatibility mode/g)).toHaveLength(1)
    expect(off.queryByRole('grid')).toBeNull()
    off.unmount()
    const noWorker = render(<OfficeFile data={workbook()} interactiveSheets />)
    await waitFor(() => expect(noWorker.container.textContent).toMatch(/opened in compatibility mode/))
    expect(noWorker.container.textContent?.match(/opened in compatibility mode/g)).toHaveLength(1)
    expect(factory).not.toHaveBeenCalled()
  })
})

describe('S5.4 bounded worker reads: pending, cancellation and replacement', () => {
  it('unknown rows paint pending (never blank) until their bounded read lands', async () => {
    const h = workerHarness(), f = frames()
    h.hold(command => command.kind === 'read')
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} {...internal({ onFrame: f.onFrame })} />)
    await waitFor(() => expect(f.last()?.pendingTiles ?? 0).toBeGreaterThan(0))
    expect(ui.queryByRole('gridcell', { name: /alpha/ })).toBeNull()
    expect(ui.getAllByRole('gridcell', { name: /Loading/ }).length).toBeGreaterThan(0)
    act(() => h.release())
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy())
    await waitFor(() => expect(f.last()?.pendingTiles).toBe(0))
  })

  it('source replacement closes the old worker document and discards its late replies', async () => {
    const first = workerHarness(), second = workerHarness()
    const ports = [first.port, second.port]
    const factory = vi.fn(() => ports.shift()!)
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={factory} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy())
    // Hold the old worker's next replies, then replace the document.
    first.hold(command => command.kind === 'read')
    go(ui, 'A3')
    ui.rerender(<OfficeFile data={workbook(CORNER.replace('alpha', 'replaced'))} interactiveSheets viewerWorker={factory} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /replaced/ })).toBeTruthy())
    expect(first.posted.some(command => command.kind === 'close')).toBe(true)
    expect(first.port.terminate).toHaveBeenCalled()
    act(() => first.release())
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 10)) })
    expect(ui.queryByRole('gridcell', { name: /alpha/ })).toBeNull()
    expect(ui.getByRole('gridcell', { name: /replaced/ })).toBeTruthy()
  })
})

describe('S5.4 far corner and copy limits on a paged fixture', () => {
  it('reaches XFD1048576 by Go To, find and copy through worker reads', async () => {
    const h = workerHarness(), write = clipboard()
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy())
    const grid = go(ui, 'XFD1048576')
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /corner/ })).toBeTruthy())
    fireEvent.keyDown(grid, { key: 'c', ctrlKey: true })
    await waitFor(() => expect(write).toHaveBeenCalledWith('corner'))
    go(ui, 'A1')
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy())
    const scans = h.reads('scan')
    fireEvent.change(ui.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'corner' } })
    await waitFor(() => expect(ui.container.textContent).toMatch(/1 of 1/))
    expect(h.reads('scan')).toBeGreaterThan(scans)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /corner/ })).toBeTruthy())
  })

  it('refuses a full-sheet copy by the separator lower bound before any worker cell read', async () => {
    const h = workerHarness(), write = clipboard()
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy())
    // The viewer is worker-backed: its visible cells came from worker row reads.
    expect(h.reads('rows')).toBeGreaterThan(0)
    fireEvent.click(ui.getByRole('button', { name: 'Select worksheet' }))
    const before = h.reads()
    fireEvent.click(ui.getByRole('button', { name: 'Copy selection' }))
    await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(/Copy is too large/))
    expect(h.reads()).toBe(before)
    expect(write).not.toHaveBeenCalled()
  })
})

describe('S5.4 view commits run in the worker', () => {
  it('header sort and a value-list filter commit through worker sort/filter commands', async () => {
    const h = workerHarness()
    const ui = render(<OfficeFile data={workbook(CORNER.replace(/<row r="1048576">.*<\/row>/, ''))} interactiveSheets viewerWorker={() => h.port} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy())
    go(ui, 'B1')
    fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' }))
    await waitFor(() => expect(ui.getByRole('columnheader', { name: /Column B; sort priority 1, ascending/ })).toBeTruthy())
    expect(h.posted.some(command => command.kind === 'sort')).toBe(true)
    go(ui, 'A2')
    fireEvent.click(ui.getByRole('button', { name: /Filter column A/ }))
    await waitFor(() => expect(ui.getByRole('checkbox', { name: 'beta' })).toBeTruthy())
    // Clearing 'beta' hides its row: two of three body rows stay visible.
    fireEvent.click(ui.getByRole('checkbox', { name: 'beta' }))
    await waitFor(() => expect((ui.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(false))
    fireEvent.click(ui.getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(ui.getByRole('button', { name: /Filter column A/ }).getAttribute('aria-pressed')).toBe('true'))
    await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(/Visible body 2 \/ complete body 3/))
    const filter = h.posted.find(command => command.kind === 'filter')
    expect(filter?.kind === 'filter' && filter.keySets?.length).toBe(1)
  })
})

import { ViewerClient } from '../src/worker/viewerClient'
import { RemoteWorkbook, openRemoteWorkbook } from '../src/worker/viewerRemoteSource'

describe('S5.4 remote source: cancellation and ownership on the main thread', () => {
  async function remoteOf(h: ReturnType<typeof workerHarness>) {
    const client = new ViewerClient(h.port, 'remote')
    const opened = await openRemoteWorkbook(client, 'remote-doc', workbook(), 'D')
    if (!('layouts' in opened)) throw new Error(`open failed: ${JSON.stringify(opened)}`)
    return { client, remote: new RemoteWorkbook(client, opened) }
  }

  it('an aborted read settles at once as cancelled, sends cancel, and its late reply is discarded', async () => {
    const h = workerHarness()
    const { client, remote } = await remoteOf(h)
    const store = remote.store('1')
    h.hold(command => command.kind === 'read')
    const abort = new AbortController()
    const pending = store.readRows({ r0: 0, r1: 2, c0: 0, c1: 1 }, undefined, { signal: abort.signal, stamp: store.stamp })
    abort.abort()
    const page = await pending
    expect(page).toMatchObject({ completeness: 'incomplete', error: { code: 'cancelled' } })
    expect(h.posted.some(command => command.kind === 'cancel')).toBe(true)
    h.release()
    // The worker's own reply for the read still arrives; nobody waits for its value.
    await waitFor(() => expect(client.pendingCount).toBe(0))
    await remote.close()
  })

  it('a worker commit re-owns every later read; a read under the old view owner is refused as stale', async () => {
    const h = workerHarness()
    const { remote } = await remoteOf(h)
    const before = remote.identity('1')
    const commit = await remote.sort(before, { range: { r0: 0, r1: 2, c0: 0, c1: 1, headerRows: 0, totalsRows: 0 }, keys: [{ col: 1, order: 'asc' }], caseSensitive: false }, new AbortController().signal)
    expect(commit.kind).toBe('committed')
    expect(remote.identity('1').viewGen).toBe(before.viewGen + 1)
    // Display order follows the worker's committed bands (values 3,1,2 → rows 1,2,0).
    expect([0, 1, 2].map(slot => remote.view('1').order.displayToSource(slot))).toEqual([1, 2, 0])
    const replies: WorkerReply[] = []
    await h.kernel.handle({ kind: 'read', requestId: 'stale-read', owner: before, query: { op: 'cell', at: { row: 0, col: 0 } } }, out => { replies.push(out) })
    expect(replies.at(-1)).toMatchObject({ kind: 'error', code: 'stale-source' })
    const store = remote.store('1')
    expect(await store.readCell({ row: 0, col: 0 }, { signal: new AbortController().signal, stamp: store.stamp })).toMatchObject({ kind: 'cell', cell: { value: 'alpha' } })
    await remote.close()
    expect(h.posted.at(-1)?.kind).toBe('close')
  })
})
