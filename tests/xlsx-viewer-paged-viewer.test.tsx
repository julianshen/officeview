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
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'

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
  it('opens a paged admission in the worker (no main-thread parse, shell document) and shows notices once', FIRST_PAINT_TEST, async () => {
    const mainPlanner = vi.spyOn(viewportPlanner, 'planViewport')
    const h = workerHarness()
    // Hold worker viewport replies: paint must fall back to the main-thread
    // planner (S13.3 review round 2) instead of waiting on plans that never land.
    h.hold(command => command.kind === 'viewport')
    const factory = vi.fn(() => h.port)
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={factory} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    // The main-thread planner is the fallback, not the painter: it runs while
    // worker plans are still pending, and must paint the same records (the
    // previous never-called assertion encoded the exclusive-plans design that
    // left blank cells "Loading…" forever).
    expect(mainPlanner).toHaveBeenCalled()
    act(() => h.release())
    expect(factory).toHaveBeenCalledTimes(1)
    const open = h.posted.find(command => command.kind === 'open')
    expect(open?.kind === 'open' && open.source.kind).toBe('blob')
    expect(h.parse).not.toHaveBeenCalled()
    // Paint and semantics consume bounded worker viewport records.
    await waitFor(() => expect(h.posted.some(command => command.kind === 'viewport')).toBe(true))
    expect(ui.container.textContent).not.toMatch(/compatibility mode/)
    await waitFor(() => expect(ui.container.textContent).toMatch(/Showing saved values/))
    expect(ui.container.textContent?.match(/Showing saved values/g)).toHaveLength(1)
  })

  it('removing the source closes the worker document AND stops rendering its grid', FIRST_PAINT_TEST, async () => {
    const h = workerHarness()
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} loading={<p>no file</p>} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    ui.rerender(<OfficeFile data={undefined} interactiveSheets viewerWorker={() => h.port} loading={<p>no file</p>} />)
    await waitFor(() => expect(h.port.terminate).toHaveBeenCalled())
    expect(ui.queryByRole('grid')).toBeNull()
    expect(ui.getByText('no file')).toBeTruthy()
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

  it('source replacement closes the old worker document and discards its late replies', FIRST_PAINT_TEST, async () => {
    const first = workerHarness(), second = workerHarness()
    const ports = [first.port, second.port]
    const factory = vi.fn(() => ports.shift()!)
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={factory} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    // Hold the old worker's next replies, then replace the document.
    first.hold(command => command.kind === 'read')
    go(ui, 'A3')
    ui.rerender(<OfficeFile data={workbook(CORNER.replace('alpha', 'replaced'))} interactiveSheets viewerWorker={factory} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /replaced/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    expect(first.posted.some(command => command.kind === 'close')).toBe(true)
    expect(first.port.terminate).toHaveBeenCalled()
    act(() => first.release())
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 10)) })
    expect(ui.queryByRole('gridcell', { name: /alpha/ })).toBeNull()
    expect(ui.getByRole('gridcell', { name: /replaced/ })).toBeTruthy()
  })
})

describe('S5.4 far corner and copy limits on a paged fixture', () => {
  it('shows the worker query deadline in the viewer status and preserves the rendered sheet', FIRST_PAINT_TEST, async () => {
    const h = workerHarness()
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    const open = h.posted.find(command => command.kind === 'open')!
    if (open.kind !== 'open') throw new Error('Expected open command')
    const session = h.kernel.session(open.docId)!
    let began!: () => void, readSignal: AbortSignal | undefined
    const started = new Promise<void>(resolve => { began = resolve })
    vi.spyOn(session.source.store('1'), 'scanPopulated').mockImplementation(async function* (_query, options) {
      readSignal = options.signal
      began()
      await new Promise(() => {})
    })
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'Date'] })
    try {
      fireEvent.change(ui.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'alpha' } })
      await act(async () => { await started })
      await act(async () => { await vi.advanceTimersByTimeAsync(120_001) })
      expect(ui.container.textContent).toMatch(/Query timed out after 120 seconds/)
      expect(ui.container.textContent?.match(/Query timed out after 120 seconds/g)).toHaveLength(1)
      expect(readSignal!.aborted).toBe(true)
      expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy()
      expect(h.kernel.lifecycleStats()).toMatchObject({ inflight: 0, transientBytes: 0 })
    } finally { ui.unmount(); vi.useRealTimers() }
  })

  it('reaches XFD1048576 by Go To, find and copy through worker reads', FIRST_PAINT_TEST, async () => {
    const h = workerHarness(), write = clipboard()
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    const grid = go(ui, 'XFD1048576')
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /corner/ })).toBeTruthy())
    fireEvent.keyDown(grid, { key: 'c', ctrlKey: true })
    await waitFor(() => expect(write).toHaveBeenCalledWith('corner'))
    go(ui, 'A1')
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy())
    const searches = h.posted.filter(command => command.kind === 'search').length
    fireEvent.change(ui.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'corner' } })
    await waitFor(() => expect(ui.container.textContent).toMatch(/1 of 1/))
    expect(h.posted.filter(command => command.kind === 'search').length).toBeGreaterThan(searches)
    expect(h.reads('scan')).toBe(0)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /corner/ })).toBeTruthy())
  })

  it('refuses a full-sheet copy by the separator lower bound before any worker cell read', FIRST_PAINT_TEST, async () => {
    const h = workerHarness(), write = clipboard()
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    // Visible records were planned by the worker; copy refusal starts no additional reads.
    expect(h.posted.some(command => command.kind === 'viewport')).toBe(true)
    fireEvent.click(ui.getByRole('button', { name: 'Select worksheet' }))
    const before = h.reads()
    fireEvent.click(ui.getByRole('button', { name: 'Copy selection' }))
    await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(/Copy is too large/))
    expect(h.reads()).toBe(before)
    expect(write).not.toHaveBeenCalled()
  })
})

describe('S5.4 view commits run in the worker', () => {
  it('header sort and a value-list filter commit through worker sort/filter commands', FIRST_PAINT_TEST, async () => {
    const h = workerHarness()
    const ui = render(<OfficeFile data={workbook(CORNER.replace(/<row r="1048576">.*<\/row>/, ''))} interactiveSheets viewerWorker={() => h.port} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    go(ui, 'B1')
    fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' }))
    await waitFor(() => expect(ui.getByRole('columnheader', { name: /Column B; sort priority 1, ascending/ })).toBeTruthy())
    expect(h.posted.some(command => command.kind === 'sort')).toBe(true)
    // The control learns the worker commit succeeded (not a stale/failed notice).
    await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(/View-only sort applied/))
    go(ui, 'A2')
    fireEvent.click(ui.getByRole('button', { name: /Filter column A/ }))
    await waitFor(() => expect(ui.getByRole('checkbox', { name: 'beta' })).toBeTruthy())
    // Clearing 'beta' hides its row: two of three body rows stay visible.
    fireEvent.click(ui.getByRole('checkbox', { name: 'beta' }))
    await waitFor(() => expect((ui.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(false))
    fireEvent.click(ui.getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(ui.getByRole('button', { name: /Filter column A/ }).getAttribute('aria-pressed')).toBe('true'))
    await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(/Visible body 2 \/ complete body 3/))
    await waitFor(() => expect(ui.queryByRole('dialog', { name: 'Column filter' })).toBeNull())
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

// ---------------------------------------------------------------------------
// Review round 1 (codex) regressions
// ---------------------------------------------------------------------------
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { createControlsClient, type ControlsClientOptions } from '../src/components/xlsx/controlsClient'
import * as viewportPlanner from '../src/xlsx/view/viewport'
import { createRowOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import type { RemoteOpen } from '../src/worker/viewerRemoteSource'

const DATE_STYLES = '<numFmts count="0"/><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs>'
const DATED = '<row r="1"><c r="A1" s="1"><v>44927</v></c><c r="B1"><v>1</v></c></row><row r="2"><c r="A2" s="1"><v>44928</v></c><c r="B2"><v>2</v></c></row>'
const HIDDEN_COLS = '<cols><col min="2" max="2" hidden="1" width="9" customWidth="1"/></cols>'
const SECRET = '<row r="1"><c r="A1" t="inlineStr"><is><t>shown</t></is></c><c r="B1" t="inlineStr"><is><t>hidden-secret</t></is></c><c r="C1" t="inlineStr"><is><t>after</t></is></c></row>'
  + '<row r="2"><c r="A2"><v>2</v></c></row>'
const pkg = (spec: Parameters<typeof xlsxBytes>[0]) => xlsxBytes(spec)

async function openRemote(h: ReturnType<typeof workerHarness>, bytes = xlsxBytes({ sheetData: CORNER }), extra: Partial<RemoteOpen> = {}) {
  const client = new ViewerClient(h.port, 'review')
  const opened = await openRemoteWorkbook(client, 'review-doc', new Blob([bytes as BlobPart]), 'D')
  if (!('layouts' in opened)) throw new Error(`open failed: ${JSON.stringify(opened)}`)
  return { client, remote: new RemoteWorkbook(client, { ...opened, ...extra }) }
}
async function residentCopy(bytes: Uint8Array, at: string): Promise<string> {
  const write = clipboard()
  const doc = await parseXlsx(await OfficePackage.load(bytes))
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  const grid = go(ui, at)
  fireEvent.keyDown(grid, { key: 'c', ctrlKey: true })
  await waitFor(() => expect(write).toHaveBeenCalled())
  ui.unmount()
  return (write.mock.calls[0] as unknown as [string])[0]
}

describe('S5.4 review: worker-backed cells keep resident semantics', () => {
  it('F1 formats paged cells with the workbook style table and date system (copy and find parity)', async () => {
    const bytes = pkg({ sheetData: DATED, styles: DATE_STYLES })
    const expected = await residentCopy(bytes, 'A1')
    expect(expected).not.toBe('44927')
    const h = workerHarness(), write = clipboard()
    const ui = render(<OfficeFile data={new Blob([bytes as BlobPart])} interactiveSheets viewerWorker={() => h.port} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: new RegExp(expected.replace(/[/.]/g, '\\$&')) })).toBeTruthy())
    const grid = go(ui, 'A1')
    fireEvent.keyDown(grid, { key: 'c', ctrlKey: true })
    await waitFor(() => expect(write).toHaveBeenCalledWith(expected))
    fireEvent.change(ui.getByRole('searchbox', { name: 'Search document' }), { target: { value: expected } })
    await waitFor(() => expect(ui.container.textContent).toMatch(/1 of 1/))
  })

  it('F2 authored-hidden columns stay collapsed: a row copy across them omits the hidden value', FIRST_PAINT_TEST, async () => {
    const bytes = pkg({ sheetData: SECRET, beforeSheetData: HIDDEN_COLS })
    const h = workerHarness(), write = clipboard()
    const ui = render(<OfficeFile data={new Blob([bytes as BlobPart])} interactiveSheets viewerWorker={() => h.port} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /shown/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    go(ui, 'A1')
    const grid = ui.getByRole('grid')
    fireEvent.keyDown(grid, { key: 'ArrowRight', shiftKey: true })
    fireEvent.keyDown(grid, { key: 'c', ctrlKey: true })
    await waitFor(() => expect(write).toHaveBeenCalled())
    const text = (write.mock.calls[0] as unknown as [string])[0]
    expect(text).not.toContain('hidden-secret')
    expect(text).toBe('shown\tafter')
    expect(ui.queryByRole('gridcell', { name: /hidden-secret/ })).toBeNull()
  })
})

describe('S5.4 review: bounded main-thread caches and reads', () => {
  it('F3 one document-wide byte ledger bounds the row cache; inactive sheets release theirs', async () => {
    const long = 'x'.repeat(2000)
    const rows = Array.from({ length: 40 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>${long}</t></is></c></row>`).join('')
    const h = workerHarness()
    const { remote } = await openRemote(h, pkg({ sheetData: rows }), { rowCacheBytes: 64 * 1024 })
    const store = remote.store('1')
    for (let row = 0; row < 40; row++) {
      expect(await store.readCell({ row, col: 0 }, { signal: new AbortController().signal, stamp: store.stamp })).toMatchObject({ kind: 'cell', cell: { value: long } })
    }
    expect(remote.ledger.peakBytes).toBeLessThanOrEqual(64 * 1024)
    expect(remote.ledger.bytes).toBeLessThanOrEqual(64 * 1024)
    expect(store.cachedRows).toBeGreaterThan(0)
    remote.activate('another-sheet')
    expect(store.cachedRows).toBe(0)
    expect(remote.ledger.bytes).toBe(0)
    await remote.close()
  })

  it('F4 tiles read only admitted detail rows, coalesced and capped per tile and per frame', async () => {
    const h = workerHarness()
    const { remote } = await openRemote(h)
    const store = remote.store('1'), axes = remote.geometry('1')
    const visibility = visibilityFromHidden(createIntervalSet([]))
    // Sub-detail rows (20 CSS px at scale 0.1 = 2 device px) are summarized: no reads.
    const identity = createSheetMetrics({ rows: axes.rows, cols: axes.cols, order: createRowOrder([]), visibility })
    expect(store.pendingIn(identity, { x: 0, y: 0, width: 5000, height: 10_000 }, 0.1)).toEqual([])
    expect(h.reads()).toBe(0)
    // Scattered display order: 512 slots map to source rows 1,000 apart.
    const scattered = Uint32Array.from({ length: 512 }, (_, i) => i * 1000 + 999 - (i % 2) * 999)
    const order = createRowOrder([{ r0: 0, r1: 511_999, sourceAtOffset: Uint32Array.from({ length: 512_000 }, (_, d) => d < 512 ? scattered[d] : 0) }].map(band => {
      // Complete permutation: the first 512 offsets show the scattered rows; the rest fill in ascending order.
      const used = new Set(scattered), rest: number[] = []
      for (let r = 0; r < 512_000; r++) if (!used.has(r)) rest.push(r)
      return { ...band, sourceAtOffset: Uint32Array.from([...scattered, ...rest]) }
    }))
    const metrics = createSheetMetrics({ rows: axes.rows, cols: axes.cols, order, visibility })
    h.hold(command => command.kind === 'read')
    const before = h.reads('rows')
    const pending = store.pendingIn(metrics, { x: 0, y: 0, width: 512, height: 512 * 20 }, 1)
    expect(pending.length).toBeGreaterThan(0)
    expect(h.reads('rows') - before).toBeLessThanOrEqual(remote.pageLimits.tile)
    for (let tile = 1; tile < 8; tile++) store.pendingIn(metrics, { x: 0, y: tile * 512 * 20, width: 512, height: 512 * 20 }, 1)
    expect(h.reads('rows') - before).toBeLessThanOrEqual(remote.pageLimits.frame)
    h.release()
    await remote.close()
  })

  it('F7 cache hits obey the source guards: aborted signal, foreign stamp and close', async () => {
    const h = workerHarness()
    const { remote } = await openRemote(h)
    const store = remote.store('1')
    expect(await store.readCell({ row: 0, col: 0 }, { signal: new AbortController().signal, stamp: store.stamp })).toMatchObject({ kind: 'cell', cell: { value: 'alpha' } })
    const aborted = new AbortController(); aborted.abort()
    expect(await store.readCell({ row: 0, col: 0 }, { signal: aborted.signal, stamp: store.stamp })).toMatchObject({ kind: 'error', code: 'cancelled' })
    expect(await store.readCell({ row: 0, col: 0 }, { signal: new AbortController().signal, stamp: { ...store.stamp, docId: 'foreign' } })).toMatchObject({ kind: 'error', code: 'stale-source' })
    const stamp = store.stamp
    await remote.close()
    expect(store.cachedRows).toBe(0)
    expect(await store.readCell({ row: 0, col: 0 }, { signal: new AbortController().signal, stamp })).toMatchObject({ kind: 'error', code: 'stale-source' })
    expect(store.cellsOf(0)).toBeUndefined()
  })

  it('F8 a column query never detaches the caller-owned membership buffer', async () => {
    const h = workerHarness()
    const { remote } = await openRemote(h)
    const store = remote.store('1')
    const rows = Uint32Array.of(0, 1, 2)
    const page = await store.queryColumn({ col: 0, rows: { kind: 'list', r0: 0, r1: 2, rows }, limit: 10 }, { signal: new AbortController().signal, stamp: store.stamp })
    expect(page.items.filter(item => item.kind === 'cell')).toHaveLength(3)
    expect(rows.byteLength).toBe(12)
    expect([...rows]).toEqual([0, 1, 2])
    await remote.close()
  })

  it('F11 an abandoned scan iteration releases its worker continuation', async () => {
    const h = workerHarness()
    const { remote } = await openRemote(h)
    const store = remote.store('1')
    const iterator = store.scanPopulated({ range: { r0: 0, r1: 1048575, c0: 0, c1: 16383 }, limit: 1 }, { signal: new AbortController().signal, stamp: store.stamp })[Symbol.asyncIterator]()
    expect((await iterator.next()).value?.items).toHaveLength(1)
    expect(h.kernel.session('review-doc')!.openScans).toBe(1)
    await iterator.return?.()
    await waitFor(() => expect(h.kernel.session('review-doc')!.openScans).toBe(0))
    expect(h.posted.some(command => command.kind === 'read' && command.query.op === 'releaseScan')).toBe(true)
    await remote.close()
  })
})

describe('S5.4 review: lifecycle and publication', () => {
  it('F6 unmounting during a pending open closes the half-open document and terminates the worker', async () => {
    const h = workerHarness()
    h.hold(command => command.kind === 'open')
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} />)
    await waitFor(() => expect(h.posted.some(command => command.kind === 'open')).toBe(true))
    ui.unmount()
    await waitFor(() => expect(h.port.terminate).toHaveBeenCalled())
    expect(h.posted.some(command => command.kind === 'close')).toBe(true)
  })

  it('F9 a layout publication repaints tiles and refreshes the scroll extent', FIRST_PAINT_TEST, async () => {
    const h = workerHarness(), f = frames()
    const { remote } = await openRemote(h)
    const ui = render(<InteractiveSheetDoc doc={{ sheets: [{ name: 'S', sheetId: '1', rows: [], cols: [], merges: [], mergeRanges: [] }] }} source={remote} onFrame={f.onFrame} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    await waitFor(() => expect(f.last()?.pendingTiles).toBe(0))
    const rowsBefore = ui.getAllByRole('row').length, paints = f.last()!.basePaints
    act(() => remote.publishLayout('1', { ...remote.layout('1'), rowOverrides: [{ start: 0, end: 0, sizeCss: 200 }] }))
    await waitFor(() => expect(f.last()!.basePaints).toBeGreaterThan(paints))
    expect(f.last()!.tileHits).toBeGreaterThanOrEqual(0)
    // The refreshed projection shows fewer rows in the same pane (row 1 is now 200px).
    await waitFor(() => expect(ui.getAllByRole('row').length).toBeLessThan(rowsBefore))
    await remote.close()
  })

  it('F10 a worker-committed stage still refuses publication when the source generation moved', async () => {
    let valueGen = 0
    const owner = () => ({ docId: 'd', sheetId: 's', valueGen, fontGen: 0, viewGen: 0 })
    const publish = vi.fn(() => ({ owner: owner(), filter: null, sort: null, records: { visible: 1, total: 1, completeness: 'complete' as const } }))
    const client = createControlsClient({
      snapshot: () => ({ owner: owner(), filter: null, sort: null, records: { visible: 1, total: 1, completeness: 'complete' } }),
      stageSort: async () => { valueGen++; return { kind: 'ready', committed: true, publish } },
    } as unknown as ControlsClientOptions)
    const result = await client.sort(owner(), { range: { r0: 0, r1: 1, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false }, new AbortController().signal)
    expect(result).toEqual({ kind: 'stale' })
    expect(publish).not.toHaveBeenCalled()
  })
})

describe('S13.3 review round 2, NIT-1: a search over an importing sheet catches up', () => {
  it('re-runs a finished search once its incomplete sheets have loaded', FIRST_PAINT_TEST, async () => {
    const h = workerHarness()
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    const open = h.posted.find(command => command.kind === 'open')!
    if (open.kind !== 'open') throw new Error('Expected open command')
    const store = h.kernel.session(open.docId)!.source.store('1')
    // The first search sees the sheet's structure still importing.
    const structural = store.queryStructural.bind(store)
    vi.spyOn(store, 'queryStructural').mockImplementationOnce(async (rect, options) => ({ ...(await structural(rect, options)), completeness: 'incomplete' }))
    const searches = () => h.posted.filter(command => command.kind === 'search').length
    fireEvent.change(ui.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'alpha' } })
    // The sheet is already loaded on the main side, so the incomplete page is
    // re-run and the "still loading" qualifier goes away instead of sticking.
    await waitFor(() => expect(searches()).toBe(2))
    await waitFor(() => expect(ui.container.textContent).toMatch(/1 of 1/))
    expect(ui.container.textContent).not.toMatch(/some sheets still loading/)
    expect(store.queryStructural).toHaveBeenCalledTimes(2)
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 50)) })
    expect(searches(), 'one refresh, not a loop').toBe(2)
    ui.unmount()
  })
})

describe('S13.3 review round 2: worker tile plans keep row-cache semantics and warm paint', () => {
  const viewportCommands = (h: ReturnType<typeof workerHarness>) => h.posted.filter(command => command.kind === 'viewport').length
  const settled = async (ui: ReturnType<typeof render>, f: ReturnType<typeof frames>) => {
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    await waitFor(() => expect(f.last()?.pendingTiles).toBe(0))
  }

  it('MAJOR-1: blank visible cells read as blank, not Loading, once their rows are known', FIRST_PAINT_TEST, async () => {
    const h = workerHarness(), f = frames()
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} {...internal({ onFrame: f.onFrame })} />)
    await settled(ui, f)
    // C1, D1, E1… are authored blank: once the bounded row reads land they are
    // known blanks, never a permanent "Loading…".
    await waitFor(() => expect(ui.queryAllByRole('gridcell', { name: /Loading/ })).toHaveLength(0))
    const blank = ui.getAllByRole('gridcell').find(cell => /\bC1\b/.test(cell.getAttribute('aria-label') ?? cell.textContent ?? ''))
    expect(blank, 'C1 is a visible gridcell').toBeTruthy()
    expect(blank!.getAttribute('aria-label') ?? blank!.textContent).not.toMatch(/Loading/)
    expect(h.reads('rows')).toBeGreaterThan(0)
    ui.unmount()
  })

  it('MAJOR-2: ring tiles are prefetched and retained over the row cache', FIRST_PAINT_TEST, async () => {
    const h = workerHarness(), f = frames()
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} {...internal({ onFrame: f.onFrame })} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    // Every visible tile of the settled first frame is retained...
    const firstSettled = f.all.find(stats => stats.pendingTiles === 0 && stats.lastRequests.length > 0 && stats.prefetched === 0)
    await waitFor(() => expect(f.last()?.pendingTiles).toBe(0))
    // ...and the ring around them is painted and retained on top.
    await waitFor(() => expect(f.last()?.prefetched ?? 0).toBeGreaterThan(0))
    const visible = firstSettled?.lastRequests.length ?? f.all.reduce((most, stats) => Math.max(most, stats.lastRequests.length), 0)
    expect(visible).toBeGreaterThan(0)
    expect(f.last()!.tiles.tiles, 'ring tiles are retained beyond the visible ones').toBeGreaterThan(visible)
    ui.unmount()
  })

  it('MAJOR-2: a warm area repaints at a new zoom with no viewport command and no pending tile', FIRST_PAINT_TEST, async () => {
    const h = workerHarness(), f = frames()
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} {...internal({ onFrame: f.onFrame })} />)
    await settled(ui, f)
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 20)) })
    const commands = viewportCommands(h), seen = f.all.length
    // Zooming in shows a subset of the rows already read: every visible tile
    // is new (the tile cache is per zoom) but the row cache proves it.
    act(() => { fireEvent.click(ui.getByRole('button', { name: 'Zoom in' })) })
    await waitFor(() => expect(f.all.length).toBeGreaterThan(seen))
    const first = f.all.slice(seen).find(stats => stats.zoom !== f.all[seen - 1].zoom)
    expect(first, 'a frame at the new zoom').toBeTruthy()
    expect(first!.pendingTiles, 'the first frame at the new zoom paints every tile at once').toBe(0)
    expect(viewportCommands(h), 'tiles the row cache proves never ask the worker').toBe(commands)
    ui.unmount()
  })
})
