/**
 * S5.4 worker-backed interactive viewer for paged workbooks (spec §§4.1, 10.1–
 * 10.4, 11). With `interactiveSheets` and a viewer worker, a paged admission
 * opens in the worker and the interactive viewer reads bounded windows through
 * worker commands; with the flag off nothing changes. The worker here is the
 * real resident kernel behind a port with worker message semantics.
 */
import { afterEach, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { OfficeFile } from '../src/components/OfficeFile'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { CacheIdentity, WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import type { CellPlacement, ViewportPlan } from '../src/xlsx/view/viewport'
import type { SheetFrameStats } from '../src/components/xlsx/InteractiveSheetDoc'
import { xlsxBytes } from './helpers/xlsx-package'
import { FIRST_PAINT_WAIT } from './helpers/viewer-timing'

import { prepareSheetAssets } from '../src/components/xlsx/sheetAssets'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import * as workerPlanModule from '../src/components/xlsx/workerViewport'
import * as viewportPlanner from '../src/xlsx/view/viewport'

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
  const replied: WorkerReply[] = []
  const held: Array<{ command: WorkerCommand; out: WorkerReply; deliver(): void }> = []
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
        replied.push(structuredClone(out))
        const deliver = () => {
          const data = structuredClone(out, { transfer: replyTransfer })
          for (const listener of [...listeners]) listener({ data } as MessageEvent)
        }
        if (hold?.(command)) held.push({ command, out, deliver })
        else deliver()
      })
    },
    addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
    removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
    terminate: vi.fn(),
  }
  return {
    port, kernel, parse, posted, replied, held,
    reads: (op?: string) => posted.filter(c => c.kind === 'read' && (!op || c.query.op === op)).length,
    hold(predicate: (command: WorkerCommand) => boolean) { hold = predicate },
    releaseWhere(predicate: (command: WorkerCommand) => boolean) {
      for (let i = held.length - 1; i >= 0; i--) if (predicate(held[i].command)) held.splice(i, 1)[0].deliver()
    },
    release() { hold = null; for (const item of held.splice(0)) item.deliver() },
  }
}
function frames() {
  const all: SheetFrameStats[] = []
  return { all, onFrame: (stats: SheetFrameStats) => { all.push(stats) }, last: () => all[all.length - 1] }
}
function go(ui: ReturnType<typeof render>, at: string) {
  fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: at } })
  fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
  const grid = ui.getByRole('grid')
  act(() => grid.focus())
  return grid
}

const OVERFLOW_TEXT = 'target1000 abcdefghijklmnopqrstuvwxyz abcdefghijklmnopqrstuvwxyz'
const overflowData = () => workbook(CORNER.replace('<row r="1048576">', `<row r="1000"><c r="A1000" t="inlineStr"><is><t>${OVERFLOW_TEXT}</t></is></c><c r="C1000" t="inlineStr"><is><t>BLOCKER</t></is></c></row><row r="1048576">`))
const coldData = () => workbook(CORNER.replace('<row r="1048576">', '<row r="1000"><c r="A1000" t="inlineStr"><is><t>target1000</t></is></c></row><row r="1048576">'))
function deferred() {
  let resolve!: () => void
  const promise = new Promise<void>(done => { resolve = done })
  return { promise, resolve }
}
function observe(h: ReturnType<typeof workerHarness>) {
  const requests: Array<{ owner: CacheIdentity; remote: CacheIdentity; session: CacheIdentity; out?: unknown }> = []
  const tiles: Array<{ fontGen: number; coverage: string; kind?: string; plan?: ViewportPlan }> = []
  let remote!: RemoteWorkbook
  const viewport = RemoteWorkbook.prototype.viewport
  vi.spyOn(RemoteWorkbook.prototype, 'viewport').mockImplementation(async function(this: RemoteWorkbook, ...args) {
    remote = this
    const request = { owner: { ...args[0] }, remote: { ...this.identity(args[0].sheetId) }, session: { ...h.kernel.session(args[0].docId)!.identity(args[0].sheetId) }, out: undefined as unknown }
    requests.push(request)
    const out = await viewport.apply(this, args)
    request.out = out
    return out
  })
  const tileSource = workerPlanModule.tilePlanSource
  vi.spyOn(workerPlanModule, 'tilePlanSource').mockImplementation(options => {
    const original = tileSource(options)
    return (window, visible) => {
      const coverage = options.store.coverage(options.metrics, window.box, options.scale)
      const plan = original(window, visible)
      tiles.push({ fontGen: options.identity.fontGen, coverage, kind: plan?.kind, plan })
      return plan
    }
  })
  return { requests, tiles, remote: () => remote }
}

it('mounted paged viewport adopts worker overflow proofs after the first font lease', async () => {
  const h = workerHarness(), f = frames(), audit = observe(h), ready = deferred()
  const assetLoader: typeof prepareSheetAssets = vi.fn(async (...args: Parameters<typeof prepareSheetAssets>) => { await ready.promise; return prepareSheetAssets(...args) })
  const planner = vi.spyOn(viewportPlanner, 'planViewport')
  const layouts: Array<{ owner: CacheIdentity; placed: CellPlacement; body: ReturnType<LayoutCache['layout']> }> = []
  const layout = LayoutCache.prototype.layout
  vi.spyOn(LayoutCache.prototype, 'layout').mockImplementation(function(this: LayoutCache, owner, placed, host) {
    const body = layout.call(this, owner, placed, host)
    if (placed.source.row === 999 && placed.source.col === 0) layouts.push({ owner: { ...owner }, placed, body })
    return body
  })
  const seams = { assetLoader, onFrame: f.onFrame } as object
  const ui = render(<OfficeFile data={overflowData()} interactiveSheets viewerWorker={() => h.port} {...seams} />)
  try {
    // The mount's first frame waits on the lazy viewer chunk, the in-process worker handshake and the first paint
    // (tests/helpers/viewer-timing.ts): a cold start, not behaviour, so it takes the first-paint wait. This was the
    // assertion that failed under load, also on the base branch (C1 flake hardening).
    await waitFor(() => expect(assetLoader).toHaveBeenCalled(), FIRST_PAINT_WAIT)
    await waitFor(() => expect(f.last()?.identity.fontGen).toBe(0), FIRST_PAINT_WAIT)
    expect(audit.requests).toHaveLength(0)
    await act(async () => { ready.resolve(); await ready.promise })
    await waitFor(() => expect(f.last()?.identity.fontGen).toBe(1), FIRST_PAINT_WAIT)
    h.hold(c => c.kind === 'read' && c.query.op === 'rows')
    const before = planner.mock.calls.length
    go(ui, 'A1000')
    await waitFor(() => expect(h.replied.some(r => r.commandKind === 'viewport' && r.kind === 'complete')).toBe(true))
    await waitFor(() => expect(audit.tiles.some(t => t.fontGen === 1 && t.kind === 'detail')).toBe(true))
    // Detail alone is insufficient: the bridge must accept the wire owner's
    // proof and carry its blank-B1000 corridor through host text layout.
    const replies = h.replied.filter(r => r.commandKind === 'viewport' && r.kind === 'complete')
    const proofReply = replies.find(r => 'overflowProof' in r.result && r.result.overflowProof?.entries.some(e => e.row === 999 && e.col === 0))
    expect(proofReply, 'worker reply must include an overflow proof for A1000').toBeDefined()
    if (!proofReply || proofReply.kind !== 'complete' || proofReply.result.kind !== 'viewport') throw new Error('Expected viewport overflow proof')
    const proof = proofReply.result.overflowProof!
    expect(proof.owner).toEqual(proofReply.owner)
    expect(proof.owner.fontGen).toBe(0)
    expect(proof.scope).toBe('tile')
    expect(proof.entries.find(e => e.row === 999 && e.col === 0)).toMatchObject({ left: 0, right: 64, rightEnd: 'blocker' })
    const placement = () => audit.tiles.flatMap(t => t.fontGen === 1 && t.plan?.kind === 'detail' ? t.plan.cells : []).find(c => c.source.row === 999 && c.source.col === 0)
    await waitFor(() => expect(placement()?.overflow).toEqual({ left: 0, right: 64 }))
    // Row reads are still held: this cell was transactionally adopted from
    // the worker plan, rather than becoming known through a fallback read.
    expect(h.held.some(x => x.command.kind === 'read' && x.command.query.op === 'rows')).toBe(true)
    expect(audit.remote().store(f.last().activeSheetId).cellAt({ row: 999, col: 0 })?.value).toBe(OVERFLOW_TEXT)
    await waitFor(() => expect(layouts.some(x => x.owner.fontGen === 1 && x.placed.overflow?.right === 64 && x.body.text?.kind === 'plain' && x.body.text.clipX?.x0 === 0 && x.body.text.clipX.x1 === 128)).toBe(true))
    expect(f.last().identity.fontGen).toBe(1)
    expect(audit.requests.length).toBeGreaterThan(0)
    for (const request of audit.requests) {
      expect(request.owner.fontGen).toBe(0)
      expect(request.owner.fontGen).toBe(request.remote.fontGen)
      expect(request.owner.fontGen).toBe(request.session.fontGen)
    }
    expect(h.replied.filter(r => r.commandKind === 'viewport').some(r => r.diagnostics.some(d => d.code === 'stale-owner'))).toBe(false)
    // Initial pending tiles may fall back; settled worker detail must be used.
    expect(planner.mock.calls.length).toBeGreaterThan(before)
    act(() => h.release())
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /target1000/ })).toBeTruthy())
    await waitFor(() => expect(f.last()?.pendingTiles).toBe(0))
  } finally {
    h.release(); ui.unmount()
  }
}, 30_000)

it('a real font lease change rejects held old-font plans before adopting new-font replies', async () => {
  const h = workerHarness(), f = frames(), audit = observe(h), next = deferred()
  const assetLoader: typeof prepareSheetAssets = vi.fn(async (...args: Parameters<typeof prepareSheetAssets>) => {
    if (args[1]?.fallbackFonts) await next.promise
    return prepareSheetAssets(...args)
  })
  const data = coldData(), factory = () => h.port, seams = { assetLoader, onFrame: f.onFrame } as object
  const ui = render(<OfficeFile data={data} interactiveSheets viewerWorker={factory} {...seams} />)
  try {
    await waitFor(() => expect(f.last()?.identity.fontGen).toBe(1), FIRST_PAINT_WAIT)
    h.hold(c => c.kind === 'viewport' || c.kind === 'read' && c.query.op === 'rows')
    go(ui, 'A1000')
    await waitFor(() => expect(h.held.some(x => x.command.kind === 'viewport' && x.out.kind === 'complete')).toBe(true))
    const oldRequests = new Set(h.held.filter(x => x.command.kind === 'viewport').map(x => x.command.requestId))
    const remote = audit.remote(), store = remote.store(f.last().activeSheetId)
    const adopt = vi.spyOn(store, 'stageViewport')
    const original = { ...f.last().identity }, zoom = f.last().zoom
    ui.rerender(<OfficeFile data={data} interactiveSheets viewerWorker={factory} fallbackFonts={{ fallbackChain: ['Arial'] }} {...seams} />)
    await waitFor(() => expect(assetLoader).toHaveBeenCalledWith(expect.anything(), expect.objectContaining({ fallbackFonts: { fallbackChain: ['Arial'] } })))
    expect(f.last().identity.fontGen).toBe(1)
    await act(async () => { next.resolve(); await next.promise })
    await waitFor(() => expect(f.last()?.identity.fontGen).toBe(2))
    const before = adopt.mock.calls.length
    await act(async () => {
      h.releaseWhere(c => oldRequests.has(c.requestId))
      for (let i = 0; i < 20; i++) await Promise.resolve()
    })
    expect(adopt.mock.calls.length).toBe(before)
    expect(audit.tiles.some(t => t.fontGen === 2 && t.kind === 'detail')).toBe(false)
    await waitFor(() => expect(h.held.some(x => x.command.kind === 'viewport' && !oldRequests.has(x.command.requestId) && x.out.kind === 'complete')).toBe(true))
    act(() => h.releaseWhere(c => c.kind === 'viewport'))
    await waitFor(() => expect(audit.tiles.some(t => t.fontGen === 2 && t.kind === 'detail')).toBe(true))
    expect(adopt.mock.calls.length).toBeGreaterThan(before)
    expect(f.last().zoom).toBe(zoom)
    expect(f.last().identity).toEqual({ ...original, fontGen: 2 })
    act(() => h.release())
    await waitFor(() => expect(f.last()?.pendingTiles).toBe(0))
  } finally { h.release(); ui.unmount() }
}, 30_000)
