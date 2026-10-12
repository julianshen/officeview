// @vitest-environment node
/**
 * S13.3 review round 2: worker viewport plans must not regress row-cache
 * semantics or warm paint.
 * - MAJOR-1 follow-up: tiles the row cache already proves skip worker requests
 *   (zero new viewport commands on revisit) and paint from the main-thread plan.
 * - MAJOR-2 follow-up: a worker plan that is not detail falls back to the
 *   main-thread planner, so warm tiles paint synchronously and stay retained.
 * - MINOR-2: retryable tile-plan diagnostics expire after a bounded backoff.
 * - MINOR-3: plan requests carry the frame owner, not the current identity.
 */
import { describe, expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { ViewerClient } from '../src/worker/viewerClient'
import { connectInProcess, ViewerKernel } from '../src/worker/viewerKernel'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import type { RemoteSheetStore } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { PLAN_RETRY_MAX_MS, PLAN_RETRY_MS, WorkerViewportPlans, tilePlanSource } from '../src/components/xlsx/workerViewport'
import { SheetTileRenderer, TILE_BYTES, type CanvasTile } from '../src/components/xlsx/sheetTiles'
import { BITMAP_BUDGETS } from '../src/xlsx/view/tileCache'
import { buildSheetGeometry, emptyWork, MergeQueryCache, type SheetGeometryResult } from '../src/xlsx/view/viewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import type { ViewportPlan } from '../src/xlsx/view/viewport'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { xlsxBytes } from './helpers/xlsx-package'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import type { CacheIdentity, LogicalCellBox, LogicalViewport } from '../src/xlsx/view/contracts'
import type { ProjectedSheetMetrics } from '../src/xlsx/view/projection'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'
import { workbookOf } from './helpers/xlsx-find'

const quiet = { yield: async () => {} }
const document = () => workbookOf([
  { sheetId: 's', cells: Array.from({ length: 40 }, (_, row) => [{ row, col: 0, value: row }, { row, col: 1, value: `t${row}` }]).flat() },
])
async function harness(doc: XlsxDocument = document()) {
  const source = createObjectWorkbookSource(doc, 'doc')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, quiet)
  const port = connectInProcess(kernel)
  const client = new ViewerClient(port)
  const open = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })
  expect(open.kind).toBe('complete')
  const session = kernel.session('doc')!
  const remote = new RemoteWorkbook(client, { docId: 'doc', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map(source.sheets.map(s => [s.sheetId, session.layout(s.sheetId)])) })
  return { doc, source, kernel, port, client, session, remote, store: remote.store('s') as RemoteSheetStore }
}
const metricsOf = (h: Awaited<ReturnType<typeof harness>>) => {
  const axes = h.remote.geometry('s')
  return createSheetMetrics({ rows: axes.rows, cols: axes.cols, order: createIdentityOrder(), visibility: visibilityFromHidden(createIntervalSet([])) })
}
const box = { x: 0, y: 0, width: 256, height: 200 }
const windowOf = (b = box): LogicalViewport => ({ box: { ...b }, zoom: 1, deviceScale: 1 })
const extentOf = (metrics: ProjectedSheetMetrics) => ({ width: metrics.cols.totalCss(), height: metrics.rows.totalCss() })

describe('row-cache coverage (RemoteSheetStore.coverage)', () => {
  it('a cold tile is unknown, a loaded one proven, one with no detail rows proven, an oversized one oversized', async () => {
    const h = await harness(), metrics = metricsOf(h)
    try {
      expect(h.store.coverage(metrics, box, 1)).toBe('unknown')
      await h.store.load([0, 1, 2, 3, 4, 5, 6, 7, 8, 9], 0, 255)
      expect(h.store.coverage(metrics, box, 1)).toBe('proven')
      // Below four device pixels nothing is read: the planner summarizes it.
      expect(h.store.coverage(metrics, { x: 0, y: 4000, width: 256, height: 200 }, 0.05)).toBe('proven')
      expect(h.store.coverage(metrics, { x: 0, y: 0, width: 256, height: 1_000_000 }, 1)).toBe('oversized')
    } finally { await h.remote.release() }
  })
})

describe('tilePlanSource: worker plans only for tiles the row cache cannot prove (MAJOR-2)', () => {
  const commands = (h: Awaited<ReturnType<typeof harness>>) => h.remote.stats.commands.viewport ?? 0
  it('a cold tile asks the worker once but plans on the main thread until a detail plan lands; a proven tile never asks', async () => {
    const h = await harness(), metrics = metricsOf(h)
    const plans = new WorkerViewportPlans(h.remote)
    const source = tilePlanSource({ store: h.store, metrics, scale: 1, identity: h.session.identity('s'), extent: extentOf(metrics), plans })
    try {
      // Pending worker plan: the main-thread planner paints (and pendingIn reads rows).
      expect(source(windowOf(), true)).toBeUndefined()
      expect(commands(h)).toBe(1)
      await h.port.idle()
      await new Promise(resolve => setTimeout(resolve, 0))
      // The landed detail plan is used while the rows are still unknown.
      expect(source(windowOf(), true)?.kind).toBe('detail')
      expect(h.remote.mainLedger.snapshot().byPool.layout, 'the stored plan is charged').toBeGreaterThan(0)
      await h.store.load([0, 1, 2, 3, 4, 5, 6, 7, 8, 9], 0, 255)
      // Proven: main-thread plan, no command, and the stored plan is released.
      expect(source(windowOf(), true)).toBeUndefined()
      expect(source({ ...windowOf(), box: { ...box, x: 0 } }, false)).toBeUndefined()
      expect(commands(h)).toBe(1)
      expect(h.remote.mainLedger.snapshot().byPool.layout, 'a proven tile releases its plan and charge').toBe(0)
    } finally { plans.dispose(); await h.remote.release() }
  })

  it('an oversized tile depends on its worker plan and shows the pending plan meanwhile', async () => {
    const h = await harness(), metrics = metricsOf(h)
    const plans = new WorkerViewportPlans(h.remote)
    const source = tilePlanSource({ store: h.store, metrics, scale: 1, identity: h.session.identity('s'), extent: extentOf(metrics), plans })
    try {
      expect(source(windowOf({ x: 0, y: 0, width: 256, height: 1_000_000 }), true)).toMatchObject({ kind: 'pending' })
      expect(commands(h)).toBe(1)
    } finally { plans.dispose(); await h.remote.release() }
  })

  it('ring tiles request plans only within their reserved share; a visible tile still finds a slot', async () => {
    const h = await harness()
    const plans = new WorkerViewportPlans(h.remote)
    const owner = h.session.identity('s'), extent = extentOf(metricsOf(h))
    const spy = vi.spyOn(h.remote, 'viewport').mockImplementation(() => new Promise(() => {}))
    const tile = (i: number) => windowOf({ x: i * 256, y: 0, width: 256, height: 200 })
    try {
      for (let i = 0; i < 4; i++) plans.get(owner, tile(i), extent, false)
      expect(spy).toHaveBeenCalledTimes(2)
      plans.get(owner, tile(10), extent, true)
      plans.get(owner, tile(11), extent, true)
      plans.get(owner, tile(12), extent, true)
      expect(spy).toHaveBeenCalledTimes(4)
    } finally { spy.mockRestore(); plans.dispose(); await h.remote.release() }
  })

  it('a pending ring request survives the next frame; one beyond the ring is cancelled', async () => {
    const h = await harness()
    const plans = new WorkerViewportPlans(h.remote)
    const owner = h.session.identity('s'), extent = extentOf(metricsOf(h))
    const signals: AbortSignal[] = []
    const spy = vi.spyOn(h.remote, 'viewport').mockImplementation((_owner, _window, signal) => { signals.push(signal); return new Promise(() => {}) })
    try {
      plans.get(owner, windowOf({ x: 512, y: 0, width: 512, height: 512 }), extent, false)
      plans.get(owner, windowOf({ x: 4096, y: 0, width: 512, height: 512 }), extent, false)
      plans.beginFrame(owner, windowOf({ x: 0, y: 0, width: 400, height: 300 }))
      expect(signals.map(signal => signal.aborted)).toEqual([false, true])
    } finally { spy.mockRestore(); plans.dispose(); await h.remote.release() }
  })
})

describe('WorkerViewportPlans: owner scope and retry (MINOR-2, MINOR-3)', () => {
  const complete = { kind: 'complete', result: { kind: 'viewport', completeness: 'complete', content: { kind: 'detail', cells: [] } } } as never

  it('MINOR-3: the request carries the frame owner its plan is stored under', async () => {
    const h = await harness()
    const plans = new WorkerViewportPlans(h.remote)
    // The frame still shows viewGen 0 while the workbook already knows a newer view.
    const frameOwner: CacheIdentity = h.session.identity('s')
    vi.spyOn(h.remote, 'identity').mockReturnValue({ ...frameOwner, viewGen: frameOwner.viewGen + 1 })
    const calls: CacheIdentity[] = []
    const spy = vi.spyOn(h.remote, 'viewport').mockImplementation(async owner => { calls.push(owner); return complete })
    try {
      plans.get(frameOwner, windowOf(), extentOf(metricsOf(h)), true)
      await new Promise(resolve => setTimeout(resolve, 0))
      expect(calls).toHaveLength(1)
      expect(calls[0]).toEqual(frameOwner)
    } finally { spy.mockRestore(); plans.dispose(); vi.restoreAllMocks(); await h.remote.release() }
  })

  it('MINOR-2: a retryable failure expires after a bounded, doubling backoff and the tile asks again', async () => {
    const h = await harness()
    const plans = new WorkerViewportPlans(h.remote)
    const owner = h.session.identity('s'), window = windowOf(), extent = extentOf(metricsOf(h))
    const timeout = { kind: 'cancelled', diagnostics: [{ code: 'query-timeout', message: 'Query timed out after 120 seconds' }] } as never
    const spy = vi.spyOn(h.remote, 'viewport').mockResolvedValue(timeout)
    const repaint = vi.fn(() => { plans.get(owner, window, extent, true) })
    const unsubscribe = h.remote.subscribe(repaint)
    vi.useFakeTimers()
    try {
      plans.get(owner, window, extent, true)
      await vi.advanceTimersByTimeAsync(0)
      expect(plans.get(owner, window, extent, true)).toMatchObject({ kind: 'diagnostic' })
      expect(spy).toHaveBeenCalledTimes(1)
      // The tile stays on screen with nothing else repainting it: the expiry
      // itself drives the retry.
      await vi.advanceTimersByTimeAsync(PLAN_RETRY_MS - 1)
      expect(spy).toHaveBeenCalledTimes(1)
      await vi.advanceTimersByTimeAsync(1)
      expect(spy).toHaveBeenCalledTimes(2)
      // The second failure backs off twice as long.
      await vi.advanceTimersByTimeAsync(PLAN_RETRY_MS * 2 - 1)
      expect(spy).toHaveBeenCalledTimes(2)
      await vi.advanceTimersByTimeAsync(1)
      expect(spy).toHaveBeenCalledTimes(3)
      // Bounded: never longer than the maximum, and a success ends it.
      spy.mockResolvedValue(complete)
      await vi.advanceTimersByTimeAsync(PLAN_RETRY_MAX_MS)
      expect(spy).toHaveBeenCalledTimes(4)
      expect(plans.get(owner, window, extent, true).kind).toBe('detail')
      await vi.advanceTimersByTimeAsync(PLAN_RETRY_MAX_MS * 4)
      expect(spy).toHaveBeenCalledTimes(4)
    } finally { vi.useRealTimers(); unsubscribe(); spy.mockRestore(); plans.dispose(); await h.remote.release() }
  })

  it('MINOR-2 control: a terminal failure is not retried while its tile stays on screen', async () => {
    const h = await harness()
    const plans = new WorkerViewportPlans(h.remote)
    const owner = h.session.identity('s'), window = windowOf(), extent = extentOf(metricsOf(h))
    const terminal = { kind: 'error', code: 'page-error', retryable: false, diagnostics: [{ code: 'page-error', message: 'bad tile' }] } as never
    const spy = vi.spyOn(h.remote, 'viewport').mockResolvedValue(terminal)
    const unsubscribe = h.remote.subscribe(() => { plans.get(owner, window, extent, true) })
    vi.useFakeTimers()
    try {
      plans.get(owner, window, extent, true)
      await vi.advanceTimersByTimeAsync(60_000)
      expect(plans.get(owner, window, extent, true)).toMatchObject({ kind: 'diagnostic' })
      expect(spy).toHaveBeenCalledTimes(1)
    } finally { vi.useRealTimers(); unsubscribe(); spy.mockRestore(); plans.dispose(); await h.remote.release() }
  })
})

describe('SheetTileRenderer and prepared plans', () => {
  const TILE = 512
  type Ready = Extract<SheetGeometryResult, { kind: 'ready' }>
  async function paintedSheet(): Promise<{ sheet: XlsxSheet; geometry: Ready }> {
    const bytes = xlsxBytes({ sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><t>alpha</t></is></c><c r="B1" t="inlineStr"><is><t>beta</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row>' })
    const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
    const geometry = buildSheetGeometry(doc.sheets[0])
    if (geometry.kind !== 'ready') throw new Error('resident geometry refused')
    return { sheet: doc.sheets[0], geometry }
  }
  function present(sheet: XlsxSheet, geometry: Ready, plan: ViewportPlan | undefined, pending: (box: LogicalCellBox) => readonly LogicalCellBox[]) {
    const createTile = (): CanvasTile => ({ canvas: createCanvas(TILE, TILE) as unknown as HTMLCanvasElement, width: TILE, height: TILE, close: () => {} })
    const r = new SheetTileRenderer('D', BITMAP_BUDGETS.D - TILE_BYTES * 4, createTile)
    const frame = {
      plan: () => plan, pending, sheet, geometry,
      identity: { docId: 'd', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 },
      origin: { x: 0, y: 0 }, viewport: { width: 128, height: 100 }, zoom: 1, dpr: 1,
      extent: { width: geometry.metrics.cols.totalCss(), height: geometry.metrics.rows.totalCss() },
      layoutCache: new LayoutCache(), merges: new MergeQueryCache(),
    }
    r.present(createCanvas(256, 128).getContext('2d') as unknown as CanvasRenderingContext2D, frame as never)
    return r
  }
  const detail = (extra: Partial<Extract<ViewportPlan, { kind: 'detail' }>> = {}): ViewportPlan => ({ kind: 'detail', cells: [], merges: [], drawings: [], rowLines: [], colLines: [], extent: { width: 1, height: 1 }, diagnostics: [], coarse: [], work: emptyWork(), ...extra })

  it('a worker detail plan paints from the plan and is retained once its rows are known', async () => {
    const { sheet, geometry } = await paintedSheet()
    const r = present(sheet, geometry, detail(), () => [])
    expect(r.cache.stats().tiles).toBe(1)
    expect(r.stats.pendingTiles).toBe(0)
    r.dispose()
  })

  it('a worker plan with unresolved cells still asks the row cache for its rows', async () => {
    const { sheet, geometry } = await paintedSheet()
    const pending = vi.fn((_box: LogicalCellBox) => [] as LogicalCellBox[])
    const r = present(sheet, geometry, detail({ pending: [{ bounds: { x: 0, y: 0, width: 64, height: 20 }, error: false }] }), pending)
    expect(pending, 'the unresolved rows are requested, or the tile never resolves').toHaveBeenCalled()
    expect(r.cache.stats().tiles).toBe(0)
    expect(r.stats.pendingTiles).toBe(1)
    r.dispose()
  })
})
