// @vitest-environment node
/**
 * View-only resizing on worker-backed (paged) sheets (spec §5.6): main-thread
 * geometry carries the session sizes, worker viewport plans (laid out in the
 * workbook geometry) are never consulted while a sheet is resized, and a plan
 * requested under an older layout generation is refused.
 */
import { describe, expect, it, vi } from 'vitest'
import { ViewerClient } from '../src/worker/viewerClient'
import { connectInProcess, ViewerKernel } from '../src/worker/viewerKernel'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import type { RemoteSheetStore } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { WorkerViewportPlans, tilePlanSource } from '../src/components/xlsx/workerViewport'
import { pagedSheetGeometry } from '../src/components/xlsx/pagedGeometry'
import * as axesModule from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createIdentityOrder, createRowOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import type { LogicalViewport } from '../src/xlsx/view/contracts'
import type { ProjectedSheetMetrics } from '../src/xlsx/view/projection'
import type { XlsxDocument } from '../src/xlsx/types'
import { workbookOf } from './helpers/xlsx-find'

const quiet = { yield: async () => {} }
const fixture = () => workbookOf([
  { sheetId: 's', cells: Array.from({ length: 40 }, (_, row) => [{ row, col: 0, value: row }, { row, col: 1, value: `t${row}` }]).flat() },
])
async function harness(doc: XlsxDocument = fixture()) {
  const source = createObjectWorkbookSource(doc, 'doc')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, quiet)
  const port = connectInProcess(kernel)
  const client = new ViewerClient(port)
  const open = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })
  expect(open.kind).toBe('complete')
  const session = kernel.session('doc')!
  const remote = new RemoteWorkbook(client, { docId: 'doc', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map(source.sheets.map(s => [s.sheetId, session.layout(s.sheetId)])) })
  return { remote, session, port, store: remote.store('s') as RemoteSheetStore }
}
const box = { x: 0, y: 0, width: 256, height: 200 }
const windowOf = (b = box): LogicalViewport => ({ box: { ...b }, zoom: 1, deviceScale: 1 })
const extentOf = (metrics: ProjectedSheetMetrics) => ({ width: metrics.cols.totalCss(), height: metrics.rows.totalCss() })
const visibility = () => visibilityFromHidden(createIntervalSet([]))
const sizes = (cols: Array<[number, number]>, rows: Array<[number, number]> = []) => ({ cols: new Map(cols), rows: new Map(rows) })

describe('paged geometry with session sizes', () => {
  it('rebuilds the axes from the worker layout with the session sizes; rows follow the source row under sort', async () => {
    const h = await harness()
    try {
      const plain = pagedSheetGeometry(h.remote, 's', createIdentityOrder(), visibility())
      expect(plain.metrics.cols.size(1)).toBe(64)
      const order = createRowOrder([{ r0: 0, r1: 2, sourceAtOffset: Uint32Array.from([2, 1, 0]) }])
      const sized = pagedSheetGeometry(h.remote, 's', order, visibility(), undefined, sizes([[1, 200]], [[2, 45]]))
      expect(sized.metrics.cols.size(1)).toBe(200)
      expect(sized.metrics.cols.start(2)).toBe(64 + 200)
      // Source row 2 is displayed first under the sort: its session height moves with it.
      expect(sized.metrics.rows.size(0)).toBe(45)
      expect(sized.metrics.rows.size(2)).toBe(20)
      // The workbook axes the worker uses are untouched.
      expect(h.remote.geometry('s').cols.size(1)).toBe(64)
    } finally { await h.remote.release() }
  })
})

describe('review #2: paged geometry never throws for a legal session size', () => {
  it('falls back to the authored axes with a typed diagnostic when the overlay is unrepresentable', async () => {
    const h = await harness()
    try {
      const layout = h.remote.layout('s')
      Object.assign(layout, { rowOverrides: [{ start: 0, end: 1, sizeCss: 1e-15 }] })
      const geometry = pagedSheetGeometry(h.remote, 's', createIdentityOrder(), visibility(), undefined, sizes([], [[0, 8]]))
      expect(geometry.metrics.rows.size(0)).toBe(1e-15)
      expect(geometry.diagnostics.some(d => d.code === 'axis-precision-exceeded')).toBe(true)
    } finally { await h.remote.release() }
  })
})

describe('confirmation r4 R2: paged drawing copies are part of the candidate', () => {
  const E = 9525
  const fakeDrawings = (count: number) => ({ state: 'ready', gen: 1, images: [], diagnostics: [], bytes: 0,
    drawings: Array.from({ length: count }, (_, i) => ({ xEmu: (100 + i) * E, yEmu: 0, widthEmu: 64 * E, heightEmu: 40 * E, geometry: 'rect', source: { partPath: 'p', treePath: `twoCellAnchor[${i}]/sp[0]`, element: 'sp', representation: 'native' } })),
    drawingsIn: () => [] }) as never
  it('a deferred build places moved drawings eagerly and prices them in its charge', async () => {
    const h = await harness()
    try {
      vi.spyOn(h.remote, 'drawings').mockReturnValue(fakeDrawings(2048))
      const plain = pagedSheetGeometry(h.remote, 's', createIdentityOrder(), visibility(), undefined, undefined, true)
      const sized = pagedSheetGeometry(h.remote, 's', createIdentityOrder(), visibility(), undefined, sizes([[0, 100]]), true)
      expect(sized.sizedDrawings?.length).toBe(2048)
      expect(sized.drawingsIn(0, 0, 1e6, 1e6, 1e6)?.length).toBe(2048)
      expect(h.remote.mainLedger.snapshot().byPool.drawings).toBe(0)
      expect(plain.sizedDrawings).toBeUndefined()
    } finally { await h.remote.release() }
  })
  it('a later drawing list is placed lazily only if it fits the main budget (null: over budget)', async () => {
    const h = await harness()
    try {
      const content = fakeDrawings(4) as { drawings: unknown[] }
      vi.spyOn(h.remote, 'drawings').mockReturnValue(content as never)
      const geometry = pagedSheetGeometry(h.remote, 's', createIdentityOrder(), visibility(), undefined, sizes([[0, 100]]))
      // Drawing content changes in place when a newer list arrives.
      content.drawings = (fakeDrawings(2048) as { drawings: unknown[] }).drawings
      h.remote.mainLedger.set('copy', 'filler', 64 * 1024 * 1024 - h.remote.mainLedger.bytes - 4096)
      const drawingsBefore = h.remote.mainLedger.snapshot().byPool.drawings
      expect(geometry.drawingsIn(0, 0, 1e6, 1e6, 1e6)).toBeNull()
      expect(h.remote.mainLedger.snapshot().byPool.drawings).toBe(drawingsBefore)
      expect(h.remote.mainLedger.bytes).toBeLessThanOrEqual(64 * 1024 * 1024)
      h.remote.mainLedger.set('copy', 'filler', 0)
    } finally { await h.remote.release() }
  })
})

describe('confirmation r5 F1: late drawing placement reserves its placement workspace first', () => {
  it('a small late list over many authored row heights is refused before any placement axis is built', async () => {
    const h = await harness()
    try {
      Object.assign(h.remote.layout('s'), { rowOverrides: Array.from({ length: 2000 }, (_, i) => ({ start: 2 * i, end: 2 * i, sizeCss: 30 })) })
      const content = { state: 'ready', gen: 1, images: [], diagnostics: [], bytes: 0, drawings: [] as unknown[], drawingsIn: () => [] }
      vi.spyOn(h.remote, 'drawings').mockReturnValue(content as never)
      const geometry = pagedSheetGeometry(h.remote, 's', createIdentityOrder(), visibility(), undefined, sizes([[0, 100]]))
      const E = 9525
      content.drawings = [{ xEmu: 10 * E, yEmu: 10 * E, widthEmu: 100 * E, heightEmu: 50 * E, geometry: 'rect', source: { partPath: 'p', treePath: 'twoCellAnchor[0]/sp[0]', element: 'sp', representation: 'native' } }]
      h.remote.mainLedger.set('copy', 'filler', 64 * 1024 * 1024 - h.remote.mainLedger.bytes - 4096)
      const axes = vi.spyOn(axesModule, 'createAxisMetrics')
      expect(geometry.drawingsIn(0, 0, 1e6, 1e6, 1e6)).toBeNull()
      expect(axes).not.toHaveBeenCalled()
      expect(h.remote.mainLedger.snapshot().peakBytes).toBeLessThanOrEqual(64 * 1024 * 1024)
      h.remote.mainLedger.set('copy', 'filler', 0)
      expect(geometry.drawingsIn(0, 0, 1e6, 1e6, 1e6)).toHaveLength(1)
    } finally { await h.remote.release() }
  })
})

describe('worker plans while a sheet is resized', () => {
  const commands = (h: Awaited<ReturnType<typeof harness>>) => h.remote.stats.commands.viewport ?? 0

  it('a resized sheet never asks the worker: cold tiles plan on the main thread, oversized ones are coarse', async () => {
    const h = await harness()
    const geometry = pagedSheetGeometry(h.remote, 's', createIdentityOrder(), visibility(), undefined, sizes([[1, 200]]))
    const plans = new WorkerViewportPlans(h.remote)
    const source = tilePlanSource({ store: h.store, metrics: geometry.metrics, scale: 1, identity: h.session.identity('s'), extent: extentOf(geometry.metrics), plans, resized: true })
    try {
      expect(source(windowOf(), true)).toBeUndefined()
      const oversized = source(windowOf({ x: 0, y: 0, width: 256, height: 1_000_000 }), true)
      expect(oversized).toMatchObject({ kind: 'coarse' })
      expect(commands(h)).toBe(0)
      // Control: the same sheet without session sizes asks the worker for its cold tile.
      const plain = tilePlanSource({ store: h.store, metrics: geometry.metrics, scale: 1, identity: h.session.identity('s'), extent: extentOf(geometry.metrics), plans })
      plain(windowOf(), true)
      expect(commands(h)).toBe(1)
    } finally { plans.dispose(); await h.remote.release() }
  })

  it('a plan requested under an older layout generation is cancelled and never published', async () => {
    const h = await harness()
    const metrics = createSheetMetrics({ rows: h.remote.geometry('s').rows, cols: h.remote.geometry('s').cols, order: createIdentityOrder(), visibility: visibility() })
    const plans = new WorkerViewportPlans(h.remote)
    const owner = h.session.identity('s')
    const signals: AbortSignal[] = []
    const real = h.remote.viewport.bind(h.remote)
    const spy = vi.spyOn(h.remote, 'viewport').mockImplementation((...args) => {
      signals.push(args[2])
      return real(...args)
    })
    try {
      plans.beginFrame(owner, windowOf(), 0)
      expect(plans.get(owner, windowOf(), extentOf(metrics), true).kind).toBe('pending')
      expect(spy).toHaveBeenCalledTimes(1)
      // The user resizes: the next frame carries layout generation 1.
      plans.beginFrame(owner, windowOf(), 1)
      expect(signals[0].aborted).toBe(true)
      await h.port.idle()
      await new Promise(resolve => setTimeout(resolve, 0))
      // The old reply landed after the scope change: nothing was stored, a new request goes out.
      expect(plans.get(owner, windowOf(), extentOf(metrics), true).kind).toBe('pending')
      expect(spy).toHaveBeenCalledTimes(2)
      expect(signals[1].aborted).toBe(false)
    } finally { spy.mockRestore(); plans.dispose(); await h.remote.release() }
  })
})
