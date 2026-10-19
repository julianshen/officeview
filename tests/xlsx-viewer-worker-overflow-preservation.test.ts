// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { ViewerClient } from '../src/worker/viewerClient'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { WorkerViewportPlans, tilePlanSource } from '../src/components/xlsx/workerViewport'
import { pagedSheetGeometry } from '../src/components/xlsx/pagedGeometry'
import { SheetTileRenderer, type SheetFrame } from '../src/components/xlsx/sheetTiles'
import { buildSheetGeometry, MergeQueryCache, queryViewport } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { extendDisplayRange } from '../src/xlsx/view/selection'
import { rowSetIntervals } from '../src/worker/viewerSession'
import { serializeSelection } from '../src/xlsx/view/copy'
import { formatValue } from '../src/xlsx/render'
import type { GridSelection, LogicalViewport } from '../src/xlsx/view/contracts'
import { workbookOf } from './helpers/xlsx-find'

const gate = () => { let open!: () => void; const promise = new Promise<void>(resolve => { open = resolve }); return { promise, open } }
const signal = () => new AbortController().signal
const visible = () => visibilityFromHidden(createIntervalSet([]))
const windowOf = (height = 512): LogicalViewport => ({ box: { x: 0, y: 0, width: 512, height }, zoom: 1, deviceScale: 1 })
async function harness(doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 'alpha-0' }, { row: 0, col: 1, value: 'BLOCKER' }] }])) {
  const objectSource = createObjectWorkbookSource(doc, 'preservation')
  const styles = [{ numFmtId: 0 }, { numFmtId: 14 }]
  const source = { ...objectSource, formatting: () => ({ dateSystem: '1900' as const, styles }) }
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const port = connectInProcess(kernel), client = new ViewerClient(port)
  const opened = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'preservation', source: { kind: 'object', document: doc }, profile: 'D' })
  if (opened.kind !== 'complete' || opened.result.kind !== 'open') throw new Error('Expected actual open')
  const session = kernel.session('preservation')!
  const remote = new RemoteWorkbook(client, { docId: 'preservation', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: opened.result.formatting, layouts: new Map([['s', session.layout('s')]]) })
  const plans = new WorkerViewportPlans(remote), store = remote.store('s'), provider = source.store('s')
  const flush = async () => { await port.idle(); for (let i = 0; i < 30; i++) await Promise.resolve() }
  remote.merges('s'); remote.drawings('s'); await flush()
  const geometry = () => pagedSheetGeometry(remote, 's', remote.view('s').order, visible())
  const route = (w: LogicalViewport) => { const g = geometry(); return tilePlanSource({ store, metrics: g.metrics, scale: w.zoom * w.deviceScale, identity: remote.identity('s'), extent: { width: g.metrics.cols.totalCss(), height: g.metrics.rows.totalCss() }, plans }) }
  const close = async () => { plans.dispose(); await remote.close() }
  return { doc, source, kernel, session, remote, plans, store, provider, flush, geometry, route, close }
}
type Harness = Awaited<ReturnType<typeof harness>>
function frame(h: Harness, w = windowOf()): SheetFrame {
  const geometry = h.geometry()
  return { sheet: { ...h.doc.sheets[0], rows: [] }, geometry, identity: h.remote.identity('s'), origin: { x: w.box.x, y: w.box.y }, viewport: { width: 128, height: 40 }, zoom: w.zoom, dpr: w.deviceScale, extent: { width: geometry.metrics.cols.totalCss(), height: geometry.metrics.rows.totalCss() }, paletteKey: '', layoutCache: new LayoutCache(), merges: new MergeQueryCache(), plan: h.route(w), pending: (box, isVisible) => h.store.pendingIn(geometry.metrics, box, w.zoom * w.deviceScale, !isVisible) }
}
function renderer() {
  return new SheetTileRenderer('D', 0, () => { const canvas = createCanvas(512, 512); return { canvas: canvas as unknown as HTMLCanvasElement, width: 512, height: 512, close: () => { canvas.width = 0; canvas.height = 0 } } })
}
function crop(h: Harness, f: SheetFrame) {
  const actual = createCanvas(128 * f.dpr, 40 * f.dpr), expected = createCanvas(128 * f.dpr, 40 * f.dpr), a = renderer(), b = renderer()
  const canonical = buildSheetGeometry(h.doc.sheets[0], { order: h.remote.view('s').order })
  if (canonical.kind !== 'ready') throw new Error('Expected canonical geometry')
  try {
    a.present(actual.getContext('2d') as unknown as CanvasRenderingContext2D, f)
    b.present(expected.getContext('2d') as unknown as CanvasRenderingContext2D, { ...f, sheet: h.doc.sheets[0], geometry: canonical, plan: undefined, pending: undefined, paletteKey: '', layoutCache: new LayoutCache(), merges: new MergeQueryCache() })
    expect(actual.getContext('2d').getImageData(0, 0, actual.width, actual.height).data).toEqual(expected.getContext('2d').getImageData(0, 0, expected.width, expected.height).data)
    expect(a.stats.pendingTiles).toBe(0)
    expect(a.cache.stats().tiles).toBeGreaterThan(0)
  } finally { a.dispose(); b.dispose() }
}

it('TASK8 cache proof cancels held plan and preserves resident paint', async () => {
  const h = await harness(), held = gate(), entered = gate(), ended = gate(), original = h.provider.readRows.bind(h.provider)
  let first = true, backing: Uint8Array | undefined
  const reads = vi.spyOn(h.provider, 'readRows').mockImplementation(async (...args) => {
    const page = await original(...args)
    if (first) { first = false; backing = new Uint8Array(1024 * 1024); entered.open(); await held.promise; const output = { ...page, backing }; backing = undefined; ended.open(); return output }
    return page
  })
  const requests = vi.spyOn(h.remote, 'viewport'), releases = vi.spyOn(h.remote, 'releasePages')
  const w = windowOf(), geometry = h.geometry()
  try {
    h.plans.beginFrame(h.remote.identity('s'), w)
    expect(h.route(w)(w, true)).toBeUndefined()
    await entered.promise
    expect(h.remote.pagesInFlight).toBe(8)
    expect(requests.mock.calls[0][2].aborted).toBe(false)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
    // A cached known row paints through the real main fallback while the
    // same worker request is pending and the following row is still unknown.
    await h.store.load([0], 0, 255)
    expect(h.store.coverage(geometry.metrics, w.box, 1)).toBe('unknown')
    expect(h.route(w)(w, true)).toBeUndefined()
    const progressive = createCanvas(128, 40), reference = createCanvas(128, 40), paint = renderer(), canonicalPaint = renderer()
    const f = frame(h), canonical = buildSheetGeometry(h.doc.sheets[0])
    if (canonical.kind !== 'ready') throw new Error('Expected canonical geometry')
    try {
      paint.present(progressive.getContext('2d') as unknown as CanvasRenderingContext2D, f)
      canonicalPaint.present(reference.getContext('2d') as unknown as CanvasRenderingContext2D, { ...f, sheet: h.doc.sheets[0], geometry: canonical, plan: undefined, pending: undefined, paletteKey: '', layoutCache: new LayoutCache(), merges: new MergeQueryCache() })
      expect(progressive.getContext('2d').getImageData(0, 0, 128, 20).data).toEqual(reference.getContext('2d').getImageData(0, 0, 128, 20).data)
      expect([...progressive.getContext('2d').getImageData(2, 35, 1, 1).data]).toEqual([236, 239, 243, 255])
      expect(paint.stats.pendingTiles).toBe(1)
      expect(paint.cache.stats().tiles).toBe(0)
      const known = renderSheetViewport(f.sheet, createCanvas(512, 512).getContext('2d') as unknown as CanvasRenderingContext2D, f.geometry, w, undefined, { owner: f.identity })
      expect(known.records.find(record => record.id === 0)).toMatchObject({ source: { row: 0, col: 0 }, body: { formattedText: 'alpha-0' } })
      expect(requests).toHaveBeenCalledTimes(1)
    } finally { paint.dispose(); canonicalPaint.dispose() }
    const progressiveLayoutBytes = h.remote.mainLedger.snapshot().byPool.layout
    expect(progressiveLayoutBytes).toBeGreaterThan(0)
    // This ordinary read is independent of the still-held viewport producer.
    await h.store.load(Array.from({ length: 26 }, (_, row) => row), 0, 255)
    expect(h.store.coverage(geometry.metrics, w.box, 1)).toBe('proven')
    expect(h.store.cellAt({ row: 0, col: 0 })?.value).toBe('alpha-0')
    const afterLoad = releases.mock.calls.length
    expect(h.route(w)(w, true)).toBeUndefined()
    expect(requests.mock.calls[0][2].aborted).toBe(true)
    await h.flush()
    expect(h.remote.pagesInFlight).toBe(0)
    expect(releases.mock.calls.slice(afterLoad)).toEqual([[8]])
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(progressiveLayoutBytes)
    expect(backing?.byteLength).toBe(1024 * 1024)
    expect(h.session.memoryStats().transientBytes).toBe(16 * 1024 * 1024)
    crop(h, frame(h))
    const residentLayoutBytes = h.remote.mainLedger.snapshot().byPool.layout
    expect(requests).toHaveBeenCalledTimes(1)
    const read = await h.store.readCell({ row: 0, col: 0 }, { signal: signal(), stamp: h.store.stamp })
    expect(read).toMatchObject({ kind: 'cell', cell: { id: 0, value: 'alpha-0' } })
    held.open(); await ended.promise; await h.flush()
    expect(h.session.memoryStats().transientBytes).toBe(0)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(residentLayoutBytes)
    expect(releases.mock.calls.slice(afterLoad)).toEqual([[8]])
    expect(h.route(w)(w, true)).toBeUndefined()
    expect(requests).toHaveBeenCalledTimes(1)
    expect(h.store.cellAt({ row: 0, col: 0 })?.value).toBe('alpha-0')
  } finally { held.open(); await h.flush(); reads.mockRestore(); requests.mockRestore(); releases.mockRestore(); await h.close() }
})

it('TASK8 genuine oversized pending refusal and completion retain known alpha inspection and paint', async () => {
  const h = await harness(), held = gate(), entered = gate(), original = h.provider.queryOverview.bind(h.provider)
  await h.store.load([0], 0, 255)
  const reads = vi.spyOn(h.provider, 'queryOverview').mockImplementation(async (...args) => { entered.open(); await held.promise; return original(...args) })
  const w = windowOf(1_000_000), g = h.geometry()
  const inspect = async () => expect(await h.store.readCell({ row: 0, col: 0 }, { signal: signal(), stamp: h.store.stamp })).toMatchObject({ kind: 'cell', cell: { id: 0, value: 'alpha-0' } })
  try {
    expect(h.store.coverage(g.metrics, w.box, 1)).toBe('oversized')
    h.plans.beginFrame(h.remote.identity('s'), w)
    expect(h.route(w)(w, true)?.kind).toBe('pending')
    await entered.promise; await inspect()
    // Resource pressure is real aggregate admission, not a synthetic reply.
    h.remote.mainLedger.set('layout', 'task8-pressure', 64 * 1024 * 1024 - h.remote.mainLedger.bytes)
    held.open(); await h.flush()
    expect(h.route(w)(w, true)).toMatchObject({ kind: 'diagnostic', diagnostics: [{ code: 'resource-limit' }] })
    await inspect()
    h.remote.mainLedger.release('layout', 'task8-pressure')
    h.plans.release(h.remote.identity('s'), w); h.plans.beginFrame(h.remote.identity('s'), w)
    expect(h.route(w)(w, true)?.kind).toBe('pending'); await h.flush()
    expect(h.route(w)(w, true)?.kind).toBe('detail')
    await inspect()
    await h.store.load(Array.from({ length: 26 }, (_, row) => row), 0, 255)
    crop(h, frame(h))
    expect(h.store.isKnown({ row: 26, col: 0 })).toBe(false)
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { held.open(); reads.mockRestore(); await h.close() }
})

it('TASK8 real viewport states preserve exact active source and group membership', async () => {
  const doc = workbookOf([{ sheetId: 's', cells: [0, 1, 2, 3].map(row => ({ row, col: 0, value: ['header', 'one', 'excluded', 'three'][row] })) }])
  // Two adjacent display slots are below detail; their source rows are {3,1}.
  for (const row of doc.sheets[0].rows) if (row.index === 1 || row.index === 3) { row.customHeight = true; row.heightPt = 0.75 }
  const h = await harness(doc)
  const sorted = await h.remote.sort(h.remote.identity('s'), { range: { r0: 0, r1: 3, c0: 0, c1: 0, headerRows: 1, totalsRows: 0 }, keys: [{ col: 0, order: 'desc' }], caseSensitive: false }, signal())
  expect(sorted.kind).toBe('committed')
  const geometry = h.geometry(), owner = h.remote.identity('s')
  const overview = vi.spyOn(h.provider, 'queryOverview')
  const selection = extendDisplayRange(geometry.metrics, { row: 3, col: 0 }, { row: 1, col: 0 })
  const w: LogicalViewport = { box: { x: 0, y: 20, width: 64, height: 2 }, zoom: 1, deviceScale: 1 }, extent = { width: geometry.metrics.cols.totalCss(), height: geometry.metrics.rows.totalCss() }
  const exact = async () => {
    const out = await queryViewport(h.provider, geometry.metrics, w, owner, { signal: signal(), activeCell: { row: 3, col: 0 } })
    expect(out.kind).toBe('viewport')
    if (out.kind !== 'viewport') throw new Error('Expected exact inspector route')
    expect(out.active).toMatchObject({ kind: 'cell', source: { row: 3, col: 0 }, cell: { id: 49152, value: 'three' }, bounds: { y: 20, height: 1 } })
    return out
  }
  try {
    expect(geometry.metrics.order.displayToSource(1)).toBe(3)
    expect(geometry.metrics.order.displayToSource(2)).toBe(1)
    expect(rowSetIntervals(selection.ranges[0].rows)).toEqual([{ r0: 1, r1: 1 }, { r0: 3, r1: 3 }])
    h.plans.beginFrame(owner, w)
    expect(h.plans.get(owner, w, extent, true).kind).toBe('pending')
    await exact(); await h.flush()
    const plan = h.plans.get(owner, w, extent, true)
    expect(plan.kind).toBe('detail')
    if (plan.kind !== 'detail') throw new Error('Expected admitted coarse regions')
    const out = await exact()
    expect(plan.coarse).toEqual(out.coarse)
    const domains = overview.mock.calls.flatMap(([request]) => request.groups)
    expect(domains.length).toBeGreaterThan(0)
    for (const domain of domains) expect(rowSetIntervals(domain.rows)).toEqual([{ r0: 1, r1: 1 }, { r0: 3, r1: 3 }])
    expect(out.groups).toHaveLength(16)
    for (const group of out.groups) expect(group).toMatchObject({ populatedCount: 2, countKind: 'exact' })
    const copy = await serializeSelection({ store: h.store, owner, currentOwner: () => h.remote.identity('s'), selection: { active: { row: 3, col: 0 }, anchor: { row: 1, col: 0 }, ranges: selection.ranges }, order: geometry.metrics.order, visibility: visible(), columns: visibilityFromHidden(createIntervalSet([]), 16384), allowed: () => true, signal: signal() })
    expect(copy).toMatchObject({ kind: 'ready', text: 'three\none' })
    h.plans.release(owner, w); h.remote.mainLedger.set('layout', 'task8-pressure', 64 * 1024 * 1024 - h.remote.mainLedger.bytes)
    h.plans.beginFrame(owner, w); expect(h.plans.get(owner, w, extent, true).kind).toBe('pending'); await h.flush()
    expect(h.plans.get(owner, w, extent, true).kind).toBe('diagnostic')
    await exact()
    expect(rowSetIntervals(selection.ranges[0].rows)).toEqual([{ r0: 1, r1: 1 }, { r0: 3, r1: 3 }])
  } finally { overview.mockRestore(); h.remote.mainLedger.release('layout', 'task8-pressure'); await h.close() }
})

it('TASK8 refresh keeps controls current and retires old plans', async () => {
  const h = await harness(), requests = vi.spyOn(h.remote, 'viewport')
  const selection: GridSelection = { active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 }, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, cols: [{ c0: 0, c1: 0 }] }] }
  const initial = JSON.stringify(selection), w = windowOf()
  try {
    let owner = h.remote.identity('s'), g = h.geometry(), extent = { width: g.metrics.cols.totalCss(), height: g.metrics.rows.totalCss() }
    h.plans.beginFrame(owner, w); h.plans.get(owner, w, extent, true); await h.flush()
    expect(h.plans.get(owner, w, extent, true).kind).toBe('detail')
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBeGreaterThan(0)
    h.session.publishFonts()
    // Real committed-owner publication carries the new font identity to the host.
    expect((await h.remote.sort(h.session.identity('s'), null, signal())).kind).toBe('committed')
    owner = h.remote.identity('s'); expect(owner.fontGen).toBe(1)
    const refreshed = { ...w, zoom: 1.25, deviceScale: 2 }
    g = h.geometry(); extent = { width: g.metrics.cols.totalCss(), height: g.metrics.rows.totalCss() }
    h.plans.beginFrame(owner, refreshed)
    expect(requests.mock.calls[0][2].aborted).toBe(true)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
    expect(h.plans.get(owner, refreshed, extent, true).kind).toBe('pending'); await h.flush()
    expect(h.plans.get(owner, refreshed, extent, true).kind).toBe('detail')
    expect(requests).toHaveBeenCalledTimes(2)
    expect(requests.mock.calls[1][0]).toEqual(owner)
    expect(requests.mock.calls[1][1]).toEqual(refreshed)
    expect(JSON.stringify(selection)).toBe(initial)
    const exact = await h.store.readCell(selection.active!, { signal: signal(), stamp: h.store.stamp })
    expect(exact).toMatchObject({ kind: 'cell', cell: { id: 0, value: 'alpha-0' } })
    await h.store.load(Array.from({ length: 26 }, (_, row) => row), 0, 255)
    crop(h, frame(h, refreshed))
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { requests.mockRestore(); await h.close() }
})

it('TASK8 worker records preserve canonical numeric date and formula consumers', async () => {
  const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 3.14159265358979 }, { row: 0, col: 1, value: 45351 }, { row: 1, col: 0, value: 42 }] }])
  const date = doc.sheets[0].rows[0].cells[1]; date.styleIndex = 1; date.style = { numFmtId: 14 }
  const formula = doc.sheets[0].rows[1].cells[0]; formula.formula = '6*7'; formula.hasCachedValue = true
  const h = await harness(doc), w = windowOf(), g = h.geometry(), extent = { width: g.metrics.cols.totalCss(), height: g.metrics.rows.totalCss() }
  try {
    const owner = h.remote.identity('s')
    h.plans.beginFrame(owner, w); h.plans.get(owner, w, extent, true); await h.flush()
    const plan = h.plans.get(owner, w, extent, true)
    expect(plan.kind).toBe('detail')
    const a = createCanvas(128, 40), b = createCanvas(128, 40), geometry = buildSheetGeometry(doc.sheets[0])
    if (geometry.kind !== 'ready') throw new Error('Expected resident')
    const direct: LogicalViewport = { ...w, box: { x: 0, y: 0, width: 128, height: 40 } }
    const transported = renderSheetViewport(doc.sheets[0], a.getContext('2d') as unknown as CanvasRenderingContext2D, undefined, direct, undefined, { owner, preparedPlan: plan })
    const resident = renderSheetViewport(doc.sheets[0], b.getContext('2d') as unknown as CanvasRenderingContext2D, geometry, direct, undefined, { owner })
    expect(transported.records.map(r => [r.id, r.source, r.body.formattedText])).toEqual(resident.records.map(r => [r.id, r.source, r.body.formattedText]))
    expect(a.getContext('2d').getImageData(0, 0, 128, 40).data).toEqual(b.getContext('2d').getImageData(0, 0, 128, 40).data)
    for (const [row, col, value] of [[0, 0, 3.14159265358979], [0, 1, 45351], [1, 0, 42]]) expect(await h.store.readCell({ row, col }, { signal: signal(), stamp: h.store.stamp })).toMatchObject({ kind: 'cell', cell: { id: row * 16384 + col, value } })
    const copied = await serializeSelection({ store: h.store, owner, currentOwner: () => owner, selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 1 }] }, cols: [{ c0: 0, c1: 1 }] }] }, order: createIdentityOrder(), visibility: visible(), columns: visibilityFromHidden(createIntervalSet([]), 16384), allowed: () => true, signal: signal(), format: cell => formatValue(cell.value, h.remote.styleOf(cell.styleIndex)?.numFmtId ?? 0) })
    expect(copied).toMatchObject({ kind: 'ready', text: '3.141592654\t2/29/2024\n42\t' })
  } finally { await h.close() }
})
