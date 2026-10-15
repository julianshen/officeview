// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { ViewerKernel, connectInProcess, type InProcessPort } from '../src/worker/viewerKernel'
import { ViewerClient } from '../src/worker/viewerClient'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { WorkerViewportPlans, tilePlanSource } from '../src/components/xlsx/workerViewport'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { buildSheetGeometry, queryViewport, type ViewportPlan } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import type { LogicalViewport, CacheIdentity } from '../src/xlsx/view/contracts'
import type { OverflowViewportResult } from '../src/xlsx/view/viewportOverflow'
import { workbookOf } from './helpers/xlsx-find'

type Wire = { kind: string; owner: CacheIdentity; result: OverflowViewportResult }
const LONG = 'abcdefghijklmnopqrstuvwxABCDEFGHIJKLMNOP'
async function validationHarness(options: { at?: { row: number; col: number }; hidden?: boolean; tiny?: boolean; merge?: boolean; mergeAnchor?: boolean; twoRows?: boolean; rows?: number; profile?: 'D' | 'M' } = {}) {
  const at = options.at ?? { row: 0, col: 0 }
  const doc = workbookOf([{ sheetId: 's', cells: [{ ...at, value: LONG }, ...Array.from({ length: (options.rows ?? (options.twoRows ? 2 : 1)) - 1 }, (_, offset) => ({ row: at.row + offset + 1, col: at.col, value: LONG })), ...(options.mergeAnchor ? [{ row: at.row, col: at.col + 1, value: LONG }] : [])] }]), sheet = doc.sheets[0]
  sheet.cols = options.hidden ? [{ min: 0, max: 16383, hidden: true }] : []
  if (options.merge) sheet.mergeRanges = [{ minRow: at.row, maxRow: at.row, minCol: at.col + 1, maxCol: at.col + 2 }]
  const styles = [{ numFmtId: 0 }, { numFmtId: 0, fillColor: 'FFFF00' }]
  if (options.mergeAnchor) { const anchor = sheet.rows[0].cells[1]; anchor.styleIndex = 1; anchor.style = styles[1] }
  const object = createObjectWorkbookSource(doc, 'slice6')
  const baseGeometry = object.geometry('s')
  const baseColumns = options.tiny ? createAxisMetrics({ logicalCount: 16384, defaultSizeCss: 64, overrides: [{ start: 0, end: 16383, sizeCss: 1 }] }) : baseGeometry.cols
  const locateRows = vi.fn(baseGeometry.rows.locate.bind(baseGeometry.rows)), locateCols = vi.fn(baseColumns.locate.bind(baseColumns))
  const observedGeometry = { ...baseGeometry, rows: { ...baseGeometry.rows, locate: locateRows }, cols: { ...baseColumns, locate: locateCols }, ...(options.tiny ? { colOverrides: [{ start: 0, end: 16383, sizeCss: 1 }] } : {}) }
  const source = { ...object, ...(options.mergeAnchor ? { formatting: () => ({ dateSystem: '1900' as const, styles }) } : {}), geometry: (_id: string) => observedGeometry }
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const native = connectInProcess(kernel)
  let mutate: ((wire: Wire) => void) | undefined
  let originalPacket: unknown
  const wrappers = new Map<(event: { data: unknown }) => void, (event: { data: unknown }) => void>()
  const port: InProcessPort = {
    postMessage: (message, transfer) => native.postMessage(message, transfer), idle: () => native.idle(),
    addEventListener(type, listener) {
      const wrapper = (event: { data: unknown }) => {
        const wire = event.data as Wire
        if (wire.kind === 'complete' && wire.result?.kind === 'viewport' && mutate) { originalPacket = structuredClone(event.data); mutate(wire) }
        listener(event)
      }
      wrappers.set(listener, wrapper); native.addEventListener(type, wrapper)
    },
    removeEventListener(type, listener) { native.removeEventListener(type, wrappers.get(listener) ?? listener); wrappers.delete(listener) },
  }
  const client = new ViewerClient(port), profile = options.profile ?? 'D'
  const opened = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'slice6', source: { kind: 'object', document: doc }, profile })
  if (opened.kind !== 'complete' || opened.result.kind !== 'open') throw new Error('Expected actual open')
  const session = kernel.session('slice6')!, owner = session.identity('s')
  const remote = new RemoteWorkbook(client, { profile, docId: 'slice6', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: opened.result.formatting, layouts: new Map([['s', session.layout('s')]]) })
  const flush = async () => { await native.idle(); for (let i = 0; i < 30; i++) await Promise.resolve() }
  remote.merges('s'); await flush()
  const store = remote.store('s'), provider = source.store('s'), geometry = source.geometry('s')
  const actualRows = provider.readRows.bind(provider), actualStructure = provider.queryStructural.bind(provider)
  const reads = vi.spyOn(provider, 'readRows'), singles = vi.spyOn(provider, 'readCell'), structural = vi.spyOn(provider, 'queryStructural'), overview = vi.spyOn(provider, 'queryOverview')
  const actualViewport = remote.viewport.bind(remote), actualStage = store.stageViewport.bind(store)
  const requests = vi.spyOn(remote, 'viewport'), releases = vi.spyOn(remote, 'releasePages'), stage = vi.spyOn(store, 'stageViewport'), settled = vi.spyOn(remote.mainLedger, 'settle')
  const metrics = createSheetMetrics({ rows: geometry.rows, cols: geometry.cols })
  const extent = { width: metrics.cols.totalCss(), height: metrics.rows.totalCss() }
  const window: LogicalViewport = { box: { x: at.col * 64, y: at.row * 20, width: 192, height: 20 }, zoom: 1, deviceScale: 1 }
  const plans = new WorkerViewportPlans(remote)
  const get = tilePlanSource({ store: { coverage: () => 'oversized' } as unknown as typeof store, metrics, scale: 1, identity: owner, extent, plans })
  const fetch = async (w = window) => { plans.beginFrame(owner, w); expect(get(w, true)?.kind).toBe('pending'); await flush(); return get(w, true)! }
  const reply = async () => {
    const out: Awaited<ReturnType<RemoteWorkbook['viewport']>> = await requests.mock.results.at(-1)!.value
    if (typeof out !== 'object' || out.kind !== 'complete') throw new Error('Expected actual transported viewport')
    return out
  }
  const paint = (plan: ViewportPlan, w = window) => {
    const canvas = createCanvas(192, 20), ctx = canvas.getContext('2d')
    const rendered = renderSheetViewport(sheet, ctx as unknown as CanvasRenderingContext2D, undefined, w, undefined, { owner, preparedPlan: plan, cache: new LayoutCache() })
    return { rendered, pixels: ctx.getImageData(0, 0, 192, 20).data }
  }
  const canonical = (w = window) => {
    const built = buildSheetGeometry(sheet)
    if (built.kind !== 'ready') throw new Error('Expected canonical geometry')
    const canvas = createCanvas(192, 20), ctx = canvas.getContext('2d')
    const rendered = renderSheetViewport(sheet, ctx as unknown as CanvasRenderingContext2D, { ...built, metrics }, w, undefined, { owner, cache: new LayoutCache() })
    return { rendered, pixels: ctx.getImageData(0, 0, 192, 20).data }
  }
  const close = async () => { plans.dispose(); for (const spy of [locateRows, locateCols, reads, singles, structural, overview, requests, releases, stage, settled]) spy.mockRestore(); await remote.release() }
  return { at, doc, sheet, source, provider, actualRows, actualStructure, actualViewport, actualStage, client, remote, deliverOriginal: () => { for (const listener of wrappers.keys()) listener({ data: originalPacket }) }, store, plans, owner, window, extent, metrics, mutate: (fn?: (wire: Wire) => void) => { mutate = fn }, flush, fetch, reply, paint, canonical, close, reads, singles, structural, overview, requests, releases, stage, settled, locateRows, locateCols }
}

const malformed: Array<[string, (wire: Wire) => void]> = [
  ['missing-sidecar', wire => { delete wire.result.overflowProof }],
  ['missing-version', wire => { delete (wire.result.overflowProof as Partial<NonNullable<OverflowViewportResult['overflowProof']>>).version }],
  ...[NaN, Infinity, -1].flatMap(value => (['left', 'right'] as const).map(field => [`${field}-${String(value)}`, (wire: Wire) => { wire.result.overflowProof!.entries[0][field] = value }] as [string, (wire: Wire) => void])),
  ['fractional-row', wire => { const e = wire.result.overflowProof!.entries[0]; e.row = 0.5; e.id = e.row * 16384 + e.col }],
  ['negative-col', wire => { const e = wire.result.overflowProof!.entries[0]; e.col = -1; e.id = -1 }],
  ['row-overflow', wire => { const e = wire.result.overflowProof!.entries[0]; e.row = 1048576; e.id = e.row * 16384 }],
  ['col-overflow', wire => { const e = wire.result.overflowProof!.entries[0]; e.col = 16384; e.id = e.col }],
  ['fractional-id', wire => { wire.result.overflowProof!.entries[0].id += 0.5 }],
  ['id-overflow', wire => { wire.result.overflowProof!.entries[0].id = 2 ** 34 }],
  ['duplicate-entry', wire => { wire.result.overflowProof!.entries.push({ ...wire.result.overflowProof!.entries[0] }) }],
  ['missing-entry', wire => { wire.result.overflowProof!.entries.pop() }],
  ['extra-entry', wire => { const e = wire.result.overflowProof!.entries[0]; wire.result.overflowProof!.entries.push({ ...e, row: 1, id: 16384 }) }],
  ['raw-source-mismatch', wire => { if (wire.result.content.kind === 'detail' && wire.result.content.cells[0].kind === 'cell') wire.result.content.cells[0].cell.col = 1 }],
  ...(['docId', 'sheetId', 'valueGen', 'fontGen', 'viewGen'] as const).flatMap(field => (['reply', 'proof'] as const).map(kind => [`${kind}-owner-${field}`, (wire: Wire) => {
    const owner = kind === 'reply' ? wire.owner : wire.result.overflowProof!.owner
    if (field === 'docId' || field === 'sheetId') owner[field] += '-wrong'
    else owner[field]++
  }] as [string, (wire: Wire) => void])),
  ...(['zoom', 'deviceScale'] as const).map(field => [`wrong-${field}-same-box`, (wire: Wire) => { wire.result.overflowProof!.window[field] = 2 }] as [string, (wire: Wire) => void]),
  ...(['x', 'y', 'width', 'height'] as const).map(field => [`wrong-window-${field}`, (wire: Wire) => { wire.result.overflowProof!.window.box[field] += 1 }] as [string, (wire: Wire) => void]),
  ['complete-with-unresolved', wire => { wire.result.overflowProof!.pending.push({ bounds: { x: 64, y: 0, width: 64, height: 20 }, error: false }) }],
]
it.each(malformed)('SLICE6 malformed sidecar or internal reply guard %s diagnoses before source adoption', async (_name, mutation) => {
  const h = await validationHarness()
  // Reply owner controls operate after the client's own owner filter. The
  // provider and clone still run; no reply or settlement is fabricated.
  if (_name.startsWith('reply-owner-')) {
    const actualViewport = h.actualViewport
    h.requests.mockImplementation(async (...args) => {
      const out = await actualViewport(...args)
      if (typeof out === 'object' && out.kind === 'complete') mutation(out as Wire)
      return out
    })
  } else h.mutate(mutation)
  try {
    const plan = await h.fetch()
    expect(plan.kind).toBe('diagnostic')
    expect(h.stage).not.toHaveBeenCalled()
    expect(h.store.cachedRows).toBe(0)
    expect(h.store.isKnown(h.at)).toBe(false)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
    expect(h.requests).toHaveBeenCalledTimes(1)
    expect(h.remote.stats.commands.viewport).toBe(1)
    expect(h.releases).toHaveBeenCalledTimes(1)
    expect(h.releases).toHaveBeenCalledWith(8)
    expect(h.remote.pagesInFlight).toBe(0)
    if (plan.kind === 'diagnostic') expect(plan.diagnostics?.length).toBeGreaterThan(0)
    expect(h.settled).not.toHaveBeenCalledWith(await h.requests.mock.results[0].value)
  } finally { await h.close() }
})
it.each([{ row: 131072, col: 0, id: 2 ** 31 }, { row: 262144, col: 0, id: 2 ** 32 }, { row: 1048575, col: 16383, id: 2 ** 34 - 1 }])('SLICE6 arithmetic ID $id survives actual transport, adoption and canonical paint', async ({ row, col, id }) => {
  const h = await validationHarness({ at: { row, col } })
  try {
    const plan = await h.fetch(), reply = await h.reply()
    expect(plan.kind).toBe('detail')
    expect((reply.result as OverflowViewportResult).overflowProof!.entries.map(entry => entry.id)).toEqual([id])
    expect(h.store.cellAt({ row, col })?.id).toBe(id)
    expect(h.store.isKnown({ row, col })).toBe(false)
    if (plan.kind !== 'detail') throw new Error('Expected detail')
    expect(plan.cells.map(item => item.source)).toEqual([{ row, col }])
    expect(plan.pending).toBeUndefined()
    expect(h.paint(plan).rendered.records[0].source).toEqual({ row, col })
    expect(h.paint(plan).pixels).toEqual(h.canonical().pixels)
    expect(h.reads).toHaveBeenCalledTimes(1)
    expect(h.singles).not.toHaveBeenCalled()
    expect(h.releases).toHaveBeenCalledTimes(1)
  } finally { await h.close() }
})

it.each(['complete', 'loading', 'structure-loading'] as const)('SLICE6 requested dependency omission is certified only with completed reads and structure: %s', async completion => {
  const h = await validationHarness()
  const actualRows = h.actualRows, actualStructure = h.actualStructure
  // Bind the original methods before installing the per-provider controls.
  h.reads.mockImplementation(async (...args) => { const page = await actualRows(...args); return { ...page, completeness: completion === 'loading' ? 'loading' : 'complete' } })
  h.structural.mockImplementation(async (...args) => { const page = await actualStructure(...args); return { ...page, completeness: completion === 'structure-loading' ? 'loading' : 'complete' } })
  try {
    const plan = await h.fetch(), reply = await h.reply()
    expect(plan.kind).toBe('detail')
    if (plan.kind !== 'detail') throw new Error('Expected detail')
    const proof = (reply.result as OverflowViewportResult).overflowProof!
    expect(h.reads.mock.calls[0][0]).toEqual({ r0: 0, r1: 0, c0: 0, c1: 34 })
    expect(plan.cells[0].source).toEqual({ row: 0, col: 0 })
    expect(h.store.cellAt(h.at)?.value).toBe(LONG)
    expect(h.store.cellAt({ row: 0, col: 1 })).toBeUndefined()
    expect(h.store.isKnown({ row: 0, col: 1 })).toBe(false)
    expect(proof.entries[0].rightEnd).toBe(completion === 'complete' ? 'paint-boundary' : 'unknown')
    expect(plan.cells[0].overflow?.right).toBe(completion === 'complete' ? 2176 : 0)
    expect(plan.pending).toEqual(completion === 'complete' ? undefined : [
      { bounds: { x: 64, y: 0, width: 128, height: 20 }, error: false },
      ...(completion === 'structure-loading' ? [{ bounds: { x: 0, y: 0, width: 64, height: 20 }, error: false }] : []),
    ])
    const painted = h.paint(plan)
    expect(Boolean(painted.rendered.pending)).toBe(completion !== 'complete')
    if (completion === 'complete') expect(painted.pixels).toEqual(h.canonical().pixels)
    expect(h.reads).toHaveBeenCalledTimes(1)
    expect(h.singles).not.toHaveBeenCalled()
    expect(h.releases).toHaveBeenCalledTimes(1)
  } finally { await h.close() }
})

it.each(['zero-width', 'zero-height', 'all-hidden', 'past-end', 'all-coarse'] as const)('SLICE6 %s empty evidence adds no occupancy reads or endpoint locates', async shape => {
  const h = await validationHarness({ hidden: shape === 'all-hidden', tiny: shape === 'all-coarse' })
  const window = structuredClone(h.window)
  if (shape === 'zero-width') window.box.width = 0
  if (shape === 'zero-height') window.box.height = 0
  if (shape === 'past-end') window.box.x = h.extent.width + 64
  try {
    if (shape === 'all-coarse') {
      const withoutCapture = await queryViewport(h.provider, h.metrics, window, h.owner, { profile: 'D', scope: 'tile', signal: new AbortController().signal, textDependencies: true })
      expect(withoutCapture.kind).toBe('viewport')
      if (withoutCapture.kind !== 'viewport') throw new Error('Expected canonical coarse query')
      expect(withoutCapture.coarse).toHaveLength(240)
      expect(h.locateCols).toHaveBeenCalledTimes(480)
      expect(h.reads).not.toHaveBeenCalled()
      expect(h.singles).not.toHaveBeenCalled()
      h.locateRows.mockClear(); h.locateCols.mockClear(); h.overview.mockClear()
    }
    const plan = await h.fetch(window), reply = await h.reply()
    expect(plan.kind).toBe('detail')
    if (plan.kind !== 'detail') throw new Error('Expected empty detail adapter')
    const proof = (reply.result as OverflowViewportResult).overflowProof!
    expect(proof).toMatchObject({ version: 1, owner: h.owner, window, entries: [], pending: [] })
    expect(plan.cells).toEqual([])
    expect(plan.pending).toBeUndefined()
    expect(plan.coarse).toHaveLength(shape === 'all-coarse' ? 240 : 0)
    expect(plan.work.pages).toBe(0)
    expect(plan.work.overviewGroups).toBe(shape === 'all-coarse' ? 240 : 0)
    expect(h.reads).not.toHaveBeenCalled()
    expect(h.singles).not.toHaveBeenCalled()
    expect(h.structural).not.toHaveBeenCalled()
    expect(h.locateRows).not.toHaveBeenCalled()
    expect(h.locateCols).toHaveBeenCalledTimes(shape === 'all-coarse' ? 480 : 0)
    expect(h.store.cachedRows).toBe(0)
    expect(h.remote.stats.commands.viewport).toBe(1)
    expect(h.overview.mock.calls.length).toBe(shape === 'all-coarse' ? 1 : 0)
    expect(h.releases).toHaveBeenCalledTimes(1)
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { await h.close() }
})
it.each([['width', -1], ['height', -1], ['width', NaN], ['height', Infinity], ['x', NaN], ['y', Infinity]] as const)('SLICE6 invalid geometry %s=%s diagnoses before any occupancy read', async (field, value) => {
  const h = await validationHarness(), window = structuredClone(h.window)
  window.box[field] = value
  try {
    const plan = await h.fetch(window)
    expect(plan.kind).toBe('diagnostic')
    expect(h.stage).not.toHaveBeenCalled()
    expect(h.reads).not.toHaveBeenCalled()
    expect(h.singles).not.toHaveBeenCalled()
    expect(h.store.cachedRows).toBe(0)
    expect(h.remote.pagesInFlight).toBe(0)
    expect(h.releases).toHaveBeenCalledTimes(1)
  } finally { await h.close() }
})

it.each(malformed.filter(([name]) => name.startsWith('reply-owner-')))('SLICE6 client rejects foreign outer %s then accepts the original cloned reply and releases credits once', async (_name, mutation) => {
  const h = await validationHarness()
  h.mutate(mutation)
  try {
    const plan = await h.fetch()
    expect(plan.kind).toBe('pending')
    expect(h.client.rejectedReplies).toBe(1)
    expect(h.stage).not.toHaveBeenCalled()
    expect(h.store.cachedRows).toBe(0)
    expect(h.remote.pagesInFlight).toBe(8)
    expect(h.releases).not.toHaveBeenCalled()
    h.deliverOriginal()
    await h.flush()
    const recovered = h.plans.get(h.owner, h.window, h.extent, true)
    expect(recovered.kind).toBe('detail')
    expect(h.stage).toHaveBeenCalledTimes(1)
    expect(h.store.cellAt(h.at)?.value).toBe(LONG)
    expect(h.store.isKnown({ row: 0, col: 1 })).toBe(false)
    expect(h.remote.pagesInFlight).toBe(0)
    expect(h.releases).toHaveBeenCalledTimes(1)
    expect(h.releases).toHaveBeenCalledWith(8)
    expect(h.remote.stats.commands.viewport).toBe(1)
    expect(h.paint(recovered).pixels).toEqual(h.canonical().pixels)
  } finally { await h.close() }
})

it.each(['epoch', 'merge'] as const)('SLICE6 actual %s change while a provider continuation is held cannot adopt or settle another consumer', async change => {
  const h = await validationHarness()
  let enter!: () => void, release!: () => void
  const entered = new Promise<void>(resolve => { enter = resolve }), held = new Promise<void>(resolve => { release = resolve })
  const other = await h.actualViewport(h.owner, { ...h.window, box: { ...h.window.box, y: 40 } }, new AbortController().signal, 'tile')
  h.reads.mockImplementation(async (...args) => { const page = await h.actualRows(...args); enter(); await held; return page })
  try {
    h.plans.beginFrame(h.owner, h.window)
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('pending')
    await entered
    expect(h.remote.pagesInFlight).toBe(8)
    if (change === 'epoch') h.store.clear()
    else h.remote.merges('s').gen++
    release(); await h.flush()
    const incoming = await h.requests.mock.results[0].value
    expect(h.store.cachedRows).toBe(0)
    expect.soft(h.settled).not.toHaveBeenCalledWith(incoming)
    expect(h.settled).not.toHaveBeenCalledWith(other)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
    expect(h.remote.pagesInFlight).toBe(0)
    expect(h.releases).toHaveBeenCalledTimes(1)
    h.reads.mockImplementation(h.actualRows)
    h.plans.release(h.owner, h.window)
    const retry = await h.fetch()
    expect(retry.kind).toBe('detail')
    expect(h.store.cellAt(h.at)?.value).toBe(LONG)
    expect(h.paint(retry).pixels).toEqual(h.canonical().pixels)
    expect(h.settled).not.toHaveBeenCalledWith(other)
    expect(h.releases).toHaveBeenCalledTimes(2)
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { release(); await h.close() }
})

it.each(['epoch', 'merge'] as const)('SLICE6 actual %s change after admission or staging rejects before publication', async change => {
  const h = await validationHarness()
  const originalStage = h.actualStage
  let staleToken: ReturnType<typeof originalStage>
  h.stage.mockImplementation((...args) => {
    staleToken = originalStage(...args)
    if (change === 'epoch') h.store.clear()
    else h.remote.merges('s').gen++
    return staleToken
  })
  try {
    h.plans.beginFrame(h.owner, h.window)
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('pending')
    await h.flush()
    const incoming = await h.requests.mock.results[0].value
    expect(h.stage).toHaveBeenCalledTimes(1)
    expect(h.store.cachedRows).toBe(0)
    expect.soft(h.settled).not.toHaveBeenCalledWith(incoming)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
    expect(h.remote.ledger.bytes).toBe(0)
    expect(h.releases).toHaveBeenCalledTimes(1)
    expect(staleToken?.commit(() => { throw new Error('stale token revived') })).toBe(false)
    h.stage.mockImplementation(originalStage)
    h.plans.release(h.owner, h.window)
    const retry = await h.fetch()
    expect(retry.kind).toBe('detail')
    expect(h.releases).toHaveBeenCalledTimes(2)
  } finally { await h.close() }
})

it.each(['layout', 'rowCache'].flatMap(pool => (['epoch', 'merge'] as const).map(change => ({ pool, change }))))('SLICE6 $change changed after actual $pool retained set rolls back source and plan publication', async ({ pool: poolToChange, change }) => {
  const h = await validationHarness()
  const first = await h.fetch(), oldRows = h.store.cellsOf(0), before = h.remote.mainLedger.snapshot(), oldRowBytes = h.remote.ledger.bytes
  expect(first.kind).toBe('detail')
  const nextWindow = { ...h.window, box: { ...h.window.box, x: 64, width: 128 } }
  const other = await h.actualViewport(h.owner, { ...h.window, box: { ...h.window.box, y: 40 } }, new AbortController().signal, 'tile')
  const actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
  let injected = false
  const charge = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (!injected && pool === poolToChange && bytes > 0 && (pool === 'rowCache' || owner.startsWith('viewport-plans:') && !owner.includes(':stage:'))) {
      injected = true
      if (change === 'epoch') h.store.clear()
      else h.remote.merges('s').gen++
    }
  })
  try {
    h.plans.beginFrame(h.owner, nextWindow)
    expect(h.plans.get(h.owner, nextWindow, h.extent, true).kind).toBe('pending')
    await h.flush()
    const incoming = await h.requests.mock.results.at(-1)!.value
    expect(injected).toBe(true)
    expect.soft(h.store.cellsOf(0)).toBe(change === 'epoch' ? undefined : oldRows)
    expect.soft(h.remote.ledger.bytes).toBe(change === 'epoch' ? 0 : oldRowBytes)
    expect.soft(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(change === 'epoch' ? 0 : before.byPool.rowCache)
    expect.soft(h.remote.mainLedger.snapshot().byPool.layout).toBe(before.byPool.layout)
    expect.soft(h.settled).not.toHaveBeenCalledWith(incoming)
    expect(h.settled).not.toHaveBeenCalledWith(other)
    expect(h.releases).toHaveBeenCalledTimes(2)
    expect(h.remote.pagesInFlight).toBe(0)
    charge.mockRestore()
    h.plans.release(h.owner, nextWindow)
    const retry = await h.fetch(nextWindow)
    expect(retry.kind).toBe('detail')
    expect(h.paint(retry, nextWindow).pixels).toEqual(h.canonical(nextWindow).pixels)
    expect(h.releases).toHaveBeenCalledTimes(3)
    expect(h.settled).not.toHaveBeenCalledWith(other)
  } finally { charge.mockRestore(); await h.close() }
})

it('SLICE6 late completed structural merge replaces pending with the correct blocker and full merged anchor paint', async () => {
  const h = await validationHarness({ merge: true, mergeAnchor: true })
  let complete = false
  h.structural.mockImplementation(async (...args) => { const page = await h.actualStructure(...args); return complete ? page : { ...page, merges: [], completeness: 'loading' } })
  vi.useFakeTimers()
  try {
    const partial = await h.fetch()
    expect(partial.kind).toBe('detail')
    if (partial.kind !== 'detail') throw new Error('Expected partial detail')
    expect(partial.cells.map(cell => cell.source)).toEqual([{ row: 0, col: 0 }, { row: 0, col: 1 }])
    expect(partial.pending).toEqual([
      { bounds: { x: 0, y: 0, width: 64, height: 20 }, error: false },
      { bounds: { x: 128, y: 0, width: 64, height: 20 }, error: false },
      { bounds: { x: 64, y: 0, width: 64, height: 20 }, error: false },
    ])
    expect(h.store.cellAt({ row: 0, col: 0 })?.value).toBe(LONG)
    expect(h.store.cellAt({ row: 0, col: 1 })?.value).toBe(LONG)
    expect(h.paint(partial).rendered.pending).toBe(true)
    expect(h.requests).toHaveBeenCalledTimes(1)
    await vi.advanceTimersByTimeAsync(99)
    expect(h.requests).toHaveBeenCalledTimes(1)
    await vi.advanceTimersByTimeAsync(1)
    complete = true
    const final = await h.fetch()
    expect(final.kind).toBe('detail')
    if (final.kind !== 'detail') throw new Error('Expected completed detail')
    expect(final.pending).toBeUndefined()
    expect(final.cells).toMatchObject([
      { source: { row: 0, col: 0 }, x: 0, y: 0, width: 64, height: 20, overflow: { left: 0, right: 0 } },
      { source: { row: 0, col: 1 }, x: 64, y: 0, width: 128, height: 20, overflow: { left: 0, right: 0 }, merge: { minRow: 0, maxRow: 0, minCol: 1, maxCol: 2 } },
    ])
    expect(final.merges).toEqual([{ range: { minRow: 0, maxRow: 0, minCol: 1, maxCol: 2 }, x: 64, y: 0, width: 128, height: 20 }])
    const painted = h.paint(final)
    expect(painted.pixels).toEqual(h.canonical().pixels)
    expect([...painted.pixels.slice((1 * 192 + 100) * 4, (1 * 192 + 100) * 4 + 4)]).toEqual([255, 255, 0, 255])
    expect(h.structural).toHaveBeenCalledTimes(2)
    expect(h.reads).toHaveBeenCalledTimes(2)
    expect(h.singles).not.toHaveBeenCalled()
    expect(h.releases).toHaveBeenCalledTimes(2)
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { vi.useRealTimers(); await h.close() }
})
it('SLICE6 returned source cell outside the requested dependency domain never becomes a paint origin or adopted blank proof', async () => {
  const h = await validationHarness()
  h.reads.mockImplementation(async (...args) => {
    const page = await h.actualRows(...args), cell = page.items[0].cells[0]
    return { ...page, items: [{ row: 0, cells: [{ ...cell, id: 100, col: 100 }] }] }
  })
  try {
    const plan = await h.fetch(), reply = await h.reply()
    expect(plan.kind).toBe('detail')
    if (plan.kind !== 'detail') throw new Error('Expected empty detail')
    expect(h.reads.mock.calls[0][0]).toEqual({ r0: 0, r1: 0, c0: 0, c1: 34 })
    expect(plan.cells).toEqual([])
    expect((reply.result as OverflowViewportResult).overflowProof!.entries).toEqual([])
    expect(h.store.cachedRows).toBe(0)
    expect(h.store.cellAt({ row: 0, col: 100 })).toBeUndefined()
    expect(h.store.isKnown({ row: 0, col: 100 })).toBe(false)
    expect(h.paint(plan).rendered.records).toEqual([])
    expect(h.reads).toHaveBeenCalledTimes(1)
    expect(h.singles).not.toHaveBeenCalled()
  } finally { await h.close() }
})

it.each(['layout', 'rowCache'].flatMap(pool => (['owner', 'plans-close', 'document-close', 'scope-change'] as const).map(change => ({ pool, change }))))('SLICE6 $change after retained $pool set cannot revive closed or stale publication owners', async ({ pool: targetPool, change }) => {
  const h = await validationHarness()
  await h.fetch()
  const oldRows = h.store.cellsOf(0), before = h.remote.mainLedger.snapshot()
  const next = { ...h.window, box: { ...h.window.box, x: 64, width: 128 } }
  const actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
  let injected = false, documentClose: Promise<void> | undefined
  const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (injected || pool !== targetPool || bytes <= 0 || pool === 'layout' && (!owner.startsWith('viewport-plans:') || owner.includes(':stage:'))) return
    injected = true
    if (change === 'owner') h.remote.identity('s').valueGen++
    if (change === 'plans-close') h.plans.dispose()
    if (change === 'document-close') documentClose = h.remote.close()
    if (change === 'scope-change') h.plans.beginFrame(h.owner, { ...h.window, zoom: 2 })
  })
  try {
    h.plans.beginFrame(h.owner, next)
    expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('pending')
    await h.flush()
    if (documentClose) await documentClose
    const incoming = await h.requests.mock.results.at(-1)!.value
    expect(injected).toBe(true)
    expect.soft(h.settled).not.toHaveBeenCalledWith(incoming)
    expect.soft(h.store.cellsOf(0)).toBe(change === 'document-close' ? undefined : oldRows)
    expect.soft(h.remote.mainLedger.snapshot().byPool.layout).toBe(change === 'owner' ? before.byPool.layout : 0)
    expect.soft(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(change === 'document-close' ? 0 : before.byPool.rowCache)
    expect(h.releases).toHaveBeenCalledTimes(2)
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { set.mockRestore(); await h.close() }
})

it.each(['epoch', 'merge'] as const)('SLICE6 %s change after actual joint admission cannot reach retained publication and retries cleanly', async change => {
  const h = await validationHarness(), actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
  let injected = false
  const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (!injected && pool === 'layout' && owner.includes(':stage:') && !owner.endsWith(':measure') && bytes > 0) {
      injected = true
      if (change === 'epoch') h.store.clear()
      else h.remote.merges('s').gen++
    }
  })
  try {
    h.plans.beginFrame(h.owner, h.window)
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('pending')
    await h.flush()
    const incoming = await h.requests.mock.results[0].value
    expect(injected).toBe(true)
    expect(h.store.cachedRows).toBe(0)
    expect(h.settled).not.toHaveBeenCalledWith(incoming)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
    expect(h.remote.ledger.bytes).toBe(0)
    expect(h.remote.pagesInFlight).toBe(0)
    expect(h.releases).toHaveBeenCalledTimes(1)
    set.mockRestore()
    h.plans.release(h.owner, h.window)
    expect((await h.fetch()).kind).toBe('detail')
    expect(h.releases).toHaveBeenCalledTimes(2)
  } finally { set.mockRestore(); await h.close() }
})

// Independent SPEC counterexample: clear/close after the first of two actual row charges.
it.each(['epoch', 'document-close', 'none'] as const)('SPEC multirow %s after first retained row charge leaves zero row accounting', async change => {
  const h = await validationHarness({ twoRows: true })
  const window = { ...h.window, box: { ...h.window.box, height: 40 } }
  const actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
  let injected = false, closing: Promise<void> | undefined
  const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (change !== 'none' && !injected && pool === 'rowCache' && bytes > 0) {
      injected = true
      if (change === 'epoch') h.store.clear()
      else closing = h.remote.close()
    }
  })
  try {
    h.plans.beginFrame(h.owner, window)
    expect(h.plans.get(h.owner, window, h.extent, true).kind).toBe('pending')
    await h.flush()
    if (closing) await closing
    const reply = await h.reply()
    if (reply.result.content.kind !== 'detail') throw new Error('Expected two actual source cells')
    expect(reply.result.content.cells.filter(c => c.kind === 'cell').map(c => c.source)).toEqual([{ row: 0, col: 0 }, { row: 1, col: 0 }])
    if (change === 'none') {
      expect(h.store.cachedRows).toBe(2)
      expect(h.remote.ledger.bytes).toBe(480)
      expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(480)
      expect(h.settled).toHaveBeenCalledWith(await h.requests.mock.results[0].value)
      return
    }
    expect(injected).toBe(true)
    expect(h.store.cachedRows).toBe(0)
    expect.soft(h.remote.ledger.bytes).toBe(0)
    expect.soft(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(0)
    expect(h.settled).not.toHaveBeenCalledWith(await h.requests.mock.results[0].value)
    expect(h.releases).toHaveBeenCalledTimes(1)
  } finally { set.mockRestore(); await h.close() }
})

const multirowCases = ([2, 3] as const).flatMap(rows => Array.from({ length: rows }, (_, index) => index + 1).flatMap(position => (['D', 'M'] as const).flatMap(profile => (['epoch', 'document-close'] as const).map(change => ({ rows, position, profile, change })))))
it.each(multirowCases)('S6-SPEC-1 $profile $rows rows $change at retained charge $position stops before the next row', async ({ rows, position, profile, change }) => {
  const h = await validationHarness({ rows, profile }), window = { ...h.window, box: { ...h.window.box, height: rows * 20 } }
  const other = await h.actualViewport(h.owner, { ...h.window, box: { ...h.window.box, y: 100 } }, new AbortController().signal, 'tile')
  const actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger), actualCharge = h.remote.ledger.charge.bind(h.remote.ledger)
  const charges = vi.spyOn(h.remote.ledger, 'charge').mockImplementation(actualCharge)
  let retainedCalls = 0, injected = false, closing: Promise<void> | undefined
  const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (pool !== 'rowCache' || bytes <= 0 || injected) return
    retainedCalls++
    if (retainedCalls !== position) return
    injected = true
    if (change === 'epoch') h.store.clear()
    else closing = h.remote.close()
  })
  try {
    h.plans.beginFrame(h.owner, window)
    expect(h.plans.get(h.owner, window, h.extent, true).kind).toBe('pending')
    await h.flush()
    if (closing) await closing
    const reply = await h.reply()
    if (reply.result.content.kind !== 'detail') throw new Error('Expected actual multirow transport')
    expect(reply.result.content.cells.map(cell => cell.source)).toEqual(Array.from({ length: rows }, (_, row) => ({ row, col: 0 })))
    expect(h.remote.profile).toBe(profile)
    expect((reply.result as OverflowViewportResult).overflowProof!.profile).toBe(profile)
    expect(injected).toBe(true)
    expect.soft(charges).toHaveBeenCalledTimes(position)
    expect.soft(charges.mock.calls.map(([, row]) => row)).toEqual(Array.from({ length: position }, (_, row) => row))
    expect(h.store.cachedRows).toBe(0)
    expect.soft(h.remote.ledger.bytes).toBe(0)
    expect.soft(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(0)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
    expect(h.settled).not.toHaveBeenCalledWith(reply)
    expect(h.settled).not.toHaveBeenCalledWith(other)
    expect(h.releases).toHaveBeenCalledTimes(1)
    expect(h.releases).toHaveBeenCalledWith(profile === 'D' ? 8 : 4)
    expect(h.remote.pagesInFlight).toBe(0)
    const token = h.stage.mock.results[0].value as ReturnType<typeof h.actualStage>
    expect(token).toBeDefined()
    expect(token!.commit(() => { throw new Error('closed multirow stage revived') })).toBe(false)
    for (let row = 0; row < rows; row++) expect(h.store.cellAt({ row, col: 0 })).toBeUndefined()
    set.mockRestore(); charges.mockRestore()
    if (change === 'epoch') {
      h.plans.release(h.owner, window)
      const retry = await h.fetch(window)
      expect(retry.kind).toBe('detail')
      expect(h.store.cachedRows).toBe(rows)
      expect(h.remote.ledger.bytes).toBe(rows * 240)
      expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(rows * 240)
      expect(h.releases).toHaveBeenCalledTimes(2)
      expect(h.remote.pagesInFlight).toBe(0)
      expect(h.settled).not.toHaveBeenCalledWith(other)
    } else {
      expect(h.remote.isClosed).toBe(true)
      expect(h.plans.get(h.owner, window, h.extent, true).kind).toBe('pending')
      expect(h.requests).toHaveBeenCalledTimes(1)
      expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
    }
  } finally { set.mockRestore(); charges.mockRestore(); await h.close() }
})
it.each(([2, 3] as const).flatMap(rows => (['D', 'M'] as const).map(profile => ({ rows, profile }))))('S6-SPEC-1 $profile $rows rows without invalidation publishes each source row and settles once', async ({ rows, profile }) => {
  const h = await validationHarness({ rows, profile }), window = { ...h.window, box: { ...h.window.box, height: rows * 20 } }
  const charges = vi.spyOn(h.remote.ledger, 'charge')
  try {
    const plan = await h.fetch(window), reply = await h.reply()
    expect(plan.kind).toBe('detail')
    if (plan.kind !== 'detail') throw new Error('Expected actual multirow plan')
    expect(plan.pending).toBeUndefined()
    expect(plan.cells.map(cell => ({ source: cell.source, y: cell.y }))).toEqual(Array.from({ length: rows }, (_, row) => ({ source: { row, col: 0 }, y: row * 20 })))
    expect(h.store.cachedRows).toBe(rows)
    expect(h.remote.ledger.bytes).toBe(rows * 240)
    expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(rows * 240)
    expect(charges).toHaveBeenCalledTimes(rows)
    expect(h.settled.mock.calls.filter(([out]) => out === reply)).toHaveLength(1)
    expect(h.releases).toHaveBeenCalledTimes(1)
    expect(h.releases).toHaveBeenCalledWith(profile === 'D' ? 8 : 4)
    expect(h.remote.pagesInFlight).toBe(0)
    expect(h.reads).toHaveBeenCalledTimes(1)
    expect(h.singles).not.toHaveBeenCalled()
  } finally { charges.mockRestore(); await h.close() }
})

it.each(([2, 3] as const).flatMap(rows => Array.from({ length: rows }, (_, index) => index + 1).flatMap(position => (['D', 'M'] as const).map(profile => ({ rows, position, profile })))))('S6-SPEC-1 $profile $rows rows same-generation fault at retained charge $position restores prior rows and accounting', async ({ rows, position, profile }) => {
  const h = await validationHarness({ rows, profile }), window = { ...h.window, box: { ...h.window.box, height: rows * 20 } }
  const prior = await h.fetch(window), priorRows = Array.from({ length: rows }, (_, row) => h.store.cellsOf(row)), before = h.remote.mainLedger.snapshot()
  const other = await h.actualViewport(h.owner, { ...h.window, box: { ...h.window.box, y: 100 } }, new AbortController().signal, 'tile')
  const next = { ...window, box: { ...window.box, x: 64, width: 128 } }, actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
  const charges = vi.spyOn(h.remote.ledger, 'charge')
  let retainedCalls = 0, injected = false
  const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (pool === 'rowCache' && bytes > 0 && !injected && ++retainedCalls === position) { injected = true; throw new Error('multirow retained fault') }
  })
  try {
    h.plans.beginFrame(h.owner, next)
    expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('pending')
    await h.flush()
    const incoming = await h.reply()
    expect(injected).toBe(true)
    expect(charges).toHaveBeenCalledTimes(position)
    expect(h.store.cachedRows).toBe(rows)
    for (let row = 0; row < rows; row++) expect(h.store.cellsOf(row)).toBe(priorRows[row])
    expect(h.remote.ledger.bytes).toBe(rows * 240)
    expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(before.byPool.rowCache)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(before.byPool.layout)
    expect(h.plans.get(h.owner, window, h.extent, true)).toBe(prior)
    expect(h.plans.get(h.owner, next, h.extent, true)).toMatchObject({ kind: 'diagnostic', diagnostics: [{ message: 'multirow retained fault' }] })
    expect(h.settled).not.toHaveBeenCalledWith(incoming)
    expect(h.settled).not.toHaveBeenCalledWith(other)
    expect(h.releases).toHaveBeenCalledTimes(2)
    expect(h.remote.pagesInFlight).toBe(0)
    set.mockRestore(); charges.mockRestore()
    h.plans.release(h.owner, next)
    expect((await h.fetch(next)).kind).toBe('detail')
    expect(h.remote.ledger.bytes).toBe(rows * 240)
    expect(h.releases).toHaveBeenCalledTimes(3)
    expect(h.settled).not.toHaveBeenCalledWith(other)
  } finally { set.mockRestore(); charges.mockRestore(); await h.close() }
})
const multirowFields = ['docId', 'sheetId', 'valueGen', 'fontGen', 'viewGen', 'merge'] as const
it.each(multirowFields)('S6-SPEC-1 full guard %s changes at middle retained row charge before the third row', async field => {
  const h = await validationHarness({ rows: 3 }), window = { ...h.window, box: { ...h.window.box, height: 60 } }
  const prior = await h.fetch(window), oldRows = [h.store.cellsOf(0), h.store.cellsOf(1), h.store.cellsOf(2)], before = h.remote.mainLedger.snapshot()
  const next = { ...window, box: { ...window.box, x: 64, width: 128 } }, actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
  const charges = vi.spyOn(h.remote.ledger, 'charge')
  let retainedCalls = 0, injected = false
  const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (pool !== 'rowCache' || bytes <= 0 || injected || ++retainedCalls !== 2) return
    injected = true
    if (field === 'merge') h.remote.merges('s').gen++
    else if (field === 'docId' || field === 'sheetId') h.remote.identity('s')[field] += '-changed'
    else h.remote.identity('s')[field]++
  })
  try {
    h.plans.beginFrame(h.owner, next)
    expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('pending')
    await h.flush()
    expect(injected).toBe(true)
    expect.soft(charges).toHaveBeenCalledTimes(2)
    expect(h.store.cachedRows).toBe(3)
    for (let row = 0; row < 3; row++) expect(h.store.cellsOf(row)).toBe(oldRows[row])
    expect(h.remote.ledger.bytes).toBe(720)
    expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(before.byPool.rowCache)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(before.byPool.layout)
    expect(prior.kind).toBe('detail')
    expect(h.settled).not.toHaveBeenCalledWith(await h.reply())
    expect(h.releases).toHaveBeenCalledTimes(2)
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { set.mockRestore(); charges.mockRestore(); await h.close() }
})
it.each((['D', 'M'] as const).flatMap(profile => (['scope', 'dispose', 'release', 'close-throw', 'ordinary-throw', 'owner', 'merge'] as const).map(change => ({ profile, change }))))('QUALITY multirow $profile $change preserves publication ownership and rollback', async ({ profile, change }) => {
  const h = await validationHarness({ rows: 3, profile })
  const window = { ...h.window, box: { ...h.window.box, height: 60 } }
  const prior = await h.fetch(window)
  const priorRows = [h.store.cellsOf(0), h.store.cellsOf(1), h.store.cellsOf(2)]
  const before = h.remote.mainLedger.snapshot()
  const other = await h.actualViewport(h.owner, { ...window, box: { ...window.box, y: 100 } }, new AbortController().signal, 'tile')
  const next = { ...window, box: { ...window.box, x: 64, width: 128 } }
  const actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
  const charges = vi.spyOn(h.remote.ledger, 'charge')
  let count = 0, injected = false
  const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (injected || pool !== 'rowCache' || bytes <= 0 || ++count !== 2) return
    injected = true
    // Inspect through the actual public route while publication is synchronous.
    expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('detail')
    if (change === 'scope') h.plans.beginFrame(h.owner, { ...window, zoom: 2 })
    else if (change === 'dispose') h.plans.dispose()
    else if (change === 'release') h.plans.release(h.owner, next)
    else if (change === 'owner') h.remote.identity('s').fontGen++
    else if (change === 'merge') h.remote.merges('s').gen++
    else {
      if (change === 'close-throw') (h.stage.mock.results.at(-1)!.value as NonNullable<ReturnType<typeof h.actualStage>>).close()
      throw new Error('QUALITY primary publication fault')
    }
  })
  try {
    h.plans.beginFrame(h.owner, next)
    expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('pending')
    await h.flush()
    expect(injected).toBe(true)
    expect(charges.mock.calls.map(([, row]) => row)).toEqual([0, 1])
    for (let row = 0; row < 3; row++) expect(h.store.cellsOf(row)).toBe(priorRows[row])
    expect(h.remote.ledger.bytes).toBe(720)
    expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(before.byPool.rowCache)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(change === 'scope' || change === 'dispose' ? 0 : before.byPool.layout)
    expect(h.settled).not.toHaveBeenCalledWith(await h.reply())
    expect(h.settled).not.toHaveBeenCalledWith(other)
    expect(h.releases).toHaveBeenCalledTimes(2)
    expect(h.releases).toHaveBeenLastCalledWith(profile === 'D' ? 8 : 4)
    expect(h.remote.pagesInFlight).toBe(0)
    const token = h.stage.mock.results.at(-1)!.value as NonNullable<ReturnType<typeof h.actualStage>>
    expect(token.commit(() => { throw new Error('revived') })).toBe(false)
    if (change === 'ordinary-throw' || change === 'close-throw') {
      expect(h.plans.get(h.owner, window, h.extent, true)).toBe(prior)
      expect(h.plans.get(h.owner, next, h.extent, true)).toMatchObject({ kind: 'diagnostic', diagnostics: [{ message: 'QUALITY primary publication fault' }] })
    }
    if (change === 'release') {
      set.mockRestore()
      h.plans.beginFrame(h.owner, next)
      expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('pending')
      await h.flush()
      expect(h.requests).toHaveBeenCalledTimes(3)
      expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('detail')
    }
  } finally { set.mockRestore(); charges.mockRestore(); await h.close() }
})
it.each(['D', 'M'] as const)('QUALITY replacement %s request survives synchronous release during publication', async profile => {
  const h = await validationHarness({ rows: 3, profile })
  const window = { ...h.window, box: { ...h.window.box, height: 60 } }
  const prior = await h.fetch(window), old = [h.store.cellsOf(0), h.store.cellsOf(1), h.store.cellsOf(2)]
  const next = { ...window, box: { ...window.box, x: 64, width: 128 } }
  const actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
  let count = 0, injected = false
  const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (injected || pool !== 'rowCache' || bytes <= 0 || ++count !== 2) return
    injected = true
    h.plans.release(h.owner, next)
    h.plans.beginFrame(h.owner, next)
    expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('pending')
  })
  try {
    h.plans.beginFrame(h.owner, next)
    expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('pending')
    await h.flush()
    expect(injected).toBe(true)
    expect(h.requests).toHaveBeenCalledTimes(3)
    expect(h.releases).toHaveBeenCalledTimes(3)
    expect(h.remote.pagesInFlight).toBe(0)
    expect(h.plans.get(h.owner, window, h.extent, true)).toBe(prior)
    // The first attempt is rejected, but its replacement must retain ownership.
    const replies = await Promise.all(h.requests.mock.results.map(r => r.value))
    expect(h.settled).not.toHaveBeenCalledWith(replies[1])
    expect.soft(h.settled).toHaveBeenCalledWith(replies[2])
    expect.soft(h.plans.get(h.owner, next, h.extent, true).kind).toBe('detail')
    expect.soft(h.store.cellsOf(0)).not.toBe(old[0])
  } finally { set.mockRestore(); await h.close() }
})
it.each((['D', 'M'] as const).flatMap(profile => (['layout', 'rowCache', 'none'] as const).map(pool => ({ profile, pool }))))('QUALITY other release $profile during $pool retained accounting keeps exact layout bytes', async ({ profile, pool: targetPool }) => {
  const { queryPayloadBytes, WORKER_QUERY_BYTES } = await import('../src/worker/viewerQueryBudget')
  const h = await validationHarness({ rows: 3, profile })
  const window = { ...h.window, box: { ...h.window.box, height: 60 } }
  const oldPlan = await h.fetch(window)
  const next = { ...window, box: { ...window.box, x: 64, width: 128 } }
  const actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
  const other = await h.actualViewport(h.owner, { ...window, box: { ...window.box, y: 100 } }, new AbortController().signal, 'tile')
  let injected = false
  const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (injected || pool !== targetPool || bytes <= 0 || pool === 'layout' && (!owner.startsWith('viewport-plans:') || owner.includes(':stage:'))) return
    injected = true
    h.plans.release(h.owner, window)
  })
  try {
    h.plans.beginFrame(h.owner, next)
    expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('pending')
    await h.flush()
    const plan = h.plans.get(h.owner, next, h.extent, true)
    expect(plan.kind).toBe('detail')
    if (plan.kind !== 'detail') throw new Error('Expected current detail')
    expect(plan.cells.map(cell => cell.source)).toEqual([{ row: 0, col: 0 }, { row: 1, col: 0 }, { row: 2, col: 0 }])
    expect(injected).toBe(targetPool !== 'none')
    const expected = queryPayloadBytes(plan, WORKER_QUERY_BYTES[profile]) + (targetPool === 'none' ? queryPayloadBytes(oldPlan, WORKER_QUERY_BYTES[profile]) : 0)
    expect.soft(h.remote.mainLedger.snapshot().byPool.layout).toBe(expected)
    expect(h.remote.ledger.bytes).toBe(720)
    expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(720)
    const incoming = await h.reply()
    expect(h.settled.mock.calls.filter(([out]) => out === incoming)).toHaveLength(1)
    expect(h.settled).not.toHaveBeenCalledWith(other)
    expect(h.releases).toHaveBeenCalledTimes(2)
    expect(h.releases).toHaveBeenLastCalledWith(profile === 'D' ? 8 : 4)
    expect(h.remote.pagesInFlight).toBe(0)
    if (targetPool !== 'none') {
      set.mockRestore()
      h.plans.beginFrame(h.owner, window)
      expect(h.plans.get(h.owner, window, h.extent, true).kind).toBe('pending')
      expect(h.requests).toHaveBeenCalledTimes(3)
      await h.flush()
    }
  } finally { set.mockRestore(); await h.close() }
})

const qfixPhases = ['layout', 'row-first', 'row-middle', 'row-last'] as const
it.each((['D', 'M'] as const).flatMap(profile => qfixPhases.flatMap(phase => (['release', 'replace'] as const).map(action => ({ profile, phase, action })))))('QFIX current $profile $action during $phase preserves public removal and matching replacement', async ({ profile, phase, action }) => {
  const h = await validationHarness({ rows: 3, profile }), window = { ...h.window, box: { ...h.window.box, height: 60 } }
  const prior = await h.fetch(window), oldRows = [h.store.cellsOf(0), h.store.cellsOf(1), h.store.cellsOf(2)]
  const other = await h.actualViewport(h.owner, { ...window, box: { ...window.box, y: 100 } }, new AbortController().signal, 'tile')
  const next = { ...window, box: { ...window.box, x: 64, width: 128 } }, actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
  let rowCharges = 0, injected = false
  const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (injected || bytes <= 0) return
    const atLayout = phase === 'layout' && pool === 'layout' && owner.startsWith('viewport-plans:') && !owner.includes(':stage:')
    const atRow = pool === 'rowCache' && ++rowCharges === (phase === 'row-first' ? 1 : phase === 'row-middle' ? 2 : phase === 'row-last' ? 3 : -1)
    if (!atLayout && !atRow) return
    injected = true
    h.plans.release(h.owner, next)
    if (action === 'replace') {
      h.plans.beginFrame(h.owner, next)
      expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('pending')
    }
  })
  try {
    h.plans.beginFrame(h.owner, next)
    expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('pending')
    await h.flush()
    expect(injected).toBe(true)
    expect(h.plans.get(h.owner, window, h.extent, true)).toBe(prior)
    const replies = await Promise.all(h.requests.mock.results.map(r => r.value))
    expect(h.settled).not.toHaveBeenCalledWith(replies[1])
    expect(h.settled).not.toHaveBeenCalledWith(other)
    if (action === 'release') {
      for (let row = 0; row < 3; row++) expect(h.store.cellsOf(row)).toBe(oldRows[row])
      set.mockRestore()
      h.plans.beginFrame(h.owner, next)
      expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('pending')
      await h.flush()
    } else expect(h.settled).toHaveBeenCalledWith(replies[2])
    expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('detail')
    expect(h.requests).toHaveBeenCalledTimes(3)
    expect(h.releases).toHaveBeenCalledTimes(3)
    expect(h.releases).toHaveBeenLastCalledWith(profile === 'D' ? 8 : 4)
    expect(h.remote.pagesInFlight).toBe(0)
    expect(h.remote.ledger.bytes).toBe(720)
    expect(h.store.cachedRows).toBe(3)
  } finally { set.mockRestore(); await h.close() }
})
it.each((['D', 'M'] as const).flatMap(profile => qfixPhases.map(phase => ({ profile, phase }))))('QFIX other $profile release during $phase retains only the live current plan bytes', async ({ profile, phase }) => {
  const { queryPayloadBytes, WORKER_QUERY_BYTES } = await import('../src/worker/viewerQueryBudget')
  const h = await validationHarness({ rows: 3, profile }), window = { ...h.window, box: { ...h.window.box, height: 60 } }
  await h.fetch(window)
  const next = { ...window, box: { ...window.box, x: 64, width: 128 } }, actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
  let rowCharges = 0, injected = false
  const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (injected || bytes <= 0) return
    const atLayout = phase === 'layout' && pool === 'layout' && owner.startsWith('viewport-plans:') && !owner.includes(':stage:')
    const atRow = pool === 'rowCache' && ++rowCharges === (phase === 'row-first' ? 1 : phase === 'row-middle' ? 2 : phase === 'row-last' ? 3 : -1)
    if (!atLayout && !atRow) return
    injected = true
    h.plans.release(h.owner, window)
  })
  try {
    const plan = await h.fetch(next), reply = await h.reply()
    expect(injected).toBe(true)
    expect(plan.kind).toBe('detail')
    if (plan.kind !== 'detail') throw new Error('Expected current live plan')
    expect(plan.cells.map(cell => cell.source)).toEqual([{ row: 0, col: 0 }, { row: 1, col: 0 }, { row: 2, col: 0 }])
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(queryPayloadBytes(plan, WORKER_QUERY_BYTES[profile]))
    expect(h.settled.mock.calls.filter(([out]) => out === reply)).toHaveLength(1)
    expect(h.remote.ledger.bytes).toBe(720)
    expect(h.releases).toHaveBeenCalledTimes(2)
    expect(h.remote.pagesInFlight).toBe(0)
    set.mockRestore()
    h.plans.beginFrame(h.owner, window)
    expect(h.plans.get(h.owner, window, h.extent, true).kind).toBe('pending')
    await h.flush()
    expect(h.releases).toHaveBeenCalledTimes(3)
  } finally { set.mockRestore(); await h.close() }
})
it.each((['D', 'M'] as const).flatMap(profile => (['layout', 'row-middle'] as const).map(phase => ({ profile, phase }))))('QFIX primary fault $profile during $phase preserves other removal and prior source rows', async ({ profile, phase }) => {
  const h = await validationHarness({ rows: 3, profile }), window = { ...h.window, box: { ...h.window.box, height: 60 } }
  await h.fetch(window)
  const oldRows = [h.store.cellsOf(0), h.store.cellsOf(1), h.store.cellsOf(2)], next = { ...window, box: { ...window.box, x: 64, width: 128 } }
  const actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
  let rowCharges = 0, injected = false
  const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (injected || bytes <= 0) return
    const atLayout = phase === 'layout' && pool === 'layout' && owner.startsWith('viewport-plans:') && !owner.includes(':stage:')
    const atRow = phase === 'row-middle' && pool === 'rowCache' && ++rowCharges === 2
    if (!atLayout && !atRow) return
    injected = true
    h.plans.release(h.owner, window)
    throw new Error('QFIX primary post-set fault')
  })
  try {
    h.plans.beginFrame(h.owner, next)
    expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('pending')
    await h.flush()
    expect(injected).toBe(true)
    const incoming = await h.reply()
    for (let row = 0; row < 3; row++) expect(h.store.cellsOf(row)).toBe(oldRows[row])
    expect(h.remote.ledger.bytes).toBe(720)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
    expect(h.plans.get(h.owner, next, h.extent, true)).toMatchObject({ kind: 'diagnostic', diagnostics: [{ message: 'QFIX primary post-set fault' }] })
    expect(h.settled).not.toHaveBeenCalledWith(incoming)
    expect(h.releases).toHaveBeenCalledTimes(2)
    expect(h.remote.pagesInFlight).toBe(0)
    set.mockRestore()
    h.plans.beginFrame(h.owner, window)
    expect(h.plans.get(h.owner, window, h.extent, true).kind).toBe('pending')
    await h.flush()
    expect(h.releases).toHaveBeenCalledTimes(3)
  } finally { set.mockRestore(); await h.close() }
})

it.each(['D', 'M'] as const)('QFIX repeated %s public removals during reconciliation refuse without an unbounded accounting loop', async profile => {
  const h = await validationHarness({ rows: 3, profile }), first = { ...h.window, box: { ...h.window.box, height: 60 } }
  const second = { ...first, box: { ...first.box, x: 192, width: 128 } }, next = { ...first, box: { ...first.box, x: 64, width: 128 } }
  await h.fetch(first); await h.fetch(second)
  const oldRows = [h.store.cellsOf(0), h.store.cellsOf(1), h.store.cellsOf(2)]
  const other = await h.actualViewport(h.owner, { ...first, box: { ...first.box, y: 100 } }, new AbortController().signal, 'tile')
  const actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger), charges = vi.spyOn(h.remote.ledger, 'charge')
  let ownedWrites = 0, insideRemoval = false
  const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (insideRemoval || pool !== 'layout' || bytes <= 0 || !owner.startsWith('viewport-plans:') || owner.includes(':stage:')) return
    ownedWrites++
    if (ownedWrites > 2) throw new Error('unbounded reconciliation')
    insideRemoval = true
    try { h.plans.release(h.owner, ownedWrites === 1 ? first : second) }
    finally { insideRemoval = false }
  })
  try {
    h.plans.beginFrame(h.owner, next)
    expect(h.plans.get(h.owner, next, h.extent, true).kind).toBe('pending')
    await h.flush()
    const incoming = await h.reply()
    expect.soft(ownedWrites).toBe(2)
    expect.soft(h.plans.get(h.owner, next, h.extent, true)).toMatchObject({ kind: 'diagnostic', diagnostics: [{ message: 'Viewport plan accounting changed during publication' }] })
    for (let row = 0; row < 3; row++) expect(h.store.cellsOf(row)).toBe(oldRows[row])
    expect(h.remote.ledger.bytes).toBe(720)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
    expect(charges).not.toHaveBeenCalled()
    expect(h.settled).not.toHaveBeenCalledWith(incoming)
    expect(h.settled).not.toHaveBeenCalledWith(other)
    expect(h.releases).toHaveBeenCalledTimes(3)
    expect(h.remote.pagesInFlight).toBe(0)
    set.mockRestore(); charges.mockRestore()
    h.plans.release(h.owner, next)
    expect((await h.fetch(next)).kind).toBe('detail')
    expect(h.releases).toHaveBeenCalledTimes(4)
    expect(h.settled).not.toHaveBeenCalledWith(other)
  } finally { set.mockRestore(); charges.mockRestore(); await h.close() }
})

async function evictionHarness(options: { at?: { row: number; col: number }; hidden?: boolean; tiny?: boolean; merge?: boolean; mergeAnchor?: boolean; twoRows?: boolean; rows?: number; profile?: 'D' | 'M'; text?: string; rowTexts?: readonly string[] } = {}) {
  const at = options.at ?? { row: 0, col: 0 }
  const doc = workbookOf([{ sheetId: 's', cells: [{ ...at, value: options.rowTexts?.[0] ?? options.text ?? LONG }, ...Array.from({ length: (options.rows ?? (options.twoRows ? 2 : 1)) - 1 }, (_, offset) => ({ row: at.row + offset + 1, col: at.col, value: options.rowTexts?.[offset + 1] ?? LONG })), ...(options.mergeAnchor ? [{ row: at.row, col: at.col + 1, value: LONG }] : [])] }]), sheet = doc.sheets[0]
  sheet.cols = options.hidden ? [{ min: 0, max: 16383, hidden: true }] : []
  if (options.merge) sheet.mergeRanges = [{ minRow: at.row, maxRow: at.row, minCol: at.col + 1, maxCol: at.col + 2 }]
  const styles = [{ numFmtId: 0 }, { numFmtId: 0, fillColor: 'FFFF00' }]
  if (options.mergeAnchor) { const anchor = sheet.rows[0].cells[1]; anchor.styleIndex = 1; anchor.style = styles[1] }
  const object = createObjectWorkbookSource(doc, 'slice6')
  const baseGeometry = object.geometry('s')
  const baseColumns = options.tiny ? createAxisMetrics({ logicalCount: 16384, defaultSizeCss: 64, overrides: [{ start: 0, end: 16383, sizeCss: 1 }] }) : baseGeometry.cols
  const locateRows = vi.fn(baseGeometry.rows.locate.bind(baseGeometry.rows)), locateCols = vi.fn(baseColumns.locate.bind(baseColumns))
  const observedGeometry = { ...baseGeometry, rows: { ...baseGeometry.rows, locate: locateRows }, cols: { ...baseColumns, locate: locateCols }, ...(options.tiny ? { colOverrides: [{ start: 0, end: 16383, sizeCss: 1 }] } : {}) }
  const source = { ...object, ...(options.mergeAnchor ? { formatting: () => ({ dateSystem: '1900' as const, styles }) } : {}), geometry: (_id: string) => observedGeometry }
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const native = connectInProcess(kernel)
  let mutate: ((wire: Wire) => void) | undefined
  let originalPacket: unknown
  const wrappers = new Map<(event: { data: unknown }) => void, (event: { data: unknown }) => void>()
  const port: InProcessPort = {
    postMessage: (message, transfer) => native.postMessage(message, transfer), idle: () => native.idle(),
    addEventListener(type, listener) {
      const wrapper = (event: { data: unknown }) => {
        const wire = event.data as Wire
        if (wire.kind === 'complete' && wire.result?.kind === 'viewport' && mutate) { originalPacket = structuredClone(event.data); mutate(wire) }
        listener(event)
      }
      wrappers.set(listener, wrapper); native.addEventListener(type, wrapper)
    },
    removeEventListener(type, listener) { native.removeEventListener(type, wrappers.get(listener) ?? listener); wrappers.delete(listener) },
  }
  const client = new ViewerClient(port), profile = options.profile ?? 'D'
  const opened = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'slice6', source: { kind: 'object', document: doc }, profile })
  if (opened.kind !== 'complete' || opened.result.kind !== 'open') throw new Error('Expected actual open')
  const session = kernel.session('slice6')!, owner = session.identity('s')
  const remote = new RemoteWorkbook(client, { profile, docId: 'slice6', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: opened.result.formatting, layouts: new Map([['s', session.layout('s')]]) })
  const flush = async () => { await native.idle(); for (let i = 0; i < 30; i++) await Promise.resolve() }
  remote.merges('s'); await flush()
  const store = remote.store('s'), provider = source.store('s'), geometry = source.geometry('s')
  const actualRows = provider.readRows.bind(provider), actualStructure = provider.queryStructural.bind(provider)
  const reads = vi.spyOn(provider, 'readRows'), singles = vi.spyOn(provider, 'readCell'), structural = vi.spyOn(provider, 'queryStructural'), overview = vi.spyOn(provider, 'queryOverview')
  const actualViewport = remote.viewport.bind(remote), actualStage = store.stageViewport.bind(store)
  const requests = vi.spyOn(remote, 'viewport'), releases = vi.spyOn(remote, 'releasePages'), stage = vi.spyOn(store, 'stageViewport'), settled = vi.spyOn(remote.mainLedger, 'settle')
  const metrics = createSheetMetrics({ rows: geometry.rows, cols: geometry.cols })
  const extent = { width: metrics.cols.totalCss(), height: metrics.rows.totalCss() }
  const window: LogicalViewport = { box: { x: at.col * 64, y: at.row * 20, width: 192, height: 20 }, zoom: 1, deviceScale: 1 }
  const plans = new WorkerViewportPlans(remote)
  const get = tilePlanSource({ store: { coverage: () => 'oversized' } as unknown as typeof store, metrics, scale: 1, identity: owner, extent, plans })
  const fetch = async (w = window) => { plans.beginFrame(owner, w); expect(get(w, true)?.kind).toBe('pending'); await flush(); return get(w, true)! }
  const reply = async () => {
    const out: Awaited<ReturnType<RemoteWorkbook['viewport']>> = await requests.mock.results.at(-1)!.value
    if (typeof out !== 'object' || out.kind !== 'complete') throw new Error('Expected actual transported viewport')
    return out
  }
  const paint = (plan: ViewportPlan, w = window) => {
    const canvas = createCanvas(192, 20), ctx = canvas.getContext('2d')
    const rendered = renderSheetViewport(sheet, ctx as unknown as CanvasRenderingContext2D, undefined, w, undefined, { owner, preparedPlan: plan, cache: new LayoutCache() })
    return { rendered, pixels: ctx.getImageData(0, 0, 192, 20).data }
  }
  const canonical = (w = window) => {
    const built = buildSheetGeometry(sheet)
    if (built.kind !== 'ready') throw new Error('Expected canonical geometry')
    const canvas = createCanvas(192, 20), ctx = canvas.getContext('2d')
    const rendered = renderSheetViewport(sheet, ctx as unknown as CanvasRenderingContext2D, { ...built, metrics }, w, undefined, { owner, cache: new LayoutCache() })
    return { rendered, pixels: ctx.getImageData(0, 0, 192, 20).data }
  }
  const close = async () => { plans.dispose(); for (const spy of [locateRows, locateCols, reads, singles, structural, overview, requests, releases, stage, settled]) spy.mockRestore(); await remote.release() }
  return { at, doc, sheet, source, provider, actualRows, actualStructure, actualViewport, actualStage, client, remote, deliverOriginal: () => { for (const listener of wrappers.keys()) listener({ data: originalPacket }) }, store, plans, owner, window, extent, metrics, mutate: (fn?: (wire: Wire) => void) => { mutate = fn }, flush, fetch, reply, paint, canonical, close, reads, singles, structural, overview, requests, releases, stage, settled, locateRows, locateCols }
}


it.each((['D', 'M'] as const).flatMap(profile => [false, true].map(release => ({ profile, release }))))('SPEC eviction selected other $profile release=$release retains exact layout ownership', async ({ profile, release }) => {
  const { queryPayloadBytes, WORKER_QUERY_BYTES } = await import('../src/worker/viewerQueryBudget')
  const h = await evictionHarness({ profile, text: 'E'.repeat(200000) })
  const count = profile === 'D' ? 20 : 10, cap = WORKER_QUERY_BYTES[profile]
  const old: Array<{ window: LogicalViewport; plan: ViewportPlan; bytes: number }> = []
  let sum = 0
  try {
    for (let i = 0; i < count; i++) {
      const window = { ...h.window, box: { ...h.window.box, x: i } }
      const plan = await h.fetch(window)
      expect(plan.kind).toBe('detail')
      if (plan.kind !== 'detail') throw new Error('Expected real admitted prior plan')
      expect(plan.cells).toHaveLength(1)
      expect(plan.cells[0].cell.value).toBe('E'.repeat(200000))
      const bytes = queryPayloadBytes(plan, cap)
      old.push({ window, plan, bytes }); sum += bytes
    }
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(sum)
    expect(sum).toBeLessThan(cap)
    expect(sum + old[0].bytes).toBeGreaterThan(cap)
    const next = { ...h.window, box: { ...h.window.box, x: count } }
    const actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
    let injected = false, proposed = 0
    const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
      actualSet(pool, owner, bytes)
      if (injected || pool !== 'layout' || bytes <= 0 || !owner.startsWith('viewport-plans:') || owner.includes(':stage:')) return
      injected = true; proposed = bytes
      if (release) h.plans.release(h.owner, old[0].window)
    })
    try {
      const plan = await h.fetch(next), incoming = await h.reply()
      expect(injected).toBe(true)
      expect(plan.kind).toBe('detail')
      if (plan.kind !== 'detail') throw new Error('Expected real admitted incoming plan')
      expect(plan.cells[0].cell.value).toBe('E'.repeat(200000))
      const expected = sum - old[0].bytes + queryPayloadBytes(plan, cap)
      expect(proposed).toBe(expected)
      for (const retained of old.slice(1)) expect(h.plans.get(h.owner, retained.window, h.extent, true)).toBe(retained.plan)
      expect.soft(h.remote.mainLedger.snapshot().byPool.layout).toBe(expected)
      expect(h.settled.mock.calls.filter(([out]) => out === incoming)).toHaveLength(1)
      expect(h.releases).toHaveBeenCalledTimes(count + 1)
      expect(h.releases).toHaveBeenLastCalledWith(profile === 'D' ? 8 : 4)
      expect(h.remote.pagesInFlight).toBe(0)
      expect(h.store.cachedRows).toBe(1)
      set.mockRestore()
      h.plans.beginFrame(h.owner, old[0].window)
      expect(h.plans.get(h.owner, old[0].window, h.extent, true).kind).toBe('pending')
      await h.flush()
      expect(h.requests).toHaveBeenCalledTimes(count + 2)
    } finally { set.mockRestore() }
  } finally { await h.close() }
})

it.each(['D', 'M'] as const)('EVICTION repeated equal %s selected removals refuse after one reconciliation without reviving either released plan', async profile => {
  const { queryPayloadBytes, WORKER_QUERY_BYTES } = await import('../src/worker/viewerQueryBudget')
  const value = 'E'.repeat(200000), h = await evictionHarness({ profile, rows: 2, rowTexts: [value, value] })
  const count = profile === 'D' ? 20 : 10, cap = WORKER_QUERY_BYTES[profile]
  const old: Array<{ window: LogicalViewport; plan: ViewportPlan; bytes: number }> = []
  let sum = 0
  try {
    for (let i = 0; i < count; i++) {
      const window = { ...h.window, box: { ...h.window.box, x: i } }, plan = await h.fetch(window)
      expect(plan.kind).toBe('detail')
      const bytes = queryPayloadBytes(plan, cap); old.push({ window, plan, bytes }); sum += bytes
    }
    const oldRow = h.store.cellsOf(0), before = h.remote.mainLedger.snapshot()
    expect(sum).toBeLessThan(cap)
    expect(sum + 2 * old[0].bytes).toBeGreaterThan(cap)
    const next = { ...h.window, box: { ...h.window.box, x: count, height: 40 } }
    const other = await h.actualViewport(h.owner, { ...h.window, box: { ...h.window.box, y: 100 } }, new AbortController().signal, 'tile')
    const actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger), charges = vi.spyOn(h.remote.ledger, 'charge')
    const ownedWrites: number[] = [], rollbackWrites: number[] = []
    const rollbackBytes = sum - old[0].bytes - old[1].bytes
    let insideRemoval = false
    const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
      actualSet(pool, owner, bytes)
      if (insideRemoval || pool !== 'layout' || bytes <= 0 || !owner.startsWith('viewport-plans:') || owner.includes(':stage:')) return
      if (bytes === rollbackBytes) { rollbackWrites.push(bytes); return }
      ownedWrites.push(bytes)
      if (ownedWrites.length > 2) throw new Error('unbounded equal-total reconciliation')
      insideRemoval = true
      try { h.plans.release(h.owner, old[ownedWrites.length - 1].window) }
      finally { insideRemoval = false }
    })
    try {
      const plan = await h.fetch(next), incoming = await h.reply()
      expect.soft(ownedWrites).toHaveLength(2)
      expect.soft(ownedWrites[1]).toBe(ownedWrites[0])
      expect(rollbackWrites).toEqual([rollbackBytes])
      expect.soft(plan).toMatchObject({ kind: 'diagnostic', diagnostics: [{ message: 'Viewport plan accounting changed during publication' }] })
      expect(h.store.cellsOf(0)).toBe(oldRow)
      expect(h.store.cellsOf(1)).toBeUndefined()
      expect(h.remote.ledger.bytes).toBe(before.byPool.rowCache)
      expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(sum - old[0].bytes - old[1].bytes)
      expect(charges).not.toHaveBeenCalled()
      expect(h.settled).not.toHaveBeenCalledWith(incoming)
      expect(h.settled).not.toHaveBeenCalledWith(other)
      expect(h.releases).toHaveBeenCalledTimes(count + 1)
      expect(h.remote.pagesInFlight).toBe(0)
      for (const retained of old.slice(2)) expect(h.plans.get(h.owner, retained.window, h.extent, true)).toBe(retained.plan)
      set.mockRestore(); charges.mockRestore()
      h.plans.beginFrame(h.owner, old[0].window)
      expect(h.plans.get(h.owner, old[0].window, h.extent, true).kind).toBe('pending')
      await h.flush()
      h.plans.beginFrame(h.owner, old[1].window)
      expect(h.plans.get(h.owner, old[1].window, h.extent, true).kind).toBe('pending')
      await h.flush()
      h.plans.release(h.owner, next)
      expect((await h.fetch(next)).kind).toBe('detail')
      expect(h.settled).not.toHaveBeenCalledWith(other)
      expect(h.releases).toHaveBeenCalledTimes(count + 4)
      expect(h.remote.pagesInFlight).toBe(0)
    } finally { set.mockRestore(); charges.mockRestore() }
  } finally { await h.close() }
})
it.each(['D', 'M'] as const)('EVICTION pending zero-byte %s other release still restores exact current layout accounting', async profile => {
  const { queryPayloadBytes, WORKER_QUERY_BYTES } = await import('../src/worker/viewerQueryBudget')
  const h = await evictionHarness({ profile }), first = h.window, pending = { ...h.window, box: { ...h.window.box, x: 384, width: 64 } }, next = { ...h.window, box: { ...h.window.box, x: 64, width: 128 } }
  const oldPlan = await h.fetch(first)
  const other = await h.actualViewport(h.owner, { ...h.window, box: { ...h.window.box, y: 100 } }, new AbortController().signal, 'tile')
  let enter!: () => void, release!: () => void, hold = true
  const entered = new Promise<void>(resolve => { enter = resolve }), held = new Promise<void>(resolve => { release = resolve })
  h.reads.mockImplementation(async (...args) => {
    const page = await h.actualRows(...args)
    if (hold) { hold = false; enter(); await held }
    return page
  })
  h.plans.beginFrame(h.owner, pending)
  expect(h.plans.get(h.owner, pending, h.extent, true).kind).toBe('pending')
  await entered
  const actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
  let injected = false
  const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (injected || pool !== 'layout' || bytes <= 0 || !owner.startsWith('viewport-plans:') || owner.includes(':stage:')) return
    injected = true
    expect(h.plans.get(h.owner, pending, h.extent, true).kind).toBe('pending')
    expect(h.requests).toHaveBeenCalledTimes(3)
    expect(h.remote.pagesInFlight).toBe(profile === 'D' ? 16 : 8)
    h.plans.release(h.owner, pending)
    release()
  })
  try {
    const plan = await h.fetch(next), incoming = await h.reply()
    expect(injected).toBe(true)
    expect(plan.kind).toBe('detail')
    expect(h.plans.get(h.owner, first, h.extent, true)).toBe(oldPlan)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(queryPayloadBytes(oldPlan, WORKER_QUERY_BYTES[profile]) + queryPayloadBytes(plan, WORKER_QUERY_BYTES[profile]))
    expect(h.settled.mock.calls.filter(([out]) => out === incoming)).toHaveLength(1)
    expect(h.settled).not.toHaveBeenCalledWith(other)
    expect(h.releases).toHaveBeenCalledTimes(3)
    expect(h.remote.pagesInFlight).toBe(0)
    set.mockRestore()
    h.plans.beginFrame(h.owner, pending)
    expect(h.plans.get(h.owner, pending, h.extent, true).kind).toBe('pending')
    await h.flush()
    expect(h.releases).toHaveBeenCalledTimes(4)
  } finally { release(); set.mockRestore(); await h.close() }
})
