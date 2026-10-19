import { SheetTileRenderer, type SheetFrame } from '../src/components/xlsx/sheetTiles'
import { LayoutCache, type CellLayout } from '../src/xlsx/view/textLayout'
import { MergeQueryCache } from '../src/xlsx/view/viewport'
import type { XlsxCellStyle, XlsxColumnSpec, XlsxMergeRange } from '../src/xlsx/types'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import type { AxisSizeInterval, LogicalViewport, Page, SparseRowPage, SortSpec } from '../src/xlsx/view/contracts'
import { viewportPlanAdmission, VIEWPORT_MEASUREMENT_BYTES, measureViewportPayload, type OverflowViewportResult } from '../src/xlsx/view/viewportOverflow'
import * as queryBudget from '../src/worker/viewerQueryBudget'
import { queryPayloadBytes, WORKER_QUERY_BYTES } from '../src/worker/viewerQueryBudget'
import { RowCacheLedger } from '../src/worker/viewerRemoteSource'
// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { ViewerClient } from '../src/worker/viewerClient'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { WorkerViewportPlans, tilePlanSource } from '../src/components/xlsx/workerViewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { buildSheetGeometry, type ViewportPlan } from '../src/xlsx/view/viewport'
import { workbookOf } from './helpers/xlsx-find'

async function paintCase(value = 'abcdefghijklmnopq', neighbour: 'occupied' | 'blank' | 'format' | 'empty' | 'formula' | 'unavailable' = 'occupied') {
  const cells = [{ row: 0, col: 0, value }, { row: 0, col: 2, value: 'C BLOCKER' }]
  if (neighbour !== 'blank') cells.push({ row: 0, col: 1, value: neighbour === 'occupied' ? 'BLOCKER' : '' } as any)
  const doc = workbookOf([{ sheetId: 's', cells }])
  const b = doc.sheets[0].rows[0].cells.find(c => c.col === 1)
  if (b && neighbour === 'format') b.value = null
  if (b && neighbour === 'format') { b.styleIndex = 1; b.style = { numFmtId: 0, fillColor: 'FFFF00' } }
  if (b && (neighbour === 'formula' || neighbour === 'unavailable')) { b.formula = '""'; if (neighbour === 'unavailable') { b.value = null; (b as any).hasCachedValue = false } }
  if (neighbour === 'unavailable') doc.valuePolicy = 'cached'
  const source = createObjectWorkbookSource(doc, 'doc')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const client = new ViewerClient(connectInProcess(kernel))
  await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })
  const session = kernel.session('doc')!
  const remote = new RemoteWorkbook(client, { docId: 'doc', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: neighbour === 'format' ? { dateSystem: '1900', styles: [{ numFmtId: 0 }, b!.style!] } : source.formatting?.(), layouts: new Map([['s', session.layout('s')]]) })
  const requests = vi.spyOn(remote, 'viewport')
  const owner = session.identity('s'), window = { box: { x: 0, y: 0, width: 192, height: 20 }, zoom: 1, deviceScale: 1 }
  const plans = new WorkerViewportPlans(remote)
  const get = tilePlanSource({ store: { coverage: () => 'oversized' } as any, metrics: {} as any, scale: 1, identity: owner, extent: { width: 192, height: 20 }, plans })
  try {
    expect(get(window, true)?.kind).toBe('pending')
    let plan: ViewportPlan | undefined
    for (let i = 0; i < 100; i++) { await new Promise(resolve => setTimeout(resolve, 0)); plan = get(window, true); if (plan?.kind !== 'pending') break }
    expect(plan?.kind).toBe('detail')
    expect(requests).toHaveBeenCalledTimes(1)
    const transported = await requests.mock.results[0].value as Awaited<ReturnType<RemoteWorkbook['viewport']>>
    if (neighbour === 'unavailable' && typeof transported === 'object' && transported.kind === 'complete' && transported.result.content.kind === 'detail') expect(transported.result.content.cells.find(cell => cell.source.col === 1)?.kind).toBe('unavailable')
    const paint = (prepared: boolean, deterministic: boolean) => {
      const canvas = createCanvas(192, 20), ctx = canvas.getContext('2d')
      if (deterministic) ctx.measureText = (s: string) => ({ width: s.length * 7 }) as any
      const out = renderSheetViewport(doc.sheets[0], ctx as any, prepared ? undefined : buildSheetGeometry(doc.sheets[0]), window, undefined, { owner, ...(prepared ? { preparedPlan: plan } : {}) })
      return { text: out.records.find(r => r.source.col === 0)!.body.text as any, pixels: ctx.getImageData(0, 0, 192, 20).data }
    }
    const prepared = paint(true, true), resident = paint(false, true)
    expect(prepared.text.clipX).toEqual(resident.text.clipX)
    const actual = paint(true, false)
    expect(actual.pixels).toEqual(paint(false, false).pixels)
    if (neighbour === 'format') expect([...actual.pixels.slice((1 * 192 + 100) * 4, (1 * 192 + 100) * 4 + 4)]).toEqual([255, 255, 0, 255])
    if (neighbour === 'format' || neighbour === 'blank') {
      let ink = false
      for (let y = 2; y < 17; y++) for (let x = 65; x < 127; x++) { const offset = (y * 192 + x) * 4; if (actual.pixels[offset] < 128 && actual.pixels[offset + 1] < 128 && actual.pixels[offset + 2] < 128 && actual.pixels[offset + 3] > 0) ink = true }
      expect(ink, 'A1 text must remain visible inside B1').toBe(true)
    }
    return prepared.text
  } finally { plans.dispose(); await remote.release() }
}

it('clips the real cloned public worker plan at occupied B1 and matches resident crop', async () => {
  expect((await paintCase()).clipX).toEqual({ x0: 0, x1: 64 })
})
it.each(['blank', 'format'] as const)('allows complete proven %s B1 and clips longer text at C1', async neighbour => {
  expect((await paintCase('abcdefghijklmnopq', neighbour)).clipX).toBeUndefined()
  expect((await paintCase('abcdefghijklmnopqrstuvw', neighbour)).clipX).toEqual({ x0: 0, x1: 128 })
})
it.each(['empty', 'formula', 'unavailable'] as const)('preserves %s B1 blockers', async neighbour => {
  expect((await paintCase('abcdefghijklmnopq', neighbour)).clipX).toEqual({ x0: 0, x1: 64 })
})

async function bridgeHarness(rowCacheBytes?: number) {
  const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 'abcdefghijklmnopq' }, { row: 0, col: 1, value: 'BLOCKER' }] }])
  const source = createObjectWorkbookSource(doc, 'doc')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const port = connectInProcess(kernel), client = new ViewerClient(port)
  await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })
  const session = kernel.session('doc')!, owner = session.identity('s')
  const remote = new RemoteWorkbook(client, { docId: 'doc', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map([['s', session.layout('s')]]), ...(rowCacheBytes === undefined ? {} : { rowCacheBytes }) })
  const store = remote.store('s'), plans = new WorkerViewportPlans(remote)
  const window = { box: { x: 0, y: 0, width: 128, height: 20 }, zoom: 1, deviceScale: 1 }, extent = { width: 128, height: 20 }
  const settle = async () => { await port.idle(); await new Promise(resolve => setTimeout(resolve, 0)) }
  const reply = await remote.viewport(owner, window, new AbortController().signal, 'tile')
  if (typeof reply !== 'object' || reply.kind !== 'complete') throw new Error('Expected real viewport reply')
  return { source, remote, store, plans, window, extent, owner, reply, settle, close: async () => { plans.dispose(); await remote.release() } }
}

it('refuses before source staging and proof indexing when the joint main budget is full, without settling a held raw reply', async () => {
  const h = await bridgeHarness(), stage = vi.spyOn(h.store, 'stageViewport'), settled = vi.spyOn(h.remote.mainLedger, 'settle')
  const viewport = vi.spyOn(h.remote, 'viewport').mockResolvedValue(h.reply)
  const proof = (h.reply.result as any).overflowProof
  const indexing = vi.fn(() => proof.entries)
  Object.defineProperty(proof, 'entries', { get: indexing })
  h.remote.merges('s')
  await h.settle()
  h.remote.mainLedger.set('copy', 'pressure', 64 * 1024 * 1024 - h.remote.mainLedger.bytes)
  try {
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('pending')
    await h.settle()
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('diagnostic')
    expect(stage).not.toHaveBeenCalled()
    expect(indexing).not.toHaveBeenCalled()
    expect(h.store.cellsOf(0)).toBeUndefined()
    expect(settled).not.toHaveBeenCalledWith(h.reply)
    expect(h.remote.mainLedger.peakBytes).toBeLessThanOrEqual(64 * 1024 * 1024)
  } finally { h.remote.mainLedger.release('copy', 'pressure'); viewport.mockRestore(); await h.close() }
})

it.each(['missing', 'nan', 'duplicate', 'source'] as const)('diagnoses %s evidence before row adoption', async malformed => {
  const h = await bridgeHarness(), stage = vi.spyOn(h.store, 'stageViewport')
  const result = h.reply.result as any
  if (malformed === 'missing') delete result.overflowProof
  if (malformed === 'nan') result.overflowProof.entries[0].right = NaN
  if (malformed === 'duplicate') result.overflowProof.entries.push({ ...result.overflowProof.entries[0] })
  if (malformed === 'source') result.content.cells[0].cell.col = 1
  const viewport = vi.spyOn(h.remote, 'viewport').mockResolvedValue(h.reply)
  try {
    h.plans.get(h.owner, h.window, h.extent, true); await h.settle()
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('diagnostic')
    expect(stage).not.toHaveBeenCalled()
    expect(h.store.cellsOf(0)).toBeUndefined()
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
  } finally { viewport.mockRestore(); await h.close() }
})

it('declines row copying at the row subcap while independently retaining the paint plan', async () => {
  const h = await bridgeHarness(1), copying = vi.spyOn(h.remote, 'styleOf')
  const viewport = vi.spyOn(h.remote, 'viewport').mockResolvedValue(h.reply)
  try {
    h.plans.get(h.owner, h.window, h.extent, true); await h.settle()
    const plan = h.plans.get(h.owner, h.window, h.extent, true)
    expect(plan.kind).toBe('detail')
    expect(h.store.cellsOf(0)).toBeUndefined()
    // Canonical layout reads each style once; source staging must add no copies.
    expect(copying).toHaveBeenCalledTimes(2)
    expect(h.remote.ledger.bytes).toBe(0)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBeGreaterThan(0)
    expect(h.remote.mainLedger.peakBytes).toBeLessThanOrEqual(64 * 1024 * 1024)
    h.plans.dispose()
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
  } finally { viewport.mockRestore(); await h.close() }
})

it('closes viewport adoption tokens once and rolls source rows back if publication throws', async () => {
  const h = await bridgeHarness()
  try {
    if (h.reply.result.content.kind !== 'detail') throw new Error('Expected detail')
    const cells = h.reply.result.content.cells
    const controller = new AbortController(), token = h.store.stageViewport(cells, controller.signal)!
    expect(h.store.cellsOf(0)).toBeUndefined()
    controller.abort()
    expect(token.commit(() => { throw new Error('Closed token revived') })).toBe(false)
    token.close(); token.close()
    const rollback = h.store.stageViewport(cells, new AbortController().signal)!
    expect(() => rollback.commit(() => { throw new Error('publication failed') })).toThrow('publication failed')
    expect(h.store.cellsOf(0)).toBeUndefined()
    expect(h.remote.ledger.bytes).toBe(0)
    expect(rollback.commit(() => {})).toBe(false)
    const success = h.store.stageViewport(cells, new AbortController().signal)!
    expect(success.commit(() => {})).toBe(true)
    expect(success.commit(() => {})).toBe(false)
    expect(h.store.cellsOf(0)?.map(c => c.value)).toEqual(['abcdefghijklmnopq', 'BLOCKER'])
  } finally { await h.close() }
})

it('keeps real same-row incomplete omissions pending and clips known A1 at the unknown frontier', async () => {
  const h = await bridgeHarness(), sourceStore = h.source.store('s'), original = sourceStore.readRows.bind(sourceStore)
  const reads = vi.spyOn(sourceStore, 'readRows').mockImplementation(async (range, cursor, options) => {
    const page = await original(range, cursor, options)
    return { ...page, completeness: 'loading', items: page.items.map(item => ({ ...item, cells: item.cells.filter(cell => cell.col === 0) })) }
  })
  try {
    h.plans.get(h.owner, h.window, h.extent, true); await h.settle()
    expect(h.plans.get(h.owner, h.window, h.extent, true)).toMatchObject({ kind: 'detail', cells: [{ cell: { value: 'abcdefghijklmnopq' }, overflow: { right: 0 } }], pending: [{ bounds: { x: 64 } }] })
    expect(reads).toHaveBeenCalledTimes(1)
  } finally { reads.mockRestore(); await h.close() }
})

it('admits bookkeeping overlap as well as the measured staging reservation at publication', async () => {
  const trial = await bridgeHarness()
  trial.remote.merges('s'); await trial.settle()
  let staging = 0
  const original = trial.remote.mainLedger.set.bind(trial.remote.mainLedger)
  const charge = vi.spyOn(trial.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => { if (owner.includes(':stage:')) staging = Math.max(staging, bytes); original(pool, owner, bytes) })
  try {
    trial.plans.get(trial.owner, trial.window, trial.extent, true); await trial.settle()
    expect(staging).toBeGreaterThan(0)
  } finally { charge.mockRestore(); await trial.close() }
  const h = await bridgeHarness()
  h.remote.merges('s'); await h.settle()
  const viewport = vi.spyOn(h.remote, 'viewport').mockResolvedValue(h.reply)
  h.remote.mainLedger.set('copy', 'handoff-pressure', 64 * 1024 * 1024 - h.remote.mainLedger.bytes - staging)
  try {
    h.plans.get(h.owner, h.window, h.extent, true); await h.settle()
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('diagnostic')
    expect(h.remote.mainLedger.peakBytes).toBeLessThanOrEqual(64 * 1024 * 1024)
    expect(h.store.cellsOf(0)).toBeUndefined()
  } finally { h.remote.mainLedger.release('copy', 'handoff-pressure'); viewport.mockRestore(); await h.close() }
})

it('owns independent pending geometry before settling the raw sidecar', async () => {
  const h = await bridgeHarness(), result = h.reply.result as any
  result.completeness = 'loading'
  result.overflowProof.entries[0].rightEnd = 'unknown'
  result.overflowProof.pending = [{ bounds: { x: 64, y: 0, width: 64, height: 20 }, error: false }]
  result.content.cells[0].cell.sharedFormula = { si: 1, ref: 'A1:A2' }
  const viewport = vi.spyOn(h.remote, 'viewport').mockResolvedValue(h.reply)
  try {
    h.plans.get(h.owner, h.window, h.extent, true); await h.settle()
    const plan = h.plans.get(h.owner, h.window, h.extent, true)
    expect(plan.kind).toBe('detail')
    result.overflowProof.pending[0].bounds.x = 999
    result.content.cells[0].source.col = 999
    result.content.cells[0].cell.sharedFormula.si = 999
    if (plan.kind !== 'detail') throw new Error('Expected detail')
    expect(plan.pending?.[0].bounds.x).toBe(64)
    expect(plan.cells[0].source.col).toBe(0)
    expect(plan.cells[0].cell.sharedFormula?.si).toBe(1)
    expect(h.store.cellsOf(0)?.[0].sharedFormula?.si).toBe(1)
  } finally { viewport.mockRestore(); await h.close() }
})

it('rolls rows back and releases staging when guarded public plan publication throws', async () => {
  const h = await bridgeHarness(), viewport = vi.spyOn(h.remote, 'viewport').mockResolvedValue(h.reply)
  const original = h.remote.mainLedger.set.bind(h.remote.mainLedger)
  const charge = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    if (pool === 'layout' && owner.startsWith('viewport-plans:') && !owner.includes(':stage:') && bytes > 0) throw new Error('publish injection')
    original(pool, owner, bytes)
  })
  try {
    h.plans.get(h.owner, h.window, h.extent, true); await h.settle()
    expect(h.plans.get(h.owner, h.window, h.extent, true)).toMatchObject({ kind: 'diagnostic', diagnostics: [{ message: 'publish injection' }] })
    expect(h.store.cellsOf(0)).toBeUndefined()
    expect(h.remote.ledger.bytes).toBe(0)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
  } finally { charge.mockRestore(); viewport.mockRestore(); await h.close() }
})

it('validates column horizons without allocating the main full-sheet row geometry', async () => {
  const h = await bridgeHarness(), geometry = vi.spyOn(h.remote, 'geometry').mockImplementation(() => { throw new Error('full row geometry allocation') })
  const viewport = vi.spyOn(h.remote, 'viewport').mockResolvedValue(h.reply)
  try {
    h.plans.get(h.owner, h.window, h.extent, true); await h.settle()
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('detail')
    expect(geometry).not.toHaveBeenCalled()
  } finally { geometry.mockRestore(); viewport.mockRestore(); await h.close() }
})

it.each(['layout', 'rowCache'] as const)('SPEC rollback after actual %s retained mutation preserves prior readable state and totals', async poolToThrow => {
 const h = await bridgeHarness()
 h.remote.merges('s'); await h.settle()
 const oldWindow = { ...h.window, box: { ...h.window.box, y: 20 } }
 const oldReply = await h.remote.viewport(h.owner, oldWindow, new AbortController().signal, 'tile')
 const viewport = vi.spyOn(h.remote, 'viewport').mockResolvedValue(oldReply)
 h.plans.get(h.owner, oldWindow, h.extent, true); await h.settle()
 const oldPlan = h.plans.get(h.owner, oldWindow, h.extent, true)
 expect(oldPlan.kind).toBe('detail')
 const seed = structuredClone(h.reply.result.content.kind === 'detail' ? h.reply.result.content.cells : [])
 for (const cell of seed) if (cell.kind === 'cell') cell.cell.value = 'old'.repeat(1000)
 h.store.adoptViewport(seed)
 const oldRows = h.store.cellsOf(0), oldValues = oldRows?.map(c => c.value), before = h.remote.mainLedger.snapshot(), rowBefore = h.remote.ledger.bytes
 viewport.mockResolvedValue(h.reply)
 const realSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
 let injected = false
 const charge = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
   realSet(pool, owner, bytes)
   if (!injected && pool === poolToThrow && bytes > 0 && (pool === 'rowCache' || owner.startsWith('viewport-plans:') && !owner.includes(':stage:'))) { injected = true; throw new Error('after actual retained mutation') }
 })
 const settled = vi.spyOn(h.remote.mainLedger, 'settle')
 try {
  h.plans.get(h.owner, h.window, h.extent, true); await h.settle()
  expect(injected).toBe(true)
  expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('diagnostic')
  expect(h.plans.get(h.owner, oldWindow, h.extent, true)).toBe(oldPlan)
  expect(h.store.cellsOf(0)).toBe(oldRows)
  expect(h.store.cellsOf(0)?.map(c => c.value)).toEqual(oldValues)
  expect(settled).not.toHaveBeenCalledWith(h.reply)
  expect(h.remote.ledger.bytes).toBe(rowBefore)
  expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(before.byPool.rowCache)
  expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(before.byPool.layout)
 } finally { charge.mockRestore(); viewport.mockRestore(); await h.close() }
})

it('SPEC source epoch changes while viewport is in flight must reject its old reply', async () => {
  const h = await bridgeHarness()
  h.remote.merges('s')
  await h.settle()
  let deliver!: (reply: typeof h.reply) => void
  const viewport = vi.spyOn(h.remote, 'viewport').mockImplementation(() => new Promise(resolve => { deliver = resolve }))
  const settled = vi.spyOn(h.remote.mainLedger, 'settle')
  try {
    h.plans.get(h.owner, h.window, h.extent, true)
    h.store.clear()
    deliver(h.reply)
    await h.settle()
    expect(h.store.cellsOf(0)).toBeUndefined()
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).not.toBe('detail')
    expect(settled).not.toHaveBeenCalledWith(h.reply)
  } finally {
    viewport.mockRestore()
    await h.close()
  }
})

it.each([0, 1])('SPEC joint preflight boundary headroom offset %s', async over => {
  const h = await bridgeHarness()
  h.remote.merges('s')
  await h.settle()
  const result = h.reply.result
  const cells = result.content.kind === 'detail' ? result.content.cells : [], cap = WORKER_QUERY_BYTES.D
  const layout = h.remote.layout('s'), adoption = h.store.viewportAdoptionBytes(cells)
  const joint = viewportPlanAdmission(h.reply, result, h.remote.formatting().styles, adoption, h.store.viewportRollbackBytes(), layout.colOverrides.length + (layout.hiddenCols?.length ?? 0) + 1, cap).joint
  const viewport = vi.spyOn(h.remote, 'viewport').mockResolvedValue(h.reply), stage = vi.spyOn(h.store, 'stageViewport')
  h.remote.mainLedger.set('copy', 'pressure', 64 * 1024 * 1024 - h.remote.mainLedger.bytes - joint + over)
  try {
    h.plans.get(h.owner, h.window, h.extent, true)
    await h.settle()
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe(over ? 'diagnostic' : 'detail')
    expect(stage.mock.calls.length).toBe(over ? 0 : 1)
    expect(h.remote.mainLedger.peakBytes).toBeLessThanOrEqual(64 * 1024 * 1024)
  } finally {
    h.remote.mainLedger.release('copy', 'pressure')
    viewport.mockRestore()
    await h.close()
  }
})

it('SPEC row ledger exact and one-over preadmission', () => {
  const ledger = new RowCacheLedger(100), full = ledger.reserve(100)
  expect(full).toBeDefined()
  expect(ledger.reserve(1)).toBeUndefined()
  full!.close()
  full!.close()
  expect(ledger.reserve(101)).toBeUndefined()
  expect(ledger.reserve(100)).toBeDefined()
})

it('SPEC two-cell plan remains usable with 1 MiB synthetic aggregate-ledger headroom', async () => {
  const h = await bridgeHarness()
  h.remote.merges('s')
  await h.settle()
  const viewport = vi.spyOn(h.remote, 'viewport').mockResolvedValue(h.reply)
  const current = h.remote.mainLedger.bytes, headroom = 1024 * 1024
  h.remote.mainLedger.set('copy', 'pressure', 64 * 1024 * 1024 - current - headroom)
  try {
    h.plans.get(h.owner, h.window, h.extent, true)
    await h.settle()
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('detail')
  } finally {
    h.remote.mainLedger.release('copy', 'pressure')
    viewport.mockRestore()
    await h.close()
  }
})

it.each(['reply-doc', 'proof-window', 'merge-after-stage', 'epoch-after-stage', 'abort-after-stage'] as const)('SPEC guard %s preserves old rows and raw ownership', async mode => {
  const h = await bridgeHarness()
  h.remote.merges('s')
  await h.settle()
  const cells = h.reply.result.content.kind === 'detail' ? h.reply.result.content.cells : []
  h.store.adoptViewport(cells)
  const old = h.store.cellsOf(0), before = h.remote.mainLedger.snapshot()
  if (mode === 'reply-doc') h.reply.owner.docId = 'wrong-doc'
  if (mode === 'proof-window') (h.reply.result as OverflowViewportResult).overflowProof!.window.box.width += 64
  const realStage = h.store.stageViewport.bind(h.store)
  const stage = vi.spyOn(h.store, 'stageViewport').mockImplementation((cells, signal, request) => {
    const token = realStage(cells, signal, request)
    if (mode === 'merge-after-stage') h.remote.merges('s').gen++
    if (mode === 'epoch-after-stage') h.store.clear()
    if (mode === 'abort-after-stage') h.plans.dispose()
    return token
  })
  const viewport = vi.spyOn(h.remote, 'viewport').mockResolvedValue(h.reply), settled = vi.spyOn(h.remote.mainLedger, 'settle')
  try {
    h.plans.get(h.owner, h.window, h.extent, true)
    await h.settle()
    expect(settled).not.toHaveBeenCalledWith(h.reply)
    expect(h.remote.mainLedger.snapshot().byPool.replies).toBe(before.byPool.replies)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
    if (mode !== 'epoch-after-stage') expect(h.store.cellsOf(0)).toBe(old)
    else expect(h.store.cellsOf(0)).toBeUndefined()
  } finally {
    stage.mockRestore()
    viewport.mockRestore()
    await h.close()
  }
})

it('SPEC measurement is admitted before reading nested proof and joint refusal precedes allocating payload walks', async () => {
  const h = await bridgeHarness()
  h.remote.merges('s'); await h.settle()
  const proof = (h.reply.result as OverflowViewportResult).overflowProof!, entries = proof.entries
  const admission = viewportPlanAdmission(h.reply, h.reply.result, h.remote.formatting().styles, h.store.viewportAdoptionBytes(h.reply.result.content.kind === 'detail' ? h.reply.result.content.cells : []), h.store.viewportRollbackBytes(), 1, WORKER_QUERY_BYTES.D)
  const observed = vi.fn(() => {
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBeGreaterThanOrEqual(VIEWPORT_MEASUREMENT_BYTES)
    return entries
  })
  Object.defineProperty(proof, 'entries', { enumerable: true, get: observed })
  observed.mockClear()
  const viewport = vi.spyOn(h.remote, 'viewport').mockResolvedValue(h.reply), walk = vi.spyOn(queryBudget, 'queryPayloadBytes'), stage = vi.spyOn(h.store, 'stageViewport'), copies = vi.spyOn(h.remote, 'styleOf')
  h.remote.mainLedger.set('copy', 'measurement-pressure', 64 * 1024 * 1024 - h.remote.mainLedger.bytes - admission.joint + 1)
  try {
    h.plans.get(h.owner, h.window, h.extent, true); await h.settle()
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('diagnostic')
    expect(observed).toHaveBeenCalled()
    expect(walk).not.toHaveBeenCalled()
    expect(copies).not.toHaveBeenCalled()
    expect(stage).not.toHaveBeenCalled()
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
  } finally { h.remote.mainLedger.release('copy', 'measurement-pressure'); walk.mockRestore(); copies.mockRestore(); stage.mockRestore(); viewport.mockRestore(); await h.close() }
})

it('SPEC row-ledger rollback restores LRU order after real replacement mutation', () => {
  const ledger = new RowCacheLedger(100), dropped: number[] = []
  const store = { sheetId: 's', drop: (row: number) => dropped.push(row) } as unknown as ReturnType<RemoteWorkbook['store']>
  ledger.charge(store, 0, 32); ledger.charge(store, 1, 32); ledger.charge(store, 2, 32)
  const transaction = ledger.beginViewportPublication()
  ledger.charge(store, 0, 16)
  transaction.rollback()
  expect(ledger.bytes).toBe(96)
  ledger.charge(store, 3, 32)
  expect(dropped).toEqual([0])
  expect(ledger.bytes).toBe(96)
})

it('SPEC scalar measurement bounds aliases, full backing buffers and cyclic payloads', () => {
  const alias = { value: 'abc' }, value = { a: alias, b: alias, view: new Uint8Array(new ArrayBuffer(100), 0, 1) }
  expect(measureViewportPayload(value, 10000)).toBeGreaterThanOrEqual(queryPayloadBytes(value, 10000))
  const cycle: { next?: unknown } = {}; cycle.next = cycle
  expect(measureViewportPayload(cycle, 10000)).toBeGreaterThan(10000)
})


it('SPEC admission keeps potentially retired plan ownership through restoration', async () => {
  const h = await bridgeHarness()
  try {
    const result = h.reply.result, cells = result.content.kind === 'detail' ? result.content.cells : [], adoption = h.store.viewportAdoptionBytes(cells)
    const base = viewportPlanAdmission(h.reply, result, h.remote.formatting().styles, adoption, h.store.viewportRollbackBytes(), 1, WORKER_QUERY_BYTES.D)
    const full = viewportPlanAdmission(h.reply, result, h.remote.formatting().styles, adoption, h.store.viewportRollbackBytes(), 1, WORKER_QUERY_BYTES.D, WORKER_QUERY_BYTES.D)
    expect(full.reserve).toBeGreaterThanOrEqual(base.reserve + WORKER_QUERY_BYTES.D)
    const small = viewportPlanAdmission(h.reply, result, h.remote.formatting().styles, adoption, h.store.viewportRollbackBytes(), 1, WORKER_QUERY_BYTES.D, 884)
    expect(small.reserve).toBe(base.reserve)
  } finally { await h.close() }
})


interface Slice3Cell { row: number; col: number; value: string; horizontal?: 'left' | 'right' | 'center' }
interface Slice4Options {
  columns?: XlsxColumnSpec[]
  merges?: XlsxMergeRange[]
  widths?: AxisSizeInterval[]
  sort?: SortSpec
  zoom?: number
  rowHeight?: number
  formatBlankCols?: readonly number[]
}
type Slice4PaintExpectation = { kind: 'complete' } | { kind: 'pending'; bounds: LogicalViewport['box'] | LogicalViewport['box'][]; compareWidth: number }
async function slice3Harness(input: readonly Slice3Cell[], options: Slice4Options = {}) {
  const doc = workbookOf([{ sheetId: 's', cells: [...input] }]), sheet = doc.sheets[0]
  sheet.cols = options.columns ?? []
  sheet.mergeRanges = options.merges ?? []
  if (options.rowHeight !== undefined) for (const row of sheet.rows) { row.customHeight = true; row.heightPt = options.rowHeight * 72 / 96 }
  const styles: XlsxCellStyle[] = [{ numFmtId: 0 }]
  for (const col of options.formatBlankCols ?? []) {
    const cell = sheet.rows[0].cells.find(cell => cell.col === col)!
    cell.value = null
    cell.styleIndex = styles.length
    cell.style = { numFmtId: 0, fillColor: 'FFFF00' }
    styles.push(cell.style)
  }
  for (const item of input) if (item.horizontal) {
    const cell = sheet.rows.find(row => row.index === item.row)!.cells.find(cell => cell.col === item.col)!
    cell.styleIndex = styles.length
    cell.style = { numFmtId: 0, horizontal: item.horizontal }
    styles.push(cell.style)
  }
  const objectSource = createObjectWorkbookSource(doc, 'slice3')
  const source = { ...objectSource, formatting: () => ({ dateSystem: '1900' as const, styles }), geometry: (sheetId: string) => {
    const base = objectSource.geometry(sheetId)
    return options.widths ? { ...base, cols: createAxisMetrics({ logicalCount: 16384, defaultSizeCss: 64, overrides: options.widths }), colOverrides: options.widths } : base
  } }
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const port = connectInProcess(kernel), client = new ViewerClient(port)
  const opened = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'slice3', source: { kind: 'object', document: doc }, profile: 'D' })
  if (opened.kind !== 'complete' || opened.result.kind !== 'open') throw new Error('Expected real open metadata')
  const session = kernel.session('slice3')!
  const remote = new RemoteWorkbook(client, { docId: 'slice3', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: opened.result.formatting, layouts: new Map([['s', session.layout('s')]]) })
  if (options.sort) expect((await remote.sort(session.identity('s'), options.sort, new AbortController().signal)).kind).toBe('committed')
  const owner = session.identity('s')
  const flush = async () => { await port.idle(); for (let i = 0; i < 30; i++) await Promise.resolve() }
  remote.merges('s'); await flush()
  const originalRows = source.store('s').readRows.bind(source.store('s'))
  const plans = new WorkerViewportPlans(remote), store = remote.store('s'), requests = vi.spyOn(remote, 'viewport'), reads = vi.spyOn(source.store('s'), 'readRows'), singles = vi.spyOn(source.store('s'), 'readCell')
  const residentGeometry = buildSheetGeometry(sheet, options.sort ? remote.view('s') : undefined)
  const axes = source.geometry('s')
  const geometry = residentGeometry.kind === 'ready' && options.widths ? { ...residentGeometry, metrics: createSheetMetrics({ rows: axes.rows, cols: axes.cols, ...(options.sort ? remote.view('s') : {}) }) } : residentGeometry
  if (geometry.kind !== 'ready') throw new Error('Expected resident geometry')
  const extent = { width: geometry.metrics.cols.totalCss(), height: geometry.metrics.rows.totalCss() }
  const zoom = options.zoom ?? 1
  const get = tilePlanSource({ store: { coverage: () => 'oversized' } as unknown as typeof store, metrics: geometry.metrics, scale: zoom, identity: owner, extent, plans })
  const layout = new LayoutCache(), deterministic = new LayoutCache()
  const makeTile = () => {
    const canvas = createCanvas(512, 512)
    return { canvas: canvas as unknown as HTMLCanvasElement, width: 512, height: 512, close: () => { canvas.width = 0; canvas.height = 0 } }
  }
  const paint = async (x = 0, y = 0, width = 512, height = 20, expectation: Slice4PaintExpectation = { kind: 'complete' }) => {
    const compareWidth = expectation.kind === 'pending' ? expectation.compareWidth : width
    const target = createCanvas(width, height), reference = createCanvas(width, height)
    const prepared = new SheetTileRenderer('D', 0, makeTile), resident = new SheetTileRenderer('D', 0, makeTile)
    const windows: LogicalViewport[] = []
    const base: SheetFrame = { sheet, geometry, identity: owner, origin: { x, y }, viewport: { width, height }, zoom, dpr: 1, extent, paletteKey: '', layoutCache: layout, merges: new MergeQueryCache() }
    const frame: SheetFrame = { ...base, plan: (window, visible) => { windows.push(window); return get(window, visible) } }
    plans.beginFrame(owner, { box: { x, y, width: width / zoom, height: height / zoom }, zoom, deviceScale: 1 })
    try {
      const before = requests.mock.calls.length
      prepared.present(target.getContext('2d') as unknown as CanvasRenderingContext2D, frame)
      if (requests.mock.calls.length > before) expect(prepared.cache.stats().tiles).toBe(0)
      await flush()
      prepared.present(target.getContext('2d') as unknown as CanvasRenderingContext2D, frame)
      resident.present(reference.getContext('2d') as unknown as CanvasRenderingContext2D, { ...base, paletteKey: '', layoutCache: new LayoutCache(), merges: new MergeQueryCache() })
      const pixels = target.getContext('2d').getImageData(0, 0, width, height).data
      expect(target.getContext('2d').getImageData(0, 0, compareWidth, height).data).toEqual(reference.getContext('2d').getImageData(0, 0, compareWidth, height).data)
      const window = windows[0]
      const plan = get(window, true)
      expect(plan?.kind).toBe('detail')
      if (plan?.kind !== 'detail') throw new Error('Expected real detail plan')
      if (expectation.kind === 'pending') {
        expect(plan.pending).toEqual((Array.isArray(expectation.bounds) ? expectation.bounds : [expectation.bounds]).map(bounds => ({ bounds, error: false })))
        expect(prepared.cache.stats().tiles).toBe(0)
      } else {
        expect(prepared.cache.stats().tiles).toBeGreaterThan(0)
        expect(plan.pending).toBeUndefined()
      }
      const fake = createCanvas(512, 512).getContext('2d')
      fake.measureText = (text: string) => ({ width: text.length * 7 }) as ReturnType<typeof fake.measureText>
      const out = renderSheetViewport(sheet, fake as unknown as CanvasRenderingContext2D, undefined, window, undefined, { owner, preparedPlan: plan, cache: deterministic })
      const canonical = createCanvas(512, 512).getContext('2d')
      canonical.measureText = (text: string) => ({ width: text.length * 7 }) as ReturnType<typeof canonical.measureText>
      const expected = renderSheetViewport(sheet, canonical as unknown as CanvasRenderingContext2D, geometry, window, undefined, { owner, cache: new LayoutCache() })
      return { pixels, residentPixels: reference.getContext('2d').getImageData(0, 0, width, height).data, plan, window, records: out.records, residentRecords: expected.records, paintPending: out.pending }
    } finally { prepared.dispose(); resident.dispose() }
  }
  return { source, originalRows, remote, store, plans, requests, reads, singles, paint, layout, deterministic, flush, get, owner, extent, close: async () => { requests.mockRestore(); reads.mockRestore(); singles.mockRestore(); plans.dispose(); await remote.release() } }
}
function sourceRecord(records: readonly CellLayout[], row: number, col: number): CellLayout {
  const record = records.find(record => record.source.row === row && record.source.col === col)
  if (!record) throw new Error(`Missing record ${row}:${col}`)
  return record
}
function plainClip(records: readonly CellLayout[], row: number, col: number) {
  const text = sourceRecord(records, row, col).body.text
  expect(text?.kind).toBe('plain')
  if (text?.kind !== 'plain') throw new Error('Expected plain cell text')
  return text.clipX
}
const SLICE3_LONG = 'abcdefghijklmnopqrstuvwxABCDEFGHIJKLMNOP'

it('SLICE3 right-aligned C1 clips at occupied B1 through the real prepared tile crop', async () => {
  const h = await slice3Harness([{ row: 0, col: 1, value: 'BLOCKER' }, { row: 0, col: 2, value: SLICE3_LONG, horizontal: 'right' }])
  try {
    const out = await h.paint()
    const text = sourceRecord(out.records, 0, 2).body.text, resident = sourceRecord(out.residentRecords, 0, 2).body.text
    expect(text).toEqual(resident)
    expect(text?.kind === 'layout' && text.clip).toMatchObject({ x: 0, width: 64 })
    expect(h.requests).toHaveBeenCalledTimes(1)
    expect(h.reads).toHaveBeenCalledTimes(1)
    expect(h.singles).not.toHaveBeenCalled()
  } finally { await h.close() }
})

it('SLICE3 centered C1 uses both blank corridors bounded by A1 and E1 in the actual tile crop', async () => {
  const h = await slice3Harness([{ row: 0, col: 0, value: 'LEFT' }, { row: 0, col: 2, value: SLICE3_LONG, horizontal: 'center' }, { row: 0, col: 4, value: 'RIGHT' }])
  try {
    const out = await h.paint()
    const text = sourceRecord(out.records, 0, 2).body.text, resident = sourceRecord(out.residentRecords, 0, 2).body.text
    expect(text).toEqual(resident)
    expect(text?.kind === 'layout' && text.clip).toMatchObject({ x: -64, width: 192 })
    expect(h.requests).toHaveBeenCalledTimes(1)
    expect(h.reads).toHaveBeenCalledTimes(1)
  } finally { await h.close() }
})


function hasInk(pixels: Uint8ClampedArray, width: number, x0: number, x1: number): boolean {
  for (let y = 2; y < 18; y++) for (let x = x0; x < x1; x++) {
    const offset = (y * width + x) * 4
    if (pixels[offset] < 128 && pixels[offset + 1] < 128 && pixels[offset + 2] < 128 && pixels[offset + 3] > 0) return true
  }
  return false
}
it('SLICE3 admitted origins on both sides enter the tile only across their certified blank corridors', async () => {
  const h = await slice3Harness([{ row: 0, col: 7, value: SLICE3_LONG }, { row: 0, col: 10, value: 'L' }, { row: 0, col: 14, value: 'R' }, { row: 0, col: 16, value: SLICE3_LONG, horizontal: 'right' }])
  try {
    const out = await h.paint(512)
    expect(out.window.box).toMatchObject({ x: 512, width: 512 })
    for (const col of [7, 16]) expect(sourceRecord(out.records, 0, col).body.text).toEqual(sourceRecord(out.residentRecords, 0, col).body.text)
    expect(out.plan.cells.find(cell => cell.source.col === 7)?.overflow?.right).toBe(128)
    expect(out.plan.cells.find(cell => cell.source.col === 16)?.overflow?.left).toBe(64)
    expect(hasInk(out.pixels, 512, 5, 120), 'left origin must paint inside the tile').toBe(true)
    expect(hasInk(out.pixels, 512, 453, 505), 'right origin must paint inside the tile').toBe(true)
    expect(h.requests).toHaveBeenCalledTimes(1)
    expect(h.reads).toHaveBeenCalledTimes(1)
    expect(h.reads.mock.calls[0][0]).toMatchObject({ c0: 0, c1: 47 })
    expect(h.singles).not.toHaveBeenCalled()
  } finally { await h.close() }
})

it.each(['left', 'right'] as const)('SLICE3 a blocker beyond the %s tile horizon cannot change the requested crop', async side => {
  let previous: Uint8ClampedArray | undefined
  for (const blocked of [false, true]) {
    const origin = side === 'left' ? 71 : 64
    const cells: Slice3Cell[] = [{ row: 0, col: origin, value: 'M'.repeat(500), ...(side === 'left' ? { horizontal: 'right' as const } : {}) }]
    if (blocked) cells.push({ row: 0, col: side === 'left' ? 31 : 104, value: 'OUTSIDE HORIZON' })
    const h = await slice3Harness(cells)
    try {
      const out = await h.paint(4096)
      expect(out.plan.cells.map(cell => cell.source.col)).toEqual([origin])
      const reply = await h.requests.mock.results[0].value as Awaited<ReturnType<RemoteWorkbook['viewport']>>
      if (typeof reply !== 'object' || reply.kind !== 'complete') throw new Error('Expected transported viewport')
      const proof = (reply.result as OverflowViewportResult).overflowProof!
      expect(proof.entries[0]).toMatchObject({ h0: 2048, h1: 6656, c0: 32, c1: 103 })
      expect(proof.entries[0][side === 'left' ? 'leftEnd' : 'rightEnd']).toBe('paint-boundary')
      if (previous) expect(out.pixels).toEqual(previous)
      previous = out.pixels
      expect(hasInk(out.pixels, 512, 4, 500)).toBe(true)
      expect(h.reads).toHaveBeenCalledTimes(1)
      expect(h.reads.mock.calls[0][0]).toMatchObject({ c0: 32, c1: 103 })
      expect(h.singles).not.toHaveBeenCalled()
    } finally { await h.close() }
  }
})

it('SLICE3 panning the same origin rebuilds its allowance key and returning reuses the original body', async () => {
  const h = await slice3Harness([{ row: 0, col: 7, value: 'M'.repeat(400) }])
  try {
    const first = await h.paint(0), firstBody = sourceRecord(first.records, 0, 7).body
    const built = h.deterministic.stats.built
    const second = await h.paint(512), secondBody = sourceRecord(second.records, 0, 7).body
    expect(first.plan.cells[0].overflow?.right).toBe(2048)
    expect(second.plan.cells[0].overflow?.right).toBe(2560)
    expect(secondBody).not.toBe(firstBody)
    expect(h.deterministic.stats.built).toBe(built + 1)
    expect(firstBody.text?.kind === 'plain' && firstBody.text.clipX?.x1).toBe(2112)
    expect(secondBody.text?.kind === 'plain' && secondBody.text.clipX?.x1).toBe(2624)
    const reused = h.deterministic.stats.reused
    const again = await h.paint(0)
    expect(sourceRecord(again.records, 0, 7).body).toBe(firstBody)
    expect(h.deterministic.stats.built).toBe(built + 1)
    expect(h.deterministic.stats.reused).toBeGreaterThan(reused)
    expect(again.pixels).toEqual(first.pixels)
    expect(h.requests).toHaveBeenCalledTimes(2)
    expect(h.reads).toHaveBeenCalledTimes(2)
    expect(h.singles).not.toHaveBeenCalled()
  } finally { await h.close() }
})

it('SLICE3 A and XFD worksheet edges remain clipped at their real arithmetic source locations', async () => {
  const maxRow = 1048575, maxCol = 16383
  const h = await slice3Harness([{ row: 0, col: 0, value: SLICE3_LONG, horizontal: 'right' }, { row: maxRow, col: maxCol, value: SLICE3_LONG }])
  try {
    const first = await h.paint(0, 0, 64)
    const a = sourceRecord(first.records, 0, 0)
    expect(a.id).toBe(0)
    expect(a.text[0].clip.x).toBeGreaterThanOrEqual(0)
    const corner = await h.paint(maxCol * 64, maxRow * 20, 64)
    const record = sourceRecord(corner.records, maxRow, maxCol)
    expect(record.id).toBe(17179869183)
    expect(record.body.text?.kind === 'plain' && record.body.text.clipX?.x1).toBe(64)
    expect(record.text[0].clip.x + record.text[0].clip.width).toBeLessThanOrEqual(16384 * 64)
    expect(record.text[0].clip.y + record.text[0].clip.height).toBeLessThanOrEqual(1048576 * 20)
    expect(h.store.cellAt({ row: maxRow, col: maxCol })?.id).toBe(17179869183)
    expect(h.store.cellAt({ row: 0, col: 0 })?.id).toBe(0)
    expect(h.reads).toHaveBeenCalledTimes(2)
    for (const [range] of h.reads.mock.calls) { expect(range.r1 - range.r0).toBeLessThanOrEqual(26); expect(range.c1 - range.c0).toBeLessThanOrEqual(71) }
    expect(h.singles).not.toHaveBeenCalled()
  } finally { await h.close() }
})

it('SLICE3 source IDs at 2^31 and 2^32 do not alias A1 or the maximum corner', async () => {
  const rows = [0, 131072, 262144]
  const h = await slice3Harness([...rows.map(row => ({ row, col: 0, value: `row-${row}` })), { row: 1048575, col: 16383, value: 'maximum-corner' }])
  try {
    const ids: number[] = []
    for (const row of rows) {
      const out = await h.paint(0, row * 20, 64)
      const record = sourceRecord(out.records, row, 0)
      ids.push(record.id)
      expect(record.id).toBe(row * 16384)
      expect(record.body.formattedText).toBe(`row-${row}`)
    }
    const corner = await h.paint(16383 * 64, 1048575 * 20, 64)
    ids.push(sourceRecord(corner.records, 1048575, 16383).id)
    expect(ids).toEqual([0, 2147483648, 4294967296, 17179869183])
    expect(new Set(ids).size).toBe(4)
    for (const row of rows) expect(h.store.cellAt({ row, col: 0 })?.value).toBe(`row-${row}`)
    expect(h.reads).toHaveBeenCalledTimes(4)
    expect(h.requests).toHaveBeenCalledTimes(4)
    expect(h.singles).not.toHaveBeenCalled()
  } finally { await h.close() }
})


it.each([false, true])('SLICE4 hidden valued B1 adds no width or blocker; visible C1 blocked=%s', async blocked => {
  const h = await slice3Harness([{ row: 0, col: 0, value: SLICE3_LONG }, { row: 0, col: 1, value: 'HIDDEN' }, ...(blocked ? [{ row: 0, col: 2, value: 'C' }] : [])], { columns: [{ min: 1, max: 1, hidden: true }] })
  try {
    const out = await h.paint()
    const a = out.plan.cells.find(cell => cell.source.col === 0)!
    expect(out.plan.cells.some(cell => cell.source.col === 1)).toBe(false)
    expect(a.width).toBe(64)
    expect(a.overflow?.right).toBe(blocked ? 0 : 2496)
    expect(sourceRecord(out.records, 0, 0).body.text).toEqual(sourceRecord(out.residentRecords, 0, 0).body.text)
    expect(plainClip(out.records, 0, 0)).toEqual(blocked ? { x0: 0, x1: 64 } : undefined)
    expect(hasInk(out.pixels, 512, 85, 125)).toBe(!blocked)
    expect(h.requests).toHaveBeenCalledTimes(1)
    expect(h.reads).toHaveBeenCalledTimes(1)
    expect(h.singles).not.toHaveBeenCalled()
  } finally { await h.close() }
})


it.each([false, true])('SLICE4 positive one-pixel B1 occupancy blocks=%s while its coarse group persists', async blocked => {
  const h = await slice3Harness([{ row: 0, col: 0, value: SLICE3_LONG }, ...(blocked ? [{ row: 0, col: 1, value: 'SUBDETAIL' }] : []), { row: 0, col: 2, value: 'C' }], { widths: [{ start: 1, end: 1, sizeCss: 1 }] })
  try {
    const out = await h.paint()
    expect(out.plan.cells.some(cell => cell.source.col === 1)).toBe(false)
    expect(out.records.some(record => record.source.col === 1)).toBe(false)
    expect(out.plan.coarse).toHaveLength(128)
    for (const group of out.plan.coarse) expect(group.bounds).toMatchObject({ x: 64, width: 1, height: 4 })
    expect(out.plan.cells.find(cell => cell.source.col === 0)?.overflow?.right).toBe(blocked ? 0 : 1)
    expect(plainClip(out.records, 0, 0)).toEqual({ x0: 0, x1: blocked ? 64 : 65 })
    expect(h.requests).toHaveBeenCalledTimes(1)
    expect(h.reads).toHaveBeenCalledTimes(1)
    expect(h.reads.mock.calls[0][0]).toMatchObject({ c0: 0, c1: 40 })
    expect(h.singles).not.toHaveBeenCalled()
  } finally { await h.close() }
})

it('SLICE4 a dependency origin behind a one-pixel run stays unadmitted', async () => {
  const h = await slice3Harness([{ row: 0, col: 0, value: SLICE3_LONG }, { row: 0, col: 2, value: 'VISIBLE' }], { widths: [{ start: 1, end: 1, sizeCss: 1 }] })
  try {
    const out = await h.paint(512)
    expect(out.plan.cells.map(cell => cell.source.col)).toEqual([2])
    expect(out.records.map(record => record.source.col)).toEqual([2])
    expect(hasInk(out.pixels, 512, 5, 100)).toBe(false)
    expect(h.reads).toHaveBeenCalledTimes(1)
    expect(h.reads.mock.calls[0][0].c0).toBe(2)
    expect(h.singles).not.toHaveBeenCalled()
  } finally { await h.close() }
})


it.each([false, true])('SLICE4 B1:C2 merge blocks A2 even with empty anchor=%s and merged anchors never overflow', async empty => {
  const h = await slice3Harness([{ row: 1, col: 0, value: SLICE3_LONG }, ...(!empty ? [{ row: 0, col: 1, value: 'M'.repeat(80) }] : [])], { merges: [{ minRow: 0, maxRow: 1, minCol: 1, maxCol: 2 }] })
  try {
    const out = await h.paint(0, 0, 512, 40)
    expect(out.plan.merges).toHaveLength(1)
    expect(out.plan.cells.find(cell => cell.source.row === 1 && cell.source.col === 0)?.overflow).toEqual({ left: 0, right: 0 })
    expect(plainClip(out.records, 1, 0)).toEqual({ x0: 0, x1: 64 })
    expect(sourceRecord(out.records, 1, 0).body.text).toEqual(sourceRecord(out.residentRecords, 1, 0).body.text)
    if (empty) {
      expect(out.plan.cells.map(cell => cell.source)).toEqual([{ row: 1, col: 0 }])
      expect(h.singles).toHaveBeenCalledTimes(1)
    } else {
      const anchor = out.plan.cells.find(cell => cell.source.row === 0 && cell.source.col === 1)!
      expect(anchor).toMatchObject({ x: 64, y: 0, width: 128, height: 40, overflow: { left: 0, right: 0 } })
      expect(sourceRecord(out.records, 0, 1).body.text).toEqual(sourceRecord(out.residentRecords, 0, 1).body.text)
      expect(sourceRecord(out.records, 0, 1).text[0].clip.x + sourceRecord(out.records, 0, 1).text[0].clip.width).toBeLessThanOrEqual(192)
      expect(h.singles).not.toHaveBeenCalled()
    }
    expect(h.reads).toHaveBeenCalledTimes(1)
  } finally { await h.close() }
})

it('SLICE4 a merge anchored above the viewport still blocks its selected intersecting row', async () => {
  const h = await slice3Harness([{ row: 25, col: 0, value: SLICE3_LONG }, { row: 24, col: 1, value: 'OFFSCREEN MERGE ANCHOR '.repeat(8) }], { merges: [{ minRow: 24, maxRow: 25, minCol: 1, maxCol: 2 }] })
  try {
    const out = await h.paint(0, 512, 512, 20)
    expect(out.window.box.y).toBe(512)
    expect(h.reads.mock.calls[0][0]).toMatchObject({ r0: 25, r1: 51 })
    expect(out.plan.cells.find(cell => cell.source.row === 25)?.overflow?.right).toBe(0)
    expect(out.plan.cells.find(cell => cell.source.row === 24)).toMatchObject({ x: 64, y: 480, width: 128, height: 40, overflow: { left: 0, right: 0 } })
    expect(sourceRecord(out.records, 25, 0).body.text).toEqual(sourceRecord(out.residentRecords, 25, 0).body.text)
    expect(h.reads).toHaveBeenCalledTimes(1)
    expect(h.singles).toHaveBeenCalledTimes(1)
    expect(h.singles.mock.calls[0][0]).toEqual({ row: 24, col: 1 })
  } finally { await h.close() }
})

it.each([false, true])('SLICE4 hidden merge span fullyHidden=%s uses its positive visible geometry', async fullyHidden => {
  const h = await slice3Harness([{ row: 1, col: 0, value: SLICE3_LONG }], { columns: [{ min: 1, max: fullyHidden ? 2 : 1, hidden: true }], merges: [{ minRow: 0, maxRow: 1, minCol: 1, maxCol: 2 }] })
  try {
    const out = await h.paint(0, 0, 512, 40)
    expect(out.plan.cells[0].overflow?.right).toBe(fullyHidden ? 2496 : 0)
    expect(sourceRecord(out.records, 1, 0).body.text).toEqual(sourceRecord(out.residentRecords, 1, 0).body.text)
    expect(out.plan.merges).toHaveLength(fullyHidden ? 0 : 1)
    if (!fullyHidden) expect(out.plan.merges[0]).toMatchObject({ x: 64, width: 64 })
    expect(h.reads).toHaveBeenCalledTimes(1)
    expect(h.singles).toHaveBeenCalledTimes(fullyHidden ? 0 : 1)
  } finally { await h.close() }
})


it('SLICE4 page-budget-denied offscreen anchor keeps its merge blocking known A26 with no extra anchor read', async () => {
  const h = await slice3Harness([{ row: 25, col: 0, value: SLICE3_LONG }, { row: 24, col: 1, value: 'UNREAD ANCHOR' }], { merges: [{ minRow: 24, maxRow: 25, minCol: 1, maxCol: 2 }] })
  let pages = 0
  h.reads.mockImplementation(async (range, _cursor, options) => {
    const page = await h.originalRows(range, undefined, options)
    pages++
    return { ...page, items: pages === 1 ? page.items : [], nextCursor: pages < 8 ? `slice4:${pages}` : undefined }
  })
  try {
    // Compare the known A-column native crop. The unread merge legitimately
    // carries pending shading in B:C, so full completed-resident equality is not claimed.
    const out = await h.paint(0, 512, 512, 20, { kind: 'pending', bounds: { x: 64, y: 480, width: 128, height: 40 }, compareWidth: 64 })
    expect(out.plan.cells.map(cell => cell.source)).toEqual([{ row: 25, col: 0 }])
    expect(out.plan.cells[0].overflow?.right).toBe(0)
    expect(plainClip(out.records, 25, 0)).toEqual({ x0: 0, x1: 64 })
    expect(sourceRecord(out.records, 25, 0).body.text).toEqual(sourceRecord(out.residentRecords, 25, 0).body.text)
    expect(out.plan.merges).toHaveLength(1)
    expect(out.plan.pending).toEqual([{ bounds: { x: 64, y: 480, width: 128, height: 40 }, error: false }])
    expect(out.plan.work.pages).toBe(8)
    expect(h.requests).toHaveBeenCalledTimes(1)
    expect(h.reads).toHaveBeenCalledTimes(8)
    for (const [range] of h.reads.mock.calls) expect(range).toEqual({ r0: 25, r1: 51, c0: 0, c1: 39 })
    expect(h.singles).not.toHaveBeenCalled()
  } finally { await h.close() }
})

it('SLICE4 occupancy and placement follow cell.row when the sparse page groups it under another row', async () => {
  const h = await slice3Harness([{ row: 0, col: 0, value: SLICE3_LONG }, { row: 1, col: 0, value: SLICE3_LONG }, { row: 1, col: 1, value: 'B' }])
  h.reads.mockImplementation(async (range, cursor, options) => {
    const page = await h.originalRows(range, cursor, options)
    return { ...page, items: [{ row: 0, cells: page.items.flatMap(item => item.cells) }] }
  })
  try {
    const out = await h.paint(0, 0, 512, 40)
    expect(out.plan.cells.map(cell => cell.source)).toEqual([{ row: 0, col: 0 }, { row: 1, col: 0 }, { row: 1, col: 1 }])
    expect(out.plan.cells[0]).toMatchObject({ y: 0, overflow: { right: 2496 } })
    expect(out.plan.cells[1]).toMatchObject({ y: 20, overflow: { right: 0 } })
    for (const row of [0, 1]) expect(sourceRecord(out.records, row, 0).body.text).toEqual(sourceRecord(out.residentRecords, row, 0).body.text)
    expect(h.reads).toHaveBeenCalledTimes(1)
    expect(h.singles).not.toHaveBeenCalled()
  } finally { await h.close() }
})

it('SLICE4 reordered source rows keep their own blockers and use current display placements', async () => {
  const h = await slice3Harness([{ row: 0, col: 0, value: 'b'.repeat(40) }, { row: 1, col: 0, value: 'a'.repeat(40) }, { row: 1, col: 1, value: 'B' }], { sort: { range: { r0: 0, r1: 1, c0: 0, c1: 1, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false } })
  try {
    const out = await h.paint(0, 0, 512, 40)
    expect(out.plan.cells.map(cell => cell.source)).toEqual([{ row: 1, col: 0 }, { row: 1, col: 1 }, { row: 0, col: 0 }])
    expect(out.plan.cells[0]).toMatchObject({ y: 0, overflow: { right: 0 } })
    expect(out.plan.cells[2]).toMatchObject({ y: 20, overflow: { right: 2496 } })
    for (const row of [0, 1]) expect(sourceRecord(out.records, row, 0).body.text).toEqual(sourceRecord(out.residentRecords, row, 0).body.text)
    expect(h.reads).toHaveBeenCalledTimes(1)
    expect(h.reads.mock.calls[0][0]).toEqual({ r0: 0, r1: 25, c0: 0, c1: 39 })
    expect(h.singles).not.toHaveBeenCalled()
  } finally { await h.close() }
})


it.each([false, true])('SLICE4 one-device-pixel B1 at half zoom admits A1 ink only for known blank blocked=%s', async blocked => {
  const cells = [...(blocked ? [{ row: 0, col: 1, value: 'SUBDETAIL' }] : []), { row: 0, col: 2, value: 'C' }]
  const options: Slice4Options = { widths: [{ start: 1, end: 1, sizeCss: 2 }], zoom: 0.5 }
  const h = await slice3Harness([{ row: 0, col: 0, value: 'M'.repeat(80) }, ...cells], options)
  const noOrigin = await slice3Harness(cells, options)
  try {
    const out = await h.paint(), control = await noOrigin.paint()
    expect(out.window.zoom * out.window.deviceScale * 2).toBe(1)
    expect(out.plan.cells.some(cell => cell.source.col === 1)).toBe(false)
    expect(out.records.some(record => record.source.col === 1)).toBe(false)
    expect(out.plan.cells.find(cell => cell.source.col === 0)?.overflow?.right).toBe(blocked ? 0 : 2)
    expect(plainClip(out.records, 0, 0)).toEqual({ x0: 0, x1: blocked ? 64 : 66 })
    expect(out.plan.coarse).toHaveLength(128)
    expect(control.plan.coarse).toHaveLength(128)
    for (const group of out.plan.coarse) expect(group.bounds).toMatchObject({ x: 64, width: 2, height: 8 })
    let changed = false
    for (let y = 1; y < 9; y++) for (let channel = 0; channel < 4; channel++) {
      const offset = (y * 512 + 32) * 4 + channel
      if (out.pixels[offset] !== control.pixels[offset]) changed = true
    }
    // Removing only A1 isolates its ink from grid/coarse paint. At half zoom
    // the gridline does not completely cover B1's one-device-pixel width.
    expect(changed, 'only known blank B1 admits A1 ink in its physical pixel').toBe(!blocked)
    for (const subject of [h, noOrigin]) {
      expect(subject.requests).toHaveBeenCalledTimes(1)
      expect(subject.reads).toHaveBeenCalledTimes(1)
      expect(subject.singles).not.toHaveBeenCalled()
    }
  } finally { await noOrigin.close(); await h.close() }
})


function slice5Gate() {
  let enter!: () => void, release!: () => void
  const entered = new Promise<void>(resolve => { enter = resolve })
  const held = new Promise<void>(resolve => { release = resolve })
  return { entered, enter, held, release }
}
const SLICE5_COLUMNS = [{ min: 3, max: 16383, hidden: true }]
function slice5Ink(pixels: Uint8ClampedArray, width: number, x0: number, x1: number, y0: number, y1: number): boolean {
  for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) {
    const offset = (y * width + x) * 4
    if (pixels[offset] < 128 && pixels[offset + 1] < 128 && pixels[offset + 2] < 128 && pixels[offset + 3] > 0) return true
  }
  return false
}
async function slice5Completion(neighbour: 'blocker' | 'blank') {
  const h = await slice3Harness([{ row: 0, col: 0, value: SLICE3_LONG }, ...(neighbour === 'blocker' ? [{ row: 0, col: 1, value: 'BLOCKER' }] : []), { row: 0, col: 2, value: 'C' }], { rowHeight: 512, columns: SLICE5_COLUMNS })
  const gates = [slice5Gate(), slice5Gate()]
  let phase = 0
  h.reads.mockImplementation(async (range, cursor, options): Promise<Page<SparseRowPage>> => {
    const page = await h.originalRows(range, undefined, options)
    if (cursor === undefined) return { ...page, items: page.items.map(item => ({ ...item, cells: item.cells.filter(cell => cell.col === 0) })), completeness: phase === 0 ? 'loading' : 'complete', nextCursor: `slice5:${phase}` }
    expect(cursor).toBe(`slice5:${phase}`)
    gates[phase].enter()
    await gates[phase].held
    return { ...page, items: phase === 0 ? [] : page.items.map(item => ({ ...item, cells: item.cells.filter(cell => cell.col > 0) })), completeness: phase === 0 ? 'loading' : 'complete', nextCursor: undefined }
  })
  vi.useFakeTimers()
  try {
    const pendingPaint = h.paint(0, 0, 192, 512, { kind: 'pending', bounds: { x: 64, y: 0, width: 128, height: 512 }, compareWidth: 64 })
    await gates[0].entered
    expect(h.remote.pagesInFlight).toBe(8)
    expect(h.get({ box: { x: 0, y: 0, width: 512, height: 512 }, zoom: 1, deviceScale: 1 }, true)?.kind).toBe('pending')
    expect(h.store.cellAt({ row: 0, col: 0 })).toBeUndefined()
    gates[0].release()
    const partial = await pendingPaint
    expect(partial.paintPending).toBe(true)
    expect(partial.plan.cells.map(cell => cell.source)).toEqual([{ row: 0, col: 0 }])
    expect(partial.plan.cells[0].overflow).toEqual({ left: 0, right: 0 })
    expect(plainClip(partial.records, 0, 0)).toEqual({ x0: 0, x1: 64 })
    expect(h.store.cellAt({ row: 0, col: 0 })?.value).toBe(SLICE3_LONG)
    expect(h.store.cellAt({ row: 0, col: 1 })).toBeUndefined()
    expect(slice5Ink(partial.pixels, 192, 4, 62, 246, 266), 'known A1 ink must remain paintable').toBe(true)
    expect(slice5Ink(partial.pixels, 192, 65, 126, 246, 266), 'A1 ink cannot cross unknown B1').toBe(false)
    expect(h.requests).toHaveBeenCalledTimes(1)
    expect(h.remote.stats.commands.viewport).toBe(1)
    expect(h.reads).toHaveBeenCalledTimes(2)
    expect(partial.plan.work.pages).toBe(2)
    expect(h.remote.currentPageFrame.used).toBe(8)
    expect(h.remote.pagesInFlight).toBe(0)
    phase = 1
    await vi.advanceTimersByTimeAsync(99)
    expect(h.requests).toHaveBeenCalledTimes(1)
    expect(h.remote.stats.commands.viewport).toBe(1)
    await vi.advanceTimersByTimeAsync(1)
    const completePaint = h.paint(0, 0, 192, 512)
    await gates[1].entered
    expect(h.remote.pagesInFlight).toBe(8)
    expect(h.store.cellAt({ row: 0, col: 1 })).toBeUndefined()
    gates[1].release()
    const complete = await completePaint
    expect(complete.paintPending).toBeFalsy()
    expect(plainClip(complete.records, 0, 0)).toEqual({ x0: 0, x1: neighbour === 'blocker' ? 64 : 128 })
    expect(complete.plan.cells[0].overflow?.right).toBe(neighbour === 'blocker' ? 0 : 64)
    expect(h.store.cellAt({ row: 0, col: 1 })?.value).toBe(neighbour === 'blocker' ? 'BLOCKER' : undefined)
    if (neighbour === 'blank') expect(slice5Ink(complete.pixels, 192, 65, 126, 246, 266), 'completed omitted B1 permits overflow ink').toBe(true)
    expect(h.requests).toHaveBeenCalledTimes(2)
    expect(h.remote.stats.commands.viewport).toBe(2)
    expect(h.reads).toHaveBeenCalledTimes(4)
    expect(complete.plan.work.pages).toBe(2)
    expect(h.remote.currentPageFrame.used).toBe(8)
    expect(h.remote.pagesInFlight).toBe(0)
    for (const [range] of h.reads.mock.calls) expect(range).toEqual({ r0: 0, r1: 0, c0: 0, c1: 2 })
    expect(h.singles).not.toHaveBeenCalled()
  } finally { for (const gate of gates) gate.release(); vi.useRealTimers(); await h.close() }
}
it('SLICE5 held same-row loading cursor keeps A1 inspectable and unretained until actual B1 blocker completion', async () => { await slice5Completion('blocker') })
it('SLICE5 completed omission certifies blank B1 and restores overflow through its width', async () => { await slice5Completion('blank') })

it('SLICE5 complete pages with a live cursor at the tile page cap cannot prove omitted B1 blank', async () => {
  const h = await slice3Harness([{ row: 0, col: 0, value: SLICE3_LONG }, { row: 0, col: 2, value: 'C' }], { rowHeight: 512, columns: SLICE5_COLUMNS })
  const gate = slice5Gate()
  let pages = 0
  h.reads.mockImplementation(async (range, cursor, options): Promise<Page<SparseRowPage>> => {
    const page = await h.originalRows(range, undefined, options)
    expect(cursor).toBe(pages === 0 ? undefined : `slice5:budget:${pages}`)
    pages++
    if (pages === 8) { gate.enter(); await gate.held }
    return { ...page, items: pages === 1 ? page.items.map(item => ({ ...item, cells: item.cells.filter(cell => cell.col === 0) })) : [], completeness: 'complete', nextCursor: `slice5:budget:${pages}` }
  })
  vi.useFakeTimers()
  try {
    const painting = h.paint(0, 0, 192, 512, { kind: 'pending', bounds: { x: 64, y: 0, width: 128, height: 512 }, compareWidth: 64 })
    await gate.entered
    expect(h.remote.pagesInFlight).toBe(8)
    expect(h.reads).toHaveBeenCalledTimes(8)
    gate.release()
    const partial = await painting
    expect(partial.paintPending).toBe(true)
    expect(plainClip(partial.records, 0, 0)).toEqual({ x0: 0, x1: 64 })
    expect(partial.plan.cells[0].overflow).toEqual({ left: 0, right: 0 })
    expect(slice5Ink(partial.pixels, 192, 4, 62, 246, 266)).toBe(true)
    expect(slice5Ink(partial.pixels, 192, 65, 126, 246, 266)).toBe(false)
    const reply = await h.requests.mock.results[0].value as Awaited<ReturnType<RemoteWorkbook['viewport']>>
    if (typeof reply !== 'object' || reply.kind !== 'complete') throw new Error('Expected real transported page-budget result')
    expect(reply.result.completeness).toBe('loading')
    expect((reply.result as OverflowViewportResult).overflowProof?.entries[0].rightEnd).toBe('unknown')
    expect(partial.plan.work.pages).toBe(8)
    expect(h.requests).toHaveBeenCalledTimes(1)
    expect(h.remote.stats.commands.viewport).toBe(1)
    expect(h.reads).toHaveBeenCalledTimes(8)
    for (const [range] of h.reads.mock.calls) expect(range).toEqual({ r0: 0, r1: 0, c0: 0, c1: 2 })
    expect(h.singles).not.toHaveBeenCalled()
    expect(h.remote.currentPageFrame.used).toBe(8)
    expect(h.remote.pageLimits).toEqual({ frame: 32, tile: 8 })
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { gate.release(); vi.useRealTimers(); await h.close() }
})

it('SLICE5 partial formatting-only B1 is known blank but absent C1 stays pending', async () => {
  const h = await slice3Harness([{ row: 0, col: 0, value: SLICE3_LONG }, { row: 0, col: 1, value: '' }, { row: 0, col: 2, value: 'C' }], { rowHeight: 512, columns: SLICE5_COLUMNS, formatBlankCols: [1] })
  h.reads.mockImplementation(async (range, cursor, options): Promise<Page<SparseRowPage>> => {
    const page = await h.originalRows(range, cursor, options)
    return { ...page, completeness: 'loading', items: page.items.map(item => ({ ...item, cells: item.cells.filter(cell => cell.col < 2) })) }
  })
  vi.useFakeTimers()
  try {
    // Both admitted A and formatting-only B identify the same unknown C region.
    const partial = await h.paint(0, 0, 192, 512, { kind: 'pending', bounds: [{ x: 128, y: 0, width: 64, height: 512 }, { x: 128, y: 0, width: 64, height: 512 }], compareWidth: 128 })
    expect(partial.paintPending).toBe(true)
    expect(partial.plan.cells.map(cell => cell.source)).toEqual([{ row: 0, col: 0 }, { row: 0, col: 1 }])
    expect(partial.plan.cells[0].overflow).toEqual({ left: 0, right: 64 })
    expect(plainClip(partial.records, 0, 0)).toEqual({ x0: 0, x1: 128 })
    expect(h.store.cellAt({ row: 0, col: 0 })?.value).toBe(SLICE3_LONG)
    expect(h.store.cellAt({ row: 0, col: 1 })).toMatchObject({ value: null, style: { fillColor: 'FFFF00' } })
    expect(h.store.cellAt({ row: 0, col: 2 })).toBeUndefined()
    expect(slice5Ink(partial.pixels, 192, 65, 126, 246, 266), 'known formatting-only B permits A ink').toBe(true)
    expect(slice5Ink(partial.pixels, 192, 129, 190, 246, 266), 'unknown C cannot receive A ink').toBe(false)
    expect([...partial.pixels.slice((100 * 192 + 100) * 4, (100 * 192 + 100) * 4 + 4)]).toEqual([255, 255, 0, 255])
    expect(partial.plan.work.pages).toBe(1)
    expect(h.requests).toHaveBeenCalledTimes(1)
    expect(h.remote.stats.commands.viewport).toBe(1)
    expect(h.reads).toHaveBeenCalledTimes(1)
    expect(h.reads.mock.calls[0][0]).toEqual({ r0: 0, r1: 0, c0: 0, c1: 2 })
    expect(h.singles).not.toHaveBeenCalled()
    expect(h.remote.currentPageFrame.used).toBe(8)
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { vi.useRealTimers(); await h.close() }
})

it('SLICE5 disjoint sorted source runs preserve complete blank proof beside an incomplete displayed row', async () => {
  const h = await slice3Harness([
    { row: 0, col: 0, value: 'b'.repeat(40) }, { row: 0, col: 1, value: 'BLOCKER' }, { row: 0, col: 2, value: 'C0' },
    { row: 1, col: 0, value: 'z'.repeat(40) },
    { row: 2, col: 0, value: 'a'.repeat(40) }, { row: 2, col: 2, value: 'C2' },
  ], { rowHeight: 256, columns: SLICE5_COLUMNS, sort: { range: { r0: 0, r1: 2, c0: 0, c1: 2, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false } })
  h.reads.mockImplementation(async (range, cursor, options): Promise<Page<SparseRowPage>> => {
    const page = await h.originalRows(range, cursor, options)
    return range.r0 === 0 ? { ...page, completeness: 'loading', items: page.items.map(item => ({ ...item, cells: item.cells.filter(cell => cell.col === 0) })) } : page
  })
  vi.useFakeTimers()
  try {
    const partial = await h.paint(0, 0, 192, 512, { kind: 'pending', bounds: { x: 64, y: 256, width: 128, height: 256 }, compareWidth: 64 })
    expect(partial.paintPending).toBe(true)
    expect(partial.plan.cells.map(cell => cell.source)).toEqual([{ row: 2, col: 0 }, { row: 2, col: 2 }, { row: 0, col: 0 }])
    expect(partial.plan.cells[0]).toMatchObject({ y: 0, height: 256, overflow: { left: 0, right: 64 } })
    expect(partial.plan.cells[2]).toMatchObject({ y: 256, height: 256, overflow: { left: 0, right: 0 } })
    expect(plainClip(partial.records, 2, 0)).toEqual({ x0: 0, x1: 128 })
    expect(plainClip(partial.records, 0, 0)).toEqual({ x0: 0, x1: 64 })
    // Keep full-width native equality for the independently completed displayed row.
    expect(partial.pixels.slice(0, 192 * 256 * 4)).toEqual(partial.residentPixels.slice(0, 192 * 256 * 4))
    expect(slice5Ink(partial.pixels, 192, 65, 126, 118, 138), 'source2 complete blank B admits source2 A ink').toBe(true)
    expect(slice5Ink(partial.pixels, 192, 65, 126, 374, 394), 'source0 partial omission stays unknown at display row1').toBe(false)
    expect(h.store.cellAt({ row: 2, col: 0 })?.value).toBe('a'.repeat(40))
    expect(h.store.cellAt({ row: 0, col: 0 })?.value).toBe('b'.repeat(40))
    expect(h.store.cellAt({ row: 0, col: 1 })).toBeUndefined()
    expect(h.requests).toHaveBeenCalledTimes(1)
    expect(h.remote.stats.commands.viewport).toBe(1)
    expect(h.reads.mock.calls.map(([range]) => range)).toEqual([{ r0: 0, r1: 0, c0: 0, c1: 2 }, { r0: 2, r1: 2, c0: 0, c1: 2 }])
    expect(partial.plan.work.pages).toBe(2)
    expect(h.singles).not.toHaveBeenCalled()
    expect(h.remote.currentPageFrame.used).toBe(8)
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { vi.useRealTimers(); await h.close() }
})
