// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { ViewerClient } from '../src/worker/viewerClient'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { WorkerViewportPlans, tilePlanSource } from '../src/components/xlsx/workerViewport'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { queryViewport } from '../src/xlsx/view/viewport'
import { ViewportOverflowCollector, type OverflowViewportResult } from '../src/xlsx/view/viewportOverflow'
import type { LogicalViewport } from '../src/xlsx/view/contracts'
import { workbookOf } from './helpers/xlsx-find'

async function sparseHarness(K: number, profile: 'D' | 'M', clock: (visits: number) => number = () => 0) {
  // The positive sub-detail sentinel belongs to H but is not a paint origin.
  const doc = workbookOf([{ sheetId: 's', cells: Array.from({ length: K + 1 }, (_, col) => ({ row: 0, col, value: col + 1 })) }])
  const source = createObjectWorkbookSource(doc, 'task7'), original = source.geometry('s')
  const colOverrides = [{ start: 0, end: K - 1, sizeCss: 4 }, { start: K, end: K, sizeCss: 1 }, { start: K + 1, end: 16383, sizeCss: 4 }]
  const cols = createAxisMetrics({ logicalCount: 16384, defaultSizeCss: 64, overrides: colOverrides })
  const locateCols = vi.fn(cols.locate.bind(cols)), locateRows = vi.fn(original.rows.locate.bind(original.rows))
  const geometry = { ...original, cols: { ...cols, locate: locateCols }, rows: { ...original.rows, locate: locateRows }, colOverrides }
  source.geometry = () => geometry
  const provider = source.store('s'), readRows = vi.spyOn(provider, 'readRows'), readCell = vi.spyOn(provider, 'readCell')
  const metrics = createSheetMetrics({ rows: geometry.rows, cols: geometry.cols })
  const window: LogicalViewport = { box: { x: 0, y: 0, width: 4 * K + 5, height: 20 }, zoom: 1, deviceScale: 1 }
  const observation = { visits: 0, comparisons: 0, checks: [] as number[], yields: [] as number[], captured: 0, bytes: 0 }
  const originalRun = ViewportOverflowCollector.prototype.run
  const capture = vi.spyOn(ViewportOverflowCollector.prototype, 'run').mockImplementation(function (this: ViewportOverflowCollector, range, complete) {
    originalRun.call(this, range, complete)
    const collector = this as unknown as { occupancy: Array<{ row: number; col: number; blocking: boolean }>; bytes: number }
    const events = collector.occupancy
    observation.captured = events.length
    let sorting = false
    collector.occupancy = new Proxy(events, { get(target, field, receiver) {
      if (field === 'sort') return (compare: (a: typeof events[number], b: typeof events[number]) => number) => {
        sorting = true
        try { return target.sort((a, b) => { observation.comparisons++; return compare(a, b) }) } finally { sorting = false }
      }
      if (typeof field === 'string' && /^\d+$/.test(field) && !sorting) observation.visits++
      return Reflect.get(target, field, receiver)
    } })
    observation.bytes = collector.bytes
  })
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, {
    now: () => { observation.checks.push(observation.visits + observation.comparisons); return clock(observation.visits + observation.comparisons) },
    yield: async () => { observation.yields.push(observation.visits + observation.comparisons) },
  })
  const port = connectInProcess(kernel), client = new ViewerClient(port)
  const opened = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'task7', source: { kind: 'object', document: doc }, profile })
  if (opened.kind !== 'complete' || opened.result.kind !== 'open') throw new Error('Expected open')
  const session = kernel.session('task7')!, owner = session.identity('s')
  const remote = new RemoteWorkbook(client, { profile, docId: 'task7', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: opened.result.formatting, layouts: new Map([['s', session.layout('s')]]) })
  const plans = new WorkerViewportPlans(remote), store = remote.store('s'), request = vi.spyOn(remote, 'viewport')
  const releases = vi.spyOn(remote, 'releasePages')
  const get = tilePlanSource({ store: { coverage: () => 'oversized' } as unknown as typeof store, metrics, scale: 1, identity: owner, extent: { width: cols.totalCss(), height: geometry.rows.totalCss() }, plans })
  const flush = async () => { await port.idle(); for (let i = 0; i < 30; i++) await Promise.resolve() }
  const fetch = async () => {
    plans.beginFrame(owner, window); expect(get(window, true)?.kind).toBe('pending'); await flush()
    const out = await request.mock.results.at(-1)!.value
    if (typeof out !== 'object' || out.kind !== 'complete' || out.result.kind !== 'viewport') throw new Error('Expected real viewport result')
    return { plan: get(window, true)!, result: out.result as OverflowViewportResult }
  }
  const close = async () => { plans.dispose(); capture.mockRestore(); readRows.mockRestore(); readCell.mockRestore(); request.mockRestore(); releases.mockRestore(); await remote.close() }
  return { K, profile, window, metrics, owner, provider, observation, remote, plans, store, fetch, close, readRows, readCell, locateCols, locateRows, releases, flush }
}

it.each([['D', 512], ['M', 1024]] as const)('TASK7 sparse-work %s K%s actual route has fixed blockers and bounded indexed visits', async (profile, K) => {
  const h = await sparseHarness(K, profile)
  try {
    const { plan, result } = await h.fetch(), proof = result.overflowProof!
    expect(plan.kind).toBe('detail')
    expect(h.observation.captured).toBe(K + 1)
    expect(proof.entries).toHaveLength(K)
    expect(proof.entries.map(e => [e.left, e.right])).toEqual(Array.from({ length: K }, () => [0, 0]))
    expect(h.observation.bytes + 448 * K).toBe(K === 512 ? 279040 : 557568)
    expect(h.readRows).toHaveBeenCalledTimes(1)
    expect(h.readRows.mock.calls[0][0]).toEqual({ r0: 0, r1: 0, c0: 0, c1: K + 513 })
    expect(h.readCell).not.toHaveBeenCalled()
    console.info('TASK7 sparse receipt', JSON.stringify({ K, N: K + 1, visits: h.observation.visits, comparisons: h.observation.comparisons, checks: h.observation.checks.length, yields: h.observation.yields }))
    expect(h.observation.visits).toBeLessThanOrEqual((K + 1) + K * (4 * Math.ceil(Math.log2(K + 1)) + 8))
  } finally { await h.close() }
})

it('TASK7 sparse-cooperation checks within256 visits and yields within4096 actual visits', async () => {
  const h = await sparseHarness(1024, 'D')
  try {
    const { result } = await h.fetch()
    expect(result.overflowProof!.entries).toHaveLength(1024)
    const gaps = (points: number[]) => { const all = [0, ...points, h.observation.visits + h.observation.comparisons]; return all.slice(1).map((n, i) => n - all[i]) }
    console.info('TASK7 cooperative receipt', JSON.stringify({ visits: h.observation.visits, checks: h.observation.checks.length, yields: h.observation.yields, checkGap: Math.max(...gaps(h.observation.checks)), yieldGap: Math.max(...gaps(h.observation.yields)) }))
    expect(Math.max(...gaps(h.observation.checks))).toBeLessThanOrEqual(256)
    expect(Math.max(...gaps(h.observation.yields))).toBeLessThanOrEqual(4096)
  } finally { await h.close() }
})

it.each(['D', 'M'] as const)('TASK7 no-extra-reads %s occupancy uses original sparse query and axis locates', async profile => {
  const h = await sparseHarness(8, profile)
  try {
    const base = await queryViewport(h.provider, h.metrics, h.window, h.owner, { profile, scope: 'tile', signal: new AbortController().signal, textDependencies: true })
    expect(base.kind).toBe('viewport')
    const rowCalls = h.readRows.mock.calls.map(call => call[0]), singles = h.readCell.mock.calls.length
    const locates = [h.locateRows.mock.calls.length, h.locateCols.mock.calls.length]
    h.readRows.mockClear(); h.readCell.mockClear(); h.locateRows.mockClear(); h.locateCols.mockClear()
    const { plan, result } = await h.fetch()
    expect(plan.kind).toBe('detail')
    expect(h.readRows.mock.calls.map(call => call[0])).toEqual(rowCalls)
    expect(h.readCell).toHaveBeenCalledTimes(singles)
    expect([h.locateRows.mock.calls.length, h.locateCols.mock.calls.length]).toEqual(locates)
    expect(result.overflowProof!.entries).toHaveLength(8)
    expect(h.remote.pagesInFlight).toBe(0)
    expect(h.releases.mock.calls).toEqual([[profile === 'D' ? 8 : 4]])
  } finally { await h.close() }
})

async function publicHarness(profile: 'D' | 'M', cells: Array<{ row: number; col: number; value: string | number }> = [{ row: 0, col: 0, value: 'old' }, { row: 1, col: 0, value: 'incoming' }], tiny = false, docId = 'bounds') {
  const doc = workbookOf([{ sheetId: 's', cells }]), source = createObjectWorkbookSource(doc, docId)
  const original = source.geometry('s')
  const columns = tiny ? createAxisMetrics({ logicalCount: 16384, defaultSizeCss: 64, overrides: [{ start: 0, end: 16383, sizeCss: 1 }] }) : original.cols
  const locateCols = vi.fn(columns.locate.bind(columns)), locateRows = vi.fn(original.rows.locate.bind(original.rows))
  const geometry = { ...original, cols: { ...columns, locate: locateCols }, rows: { ...original.rows, locate: locateRows }, ...(tiny ? { colOverrides: [{ start: 0, end: 16383, sizeCss: 1 }] } : {}) }
  source.geometry = () => geometry
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} }), port = connectInProcess(kernel), client = new ViewerClient(port)
  const opened = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: docId, source: { kind: 'object', document: doc }, profile })
  if (opened.kind !== 'complete' || opened.result.kind !== 'open') throw new Error('Expected open')
  const session = kernel.session(docId)!, owner = session.identity('s'), remote = new RemoteWorkbook(client, { profile, docId: docId, valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: opened.result.formatting, layouts: new Map([['s', session.layout('s')]]) })
  const plans = new WorkerViewportPlans(remote), store = remote.store('s'), provider = source.store('s'), metrics = createSheetMetrics({ rows: geometry.rows, cols: geometry.cols })
  const extent = { width: columns.totalCss(), height: geometry.rows.totalCss() }
  const flush = async () => { await port.idle(); for (let i = 0; i < 30; i++) await Promise.resolve() }
  remote.merges('s'); await flush()
  const window = (row = 0, x = 0): LogicalViewport => ({ box: { x, y: row * 20, width: 64, height: 20 }, zoom: 1, deviceScale: 1 })
  const fetch = async (w = window()) => { plans.beginFrame(owner, w); expect(plans.get(owner, w, extent, true).kind).toBe('pending'); await flush(); return plans.get(owner, w, extent, true) }
  const close = async () => { plans.dispose(); await remote.close() }
  return { doc, source, kernel, session, port, client, owner, remote, plans, store, provider, metrics, extent, flush, window, fetch, close, locateRows, locateCols }
}

it.each(['D', 'M'] as const)('TASK7 tiny-lazy %s full-width coarse query never attempts logical-cell enumeration', async profile => {
  const h = await publicHarness(profile, [], true), reads = vi.spyOn(h.provider, 'readRows'), singles = vi.spyOn(h.provider, 'readCell'), overview = vi.spyOn(h.provider, 'queryOverview')
  const w = { ...h.window(), box: { x: 0, y: 0, width: 16384, height: 20 } }
  try {
    const baseline = await queryViewport(h.provider, h.metrics, w, h.owner, { profile, scope: 'tile', signal: new AbortController().signal, textDependencies: true })
    expect(baseline.kind).toBe('viewport')
    // Bucket4→8→16(D)→32(M), counting ceil(20/size) * ceil(16384/size).
    const expectedGroups = profile === 'D' ? 2048 : 512
    if (baseline.kind !== 'viewport') throw new Error('Expected coarse baseline')
    expect(baseline.cells).toEqual([])
    console.info('TASK7 tiny baseline', profile, baseline.coarse.length, h.locateCols.mock.calls.length)
    expect(reads).not.toHaveBeenCalled(); expect(singles).not.toHaveBeenCalled()
    expect(baseline.coarse).toHaveLength(expectedGroups)
    expect(h.locateCols).toHaveBeenCalledTimes(2 * expectedGroups)
    const baselineGroups = expectedGroups, baselineLocates = 2 * expectedGroups
    h.locateCols.mockClear(); h.locateRows.mockClear(); overview.mockClear()
    const plan = await h.fetch(w)
    expect(plan.kind).toBe('detail')
    if (plan.kind !== 'detail') throw new Error('Expected empty detail bridge')
    expect(plan.cells).toEqual([])
    expect(plan.coarse).toHaveLength(baselineGroups)
    expect(h.locateCols).toHaveBeenCalledTimes(baselineLocates)
    expect(reads).not.toHaveBeenCalled(); expect(singles).not.toHaveBeenCalled()
    expect(overview).toHaveBeenCalledTimes(1)
    const query = overview.mock.calls[0][0]
    expect(query.groups.length).toBeLessThanOrEqual(profile === 'D' ? 2048 : 1024)
    expect(query.groups.length).toBe(baselineGroups)
    const page = await overview.mock.results[0].value
    expect(page.items).toHaveLength(baselineGroups)
    expect(h.store.cachedRows).toBe(0)
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { reads.mockRestore(); singles.mockRestore(); overview.mockRestore(); await h.close() }
})

it.each(['D', 'M'] as const)('TASK7 joint-pressure %s actual new reply retains held raw old rows and old plans before allocation refusal and retry', async profile => {
  const h = await publicHarness(profile), limit = (profile === 'D' ? 64 : 32) * 1024 * 1024
  const oldWindow = h.window(), incomingWindow = h.window(1)
  const oldPlan = await h.fetch(), oldRows = h.store.cellsOf(0), oldBytes = h.remote.mainLedger.snapshot()
  const held = await h.remote.viewport(h.owner, oldWindow, new AbortController().signal, 'tile')
  if (typeof held !== 'object' || held.kind !== 'complete') throw new Error('Expected independently held actual raw reply')
  const heldBytes = h.remote.mainLedger.snapshot().byPool.replies
  const originalViewport = h.remote.viewport.bind(h.remote), stage = vi.spyOn(h.store, 'stageViewport'), copies = vi.spyOn(h.remote, 'styleOf'), settle = vi.spyOn(h.remote.mainLedger, 'settle')
  let incoming: object | undefined, proofReads = 0
  const viewport = vi.spyOn(h.remote, 'viewport').mockImplementation(async (...args) => {
    const out = await originalViewport(...args)
    if (typeof out === 'object' && out.kind === 'complete') {
      incoming = out
      const proof = (out.result as OverflowViewportResult).overflowProof!, entries = proof.entries
      Object.defineProperty(proof, 'entries', { enumerable: true, get: () => { proofReads++; return entries } })
      h.remote.mainLedger.set('copy', 'task7-joint-pressure', limit - h.remote.mainLedger.bytes)
    }
    return out
  })
  try {
    expect(oldPlan.kind).toBe('detail')
    expect(oldBytes.byPool.rowCache).toBeGreaterThan(0); expect(oldBytes.byPool.layout).toBeGreaterThan(0); expect(heldBytes).toBeGreaterThan(0)
    const refused = await h.fetch(incomingWindow)
    expect(refused.kind).toBe('diagnostic')
    expect(viewport).toHaveBeenCalledTimes(1)
    expect(stage).not.toHaveBeenCalled(); expect(copies).not.toHaveBeenCalled(); expect(proofReads).toBe(0)
    expect(h.store.cellsOf(0)).toBe(oldRows); expect(h.store.cellsOf(1)).toBeUndefined()
    expect(h.plans.get(h.owner, oldWindow, h.extent, true)).toBe(oldPlan)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(oldBytes.byPool.layout)
    expect(settle).not.toHaveBeenCalledWith(held); expect(settle).not.toHaveBeenCalledWith(incoming)
    expect(h.remote.mainLedger.snapshot().byPool.replies).toBeGreaterThan(heldBytes)
    expect(h.remote.mainLedger.peakBytes).toBe(limit)
    expect(h.remote.pagesInFlight).toBe(0)
    h.remote.mainLedger.release('copy', 'task7-joint-pressure'); viewport.mockRestore()
    h.plans.release(h.owner, incomingWindow)
    const retry = await h.fetch(incomingWindow)
    expect(retry.kind).toBe('detail'); expect(h.store.cellAt({ row: 1, col: 0 })?.value).toBe('incoming')
    expect(settle).not.toHaveBeenCalledWith(held)
    expect(h.remote.mainLedger.peakBytes).toBeLessThanOrEqual(limit)
    console.info('TASK7 joint receipt', JSON.stringify({ profile, oldBytes, heldBytes, retry: h.remote.mainLedger.snapshot() }))
  } finally {
    viewport.mockRestore(); stage.mockRestore(); copies.mockRestore(); settle.mockRestore()
    h.remote.mainLedger.release('copy', 'task7-joint-pressure')
    // Explicitly end these two consumers only after their assertions/drain.
    h.remote.mainLedger.settle(held); if (incoming) h.remote.mainLedger.settle(incoming)
    await h.close()
  }
})

it.each(['D', 'M'] as const)('TASK7 credits %s held actual ring plans leave visible share and return each full reservation once', async profile => {
  const h = await publicHarness(profile), tile = profile === 'D' ? 8 : 4, frame = profile === 'D' ? 32 : 16
  const actualRows = h.provider.readRows.bind(h.provider)
  let drain!: () => void
  const gate = new Promise<void>(resolve => { drain = resolve })
  const held: object[] = []
  const reads = vi.spyOn(h.provider, 'readRows').mockImplementation(async (...args) => { const page = await actualRows(...args); held.push(page); await gate; return page })
  const requests = vi.spyOn(h.remote, 'viewport'), releases = vi.spyOn(h.remote, 'releasePages')
  const w = (row: number) => h.window(row)
  try {
    h.plans.beginFrame(h.owner, { ...w(0), box: { x: 0, y: 0, width: 64, height: 80 } })
    h.plans.get(h.owner, w(0), h.extent, false); h.plans.get(h.owner, w(1), h.extent, false)
    expect(h.remote.pagesInFlight).toBe(frame / 2)
    h.plans.get(h.owner, w(2), h.extent, false)
    expect(requests).toHaveBeenCalledTimes(2)
    h.plans.get(h.owner, w(2), h.extent, true); h.plans.get(h.owner, w(3), h.extent, true)
    expect(requests).toHaveBeenCalledTimes(4); expect(h.remote.pagesInFlight).toBe(frame)
    h.plans.get(h.owner, w(4), h.extent, true)
    expect(requests).toHaveBeenCalledTimes(4)
    for (let i = 0; i < 200 && held.length < 4; i++) await Promise.resolve()
    expect(held).toHaveLength(4)
    expect(releases).not.toHaveBeenCalled()
    drain(); await h.flush()
    expect(h.remote.pagesInFlight).toBe(0)
    expect(releases.mock.calls).toEqual(Array.from({ length: 4 }, () => [tile]))
    h.plans.get(h.owner, w(4), h.extent, true)
    expect(requests).toHaveBeenCalledTimes(4)
    h.plans.beginFrame(h.owner, w(4)); h.plans.get(h.owner, w(4), h.extent, true); await h.flush()
    expect(requests).toHaveBeenCalledTimes(5)
    expect(releases.mock.calls).toEqual(Array.from({ length: 5 }, () => [tile]))
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { drain(); await h.flush(); reads.mockRestore(); requests.mockRestore(); releases.mockRestore(); await h.close() }
})

it.each(['D', 'M'] as const)('TASK7 shared-row-credits %s a real held row blocks the same tile plan and spends cumulative frame credit', async profile => {
  const h = await publicHarness(profile), actualRows = h.provider.readRows.bind(h.provider)
  let drain!: () => void, entered!: () => void
  const gate = new Promise<void>(resolve => { drain = resolve }), held = new Promise<void>(resolve => { entered = resolve })
  const reads = vi.spyOn(h.provider, 'readRows').mockImplementation(async (...args) => { const page = await actualRows(...args); entered(); await gate; return page })
  const requests = vi.spyOn(h.remote, 'viewport'), releases = vi.spyOn(h.remote, 'releasePages')
  try {
    const w = h.window(), tile = profile === 'D' ? 8 : 4
    h.plans.beginFrame(h.owner, w)
    h.store.pendingIn(h.metrics, w.box, 1)
    await held
    expect(h.remote.pagesInFlight).toBe(1)
    h.plans.get(h.owner, w, h.extent, true)
    expect(requests).not.toHaveBeenCalled()
    drain(); await h.flush()
    expect(h.remote.pagesInFlight).toBe(0); expect(releases.mock.calls).toEqual([[1]])
    h.plans.get(h.owner, w, h.extent, true)
    expect(requests).not.toHaveBeenCalled()
    h.plans.beginFrame(h.owner, w); h.plans.get(h.owner, w, h.extent, true); await h.flush()
    expect(requests).toHaveBeenCalledTimes(1); expect(releases.mock.calls).toEqual([[1], [tile]])
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { drain(); await h.flush(); reads.mockRestore(); requests.mockRestore(); releases.mockRestore(); await h.close() }
})

it.each([['D', 'cancel'], ['M', 'cancel'], ['D', 'close'], ['M', 'close']] as const)('TASK7 provider-lifetime %s %s preserves charge for signal-ignoring allocated page until producer drain', async (profile, action) => {
  const h = await publicHarness(profile), actualRows = h.provider.readRows.bind(h.provider), allocation = 1024 * 1024
  let drain!: () => void, entered!: () => void, completed!: () => void
  const gate = new Promise<void>(resolve => { drain = resolve }), held = new Promise<void>(resolve => { entered = resolve }), ended = new Promise<void>(resolve => { completed = resolve })
  let providerOwned: { page: object; backing: Uint8Array } | undefined
  const reads = vi.spyOn(h.provider, 'readRows').mockImplementation(async (...args) => {
    const page = await actualRows(...args), backing = new Uint8Array(allocation)
    providerOwned = { page, backing }; entered()
    await gate
    const output = { ...page, backing }
    providerOwned = undefined; completed()
    return output
  })
  const releases = vi.spyOn(h.remote, 'releasePages')
  try {
    const w = h.window(), tile = profile === 'D' ? 8 : 4
    h.plans.beginFrame(h.owner, w); h.plans.get(h.owner, w, h.extent, true); await held
    expect(providerOwned!.backing.byteLength).toBe(allocation)
    expect(h.session.memoryStats().transientBytes).toBeGreaterThanOrEqual(allocation)
    if (action === 'cancel') h.plans.release(h.owner, w)
    else await h.remote.close()
    await h.flush()
    expect(providerOwned!.backing.byteLength).toBe(allocation)
    expect(h.remote.pagesInFlight).toBe(0)
    expect(releases.mock.calls).toEqual([[tile]])
    console.info('TASK7 lifetime receipt', JSON.stringify({ profile, action, beforeDrain: h.session.memoryStats(), allocation, providerReachable: Boolean(providerOwned) }))
    expect(h.session.memoryStats().transientBytes).toBeGreaterThanOrEqual(allocation)
    drain(); await ended; await h.flush()
    expect(providerOwned).toBeUndefined()
    expect(h.session.memoryStats().transientBytes).toBe(0)
    expect(h.store.cellsOf(0)).toBeUndefined()
    expect(releases.mock.calls).toEqual([[tile]])
  } finally { drain(); await ended; await h.flush(); reads.mockRestore(); releases.mockRestore(); await h.close() }
})

it.each(['D', 'M'] as const)('TASK7 row-subcap %s unchanged real row budget preserves warm rows when retained growth fits while main and paint fit', async profile => {
  const oldCount = profile === 'D' ? 11 : 5, incomingRow = oldCount, cells = Array.from({ length: oldCount }, (_, row) => ({ row, col: 0, value: 'o'.repeat(700000) }))
  cells.push({ row: incomingRow, col: 0, value: 'n'.repeat(300000) })
  const h = await publicHarness(profile, cells), cap = (profile === 'D' ? 64 : 32) * 1024 * 1024
  try {
    expect(h.remote.ledger.limitBytes).toBe(cap / 4)
    for (let row = 0; row < oldCount; row++) { h.remote.beginPageFrame(); h.store.pendingIn(h.metrics, h.window(row).box, 1); await h.flush() }
    const oldRows = Array.from({ length: oldCount }, (_, row) => h.store.cellsOf(row))
    expect(oldRows.every(row => row?.length === 1)).toBe(true)
    const retained = h.remote.ledger.bytes
    // Row accounting is64+cachedCellBytes=64+96+UTF16value; independently fixed.
    expect(retained).toBe(oldCount * (64 + 96 + 1400000))
    const style = vi.spyOn(h.remote, 'styleOf'), stage = vi.spyOn(h.store, 'stageViewport')
    try {
      const plan = await h.fetch(h.window(incomingRow))
      expect(plan.kind).toBe('detail')
      expect(stage).toHaveBeenCalledTimes(1)
      expect(style).toHaveBeenCalledTimes(2)
      expect(h.store.cellsOf(incomingRow)?.map(cell => cell.value)).toEqual(['n'.repeat(300000)])
      const oldBytes = 64 + 96 + 1400000
      // The 0.6 MB retained row fits the existing row-cache headroom.
      // The larger transient stage belongs to the separately admitted main budget.
      const evicted = Math.max(0, Math.ceil((retained + 64 + 96 + 600000 - h.remote.ledger.limitBytes) / oldBytes))
      expect(evicted).toBe(0)
      expect(h.remote.ledger.bytes).toBe(retained - evicted * oldBytes + 64 + 96 + 600000)
      for (let row = 0; row < oldCount; row++) {
        if (row < evicted) { expect(h.store.cellsOf(row)).toBeUndefined(); expect(h.store.isKnown({ row, col: 63 })).toBe(false) }
        else { expect(h.store.cellsOf(row)).toBe(oldRows[row]); expect(h.store.isKnown({ row, col: 63 })).toBe(true) }
      }
      expect(h.store.isKnown({ row: incomingRow, col: 0 })).toBe(false)
      expect(h.remote.mainLedger.snapshot().byPool.layout).toBeGreaterThan(0)
      expect(h.remote.mainLedger.peakBytes).toBeLessThanOrEqual(cap)
      expect(h.remote.pagesInFlight).toBe(0)
    } finally { style.mockRestore(); stage.mockRestore() }
  } finally { await h.close() }
})

it.each(['D', 'M'] as const)('TASK7 eviction %s retains32 settled actual plans and touches oldest before33rd', async profile => {
  const h = await publicHarness(profile, Array.from({ length: 34 }, (_, row) => ({ row, col: 0, value: row + 1 }))), requests = vi.spyOn(h.remote, 'viewport')
  const stages = new Map<string, number>(), actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
  const writes = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => { if (owner.includes(':stage:')) stages.set(owner, bytes); actualSet(pool, owner, bytes) })
  try {
    const initial = []
    for (let row = 0; row < 32; row++) { const plan = await h.fetch(h.window(row)); expect(plan.kind).toBe('detail'); initial.push(plan) }
    expect(requests).toHaveBeenCalledTimes(32)
    expect(h.plans.get(h.owner, h.window(0), h.extent, true)).toBe(initial[0])
    const newest = await h.fetch(h.window(32))
    expect(newest.kind).toBe('detail'); expect(requests).toHaveBeenCalledTimes(33)
    expect(h.plans.get(h.owner, h.window(0), h.extent, true)).toBe(initial[0])
    const next = await h.fetch(h.window(1))
    expect(next.kind).toBe('detail'); expect(next).not.toBe(initial[1]); expect(requests).toHaveBeenCalledTimes(34)
    const entries = (h.plans as unknown as { entries: Map<string, { bytes: number; plan?: object }> }).entries
    expect(entries.size).toBe(32)
    expect([...entries.values()].every(entry => entry.plan)).toBe(true)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe([...entries.values()].reduce((sum, entry) => sum + entry.bytes, 0))
    expect([...stages.values()].every(bytes => bytes === 0)).toBe(true)
    expect(h.remote.pagesInFlight).toBe(0)
    h.plans.dispose()
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
  } finally { requests.mockRestore(); writes.mockRestore(); await h.close() }
})

it.each(['D', 'M'] as const)('TASK7 prepared %s overflow requests preserve actual profile preparation capacity and owners', async profile => {
  const count = profile === 'D' ? 8 : 4, h = await publicHarness(profile, Array.from({ length: count }, (_, col) => ({ row: 0, col, value: col + 1 })))
  try {
    const prepared = await h.client.send({ kind: 'prepareColumns', requestId: h.client.nextRequestId(), owner: h.owner, cols: Uint16Array.from({ length: count }, (_, col) => col) })
    expect(prepared.kind).toBe('complete')
    const before = h.session.memoryStats().preparedBytes
    expect(before).toBeGreaterThan(0)
    expect(await h.fetch()).toMatchObject({ kind: 'detail' })
    expect(h.session.memoryStats().preparedBytes).toBe(before)
    const excess = await h.client.send({ kind: 'prepareColumns', requestId: h.client.nextRequestId(), owner: h.owner, cols: Uint16Array.from({ length: count + 1 }, (_, col) => col) })
    expect(excess.kind).toBe(profile === 'D' ? 'error' : 'quota')
    expect(h.session.memoryStats().preparedBytes).toBe(before)
    expect(h.session.memoryStats().limitBytes).toBe((profile === 'D' ? 512 : 256) * 1024 * 1024)
    expect(h.remote.pagesInFlight).toBe(0)
    await h.remote.close()
    expect(h.session.memoryStats().preparedBytes).toBe(0)
  } finally { await h.close() }
})

it.each([['D', 3855], ['M', 1927]] as const)('TASK7 joint-workspace %s K%s admits legal sparse evidence/output', async (profile, K) => {
  const h = await sparseHarness(K, profile)
  try {
    h.plans.beginFrame(h.owner, h.window)
    const extent = { width: h.metrics.cols.totalCss(), height: h.metrics.rows.totalCss() }
    expect(h.plans.get(h.owner, h.window, extent, true).kind).toBe('pending')
    // Use actual transport completion instead of a fabricated overflow payload.
    await h.flush()
    const plan = h.plans.get(h.owner, h.window, extent, true)
    expect(plan.kind).toBe('detail')
    expect(h.observation.captured).toBe(K + 1)
    expect(h.readRows).toHaveBeenCalledTimes(1)
    expect(h.store.cachedRows).toBe(1)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBeGreaterThan(0)
    expect(h.remote.pagesInFlight).toBe(0)
    expect(h.releases.mock.calls).toEqual([[profile === 'D' ? 8 : 4]])
  } finally { await h.close() }
})

it('TASK7 clock-cooperation injected10ms forces preprocessing yields before4096 work', async () => {
  const h = await sparseHarness(512, 'D', visits => Math.floor(visits / 64) * 10)
  try {
    const { plan } = await h.fetch()
    expect(plan.kind).toBe('detail')
    expect(h.observation.yields.length).toBeGreaterThan(0)
    const points = [0, ...h.observation.yields, h.observation.visits + h.observation.comparisons]
    expect(Math.max(...points.slice(1).map((point, index) => point - points[index]))).toBeLessThanOrEqual(68)
  } finally { await h.close() }
})

it.each([['D', 'cancel'], ['M', 'cancel'], ['D', 'close'], ['M', 'close']] as const)('TASK7 late-rejection %s %s drains actual producer once with no owner revival', async (profile, action) => {
  const h = await publicHarness(profile)
  let reject!: (cause: unknown) => void, entered!: () => void
  const work = new Promise<never>((_resolve, fail) => { reject = fail }), held = new Promise<void>(resolve => { entered = resolve })
  const reads = vi.spyOn(h.provider, 'readRows').mockImplementation(() => { entered(); return work })
  const releases = vi.spyOn(h.remote, 'releasePages'), cause = new Error('actual held producer failure')
  try {
    const w = h.window(), tile = profile === 'D' ? 8 : 4
    h.plans.beginFrame(h.owner, w); h.plans.get(h.owner, w, h.extent, true); await held
    if (action === 'cancel') h.plans.release(h.owner, w)
    else await h.remote.close()
    await h.flush()
    expect(h.session.memoryStats().transientBytes).toBe(2 * (profile === 'D' ? 8 : 4) * 1024 * 1024)
    expect(h.remote.pagesInFlight).toBe(0); expect(releases.mock.calls).toEqual([[tile]])
    reject(cause); await work.catch(error => { expect(error).toBe(cause) }); await h.flush()
    expect(h.session.memoryStats().transientBytes).toBe(0)
    expect(h.store.cellsOf(0)).toBeUndefined()
    expect(releases.mock.calls).toEqual([[tile]])
    if (action === 'cancel') {
      reads.mockRestore()
      expect(await h.fetch()).toMatchObject({ kind: 'detail' })
      expect(h.remote.pagesInFlight).toBe(0)
    } else {
      expect(h.kernel.session('bounds')).toBeUndefined()
      expect(h.plans.get(h.owner, w, h.extent, true).kind).toBe('pending')
    }
  } finally { reject(cause); await work.catch(() => {}); await h.flush(); reads.mockRestore(); releases.mockRestore(); await h.close() }
})

it.each(['ordinary', 'falsy', 'synchronous'] as const)('TASK7 producer-error %s preserves actual session rejection identity and drainage', async kind => {
  const h = await publicHarness('D'), cause = kind === 'falsy' ? false : new Error(`source ${kind}`)
  const reads = vi.spyOn(h.provider, 'readRows').mockImplementation(() => { if (kind === 'synchronous') throw cause; return Promise.reject(cause) })
  try {
    let caught = false, received: unknown
    try { await h.session.viewport(h.owner, h.window(), new AbortController().signal, 'tile') } catch (error) { caught = true; received = error }
    expect(caught).toBe(true); expect(received).toBe(cause)
    await h.flush()
    expect(h.session.memoryStats().transientBytes).toBe(0)
  } finally { reads.mockRestore(); await h.close() }
})

it.each(['D', 'M'] as const)('TASK7 coexistence %s holds old new source arrays plans raw and staged handoff owners at publication peak', async profile => {
  const h = await publicHarness(profile, [{ row: 0, col: 0, value: 'old' }, { row: 0, col: 1, value: 'new' }])
  const oldPlan = await h.fetch(), oldRows = h.store.cellsOf(0)!, oldLayout = h.remote.mainLedger.snapshot().byPool.layout
  const held = await h.remote.viewport(h.owner, h.window(), new AbortController().signal, 'tile')
  if (typeof held !== 'object' || held.kind !== 'complete') throw new Error('Expected held actual raw reply')
  const names = new Map<string, number>(), observations: Array<{ reserve: number; snapshot: ReturnType<typeof h.remote.mainLedger.snapshot>; replacement: boolean }> = []
  const actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger), writes = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
    actualSet(pool, owner, bytes)
    if (pool === 'layout') {
      names.set(owner, bytes)
      if (owner.startsWith('viewport-plans:') && !owner.includes(':stage:') && bytes > oldLayout) observations.push({ reserve: [...names].filter(([name]) => name.includes(':stage:')).reduce((sum, [, amount]) => sum + amount, 0), snapshot: h.remote.mainLedger.snapshot(), replacement: h.store.cellsOf(0) !== oldRows })
    }
  })
  const settled = vi.spyOn(h.remote.mainLedger, 'settle')
  try {
    const next = await h.fetch(h.window(0, 64))
    expect(next.kind).toBe('detail'); expect(next).not.toBe(oldPlan)
    expect(h.plans.get(h.owner, h.window(), h.extent, true)).toBe(oldPlan)
    expect(oldRows.map(cell => cell.value)).toEqual(['old', 'new'])
    expect(h.store.cellsOf(0)).not.toBe(oldRows)
    expect(observations.length).toBeGreaterThan(0)
    const publication = observations[0]
    expect(publication.replacement).toBe(true)
    expect(publication.reserve).toBeGreaterThan(0)
    expect(publication.snapshot.byPool.rowCache).toBe(64 + 2 * 96 + 2 * (3 + 3))
    expect(publication.snapshot.byPool.replies).toBeGreaterThan(0)
    expect(publication.snapshot.byPool.layout).toBeGreaterThan(oldLayout + publication.reserve)
    expect([...names].filter(([name]) => name.includes(':stage:')).every(([, amount]) => amount === 0)).toBe(true)
    expect(settled).not.toHaveBeenCalledWith(held)
    expect(h.remote.pagesInFlight).toBe(0)
    expect(h.remote.mainLedger.peakBytes).toBeLessThanOrEqual((profile === 'D' ? 64 : 32) * 1024 * 1024)
    console.info('TASK7 coexistence receipt', JSON.stringify({ profile, publication, peak: h.remote.mainLedger.snapshot() }))
  } finally { writes.mockRestore(); settled.mockRestore(); h.remote.mainLedger.settle(held); await h.close() }
})

it('TASK7 index-accounting indexes and sort coexist inside unchanged mobile2Q', async () => {
  const h = await sparseHarness(1800, 'M')
  try {
    h.plans.beginFrame(h.owner, h.window)
    const extent = { width: h.metrics.cols.totalCss(), height: h.metrics.rows.totalCss() }
    h.plans.get(h.owner, h.window, extent, true); await h.flush()
    // This legal tile exceeded the old fixed Q/4 evidence partition. Its
    // provider page, output, sort and indexes fit the unchanged joint 2Q owner.
    expect(h.observation.captured).toBe(1801)
    expect(h.plans.get(h.owner, h.window, extent, true).kind).toBe('detail')
    expect(h.store.cachedRows).toBe(1)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBeGreaterThan(0)
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { await h.close() }
})

it.each([['D', 'M'], ['M', 'D']] as const)('TASK7 replacement %s→%s new document/profile has independent source plan owner after old close', async (from, to) => {
  const old = await publicHarness(from, [{ row: 0, col: 0, value: 'old document' }], false, 'old-document')
  const oldPlan = await old.fetch()
  expect(oldPlan.kind).toBe('detail')
  await old.remote.close(); old.plans.dispose()
  expect(old.remote.mainLedger.snapshot().byPool.layout).toBe(0)
  expect(old.session.memoryStats().totalBytes).toBe(0)
  const next = await publicHarness(to, [{ row: 0, col: 0, value: 'new document' }], false, 'new-document')
  try {
    const plan = await next.fetch()
    expect(plan.kind).toBe('detail'); expect(plan).not.toBe(oldPlan)
    expect(next.owner.docId).toBe('new-document'); expect(old.owner.docId).toBe('old-document')
    expect(next.remote.profile).toBe(to)
    expect(next.store.cellAt({ row: 0, col: 0 })?.value).toBe('new document')
    expect(old.store.cellAt({ row: 0, col: 0 })).toBeUndefined()
    expect(old.plans.get(next.owner, next.window(), next.extent, true).kind).toBe('pending')
    expect(old.session.memoryStats().totalBytes).toBe(0)
    expect(next.remote.pagesInFlight).toBe(0)
  } finally { await old.close(); await next.close() }
})
