// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { WORKER_QUERY_BYTES, queryPayloadBytes } from '../src/worker/viewerQueryBudget'
import { ViewportOverflowCollector, validateOverflowProof, viewportPlanAdmission, type OverflowViewportResult } from '../src/xlsx/view/viewportOverflow'
import { queryViewport, VIEWPORT_DETAIL_LIMITS } from '../src/xlsx/view/viewport'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import type { LogicalViewport, Page, SparseRowPage } from '../src/xlsx/view/contracts'
import { workbookOf } from './helpers/xlsx-find'

async function harness(profile: 'D' | 'M', rows = 1, cols = 1, text = '1') {
  const document = workbookOf([{ sheetId: 's', cells: Array.from({ length: rows * cols }, (_, at) => ({ row: Math.floor(at / cols), col: at % cols, value: text })) }])
  const source = createObjectWorkbookSource(document, 'slice9-collector')
  const kernel = new ViewerKernel({ parse: async () => document, openSource: async () => source }, { yield: async () => {} })
  await kernel.handle({ kind: 'open', requestId: 'open', docId: 'slice9-collector', source: { kind: 'object', document }, profile }, () => {})
  const session = kernel.session('slice9-collector')!, owner = session.identity('s'), geometry = source.geometry('s')
  const metrics = createSheetMetrics({ rows: geometry.rows, cols: geometry.cols })
  const window: LogicalViewport = { box: { x: 0, y: 0, width: 64, height: 20 }, zoom: 1, deviceScale: 1 }
  const close = () => kernel.handle({ kind: 'close', requestId: 'close', docId: owner.docId }, () => {})
  return { document, source, kernel, session, owner, metrics, window, close }
}

it.each([['D', 0.6, 120, 170, 3], ['M', 1, 120, 170, 4]] as const)('dense %s real kernel keeps legal tile detail', async (profile, zoom, rows, cols, tx) => {
  const h = await harness(profile, rows, cols)
  const window = { box: { x: tx * 512 / zoom, y: 0, width: 512 / zoom, height: 512 / zoom }, zoom, deviceScale: 1 }
  try {
    const reply = await h.session.viewport(h.owner, window, new AbortController().signal, 'tile')
    expect(reply).toHaveProperty('result.content.kind', 'detail')
    if (!('result' in reply) || reply.result.content.kind !== 'detail') throw new Error('Expected detail')
    expect(reply.result.content.cells.length).toBeGreaterThan(profile === 'D' ? 3300 : 1800)
    expect(reply.result.content.cells.length).toBeLessThanOrEqual(profile === 'D' ? 5000 : 2500)
    expect(reply.result.overflowProof?.entries).toHaveLength(reply.result.content.cells.length)
    expect(queryPayloadBytes(reply, WORKER_QUERY_BYTES[profile])).toBeLessThanOrEqual(WORKER_QUERY_BYTES[profile])
    expect(h.session.memoryStats().transientBytes).toBe(0)
  } finally { await h.close() }
})

it.each([['D', 100], ['M', 50]] as const)('%s accepts accumulated output above the former half-query partition', async (profile, rows) => {
  const h = await harness(profile, rows, 1, 'a'.repeat(21000))
  const window = { ...h.window, box: { x: 0, y: 0, width: 64, height: rows * 20 } }
  try {
    const reply = await h.session.viewport(h.owner, window, new AbortController().signal)
    expect(reply).toHaveProperty('result.content.kind', 'detail')
    expect(queryPayloadBytes(reply, WORKER_QUERY_BYTES[profile])).toBeGreaterThan(WORKER_QUERY_BYTES[profile] / 2)
    expect(queryPayloadBytes(reply, WORKER_QUERY_BYTES[profile])).toBeLessThanOrEqual(WORKER_QUERY_BYTES[profile])
  } finally { await h.close() }
})

it('pending influence stops at nearer known blockers on both sides', async () => {
  const h = await harness('D', 1, 7)
  try {
    const collector = new ViewportOverflowCollector(WORKER_QUERY_BYTES.D)
    const cells = h.source.store('s')
    const page = await cells.readRows({ r0: 0, r1: 0, c0: 0, c1: 6 }, undefined, { stamp: cells.stamp, signal: new AbortController().signal })
    collector.structure('complete', [])
    for (const col of [1, 3, 5]) collector.cell(page.items[0].cells[col])
    collector.run({ r0: 0, r1: 0, c0: 0, c1: 6 }, false)
    const { unavailable: _unavailable, ...cell } = page.items[0].cells[3]
    const item = { kind: 'cell' as const, cell, source: { row: 0, col: 3 }, bounds: { x: 192, y: 0, width: 64, height: 20 } }
    const window = { ...h.window, box: { x: 0, y: 0, width: 448, height: 20 } }
    const steps = collector.finish(h.metrics, [item], h.owner, window, 'D', 'tile')
    let step = steps.next()
    while (!step.done) step = steps.next()
    expect(step.value.pending.map(p => p.bounds)).toEqual([
      { x: 128, y: 0, width: 64, height: 20 },
      { x: 256, y: 0, width: 64, height: 20 },
    ])
  } finally { await h.close() }
})

it.each(['h0', 'h1'] as const)('rejects forged %s even when corridors and clip arithmetic stay self-consistent', async field => {
  const h = await harness('D')
  try {
    const reply = await h.session.viewport(h.owner, h.window, new AbortController().signal, 'tile')
    if (!('result' in reply)) throw new Error('Expected result')
    const result = reply.result as OverflowViewportResult
    result.overflowProof!.entries[0][field] += field === 'h0' ? -64 : 64
    expect(() => validateOverflowProof(result, h.owner, h.window, 'D', 'tile', h.metrics.cols)).toThrow('Invalid or missing worker overflow evidence')
  } finally { await h.close() }
})

it('holds workspace through the late facade observer before deferred producer drainage', async () => {
  const h = await harness('D'), store = h.source.store('s'), controller = new AbortController()
  const original = store.readRows.bind(store)
  const page = await original({ r0: 0, r1: 0, c0: 0, c1: 0 }, undefined, { stamp: store.stamp, signal: new AbortController().signal })
  let deliver!: (page: Page<SparseRowPage>) => void
  const held = new Promise<Page<SparseRowPage>>(resolve => { deliver = resolve }), nativeThen = held.then.bind(held)
  let observers = 0, lateObserverBytes: number | undefined
  held.then = ((fulfilled, rejected) => {
    const at = ++observers
    return nativeThen(value => { if (at === 2) lateObserverBytes = h.session.memoryStats().transientBytes; return fulfilled!(value) }, rejected)
  }) as typeof held.then
  const read = store.readRows
  store.readRows = () => held
  try {
    const pending = h.session.viewport(h.owner, h.window, controller.signal, 'tile')
    for (let i = 0; i < 100 && observers < 2; i++) await Promise.resolve()
    expect(observers).toBe(2)
    controller.abort()
    expect(await pending).toEqual({ kind: 'cancelled' })
    expect(h.session.memoryStats().transientBytes).toBe(2 * WORKER_QUERY_BYTES.D)
    deliver(page)
    for (let i = 0; i < 30; i++) await Promise.resolve()
    expect(lateObserverBytes).toBe(2 * WORKER_QUERY_BYTES.D)
    expect(h.session.memoryStats().transientBytes).toBe(0)
  } finally { deliver(page); store.readRows = read; await h.close() }
})

it('measures each distinct style once for admission while retaining each layout record charge', async () => {
  const h = await harness('D', 1, 3)
  try {
    const reply = await h.session.viewport(h.owner, { ...h.window, box: { x: 0, y: 0, width: 192, height: 20 } }, new AbortController().signal, 'tile')
    if (!('result' in reply)) throw new Error('Expected result')
    const style = { numFmtId: 0, font: { name: 'same', size: 11 } }
    const observed = vi.fn(() => style)
    const styles = new Proxy([style], { get: (target, key, receiver) => key === '0' ? observed() : Reflect.get(target, key, receiver) })
    viewportPlanAdmission(reply, reply.result, styles, 0, 0, 1, WORKER_QUERY_BYTES.D)
    expect(observed).toHaveBeenCalledTimes(1)
  } finally { await h.close() }
})

it('a merge to the left blocks a right-aligned origin without populated covered cells', async () => {
  const h = await harness('D', 1, 6)
  try {
    const collector = new ViewportOverflowCollector(WORKER_QUERY_BYTES.D)
    collector.structure('complete', [{ range: { minRow: 0, maxRow: 1, minCol: 1, maxCol: 2 }, bounds: { x: 64, y: 0, width: 128, height: 40 } }])
    collector.run({ r0: 0, r1: 0, c0: 0, c1: 10 }, true)
    const store = h.source.store('s')
    const read = await store.readCell({ row: 0, col: 5 }, { signal: new AbortController().signal, stamp: store.stamp })
    if (read.kind !== 'cell') throw new Error('Expected origin')
    const origin = { ...read, source: { row: 0, col: 5 }, bounds: { x: 320, y: 0, width: 64, height: 20 } }
    collector.cell(read.cell)
    const steps = collector.finish(h.metrics, [origin], h.owner, { ...h.window, box: { x: 0, y: 0, width: 512, height: 20 } }, 'D', 'tile')
    let step = steps.next()
    while (!step.done) step = steps.next()
    expect(step.value.entries[0]).toMatchObject({ left: 128, x0: 192, leftEnd: 'blocker' })
  } finally { await h.close() }
})

it('bounded style-cache collisions stay conservative with many unique and repeated styles', async () => {
  const h = await harness('D')
  try {
    const reply = await h.session.viewport(h.owner, h.window, new AbortController().signal, 'tile')
    if (!('result' in reply) || reply.result.content.kind !== 'detail') throw new Error('Expected detail')
    const origin = reply.result.content.cells[0]
    if (origin.kind !== 'cell') throw new Error('Expected origin')
    const indices = [...Array.from({ length: 300 }, (_, index) => index), ...Array.from({ length: 300 }, () => 0)]
    const cells = indices.map((styleIndex, col) => ({ ...origin, cell: { ...origin.cell, styleIndex }, source: { row: 0, col } }))
    const result = { ...reply.result, content: { kind: 'detail' as const, cells } }
    const styles = Array.from({ length: 300 }, (_, index) => ({ numFmtId: index, font: { name: `style${index}`, size: 11 } }))
    let reads = 0
    const observed = new Proxy(styles, { get(target, key, receiver) { if (typeof key === 'string' && /^\d+$/.test(key)) reads++; return Reflect.get(target, key, receiver) } })
    const admission = viewportPlanAdmission({ result }, result, observed, 0, 0, 1, WORKER_QUERY_BYTES.D)
    expect(reads).toBe(301)
    const withoutStyles = viewportPlanAdmission({ result }, result, [], 0, 0, 1, WORKER_QUERY_BYTES.D)
    expect(admission.planUpper - withoutStyles.planUpper).toBeGreaterThanOrEqual(styles.reduce((sum, style) => sum + queryPayloadBytes(style, WORKER_QUERY_BYTES.D), 0))
    expect(admission.joint).toBe(8192 + admission.reserve + admission.handoff)
  } finally { await h.close() }
})


it.each([['D', 'tile', 5000], ['M', 'tile', 2500], ['D', 'frame', 20000], ['M', 'frame', 10000]] as const)('%s %s keeps its %s record boundary before byte admission', async (profile, scope, cap) => {
  const rows = cap / 100, h = await harness(profile, rows + 1, 100)
  const window = { ...h.window, box: { x: 0, y: 0, width: 68 * 64, height: rows * 20 } }
  const store = h.source.store('s'), reads = vi.spyOn(store, 'readRows')
  try {
    expect(VIEWPORT_DETAIL_LIMITS[profile][scope === 'tile' ? 'tile' : 'records']).toEqual(scope === 'tile' ? { records: cap, pages: profile === 'D' ? 8 : 4 } : cap)
    const atCap = await queryViewport(store, h.metrics, window, h.owner, { profile, scope, signal: new AbortController().signal, textDependencies: true })
    expect(atCap.kind).toBe('viewport')
    if (atCap.kind !== 'viewport') throw new Error('Expected viewport')
    expect(atCap.cells).toHaveLength(cap)
    reads.mockClear()
    const above = await queryViewport(store, h.metrics, { ...window, box: { ...window.box, height: (rows + 1) * 20 } }, h.owner, { profile, scope, signal: new AbortController().signal, textDependencies: true })
    expect(above.kind).toBe('viewport')
    if (above.kind !== 'viewport') throw new Error('Expected viewport')
    expect(above.cells).toHaveLength(0)
    expect(above.diagnostics).toContainEqual(expect.objectContaining({ code: 'detail-limit' }))
    expect(reads).not.toHaveBeenCalled()
    if (scope === 'tile') {
      const reply = await h.session.viewport(h.owner, window, new AbortController().signal, scope)
      expect(reply).toHaveProperty('result.content.cells.length', cap)
    }
  } finally { reads.mockRestore(); await h.close() }
})

it.each(['detail', 'coarse'] as const)('charges %s source pages before constructing their structural/overview output', async kind => {
  const h = await harness('D')
  const collector = new ViewportOverflowCollector(WORKER_QUERY_BYTES.D), captured = vi.spyOn(collector, 'page')
  const structural = vi.spyOn(collector, 'structure')
  const store = h.source.store('s')
  const window = kind === 'detail' ? h.window : { ...h.window, zoom: 0.05 }
  try {
    await queryViewport(store, h.metrics, window, h.owner, { profile: 'D', signal: new AbortController().signal, overflowCapture: collector })
    if (kind === 'detail') {
      expect(captured.mock.calls[0][0]).toHaveProperty('merges')
      expect(captured.mock.invocationCallOrder[0]).toBeLessThan(structural.mock.invocationCallOrder[0])
    } else {
      expect(captured).toHaveBeenCalled()
      expect(captured.mock.calls[0][0]).toHaveProperty('items')
    }
  } finally { captured.mockRestore(); structural.mockRestore(); await h.close() }
})

it.each(['D', 'M'] as const)('%s retained late producers exhaust the existing per-session pool and refuse the next workspace', async profile => {
  const h = await harness(profile), store = h.source.store('s'), original = store.readRows
  const page = await original.call(store, { r0: 0, r1: 0, c0: 0, c1: 0 }, undefined, { signal: new AbortController().signal, stamp: store.stamp })
  const producers: Array<() => void> = []
  const workspace = 2 * WORKER_QUERY_BYTES[profile], budget = (profile === 'D' ? 512 : 256) * 1024 * 1024
  try {
    await h.session.viewport(h.owner, h.window, new AbortController().signal, 'tile')
    store.readRows = () => new Promise(resolve => { producers.push(() => resolve(page)) })
    let refused = false
    for (let request = 0; request < budget / workspace + 1; request++) {
      const before = producers.length, controller = new AbortController()
      let outcome: Awaited<ReturnType<typeof h.session.viewport>> | undefined
      const pending = h.session.viewport(h.owner, h.window, controller.signal, 'tile').then(result => { outcome = result; return result })
      for (let turn = 0; turn < 100 && !outcome && producers.length === before; turn++) await Promise.resolve()
      if (producers.length === before) {
        expect(await pending).toMatchObject({ kind: 'quota', requiredBytes: workspace, limitBytes: budget })
        refused = true
        break
      }
      controller.abort()
      expect(await pending).toEqual({ kind: 'cancelled' })
      expect(h.session.memoryStats().transientBytes).toBe(producers.length * workspace)
      expect(h.session.memoryStats().totalBytes).toBeLessThanOrEqual(budget)
    }
    expect(refused).toBe(true)
    expect(producers).toHaveLength(31)
    await h.close()
    expect(h.kernel.session(h.owner.docId)).toBeUndefined()
    expect(h.session.memoryStats().transientBytes).toBe(31 * workspace)
    // A retired session keeps its own pool; this asserts no cross-document cap.
    for (const drain of producers) drain()
    for (let turn = 0; turn < 30; turn++) await Promise.resolve()
    expect(h.session.memoryStats().totalBytes).toBe(0)
  } finally {
    for (const drain of producers) drain()
    store.readRows = original
    await h.close()
  }
})
