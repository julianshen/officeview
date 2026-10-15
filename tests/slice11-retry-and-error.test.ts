// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { ViewerClient } from '../src/worker/viewerClient'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { WorkerViewportPlans, PLAN_RETRY_MS } from '../src/components/xlsx/workerViewport'
import type { ViewportCell } from '../src/xlsx/view/contracts'
import { sameIdentity } from '../src/xlsx/view/cacheIdentity'
import { workbookOf } from './helpers/xlsx-find'

const cell = (): ViewportCell => ({ kind: 'cell', source: { row: 0, col: 0 }, bounds: { x: 0, y: 0, width: 64, height: 20 }, cell: { id: 0, row: 0, col: 0, value: 'origin', styleIndex: 0 } })
async function harness() {
  const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 'origin' }, { row: 0, col: 256, value: 'fresh chunk' }] }])
  const source = createObjectWorkbookSource(doc, 'deferred')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const port = connectInProcess(kernel), client = new ViewerClient(port)
  await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'deferred', source: { kind: 'object', document: doc }, profile: 'D' })
  const session = kernel.session('deferred')!, owner = session.identity('s')
  const remote = new RemoteWorkbook(client, { docId: 'deferred', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map([['s', session.layout('s')]]) })
  const store = remote.store('s'), plans = new WorkerViewportPlans(remote)
  const flush = async () => { await port.idle(); for (let i = 0; i < 50; i++) await Promise.resolve() }
  remote.merges('s'); await flush()
  const window = { box: { x: 0, y: 0, width: 128, height: 20 }, zoom: 1, deviceScale: 1 }, extent = { width: 257 * 64, height: 20 }
  return { remote, store, plans, owner, session, window, extent, flush, close: async () => { plans.dispose(); await remote.release() } }
}

it('SLICE11 stale refusal exposes a failure that survives copying and serialization', async () => {
  const h = await harness()
  try {
    await h.store.load([0], 0, 63)
    const token = h.store.stageViewport([cell()], new AbortController().signal)!
    await h.store.load([0], 256, 511)
    const publish = vi.fn()
    expect(token.commit(publish)).toBe(false)
    const expected = { code: 'stale-source', retryable: true, message: 'Viewport row cache changed; re-stage the publication' }
    const error = token.error!
    expect(token.error, 'repeated access preserves the exposed failure identity').toBe(error)
    expect({ ...error }).toEqual(expected)
    expect(JSON.parse(JSON.stringify(error))).toEqual(expected)
    expect(structuredClone(error)).toEqual(expected)
    expect(error).not.toBeInstanceOf(Error)
    expect(publish).not.toHaveBeenCalled()
  } finally { await h.close() }
})

it('SLICE11 a refused token without an error gets a retryable fallback and re-stages', async () => {
  const h = await harness(), query = vi.spyOn(h.remote, 'viewport'), realStage = h.store.stageViewport.bind(h.store)
  let refused = false
  const stage = vi.spyOn(h.store, 'stageViewport').mockImplementation((...args) => {
    const token = realStage(...args)
    if (!token || refused) return token
    refused = true
    return { close: token.close, commit: () => { token.close(); return false } }
  })
  try {
    h.plans.beginFrame(h.owner, h.window)
    h.plans.get(h.owner, h.window, h.extent, true)
    await h.flush()
    expect(h.plans.get(h.owner, h.window, h.extent, true)).toMatchObject({ kind: 'diagnostic', diagnostics: [{ code: 'stale-source', message: 'Viewport could not be updated. Retrying.' }] })
    expect(h.remote.pagesInFlight).toBe(0)
    expect((h.remote.ledger as unknown as { reserved: number }).reserved).toBe(0)
    await new Promise(resolve => setTimeout(resolve, PLAN_RETRY_MS + 20))
    h.plans.beginFrame(h.owner, h.window)
    h.plans.get(h.owner, h.window, h.extent, true)
    await h.flush()
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('detail')
    expect(query).toHaveBeenCalledTimes(2)
    expect(stage).toHaveBeenCalledTimes(2)
  } finally { query.mockRestore(); stage.mockRestore(); await h.close() }
})


it('LANDING a cleared source request backs off and retries like a silent stale refusal', async () => {
  const h = await harness(), viewport = h.remote.viewport.bind(h.remote)
  const query = vi.spyOn(h.remote, 'viewport').mockImplementationOnce(async (...args) => {
    const out = await viewport(...args)
    h.store.clear()
    return out
  })
  try {
    h.plans.beginFrame(h.owner, h.window)
    h.plans.get(h.owner, h.window, h.extent, true)
    await h.flush()
    expect(h.plans.get(h.owner, h.window, h.extent, true)).toMatchObject({ kind: 'diagnostic', diagnostics: [{ code: 'stale-source', message: 'Viewport source cache was cleared' }] })
    expect(query).toHaveBeenCalledTimes(1)
    expect(h.remote.pagesInFlight).toBe(0)
    await new Promise(resolve => setTimeout(resolve, PLAN_RETRY_MS + 20))
    h.plans.beginFrame(h.owner, h.window)
    h.plans.get(h.owner, h.window, h.extent, true)
    await h.flush()
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('detail')
    expect(query).toHaveBeenCalledTimes(2)
  } finally { query.mockRestore(); await h.close() }
})


it('LANDING a source cleared during staging reports the cleared cache, not an owner change, and retries', async () => {
  const h = await harness(), stageViewport = h.store.stageViewport.bind(h.store)
  // Clear after the post-reply live check but before the post-stage guard:
  // the guard fails only because the source cache was cleared.
  const stage = vi.spyOn(h.store, 'stageViewport').mockImplementationOnce((...args) => {
    h.store.clear()
    return stageViewport(...args)
  })
  try {
    h.plans.beginFrame(h.owner, h.window)
    h.plans.get(h.owner, h.window, h.extent, true)
    await h.flush()
    expect(stage).toHaveBeenCalledTimes(1)
    expect(sameIdentity(h.remote.identity('s'), h.owner), 'the owner itself did not change').toBe(true)
    expect(h.plans.get(h.owner, h.window, h.extent, true)).toMatchObject({ kind: 'diagnostic', diagnostics: [{ code: 'stale-source', message: 'Viewport source cache was cleared' }] })
    expect(h.remote.pagesInFlight).toBe(0)
    expect((h.remote.ledger as unknown as { reserved: number }).reserved).toBe(0)
    await new Promise(resolve => setTimeout(resolve, PLAN_RETRY_MS + 20))
    h.plans.beginFrame(h.owner, h.window)
    h.plans.get(h.owner, h.window, h.extent, true)
    await h.flush()
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('detail')
    expect(stage).toHaveBeenCalledTimes(2)
  } finally { stage.mockRestore(); await h.close() }
})

it('LANDING retries a mid-flight worker font change while the host font remains unchanged', async () => {
  const h = await harness(), viewport = h.remote.viewport.bind(h.remote)
  const identity = vi.spyOn(h.remote, 'identity')
  const query = vi.spyOn(h.remote, 'viewport').mockImplementationOnce(async (...args) => {
    const out = await viewport(...args)
    // Deliver the old reply only after the worker and its remote mirror have
    // moved to new fonts; the frame's independent host owner stays the same.
    h.session.publishFonts()
    identity.mockReturnValue(h.session.identity('s'))
    return out
  })
  try {
    h.plans.beginFrame(h.owner, h.window)
    h.plans.get(h.owner, h.window, h.extent, true)
    await h.flush()
    expect(h.owner.fontGen).toBe(0)
    expect(h.remote.identity('s').fontGen).toBe(1)
    expect(h.plans.get(h.owner, h.window, h.extent, true)).toMatchObject({ kind: 'diagnostic', diagnostics: [{ code: 'stale-source', message: 'Viewport owner changed. Retrying.' }] })
    expect(h.remote.pagesInFlight).toBe(0)
    expect((h.remote.ledger as unknown as { reserved: number }).reserved).toBe(0)
    await new Promise(resolve => setTimeout(resolve, PLAN_RETRY_MS + 20))
    h.plans.beginFrame(h.owner, h.window)
    h.plans.get(h.owner, h.window, h.extent, true)
    await h.flush()
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('detail')
    expect(query).toHaveBeenCalledTimes(2)
    expect(query.mock.calls.map(([owner]) => owner.fontGen)).toEqual([0, 1])
    expect(h.owner.fontGen).toBe(0)
    expect(h.store.cellAt({ row: 0, col: 0 })?.value).toBe('origin')
  } finally { query.mockRestore(); identity.mockRestore(); await h.close() }
})
