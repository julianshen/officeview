// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { ViewerClient } from '../src/worker/viewerClient'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { WorkerViewportPlans, PLAN_RETRY_MS } from '../src/components/xlsx/workerViewport'
import type { ViewportCell } from '../src/xlsx/view/contracts'
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
  return { remote, store, plans, owner, window, extent, flush, close: async () => { plans.dispose(); await remote.release() } }
}

it('SLICE10 deferred cross-chunk publication refuses stale state and re-stages without losing data', async () => {
  const h = await harness()
  try {
    await h.store.load([0], 0, 63)
    const stage = h.store.stageViewport([cell()], new AbortController().signal)!
    expect(stage).toBeDefined()
    await h.store.load([0], 256, 511)
    const current = h.store.cellsOf(0), bytes = h.remote.ledger.bytes, publish = vi.fn()
    expect(h.store.cellAt({ row: 0, col: 256 })?.value).toBe('fresh chunk')
    expect(stage.commit(publish)).toBe(false)
    expect(stage.error).toMatchObject({ code: 'stale-source', retryable: true, message: expect.stringMatching(/cache changed/i) })
    expect(publish).not.toHaveBeenCalled()
    expect(h.store.cellsOf(0)).toBe(current)
    expect(h.store.isKnown({ row: 0, col: 256 })).toBe(true)
    expect(h.store.cellAt({ row: 0, col: 256 })?.value).toBe('fresh chunk')
    expect(h.remote.ledger.bytes).toBe(bytes)
    expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(bytes)
    expect((h.remote.ledger as unknown as { reserved: number }).reserved).toBe(0)
    const next = h.store.stageViewport([cell()], new AbortController().signal)!
    expect(next.commit(publish)).toBe(true)
    expect(next.error).toBeUndefined()
    expect(publish).toHaveBeenCalledTimes(1)
    expect(h.store.cellAt({ row: 0, col: 256 })?.value).toBe('fresh chunk')
  } finally { await h.close() }
})

it('SLICE10 an unchanged deferred token commits across reads, LRU touches and reservation release', async () => {
  const h = await harness()
  try {
    await h.store.load([0], 0, 63)
    const stage = h.store.stageViewport([cell()], new AbortController().signal)!, publish = vi.fn()
    await Promise.resolve()
    h.store.cellsOf(0)
    h.remote.ledger.reserve(1)!.close()
    expect(stage.commit(publish)).toBe(true)
    expect(stage.error).toBeUndefined()
    expect(publish).toHaveBeenCalledTimes(1)
    expect(h.store.cellAt({ row: 0, col: 0 })?.value).toBe('origin')
  } finally { await h.close() }
})

it.each(['charge', 'release-recharge', 'drop', 'rollback'] as const)('SLICE10 deferred token refuses a touched-row %s without publication', async change => {
  const h = await harness()
  try {
    await h.store.load([0], 0, 63)
    const stage = h.store.stageViewport([cell()], new AbortController().signal)!, publish = vi.fn(), bytes = h.remote.ledger.bytes
    if (change === 'drop') h.store.drop(0)
    else if (change === 'release-recharge') { h.remote.ledger.release(h.store, 0); h.remote.ledger.charge(h.store, 0, bytes) }
    else if (change === 'rollback') {
      const transaction = h.remote.ledger.beginViewportPublication(h.store, [0])
      h.remote.ledger.charge(h.store, 0, bytes)
      transaction.rollback()
    } else h.remote.ledger.charge(h.store, 0, bytes)
    const current = h.store.cellsOf(0)
    expect(stage.commit(publish)).toBe(false)
    expect(stage.error).toMatchObject({ code: 'stale-source', retryable: true })
    expect(publish).not.toHaveBeenCalled()
    expect(h.store.cellsOf(0)).toBe(current)
    expect(h.remote.ledger.bytes).toBe(bytes)
    expect((h.remote.ledger as unknown as { reserved: number }).reserved).toBe(0)
  } finally { h.remote.ledger.release(h.store, 0); await h.close() }
})

it('SLICE10 synchronous WorkerViewportPlans publication remains a single cached detail request', async () => {
  const h = await harness(), query = vi.spyOn(h.remote, 'viewport'), stage = vi.spyOn(h.store, 'stageViewport')
  try {
    h.plans.beginFrame(h.owner, h.window)
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('pending')
    await h.flush()
    const plan = h.plans.get(h.owner, h.window, h.extent, true)
    expect(plan.kind).toBe('detail')
    h.plans.beginFrame(h.owner, h.window)
    expect(h.plans.get(h.owner, h.window, h.extent, true)).toBe(plan)
    expect(query).toHaveBeenCalledTimes(1)
    expect(stage).toHaveBeenCalledTimes(1)
    expect(h.store.cellAt({ row: 0, col: 0 })?.value).toBe('origin')
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { query.mockRestore(); stage.mockRestore(); await h.close() }
})

it('SLICE10 WorkerViewportPlans retries a stale publication by requesting and staging again', async () => {
  const h = await harness(), query = vi.spyOn(h.remote, 'viewport'), realStage = h.store.stageViewport.bind(h.store)
  let changed = false
  const stage = vi.spyOn(h.store, 'stageViewport').mockImplementation((...args) => {
    const token = realStage(...args)
    if (!changed && token) { changed = true; h.remote.ledger.charge(h.store, 0, h.remote.ledger.rowBytes(h.store, 0) ?? 64) }
    return token
  })
  try {
    h.plans.beginFrame(h.owner, h.window)
    h.plans.get(h.owner, h.window, h.extent, true)
    await h.flush()
    expect(h.plans.get(h.owner, h.window, h.extent, true)).toMatchObject({ kind: 'diagnostic', diagnostics: [{ code: 'stale-source' }] })
    await new Promise(resolve => setTimeout(resolve, PLAN_RETRY_MS + 20))
    h.plans.beginFrame(h.owner, h.window)
    h.plans.get(h.owner, h.window, h.extent, true)
    await h.flush()
    expect(h.plans.get(h.owner, h.window, h.extent, true).kind).toBe('detail')
    expect(query).toHaveBeenCalledTimes(2)
    expect(stage).toHaveBeenCalledTimes(2)
    expect(h.remote.pagesInFlight).toBe(0)
  } finally { query.mockRestore(); stage.mockRestore(); await h.close() }
})


it('SLICE10 stage installation invalidates another held token before reentrant publication', async () => {
  const h = await harness()
  try {
    await h.store.load([0], 0, 63)
    const first = h.store.stageViewport([cell()], new AbortController().signal)!
    const second = h.store.stageViewport([cell()], new AbortController().signal)!, publish = vi.fn()
    expect(first.commit(() => {
      expect(second.commit(publish)).toBe(false)
      expect(second.error).toMatchObject({ code: 'stale-source', retryable: true })
    })).toBe(true)
    expect(publish).not.toHaveBeenCalled()
  } finally { await h.close() }
})

it('SLICE10 reentrant cache mutation refuses the outer token without overwriting newer charges', async () => {
  const h = await harness()
  try {
    await h.store.load([0], 0, 63)
    const stage = h.store.stageViewport([cell()], new AbortController().signal)!
    let current: ReturnType<typeof h.store.cellsOf>, bytes = 0
    expect(stage.commit(() => {
      const newer = cell()
      if (newer.kind !== 'cell') throw new Error('expected cell')
      newer.cell.value = 'newer'.repeat(100)
      h.store.adoptViewport([newer])
      current = h.store.cellsOf(0)
      bytes = h.remote.ledger.bytes
    })).toBe(false)
    expect(stage.error).toMatchObject({ code: 'stale-source', retryable: true })
    expect(h.store.cellsOf(0)).toBe(current!)
    expect(h.remote.ledger.bytes).toBe(bytes)
    expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(bytes)
  } finally { await h.close() }
})
