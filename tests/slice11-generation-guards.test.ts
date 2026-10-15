// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { ViewerClient } from '../src/worker/viewerClient'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { WorkerViewportPlans } from '../src/components/xlsx/workerViewport'
import type { ViewportCell } from '../src/xlsx/view/contracts'
import { workbookOf } from './helpers/xlsx-find'

function cell(value = 'origin'): ViewportCell {
  return { kind: 'cell', source: { row: 0, col: 0 }, bounds: { x: 0, y: 0, width: 64, height: 20 }, cell: { id: 0, row: 0, col: 0, value, styleIndex: 0 } }
}
async function harness(value = 'origin', rowCacheBytes?: number) {
  const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value }, { row: 0, col: 256, value: 'x'.repeat(4096) }] }])
  const source = createObjectWorkbookSource(doc, 'guards')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const port = connectInProcess(kernel), client = new ViewerClient(port)
  await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'guards', source: { kind: 'object', document: doc }, profile: 'D' })
  const session = kernel.session('guards')!, owner = session.identity('s')
  const remote = new RemoteWorkbook(client, { docId: 'guards', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map([['s', session.layout('s')]]), rowCacheBytes })
  const store = remote.store('s'), plans = new WorkerViewportPlans(remote)
  const flush = async () => { await port.idle(); for (let i = 0; i < 50; i++) await Promise.resolve() }
  remote.merges('s'); await flush()
  return { remote, store, plans, owner, flush, close: async () => { plans.dispose(); await remote.release() } }
}

it('SLICE11 G_release refuses a held token after an ordinary oversized chunk removes its row', async () => {
  const h = await harness('origin', 2048)
  try {
    await h.store.load([0], 0, 63)
    expect(h.store.cellAt({ row: 0, col: 0 })?.value).toBe('origin')
    const token = h.store.stageViewport([cell()], new AbortController().signal)!, publish = vi.fn()
    expect(token).toBeDefined()
    // put() takes its oversized-row release path: no charge() or drop() can
    // mask release() failing to invalidate the held snapshot.
    await h.store.load([0], 256, 511)
    expect(h.store.cachedRows).toBe(0)
    expect(h.remote.ledger.bytes).toBe(0)
    expect(token.commit(publish)).toBe(false)
    expect(token.error).toMatchObject({ code: 'stale-source', retryable: true })
    expect(publish).not.toHaveBeenCalled()
    expect(h.store.cellAt({ row: 0, col: 0 })).toBeUndefined()
    expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(0)
    expect((h.remote.ledger as unknown as { reserved: number }).reserved).toBe(0)
  } finally { await h.close() }
})

it('SLICE11 G_rollback invalidates a held token even when the touched transaction never charged', async () => {
  const h = await harness('origin', 2048)
  try {
    await h.store.load([0], 0, 63)
    const before = h.store.cellsOf(0), bytes = h.remote.ledger.bytes
    const token = h.store.stageViewport([cell()], new AbortController().signal)!, publish = vi.fn()
    expect(token).toBeDefined()
    const transaction = h.remote.ledger.beginViewportPublication(h.store, [0])
    // No installation or charge occurs between this snapshot and rollback.
    transaction.rollback()
    expect(h.store.cellsOf(0)).toBe(before)
    expect(h.remote.ledger.bytes).toBe(bytes)
    expect(token.commit(publish)).toBe(false)
    expect(token.error).toMatchObject({ code: 'stale-source', retryable: true })
    expect(publish).not.toHaveBeenCalled()
    expect(h.store.cellsOf(0)).toBe(before)
    expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(bytes)
    expect((h.remote.ledger as unknown as { reserved: number }).reserved).toBe(0)
  } finally { await h.close() }
})

it('SLICE11 W_norollback restores evicted plans and layout charge after publication is refused', async () => {
  const h = await harness('E'.repeat(200000))
  const windowAt = (x: number) => ({ box: { x, y: 0, width: 192, height: 20 }, zoom: 1, deviceScale: 1 })
  const extent = { width: 257 * 64, height: 20 }
  const fetch = async (x: number) => {
    const window = windowAt(x)
    h.plans.beginFrame(h.owner, window)
    h.plans.get(h.owner, window, extent, true)
    await h.flush()
    return h.plans.get(h.owner, window, extent, true)
  }
  try {
    const oldPlans = []
    for (let i = 0; i < 20; i++) {
      const plan = await fetch(i)
      expect(plan.kind).toBe('detail')
      oldPlans.push(plan)
    }
    const layoutBytes = h.remote.mainLedger.snapshot().byPool.layout
    const actualSet = h.remote.mainLedger.set.bind(h.remote.mainLedger)
    let changed = false, nestedBytes = 0
    const set = vi.spyOn(h.remote.mainLedger, 'set').mockImplementation((pool, owner, bytes) => {
      actualSet(pool, owner, bytes)
      if (!changed && pool === 'layout' && /^viewport-plans:\d+$/.test(owner)) {
        changed = true
        h.store.adoptViewport([cell('nested new owner')])
        nestedBytes = h.remote.ledger.bytes
      }
    })
    const query = vi.spyOn(h.remote, 'viewport')
    try {
      expect(await fetch(20)).toMatchObject({ kind: 'diagnostic', diagnostics: [{ code: 'stale-source' }] })
      expect(changed).toBe(true)
      expect(h.store.cellAt({ row: 0, col: 0 })?.value).toBe('nested new owner')
      expect(h.remote.ledger.bytes).toBe(nestedBytes)
      // Retain the exact old objects, including the oldest plan selected for
      // provisional eviction; a replacement request is not a restoration.
      for (let i = 0; i < 20; i++) expect(h.plans.get(h.owner, windowAt(i), extent, true) === oldPlans[i], `Original plan ${i} must be restored`).toBe(true)
      expect(query).toHaveBeenCalledTimes(1)
      expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(layoutBytes)
      expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(nestedBytes)
      expect(h.remote.pagesInFlight).toBe(0)
      expect((h.remote.ledger as unknown as { reserved: number }).reserved).toBe(0)
    } finally { query.mockRestore(); set.mockRestore() }
  } finally { await h.close() }
})
