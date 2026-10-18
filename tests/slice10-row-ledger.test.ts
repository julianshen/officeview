// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { ViewerClient } from '../src/worker/viewerClient'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { RemoteWorkbook, RowCacheLedger, cachedCellBytes } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { WorkerViewportPlans } from '../src/components/xlsx/workerViewport'
import type { ViewportCell } from '../src/xlsx/view/contracts'
import { workbookOf } from './helpers/xlsx-find'

function cell(row: number, col = 0, value: string | number = row): ViewportCell {
  return { kind: 'cell', source: { row, col }, bounds: { x: col * 64, y: row * 20, width: 64, height: 20 }, cell: { id: row * 16384 + col, row, col, value, styleIndex: 0 } }
}
async function harness(profile: 'D' | 'M' = 'D', rowCacheBytes?: number, dense: boolean | string = false) {
  const cells = typeof dense === 'string' ? [{ row: 0, col: 0, value: dense }] : dense ? Array.from({ length: 120 }, (_, row) => Array.from({ length: 170 }, (_, col) => ({ row, col, value: `${row}` }))).flat() : Array.from({ length: 400 }, (_, row) => ({ row, col: 0, value: `row${row}${'x'.repeat(40)}` }))
  const doc = workbookOf([{ sheetId: 's', cells }, { sheetId: 't', cells: [{ row: 0, col: 0, value: 1 }] }])
  const source = createObjectWorkbookSource(doc, 'doc')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const port = connectInProcess(kernel), client = new ViewerClient(port)
  await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'doc', source: { kind: 'object', document: doc }, profile })
  const session = kernel.session('doc')!, owner = session.identity('s')
  const remote = new RemoteWorkbook(client, { profile, docId: 'doc', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map(source.sheets.map(s => [s.sheetId, session.layout(s.sheetId)])), rowCacheBytes })
  const store = remote.store('s'), plans = new WorkerViewportPlans(remote)
  const flush = async () => { await port.idle(); for (let i = 0; i < 50; i++) await Promise.resolve() }
  remote.merges('s'); await flush()
  const extent = { width: 170 * 64, height: 400 * 20 }
  return { remote, store, plans, owner, flush, extent, close: async () => { plans.dispose(); await remote.release() } }
}

it.each(['D', 'M'] as const)('SLICE10 %s dense bridge evicts only retained growth plus one warm row', async profile => {
  const h = await harness(profile, undefined, true)
  try {
    const ledger = h.remote.ledger, limit = ledger.limitBytes, chunk = 4096
    const dropped: number[] = []
    const warm = { sheetId: 'warm', drop: (row: number) => dropped.push(row) } as unknown as Parameters<RowCacheLedger['charge']>[0]
    for (let row = 0; row < Math.floor((limit - 2048) / chunk); row++) ledger.charge(warm, row, chunk)
    const before = ledger.bytes, zoom = profile === 'D' ? 0.6 : 1
    const w = { box: { x: 0, y: 0, width: 512 / zoom, height: 512 / zoom }, zoom, deviceScale: 1 }
    h.plans.beginFrame(h.owner, w)
    expect(h.plans.get(h.owner, w, h.extent, true).kind).toBe('pending')
    await h.flush()
    expect(h.plans.get(h.owner, w, h.extent, true).kind).toBe('detail')
    let retained = 0
    for (let row = 0; row < 200; row++) {
      const cells = h.store.cellsOf(row)
      if (cells) retained += 64 + cells.reduce((sum, c) => sum + cachedCellBytes(c), 0)
    }
    expect(retained).toBeGreaterThan(100000)
    const needed = Math.max(0, retained - (limit - before))
    expect(dropped.length * chunk).toBeGreaterThanOrEqual(needed)
    expect(dropped.length * chunk).toBeLessThan(needed + chunk)
    expect(ledger.bytes).toBe(before - dropped.length * chunk + retained)
    expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(ledger.bytes)
    expect(h.remote.mainLedger.peakBytes).toBeLessThanOrEqual((profile === 'D' ? 64 : 32) * 1024 * 1024)
  } finally { await h.close() }
})

it.each(['close', 'abort', 'refuse'] as const)('SLICE10 %s stage preserves a near-full warm cache before publication', async outcome => {
  const h = await harness('D', 4096)
  try {
    await h.store.load(Array.from({ length: 30 }, (_, row) => 200 + row), 0, 10)
    const before = h.remote.ledger.bytes, rows = h.store.cachedRows
    expect(before).toBeGreaterThan(3800)
    const controller = new AbortController()
    const stage = h.store.stageViewport([cell(1000, 0, outcome === 'refuse' ? 'x'.repeat(10000) : 7)], controller.signal)
    if (outcome === 'refuse') expect(stage).toBeUndefined()
    else {
      expect(stage).toBeDefined()
      expect(h.remote.ledger.bytes).toBe(before)
      if (outcome === 'abort') controller.abort()
      stage!.close()
    }
    expect(h.store.cachedRows).toBe(rows)
    expect(h.remote.ledger.bytes).toBe(before)
    expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(before)
  } finally { await h.close() }
})

it('SLICE10 refused pinned reservation leaves evictable rows untouched', () => {
  const ledger = new RowCacheLedger(1000), dropped: number[] = []
  const store = { sheetId: 's', drop: (row: number) => dropped.push(row) } as unknown as Parameters<RowCacheLedger['charge']>[0]
  ledger.charge(store, 0, 800)
  ledger.charge(store, 1, 100)
  expect(ledger.reserve(300, store, [0])).toBeUndefined()
  expect(dropped).toEqual([])
  expect(ledger.bytes).toBe(900)
})

it('SLICE10 exact reviewer deferred pin probe never resurrects uncharged cached data', async () => {
  const h = await harness('D', 4096)
  try {
    h.store.adoptViewport([cell(10)])
    const stage = h.store.stageViewport([cell(10, 1)], new AbortController().signal)!
    expect(stage).toBeDefined()
    const reserved = (h.remote.ledger as unknown as { reserved: number }).reserved
    h.remote.ledger.charge(h.store, 10, 4096 - reserved + 1)
    const publish = vi.fn()
    expect(stage.commit(publish)).toBe(false)
    expect(publish).not.toHaveBeenCalled()
    expect(stage.error).toMatchObject({ code: 'stale-source', retryable: true })
    stage.close()
    expect(h.store.cachedRows).toBe(0)
    expect(h.store.cellsOf(10)).toBeUndefined()
    expect(h.remote.ledger.bytes).toBe(0)
    expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(0)
  } finally { await h.close() }
})

it('SLICE10 rollback rotates with one iterator regardless of row count', () => {
  const ledger = new RowCacheLedger(2000000)
  const store = { sheetId: 's', drop: () => {} } as unknown as Parameters<RowCacheLedger['charge']>[0]
  for (let row = 0; row < 10000; row++) ledger.charge(store, row, 100)
  const transaction = ledger.beginViewportPublication(store, [0])
  ledger.charge(store, 0, 101)
  const entries = (ledger as unknown as { entries: Map<string, unknown> }).entries
  const iterators = vi.spyOn(entries, 'entries')
  transaction.rollback()
  expect(iterators.mock.calls.length).toBeLessThanOrEqual(1)
  expect(ledger.bytes).toBe(1000000)
  iterators.mockRestore()
})

it('SLICE10 rollback with no restored rows never rotates unrelated entries', () => {
  const ledger = new RowCacheLedger(1000)
  const store = { sheetId: 's', drop: () => {} } as unknown as Parameters<RowCacheLedger['charge']>[0]
  ledger.charge(store, 0, 100)
  const transaction = ledger.beginViewportPublication(store, [1])
  ledger.charge(store, 1, 100)
  const entries = (ledger as unknown as { entries: Map<string, unknown> }).entries
  const iterators = vi.spyOn(entries, 'entries')
  transaction.rollback()
  expect(iterators).not.toHaveBeenCalled()
  expect(ledger.bytes).toBe(100)
  iterators.mockRestore()
})

it('SLICE10 MJ3pin two-row eviction preserves an older touched row', () => {
  const ledger = new RowCacheLedger(1000), dropped: number[] = []
  const store = { sheetId: 's', drop: (row: number) => dropped.push(row) } as unknown as Parameters<RowCacheLedger['charge']>[0]
  ledger.charge(store, 0, 200)
  ledger.charge(store, 1, 200)
  const reservation = ledger.reserve(600, store, [0])!
  expect(reservation).toBeDefined()
  ledger.charge(store, 1, 500)
  expect(dropped).toEqual([1])
  expect(ledger.bytes).toBe(200)
  reservation.close()
})

it('SLICE10 MJ3chargeIgnoresRes two-row eviction retains the new ordinary row', () => {
  const ledger = new RowCacheLedger(1000), dropped: number[] = []
  const store = { sheetId: 's', drop: (row: number) => dropped.push(row) } as unknown as Parameters<RowCacheLedger['charge']>[0]
  ledger.charge(store, 0, 300)
  const reservation = ledger.reserve(600)!
  ledger.charge(store, 1, 300)
  expect(dropped).toEqual([0])
  expect(ledger.bytes).toBe(300)
  reservation.close()
})

it('SLICE10 reserves merged retained growth once for duplicate coordinates', async () => {
  const h = await harness('D', 4096)
  try {
    h.store.adoptViewport([cell(10, 0, 'x'.repeat(100))])
    const before = h.remote.ledger.bytes
    const stage = h.store.stageViewport([cell(10, 1, 'large'.repeat(100)), cell(10, 1, 7), cell(11)], new AbortController().signal)!
    expect(stage).toBeDefined()
    const retained = cachedCellBytes({ value: 7 }) + 64 + cachedCellBytes({ value: 11 })
    expect((h.remote.ledger as unknown as { reserved: number }).reserved).toBe(retained)
    expect(stage.commit(() => {})).toBe(true)
    expect(h.remote.ledger.bytes).toBe(before + retained)
    expect(h.store.cellsOf(10)?.map(c => c.value)).toEqual(['x'.repeat(100), 7])
  } finally { await h.close() }
})

it('SLICE10 shrink before growth preserves warm rows when merged retained growth is zero', async () => {
  const h = await harness('D', 4096)
  try {
    h.store.adoptViewport([cell(0), cell(1, 0, 'x'.repeat(1000)), cell(2, 0, 'y'.repeat(800))])
    const warm = h.store.cellAt({ row: 2, col: 0 }), before = h.remote.ledger.bytes
    expect(before).toBe(4080)
    // Incoming row order grows row0 before row1 shrinks by the same amount.
    const stage = h.store.stageViewport([cell(0, 0, 'x'.repeat(1000)), cell(1)], new AbortController().signal)!
    expect(stage).toBeDefined()
    expect((h.remote.ledger as unknown as { reserved: number }).reserved).toBe(0)
    expect(stage.commit(() => {})).toBe(true)
    expect(h.remote.ledger.bytes).toBe(before)
    expect(h.store.cellAt({ row: 2, col: 0 })).toBe(warm)
  } finally { await h.close() }
})

it('SLICE10 stale deferred publication preserves the independently newer charged owner', async () => {
  const h = await harness('D', 4096)
  try {
    h.store.adoptViewport([cell(10)])
    const stage = h.store.stageViewport([cell(10, 1)], new AbortController().signal)!
    h.store.adoptViewport([cell(10, 0, 'newer')])
    const current = h.store.cellsOf(10), bytes = h.remote.ledger.bytes
    const publish = vi.fn()
    expect(stage.commit(publish)).toBe(false)
    expect(publish).not.toHaveBeenCalled()
    expect(stage.error).toMatchObject({ code: 'stale-source', retryable: true })
    expect(h.store.cellsOf(10)).toBe(current)
    expect(h.remote.ledger.bytes).toBe(bytes)
    expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(bytes)
  } finally { await h.close() }
})

it('SLICE10 rollback preserves a newer public owner installed during publication', async () => {
  const h = await harness('D', 4096)
  try {
    h.store.adoptViewport([cell(10)])
    const stage = h.store.stageViewport([cell(10, 1)], new AbortController().signal)!
    let current: ReturnType<typeof h.store.cellsOf>, bytes = 0
    expect(() => stage.commit(() => {
      h.store.adoptViewport([cell(10, 0, 'newer')])
      current = h.store.cellsOf(10)
      bytes = h.remote.ledger.bytes
      throw new Error('publication failed')
    })).toThrow('publication failed')
    expect(h.store.cellsOf(10)).toBe(current!)
    expect(h.remote.ledger.bytes).toBe(bytes)
    expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(bytes)
  } finally { await h.close() }
})

it('SLICE10 retained reservation subtracts current ledger ownership while a stage waits', async () => {
  const h = await harness('D', 4096)
  try {
    h.store.adoptViewport([cell(10)])
    h.remote.ledger.charge(h.store, 10, 500)
    const stage = h.store.stageViewport([cell(10, 1)], new AbortController().signal)!
    expect(stage).toBeDefined()
    expect((h.remote.ledger as unknown as { reserved: number }).reserved).toBe(0)
    expect(stage.commit(() => {})).toBe(true)
    expect(h.remote.ledger.bytes).toBe(256)
    expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(256)
  } finally { await h.close() }
})

it('SLICE10 aggregate pins and reservations cannot promise occupied retained capacity', () => {
  const ledger = new RowCacheLedger(1000), dropped: number[] = []
  const store = { sheetId: 's', drop: (row: number) => dropped.push(row) } as unknown as Parameters<RowCacheLedger['charge']>[0]
  ledger.charge(store, 0, 300)
  ledger.charge(store, 1, 300)
  const first = ledger.reserve(300, store, [0])!
  expect(first).toBeDefined()
  expect(ledger.reserve(200, store, [1])).toBeUndefined()
  expect(dropped).toEqual([])
  expect(ledger.bytes).toBe(600)
  first.close()
  const second = ledger.reserve(200, store, [1])!
  expect(second).toBeDefined()
  second.close()
  expect(ledger.reserve(1000)).toBeDefined()
})
