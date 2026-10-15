// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { ViewerClient } from '../src/worker/viewerClient'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { RemoteWorkbook, RowCacheLedger, cachedCellBytes } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { WorkerViewportPlans } from '../src/components/xlsx/workerViewport'
import { queryPayloadBytes, WORKER_QUERY_BYTES } from '../src/worker/viewerQueryBudget'
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

it('SLICE9 counts old cells once per touched row and visits linear old-row work', async () => {
  const h = await harness()
  try {
    const original = Array.from({ length: 200 }, (_, col) => cell(0, col))
    h.store.adoptViewport([...original, cell(3)])
    const old = h.store.cellsOf(0)!, other = h.store.cellsOf(3)!
    // The worker may return source rows in display order after sorting.
    const next = Array.from({ length: 100 }, (_, col) => [cell(0, col + 200), cell(3, col + 200)]).flat()
    let visits = 0
    for (const row of [old, other]) {
      const iterator = row[Symbol.iterator].bind(row)
      Object.defineProperty(row, Symbol.iterator, { configurable: true, value: function* () { for (const item of iterator()) { visits++; yield item } } })
    }
    const expected = 128 + next.length * 192 + next.reduce((n, c) => n + 256 + 3 * cachedCellBytes('cell' in c ? c.cell : { value: null }), 0) + [...old, ...other].reduce((n, c) => n + 128 + cachedCellBytes(c), 0)
    visits = 0
    expect(h.store.viewportAdoptionBytes(next)).toBe(expected)
    expect(visits).toBe(old.length + other.length)
  } finally { await h.close() }
})

it('SLICE9 full LRU still adopts a real tile for inspection and accessibility', async () => {
  const h = await harness('D', 16384)
  try {
    await h.store.load(Array.from({ length: 200 }, (_, i) => 200 + i), 0, 10)
    expect(h.remote.ledger.bytes).toBeGreaterThan(16000)
    const oldRows = h.store.cachedRows
    const w = { box: { x: 0, y: 0, width: 128, height: 20 }, zoom: 1, deviceScale: 1 }
    h.plans.beginFrame(h.owner, w); expect(h.plans.get(h.owner, w, h.extent, true).kind).toBe('pending'); await h.flush()
    expect(h.plans.get(h.owner, w, h.extent, true).kind).toBe('detail')
    expect(h.store.cellsOf(0)?.map(c => c.value)).toEqual([`row0${'x'.repeat(40)}`])
    expect(h.store.cellAt({ row: 0, col: 0 })?.value).toBe(`row0${'x'.repeat(40)}`)
    // One same-sized adopted row replaces exactly one warm row at capacity.
    expect(h.store.cachedRows).toBe(oldRows)
    expect(h.remote.ledger.bytes).toBeLessThanOrEqual(16384)
    h.store.clear(); expect(h.remote.ledger.bytes).toBe(0)
  } finally { await h.close() }
})

it('SLICE9 other charges preserve promised reservation capacity', () => {
  const ledger = new RowCacheLedger(1000)
  const dropped: number[] = []
  const store = { sheetId: 's', drop: (row: number) => dropped.push(row) } as unknown as Parameters<RowCacheLedger['charge']>[0]
  const reserved = ledger.reserve(600)!
  ledger.charge(store, 0, 500)
  expect(ledger.bytes).toBeLessThanOrEqual(400)
  expect(dropped).toEqual([0])
  reserved.close()
  ledger.charge(store, 1, 1000); expect(ledger.bytes).toBe(1000)
})

it('SLICE9 rollback cannot recreate an evicted row charge', async () => {
  const h = await harness('D', 4096)
  try {
    h.store.adoptViewport([cell(10)])
    const live = h.store.cellsOf(10), stage = h.store.stageViewport([cell(0)], new AbortController().signal)!
    const actualCharge = h.remote.ledger.charge.bind(h.remote.ledger)
    const spy = vi.spyOn(h.remote.ledger, 'charge').mockImplementation((...args) => {
      actualCharge(...args)
      if (args[1] === 0) { h.store.drop(10); h.remote.ledger.release(h.store, 10); throw new Error('after independent eviction') }
    })
    expect(() => stage.commit(() => {})).toThrow('after independent eviction')
    spy.mockRestore(); stage.close()
    expect(live?.length).toBe(1)
    expect(h.store.cellsOf(10)).toBeUndefined()
    expect(h.store.cellsOf(0)).toBeUndefined()
    expect(h.remote.ledger.bytes).toBe(0)
    h.store.clear(); expect(h.remote.mainLedger.snapshot().byPool.rowCache).toBe(0)
  } finally { await h.close() }
})

it('SLICE9 activate clears an in-flight reply into a settled stale-source tile', async () => {
  const h = await harness()
  try {
    h.remote.store('t')
    const w = { box: { x: 0, y: 0, width: 128, height: 20 }, zoom: 1, deviceScale: 1 }
    h.plans.beginFrame(h.owner, w); expect(h.plans.get(h.owner, w, h.extent, true).kind).toBe('pending')
    h.remote.activate('t'); await h.flush()
    expect(h.plans.get(h.owner, w, h.extent, true)).toMatchObject({ kind: 'diagnostic', diagnostics: [{ code: 'stale-source' }] })
    expect(h.store.cachedRows).toBe(0); expect(h.remote.pagesInFlight).toBe(0)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
  } finally { await h.close() }
})

it('SLICE9 publication accepts field-equivalent identity with different key order', async () => {
  const h = await harness()
  const owner = { viewGen: h.owner.viewGen, fontGen: h.owner.fontGen, valueGen: h.owner.valueGen, sheetId: h.owner.sheetId, docId: h.owner.docId }
  try {
    const w = { box: { x: 0, y: 0, width: 128, height: 20 }, zoom: 1, deviceScale: 1 }
    h.plans.beginFrame(owner, w); h.plans.get(owner, w, h.extent, true); await h.flush()
    expect(h.plans.get(owner, w, h.extent, true).kind).toBe('detail')
    expect(h.store.cellAt({ row: 0, col: 0 })).toBeDefined()
  } finally { await h.close() }
})

it.each(['D', 'M'] as const)('SLICE9 real %s adjacent dense tiles paint and adopt within unchanged main budgets', async profile => {
  const h = await harness(profile, undefined, true)
  const requests = vi.spyOn(h.remote, 'viewport')
  try {
    const zoom = profile === 'D' ? 0.6 : 1, size = 512 / zoom
    for (const tile of [2, 3, 4, 5]) {
      const w = { box: { x: tile * size, y: 0, width: size, height: size }, zoom, deviceScale: 1 }
      h.plans.beginFrame(h.owner, w)
      expect(h.plans.get(h.owner, w, h.extent, true).kind).toBe('pending')
      await h.flush()
      const plan = h.plans.get(h.owner, w, h.extent, true)
      expect(plan.kind, JSON.stringify(plan.diagnostics)).toBe('detail')
      if (plan.kind !== 'detail') throw new Error('Expected real dense tile')
      expect(plan.pending).toBeUndefined()
      const painted = plan.cells.filter(c => c.x < w.box.x + size && c.x + c.width > w.box.x)
      expect(painted.length).toBeGreaterThan(100)
      for (const placed of painted) expect(h.store.cellAt(placed.source)?.value).toBe(placed.cell.value)
      expect(h.remote.mainLedger.peakBytes).toBeLessThanOrEqual((profile === 'D' ? 64 : 32) * 1024 * 1024)
      expect(h.remote.pagesInFlight).toBe(0)
    }
    expect(requests).toHaveBeenCalledTimes(4)
  } finally { requests.mockRestore(); await h.close() }
})

it.each(['D', 'M'] as const)('SLICE9 %s selected plan eviction rollback retains live old data and release owners', async profile => {
  const h = await harness(profile, undefined, 'E'.repeat(200000))
  const count = profile === 'D' ? 20 : 10, cap = WORKER_QUERY_BYTES[profile]
  const old = []
  const windowAt = (x: number) => ({ box: { x, y: 0, width: 192, height: 20 }, zoom: 1, deviceScale: 1 })
  const fetch = async (x: number) => { const w = windowAt(x); h.plans.beginFrame(h.owner, w); h.plans.get(h.owner, w, h.extent, true); await h.flush(); return h.plans.get(h.owner, w, h.extent, true) }
  try {
    for (let i = 0; i < count; i++) { const plan = await fetch(i); expect(plan.kind).toBe('detail'); old.push(plan) }
    const before = h.remote.mainLedger.snapshot(), oldRow = h.store.cellsOf(0)
    const firstBytes = queryPayloadBytes(old[0], cap)
    expect(before.byPool.layout).toBeLessThan(cap)
    expect(before.byPool.layout + firstBytes).toBeGreaterThan(cap)
    let failed = false
    const actualCharge = h.remote.ledger.charge.bind(h.remote.ledger)
    const charge = vi.spyOn(h.remote.ledger, 'charge').mockImplementation((...args) => {
      actualCharge(...args)
      if (!failed) { failed = true; throw new Error('rollback selected eviction') }
    })
    try {
      expect(await fetch(count)).toMatchObject({ kind: 'diagnostic', diagnostics: [{ message: 'rollback selected eviction' }] })
      expect(failed).toBe(true)
      expect(h.store.cellsOf(0)).toBe(oldRow)
      expect(h.remote.ledger.bytes).toBe(before.byPool.rowCache)
      for (let i = 0; i < count; i++) expect(h.plans.get(h.owner, windowAt(i), h.extent, true)).toBe(old[i])
      expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(before.byPool.layout)
      h.plans.release(h.owner, windowAt(0))
      expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(before.byPool.layout - firstBytes)
      expect(h.remote.pagesInFlight).toBe(0)
    } finally { charge.mockRestore() }
  } finally { await h.close() }
})

it('SLICE9 an evicted staged owner stays gone while unrelated live owners remain accounted', async () => {
  const h = await harness('D', 4096)
  try {
    h.store.adoptViewport([cell(0), cell(10)])
    const unrelated = h.store.cellsOf(10), before = h.remote.ledger.bytes
    const stage = h.store.stageViewport([cell(0, 1)], new AbortController().signal)!
    const actualCharge = h.remote.ledger.charge.bind(h.remote.ledger)
    const charge = vi.spyOn(h.remote.ledger, 'charge').mockImplementation((...args) => {
      actualCharge(...args)
      h.store.drop(0)
      throw new Error('staged row evicted')
    })
    try { expect(() => stage.commit(() => {})).toThrow('staged row evicted') }
    finally { charge.mockRestore(); stage.close() }
    expect(h.store.cellsOf(0)).toBeUndefined()
    expect(h.store.cellsOf(10)).toBe(unrelated)
    expect(h.remote.ledger.bytes).toBe(before / 2)
    h.store.clear(); expect(h.remote.ledger.bytes).toBe(0)
  } finally { await h.close() }
})

it('SLICE9 a truly oversized row declines caching without evicting a fitting live row', async () => {
  const h = await harness('D', 4096)
  try {
    h.store.adoptViewport([cell(10)])
    const retained = h.store.cellsOf(10), bytes = h.remote.ledger.bytes
    h.store.adoptViewport([cell(0, 0, 'x'.repeat(4096))])
    expect(h.store.cellsOf(0)).toBeUndefined()
    expect(h.store.cellsOf(10)).toBe(retained)
    expect(h.remote.ledger.bytes).toBe(bytes)
  } finally { await h.close() }
})

it.each(['none', 'same', 'other'] as const)('SLICE9 touched-row rollback preserves old LRU unless a public touch supersedes it (%s)', async touched => {
  const h = await harness('D', 4096)
  try {
    h.store.adoptViewport([cell(0), cell(1), cell(2)])
    const before = h.remote.ledger.bytes
    expect(before).toBe(480)
    const stage = h.store.stageViewport([cell(0, 1)], new AbortController().signal)!
    const actualCharge = h.remote.ledger.charge.bind(h.remote.ledger)
    const charge = vi.spyOn(h.remote.ledger, 'charge').mockImplementation((...args) => {
      actualCharge(...args)
      if (touched !== 'none') h.store.cellsOf(touched === 'same' ? 0 : 1)
      throw new Error('rollback prior recency')
    })
    try { expect(() => stage.commit(() => {})).toThrow('rollback prior recency') }
    finally { charge.mockRestore(); stage.close() }
    expect(h.remote.ledger.bytes).toBe(before)
    // A real retained charge forces exactly one 160 B old row out.
    const pressure = { sheetId: 'pressure', drop: () => {} } as unknown as Parameters<RowCacheLedger['charge']>[0]
    h.remote.ledger.charge(pressure, 0, 4096 - before + 1)
    expect(h.store.cellAt({ row: touched === 'same' ? 1 : 0, col: 0 })).toBeUndefined()
    expect(h.store.cellAt({ row: touched === 'same' ? 0 : 1, col: 0 })).toBeDefined()
    expect(h.store.cellAt({ row: 2, col: 0 })).toBeDefined()
    h.remote.ledger.charge(pressure, 1, 4096 - h.remote.ledger.bytes + 1)
    expect(h.store.cellAt({ row: touched === 'none' ? 1 : 2, col: 0 })).toBeUndefined()
    expect(h.store.cellAt({ row: touched === 'none' ? 2 : touched === 'same' ? 0 : 1, col: 0 })).toBeDefined()
  } finally { await h.close() }
})
