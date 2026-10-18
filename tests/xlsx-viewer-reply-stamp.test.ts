// @vitest-environment node
/** Codex D4 (spec §10.2): main charges a non-granted reply by the size the
 * worker measured in steps and stamped on it; it never walks the reply. */
import { describe, expect, it, vi } from 'vitest'
import * as payload from '../src/xlsx/view/payloadBytes'
import { ViewerClient } from '../src/worker/viewerClient'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { measurePayloadBytesStepped, measureWithinSlice } from '../src/worker/replyMeasure'
import { CheckpointBudget } from '../src/xlsx/view/scheduler'
import { workbookOf } from './helpers/xlsx-find'
import { MAIN_THREAD_BUDGETS } from '../src/xlsx/view/tileCache'

vi.mock('../src/xlsx/view/payloadBytes', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/view/payloadBytes')>()
  return { ...actual, estimatePayloadBytes: vi.fn(actual.estimatePayloadBytes), measurePayloadBytes: vi.fn(actual.measurePayloadBytes) }
})
const estimate = vi.mocked(payload.estimatePayloadBytes), measure = vi.mocked(payload.measurePayloadBytes)

const cells = (count: number) => ({ items: Array.from({ length: count }, (_, i) => ({ row: i, col: 0, value: `v${i}`, styleIndex: 0, nested: [i, { flag: true }] })), stamp: { docId: 'd', sheetId: 's', valueGen: 0 } })

describe('worker reply measure', () => {
  it('measures exactly what measurePayloadBytes does, in steps of at most 4,096 visits', async () => {
    const value = cells(8192)
    let visits = 0, since = 0
    const slices: number[] = []
    const budget = new CheckpointBudget(() => {}, undefined, { now: () => 0, yield: async () => { slices.push(visits - since); since = visits } })
    const original = budget.spend.bind(budget)
    budget.spend = units => { visits += units; return original(units) }
    expect(await measurePayloadBytesStepped(value, () => true, budget)).toBe(payload.measurePayloadBytes(value))
    slices.push(visits - since)
    expect(slices.length - 1).toBeGreaterThanOrEqual(2)
    expect(Math.max(...slices)).toBeLessThanOrEqual(4096)
    expect(measureWithinSlice(value, () => true)).toBeNull()
    expect(measureWithinSlice(cells(10), () => true)).toEqual({ bytes: payload.measurePayloadBytes(cells(10)) })
  })
})

describe('main charges a stamped reply without walking it', () => {
  it('an 8,192-cell readRows reply: no main walk, and the charge covers the reply', async () => {
    const doc = workbookOf([{ sheetId: 's', cells: Array.from({ length: 8192 }, (_, i) => ({ row: Math.floor(i / 64), col: i % 64, value: `cell-${i}` })) }])
    const source = createObjectWorkbookSource(doc, 'stamp')
    const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source })
    const port = connectInProcess(kernel), client = new ViewerClient(port)
    await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'stamp', source: { kind: 'object', document: doc }, profile: 'D' })
    const session = kernel.session('stamp')!
    const remote = new RemoteWorkbook(client, { docId: 'stamp', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map([['s', session.layout('s')]]) })
    try {
      const store = remote.store('s')
      estimate.mockClear(); measure.mockClear()
      const page = await store.readRows({ r0: 0, r1: 127, c0: 0, c1: 63 }, undefined, { signal: new AbortController().signal, stamp: { ...store.stamp } })
      expect(page.error).toBeUndefined()
      expect(page.items.reduce((sum, item) => sum + item.cells.length, 0)).toBe(8192)
      // Nothing on main (or in the in-process worker's posting path) walked the reply synchronously.
      expect(estimate).not.toHaveBeenCalled()
      expect(measure).not.toHaveBeenCalled()
      const charged = remote.mainLedger.snapshot().byPool.replies
      expect(charged).toBeGreaterThanOrEqual(payload.measurePayloadBytes(page)!)
      // The worker's stamp, not the conservative unstamped maximum (the whole main budget).
      expect(charged).toBeLessThan(MAIN_THREAD_BUDGETS.D)
    } finally { await remote.release() }
  })
})
