// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { ViewerClient } from '../src/worker/viewerClient'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { WorkerViewportPlans } from '../src/components/xlsx/workerViewport'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { axisBytes } from '../src/xlsx/view/mainLedger'
import type { SheetGeometry } from '../src/worker/viewerSession'
import { workbookOf } from './helpers/xlsx-find'

async function harness(empty = false, missingLayout = false) {
  const doc = workbookOf([{ sheetId: 's', cells: empty ? [] : [{ row: 0, col: 0, value: 'abcdefghijklmnopqrstuvwxyz' }, { row: 0, col: 2, value: 'blocker' }] }])
  const objectSource = createObjectWorkbookSource(doc, 'columns')
  let geometry = objectSource.geometry('s')
  const source = { ...objectSource, geometry: () => geometry }
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const port = connectInProcess(kernel), client = new ViewerClient(port)
  await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'columns', source: { kind: 'object', document: doc }, profile: 'D' })
  const session = kernel.session('columns')!, owner = session.identity('s')
  const remote = new RemoteWorkbook(client, { docId: 'columns', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: missingLayout ? new Map() : new Map([['s', session.layout('s')]]) })
  const plans = new WorkerViewportPlans(remote)
  const flush = async () => { await port.idle(); for (let i = 0; i < 50; i++) await Promise.resolve() }
  remote.merges('s'); await flush()
  const window = { box: { x: 0, y: 0, width: 384, height: 20 }, zoom: 1, deviceScale: 1 }, extent = { width: 384, height: 20 }
  const plan = async () => { plans.beginFrame(owner, window); plans.get(owner, window, extent, true); await flush(); return plans.get(owner, window, extent, true) }
  return { remote, plans, owner, window, extent, flush, plan, finalGeometry: (next: SheetGeometry) => { geometry = next; remote.publishLayout('s', session.layout('s')) }, close: async () => { plans.dispose(); await remote.release() } }
}

it('SLICE10 reuses already-charged remote geometry columns during worker proof validation', async () => {
  const h = await harness()
  try {
    h.remote.geometry('s')
    const before = h.remote.mainLedger.snapshot().byPool.geometry
    expect((await h.plan()).kind).toBe('detail')
    expect(h.remote.mainLedger.snapshot().byPool.geometry).toBe(before)
  } finally { await h.close() }
})

it('SLICE10 a missing layout has stable identity', async () => {
  const h = await harness(true, true)
  try { expect(h.remote.layout('s')).toBe(h.remote.layout('s')) } finally { await h.close() }
})

it('SLICE10 repeated missing-layout replies retain exactly one column axis charge', async () => {
  const h = await harness(true, true)
  try {
    expect((await h.plan()).kind).toBe('detail')
    const before = h.remote.mainLedger.snapshot().byPool.geometry
    expect(before).toBe(axisBytes(createAxisMetrics({ logicalCount: 16384, defaultSizeCss: 64 })))
    h.plans.release(h.owner, h.window)
    expect((await h.plan()).kind).toBe('detail')
    expect(h.remote.mainLedger.snapshot().byPool.geometry).toBe(before)
  } finally { await h.close() }
})

it('SLICE10 final import metadata refreshes proof columns and accepts new column starts', async () => {
  const h = await harness()
  try {
    const initial = await h.plan()
    expect(initial.kind).toBe('detail')
    if (initial.kind !== 'detail') throw new Error('expected initial detail')
    expect(initial.cells.find(c => c.source.col === 2)?.x).toBe(128)
    const base = h.remote.geometry('s'), colOverrides = [{ start: 0, end: 0, sizeCss: 192 }]
    h.finalGeometry({ ...base, colOverrides, cols: createAxisMetrics({ logicalCount: 16384, defaultSizeCss: 64, overrides: colOverrides }) })
    const final = await h.plan()
    expect(final.kind).toBe('detail')
    if (final.kind !== 'detail') throw new Error('expected final detail')
    expect(final.cells.find(c => c.source.col === 2)?.x).toBe(256)
    expect(final.cells.find(c => c.source.col === 0)?.overflow?.right).toBe(64)
  } finally { await h.close() }
})

it('SLICE10 column cache construction waits for joint admission without creating row geometry', async () => {
  const h = await harness(), charge = vi.spyOn(h.remote.mainLedger, 'chargeWhileReachable'), geometry = vi.spyOn(h.remote, 'geometry').mockImplementation(() => { throw new Error('row geometry must stay lazy') })
  try {
    h.remote.mainLedger.set('copy', 'pressure', 64 * 1024 * 1024 - h.remote.mainLedger.bytes)
    expect((await h.plan()).kind).toBe('diagnostic')
    expect(charge.mock.calls.filter(([pool]) => pool === 'geometry')).toHaveLength(0)
    expect(geometry).not.toHaveBeenCalled()
  } finally { h.remote.mainLedger.release('copy', 'pressure'); charge.mockRestore(); geometry.mockRestore(); await h.close() }
})

it('SLICE10 geometry constructed after worker proof adopts the same already-charged column axis', async () => {
  const h = await harness()
  try {
    expect((await h.plan()).kind).toBe('detail')
    const before = h.remote.mainLedger.snapshot().byPool.geometry
    const geometry = h.remote.geometry('s')
    expect(geometry.cols).toBe(h.remote.columns('s'))
    expect(h.remote.mainLedger.snapshot().byPool.geometry).toBe(before + axisBytes(geometry.rows))
    const oldColumns = geometry.cols
    h.finalGeometry({ ...geometry, colOverrides: [{ start: 0, end: 0, sizeCss: 192 }], cols: createAxisMetrics({ logicalCount: 16384, defaultSizeCss: 64, overrides: [{ start: 0, end: 0, sizeCss: 192 }] }) })
    expect(h.remote.columns('s')).not.toBe(oldColumns)
    expect(oldColumns.start(1)).toBe(64)
    expect(h.remote.columns('s').start(1)).toBe(192)
    // A consumer still holds the old axis: replacing metadata must not
    // settle its charge early. Collection remains the ledger's release path.
    expect(h.remote.mainLedger.snapshot().byPool.geometry).toBe(before + axisBytes(geometry.rows) + axisBytes(h.remote.columns('s')))
  } finally { await h.close() }
})
