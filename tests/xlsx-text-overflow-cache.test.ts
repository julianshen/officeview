import { describe, expect, it } from 'vitest'
import { pagedSheetGeometry } from '../src/components/xlsx/pagedGeometry'
import { ViewerClient } from '../src/worker/viewerClient'
import { connectInProcess, ViewerKernel } from '../src/worker/viewerKernel'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { cellTextOverflow, rowOverflowInput } from '../src/xlsx/view/viewport'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { workbookOf } from './helpers/xlsx-find'

async function fixture() {
  const doc = workbookOf([{ sheetId: 's', cells: [
    { row: 0, col: 0, value: 'early' },
    { row: 0, col: 500, value: 'later long text' },
    { row: 0, col: 501, value: 'BLOCK' },
    { row: 1, col: 0, value: 'x'.repeat(64) },
  ] }])
  const source = createObjectWorkbookSource(doc, 'd')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const client = new ViewerClient(connectInProcess(kernel))
  await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'd', source: { kind: 'object', document: doc }, profile: 'D' })
  const remote = new RemoteWorkbook(client, {
    docId: 'd', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), rowCacheBytes: 512,
    layouts: new Map(source.sheets.map(sheet => [sheet.sheetId, kernel.session('d')!.layout(sheet.sheetId)])),
  })
  const order = createIdentityOrder(), visibility = visibilityFromHidden(createIntervalSet([]))
  const fresh = () => pagedSheetGeometry(remote, 's', order, visibility)
  return { store: remote.store('s'), geometry: fresh(), fresh, close: async () => { await remote.close(); await client.dispose() } }
}

describe('review round 5: overflow memo follows the real remote row cache', () => {
  it('sees a blocker in a late column chunk without replacing geometry', async () => {
    const f = await fixture()
    try {
      await f.store.load([0], 0, 0)
      const early = rowOverflowInput(f.geometry, 0)
      expect(cellTextOverflow(f.geometry, { row: 0, col: 0 })!.right).toBeGreaterThan(0)
      await f.store.load([0], 500, 501)
      expect(f.store.isKnown({ row: 0, col: 501 })).toBe(true)
      expect(cellTextOverflow(f.fresh(), { row: 0, col: 500 })!.right).toBe(0)
      expect(cellTextOverflow(f.geometry, { row: 0, col: 500 })!.right).toBe(0)
      expect(rowOverflowInput(f.geometry, 0)).not.toBe(early)
    } finally { await f.close() }
  })

  it.each(['eviction', 'clear'] as const)('drops stale blockers after %s and reload', async operation => {
    const f = await fixture()
    try {
      await f.store.load([0], 0, 501)
      expect(cellTextOverflow(f.geometry, { row: 0, col: 500 })!.right).toBe(0)
      if (operation === 'eviction') await f.store.load([1], 0, 0)
      else f.store.clear()
      expect(f.store.cellsOf(0)).toBeUndefined()
      expect(cellTextOverflow(f.geometry, { row: 0, col: 500 })!.right).toBeGreaterThan(0)
      await f.store.load([0], 0, 0)
      expect(f.store.isKnown({ row: 0, col: 501 })).toBe(false)
      expect(cellTextOverflow(f.geometry, { row: 0, col: 500 })).toEqual(cellTextOverflow(f.fresh(), { row: 0, col: 500 }))
      await f.store.load([0], 500, 501)
      expect(cellTextOverflow(f.geometry, { row: 0, col: 500 })!.right).toBe(0)
    } finally { await f.close() }
  })

  it('reuses occupancy without reading cells again for a stable snapshot', async () => {
    const f = await fixture()
    try {
      await f.store.load([0], 0, 501)
      const cells = f.store.cellsOf(0)!
      let reads = 0
      const counted = new Proxy(cells, { get: (target, key, receiver) => {
        if (typeof key === 'string' && /^\d+$/.test(key)) reads++
        return Reflect.get(target, key, receiver)
      } })
      const geometry = { ...f.geometry, cellsOf: () => counted }
      const input = rowOverflowInput(geometry, 0)
      reads = 0
      for (let i = 0; i < 64; i++) {
        expect(rowOverflowInput(geometry, 0)).toBe(input)
        expect(cellTextOverflow(geometry, { row: 0, col: 500 })!.right).toBe(0)
      }
      expect(reads).toBe(0)
    } finally { await f.close() }
  })
})
