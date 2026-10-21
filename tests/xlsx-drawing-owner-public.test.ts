// @vitest-environment node
import { expect, it } from 'vitest'
import v8 from 'node:v8'
import vm from 'node:vm'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { ViewerClient } from '../src/worker/viewerClient'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { RemoteDrawingCache, drawingNotices } from '../src/worker/viewerRemoteDrawings'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { pagedSheetGeometry } from '../src/components/xlsx/pagedGeometry'
import { SheetImageCache } from '../src/components/xlsx/sheetAssets'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { remapDrawings, SIZED_DRAWING_BYTES } from '../src/xlsx/view/axisSizes'
import { MAIN_THREAD_BUDGETS } from '../src/xlsx/view/tileCache'
import type { SheetGeometry } from '../src/xlsx/view/viewport'
import type { XlsxDrawing } from '../src/xlsx/types'
import { workbookOf } from './helpers/xlsx-find'

v8.setFlagsFromString('--expose_gc')
const gc = vm.runInNewContext('gc') as () => void
async function collect() { for (let i = 0; i < 12; i++) { gc(); await new Promise(r => setTimeout(r, 10)) } }
const E = 9525, imageBytes = 1024 * 1024
async function harness(options: { empty?: boolean; image?: boolean; late?: boolean } = {}) {
  const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 42 }] }])
  doc.sheets[0].columnMetrics = { maxDigitWidthPx: 7, defaultWidthPx: 80 }
  doc.sheets[0].frozen = { rows: 1, cols: 1, topRow: 1, leftCol: 1 }
  if (!options.empty) doc.sheets[0].drawings = [{
    xEmu: 80 * E, yEmu: 0, widthEmu: 80 * E, heightEmu: 20 * E,
    source: { partPath: 'd', treePath: 'twoCellAnchor[0]/grpSp[0]', element: 'grpSp', representation: 'native' },
    group: { off: { x: 80 * E, y: 0 }, ext: { width: 80 * E, height: 20 * E }, chOff: { x: 0, y: 0 }, chExt: { width: 80 * E, height: 20 * E } },
    children: [{ xEmu: 0, yEmu: 0, widthEmu: 80 * E, heightEmu: 20 * E,
      source: { partPath: 'd', treePath: 'twoCellAnchor[0]/grpSp[0]/graphicFrame[0]', element: 'graphicFrame', representation: 'native' },
      // A transport-opaque payload (never painted): its byte size calibrates the main-pressure tests, so it keeps
      // its pre-D0 number array rather than the point objects (which would move it past the 64 KiB headroom).
      content: { kind: 'xlsxChart', chart: { groups: [{ series: [{ values: Array.from({ length: 1000 }, (_, i) => i + 42) }] }] } } as never,
      ...(options.image ? { image: { data: new Uint8Array(imageBytes), mime: 'image/png', partPath: 'p.png' } } : {}),
    }],
  }]
  doc.sheets[0].drawingAnchors = { unresolved: 0, records: options.empty ? [] : [{ id: 'd', kind: 'twoCell', from: { col: 1, row: 0, colOffEmu: 0, rowOffEmu: 0 }, to: { col: 2, row: 1, colOffEmu: 0, rowOffEmu: 0 } }] }
  const source = createObjectWorkbookSource(doc, 'r')
  let release = () => {}
  if (options.late) {
    const gate = new Promise<void>(resolve => { release = resolve })
    const store = source.store('s'), read = store.readDrawings!.bind(store)
    store.readDrawings = (...args) => gate.then(() => read(...args))
  }
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const client = new ViewerClient(connectInProcess(kernel))
  expect((await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'r', source: { kind: 'object', document: doc }, profile: 'D' })).kind).toBe('complete')
  const session = kernel.session('r')!
  const remote = new RemoteWorkbook(client, { docId: 'r', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map([['s', session.layout('s')]]) })
  remote.merges('s')
  for (let i = 0; i < 200 && remote.loadState('s').completeness !== 'complete'; i++) await new Promise(r => setTimeout(r, 2))
  expect(remote.loadState('s').completeness).toBe('complete')
  return { remote, release }
}
async function ready(remote: RemoteWorkbook) {
  for (let i = 0; i < 200 && remote.drawings('s').state !== 'ready'; i++) await new Promise(r => setTimeout(r, 2))
  expect(remote.drawings('s').state).toBe('ready')
}
const geometry = (remote: RemoteWorkbook) => pagedSheetGeometry(remote, 's', createIdentityOrder(), visibilityFromHidden(createIntervalSet([])), undefined, { cols: new Map([[0, 120]]), rows: new Map() })

it('drawing metadata pressure is a typed visible refusal within the existing main budget', async () => {
  const { remote } = await harness(), ledger = remote.mainLedger, cap = MAIN_THREAD_BUDGETS.D
  ledger.set('dom', 'pressure', cap - ledger.bytes - 64 * 1024)
  const cache = new RemoteDrawingCache({ identity: id => remote.identity(id), loadState: () => ({ completeness: 'complete' }), nextRequestId: () => remote.nextRequestId(), request: (c, signal) => remote.request(c, signal), changed: () => {} }, cap)
  const content = cache.get('s')
  for (let i = 0; i < 100 && content.state === 'loading'; i++) await new Promise(r => setTimeout(r, 2))
  expect(content.state).toBe('failed')
  expect(drawingNotices(content).join(' ')).toMatch(/more main-thread memory/)
  expect(ledger.bytes).toBeLessThanOrEqual(cap)
  expect(ledger.peakBytes).toBeLessThanOrEqual(cap)
  cache.close(); ledger.release('dom', 'pressure'); await remote.close()
})

it('an empty actual drawing reply stays ready without a new memory notice under the same pressure', async () => {
  const { remote } = await harness({ empty: true }), ledger = remote.mainLedger, cap = MAIN_THREAD_BUDGETS.D
  ledger.set('dom', 'pressure', cap - ledger.bytes - 64 * 1024)
  await ready(remote)
  expect(drawingNotices(remote.drawings('s'))).toEqual([])
  expect(ledger.bytes).toBeLessThanOrEqual(cap)
  ledger.release('dom', 'pressure'); await remote.close()
})

it('a real image-cache key retains its source bytes after source close and refunds after final cache release', async () => {
  const { remote } = await harness({ image: true }), ledger = remote.mainLedger, cache = new SheetImageCache(64)
  await ready(remote)
  const key = await (async () => {
    const asset = remote.drawings('s').images[0]
    expect(asset.data.byteLength).toBe(imageBytes)
    await cache.acquire(asset, asset, 4, async () => ({ width: 1, height: 1 }) as CanvasImageSource)
    return new WeakRef(asset)
  })()
  await remote.close(); await collect()
  expect(key.deref()!.data.byteLength).toBe(imageBytes)
  expect(ledger.snapshot().byPool.drawings).toBeGreaterThanOrEqual(imageBytes)
  cache.release(key.deref()!); await collect()
  expect(ledger.snapshot().byPool.drawings).toBe(0)
})

it('a detached copied group preserves its copy charge across repeated remaps without keeping the old copy array', async () => {
  const { remote } = await harness(), ledger = remote.mainLedger
  await ready(remote)
  let first: SheetGeometry | null = geometry(remote)
  const old = new WeakRef(first.sizedDrawings!)
  const axes = remote.geometry('s'), cols = createAxisMetrics({ logicalCount: 16384, defaultSizeCss: 80, overrides: [{ start: 0, end: 0, sizeCss: 160 }] })
  ledger.set('drawings', 'test-remap-work', SIZED_DRAWING_BYTES)
  let second: XlsxDrawing[] | null = remapDrawings(first.sizedDrawings!, { rows: axes.rows, cols: first.metrics.cols }, { rows: axes.rows, cols })
  ledger.adoptWhileReachable('drawings', 'test-remap-work', 'drawings', [second], SIZED_DRAWING_BYTES)
  expect(second[0].xEmu / E).toBe(160)
  first = null; await collect()
  expect(old.deref()).toBeUndefined()
  let group: object | null = second[0].group!
  second = null; await remote.close(); await collect()
  expect(group).toBeTruthy()
  expect(ledger.snapshot().byPool.drawings).toBeGreaterThanOrEqual(SIZED_DRAWING_BYTES)
  group = null; await collect()
  expect(ledger.snapshot().byPool.drawings).toBe(0)
})

it('a late lazy-placement query node retains its source data and copy after the query geometry and source close', async () => {
  const { remote, release } = await harness({ image: true, late: true }), ledger = remote.mainLedger
  let delayed: SheetGeometry | null = geometry(remote)
  expect(delayed.sizedDrawings).toBeUndefined()
  release(); await ready(remote)
  let node: XlsxDrawing | null = delayed.drawingsIn(0, 0, 4000, 4000, 20_000)![0]
  expect(node.xEmu / E).toBe(120)
  delayed = null; await remote.close(); await collect()
  expect(node.children![0].image!.data!.byteLength).toBe(imageBytes)
  expect(ledger.snapshot().byPool.drawings).toBeGreaterThanOrEqual(imageBytes + SIZED_DRAWING_BYTES)
  node = null; await collect()
  expect(ledger.snapshot().byPool.drawings).toBe(0)
})
