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

it('peer probe: actual large image reply under existing aggregate pressure stays inside the main budget', async () => {
 const {remote}=await harness({image:true}),ledger=remote.mainLedger,cap=MAIN_THREAD_BUDGETS.D
 ledger.set('dom','peer-raw-pressure',cap-ledger.bytes-64*1024)
 const cache=new RemoteDrawingCache({identity:id=>remote.identity(id),loadState:()=>({completeness:'complete'}),nextRequestId:()=>remote.nextRequestId(),request:(c,signal)=>remote.request(c,signal),changed:()=>{}},cap)
 try{
  const content=cache.get('s')
  for(let i=0;i<100&&content.state==='loading';i++)await new Promise(r=>setTimeout(r,2))
  console.log(JSON.stringify({state:content.state,mainBytes:ledger.bytes,mainPeak:ledger.peakBytes,cap,overBy:ledger.peakBytes-cap}))
  expect(ledger.peakBytes).toBeLessThanOrEqual(cap)
 }finally{cache.close();ledger.release('dom','peer-raw-pressure');await remote.close()}
})
