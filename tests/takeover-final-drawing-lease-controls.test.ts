// @vitest-environment node
import { expect, it } from 'vitest'
import v8 from 'node:v8'
import vm from 'node:vm'
import { chartWorkbookBytes } from './helpers/xlsx-chart-fixture'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { createViewerClient } from '../src/worker/viewerClient'
import { RemoteWorkbook, openRemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { pagedSheetGeometry } from '../src/components/xlsx/pagedGeometry'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { visibilityFromHidden, createIntervalSet } from '../src/xlsx/view/visibility'
import { sheetWithDrawings } from '../src/worker/viewerRemoteDrawings'

v8.setFlagsFromString('--expose_gc')
const gc = vm.runInNewContext('gc') as () => void
async function collect() { for (let i=0;i<12;i++) { gc(); await new Promise(r=>setTimeout(r,10)) } }
const chart = '<chartSpace><chart><plotArea><barChart><ser><idx val="0"/><val><numRef><f>Styled!A1:A1000</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>'
async function open() {
 const bytes=await chartWorkbookBytes(chart,{cells:Array.from({length:1000},(_,i)=>({ref:`A${i+1}`,v:i+42})),patch:async zip=>{
  const p='xl/worksheets/sheet1.xml';const s=await zip.file(p)!.async('string')
  zip.file(p,s.replace('<sheetData>','<sheetFormatPr baseColWidth="10"/><sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" state="frozen"/></sheetView></sheetViews><sheetData>'))
 }})
 const blob=new Blob([Uint8Array.from(bytes)]),session=await openImportSession({kind:'blob',blob},{docId:'P',profile:'D'})
 expect((await session.done).kind).toBe('complete')
 const source=createPagedWorkbookSource(session)
 const kernel=new ViewerKernel({parse:async()=>{throw Error('must use paged source')},openSource:async()=>source},{yield:async()=>{}})
 const client=createViewerClient(connectInProcess(kernel)),opened=await openRemoteWorkbook(client,'P',blob,'D')
 if(!('layouts'in opened)) throw Error('open failed')
 const remote=new RemoteWorkbook(client,opened)
 remote.merges('1')
 for(let i=0;i<200&&remote.loadState('1').completeness!=='complete';i++)await new Promise(r=>setTimeout(r,2))
 expect(remote.loadState('1').completeness).toBe('complete')
 for(let i=0;i<200&&remote.drawings('1').state!=='ready';i++)await new Promise(r=>setTimeout(r,2))
 expect(remote.drawings('1').state).toBe('ready')
 return remote
}


it('control: an original sheet drawing-array owner keeps the payload charged through source close',async()=>{
 const remote=await open(),ledger=remote.mainLedger
 let sheet=sheetWithDrawings({name:'Styled',rows:[],cols:[],merges:[],mergeRanges:[]},remote.drawings('1'))
 const charged=ledger.snapshot().byPool.drawings
 expect(charged).toBeGreaterThanOrEqual(8000)
 await remote.close();await collect()
 expect((sheet.drawings![0].content as any).chart.groups[0].series[0].values.values[999]).toBe(1041)
 expect(ledger.snapshot().byPool.drawings).toBe(charged)
 sheet=null as any;await collect()
 expect(ledger.snapshot().byPool.drawings).toBe(0)
},15000)
it('control: source close releases unescaped chart content after real collection',async()=>{
 const remote=await open(),ledger=remote.mainLedger
 expect(ledger.snapshot().byPool.drawings).toBeGreaterThanOrEqual(8000)
 await remote.close();await collect()
 expect(ledger.snapshot().byPool.drawings).toBe(0)
},15000)
it('a raw paged geometry query result also retains its reachable chart data through source close',async()=>{
 const remote=await open(),ledger=remote.mainLedger
 let nodes=(()=>{const geometry=pagedSheetGeometry(remote,'1',createIdentityOrder(),visibilityFromHidden(createIntervalSet([])));return geometry.drawingsIn(0,0,4000,4000,20000)!})()
 expect(nodes).toHaveLength(1)
 await remote.close();await collect()
 expect((nodes[0].content as any).chart.groups[0].series[0].values.values[999]).toBe(1041)
 console.log('retained unsized query chart after source close',JSON.stringify({points:1000,minimumDataBytes:8000,drawingsCharged:ledger.snapshot().byPool.drawings}))
 expect(ledger.snapshot().byPool.drawings).toBeGreaterThanOrEqual(8000)
 nodes=null as any;await collect()
 expect(ledger.snapshot().byPool.drawings).toBe(0)
},15000)
