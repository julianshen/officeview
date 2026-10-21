// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { drawingFootprint } from '../src/xlsx/view/drawingPayload'
import { fillChartRanges, parseXlsxChart } from '../src/charts'
import { parseXmlOrdered } from '../src/core/xml'
import { WorkerViewportPlans } from '../src/components/xlsx/workerViewport'
import { paneLogicalBoxes } from '../src/xlsx/view/frozenPanes'

it('numeric chart arrays are accounted by drawing footprint', () => {
  const model = parseXlsxChart(parseXmlOrdered('<chartSpace><chart><plotArea><lineChart><ser><val><numRef><f>A1:A4096</f></numRef></val></ser></lineChart></plotArea></chart></chartSpace>'), undefined)!
  const full = fillChartRanges(model, 'Data', (_s,r) => ({value:r}))
  const small = {...full, groups:full.groups.map(g=>({...g,series:g.series.map(s=>({...s,values:[0]}))}))}
  const extra = drawingFootprint(full)-drawingFootprint(small)
  console.log({full:drawingFootprint(full),small:drawingFootprint(small),extra,minimumExtra:4095*8})
  expect(extra).toBeGreaterThanOrEqual(4095*8)
})

const owner = {docId:'d',sheetId:'s',valueGen:0,fontGen:0,viewGen:0}
const windowOf = (box: {x:number;y:number;width:number;height:number}) => ({box,zoom:1,deviceScale:1})
function harness() {
  const signals:AbortSignal[]=[]
  const remote = {
    beginPageFrame:vi.fn(), merges:()=>({gen:0}), pageLimits:{tile:8,frame:32}, currentPageFrame:0,
    admitPages:()=>true, releasePages:vi.fn(), store:()=>({captureViewportRequest:()=>({live:()=>true,close:vi.fn()})}),
    identity:()=>owner,viewport:(_o:unknown,_w:unknown,signal:AbortSignal)=>{signals.push(signal);return new Promise(resolve=>signal.addEventListener('abort',()=>resolve({kind:'cancelled'})))},
    mainLedger:{release:vi.fn(),set:vi.fn()}, notify:vi.fn(), profile:'D', isClosed:false,
  }
  return {remote,signals,plans:new WorkerViewportPlans(remote as never)}
}
// The viewer hands beginFrame the visible regions (paneLogicalBoxes), as InteractiveSheetDoc does.
const frozen = { rows: 1, cols: 1, left: 64, top: 20 }, pane = { width: 640, height: 400 }
const frameAt = (h: ReturnType<typeof harness>, x: number, y: number) => h.plans.beginFrame(owner, windowOf({ x, y, width: 640, height: 400 }), paneLogicalBoxes(frozen, { x, y }, 1, pane))
it('frozen scrolling cancels pending plans in the invisible gap', () => {
  const h=harness(), old=windowOf({x:5120,y:5120,width:512,height:512})
  frameAt(h, 5120, 5120)
  h.plans.get(owner,old,{width:1e6,height:1e6},true)
  frameAt(h, 51200, 51200)
  const aborted=h.signals[0].aborted
  h.plans.dispose()
  expect(aborted).toBe(true)
})
it('frozen scrolling keeps the pinned corner and strip requests alive', () => {
  const h=harness(), corner=windowOf({x:0,y:0,width:512,height:512})
  frameAt(h, 5120, 5120)
  h.plans.get(owner,corner,{width:1e6,height:1e6},true)
  frameAt(h, 51200, 51200)
  const aborted=h.signals[0].aborted
  h.plans.dispose()
  expect(aborted).toBe(false)
})
it('vertical-only scrolling keeps top-strip requests; horizontal-only keeps left-strip requests', () => {
  const top=harness(), topRequest=windowOf({x:5120,y:0,width:512,height:512})
  frameAt(top, 5120, 5120)
  top.plans.get(owner,topRequest,{width:1e6,height:1e6},true)
  frameAt(top, 5120, 51200)
  const topAborted=top.signals[0].aborted
  top.plans.dispose()
  const left=harness(), leftRequest=windowOf({x:0,y:5120,width:512,height:512})
  frameAt(left, 5120, 5120)
  left.plans.get(owner,leftRequest,{width:1e6,height:1e6},true)
  frameAt(left, 51200, 5120)
  const leftAborted=left.signals[0].aborted
  left.plans.dispose()
  expect({topAborted,leftAborted}).toEqual({topAborted:false,leftAborted:false})
  // and the same strip requests are dropped once the strip itself scrolls past them
  const gone=harness()
  frameAt(gone, 5120, 5120)
  gone.plans.get(owner,topRequest,{width:1e6,height:1e6},true)
  frameAt(gone, 51200, 51200)
  expect(gone.signals[0].aborted).toBe(true)
  gone.plans.dispose()
})
it('control: ordinary scrolling cancels the same obsolete pending plan', () => {
  const h=harness(), old=windowOf({x:5120,y:5120,width:512,height:512})
  h.plans.beginFrame(owner,windowOf({x:5120,y:5120,width:640,height:400}))
  h.plans.get(owner,old,{width:1e6,height:1e6},true)
  h.plans.beginFrame(owner,windowOf({x:51200,y:51200,width:640,height:400}))
  expect(h.signals[0].aborted).toBe(true)
  h.plans.dispose()
})

import { parseXlsx } from '../src/xlsx/parse'
import { chartWorkbook, addNumberFormat } from './helpers/xlsx-chart-fixture'
import { cellDisplayText } from '../src/xlsx/render'
it('1904 chart category labels equal the cell display', async () => {
 const xml='<chartSpace><chart><plotArea><barChart><ser><cat><numRef><f>Styled!A1</f></numRef></cat><val><numRef><f>Styled!B1</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>'
 const pkg=await chartWorkbook(xml,{cells:[{ref:'A1',v:1,style:3},{ref:'B1',v:5}],patch:async zip=>{
   await addNumberFormat(zip,14)
   const p='xl/workbook.xml',s=await zip.file(p)!.async('string')
   zip.file(p,s.replace('<sheets>','<workbookPr date1904="1"/><sheets>'))
 }})
 const doc=await parseXlsx(pkg),sheet=doc.sheets[0],cell=sheet.rows[0].cells.find(c=>c.col===0)!
 const content=sheet.drawings![0].content!
 expect(content.kind).toBe('xlsxChart')
 if(content.kind!=='xlsxChart')throw new Error('no chart')
 console.log({grid:cellDisplayText(sheet,cell),chart:content.chart.categories})
 expect(content.chart.categories).toEqual([cellDisplayText(sheet,cell)])
})
it('cached data beyond the 4096 point cap is announced as truncated', () => {
 const model=parseXlsxChart(parseXmlOrdered('<chartSpace><chart><plotArea><lineChart><ser><val><numLit><ptCount val="5000"/><pt idx="0"><v>1</v></pt><pt idx="4999"><v>2</v></pt></numLit></val></ser></lineChart></plotArea></chart></chartSpace>'),undefined)!
 expect(model.groups[0].series[0].values).toHaveLength(4096)
 expect(model.truncated).toBe(true)
})
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
it('frozen gap requests cannot monopolize the real frame page allowance', async () => {
 const h=harness(),remote:any=h.remote
 remote.pagesInFlight=0
 remote.beginPageFrame=RemoteWorkbook.prototype.beginPageFrame
 remote.admitPages=RemoteWorkbook.prototype.admitPages
 remote.releasePages=RemoteWorkbook.prototype.releasePages
 Object.defineProperty(remote,'currentPageFrame',{get:()=>remote.pageFrame})
 frameAt(h, 5430, 5430)
 for(const [x,y] of [[5120,5120],[5632,5120],[5120,5632],[5632,5632]])h.plans.get(owner,windowOf({x,y,width:512,height:512}),{width:1e6,height:1e6},true)
 expect(remote.pagesInFlight).toBe(32)
 frameAt(h, 51200, 51200)
 // Aborted requests settle (as real worker requests do) and give their page credits back.
 await new Promise(resolve => setTimeout(resolve, 0))
 h.plans.get(owner,windowOf({x:51200,y:51200,width:512,height:512}),{width:1e6,height:1e6},true)
 const calls=h.signals.length
 h.plans.dispose()
 expect(calls).toBe(5)
})
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
it('drawing reads refuse numeric chart storage above their admitted bytes', async () => {
 const xml='<chartSpace><chart><plotArea><barChart><ser><val><numRef><f>Styled!A1:A4096</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>'
 const doc=await parseXlsx(await chartWorkbook(xml,{cells:Array.from({length:4096},(_,i)=>({ref:`A${i+1}`,v:i}))}))
 const source=createObjectWorkbookSource(doc,'review'),store=source.store(doc.sheets[0].sheetId!)
 const charged:number[]=[]
 const out=await store.readDrawings!({stamp:store.stamp,signal:new AbortController().signal},'D',{set:bytes=>{charged.push(bytes);return bytes<=8192}})
 console.log({charged,completeness:out.completeness,pointCount:out.drawings[0]?.content?.kind==='xlsxChart'?out.drawings[0].content.chart.groups[0].series[0].values.length:0})
 await source.release()
 expect(out.completeness).toBe('incomplete')
})
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createScrollWindow, setScrollOffset, setScrollZoom, viewportToLogical } from '../src/xlsx/view/scrollWindow'
it('zoom dropping a frozen strip preserves an attainable pointer anchor', () => {
 const metrics=createSheetMetrics({rows:createAxisMetrics({logicalCount:1000,defaultSizeCss:20}),cols:createAxisMetrics({logicalCount:100,defaultSizeCss:64})})
 const size={width:200,height:200},anchor={x:10,y:50}
 const before=setScrollOffset(createScrollWindow(metrics,size,{frozen:{rows:0,cols:1}}),metrics,size,{x:640,y:400})
 const world=viewportToLogical(before,anchor)
 const after=setScrollZoom(before,metrics,size,2,anchor)
 expect(after.frozen?.left).toBe(0)
 console.log({beforeWorld:world,afterWorld:viewportToLogical(after,anchor),afterOrigin:after.logicalOrigin})
 expect(viewportToLogical(after,anchor)).toEqual(world)
})
