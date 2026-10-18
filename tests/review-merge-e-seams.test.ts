import { expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { createCellLayoutHost } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { handlePlacement, guidePlacement } from '../src/components/xlsx/sheetResize'
import { columnWidthPx } from '../src/xlsx/columnWidth'
import type { XlsxCell, XlsxSheet, XlsxDrawing } from '../src/xlsx/types'
import { createAdmittedLayoutContext } from '../src/components/xlsx/measureContext'
import { createLocalTextMeasurer, inheritedCanvasLang } from '../src/core/text-metrics'
import * as common from '../src/xlsx/formula/format'
const sheet: XlsxSheet = { name:'S',rows:[],cols:[],merges:[],mergeRanges:[] }
const owner={docId:'d',sheetId:'s',viewGen:0,valueGen:0,fontGen:0}
const gate={admit:()=>true,release:()=>{},hold:()=>{},cacheChanged:()=>{}}
it('admitted A/B/E formatting has canonical parity across selected and fallback sections',()=>{
 const formats=[undefined,'General','[Red]General','[>0]0.00" units";0','[>10]0','[>10]General;0','0.00"General"','0.00E+00','# ?/?','[h]:mm:ss','[Purple]0','0;0;0;0;0']
 const failures:any[]=[]
 for(const formatCode of formats)for(const numFmtId of [0,11,164])for(const value of [0.000000123456789,123456789012345,-12345.67])for(const width of [24,64,160]){
  const cell:XlsxCell={ref:'A1',row:0,col:0,styleIndex:0,value,style:{numFmtId,...(formatCode===undefined?{}:{formatCode})}}
  const host=createCellLayoutHost(createCanvas(1,1).getContext('2d') as never,sheet)
  const canonical=host.layoutBody(cell,width,20)
  const cache=new LayoutCache();cache.beginBatch(gate)
  try {const actual=cache.layout(owner,{cell,source:{row:0,col:0},x:0,y:0,width,height:20},host);if(actual.formattedText!==canonical.formattedText)failures.push({formatCode,numFmtId,value,width,actual:actual.formattedText,canonical:canonical.formattedText})}
  finally{cache.endBatch()}
 }
 console.log('format parity failures',JSON.stringify(failures));expect(failures).toEqual([])
})
it('conditional and quoted General fixed outputs do not repeat unobserved formatting',()=>{
 const spy=vi.spyOn(common,'formatCommon')
 try {
  const events:any[]=[]
  for(const code of ['[>0]0.00" units";0','0.00"General"']){
   spy.mockClear();const cell:XlsxCell={ref:'A1',row:0,col:0,styleIndex:0,value:3,style:{numFmtId:164,formatCode:code}}
   const host=createCellLayoutHost(createCanvas(1,1).getContext('2d') as never,sheet),cache=new LayoutCache();cache.beginBatch(gate)
   try {cache.layout(owner,{cell,source:{row:0,col:0},x:0,y:0,width:160,height:20},host)}finally{cache.endBatch()}
   events.push({code,calls:spy.mock.calls.map(c=>({code:c[1],observed:!!c[3]}))})
  }
  console.log('format observations',JSON.stringify(events));expect(events.every(e=>e.calls.every((c:any)=>c.observed))).toBe(true)
 }finally{spy.mockRestore()}
})
it('gutter-relative handles and guides preserve full targets at far-row gutter width',()=>{
 const layout={viewport:{width:300,height:200},headerHeight:132,gutter:61}
 for(const touch of [false,true])for(const edge of [0,5,150,299,300]){
  const b=handlePlacement('column',edge,{...layout,touch});expect(b.x).toBeGreaterThanOrEqual(61);expect(b.x+b.width).toBeLessThanOrEqual(361);expect(b.x).toBeLessThanOrEqual(61+edge);expect(b.x+b.width).toBeGreaterThanOrEqual(61+edge)
  expect(guidePlacement('column',edge,{...layout,touch}).x).toBe(61+edge-1)
 }
 expect(handlePlacement('row',100,{...layout,touch:true})).toEqual({x:0,y:44,width:61,height:44})
 expect(guidePlacement('row',100,{...layout,touch:false}).width).toBe(361)
})
it('MDW fixtures retain the original pixel widths',()=>{
 expect([8/7,75/7,1.7142857,57.7142857].map(w=>columnWidthPx(w,7))).toEqual([8,75,12,404])
})
it('stored row heights and non-64 default columns survive the resize drawing remap',()=>{
 const E=9525,drawing={kind:'shape',xEmu:88*E,yEmu:17*E,widthEmu:176*E,heightEmu:17*E,source:{treePath:'twoCellAnchor[0]/sp[0]'}} as unknown as XlsxDrawing
 const s:XlsxSheet={...sheet,columnMetrics:{maxDigitWidthPx:8,defaultWidthPx:88},rows:[{index:0,heightPt:12.75,cells:[]}],drawings:[drawing]}
 const g=buildSheetGeometry(s,{sizes:{rows:new Map(),cols:new Map([[1,120]])}});expect(g.kind).toBe('ready');if(g.kind!=='ready')return
 expect(g.metrics.rows.size(0)).toBe(17);expect(g.metrics.cols.size(0)).toBe(88);expect(g.metrics.cols.size(1)).toBe(120)
 const d=g.drawingsIn(0,0,1000,1000,100)![0];expect(d.xEmu).toBe(88*E);expect(d.widthEmu).toBe((120+88)*E)
})
it('both admitted native measurement contexts inherit the painting document language',()=>{
 const old=document.documentElement.lang;document.documentElement.lang='zh-TW'
 const nativeCreate=document.createElement.bind(document),observations:any[]=[]
 const make=vi.spyOn(document,'createElement').mockImplementation(((tag:string,...args:any[])=>{
  const el=nativeCreate(tag,...args as []);if(tag==='canvas'){
   const ctx=(el as HTMLCanvasElement).getContext('2d')! as CanvasRenderingContext2D&{lang:string};Object.defineProperty(ctx,'lang',{configurable:true,writable:true,value:'inherit'})
   const original=ctx.measureText.bind(ctx);ctx.measureText=(t:string)=>{observations.push({lang:ctx.lang,effective:ctx.lang==='inherit'?inheritedCanvasLang(ctx.canvas):ctx.lang,text:t});return original(t)}
  }return el
 }) as typeof document.createElement)
 try {
  const ctx=createAdmittedLayoutContext(()=>{})!;ctx.font='12px sans-serif';ctx.measureText('plain');createLocalTextMeasurer(ctx)('local')
  console.log('languages',JSON.stringify(observations));expect(observations.every(o=>o.effective==='zh-TW')).toBe(true)
 }finally{make.mockRestore();document.documentElement.lang=old}
})
