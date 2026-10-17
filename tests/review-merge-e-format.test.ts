// @vitest-environment node
import {expect,it,vi} from 'vitest'
import {createCanvas} from 'canvas'
import {createCellLayoutHost} from '../src/xlsx/view/renderViewport'
import {LayoutCache} from '../src/xlsx/view/textLayout'
import type {XlsxCell,XlsxSheet} from '../src/xlsx/types'
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
