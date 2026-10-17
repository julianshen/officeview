import { expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import * as format from '../src/xlsx/formula/format'
import * as extended from '../src/xlsx/cellFormat'
import { createCellLayoutHost } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import { createFitMeasure } from '../src/components/xlsx/sheetResize'
import type { XlsxCell, XlsxSheet } from '../src/xlsx/types'
const sheet: XlsxSheet = { name:'S', rows:[], cols:[], merges:[], mergeRanges:[] }
const owner = { docId:'d',sheetId:'s',viewGen:0,valueGen:0,fontGen:0 }
function observe() {
 const color=vi.spyOn(extended,'formatColor'), ext=vi.spyOn(extended,'formatExtended'), common=vi.spyOn(format,'formatCommon')
 return { color, ext, common, events:()=>[...color.mock.calls.map(c=>({fn:'color',observed:!!c[2]})),...ext.mock.calls.map(c=>({fn:'extended',observed:!!c[4]})),...common.mock.calls.map(c=>({fn:'common',observed:!!c[3]}))], restore:()=>{color.mockRestore();ext.mockRestore();common.mockRestore()} }
}
it('width-independent coloured paint keeps the observer for every format builder',()=>{
 const obs=observe(), cache=new LayoutCache(), gate={admit:()=>true,release:()=>{},hold:()=>{},cacheChanged:()=>{}}
 const cell:XlsxCell={ref:'A1',row:0,col:0,styleIndex:0,value:3,style:{numFmtId:164,formatCode:'[Red]0.00'}}
 cache.beginBatch(gate)
 try {
  const body=cache.layout(owner,{cell,source:{row:0,col:0},x:0,y:0,width:400,height:20},createCellLayoutHost(createCanvas(1,1).getContext('2d') as never,sheet))
  expect(body.formattedText).toBe('3.00');console.log('paint builders',JSON.stringify(obs.events()))
  expect(obs.events().every(c=>c.observed)).toBe(true)
 }finally{cache.endBatch();obs.restore()}
})
it('wrapped auto-fit keeps its allocation observer for its canonical layout pass',()=>{
 const obs=observe(),fit=createFitMeasure({sheet,admit:()=>true})
 const cell:XlsxCell={ref:'A1',row:0,col:0,styleIndex:0,value:3,style:{numFmtId:164,formatCode:'[Red]0.00',wrapText:true}}
 try {
  expect(fit.layoutHeight!(cell,100)).toBeGreaterThan(0)
  console.log('fit builders',JSON.stringify(obs.events()))
  expect(obs.events().every(c=>c.observed)).toBe(true)
 }finally{obs.restore()}
})
