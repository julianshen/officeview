// @vitest-environment node
import { expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { cellDisplayText } from '../src/xlsx/render'
import { createCellLayoutHost } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import type { XlsxCell, XlsxSheet } from '../src/xlsx/types'
it('selected General, no-section, fallback, zero and style variants match canonical output across widths',()=>{
 const sheet:XlsxSheet={name:'S',rows:[],cols:[],merges:[],mergeRanges:[]},owner={docId:'d',sheetId:'s',viewGen:0,valueGen:0,fontGen:0},failures:unknown[]=[]
 const codes=[undefined,'General','[Red]General','[>0]General;[Blue]0.00','[>10]0','[>10]General;[<0]0.00','0.00"General"','[Purple]0','[Red]0.00','0.00E+00','[h]:mm:ss']
 let cases=0,missed=0
 for(const code of codes)for(const numFmtId of [0,11,164])for(const value of [0,123456789012345,-12345.6789,1.23456789e-7])for(const extra of [{},{shrinkToFit:true},{textRotation:45},{wrapText:true}]){
  const cell:XlsxCell={ref:'A1',row:0,col:0,styleIndex:0,value,style:{numFmtId,...(code===undefined?{}:{formatCode:code}),...extra}}
  let flagged=false;const text=cellDisplayText(sheet,cell,undefined,()=>{flagged=true})
  for(const width of [24,64,160]){
   const host=createCellLayoutHost(createCanvas(1,1).getContext('2d') as never,sheet),canonical=host.layoutBody(cell,width,20)
   if(!flagged&&host.layoutBody(cell,width,20,undefined,text).formattedText!==canonical.formattedText)missed++
   const cache=new LayoutCache();cache.beginBatch({admit:()=>true,release:()=>{},hold:()=>{},cacheChanged:()=>{}})
   try{const actual=cache.layout(owner,{cell,source:{row:0,col:0},x:0,y:0,width,height:20},host);if(actual.formattedText!==canonical.formattedText)failures.push({code,numFmtId,value,extra,width,actual:actual.formattedText,canonical:canonical.formattedText});cases++}finally{cache.endBatch()}
  }
 }
 console.log('width matrix',JSON.stringify({cases,missed,parityFailures:failures.length}));expect(missed).toBe(0);expect(failures).toEqual([])
})
