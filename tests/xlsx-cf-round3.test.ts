// @vitest-environment node
import { expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { prepareCfPaint } from '../src/xlsx/cf/render'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { computeMetrics, renderSheet } from '../src/xlsx/render'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import type { XlsxSheet } from '../src/xlsx/types'
function sheet(count = 5, end = 1048575): XlsxSheet {
  return {name:'CF',sheetId:'1',cols:[],merges:[],mergeRanges:[],rows:[{index:0,cells:[{row:0,col:1,ref:'B1',value:1,styleIndex:0}]},{index:99,cells:[{row:99,col:1,ref:'B100',value:1,styleIndex:0}]}],conditionalFormatting:{diagnostics:[],rules:Array.from({length:count},(_,id)=>({id,priority:id+1,documentOrder:id,stopIfTrue:false,ranges:[{firstRow:0,lastRow:end,firstCol:1,lastCol:1}],writes:new Set(['fill']),dxf:{'fill.bg':'FF00FF00'},definition:{kind:'cellIs',operator:'greaterThan',operands:[{kind:'number',value:5}],dialect:'base',formulas:[]}}))}}
}
it('coalesces overlapping rules and panes into one source-space slab',async()=>{
  const s=sheet(),g=buildSheetGeometry(s);if(g.kind!=='ready')throw Error('geometry')
  const b=await prepareCfPaint(s,{viewport:{metrics:g.metrics,boxes:[{x:0,y:0,width:4608,height:6144},{x:0,y:0,width:128,height:100}]}})
  try {expect(b).toBeDefined();expect(b?.layer.rects).toEqual([{minRow:0,maxRow:307,minCol:1,maxCol:1}]);expect(s.diagnostics?.some(d=>d.feature==='cf-paint-unavailable')).not.toBe(true)}finally{b?.dispose()}
})
it('scopes large finite ranges to the tile ring instead of preparing 99990 cells',async()=>{
  const s=sheet(3,9999);s.rows=[{index:1,cells:[{row:1,col:0,ref:'A2',value:10,styleIndex:0}]},{index:9999,cells:[{row:9999,col:9,ref:'J10000',value:10,styleIndex:0}]}]
  s.conditionalFormatting={...s.conditionalFormatting!,rules:s.conditionalFormatting!.rules.map((r,i)=>({...r,ranges:[{firstRow:1,lastRow:9999,firstCol:0,lastCol:9}],definition:{kind:'cellIs',operator:'greaterThan',operands:[{kind:'number',value:i===0?5:20}],dialect:'base',formulas:[]}}))}
  const g=buildSheetGeometry(s);if(g.kind!=='ready')throw Error('geometry');const b=await prepareCfPaint(s,{viewport:{metrics:g.metrics,boxes:[{x:0,y:0,width:1024,height:1024}]}})
  try{expect(b?.layer.styleAt(1,0,undefined)?.fillColor).toBe('FF00FF00');expect(b?.layer.rects).toEqual([{minRow:1,maxRow:51,minCol:0,maxCol:9}])}finally{b?.dispose()}
})
it('blank CF borders follow populated borders in static and viewport paint',async()=>{
  const s=sheet(1,1);s.rows=[{index:0,cells:[{row:0,col:0,ref:'A1',value:10,styleIndex:0,style:{numFmtId:0,borders:{bottom:'thin'},borderColors:{bottom:'#000000'}}}]},{index:1,cells:[]}];s.conditionalFormatting={...s.conditionalFormatting!,rules:s.conditionalFormatting!.rules.map(r=>({...r,ranges:[{firstRow:1,lastRow:1,firstCol:0,lastCol:0}],writes:new Set(['border.top']),dxf:{'border.top.style':'thin','border.top.color':'FFFF0000'},definition:{kind:'predicate',type:'containsBlanks',dialect:'base',formulas:[]}}))}
  const b=await prepareCfPaint(s),m=computeMetrics(s),a=createCanvas(128,80).getContext('2d'),v=createCanvas(128,80).getContext('2d');try{renderSheet(s,a as unknown as CanvasRenderingContext2D,m,undefined,{conditionalFormatting:b});renderSheetViewport(s,v as unknown as CanvasRenderingContext2D,buildSheetGeometry(s),{box:{x:0,y:0,width:128,height:80},zoom:1,deviceScale:1},{conditionalFormatting:b});expect([...v.getImageData(10,20,1,1).data]).toEqual([255,0,0,255]);expect([...a.getImageData(10,20,1,1).data]).toEqual([...v.getImageData(10,20,1,1).data])}finally{b?.dispose()}
})
