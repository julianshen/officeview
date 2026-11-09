// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { prepareCfPaint } from '../src/xlsx/cf/render'
import { computeMetrics, renderSheet } from '../src/xlsx/render'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import type { XlsxSheet } from '../src/xlsx/types'
function sheet(): XlsxSheet & { conditionalFormatting: { diagnostics: []; rules: import('../src/xlsx/cf/model').CfRule[] } } {
  return { name: 'CF', sheetId: '1', merges: [], mergeRanges: [], cols: [], rows: [{ index: 0, cells: [{ row: 0, col: 0, ref: 'A1', value: 10, styleIndex: 0 }] }], conditionalFormatting: { diagnostics: [], rules: [{ id: 0, priority: 1, documentOrder: 0, stopIfTrue: false, ranges: [{ firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }], writes: new Set(['fill']), dxf: { 'fill.bg': 'FF00FF00' }, definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 5 }], dialect: 'base', formulas: [] } }] } }
}
it('preserves overflow and watermark work across former 8-column seams', async () => {
 const s=sheet(); s.rows[0].cells=[{row:0,col:7,ref:'H1',value:'Overflowing text across tile seam into I1 and J1',styleIndex:0}]; s.rows.push({index:19,cells:[{row:19,col:14,ref:'O20',value:1,styleIndex:0}]})
 const b=await prepareCfPaint(s); expect(b).toBeDefined(); const m=computeMetrics(s)
 const paint=(cf:boolean)=>{const c=createCanvas(m.widthPx,m.heightPx),ctx=c.getContext('2d'),spy=vi.spyOn(ctx,'fillText'); renderSheet(s,ctx as unknown as CanvasRenderingContext2D,m,{text:'CONFIDENTIAL'},cf?{conditionalFormatting:b}:undefined); return {ctx,calls:spy.mock.calls.length}}
 try {const base=paint(false),cf=paint(true); expect(cf.calls).toBe(base.calls); expect([...cf.ctx.getImageData(516,2,56,16).data]).toEqual([...base.ctx.getImageData(516,2,56,16).data])} finally {b?.dispose()}
})
it('retains both halves of a thick CF border at a former row seam', async()=>{
 const s=sheet(); s.rows.push({index:19,cells:[{row:19,col:14,ref:'O20',value:1,styleIndex:0}]}); const r=s.conditionalFormatting.rules[0]; s.conditionalFormatting.rules[0]={...r,ranges:[{firstRow:7,lastRow:7,firstCol:0,lastCol:0}],writes:new Set(['border.bottom']),dxf:{'border.bottom.style':'thick','border.bottom.color':'FFFF0000'},definition:{kind:'predicate',type:'containsBlanks',dialect:'base',formulas:[]}} as typeof r
 const b=await prepareCfPaint(s),m=computeMetrics(s),ctx=createCanvas(m.widthPx,m.heightPx).getContext('2d'); try {renderSheet(s,ctx as unknown as CanvasRenderingContext2D,m,undefined,{conditionalFormatting:b}); expect([...ctx.getImageData(10,161,1,1).data]).toEqual([255,0,0,255])}finally{b?.dispose()}
})
it('refuses aggregate rule work across paint batches per-range rather than whole sheet (I4)', async () => {
 const s=sheet(),r=s.conditionalFormatting.rules[0]; s.conditionalFormatting.rules=Array.from({length:600},(_,id)=>({...r,id,priority:id+1,ranges:[{firstRow:0,lastRow:499,firstCol:0,lastCol:9}]})); const b=await prepareCfPaint(s); try{expect(b).toBeDefined();expect(b?.layer.styleAt(0,0,undefined)?.fillColor).toBe('FF00FF00')}finally{b?.dispose()}
})
it('prepares finite CF inside a 1080p half-zoom viewport and its tile remainder',async()=>{
 const s=sheet(),r=s.conditionalFormatting.rules[0]; s.rows.push({index:99,cells:[{row:99,col:51,ref:'AZ100',value:1,styleIndex:0}]}); s.conditionalFormatting.rules[0]={...r,ranges:[{firstRow:0,lastRow:19,firstCol:0,lastCol:7}],definition:{kind:'predicate',type:'containsBlanks',dialect:'base',formulas:[]}}; const g=buildSheetGeometry(s); if(g.kind!=='ready')throw Error('geometry')
 for(const box of [{x:0,y:0,width:3840,height:2160},{x:0,y:0,width:200,height:100}]) {const b=await prepareCfPaint(s,{viewport:{metrics:g.metrics,boxes:[box]}});try{expect(b?.layer.styleAt(6,4,undefined)?.fillColor).toBe('FF00FF00');expect(b?.layer.rects.some(r=>r.maxCol>7)).toBe(false)}finally{b?.dispose()}}
})
it('short-axis CF expansion does not reject an unchanged tall axis',()=>{
 const s=sheet();s.rows.push({index:999,cells:[{row:999,col:0,ref:'A1000',value:2,styleIndex:0}]});s.conditionalFormatting.rules[0]={...s.conditionalFormatting.rules[0],ranges:[{firstRow:0,lastRow:0,firstCol:1,lastCol:1}]};expect(computeMetrics(s).exportLimits).toBeUndefined()
})
it('statistics scan one shared scope once across row batches',async()=>{
 const {createCfMemory,createCfWorkBudget}=await import('../src/xlsx/cf/limits'); const {createCfStatistics}=await import('../src/xlsx/cf/statistics');const {createCfService}=await import('../src/xlsx/cf/service');
 const s=sheet(),r=s.conditionalFormatting.rules[0]; s.conditionalFormatting.rules[0]={...r,ranges:[{firstRow:0,lastRow:19,firstCol:0,lastCol:0}],definition:{kind:'aboveAverage',above:true,equal:false,dialect:'base',formulas:[]}} as typeof r;
 let reads=0; const stamp={owner:'stats',sheetId:'1',workbookValueGen:0,cfGen:1,viewGen:0,structureGen:0,clockGen:0};const source={stamp:()=>stamp,async bounds(){return {firstRow:0,lastRow:19,firstCol:0,lastCol:0}},async *pages(_id: string,rects: readonly import('../src/xlsx/cf/model').SourceRect[]){for(const b of rects)for(let row=b.firstRow;row<=b.lastRow;row++){reads++;yield {value:{rows:new Uint32Array([row]),cols:new Uint16Array([0]),values:[{kind:'number' as const,value:row}]},bytes:0,release(){},transfer(){throw Error('unused')}}}}};
 const memory=createCfMemory(),statistics={session:createCfStatistics(memory),stamp,results:new Map()},budget=createCfWorkBudget();
 try{for(let row=0;row<20;row++){const service=createCfService(s.conditionalFormatting,memory,{statistics,budget:()=>budget});try{const b=await service.prepareCf({stamp,sourceRects:[{firstRow:row,lastRow:row,firstCol:0,lastCol:0}],grantBytes:1024*1024,signal:new AbortController().signal},source);b.release()}finally{service.dispose()}}expect(reads).toBe(40)}finally{for(const out of statistics.results.values())if(out.kind==='ready')out.value.release();statistics.session.dispose()}expect(memory.currentBytes).toBe(0)
})
