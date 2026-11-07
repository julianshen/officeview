import { expect, it } from 'vitest'
import { prepareCfWireScope, prepareCfTextScope } from '../src/xlsx/cf/text'
import { storeCfSource } from '../src/xlsx/cf/render'
import { createCfMemory } from '../src/xlsx/cf/limits'
import type { CfReadResult, SourceStore } from '../src/xlsx/view/contracts'
const stamp={owner:'t',sheetId:'1',workbookValueGen:1,cfGen:1,viewGen:1,structureGen:1,clockGen:0}
const budget={reserve(_n:number){},release(_n:number){}}
const result:CfReadResult={kind:'ready',stamp,sheetId:'1',states:[{row:2,col:2,properties:{'fill.bg':'theme:4:0'}}],definitions:[],definitionsVersion:1,diagnostics:[]}
it('wire paint coverage includes winning cells only, not the requested viewport',async()=>{
 const scope=await prepareCfWireScope(async()=>({result,release(){}}),[{r0:0,r1:50,c0:0,c1:10}],{budget})
 expect(scope!.regions).toEqual([{minRow:2,maxRow:2,minCol:2,maxCol:2}]);scope!.dispose()
})
it('wire scope synchronously withdraws on structure change or close',async()=>{
 let current=true
 const scope=await prepareCfWireScope(async()=>({result,release(){}}),[{r0:2,r1:2,c0:2,c1:2}],{budget,current:()=>current})
 expect(scope!.effectiveStyle(2,2,undefined)?.fillColor).toBeDefined()
 current=false
 expect(scope!.effectiveStyle(2,2,undefined)).toBeUndefined()
 expect(scope!.describe(2,2)).toBeUndefined();scope!.dispose()
})
it('importing bounds cannot certify a prefix as a complete statistics range',async()=>{
 const store={stamp:{docId:'t',sheetId:'1',valueGen:1},completeness:'loading',extent:{},readRows:async()=>{throw Error('must not read')}} as unknown as SourceStore
 const memory=createCfMemory()
 expect(await storeCfSource(store,stamp,memory).bounds!('1')).toBeUndefined()
})
it('resident text uses the workbook colour context by default',async()=>{
 const sheet={name:'s',sheetId:'1',cols:[],rows:[{index:2,cells:[{row:2,col:2,ref:'C3',value:1,styleIndex:0}]}],merges:[],mergeRanges:[],tableStyleSource:{colors:{theme:{accent1:'#ff0000'}}},conditionalFormatting:{diagnostics:[],rules:[{id:0,priority:1,documentOrder:0,stopIfTrue:false,ranges:[{firstRow:2,lastRow:2,firstCol:2,lastCol:2}],writes:new Set(['fill']),dxf:{'fill.bg':'theme:4:0'},definition:{kind:'cellIs',operator:'greaterThan',operands:[{kind:'number',value:0}],dialect:'base',formulas:[]}}]}} as unknown as import('../src/xlsx/types').XlsxSheet
 const scope=await prepareCfTextScope(sheet,{rects:[{r0:2,r1:2,c0:2,c1:2}],signal:new AbortController().signal,budget})
 expect(scope!.effectiveStyle(2,2,undefined)?.fillColor?.toLowerCase()).toMatch(/ff0000$/);scope!.dispose()
})
it('paged region layer preserves table styles and applies custom CF colours only to winning cells',async()=>{
 const base={rects:[{minRow:0,maxRow:9,minCol:0,maxCol:3}],retainedBytes:100,styleAt:(_r:number,_c:number,own:unknown)=>({...own as object,numFmtId:0,fillColor:'table-fill'})}
 const scope=await prepareCfWireScope(async()=>({result,release(){}}),[{r0:0,r1:50,c0:0,c1:10}],{budget,base,colors:{theme:{accent1:'#ff0000'}}})
 expect(scope!.paintLayer!.styleAt(0,0,undefined)?.fillColor).toBe('table-fill')
 expect(scope!.paintLayer!.styleAt(2,2,undefined)?.fillColor?.toLowerCase()).toMatch(/ff0000$/)
 expect(scope!.paintLayer!.rects).toHaveLength(2);scope!.dispose()
})
it('paged formatting transports custom theme and indexed colours for CF even without drawings',async()=>{
 const {workbookFormattingWithColorContext}=await import('../src/xlsx/storage/pagedFormatting')
 const out=await workbookFormattingWithColorContext({},'<styleSheet><colors><indexedColors><rgbColor rgb="FF123456"/></indexedColors></colors></styleSheet>','<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="custom"><a:accent1><a:srgbClr val="FF0000"/></a:accent1></a:clrScheme></a:themeElements></a:theme>')
 expect(out.formatting.colors).toEqual(out.colorContext)
 expect(out.formatting.colors?.theme?.accent1).toBeDefined()
 expect(out.formatting.colors?.indexed?.[0]).toBeDefined()
})
it('worker compatibility formatting finds the workbook palette on a later CF sheet',async()=>{
 const {createObjectWorkbookSource}=await import('../src/xlsx/view/source')
 const plain={name:'plain',sheetId:'1',cols:[],rows:[],merges:[],mergeRanges:[]}
 const colors={theme:{accent1:'#ff0000'},indexed:['FF123456']}
 const source=createObjectWorkbookSource({cellStyles:[],sheets:[plain,{...plain,name:'cf',sheetId:'2',tableStyleSource:{colors,root:{}}}]},'palette')
 expect(source.formatting?.().colors).toEqual(colors)
 await source.release()
})
