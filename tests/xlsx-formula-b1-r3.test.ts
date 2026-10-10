/** B1 r3 sealed SPEC counterexamples: structural contract, not native parity. */
import { describe, expect, it, vi } from 'vitest'
import * as referenceModule from '../src/xlsx/formula/refs'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import type { XlsxCell, XlsxDocument } from '../src/xlsx/types'
import { evaluateFormulaInternal } from '../src/xlsx/formula/evaluator'
import { parseFormula } from '../src/xlsx/formula/parser'
import { createReferenceServices, isReferenceNode, type StoredCell } from '../src/xlsx/formula/refs'
import type { AstNode, EvaluationContext, EvaluationValue, ResolvedRef } from '../src/xlsx/formula/types'

function setup() {
  let generation = 1
  let pending = true
  let reads = 0
  const marker = { pending: true }
  const values = new Map([['0:0',10],['0:1',11],['2:0',3],['2:1',4]])
  const services = createReferenceServices({
    generation: () => generation,
    sheets: [{ name:'S',sheetId:'s',workbookIndex:0 }],definedNames:[],tables:[],
    sheetIdOfName: () => 's',sheetNameOfId: () => 'S',
    store: {
      stored: () => Array.from(values,([key,value]): StoredCell => {
        const [col,row]=key.split(':').map(Number)
        return {address:{sheetId:'s',col,row},value,origin:'input'}
      }),
      read: address => {
        reads++
        if(address.col===2 && address.row===1 && pending) throw marker
        const value=values.get(`${address.col}:${address.row}`)
        return value===undefined ? undefined : {value,origin:'input'}
      },
    },
  })
  const ctx:EvaluationContext={currentSheet:'S',currentCell:{col:4,row:0,absCol:false,absRow:false},references:services,flatArgs:new WeakMap(),functionWork:new WeakMap()}
  return {ctx,services,marker,values,get reads(){return reads},resume:()=>{pending=false},next:()=>{generation++}}
}
function resolved(source:string, f:ReturnType<typeof setup>): ResolvedRef {
  const ast=parseFormula(source)
  if(!isReferenceNode(ast)) throw new Error('Invalid reference fixture')
  const ref=f.services.resolve(ast,f.ctx)
  if(!ref || ref.kind!=='resolved-reference') throw new Error('Missing resolved fixture')
  return ref
}

describe('B1 r3 sealed minimal controls',()=>{
  it.each([false,true])('generation refresh discards completed sibling references, provided memo=%s',provided=>{
    const f=setup();const ast=parseFormula('SUM(A1:A2,C1:C2)')
    const memo=new WeakMap<AstNode,EvaluationValue>();const sentinel=parseFormula('42');memo.set(sentinel,42)
    if(provided)f.ctx.nodeValues=memo
    let caught:unknown;try{evaluateFormulaInternal(ast,f.ctx)}catch(error){caught=error}
    expect(caught).toBe(f.marker)
    for(const [key,value] of [['0:0',100],['0:1',200],['2:0',300],['2:1',400]] as const)f.values.set(key,value)
    f.next();f.resume()
    expect(evaluateFormulaInternal(ast,f.ctx)).toBe(1000)
    if(provided){expect(f.ctx.nodeValues).toBe(memo);expect(memo.get(sentinel)).toBe(42)}else expect(f.ctx.nodeValues).toBeUndefined()
  })
  it('user stale-text read exception escapes with exact identity after one read',()=>{
    const marker=new Error('stale cursor generation (user callback exception)')
    const abort=new Error('harness bound after twenty reads')
    let reads=0
    const services=createReferenceServices({generation:()=>1,sheets:[{name:'S',sheetId:'s',workbookIndex:0}],definedNames:[],tables:[],sheetIdOfName:()=> 's',sheetNameOfId:()=> 'S',store:{stored:()=>[{address:{sheetId:'s',col:0,row:0},value:1,origin:'input'}],read:()=>{if(++reads>20)throw abort;throw marker}}})
    let caught:unknown
    try{evaluateFormulaInternal(parseFormula('SUM(A1:A2)'),{references:services,currentSheet:'S',flatArgs:new WeakMap()})}catch(error){caught=error}
    expect(caught).toBe(marker);expect(reads).toBe(1)
  })
  it.each(['all','populated'] as const)('%s duplicate single-cell union advances to its second occurrence then EOF',mode=>{
    const f=setup();const ref=resolved('(A1,A1)',f);const cursor=ref.openCursor(mode)
    expect(ref.peek(cursor)).toMatchObject({areaIndex:0,value:10})
    ref.advance(cursor);expect(ref.peek(cursor)).toMatchObject({areaIndex:1,value:10})
    ref.advance(cursor);expect(ref.peek(cursor)).toBeUndefined()
  })
  it.each([false,true])('same-generation sibling suspension keeps result28 and memo ownership, provided memo=%s',provided=>{
    const f=setup();const ast=parseFormula('SUM(A1:A2,C1:C2)')
    const memo=new WeakMap<AstNode,EvaluationValue>();const sentinel=parseFormula('42');memo.set(sentinel,42)
    if(provided)f.ctx.nodeValues=memo
    let caught:unknown;try{evaluateFormulaInternal(ast,f.ctx)}catch(error){caught=error}
    expect(caught).toBe(f.marker);f.resume();expect(evaluateFormulaInternal(ast,f.ctx)).toBe(28)
    expect(f.reads).toBe(5)
    if(provided){expect(f.ctx.nodeValues).toBe(memo);expect(memo.get(sentinel)).toBe(42)}else expect(f.ctx.nodeValues).toBeUndefined()
  })
})

describe('B1 r3 transition and callback generation controls',()=>{
  it.each(['all','populated'] as const)('%s retains offset zero of every new area, including both duplicate A1 occurrences',mode=>{
    const f=setup();f.resume();f.values.delete('2:0');f.values.delete('2:1')
    const ref=resolved('(B1:B2,A1,C1:C2,A1)',f);const cursor=ref.openCursor(mode)
    const entries:Array<{area:number;row:number;col:number;value:EvaluationValue}>=[]
    for(let attempts=0;attempts<10;attempts++){
      const entry=ref.peek(cursor);if(!entry)break
      entries.push({area:entry.areaIndex,row:entry.offsetRow,col:entry.offsetCol,value:entry.value});ref.advance(cursor)
    }
    expect(entries).toEqual(mode==='all' ? [
      {area:0,row:0,col:0,value:null},{area:0,row:1,col:0,value:null},
      {area:1,row:0,col:0,value:10},{area:2,row:0,col:0,value:null},{area:2,row:1,col:0,value:null},{area:3,row:0,col:0,value:10},
    ] : [{area:1,row:0,col:0,value:10},{area:3,row:0,col:0,value:10}])
    expect(ref.peek(cursor)).toBeUndefined()
  })
  it.each(['all','populated'] as const)('%s skips zero-size areas without consuming the following real occurrence',mode=>{
    const f=setup()
    const services=createReferenceServices({generation:()=>1,sheets:[{name:'S',sheetId:'s',workbookIndex:0}],definedNames:[],tables:[{id:'empty',name:'Empty',displayName:'Empty',sheetId:'s',partPath:'empty.xml',extent:{sheetId:'s',firstCol:0,firstRow:0,cols:1,rows:1},headerRowCount:1,totalsRowCount:0,columns:[{id:'a',name:'Amount',index:0}]}],sheetIdOfName:()=> 's',sheetNameOfId:()=> 'S',store:{stored:()=>[{address:{sheetId:'s',col:0,row:0},value:10,origin:'input'}],read:()=>({value:10,origin:'input'})}})
    f.services=services;f.ctx.references=services
    const ref=resolved('(Empty[#Data],A1,Empty[#Data],A1)',f);const cursor=ref.openCursor(mode)
    expect(ref.peek(cursor)).toMatchObject({areaIndex:1,offsetRow:0,offsetCol:0,value:10})
    ref.advance(cursor);expect(ref.peek(cursor)).toMatchObject({areaIndex:3,offsetRow:0,offsetCol:0,value:10})
    ref.advance(cursor);expect(ref.peek(cursor)).toBeUndefined()
  })
  it('ordinary callback errors also preserve exact identity with a single read',()=>{
    const marker=new Error('ordinary user callback failure');let reads=0
    const services=createReferenceServices({generation:()=>1,sheets:[{name:'S',sheetId:'s',workbookIndex:0}],definedNames:[],tables:[],sheetIdOfName:()=> 's',sheetNameOfId:()=> 'S',store:{stored:()=>[{address:{sheetId:'s',col:0,row:0},value:1,origin:'input'}],read:()=>{reads++;throw marker}}})
    let caught:unknown;try{evaluateFormulaInternal(parseFormula('SUM(A1:A2)'),{references:services,currentSheet:'S',flatArgs:new WeakMap()})}catch(error){caught=error}
    expect(caught).toBe(marker);expect(reads).toBe(1)
  })
  it('generation changed from a store.read callback restarts the complete footprint and then succeeds',()=>{
    let generation=1;let change=true;let reads=0
    const values=new Map([[0,10],[1,20]])
    const services=createReferenceServices({generation:()=>generation,sheets:[{name:'S',sheetId:'s',workbookIndex:0}],definedNames:[],tables:[],sheetIdOfName:()=> 's',sheetNameOfId:()=> 'S',store:{stored:()=>Array.from(values,([row,value])=>({address:{sheetId:'s',col:0,row},value,origin:'input' as const})),read:address=>{reads++;const value=values.get(address.row);if(change){change=false;generation++;values.set(0,100);values.set(1,200)}return value===undefined?undefined:{value,origin:'input'}}}})
    const ctx:EvaluationContext={references:services,currentSheet:'S',flatArgs:new WeakMap()};let result:unknown
    expect(()=>{result=evaluateFormulaInternal(parseFormula('SUM(A1:A2)'),ctx)}).not.toThrow()
    expect(result).toBe(300);expect(reads).toBe(3)
    expect(ctx.unsupportedFeatures?.has('reference-generation-unstable')??false).toBe(false)
  })
  it.each(['read','stored'] as const)('persistent generation changes inside store.%s stop with bounded resource diagnostic, no invented native result',channel=>{
    let generation=1;let callbacks=0;const abort=new Error('harness callback bound after twenty')
    const change=()=>{if(++callbacks>20)throw abort;generation++}
    const services=createReferenceServices({generation:()=>generation,sheets:[{name:'S',sheetId:'s',workbookIndex:0}],definedNames:[],tables:[],sheetIdOfName:()=> 's',sheetNameOfId:()=> 'S',store:{stored:()=>{if(channel==='stored')change();return [{address:{sheetId:'s',col:0,row:0},value:1,origin:'input'}]},read:()=>{if(channel==='read')change();return {value:1,origin:'input'}}}})
    const issues:string[]=[];const ctx:EvaluationContext={references:services,currentSheet:'S',flatArgs:new WeakMap(),reportFormulaIssue:issue=>issues.push(issue.feature??'')}
    let result:unknown;expect(()=>{result=evaluateFormulaInternal(parseFormula('SUM(A1:A2)'),ctx)}).not.toThrow()
    expect(result).toEqual({kind:'formula-error',code:'#NAME?'})
    expect(ctx.unsupportedFeatures?.has('reference-generation-unstable')).toBe(true)
    expect(issues).toEqual(['reference-generation-unstable'])
    expect(callbacks).toBeGreaterThan(0);expect(callbacks).toBeLessThanOrEqual(16)
  })
  it('completed full argument progress refreshes on the next generation while preserving its owning map',()=>{
    const f=setup();f.resume();const ast=parseFormula('SUM(A1:A2,C1:C2)');const flatArgs=f.ctx.flatArgs
    expect(evaluateFormulaInternal(ast,f.ctx)).toBe(28)
    for(const [key,value] of [['0:0',100],['0:1',200],['2:0',300],['2:1',400]] as const)f.values.set(key,value)
    f.next();f.ctx.nodeValues=new WeakMap()
    expect(evaluateFormulaInternal(ast,f.ctx)).toBe(1000);expect(f.ctx.flatArgs).toBe(flatArgs)
  })
})


describe('B1 r3 resource cache and identity controls',()=>{
  it('private-class lookalike name/message still escapes unchanged after one read',()=>{
    const marker=new Error('stale reference generation (bound 1)');marker.name='ReferenceGenerationChanged';let reads=0
    const abort=new Error('lookalike harness bound after twenty reads')
    const services=createReferenceServices({generation:()=>1,sheets:[{name:'S',sheetId:'s',workbookIndex:0}],definedNames:[],tables:[],sheetIdOfName:()=> 's',sheetNameOfId:()=> 'S',store:{stored:()=>[{address:{sheetId:'s',col:0,row:0},value:1,origin:'input'}],read:()=>{if(++reads>20)throw abort;throw marker}}})
    let caught:unknown;try{evaluateFormulaInternal(parseFormula('SUM(A1:A2)'),{references:services,currentSheet:'S',flatArgs:new WeakMap()})}catch(error){caught=error}
    expect(caught).toBe(marker);expect(reads).toBe(1)
  })
  it('bounded instability retains workbook cache77 and keeps the dead8 frame silent',()=>{
    const live:XlsxCell={ref:'B3',col:1,row:2,value:77,hasCachedValue:true,styleIndex:0,formula:'SUM(A1:A2)'}
    const dead:XlsxCell={ref:'B4',col:1,row:3,value:77,hasCachedValue:true,styleIndex:0,formula:'IF(FALSE,SUM(A1:A2),8)'}
    const doc:XlsxDocument={sheets:[{name:'S',sheetId:'s',workbookIndex:0,cols:[],merges:[],mergeRanges:[],rows:[{index:0,cells:[{ref:'A1',col:0,row:0,value:1,styleIndex:0}]},{index:1,cells:[{ref:'A2',col:0,row:1,value:2,styleIndex:0}]},{index:2,cells:[live]},{index:3,cells:[dead]}]}],definedNames:[],tables:[]}
    const actual=createReferenceServices;let reads=0
    const spy=vi.spyOn(referenceModule,'createReferenceServices').mockImplementation(deps=>{
      let generation=deps.generation()
      return actual({...deps,generation:()=>generation,store:{stored:()=>deps.store.stored(),read:address=>{reads++;generation++;return deps.store.read(address)}}})
    })
    try{
      expect(()=>evaluateWorkbookFormulas(doc,{forceRecalc:true})).not.toThrow()
      expect(live.value).toBe(77);expect(dead.value).toBe(8)
      expect(doc.diagnostics?.some(issue=>issue.feature==='reference-generation-unstable')).toBe(true)
      expect(doc.diagnostics?.some(issue=>issue.message.includes('B4'))).toBe(false)
      expect(reads).toBeGreaterThan(0);expect(reads).toBeLessThanOrEqual(16)
    }finally{spy.mockRestore()}
  })
})
