/** Synthetic sealed QUALITY grammar/source/API boundaries; no native parity claim. */
import { describe, it, expect } from 'vitest'
import { parseFormula } from '../src/xlsx/formula/parser'
import { formatFormula, translateSharedFormula } from '../src/xlsx/formula/shared'
import { evaluateFormulaInternal, isEvaluationError } from '../src/xlsx/formula/evaluator'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import { createReferenceServices, isReferenceNode, type StoredCell } from '../src/xlsx/formula/refs'
import type { CellAddress, EvaluationContext, ElementValue } from '../src/xlsx/formula/types'
import type { XlsxCell, XlsxDocument } from '../src/xlsx/types'

function cell(ref:string,col:number,row:number,value:number,extra:Partial<XlsxCell>={}):XlsxCell{return {ref,col,row,value,styleIndex:0,...extra}}
function fixture(size=3){
  let reads=0;let visits=0;const addresses:CellAddress[]=[]
  const cells:StoredCell[]=Array.from({length:size},(_,row)=>({address:{sheetId:'s',col:0,row},value:1,origin:'input'}))
  const services=createReferenceServices({generation:()=>1,sheets:[{sheetId:'s',name:'S',workbookIndex:0}],sheetIdOfName:()=> 's',sheetNameOfId:()=> 'S',definedNames:[],tables:[],store:{
    stored:function*(){for(const stored of cells){visits++;yield stored}},
    read:address=>{reads++;addresses.push(address);return Number.isInteger(address.row)&&address.row<size?{value:1,origin:'input'}:undefined},
  }})
  const ctx:EvaluationContext={references:services,currentSheet:'S',flatArgs:new WeakMap()}
  return {ctx,services,addresses,get reads(){return reads},get visits(){return visits},ref(source:string){const ast=parseFormula(source);if(!isReferenceNode(ast))throw new Error('Invalid reference fixture');const ref=services.resolve(ast,ctx);if(!ref||ref.kind!=='resolved-reference')throw new Error('Missing resolved fixture');return ref}}
}

describe('B1 r5 Q1 compound implicit intersection source',()=>{
  it.each(['@(A1:A3 A2:A4)','@(A1 A1)',"@('S:T'!$A1)",'@(A1+1)','@(-A1)'])('translated %s preserves operand AST after formatting/reparse',source=>{
    const moved=translateSharedFormula(source,1,1)
    expect(parseFormula(moved.formula)).toEqual(moved.ast)
  })
  it('actual shared intersection master computes20 and follower30 after translation',()=>{
    const master=cell('B2',1,1,77,{hasCachedValue:true,formula:'@(A1:A3 A2:A4)',sharedFormula:{si:0,ref:'B2:B3'}})
    const follower=cell('B3',1,2,77,{hasCachedValue:true,sharedFormula:{si:0}})
    const doc:XlsxDocument={sheets:[{name:'S',sheetId:'s',workbookIndex:0,cols:[],merges:[],mergeRanges:[],rows:[{index:0,cells:[cell('A1',0,0,10)]},{index:1,cells:[cell('A2',0,1,20),master]},{index:2,cells:[cell('A3',0,2,30),follower]},{index:3,cells:[cell('A4',0,3,40)]}]}],definedNames:[],tables:[]}
    evaluateWorkbookFormulas(doc,{forceRecalc:true})
    expect(master.value).toBe(20);expect(follower.value).toBe(30)
    expect(parseFormula(follower.formula??'')).toEqual(translateSharedFormula(master.formula??'',0,1).ast)
    expect(doc.diagnostics?.some(issue=>issue.feature==='formula-syntax')??false).toBe(false)
  })
})

describe('B1 r5 Q2 cell-like 3D sheet-name axes',()=>{
  it.each(['S1:S2!A:A','S1:S2!$A:C','S1:S2!1:1','S1:S2!$1:3'])('legal %s uses the full approved GridReference target grammar',source=>{
    const ast=parseFormula(source);expect(ast.type).toBe('ref3d')
    expect(ast).toEqual(parseFormula(`'S1:S2'!${source.slice(source.indexOf('!')+1)}`))
  })
  it.each(["'S1:S2'!A:A","'S1:S2'!1:1"])('shared quoted %s retains its translated syntax after canonical formatting',source=>{
    const moved=translateSharedFormula(source,1,1)
    expect(parseFormula(moved.formula)).toEqual(moved.ast)
  })
  it('workbook SUM of quoted/unquoted whole-axis spellings both computes sparse10+20=30',()=>{
    const unquoted=cell('B1',1,0,77,{formula:'SUM(S1:S2!A:A)',hasCachedValue:true})
    const quoted=cell('C1',2,0,77,{formula:"SUM('S1:S2'!A:A)",hasCachedValue:true})
    const doc:XlsxDocument={sheets:['S1','S2'].map((name,i)=>({name,sheetId:name,workbookIndex:i,cols:[],merges:[],mergeRanges:[],rows:[{index:0,cells:[cell('A1',0,0,i===0?10:20),...(i===0?[unquoted,quoted]:[])]}]})),definedNames:[],tables:[]}
    evaluateWorkbookFormulas(doc,{forceRecalc:true})
    expect(quoted.value).toBe(30);expect(unquoted.value).toBe(30)
    expect(doc.diagnostics?.some(issue=>issue.feature==='formula-syntax')??false).toBe(false)
  })
  it('ordinary cell ranges remain ordinary and full3D cells still roundtrip',()=>{
    expect(parseFormula('S1:S2')).toMatchObject({type:'range'})
    const ast=parseFormula('S1:S2!$A$1');expect(ast.type).toBe('ref3d');expect(parseFormula(formatFormula(ast))).toEqual(ast)
  })
})

describe('B1 r5 Q3 numeric readAt API boundary',()=>{
  it.each([[NaN,0],[0.5,0],[0,NaN],[0,0.5],[Infinity,0],[0,Infinity],[-1,0],[3,0]] as const)('invalid offset row=%s col=%s rejects before backend callback',(row,col)=>{
    const f=fixture();let value:ElementValue|undefined;let caught:unknown
    try{value=f.ref('A1:A3').readAt(0,row,col)}catch(error){caught=error}
    expect(f.reads).toBe(0);expect(f.addresses).toEqual([])
    expect(isEvaluationError(value)||caught instanceof Error).toBe(true) // no native error-code assumption
  })
  it.each([NaN,Infinity,0.5,-1])('malformed areaIndex=%s is rejected before a backend lookup',index=>{
    const f=fixture();let value:ElementValue|undefined;let caught:unknown
    try{value=f.ref('A1:A3').readAt(index,0,0)}catch(error){caught=error}
    expect(f.reads).toBe(0);expect(f.addresses).toEqual([])
    expect(isEvaluationError(value)||caught instanceof Error).toBe(true)
  })
  it('valid500000th row preserves full geometry and exactly one lazy read, no stored enumeration',()=>{
    const f=fixture();const ref=f.ref('A:A')
    expect(ref.areas[0].rows).toBe(1048576);expect(ref.readAt(0,499999,0)).toBeNull()
    expect(f.reads).toBe(1);expect(f.visits).toBe(0);expect(f.addresses).toEqual([{sheetId:'s',col:0,row:499999}])
  })
  it.each([100,1000])('%s sparse populated cells still require only population reads/visits',size=>{
    const f=fixture(size);expect(evaluateFormulaInternal(parseFormula('SUM(A:A)'),f.ctx)).toBe(size)
    expect(f.reads).toBe(size);expect(f.visits).toBe(size)
  })
})
