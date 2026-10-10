/** Synthetic sealed B1 r4 owned-memo generation controls; not native parity. */
import { describe, it, expect } from 'vitest'
import { parseFormula } from '../src/xlsx/formula/parser'
import { evaluateFormulaInternal } from '../src/xlsx/formula/evaluator'
import { createReferenceServices, type StoredCell } from '../src/xlsx/formula/refs'
import type { AstNode, EvaluationContext, EvaluationValue, ReferenceServices } from '../src/xlsx/formula/types'

function setup(changeDuringRead:boolean, wrapped=false){
  let generation=1;let change=changeDuringRead;let reads=0
  const values=new Map([['0:0',10],['2:0',3],['2:1',4]])
  const bump=()=>{generation++;values.set('0:0',100);values.set('2:0',300);values.set('2:1',400)}
  const actual=createReferenceServices({generation:()=>generation,sheets:[{name:'S',sheetId:'s',workbookIndex:0}],definedNames:[{name:'Value',source:'$A$1+1',baseProvenance:'unknown'}],tables:[],sheetIdOfName:()=> 's',sheetNameOfId:()=> 'S',store:{
    stored:()=>Array.from(values,([key,value]):StoredCell=>{const[col,row]=key.split(':').map(Number);return {address:{sheetId:'s',col,row},value,origin:'input'}}),
    read:address=>{reads++;const value=values.get(`${address.col}:${address.row}`);if(change && address.col===2){change=false;bump()}return value===undefined?undefined:{value,origin:'input'}},
  }})
  const services:ReferenceServices=wrapped?{resolve:actual.resolve.bind(actual),bindName:actual.bindName.bind(actual),prepare:actual.prepare.bind(actual)}:actual
  const ctx:EvaluationContext={references:services,currentSheet:'S',currentCell:{col:4,row:0,absCol:false,absRow:false},currentAddress:{sheetId:'s',col:4,row:0},flatArgs:new WeakMap(),getCellValue:(_sheet,col,row)=>values.get(`${col}:${row}`)??null}
  return {ctx,services,values,bump,get reads(){return reads}}
}
const firstSiblings=[['A1+1',801],['A1',800],['Value',801]] as const

describe('B1 r4 sealed owned-memo controls',()=>{
  it.each(firstSiblings)('same invocation local memo refreshes completed %s sibling after read changes generation',(first,expected)=>{
    const f=setup(true)
    expect(f.ctx.nodeValues).toBeUndefined()
    expect(evaluateFormulaInternal(parseFormula(`SUM(${first},C1:C2)`),f.ctx)).toBe(expected)
    expect(f.ctx.nodeValues).toBeUndefined()
  })
  it.each(firstSiblings)('external Pending unwind restores local ownership before fresh %s sibling retry',(first,expected)=>{
    const f=setup(false);const marker={pending:true};const prepare=f.services.prepare;let pending=true
    f.services.prepare=(ref,ctx)=>{if(pending){pending=false;throw marker}return prepare(ref,ctx)}
    const ast=parseFormula(`SUM(${first},C1:C2)`)
    let caught:unknown;try{evaluateFormulaInternal(ast,f.ctx)}catch(error){caught=error}
    expect(caught).toBe(marker);expect(f.ctx.nodeValues).toBeUndefined()
    f.bump();expect(evaluateFormulaInternal(ast,f.ctx)).toBe(expected);expect(f.ctx.nodeValues).toBeUndefined()
  })
})

describe('B1 r4 ownership and scheduler paired controls',()=>{
  it.each(firstSiblings)('deliberately supplied memo entry for %s survives generation restart unchanged',(first)=>{
    const f=setup(true);const ast=parseFormula(`SUM(${first},C1:C2)`)
    if(ast.type!=='call')throw new Error('Invalid call fixture')
    const memo=new WeakMap<AstNode,EvaluationValue>();const sentinel=parseFormula('42')
    memo.set(sentinel,42);memo.set(ast.args[0],999);f.ctx.nodeValues=memo
    expect(evaluateFormulaInternal(ast,f.ctx)).toBe(1699)
    expect(f.ctx.nodeValues).toBe(memo);expect(memo.get(sentinel)).toBe(42);expect(memo.get(ast.args[0])).toBe(999)
  })
  it.each(firstSiblings)('unregistered service wrapper refreshes owned %s sibling after branded invalidation',(first,expected)=>{
    const f=setup(true,true)
    expect(evaluateFormulaInternal(parseFormula(`SUM(${first},C1:C2)`),f.ctx)).toBe(expected)
    expect(f.ctx.nodeValues).toBeUndefined()
  })
  it.each([false,true])('owned nested handler progress refreshes with the owned epoch, wrapper=%s',wrapped=>{
    const f=setup(true,wrapped)
    expect(evaluateFormulaInternal(parseFormula('SUM(SUM(A1+1),C1:C2)'),f.ctx)).toBe(801)
    expect(f.ctx.nodeValues).toBeUndefined()
  })
  it('completed outer value-name sibling and nested call recompute together without false active-name identity',()=>{
    const f=setup(true)
    expect(evaluateFormulaInternal(parseFormula('Value+SUM(A1+1,C1:C2)'),f.ctx)).toBe(902)
    expect(f.ctx.nodeValues).toBeUndefined()
    expect(f.ctx.unsupportedFeatures?.has('name-cycle')??false).toBe(false)
  })
  it('dead IF and IFERROR branches cause no generation reads/diagnostics and restore local context',()=>{
    const f=setup(true)
    expect(evaluateFormulaInternal(parseFormula('IF(FALSE,SUM(A1+1,C1:C2),8)'),f.ctx)).toBe(8)
    expect(evaluateFormulaInternal(parseFormula('IFERROR(8,SUM(A1+1,C1:C2))'),f.ctx)).toBe(8)
    expect(f.reads).toBe(0);expect(f.ctx.nodeValues).toBeUndefined();expect(f.ctx.unsupportedFeatures?.size??0).toBe(0)
  })
  it('user unwind error resembling an owned restart preserves exact identity and restores local memo',()=>{
    const f=setup(false);const marker=new Error('owned memo epoch restart');marker.name='OwnedMemoGenerationRestart'
    f.services.prepare=()=>{throw marker}
    let caught:unknown;try{evaluateFormulaInternal(parseFormula('SUM(A1+1,C1:C2)'),f.ctx)}catch(error){caught=error}
    expect(caught).toBe(marker);expect(f.ctx.nodeValues).toBeUndefined()
  })
  it('a previously scoped map deliberately supplied later remains caller-owned',()=>{
    const f=setup(false)
    let saved:WeakMap<AstNode,EvaluationValue>|undefined
    const read=f.ctx.getCellValue
    f.ctx.getCellValue=(sheet,col,row)=>{saved=f.ctx.nodeValues;return read?.(sheet,col,row)??null}
    expect(evaluateFormulaInternal(parseFormula('A1+1'),f.ctx)).toBe(11);expect(f.ctx.nodeValues).toBeUndefined()
    if(!saved)throw new Error('Missing captured scoped memo fixture')
    const memo=saved;const sentinel=parseFormula('42');const ast=parseFormula('SUM(A1+1,C1:C2)')
    if(ast.type!=='call')throw new Error('Invalid call fixture')
    memo.set(sentinel,42);memo.set(ast.args[0],999);f.ctx.nodeValues=memo
    const prepare=f.services.prepare;let change=true
    f.services.prepare=(ref,ctx)=>{if(change){change=false;f.bump()}return prepare(ref,ctx)}
    expect(evaluateFormulaInternal(ast,f.ctx)).toBe(1699)
    expect(f.ctx.nodeValues).toBe(memo);expect(memo.get(sentinel)).toBe(42);expect(memo.get(ast.args[0])).toBe(999)
  })
  it('owned restarts without an argument store retain the shared bounded resource policy',()=>{
    let generation=1;let reads=0;const abort=new Error('harness bounded after twenty generation changes')
    const services=createReferenceServices({generation:()=>generation,sheets:[{name:'S',sheetId:'s',workbookIndex:0}],definedNames:[],tables:[],sheetIdOfName:()=> 's',sheetNameOfId:()=> 'S',store:{stored:()=>[{address:{sheetId:'s',col:2,row:0},value:3,origin:'input'}],read:()=>{if(++reads>20)throw abort;generation++;return {value:3,origin:'input'}}}})
    const issues:string[]=[];const ctx:EvaluationContext={references:services,currentSheet:'S',getCellValue:()=>10,reportFormulaIssue:issue=>issues.push(issue.feature??'')}
    let result:unknown;expect(()=>{result=evaluateFormulaInternal(parseFormula('SUM(A1+1,C1:C2)'),ctx)}).not.toThrow()
    expect(result).toEqual({kind:'formula-error',code:'#NAME?'})
    expect(issues).toEqual(['reference-generation-unstable']);expect(reads).toBeLessThanOrEqual(9)
    expect(ctx.nodeValues).toBeUndefined()
  })
})
