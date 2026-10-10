/** B2 r2: official populated-text datasets and bounded private lifecycle controls.
 * Official expected results are documented-derived, not new native measurements.
 * https://support.microsoft.com/en-us/excel/functions/sumifs-function
 * https://support.microsoft.com/en-us/excel/functions/averageif-function */
import { describe, it, expect } from 'vitest'
import { evaluateFormula, evaluateFormulaInternal, formulaError } from '../src/xlsx/formula/evaluator'
import { parseFormula } from '../src/xlsx/formula/parser'
import { createReferenceServices, type StoredCell } from '../src/xlsx/formula/refs'
import type { AstNode, ElementValue, EvaluationContext, EvaluationValue } from '../src/xlsx/formula/types'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import type { XlsxCell, XlsxDocument } from '../src/xlsx/types'

function cell(col:number,row:number,value:ElementValue,origin:StoredCell['origin']='input'):StoredCell {
  return {address:{sheetId:'s',col,row},value,origin}
}
function fixture(cells:StoredCell[], pendingAt?:{col:number;row:number}) {
  const reads=new Map<string,number>();let pending=!!pendingAt;const marker=new Error('negative text target pending')
  const services=createReferenceServices({generation:()=>1,sheets:[{sheetId:'s',name:'S',workbookIndex:0}],sheetIdOfName:n=>n==='S'?'s':undefined,sheetNameOfId:()=> 'S',definedNames:[{name:'Criterion',source:'"<>Bananas"',baseProvenance:'unknown'}],tables:[],store:{
    stored:()=>cells,read:a=>{const key=`${a.col}:${a.row}`;reads.set(key,(reads.get(key)??0)+1);if(pending&&a.col===pendingAt?.col&&a.row===pendingAt.row){pending=false;throw marker}const c=cells.find(c=>c.address.col===a.col&&c.address.row===a.row);return c?{value:c.value,origin:c.origin}:undefined},
  }})
  const ctx:EvaluationContext={typedValues:true,references:services,currentSheet:'S',currentAddress:{sheetId:'s',col:6,row:0},currentCell:{col:6,row:0,absCol:false,absRow:false},flatArgs:new WeakMap(),functionWork:new WeakMap(),getCellValue:(_s,col,row)=>cells.find(c=>c.address.col===col&&c.address.row===row)?.value??null}
  return {ctx,reads,marker}
}
function textFixture(text=['Apples','Bananas','Carrots']) {return text.flatMap((value,row)=>[cell(0,row,value),cell(1,row,(row+1)*10)])}

describe('B2 r2 negative populated text criteria',()=>{
  it('Microsoft SUMIFS <>Bananas exact documented result30',()=>{
    const quantity=[5,4,15,3,22,12,10,33],product=['Apples','Apples','Artichokes','Artichokes','Bananas','Bananas','Carrots','Carrots'],seller=['Tom','Sarah','Tom','Sarah','Tom','Sarah','Tom','Sarah']
    const cells=[cell(0,0,'Quantity Sold'),cell(1,0,'Product'),cell(2,0,'Salesperson')]
    for(let i=0;i<8;i++)cells.push(cell(0,i+1,quantity[i]),cell(1,i+1,product[i]),cell(2,i+1,seller[i]))
    const f=fixture(cells);expect(evaluateFormula('SUMIFS(A2:A9,B2:B9,"<>Bananas",C2:C9,"Tom")',f.ctx)).toBe(30);expect(f.ctx.unsupportedFeatures?.size??0).toBe(0)
  })
  it('Microsoft AVERAGEIF negative wildcard exact documented result18589',()=>{
    const region=['East','West','North','South (New Office)','MidWest'],profit=[45678,23789,-4789,0,9678]
    const cells=[cell(0,0,'Region'),cell(1,0,'Profits (Thousands)')]
    for(let i=0;i<5;i++)cells.push(cell(0,i+1,region[i]),cell(1,i+1,profit[i]))
    const f=fixture(cells);expect(evaluateFormula('AVERAGEIF(A2:A6,"<>*(New Office)",B2:B6)',f.ctx)).toBe(18589);expect(f.ctx.unsupportedFeatures?.size??0).toBe(0)
  })
  it.each([
    ['COUNTIF(A1:A3,"%s")',1,2],['COUNTIFS(A1:A3,"%s")',1,2],
    ['SUMIF(A1:A3,"%s",B1:B3)',20,40],['SUMIFS(B1:B3,A1:A3,"%s")',20,40],
    ['AVERAGEIF(A1:A3,"%s",B1:B3)',20,20],['AVERAGEIFS(B1:B3,A1:A3,"%s")',20,20],
  ] as const)('%s pairs equality and negation on only populated text',(source,positive,negative)=>{
    expect(evaluateFormula(source.replace('%s','=Bananas'),fixture(textFixture()).ctx)).toBe(positive)
    expect(evaluateFormula(source.replace('%s','<>Bananas'),fixture(textFixture()).ctx)).toBe(negative)
  })
  it.each(['a~*b','a~?b','a~~b'])('negated escaped literal %s keeps accepted ASCII semantics',pattern=>{
    const cells=textFixture(['a*b','a?b','a~b','other']);expect(evaluateFormula(`COUNTIF(A1:A4,"=${pattern}")`,fixture(cells).ctx)).toBe(1);expect(evaluateFormula(`COUNTIF(A1:A4,"<>${pattern}")`,fixture(cells).ctx)).toBe(3)
  })
  it('negative wildcard/case is the complement of the paired ASCII text match',()=>{
    const cells=textFixture(['Alpha','alphabet','Beta']);expect(evaluateFormula('COUNTIF(A1:A3,"=AL*")',fixture(cells).ctx)).toBe(2);expect(evaluateFormula('COUNTIF(A1:A3,"<>AL*")',fixture(cells).ctx)).toBe(1)
  })
  it.each(['"<>Bananas"','D1','Criterion'])('criterion source %s preserves scalar input origin',source=>{
    const f=fixture([...textFixture(),cell(3,0,'<>Bananas','formula')]);expect(evaluateFormula(`COUNTIF(A1:A3,${source})`,f.ctx)).toBe(2);expect(f.ctx.unsupportedFeatures?.size??0).toBe(0)
  })
  it('only selected negative target suspends; resume does not re-add earlier target or clear provided memo',()=>{
    const f=fixture(textFixture(['Apples','Bananas','Carrots']),{col:1,row:2});const ast=parseFormula('SUMIF(A1:A3,"<>Bananas",B1:B3)');const memo=new WeakMap<AstNode,EvaluationValue>();const sentinel=parseFormula('91');memo.set(sentinel,91);f.ctx.nodeValues=memo
    let caught:unknown;try{evaluateFormulaInternal(ast,f.ctx)}catch(error){caught=error};expect(caught).toBe(f.marker);expect(f.reads.get('1:1')).toBeUndefined();expect(f.reads.get('1:0')).toBe(1)
    expect(evaluateFormulaInternal(ast,f.ctx)).toBe(40);expect(f.reads.get('1:0')).toBe(1);expect(f.reads.get('1:2')).toBe(2);expect(f.ctx.nodeValues).toBe(memo);expect(memo.get(sentinel)).toBe(91)
  })
  it('unselected target Pending is never visited',()=>{
    const f=fixture(textFixture(),{col:1,row:1});expect(evaluateFormula('SUMIF(A1:A3,"<>Bananas",B1:B3)',f.ctx)).toBe(40);expect(f.reads.get('1:1')).toBeUndefined()
  })
  it.each([null,'',2,true,formulaError('#N/A')] as const)('unverified non-plain-text criterion cell %s is gated without a guessed negative result',value=>{
    const cells=[cell(0,0,'Apples'),...(value===null?[]:[cell(0,1,value,value===''?'formula':'input')])];const f=fixture(cells)
    expect(evaluateFormula('COUNTIF(A1:A2,"<>Bananas")',f.ctx)).toBe('#NAME?');expect(f.ctx.unsupportedFeatures?.has('conditional-negative-text-input-unverified')).toBe(true)
  })
  it('live mixed-input capability retains cache77, dead branch8 stays silent',()=>{
    const live:XlsxCell={ref:'C1',col:2,row:0,value:77,styleIndex:0,hasCachedValue:true,formula:'COUNTIF(A1:A2,"<>Bananas")'},dead:XlsxCell={ref:'C2',col:2,row:1,value:77,styleIndex:0,hasCachedValue:true,formula:'IF(FALSE,COUNTIF(A1:A2,"<>Bananas"),8)'}
    const doc:XlsxDocument={sheets:[{name:'S',sheetId:'s',workbookIndex:0,cols:[],merges:[],mergeRanges:[],rows:[{index:0,cells:[{ref:'A1',col:0,row:0,value:'Apples',styleIndex:0},live]},{index:1,cells:[{ref:'A2',col:0,row:1,value:2,styleIndex:0},dead]}]}],definedNames:[],tables:[]}
    evaluateWorkbookFormulas(doc,{forceRecalc:true});expect(live.value).toBe(77);expect(dead.value).toBe(8);expect(doc.diagnostics?.some(d=>d.feature==='conditional-negative-text-input-unverified')).toBe(true)
    doc.sheets[0].rows[0].cells=doc.sheets[0].rows[0].cells.filter(c=>c!==live);delete doc.diagnostics;dead.value=77
    evaluateWorkbookFormulas(doc,{forceRecalc:true});expect(dead.value).toBe(8);expect(doc.diagnostics??[]).toEqual([])
  })
  it('supported workbook negative text replaces cache77 with computed2, no capability diagnostic',()=>{
    const result:XlsxCell={ref:'C1',col:2,row:0,value:77,styleIndex:0,hasCachedValue:true,formula:'COUNTIF(A1:A3,"<>Bananas")'}
    const doc:XlsxDocument={sheets:[{name:'S',sheetId:'s',workbookIndex:0,cols:[],merges:[],mergeRanges:[],rows:['Apples','Bananas','Carrots'].map((value,row)=>({index:row,cells:[{ref:`A${row+1}`,col:0,row,value,styleIndex:0},...(row===0?[result]:[])]}))}],definedNames:[],tables:[]};evaluateWorkbookFormulas(doc,{forceRecalc:true});expect(result.value).toBe(2);expect(doc.diagnostics??[]).toEqual([])
  })
})
