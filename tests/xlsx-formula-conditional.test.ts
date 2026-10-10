/** B2 measured scalar/reference observations and explicitly documented controls.
 * Native-reference cases1..11,22..26 and native-families0..17 (TYPE array grammar
 * remains C2). Fixtures are embedded/typed; no external /tmp dependency. */
import { describe, it, expect } from 'vitest'
import { evaluateFormula, evaluateFormulaInternal, formulaError, isEvaluationError } from '../src/xlsx/formula/evaluator'
import { parseFormula } from '../src/xlsx/formula/parser'
import { createReferenceServices, type StoredCell } from '../src/xlsx/formula/refs'
import type { AstNode, EvaluationContext, EvaluationValue, ElementValue } from '../src/xlsx/formula/types'
import { conditionalPredicate } from '../src/xlsx/formula/functions/cond'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import type { XlsxCell, XlsxDocument } from '../src/xlsx/types'

function stored(col:number,row:number,value:ElementValue,origin:StoredCell['origin']='input'):StoredCell{return {address:{sheetId:'s',col,row},value,origin}}
function nativeCells():StoredCell[]{
  const cells=[stored(0,0,'Key'),stored(1,0,'Value'),stored(2,0,'Label'),stored(3,0,'Criterion')]
  const av:ElementValue[]=[1,2,2,'',0,true,formulaError('#N/A')]
  const bv=[10,20,30,0,40,50,60];const cv=['one','two','two-last','empty-text','zero','bool','error']
  for(let i=0;i<7;i++){cells.push(stored(0,i+1,av[i],i===3||i===5||i===6?'formula':'input'),stored(1,i+1,bv[i]),stored(2,i+1,cv[i]))}
  return cells // A9/A10/D2 absent; A5 formula empty remains populated
}
function fixture(cells:StoredCell[]=nativeCells(),pendingAt?:{col:number;row:number}){
  let generation=1;const reads=new Map<string,number>();let visits=0;let pending=!!pendingAt;const marker=new Error('conditional target pending')
  const services=createReferenceServices({generation:()=>generation,sheets:[{sheetId:'s',name:'S',workbookIndex:0}],sheetIdOfName:()=> 's',sheetNameOfId:()=> 'S',definedNames:[],tables:[],store:{
    stored:function*(){for(const c of cells){visits++;yield c}},read:a=>{const key=`${a.col}:${a.row}`;reads.set(key,(reads.get(key)??0)+1);if(pending&&a.col===pendingAt?.col&&a.row===pendingAt.row){pending=false;throw marker}const c=cells.find(c=>c.address.col===a.col&&c.address.row===a.row);return c?{value:c.value,origin:c.origin}:undefined},
  }})
  const ctx:EvaluationContext={typedValues:true,references:services,currentSheet:'S',currentCell:{col:6,row:0,absCol:false,absRow:false},currentAddress:{sheetId:'s',col:6,row:0},flatArgs:new WeakMap(),functionWork:new WeakMap(),getCellValue:(_sheet,col,row)=>cells.find(c=>c.address.col===col&&c.address.row===row)?.value??null}
  return {ctx,services,cells,reads,marker,get visits(){return visits},next:()=>{generation++}}
}
const referenceOracles:ReadonlyArray<readonly[string,number]>=[
 ['SUMIF(A2:A10,2,B2:B10)',50],['SUMIF(A2:A10,">1",B2:B10)',50],['COUNTIF(A2:A10,2)',2],['COUNTIF(A2:A10,0)',1],['COUNTIF(A2:A10,"")',3],['COUNTIFS(A2:A10,D2)',1],['COUNTIFS(A2:A10,"")',3],['AVERAGEIF(A2:A10,2,B2:B10)',25],['SUMIFS(B2:B10,A2:A10,2,B2:B10,">20")',30],['COUNTIFS(A2:A10,2,B2:B10,">20")',1],['COUNTIF(A:A,"")',1048569],['IFS(TRUE,1,TRUE,1/0)',1],['SWITCH(1,1,5,2,1/0)',5],['IFNA(NA(),7)',7],['COUNTIF(C2:C8,"two*")',2],['COUNTIF(A2:A10,"#N/A")',1],
]
const familyOracles:ReadonlyArray<readonly[string,number|boolean|string]>=[
 ['IFNA(1/0,7)','#DIV/0!'],['IFS(FALSE,1,TRUE,2)',2],['IFS(FALSE,1)','#N/A'],['SWITCH(2,1,5,2,7,9)',7],['SWITCH(3,1,5,2,7)','#N/A'],['ISBLANK("")',false],['ISNUMBER(TRUE)',false],['ISTEXT("#N/A")',true],['ISLOGICAL(FALSE)',true],['ISERROR(NA())',true],['ISERR(NA())',false],['ISNA(NA())',true],['TYPE(NA())',16],['ERROR.TYPE(NA())',7],['N("5")',0],['N(TRUE)',1],['N(NA())','#N/A'],
]
describe('B2 native references1..11 and22..26',()=>{it.each(referenceOracles)('%s = %s',(source,expected)=>{const f=fixture();expect(evaluateFormula(source,f.ctx)).toBe(expected)})})
describe('B2 native family scalar predicates/branch controls',()=>{it.each(familyOracles)('%s = %s',(source,expected)=>{expect(evaluateFormula(source,fixture().ctx)).toBe(expected)})})

describe('B2 primary documented geometry and predicates',()=>{
  it.each([['SUMIF(A2:A4,2)',4],['AVERAGEIF(A2:A4,2)',2],['AVERAGEIFS(B2:B4,A2:A4,2,B2:B4,">20")',30],['AVERAGEIF(A2:A4,99,B2:B4)','#DIV/0!'],['AVERAGEIFS(B2:B4,A2:A4,99)','#DIV/0!']] as const)('%s = %s',(source,expected)=>{expect(evaluateFormula(source,fixture().ctx)).toBe(expected)})
  it.each(['SUMIF','AVERAGEIF'])('%s optional target anchors resize to criteria geometry, not declared extent',name=>{const f=fixture();expect(evaluateFormula(`${name}(A2:A4,2,B2:B2)`,f.ctx)).toBe(name==='SUMIF'?50:25)})
  it('SUMIFS boolean targets explicitly convert TRUE1/FALSE0',()=>{const f=fixture([stored(0,0,1),stored(0,1,1),stored(0,2,1),stored(1,0,true),stored(1,1,false),stored(1,2,2)]);expect(evaluateFormula('SUMIFS(B1:B3,A1:A3,1)',f.ctx)).toBe(3)})
  it.each([['ISBLANK(D2)',true],['ISBLANK(A5)',false],['ISNUMBER(A6)',true],['ISNUMBER("5")',false],['ISTEXT(A5)',true],['ISTEXT(A8)',false],['ISLOGICAL(A7)',true],['ISLOGICAL(1)',false],['ISERROR(A8)',true],['ISERROR("#N/A")',false],['ISERR(1/0)',true],['ISNA(1/0)',false],['TYPE(4)',1],['TYPE("")',2],['TYPE(TRUE)',4],['N(FALSE)',0],['N(D2)',0],['N(4)',4]] as const)('per-origin %s = %s',(source,expected)=>{expect(evaluateFormula(source,fixture().ctx)).toBe(expected)})
  it('COUNTIF whole-axis blank count is structural, with sparse reads and formula-empty distinction',()=>{const f=fixture();expect(evaluateFormula('COUNTIF(A:A,"")',f.ctx)).toBe(1048569);expect([...f.reads.values()].reduce((a,b)=>a+b,0)).toBeLessThanOrEqual(16);expect(f.visits).toBeLessThanOrEqual(nativeCells().length)})
})

describe('B2 transactional and lazy application controls',()=>{
  it('pending matching sum value resumes once without adding prior matches again; supplied memo preserved',()=>{
    const f=fixture(nativeCells(),{col:1,row:3});const ast=parseFormula('SUMIF(A2:A10,2,B2:B10)');const memo=new WeakMap<AstNode,EvaluationValue>();const sentinel=parseFormula('42');memo.set(sentinel,42);f.ctx.nodeValues=memo
    let caught:unknown;try{evaluateFormulaInternal(ast,f.ctx)}catch(error){caught=error};expect(caught).toBe(f.marker)
    expect(evaluateFormulaInternal(ast,f.ctx)).toBe(50);expect(f.ctx.nodeValues).toBe(memo);expect(memo.get(sentinel)).toBe(42)
    expect(f.reads.get('1:2')).toBe(1);expect(f.reads.get('1:3')).toBe(2)
  })
  it.each(['IFNA(5,UnknownName)','IFS(TRUE,5,TRUE,UnknownName)','SWITCH(1,1,5,2,UnknownName)'])('%s keeps unchosen unsupported branches silent',source=>{const f=fixture();expect(evaluateFormula(source,f.ctx)).toBe(5);expect(f.ctx.unsupportedFeatures?.size??0).toBe(0)})
  it('workbook dead conditional branch computes8 without a conditional diagnostic',()=>{
    const target:XlsxCell={ref:'B1',col:1,row:0,value:77,styleIndex:0,hasCachedValue:true,formula:'IFS(TRUE,8,TRUE,UnknownName)'}
    const doc:XlsxDocument={sheets:[{name:'S',sheetId:'s',workbookIndex:0,cols:[],merges:[],mergeRanges:[],rows:[{index:0,cells:[target]}]}],definedNames:[],tables:[]};evaluateWorkbookFormulas(doc,{forceRecalc:true});expect(target.value).toBe(8);expect(doc.diagnostics??[]).toEqual([])
  })
  it('typed NA identity differs from error-looking text under predicates and IFNA',()=>{const f=fixture();expect(isEvaluationError(evaluateFormulaInternal(parseFormula('NA()'),f.ctx))).toBe(true);expect(evaluateFormula('IFNA("#N/A",7)',f.ctx)).toBe('#N/A');expect(evaluateFormula('ISNA("#N/A")',f.ctx)).toBe(false)})
})


describe('B2 complete documented operators and status controls',()=>{
  it.each([['<2',1],['<=2',3],['>1',2],['>=2',2],['<>2',1],['=2',2],['2',2]] as const)('COUNTIF numeric-only criterion %s follows its operator', (criterion,expected)=>{const f=fixture([stored(0,0,1),stored(0,1,2),stored(0,2,2)]);expect(evaluateFormula(`COUNTIF(A1:A3,"${criterion}")`,f.ctx)).toBe(expected)})
  it.each([['a~*b',1],['a~?b',1],['a~~b',1],['AL*',2],['a?b',3]] as const)('documented ASCII wildcard %s retains escaping/case', (criterion,expected)=>{const f=fixture(['a*b','a?b','a~b','Alpha','alphabet','other'].map((v,row)=>stored(0,row,v)));expect(evaluateFormula(`COUNTIF(A1:A6,"${criterion}")`,f.ctx)).toBe(expected)})
  it.each([['#NULL!',1],['#DIV/0!',2],['#VALUE!',3],['#REF!',4],['#NAME?',5],['#NUM!',6],['#N/A',7]] as const)('classical ERROR.TYPE %s maps to documented %s',(code,number)=>{expect(evaluateFormula(`ERROR.TYPE(${code})`,fixture().ctx)).toBe(number)})
  it('ERROR.TYPE normal number produces genuine NA and IFNA catches only that error',()=>{expect(evaluateFormula('ERROR.TYPE(1)',fixture().ctx)).toBe('#N/A');expect(evaluateFormula('IFNA(ERROR.TYPE(1),7)',fixture().ctx)).toBe(7)})
  it('SUMIFS shape mismatch is documented VALUE, not a clamped partial sum',()=>{expect(evaluateFormula('SUMIFS(B2:B4,A2:A3,2)',fixture().ctx)).toBe('#VALUE!')})
  it.each([['~x','conditional-wildcard-unsupported-tilde'],['~','conditional-wildcard-unsupported-tilde'],['漢*','conditional-wildcard-unsupported-non-ascii'],['=>2','conditional-criteria-unverified'],['<>foo','conditional-negative-text-input-unverified']] as const)('unsettled criterion %s has an attributable capability status',(criterion,feature)=>{
    const f=fixture();expect(evaluateFormula(`COUNTIF(A2:A10,"${criterion}")`,f.ctx)).toBe('#NAME?');expect(f.ctx.unsupportedFeatures?.has(feature)).toBe(true)
  })
  it('numeric text range matching is gated instead of guessed, including cache77/dead8',()=>{
    const f=fixture([stored(0,0,'2'),stored(0,1,2)]);expect(evaluateFormula('COUNTIF(A1:A2,2)',f.ctx)).toBe('#NAME?');expect(f.ctx.unsupportedFeatures?.has('conditional-numeric-text-unverified')).toBe(true)
    const live:XlsxCell={ref:'B1',col:1,row:0,value:77,styleIndex:0,hasCachedValue:true,formula:'COUNTIF(A1:A2,2)'};const dead:XlsxCell={ref:'B2',col:1,row:1,value:77,styleIndex:0,hasCachedValue:true,formula:'IF(FALSE,COUNTIF(A1:A2,2),8)'}
    const doc:XlsxDocument={sheets:[{name:'S',sheetId:'s',workbookIndex:0,cols:[],merges:[],mergeRanges:[],rows:[{index:0,cells:[{ref:'A1',col:0,row:0,value:'2',styleIndex:0},live]},{index:1,cells:[{ref:'A2',col:0,row:1,value:2,styleIndex:0},dead]}]}],definedNames:[],tables:[]}
    evaluateWorkbookFormulas(doc,{forceRecalc:true});expect(live.value).toBe(77);expect(dead.value).toBe(8);expect(doc.diagnostics?.some(d=>d.feature==='conditional-numeric-text-unverified')).toBe(true)
  })
  it('selected target error has an explicit unsettled profile; unselected target is not propagated',()=>{
    const f=fixture([stored(0,0,1),stored(0,1,2),stored(1,0,formulaError('#DIV/0!')),stored(1,1,20)])
    expect(evaluateFormula('SUMIF(A1:A2,2,B1:B2)',f.ctx)).toBe(20)
    expect(evaluateFormula('SUMIF(A1:A2,1,B1:B2)',f.ctx)).toBe('#NAME?');expect(f.ctx.unsupportedFeatures?.has('conditional-selected-target-error-unverified')).toBe(true)
  })
  it('aggregate private work preserves independent family results and evaluation order',()=>{
    const f=fixture();expect(evaluateFormula('SUM(COUNTIF(A2:A4,2),SUMIF(A2:A4,2),AVERAGEIF(A2:A4,2))',f.ctx)).toBe(8)
    expect(evaluateFormula('SUM(AVERAGEIF(A2:A4,2),SUMIF(A2:A4,2),COUNTIF(A2:A4,2))',f.ctx)).toBe(8)
  })
})


describe('B2 private value adapter and full scalar branch modes',()=>{
  it('measured TYPE matrix64 stays scalar while known predicates classify elementwise (C2)',()=>{
    const matrix={kind:'formula-matrix',rows:1,cols:2,values:[[1,2]]}
    expect(conditionalPredicate('TYPE',matrix,'array')).toBe(64)
    expect(conditionalPredicate('N',matrix,'array')).toEqual({kind:'formula-matrix',rows:1,cols:2,values:[[1,2]]})
    const f=fixture();expect(conditionalPredicate('ISNUMBER',matrix,'array',f.ctx)).toEqual({kind:'formula-matrix',rows:1,cols:2,values:[[true,true]]})
    expect(f.ctx.unsupportedFeatures?.has('conditional-array-unavailable')??false).toBe(false)
  })
  it('modern ERROR.TYPE numbers remain explicitly unverified while classical genuine errors and text differ',()=>{
    const f=fixture();const modern={kind:'formula-error',code:'#SPILL!'}
    expect(conditionalPredicate('ERROR.TYPE',modern,'direct',f.ctx)).toEqual({kind:'formula-error',code:'#NAME?'})
    expect(f.ctx.unsupportedFeatures?.has('conditional-modern-error-unverified')).toBe(true)
    expect(conditionalPredicate('ERROR.TYPE','#N/A','direct')).toEqual({kind:'formula-error',code:'#N/A'})
  })
  it('omitted predicate input is retained as an explicit distinct profile',()=>{
    const f=fixture();expect(evaluateFormula('ISBLANK(,)',f.ctx)).toBe('#VALUE!')
    expect(conditionalPredicate('ISBLANK',null,'omitted',f.ctx)).toEqual({kind:'formula-error',code:'#NAME?'})
    expect(f.ctx.unsupportedFeatures?.has('conditional-omitted-predicate-unverified')).toBe(true)
  })
  it.each(['SUMIF()','SUMIFS(A1)','COUNTIF(A1)','COUNTIFS(A1)','AVERAGEIF(A1)','AVERAGEIFS(A1)','IFNA(1)','IFS(TRUE)','SWITCH(1)','ISBLANK(1,2)','ISNUMBER(1,2)','ISTEXT(1,2)','ISLOGICAL(1,2)','ISERROR(1,2)','ISERR(1,2)','ISNA(1,2)','TYPE(1,2)','ERROR.TYPE(1,2)','N(1,2)'])('wrong-arity %s has a typed signature error, not unknown function capability',source=>{const f=fixture();expect(evaluateFormula(source,f.ctx)).toBe('#VALUE!');expect(f.ctx.unsupportedFeatures?.size??0).toBe(0)})
  it.each([['IFNA(5,1/0)',5],['IFNA(NA(),1/0)','#DIV/0!'],['IFS(FALSE,1/0,TRUE,7)',7],['IFS(TRUE,1/0,TRUE,7)','#DIV/0!'],['IFS("not logical",1)','#VALUE!'],['SWITCH(3,1,1/0,2,1/0,9)',9],['SWITCH(1,1,1/0,9)','#DIV/0!']] as const)('opposite selected/error branch %s preserves its rule', (source,expected)=>{expect(evaluateFormula(source,fixture().ctx)).toBe(expected)})
  it('lazy selected branch Pending remains the exact object and does not evaluate later conditions/results',()=>{
    const f=fixture();const marker=new Error('lazy conditional pending');let first=true;let extraReads=0
    f.ctx.getCellValue=(_sheet,col)=>{if(col===0&&first){first=false;throw marker}if(col===1)extraReads++;return 7}
    const ast=parseFormula('IFS(TRUE,A1,TRUE,B1)');const memo=new WeakMap<AstNode,EvaluationValue>();const sentinel=parseFormula('42');memo.set(sentinel,42);f.ctx.nodeValues=memo
    let caught:unknown;try{evaluateFormulaInternal(ast,f.ctx)}catch(error){caught=error};expect(caught).toBe(marker)
    expect(evaluateFormulaInternal(ast,f.ctx)).toBe(7);expect(extraReads).toBe(0);expect(f.ctx.nodeValues).toBe(memo);expect(memo.get(sentinel)).toBe(42)
  })
})

describe('B2 reference origin, epoch and documented IFNA blank rules',()=>{
  it.each(['IFNA(D2,7)','IFNA(NA(),D2)'])('primary documented empty-cell %s returns empty text, not omitted/zero',source=>{expect(evaluateFormula(source,fixture().ctx)).toBe('')})
  it.each(['SUMIF','AVERAGEIF'])('%s resizing preserves stable cross-sheet IDs through an absolute name target',family=>{
    const cells:StoredCell[]=[{address:{sheetId:'criteria-uuid',col:0,row:0},value:1,origin:'input'},{address:{sheetId:'criteria-uuid',col:0,row:1},value:2,origin:'input'},{address:{sheetId:'criteria-uuid',col:0,row:2},value:2,origin:'input'},...[10,20,30].map((value,row):StoredCell=>({address:{sheetId:'values-uuid',col:1,row},value,origin:'input'}))]
    const sheets=[{sheetId:'criteria-uuid',name:'Criteria',workbookIndex:0},{sheetId:'values-uuid',name:'Values',workbookIndex:1}]
    const services=createReferenceServices({generation:()=>1,sheets,definedNames:[{name:'Target',source:'Values!$B$1',baseProvenance:'unknown'}],tables:[],sheetIdOfName:name=>sheets.find(s=>s.name===name)?.sheetId,sheetNameOfId:id=>sheets.find(s=>s.sheetId===id)?.name,store:{stored:()=>cells,read:address=>{const c=cells.find(c=>c.address.sheetId===address.sheetId&&c.address.col===address.col&&c.address.row===address.row);return c?{value:c.value,origin:c.origin}:undefined}}})
    const ctx:EvaluationContext={references:services,currentSheet:'Criteria',flatArgs:new WeakMap(),functionWork:new WeakMap()}
    expect(evaluateFormula(`${family}(Criteria!A1:A3,2,Target)`,ctx)).toBe(family==='SUMIF'?50:25)
  })
  it('two blank-criteria ranges merge sparse positions and structural absents without scanning the axis',()=>{
    const f=fixture([stored(0,0,'Header'),stored(1,0,'Header'),stored(0,1,'' ,'formula'),stored(1,2,9)])
    expect(evaluateFormula('COUNTIFS(A:A,"",B:B,"")',f.ctx)).toBe(1048574)
    expect([...f.reads.values()].reduce((a,b)=>a+b,0)).toBeLessThanOrEqual(12)
  })
  it('changed generation after pending sum target drops the full conditional accumulator, preserving provided memo',()=>{
    const f=fixture([stored(0,0,2),stored(0,1,2),stored(0,2,2),stored(1,0,10),stored(1,1,20),stored(1,2,30)],{col:1,row:1})
    const ast=parseFormula('SUMIF(A1:A3,2,B1:B3)');const memo=new WeakMap<AstNode,EvaluationValue>();const sentinel=parseFormula('42');memo.set(sentinel,42);f.ctx.nodeValues=memo
    let caught:unknown;try{evaluateFormulaInternal(ast,f.ctx)}catch(error){caught=error};expect(caught).toBe(f.marker)
    for(const c of f.cells)if(c.address.col===1)c.value=100*(c.address.row+1)
    f.next();expect(evaluateFormulaInternal(ast,f.ctx)).toBe(600)
    expect(f.ctx.nodeValues).toBe(memo);expect(memo.get(sentinel)).toBe(42)
  })
  it('non-ASCII wildcard capability propagates through predicates and retains valid cache, not a fake native NAME test',()=>{
    const f=fixture([stored(0,0,'漢')]);expect(evaluateFormula('ISERROR(COUNTIF(A1,"*"))',f.ctx)).toBe('#NAME?')
    expect(f.ctx.unsupportedFeatures?.has('conditional-wildcard-unsupported-non-ascii')).toBe(true)
  })
  it('standalone whole-axis conditional aggregation explicitly reports unavailable sparse services',()=>{
    let reads=0;const ctx:EvaluationContext={getCellValue:()=>{reads++;return null}}
    expect(evaluateFormula('COUNTIF(A:A,"")',ctx)).toBe('#NAME?');expect(ctx.unsupportedFeatures?.has('conditional-reference-unavailable')).toBe(true);expect(reads).toBe(0)
  })
})
