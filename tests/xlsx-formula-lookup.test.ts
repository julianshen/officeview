/** C1 finite reference lookups. Native references12..21 embedded verbatim.
 * Family array observations are materialized as reference fixtures: only the
 * shared lookup algorithm is tested here, not literal grammar/array publication.
 * Primary mode definitions: support.microsoft.com/en-us/excel/functions/{index,
 * match,vlookup,hlookup,xlookup,xmatch}-function. Binary requires sorted inputs. */
import { describe, it, expect } from 'vitest'
import { evaluateFormula, evaluateFormulaInternal, formulaError, isEvaluationError } from '../src/xlsx/formula/evaluator'
import { indexSelection, xlookupSelection } from '../src/xlsx/formula/functions/lookup'
import { parseFormula } from '../src/xlsx/formula/parser'
import { createReferenceServices, isReferenceNode, type StoredCell } from '../src/xlsx/formula/refs'
import type { AstNode, ElementValue, EvaluationContext, EvaluationValue } from '../src/xlsx/formula/types'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import type { XlsxCell, XlsxDocument } from '../src/xlsx/types'
function cell(col:number,row:number,value:ElementValue,origin:StoredCell['origin']='input'):StoredCell{return {address:{sheetId:'s',col,row},value,origin}}
function fixture(cells:StoredCell[]=nativeCells(),pendingAt?:{col:number;row:number}) {
 let generation=1,pending=!!pendingAt,visits=0;const marker=new Error('lookup pending'),reads=new Map<string,number>();const issues:string[]=[]
 const services=createReferenceServices({generation:()=>generation,sheets:[{sheetId:'s',name:'S',workbookIndex:0}],sheetIdOfName:n=>n==='S'?'s':undefined,sheetNameOfId:()=> 'S',definedNames:[{name:'LookupKeys',source:'$A$1:$A$3',baseProvenance:'unknown'}],tables:[],store:{stored:function*(){for(const c of cells){visits++;yield c}},read:a=>{const key=`${a.col}:${a.row}`;reads.set(key,(reads.get(key)??0)+1);if(pending&&a.col===pendingAt?.col&&a.row===pendingAt.row){pending=false;throw marker}const c=cells.find(c=>c.address.col===a.col&&c.address.row===a.row);return c?{value:c.value,origin:c.origin}:undefined}}})
 const ctx:EvaluationContext={typedValues:true,references:services,currentSheet:'S',currentAddress:{sheetId:'s',col:6,row:0},currentCell:{col:6,row:0,absCol:false,absRow:false},flatArgs:new WeakMap(),functionWork:new WeakMap(),getCellValue:(_s,col,row)=>cells.find(c=>c.address.col===col&&c.address.row===row)?.value??null,reportFormulaIssue:i=>issues.push(i.feature??i.kind)}
 return {ctx,services,cells,reads,issues,marker,next:()=>generation++,get visits(){return visits}}
}
function nativeCells():StoredCell[]{const cells=[cell(0,0,'Key'),cell(1,0,'Value'),cell(2,0,'Label'),cell(3,0,'Criterion')];const a:ElementValue[]=[1,2,2,'',0,true,formulaError('#N/A')],b=[10,20,30,0,40,50,60],c=['one','two','two-last','empty-text','zero','bool','error'];for(let i=0;i<7;i++)cells.push(cell(0,i+1,a[i],[3,5,6].includes(i)?'formula':'input'),cell(1,i+1,b[i]),cell(2,i+1,c[i]));return cells}
function vector(values:ElementValue[],horizontal=false):StoredCell[]{return values.flatMap((value,i)=>horizontal?[cell(i,0,value),cell(i,1,(i+1)*10)]:[cell(0,i,value),cell(1,i,(i+1)*10)])}
const nativeReference:ReadonlyArray<readonly[string,number|string]>=[['INDEX(A:A,500000)',0],['MATCH(2,A2:A4,0)',2],['MATCH(2,A2:A4,1)',3],['VLOOKUP(2,A2:C4,3,FALSE)','two'],['XLOOKUP(2,A2:A4,C2:C4)','two'],['XLOOKUP(2,A2:A4,C2:C4,,0,-1)','two-last'],['XMATCH(2,A2:A4,0,-1)',3],['XLOOKUP(1.5,A2:A4,C2:C4,,1)','two'],['XLOOKUP(1.5,A2:A4,C2:C4,,-1)','one'],['XLOOKUP(2,A2:A4,C2:C4,1/0)','two']]
describe('C1 measured native reference12..21',()=>it.each(nativeReference)('%s = %s',(formula,expected)=>expect(evaluateFormula(formula,fixture().ctx)).toBe(expected)))
describe('C1 family algorithm controls and documented finite modes',()=>{
 it('measured family descending MATCH duplicates choose2',()=>expect(evaluateFormula('MATCH(2,A1:A4,-1)',fixture(vector([3,2,2,1])).ctx)).toBe(2))
 it('measured family MATCH wildcard first2',()=>expect(evaluateFormula('MATCH("t*",A1:A3,0)',fixture(vector(['one','two','three'])).ctx)).toBe(2))
 it.each([[2,[1,2,2,3],2,20],[-2,[3,2,2,1],3,30]] as const)('family sorted binary%s measured duplicate midpoint',(search,values,expectedIndex,expectedReturn)=>{const f=fixture(vector([...values]));expect(evaluateFormula(`XMATCH(2,A1:A4,0,${search})`,f.ctx)).toBe(expectedIndex);expect(evaluateFormula(`XLOOKUP(2,A1:A4,B1:B4,,0,${search})`,f.ctx)).toBe(expectedReturn)})
 it('measured family XLOOKUP wildcard',()=>expect(evaluateFormula('XLOOKUP("t*",A1:A3,B1:B3,,2)',fixture(vector(['one','two','three'])).ctx)).toBe(20))
 it('measured family HLOOKUP exact2',()=>expect(evaluateFormula('HLOOKUP(2,A1:C2,2,FALSE)',fixture(vector([1,2,3],true)).ctx)).toBe(20))
 it.each([-1,1,2,-2])('XMATCH search%s supports documented exact/nearest modes on unique sorted numbers',search=>{
  const values=search===-2?[7,5,3,1]:[1,3,5,7],f=fixture(vector(values));const index=(value:number)=>values.indexOf(value)+1
  expect(evaluateFormula(`XMATCH(5,A1:A4,0,${search})`,f.ctx)).toBe(index(5));expect(evaluateFormula(`XMATCH(4,A1:A4,-1,${search})`,f.ctx)).toBe(index(3));expect(evaluateFormula(`XMATCH(4,A1:A4,1,${search})`,f.ctx)).toBe(index(5))
  expect(evaluateFormula(`XLOOKUP(4,A1:A4,B1:B4,,-1,${search})`,fixture(vector(values)).ctx)).toBe(index(3)*10);expect(evaluateFormula(`XLOOKUP(4,A1:A4,B1:B4,,1,${search})`,fixture(vector(values)).ctx)).toBe(index(5)*10)
 })
 it.each([1,3,8,17])('true sorted binary unique length%s has logarithmic physical reads',length=>{
  const values=Array.from({length},(_,i)=>i*2+1);for(const search of [2,-2]){const ordered=search===2?values:[...values].reverse(),f=fixture(vector(ordered)),wanted=values[Math.floor(length/2)];expect(evaluateFormula(`XMATCH(${wanted},A1:A${length},0,${search})`,f.ctx)).toBe(ordered.indexOf(wanted)+1);expect([...f.reads.values()].reduce((a,b)=>a+b,0)).toBeLessThanOrEqual(Math.ceil(Math.log2(length+1))+1)}
 })
 it.each([['MATCH(4,A1:A4,1)',2],['MATCH(4,A1:A4,-1)',2],['VLOOKUP(4,A1:B4,2,TRUE)',20],['VLOOKUP(4,A1:B4,2)',20],['HLOOKUP(4,A1:D2,2,TRUE)',20]] as const)('documented approximate %s',(formula,expected)=>{const horizontal=formula.startsWith('H'),values=formula.includes(',-1)')?[7,5,3,1]:[1,3,5,7];expect(evaluateFormula(formula,fixture(vector(values,horizontal)).ctx)).toBe(expected)})
 it.each(['MATCH(99,A1:A3,0)','XMATCH(99,A1:A3)','VLOOKUP(99,A1:B3,2,FALSE)','HLOOKUP(99,A1:C2,2,FALSE)','XLOOKUP(99,A1:A3,B1:B3)'])('documented missing %s is genuineNA',formula=>{const f=fixture(vector([1,2,3],formula.startsWith('H')));const result=evaluateFormulaInternal(parseFormula(formula),f.ctx);expect(isEvaluationError(result)).toBe(true);expect(evaluateFormula(formula,f.ctx)).toBe('#N/A')})
 it.each(['MATCH("AL*",A1:A4,0)','XMATCH("AL*",A1:A4,2)','XLOOKUP("AL*",A1:A4,B1:B4,,2)','VLOOKUP("AL*",A1:B4,2,FALSE)'])('ASCII wildcard/case %s',(formula)=>expect(evaluateFormula(formula,fixture(vector(['Beta','Alpha','alphabet','other'])).ctx)).toBe(formula.startsWith('MATCH')||formula.startsWith('XMATCH')?2:20))
 it.each(['a~*b','a~?b','a~~b'])('MATCH escaped%s uses acceptedASCII helper',pattern=>{const values=['a*b','a?b','a~b'];expect(evaluateFormula(`MATCH("${pattern}",A1:A3,0)`,fixture(vector(values)).ctx)).toBe(pattern==='a~*b'?1:pattern==='a~?b'?2:3)})
 it('XMATCH exact0 treats star literally; wildcard2 separately matchespattern',()=>{const f=fixture(vector(['tree','t*','two']));expect(evaluateFormula('XMATCH("t*",A1:A3,0)',f.ctx)).toBe(2);expect(evaluateFormula('XMATCH("t*",A1:A3,2)',f.ctx)).toBe(1)})
 it('INDEX fullphysical farrow performs one addressedread and no millioncellmaterialization',()=>{const f=fixture([cell(0,499999,42)]);expect(evaluateFormula('INDEX(A:A,500000)',f.ctx)).toBe(42);expect(f.reads.get('0:499999')).toBe(1);expect([...f.reads.values()].reduce((a,b)=>a+b,0)).toBe(1);expect(f.visits).toBeLessThanOrEqual(1)})
 it.each([['INDEX(A1:B2,2,2)',4],['INDEX(A1:B2,0,2)',2],['INDEX(A1:B2,2,0)',3],['INDEX(A1:B2,3,1)','#REF!'],['INDEX(A1:B2,-1,1)','#VALUE!'],['INDEX(A1:B2,1,)','#NAME?']] as const)('INDEX coordinate/zeroaxis/explicitomission%s',(formula,expected)=>expect(evaluateFormula(formula,fixture([cell(0,0,1),cell(1,0,2),cell(0,1,3),cell(1,1,4)]).ctx)).toBe(expected))
 it.each(['INDEX()','MATCH()','VLOOKUP()','HLOOKUP()','XLOOKUP()','XMATCH()'])('wrongarity%s hasVALUE notunknowncapability',formula=>expect(evaluateFormula(formula,fixture().ctx)).toBe('#VALUE!'))
 it.each(['XMATCH(2,A1:A3,3)','XMATCH(2,A1:A3,0,0)','XLOOKUP(2,A1:A3,B1:B3,,3)','XLOOKUP(2,A1:A3,B1:B3,,0,0)','MATCH(2,A1:A3,2)'])('invalid documentedmode%s',formula=>expect(evaluateFormula(formula,fixture(vector([1,2,3])).ctx)).toBe('#VALUE!'))
 it('INDEX genuine selectederror stays tagged while textlookalike stays text',()=>{const f=fixture([cell(0,0,formulaError('#N/A')),cell(0,1,'#N/A')]);expect(evaluateFormulaInternal(parseFormula('INDEX(A1:A2,1)'),f.ctx)).toEqual(formulaError('#N/A'));expect(evaluateFormulaInternal(parseFormula('INDEX(A1:A2,2)'),f.ctx)).toBe('#N/A')})
 it('lookupsource absolutealias preservesgeometry',()=>expect(evaluateFormula('XMATCH(2,LookupKeys)',fixture(vector([1,2,3])).ctx)).toBe(2))
})
describe('C1 lazy/Pending/generation and explicit profile boundaries',()=>{
 it.each([['XLOOKUP(2,A1:A3,B1:B3,1/0)',20],['XLOOKUP(99,A1:A3,B1:B3,7)',7],['XLOOKUP(99,A1:A3,B1:B3,1/0)','#DIV/0!']] as const)('notFound evaluatesonlydocumentedmiss%s',(formula,expected)=>expect(evaluateFormula(formula,fixture(vector([1,2,3])).ctx)).toBe(expected))
 it('selectedreturnPending preserves earliersearchreads and caller memo; unselected returns untouched',()=>{
  const f=fixture(vector([1,2,3]),{col:1,row:1}),ast=parseFormula('XLOOKUP(2,A1:A3,B1:B3,UnknownName)'),memo=new WeakMap<AstNode,EvaluationValue>(),sentinel=parseFormula('91');memo.set(sentinel,91);f.ctx.nodeValues=memo;let caught:unknown;try{evaluateFormulaInternal(ast,f.ctx)}catch(e){caught=e};expect(caught).toBe(f.marker);expect(f.reads.get('1:0')).toBeUndefined();expect(f.reads.get('1:2')).toBeUndefined();expect(evaluateFormulaInternal(ast,f.ctx)).toBe(20);expect(f.reads.get('0:0')).toBe(1);expect(f.reads.get('0:1')).toBe(1);expect(f.reads.get('1:1')).toBe(2);expect(f.ctx.nodeValues).toBe(memo);expect(memo.get(sentinel)).toBe(91);expect(f.ctx.unsupportedFeatures?.size??0).toBe(0)
 })
 it('missfallbackPending retains exactobject and resumes fallback only',()=>{
  const f=fixture(vector([1,2,3])),marker={pending:'fallback'};let first=true;f.ctx.getCellValue=()=>{if(first){first=false;throw marker}return 8};const ast=parseFormula('XLOOKUP(99,A1:A3,B1:B3,D1)');let caught:unknown;try{evaluateFormulaInternal(ast,f.ctx)}catch(e){caught=e};expect(caught).toBe(marker);expect(evaluateFormulaInternal(ast,f.ctx)).toBe(8);expect(f.reads.get('0:0')).toBe(1)
 })
 it('genuine generationchange resets fullsearch including owned lookupvalue memo',()=>{
  const f=fixture([...vector([1,2,3]),cell(3,0,2)],{col:1,row:1}),ast=parseFormula('XLOOKUP(D1+0,A1:A3,B1:B3)');let caught:unknown;try{evaluateFormulaInternal(ast,f.ctx)}catch(e){caught=e};expect(caught).toBe(f.marker);for(const c of f.cells)if(c.address.col===1)c.value=100*(c.address.row+1);else if(c.address.col===3)c.value=3;f.next();expect(evaluateFormulaInternal(ast,f.ctx)).toBe(300);expect(f.ctx.nodeValues).toBeUndefined()
 })
 it.each([['XMATCH(2,A1:A2)',[cell(0,0,1)],'lookup-blank-vector-unverified'],['XMATCH(2,A1:A2)',[cell(0,0,'2'),cell(0,1,2)],'lookup-numeric-text-unverified'],['XMATCH(2,A1:A2)',[cell(0,0,formulaError('#N/A')),cell(0,1,2)],'lookup-vector-error-unverified'],['XMATCH("漢*",A1:A2,2)',vector(['漢','other']),'lookup-wildcard-unsupported-non-ascii'],['XMATCH("t*",A1:A2,2,2)',vector(['one','two']),'lookup-binary-wildcard-unverified']] as const)('unverifiedprofile%s staysspecific',(formula,cells,feature)=>{const f=fixture([...cells]);expect(evaluateFormula(formula,f.ctx)).toBe('#NAME?');expect(f.ctx.unsupportedFeatures?.has(feature)).toBe(true)})
 it('unknownprofilecache77 and deadbranch8 with attributablelive issue only',()=>{
  const live:XlsxCell={ref:'C1',col:2,row:0,value:77,styleIndex:0,hasCachedValue:true,formula:'XMATCH(2,A1:A2)'},dead:XlsxCell={ref:'C2',col:2,row:1,value:77,styleIndex:0,hasCachedValue:true,formula:'IF(FALSE,XMATCH(2,A1:A2),8)'},doc:XlsxDocument={sheets:[{name:'S',sheetId:'s',workbookIndex:0,cols:[],merges:[],mergeRanges:[],rows:[{index:0,cells:[{ref:'A1',col:0,row:0,value:'2',styleIndex:0},live]},{index:1,cells:[{ref:'A2',col:0,row:1,value:2,styleIndex:0},dead]}]}],definedNames:[],tables:[]};evaluateWorkbookFormulas(doc,{forceRecalc:true});expect(live.value).toBe(77);expect(dead.value).toBe(8);expect(doc.diagnostics?.some(d=>d.feature==='lookup-numeric-text-unverified')).toBe(true)
 })
})

// Staged private matrices: genuine dimensions/values, no public C2 array claim.
describe('C1 private selection shape boundary',()=>{
 function ref(f:ReturnType<typeof fixture>,source:string){const node=parseFormula(source);if(!isReferenceNode(node))throw new Error('non-reference fixture');const r=f.services.resolve(node,f.ctx);if(!r||isEvaluationError(r))throw new Error('invalid reference fixture');return r}
 it.each([[0,2,2,1,[[2],[4]]],[2,0,1,2,[[3,4]]],[0,0,2,2,[[1,2],[3,4]]]] as const)('INDEX zero row%s/col%s retains shape',(row,col,rows,cols,values)=>{const f=fixture([cell(0,0,1),cell(1,0,2),cell(0,1,3),cell(1,1,4)]);expect(indexSelection(ref(f,'A1:B2'),row,col,0,f.ctx)).toEqual({kind:'formula-matrix',rows,cols,values})})
 it('INDEX selected union area is positional and keeps duplicates/true geometry',()=>{const f=fixture([cell(0,0,10),cell(1,0,20)]);expect(indexSelection(ref(f,'(A1,B1,A1)'),1,1,1,f.ctx)).toBe(20);expect(indexSelection(ref(f,'(A1,B1,A1)'),1,1,2,f.ctx)).toBe(10)})
 it('XLOOKUP vertical keys preserve entire selected row',()=>{const f=fixture([cell(1,0,10),cell(2,0,11),cell(1,1,20),cell(2,1,21)]);expect(xlookupSelection(ref(f,'B1:C2'),1,false,f.ctx)).toEqual({kind:'formula-matrix',rows:1,cols:2,values:[[20,21]]})})
 it('XLOOKUP horizontal keys preserve entire selected column',()=>{const f=fixture([cell(0,1,10),cell(1,1,20),cell(0,2,11),cell(1,2,21)]);expect(xlookupSelection(ref(f,'A2:B3'),1,true,f.ctx)).toEqual({kind:'formula-matrix',rows:2,cols:1,values:[[20],[21]]})})
 it('fullaxis materialization resource limit is specific; positional farrow remains separate',()=>{const f=fixture([]);f.ctx.maxArrayCells=100000;expect(indexSelection(ref(f,'A:A'),0,1,0,f.ctx)).toEqual(formulaError('#NAME?'));expect(f.ctx.unsupportedFeatures?.has('lookup-matrix-resource-limit')).toBe(true);expect([...f.reads.values()].reduce((a,b)=>a+b,0)).toBe(0)})
})

describe('C1 exact lookup key-axis scope',()=>{
 it('VLOOKUP unrelated return blanks do not classify fully populated key column as blank',()=>{const f=fixture([cell(0,0,1),cell(0,1,2),cell(0,2,3),cell(1,1,20)]);expect(evaluateFormula('VLOOKUP(2,A1:B3,2,FALSE)',f.ctx)).toBe(20);expect(f.ctx.unsupportedFeatures?.size??0).toBe(0);expect(f.reads.get('1:0')).toBeUndefined();expect(f.reads.get('1:2')).toBeUndefined()})
 it('HLOOKUP unrelated return blanks do not classify fully populated key row as blank',()=>{const f=fixture([cell(0,0,1),cell(1,0,2),cell(2,0,3),cell(1,1,20)]);expect(evaluateFormula('HLOOKUP(2,A1:C2,2,FALSE)',f.ctx)).toBe(20);expect(f.ctx.unsupportedFeatures?.size??0).toBe(0);expect(f.reads.get('0:1')).toBeUndefined();expect(f.reads.get('2:1')).toBeUndefined()})
})
