/** Finite reference lookups. Selection matrices are private until C2 array API. */
import type { FunctionHandler, EvaluatorFn } from '../functions'
import type { AstNode, ElementValue, EvaluationContext, EvaluationError, EvaluationValue, MatrixValue, RefArea, ResolvedRef } from '../types'
import { coerceToBoolean, coerceToNumber, formulaError, isEvaluationError, isMatrixValue, markUnsupported, matrixOf, scalarProjection } from '../evaluator'
import { isReferenceNode } from '../refs'
import { matchWildcard } from './wildcard'
import { currentReferenceGeneration, isReferenceGenerationChanged, referenceGenerationChanged, ownedMemoEpoch, registerReferenceCreator, requestOwnedMemoRestart, createReferenceRegion, MAX_REFERENCE_GENERATION_RESTARTS } from '../reference-generation'

export interface StructuralLookupMatrix { readonly kind:'formula-matrix'; readonly rows:number; readonly cols:number; readonly values:readonly (readonly ElementValue[])[] }
export type LookupSelection = EvaluationValue | StructuralLookupMatrix
interface SelectionProgress { prepared?:boolean; next:number; values:ElementValue[][] }
const DEFAULT_MATRIX_CELL_BUDGET=1048576 // approved C2 materialization policy default, not Excel dimensions
function unavailable(ctx:EvaluationContext|undefined,feature:string,message:string):EvaluationError {
 markUnsupported(ctx,feature);ctx?.reportFormulaIssue?.({kind:'unsupported-reference',feature,message});return formulaError('#NAME?')
}
function isMatrix(value:LookupSelection):value is StructuralLookupMatrix{return typeof value==='object'&&value!==null&&value.kind==='formula-matrix'}
function selection(ref:ResolvedRef,area:number,row:number,col:number,rows:number,cols:number,ctx:EvaluationContext|undefined,progress:SelectionProgress):LookupSelection {
 const source=ref.areas[area]
 if(!source||![area,row,col,rows,cols].every(Number.isSafeInteger)||row<0||col<0||rows<1||cols<1||row+rows>source.rows||col+cols>source.cols)return formulaError('#REF!')
 if(rows*cols>(ctx?.maxArrayCells??DEFAULT_MATRIX_CELL_BUDGET))return unavailable(ctx,'lookup-matrix-resource-limit',`Selection exceeds the ${ctx?.maxArrayCells??DEFAULT_MATRIX_CELL_BUDGET}-cell materialization policy; positional reads remain available`)
 if(!progress.prepared){if(!privateMatrixRefs.has(ref))ctx?.references?.prepare(ref,ctx);progress.prepared=true}
 while(progress.next<rows*cols){const r=Math.floor(progress.next/cols),c=progress.next%cols;const value=ref.readAt(area,row+r,col+c);(progress.values[r]??=[])[c]=value===null?0:value;progress.next++}
 return rows===1&&cols===1?progress.values[0][0]:{kind:'formula-matrix',rows,cols,values:progress.values}
}
/** Zero axes select full dimensions. Omission is handled separately by caller. */
export function indexSelection(ref:ResolvedRef,row:number,col:number,area:number,ctx:EvaluationContext|undefined,progress:SelectionProgress={next:0,values:[]}):LookupSelection {
 const a=ref.areas[area];if(!a||!Number.isSafeInteger(row)||!Number.isSafeInteger(col)||row<0||col<0)return formulaError('#REF!')
 return selection(ref,area,row===0?0:row-1,col===0?0:col-1,row===0?a.rows:1,col===0?a.cols:1,ctx,progress)
}
export function xlookupSelection(ref:ResolvedRef,index:number,horizontal:boolean,ctx:EvaluationContext|undefined,progress:SelectionProgress={next:0,values:[]}):LookupSelection {
 const a=ref.areas[0];if(!a||ref.areas.length!==1)return formulaError('#VALUE!')
 return selection(ref,0,horizontal?0:index,horizontal?index:0,horizontal?a.rows:1,horizontal?1:a.cols,ctx,progress)
}

/** Integer positional argument for the array-constant INDEX path. */
function arrayIndexInteger(arg:AstNode|undefined,ctx:EvaluationContext|undefined,evalNode:EvaluatorFn):number|EvaluationError {
 if(!arg)return formulaError('#VALUE!')
 const value=evalNode(arg,ctx);if(isEvaluationError(value))return value
 const num=coerceToNumber(scalarProjection(value));if(isEvaluationError(num))return num
 return Number.isSafeInteger(num)?num:unavailable(ctx,'lookup-integer-argument-unverified','Noninteger lookup positions/modes have no accepted profile')
}

/**
 * C1 carried control: INDEX over a literal/array value (INDEX({10;20;30},2)=20)
 * rather than a reference. Zero axes select the full dimension and the
 * one-row/one-column omission rule is preserved; matrices stay rectangular.
 */
function indexArrayValue(args:AstNode[],ctx:EvaluationContext|undefined,evalNode:EvaluatorFn):EvaluationValue {
 const source=evalNode(args[0],ctx);if(isEvaluationError(source))return source
 const matrix:MatrixValue=isMatrixValue(source)?source:matrixOf(1,1,[[source]])
 if(args[1]?.type==='empty'||args[2]?.type==='empty')return unavailable(ctx,'lookup-index-omission-unverified','Explicit INDEX omission is distinct from zero-axis selection')
 let row=arrayIndexInteger(args[1],ctx,evalNode);if(isEvaluationError(row))return row
 if(row<0)return formulaError('#VALUE!')
 let col:number|EvaluationError
 if(args[2])col=arrayIndexInteger(args[2],ctx,evalNode)
 else if(matrix.cols===1)col=1
 else if(matrix.rows===1){col=row;row=1}
 else col=0 // documented array form: row_num only on a 2D array returns the whole selected row
 if(isEvaluationError(col))return col
 if(col<0)return formulaError('#VALUE!')
 const rows=row===0?matrix.rows:1,cols=col===0?matrix.cols:1
 const startRow=row===0?0:row-1,startCol=col===0?0:col-1
 if(startRow<0||startCol<0||startRow+rows>matrix.rows||startCol+cols>matrix.cols)return formulaError('#REF!')
 if(rows===1&&cols===1)return matrix.values[startRow][startCol]
 const values:ElementValue[][]=[]
 for(let r=0;r<rows;r++){const out:ElementValue[]=[];for(let c=0;c<cols;c++)out.push(matrix.values[startRow+r][startCol+c]);values.push(out)}
 return matrixOf(rows,cols,values)
}

// Completed reference/matrix projection provenance is scoped to the active memo.
// Flattening returned axes/reference results needs C2 modes, never scalar reuse.
const projectedByMemo=new WeakMap<WeakMap<AstNode,EvaluationValue>,WeakSet<AstNode[]>>()
function recordProjection(args:AstNode[],ctx:EvaluationContext|undefined):void {const memo=ctx?.nodeValues;if(!memo)return;let set=projectedByMemo.get(memo);if(!set){set=new WeakSet();projectedByMemo.set(memo,set)}set.add(args)}
export function lookupProjectionIn(root:AstNode,ctx:EvaluationContext|undefined):boolean {
 const memo=ctx?.nodeValues,set=memo?projectedByMemo.get(memo):undefined;if(!memo||!set)return false
 const stack=[root],seen=new Set<AstNode>()
 while(stack.length){const n=stack.pop()!;if(seen.has(n)||!memo.has(n))continue;seen.add(n)
  if(n.type==='call'){if(set.has(n.args))return true;stack.push(...n.args)}
  else if(n.type==='unary'||n.type==='implicitIntersect')stack.push(n.expr)
  else if(n.type==='binary')stack.push(n.left,n.right)
 }
 return false
}

type Family='INDEX'|'MATCH'|'VLOOKUP'|'HLOOKUP'|'XLOOKUP'|'XMATCH'
type MatchMode=-1|0|1|2
type SearchMode=-2|-1|1|2
const LOOKUP=Symbol('lookup progress')
interface View {ref:ResolvedRef;area:RefArea;prepared:boolean}
interface State {
 tag:typeof LOOKUP;family:Family;generation?:number;owned:ReturnType<typeof ownedMemoEpoch>;restarts:number
 scalars:Map<number,EvaluationValue>;views:Map<number,View>;configured:boolean;lookup?:EvaluationValue
 key?:View;target?:View;horizontal:boolean;length:number;mode:MatchMode;search:SearchMode;wildcard:boolean
 next:number;lo:number;hi:number;best?:number;bestValue?:ElementValue;found?:number;searched:boolean
 keys:Map<number,ElementValue>;selection:SelectionProgress;indexRow:number;indexCol:number;indexArea:number
 complete:boolean;result?:EvaluationValue;projected:boolean
}
const fallbackWork=new WeakMap<EvaluationContext,WeakMap<AstNode[],State>>()
// No EvaluationContext at all (public evaluateFormula without services): a
// module-level store keeps the same frame semantics without inventing a ctx.
const anonymousWork=new WeakMap<AstNode[],State>()
function stateAt(args:AstNode[],ctx:EvaluationContext|undefined):State|undefined {const v=ctx?(ctx.functionWork?ctx.functionWork.get(args):fallbackWork.get(ctx)?.get(args)):anonymousWork.get(args);return typeof v==='object'&&v!==null&&'tag'in v&&v.tag===LOOKUP?v as State:undefined}
function save(args:AstNode[],ctx:EvaluationContext|undefined,state:State):void {if(ctx?.functionWork)ctx.functionWork.set(args,state);else if(ctx){let m=fallbackWork.get(ctx);if(!m){m=new WeakMap();fallbackWork.set(ctx,m)}m.set(args,state)}else anonymousWork.set(args,state)}
function fresh(family:Family,ctx:EvaluationContext|undefined,restarts=0):State{return {tag:LOOKUP,family,generation:ctx?.references?currentReferenceGeneration(ctx.references):undefined,owned:ownedMemoEpoch(ctx),restarts,scalars:new Map(),views:new Map(),configured:false,horizontal:false,length:0,mode:0,search:1,wildcard:false,next:0,lo:0,hi:-1,searched:false,keys:new Map(),selection:{next:0,values:[]},indexRow:1,indexCol:1,indexArea:0,complete:false,projected:false}}
function argValue(i:number,args:AstNode[],state:State,ctx:EvaluationContext|undefined,evalNode:EvaluatorFn,defaultValue?:EvaluationValue):EvaluationValue {
 if(state.scalars.has(i))return state.scalars.get(i)!
 const n=args[i];const value=!n||n.type==='empty'?defaultValue===undefined?null:defaultValue:evalNode(n,ctx);state.scalars.set(i,value);return value
}
function numeric(i:number,args:AstNode[],state:State,ctx:EvaluationContext|undefined,evalNode:EvaluatorFn,defaultValue?:number):number|EvaluationError {
 const v=argValue(i,args,state,ctx,evalNode,defaultValue);if(isEvaluationError(v))return v;const n=coerceToNumber(v);if(isEvaluationError(n))return n;return Number.isSafeInteger(n)?n:unavailable(ctx,'lookup-integer-argument-unverified','Noninteger lookup positions/modes have no accepted profile')
}

/**
 * Private matrix/rectangle view: presents a literal/value matrix as a
 * ResolvedRef so the existing finite lookup algorithms run unchanged. No fake
 * workbook sheet identity is invented (the area sheetId is the empty string)
 * and no public ABI is added. Sub-regions (V/HLOOKUP key axis) are produced by
 * the registered private region creator.
 *
 * Membership of `privateMatrixRefs` (this ref plus every creator sub-region) is
 * the ONLY signal that a ref is in-memory: such a ref must never invoke the
 * physical workbook preparation/discovery hook. Genuine references are never
 * added, and privacy is never inferred from an empty sheetId.
 */
const privateMatrixRefs=new WeakSet<ResolvedRef>()
function createMatrixRef(matrix:MatrixValue):ResolvedRef {
 const area:RefArea={sheetId:'',firstCol:0,firstRow:0,cols:matrix.cols,rows:matrix.rows}
 const populated:Array<{row:number;col:number}>=[]
 for(let r=0;r<matrix.rows;r++)for(let c=0;c<matrix.cols;c++)if(matrix.values[r][c]!==null)populated.push({row:r,col:c})
 const cursors=new Map<string,{mode:'populated'|'all';row:number;col:number;ordinal:number}>()
 let seq=0
 const ref:ResolvedRef={
  kind:'resolved-reference',
  areas:[area],
  generation:0,
  readAt(areaIndex,offsetRow,offsetCol){
   if(areaIndex!==0||!Number.isInteger(offsetRow)||!Number.isInteger(offsetCol))return formulaError('#REF!')
   if(offsetRow<0||offsetCol<0||offsetRow>=matrix.rows||offsetCol>=matrix.cols)return formulaError('#REF!')
   return matrix.values[offsetRow][offsetCol]
  },
  openCursor(mode){const id=`m${++seq}`;cursors.set(id,{mode,row:0,col:0,ordinal:0});return {id,mode,generation:0}},
  peek(cursor){
   const st=cursors.get(cursor.id);if(!st)return undefined
   if(st.mode==='populated'){if(st.ordinal>=populated.length)return undefined;const p=populated[st.ordinal];return {areaIndex:0,offsetRow:p.row,offsetCol:p.col,address:{sheetId:'',col:p.col,row:p.row},value:matrix.values[p.row][p.col],origin:'input'}}
   if(st.row>=matrix.rows)return undefined
   return {areaIndex:0,offsetRow:st.row,offsetCol:st.col,address:{sheetId:'',col:st.col,row:st.row},value:matrix.values[st.row][st.col],origin:'input'}
  },
  advance(cursor){const st=cursors.get(cursor.id);if(!st)return;if(st.mode==='populated'){st.ordinal++;return}st.col++;if(st.col>=matrix.cols){st.col=0;st.row++}},
  countAbsent(){return matrix.rows*matrix.cols-populated.length},
 }
 registerReferenceCreator(ref,(newAreas)=>{const a=newAreas.length===1?newAreas[0]:{sheetId:'',firstCol:0,firstRow:0,cols:0,rows:0};const sub:ElementValue[][]=[];for(let r=0;r<a.rows;r++){const out:ElementValue[]=[];for(let c=0;c<a.cols;c++){const rr=a.firstRow+r,cc=a.firstCol+c;out.push(rr>=0&&rr<matrix.rows&&cc>=0&&cc<matrix.cols?matrix.values[rr][cc]:null)}sub.push(out)}return createMatrixRef(matrixOf(a.rows,a.cols,sub))})
 privateMatrixRefs.add(ref)
 return ref
}
function matrixArgument(arg:AstNode,ctx:EvaluationContext|undefined,evalNode:EvaluatorFn):MatrixValue|EvaluationError {
 if(arg.type==='arrayConst')return matrixOf(arg.rows.length,arg.rows[0].length,arg.rows as readonly (readonly ElementValue[])[])
 // Evaluate the expression in array mode through the approved hook (separate
 // stores); a value-denoting name keeps its OWN name node as the request.
 const value=ctx?.evaluateInMode?ctx.evaluateInMode(arg,'array'):evalNode(arg,ctx)
 if(isEvaluationError(value))return value
 if(isMatrixValue(value))return value
 return matrixOf(1,1,[[value]])
}
function matrixView(i:number,arg:AstNode,state:State,ctx:EvaluationContext|undefined,evalNode:EvaluatorFn):View|EvaluationError {
 const matrix=matrixArgument(arg,ctx,evalNode);if(isEvaluationError(matrix))return matrix
 const ref=createMatrixRef(matrix)
 const v={ref,area:ref.areas[0],prepared:true}
 state.views.set(i,v)
 return v
}

/** View of a lookup argument: a real reference when services resolve one; a
 * value/matrix view otherwise. A value-denoting name is evaluated through its
 * OWN name node (evalNode) so active-name cycle identity, laziness, Pending and
 * generation behavior stay intact. */
function view(i:number,args:AstNode[],state:State,ctx:EvaluationContext|undefined,evalNode:EvaluatorFn):View|EvaluationError {
 const cached=state.views.get(i);if(cached)return cached
 const arg=args[i];if(!arg)return formulaError('#VALUE!')
 const services=ctx?.references
 if(services&&ctx&&isReferenceNode(arg)){
  const r=services.resolve(arg,ctx)
  if(r===undefined)return matrixView(i,arg,state,ctx,evalNode)
  if(isEvaluationError(r))return r
  if(!r.areas.length)return formulaError('#REF!')
  const v={ref:r,area:r.areas[0],prepared:false};state.views.set(i,v);return v
 }
 if(!services&&isReferenceNode(arg))return unavailable(ctx,'lookup-reference-unavailable','C1 reference lookup requires full-geometry reference services; public arrays belong to C2')
 return matrixView(i,arg,state,ctx,evalNode)
}
function prepare(v:View,ctx:EvaluationContext|undefined):void {if(!v.prepared){if(!privateMatrixRefs.has(v.ref))ctx?.references?.prepare(v.ref,ctx);v.prepared=true}}
function initialize(args:AstNode[],state:State,ctx:EvaluationContext|undefined,evalNode:EvaluatorFn):EvaluationError|undefined {
 const family=state.family
 if(family==='INDEX'){
  const v=view(0,args,state,ctx,evalNode);if(isEvaluationError(v))return v;state.key=v
  if(args[1]?.type==='empty'||args[2]?.type==='empty')return unavailable(ctx,'lookup-index-omission-unverified','Explicit INDEX omission is distinct from zero-axis selection')
  const row=numeric(1,args,state,ctx,evalNode);if(isEvaluationError(row))return row
  let col:number|EvaluationError
  if(args[2])col=numeric(2,args,state,ctx,evalNode)
  else if(v.area.cols===1)col=1
  else return unavailable(ctx,'lookup-index-omission-unverified','Omitted INDEX column/reference-form behavior is pending its function-specific array/area profile')
  if(isEvaluationError(col))return col
  const area=numeric(3,args,state,ctx,evalNode,1);if(isEvaluationError(area))return area
  if(row<0||col<0)return formulaError('#VALUE!')
  state.indexRow=row;state.indexCol=col;state.indexArea=area-1;state.configured=true;return undefined
 }
 if(args[0]?.type==='empty')return unavailable(ctx,'lookup-omitted-value-unverified','Omitted lookup value has a distinct function-specific blank profile')
 const wanted=argValue(0,args,state,ctx,evalNode);if(isEvaluationError(wanted))return wanted
 if(wanted===null||wanted==='')return unavailable(ctx,'lookup-blank-value-unverified','Blank lookup input matching needs a settled function-specific profile')
 state.lookup=wanted
 const keys=view(1,args,state,ctx,evalNode);if(isEvaluationError(keys))return keys;if(keys.ref.areas.length!==1)return unavailable(ctx,'lookup-reference-shape-unverified','Lookup vector/table must have one rectangular area')
 state.key=keys
 if(family==='VLOOKUP'||family==='HLOOKUP'){
  state.horizontal=family==='HLOOKUP';state.target=keys
  const index=numeric(2,args,state,ctx,evalNode);if(isEvaluationError(index))return index;if(index<1)return formulaError('#VALUE!');if(index>(state.horizontal?keys.area.rows:keys.area.cols))return formulaError('#REF!')
  state.indexRow=state.horizontal?index-1:0;state.indexCol=state.horizontal?0:index-1
  // Key axis is a private 1-col/1-row region over the original table geometry:
  // absence counting/search/preparation use only the key cells; state.target
  // keeps the original table for return reads.
  const keyArea:RefArea={sheetId:keys.area.sheetId,firstCol:keys.area.firstCol,firstRow:keys.area.firstRow,cols:state.horizontal?keys.area.cols:1,rows:state.horizontal?1:keys.area.rows}
  const keyRef=createReferenceRegion(keys.ref,[keyArea])
  if(!keyRef)return unavailable(ctx,'lookup-key-axis-region-unavailable','Key-axis region creation is unavailable for the resolved lookup table')
  state.key={ref:keyRef,area:keyArea,prepared:false}
  const option=argValue(3,args,state,ctx,evalNode,true);if(isEvaluationError(option))return option;const approximate=coerceToBoolean(option);if(isEvaluationError(approximate))return approximate
  state.mode=approximate?-1:0;state.wildcard=!approximate&&typeof wanted==='string'
 }else{
  if(keys.area.rows!==1&&keys.area.cols!==1)return formulaError('#VALUE!');state.horizontal=keys.area.rows===1&&keys.area.cols>1
  const mode=numeric(family==='XLOOKUP'?4:2,args,state,ctx,evalNode,family==='MATCH'?1:0);if(isEvaluationError(mode))return mode
  if(family==='MATCH'){
   if(![-1,0,1].includes(mode))return formulaError('#VALUE!');state.mode=mode===1?-1:mode===-1?1:0;state.search=1;state.wildcard=mode===0&&typeof wanted==='string'
  }else{
   if(![-1,0,1,2].includes(mode))return formulaError('#VALUE!');state.mode=mode as MatchMode
   const search=numeric(family==='XLOOKUP'?5:3,args,state,ctx,evalNode,1);if(isEvaluationError(search))return search;if(![1,-1,2,-2].includes(search))return formulaError('#VALUE!');state.search=search as SearchMode;state.wildcard=mode===2
  }
  if(family==='XLOOKUP'){
   const target=view(2,args,state,ctx,evalNode);if(isEvaluationError(target))return target;if(target.ref.areas.length!==1)return unavailable(ctx,'lookup-reference-shape-unverified','Return reference must have one rectangular area')
   if((state.horizontal?target.area.cols:target.area.rows)!==(state.horizontal?keys.area.cols:keys.area.rows))return formulaError('#VALUE!');state.target=target
  }
 }
 state.length=state.horizontal?keys.area.cols:keys.area.rows
 if(state.wildcard){if(typeof wanted!=='string')return formulaError('#VALUE!');const check=matchWildcard(wanted,'');if(check.status!=='matched'&&check.status!=='not-matched')return unavailable(ctx,`lookup-wildcard-${check.status}`,check.reason);if(Math.abs(state.search)===2)return unavailable(ctx,'lookup-binary-wildcard-unverified','Binary wildcard topology has no accepted native execution profile')}
 prepare(state.key!,ctx);if(state.target&&state.target.ref!==state.key!.ref)prepare(state.target,ctx)
 // No blank scan or guessed absent-cell matching. Absence is counted on the key
 // axis only, so unrelated blank return cells cannot gate the lookup.
 if(state.key!.ref.countAbsent()>0)return unavailable(ctx,'lookup-blank-vector-unverified','Structural absent lookup cells need a function-specific comparison profile')
 state.hi=state.length-1;state.next=state.search===-1?state.length-1:0;state.configured=true;return undefined
}
function keyAt(index:number,state:State):ElementValue {if(state.keys.has(index))return state.keys.get(index)!;const value=state.key!.ref.readAt(0,state.horizontal?0:index,state.horizontal?index:0);state.keys.set(index,value);return value}
function compare(value:ElementValue,wanted:EvaluationValue,state:State,ctx:EvaluationContext|undefined):number|EvaluationError {
 if(isEvaluationError(value))return unavailable(ctx,'lookup-vector-error-unverified','Lookup-vector error inclusion/propagation requires a function-specific profile')
 if(value===null||value==='')return unavailable(ctx,'lookup-blank-vector-unverified','Empty lookup cells are not silently conflated with zero or empty text')
 if(isEvaluationError(wanted))return wanted
 if(state.wildcard){if(typeof value!=='string'||typeof wanted!=='string')return unavailable(ctx,'lookup-wildcard-type-unverified','Mixed wildcard/text lookup rows have no accepted profile');const m=matchWildcard(wanted,value);if(m.status==='matched')return 0;if(m.status==='not-matched')return 1;return unavailable(ctx,`lookup-wildcard-${m.status}`,m.reason)}
 if(typeof value!==typeof wanted){
  if((typeof value==='string'&&typeof wanted==='number'&&!isEvaluationError(coerceToNumber(value)))||(typeof value==='number'&&typeof wanted==='string'&&!isEvaluationError(coerceToNumber(wanted))))return unavailable(ctx,'lookup-numeric-text-unverified','Stored numeric text is not converted to an unverified numeric lookup match')
  return unavailable(ctx,'lookup-mixed-type-unverified','Mixed lookup comparison/order has no accepted finite profile')
 }
 if(typeof value==='number'&&typeof wanted==='number')return value===wanted?0:value<wanted?-1:1
 if(typeof value==='boolean'&&typeof wanted==='boolean')return value===wanted?0:value?1:-1
 if(typeof value==='string'&&typeof wanted==='string'){
  if(/[^\x00-\x7f]/.test(value+wanted))return unavailable(ctx,'lookup-text-unicode-unverified','Only the accepted finite ASCII case profile is implemented')
  const a=value.toLowerCase(),b=wanted.toLowerCase();return a===b?0:a<b?-1:1
 }
 return unavailable(ctx,'lookup-value-unverified','Lookup values are outside the finite scalar comparison contract')
}
function search(state:State,ctx:EvaluationContext|undefined):number|EvaluationError|undefined {
 const wanted=state.lookup!
 if(Math.abs(state.search)===2){
  // Ascending virtual coordinates; descending physical vectors reverse indexing.
  // Genuine logarithmic midpoint search, preserving measured four-row tie choices.
  while(state.lo<=state.hi){const middle=Math.floor((state.lo+state.hi)/2),index=state.search===-2?state.length-1-middle:middle,order=compare(keyAt(index,state),wanted,state,ctx);if(isEvaluationError(order))return order
   if(order===0){state.searched=true;state.found=index;return index}
   if(order<0){if(state.mode===-1)state.best=index;state.lo=middle+1}else{if(state.mode===1)state.best=index;state.hi=middle-1}
  }
  state.searched=true;state.found=state.best;return state.best
 }
 const direction=state.search===-1?-1:1
 while(state.next>=0&&state.next<state.length){const index=state.next,value=keyAt(index,state),order=compare(value,wanted,state,ctx);if(isEvaluationError(order))return order
  if(order===0){
   if(state.family==='MATCH'&&state.mode===-1){state.best=index;state.bestValue=value} // native ascending approximate selects last duplicate
   else {state.searched=true;state.found=index;return index}
  }else if((state.mode===-1&&order<0)||(state.mode===1&&order>0)){
   if(state.best===undefined){state.best=index;state.bestValue=value}else{
    const previous=compare(value,state.bestValue!,{...state,wildcard:false},ctx);if(isEvaluationError(previous))return previous
    if((state.mode===-1&&previous>0)||(state.mode===1&&previous<0)){state.best=index;state.bestValue=value}
   }
  }
  state.next+=direction
 }
 state.searched=true;state.found=state.best;return state.best
}
function run(args:AstNode[],state:State,ctx:EvaluationContext|undefined,evalNode:EvaluatorFn):EvaluationValue {
 if(!state.configured){const e=initialize(args,state,ctx,evalNode);if(e)return e}
 if(state.family==='INDEX'){const result=indexSelection(state.key!.ref,state.indexRow,state.indexCol,state.indexArea,ctx,state.selection);if(!isEvaluationError(result)){state.projected=true;recordProjection(args,ctx)}return result}
 const found=state.searched?state.found:search(state,ctx);if(isEvaluationError(found))return found
 if(found===undefined){if(state.family==='XLOOKUP'&&args[3]&&args[3].type!=='empty')return evalNode(args[3],ctx);return formulaError('#N/A')}
 if(state.family==='MATCH'||state.family==='XMATCH')return found+1
 if(state.family==='XLOOKUP'){const result=xlookupSelection(state.target!.ref,found,state.horizontal,ctx,state.selection);if(isMatrix(result)){state.projected=true;recordProjection(args,ctx)}return result}
 const result=selection(state.target!.ref,0,state.horizontal?state.indexRow:found,state.horizontal?found:state.indexCol,1,1,ctx,state.selection);return result
}
function handler(family:Family,min:number,max:number):FunctionHandler {
 return (args,ctx,evalNode)=>{
  if(args.length<min||args.length>max)return formulaError('#VALUE!')
  // INDEX over a literal/array value (no reference topology) is a C1 carried
  // control: it is a pure positional selection and needs no services.
  if(family==='INDEX'&&args[0]&&!isReferenceNode(args[0]))return indexArrayValue(args,ctx,evalNode)
  let state=stateAt(args,ctx);if(!state||state.family!==family){state=fresh(family,ctx);save(args,ctx,state)}
  for(;;){try{
   const owner=ownedMemoEpoch(ctx);if(owner&&(state.owned?.owner!==owner.owner||state.owned.epoch!==owner.epoch)){if(owner.epoch>0){state=fresh(family,ctx,state.restarts);save(args,ctx,state)}else state.owned=owner}
   const before=ctx?.references?currentReferenceGeneration(ctx.references):undefined;if(state.generation!==undefined&&before!==undefined&&state.generation!==before)throw referenceGenerationChanged(state.generation,before)
   if(state.complete){if(state.projected)recordProjection(args,ctx);return state.result!}
   const result=run(args,state,ctx,evalNode);const after=ctx?.references?currentReferenceGeneration(ctx.references):undefined;if(state.generation!==undefined&&after!==undefined&&state.generation!==after)throw referenceGenerationChanged(state.generation,after)
   state.complete=true;state.result=result;if(!ctx?.functionWork&&ctx)fallbackWork.get(ctx)?.delete(args);return result
  }catch(error){if(!isReferenceGenerationChanged(error))throw error;const count=state.restarts+1;state=fresh(family,ctx,count);save(args,ctx,state);if(count>MAX_REFERENCE_GENERATION_RESTARTS)return unavailable(ctx,'reference-generation-unstable',`Lookup footprint changed beyond the ${MAX_REFERENCE_GENERATION_RESTARTS}-restart application budget`);requestOwnedMemoRestart(ctx)}}
 }
}
export const LOOKUP_FUNCTIONS:Readonly<Record<string,FunctionHandler>>={INDEX:handler('INDEX',2,4),MATCH:handler('MATCH',2,3),VLOOKUP:handler('VLOOKUP',3,4),HLOOKUP:handler('HLOOKUP',3,4),XLOOKUP:handler('XLOOKUP',3,6),XMATCH:handler('XMATCH',2,4)}
