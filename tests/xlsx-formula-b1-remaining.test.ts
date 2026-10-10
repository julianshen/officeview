/** Synthetic B1 contract controls, not claims of native measured parity. */
import { describe, expect, it } from 'vitest'
import { evaluateFormula, isEvaluationError } from '../src/xlsx/formula/evaluator'
import { parseFormula } from '../src/xlsx/formula/parser'
import { formatFormula, translateSharedFormula } from '../src/xlsx/formula/shared'
import { createReferenceServices, type StoredCell } from '../src/xlsx/formula/refs'
import type { AstNode, DefinedNameMetadata, EvaluationContext, EvaluationValue, ReferenceNode, TableMetadata } from '../src/xlsx/formula/types'
import type { XlsxCell, XlsxDocument } from '../src/xlsx/types'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'

const sheets = [{ sheetId: 's', name: 'S', workbookIndex: 0 }, { sheetId: 't', name: 'T', workbookIndex: 2 }]
function fixture(names: DefinedNameMetadata[] = [], tables: TableMetadata[] = [], cells: StoredCell[] = []) {
  const issues: string[] = []
  const reads = { count: 0 }
  const generation = { n: 1 }
  const services = createReferenceServices({
    generation: () => generation.n, sheets,
    sheetIdOfName: name => sheets.find(s => s.name.toLowerCase() === name.toLowerCase())?.sheetId,
    sheetNameOfId: id => sheets.find(s => s.sheetId === id)?.name,
    definedNames: names, tables,
    store: { stored: () => cells, read: a => {
      reads.count++
      const c = cells.find(c => c.address.sheetId === a.sheetId && c.address.col === a.col && c.address.row === a.row)
      return c ? { value: c.value, origin: c.origin } : undefined
    } },
    onIssue: feature => issues.push(feature),
  })
  const ctx: EvaluationContext = {
    references: services, currentSheet: 'S', currentCell: { col: 1, row: 1, absCol: false, absRow: false },
    currentAddress: { sheetId: 's', col: 1, row: 1 }, flatArgs: new WeakMap(), unsupportedFeatures: new Set(),
    reportFormulaIssue: issue => issues.push(issue.feature ?? ''),
  }
  return { services, ctx, issues, reads, generation }
}
function reference(source: string): ReferenceNode {
  const ast = parseFormula(source)
  switch (ast.type) {
    case 'cell': case 'range': case 'wholeCol': case 'wholeRow': case 'name':
    case 'table': case 'ref3d': case 'union': case 'intersect': case 'spill': return ast
    default: throw new Error(`Invalid reference fixture: ${source}`)
  }
}
function cell(sheetId: string, col: number, row: number, value: number): StoredCell {
  return { address: { sheetId, col, row }, value, origin: 'input' }
}
function workbook(formula: string, names: DefinedNameMetadata[], tables: TableMetadata[] = []): { doc: XlsxDocument; live: XlsxCell; dead: XlsxCell } {
  const live: XlsxCell = { ref: 'B2', col: 1, row: 1, value: 77, styleIndex: 0, hasCachedValue: true, formula }
  const dead: XlsxCell = { ref: 'B3', col: 1, row: 2, value: 77, styleIndex: 0, hasCachedValue: true, formula: `IF(FALSE,${formula},8)` }
  const doc: XlsxDocument = { sheets: sheets.map((s,i) => ({ ...s, cols: [], merges: [], mergeRanges: [], rows: i === 0 ? [{ index: 1, cells: [live] }, { index: 2, cells: [dead] }] : [] })), definedNames: names, tables }
  return { doc, live, dead }
}

describe('B1 remaining group5 source names', () => {
  it.each(['_xlfn.MyName','_xlfn._xlws.MyName','_xlfn.TRUE','S!MyName',"'T'!_xlfn.MixedName"])('preserves bare/qualified %s in parse, format and shared copies', source => {
    const bang = source.indexOf('!')
    const name = bang < 0 ? source : source.slice(bang + 1)
    const sheet = bang < 0 ? undefined : source.slice(0,bang).replace(/'/g,'')
    const ast = { type: 'name', ref: { name, ...(sheet ? { sheet } : {}) } }
    expect(parseFormula(source)).toEqual(ast)
    const moved = translateSharedFormula(source, 3, 4)
    expect(moved.ast).toEqual(ast)
    expect(parseFormula(moved.formula)).toEqual(ast)
    expect(formatFormula(parseFormula(source))).toContain(name)
  })
  it('prefix names bind case-insensitively; untaken unknown name stays silent; call prefixes normalize', () => {
    const f = fixture([{ name: '_xlfn.MyName', source: '123', baseProvenance: 'unknown' }, { name: 'MixedName', source: '9', localSheetIndex: 2, baseProvenance: 'unknown' }])
    expect(evaluateFormula('_xlfn.myname',f.ctx)).toBe(123)
    expect(evaluateFormula('T!mIxEdNaMe',f.ctx)).toBe(9)
    expect(evaluateFormula('IF(FALSE,_xlfn.Missing,8)',f.ctx)).toBe(8)
    expect(f.issues).toEqual([])
    expect(evaluateFormula('_xlfn.SUM(1,2)',f.ctx)).toBe(3)
    expect(evaluateFormula('_xlfn._xlws.SUM(1,2)',f.ctx)).toBe(3)
  })
})

describe('B1 remaining group4 relative names', () => {
  it.each(['explicit','verified-file'] as const)('preserves %s metadata/source identity and shifts per endpoint at two use sites', baseProvenance => {
    const base = { sheetId: 's', col: 0, row: 0 }
    const f = fixture([
      { name: 'Rel', source: 'A1', relativeBase: base, baseProvenance },
      { name: 'Mixed', source: 'B$1:$D3', relativeBase: base, baseProvenance },
      { name: 'Cols', source: '$A:C', relativeBase: base, baseProvenance },
      { name: 'Rows', source: '$1:3', relativeBase: base, baseProvenance },
    ],[],[cell('s',1,1,42),cell('s',2,2,93)])
    const binding = f.services.bindName({ name: 'Rel' }, f.ctx)
    expect(binding).toMatchObject({ relativeBase: base, baseProvenance, source: 'A1' })
    expect(evaluateFormula('Rel',f.ctx)).toBe(42)
    for (const [col,row,expected] of [[1,1,42],[2,2,93]]) {
      f.ctx.currentCell = { col,row,absCol:false,absRow:false }; f.ctx.currentAddress = { sheetId:'s',col,row }
      expect(evaluateFormula('Rel',f.ctx)).toBe(expected)
      const again = f.services.bindName({ name:'rel' },f.ctx)
      if (isEvaluationError(binding) || isEvaluationError(again)) throw new Error('Missing valid binding fixture')
      expect(again.ast).toBe(binding.ast)
      expect(again.source).toBe('A1')
      const mixed = f.services.resolve(reference('Mixed'),f.ctx)
      expect(mixed).toMatchObject({ areas: [{ sheetId:'s',firstCol:Math.min(1+col,3),firstRow:0,cols:Math.abs(3-(1+col))+1,rows:3+row }] })
      expect(f.services.resolve(reference('Cols'),f.ctx)).toMatchObject({ areas:[{ firstCol:0,cols:3+col,rows:1048576 }] })
      expect(f.services.resolve(reference('Rows'),f.ctx)).toMatchObject({ areas:[{ firstRow:0,rows:3+row,cols:16384 }] })
    }
    expect(f.issues).toEqual([])
  })
  it('qualified/original-index scope keeps its metadata, explicit source sheet and use-site offset', () => {
    const f = fixture([
      { name:'Rel',source:'S!A1',relativeBase:{sheetId:'s',col:0,row:0},baseProvenance:'explicit' },
      { name:'Rel',source:'T!$A1',relativeBase:{sheetId:'t',col:3,row:0},baseProvenance:'verified-file',localSheetIndex:2 },
    ],[],[cell('s',1,1,42),cell('t',0,1,19)])
    expect(evaluateFormula('Rel',f.ctx)).toBe(42)
    expect(evaluateFormula('T!Rel',f.ctx)).toBe(19)
    expect(f.services.bindName({name:'Rel',sheet:'T'},f.ctx)).toMatchObject({scopeSheetId:'t',relativeBase:{sheetId:'t',col:3,row:0},baseProvenance:'verified-file'})
    expect(f.issues).toEqual([])
  })
  it('unknown base retains gate/cache77/dead8; known out-of-domain shift is a genuine REF; absolute names need no base', () => {
    const names: DefinedNameMetadata[] = [{name:'UnknownRel',source:'A1',baseProvenance:'unknown'}, {name:'Underflow',source:'A1',relativeBase:{sheetId:'s',col:4,row:4},baseProvenance:'explicit'}, {name:'Absolute',source:'$A$1',baseProvenance:'unknown'}]
    const f = fixture(names,[],[cell('s',0,0,5)])
    expect(evaluateFormula('UnknownRel',f.ctx)).toBe('#NAME?')
    expect(f.ctx.unsupportedFeatures?.has('relative-name')).toBe(true)
    expect(f.issues).toContain('relative-name')
    expect(evaluateFormula('Underflow',f.ctx)).toBe('#REF!')
    expect(evaluateFormula('Absolute',f.ctx)).toBe(5)
    const {doc,live,dead} = workbook('UnknownRel',names)
    evaluateWorkbookFormulas(doc,{forceRecalc:true})
    expect(live.value).toBe(77); expect(dead.value).toBe(8)
    expect(doc.diagnostics?.some(i=>i.feature==='relative-name')).toBe(true)
  })
})

const table: TableMetadata = { id:'table',name:'Table1',displayName:'Table1',sheetId:'t',partPath:'table.xml',extent:{sheetId:'t',firstCol:0,firstRow:0,cols:2,rows:4},headerRowCount:1,totalsRowCount:1,columns:[{id:'a',name:'Amount',index:0},{id:'b',name:'Other',index:1}] }
describe('B1 remaining group7 thisRow placement', () => {
  it.each(['Table1[@Amount]','Table1[[#This Row],[Amount]]'])('cross-sheet %s gates without reading despite matching row, retains cache77/dead8', formula => {
    const f=fixture([],[table],[cell('t',0,1,19)])
    expect(evaluateFormula(formula,f.ctx)).toBe('#NAME?')
    expect(f.ctx.unsupportedFeatures?.has('table-this-row-placement')).toBe(true)
    expect(f.issues).toContain('table-this-row-placement'); expect(f.reads.count).toBe(0)
    const {doc,live,dead}=workbook(formula,[],[table])
    evaluateWorkbookFormulas(doc,{forceRecalc:true})
    expect(live.value).toBe(77); expect(dead.value).toBe(8)
    expect(doc.diagnostics?.some(i=>i.feature==='table-this-row-placement')).toBe(true)
    const quiet=fixture([],[table]); expect(evaluateFormula(`IF(FALSE,${formula},8)`,quiet.ctx)).toBe(8); expect(quiet.issues).toEqual([])
  })
  it('owning stable address succeeds, currentSheet fallback succeeds, wrong table-context sheet gates', () => {
    const f=fixture([],[table],[cell('t',0,1,19)])
    f.ctx.currentSheet='T';f.ctx.currentAddress={sheetId:'t',col:1,row:1}
    expect(evaluateFormula('Table1[@Amount]',f.ctx)).toBe(19)
    f.ctx.tableContext={tableId:table.id,currentCell:f.ctx.currentAddress}
    expect(evaluateFormula('[@Amount]',f.ctx)).toBe(19)
    delete f.ctx.currentAddress;delete f.ctx.tableContext
    expect(evaluateFormula('Table1[@Amount]',f.ctx)).toBe(19)
    f.ctx.currentSheet='S';f.ctx.tableContext={tableId:table.id,currentCell:{sheetId:'s',col:1,row:1}}
    expect(evaluateFormula('[@Amount]',f.ctx)).toBe('#NAME?')
    expect(f.issues).toContain('table-this-row-placement')
  })
  it('stable currentAddress is authoritative; header/totals/outside body placement stays diagnosed', () => {
    const f=fixture([],[table],[cell('t',0,1,19)])
    f.ctx.currentSheet='T' // contradicting legacy name cannot override resolved address on S
    expect(evaluateFormula('Table1[@Amount]',f.ctx)).toBe('#NAME?')
    for (const row of [0,3,9]) {
      f.ctx.currentAddress={sheetId:'t',col:1,row};f.ctx.currentCell={col:1,row,absCol:false,absRow:false}
      expect(evaluateFormula('Table1[@Amount]',f.ctx)).toBe('#NAME?')
      expect(f.ctx.unsupportedFeatures?.has('table-this-row-placement')).toBe(true)
    }
    expect(f.reads.count).toBe(0)
  })
})

describe('B1 remaining group6 compound escaped selectors', () => {
  const headers=['Amount','A]B','A[B',"O'",'@Rate','#Tag']
  const t: TableMetadata={...table,sheetId:'s',extent:{...table.extent,sheetId:'s',cols:6},columns:headers.map((name,index)=>({id:String(index),name,index}))}
  it.each([
    ["Table1[[A']B],[A'[B]]",'union',16],
    ["Table1[[A']B]:[A'[B]]",'table',16],
    ["Table1[[O''],['#Tag]]",'union',40],
    ["Table1[['@Rate]:['#Tag]]",'table',48],
    ['Table1[@[Amount]]','table',2],
    ["Table1[@[A']B]:[A'[B]]",'table',8],
    ["Table1[@O'']",'table',7],
  ] as const)('decodes %s, resolves selectors and preserves raw/shared roundtrip', (source,type,total) => {
    const f=fixture([],[t],headers.flatMap((_name,col)=>[cell('s',col,1,[2,3,5,7,11,13][col]),cell('s',col,2,[2,3,5,7,11,13][col])]))
    const ast=parseFormula(source);expect(ast.type).toBe(type)
    if (ast.type==='table') expect(ast.ref.raw).toBe(source)
    if (ast.type==='union') {
      expect(ast.refs).toHaveLength(2)
      expect(ast.refs.map(n=>n.type==='table'?n.ref.columns:null)).not.toContain(null)
    }
    expect(evaluateFormula(`SUM(${source})`,f.ctx)).toBe(total)
    const translated=translateSharedFormula(source,3,4)
    expect(translated.ast).toEqual(ast);expect(parseFormula(translated.formula)).toEqual(ast)
    expect(translated.formula).toContain("Table1")
  })
  it.each(["Table1[A']B]","Table1[A'[B]","Table1[O'']","Table1['@Rate]","Table1['#Tag]"] as const)('retains exact decoded/header and raw spelling for %s',source=>{
    const ast=parseFormula(source)
    expect(ast.type).toBe('table')
    if(ast.type!=='table') throw new Error('Invalid escaped table fixture')
    const expected:Record<string,string>={"Table1[A']B]":'A]B',"Table1[A'[B]":'A[B',"Table1[O'']":"O'","Table1['@Rate]":'@Rate',"Table1['#Tag]":'#Tag'}
    expect(ast.ref.columns).toEqual({kind:'single',name:expected[source]})
    expect(ast.ref.raw).toBe(source)
    expect(formatFormula(ast)).toBe(source)
  })
})

describe('B1 remaining actual @ value-name ownership', () => {
  it.each(['@Loop','SUM(@Loop)','SUM(Loop)'])('actual %s with Loop=@Loop is bounded, specifically diagnosed, cache77 retained and dead8 silent', formula => {
    const names: DefinedNameMetadata[]=[{name:'Loop',source:'@Loop',baseProvenance:'unknown'}]
    const f=fixture(names)
    let result: unknown
    expect(()=>{result=evaluateFormula(formula,f.ctx)}).not.toThrow()
    expect(result).toBe('#NAME?');expect(f.ctx.unsupportedFeatures?.has('name-cycle')).toBe(true);expect(f.issues).toContain('name-cycle')
    const quiet=fixture(names);expect(evaluateFormula(`IF(FALSE,${formula},8)`,quiet.ctx)).toBe(8);expect(quiet.issues).toEqual([])
    const {doc,live,dead}=workbook(formula,names);evaluateWorkbookFormulas(doc,{forceRecalc:true})
    expect(live.value).toBe(77);expect(dead.value).toBe(8);expect(doc.diagnostics?.some(i=>i.feature==='name-cycle')).toBe(true)
  })
  it('pending preparation before an @ value-name cycle keeps marker/memo identity; retry reaches cycle; ordinary @Five stays supported',()=>{
    const f=fixture([{name:'R',source:'$A$1',baseProvenance:'unknown'},{name:'Loop',source:'SUM(R,@Loop)',baseProvenance:'unknown'},{name:'Five',source:'5',baseProvenance:'unknown'}],[],[cell('s',0,0,5)])
    const pending=new Error('pending @ name fixture');let first=true
    f.services.prepare=()=>{if(first){first=false;throw pending}}
    const memo=new WeakMap<AstNode,EvaluationValue>();const sentinel=parseFormula('42');memo.set(sentinel,42);f.ctx.nodeValues=memo
    let caught: unknown;try{evaluateFormula('@Loop',f.ctx)}catch(error){caught=error}
    expect(caught).toBe(pending);expect(f.ctx.nodeValues).toBe(memo);expect(memo.get(sentinel)).toBe(42)
    expect(()=>evaluateFormula('@Loop',f.ctx)).not.toThrow()
    expect(evaluateFormula('@Loop',f.ctx)).toBe('#NAME?');expect(f.issues).toContain('name-cycle')
    expect(f.ctx.nodeValues).toBe(memo);expect(memo.get(sentinel)).toBe(42)
    expect(evaluateFormula('@Five',f.ctx)).toBe(5)
    const local=fixture([{name:'Five',source:'5',baseProvenance:'unknown'}]);expect(evaluateFormula('@Five',local.ctx)).toBe(5);expect(local.ctx.nodeValues).toBeUndefined()
  })
})

describe('B1 remaining alias reference geometry', () => {
  it('aliases to ranges preserve full SUM/ROWS/intersection/union geometry instead of scalar projection',()=>{
    const f=fixture([{name:'Alias',source:'Inner',baseProvenance:'unknown'},{name:'Inner',source:'$A$1:$A$2',baseProvenance:'unknown'}],[],[cell('s',0,0,10),cell('s',0,1,20)])
    expect(evaluateFormula('SUM(Alias)',f.ctx)).toBe(30)
    expect(evaluateFormula('ROWS(Alias)',f.ctx)).toBe(2)
    expect(evaluateFormula('SUM(Alias $A$2:$A$3)',f.ctx)).toBe(20)
    expect(evaluateFormula('SUM((Alias,Alias))',f.ctx)).toBe(60)
    expect(evaluateFormula('@Alias',f.ctx)).toBe(20)
  })
  it('absolute3000 alias-to-whole-axis stays bounded, sparse and full-sized at two use sites',()=>{
    const names:DefinedNameMetadata[]=Array.from({length:3000},(_,i)=>({name:`Shape_${i}`,source:i===2999?'$A:$A':`Shape_${i+1}`,baseProvenance:'unknown'}))
    const f=fixture(names,[],[cell('s',0,0,10),cell('s',0,499999,20)])
    let value:unknown
    expect(()=>{value=evaluateFormula('SUM(Shape_0)',f.ctx)}).not.toThrow();expect(value).toBe(30)
    expect(f.reads.count).toBe(2)
    expect(evaluateFormula('ROWS(Shape_0)',f.ctx)).toBe(1048576)
    expect(evaluateFormula('SUM(Shape_0 $A$500000)',f.ctx)).toBe(20)
    f.ctx.currentCell={col:3,row:499999,absCol:false,absRow:false};f.ctx.currentAddress={sheetId:'s',col:3,row:499999}
    expect(evaluateFormula('@Shape_0',f.ctx)).toBe(20)
  })
  it.each(['AliasLoop','RefLoop'])('reference topology cycle %s gates boundedly/cache77/dead8 without claiming native cycle value',name=>{
    const names:DefinedNameMetadata[]=[{name:'AliasLoop',source:'AliasLoop',baseProvenance:'unknown'},{name:'RefLoop',source:'(RefLoop,$A$1)',baseProvenance:'unknown'}]
    const f=fixture(names,[],[cell('s',0,0,5)]);let value:unknown
    expect(()=>{value=evaluateFormula(`SUM(${name})`,f.ctx)}).not.toThrow();expect(value).toBe('#NAME?');expect(f.issues).toContain('name-cycle')
    const {doc,live,dead}=workbook(`SUM(${name})`,names);evaluateWorkbookFormulas(doc,{forceRecalc:true});expect(live.value).toBe(77);expect(dead.value).toBe(8)
    expect(doc.diagnostics?.some(i=>i.feature==='name-cycle')).toBe(true)
  })
  it('aliased range pending drain retains marker/memo and exact aggregate on retry',()=>{
    const f=fixture([{name:'Alias',source:'Inner',baseProvenance:'unknown'},{name:'Inner',source:'$A$1:$A$2',baseProvenance:'unknown'}],[],[cell('s',0,0,10),cell('s',0,1,20)])
    const pending=new Error('pending aliased range');let first=true
    f.services.prepare=()=>{if(first){first=false;throw pending}}
    const memo=new WeakMap<AstNode,EvaluationValue>();const sentinel=parseFormula('42');memo.set(sentinel,42);f.ctx.nodeValues=memo
    let caught:unknown;try{evaluateFormula('SUM(Alias)',f.ctx)}catch(error){caught=error}
    expect(caught).toBe(pending);expect(f.ctx.nodeValues).toBe(memo);expect(memo.get(sentinel)).toBe(42)
    expect(evaluateFormula('SUM(Alias)',f.ctx)).toBe(30);expect(f.reads.count).toBe(2)
    expect(f.ctx.nodeValues).toBe(memo);expect(memo.get(sentinel)).toBe(42);expect(f.issues).toEqual([])
  })
})

describe('B1 remaining resolver classification ownership',()=>{
  it.each([['S','T'],['T','S']] as const)('same global alias visits %s then %s without poisoning VALUE/reference classification',(...order)=>{
    const f=fixture([{name:'Alias',source:'Target',baseProvenance:'unknown'},{name:'Target',source:'5',localSheetIndex:0,baseProvenance:'unknown'},{name:'Target',source:'$A$1:$A$2',localSheetIndex:2,baseProvenance:'unknown'}],[],[cell('t',0,0,10),cell('t',0,1,20)])
    for(const sheet of order){f.ctx.currentSheet=sheet;f.ctx.currentAddress={sheetId:sheet==='S'?'s':'t',col:1,row:1};delete f.ctx.nodeValues
      expect(evaluateFormula('SUM(Alias)',f.ctx)).toBe(sheet==='S'?5:30)
      expect(f.ctx.nodeValues).toBeUndefined()
    }
    expect(f.issues).toEqual([])
  })
  it('changed child source, generation and base metadata each refresh with a fresh memo',()=>{
    const target:DefinedNameMetadata={name:'Target',source:'5',baseProvenance:'unknown'}
    const f=fixture([{name:'Alias',source:'Target',baseProvenance:'unknown'},target],[],[cell('s',0,0,10),cell('s',0,1,20),cell('s',1,1,42)])
    expect(evaluateFormula('SUM(Alias)',f.ctx)).toBe(5)
    target.source='$A$1:$A$2' // source changes must not require callers to invent a generation bump
    f.ctx.nodeValues=new WeakMap()
    expect(evaluateFormula('SUM(Alias)',f.ctx)).toBe(30)
    target.source='A1';target.relativeBase={sheetId:'s',col:0,row:0};target.baseProvenance='explicit';f.generation.n++;f.ctx.nodeValues=new WeakMap()
    expect(evaluateFormula('Alias',f.ctx)).toBe(42)
    target.relativeBase={sheetId:'s',col:1,row:1};f.ctx.nodeValues=new WeakMap()
    expect(evaluateFormula('Alias',f.ctx)).toBe(10)
    expect(f.issues).toEqual([])
  })
  it.each(['5','$A:$A'])('3000 aliases ending %s use a linear number of bindings and correct full result',source=>{
    const n=3000
    const f=fixture(Array.from({length:n},(_,i)=>({name:`Bound_${i}`,source:i===n-1?source:`Bound_${i+1}`,baseProvenance:'unknown' as const})),[],[cell('s',0,0,10),cell('s',0,499999,20)])
    let bindings=0;const bind=f.services.bindName.bind(f.services)
    f.services.bindName=(name,ctx)=>{bindings++;return bind(name,ctx)}
    expect(evaluateFormula('SUM(Bound_0)',f.ctx)).toBe(source==='5'?5:30)
    expect(bindings).toBeLessThanOrEqual(3*n+5)
    expect(f.reads.count).toBe(source==='5'?0:2)
    expect(f.issues).toEqual([])
  })
  it('shared copies of duplicate ordinary table operands preserve both occurrences',()=>{
    const t:TableMetadata={...table,sheetId:'s',extent:{...table.extent,sheetId:'s'}}
    const f=fixture([],[t],[cell('s',0,1,5)])
    const moved=translateSharedFormula('(Table1[Amount],Table1[Amount])',2,3)
    expect(parseFormula(moved.formula)).toEqual(moved.ast)
    expect(evaluateFormula(`SUM(${moved.formula})`,f.ctx)).toBe(10)
  })
})

describe('B1 remaining relative value-name bodies',()=>{
  it.each(['explicit','verified-file'] as const)('%s base shifts CALL and arithmetic sources with mixed anchors at two sites',baseProvenance=>{
    const base={sheetId:'s',col:0,row:0}
    const f=fixture([{name:'ValueRel',source:'SUM(A1:A2)',relativeBase:base,baseProvenance},{name:'ScalarRel',source:'A1+1',relativeBase:base,baseProvenance},{name:'MixedValue',source:'SUM($A1:B$2)',relativeBase:base,baseProvenance},{name:'MixedRef',source:'$A1:B$2',relativeBase:base,baseProvenance}],[],[cell('s',0,1,3),cell('s',1,1,42),cell('s',2,1,7),cell('s',0,2,1),cell('s',1,2,9),cell('s',2,2,93),cell('s',3,2,2),cell('s',2,3,4)])
    f.ctx.getCellValue=(sheet,col,row)=>{const values:Record<string,number>={'S:1:1':42,'S:2:2':93};return values[`${sheet}:${col}:${row}`]??null}
    for(const [col,row,sum,scalar,mixed] of [[1,1,51,43,52],[2,2,97,94,157]]){
      f.ctx.currentCell={col,row,absCol:false,absRow:false};f.ctx.currentAddress={sheetId:'s',col,row}
      expect(evaluateFormula('ValueRel',f.ctx)).toBe(sum)
      expect(evaluateFormula('ScalarRel',f.ctx)).toBe(scalar)
      expect(evaluateFormula('MixedValue',f.ctx)).toBe(mixed)
      expect(evaluateFormula('SUM(MixedRef)',f.ctx)).toBe(mixed)
      expect(f.services.bindName({name:'ValueRel'},f.ctx)).toMatchObject({source:'SUM(A1:A2)',ast:parseFormula('SUM(A1:A2)'),relativeBase:base,baseProvenance})
    }
    expect(f.issues).toEqual([])
  })
  it('unknown relative body gates only reached references; CALL dead path8 stays silent and live workbook cache77 retained',()=>{
    const names:DefinedNameMetadata[]=[{name:'UnknownValue',source:'SUM(A1:A2)',baseProvenance:'unknown'},{name:'QuietValue',source:'IF(FALSE,SUM(A1:A2),8)',baseProvenance:'unknown'}]
    const f=fixture(names,[],[cell('s',0,0,5)])
    expect(evaluateFormula('UnknownValue',f.ctx)).toBe('#NAME?');expect(f.issues).toContain('relative-name');expect(f.reads.count).toBe(0)
    const quiet=fixture(names);expect(evaluateFormula('QuietValue',quiet.ctx)).toBe(8);expect(quiet.issues).toEqual([])
    const {doc,live,dead}=workbook('UnknownValue',names);evaluateWorkbookFormulas(doc,{forceRecalc:true});expect(live.value).toBe(77);expect(dead.value).toBe(8);expect(doc.diagnostics?.some(i=>i.feature==='relative-name')).toBe(true)
  })
  it('translated CALL child identity survives pending replay under the supplied memo',()=>{
    const f=fixture([{name:'ValueRel',source:'SUM(A1:A2)',relativeBase:{sheetId:'s',col:0,row:0},baseProvenance:'explicit'}],[],[cell('s',1,1,42),cell('s',1,2,9)])
    const pending=new Error('relative body pending');let first=true;f.services.prepare=()=>{if(first){first=false;throw pending}}
    const memo=new WeakMap<AstNode,EvaluationValue>();const sentinel=parseFormula('42');memo.set(sentinel,42);f.ctx.nodeValues=memo
    let caught:unknown;try{evaluateFormula('ValueRel',f.ctx)}catch(error){caught=error};expect(caught).toBe(pending)
    expect(evaluateFormula('ValueRel',f.ctx)).toBe(51);expect(f.reads.count).toBe(2)
    expect(f.ctx.nodeValues).toBe(memo);expect(memo.get(sentinel)).toBe(42);expect(f.issues).toEqual([])
  })
})

describe('B1 remaining relative value-name scope',()=>{
  it('qualified local CALL keeps original workbook index/source sheet while offsets use each caller site',()=>{
    const f=fixture([{name:'ValueRel',source:'SUM(A1:A2)',relativeBase:{sheetId:'s',col:0,row:0},baseProvenance:'explicit'},{name:'ValueRel',source:'SUM(T!$A1:$A2)',relativeBase:{sheetId:'t',col:5,row:0},baseProvenance:'verified-file',localSheetIndex:2}],[],[cell('s',1,1,42),cell('s',1,2,9),cell('t',0,1,19),cell('t',0,2,23),cell('t',0,3,31)])
    delete f.ctx.currentAddress // approved legacy currentCell/currentSheet basis remains supported
    expect(evaluateFormula('ValueRel',f.ctx)).toBe(51)
    expect(evaluateFormula('T!ValueRel',f.ctx)).toBe(42)
    f.ctx.currentCell={col:2,row:2,absCol:false,absRow:false}
    expect(evaluateFormula('T!ValueRel',f.ctx)).toBe(54)
    expect(f.services.bindName({name:'ValueRel',sheet:'T'},f.ctx)).toMatchObject({scopeSheetId:'t',source:'SUM(T!$A1:$A2)',baseProvenance:'verified-file'})
    expect(f.issues).toEqual([])
  })
})
