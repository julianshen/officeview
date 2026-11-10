import { describe, expect, it, vi } from 'vitest'
import { prepareCfPaint, disposeSheetCfStatistics } from '../src/xlsx/cf/render'
import { createCfMemory } from '../src/xlsx/cf/limits'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { act, cleanup, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import * as paint from '../src/xlsx/cf/render'
import type { XlsxSheet } from '../src/xlsx/types'

function sheet(): XlsxSheet {
  return { name: 'Live', sheetId: 's', cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [{ row: 0, col: 0, ref: 'A1', styleIndex: 0, value: 1 }] }], conditionalFormatting: { diagnostics: [], rules: [
    { id: 0, priority: 1, documentOrder: 0, stopIfTrue: false, ranges: [{ firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }], writes: new Set(['fill']), dxf: { 'fill.bg': 'FF00FF00' }, definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] } },
    { id: 1, priority: 2, documentOrder: 1, stopIfTrue: false, ranges: [{ firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }], writes: new Set(['fill']), dxf: { 'fill.bg': 'FFFF0000' }, definition: { kind: 'cellIs', operator: 'lessThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] } },
  ] } } as XlsxSheet
}
describe('live CF source generation', () => {
  it('value commit withdraws positive green immediately, then prepares negative red', async () => {
    const s = sheet(), source = createObjectWorkbookSource({ sheets: [s] }, 'live'), store = source.store('s')
    const memory = createCfMemory()
    const old = await prepareCfPaint(s, { memory, store })
    expect(old?.layer.styleAt(0, 0, undefined)?.fillColor).toBe('FF00FF00')
    s.rows[0]!.cells[0]!.value = -1
    store.noteValuesChanged?.()
    expect(old?.layer.styleAt(0, 0, undefined)?.fillColor).toBeUndefined()
    const next = await prepareCfPaint(s, { memory, store })
    expect(next?.layer.styleAt(0, 0, undefined)?.fillColor).toBe('FFFF0000')
    old?.dispose(); next?.dispose(); expect(memory.currentBytes).toBe(0)
  })
})

it('mounted resident commit retires old paint and prepares the committed store', async () => {
  const s = sheet(), doc = { sheets: [s] }, source = createObjectWorkbookSource(doc, 'mounted')
  const batches: import('../src/xlsx/cf/render').CfPaintBatch[] = []
  const actual = paint.prepareCfPaint
  const spy = vi.spyOn(paint, 'prepareCfPaint').mockImplementation(async (sheet, options) => {
    const out = await actual(sheet, options)
    if (out && options?.store) batches.push(out)
    return out
  })
  const ui = render(<InteractiveSheetDoc doc={doc} source={source} />)
  try {
    await waitFor(() => expect(batches.some(b => b.layer.styleAt(0,0,undefined)?.fillColor === 'FF00FF00')).toBe(true))
    const before = batches.slice()
    act(() => { s.rows[0]!.cells[0]!.value = -1; source.store('s').noteValuesChanged?.(); ui.rerender(<InteractiveSheetDoc doc={doc} source={source} />) })
    for (const b of before) expect(b.layer.styleAt(0,0,undefined)?.fillColor).toBeUndefined()
    await waitFor(() => expect(batches.some(b => b.layer.styleAt(0,0,undefined)?.fillColor === 'FFFF0000')).toBe(true))
  } finally { ui.unmount(); cleanup(); spy.mockRestore() }
})

it('changing an authored-hidden source cell in the statistics range withdraws the host style', async () => {
  const s = sheet()
  s.rows.push({ index: 1, hidden: true, cells: [{ ref:'A2', row:1, col:0, styleIndex:0, value:0 }] })
  const rule = s.conditionalFormatting!.rules[0]!
  s.conditionalFormatting = { diagnostics: [], rules: [{ ...rule, ranges:[{firstRow:0,lastRow:1,firstCol:0,lastCol:0}], definition:{kind:'aboveAverage', above:true, equal:false, dialect:'base', formulas:[]} }] }
  const source=createObjectWorkbookSource({sheets:[s]},'statistics'),store=source.store('s'), memory=createCfMemory()
  const allowGate = { allows: () => true }
  const before=await prepareCfPaint(s,{store,memory,evidence:allowGate})
  expect(before?.layer.styleAt(0,0,undefined)?.fillColor).toBe('FF00FF00')
  s.rows[1]!.cells[0]!.value=10; store.noteValuesChanged?.()
  expect(before?.layer.styleAt(0,0,undefined)?.fillColor).toBeUndefined()
  const after=await prepareCfPaint(s,{store,memory,evidence:allowGate})
  expect(after?.layer.styleAt(0,0,undefined)?.fillColor).toBeUndefined()
  before?.dispose();after?.dispose();disposeSheetCfStatistics(s);expect(memory.currentBytes).toBe(0)
})

it('a warm mounted CF frame reuses its prepared state', async () => {
  const s=sheet(),doc={sheets:[s]},source=createObjectWorkbookSource(doc,'warm')
  const spy=vi.spyOn(paint,'prepareCfPaint')
  const ui=render(<InteractiveSheetDoc doc={doc} source={source} />)
  try {
    await waitFor(()=>expect(spy).toHaveBeenCalled())
    await act(async()=>{await new Promise(resolve=>setTimeout(resolve,50))})
    const reads=vi.spyOn(source.store('s'),'readRows'),scan=vi.spyOn(source.store('s'),'scanPopulated')
    const count=spy.mock.calls.length
    for(let i=0;i<100;i++) ui.rerender(<InteractiveSheetDoc doc={doc} source={source} />)
    await act(async()=>{await new Promise(r=>setTimeout(r,50))})
    expect(spy).toHaveBeenCalledTimes(count)
    expect(reads).not.toHaveBeenCalled();expect(scan).not.toHaveBeenCalled();reads.mockRestore();scan.mockRestore()
  } finally {ui.unmount();spy.mockRestore()}
})

it('print preparation crossing a value commit is discarded with a retry notice', async () => {
  const { fireEvent } = await import('@testing-library/react')
  const s=sheet(), doc={sheets:[s]}, source=createObjectWorkbookSource(doc,'print')
  let resume!: (value: never) => void
  const disposed=vi.fn(), pages=Object.assign([], {dispose:disposed})
  const loader=vi.fn(()=>new Promise<never>(resolve=>{resume=resolve}))
  const ui=render(<InteractiveSheetDoc doc={doc} source={source} printLoader={loader} />)
  try {
    fireEvent.keyDown(ui.getByRole('region',{name:'Spreadsheet viewer'}),{key:'p',ctrlKey:true})
    await waitFor(()=>expect(loader).toHaveBeenCalled())
    source.store('s').noteValuesChanged?.()
    await act(async()=>resume(pages as never))
    expect(disposed).toHaveBeenCalledTimes(1)
    expect(ui.getByText(/changed while print was preparing/)).toBeDefined()
  } finally {ui.unmount()}
})

it('pivot overlay supplies the new value without mutating saved cells', async () => {
  const s=sheet(), source=createObjectWorkbookSource({sheets:[s]},'pivot'),store=source.store('s'), memory=createCfMemory()
  const before=await prepareCfPaint(s,{store,memory})
  store.setValueOverlay!({rects:[{r0:0,r1:0,c0:0,c1:0}],protectedRects:[{r0:0,r1:0,c0:0,c1:0}],merges:[],cell:()=>({value:-1})})
  expect(s.rows[0]!.cells[0]!.value).toBe(1)
  expect(before?.layer.styleAt(0,0,undefined)?.fillColor).toBeUndefined()
  const after=await prepareCfPaint(s,{store,memory})
  expect(after?.layer.styleAt(0,0,undefined)?.fillColor).toBe('FFFF0000')
  before?.dispose();after?.dispose();disposeSheetCfStatistics(s);expect(memory.currentBytes).toBe(0)
})

it('paged store overlay follows the same green-to-red generation rule', async () => {
  const { xlsxBytes } = await import('./helpers/xlsx-package')
  const { SourcePageStore, memoryPageBackend } = await import('../src/xlsx/storage/sourcePages')
  const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
  const { createPagedWorkbookSource } = await import('../src/xlsx/view/pagedSource')
  const bytes=xlsxBytes({sheetData:'<row r="1"><c r="A1"><v>1</v></c></row>'})
  const pages=new SourcePageStore({docId:'cf-live-paged',backend:memoryPageBackend()})
  const session=await openImportSession({kind:'blob',blob:new Blob([bytes.slice().buffer as ArrayBuffer])},{docId:'cf-live-paged',profile:'D',store:pages})
  expect((await session.done).kind).toBe('complete')
  const source=createPagedWorkbookSource(session),store=source.store('1'),s=sheet(),memory=createCfMemory()
  s.sheetId='1'
  try {
    const before=await prepareCfPaint(s,{store,memory})
    expect(before?.layer.styleAt(0,0,undefined)?.fillColor).toBe('FF00FF00')
    store.setValueOverlay!({rects:[{r0:0,r1:0,c0:0,c1:0}],protectedRects:[{r0:0,r1:0,c0:0,c1:0}],merges:[],cell:()=>({value:-1})})
    expect(before?.layer.styleAt(0,0,undefined)?.fillColor).toBeUndefined()
    const after=await prepareCfPaint(s,{store,memory})
    expect(after?.layer.styleAt(0,0,undefined)?.fillColor).toBe('FFFF0000')
    before?.dispose();after?.dispose();disposeSheetCfStatistics(s);expect(memory.currentBytes).toBe(0)
  } finally {await source.release()}
})

it('finite CF value commits do not reload image/chart assets, and display-only revisions do not rescan', async()=>{
 const {prepareSheetAssets}=await import('../src/components/xlsx/sheetAssets')
 const s=sheet(),doc={sheets:[s,{...sheet(),name:'Unrelated',sheetId:'other',conditionalFormatting:undefined}]},source=createObjectWorkbookSource(doc,'separate'),load=vi.fn(prepareSheetAssets)
 const read=vi.spyOn(source.store('s'),'readRows'), prepare=vi.spyOn(paint,'prepareCfPaint')
 const ui=render(<InteractiveSheetDoc doc={doc} source={source} assetLoader={load}/>)
 try {
  await waitFor(()=>expect(prepare.mock.results.length).toBeGreaterThan(0))
  await act(async()=>{await new Promise(r=>setTimeout(r,100))})
  const n=prepare.mock.calls.length
  ui.rerender(<InteractiveSheetDoc doc={doc} source={source} assetLoader={load} views={{s:{revision:2}}}/>)
  await act(async()=>{await new Promise(r=>setTimeout(r,100))})
  expect(prepare).toHaveBeenCalledTimes(n)
  source.store('other').noteValuesChanged!()
  ui.rerender(<InteractiveSheetDoc doc={doc} source={source} assetLoader={load} views={{s:{revision:2}}}/>)
  await act(async()=>{await new Promise(r=>setTimeout(r,50))})
  expect(prepare).toHaveBeenCalledTimes(n)
  source.store('s').noteValuesChanged!()
  ui.rerender(<InteractiveSheetDoc doc={doc} source={source} assetLoader={load} views={{s:{revision:2}}}/>)
  await waitFor(()=>expect(prepare.mock.calls.length).toBeGreaterThan(n))
  expect(load).toHaveBeenCalledTimes(1)
 } finally {ui.unmount();read.mockRestore();prepare.mockRestore()}
})

it('importing store withholds paint without a permanent unsupported diagnostic',async()=>{
 const s=sheet(),source=createObjectWorkbookSource({sheets:[s]},'importing'),store=source.store('s')
 Object.defineProperty(store,'completeness',{value:'loading'})
 const read=vi.spyOn(store,'readRows')
 expect(await prepareCfPaint(s,{store})).toBeUndefined()
 expect(read).not.toHaveBeenCalled()
 expect(s.diagnostics?.some(d=>d.feature==='cf-paint-unavailable')).not.toBe(true)
})

it('in-flight resident g1 cannot publish after g2, and removal retires the final scope',async()=>{
 const s=sheet(),doc={sheets:[s]},source=createObjectWorkbookSource(doc,'race'),store=source.store('s')
 const actual=paint.prepareCfPaint
 let resume!:()=>void
 const gate=new Promise<void>(r=>{resume=r}),disposed:import('../src/xlsx/cf/render').CfPaintBatch[]=[]
 const spy=vi.spyOn(paint,'prepareCfPaint').mockImplementation(async(s,o)=>{
  const b=await actual(s,o)
  if(o?.store?.stamp.valueGen===0){await gate}
  if(b)disposed.push(b)
  return b
 })
 const ui=render(<InteractiveSheetDoc doc={doc} source={source}/>)
 try{
  await waitFor(()=>expect(spy).toHaveBeenCalled())
  act(()=>{s.rows[0].cells[0].value=-1;store.noteValuesChanged!();ui.rerender(<InteractiveSheetDoc doc={doc} source={source}/>)})
  await act(async()=>resume())
  await waitFor(()=>expect(disposed.some(b=>b.layer.styleAt(0,0,undefined)?.fillColor==='FFFF0000')).toBe(true))
  expect(disposed.some(b=>b.layer.styleAt(0,0,undefined)?.fillColor==='FF00FF00')).toBe(false)
  ui.unmount()
  expect(disposed.every(b=>!b.layer.styleAt(0,0,undefined)?.fillColor)).toBe(true)
 }finally{resume();cleanup();spy.mockRestore()}
})

it('sorted filtered frozen panes keep source CF and reject an older external view revision',async()=>{
 const {createRowOrder,createIdentityOrder}=await import('../src/xlsx/view/order')
 const {createIntervalSet,visibilityFromHidden}=await import('../src/xlsx/view/visibility')
 const {SheetTileRenderer}=await import('../src/components/xlsx/sheetTiles')
 const s=sheet();s.frozen={rows:1,cols:1}
 s.rows.push({index:1,cells:[{row:1,col:0,ref:'A2',styleIndex:0,value:-1}]})
 s.conditionalFormatting={...s.conditionalFormatting!,rules:s.conditionalFormatting!.rules.map(r=>({...r,ranges:r.ranges.map(x=>({...x,lastRow:1}))}))}
 const doc={sheets:[s]},source=createObjectWorkbookSource(doc,'sorted')
 const order=createRowOrder([{r0:0,r1:1,sourceAtOffset:new Uint32Array([1,0])}])
 const view={revision:2,order,visibility:visibilityFromHidden(createIntervalSet([{r0:1,r1:1}]))}
 const frames:import('../src/components/xlsx/sheetTiles').SheetFrame[]=[]
 const original=SheetTileRenderer.prototype.present
 const spy=vi.spyOn(SheetTileRenderer.prototype,'present').mockImplementation(function(this:InstanceType<typeof SheetTileRenderer>,ctx,frame){frames.push(frame);return original.call(this,ctx,frame)})
 const ui=render(<InteractiveSheetDoc doc={doc} source={source} views={{s:view}}/>)
 try{
  await waitFor(()=>expect(frames.some(f=>f.assets?.conditionalFormatting?.layer.styleAt(1,0,undefined)?.fillColor==='FFFF0000')).toBe(true))
  ui.rerender(<InteractiveSheetDoc doc={doc} source={source} views={{s:{revision:1,order:createIdentityOrder()}}}/>)
  await act(async()=>{await new Promise(r=>setTimeout(r,100))})
  const geometry=frames.at(-1)!.geometry
  if(geometry.kind!=='ready')throw Error('expected ready geometry')
  expect(geometry.metrics.order.displayToSource(0)).toBe(1)
  expect(frames.at(-1)!.assets!.conditionalFormatting!.layer.styleAt(0,0,undefined)?.fillColor).toBe('FF00FF00')
  expect(frames.at(-1)!.assets!.conditionalFormatting!.layer.styleAt(1,0,undefined)?.fillColor).toBe('FFFF0000')
 }finally{ui.unmount();spy.mockRestore()}
})

it('rapid sheet switches discard delayed preparations, including removal before completion',async()=>{
 const {fireEvent}=await import('@testing-library/react')
 const {prepareSheetAssets}=await import('../src/components/xlsx/sheetAssets')
 const a=sheet(),b={...sheet(),sheetId:'b',name:'Other'},doc={sheets:[a,b]},source=createObjectWorkbookSource(doc,'switch')
 let release!:()=>void
 const gate=new Promise<void>(r=>{release=r}),leases:Awaited<ReturnType<typeof prepareSheetAssets>>[]=[]
 const loader:typeof prepareSheetAssets=async(...args)=>{const out=await prepareSheetAssets(...args);leases.push(out);await gate;return out}
 const ui=render(<InteractiveSheetDoc doc={doc} source={source} assetLoader={loader}/>)
 try{
  await waitFor(()=>expect(leases.length).toBeGreaterThan(0))
  fireEvent.change(ui.getByRole('combobox',{name:'Worksheet'}),{target:{value:'b'}})
  await waitFor(()=>expect(leases.length).toBeGreaterThan(1))
  fireEvent.change(ui.getByRole('combobox',{name:'Worksheet'}),{target:{value:'s'}})
  await waitFor(()=>expect(leases.length).toBeGreaterThan(2))
  ui.unmount()
  await act(async()=>release())
  expect(leases.every(l=>l.assets.conditionalFormatting?.layer.styleAt(0,0,undefined)?.fillColor===undefined)).toBe(true)
 }finally{release();cleanup()}
})

it('worker whole-column statistics include the expanded pivot overlay extent',async()=>{
 const {xlsxBytes}=await import('./helpers/xlsx-package')
 const {SourcePageStore,memoryPageBackend}=await import('../src/xlsx/storage/sourcePages')
 const {openImportSession}=await import('../src/xlsx/storage/pagedImport')
 const {createPagedWorkbookSource}=await import('../src/xlsx/view/pagedSource')
 const bytes=xlsxBytes({sheetData:'<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>0</v></c></row>',styles:'<cellXfs count="1"><xf/></cellXfs><dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill></dxf></dxfs>',afterSheetData:'<conditionalFormatting sqref="A:A"><cfRule type="aboveAverage" dxfId="0" priority="1" aboveAverage="1"/></conditionalFormatting>'})
 const pages=new SourcePageStore({docId:'extent',backend:memoryPageBackend()})
 const session=await openImportSession({kind:'blob',blob:new Blob([bytes.slice().buffer as ArrayBuffer])},{docId:'extent',profile:'D',store:pages})
 await session.done
 const source=createPagedWorkbookSource(session),store=source.store('1')
 const read=()=>store.readConditionalFormatting!({rects:[{r0:0,r1:0,c0:0,c1:0}],stamp:{owner:'extent',sheetId:'1',workbookValueGen:store.stamp.valueGen,cfGen:0,viewGen:0,structureGen:store.structureGen??0,clockGen:0},grantBytes:8*1024*1024},createCfMemory())
 try{
  const before=await read();expect(before.kind).toBe('ready');expect(before.states[0]?.properties['fill.fg']).toBe('FF00FF00')
  store.setValueOverlay!({rects:[{r0:10,r1:10,c0:0,c1:0}],protectedRects:[{r0:10,r1:10,c0:0,c1:0}],merges:[],cell:()=>({value:100})})
  const after=await read();expect(after.kind).toBe('ready');expect(after.states.find(s=>s.row===0)?.properties['fill.fg']).toBeUndefined()
 }finally{await source.release()}
})

it('a refused resident scroll scope never lends its new coverage to the old batch',async()=>{
 const {fireEvent}=await import('@testing-library/react')
 const {SheetTileRenderer}=await import('../src/components/xlsx/sheetTiles')
 const s=sheet();s.rows=Array.from({length:200},(_,r)=>({index:r,cells:[{row:r,col:0,ref:`A${r+1}`,styleIndex:0,value:2}]}))
 s.conditionalFormatting={...s.conditionalFormatting!,rules:[{...s.conditionalFormatting!.rules[0],ranges:[{firstRow:0,lastRow:1048575,firstCol:0,lastCol:0}]}]}
 const doc={sheets:[s]},source=createObjectWorkbookSource(doc,'refused-scroll')
 let refuse=false,refusals=0
 const original=paint.prepareCfPaint, present=SheetTileRenderer.prototype.present
 const prep=vi.spyOn(paint,'prepareCfPaint').mockImplementation(async(...args)=>{if(refuse){refusals++;return undefined}return original(...args)})
 const frames:import('../src/components/xlsx/sheetTiles').SheetFrame[]=[]
 const spy=vi.spyOn(SheetTileRenderer.prototype,'present').mockImplementation(function(this:InstanceType<typeof SheetTileRenderer>,ctx,f){frames.push(f);return present.call(this,ctx,f)})
 const ui=render(<InteractiveSheetDoc doc={doc} source={source}/>)
 try{
  await waitFor(()=>expect(frames.some(f=>f.assets?.conditionalFormatting?.layer.styleAt(0,0,undefined)?.fillColor==='FF00FF00')).toBe(true))
  refuse=true
  const scroll=ui.getByTestId('officeview-sheet-scroller')
  act(()=>{scroll.scrollTop=1200;fireEvent.scroll(scroll)})
  await waitFor(()=>expect(refusals).toBeGreaterThan(0))
  await act(async()=>{await new Promise(r=>setTimeout(r,50))})
  const f=frames.at(-1)!
  expect(f.cfPending || !f.cfCoverage?.({x:0,y:1024,width:512,height:512})).toBe(true)
 }finally{ui.unmount();prep.mockRestore();spy.mockRestore()}
})

it('max edit recolours the whole scale scope in one generation', async () => {
  const scaleSheet: XlsxSheet = {
    name: 'ScaleLive',
    sheetId: 's',
    cols: [],
    merges: [],
    mergeRanges: [],
    rows: [
      { index: 0, cells: [{ row: 0, col: 0, ref: 'A1', styleIndex: 0, value: 0 }] },
      { index: 1, cells: [{ row: 1, col: 0, ref: 'A2', styleIndex: 0, value: 10 }] },
      { index: 2, cells: [{ row: 2, col: 0, ref: 'A3', styleIndex: 0, value: 0 }] },
    ],
    conditionalFormatting: {
      diagnostics: [],
      rules: [
        {
          id: 0,
          priority: 1,
          documentOrder: 0,
          stopIfTrue: false,
          ranges: [{ firstRow: 0, lastRow: 2, firstCol: 0, lastCol: 0 }],
          writes: new Set(['colorScale', 'fill']),
          definition: {
            kind: 'colorScale',
            points: [
              { bound: { type: 'min' }, argb: 0xfff8696b },
              { bound: { type: 'max' }, argb: 0xff63be7b },
            ],
            dialect: 'base',
            formulas: [],
          },
        },
      ],
    },
  } as XlsxSheet

  const doc = { sheets: [scaleSheet] }
  const source = createObjectWorkbookSource(doc, 'scale-live')
  const store = source.store('s')
  const memory = createCfMemory()
  const allowGate = { allows: () => true }

  const g1 = await prepareCfPaint(scaleSheet, { memory, store, evidence: allowGate })
  // In g1, max is 10, so A2 (10) gets max color FF63BE7B
  expect(g1?.layer.styleAt(1, 0, undefined)?.fillColor).toBe('FF63BE7B')

  // Commit value 20 at A3 (so max becomes 20)
  scaleSheet.rows[2]!.cells[0]!.value = 20
  store.noteValuesChanged?.()

  // At g2 while preparation is pending: no cell paints its g1 colour at g2!
  expect(g1?.layer.styleAt(1, 0, undefined)?.fillColor).toBeUndefined()

  // Prepare next generation g2
  const g2 = await prepareCfPaint(scaleSheet, { memory, store, evidence: allowGate })
  // In g2, max is 20, so A2 (10) is 50% midpoint (recoloured!), A3 (20) gets max color FF63BE7B
  expect(g2?.layer.styleAt(2, 0, undefined)?.fillColor).toBe('FF63BE7B')
  expect(g2?.layer.styleAt(1, 0, undefined)?.fillColor).not.toBe('FF63BE7B')

  // Pivot commit, sort/filter, sheet switch and close all release cleanly
  store.setValueOverlay?.({ rects: [{ r0: 0, r1: 2, c0: 0, c1: 0 }], protectedRects: [{ r0: 0, r1: 2, c0: 0, c1: 0 }], merges: [], cell: () => ({ value: 10 }) })
  const g3 = await prepareCfPaint(scaleSheet, { memory, store, evidence: allowGate })
  g1?.dispose(); g2?.dispose(); g3?.dispose()
  disposeSheetCfStatistics(scaleSheet)
  expect(memory.currentBytes).toBe(0)
})

it('max edit invalidates ONLY the tiles of the scale scope (tile-level) (T15-A)', async () => {
  const { TileCache } = await import('../src/xlsx/view/tileCache')
  const { SheetTileRenderer } = await import('../src/components/xlsx/sheetTiles')

  const rows = Array.from({ length: 60 }, (_, r) => ({
    index: r,
    cells: [{ row: r, col: 0, ref: `A${r + 1}`, styleIndex: 0, value: r }],
  }))

  const scaleSheet: XlsxSheet = {
    name: 'ScaleTileSheet',
    sheetId: 's',
    cols: [{ min: 0, max: 0, widthChars: 12 }],
    merges: [],
    mergeRanges: [],
    rows,
    conditionalFormatting: {
      diagnostics: [],
      rules: [
        {
          id: 1,
          priority: 1,
          documentOrder: 0,
          stopIfTrue: false,
          ranges: [{ firstRow: 0, lastRow: 59, firstCol: 0, lastCol: 0 }], // A1:A60
          writes: new Set(['colorScale', 'fill']),
          definition: {
            kind: 'colorScale',
            points: [
              { bound: { type: 'min' }, argb: 0xfff8696b },
              { bound: { type: 'max' }, argb: 0xff63be7b },
            ],
            dialect: 'base',
            formulas: [],
          },
        },
      ],
    },
  } as XlsxSheet

  const doc = { sheets: [scaleSheet] }
  const source = createObjectWorkbookSource(doc, 'tile-live')
  const store = source.store('s')

  let activeRenderer: InstanceType<typeof SheetTileRenderer> | undefined
  const origPresent = SheetTileRenderer.prototype.present
  const presentSpy = vi.spyOn(SheetTileRenderer.prototype, 'present').mockImplementation(function (
    this: InstanceType<typeof SheetTileRenderer>,
    ctx,
    frame,
  ) {
    activeRenderer = this
    return origPresent.call(this, ctx, frame)
  })

  const invalidateCalls: Array<(key: import('../src/xlsx/view/tileCache').TileKey) => boolean> = []
  const origInvalidate = TileCache.prototype.invalidate
  const invalidateSpy = vi.spyOn(TileCache.prototype, 'invalidate').mockImplementation(function (
    this: InstanceType<typeof TileCache>,
    predicate,
  ) {
    invalidateCalls.push(predicate)
    return origInvalidate.call(this, predicate)
  })

  const ui = render(<InteractiveSheetDoc doc={doc} source={source} />)
  try {
    await waitFor(() => expect(activeRenderer).toBeDefined())

    const identity = { docId: 'tile-live', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }
    const tileInScope1 = { paletteKey: 'classic', identity, backend: 'canvas2d', zoom: 1, dpr: 1, col: 0, row: 0 }
    const tileInScope2 = { paletteKey: 'classic', identity, backend: 'canvas2d', zoom: 1, dpr: 1, col: 0, row: 1 }
    const tileOutOfScope = { paletteKey: 'classic', identity, backend: 'canvas2d', zoom: 1, dpr: 1, col: 5, row: 0 }
    const tileOtherSheet = {
      paletteKey: 'classic',
      identity: { ...identity, sheetId: 'other' },
      backend: 'canvas2d',
      zoom: 1,
      dpr: 1,
      col: 0,
      row: 0,
    }

    const dummySurface = (w = 512, h = 512) => ({ width: w, height: h, close: vi.fn() })
    activeRenderer!.cache.put(tileInScope1, dummySurface())
    activeRenderer!.cache.put(tileInScope2, dummySurface())
    activeRenderer!.cache.put(tileOutOfScope, dummySurface())

    expect(activeRenderer!.cache.get(tileInScope1)).toBeDefined()
    expect(activeRenderer!.cache.get(tileInScope2)).toBeDefined()
    expect(activeRenderer!.cache.get(tileOutOfScope)).toBeDefined()

    const beforeCallsCount = invalidateCalls.length

    // Perform max edit on cell A50 (row 49, col 0)
    act(() => {
      scaleSheet.rows[49]!.cells[0]!.value = 1000
      store.noteValuesChanged?.()
      ui.rerender(<InteractiveSheetDoc doc={doc} source={source} />)
    })

    await waitFor(() => expect(invalidateCalls.length).toBeGreaterThan(beforeCallsCount))
    const cfPredicate = invalidateCalls[beforeCallsCount]!

    // (T15-A) Assertions:
    // 1. Max edit invalidates ONLY the tiles of the scale scope (tile-level)
    expect(cfPredicate(tileInScope1), 'Tile in scope (rows 0..25) must be invalidated').toBe(true)
    expect(cfPredicate(tileInScope2), 'Tile in scope (rows 26..51) must be invalidated').toBe(true)
    expect(cfPredicate(tileOutOfScope), 'Tile out of scope (col 5) must NOT be invalidated').toBe(false)
    expect(cfPredicate(tileOtherSheet), 'Tile on another sheet must NOT be invalidated').toBe(false)

    // 1b. A value edit bumps valueGen, which is part of the tile identity: what an edit actually relies on is
    // the identity sweep in present() (every old-identity tile of this renderer is dropped, scope or not).
    // The CF-scope predicate above is the fallback for identity-PRESERVING changes (clock, text generation), which
    // the sorted-view clock test exercises. Assert the sweep so this test stops proving only the fallback.
    await waitFor(() => {
      expect(activeRenderer!.cache.get(tileInScope1), 'old-identity in-scope tile swept').toBeUndefined()
      expect(activeRenderer!.cache.get(tileOutOfScope), 'old-identity out-of-scope tile swept by identity, not by CF scope').toBeUndefined()
    })

    // 2. Cache retention verification: invalidates ONLY the tiles of the scale scope
    const testCache = new TileCache(activeRenderer!.cache.pool)
    testCache.put(tileInScope1, dummySurface())
    testCache.put(tileInScope2, dummySurface())
    testCache.put(tileOutOfScope, dummySurface())

    testCache.invalidate(cfPredicate)

    expect(testCache.get(tileInScope1), 'tileInScope1 dropped').toBeUndefined()
    expect(testCache.get(tileInScope2), 'tileInScope2 dropped').toBeUndefined()
    expect(testCache.get(tileOutOfScope), 'tileOutOfScope retained').toBeDefined()
  } finally {
    ui.unmount()
    presentSpy.mockRestore()
    invalidateSpy.mockRestore()
  }
})

it('clock-driven CF invalidation follows the DISPLAYED rows of a sorted view (I1)', async () => {
  const { TileCache } = await import('../src/xlsx/view/tileCache')
  const { SheetTileRenderer } = await import('../src/components/xlsx/sheetTiles')
  const { createRowOrder } = await import('../src/xlsx/view/order')
  const N = 60
  const rows = Array.from({ length: N }, (_, r) => ({ index: r, cells: [{ row: r, col: 0, ref: `A${r + 1}`, styleIndex: 0, value: r }] }))
  const todaySheet: XlsxSheet = {
    name: 'SortedClock', sheetId: 's', cols: [{ min: 0, max: 0, widthChars: 12 }], merges: [], mergeRanges: [], rows,
    conditionalFormatting: { diagnostics: [], rules: [{
      id: 1, priority: 1, documentOrder: 0, stopIfTrue: false, ranges: [{ firstRow: 0, lastRow: 2, firstCol: 0, lastCol: 0 }],
      writes: new Set(['fill']), dxf: { 'fill.bg': 'FF00FF00' },
      definition: { kind: 'predicate', type: 'timePeriod', timePeriod: 'today', dialect: 'base', formulas: [] },
    }] },
  } as unknown as XlsxSheet
  // Reversed order: source rows 0..2 are DISPLAYED at slots 59..57 (tile row 2), not tile row 0.
  const order = createRowOrder([{ r0: 0, r1: N - 1, sourceAtOffset: Uint32Array.from({ length: N }, (_, i) => N - 1 - i) }])
  const doc = { sheets: [todaySheet] }
  const source = createObjectWorkbookSource(doc, 'sorted-clock')
  vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'Date'], now: new Date(2026, 5, 10, 23, 59, 50) })
  const invalidateCalls: Array<(key: import('../src/xlsx/view/tileCache').TileKey) => boolean> = []
  const origInvalidate = TileCache.prototype.invalidate
  const spy = vi.spyOn(TileCache.prototype, 'invalidate').mockImplementation(function (this: InstanceType<typeof TileCache>, predicate) {
    invalidateCalls.push(predicate)
    return origInvalidate.call(this, predicate)
  })
  const presented = vi.spyOn(SheetTileRenderer.prototype, 'present')
  const ui = render(<InteractiveSheetDoc doc={doc} source={source} views={{ s: { revision: 1, order } }} />)
  try {
    await vi.waitFor(() => expect(presented).toHaveBeenCalled())
    const before = invalidateCalls.length
    await act(async () => { await vi.advanceTimersByTimeAsync(15_000) })
    await vi.waitFor(() => expect(invalidateCalls.length).toBeGreaterThan(before))
    const identity = { docId: 'sorted-clock', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }
    const tile = (row: number) => ({ paletteKey: 'classic', identity, backend: 'canvas2d', zoom: 1, dpr: 1, col: 0, row })
    // The first invalidation after the clock tick is the CF scope predicate (later ones are identity sweeps).
    const cfPredicate = invalidateCalls[before]!
    // The identity-preserving clock change must drop the tile that DISPLAYS the scope (row 2) and no other.
    expect(cfPredicate(tile(2) as never), 'tile row 2 displays source rows 0..2').toBe(true)
    expect(cfPredicate(tile(0) as never), 'tile row 0 shows nothing of the scope').toBe(false)
  } finally {
    ui.unmount(); spy.mockRestore(); presented.mockRestore(); vi.useRealTimers()
  }
})
