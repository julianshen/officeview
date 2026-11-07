import { describe, expect, it, vi } from 'vitest'
import { prepareCfPaint } from '../src/xlsx/cf/render'
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

it('changing another cell in the statistics range withdraws the host style', async () => {
  const s = sheet()
  s.rows.push({ index: 1, cells: [{ ref:'A2', row:1, col:0, styleIndex:0, value:0 }] })
  const rule = s.conditionalFormatting!.rules[0]!
  s.conditionalFormatting = { diagnostics: [], rules: [{ ...rule, ranges:[{firstRow:0,lastRow:1,firstCol:0,lastCol:0}], definition:{kind:'aboveAverage', above:true, equal:false, dialect:'base', formulas:[]} }] }
  const source=createObjectWorkbookSource({sheets:[s]},'statistics'),store=source.store('s'), memory=createCfMemory()
  const before=await prepareCfPaint(s,{store,memory})
  expect(before?.layer.styleAt(0,0,undefined)?.fillColor).toBe('FF00FF00')
  s.rows[1]!.cells[0]!.value=10; store.noteValuesChanged?.()
  expect(before?.layer.styleAt(0,0,undefined)?.fillColor).toBeUndefined()
  const after=await prepareCfPaint(s,{store,memory})
  expect(after?.layer.styleAt(0,0,undefined)?.fillColor).toBeUndefined()
  before?.dispose();after?.dispose();expect(memory.currentBytes).toBe(0)
})

it('a warm mounted CF frame reuses its prepared state', async () => {
  const s=sheet(),doc={sheets:[s]},source=createObjectWorkbookSource(doc,'warm')
  const spy=vi.spyOn(paint,'prepareCfPaint')
  const ui=render(<InteractiveSheetDoc doc={doc} source={source} />)
  try {
    await waitFor(()=>expect(spy).toHaveBeenCalled())
    await act(async()=>{await new Promise(resolve=>setTimeout(resolve,50))})
    const count=spy.mock.calls.length
    for(let i=0;i<100;i++) ui.rerender(<InteractiveSheetDoc doc={doc} source={source} />)
    expect(spy).toHaveBeenCalledTimes(count)
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
  before?.dispose();after?.dispose();expect(memory.currentBytes).toBe(0)
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
    before?.dispose();after?.dispose();expect(memory.currentBytes).toBe(0)
  } finally {await source.release()}
})
