import { describe, test, expect, vi } from 'vitest'
import { chartWorkbookBytes, chartWorkbook } from './helpers/xlsx-chart-fixture'
import { parseXlsx } from '../src/xlsx/parse'
import { parseXlsxChart, fillWorkbookCharts } from '../src/charts'
import { parseXmlOrdered } from '../src/core/xml'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { MemoryPool } from '../src/xlsx/view/tileCache'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
const xml = (ref: string) => `<chartSpace><chart><plotArea><barChart><ser><idx val="0"/><val><numRef><f>${ref}</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>`
const chartOf = (drawings: any[]) => drawings[0].content.chart
async function paged(bytes: Uint8Array) {
 const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId:'review', profile:'D' })
 expect((await session.done).kind).toBe('complete')
 const source = createPagedWorkbookSource(session)
 const adapter = source.store('1') as any
 return { source, adapter, session }
}

describe('independent review counterexamples', () => {
 test('paged loading reads only the single line consumed by chart rangeCells', async () => {
  const cells = Array.from({length:100}, (_,r) => Array.from({length:26},(_,c) => ({ref:`${String.fromCharCode(65+c)}${r+1}`,v: r*26+c}))).flat()
  const {source,adapter} = await paged(await chartWorkbookBytes(xml('Styled!A1:Z100'), {cells}))
  try {
   const valueOf = vi.spyOn(adapter, 'valueOf')
   const result = await adapter.readDrawings({stamp:adapter.stamp,signal:new AbortController().signal},'D')
   expect(chartOf(result.drawings).groups[0].series[0].values).toHaveLength(100)
   console.log('resolved cell count',valueOf.mock.calls.length,'output point count',100)
   expect(valueOf.mock.calls.length).toBeLessThanOrEqual(100)
  } finally { await source.release() }
 })
 test('paged and resident preserve last column of a large two dimensional source', async () => {
  const bytes=await chartWorkbookBytes(xml('Styled!A1:GJH5000'),{cells:[{ref:'GJH1',v:42},{ref:'GJH2',v:84}]})
  const resident=await parseXlsx(await chartWorkbook(xml('Styled!A1:GJH5000'),{cells:[{ref:'GJH1',v:42},{ref:'GJH2',v:84}]}))
  const expected=chartOf(resident.sheets[0].drawings!).groups[0].series[0].values
  expect(expected.slice(0,2)).toEqual([42,84])
  const {source,adapter}=await paged(bytes)
  try {
   const result=await adapter.readDrawings({stamp:adapter.stamp,signal:new AbortController().signal},'D')
   expect(chartOf(result.drawings).groups[0].series[0].values.slice(0,2)).toEqual([42,84])
  } finally { await source.release() }
 })
 test('malformed series index is contained as a placeholder instead of aborting XLSX parsing',async()=>{
  await expect(parseXlsx(await chartWorkbook(xml('Styled!A1').replace('idx val="0"','idx val="-1"')))).resolves.toBeDefined()
 })
 test('a one-cell chart reference does not format every cell on a resident sheet',()=>{
  const model=parseXlsxChart(parseXmlOrdered(xml('Styled!A1')),undefined)!
  const sheets=[{name:'Styled',rows:Array.from({length:10000},(_,r)=>({index:r,cells:[{row:r,col:0,value:r}]})),drawings:[{content:{kind:'xlsxChart',chart:model}}]}]
  const text=vi.fn((_:unknown,cell:any)=>String(cell.value))
  fillWorkbookCharts(sheets,{text})
  expect(sheets[0].drawings[0].content.chart.groups[0].series[0].values).toEqual([0])
  console.log('resident formatting calls',text.mock.calls.length,'output point count',1)
  expect(text.mock.calls.length).toBeLessThanOrEqual(1)
 })
})

test('uncached chart follows the resident formula engine value',async()=>{
 const pkg=await chartWorkbook(xml('Styled!B1'),{cells:[{ref:'A1',v:3},{ref:'B1',v:99}],patch:async zip=>{
  const path='xl/worksheets/sheet1.xml'
  const sheet=await zip.file(path)!.async('string')
  zip.file(path,sheet.replace('<v>99</v>','<f ca="1">A1+4</f><v>99</v>'))
 }})
 const doc=await parseXlsx(pkg)
 const formula=doc.sheets[0].rows.flatMap(r=>r.cells).find(c=>c.col===1)!
 expect(formula.value).toBe(7)
 expect(chartOf(doc.sheets[0].drawings!).groups[0].series[0].values).toEqual([7])
})


test('cross-sheet chart waits for referenced sheet and never caches incomplete data as complete',async()=>{
 const bytes=await chartWorkbookBytes(xml('Data!A1:A2'),{patch:async zip=>{
  const book='xl/workbook.xml',rels='xl/_rels/workbook.xml.rels'
  zip.file(book,(await zip.file(book)!.async('string')).replace('</sheets>','<sheet name="Data" sheetId="2" r:id="rData"/></sheets>'))
  zip.file(rels,(await zip.file(rels)!.async('string')).replace('</Relationships>','<Relationship Id="rData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>'))
  zip.file('xl/worksheets/sheet2.xml','<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>42</v></c></row><row r="2"><c r="A2"><v>84</v></c></row></sheetData></worksheet>')
 }})
 const {source,adapter,session}=await paged(bytes)
 const pool=new MemoryPool({limitBytes:64*1024*1024}); source.bindMemory!(pool)
 const data=session.workbook.sheets[1]
 const savedRow=data.importedThroughRow
 data.importedThroughRow=-1; data.state='loading'
 try {
  const first=await adapter.readDrawings({stamp:adapter.stamp,signal:new AbortController().signal},'D')
  expect(first.completeness).toBe('complete')
  console.log('cross-sheet first incomplete values',chartOf(first.drawings).groups[0].series[0].values)
  data.importedThroughRow=savedRow; data.state='complete'
  const second=await adapter.readDrawings({stamp:adapter.stamp,signal:new AbortController().signal},'D')
  expect(chartOf(second.drawings).groups[0].series[0].values).toEqual([42,84])
 } finally {await source.release();pool.close()}
})
