// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { chartWorkbook, chartWorkbookBytes, addNumberFormat } from './helpers/xlsx-chart-fixture'
import { parseXlsx } from '../src/xlsx/parse'
import { parseXlsxChart } from '../src/xlsx/chart'
import { parseXmlOrdered } from '../src/core/xml'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { MemoryPool } from '../src/xlsx/view/tileCache'

const xml = (cat: string, val: string) => `<chartSpace><chart><plotArea><barChart><ser><cat><numRef><f>${cat}</f></numRef></cat><val><numRef><f>${val}</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>`
const model = (result: any) => result.drawings[0].content.chart
async function open(bytes: Uint8Array, opts: object = {}) {
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId:'edge', profile:'D', ...opts })
  expect((await session.done).kind).toBe('complete')
  const source = createPagedWorkbookSource(session)
  const adapter = source.store('1') as any
  const read = (memory?: any) => adapter.readDrawings({stamp:adapter.stamp, signal:new AbortController().signal}, 'D', memory)
  return {session, source, adapter, read}
}

it('paged categories retain number formatting when the value range overlaps', async () => {
  const chart=xml('Styled!B1:B2','Styled!B1:B2')
  const options={ cells:[{ref:'B1',v:1,style:3},{ref:'B2',v:2,style:3}], patch:async(zip:any)=>addNumberFormat(zip,14) }
  const resident=await parseXlsx(await chartWorkbook(chart,options))
  const expected=model({drawings:resident.sheets[0].drawings}).categories
  const h=await open(await chartWorkbookBytes(chart,options))
  try {
    const result=await h.read()
    console.log('overlap categories', {resident:expected, paged:model(result).categories})
    expect(model(result).categories).toEqual(expected)
  } finally {await h.source.release()}
})

it('failed referenced sheet is reported as unknown instead of cached complete blanks', async () => {
 const bytes=await chartWorkbookBytes(xml('Data!A1:A2','Styled!B1:B2'),{cells:[{ref:'B1',v:3},{ref:'B2',v:4}],patch:async zip=>{
  const book='xl/workbook.xml',rels='xl/_rels/workbook.xml.rels'
  zip.file(book,(await zip.file(book)!.async('string')).replace('</sheets>','<sheet name="Data" sheetId="2" r:id="rData"/></sheets>'))
  zip.file(rels,(await zip.file(rels)!.async('string')).replace('</Relationships>','<Relationship Id="rData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>'))
  zip.file('xl/worksheets/sheet2.xml','<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>42</v></c></row><row r="2"><c r="A2"><v>84</v></c></row></sheetData></worksheet>')
 }})
 const h=await open(bytes), data=h.session.workbook.sheets[1]
 const pool=new MemoryPool({limitBytes:64*1024*1024}); h.source.bindMemory!(pool)
 const through=data.importedThroughRow
 data.importedThroughRow=-1; data.state='incomplete';data.failure={code:'storage-error',message:'data unavailable'} as any
 try {
  const first=await h.read()
  console.log('failed reference', {completeness:first.completeness,pendingSheets:first.pendingSheets,diagnostics:first.diagnostics,categories:model(first).categories})
  data.importedThroughRow=through;data.state='complete';data.failure=undefined
  const second=await h.read()
  console.log('failed reference restored',model(second).categories)
  expect(first.completeness==='incomplete'||first.pendingSheets?.includes('2')||model(first).unsupported).toBeTruthy()
  // Not cached as complete: a later read after the sheet recovers sees its cells.
  expect(first.diagnostics.some((d:any)=>d.message.includes('data unavailable'))).toBe(true)
  expect(model(second).categories).toEqual(['42','84'])
 } finally {await h.source.release();pool.close()}
})

it('reads each distinct source page once when multiple series share one range', async () => {
 const chart='<chartSpace><chart><plotArea><barChart>'+Array.from({length:64},()=>'<ser><val><numRef><f>Styled!A1:A100</f></numRef></val></ser>').join('')+'</barChart></plotArea></chart></chartSpace>'
 const h=await open(await chartWorkbookBytes(chart,{cells:Array.from({length:100},(_,i)=>({ref:`A${i+1}`,v:i}))}),{pageRows:1})
 try {
  let summaryProbes=0
  const summaries=h.session.workbook.store.pages('1').map(p=>({...p,get r0(){summaryProbes++;return p.r0}}))
  vi.spyOn(h.session.workbook.store,'pages').mockReturnValue(summaries)
  const before=h.session.workbook.store.stats
  const page=vi.spyOn(h.session.workbook.store,'page')
  const result=await h.read()
  const after=h.session.workbook.store.stats
  console.log('page calls',page.mock.calls.length,'distinct',new Set(page.mock.calls.map(c=>c.join('/'))).size,'summary probes',summaryProbes,'backend loads',after.loads-before.loads,'cache hits',after.hits-before.hits)
  expect(model(result).groups[0].series).toHaveLength(64)
  expect(summaryProbes).toBeLessThanOrEqual(10_000)
  expect(page.mock.calls.length).toBeLessThanOrEqual(101)
 } finally {await h.source.release()}
})

it('truncation includes category indices and scatter x values', () => {
 for (const [kind,part] of [['barChart','<cat><strLit><pt idx="5000"><v>x</v></pt></strLit></cat><val><numLit><pt idx="0"><v>1</v></pt></numLit></val>'],['scatterChart','<xVal><numLit><ptCount val="4097"/></numLit></xVal><yVal><numLit><pt idx="0"><v>1</v></pt></numLit></yVal>']]) {
  const chart=parseXlsxChart(parseXmlOrdered(`<chartSpace><chart><plotArea><${kind}><ser>${part}</ser></${kind}></plotArea></chart></chartSpace>`),undefined)!
  expect(chart.truncated).toBe(true)
 }
})

// Adapted to the bounded-stream design: a label is capped at CHART_LABEL_LIMIT
// (255) characters and the chart marked truncated, so a 100,000-character
// category is never built in full and fits a small allowance. Only ACCEPTED
// reservations count as admission (F3).
it('a long category label is capped (truncated), admitted by its capped bytes and never built in full', async () => {
 const label='x'.repeat(100_000)
 const h=await open(await chartWorkbookBytes(xml('Styled!A1','Styled!B1'),{cells:[{ref:'A1',v:label},{ref:'B1',v:5}]}))
 try {
  const accepted:number[]=[], built:number[]=[], rawValue=h.adapter.valueOf.bind(h.adapter)
  vi.spyOn(h.adapter,'valueOf').mockImplementation(async (...args:any[])=>{const v=await rawValue(...args);if(typeof v==='string')built.push(v.length);return v})
  const result=await h.read({set:(bytes:number)=>{const ok=bytes<=100_000;if(ok)accepted.push(bytes);return ok}})
  expect(result.completeness).toBe('complete')
  expect(model(result).categories[0]).toBe('x'.repeat(255))
  expect(model(result).truncated).toBe(true)
  expect(built.every(length=>length<=256)).toBe(true)
  expect(Math.max(...accepted)).toBeLessThan(2*label.length)
 } finally {await h.source.release()}
})

it('a refused label reservation materializes nothing and is not counted as admission', async () => {
 const h=await open(await chartWorkbookBytes(xml('Styled!A1','Styled!B1'),{cells:[{ref:'A1',v:'label text'},{ref:'B1',v:5}]}))
 try {
  // Learn the reservation sequence with room, then refuse the first chart-label reservation.
  const roomy:number[]=[]
  await h.read({set:(bytes:number)=>{roomy.push(bytes);return true}})
  const accepted:number[]=[], refusedAt:number[]=[], afterRefusal:string[]=[]
  let refused=false
  const rawValue=h.adapter.valueOf.bind(h.adapter)
  vi.spyOn(h.adapter,'valueOf').mockImplementation(async (...args:any[])=>{if(refused)afterRefusal.push('valueOf');return rawValue(...args)})
  const fresh=await open(await chartWorkbookBytes(xml('Styled!A1','Styled!B1'),{cells:[{ref:'A1',v:'label text'},{ref:'B1',v:5}]}))
  try {
   const freshRaw=fresh.adapter.valueOf.bind(fresh.adapter)
   vi.spyOn(fresh.adapter,'valueOf').mockImplementation(async (...args:any[])=>{if(refused)afterRefusal.push('valueOf');return freshRaw(...args)})
   // The label's prefix reservation adds exactly 2 B per character ('label text': 20 B) on top of the workspace.
   const at=roomy.findIndex((bytes,i)=>i+1<roomy.length&&roomy[i+1]-bytes===2*'label text'.length)
   expect(at).toBeGreaterThanOrEqual(0)
   const workspace=roomy[at]
   const result=await fresh.read({set:(bytes:number)=>{const ok=!refused&&bytes<=workspace;if(ok)accepted.push(bytes);else{refused=true;refusedAt.push(bytes)}return ok}})
   expect(refusedAt.length).toBeGreaterThan(0)
   expect(refusedAt[0]-workspace).toBe(2*'label text'.length)
   expect(accepted.every(bytes=>bytes<refusedAt[0])).toBe(true)
   expect(result.completeness).toBe('incomplete')
   expect(afterRefusal).toEqual([])
  } finally {await fresh.source.release()}
 } finally {await h.source.release()}
})

it('numeric chart workspace refusal happens before cell loading', async () => {
 const chart='<chartSpace><chart><plotArea><barChart><ser><val><numRef><f>Styled!A1:A4096</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>'
 const h=await open(await chartWorkbookBytes(chart))
 try {
  const values=vi.spyOn(h.adapter,'valueOf'), charged:number[]=[]
  const result=await h.read({set:(bytes:number)=>{charged.push(bytes);return bytes<=100_000}})
  console.log('numeric workspace refusal',{charged,values:values.mock.calls.length,unsupported:model(result).unsupported})
  expect(values).not.toHaveBeenCalled()
  expect(model(result).unsupported).toContain('memory')
 } finally {await h.source.release()}
})
