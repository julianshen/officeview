// @vitest-environment node
import { pts } from './helpers/chart-points'
import {expect,it} from 'vitest'
import {writeFile} from 'node:fs/promises'
import {chartWorkbookBytes} from './helpers/xlsx-chart-fixture'
import {openImportSession} from '../src/xlsx/storage/pagedImport'
import {createPagedWorkbookSource} from '../src/xlsx/view/pagedSource'
import {parseXlsx} from '../src/xlsx/parse'
import {OfficePackage} from '../src/core/zip'
import {cellDisplayText} from '../src/xlsx/render'
const evidence:Record<string,unknown>={},save=async()=>{void evidence;void writeFile}
const xml=(cat:string|undefined,val:string)=>`<chartSpace><chart><plotArea><barChart><ser>${cat?`<cat><numRef><f>${cat}</f></numRef></cat>`:''}<val><numRef><f>${val}</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>`
async function open(bytes:Uint8Array){
 const doc=await parseXlsx(await OfficePackage.load(bytes))
 const session=await openImportSession({kind:'blob',blob:new Blob([new Uint8Array(bytes)])},{docId:'p5',profile:'D'})
 expect((await session.done).kind).toBe('complete')
 const source=createPagedWorkbookSource(session),adapter=source.store('1') as any
 return{doc,source,read:()=>adapter.readDrawings({stamp:adapter.stamp,signal:new AbortController().signal},'D')}
}
// Adapted: Excel does not read radix text as a number (coordinator decision), so the
// safe 2^52/2^51 radix cases are not numbers on either path; the control case stays not a number.
for(const [name,text,expected]of [['safe-hex','0x10000000000000',undefined],['safe-octal','0o100000000000000000',undefined],['control','0x_x0011_',undefined]]as const)it(`public numeric reading ${name} follows Excel on both paths`,async()=>{
 const h=await open(await chartWorkbookBytes(xml(undefined,'Styled!A1'),{cells:[{ref:'A1',v:text}]}))
 try{
  const raw=h.doc.sheets[0].rows[0].cells[0].value
  if(name==='control')expect(raw).toBe('0x\u0011')
  const resident=(h.doc.sheets[0].drawings![0].content as any).chart,paged=(await h.read()).drawings[0].content.chart
  evidence[name]={raw,expected,resident:resident.groups[0].series[0].values,paged:paged.groups[0].series[0].values,truncated:resident.truncated};await save()
  expect(resident.groups[0].series[0].values).toEqual(pts(expected))
  expect(paged.groups[0].series[0].values).toEqual(pts(expected))
 }finally{await h.source.release()}
})
it('grouped chart categories keep the canonical formatted prefix on both paths',async()=>{
 const h=await open(await chartWorkbookBytes(xml('Styled!A1','Styled!B1'),{cells:[{ref:'A1',v:1e299,style:3},{ref:'B1',v:5}],patch:async zip=>{
  const styles=await zip.file('xl/styles.xml')!.async('string')
  zip.file('xl/styles.xml',styles.replace('<cellXfs count="3">','<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0"/></numFmts><cellXfs count="4">').replace('</cellXfs>','<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'))
 }}))
 try{
  const sheet=h.doc.sheets[0],full=cellDisplayText(sheet,sheet.rows[0].cells[0]),expected=full.slice(0,255),resident=(sheet.drawings![0].content as any).chart.categories[0],paged=(await h.read()).drawings[0].content.chart.categories[0]
  evidence.grouped={expected,resident,paged};await save()
  expect(resident).toBe(expected);expect(paged).toBe(expected)
 }finally{await h.source.release()}
})
