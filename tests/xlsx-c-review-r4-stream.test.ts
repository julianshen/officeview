// @vitest-environment node
import {expect,it,vi} from 'vitest'
import {writeFile} from 'node:fs/promises'
import {chartWorkbookBytes} from './helpers/xlsx-chart-fixture'
import {openImportSession} from '../src/xlsx/storage/pagedImport'
import {createPagedWorkbookSource} from '../src/xlsx/view/pagedSource'
import {parseXlsx} from '../src/xlsx/parse'
import {OfficePackage} from '../src/core/zip'
import * as render from '../src/xlsx/render'
import * as pages from '../src/xlsx/storage/sourcePages'
const evidence:Record<string,unknown>={}
const xml=(cat:string|undefined,val:string)=>`<chartSpace><chart><plotArea><barChart><ser>${cat?`<cat><numRef><f>${cat}</f></numRef></cat>`:''}<val><numRef><f>${val}</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>`
async function open(bytes:Uint8Array,opts:object={}){
 const session=await openImportSession({kind:'blob',blob:new Blob([new Uint8Array(bytes)])},{docId:'r4',profile:'D',...opts})
 expect((await session.done).kind).toBe('complete')
 const source=createPagedWorkbookSource(session),adapter=source.store('1') as any
 return{session,source,adapter,read:(memory?:any)=>adapter.readDrawings({stamp:adapter.stamp,signal:new AbortController().signal},'D',memory)}
}
// Evidence stays in memory (no file written into the worktree).
async function save(){void evidence; void writeFile}
for(const overlap of [false,true])it(`long numeric strings preserve numeric chart values overlap=${overlap}`,async()=>{
 const value='0'.repeat(300)+'5'
 const bytes=await chartWorkbookBytes(xml(overlap?'Styled!A1':undefined,'Styled!A1'),{cells:[{ref:'A1',v:value}]})
 const h=await open(bytes)
 try{
  const resident=await parseXlsx(await OfficePackage.load(bytes))
  const expected=(resident.sheets[0].drawings![0].content as any).chart.groups[0].series[0].values
  const result=await h.read({set:()=>true}),actual=result.drawings[0].content.chart.groups[0].series[0].values
  evidence[`numeric-${overlap}`]={expected,actual,complete:result.completeness,truncated:result.drawings[0].content.chart.truncated};await save()
  expect(actual).toEqual(expected)
 }finally{await h.source.release()}
})
for (const repeats of [200,10_000]) it(`format expansion is covered by accepted reservation before formatting repeat=${repeats}`,async()=>{
 const text='X'.repeat(255),fmt='0;0;0;'+'@'.repeat(repeats)
 const bytes=await chartWorkbookBytes(xml('Styled!A1','Styled!B1'),{cells:[{ref:'A1',v:text,style:3},{ref:'B1',v:5}],patch:async zip=>{
  const styles=await zip.file('xl/styles.xml')!.async('string')
  zip.file('xl/styles.xml',styles.replace('<cellXfs count="3">',`<numFmts count="1"><numFmt numFmtId="164" formatCode="${fmt}"/></numFmts><cellXfs count="4">`).replace('</cellXfs>','<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'))
 }})
 const h=await open(bytes)
 try{
  const accepted:number[]=[],formatted:any[]=[],raw=render.cellDisplayText
  const spy=vi.spyOn(render,'cellDisplayText').mockImplementation((...args:Parameters<typeof raw>)=>{
   const admitted=accepted.at(-1)??0,out=raw(...args);formatted.push({admitted,bytes:2*out.length});return out
  })
  let result:any
  try{result=await h.read({set:(n:number)=>{const ok=n<=100_000;if(ok)accepted.push(n);return ok}})}finally{spy.mockRestore()}
  evidence[`format-${repeats}`]={accepted,formatted,complete:result.completeness,kept:result.drawings[0]?.content?.chart?.categories[0]?.length,truncated:result.drawings[0]?.content?.chart?.truncated};await save()
  // Adapted: the formatter's output is capped (LIMIT + 1 units) and every transient it
  // reports is admitted before allocation. A format whose parse alone exceeds the
  // allowance is refused with nothing formatted; otherwise every formatted output is
  // covered by an accepted reservation and the kept label is capped and truncated.
  if (repeats > 1000) { expect(result.completeness).toBe('incomplete'); expect(formatted).toEqual([]) }
  else {
   expect(formatted.length).toBeGreaterThan(0)
   expect(formatted.every(v=>v.admitted>=v.bytes)).toBe(true)
   expect(result.drawings[0].content.chart.categories[0]).toHaveLength(255)
   expect(result.drawings[0].content.chart.truncated).toBe(true)
  }
 }finally{await h.source.release()}
})
it('refusing a raw-prefix reservation builds no prefix string',async()=>{
 const bytes=await chartWorkbookBytes(xml('Styled!A1','Styled!B1'),{cells:[{ref:'A1',v:'label text'},{ref:'B1',v:5}]})
 const h=await open(bytes)
 let workspace=0,last=0
 const raw=pages.textPrefix
 try{
  // Adapted: the prefix is now reserved immediately before it is built, so the point
  // workspace is the charge before that reservation (the last two charges before the build).
  const charges:number[]=[]
  const spy=vi.spyOn(pages,'textPrefix').mockImplementation((...args:Parameters<typeof raw>)=>{workspace=charges[charges.length-2];return raw(...args)})
  try{await h.read({set:(n:number)=>{last=n;charges.push(n);return true}})}finally{spy.mockRestore()}
  void last
 }finally{await h.source.release()}
 expect(workspace).toBeGreaterThan(0)
 const fresh=await open(bytes)
 try{
  const built:number[]=[],reserves:any[]=[]
  const spy=vi.spyOn(pages,'textPrefix').mockImplementation((...args:Parameters<typeof raw>)=>{const out=raw(...args);built.push(out.text.length);return out})
  let result:any
  try{result=await fresh.read({set:(n:number)=>{reserves.push({n,accepted:n<=workspace});return n<=workspace}})}finally{spy.mockRestore()}
  evidence.prefix={workspace,built,reserves,complete:result.completeness};await save()
  expect(result.completeness).toBe('incomplete')
  expect(reserves.some(e=>!e.accepted&&e.n===workspace+20)).toBe(true)
  expect(built).toEqual([])
 }finally{await fresh.source.release()}
})
