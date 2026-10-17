// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { mkdtemp, writeFile, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { chartWorkbookBytes } from './helpers/xlsx-chart-fixture'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { SourcePageStore, payloadBytes, type PageBackend } from '../src/xlsx/storage/sourcePages'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { MemoryPool } from '../src/xlsx/view/tileCache'
import * as render from '../src/xlsx/render'
const evidence:Record<string,unknown>={}
const model=(result:any)=>result.drawings[0].content.chart
const xml=(cat:string|undefined,val:string)=>`<chartSpace><chart><plotArea><barChart><ser>${cat?`<cat><numRef><f>${cat}</f></numRef></cat>`:''}<val><numRef><f>${val}</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>`
async function open(bytes:Uint8Array, opts:object={}) {
 const session=await openImportSession({kind:'blob',blob:new Blob([new Uint8Array(bytes)])},{docId:'r3',profile:'D',...opts})
 expect((await session.done).kind).toBe('complete')
 const source=createPagedWorkbookSource(session), adapter=source.store('1') as any
 return {session,source,adapter,read:(memory?:any)=>adapter.readDrawings({stamp:adapter.stamp,signal:new AbortController().signal},'D',memory)}
}
// Evidence is kept in memory (no file written into the worktree).
async function save(){void evidence}
for (const storage of ['inline','sst'] as const) it(`numeric series ${storage} strings are admitted before materialization`,async()=>{
 const long='0'.repeat(100_000)+'5'
 const h=await open(await chartWorkbookBytes(xml(undefined,'Styled!A1'),{cells:[{ref:'A1',v:long}],patch: storage==='sst'?async zip=>{
  const name='xl/worksheets/sheet1.xml', sheet=await zip.file(name)!.async('string')
  zip.file(name,sheet.replace(/<c r="A1"[^>]*>[\s\S]*?<\/c>/,'<c r="A1" t="s"><v>0</v></c>'))
  zip.file('xl/sharedStrings.xml',`<sst><si><t>${long}</t></si></sst>`)
  const rel='xl/_rels/workbook.xml.rels'
  zip.file(rel,(await zip.file(rel)!.async('string')).replace('</Relationships>','<Relationship Id="rSst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'))
 }:undefined}))
 try {
  const charges:number[]=[], materialized:Array<{bytes:number;admitted:number}>=[]
  const raw=h.adapter.valueOf.bind(h.adapter)
  vi.spyOn(h.adapter,'valueOf').mockImplementation(async(...args:any[])=>{
   const admitted=charges.at(-1)??0, value=await raw(...args)
   if(typeof value==='string')materialized.push({bytes:value.length*2,admitted})
   return value
  })
  const result=await h.read({set:(n:number)=>{charges.push(n);return n<=100_000}})
  evidence[storage]={charges,materialized,completeness:result.completeness,values:result.drawings[0]?.content?.chart?.groups[0]?.series[0]?.values};await save()
  expect(materialized.every(x=>x.admitted>=x.bytes)).toBe(true)
 }finally{await h.source.release()}
})
// Adapted: the label's format is bounded before formatting (format text plus
// value), and formatting follows an ACCEPTED reservation covering it; the kept
// label is capped at 255 characters (truncated). Amended (255-unit format-code
// cap, user decision 2026-10-17): Excel cannot author a code over 255 UTF-16
// units, so the 100,000-character literal code is now General at style parse
// and there is no long format left to refuse; a 253-unit literal code with a
// six-digit value still exercises the admitted, capped, truncated label.
it('formatted numeric category length is admitted before formatting',async()=>{
 const workbookWith=(code:string)=>chartWorkbookBytes(xml('Styled!A1','Styled!B1'),{cells:[{ref:'A1',v:123456,style:3},{ref:'B1',v:5}],patch:async zip=>{
  const styles=await zip.file('xl/styles.xml')!.async('string')
  zip.file('xl/styles.xml',styles.replace('<cellXfs count="3">',`<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;${code}&quot;0"/></numFmts><cellXfs count="4">`).replace('</cellXfs>','<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'))
 }})
 const raw=render.cellDisplayText
 for (const overlong of [true,false]) {
  const h=await open(await workbookWith('L'.repeat(overlong?100_000:250)))
  try{
   const accepted:number[]=[],formatted:Array<{bytes:number;admitted:number}>=[]
   const spy=vi.spyOn(render,'cellDisplayText').mockImplementation((...args:Parameters<typeof raw>)=>{
    const admitted=accepted.at(-1)??0, out=raw(...args);formatted.push({bytes:out.length*2,admitted});return out
   })
   const result=await h.read({set:(n:number)=>{const ok=n<=100_000;if(ok)accepted.push(n);return ok}})
   spy.mockRestore()
   evidence[overlong?'formattedOverlong':'formatted']={accepted,formatted,completeness:result.completeness}
   expect(result.completeness).toBe('complete')
   expect(formatted.length).toBeGreaterThan(0)
   expect(formatted.every(x=>x.admitted>=x.bytes)).toBe(true)
   // The formatter's output is capped (r4): it never builds more than LIMIT + 1 units.
   expect(formatted.every(x=>x.bytes<=2*256)).toBe(true)
   if(overlong){expect(model(result).categories[0]).toBe('123456');expect(model(result).truncated).toBeFalsy()}
   else{expect(model(result).categories[0]).toBe('L'.repeat(250)+'12345');expect(model(result).truncated).toBe(true)}
  }finally{await h.source.release()}
 }
})
it('decoded chart pages remain charged until the chart releases its working references',async()=>{
 const dir=await mkdtemp(join(tmpdir(),'r3-pages-'))
 const file=(key:string)=>join(dir,encodeURIComponent(key))
 const backend:PageBackend={kind:'opfs',put:async(k,b)=>{await writeFile(file(k),b)},get:async k=>{try{return new Uint8Array(await readFile(file(k)))}catch{return undefined}},delete:async k=>{await rm(file(k),{force:true})},removeAll:async()=>{await rm(dir,{recursive:true,force:true})}}
 const store=new SourcePageStore({docId:'r3',backend})
 const col=(n:number)=>{let out='';for(n++;n;n=Math.floor((n-1)/26))out=String.fromCharCode(65+(n-1)%26)+out;return out}
 const cells=Array.from({length:100},(_,row)=>Array.from({length:1000},(_,c)=>({ref:`${col(c)}${row+1}`,v:row+c}))).flat()
 const h=await open(await chartWorkbookBytes(xml(undefined,'Styled!A1:A100'),{cells}),{store,pageRows:1})
 const pool=new MemoryPool({limitBytes:500_000});h.source.bindMemory!(pool)
 const charge=pool.charge(), charges:number[]=[], observations:any[]=[], loadedIds=new Set<string>(), materializedIds=new Set<string>()
 let decodedBytes=0
 const rawPage=store.page.bind(store),rawValue=h.adapter.valueOf.bind(h.adapter)
 store.page=async(s,id)=>{const p=await rawPage(s,id);loadedIds.add(id);decodedBytes+=payloadBytes(p);return p}
 h.adapter.valueOf=async(p:any,i:number)=>{
  materializedIds.add(String(p.rows[i]));observations.push({loadedPages:loadedIds.size,decodedBytes,sourceResident:store.stats.residentBytes,pool:pool.stats(),charged:charges.at(-1),row:p.rows[i]});return rawValue(p,i)
 }
 try{
  const result=await h.read({set:(n:number)=>{charges.push(n);return charge.set(n)}})
  evidence.pages={charges,first:observations[0],last:observations.at(-1),materializedRows:materializedIds.size,stats:store.stats,completeness:result.completeness,values:result.drawings[0]?.content?.chart?.groups[0]?.series[0]?.values};await save()
  expect(observations.length).toBe(100)
  // The first materialization proves all 100 decoded payloads are retained in
  // the chart's needs map. Typed payload bytes alone exceed the whole pool.
  expect(observations[0].decodedBytes).toBeLessThanOrEqual(pool.limitBytes)
 }finally{charge.release();await h.source.release();pool.close();await rm(dir,{recursive:true,force:true})}
},30_000)
