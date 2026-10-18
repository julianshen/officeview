// @vitest-environment node
import {expect,it} from 'vitest'
import {mkdtemp,writeFile,readFile,rm} from 'node:fs/promises'
import {tmpdir} from 'node:os'
import {join} from 'node:path'
import {chartWorkbookBytes} from './helpers/xlsx-chart-fixture'
import {openImportSession} from '../src/xlsx/storage/pagedImport'
import {SourcePageStore,type PageBackend} from '../src/xlsx/storage/sourcePages'
import {createPagedWorkbookSource} from '../src/xlsx/view/pagedSource'
import {MemoryPool} from '../src/xlsx/view/tileCache'
it('the SST chunk stays owned while the current page is pinned',async()=>{
 const dir=await mkdtemp(join(tmpdir(),'r4-current-page-')),file=(k:string)=>join(dir,encodeURIComponent(k))
 const backend:PageBackend={kind:'opfs',put:async(k,b)=>{await writeFile(file(k),b)},get:async k=>{try{return new Uint8Array(await readFile(file(k)))}catch{return undefined}},delete:async k=>{await rm(file(k),{force:true})},removeAll:async()=>{await rm(dir,{recursive:true,force:true})}}
 const store=new SourcePageStore({docId:'r4-page',backend})
 const col=(n:number)=>{let out='';for(n++;n;n=Math.floor((n-1)/26))out=String.fromCharCode(65+(n-1)%26)+out;return out}
 const cells=Array.from({length:10_000},(_,c)=>({ref:`${col(c)}1`,v:c+4}))
 const xml='<chartSpace><chart><plotArea><barChart><ser><cat><strRef><f>Styled!A1</f></strRef></cat><val><numRef><f>Styled!B1</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>'
 const bytes=await chartWorkbookBytes(xml,{cells,patch:async zip=>{
  const path='xl/worksheets/sheet1.xml',sheet=await zip.file(path)!.async('string')
  zip.file(path,sheet.replace(/<c r="A1"[^>]*>[\s\S]*?<\/c>/,'<c r="A1" t="s"><v>0</v></c>'))
  zip.file('xl/sharedStrings.xml',`<sst><si><t>Label</t></si><si><t>${'Z'.repeat(150_000)}</t></si></sst>`)
  const rel='xl/_rels/workbook.xml.rels'
  zip.file(rel,(await zip.file(rel)!.async('string')).replace('</Relationships>','<Relationship Id="rSst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'))
 }})
 const session=await openImportSession({kind:'blob',blob:new Blob([new Uint8Array(bytes)])},{docId:'r4-page',profile:'D',store,pageRows:1})
 expect((await session.done).kind).toBe('complete')
 const source=createPagedWorkbookSource(session),adapter=source.store('1') as any,pool=new MemoryPool({limitBytes:500_000});source.bindMemory!(pool)
 const charge=pool.charge(),observations:any[]=[]
 const rawLoad=(store as any).load.bind(store)
 ;(store as any).load=async(...args:any[])=>{
  const value=await rawLoad(...args)
  if(String(args[0]).startsWith('sst/')){
   const retained=[...(store as any).resident.values()].some((r:any)=>r.value===value)
   const stats=pool.stats(),bytes=value.text.byteLength+value.textStart.byteLength
   observations.push({retained,stringBytes:bytes,pool:stats,totalLiveAtLeast:stats.totalBytes+(retained?0:bytes)})
  }
  return value
 }
 try{
  const result=await adapter.readDrawings({stamp:adapter.stamp,signal:new AbortController().signal},'D',{set:(n:number)=>charge.set(n)})
  void observations
  expect(result.completeness).toBe('complete')
  expect(observations.length).toBeGreaterThan(0)
  expect(observations.every(o=>o.totalLiveAtLeast<=pool.limitBytes)).toBe(true)
 }finally{charge.release();await source.release();pool.close();await rm(dir,{recursive:true,force:true})}
},30_000)
