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
// Adapted: the store now reserves the incoming buffer in the pool BEFORE the read
// (SourcePageStore.readingBytes), so the pool total already counts that buffer once;
// the original formula (pool total + incoming) would count it twice.
it('disk string-record reads reserve the incoming buffer beside cached records',{timeout:30000},async()=>{
 const dir=await mkdtemp(join(tmpdir(),'r6-string-handoff-')),file=(k:string)=>join(dir,encodeURIComponent(k))
 let poolForReads:MemoryPool|undefined;const readPeaks:any[]=[]
 const backend:PageBackend={kind:'opfs',put:async(k,b)=>{await writeFile(file(k),b)},get:async k=>{try{const bytes=new Uint8Array(await readFile(file(k)));if(poolForReads&&k.startsWith('sst/')){const stats=poolForReads.stats();const cached=[...(store as any).resident.entries()].filter(([key]:any)=>key.startsWith('sst/')).map(([key,held]:any)=>({key,typedBytes:held.value.text.byteLength+held.value.textStart.byteLength,sharesIncoming:held.value.text.buffer===bytes.buffer}));readPeaks.push({key:k,bytes:bytes.byteLength,cached,pool:stats,distinctPayloadBytes:bytes.byteLength+cached.filter((o:any)=>!o.sharesIncoming).reduce((n:number,o:any)=>n+o.typedBytes,0),liveAtLeast:stats.totalBytes+Math.max(0,bytes.byteLength-((store as any).readingBytes??0)),reserved:(store as any).readingBytes??0})}return bytes}catch{return undefined}},delete:async k=>{await rm(file(k),{force:true})},removeAll:async()=>{await rm(dir,{recursive:true,force:true})}}
 const store=new SourcePageStore({docId:'r6-handoff',backend})
 const xml='<chartSpace><chart><plotArea><barChart><ser><cat><strRef><f>Styled!A1:A3</f></strRef></cat><val><numRef><f>Styled!B1:B3</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>'
 const bytes=await chartWorkbookBytes(xml,{cells:[{ref:'A1',v:'A'},{ref:'A2',v:'B'},{ref:'A3',v:'C'},{ref:'B1',v:1},{ref:'B2',v:2},{ref:'B3',v:3}],patch:async zip=>{
  const path='xl/worksheets/sheet1.xml';let sheet=await zip.file(path)!.async('string')
  for(let i=0;i<3;i++)sheet=sheet.replace(new RegExp(`<c r="A${i+1}"[^>]*>[\\s\\S]*?<\\/c>`),`<c r="A${i+1}" t="s"><v>${i*4096}</v></c>`)
  zip.file(path,sheet)
  zip.file('xl/sharedStrings.xml','<sst>'+Array.from({length:12288},(_,i)=>`<si><t>${i%4096===0?String.fromCharCode(65+i/4096):'Z'.repeat(37)}</t></si>`).join('')+'</sst>')
  const rel='xl/_rels/workbook.xml.rels';zip.file(rel,(await zip.file(rel)!.async('string')).replace('</Relationships>','<Relationship Id="rSst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'))
 }})
 const session=await openImportSession({kind:'blob',blob:new Blob([new Uint8Array(bytes)])},{docId:'r6-handoff',profile:'D',store,pageRows:3})
 expect((await session.done).kind).toBe('complete')
 const source=createPagedWorkbookSource(session),adapter=source.store('1') as any,pool=new MemoryPool({limitBytes:500000});source.bindMemory!(pool);poolForReads=pool
 const charge=pool.charge(),active=new Map<string,{bytes:number;retained:boolean}>(),loads:any[]=[],admissions:any[]=[],pins:any[]=[]
 const rawPin=store.pinnedString?.bind(store),rawLoad=(store as any).load.bind(store)
 ;(store as any).load=async(...args:any[])=>{
  const value=await rawLoad(...args)
  if(String(args[0]).startsWith('sst/')){
   const retained=[...(store as any).resident.values()].some((r:any)=>r.value===value)
   const bytes=value.text.byteLength+value.textStart.byteLength,unowned=[...active.values()].filter(v=>!v.retained).reduce((n,v)=>n+v.bytes,0)
   const stats=pool.stats();loads.push({key:args[0],retained,bytes,active:[...active],stats,liveAtLeast:stats.totalBytes+unowned+(retained?0:bytes)})
  }
  return value
 }
 if(rawPin)store.pinnedString=async(...args)=>{
  const next=await rawPin(...args);if(!next)return next
  pins.push({key:next.key,retained:next.retained,bytes:next.bytes});active.set(next.key,{retained:next.retained,bytes:next.bytes})
  return {...next,release:()=>{next.release();active.delete(next.key)}}
 }
 try{
  const result=await adapter.readDrawings({stamp:adapter.stamp,signal:new AbortController().signal},'D',{set:(n:number)=>{
   const ok=charge.set(n),unowned=[...active.values()].filter(v=>!v.retained).reduce((n,v)=>n+v.bytes,0)
   if(unowned)admissions.push({requested:n,accepted:ok,held:charge.bytes,unowned,stats:pool.stats()})
   return ok
  }})
  void writeFile; void result
  expect(readPeaks.length).toBeGreaterThan(1);expect(readPeaks.every(o=>o.reserved>=o.bytes)).toBe(true);expect(readPeaks.every(o=>o.liveAtLeast<=pool.limitBytes)).toBe(true)
  expect(loads.length).toBeGreaterThan(1)
  expect(loads.every(o=>o.liveAtLeast<=pool.limitBytes)).toBe(true)
  expect(admissions.every(o=>!o.accepted||o.held>=o.unowned)).toBe(true)
 }finally{charge.release();await source.release();pool.close();await rm(dir,{recursive:true,force:true})}
})
