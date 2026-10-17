// @vitest-environment node
import {expect,it} from 'vitest'
import {writeFile,mkdtemp,readFile,rm} from 'node:fs/promises'
import {join} from 'node:path'
import {tmpdir} from 'node:os'
import {xlsxBytes} from './helpers/xlsx-package'
import {SourcePageStore,memoryPageBackend,type PageBackend} from '../src/xlsx/storage/sourcePages'
import {openImportSession} from '../src/xlsx/storage/pagedImport'
import {createPagedWorkbookSource} from '../src/xlsx/view/pagedSource'
import {MemoryPool} from '../src/xlsx/view/tileCache'

const input=(bytes:Uint8Array)=>({kind:'blob' as const,blob:new Blob([new Uint8Array(bytes)])})
it('RAM source reads do not reserve a second copy of their already-owned encoded page',async()=>{
 const bytes=xlsxBytes({sheetData:'<row r="1">'+Array.from({length:2000},(_,i)=>{let n=i+1,col='';for(;n;n=Math.floor((n-1)/26))col=String.fromCharCode(65+(n-1)%26)+col;return `<c r="${col}1"><v>${i}</v></c>`}).join('')+'</row>'})
 const backend=memoryPageBackend(),store=new SourcePageStore({docId:'ram-read',backend})
 const session=await openImportSession(input(bytes),{docId:'ram-read',profile:'D',store,pageRows:1});expect((await session.done).kind).toBe('complete')
 const source=createPagedWorkbookSource(session),adapter=source.store('1')!,held=backend.heldBytes!
 const pool=new MemoryPool({limitBytes:held+10000});source.bindMemory!(pool)
 try{
  const before=pool.stats(),result=await adapter.readCell({row:0,col:1999},{stamp:adapter.stamp,signal:new AbortController().signal})
  void writeFile; void before
  expect(result).toMatchObject({kind:'cell',cell:{value:1999}})
  expect(pool.stats().totalBytes).toBeLessThanOrEqual(pool.limitBytes)
 }finally{await source.release();pool.close()}
})

// Adapted: the read is now admitted for its exact decode cost (recorded at write:
// the metadata JSON string and parsed objects). At the reviewer's tight pool
// (stored + 10,000 B) that cost cannot fit, so the read is refused with a typed
// resource result before the page is read or parsed; with room, the parse happens
// inside an admitted reservation that covers the incoming buffer and the JSON string.
for (const roomy of [false,true]) it(`formula metadata reconstruction is covered before JSON.parse (roomy=${roomy})`,async()=>{
 const dir=await mkdtemp(join(tmpdir(),'r8-formula-')),file=(k:string)=>join(dir,encodeURIComponent(k))
 let incomingBytes=0
 const backend:PageBackend={kind:'opfs',put:async(k,b)=>{await writeFile(file(k),b)},get:async k=>{try{const bytes=new Uint8Array(await readFile(file(k)));incomingBytes=bytes.byteLength;return bytes}catch{return undefined}},delete:async k=>{await rm(file(k),{force:true})},removeAll:async()=>{await rm(dir,{recursive:true,force:true})}}
 const store=new SourcePageStore({docId:'formula-read',backend})
 const formula='N(&quot;'+ '漢'.repeat(500)+'&quot;)+1'
 const bytes=xlsxBytes({sheetData:Array.from({length:200},(_,i)=>`<row r="${i+1}"><c r="A${i+1}"><f>${formula}</f><v>1</v></c></row>`).join('')})
 const session=await openImportSession(input(bytes),{docId:'formula-read',profile:'D',store,pageRows:200});expect((await session.done).kind).toBe('complete')
 const source=createPagedWorkbookSource(session),adapter=source.store('1')!,stored=store.pages('1')[0].bytes,pool=new MemoryPool({limitBytes:roomy?4*stored+100000:stored+10000});source.bindMemory!(pool)
 const rawParse=JSON.parse,observations:any[]=[]
 JSON.parse=function(text:any,...args:any[]){
  if(typeof text==='string'&&text.startsWith('[[0,{"formula":'))observations.push({incomingBytes,jsonUnits:text.length,knownPayloadBytes:incomingBytes+2*text.length,pool:pool.stats(),reading:(store as any).readingBytes})
  return Reflect.apply(rawParse,JSON,[text,...args])
 } as typeof JSON.parse
 try{
  const result=await adapter.readCell({row:0,col:0},{stamp:adapter.stamp,signal:new AbortController().signal})
  if(!roomy){expect(result).toMatchObject({kind:'error',code:'resource-limit'});expect(observations).toEqual([])}
  else{
   expect(result).toMatchObject({kind:'cell'})
   expect(observations.length).toBe(1)
   expect(observations.every(o=>o.knownPayloadBytes<=pool.limitBytes)).toBe(true)
   // The reservation in force at JSON.parse covers the incoming buffer and the JSON string.
   expect(observations.every(o=>o.reading>=o.knownPayloadBytes)).toBe(true)
  }
 }finally{JSON.parse=rawParse;await source.release();pool.close();await rm(dir,{recursive:true,force:true})}
})
