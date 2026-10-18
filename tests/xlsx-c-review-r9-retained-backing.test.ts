// @vitest-environment node
import {expect,it} from 'vitest'
import {writeFile,mkdtemp,readFile,rm} from 'node:fs/promises'
import {join} from 'node:path'
import {tmpdir} from 'node:os'
import {xlsxBytes} from './helpers/xlsx-package'
import {SourcePageStore,type PageBackend,DECODED_PAGE_OVERHEAD_BYTES,payloadBytes} from '../src/xlsx/storage/sourcePages'
import {openImportSession} from '../src/xlsx/storage/pagedImport'
import {createPagedWorkbookSource} from '../src/xlsx/view/pagedSource'
import {MemoryPool} from '../src/xlsx/view/tileCache'

it('cached formula pages own their entire retained backing buffer across the next read',async()=>{
 const dir=await mkdtemp(join(tmpdir(),'r9-cache-')),file=(k:string)=>join(dir,encodeURIComponent(k))
 let pool:MemoryPool|undefined;const reads:any[]=[]
 const backend:PageBackend={kind:'opfs',put:async(k,b)=>{await writeFile(file(k),b)},get:async k=>{
  const bytes=new Uint8Array(await readFile(file(k)))
  if(pool){
   const cached=[...(store as any).resident.entries()].map(([key,res]:any)=>{
    const p=res.value,views=[p.rows,p.cols,p.kinds,p.flags,p.styles,p.nums,p.text,p.textStart],buffers=new Set<ArrayBuffer>(views.map(v=>v.buffer)),lane=views.reduce((n,v)=>n+v.byteLength,0),backing=[...buffers].reduce((n,b)=>n+b.byteLength,0),metaModel=payloadBytes(p)-lane
    return {key,charged:res.bytes,lane,backing,metaModel,lower:backing+metaModel+DECODED_PAGE_OVERHEAD_BYTES,formulaUnits:[...p.meta.values()].reduce((n:number,v:any)=>n+(v.formula?.length??0),0)}
   })
   const stats=pool.stats(),missing=cached.reduce((n,r)=>n+Math.max(0,r.lower-r.charged),0)
   reads.push({key:k,reading:store.readingBytes,incoming:bytes.byteLength,cached,stats,missing,liveAtLeast:stats.totalBytes+missing})
  }
  return bytes
 },delete:async k=>{await rm(file(k),{force:true})},removeAll:async()=>{await rm(dir,{recursive:true,force:true})}}
 const store=new SourcePageStore({docId:'r9-cache',backend})
 const formula='N(&quot;'+'漢'.repeat(500)+'&quot;)+1'
 const bytes=xlsxBytes({sheetData:Array.from({length:400},(_,i)=>`<row r="${i+1}"><c r="A${i+1}"><f>${formula}</f><v>1</v></c></row>`).join('')})
 const session=await openImportSession({kind:'blob',blob:new Blob([new Uint8Array(bytes)])},{docId:'r9-cache',profile:'D',store,pageRows:200})
 expect((await session.done).kind).toBe('complete')
 const source=createPagedWorkbookSource(session),adapter=source.store('1')!
 pool=new MemoryPool({limitBytes:Number(process.env.IV_REVIEW_POOL_LIMIT??1_000_000)});source.bindMemory!(pool)
 const rawParse=JSON.parse,parses:any[]=[]
 JSON.parse=function(text:any,...args:any[]){
  const value=Reflect.apply(rawParse,JSON,[text,...args])
  if(typeof text==='string'&&text.startsWith('[[0,{"formula":')){
   const last=reads.at(-1),cached=last.cached
   const parsedFormulaBytes=2*value.reduce((n:number,entry:any)=>n+(entry[1].formula?.length??0),0)
   const cachedKnown=cached.reduce((n:number,p:any)=>n+p.backing+2*p.formulaUnits,0)
   parses.push({key:last.key,reading:store.readingBytes,pool:pool!.stats(),knownPayload:cachedKnown+last.incoming+2*text.length+parsedFormulaBytes,cachedKnown,incoming:last.incoming,jsonBytes:2*text.length,parsedFormulaBytes,modelTotal:pool!.stats().totalBytes+last.missing})
  }
  return value
 } as typeof JSON.parse
 let pin:Awaited<ReturnType<typeof store.pinnedPage>>|undefined
 try{
  const results=[]
  results.push(await adapter.readCell({row:0,col:0},{stamp:adapter.stamp,signal:new AbortController().signal}))
  pin=await store.pinnedPage('1','p0');expect(pin.retained).toBe(true)
  results.push(await adapter.readCell({row:200,col:0},{stamp:adapter.stamp,signal:new AbortController().signal}))
    expect(results[0]).toMatchObject({kind:'cell',cell:{value:1}})
  expect(pin.page.meta.size).toBe(200)
  if(results[1].kind==='error'){expect(results[1]).toMatchObject({code:'resource-limit'});expect(parses).toHaveLength(1)}
  else{expect(results[1]).toMatchObject({kind:'cell',cell:{value:1}});expect(reads.some(r=>r.cached.some((p:any)=>p.metaModel>200000))).toBe(true);expect(parses).toHaveLength(2)}
  expect(parses.every(r=>r.knownPayload<=pool!.limitBytes)).toBe(true)
 }finally{JSON.parse=rawParse;pin?.release();await source.release();pool.close();await rm(dir,{recursive:true,force:true})}
})
