// @vitest-environment node
import {expect,it} from 'vitest'
import {writeFile} from 'node:fs/promises'
import {chartWorkbookBytes} from './helpers/xlsx-chart-fixture'
import {SourcePageStore,memoryPageBackend} from '../src/xlsx/storage/sourcePages'
import {openImportSession} from '../src/xlsx/storage/pagedImport'
import {createPagedWorkbookSource} from '../src/xlsx/view/pagedSource'
import {columnToLetters} from '../src/xlsx/view/coordinates'
import {MemoryPool} from '../src/xlsx/view/tileCache'
for(const poolClosed of [false,true])it(`an active public chart read preserves the accepted record charge across close, poolClosed=${poolClosed}`,async()=>{
 const xml='<chartSpace><chart><plotArea><barChart><ser><val><numRef><f>Styled!B1</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>'
 const bytes=await chartWorkbookBytes(xml,{cells:Array.from({length:10000},(_,i)=>({ref:`${columnToLetters(i)}1`,v:i+1}))})
 const backend=memoryPageBackend(),store=new SourcePageStore({docId:'chart-close',backend})
 const session=await openImportSession({kind:'blob',blob:new Blob([new Uint8Array(bytes)])},{docId:'chart-close',profile:'D',store,pageRows:1})
 expect((await session.done).kind).toBe('complete')
 const source=createPagedWorkbookSource(session),adapter=source.store('1')!,pool=new MemoryPool({limitBytes:2000000});source.bindMemory!(pool)
 const summary=store.pages('1')[0];await store.page('1',summary.pageId)
 const rawPin=store.pinnedPage.bind(store);let active:Awaited<ReturnType<typeof store.pinnedPage>>|undefined
 store.pinnedPage=async(...args)=>{
  const pending=rawPin(...args);pool.trim(0);const pin=await pending;expect(pin.retained).toBe(false)
  active=pin;const release=pin.release;pin.release=()=>{active=undefined;release()};return pin
 }
 const charge=pool.charge(),other=pool.charge();let scheduled=false,closing:Promise<void>|undefined,workspace:Uint8Array|undefined,observation:any
 try{
  const result=await adapter.readDrawings!({stamp:adapter.stamp,signal:new AbortController().signal},'D',{set:n=>{
   const ok=charge.set(n)
   if(ok&&active&&!scheduled){scheduled=true;queueMicrotask(()=>{
    expect(active).toBeDefined()
    const before=pool.stats(),beforeBytes=active!.bytes,backing=active!.page.rows.buffer.byteLength
    if(poolClosed)pool.close()
    closing=store.close()
    const after=pool.stats(),wanted=pool.limitBytes-after.totalBytes,accepted=other.set(wanted)
    if(accepted)workspace=new Uint8Array(wanted)
    observation={before,beforeBytes,backing,backendHeld:backend.heldBytes,afterBytes:active!.bytes,accepted,heldShare:charge.bytes,stats:pool.stats(),workspace:workspace?.byteLength??0,liveBuffers:backing+(workspace?.byteLength??0)}
   })}
   return ok
  }})
  await closing
  void writeFile; void result
  expect(scheduled).toBe(true);expect(active).toBeUndefined();expect(observation.backendHeld).toBe(0)
  if(poolClosed)expect(observation.accepted).toBe(false)
  expect(observation.liveBuffers).toBeLessThanOrEqual(pool.limitBytes)
 }finally{await closing;charge.release();other.release();await source.release();pool.close();void workspace}
},30000)
