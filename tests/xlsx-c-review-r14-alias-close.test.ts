// @vitest-environment node
import {expect,it} from 'vitest'
import {SourcePageStore,memoryPageBackend,PageBuilder,K_NUMBER} from '../src/xlsx/storage/sourcePages'
import {MemoryPool} from '../src/xlsx/view/tileCache'
for(const kind of ['sst','page']as const)for(const terminal of ['close','abandon']as const)for(const order of ['a-first','b-first']as const)it(`${kind} distinct decoded records sharing one RAM backing survive ${terminal} ${order}`,async()=>{
 const backend=memoryPageBackend(),store=new SourcePageStore({docId:'alias-close',backend})
 if(kind==='sst')await store.commitStrings(0,['漢'.repeat(20000)],1)
 else{const b=new PageBuilder();for(let c=0;c<2000;c++)b.add(0,c,K_NUMBER,0,c,'',false,undefined);await store.commit({sheetId:'1',pageId:'p0',r0:0,r1:0,cellCount:2000,bytes:b.bytes,payload:b.build()})}
 const pool=new MemoryPool({limitBytes:60000});store.bindMemory(pool)
 const pin=()=>kind==='sst'?store.pinnedString('sst',0):store.pinnedPage('1','p0')
 await pin().then(p=>p!.release())
 const pending=pin();pool.trim(0);const a=(await pending)!
 expect(a.retained).toBe(false)
 const caller=pool.charge(),other=pool.charge();expect(caller.set(a.bytes)).toBe(true)
 const b=(await pin())!;expect(b.retained).toBe(true)
 const ar='record'in a?a.record:a.page,br='record'in b?b.record:b.page
 const ab=('record'in a?a.record.text:a.page.rows).buffer,bb=('record'in b?b.record.text:b.page.rows).buffer
 expect(ar).not.toBe(br);expect(ab).toBe(bb)
 const backing=bb.byteLength
 try{
  if(terminal==='close')await store.close();else await store.abandon('alias-close')
  const chargedAfterClose=pool.stats().totalBytes
  if(order==='a-first'){a.release();caller.release()}else b.release()
  const chargedAfterRelease=pool.stats().totalBytes,free=pool.limitBytes-chargedAfterRelease,accepted=other.set(free)
  const workspace=accepted?new Uint8Array(free):undefined
  console.log(JSON.stringify({kind,terminal,order,backing,chargedAfterClose,chargedAfterRelease,free,accepted,liveBuffers:backing+(workspace?.byteLength??0)}))
  expect(backing+(workspace?.byteLength??0)).toBeLessThanOrEqual(pool.limitBytes)
 }finally{a.release();b.release();caller.release();other.release();await store.close();pool.close()}
})
