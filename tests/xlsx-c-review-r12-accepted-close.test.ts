// @vitest-environment node
import {expect,it} from 'vitest'
import {writeFile} from 'node:fs/promises'
import {SourcePageStore,memoryPageBackend,PageBuilder,K_NUMBER} from '../src/xlsx/storage/sourcePages'
import {MemoryPool} from '../src/xlsx/view/tileCache'
for(const kind of ['sst','page']as const)for(const terminalPool of [false,true])it(`accepted ${kind} lease survives close without exposing uncharged backing, poolClosed=${terminalPool}`,async()=>{
 const backend=memoryPageBackend(),store=new SourcePageStore({docId:'accepted-close',backend})
 if(kind==='sst')await store.commitStrings(0,['漢'.repeat(20000)],1)
 else{const b=new PageBuilder();for(let c=0;c<2000;c++)b.add(0,c,K_NUMBER,0,c,'',false,undefined);await store.commit({sheetId:'1',pageId:'p0',r0:0,r1:0,cellCount:2000,bytes:b.bytes,payload:b.build()})}
 const pool=new MemoryPool({limitBytes:60000});store.bindMemory(pool)
 if(kind==='sst')await store.sharedStringRecord(0);else await store.page('1','p0')
 const pending=kind==='sst'?store.pinnedString('sst',0):store.pinnedPage('1','p0');pool.trim(0);const lease=(await pending)!
 const caller=pool.charge(),other=pool.charge()
 try{
  expect(lease.retained).toBe(false);const beforeBytes=lease.bytes,backing='record'in lease?lease.record.text.buffer.byteLength:lease.page.rows.buffer.byteLength
  expect(caller.set(lease.bytes)).toBe(true);expect(pool.stats().totalBytes).toBeGreaterThanOrEqual(backing+beforeBytes)
  if(terminalPool)pool.close()
  await store.close();expect(backend.heldBytes).toBe(0)
  const afterBytes=lease.bytes,free=pool.limitBytes-pool.stats().totalBytes,accepted=other.set(free)
  const workspace=accepted?new Uint8Array(free):undefined;const liveBuffers=backing+(workspace?.buffer.byteLength??0)
  void writeFile; void afterBytes
  if(terminalPool)expect(accepted).toBe(false)
  expect(liveBuffers).toBeLessThanOrEqual(pool.limitBytes)
 }finally{lease.release();caller.release();other.release();await store.close();pool.close()}
})
