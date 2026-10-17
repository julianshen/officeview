// @vitest-environment node
import {expect,it} from 'vitest'
import {writeFile} from 'node:fs/promises'
import {SourcePageStore,memoryPageBackend,PageBuilder,K_NUMBER} from '../src/xlsx/storage/sourcePages'
import {MemoryPool} from '../src/xlsx/view/tileCache'
for(const kind of ['sst','page'] as const)it(`a ${kind} lease crossing close either cancels or charges its now-unowned buffer`,async()=>{
 const backend=memoryPageBackend(),store=new SourcePageStore({docId:'close-handoff',backend})
 if(kind==='sst')await store.commitStrings(0,['漢'.repeat(20000)],1)
 else{const b=new PageBuilder();for(let c=0;c<2000;c++)b.add(0,c,K_NUMBER,0,c,'',false,undefined);await store.commit({sheetId:'1',pageId:'p0',r0:0,r1:0,cellCount:2000,bytes:b.bytes,payload:b.build()})}
 const pool=new MemoryPool({limitBytes:60000});store.bindMemory(pool)
 if(kind==='sst')await store.sharedStringRecord(0);else await store.page('1','p0')
 const pending=kind==='sst'?store.pinnedString('sst',0):store.pinnedPage('1','p0')
 const settled=pending.then(value=>({value}),error=>({error}))
 await store.close()
 const result=await settled
 try{
  expect(backend.heldBytes).toBe(0)
  if('error'in result){expect(result.error.code).toBe('cancelled');return}
  const lease=result.value!;expect(lease.retained).toBe(false)
  const backing='record'in lease?lease.record.text.buffer.byteLength:lease.page.rows.buffer.byteLength
  const expected=backing+(kind==='sst'?1280:2048)
  const charge=pool.charge(),replacement=pool.charge();expect(charge.set(lease.bytes)).toBe(true)
  const replacementBytes=pool.limitBytes-lease.bytes;expect(replacement.set(replacementBytes)).toBe(true)
  const workspace=new Uint8Array(replacementBytes);workspace[0]=1
  const knownLiveBuffers=backing+workspace.buffer.byteLength
  void writeFile
  try{expect(knownLiveBuffers).toBeLessThanOrEqual(pool.limitBytes);expect(lease.bytes).toBeGreaterThanOrEqual(expected)}finally{lease.release();charge.release();replacement.release()}
 }finally{await store.close();pool.close()}
})

for(const kind of ['sst','page'] as const)it(`a ${kind} lease whose cache entry is evicted mid-acquisition keeps the RAM deduction while RAM owns the buffer, and charges it in full once RAM ownership ends`,async()=>{
 const backend=memoryPageBackend(),store=new SourcePageStore({docId:'evict-handoff',backend})
 if(kind==='sst')await store.commitStrings(0,['漢'.repeat(20000)],1)
 else{const b=new PageBuilder();for(let c=0;c<2000;c++)b.add(0,c,K_NUMBER,0,c,'',false,undefined);await store.commit({sheetId:'1',pageId:'p0',r0:0,r1:0,cellCount:2000,bytes:b.bytes,payload:b.build()})}
 const pool=new MemoryPool({limitBytes:200000});store.bindMemory(pool)
 if(kind==='sst')await store.sharedStringRecord(0);else await store.page('1','p0')
 const pending=kind==='sst'?store.pinnedString('sst',0):store.pinnedPage('1','p0')
 // The decoded cache entry is evicted before the lease is acquired (pressure from another owner).
 pool.trim(0)
 const lease=(await pending)!
 try{
  expect(lease.retained).toBe(false)
  const backing='record'in lease?lease.record.text.buffer.byteLength:lease.page.rows.buffer.byteLength
  const wrapper=kind==='sst'?1280:2048
  // RAM still owns (and the pool still charges) the buffer: only the wrapper is new.
  expect(backend.heldBytes).toBeGreaterThanOrEqual(backing)
  expect(pool.stats().totalBytes).toBeGreaterThanOrEqual(backing)
  expect(lease.bytes).toBe(wrapper)
  // RAM ownership ends: the holder's charge now includes the backing buffer.
  await store.close()
  expect(backend.heldBytes).toBe(0)
  expect(lease.bytes).toBe(backing+wrapper)
 }finally{lease.release();await store.close();pool.close()}
})
