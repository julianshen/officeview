// @vitest-environment node
import {expect,it} from 'vitest'
import {SourcePageStore,memoryPageBackend} from '../src/xlsx/storage/sourcePages'
import {MemoryPool} from '../src/xlsx/view/tileCache'
it('closing shared retained leases charges their one record once',async()=>{
 const store=new SourcePageStore({docId:'shared-wide',backend:memoryPageBackend()})
 await store.commitStrings(0,['漢'.repeat(20000)],1)
 const pool=new MemoryPool({limitBytes:120000});store.bindMemory(pool)
 const a=(await store.pinnedString('sst',0))!,b=(await store.pinnedString('sst',0))!,other=pool.charge()
 expect(a.record).toBe(b.record)
 try {
  await store.close()
  const accepted=other.set(60000)
  console.log(JSON.stringify({sameRecord:a.record===b.record,backing:a.record.text.buffer.byteLength,chargedBeforeOther:pool.stats().totalBytes,accepted}))
  expect(accepted).toBe(true)
 } finally {a.release();b.release();other.release();await store.close();pool.close()}
})
