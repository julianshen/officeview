// @vitest-environment node
import {expect,it} from 'vitest'
import {writeFile} from 'node:fs/promises'
import {SourcePageStore,memoryPageBackend} from '../src/xlsx/storage/sourcePages'
import {MemoryPool} from '../src/xlsx/view/tileCache'
for(const turns of [0,1,2,3,4,5])it(`RAM overlapping leases, microtask turns=${turns}`,async()=>{
 const backend=memoryPageBackend(),store=new SourcePageStore({docId:'ram-overlap',backend})
 await store.commitStrings(0,['漢'.repeat(20000)],1);await store.commitStrings(1,['字'.repeat(20000)],1)
 const held=backend.heldBytes,pool=new MemoryPool({limitBytes:held+3000});store.bindMemory(pool)
 const a=store.pinnedString('sst',0);for(let i=0;i<turns;i++)await Promise.resolve();const b=store.pinnedString('sst',1)
 const result=await Promise.all([a,b]);const charge=pool.charge()
 try{
  const unretained=result.find(x=>x&&!x.retained),snapshot=pool.stats();let full:boolean|undefined,incremental:boolean|undefined
  if(unretained){expect(unretained.record.text.buffer).toBe((await backend.get(unretained.key))!.buffer);full=charge.set(unretained.bytes);charge.set(0);incremental=charge.set(1280)}
  void writeFile; void snapshot
  if(unretained){expect(incremental).toBe(true);expect(full).toBe(true);expect(unretained.bytes).toBe(1280)}
  else expect(result.every(x=>x?.retained)).toBe(true)
 }finally{charge.release();for(const x of result)x?.release();await store.close();pool.close()}
})
