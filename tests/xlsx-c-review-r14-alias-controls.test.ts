// @vitest-environment node
import {expect,it} from 'vitest'
import {SourcePageStore,memoryPageBackend,PageBuilder,K_NUMBER} from '../src/xlsx/storage/sourcePages'
import {MemoryPool} from '../src/xlsx/view/tileCache'
for(const kind of ['sst','page']as const)for(const closedPool of [false,true])it(`${kind} aliased records: closed-pool safety / roomy no-duplicate control, closedPool=${closedPool}`,async()=>{
 const store=new SourcePageStore({docId:'alias-control',backend:memoryPageBackend()})
 if(kind==='sst')await store.commitStrings(0,['漢'.repeat(20000)],1)
 else{const p=new PageBuilder();for(let c=0;c<2000;c++)p.add(0,c,K_NUMBER,0,c,'',false,undefined);await store.commit({sheetId:'1',pageId:'p0',r0:0,r1:0,cellCount:2000,bytes:p.bytes,payload:p.build()})}
 const pool=new MemoryPool({limitBytes:120000});store.bindMemory(pool)
 const pin=()=>kind==='sst'?store.pinnedString('sst',0):store.pinnedPage('1','p0')
 ;(await pin())!.release();const pending=pin();pool.trim(0);const a=(await pending)!
 const caller=pool.charge(),other=pool.charge();expect(caller.set(a.bytes)).toBe(true);const b=(await pin())!
 const ar='record'in a?a.record:a.page,br='record'in b?b.record:b.page;expect(ar).not.toBe(br)
 const av='record'in a?a.record.text:a.page.rows,bv='record'in b?b.record.text:b.page.rows;const backing=bv.buffer.byteLength;expect(av.buffer).toBe(bv.buffer)
 try{
  if(closedPool)pool.close();await store.close()
  const charged=pool.stats().totalBytes
  if(closedPool){a.release();caller.release();const accepted=other.set(pool.limitBytes-pool.stats().totalBytes);expect(accepted).toBe(false)}
  else{expect(backing+a.bytes+b.bytes+60000).toBeLessThan(pool.limitBytes);const accepted=other.set(60000);console.log(JSON.stringify({kind,backing,recordOverhead:a.bytes+b.bytes,charged,wanted:60000,accepted}));expect(accepted).toBe(true)}
 }finally{a.release();b.release();caller.release();other.release();await store.close();pool.close()}
})
