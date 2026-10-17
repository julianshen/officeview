// @vitest-environment node
import { expect, it } from 'vitest'
import { SourcePageStore, memoryPageBackend, PageBuilder, K_NUMBER } from '../src/xlsx/storage/sourcePages'
import { MemoryPool } from '../src/xlsx/view/tileCache'
for (const kind of ['sst', 'page'] as const) for (const terminal of ['close', 'abandon'] as const) it(`${kind} two leases survive ${terminal} until LAST release`, async () => {
 const backend = memoryPageBackend(), store = new SourcePageStore({ docId: 'multiple', backend })
 if (kind === 'sst') await store.commitStrings(0, ['漢'.repeat(20000)], 1)
 else { const b = new PageBuilder(); for(let c=0;c<2000;c++) b.add(0,c,K_NUMBER,0,c,'',false,undefined); await store.commit({sheetId:'1',pageId:'p0',r0:0,r1:0,cellCount:2000,bytes:b.bytes,payload:b.build()}) }
 const pool = new MemoryPool({ limitBytes: 60000 }); store.bindMemory(pool)
 const pin = () => kind === 'sst' ? store.pinnedString('sst',0) : store.pinnedPage('1','p0')
 const a = (await pin())!, b = (await pin())!; expect(a.retained).toBe(true); expect(b.retained).toBe(true)
 const backing = 'record' in b ? b.record.text.buffer.byteLength : b.page.rows.buffer.byteLength
 const other = pool.charge()
 try {
  if(terminal==='close') await store.close(); else await store.abandon('multiple')
  const beforeRelease=pool.stats().totalBytes; a.release()
  const afterRelease=pool.stats().totalBytes, free = pool.limitBytes-afterRelease
  const accepted=other.set(free), workspace=accepted ? new Uint8Array(free) : undefined
  console.log(JSON.stringify({kind,terminal,backing,beforeRelease,afterRelease,free,accepted,liveBuffers:backing+(workspace?.byteLength??0)}))
  expect(backing+(workspace?.byteLength??0)).toBeLessThanOrEqual(pool.limitBytes)
 } finally { a.release(); b.release(); other.release(); await store.close(); pool.close() }
})

// Both release orders, a roomy no-duplicate control and an explicit pool refusal.
for (const order of ['a-first', 'b-first'] as const) it(`two leases on one record keep ONE closing charge until the last release (${order})`, async () => {
 const backend = memoryPageBackend(), store = new SourcePageStore({ docId: 'refcount', backend })
 await store.commitStrings(0, ['漢'.repeat(20000)], 1)
 const pool = new MemoryPool({ limitBytes: 200000 }); store.bindMemory(pool)
 const a = (await store.pinnedString('sst', 0))!, b = (await store.pinnedString('sst', 0))!
 expect(a.record).toBe(b.record)
 const full = 1280 + a.record.text.buffer.byteLength
 try {
  await store.close()
  // One record, charged once: its kept cache entry plus its refunded RAM backing.
  expect(pool.stats().totalBytes).toBe(full)
  const [first, second] = order === 'a-first' ? [a, b] : [b, a]
  first.release()
  expect(pool.stats().totalBytes).toBe(full)
  second.release()
  expect(pool.stats().totalBytes).toBe(0)
 } finally { a.release(); b.release(); pool.close() }
})

// r14: close never refunds and re-admits funded bytes, so a pool refusing every
// admission at close cannot drop or overdraw a leased record (no overdraft scalar).
it('close hands leased bytes over without any admission: a refusing pool still counts them', async () => {
 const backend = memoryPageBackend(), store = new SourcePageStore({ docId: 'overdraft', backend })
 await store.commitStrings(0, ['漢'.repeat(20000)], 1)
 const pool = new MemoryPool({ limitBytes: 200000 }); store.bindMemory(pool)
 const a = (await store.pinnedString('sst', 0))!, b = (await store.pinnedString('sst', 0))!
 const full = 1280 + a.record.text.buffer.byteLength
 pool.retain = (() => null) as typeof pool.retain
 const rawCharge = pool.charge.bind(pool)
 pool.charge = (() => { const c = rawCharge(), set = c.set.bind(c); return Object.assign(c, { set: (n: number) => n <= c.bytes && set(n) }) }) as typeof pool.charge
 try {
  await store.close()
  expect(pool.stats().totalBytes).toBe(full)
  a.release()
  expect(pool.stats().totalBytes).toBe(full)
  b.release()
  expect(pool.stats().totalBytes).toBe(0)
 } finally { a.release(); b.release(); pool.close() }
})
