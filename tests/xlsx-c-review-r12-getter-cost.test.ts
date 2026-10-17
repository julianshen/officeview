// @vitest-environment node
import {expect,it} from 'vitest'
import {writeFile} from 'node:fs/promises'
import {chartWorkbookBytes} from './helpers/xlsx-chart-fixture'
import {SourcePageStore,memoryPageBackend} from '../src/xlsx/storage/sourcePages'
import {openImportSession} from '../src/xlsx/storage/pagedImport'
import {createPagedWorkbookSource} from '../src/xlsx/view/pagedSource'
import {MemoryPool} from '../src/xlsx/view/tileCache'
it('charge refresh does not repeatedly traverse an unchanged page metadata map for every chart label',async()=>{
 const receipts:any[]=[]
 for(const count of [32,64,128]){
  const xml=`<chartSpace><chart><plotArea><barChart><ser><cat><strRef><f>Styled!A1:A${count}</f></strRef></cat><val><numRef><f>Styled!B1:B${count}</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>`
  const bytes=await chartWorkbookBytes(xml,{cells:Array.from({length:count},(_,i)=>[{ref:`A${i+1}`,v:i+1},{ref:`B${i+1}`,v:i+1}]).flat(),patch:async zip=>{
   const path='xl/worksheets/sheet1.xml';let text=await zip.file(path)!.async('string')
   for(let i=1;i<=count;i++)text=text.replace(new RegExp(`(<c r="A${i}"[^>]*>)`),'$1<f>1+N("漢")</f>')
   zip.file(path,text)
  }})
  const backend=memoryPageBackend(),store=new SourcePageStore({docId:'getter-cost',backend}),session=await openImportSession({kind:'blob',blob:new Blob([new Uint8Array(bytes)])},{docId:'getter-cost',profile:'D',store,pageRows:count})
  expect((await session.done).kind).toBe('complete');const source=createPagedWorkbookSource(session),adapter=source.store('1')!,pool=new MemoryPool({limitBytes:4000000});source.bindMemory!(pool)
  const p=await store.page('1',store.pages('1')[0].pageId);expect(p.meta.size).toBe(count)
  let visits=0,scans=0;const iterator=p.meta[Symbol.iterator].bind(p.meta)
  Object.defineProperty(p.meta,Symbol.iterator,{configurable:true,value:function*(){scans++;for(const entry of iterator()){visits++;yield entry}}})
  const rawPin=store.pinnedPage.bind(store);store.pinnedPage=async(...args)=>{const pending=rawPin(...args);pool.trim(0);const pin=await pending;expect(pin.retained).toBe(false);return pin}
  const charge=pool.charge();let admissions=0
  try{
   const result=await adapter.readDrawings!({stamp:adapter.stamp,signal:new AbortController().signal},'D',{set:n=>{admissions++;return charge.set(n)}})
   expect(result.completeness).toBe('complete');const chart=result.drawings[0].content as any;expect(chart.chart.categories).toEqual(Array.from({length:count},(_,i)=>String(i+1)))
   receipts.push({count,visits,scans,admissions,labels:chart.chart.categories.length})
  }finally{delete (p.meta as any)[Symbol.iterator];charge.release();await source.release();pool.close()}
 }
  void writeFile
 // One read's unchanged metadata can be priced once (allow a small constant number of traversals).
 expect(receipts.every(r=>r.visits<=4*r.count)).toBe(true)
},30000)
