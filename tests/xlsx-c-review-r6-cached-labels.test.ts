// @vitest-environment node
import {expect,it} from 'vitest'
import {writeFile} from 'node:fs/promises'
import {CACHED_COLUMN,chartWorkbookBytes} from './helpers/xlsx-chart-fixture'
import {parseXlsx} from '../src/xlsx/parse'
import {OfficePackage} from '../src/core/zip'
import {openImportSession} from '../src/xlsx/storage/pagedImport'
import {createPagedWorkbookSource} from '../src/xlsx/view/pagedSource'
it('cached category and series labels also honor the amended 255-unit cap',async()=>{
 const long='Q'.repeat(300),xml=CACHED_COLUMN.replace('<v>A</v>',`<v>${long}</v>`).replace('<v>2024</v>',`<v>${long}</v>`)
 const bytes=await chartWorkbookBytes(xml),doc=await parseXlsx(await OfficePackage.load(bytes))
 const session=await openImportSession({kind:'blob',blob:new Blob([new Uint8Array(bytes)])},{docId:'cache-cap',profile:'D'});expect((await session.done).kind).toBe('complete')
 const source=createPagedWorkbookSource(session),adapter=source.store('1') as any
 try{
  const resident=(doc.sheets[0].drawings![0].content as any).chart,paged=(await adapter.readDrawings({stamp:adapter.stamp,signal:new AbortController().signal},'D')).drawings[0].content.chart
  void writeFile
  for(const chart of [resident,paged]){expect(chart.categories[0]).toBe(long.slice(0,255));expect(chart.groups[0].series[0].name).toBe(long.slice(0,255));expect(chart.truncated).toBe(true)}
 }finally{await source.release()}
})
