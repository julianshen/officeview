import { afterEach, expect, it, vi } from 'vitest'
import { cleanup, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import * as display from '../src/xlsx/render'
import type { XlsxDocument } from '../src/xlsx/types'
afterEach(()=>{cleanup();vi.restoreAllMocks()})
it('mounted semantic values use an admitted formatter rather than parsing a huge format outside admission',async()=>{
 const code='[>0]0;[<=0]"'+'q'.repeat(65536)+'"0'
 const doc:XlsxDocument={sheets:[{name:'S',rows:[{index:0,cells:[{ref:'A1',row:0,col:0,value:3,styleIndex:0,style:{numFmtId:164,formatCode:code}}]}],cols:[],merges:[],mergeRanges:[]}]}
 const native=display.cellDisplayText,events:{observed:boolean,stack:string}[]=[]
 vi.spyOn(display,'cellDisplayText').mockImplementation((...args)=>{if(args[1].style?.formatCode===code)events.push({observed:!!args[2],stack:new Error().stack??''});return native(...args)})
 const ui=render(<InteractiveSheetDoc doc={doc}/> )
 await waitFor(()=>expect(ui.getAllByRole('gridcell').length).toBeGreaterThan(0))
 const semantic=events.filter(e=>e.stack.includes('cellPresentation'))
 console.log('mounted semantic formatting',JSON.stringify({calls:semantic.length,observed:semantic.filter(e=>e.observed).length,codeUnits:code.length}))
 expect(semantic.length).toBeGreaterThan(0)
 expect(semantic.every(e=>e.observed)).toBe(true)
})
