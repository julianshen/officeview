import { afterEach, expect, it, vi } from 'vitest'
import { act, cleanup, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import * as loader from '../src/xlsx/cfLoader'
import { TileCache } from '../src/xlsx/view/tileCache'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'
import type { XlsxSheet } from '../src/xlsx/types'
afterEach(() => { cleanup(); vi.restoreAllMocks() })
it('mounted viewer cannot retain blank host tiles while the CF asset load is delayed', FIRST_PAINT_TEST, async () => {
  const sheet: XlsxSheet = { name: 'CF', sheetId: '1', rows: [], cols: [], merges: [], mergeRanges: [], conditionalFormatting: { diagnostics: [], rules: [{ id: 0, priority: 1, documentOrder: 0, stopIfTrue: false, ranges: [{ firstRow: 1, lastRow: 1, firstCol: 1, lastCol: 1 }], writes: new Set(['fill']), dxf: { 'fill.bg': 'FF00FF00' }, definition: { kind: 'predicate', type: 'containsBlanks', dialect: 'base', formulas: [] } }] } }
  const module = await import('../src/xlsx/cf/render')
  // Preparation exists first, but its optional rendering asset is still delayed.
  const prepared = await module.prepareCfPaint(sheet)
  let release!: () => void
  const gate = new Promise<void>(r => { release = r })
  const load = vi.spyOn(loader, 'ensureConditionalFormatting').mockImplementation(async () => { await gate; return { ...module, prepareCfPaint: async () => prepared } })
  const put = vi.spyOn(TileCache.prototype, 'put')
  const ui = render(<InteractiveSheetDoc doc={{ sheets: [sheet] }} />)
  try {
    await waitFor(() => expect(load).toHaveBeenCalled(), FIRST_PAINT_WAIT)
    expect(put).not.toHaveBeenCalled()
    await act(async () => { release() })
    await waitFor(() => expect(put).toHaveBeenCalled(), FIRST_PAINT_WAIT)
    expect(load).toHaveBeenCalledTimes(1)
  } finally { release(); ui.unmount(); prepared?.dispose() }
})
it('mounted CF assets receive projected blank coverage beyond the stored rows', FIRST_PAINT_TEST, async () => {
  const { prepareSheetAssets } = await import('../src/components/xlsx/sheetAssets')
  const sheet: XlsxSheet = { name: 'CF', sheetId: '1', rows: [{ index: 0, cells: [{ row: 0, col: 0, ref: 'A1', value: 10, styleIndex: 0 }] }], cols: [], merges: [], mergeRanges: [], conditionalFormatting: { diagnostics: [], rules: [{ id: 0, priority: 1, documentOrder: 0, stopIfTrue: false, ranges: [{ firstRow: 0, lastRow: 1048575, firstCol: 1, lastCol: 1 }], writes: new Set(['fill']), dxf: { 'fill.bg': 'FF00FF00' }, definition: { kind: 'predicate', type: 'containsBlanks', dialect: 'base', formulas: [] } }] } }
  let fill: string | undefined, parentSeen = false
  const assetLoader: typeof prepareSheetAssets = async (doc, options) => {
    parentSeen = Boolean(options?.chartParent)
    const lease = await prepareSheetAssets(doc, options)
    fill = lease.assets.conditionalFormatting?.layer.styleAt(5, 1, undefined)?.fillColor
    return lease
  }
  const ui = render(<InteractiveSheetDoc doc={{ sheets: [sheet] }} assetLoader={assetLoader} />)
  try { await waitFor(() => expect(fill).toBe('FF00FF00'), FIRST_PAINT_WAIT); expect(parentSeen).toBe(true) } finally { ui.unmount() }
})
it('one-pixel scrolling does not reload fonts/images or reprepare finite CF', FIRST_PAINT_TEST, async () => {
  const { fireEvent } = await import('@testing-library/react')
  const { prepareSheetAssets } = await import('../src/components/xlsx/sheetAssets')
  const sheet: XlsxSheet = { name:'CF',sheetId:'1',rows:[{index:999,cells:[{row:999,col:20,ref:'U1000',value:1,styleIndex:0}]}],cols:[],merges:[],mergeRanges:[],conditionalFormatting:{diagnostics:[],rules:[{id:0,priority:1,documentOrder:0,stopIfTrue:false,ranges:[{firstRow:0,lastRow:19,firstCol:0,lastCol:7}],writes:new Set(['fill']),dxf:{'fill.bg':'FF00FF00'},definition:{kind:'predicate',type:'containsBlanks',dialect:'base',formulas:[]}}]}}
  const load=vi.fn(prepareSheetAssets),ui=render(<InteractiveSheetDoc doc={{sheets:[sheet]}} assetLoader={load}/>);
  try { await waitFor(()=>expect(load).toHaveBeenCalledTimes(1),FIRST_PAINT_WAIT); await act(async()=>{await new Promise(r=>setTimeout(r,100))}); const scroller=ui.getByTestId('officeview-sheet-scroller'); await act(async()=>{scroller.scrollTop=1;fireEvent.scroll(scroller);await new Promise(r=>setTimeout(r,100))}); expect(load).toHaveBeenCalledTimes(1) } finally {ui.unmount()}
})
it('open-axis CF covers the tile ring and blocks cache publication during a scoped refresh', FIRST_PAINT_TEST, async () => {
  const { fireEvent } = await import('@testing-library/react')
  const { prepareSheetAssets } = await import('../src/components/xlsx/sheetAssets')
  const module = await import('../src/xlsx/cf/render')
  const sheet: XlsxSheet = { name:'CF',sheetId:'1',rows:[{index:999,cells:[{row:999,col:20,ref:'U1000',value:1,styleIndex:0}]}],cols:[],merges:[],mergeRanges:[],conditionalFormatting:{diagnostics:[],rules:[{id:0,priority:1,documentOrder:0,stopIfTrue:false,ranges:[{firstRow:0,lastRow:1048575,firstCol:1,lastCol:1}],writes:new Set(['fill']),dxf:{'fill.bg':'FF00FF00'},definition:{kind:'predicate',type:'containsBlanks',dialect:'base',formulas:[]}}]}}
  const real=module.prepareCfPaint;let release!:()=>void;const gate=new Promise<void>(r=>release=r);let latest: Awaited<ReturnType<typeof real>>
  const prepare=vi.spyOn(module,'prepareCfPaint').mockImplementation(async(...args)=>{if(prepare.mock.calls.length>1)await gate;return latest=await real(...args)})
  const put=vi.spyOn(TileCache.prototype,'put'),load=vi.fn(prepareSheetAssets),ui=render(<InteractiveSheetDoc doc={{sheets:[sheet]}} assetLoader={load}/>);
  try {
    await waitFor(()=>expect(put).toHaveBeenCalled(),FIRST_PAINT_WAIT)
    // Outside the pane, inside its first 512-device-pixel tile/ring.
    expect(latest?.layer.styleAt(40,1,undefined)?.fillColor).toBe('FF00FF00')
    const scroller=ui.getByTestId('officeview-sheet-scroller')
    await act(async()=>{scroller.scrollTop=1;fireEvent.scroll(scroller);await new Promise(r=>setTimeout(r,100))});expect(prepare).toHaveBeenCalledTimes(1)
    const before=put.mock.calls.length
    await act(async()=>{scroller.scrollTop=600;fireEvent.scroll(scroller);await new Promise(r=>setTimeout(r,100))})
    await waitFor(()=>expect(prepare).toHaveBeenCalledTimes(2),FIRST_PAINT_WAIT);expect(load).toHaveBeenCalledTimes(1);expect(put.mock.calls.length).toBe(before)
    await act(async()=>release());await waitFor(()=>expect(put.mock.calls.length).toBeGreaterThan(before),FIRST_PAINT_WAIT)
    expect(latest?.layer.styleAt(60,1,undefined)?.fillColor).toBe('FF00FF00')
  } finally {release();ui.unmount()}
})
