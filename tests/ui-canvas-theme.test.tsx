import { afterEach, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { useRef } from 'react'
import { OfficeDoc } from '../src/components/OfficeDoc'
import * as paint from '../src/render/paint'
import { docxModel, fakePaintables, mockViewerLayout } from './helpers/paged-viewer'
import { initialIdentity } from '../src/worker/viewerProtocol'
import { identityFields, sameIdentity } from '../src/xlsx/view/cacheIdentity'
import { paintGridSelectionOverlay, paintSearchOverlay } from '../src/core/overlay'
import { paintFindOverlay } from '../src/components/xlsx/selectionOverlay'
import { MemoryPool, TileCache } from '../src/xlsx/view/tileCache'
import { SheetTileRenderer, type SheetFrame } from '../src/components/xlsx/sheetTiles'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { buildSheetGeometry, MergeQueryCache } from '../src/xlsx/view/viewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import { paletteKey, readCanvasColors, useCanvasColors } from '../src/components/canvasColors'
import { THEME_DEFAULTS } from '../src/styles/tokens'
import { workbookOf } from './helpers/xlsx-find'
import type { ThemeMode } from '../src/components/theme'

const identity = {docId:'d',sheetId:'s',valueGen:0,fontGen:0,viewGen:0}
const rect = {x:2,y:2,width:20,height:20}
afterEach(() => { cleanup(); vi.restoreAllMocks(); vi.unstubAllGlobals() })
function context() {
  const canvas = document.createElement('canvas'); canvas.width = 100; canvas.height = 100
  const ctx = canvas.getContext('2d')!
  const fills: string[] = [], strokes: string[] = [], texts: string[] = []
  const fill = ctx.fillRect.bind(ctx), stroke = ctx.stroke.bind(ctx), sr = ctx.strokeRect.bind(ctx), text = ctx.fillText.bind(ctx)
  vi.spyOn(ctx, 'fillRect').mockImplementation((...a) => { fills.push(String(ctx.fillStyle)); fill(...a) })
  vi.spyOn(ctx, 'stroke').mockImplementation(() => { strokes.push(String(ctx.strokeStyle)); stroke() })
  vi.spyOn(ctx, 'strokeRect').mockImplementation((...a) => { strokes.push(String(ctx.strokeStyle)); sr(...a) })
  vi.spyOn(ctx, 'fillText').mockImplementation((...a) => { texts.push(String(ctx.fillStyle)); text(...a) })
  return {ctx, fills, strokes, texts}
}
function frame(): SheetFrame {
  const sheet = workbookOf([{sheetId:'s',cells:[{row:0,col:0,value:'test'}]}]).sheets[0]
  return {sheet,geometry:buildSheetGeometry(sheet),identity,paletteKey:'a',origin:{x:0,y:0},viewport:{width:100,height:100},zoom:1,dpr:1,extent:{width:100,height:100},layoutCache:new LayoutCache(),merges:new MergeQueryCache()}
}
it('painters use passed colours and keep legacy literals by default', () => {
  const {ctx, fills, strokes} = context()
  const geometry = {fills:[rect],borders:[{x0:0,y0:0,x1:10,y1:10}]}
  paintGridSelectionOverlay(ctx, geometry)
  expect(strokes.pop()).toBe('#008000')
  paintGridSelectionOverlay(ctx, geometry, 1, {selection:'#107c41','selection-fill':'#123456'})
  expect(strokes.pop()).toBe('#107c41'); expect(fills.pop()).toBe('#123456')
  paintSearchOverlay(ctx,[rect],[rect],1,{find:'#234567','find-active':'#345678'})
  expect(fills.pop()).toBe('#234567'); expect(strokes.pop()).toBe('#345678')
  paintFindOverlay(ctx,[rect],[rect],{find:'#234567','find-active':'#345678'})
  expect(fills.slice(-2)).toEqual(['#234567','#345678'])
})
it('paper stays paper: worksheet white and gridlines are never themed', () => {
  const f = frame(), {ctx,fills,strokes} = context()
  renderSheetViewport(f.sheet,ctx,f.geometry,{box:{x:0,y:0,width:100,height:100},zoom:1,deviceScale:1},undefined,{colors:THEME_DEFAULTS.dark})
  expect(fills).toContain('#ffffff'); expect(strokes).toContain('#d0d0d0')
  const renderer = new SheetTileRenderer('D',0)
  fills.length = 0
  renderer.present(ctx,{...f,colors:THEME_DEFAULTS.dark})
  expect(fills).toContain('#ffffff'); renderer.dispose()
})
it('pending and freeze colours fall back and accept overrides', () => {
  const f = frame(), {ctx,fills,texts} = context()
  const window = {box:{x:0,y:0,width:100,height:100},zoom:1,deviceScale:1}
  const preparedPlan = {kind:'coarse' as const,groups:[{groupId:'g',bounds:window.box,slots:{start:0,end:0},cols:{c0:0,c1:0}}],diagnostics:[],work:{records:0,gridlines:0,merges:0},extent:f.extent}
  renderSheetViewport(f.sheet,ctx,undefined,window,undefined,{preparedPlan:preparedPlan as never})
  expect(fills).toContain('#eceff3'); expect(texts).toContain('#666666')
  renderSheetViewport(f.sheet,ctx,undefined,window,undefined,{preparedPlan:preparedPlan as never,colors:{pending:'#123456','pending-text':'#234567'}})
  expect(fills).toContain('#123456'); expect(texts).toContain('#234567')
  const renderer = new SheetTileRenderer('D',0)
  const frozen = {rows:1,cols:1,left:20,top:20}
  renderer.present(ctx,{...f,frozen}); expect(fills).toContain('#5f6368')
  renderer.present(ctx,{...f,frozen,paletteKey:'b',colors:{freeze:'#345678'}}); expect(fills).toContain('#345678')
  renderer.dispose()
})
it('TileCache keys differ by palette', () => {
  const cache = new TileCache(new MemoryPool({limitBytes:1024*1024}))
  const key = {identity,backend:'canvas',zoom:1,dpr:1,col:0,row:0,paletteKey:'a'}
  const surface = {width:10,height:10,close:vi.fn()}
  expect(cache.put(key,surface)).toBe(true)
  expect(cache.get({...key,paletteKey:'b'})).toBeUndefined()
  expect(cache.get(key)).toBe(surface); cache.dispose()
})
it('a palette change misses the tile cache and drops old tiles after presenting', () => {
  const renderer = new SheetTileRenderer('D',0), {ctx} = context(), f = frame()
  renderer.present(ctx,f)
  const paints = renderer.stats.basePaints
  renderer.present(ctx,f); expect(renderer.stats.tileHits).toBeGreaterThan(0)
  const hits = renderer.stats.tileHits, closed = renderer.cache.stats().closed
  const draw = ctx.drawImage.bind(ctx)
  vi.spyOn(ctx,'drawImage').mockImplementation((...args: any[]) => { expect(renderer.cache.stats().closed).toBe(closed); (draw as any)(...args) })
  renderer.present(ctx,{...f,paletteKey:'b'})
  expect(renderer.stats.tileHits).toBe(hits); expect(renderer.stats.basePaints).toBeGreaterThan(paints)
  expect(renderer.cache.stats().closed).toBeGreaterThan(closed); renderer.dispose()
})
it('the shared CacheIdentity and worker ownership checks are unchanged', () => {
  expect(Object.keys(initialIdentity('d','s')).sort()).toEqual(['docId','sheetId','valueGen','fontGen','viewGen'].sort())
  expect([...identityFields].sort()).toEqual(Object.keys(identity).sort())
  expect(sameIdentity({...identity, paletteKey:'a'} as typeof identity, {...identity, paletteKey:'b'} as typeof identity)).toBe(true)
})
it('readCanvasColors returns only resolved tokens so painters fall back', () => {
  const el = document.createElement('div'); document.body.append(el)
  expect(readCanvasColors(el)).toEqual({})
  el.style.setProperty('--ov-color-find','#12345678')
  expect(readCanvasColors(el)).toEqual({find:'#12345678'})
  expect(paletteKey({find:'x',freeze:'y'})).toBe(paletteKey({freeze:'y',find:'x'}))
  el.remove()
})
it('observes root attributes and system changes with stable palette identity and cleanup', async () => {
  let change!: () => void
  const remove = vi.fn()
  vi.stubGlobal('matchMedia',vi.fn(() => ({addEventListener:(_s:string,fn:()=>void) => {change=fn},removeEventListener:remove})))
  const snapshots: ReturnType<typeof useCanvasColors>[] = []
  function Probe({mode}:{mode:ThemeMode}) { const root=useRef<HTMLDivElement>(null); const p=useCanvasColors(root,mode); snapshots.push(p); return <div ref={root} data-testid="root" /> }
  const ui=render(<Probe mode="light"/>), root=ui.getByTestId('root')
  await act(async () => root.style.setProperty('--ov-color-find','#123456'))
  await waitFor(() => expect(snapshots.at(-1)!.colors.find).toBe('#123456'))
  const last=snapshots.at(-1)
  ui.rerender(<Probe mode="light"/>); expect(snapshots.at(-1)).toBe(last)
  await act(async () => change()); expect(snapshots.at(-1)).toBe(last)
  ui.rerender(<Probe mode="dark"/>); expect(snapshots.at(-1)).toBe(last)
  ui.unmount(); expect(remove).toHaveBeenCalled()
})
it('canvas token defaults are legible on white paper in both themes', () => {
  const luminance=(hex:string) => { const s=hex.slice(1); const rgb=[0,2,4].map(i=>parseInt(s.slice(i,i+2),16)/255).map(v=>v<=.04045?v/12.92:((v+.055)/1.055)**2.4); return rgb[0]*.2126+rgb[1]*.7152+rgb[2]*.0722 }
  const contrast=(a:string,b:string) => {const x=luminance(a),y=luminance(b);return (Math.max(x,y)+.05)/(Math.min(x,y)+.05)}
  for(const p of Object.values(THEME_DEFAULTS)) {
    for(const k of ['selection','freeze','find-active'] as const) expect(contrast(p[k],'#ffffff')).toBeGreaterThanOrEqual(3)
    expect(contrast(p['pending-text'],p.pending)).toBeGreaterThanOrEqual(4.5)
  }
})

function recordFills() {
  const proto = Object.getPrototypeOf(document.createElement('canvas').getContext('2d'))
  const original = proto.fillRect
  const fills: string[] = []
  vi.spyOn(proto, 'fillRect').mockImplementation(function(this: CanvasRenderingContext2D, ...args: unknown[]) {
    fills.push(String(this.fillStyle)); return original.apply(this, args)
  })
  return fills
}
it('paged overlay repaints when canvas colours change without remount', async () => {
  const restore = mockViewerLayout({width:800,height:600})
  vi.spyOn(paint, 'getPaintables').mockImplementation(async () => fakePaintables([{w:800,h:1000,text:'test words'}]))
  const fills = recordFills()
  try {
    const ui = render(<OfficeDoc document={docxModel(['test words'])} showSearch />)
    const input = await ui.findByRole('searchbox', {name:'Search document'})
    fireEvent.change(input, {target:{value:'test'}})
    await waitFor(() => expect(fills.some(c=>c.includes('255'))).toBe(true))
    const canvas = ui.container.querySelector('canvas')!
    const root = ui.container.querySelector<HTMLElement>('[data-ov-theme-root]')!
    fills.length = 0
    await act(async () => root.style.setProperty('--ov-color-find','rgba(0,0,255,.4)'))
    await waitFor(() => expect(fills.some(c=>c.replace(/\s/g,'').startsWith('rgba(0,0,255,'))).toBe(true))
    expect(ui.container.querySelector('canvas')).toBe(canvas)
  } finally { cleanup(); restore() }
})
it('switching theme repaints the sheet with new overlay colours on the same root node', async () => {
  const fills = recordFills()
  const doc = workbookOf([{sheetId:'s',cells:[{row:0,col:0,value:'test words'}]}])
  const style = document.createElement('style')
  style.textContent = '[data-ov-theme-root][data-ov-theme="light"]{--ov-color-find:#1122aa}[data-ov-theme-root][data-ov-theme="dark"]{--ov-color-find:#3322bb}'
  document.head.append(style)
  try {
    const ui = render(<OfficeDoc document={doc} interactiveSheets showSearch theme="light" />)
    const input = await ui.findByRole('searchbox')
    fireEvent.change(input,{target:{value:'test'}})
    await waitFor(() => expect(fills).toContain('#1122aa'))
    const root = ui.container.querySelector('[data-officeview-sheet-instance]')
    const canvas = ui.container.querySelector('[data-testid="officeview-sheet-overlay"]')
    fills.length = 0
    ui.rerender(<OfficeDoc document={doc} interactiveSheets showSearch theme="dark" />)
    await waitFor(() => expect(fills).toContain('#3322bb'))
    expect(ui.container.querySelector('[data-officeview-sheet-instance]')).toBe(root)
    expect(ui.container.querySelector('[data-testid="officeview-sheet-overlay"]')).toBe(canvas)
  } finally { cleanup(); style.remove() }
})
it('resamples class, root theme, system scheme and explicit mode changes', async () => {
  let change!: () => void
  vi.stubGlobal('matchMedia',vi.fn(() => ({addEventListener:(_s:string,fn:()=>void)=>{change=fn},removeEventListener:vi.fn()})))
  const style=document.createElement('style')
  style.textContent='.palette-a{--ov-color-find:#111111}.palette-b{--ov-color-find:#222222}[data-ov-theme="dark"]{--ov-color-find:#333333}'
  document.head.append(style)
  let latest: ReturnType<typeof useCanvasColors>
  function Probe({mode}:{mode:ThemeMode}) {const root=useRef<HTMLDivElement>(null);latest=useCanvasColors(root,mode);return <div ref={root} className="palette-a" data-testid="root"/>}
  try {
    const ui=render(<Probe mode="light"/>),root=ui.getByTestId('root')
    await waitFor(()=>expect(latest.colors.find).toBe('#111111'))
    await act(async()=>{root.className='palette-b'})
    await waitFor(()=>expect(latest.colors.find).toBe('#222222'))
    await act(async()=>{root.dataset.ovTheme='dark'})
    await waitFor(()=>expect(latest.colors.find).toBe('#333333'))
    style.textContent='[data-ov-theme="dark"]{--ov-color-find:#444444}'
    await act(async()=>change())
    await waitFor(()=>expect(latest.colors.find).toBe('#444444'))
    style.textContent='[data-ov-theme="dark"]{--ov-color-find:#555555}'
    ui.rerender(<Probe mode="dark"/>)
    await waitFor(()=>expect(latest.colors.find).toBe('#555555'))
  } finally {cleanup();style.remove()}
})
