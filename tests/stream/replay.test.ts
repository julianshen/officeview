// @vitest-environment node
import { createCanvas } from 'canvas'
import { describe, expect, it, vi } from 'vitest'
import { OpWriter, Opcode as O, type Asset } from '../../src/stream/ops'
import { replayUnit, ReplayError } from '../../src/stream/replay'
import type { CanvasLike } from '../../src/core/canvas-backend'

const surface = () => createCanvas(96, 80)
const context = (c: ReturnType<typeof surface>) => c.getContext('2d') as unknown as CanvasRenderingContext2D
const pixels = (c: ReturnType<typeof surface>) => c.getContext('2d').getImageData(0, 0, c.width, c.height).data
const backend = { createCanvas: (w: number, h: number) => createCanvas(w, h) as unknown as CanvasLike }
const map = (...assets: Asset[]) => new Map(assets.map(a => [a.id, a]))
const bitmap: Asset = { id: 1, hash: 'bitmap', type: 'bitmap', w: 2, h: 2, rgba: Uint8Array.of(255,0,0,255, 0,255,0,255, 0,0,255,255, 255,255,0,255) }
function source() { const c = createCanvas(2, 2), ctx = c.getContext('2d'), d = ctx.createImageData(2, 2); d.data.set(bitmap.type === 'bitmap' ? bitmap.rgba : []); ctx.putImageData(d, 0, 0); return c }
async function compare(w: OpWriter, paint: (ctx: CanvasRenderingContext2D) => void, assets = map()) {
  const direct = surface(), replay = surface(); paint(context(direct))
  await replayUnit(context(replay), { ops: w.bytes(), assets }, backend)
  expect(Buffer.compare(Buffer.from(pixels(replay)), Buffer.from(pixels(direct)))).toBe(0)
}

describe('stream replayer', () => {
  it('replays state, nested clips, transform, shadows, dash, paths and rect modes exactly', async () => {
    const w = new OpWriter().op(O.STATE_PUSH).op(O.TRANSFORM, { mode: 0, matrix: [1,0,0,1,3,2] })
      .op(O.STYLE_COLOR, { target: 0, color: '#123456' }).op(O.STYLE_COLOR, { target: 1, color: '#ff8844' })
      .op(O.STYLE_COLOR, { target: 2, color: 'rgba(0,0,0,.5)' })
      .op(O.STATE_SETUP, { fields: { globalAlpha: .8, lineWidth: 2, lineCap: 'round', lineJoin: 'bevel', miterLimit: 4, lineDash: [3,2], lineDashOffset: 1, shadowBlur: 2, shadowOffsetX: 1, shadowOffsetY: 2 } })
      .op(O.PATH, { sub: 0, args: [] }).op(O.PATH, { sub: 8, args: [1,1,70,60] }).op(O.PATH_END, { mode: 3, rule: 0 })
      .op(O.STATE_PUSH).op(O.TRANSFORM, { mode: 1, matrix: [1,.1,0,1,1,0] })
      .op(O.RECT, { mode: 0, rect: [2,2,50,40] }).op(O.RECT, { mode: 1, rect: [4,4,40,30] }).op(O.RECT, { mode: 2, rect: [9,9,5,5] })
      .op(O.STATE_POP).op(O.STATE_POP)
    await compare(w, c => {
      c.save(); c.setTransform(1,0,0,1,3,2); c.fillStyle='#123456'; c.strokeStyle='#ff8844'; c.shadowColor='rgba(0,0,0,.5)'
      c.globalAlpha=.8; c.lineWidth=2; c.lineCap='round'; c.lineJoin='bevel'; c.miterLimit=4; c.setLineDash([3,2]); c.lineDashOffset=1; c.shadowBlur=2; c.shadowOffsetX=1; c.shadowOffsetY=2
      c.beginPath(); c.rect(1,1,70,60); c.clip(); c.save(); c.transform(1,.1,0,1,1,0); c.fillRect(2,2,50,40); c.strokeRect(4,4,40,30); c.clearRect(9,9,5,5); c.restore(); c.restore()
    })
  })
  it('replays every path verb and fill/stroke rules', async () => {
    const w = new OpWriter().op(O.PATH,{sub:0,args:[]}).op(O.PATH,{sub:1,args:[5,5]}).op(O.PATH,{sub:2,args:[40,5]})
      .op(O.PATH,{sub:3,args:[50,5,50,15]}).op(O.PATH,{sub:4,args:[50,20,40,30,30,30]})
      .op(O.PATH,{sub:5,args:[30,30,10,0,Math.PI],ccw:false}).op(O.PATH,{sub:6,args:[10,30,10,10,5]})
      .op(O.PATH,{sub:7,args:[60,40,9,6,.2,0,Math.PI*2],ccw:true})
      .op(O.PATH,{sub:9,args:[10,40,25,20],radii:[2,4]}).op(O.PATH,{sub:10,args:[]}).op(O.PATH_END,{mode:2,rule:1})
    await compare(w,c=>{c.beginPath();c.moveTo(5,5);c.lineTo(40,5);c.quadraticCurveTo(50,5,50,15);c.bezierCurveTo(50,20,40,30,30,30);c.arc(30,30,10,0,Math.PI,false);c.arcTo(10,30,10,10,5);c.ellipse(60,40,9,6,.2,0,Math.PI*2,true);c.roundRect(10,40,25,20,[2,4]);c.closePath();c.fill('evenodd');c.stroke()})
  })
  it('replays full shaped text with maxWidth and strokeText', async () => {
    const w = new OpWriter().op(O.STATE_SETUP,{fields:{font:'16px sans-serif',textAlign:'left',textBaseline:'alphabetic'}})
      .op(O.GLYPH_RUN,{mode:0,text:'Hello ffi',x:3,y:25,maxWidth:45,advances:[1,2,3]})
      .op(O.GLYPH_RUN,{mode:1,text:'World',x:4,y:50,advances:[]})
    await compare(w,c=>{c.font='16px sans-serif';c.textAlign='left';c.textBaseline='alphabetic';c.fillText('Hello ffi',3,25,45);c.strokeText('World',4,50)})
  })
  it.each(['linear','radial'] as const)('replays %s gradient stops',async type=>{
    const a: Asset = type==='linear' ? {id:2,hash:'linear',type,x0:0,y0:0,x1:80,y1:30,stops:[[0,'red'],[.4,'blue'],[1,'green']]} : {id:2,hash:'radial',type,x0:30,y0:30,r0:0,x1:40,y1:30,r1:40,stops:[[0,'red'],[1,'blue']]}
    const w = new OpWriter().op(O.STYLE_OBJ,{target:0,assetId:2}).op(O.RECT,{mode:0,rect:[0,0,80,70]})
    await compare(w,c=>{const g=type==='linear'?c.createLinearGradient(0,0,80,30):c.createRadialGradient(30,30,0,40,30,40);for(const [n,color] of a.stops)g.addColorStop(n,color);c.fillStyle=g;c.fillRect(0,0,80,70)},map(a))
  })
  it('replays pattern transforms and all drawImage arities; creates each bitmap once',async()=>{
    const pattern: Asset={id:2,hash:'pattern',type:'pattern',bitmap:1,repetition:'repeat',transform:[3,0,0,2,1,1]}
    const w=new OpWriter().op(O.STYLE_OBJ,{target:0,assetId:2}).op(O.RECT,{mode:0,rect:[0,0,30,30]})
      .op(O.IMAGE_DRAW,{assetId:1,args:[40,0]}).op(O.IMAGE_DRAW,{assetId:1,args:[40,10,12,12]}).op(O.IMAGE_DRAW,{assetId:1,args:[0,0,1,2,60,10,12,20]})
    const direct=surface(),out=surface(),c=context(direct),src=source(),p=c.createPattern(src as unknown as CanvasImageSource,'repeat')!
    const m=c.getTransform();Object.assign(m,{a:3,b:0,c:0,d:2,e:1,f:1});p.setTransform(m);c.fillStyle=p;c.fillRect(0,0,30,30);c.drawImage(src as unknown as CanvasImageSource,40,0);c.drawImage(src as unknown as CanvasImageSource,40,10,12,12);c.drawImage(src as unknown as CanvasImageSource,0,0,1,2,60,10,12,20)
    const create=vi.fn(async(w:number,h:number)=>backend.createCanvas(w,h))
    await replayUnit(context(out),{ops:w.bytes(),assets:map(bitmap,pattern)},{createCanvas:create})
    expect(create).toHaveBeenCalledTimes(1);expect(Buffer.compare(Buffer.from(pixels(out)),Buffer.from(pixels(direct)))).toBe(0)
  })
  it('replays inline TILE under the current transform',async()=>{
    const rgba=bitmap.type==='bitmap'?bitmap.rgba:new Uint8Array()
    const w=new OpWriter().op(O.TRANSFORM,{mode:0,matrix:[2,0,0,2,4,3]}).op(O.TILE,{w:2,h:2,dx:1,dy:2,rgba})
    await compare(w,c=>{c.setTransform(2,0,0,2,4,3);c.drawImage(source() as unknown as CanvasImageSource,1,2)})
  })
  it.each(['asset','opcode','pattern','asset-kind','bitmap-size'])('rejects late %s failure before any pixels change',async bad=>{
    const w=new OpWriter().op(O.RECT,{mode:0,rect:[0,0,20,20]});let assets=map();let ops:Uint8Array
    if(bad==='opcode')ops=Uint8Array.from([...w.bytes(),255])
    else if(bad==='pattern'){assets=map({id:2,hash:'p',type:'pattern',bitmap:999,repetition:'repeat'});ops=w.op(O.STYLE_OBJ,{target:0,assetId:2}).bytes()}
    else if(bad==='asset-kind'){assets=map({id:2,hash:'g',type:'linear',x0:0,y0:0,x1:1,y1:1,stops:[]});ops=w.op(O.IMAGE_DRAW,{assetId:2,args:[0,0]}).bytes()}
    else if(bad==='bitmap-size'){assets=map({...bitmap,type:'bitmap',w:5,h:5,rgba:new Uint8Array(4)});ops=w.op(O.IMAGE_DRAW,{assetId:1,args:[0,0]}).bytes()}
    else ops=w.op(O.IMAGE_DRAW,{assetId:999,args:[0,0]}).bytes()
    const out=surface();await expect(replayUnit(context(out),{ops,assets},backend)).rejects.toBeInstanceOf(ReplayError);expect(pixels(out).some(n=>n!==0)).toBe(false)
  })
  it('a rejected asynchronous surface allocation leaves output untouched',async()=>{
    const w=new OpWriter().op(O.RECT,{mode:0,rect:[0,0,20,20]}).op(O.IMAGE_DRAW,{assetId:1,args:[0,0]});const out=surface()
    await expect(replayUnit(context(out),{ops:w.bytes(),assets:map(bitmap)},{createCanvas:async()=>{throw new Error('refused')}})).rejects.toBeInstanceOf(ReplayError)
    expect(pixels(out).some(n=>n!==0)).toBe(false)
  })
})

it('rejects a cyclic pattern-to-pattern reference without traversing the cycle', async () => {
  const assets = map({id:1,hash:'p',type:'pattern',bitmap:1,repetition:'repeat'})
  const lookup = vi.spyOn(assets,'get'), out = surface()
  await expect(replayUnit(context(out),{ops:new OpWriter().op(O.STYLE_OBJ,{target:0,assetId:1}).bytes(),assets},backend)).rejects.toBeInstanceOf(ReplayError)
  expect(lookup.mock.calls.length).toBeLessThanOrEqual(2)
  expect(pixels(out).some(n=>n!==0)).toBe(false)
})
it('owns validated image bytes across asynchronous backend preparation', async () => {
  const rgba = bitmap.type==='bitmap' ? bitmap.rgba.slice() : new Uint8Array()
  const assets = map({...bitmap,type:'bitmap',w:2,h:2,rgba})
  const w = new OpWriter().op(O.IMAGE_DRAW,{assetId:1,args:[0,0]})
  let resolve!: (canvas:CanvasLike)=>void
  const out = surface(), run = replayUnit(context(out),{ops:w.bytes(),assets},{createCanvas:()=>new Promise(r=>{resolve=r})})
  rgba.fill(0); assets.clear(); resolve(backend.createCanvas(2,2)); await run
  const expected=surface();context(expected).drawImage(source() as unknown as CanvasImageSource,0,0)
  expect(Buffer.compare(Buffer.from(pixels(out)),Buffer.from(pixels(expected)))).toBe(0)
})

it.each([5,6,7,9] as const)('rejects a late negative-radius path %i before drawing',async sub=>{
  const args=sub===5?[20,20,-1,0,1]:sub===6?[1,2,3,4,-1]:sub===7?[20,20,-1,2,0,0,1]:[1,2,3,4]
  const w=new OpWriter().op(O.RECT,{mode:0,rect:[0,0,20,20]}).op(O.PATH,{sub,args,...(sub===5||sub===7?{ccw:false}:{}),...(sub===9?{radii:[-1]}:{})})
  const out=surface();await expect(replayUnit(context(out),{ops:w.bytes(),assets:map()},backend)).rejects.toBeInstanceOf(ReplayError)
  expect(pixels(out).some(n=>n!==0)).toBe(false)
})

describe('native rejection preflight', () => {
  const cases = [
    ...[NaN, Infinity, -Infinity, -1].map(value => ({ name:`dash ${value}`, code:O.STATE_SETUP as const, data:{fields:{lineDash:[2,value]}} })),
    ...[NaN, Infinity, -Infinity].map(value => ({ name:`matrix ${value}`, code:O.TRANSFORM as const, data:{mode:0 as const,matrix:[1,0,0,1,value,0] as [number,number,number,number,number,number]} })),
    ...[NaN, Infinity, -Infinity].map(value => ({ name:`text x ${value}`, code:O.GLYPH_RUN as const, data:{mode:0 as const,text:'test',x:value,y:10,advances:[]} })),
    ...[NaN, Infinity, -Infinity, -1].map(value => ({ name:`text maxWidth ${value}`, code:O.GLYPH_RUN as const, data:{mode:0 as const,text:'test',x:1,y:10,maxWidth:value,advances:[]} })),
    ...[NaN, Infinity, -Infinity].map(value => ({ name:`arc ${value}`, code:O.PATH as const, data:{sub:5 as const,args:[10,10,value,0,2],ccw:false} })),
    ...[NaN, Infinity, -Infinity].map(value => ({ name:`roundRect ${value}`, code:O.PATH as const, data:{sub:9 as const,args:[1,1,10,10],radii:[value]} })),
  ]
  it.each(cases)('never partially paints when the backend rejects $name', async ({code,data}) => {
    const w=new OpWriter().op(O.RECT,{mode:0,rect:[0,0,20,20]});w.op(code,data as never)
    const out=surface()
    try { await replayUnit(context(out),{ops:w.bytes(),assets:map()},backend) }
    catch(error) { expect(error).toBeInstanceOf(ReplayError); expect(pixels(out).some(n=>n!==0)).toBe(false) }
  })
  it('preserves native ignored style-string semantics',async()=>{
    const w=new OpWriter().op(O.STYLE_COLOR,{target:0,color:'red'}).op(O.STYLE_COLOR,{target:0,color:'not-a-color'})
      .op(O.STATE_SETUP,{fields:{font:'16px sans-serif'}}).op(O.STATE_SETUP,{fields:{font:'not-a-font',globalCompositeOperation:'not-a-composite'}})
      .op(O.RECT,{mode:0,rect:[0,0,20,20]}).op(O.GLYPH_RUN,{mode:0,text:'A',x:2,y:40,advances:[]})
    await compare(w,c=>{c.fillStyle='red';c.fillStyle='not-a-color';c.font='16px sans-serif';c.font='not-a-font';c.globalCompositeOperation='not-a-composite' as GlobalCompositeOperation;c.fillRect(0,0,20,20);c.fillText('A',2,40)})
  })
})
