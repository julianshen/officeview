import { createHash } from 'node:crypto'
import { messageFor } from '../../scripts/stream-s0/mark/message'
import { extendedCodeword } from '../../scripts/stream-s0/mark/bch'
import { expect, it } from 'vitest'
import { basis } from '../../scripts/stream-s0/mark/basis'
import { embed } from '../../scripts/stream-s0/mark/embed'
import { crop, scale } from '../../scripts/stream-s0/mark/attacks'
import { register, lum } from '../../scripts/stream-s0/mark/register'
import { listCorpus, loadUnits, renderUnit } from '../../scripts/stream-s0/corpus'
const params={nChips:512,cellPx:4,tilePx:64,amplitude:2,pedestal:2}
const code=Int8Array.from({length:512},(_,i)=>i%3?1:-1)
it('recovers the DOCX 25-percent crop at original offset 102,132',async()=>{
 const file=listCorpus().find(f=>f.name==='poi-55733.docx')!
 expect(file).toBeDefined()
 const units=await loadUnits(file)
 try {
 const ref=await renderUnit(units[0]!,1);if('skipped'in ref)throw Error(ref.skipped)
 expect([ref.width,ref.height]).toEqual([816,1056])
 const key=(s:string)=>new Uint8Array(createHash('sha256').update(s).digest())
 const chips=basis(key('survival:poi-55733.docx#0'),ref.width,ref.height,params)
 const payload=extendedCodeword(messageFor(key('x'),1,1,0))
 const cap=crop(embed(ref,chips,payload,params),.25)
 const content=register(cap,ref)!
 let trueMse=0;for(let y=0;y<cap.height;y++)for(let x=0;x<cap.width;x++)trueMse+=(lum(cap,x,y)-lum(ref,x+102,y+132))**2
 trueMse/=cap.width*cap.height
 expect(content.error).toBeLessThanOrEqual(trueMse+1e-9)
 // This sparse page has tied blank content regions; page support resolves them.
 const pose=register(cap,ref,chips)
 expect([pose?.x,pose?.y]).toEqual([102,132])
 }finally{units.dispose()}
},60000)
it('uses page support to synchronize a flat slide without candidate codes',()=>{
 const ref={width:1280,height:720,data:new Uint8ClampedArray(1280*720*4).fill(255)}
 const chips=basis(new Uint8Array(32).fill(7),ref.width,ref.height,params)
 const pose=register(crop(embed(ref,chips,code,params),.25),ref,chips)
 expect([pose?.x,pose?.y]).toEqual([160,90])
},60000)

it.each([.5,2])('flat crop-then-scale %s retains original pose',factor=>{
 const ref={width:1280,height:720,data:new Uint8ClampedArray(1280*720*4).fill(255)}
 const chips=basis(new Uint8Array(32).fill(7),ref.width,ref.height,params)
 const pose=register(scale(crop(embed(ref,chips,code,params),.25),factor),ref,chips)
 expect([pose?.x,pose?.y]).toEqual([160,90])
},60000)
it.each(['top-left','bottom-right'] as const)('content minimum also recovers a non-centred %s crop',anchor=>{
 const width=96,height=80,data=new Uint8ClampedArray(width*height*4)
 for(let y=0;y<height;y++)for(let x=0;x<width;x++){const i=(y*width+x)*4,v=(x*37+y*61+x*y*7)%200;data[i]=v;data[i+1]=v;data[i+2]=v;data[i+3]=255}
 const ref={width,height,data},cap=crop(ref,.25,anchor),pose=register(cap,ref)
 expect([pose?.x,pose?.y]).toEqual(anchor==='top-left'?[0,0]:[24,20])
 expect(pose?.error).toBe(0)
})
it('the low-contrast narration rectangle registers the payload-23 enlarged crop',async()=>{
 const file=listCorpus().find(f=>f.name==='lo-narration.pptx')!,units=await loadUnits(file)
 try {
  const ref=await renderUnit(units[0]!,1);if('skipped'in ref)throw Error(ref.skipped)
  const key=(s:string)=>new Uint8Array(createHash('sha256').update(s).digest())
  const chips=basis(key('survival:lo-narration.pptx#0'),ref.width,ref.height,params)
  const payload=extendedCodeword(messageFor(key('stage0-t8-public-reproducibility-message-key'),1,23,23))
  const pose=register(scale(crop(embed(ref,chips,payload,params),.25),2),ref,chips)
  expect([pose?.x,pose?.y]).toEqual([160,90])
  expect(pose?.error).toBeLessThan(.1)
 }finally{units.dispose()}
},60000)
