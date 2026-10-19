import { expect,it } from 'vitest'
import { createCanvas } from 'canvas'
import { SheetTileRenderer } from '../src/components/xlsx/sheetTiles'
import { buildSheetGeometry,MergeQueryCache } from '../src/xlsx/view/viewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import type { XlsxSheet } from '../src/xlsx/types'
it('a per-frame reused frozen tile remains live through all clips under eviction',()=>{
 const sheet:XlsxSheet={name:'S',rows:[],cols:[],merges:[],mergeRanges:[],drawings:[{xEmu:0,yEmu:0,widthEmu:2048*9525,heightEmu:2048*9525,geometry:'rect',fill:'#00ff00',source:{partPath:'d',treePath:'s',element:'sp',representation:'native'}}]}
 const geometry=buildSheetGeometry(sheet);if(geometry.kind!=='ready')throw Error('geometry')
 let closed=0
 const renderer=new SheetTileRenderer('D',124*1024*1024,()=>{
  const tile=createCanvas(512,512);return {canvas:tile as unknown as HTMLCanvasElement,width:512,height:512,close:()=>{closed++;tile.width=0;tile.height=0}}
 })
 const canvas=createCanvas(640,400),ctx=canvas.getContext('2d'),dead:number[][]=[],draw=ctx.drawImage.bind(ctx)
 ctx.drawImage=((...args:any[])=>{if(args[0].width===0)dead.push([args[1],args[2]]);else (draw as any)(...args)}) as any
 renderer.present(ctx as unknown as CanvasRenderingContext2D,{sheet,geometry,identity:{docId:'review',sheetId:'s',valueGen:0,fontGen:0,viewGen:0},origin:{x:480,y:480},viewport:{width:640,height:400},zoom:1,dpr:1,extent:{width:2048,height:2048},frozen:{rows:1,cols:1,left:64,top:20},paletteKey: '', layoutCache:new LayoutCache(),merges:new MergeQueryCache()})
 const pixel=(x:number,y:number)=>Array.from(ctx.getImageData(x,y,1,1).data)
 console.log({closed,dead,pixels:[pixel(100,10),pixel(10,100),pixel(10,10)],bitmapBytes:renderer.bitmapBytes(),peak:renderer.stats.peakBitmapBytes,budget:renderer.budgetBytes})
 renderer.dispose();expect(dead).toEqual([])
})
