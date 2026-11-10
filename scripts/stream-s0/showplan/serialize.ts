import type {PptxSlide} from '../../../src/pptx/types'
import type {RenderPlan} from '../../../src/slideshow/render/plan'
import {rasterLayer,type RasterContext} from '../../../src/slideshow/render/raster'
import {slideSize} from '../../../src/slideshow/render/scene'
import {createCanvas} from '../napi-shim'
import {CAP,type Bitmap,type OvspJson} from './schema'
/** Only final projected frames cross the prototype boundary, never original
 * rasters alongside an unsafe interval. One call releases one step's samples. */
export function serializeFrames(frames:{t:number;bitmap:Bitmap}[],eligibility:'proof-subset'|'step-composited',reason:string,trigger:'arrive'|'click'='arrive') {
 if(!frames.length||frames.length>128)throw new Error('Frame count cap')
 const {width,height}=frames[0]!.bitmap,assets=new Map<string,Bitmap>()
 const plan:OvspJson={version:1,width,height,scope:{session:"stage0",slide:0,generation:1},layers:[],steps:[{order:0,eligibility,reason,trigger,hitRegions:[]}],keyframes:[],morph:[],transition:[],links:[]}
 for(const [i,f]of frames.entries()){
  if(f.bitmap.width!==width||f.bitmap.height!==height)throw new Error('Frame dimensions changed')
  const id=`appearance-${i}`;plan.layers.push({id,assetId:id,bounds:{x:0,y:0,w:width,h:height},z:i,partTag:'final-appearance'})
  plan.keyframes.push({step:0,t:f.t,layers:[id]})
 }
 const bytes=JSON.stringify(plan).length*2+frames.reduce((n,f)=>n+f.bitmap.data.byteLength,0)
 if(bytes>CAP)throw new Error('OVSP combined admission refused')
 for(const [i,f]of frames.entries())assets.set(`appearance-${i}`,f.bitmap)
 return {plan,assets}
}
/** Static-only independent server raster route. Animated/part layers require
 * the local player's composed appearance projection instead of raw exports. */
export function serializePlan(slide:PptxSlide,plan:RenderPlan,ctx:RasterContext) {
 if(plan.layers.some(l=>l.kind==='target'))throw new Error('Animated layers require appearance projection')
 const size=slideSize(slide),w=Math.ceil(size.w),h=Math.ceil(size.h)
 if(w*h*8>CAP)throw new Error('Static projection admission refused')
 const hold=ctx.workspace.lease(w*h*8,'OVSP static projection');if(!hold)throw new Error('Projection refused')
 try{
 const output=createCanvas(w,h),c=output.getContext('2d')
 for(const layer of plan.layers){const r=rasterLayer(layer,slide,ctx);if('refused'in r){if(r.refused==='empty')continue;throw new Error('Raster refused')}
 try{c.drawImage(r.canvas as any,r.bounds.x,r.bounds.y,r.bounds.w,r.bounds.h)}finally{r.lease.release()}}
 const bitmap={width:w,height:h,data:new Uint8ClampedArray(c.getImageData(0,0,w,h).data)}
 return serializeFrames([{t:0,bitmap}],'proof-subset','fully flattened static appearance; no separate covered or off-slide pixels')
 }finally{hold.release()}
}
