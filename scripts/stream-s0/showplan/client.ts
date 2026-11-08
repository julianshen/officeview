import {createCanvas} from '../napi-shim'
import {createCanvas2DCompositor} from '../../../src/slideshow/render/canvas2d'
import {createSlideShowWorkspace} from '../../../src/slideshow/workspace'
import {CAP,type OvspJson,type Bitmap} from './schema'
/** Package-free Stage-0 client. Fallback samples are held, never interpolated
 * with unreleased original layers. t is normalized within the selected step. */
export function composeFrame(plan:OvspJson,assets:ReadonlyMap<string,Bitmap>,stepIndex:number,t:number):Bitmap {
 if(plan.version!==1)throw new Error('Unsupported OVSP version')
 if(!Number.isFinite(t)||t<0||t>1||!Number.isInteger(stepIndex)||!plan.steps[stepIndex])throw new Error('Invalid step/time')
 if(plan.layers.length>128||plan.steps.length>64||plan.keyframes.length>512)throw new Error('OVSP record cap')
 let bytes=JSON.stringify(plan).length*2
 for(const b of assets.values()){if(!Number.isSafeInteger(b.width)||!Number.isSafeInteger(b.height)||b.width<1||b.height<1||b.data.length!==4*b.width*b.height)throw new Error('Invalid bitmap');bytes+=b.data.byteLength}
 if(bytes>CAP)throw new Error('OVSP combined admission refused')
 if(!Number.isSafeInteger(plan.width)||!Number.isSafeInteger(plan.height)||plan.width<1||plan.height<1||plan.width*plan.height*4>CAP/3)throw new Error('Output admission refused')
 const frame=plan.keyframes.filter(k=>k.step===stepIndex&&k.t<=t).sort((a,b)=>b.t-a.t)[0];if(!frame)throw new Error('No released frame')
 const workspace=createSlideShowWorkspace('desktop',{limit:CAP}),factory=createCanvas as any
 const outputLease=workspace.lease(plan.width*plan.height*4,'output');if(!outputLease)throw new Error('Output refused')
 const canvas=createCanvas(plan.width,plan.height),compositor=createCanvas2DCompositor(canvas as any,factory,workspace)
 try {
 const draws=[]
 for(const layer of [...plan.layers].sort((a,b)=>a.z-b.z)) {
  if(!frame.layers.includes(layer.id))continue
  const b=assets.get(layer.assetId);if(!b)throw new Error('Missing asset')
  const lease=workspace.lease(b.data.byteLength*2,'asset upload');if(!lease)throw new Error('Asset refused')
  const c=createCanvas(b.width,b.height),ctx=c.getContext('2d'),im=ctx.createImageData(b.width,b.height);im.data.set(b.data);ctx.putImageData(im,0,0)
  const handle=compositor.upload({canvas:c as any,bounds:layer.bounds,scale:1,bytes:b.data.byteLength},lease)
  if('refused'in handle){lease.release();throw new Error('Upload refused')}
  draws.push({handle,state:{visible:true,opacity:1,transform:[1,0,0,1,0,0] as const}})
 }
 compositor.frame(draws)
 return {width:canvas.width,height:canvas.height,data:new Uint8ClampedArray(canvas.getContext('2d').getImageData(0,0,canvas.width,canvas.height).data)}
 }finally{compositor.dispose();outputLease.release()}
}
