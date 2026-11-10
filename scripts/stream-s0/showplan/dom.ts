/** Task-local napi bridge for the existing DOM-driven local-player rig. */
import {createCanvas,Image,ImageData} from '../napi-shim'
export async function installNapiDom(){
 const name='jsdom', {JSDOM}=await import(name),dom=new JSDOM('<!doctype html><html><body></body></html>',{pretendToBeVisual:true,url:'http://localhost/'})
 const win=dom.window,g=globalThis as any
 const realm=new Set(['AbortController','AbortSignal','Event','EventTarget','CustomEvent'])
 for(const key of Object.getOwnPropertyNames(win)){if(key in g&&!realm.has(key))continue;try{g[key]=win[key]}catch{}}
 g.window=win;g.document=win.document;g.Image=Image;g.ImageData=ImageData;win.Image=Image;win.ImageData=ImageData
 const surfaces=new WeakMap<object,ReturnType<typeof createCanvas>>()
 function surface(el:any){let c=surfaces.get(el);if(!c){c=createCanvas(el.width||1,el.height||1);surfaces.set(el,c)}if(c.width!==el.width)c.width=Math.max(1,el.width);if(c.height!==el.height)c.height=Math.max(1,el.height);return c}
 win.HTMLCanvasElement.prototype.getContext=function(kind:string){if(kind!=='2d')return null;const c=surface(this),ctx=c.getContext('2d'),el=this
 return new Proxy(ctx,{get(target,key){if(key==='canvas')return el;const v=Reflect.get(target,key,target);if(typeof v!=='function')return v;if(key==='drawImage'||key==='createPattern')return (...args:any[])=>{if(args[0] instanceof win.HTMLCanvasElement)args[0]=surface(args[0]);return (v as Function).apply(target,args)};return v.bind(target)},set(target,key,value){return Reflect.set(target,key,value,target)}})}
 win.HTMLCanvasElement.prototype.toDataURL=function(){return surface(this).toDataURL()}
 Object.defineProperty(win,'devicePixelRatio',{value:1,configurable:true});g.devicePixelRatio=1
 return ()=>dom.window.close()
}
