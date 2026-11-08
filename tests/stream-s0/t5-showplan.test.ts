import {expect,it} from 'vitest'
import {readFileSync,existsSync} from 'node:fs'
import {resolve,dirname} from 'node:path'
import ts from 'typescript'
import {classifyStep} from '../../scripts/stream-s0/showplan/classify'
import {composeFrame} from '../../scripts/stream-s0/showplan/client'
import {serializePlan} from '../../scripts/stream-s0/showplan/serialize'
import {buildRenderPlan} from '../../src/slideshow/render/plan'
import {rasterLayer} from '../../src/slideshow/render/raster'
import {createSlideShowWorkspace} from '../../src/slideshow/workspace'
import {createCanvas} from '../../scripts/stream-s0/napi-shim'
it('client transitive executable imports reach no PPTX or slideshow reader',()=>{
 const seen=new Set<string>();function visit(file:string){if(seen.has(file))return;seen.add(file);expect(file).not.toMatch(/src\/(pptx\/|slideshow\/read\/)/)
 const ast=ts.createSourceFile(file,readFileSync(file,'utf8'),ts.ScriptTarget.Latest,true)
 function walk(n:ts.Node){let spec:string|undefined
 if(ts.isImportDeclaration(n)&&!n.importClause?.isTypeOnly&&ts.isStringLiteral(n.moduleSpecifier))spec=n.moduleSpecifier.text
 if(ts.isExportDeclaration(n)&&!n.isTypeOnly&&n.moduleSpecifier&&ts.isStringLiteral(n.moduleSpecifier))spec=n.moduleSpecifier.text
 if(ts.isCallExpression(n)&&(n.expression.kind===ts.SyntaxKind.ImportKeyword||n.expression.getText(ast)==='require')){expect(n.arguments[0]&&ts.isStringLiteral(n.arguments[0])).toBe(true);if(n.arguments[0]&&ts.isStringLiteral(n.arguments[0]))spec=n.arguments[0].text}
 if(spec?.startsWith('.')){const base=resolve(dirname(file),spec),target=[base,base+'.ts',base+'.tsx',base+'/index.ts'].find(existsSync);expect(target).toBeTruthy();visit(target!)}ts.forEachChild(n,walk)}walk(ast)}
 visit(resolve('scripts/stream-s0/showplan/client.ts'));expect(seen.size).toBeGreaterThan(2)
})
it('opacity-0 and unproven motion are step-composited; safe static qualifies',()=>{
 expect(classifyStep({effect:'visibility',safeRasters:true,opacity:0})).toBe('step-composited')
 expect(classifyStep({effect:'motion',safeRasters:true})).toBe('step-composited')
 expect(classifyStep({effect:'static',safeRasters:true})).toBe('proof-subset')
})
it('static slide composes identically to the local raster path',()=>{
 const slide:any={index:0,widthEmu:914400,heightEmu:914400,background:'#345678',shapes:[]}
 const localPlan=buildRenderPlan(slide,new Set(),new Map(),[])
 const ctx:any={scale:1,images:[],resolveFont:(s:string)=>s,canvas:createCanvas,workspace:createSlideShowWorkspace('desktop'),maxTextureSize:8192}
 const local=rasterLayer(localPlan.layers[0]!,slide,ctx);expect('refused'in local).toBe(false);if('refused'in local)return
 const encoded=serializePlan(slide,localPlan,ctx)
 const result=composeFrame(encoded.plan,encoded.assets,0,1)
 const c:any=local.canvas;expect(result.data).toEqual(new Uint8ClampedArray(c.getContext('2d').getImageData(0,0,c.width,c.height).data));local.lease.release()
 expect(JSON.stringify(encoded.plan)).not.toMatch(/shapes|textBody|nodes/)
})

import {serializeFrames} from '../../scripts/stream-s0/showplan/serialize'
it('unsafe original target layers never ship beside fallback frames',()=>{
 expect(()=>serializePlan({} as any,{layers:[{kind:'target',key:'secret',nodes:[]}]},{} as any)).toThrow(/appearance projection/)
 const bitmap={width:1,height:1,data:Uint8ClampedArray.of(0,0,0,255)}
 const e=serializeFrames([{t:0,bitmap}],'step-composited','opacity/recolour projection')
 expect([...e.assets.keys()]).toEqual(['appearance-0'])
 expect(e.plan.layers.every(l=>l.partTag==='final-appearance')).toBe(true)
})
it('client rejects an unknown OVSP version before compositing',()=>{
 const bitmap={width:1,height:1,data:Uint8ClampedArray.of(0,0,0,255)},e=serializeFrames([{t:0,bitmap}],'proof-subset','static')
 expect(()=>composeFrame({...e.plan,version:2} as any,e.assets,0,1)).toThrow(/version/)
})
it('unlisted timing cannot qualify as static merely because effects are empty',()=>{
 expect(classifyStep({effect:'static',safeRasters:true,timingPresent:true} as any)).toBe('step-composited')
})
it('client refuses an unavailable step of a multi-step slide plan',()=>{
 const bitmap={width:1,height:1,data:Uint8ClampedArray.of(0,0,0,255)},e=serializeFrames([{t:0,bitmap}],'step-composited','projected')
 e.plan.steps.push({...e.plan.steps[0]!,order:1});e.plan.keyframes[0]!.step=1
 expect(()=>composeFrame(e.plan,e.assets,0,1)).toThrow(/No released frame/)
 expect(composeFrame(e.plan,e.assets,1,1).data).toEqual(bitmap.data)
})
