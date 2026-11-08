import {readFileSync} from 'node:fs'
import {listCorpus} from './corpus'
import {writeResult} from './results'
import {installNapiDom} from './showplan/dom'
import {classifyStep,reasonFor} from './showplan/classify'
import {serializePlan,serializeFrames} from './showplan/serialize'
import {composeFrame} from './showplan/client'
import {buildRenderPlan} from '../../src/slideshow/render/plan'
import {createSlideShowWorkspace} from '../../src/slideshow/workspace'
import {createCanvas} from './napi-shim'
import type {Bitmap} from './showplan/schema'
function diff(a:Bitmap,b:Bitmap){if(a.width!==b.width||a.height!==b.height)return {fraction:1,pixels:a.width*a.height,sizeMismatch:true};let pixels=0;for(let i=0;i<a.data.length;i+=4)if(a.data[i]!==b.data[i]||a.data[i+1]!==b.data[i+1]||a.data[i+2]!==b.data[i+2]||a.data[i+3]!==b.data[i+3])pixels++;return {fraction:pixels/(a.width*a.height),pixels,sizeMismatch:false}}
const close=await installNapiDom()
const {installShowSeams,loadDeck,startShow}=await import('../../tests/helpers/slideshow-show-rig')
const {slideXmls}=await import('../../tests/helpers/slideshow-coverage')
const {planDeck}=await import('../slideshow-capture-plan')
const {sessionDeps}=await import('../../src/slideshow/player/session')
const decks:any[]=[]
try {
 for(const file of listCorpus().filter(f=>f.kind==='pptx').slice(0,process.env.S0_T5_LIMIT?Number(process.env.S0_T5_LIMIT):undefined)){
 const rows:any[]=[],blocked:any[]=[];let scheduledSteps=0;let seams:ReturnType<typeof installShowSeams>|undefined,rig:Awaited<ReturnType<typeof startShow>>|undefined
 try{
  const bytes=new Uint8Array(readFileSync(file.path)),schedule=await planDeck(file.name,bytes),doc=await loadDeck(bytes)
  scheduledSteps=schedule.slides.filter(s=>!s.hidden).reduce((n,s)=>n+s.steps.length,0)
  const xmls=await slideXmls(bytes)
  const hasTiming=(v:unknown):boolean=>!!v&&typeof v==='object'&&Object.entries(v).some(([k,x])=>k.split(':').pop()==='timing'||hasTiming(x))
  const width=Math.ceil(schedule.slideSize.cx/9525),height=Math.ceil(schedule.slideSize.cy/9525)
  seams=installShowSeams({width,height});sessionDeps.backend='canvas2d';rig=await startShow(doc,{seams,size:{width,height},motion:'full'})
  for(const slide of schedule.slides){
   if(slide.hidden){blocked.push({slide:slide.slide,reason:'hidden slide deliberately not delivered'});continue}
   const staticSlide=!hasTiming(xmls[slide.slide-1])&&slide.steps.length===1&&!slide.steps[0]!.effects.length&&!slide.steps[0]!.transition&&slide.interactive.length===0
   for(const step of slide.steps){
    if(step.step>0||rig.s.state.slide!==slide.slide-1){if(step.step===0)rig.s.goTo(slide.slide-1);else rig.s.next();await rig.tick(0)}
    const t0=seams.clock.now(),frames:{t:number;bitmap:Bitmap}[]=[]
    for(const sample of step.frames){await rig.tick(Math.max(0,t0+sample.ms-seams.clock.now()));if(rig.s.state.slide!==slide.slide-1)throw new Error('Local player advanced away from scheduled slide');const canvas=rig.frameCanvas(),ctx=canvas?.getContext('2d');if(!canvas||!ctx)throw new Error('No local player frame');frames.push({t:sample.pct/100,bitmap:{width:canvas.width,height:canvas.height,data:new Uint8ClampedArray(ctx.getImageData(0,0,canvas.width,canvas.height).data)}})}
    if(!frames.length)throw new Error('No bounded capture samples')
    const evidence={effect:staticSlide?'static':'unknown',safeRasters:staticSlide},eligibility=classifyStep(evidence),reason=staticSlide?'flattened static final appearance; no independent occluded layers':reasonFor(evidence)
    let encoded
    if(staticSlide){const preview=rig.s.presenterSource.previewOf(slide.slide-1);encoded=serializePlan(preview.slide,buildRenderPlan(preview.slide,new Set(),new Map(),[]),{scale:1,images:preview.images,resolveFont:preview.resolveFont,canvas:createCanvas as any,workspace:createSlideShowWorkspace('desktop'),maxTextureSize:8192})}
    else encoded=serializeFrames(frames,eligibility,reason,step.trigger)
    // One slide plan, with only the current released step's appearance assets.
    // Earlier/future steps have no available frame and the client refuses them.
    encoded.plan.scope={session:'stage0:'+file.name,slide:slide.slide-1,generation:step.step+1}
    encoded.plan.steps=slide.steps.map(s=>({order:s.step,eligibility,reason,trigger:s.trigger,hitRegions:[]}))
    const stepIndex=slide.steps.indexOf(step)
    for(const frame of encoded.plan.keyframes)frame.step=stepIndex
    const wire=JSON.parse(JSON.stringify(encoded.plan)),actual=composeFrame(wire,encoded.assets,stepIndex,1),comparison=diff(actual,frames[frames.length-1]!.bitmap)
    rows.push({slide:slide.slide,step:step.step,eligibility,reason,staticSlide,spanMs:step.spanMs,finalFrameMs:step.frames.at(-1)?.ms,...comparison,layers:wire.layers.length,assetBytes:[...encoded.assets.values()].reduce((n,b)=>n+b.data.byteLength,0),comparisonKind:staticSlide?'independent static raster vs local player':'client replay of server-composed local-player samples (transport/composition check only)'})
   }
   for(const id of slide.interactive)blocked.push({slide:slide.slide,trigger:id,reason:'interactive-trigger branch not captured by linear click-step schedule; step-composited, not counted as verified'})
  }
 }catch(e){blocked.push({reason:String(e)})}finally{rig?.exit();seams?.restore();document.body.replaceChildren()}
 const proof=rows.filter(r=>r.eligibility==='proof-subset').length
 decks.push({deck:file.name,steps:scheduledSteps,measuredSteps:rows.length,unmeasuredSteps:scheduledSteps-rows.length,proofSubset:proof,stepComposited:scheduledSteps-proof,eligibilityRate:scheduledSteps?proof/scheduledSteps:0,rows,blocked})
 console.log(`${file.name}: ${proof}/${rows.length} proof, blocked=${blocked.length}, worst=${Math.max(0,...rows.map(r=>r.fraction))}`)
 }
 const totalSteps=decks.reduce((n,d)=>n+d.steps,0)
 const rows=decks.flatMap(d=>d.rows),staticRows=rows.filter(r=>r.staticSlide),proof=rows.filter(r=>r.eligibility==='proof-subset').length
 writeResult('t5',{status:decks.some(d=>d.blocked.some((b:any)=>!b.reason.startsWith('hidden slide')))||!staticRows.length||staticRows.some(r=>r.fraction>.001)?'RED':'GREEN',decks,steps:totalSteps,measuredSteps:rows.length,unmeasuredSteps:totalSteps-rows.length,proofSubset:proof,stepComposited:totalSteps-proof,eligibilityRate:totalSteps?proof/totalSteps:0,staticSlides:staticRows.length,staticFailures:staticRows.filter(r=>r.fraction>.001).length,worstStaticDiff:Math.max(0,...staticRows.map(r=>r.fraction)),tolerance:.001,scope:'Static independent raster vs actual local player; unsafe animation intervals release only composed samples, never original layers. No analytic motion proof; interactive branches and unresolved durations are explicit limitations. No original PPTX, shapes, text, URLs or notes in client plan. Client graph gate is the separate automated test.'})
}finally{close()}
