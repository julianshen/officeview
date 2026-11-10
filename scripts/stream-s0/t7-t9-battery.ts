/** Offline Stage-0 measurement only. No product source or harness changes. */
import { createHash } from 'node:crypto'
import { pathToFileURL } from 'node:url'
import { listCorpus, loadUnits, renderUnit } from './corpus'
import { writeResult } from './results'
import { messageFor, insertMessage } from './mark/message'
import { extendedCodeword } from './mark/bch'
import { basis } from './mark/basis'
import { embed } from './mark/embed'
import { sweepParams } from './mark/params'
import { pageDeltaE } from './mark/ciede2000'
import { scoreCapture, decideScores, prepareCandidates, type Scores } from './mark/detect'
import { jpeg, scale, crop, average, median, difference } from './mark/attacks'
import type { Bitmap, MarkParams } from './mark/types'
export interface Page {id:string; kind:string; bitmap:Bitmap}
const key=(s:string)=>new Uint8Array(createHash('sha256').update(s).digest())
const kMsg=key('stage0-t8-public-reproducibility-message-key')
const quantile=(a:number[],q:number)=>{const v=[...a].sort((a,b)=>a-b);return v[Math.max(0,Math.ceil(q*v.length)-1)] ?? 0}
// Abramowitz-Stegun normal CDF; numerical inversion makes tail extrapolation
// explicit. This model is not a distribution-free empirical certificate.
function cdf(z:number):number {
  const x=Math.abs(z),t=1/(1+0.2316419*x),d=Math.exp(-x*x/2)/Math.sqrt(2*Math.PI)
  const tail=d*t*(0.319381530+t*(-0.356563782+t*(1.781477937+t*(-1.821255978+t*1.330274429))))
  return z>=0?1-tail:tail
}
function inv(p:number):number {let lo=-12,hi=12;for(let i=0;i<70;i++){const mid=(lo+hi)/2;if(cdf(mid)<p)lo=mid;else hi=mid}return (lo+hi)/2}
function tailFit(samples:number[]) {
  const sorted=samples.filter(Number.isFinite).sort((a,b)=>a-b)
  if(sorted.length<3)return {mean:0,sigma:0,threshold:0,n:sorted.length,rms:0}
  const start=Math.floor(sorted.length*0.9), xs:number[]=[],ys:number[]=[]
  for(let i=start;i<sorted.length;i++){xs.push(inv((i+0.5)/sorted.length));ys.push(sorted[i]!)}
  const mx=xs.reduce((a,b)=>a+b,0)/xs.length,my=ys.reduce((a,b)=>a+b,0)/ys.length
  let xx=0,xy=0;for(let i=0;i<xs.length;i++){xx+=(xs[i]!-mx)**2;xy+=(xs[i]!-mx)*(ys[i]!-my)}
  const sigma=Math.max(0,xx?xy/xx:0),mean=my-sigma*mx
  const rms=Math.sqrt(ys.reduce((sum,y,i)=>sum+(y-mean-sigma*xs[i]!)**2,0)/ys.length)
  return {mean,sigma,threshold:mean+sigma*inv(1-1e-6),n:sorted.length,rms}
}
export function calibrate(samples:{best:number;margin:number}[],m:number) {
  const absFit=tailFit(samples.map(x=>x.best)),marginFit=tailFit(samples.map(x=>x.margin))
  const union=Math.sqrt(2*Math.log(m/1e-6)),union24=Math.sqrt(2*Math.log(2**24/1e-6))
  const tau={abs:Math.max(union,absFit.threshold),margin:Math.max(0,marginFit.threshold)}
  const tail=(v:number,f:ReturnType<typeof tailFit>)=>f.sigma?1-cdf((v-f.mean)/f.sigma):0
  return {tau,absFit,marginFit,union,union24,tau24:Math.sqrt(tau.abs**2+2*Math.log(2**24/m)),fittedAbsTail:tail(tau.abs,absFit),fittedMarginTail:tail(tau.margin,marginFit),cdf:[0.5,0.9,0.95,0.99,1].map(q=>({q,best:quantile(samples.map(x=>x.best).filter(Number.isFinite),q),margin:quantile(samples.map(x=>x.margin),q)})),caveat:'Upper-decile Gaussian QQ extrapolation; 2000 decisions do not empirically certify 1e-6. Union floor assumes sub-Gaussian candidate scores; structured BCH/registration nulls must be checked, not asserted independent.'}
}
const attacks: {name:string; target:number|null; apply:(b:Bitmap)=>Bitmap|Promise<Bitmap>}[]=[
  {name:'default',target:.9,apply:b=>b},...[95,90,80,70,30,10].map(q=>({name:`q${q}`,target:q===70?.75:null,apply:(b:Bitmap)=>jpeg(b,q)})),
  {name:'up2',target:null,apply:b=>scale(b,2)},{name:'down0.5',target:.75,apply:b=>scale(b,.5)},
  {name:'q70-down0.5',target:.75,apply:async b=>scale(await jpeg(b,70),.5)},
  {name:'crop5',target:null,apply:b=>crop(b,.05)},{name:'crop25',target:.7,apply:b=>crop(b,.25)},
  {name:'crop25-down0.5',target:null,apply:b=>scale(crop(b,.25),.5)},
  {name:'crop25-up2',target:null,apply:b=>scale(crop(b,.25),2)},
]
const empty=():Scores=>({scores:new Float64Array(),support:0,flat:true})
function summary(s:Scores){let best=-Infinity,second=-Infinity;for(const z of s.scores){if(z>best){second=best;best=z}else if(z>second)second=z}return {best,margin:Number.isFinite(second)?best-second:0}}
interface Options {candidateCount?:number;nullCount?:number;tokens?:number;params?:MarkParams[];progress?:(s:string)=>void;/** reuse a calibrated tau (subset runs); nullCount should then be 0 */tau?:{abs:number;margin:number}}
export async function runBattery(pages:Page[],options:Options={}) {
  const M=options.candidateCount??65536,N=options.nullCount??2000,tokens=options.tokens??5,log=options.progress??console.log
  if(!pages.length||M<6||M>2**24)throw new Error('Invalid battery size')
  const used=new Set<string>(),codes:Int8Array[]=[]
  for(let i=0;i<M;i++) {let salt=i%256;const row=insertMessage(kMsg,1,i,used,()=>salt++%256);codes.push(extendedCodeword(row.msg40))}
  const matrix=prepareCandidates(codes), floor={abs:Math.sqrt(2*Math.log(M/1e-6)),margin:0}
  log(`prepared ${M} keyed BCH candidates (${matrix.bytes} bytes)`)
  const sweep=[]
  for(const params of options.params??[...sweepParams()]) {
    const measurements=[];let correct=0
    for(const page of pages) {
      try {
        const chips=basis(key(`sweep:${page.id}`),page.bitmap.width,page.bitmap.height,params)
        const marked=embed(page.bitmap,chips,codes[0]!,params), delta=pageDeltaE(page.bitmap,marked)
        const s=scoreCapture(scale(await jpeg(marked,70),.5),page.bitmap,chips,matrix,params)
        const d=decideScores(s,floor);if(d.kind==='attributed'&&d.index===0)correct++
        measurements.push({id:page.id,...delta,carriers:params.nChips-marked.mark.maskedCells,pedestalCells:marked.mark.pedestalCells,variant:marked.mark.variant})
      }catch(e){measurements.push({id:page.id,error:String(e),carriers:0})}
    }
    const max=Math.max(...measurements.map(x=>'max'in x?x.max:0)),eligible=measurements.some(x=>x.carriers>0)&&max<=3
    sweep.push({params,measurements,pageMax:max,eligible,q70HalfCorrect:correct,q70HalfRate:correct/pages.length})
    log(`sweep ${sweep.length}/24 ${JSON.stringify(params)} maxDE=${max.toFixed(3)} q70-half=${correct}/${pages.length}`)
  }
  const choice=[...sweep].filter(x=>x.eligible).sort((a,b)=>b.q70HalfRate-a.q70HalfRate||a.pageMax-b.pageMax)[0]
  // If no setting is eligible, continue diagnostic measurements, never label
  // the fallback parameters chosen or the T7 gate GREEN.
  const params=choice?.params??sweep[0]!.params
  const nulls: {kind:string;page:string;basisKey:string;best:number;margin:number;support:number;scores:number;error?:string}[]=[]
  let scored=0
  for(let i=0;scored<N&&i<N*6;i++) {
    const page=pages[i%pages.length]!,id=`null:${i}`,k=key(id),kind=['unmarked','wrong-key','out-of-table'][i%3]!
    let s=empty(),error:string|undefined
    try {
      const chips=basis(k,page.bitmap.width,page.bitmap.height,params)
      let cap:Bitmap=page.bitmap
      if(kind==='wrong-key')cap=embed(page.bitmap,basis(key(`wrong:${i}`),page.bitmap.width,page.bitmap.height,params),codes[i%M]!,params)
      if(kind==='out-of-table')cap=embed(page.bitmap,chips,extendedCodeword(messageFor(kMsg,1,M+i,0)),params)
      // Vary acquisition and registration search, not only carrier keys.
      cap=await attacks[(Math.floor(i/3))%attacks.length]!.apply(cap)
      s=scoreCapture(cap,page.bitmap,chips,matrix,params)
    }catch(e){error=String(e)}
    if(s.scores.length===M)scored++
    nulls.push({kind,page:page.id,basisKey:Buffer.from(k).toString('hex'),...summary(s),support:s.support,scores:s.scores.length,...(error?{error}:{})})
    if((i+1)%25===0)log(`null decisions ${i+1}, nonflat full-M ${scored}/${N}`)
  }
  // Per-class fits protect the rarest class from being hidden by pooled fits.
  const fits=['all','unmarked','wrong-key','out-of-table'].map(kind=>({kind,...calibrate(kind==='all'?nulls:nulls.filter(x=>x.kind===kind),M)}))
  const tau=options.tau??{abs:Math.max(...fits.map(f=>f.tau.abs)),margin:Math.max(...fits.map(f=>f.tau.margin))}
  log(`calibration: full-M=${scored}/${N} attempts=${nulls.length} tau=${JSON.stringify(tau)}`)
  // Full multi-hit rule: exactly one threshold crossing, followed by margin.
  // Only top two are needed to re-evaluate that rule after calibration.
  const falseCount=nulls.filter(s=>s.best>=tau.abs&&s.best-s.margin<tau.abs&&s.margin>=tau.margin).length
  const suspectCount=nulls.filter(s=>s.best-s.margin>=tau.abs).length
  const captureResults: {page:string;token:number;attack:string;registeredOffset:{x:number;y:number}|null;trueOffset:{x:number;y:number};support:number;zTrue:number|null;zBest:number|null;decision:string;cause:string}[]=[]
  const survival=attacks.map(a=>({attack:a.name,target:a.target,total:0,correct:0,innocent:0,suspected:0,none:0,rate:0,byClass:{} as Record<string,{total:number;correct:number}>,status:'REPORTED'}))
  const collusion=[...[2,4,6].map(k=>({attack:`average-${k}`,k,apply:average})),...[3,4,5,6].map(k=>({attack:`median-${k}`,k,apply:median})),{attack:'difference-2',k:2,apply:(b:Bitmap[])=>difference(b[0]!,b[1]!)}].map(a=>({...a,total:0,colluder:0,suspected:0,innocent:0,none:0,colluderRate:0,suspectedRate:0,innocentRate:0}))
  for(const [pageIndex,page] of pages.entries()) {
    let chips:ReturnType<typeof basis>|undefined
    try{chips=basis(key(`survival:${page.id}`),page.bitmap.width,page.bitmap.height,params)}catch{}
    for(let token=0;token<tokens;token++) {
      const index=(pageIndex*tokens+token)%M, marked=chips?embed(page.bitmap,chips,codes[index]!,params):undefined
      for(let a=0;a<attacks.length;a++) {
        const row=survival[a]!,cls=row.byClass[page.kind]??={total:0,correct:0};row.total++;cls.total++
        const s=chips&&marked?scoreCapture(await attacks[a]!.apply(marked),page.bitmap,chips,matrix,params):empty(),d=decideScores(s,tau)
        const frac=attacks[a]!.name.startsWith('crop25')?.25:attacks[a]!.name==='crop5'?.05:0
        const trueOffset={x:Math.floor((page.bitmap.width-Math.floor(page.bitmap.width*(1-frac)))/2),y:Math.floor((page.bitmap.height-Math.floor(page.bitmap.height*(1-frac)))/2)}
        const registeredOffset=s.registration?{x:s.registration.x,y:s.registration.y}:null
        const correct=d.kind==='attributed'&&d.index===index
        const cause=correct?'correct':!chips?'capacity':!registeredOffset?'registration-missing':Math.abs(registeredOffset.x-trueOffset.x)>.51||Math.abs(registeredOffset.y-trueOffset.y)>.51?'registration-offset':s.flat?'flat-residual':d.kind==='collusion-suspected'?'multiple-hits':'score-or-margin'
        if(attacks[a]!.name.startsWith('crop'))log(`capture ${page.id} token=${index} ${attacks[a]!.name} pose=${JSON.stringify(registeredOffset)} support=${s.support} zTrue=${s.scores[index]??'flat'} ${cause}`)
        captureResults.push({page:page.id,token:index,attack:attacks[a]!.name,registeredOffset,trueOffset,support:s.support,zTrue:Number.isFinite(s.scores[index])?s.scores[index]!:null,zBest:Number.isFinite(summary(s).best)?summary(s).best:null,decision:d.kind,cause})
        if(d.kind==='attributed'){if(d.index===index){row.correct++;cls.correct++}else row.innocent++}else if(d.kind==='collusion-suspected')row.suspected++;else row.none++
      }
    }
    const offset=(pageIndex*7)%M,ids=Array.from({length:6},(_,i)=>(offset+i)%M)
    const captures=chips?ids.map(i=>embed(page.bitmap,chips!,codes[i]!,params)):[]
    for(const row of collusion) {
      row.total++
      const s=chips?scoreCapture(row.apply(captures.slice(0,row.k)),page.bitmap,chips,matrix,params):empty(),d=decideScores(s,tau)
      if(d.kind==='attributed'){if(ids.slice(0,row.k).includes(d.index))row.colluder++;else row.innocent++}else if(d.kind==='collusion-suspected')row.suspected++;else row.none++
    }
    log(`survival+collusion ${pageIndex+1}/${pages.length} ${page.id}`)
  }
  for(const row of survival){row.rate=row.correct/row.total;row.status=row.target===null?'REPORTED':row.rate>=row.target?'GREEN':'RED'}
  for(const row of collusion){row.colluderRate=row.colluder/row.total;row.suspectedRate=row.suspected/row.total;row.innocentRate=row.innocent/row.total}
  const counts=Object.fromEntries(['docx','xlsx','pptx'].map(k=>[k,pages.filter(p=>p.kind===k).length]))
  const coverage={units:pages.length,counts,imageHeavyVerified:0,shortfall:{units:Math.max(0,40-pages.length),docx:Math.max(0,15-counts.docx!),xlsx:Math.max(0,10-counts.xlsx!),pptx:Math.max(0,10-counts.pptx!),imageHeavy:5},note:'All corpus units retained, including insufficient-carrier failures. No synthetic padding. Image-heavy unit classification has not been independently verified; minimum is NOT established.'}
  return {
    t7:{status:choice?'GREEN':'RED',coverage,pages:pages.map(p=>({id:p.id,kind:p.kind,width:p.bitmap.width,height:p.bitmap.height})),sweep,chosen:choice?.params??null,diagnosticParams:params,selection:'highest q70-half correct rate at union floor, one fixed token/page; final survival recalculated with fitted tau and five tokens/page',pageGate:'max over all pixels of all pages <=3'},
    t9:{status:choice&&scored>=N&&falseCount===0&&survival.every(r=>r.status!=='RED')?'GREEN':'RED',coverage,candidateCount:M,matrixBytes:matrix.bytes,fullScoredDecisions:scored,candidateScores:nulls.reduce((a,s)=>a+s.scores,0),nulls,fits,tau,tau24:Math.sqrt(tau.abs**2+2*Math.log(2**24/M)),falseCount,suspectCount,zeroFalse95Upper:1-Math.pow(.05,1/nulls.length),captures:captureResults,survival,collusion:collusion.map(({apply,...row})=>row),scope:'Layer-1 Stage-0 only, no Layer-2 consensus claim. Seed/key are public reproducibility inputs, not production secrets. Crops remove the stated fraction of each axis; centered. No wall-clock gate.'},
  }
}
async function main() {
  const pages:Page[]=[],failures:{id:string;reason:string}[]=[]
  for(const file of listCorpus()) {
    try {
      const units=await loadUnits(file)
      try {for(const unit of units){const id=`${file.name}#${unit.index}`;const bitmap=await renderUnit(unit,1);if('skipped'in bitmap)failures.push({id,reason:bitmap.skipped});else pages.push({id,kind:file.kind,bitmap})}}finally{units.dispose()}
      console.log(`loaded ${file.name}, ${pages.length} total units`)
    }catch(e){failures.push({id:file.name,reason:String(e)})}
  }
  const result=await runBattery(pages)
  writeResult('t7',{...result.t7,status:result.t7.status as 'GREEN'|'RED',loadFailures:failures})
  writeResult('t9',{...result.t9,status:result.t9.status as 'GREEN'|'RED',loadFailures:failures})
  console.log(JSON.stringify({t7:result.t7.status,t9:result.t9.status,params:result.t7.chosen,tau:result.t9.tau,survival:result.t9.survival,collusion:result.t9.collusion}))
}
if(process.argv[1]&&import.meta.url===pathToFileURL(process.argv[1]).href)await main()
