export interface AppearanceStep {effect:string;safeRasters:boolean;opacity?:number;masked?:boolean;recolour?:boolean;timingPresent?:boolean}
export function reasonFor(s:AppearanceStep):string {
 if(s.effect==='static'&&s.timingPresent)return 'timing present without a proved static state'
 if(!s.safeRasters)return 'no final-observable-appearance proof for source rasters'
 if(s.opacity!==undefined&&s.opacity!==1)return 'opacity modifies appearance; original raster must not ship'
 if(s.masked||s.recolour)return 'mask/recolour requires server appearance projection'
 if(['static','visibility','transition'].includes(s.effect))return 'only already-projected safe appearance rasters are exposed'
 return 'no analytic continuous swept-area proof; sampled times are not a proof'
}
export function classifyStep(s:AppearanceStep):'proof-subset'|'step-composited' {
 return !(s.effect==='static'&&s.timingPresent)&&s.safeRasters&&(s.opacity===undefined||s.opacity===1)&&!s.masked&&!s.recolour&&['static','visibility','transition'].includes(s.effect)?'proof-subset':'step-composited'
}
