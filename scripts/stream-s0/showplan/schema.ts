export interface Bitmap {width:number;height:number;data:Uint8ClampedArray}
export interface OvspJson {
 version:1; width:number;height:number;
 scope:{session:string;slide:number;generation:number};
 layers:{id:string;assetId:string;bounds:{x:number;y:number;w:number;h:number};z:number;partTag:string}[];
 steps:{order:number;eligibility:'proof-subset'|'step-composited';reason:string;trigger:'arrive'|'click';hitRegions:{x:number;y:number;w:number;h:number}[]}[];
 keyframes:{step:number;t:number;layers:string[]}[];
 morph:never[];transition:never[];links:never[];
}
export const CAP=128*1024*1024
