/** Lazy chart display formatting. The caller admits retained output; this adapter owns transient work. */
import { formatExtended, NO_SECTION, BUILTIN_EXTENDED_FORMATS } from '../xlsx/cellFormat'
import type { ChartFormatter } from './format-api'
import type { ChartPoints } from './model'
import { POINT_BLANK } from './points'
import { ChartRefusal } from './own'
import { createChartWorkspace, type ChartWorkspace } from './workspace'
export interface ChartNumberFormat { formatCode?: string; sourceLinked?: boolean }
const general=(n:number)=>String(Number(n.toPrecision(15)))
let serial=0
export function createChartFormatter(workspace: ChartWorkspace = createChartWorkspace('mobile')): ChartFormatter {
  return { format(value, code, epoch) {
    const tag=`chart-format#${++serial}`
    let live=0, admitted=0
    const allocation={maxOutput:256, allocate(bytes:number){
      live+=bytes
      if(live>admitted){if(!workspace.reserve(live-admitted,tag))throw new ChartRefusal();admitted=live}
    },release(bytes:number){live-=bytes}}
    try {
      allocation.allocate(1024)
      if(!code || code.length>255 || /^general$/i.test(code))return general(value).slice(0,256)
      const text=formatExtended(value,code,{dateSystem:epoch,unicode:{version:2,source:'standalone-default'},locale:'en-US',timeZone:'UTC',epochNowMs:0},general,allocation)
      return (text===NO_SECTION?'#':text??general(value)).slice(0,256)
    } finally {workspace.release(tag)}
  }}
}
export function pointFormat(format: ChartNumberFormat, points: ChartPoints | undefined, index: number): string | undefined {
  if(format.sourceLinked!==true && format.formatCode!==undefined)return format.formatCode
  if(format.sourceLinked===true){const k=points?.formatIndex?.[index];return k?points?.formats?.[k-1]:undefined}
  return undefined
}
/** First nonblank in the first bound series is the specified, Office-unverified mixed-format axis rule. */
export function axisFormat(format: ChartNumberFormat, inputs: readonly ChartPoints[]): string | undefined {
  if(format.sourceLinked!==true)return format.formatCode
  const points=inputs[0]
  if(points)for(let i=0;i<points.length;i++)if(points.states[i]!==POINT_BLANK)return pointFormat(format,points,i)
  return undefined
}

/** Renderer-compatible source codes, without formatting/materializing the cell value. */
export function sourceFormat(style: {numFmtId:number;formatCode?:string}|undefined):string|undefined {
  if(!style)return undefined
  if(style.formatCode!==undefined)return style.formatCode.length>255?'General':style.formatCode
  return BUILTIN_EXTENDED_FORMATS[style.numFmtId] ?? BUILTIN_SOURCE_FORMATS[style.numFmtId]
}

const BUILTIN_SOURCE_FORMATS:Record<number,string>=({2:'0.00',3:'#,##0',4:'#,##0.00',9:'0%',10:'0.00%',14:'m/d/yyyy',15:'d-mmm-yy',16:'d-mmm',17:'mmm-yy',18:'h:mm AM/PM',19:'h:mm:ss AM/PM',20:'h:mm',21:'h:mm:ss',22:'m/d/yyyy h:mm',49:'@'} as Record<number,string>)
