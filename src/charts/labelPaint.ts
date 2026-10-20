/** Placement of prepared data-label strings. Formatting never happens during paint. */
import type {ChartModel} from './model'
import type {ChartLayout} from './layout'
import {chartFont} from './layout'
import {axisFraction,categoryFraction} from './axes'
import {valueAt,hiddenAt} from './points'
import type {LabelRun} from './labels'
export const labelFont=(run:LabelRun|undefined,px:number,family:string)=>`${run?.italic?'italic ':''}${run?.bold?'bold ':''}${chartFont((run?.sizePt??px*72/96)*96/72,family)}`
export function paintDataLabels(model:ChartModel,ctx:CanvasRenderingContext2D,layout:ChartLayout,family:string):void {
  const labels=model.preparedAxes?.dataLabels
  if(!labels?.length)return
  const {view,plot}=layout,slots=new Map<number,number>()
  for(let i=0;i<view.slots;i++)slots.set(view.source(i),i)
  const along=(a:NonNullable<(typeof layout.graph.groups)[number]>['value'],v:number)=>a.orientation==='h'?plot.x+axisFraction(a,v)*plot.w:plot.y+plot.h-axisFraction(a,v)*plot.h
  let pieSeries:object|undefined,pieBefore=0,pieIndex=0,pieTotal=0
  for(const label of labels){
    const g=model.groups[label.group],s=g.series[label.series],slot=slots.get(label.point)
    if(slot===undefined || (model.plotVisOnly!==false&&(hiddenAt(s.values.hidden,label.point)||hiddenAt(s.xValues?.hidden,label.point))))continue
    const raw=valueAt(s.values,label.point)??0
    let x=0,y=0,anchorX=0,anchorY=0,outside=false
    if(g.family==='pie'||g.family==='doughnut'){
      if(pieSeries!==s){pieSeries=s;pieBefore=0;pieIndex=0;pieTotal=0;for(let i=0;i<view.slots;i++)pieTotal+=Math.max(0,valueAt(s.values,view.source(i))??0)}
      while(pieIndex<slot){pieBefore+=Math.max(0,valueAt(s.values,view.source(pieIndex++))??0)}
      if(!(pieTotal>0)||raw<=0)continue
      const box=layout.area,r=Math.max(1,Math.min(box.w,box.h)/2-4),cx=box.x+box.w/2,cy=box.y+box.h/2
      const angle=((g.firstSliceAngle??0)-90)*Math.PI/180+(pieBefore+raw/2)/pieTotal*Math.PI*2
      outside=label.position==='outEnd'||(label.position==='bestFit'&&raw/pieTotal<.12)
      anchorX=cx+Math.cos(angle)*r;anchorY=cy+Math.sin(angle)*r
      x=cx+Math.cos(angle)*(outside?r+12:r*.65);y=cy+Math.sin(angle)*(outside?r+12:r*.65)
    }else{
      const a=layout.graph.groups[label.group];if(!a)continue
      let v=raw,base=0
      if(g.grouping==='stacked'||g.grouping==='percentStacked'){
        let total=0;for(let j=0;j<g.series.length;j++){const n=valueAt(g.series[j].values,label.point)??0;total+=Math.abs(n);if(j<label.series&&Math.sign(n)===Math.sign(raw))base+=n}
        if(g.grouping==='percentStacked'){v=total?v/total:0;base=total?base/total:0}
      }
      const fraction=a.category?categoryFraction(a.category,layout.graph.categoryIndex(a.category,slot),a.between):axisFraction(a.x!,valueAt(s.xValues,label.point)??slot+1)
      const category=a.horizontal===a.value?plot.y+plot.h-fraction*plot.h:plot.x+fraction*plot.w
      const start=along(a.value,base),end=along(a.value,base+v)
      let position=end
      if(label.position==='ctr')position=(start+end)/2
      else if(label.position==='inBase')position=start+(end>=start?6:-6)
      else if(label.position==='inEnd')position=end+(end>=start?-6:6)
      else if(label.position==='outEnd')position=end+(end>=start?8:-8)
      if(a.value.orientation==='h'){x=position;y=category}else{x=category;y=position}
      if(g.family==='bar'&&g.grouping==='clustered'){
        const band=(a.category?.orientation==='h'?plot.w:plot.h)/Math.max(1,view.slots)
        const stride=1-g.overlap/100,span=1+stride*(g.series.length-1),width=band/(span+g.gapWidth/100)
        const offset=(label.series-(g.series.length-1)/2)*width*stride
        if(a.value.orientation==='h')y-=offset;else x+=offset
      }
      if(label.position==='t')y-=8;if(label.position==='b')y+=8;if(label.position==='l')x-=8;if(label.position==='r')x+=8
    }
    ctx.save()
    if(outside&&label.leaderLines){ctx.strokeStyle='#595959';ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(anchorX,anchorY);ctx.lineTo(x,y);ctx.stroke()}
    ctx.textBaseline='middle';ctx.textAlign='left'
    const runs=label.runs.length?label.runs:[{text:label.text}]
    let width=label.legendKey?12:0
    for(const run of runs)width+=layout.width(run.text,labelFont(run,layout.fontPx,family))
    let at=x-width/2
    if(label.legendKey){ctx.fillStyle='#'+s.color;ctx.fillRect(at,y-4,8,8);at+=12}
    for(const run of runs){ctx.font=labelFont(run,layout.fontPx,family);ctx.fillStyle=run.color??'#595959';ctx.fillText(run.text,at,y);at+=layout.width(run.text,ctx.font)}
    ctx.restore()
  }
}
