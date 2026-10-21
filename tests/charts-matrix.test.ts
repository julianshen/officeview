// @vitest-environment node
import { expect,it } from 'vitest'
import { plotPoint } from '../src/charts/plotPoint'
import type { ChartFamily,ChartPoint } from '../src/charts/model'
const points:ChartPoint[]=[{state:'blank'},{state:'error',code:'#N/A'},{state:'error',code:'#DIV/0!'},{state:'unavailable'}]
for(const family of ['line','scatter','bar','area','pie','doughnut'] as ChartFamily[])for(const mode of ['gap','zero','span'] as const){
 it(`${family} ${mode}: all four non-value matrix cells`,()=>{
  const blank=mode==='zero'?0:mode
  const want=family==='line'||family==='scatter'?[blank,'span','span','gap']:family==='area'?[mode==='gap'?'gap':0,mode==='gap'?'gap':0,mode==='gap'?'gap':0,mode==='gap'?'gap':0]:['gap','gap','gap','gap']
  expect(points.map(p=>plotPoint(p,family,'standard',mode,false))).toEqual(want)
  expect(plotPoint(points[1],family,'standard',mode,true)).toBe(want[0])
 })
}
it('#N/A spans a line regardless of dispBlanksAs',()=>{for(const mode of ['gap','zero','span'] as const)expect(plotPoint(points[1],'line','standard',mode,false)).toBe('span')})
it('stacked and 100% override every non-value state',()=>{for(const grouping of ['stacked','percentStacked'] as const)for(const family of ['bar','line','area'] as const)expect(points.map(p=>plotPoint(p,family,grouping,'gap',false))).toEqual([0,0,0,0])})
import {parseXmlOrdered} from '../src/core/xml'
import {parseXlsxChart} from '../src/charts/parse'
import {paintXlsxChart} from '../src/charts/paint'
import {createCanvas} from 'canvas'
import {prepareChartAxes} from '../src/charts/prepare'
import {PointWriter} from '../src/charts/points'
function painted(family:string,middle:string,mode:string,log=false){
 const values=`<numLit><pt idx="0"><v>2</v></pt>${middle?`<pt idx="1"><v>${middle}</v></pt>`:''}<pt idx="2"><v>4</v></pt></numLit>`
 const m=parseXlsxChart(parseXmlOrdered(`<chartSpace><chart><dispBlanksAs val="${mode}"/><plotArea><${family}Chart>${family==='scatter'?'<scatterStyle val="line"/>':''}<ser><spPr><ln><solidFill><srgbClr val="FF0000"/></solidFill></ln></spPr><${family==='scatter'?'yVal':'val'}>${values}</${family==='scatter'?'yVal':'val'}></ser></${family}Chart></plotArea></chart></chartSpace>`),undefined)!
 expect(m.unsupported).toBeUndefined();if(log){const binding=m.groups[0].binding;if(binding.kind!=='xy')throw Error('not xy');m.axes.find(axis=>axis.id===binding.y)!.logBase=10; m.preparedAxes=prepareChartAxes({...m,preparedAxes:undefined}).preparedAxes}
 const c=createCanvas(400,300).getContext('2d'),paths:string[][]=[];let path:string[]=[]
 const begin=c.beginPath.bind(c),move=c.moveTo.bind(c),line=c.lineTo.bind(c),stroke=c.stroke.bind(c)
 c.beginPath=()=>{path=[];begin()};c.moveTo=(x,y)=>{path.push('M');move(x,y)};c.lineTo=(x,y)=>{path.push('L');line(x,y)};c.stroke=()=>{if(c.strokeStyle==='#ff0000')paths.push([...path]);stroke()}
 paintXlsxChart(m,c as any,400,300);return paths
}
for(const family of ['line','scatter'])for(const mode of ['gap','zero','span'])it(`${family} ${mode} native stroke oracle for blanks and errors`,()=>{
 expect(painted(family,'#N/A',mode)).toContainEqual(['M','L'])
 expect(painted(family,'',mode)).toContainEqual(mode==='gap'?['M','M']:mode==='zero'?['M','L','L']:['M','L'])
})
it('unavailable plot inputs publish the required diagnostic',()=>{
 const m=parseXlsxChart(parseXmlOrdered('<chartSpace><chart><plotArea><lineChart><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></lineChart></plotArea></chart></chartSpace>'),undefined)!
 const p=new PointWriter(2);p.value(0,1);p.unavailable(1);m.groups[0].series[0].values=p.done()
 expect(prepareChartAxes(m).preparedAxes!.diagnostics).toContainEqual(expect.objectContaining({code:'chart-point-unavailable'}))
})

it('M1: a nonpositive scatter value on a log axis breaks the line',()=>{expect(painted('scatter','-1','span',true)).toContainEqual(['M','M'])})
