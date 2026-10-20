// @vitest-environment node
import {expect,it} from 'vitest'
import {parseXmlOrdered} from '../src/core/xml'
import {parseXlsxChart} from '../src/charts/parse'
import {createCanvas} from 'canvas'
import {paintXlsxChart} from '../src/charts/paint'
const xml=(group='',series='',point='',family='barChart',value='3')=>`<chartSpace><chart><plotArea><${family}>${group}<ser><tx><v>Sales</v></tx><cat><strLit><pt idx="0"><v>North</v></pt></strLit></cat><val><numLit><pt idx="0"><v>${value}</v></pt></numLit></val>${series}${point}</ser></${family}></plotArea></chart></chartSpace>`
const parse=(x:string)=>parseXlsxChart(parseXmlOrdered(x),undefined)!
const labels=(x:string)=>parse(x).preparedAxes?.dataLabels??[]
it('label precedence group then series then point; explicit false overrides; delete hides',()=>{
 const g='<dLbls><showVal val="1"/><showCatName val="1"/></dLbls>'
 const s='<dLbls><showVal val="0"/><dLbl><idx val="0"/><showSerName val="1"/></dLbl></dLbls>'
 expect(labels(xml(g,s))[0].text).toBe('North, Sales')
 expect(labels(xml(g,'<dLbls><dLbl><idx val="0"/><delete val="1"/></dLbl></dLbls>'))).toHaveLength(0)
})
it('content flags, separator, rich text, leader lines; dLblPos defaults; bestFit for pie',()=>{
 const l=labels(xml('<dLbls><showVal val="1"/><showCatName val="1"/><showSerName val="1"/><showPercent val="1"/><showLegendKey val="1"/><separator> | </separator></dLbls>'))[0]
 expect(l.text).toBe('North | Sales | 3');expect(l.legendKey).toBe(true);expect(l.position).toBe('outEnd')
 const rich='<dLbls><showVal val="1"/><showLeaderLines val="1"/><dLbl><idx val="0"/><tx><rich><p><r><rPr b="1"/><t>Custom</t></r></p></rich></tx></dLbl></dLbls>'
 const r=labels(xml('',rich,'','pieChart'))[0];expect(r.text).toBe('Custom');expect(r.runs[0].bold).toBe(true);expect(r.position).toBe('bestFit');expect(r.leaderLines).toBe(true)
})
it('field-reference and manual-layout labels fall back with chart-label-feature-unsupported',()=>{
 const m=parse(xml('','<dLbls><showVal val="1"/><dLbl><idx val="0"/><layout><manualLayout><x val="0.2"/></manualLayout></layout><extLst><ext><c15:datalabelsRange xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"/></ext></extLst></dLbl></dLbls>'))
 expect(m.preparedAxes!.diagnostics).toContainEqual(expect.objectContaining({code:'chart-label-feature-unsupported'}));expect(m.preparedAxes!.dataLabels[0].text).toBe('3')
})
it('a label on an #N/A point shows #N/A',()=>{expect(labels(xml('<dLbls><showVal val="1"/></dLbls>','','','lineChart','#N/A'))[0].text).toBe('#N/A')})
it('native painter draws prepared label text and legend key',()=>{
 const m=parse(xml('<dLbls><showVal val="1"/><showCatName val="1"/><showLegendKey val="1"/></dLbls>'));expect(m.unsupported).toBeUndefined()
 const c=createCanvas(400,300).getContext('2d'),seen:string[]=[];const f=c.fillText.bind(c);c.fillText=(t,x,y,...r)=>{seen.push(t);f(t,x,y,...r)}
 paintXlsxChart(m,c as any,400,300);expect(seen).toContain('North, 3')
})
import {prepareChartAxes} from '../src/charts/prepare'
import {layoutChart} from '../src/charts/layout'
import {PointWriter} from '../src/charts/points'
import {chartWorkbook} from './helpers/xlsx-chart-fixture'
import {parseXlsx} from '../src/xlsx/parse'
it('bubble size content and point-level formatting are prepared before paint',()=>{
 const m=parse(xml('<dLbls><showVal val="1"/><numFmt sourceLinked="1"/></dLbls>'))
 const values=new PointWriter(1,true);values.value(0,.25,'0.0%');m.groups[0].series[0].values=values.done()
 const sizes=new PointWriter(1);sizes.value(0,8);m.groups[0].series[0].bubbleSizes=sizes.done();m.groups[0].series[0].labels={showBubbleSize:true}
 expect(prepareChartAxes(m).preparedAxes!.dataLabels[0].text).toBe('25.0%, 8')
})
it('workbook coverage is native chart, with distinct bar geometry and prepared labels',async()=>{
 const doc=await parseXlsx(await chartWorkbook(xml('<dLbls><showCatName val="1"/><showVal val="1"/></dLbls>')))
 expect(doc.drawingCoverage).toContainEqual(expect.objectContaining({status:'native',selectedRepresentation:'chart'}))
 const content=doc.sheets[0].drawings![0].content!;expect(content.kind).toBe('xlsxChart');if(content.kind!=='xlsxChart')throw Error('fallback')
 expect(content.chart.unsupported).toBeUndefined()
 const c=createCanvas(400,300).getContext('2d'),rects:number[][]=[],texts:string[]=[];const rect=c.fillRect.bind(c),text=c.fillText.bind(c)
 c.fillRect=(x,y,w,h)=>{if(c.fillStyle!=='#ffffff')rects.push([x,y,w,h]);rect(x,y,w,h)};c.fillText=(t,x,y,...r)=>{texts.push(t);text(t,x,y,...r)}
 paintXlsxChart(content.chart,c as any,400,300);expect(rects.some(r=>r[2]>10&&r[3]>30)).toBe(true);expect(texts).toContain('North, 3')
})
it('bestFit moves small pie labels outside and draws the requested leader',()=>{
 const m=parse(xml('<dLbls><showVal val="1"/><showLeaderLines val="1"/></dLbls>','','','pieChart').replace('</numLit></val>','<pt idx="1"><v>97</v></pt></numLit></val>'))
 const c=createCanvas(400,300).getContext('2d'),lines:number[][]=[],labels=new Map<string,number[]>();const line=c.lineTo.bind(c),text=c.fillText.bind(c);c.lineTo=(x,y)=>{lines.push([x,y]);line(x,y)}
 c.fillText=(t,x,y,...rest)=>{if(t==='3'||t==='97')labels.set(t,[x+c.measureText(t).width/2,y]);text(t,x,y,...rest)}
 paintXlsxChart(m,c as any,400,300)
 const layout=layoutChart(m,400,300,1,0,{width:(t,font)=>{c.font=font;return c.measureText(t).width},lineHeight:()=>12})
 const box=layout.area,r=Math.min(box.w,box.h)/2-4,cx=box.x+box.w/2,cy=box.y+box.h/2
 const small=labels.get('3')!,large=labels.get('97')!
 expect(Math.hypot(small[0]-cx,small[1]-cy)).toBeCloseTo(r+12,4)
 expect(Math.hypot(large[0]-cx,large[1]-cy)).toBeCloseTo(r*.65,4)
 expect(lines).toContainEqual([expect.closeTo(small[0],4),expect.closeTo(small[1],4)])
})
it('category-name labels use each series own cached categories',()=>{
 const m=parse(xml('<dLbls><showCatName val="1"/></dLbls>').replace('</barChart>','<ser><cat><strLit><pt idx="0"><v>South</v></pt></strLit></cat><val><numLit><pt idx="0"><v>4</v></pt></numLit></val></ser></barChart>'))
 expect(m.preparedAxes!.dataLabels.map(l=>l.text)).toEqual(['North','South'])
})
it('field-reference-only labels fall back to the saved point value at the ordinary position',()=>{
 const m=parse(xml('','<dLbls><dLbl><idx val="0"/><tx><strRef><f>S!A1</f></strRef></tx></dLbl></dLbls>'))
 expect(m.preparedAxes!.dataLabels[0]).toMatchObject({text:'3',position:'outEnd'})
})
