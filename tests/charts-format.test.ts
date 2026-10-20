// @vitest-environment node
import { expect, it } from 'vitest'
import { createChartFormatter, axisFormat, pointFormat } from '../src/charts/format'
import { createChartWorkspace } from '../src/charts/workspace'
import { PointWriter } from '../src/charts/points'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
it('formats: scientific, fraction, elapsed, date, percent, locale prefix', () => {
  const f=createChartFormatter(createChartWorkspace('mobile'))
  for (const [v,code,want] of [[12345,'0.00E+00','1.23E+04'],[1.75,'# ?/?','1 3/4'],[1.5,'[h]:mm:ss','36:00:00'],[1,'yyyy-mm-dd','1900-01-01'],[.125,'0.0%','12.5%'],[12.5,'[$€-407]0.00','€12.50']] as const) expect(f.format(v,code,'1900')).toBe(want)
  expect(f.format(0,'yyyy-mm-dd','1904')).toBe('1904-01-01')
})
it('sourceLinked per-point label formats; axis uses the first non-blank (acceptance, unverified until reference)',()=>{
  const p=new PointWriter(3,true);p.value(1, .5,'0%');p.value(2,2,'0.00');const points=p.done()
  expect(pointFormat({formatCode:'0.0',sourceLinked:false},points,1)).toBe('0.0')
  expect(pointFormat({formatCode:'0.0',sourceLinked:true},points,1)).toBe('0%')
  expect(pointFormat({sourceLinked:true},points,2)).toBe('0.00')
  expect(axisFormat({sourceLinked:true},[points])).toBe('0%')
})
it('codes over 255 units are General; output capped at 256',()=>{
  const f=createChartFormatter(createChartWorkspace('mobile'))
  expect(f.format(12,'0'.repeat(256),'1900')).toBe('12')
  expect(f.format(1e308,'0'.repeat(255),'1900')).toHaveLength(256)
})
it('admits formatter work before allocation and refunds on success or refusal',()=>{
  let peak=0,held=0
  const w=createChartWorkspace('mobile',{reserve:b=>{held+=b;peak=Math.max(peak,held);return true},release:(_t,b)=>{held-=b}})
  createChartFormatter(w).format(12345,'0.00E+00','1900');expect(peak).toBeGreaterThan(2048);expect(held).toBe(0)
  const tight=createChartWorkspace('mobile',{reserve:()=>false,release:()=>{}})
  expect(()=>createChartFormatter(tight).format(12,'0.00','1900')).toThrow();expect(tight.used).toBe(0)
})
it('prepares explicit axis display formats and retains cache formatCode provenance',()=>{
  const m=parseXlsxChart(parseXmlOrdered('<chartSpace><chart><plotArea><barChart><ser><val><numLit><formatCode>0%</formatCode><pt idx="0"><v>0.5</v></pt></numLit></val></ser><axId val="1"/><axId val="2"/></barChart><catAx><axId val="1"/><crossAx val="2"/></catAx><valAx><axId val="2"/><crossAx val="1"/><numFmt formatCode="0.00E+00" sourceLinked="0"/></valAx></plotArea></chart></chartSpace>'),undefined)!
  expect(m.unsupported).toBeUndefined();expect(m.preparedAxes!.axes.find(a=>a.id==='2')!.labels.some(s=>s.includes('E+'))).toBe(true)
  expect(pointFormat({sourceLinked:true},m.groups[0].series[0].values,0)).toBe('0%')
})
it('mixed-source axis rule never borrows the second series format when the first is all blank',()=>{
 const p=new PointWriter(1,true);p.value(0,1,'0.00')
 expect(axisFormat({sourceLinked:true},[new PointWriter(1).done(),p.done()])).toBeUndefined()
})
it('explicit General overrides the percent-axis default',()=>{
 const m=parseXlsxChart(parseXmlOrdered('<chartSpace><chart><plotArea><barChart><grouping val="percentStacked"/><ser><val><numLit><pt idx="0"><v>3</v></pt></numLit></val></ser><axId val="1"/><axId val="2"/></barChart><catAx><axId val="1"/><crossAx val="2"/></catAx><valAx><axId val="2"/><crossAx val="1"/><numFmt formatCode="General" sourceLinked="0"/></valAx></plotArea></chart></chartSpace>'),undefined)!
 expect(m.axes.find(a=>a.id==='2')!.formatCode).toBe('General');expect(m.preparedAxes!.axes.find(a=>a.id==='2')!.labels.every(s=>!s.includes('%'))).toBe(true)
})
import {fillChartRanges} from '../src/charts/fill'
it('uncached worksheet values retain per-point source formats for labels',()=>{
 const m=parseXlsxChart(parseXmlOrdered('<chartSpace><chart><plotArea><barChart><dLbls><showVal val="1"/><numFmt sourceLinked="1"/></dLbls><ser><val><numRef><f>S!A1:A2</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>'),undefined)!
 const filled=fillChartRanges(m,'S',(_s,r)=>({value:r?.5:12,formatCode:r?'0%':'0.00'}))
 expect(filled.preparedAxes!.dataLabels.map(l=>l.text)).toEqual(['12.00','50%'])
})
it('General follows display numeric text without adding grouping separators',()=>{
 const f=createChartFormatter(createChartWorkspace('mobile'))
 expect(f.format(12345,'General','1900')).toBe('12345')
 expect(f.format(1.23456789012345,undefined,'1900')).toBe('1.23456789012345')
})
it('overlong cached source codes become General before entering the retained format table',()=>{
 const p=new PointWriter(1,true);p.value(0,12,'0'.repeat(256))
 expect(p.done().formats).toEqual(['General'])
})
