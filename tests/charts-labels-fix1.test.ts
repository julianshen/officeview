// @vitest-environment node
import { expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { createChartWorkspace } from '../src/charts/workspace'
import { labelPreparationBound, labelBytes } from '../src/charts/labels'
import { releaseChartModel } from '../src/charts/own'

const flags = ['showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize', 'showLegendKey']
const controls = (on: boolean) => flags.map(flag => `<c:${flag} val="${on ? 1 : 0}"/>`).join('')
const points = (count: number) => Array.from({ length: count }, (_, i) => `<c:pt idx="${i}"><c:v>${i + 1}</c:v></c:pt>`).join('')
const chart = (labels: string, count = 2, series = 1) => `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"><c:chart><c:plotArea><c:lineChart><c:dLbls>${labels}</c:dLbls>${Array.from({ length: series }, () => `<c:ser><c:val><c:numLit>${points(count)}</c:numLit></c:val></c:ser>`).join('')}</c:lineChart></c:plotArea></c:chart></c:chartSpace>`
const parse = (xml: string) => parseXlsxChart(parseXmlOrdered(xml), undefined)!

it('I1: all-off Office labels on three 4096-point series remain native on mobile', () => {
  const workspace = createChartWorkspace('mobile')
  const model = parseXlsxChart(parseXmlOrdered(chart(controls(false), 4096, 3)), undefined, { workspace, profile: 'mobile' } as never)!
  expect(model.refused).not.toBe(true)
  expect(model.unsupported).toBeUndefined()
  expect(model.preparedAxes?.dataLabels).toEqual([])
  expect(labelPreparationBound(model)).toBe(0)
  const ctx = createCanvas(400, 300).getContext('2d')
  const stroke = ctx.stroke.bind(ctx), colors = new Set(model.groups[0].series.map(s => '#' + s.color.toLowerCase()))
  let seriesStrokes = 0
  ctx.stroke = () => { if (colors.has(ctx.strokeStyle.toString())) seriesStrokes++; stroke() }
  paintXlsxChart(model, ctx as any, 400, 300)
  expect(seriesStrokes).toBe(3)
  expect(model.unsupported).toBeUndefined()
  releaseChartModel(model)
  expect(workspace.used).toBe(0)
})

it('I1: all-on labels have a tight admitted bound and release their mobile ownership', () => {
  let peak = 0, live = 0
  const workspace = createChartWorkspace('mobile', { reserve(bytes) { live += bytes; peak = Math.max(peak, live); return true }, release(_tag, bytes) { live -= bytes } })
  const model = parseXlsxChart(parseXmlOrdered(chart(controls(true), 4096, 2)), undefined, { workspace, profile: 'mobile' } as never)!
  expect(model.refused).not.toBe(true)
  expect(model.preparedAxes?.dataLabels).toHaveLength(8192)
  expect(labelPreparationBound(model)).toBeGreaterThanOrEqual(labelBytes(model.preparedAxes!.dataLabels))
  expect(labelPreparationBound(model)).toBeLessThan(8192 * 1800)
  expect(peak).toBeLessThanOrEqual(workspace.limit)
  releaseChartModel(model)
  expect(live).toBe(0)
})

it('I2: Office 2013 leader-line extensions and empty layouts do not enable disabled pie labels', () => {
  const extension = '<c:extLst><c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}"><c15:leaderLines/><c15:showLeaderLines val="1"/></c:ext></c:extLst>'
  const model = parse(chart(`${controls(false)}<c:showLeaderLines val="1"/>${extension}<c:dLbl><c:idx val="0"/><c:layout/></c:dLbl>`).replace(/lineChart/g, 'pieChart'))
  expect(model.preparedAxes!.dataLabels).toEqual([])
  expect(model.preparedAxes!.diagnostics.some(d => d.code === 'chart-label-feature-unsupported')).toBe(false)
})

it('I2: disabled labels stay disabled even with a real unsupported text source', () => {
  const ext = '<c:extLst><c:ext><c15:datalabelsRange/><c15:showDataLabelsRange val="1"/></c:ext></c:extLst>'
  expect(parse(chart(controls(false) + ext)).preparedAxes!.dataLabels).toEqual([])
  expect(parse(chart(ext)).preparedAxes!.dataLabels.map(l => l.text)).toEqual(['1', '2'])
  const manual = '<c:dLbl><c:idx val="0"/><c:layout><c:manualLayout><c:x val="0.2"/></c:manualLayout></c:layout></c:dLbl>'
  expect(parse(chart(manual)).preparedAxes!.dataLabels).toEqual([])
})

it('M2: showPercent is pie/doughnut-only and source formats do not turn percent into a decimal', () => {
  expect(parse(chart('<c:showPercent val="1"/>')).preparedAxes!.dataLabels).toEqual([])
  const pie = chart('<c:showPercent val="1"/><c:numFmt sourceLinked="1"/>', 1).replace(/lineChart/g, 'pieChart').replace('<c:numLit>', '<c:numLit><c:formatCode>0.00</c:formatCode>')
  expect(parse(pie).preparedAxes!.dataLabels[0].text).toBe('100%')
})

it('M4: clustered labels align with the painted bars when crossBetween is midCat', () => {
  const model = parse(chart('<c:showSerName val="1"/><c:showVal val="1"/><c:separator>|</c:separator>', 2, 2).replace(/lineChart/g, 'barChart'))
  for (const axis of model.axes) if (axis.kind === 'value') axis.crossBetween = 'midCat'
  const prepared = prepareChartAxes({ ...model, preparedAxes: undefined })
  const ctx = createCanvas(400, 300).getContext('2d')
  const bars: number[] = [], texts: number[] = []
  const rect = ctx.fillRect.bind(ctx), text = ctx.fillText.bind(ctx)
  ctx.fillRect = (x, y, w, h) => { if (ctx.fillStyle !== '#ffffff' && h > 5 && w > 5) bars.push(x + w / 2); rect(x, y, w, h) }
  ctx.fillText = (t, x, y, ...rest) => { if (t.includes('|')) texts.push(x + ctx.measureText(t).width / 2); text(t, x, y, ...rest) }
  paintXlsxChart(prepared, ctx as any, 400, 300)
  expect(bars).toHaveLength(4)
  expect(texts).toHaveLength(4)
  texts.forEach((x, i) => expect(x).toBeCloseTo(bars[i], 5))
})
import { prepareChartAxes } from '../src/charts/prepare'

it('M7: identical category caches and pending ranges share the model category role', () => {
  const cached = '<c:cat><c:strLit><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>South</c:v></c:pt></c:strLit></c:cat>'
  const same = parse(chart('<c:showCatName val="1"/>', 2, 2).replace(/<c:ser>/g, '<c:ser>' + cached))
  expect(same.groups[0].series.map(s => s.categories)).toEqual([undefined, undefined])
  expect(same.preparedAxes!.dataLabels.map(l => l.text)).toEqual(['North', 'South', 'North', 'South'])
  const pending = '<c:cat><c:strRef><c:f>S!$A$1:$A$2</c:f></c:strRef></c:cat>'
  const shared = parse(chart('', 2, 2).replace(/<c:ser>/g, '<c:ser>' + pending))
  expect(shared.categoryPending).toBeDefined()
  expect(shared.groups[0].series[1].pending?.categories).toBeUndefined()
})

it('I1: one enabled point overrides an all-off group without reserving its disabled neighbours', () => {
  const model = parse(chart(controls(false) + '<c:dLbl><c:idx val="0"/><c:showVal val="1"/></c:dLbl>', 4096))
  expect(model.preparedAxes!.dataLabels.map(l => l.text)).toEqual(['1'])
  expect(labelPreparationBound(model)).toBeLessThan(10_000)
})
