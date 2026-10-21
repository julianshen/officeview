// @vitest-environment node
import { expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { labelBytes, labelPreparationBound } from '../src/charts/labels'
import { createChartWorkspace } from '../src/charts/workspace'
import { releaseChartModel } from '../src/charts/own'

const off = ['showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize', 'showLegendKey'].map(key => `<c:${key} val="0"/>`).join('')
const rangeSwitch = (on: boolean) => `<c:extLst><c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}"><c15:showDataLabelsRange val="${on ? 1 : 0}"/></c:ext></c:extLst>`
const xml = (on?: boolean, extra = '') => `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:tx><c:v>Sales</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>South</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt></c:numLit></c:val><c:dLbls>${off}<c:dLblPos val="inEnd"/><c:numFmt formatCode="0.00" sourceLinked="0"/>${on === undefined ? '' : rangeSwitch(on)}${extra}</c:dLbls><c:extLst><c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}"><c15:datalabelsRange><c15:f>S!$B$1:$B$2</c15:f><c15:dlblRangeCache><c:pt idx="0"><c:v>Authored cell text</c:v></c:pt></c15:dlblRangeCache></c15:datalabelsRange></c:ext></c:extLst></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`
const diagnostic = (model: ReturnType<typeof parseXlsxChart>) => model!.preparedAxes!.diagnostics.filter(d => d.code === 'chart-label-feature-unsupported')

it('Office Value From Cells with all ordinary flags off paints formatted fallback, admits it, and diagnoses once', () => {
  const workspace = createChartWorkspace('mobile')
  const model = parseXlsxChart(parseXmlOrdered(xml(true)), undefined, { profile: 'mobile', workspace } as never)!
  expect(model.refused).not.toBe(true)
  expect(model.preparedAxes!.dataLabels.map(l => [l.text, l.position])).toEqual([['3.00', 'inEnd'], ['5.00', 'inEnd']])
  expect(diagnostic(model)).toHaveLength(1)
  expect(labelPreparationBound(model)).toBeGreaterThanOrEqual(labelBytes(model.preparedAxes!.dataLabels))
  const ctx = createCanvas(400, 300).getContext('2d'), painted: string[] = []
  const fill = ctx.fillText.bind(ctx)
  ctx.fillText = (text, x, y, ...rest) => { painted.push(text); fill(text, x, y, ...rest) }
  paintXlsxChart(model, ctx as any, 400, 300)
  expect(painted).toContain('3.00')
  expect(painted).toContain('5.00')
  expect(painted).not.toContain('Authored cell text')
  releaseChartModel(model)
  expect(workspace.used).toBe(0)
})

for (const on of [undefined, false]) it(`ordinary flags off and range switch ${String(on)} stays off without diagnostic`, () => {
  const model = parseXlsxChart(parseXmlOrdered(xml(on)), undefined)!
  expect(model.preparedAxes!.dataLabels).toEqual([])
  expect(diagnostic(model)).toHaveLength(0)
  expect(labelPreparationBound(model)).toBe(0)
})

it('range fallback includes other enabled content once, without duplicating showVal', () => {
  const input = xml(true, '<c:separator> | </c:separator>').replace('<c:showCatName val="0"/>', '<c:showCatName val="1"/>')
  for (const data of [input, input.replace('<c:showVal val="0"/>', '<c:showVal val="1"/>')]) {
    const model = parseXlsxChart(parseXmlOrdered(data), undefined)!
    expect(model.preparedAxes!.dataLabels.map(l => l.text)).toEqual(['North | 3.00', 'South | 5.00'])
    expect(diagnostic(model)).toHaveLength(1)
  }
})

it('a point can explicitly disable an inherited range source without a stale diagnostic', () => {
  const point = '<c:dLbl><c:idx val="0"/><c:showCatName val="1"/>' + rangeSwitch(false) + '</c:dLbl><c:dLbl><c:idx val="1"/><c:delete val="1"/></c:dLbl>'
  const model = parseXlsxChart(parseXmlOrdered(xml(true, point)), undefined)!
  expect(model.preparedAxes!.dataLabels.map(l => l.text)).toEqual(['North'])
  expect(diagnostic(model)).toHaveLength(0)
})

it('an enabled range source is diagnosed even when no saved point value can form fallback text', () => {
  const input = xml(true).replace(/<c:numLit>.*?<\/c:numLit>/, '<c:numLit><c:ptCount val="2"/></c:numLit>')
  const model = parseXlsxChart(parseXmlOrdered(input), undefined)!
  expect(model.preparedAxes!.dataLabels).toEqual([])
  expect(diagnostic(model)).toHaveLength(1)
})
