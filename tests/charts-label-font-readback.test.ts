// @vitest-environment node
import { expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'

// Browsers normalise ctx.font on read-back (quoting, weight keywords), so the
// string read back differs from the one assigned; node-canvas returns it as set.
// Data labels must be painted from the font layout measured with, not ctx.font.
function browserLikeContext() {
  const ctx = createCanvas(400, 300).getContext('2d')
  let font = ctx.font
  Object.defineProperty(ctx, 'font', {
    get: () => `normalised(${font})`,
    set: (value: string) => { font = value },
  })
  return ctx
}

const flags = ['showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize', 'showLegendKey']
const labels = flags.map(flag => `<c:${flag} val="${flag === 'showVal' ? 1 : 0}"/>`).join('')
const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:dLbls>${labels}</c:dLbls><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>34</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`

it('paints data labels when the context normalises ctx.font on read-back', () => {
  const model = parseXlsxChart(parseXmlOrdered(xml), undefined)!
  expect(model.preparedAxes?.dataLabels.map(l => l.text)).toEqual(['12', '34'])
  const ctx = browserLikeContext()
  const drawn: string[] = []
  const fillText = ctx.fillText.bind(ctx)
  ctx.fillText = (text: string, x: number, y: number) => { drawn.push(text); fillText(text, x, y) }
  expect(() => paintXlsxChart(model, ctx as never, 400, 300)).not.toThrow()
  expect(drawn).toEqual(expect.arrayContaining(['12', '34']))
})
