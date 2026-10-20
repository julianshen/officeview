// @vitest-environment node
/**
 * Task A4b (spec 5.9, 5.10): chart titles (content order, run formatting precedence), legends (position, overlay,
 * legendEntry delete) and multi-level categories (grouped tiers, noMultiLvlLbl). Each case asserts painted output
 * and geometry, not only the model. Synthetic ("Office-shaped synthetic") fixtures only.
 */
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { layoutChart, type ChartLayout } from '../src/charts/layout'
import { createChartWorkspace } from '../src/charts/workspace'
import { displayView } from '../src/charts/display'
import { categoryTiers } from '../src/charts/categories'
import { parseXlsx } from '../src/xlsx/parse'
import type { ChartParseContext } from '../src/charts/context'
import type { ChartModel } from '../src/charts/model'
import { chartWorkbook } from './helpers/xlsx-chart-fixture'

const W = 400, H = 260
const lit = (values: number[]) => `<numLit><ptCount val="${values.length}"/>${values.map((v, i) => `<pt idx="${i}"><v>${v}</v></pt>`).join('')}</numLit>`
const strLit = (labels: string[]) => `<strLit><ptCount val="${labels.length}"/>${labels.map((v, i) => `<pt idx="${i}"><v>${v}</v></pt>`).join('')}</strLit>`
const COLORS = ['FF0000', '00FF00', '0000FF']
const fill = (hex: string) => `<spPr><solidFill><srgbClr val="${hex}"/></solidFill></spPr>`
const ser = (i: number, values: number[], cat = `<cat>${strLit(values.map((_, k) => `C${k}`))}</cat>`) =>
  `<ser><idx val="${i}"/><order val="${i}"/><tx><v>Series ${i}</v></tx>${fill(COLORS[i % 3])}${cat}<val>${lit(values)}</val></ser>`
const AX = '<axId val="1"/><axId val="2"/>'
const axes = (cat = '') => `<catAx><axId val="1"/>${cat}<crossAx val="2"/></catAx><valAx><axId val="2"/><crossAx val="1"/></valAx>`
const bars = (series: string, dir = 'col') => `<barChart><barDir val="${dir}"/><grouping val="clustered"/>${series}${AX}</barChart>`
const space = (plot: string, chart = '<autoTitleDeleted val="1"/>', legend = '', root = '') => `<chartSpace>${root}<chart>${chart}<plotArea>${plot}</plotArea>${legend}</chart></chartSpace>`
const V = [[3, 7, 2, 9], [5, 2, 4, 1], [6, 4, 8, 3]]
const three = V.map((v, i) => ser(i, v)).join('')
const parse = (xml: string, context?: ChartParseContext) => parseXlsxChart(parseXmlOrdered(xml), undefined, context)!

interface Text { text: string; x: number; y: number; font: string; fill: string; maxWidth?: number }
interface Segment { x1: number; y1: number; x2: number; y2: number }
/** Paints on a real canvas, recording the text drawn (with the font and colour in force) and stroked segments. */
function paint(model: ChartModel, w = W, h = H) {
  const canvas = createCanvas(w, h), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  const texts: Text[] = [], segments: Segment[] = []
  let path: Segment[] = [], pen: [number, number] = [0, 0]
  const o = { fillText: ctx.fillText.bind(ctx), moveTo: ctx.moveTo.bind(ctx), lineTo: ctx.lineTo.bind(ctx), beginPath: ctx.beginPath.bind(ctx), stroke: ctx.stroke.bind(ctx) }
  ctx.fillText = (text: string, x: number, y: number, maxWidth?: number) => {
    const m = ctx.getTransform()
    texts.push({ text, x: m.a * x + m.c * y + m.e, y: m.b * x + m.d * y + m.f, font: ctx.font, fill: String(ctx.fillStyle).toUpperCase(), ...(maxWidth !== undefined ? { maxWidth } : {}) })
    if (maxWidth === undefined) o.fillText(text, x, y); else o.fillText(text, x, y, maxWidth)
  }
  ctx.beginPath = () => { path = []; o.beginPath() }
  ctx.moveTo = (x, y) => { pen = [x, y]; o.moveTo(x, y) }
  ctx.lineTo = (x, y) => { path.push({ x1: pen[0], y1: pen[1], x2: x, y2: y }); pen = [x, y]; o.lineTo(x, y) }
  ctx.stroke = (...a: []) => { segments.push(...path); o.stroke(...a) }
  paintXlsxChart(model, ctx, w, h)
  const layout: ChartLayout = layoutChart(model, w, h, 1, 0, { width: (text, css) => { ctx.font = css; return ctx.measureText(text).width }, lineHeight: () => model.fontSizePt * 96 / 72 })
  const data = ctx.getImageData(0, 0, w, h).data
  const near = (hex: string, x: number, y: number, tolerance = 40) => {
    const i = (Math.round(y) * w + Math.round(x)) * 4
    return [0, 2, 4].every((k, c) => Math.abs(data[i + c] - parseInt(hex.slice(k, k + 2), 16)) <= tolerance)
  }
  /** Pixels of a colour inside a box. */
  const count = (hex: string, box: { x: number; y: number; w: number; h: number }, tolerance = 40) => {
    let n = 0
    for (let y = Math.max(0, Math.floor(box.y)); y < Math.min(h, Math.ceil(box.y + box.h)); y++)
      for (let x = Math.max(0, Math.floor(box.x)); x < Math.min(w, Math.ceil(box.x + box.w)); x++) if (near(hex, x, y, tolerance)) n++
    return n
  }
  const text = (t: string) => texts.filter(e => e.text === t)
  return { texts, text, segments, layout, data, near, count }
}

describe('charts-text (Task A4b)', () => {
  it('title content order: rich, strRef cache, automatic, autoTitleDeleted', () => {
    const titled = (title: string, series = three) => paint(parse(space(bars(series) + axes(), title)))
    // Rich text wins over a reference in the same tx (malformed, but the order is the rule).
    const rich = titled('<title><tx><rich><bodyPr/><p><r><t>Rich title</t></r></p></rich><strRef><f>Sheet1!$A$1</f><strCache><ptCount val="1"/><pt idx="0"><v>Cached</v></pt></strCache></strRef></tx></title>')
    expect(rich.text('Rich title')).toHaveLength(1)
    expect(rich.text('Cached')).toHaveLength(0)
    const titleText = rich.text('Rich title')[0]
    expect(titleText.y).toBeCloseTo(rich.layout.title.y, 5)
    expect(rich.count('595959', rich.layout.title, 60)).toBeGreaterThan(10)
    // The reference's cache, then the automatic title (single series: its name; several: "Chart Title").
    expect(titled('<title><tx><strRef><f>Sheet1!$A$1</f><strCache><ptCount val="1"/><pt idx="0"><v>Cached</v></pt></strCache></strRef></tx></title>').text('Cached')).toHaveLength(1)
    expect(titled('<title/><autoTitleDeleted val="0"/>', ser(0, V[0])).text('Series 0')).toHaveLength(1)
    expect(titled('<title/><autoTitleDeleted val="0"/>').text('Chart Title')).toHaveLength(1)
    // autoTitleDeleted removes the automatic title (and its band), never an authored one.
    const deleted = titled('<title/><autoTitleDeleted val="1"/>')
    expect(deleted.texts.some(t => t.text === 'Chart Title')).toBe(false)
    const none = titled('<autoTitleDeleted val="1"/>')
    expect(deleted.layout.plot).toEqual(none.layout.plot)
    expect(titled('<title><tx><rich><p><r><t>Kept</t></r></p></rich></tx></title><autoTitleDeleted val="1"/>').text('Kept')).toHaveLength(1)
    // Paragraphs are lines: each drawn on its own line, the title band grows and the plot moves down.
    const two = titled('<title><tx><rich><bodyPr/><p><r><t>Line one</t></r></p><p><r><t>Line two</t></r></p></rich></tx></title>')
    const [l1, l2] = [two.text('Line one')[0], two.text('Line two')[0]]
    expect(l1).toBeDefined(); expect(l2).toBeDefined()
    expect(l2.y - l1.y).toBeGreaterThan(15)
    expect(two.layout.title.h).toBeGreaterThan(rich.layout.title.h + 15)
    expect(two.layout.plot.y).toBeGreaterThan(rich.layout.plot.y + 15)
    // Both lines are painted inside the grown title box.
    const lower = { ...two.layout.title, y: l2.y, h: two.layout.title.y + two.layout.title.h - l2.y }
    expect(two.count('595959', lower, 60)).toBeGreaterThan(10)
  })

  it('run formatting precedence rPr, paragraph, txPr, style, theme', async () => {
    // rPr (run 1: 12 pt bold red) over the paragraph default (24 pt green), which run 2 inherits; the title txPr
    // supplies what neither states (italic, and a blue colour both override).
    const rich = '<title><tx><rich><bodyPr/><p><pPr><defRPr sz="2400"><solidFill><srgbClr val="00B050"/></solidFill></defRPr></pPr>'
      + '<r><rPr lang="en-US" sz="1200" b="1"><solidFill><srgbClr val="FF0000"/></solidFill></rPr><t>Run</t></r><r><rPr lang="en-US"/><t>Para</t></r></p></rich></tx>'
      + '<txPr><bodyPr/><p><pPr><defRPr i="1"><solidFill><srgbClr val="0000FF"/></solidFill></defRPr></pPr></p></txPr></title>'
    const p = paint(parse(space(bars(three) + axes(), rich)))
    const run = p.text('Run')[0], para = p.text('Para')[0]
    expect(run).toMatchObject({ font: 'italic bold 16px "Calibri"', fill: '#FF0000' })
    expect(para).toMatchObject({ font: 'italic 32px "Calibri"', fill: '#00B050' })
    // Drawn left to right on one line, centred as a whole in the title box.
    expect(para.x).toBeGreaterThan(run.x)
    expect(run.y).toBeCloseTo(para.y, 5)
    const box = p.layout.title
    expect(p.count('FF0000', { x: box.x, y: box.y, w: box.w, h: box.h })).toBeGreaterThan(10)
    expect(p.count('00B050', { x: box.x, y: box.y, w: box.w, h: box.h })).toBeGreaterThan(30)
    // The band fits the largest run (32 px) instead of the 14 pt default.
    expect(box.h).toBeGreaterThanOrEqual(32 + 6)

    // txPr alone (a referenced title): 18 pt bold blue.
    const cached = '<tx><strRef><f>Sheet1!$A$1</f><strCache><ptCount val="1"/><pt idx="0"><v>Cached</v></pt></strCache></strRef></tx>'
    const tx = paint(parse(space(bars(three) + axes(), `<title>${cached}<txPr><bodyPr/><p><pPr><defRPr sz="1800" b="1"><solidFill><srgbClr val="0000FF"/></solidFill></defRPr></pPr></p></txPr></title>`)))
    expect(tx.text('Cached')[0]).toMatchObject({ font: 'bold 24px "Calibri"', fill: '#0000FF' })
    expect(tx.count('0000FF', tx.layout.title)).toBeGreaterThan(10)

    // Theme (defaults): the theme minor font, 1.4 x the chart font, #595959.
    const theme = paint(parse(space(bars(three) + axes(), `<title>${cached}</title>`, '', '<txPr><bodyPr/><p><pPr><defRPr sz="1200"/></pPr></p></txPr>')))
    expect(theme.text('Cached')[0]).toMatchObject({ fill: '#595959' })
    expect(theme.text('Cached')[0].font).toBe(`${16 * 1.4}px "Calibri"`)

    // The chart style part (cs:title) sits between txPr and the theme: its defRPr size/bold and fontRef colour
    // apply when the chart states nothing; a txPr size overrides the style's size but not its colour.
    const CS = 'http://schemas.microsoft.com/office/drawing/2012/chartStyle', A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    const styleXml = `<cs:chartStyle xmlns:cs="${CS}" xmlns:a="${A}" id="201"><cs:title><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:srgbClr val="C00000"/></cs:fontRef><cs:defRPr sz="2000" b="1"/></cs:title></cs:chartStyle>`
    const packaged = async (title: string) => {
      const pkg = await chartWorkbook(space(bars(three) + axes(), title), { patch: (zip: JSZip) => {
        zip.file('xl/charts/_rels/chart1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/></Relationships>`)
        zip.file('xl/charts/style1.xml', styleXml)
      } })
      const content = (await parseXlsx(pkg)).sheets[0].drawings?.[0].content
      if (content?.kind !== 'xlsxChart') throw new Error('no chart')
      return paint(content.chart)
    }
    const styled = await packaged(`<title>${cached}</title>`)
    expect(styled.text('Cached')[0]).toMatchObject({ font: `bold ${20 * 96 / 72}px "Calibri"`, fill: '#C00000' })
    expect(styled.count('C00000', styled.layout.title)).toBeGreaterThan(10)
    const overridden = await packaged(`<title>${cached}<txPr><bodyPr/><p><pPr><defRPr sz="1000" b="0"/></pPr></p></txPr></title>`)
    expect(overridden.text('Cached')[0]).toMatchObject({ font: `${10 * 96 / 72}px "Calibri"`, fill: '#C00000' })
  })

  it('legend position, overlay and legendEntry delete', () => {
    const at = (legend: string) => paint(parse(space(bars(three) + axes(), undefined, legend)))
    const none = at('')
    const pos = (p: string) => at(`<legend><legendPos val="${p}"/></legend>`)
    const swatch = (r: ReturnType<typeof paint>, name: string) => {
      const label = r.text(name)[0]
      expect(label).toBeDefined()
      return label
    }
    const r = pos('r'), l = pos('l'), t = pos('t'), b = pos('b'), tr = pos('tr')
    // Right / left: a column beside the plot, vertically centred (tr: at the top); top / bottom: a row.
    expect(r.layout.legend.x).toBeGreaterThanOrEqual(r.layout.plot.x + r.layout.plot.w)
    expect(l.layout.legend.x + l.layout.legend.w).toBeLessThanOrEqual(l.layout.plot.x)
    expect(t.layout.legend.y + t.layout.legend.h).toBeLessThanOrEqual(t.layout.plot.y)
    expect(b.layout.legend.y).toBeGreaterThanOrEqual(b.layout.plot.y + b.layout.plot.h)
    expect(tr.layout.legend.x).toBeGreaterThanOrEqual(tr.layout.plot.x + tr.layout.plot.w)
    expect(tr.layout.legend.y).toBeLessThan(r.layout.legend.y)
    for (const p of [r, l, tr]) { const [a, c] = [swatch(p, 'Series 0'), swatch(p, 'Series 2')]; expect(c.y).toBeGreaterThan(a.y); expect(c.x).toBeCloseTo(a.x, 5) }
    for (const p of [t, b]) { const [a, c] = [swatch(p, 'Series 0'), swatch(p, 'Series 2')]; expect(c.x).toBeGreaterThan(a.x); expect(c.y).toBeCloseTo(a.y, 5) }
    // Each legend's swatch for series 1 is painted green inside its box.
    for (const p of [r, l, t, b, tr]) expect(p.count('00FF00', p.layout.legend, 10)).toBeGreaterThan(20)

    // overlay=1: the legend takes no space (the plot is the legend-less plot) and is painted over the plot.
    const over = at('<legend><legendPos val="r"/><overlay val="1"/></legend>')
    expect(over.layout.plot).toEqual(none.layout.plot)
    expect(at('<legend><legendPos val="r"/><overlay val="0"/></legend>').layout.plot).toEqual(r.layout.plot)
    const lg = over.layout.legend
    expect(lg.x).toBeLessThan(over.layout.plot.x + over.layout.plot.w)
    const name = over.text('Series 1')[0]
    // The series-1 swatch (left of its name) is green although it lies over the plot.
    expect(over.count('00FF00', { x: lg.x, y: name.y - 6, w: name.x - lg.x, h: 12 }, 10)).toBeGreaterThan(20)

    // legendEntry/delete: the entry leaves the legend (name and swatch); the series is still drawn.
    const del = at('<legend><legendPos val="r"/><legendEntry><idx val="1"/><delete val="1"/></legendEntry></legend>')
    expect(del.text('Series 1')).toHaveLength(0)
    expect(del.text('Series 0')).toHaveLength(1)
    expect(del.text('Series 2')).toHaveLength(1)
    expect(del.count('00FF00', del.layout.legend, 10)).toBe(0)
    expect(del.count('00FF00', del.layout.plot, 10)).toBeGreaterThan(100)
    expect(del.layout.legend.h).toBeLessThan(r.layout.legend.h)
    // delete=0 keeps it.
    expect(at('<legend><legendPos val="r"/><legendEntry><idx val="1"/><delete val="0"/></legendEntry></legend>').text('Series 1')).toHaveLength(1)
    // A varyColors pie legend lists points: idx is the point index.
    const pie = paint(parse(space(`<pieChart><varyColors val="1"/>${ser(0, V[0])}</pieChart>`, undefined, '<legend><legendPos val="r"/><legendEntry><idx val="2"/><delete val="1"/></legendEntry></legend>')))
    expect(pie.texts.map(e => e.text).filter(e => /^C\d$/.test(e))).toEqual(['C0', 'C1', 'C3'])
    // Every entry deleted: no legend band at all.
    const all = at(`<legend><legendPos val="r"/>${[0, 1, 2].map(i => `<legendEntry><idx val="${i}"/><delete val="1"/></legendEntry>`).join('')}</legend>`)
    expect(all.layout.plot).toEqual(none.layout.plot)
  })

  it('multi-level categories draw grouped tiers; noMultiLvlLbl shows the innermost only', () => {
    const inner = ['Q1', 'Q2', 'Q1', 'Q2']
    const level = (labels: Array<string | undefined>) => `<lvl>${labels.map((t, k) => t ? `<pt idx="${k}"><v>${t}</v></pt>` : '').join('')}</lvl>`
    const multi = (outer: Array<Array<string | undefined>>) => `<cat><multiLvlStrRef><f>Sheet1!$A$2:$C$5</f><multiLvlStrCache><ptCount val="4"/>${level(inner)}${outer.map(level).join('')}</multiLvlStrCache></multiLvlStrRef></cat>`
    const years = ['2023', undefined, '2024', undefined]
    const chart = (cat: string, catAx = '', dir = 'col') => space(bars(ser(0, V[0], cat) + ser(1, V[1], cat), dir) + (dir === 'bar' ? `<catAx><axId val="1"/>${catAx}<axPos val="l"/><crossAx val="2"/></catAx><valAx><axId val="2"/><axPos val="b"/><crossAx val="1"/></valAx>` : axes(catAx)))
    const two = parse(chart(multi([years])))
    expect(two.categories).toEqual(inner)
    const p = paint(two), plot = p.layout.plot
    // Group labels centred over their children, in a tier below the innermost labels.
    const [y23, y24] = [p.text('2023')[0], p.text('2024')[0]]
    expect(y23).toBeDefined(); expect(y24).toBeDefined()
    expect(y23.x).toBeCloseTo(plot.x + plot.w * .25, 0)
    expect(y24.x).toBeCloseTo(plot.x + plot.w * .75, 0)
    const q = p.text('Q1')[0]
    expect(y23.y).toBeGreaterThan(q.y + 8)
    expect(p.count('595959', { x: y23.x - 15, y: y23.y, w: 30, h: 14 }, 60)).toBeGreaterThan(5)
    // A separator at the group boundary (between categories 1 and 2) runs from the axis down through the tier.
    const mid = plot.x + plot.w / 2
    const separator = p.segments.find(s => Math.abs(s.x1 - mid) <= 1 && Math.abs(s.x2 - mid) <= 1 && Math.min(s.y1, s.y2) <= plot.y + plot.h + 1 && Math.max(s.y1, s.y2) >= y23.y + 10)
    expect(separator).toBeDefined()
    // The tier takes room from the plot.
    const single = paint(parse(chart(`<cat>${strLit(inner)}</cat>`)))
    expect(plot.h).toBeLessThan(single.layout.plot.h - 10)

    // Spans follow the display view: a category hidden by plotVisOnly leaves its group (slots, not source indices).
    const hidden = displayView({ ...two, categoryHidden: Uint32Array.of(1) })
    expect(categoryTiers(two.categoryLevels!, 4, hidden)[0]).toEqual([{ start: 0, end: 1, text: '2023' }, { start: 1, end: 3, text: '2024' }])

    // Three levels: the outermost spans every child; tiers stack outward.
    const three = paint(parse(chart(multi([years, ['Region', undefined, undefined, undefined]]))))
    const region = three.text('Region')[0]
    expect(region.x).toBeCloseTo(three.layout.plot.x + three.layout.plot.w / 2, 0)
    expect(region.y).toBeGreaterThan(three.text('2023')[0].y + 8)
    expect(three.layout.plot.h).toBeLessThan(plot.h - 10)

    // A vertical category axis (bar): tiers stack leftward of the innermost labels, centred on their spans.
    const bar = paint(parse(chart(multi([years]), '', 'bar')))
    const bq = bar.text('Q1')[0], b23 = bar.text('2023')[0]
    expect(b23.x).toBeLessThan(bq.x - 10)
    expect(b23.y).toBeCloseTo(bar.layout.plot.y + bar.layout.plot.h * .75, 0)

    // noMultiLvlLbl=1: only the innermost level, pixel-identical to a single-level axis.
    const flat = paint(parse(chart(multi([years]), '<noMultiLvlLbl val="1"/>')))
    expect(flat.text('2023')).toHaveLength(0)
    expect(flat.layout.plot).toEqual(single.layout.plot)
    expect(Buffer.from(flat.data).equals(Buffer.from(single.data))).toBe(true)

    // Retained level data is charged to the model before it is built; a refusal is a refused chart, not a partial one.
    const ws = (headroom?: number) => {
      const workspace = createChartWorkspace('mobile')
      if (headroom !== undefined) workspace.reserve(workspace.limit - headroom, 'filler')
      return { workspace } as ChartParseContext
    }
    const plain = ws(), levels = ws()
    parse(chart(`<cat>${strLit(inner)}</cat>`), plain)
    const charged = parse(chart(multi([years])), levels)
    expect(charged.refused).toBeUndefined()
    expect(levels.workspace.used - plain.workspace.used).toBeGreaterThanOrEqual(2 * 2 * 32)
    // The smallest headroom that admits each parse: the multi-level chart needs at least its level bytes more.
    const minimum = (xml: string) => { let lo = 0, hi = 1 << 24; while (lo < hi) { const mid = (lo + hi) >>> 1; if (parse(xml, ws(mid)).refused) lo = mid + 1; else hi = mid } return lo }
    const singleMin = minimum(chart(`<cat>${strLit(inner)}</cat>`)), multiple = minimum(chart(multi([years])))
    expect(multiple - singleMin).toBeGreaterThanOrEqual(64 + 2 * (4 + 32 + 8))
    expect(parse(chart(multi([years])), ws(multiple - 1)).refused).toBe(true)
  })
})
