// @vitest-environment node
/**
 * Task A4b (spec 5.9, 5.10): chart titles (content order, run formatting precedence), legends (position, overlay,
 * legendEntry delete) and multi-level categories (grouped tiers, noMultiLvlLbl). Each case asserts painted output
 * and geometry, not only the model. Synthetic ("Office-shaped synthetic") fixtures only.
 */
import { afterEach, describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { layoutChart, type ChartLayout } from '../src/charts/layout'
import { createChartWorkspace } from '../src/charts/workspace'
import { pendingLines } from '../src/charts/fill'
import { displayView } from '../src/charts/display'
import { categoryTiers, readCategoryLevels, CHART_CATEGORY_LEVEL_UNREADABLE } from '../src/charts/categories'
import { getChildren, type XmlNode } from '../src/core/xml'
import { titleRuns, CHART_TITLE_RUN_LIMIT, CHART_TITLE_LINE_LIMIT } from '../src/charts/titles'
import { parseXlsx } from '../src/xlsx/parse'
import type { ChartParseContext } from '../src/charts/context'
import type { ChartModel } from '../src/charts/model'
import type { ThemeContext } from '../src/drawing/style'
import { chartWorkbook, chartWorkbookBytes } from './helpers/xlsx-chart-fixture'
import { OfficePackage } from '../src/core/zip'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'

const opened: Array<{ release(): Promise<void> }> = []
afterEach(async () => { for (const source of opened.splice(0)) await source.release() })
/** The same workbook read resident (parseXlsx) and paged (the import session's drawing read): the chart models. */
async function residentAndPaged(bytes: Uint8Array): Promise<[ChartModel, ChartModel]> {
  const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
  const resident = doc.sheets[0].drawings?.[0].content
  // Native success: the coverage representation is the chart itself, not a card or fallback picture.
  expect(doc.drawingCoverage?.[0]).toMatchObject({ selectedRepresentation: 'chart' })
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'A4b', profile: 'D' })
  expect((await session.done).kind).toBe('complete')
  const source = createPagedWorkbookSource(session)
  opened.push(source)
  const store = source.store('1')
  const read = await store.readDrawings!({ stamp: store.stamp, signal: new AbortController().signal }, 'D')
  const paged = read.drawings[0].content
  if (resident?.kind !== 'xlsxChart' || paged?.kind !== 'xlsxChart') throw new Error('no chart')
  return [resident.chart, paged.chart]
}

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
    // Resident and paged reads see the same style part: the same resolved runs.
    const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/></Relationships>`
    const [resident, paged] = await residentAndPaged(await chartWorkbookBytes(space(bars(three) + axes(), `<title>${cached}</title>`), { patch: (zip: JSZip) => { zip.file('xl/charts/_rels/chart1.xml.rels', rels); zip.file('xl/charts/style1.xml', styleXml) } }))
    expect(resident.titleRuns).toEqual([{ text: 'Cached', sizePt: 20, bold: true, color: '#C00000' }])
    expect(paged.titleRuns).toEqual(resident.titleRuns)
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

describe('charts-text (Task A4b fix round 1)', () => {
  const level = (labels: Array<string | undefined>) => `<lvl>${labels.map((t, k) => t ? `<pt idx="${k}"><v>${t}</v></pt>` : '').join('')}</lvl>`
  /** A `<cat>` multi-level cache whose label reads are counted (each `v` text read through a getter). */
  function countedCat(inner: number, outer: Array<Array<string | undefined>>) {
    const cat = parseXmlOrdered(`<cat><multiLvlStrRef><multiLvlStrCache><ptCount val="${inner}"/>${level(Array.from({ length: inner }, (_, i) => `L${i}`))}${outer.map(level).join('')}</multiLvlStrCache></multiLvlStrRef></cat>`)
    const counter = { reads: 0 }
    const lvls = getChildren(getChildren(getChildren(cat, 'multiLvlStrRef')[0], 'multiLvlStrCache')[0], 'lvl')
    for (const lvl of lvls.slice(1)) for (const pt of getChildren(lvl, 'pt')) {
      const v = getChildren(pt, 'v')[0] as XmlNode, text = v['#text']
      Object.defineProperty(v, '#text', { get: () => { counter.reads++; return text } })
    }
    return { cat, counter }
  }

  it('I1: outer category levels are admitted before they are built, and levels past the bounds are never read', () => {
    const outer = Array.from({ length: 20 }, (_, l) => Array.from({ length: 100 }, (_, k) => `G${l}-${k}`))
    // A refusing charge: the first admission comes after sizing only the first outer level.
    const refused = countedCat(100, outer)
    let readsAtCharge = -1
    expect(() => readCategoryLevels(refused.cat, 100, () => { readsAtCharge = refused.counter.reads; throw new Error('refused') }, { truncated: false }, 255)).toThrow('refused')
    expect(readsAtCharge).toBeLessThanOrEqual(100)
    // Accepted: 8 levels kept; the ninth and later are never read (sizing + build of the 8 kept only).
    const kept = countedCat(100, outer), cut = { truncated: false }, charges: number[] = []
    const levels = readCategoryLevels(kept.cat, 100, b => { charges.push(b) }, cut, 255)!
    expect(levels).toHaveLength(8)
    expect(cut.truncated).toBe(true)
    expect(kept.counter.reads).toBeLessThanOrEqual(2 * 8 * 100)
    expect(charges).toHaveLength(8)
    // The label bound: a level that would pass 4,096 labels stops its sizing walk at the first label over.
    const big = Array.from({ length: 3000 }, (_, k) => `B${k}`)
    const wide = countedCat(3000, [big, big]), wideCut = { truncated: false }
    expect(readCategoryLevels(wide.cat, 3000, () => {}, wideCut, 255)).toHaveLength(1)
    expect(wideCut.truncated).toBe(true)
    expect(wide.counter.reads).toBeLessThanOrEqual(3000 * 2 + 4096 - 3000 + 1)
    // Out-of-order and repeated indices are sorted (first label of an index kept); a long label is cut and reported.
    const shuffled = parseXmlOrdered(`<cat><multiLvlStrRef><multiLvlStrCache><ptCount val="4"/>${level(['a', 'b', 'c', 'd'])}<lvl><pt idx="2"><v>Y2</v></pt><pt idx="0"><v>Y0</v></pt><pt idx="2"><v>dup</v></pt><pt idx="3"><v>${'x'.repeat(300)}</v></pt></lvl></multiLvlStrCache></multiLvlStrRef></cat>`)
    const shuffledCut = { truncated: false }
    const sorted = readCategoryLevels(shuffled, 4, () => {}, shuffledCut, 255)!
    expect([...sorted[0].starts]).toEqual([0, 2, 3])
    expect(sorted[0].labels.slice(0, 2)).toEqual(['Y0', 'Y2'])
    expect(sorted[0].labels[2]).toHaveLength(255)
    expect(shuffledCut.truncated).toBe(true)
  })

  it('I2/M1: every title record counts against the bound, the band stays within the chart, and cuts are reported', () => {
    const brs = '<title><tx><rich><bodyPr/><p><r><t>A</t></r>' + '<br/>'.repeat(1000) + '<r><t>B</t></r></p></rich></tx></title>'
    const records = titleRuns(parseXmlOrdered(brs), 'A B', undefined, undefined, () => {})!
    expect(records.runs.length).toBeLessThanOrEqual(CHART_TITLE_RUN_LIMIT)
    expect(records.runs.filter(r => r.text === '\n').length).toBeLessThan(CHART_TITLE_LINE_LIMIT)
    expect(records.truncated).toBe(true)
    const model = parse(space(bars(three) + axes(), brs))
    expect(model.truncated).toBe(true)
    expect(model.diagnostics?.some(d => d.code === 'chart-title-truncated')).toBe(true)
    const p = paint(model)
    expect(p.layout.title.h).toBeLessThanOrEqual(H / 2)
    expect(p.layout.area.h).toBeGreaterThan(H / 3)
    // The data is still drawn.
    expect(p.count('FF0000', p.layout.plot, 10)).toBeGreaterThan(100)
    // Paragraph separators count too.
    const paras = '<title><tx><rich><bodyPr/>' + Array.from({ length: 40 }, (_, i) => `<p><r><t>P${i}</t></r></p>`).join('') + '</rich></tx></title>'
    const many = parse(space(bars(three) + axes(), paras))
    expect(many.titleRuns!.filter(r => r.text === '\n').length).toBe(CHART_TITLE_LINE_LIMIT - 1)
    const drawn = paint(many)
    expect(drawn.layout.title.h).toBeLessThanOrEqual(H / 2)
    expect(drawn.layout.diagnostics.some(d => d.code === 'chart-title-truncated')).toBe(true)
    // 65 text runs: the 65th is dropped, the model says so and the shown title is what the runs keep.
    const runs65 = '<title><tx><rich><bodyPr/><p>' + Array.from({ length: 64 }, (_, i) => `<r><rPr b="${i % 2}"/><t>r${i} </t></r>`).join('') + '<r><t>LOST</t></r></p></rich></tx></title>'
    const cut = parse(space(bars(three) + axes(), runs65))
    expect(cut.titleRuns).toHaveLength(64)
    expect(cut.truncated).toBe(true)
    expect(cut.title).not.toContain('LOST')
    expect(paint(cut).texts.some(t => t.text.includes('LOST'))).toBe(false)
  })

  it('I3: legend entries are deleted by series c:idx, never by position', () => {
    const named = (idx: number, values: number[], color: string) => `<ser><idx val="${idx}"/><order val="${idx}"/><tx><v>S${idx}</v></tx>${fill(color)}<cat>${strLit(['a', 'b'])}</cat><val>${lit(values)}</val></ser>`
    const legend = (idx: number) => `<legend><legendPos val="r"/><legendEntry><idx val="${idx}"/><delete val="1"/></legendEntry></legend>`
    const chart = (idx: number) => paint(parse(space(bars(named(7, [1, 2], 'FF0000') + named(42, [2, 1], '00FF00')) + axes(), undefined, legend(idx))))
    expect(chart(42).texts.map(t => t.text).filter(t => /^S\d+$/.test(t))).toEqual(['S7'])
    expect(chart(7).texts.map(t => t.text).filter(t => /^S\d+$/.test(t))).toEqual(['S42'])
    expect(chart(0).texts.map(t => t.text).filter(t => /^S\d+$/.test(t))).toEqual(['S7', 'S42'])
    // A combination chart: the line series (c:idx 2, second in document order) is the deleted entry.
    const combo = paint(parse(space(`${bars(named(5, [1, 2], 'FF0000'))}<lineChart><grouping val="standard"/>${named(2, [2, 1], '0000FF')}${AX}</lineChart>` + axes(), undefined, legend(2))))
    expect(combo.texts.map(t => t.text).filter(t => /^S\d+$/.test(t))).toEqual(['S5'])
    expect(combo.count('0000FF', combo.layout.legend, 10)).toBe(0)
    expect(combo.count('0000FF', combo.layout.plot, 10)).toBeGreaterThan(20)
    // A large series index is a valid key (not bounded like a point offset).
    expect(paint(parse(space(bars(named(70000, [1, 2], 'FF0000') + named(1, [2, 1], '00FF00')) + axes(), undefined, legend(70000)))).texts.map(t => t.text).filter(t => /^S\d+$/.test(t))).toEqual(['S1'])
  })

  it('I4: authored typeface and underline resolve through the chain and are painted', async () => {
    // A run whose only formatting is a typeface and an underline keeps its runs.
    const only = parse(space(bars(three) + axes(), '<title><tx><rich><bodyPr/><p><r><rPr u="sng"><latin typeface="Courier New"/></rPr><t>Mono</t></r></p></rich></tx></title>'))
    expect(only.titleRuns).toEqual([{ text: 'Mono', font: 'Courier New', underline: true }])
    const p = paint(only), mono = p.text('Mono')[0]
    expect(mono.font).toBe(`${10 * 96 / 72 * 1.4}px "Courier New"`)
    // The underline: a #595959 band under the text, inside the title box, below the glyphs' baseline.
    const box = p.layout.title
    let rows = 0
    for (let y = Math.floor(box.y); y < box.y + box.h; y++) if (p.count('595959', { x: mono.x - 20, y, w: 40, h: 1 }, 30) >= 30) rows++
    expect(rows).toBeGreaterThanOrEqual(1)
    const plain = paint(parse(space(bars(three) + axes(), '<title><tx><rich><bodyPr/><p><r><rPr><latin typeface="Courier New"/></rPr><t>Mono</t></r></p></rich></tx></title>')))
    let plainRows = 0
    for (let y = Math.floor(box.y); y < box.y + box.h; y++) if (plain.count('595959', { x: mono.x - 20, y, w: 40, h: 1 }, 30) >= 30) plainRows++
    expect(plainRows).toBe(0)
    // Inheritance and override: the paragraph's typeface reaches a run that states none; a run's u="none" overrides txPr's.
    const inherit = parse(space(bars(three) + axes(), '<title><tx><rich><bodyPr/><p><pPr><defRPr><latin typeface="Georgia"/></defRPr></pPr><r><rPr/><t>One</t></r><r><rPr u="none"><latin typeface="Arial"/></rPr><t>Two</t></r></p></rich></tx><txPr><bodyPr/><p><pPr><defRPr u="sng"/></pPr></p></txPr></title>'))
    expect(inherit.titleRuns).toEqual([{ text: 'One', font: 'Georgia', underline: true }, { text: 'Two', font: 'Arial', underline: false }])
    const two = paint(inherit)
    expect(two.text('One')[0].font).toBe(`${10 * 96 / 72 * 1.4}px "Georgia"`)
    expect(two.text('Two')[0].font).toBe(`${10 * 96 / 72 * 1.4}px "Arial"`)
    // Theme references resolve to the theme's fonts (+mj-lt: major), and the style part's fontRef idx="major" too.
    const theme = { palette: {}, colors: {}, colorMap: {}, fonts: { minor: { latin: 'Minor Face' }, major: { latin: 'Major Face' } } } as unknown as ThemeContext
    const themed = parseXlsxChart(parseXmlOrdered(space(bars(three) + axes(), '<title><tx><rich><bodyPr/><p><r><rPr><latin typeface="+mj-lt"/></rPr><t>T</t></r></p></rich></tx></title>')), theme)!
    expect(themed.titleRuns).toEqual([{ text: 'T', font: 'Major Face' }])
    const styled = parseXlsxChart(parseXmlOrdered(space(bars(three) + axes(), '<title><tx><strRef><f>A1</f><strCache><ptCount val="1"/><pt idx="0"><v>S</v></pt></strCache></strRef></tx></title>')), theme, { styleParts: { style: { title: { font: 'major', underline: true } } } } as unknown as ChartParseContext)!
    expect(styled.titleRuns).toEqual([{ text: 'S', font: 'Major Face', underline: true }])
  })

  it('M2: mixed-size runs share one baseline', () => {
    const rich = '<title><tx><rich><bodyPr/><p><r><rPr sz="1200"><solidFill><srgbClr val="FF0000"/></solidFill></rPr><t>Run</t></r><r><rPr sz="2800"><solidFill><srgbClr val="00B050"/></solidFill></rPr><t>Para</t></r></p></rich></tx></title>'
    const p = paint(parse(space(bars(three) + axes(), rich)))
    const run = p.text('Run')[0], para = p.text('Para')[0]
    expect(run.y).toBeCloseTo(para.y, 5)
    // Ink: both words (no descenders) end on the same row; the small word's top is well below the large one's.
    const box = p.layout.title, rows = (hex: string) => { const out: number[] = []; for (let y = Math.floor(box.y); y < box.y + box.h; y++) if (p.count(hex, { x: box.x, y, w: box.w, h: 1 }, 40) > 0) out.push(y); return out }
    const red = rows('FF0000'), green = rows('00B050')
    expect(Math.abs(red[red.length - 1] - green[green.length - 1])).toBeLessThanOrEqual(2)
    expect(red[0] - green[0]).toBeGreaterThan(8)
  })

  it('I5: uncached multi-level references read their outer levels from the worksheet (resident and paged)', async () => {
    const cells = [
      { ref: 'A1', v: '2023' }, { ref: 'B1', v: 'Q1' }, { ref: 'B2', v: 'Q2' }, { ref: 'A3', v: '2024' }, { ref: 'B3', v: 'Q1' }, { ref: 'B4', v: 'Q2' },
      { ref: 'C1', v: 3 }, { ref: 'C2', v: 7 }, { ref: 'C3', v: 2 }, { ref: 'C4', v: 9 },
    ]
    const xml = space(bars(`<ser><idx val="0"/><order val="0"/><tx><v>S</v></tx><cat><multiLvlStrRef><f>Styled!$A$1:$B$4</f></multiLvlStrRef></cat><val><numRef><f>Styled!$C$1:$C$4</f></numRef></val></ser>`) + axes())
    const [resident, paged] = await residentAndPaged(await chartWorkbookBytes(xml, { cells }))
    for (const model of [resident, paged]) {
      expect(model.categories).toEqual(['Q1', 'Q2', 'Q1', 'Q2'])
      expect(model.categoryLevels?.map(l => ({ starts: [...l.starts], labels: l.labels }))).toEqual([{ starts: [0, 2], labels: ['2023', '2024'] }])
      expect(model.categoryMultiLevel).toBeUndefined()
    }
    // Drawn as the cached hierarchy is.
    const p = paint(resident)
    expect(p.text('2023')[0].x).toBeCloseTo(p.layout.plot.x + p.layout.plot.w * .25, 0)
    expect(p.text('2024')[0].x).toBeCloseTo(p.layout.plot.x + p.layout.plot.w * .75, 0)
    // An outer cell without a saved value: the level is not drawn and a diagnostic says so (never silently flat).
    const unreadable = await chartWorkbookBytes(xml, { cells: cells.filter(c => c.ref !== 'A3'), patch: async (zip: JSZip) => {
      const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      zip.file('xl/worksheets/sheet1.xml', sheet.replace(/(<row r="3"[^>]*>)/, '$1<c r="A3"><f>B9&amp;"x"</f></c>'))
    } })
    const [r2, p2] = await residentAndPaged(unreadable)
    for (const model of [r2, p2]) {
      expect(model.categoryLevels).toBeUndefined()
      expect(model.diagnostics?.some(d => d.code === CHART_CATEGORY_LEVEL_UNREADABLE)).toBe(true)
      expect(model.categories).toEqual(['Q1', 'Q2', 'Q1', 'Q2'])
    }
  })
})

describe('charts-text (Task A4b fix round 2)', () => {
  it('I1 residual: unordered levels build only bounded valid entries, never per-source-point temporaries', () => {
    // Two valid labels in descending order among 10,000 invalid indices (past the 2 categories).
    const pts = `<pt idx="1"><v>Second</v></pt>${Array.from({ length: 10000 }, (_, k) => `<pt idx="${5 + k}"><v>x</v></pt>`).join('')}<pt idx="0"><v>First</v></pt>`
    const cat = parseXmlOrdered(`<cat><multiLvlStrRef><multiLvlStrCache><ptCount val="2"/><lvl><pt idx="0"><v>a</v></pt><pt idx="1"><v>b</v></pt></lvl><lvl>${pts}</lvl></multiLvlStrCache></multiLvlStrRef></cat>`)
    // Record the largest array any map/filter/slice/from call produces while the level is read.
    const proto = Array.prototype as unknown as Record<string, (...a: unknown[]) => unknown>
    const names = ['map', 'filter', 'slice', 'concat'] as const, saved = names.map(n => proto[n])
    const from = Array.from
    let largest = 0
    names.forEach((n, i) => { proto[n] = function (this: unknown[], ...a: unknown[]) { const r = saved[i].apply(this, a) as unknown[]; largest = Math.max(largest, r.length); return r } })
    Array.from = ((...a: Parameters<typeof Array.from>) => { const r = from(...a); largest = Math.max(largest, r.length); return r }) as typeof Array.from
    let charged = 0, levels
    try { levels = readCategoryLevels(cat, 2, b => { charged += b }, { truncated: false }, 255) } finally { names.forEach((n, i) => { proto[n] = saved[i] }); Array.from = from }
    expect(largest).toBeLessThan(100)
    expect(levels!.map(l => ({ starts: [...l.starts], labels: l.labels }))).toEqual([{ starts: [0, 1], labels: ['First', 'Second'] }])
    // The admission covers what is built: the retained level and the two-entry sort workspace only.
    expect(charged).toBeLessThan(1024)
  })

  it('I2 residual: a first line that does not fit half the chart is omitted; the plot never goes negative', () => {
    const big = '<title><tx><rich><bodyPr/><p><r><rPr sz="7200"/><t>Huge</t></r></p><p><r><t>Second</t></r></p></rich></tx></title>'
    const small = paint(parse(space(bars(three) + axes(), big)), 400, 100)
    expect(small.text('Huge')).toHaveLength(0)
    expect(small.text('Second')).toHaveLength(0)
    expect(small.layout.diagnostics.some(d => d.code === 'chart-title-truncated')).toBe(true)
    expect(small.layout.area.h).toBeGreaterThan(0)
    expect(small.layout.plot.h).toBeGreaterThan(0)
    // The same title fits in a tall chart (both lines drawn).
    const tall = paint(parse(space(bars(three) + axes(), big)), 400, 400)
    expect(tall.text('Huge')).toHaveLength(1)
    expect(tall.text('Second')).toHaveLength(1)
    expect(tall.layout.title.h).toBeLessThanOrEqual(400 / 2)
    // A plain title in a very short chart is omitted the same way.
    const plain = paint(parse(space(bars(three) + axes(), '<title><tx><rich><p><r><t>Plain</t></r></p></rich></tx></title>')), 400, 44)
    expect(plain.text('Plain')).toHaveLength(0)
    expect(plain.layout.area.h).toBeGreaterThanOrEqual(0)
  })

  it('I3 residual: a single varyColors series lists and deletes legend entries by point index', () => {
    const vary = (family: string, legend: string) => {
      const series = `<ser><idx val="7"/><order val="0"/><tx><v>S7</v></tx><cat>${strLit(['A', 'B', 'C'])}</cat><val>${lit([1, 2, 3])}</val></ser>`
      const group = family === 'bar' ? `<barChart><barDir val="col"/><grouping val="clustered"/><varyColors val="1"/>${series}${AX}</barChart>` : `<lineChart><grouping val="standard"/><varyColors val="1"/>${series}${AX}</lineChart>`
      return paint(parse(space(group + axes(), undefined, legend)))
    }
    const names = (p: ReturnType<typeof paint>) => p.texts.map(t => t.text).filter(t => /^(A|B|C|S7)$/.test(t))
    const legend = (idx?: number) => `<legend><legendPos val="r"/>${idx === undefined ? '' : `<legendEntry><idx val="${idx}"/><delete val="1"/></legendEntry>`}</legend>`
    for (const family of ['bar', 'line']) {
      const all = vary(family, legend())
      // Category axis labels A, B, C plus the legend's A, B, C: no series entry.
      expect(names(all).filter(t => t === 'S7'), family).toEqual([])
      expect(names(all).filter(t => t === 'B'), family).toHaveLength(2)
      const del = vary(family, legend(1))
      expect(names(del).filter(t => t === 'B'), family).toHaveLength(1)
      expect(names(del).filter(t => t === 'A'), family).toHaveLength(2)
      // The series c:idx is not a point: nothing is deleted.
      expect(names(vary(family, legend(7))).filter(t => t === 'C'), family).toHaveLength(2)
      // Each entry's swatch is its point's colour: the legend shows the second accent (point 1).
      expect(all.count('ED7D31', all.layout.legend, 12), family).toBeGreaterThan(20)
    }
  })

  it('R1: paged acquisition reads only the outer levels the fill keeps (resident and paged agree)', async () => {
    // A vertical reference of 20 columns (19 outer levels), each outer level with one label at row 1.
    const letters = 'ABCDEFGHIJKLMNOPQRST'.split('')
    const cells = [
      ...letters.slice(0, 19).map((c, k) => ({ ref: `${c}1`, v: `L${19 - k}` })),
      ...Array.from({ length: 40 }, (_, r) => ({ ref: `T${r + 1}`, v: `I${r}` })),
      ...Array.from({ length: 40 }, (_, r) => ({ ref: `U${r + 1}`, v: r + 1 })),
    ]
    const xml = space(bars(`<ser><idx val="0"/><order val="0"/><tx><v>S</v></tx><cat><multiLvlStrRef><f>Styled!$A$1:$T$40</f></multiLvlStrRef></cat><val><numRef><f>Styled!$U$1:$U$40</f></numRef></val></ser>`) + axes())
    const parsed = parse(xml)
    expect(parsed.categoryMultiLevel).toBe(true)
    const lines = pendingLines([{ content: { kind: 'xlsxChart', chart: parsed } as never }], 'Styled')
    // The value line, the innermost category line and eight outer lines: never the 11 levels the fill cannot keep.
    expect(lines).toHaveLength(1 + 1 + 8)
    expect(lines.reduce((n, l) => n + l.count, 0)).toBe(10 * 40)
    const [resident, paged] = await residentAndPaged(await chartWorkbookBytes(xml, { cells }))
    for (const model of [resident, paged]) {
      expect(model.categoryLevels).toHaveLength(8)
      expect(model.categoryLevels!.map(l => l.labels)).toEqual(Array.from({ length: 8 }, (_, k) => [`L${k + 1}`]))
      expect(model.diagnostics?.some(d => d.code === 'chart-category-levels-truncated')).toBe(true)
      expect(model.truncated).toBe(true)
    }
  })
})

describe('charts-text (Task A4b fix round 3)', () => {
  it('axis bands that exhaust a small placement leave an empty, never negative, plot (automatic and fallback)', () => {
    const big = '<title><tx><rich><bodyPr/><p><r><rPr sz="7200"/><t>Huge</t></r></p></rich></tx></title>'
    const manual = '<layout><manualLayout><x val="0.9"/><y val="0.9"/><w val="0.05"/><h val="0.05"/></manualLayout></layout>'
    const cases: Array<[string, number, number]> = [
      [space(bars(three) + axes(), big), 400, 20],
      [space(bars(three) + axes(), big), 20, 400],
      [space(bars(three) + axes(), undefined, '<legend><legendPos val="b"/></legend>'), 60, 24],
      // A manual outer plot that leaves no room falls back to the (also exhausted) automatic plot.
      [space(manual + bars(three) + axes(), big), 400, 20],
    ]
    for (const [xml, w, h] of cases) {
      const model = parse(xml)
      const p = paint(model, w, h)
      expect(p.layout.plot.h, `${w}x${h}`).toBeGreaterThanOrEqual(0)
      expect(p.layout.plot.w, `${w}x${h}`).toBeGreaterThanOrEqual(0)
      // Nothing is plotted in an empty plot: no bar colour appears.
      if (p.layout.plot.h === 0 || p.layout.plot.w === 0) expect(p.count('FF0000', { x: 0, y: 0, w, h }, 10), `${w}x${h}`).toBe(0)
    }
    const fallback = paint(parse(cases[3][0]), 400, 20)
    expect(fallback.layout.diagnostics.some(d => d.code === 'chart-layout-invalid')).toBe(true)
  })
})

describe('charts-text (A4b merged with feature: trendlines, contour, forecast)', () => {
  const trend = (body = '<trendlineType val="linear"/>') => `<trendline>${body}</trendline>`
  const legendNames = (p: ReturnType<typeof paint>) => p.texts.filter(t => t.x >= p.layout.legend.x - 1 && t.y >= p.layout.legend.y - 1 && t.y <= p.layout.legend.y + p.layout.legend.h + 1).map(t => t.text)
  it('legend: point entries and legendEntry deletion combine with trendline entries and contour bands', () => {
    // A single varyColors series with a trendline: its points, then the trendline entry; point 1 deleted.
    const vary = (legend: string) => paint(parse(space(`<barChart><barDir val="col"/><grouping val="clustered"/><varyColors val="1"/><ser><idx val="7"/><order val="0"/><tx><v>S7</v></tx>${trend()}<cat>${strLit(['A', 'B', 'C'])}</cat><val>${lit([1, 2, 3])}</val></ser>${AX}</barChart>` + axes(), undefined, legend)))
    expect(legendNames(vary('<legend><legendPos val="r"/></legend>'))).toEqual(['A', 'B', 'C', 'Linear (S7)'])
    expect(legendNames(vary('<legend><legendPos val="r"/><legendEntry><idx val="1"/><delete val="1"/></legendEntry></legend>'))).toEqual(['A', 'C', 'Linear (S7)'])
    // Series keyed by c:idx, the trendline by A6a's numbering (after the two series: entry 2).
    const two = (entries: number[]) => paint(parse(space(bars(`<ser><idx val="0"/><order val="0"/><tx><v>Series 0</v></tx>${fill('FF0000')}${trend()}<cat>${strLit(['a', 'b', 'c'])}</cat><val>${lit([1, 3, 2])}</val></ser>` + ser(1, [2, 1, 3])) + axes(), undefined, `<legend><legendPos val="r"/>${entries.map(i => `<legendEntry><idx val="${i}"/><delete val="1"/></legendEntry>`).join('')}</legend>`)))
    expect(legendNames(two([]))).toEqual(['Series 0', 'Series 1', 'Linear (Series 0)'])
    expect(legendNames(two([1]))).toEqual(['Series 0', 'Linear (Series 0)'])
    expect(legendNames(two([2]))).toEqual(['Series 0', 'Series 1'])
    // A contour's band legend: band entries deleted by band index.
    const contour = (legend: string) => {
      const cats = `<cat>${strLit(['C0', 'C1', 'C2'])}</cat>`
      const row = (i: number, v: number[]) => `<ser><idx val="${i}"/><order val="${i}"/><tx><v>R${i}</v></tx>${cats}<val>${lit(v)}</val></ser>`
      return paint(parse(`<chartSpace><chart><autoTitleDeleted val="1"/><plotArea><surfaceChart>${row(0, [1, 3, 5])}${row(1, [2, 4, 6])}<axId val="1"/><axId val="2"/><axId val="3"/></surfaceChart><catAx><axId val="1"/><crossAx val="3"/></catAx><valAx><axId val="3"/><majorUnit val="2"/><crossAx val="1"/></valAx><serAx><axId val="2"/><crossAx val="3"/></serAx></plotArea>${legend}</chart></chartSpace>`))
    }
    const bands = legendNames(contour('<legend><legendPos val="r"/></legend>'))
    expect(bands.length).toBeGreaterThan(1)
    expect(legendNames(contour('<legend><legendPos val="r"/><legendEntry><idx val="0"/><delete val="1"/></legendEntry></legend>'))).toEqual(bands.slice(1))
  })

  it('category tiers combine with forecast slots, and titles with contour axes', () => {
    const level = (labels: Array<string | undefined>) => `<lvl>${labels.map((t, k) => t ? `<pt idx="${k}"><v>${t}</v></pt>` : '').join('')}</lvl>`
    const cat = `<cat><multiLvlStrRef><f>Sheet1!$A$1:$B$4</f><multiLvlStrCache><ptCount val="4"/>${level(['Q1', 'Q2', 'Q1', 'Q2'])}${level(['2023', undefined, '2024', undefined])}</multiLvlStrCache></multiLvlStrRef></cat>`
    // A line with a 2-period forward forecast: the axis has 6 slots, the tiers span the 4 real categories only.
    const p = paint(parse(space(`<lineChart><grouping val="standard"/><ser><idx val="0"/><order val="0"/><tx><v>S</v></tx>${trend('<trendlineType val="linear"/><forward val="2"/>')}${cat}<val>${lit([1, 2, 3, 4])}</val></ser>${AX}</lineChart>` + axes())))
    const plot = p.layout.plot, slots = 6
    expect(p.text('2023')[0].x).toBeCloseTo(plot.x + plot.w * (1 / slots), 0)
    expect(p.text('2024')[0].x).toBeCloseTo(plot.x + plot.w * (3 / slots), 0)
    expect(p.text('2023')[0].y).toBeGreaterThan(p.text('Q1')[0].y + 8)
    // A titled contour keeps its series-axis labels and band legend below a run-formatted title.
    const contour = paint(parse(`<chartSpace><chart><title><tx><rich><bodyPr/><p><r><rPr b="1"><solidFill><srgbClr val="C00000"/></solidFill></rPr><t>Surface</t></r><r><t> map</t></r></p></rich></tx></title><plotArea><surfaceChart><ser><idx val="0"/><order val="0"/><tx><v>R0</v></tx><cat>${strLit(['C0', 'C1'])}</cat><val>${lit([1, 3])}</val></ser><ser><idx val="1"/><order val="1"/><tx><v>R1</v></tx><cat>${strLit(['C0', 'C1'])}</cat><val>${lit([2, 4])}</val></ser><axId val="1"/><axId val="2"/><axId val="3"/></surfaceChart><catAx><axId val="1"/><crossAx val="3"/></catAx><valAx><axId val="3"/><crossAx val="1"/></valAx><serAx><axId val="2"/><crossAx val="3"/></serAx></plotArea></chart></chartSpace>`))
    expect(contour.text('Surface')[0]).toMatchObject({ fill: '#C00000' })
    expect(contour.text('R0')).toHaveLength(1)
    expect(contour.text('R1')).toHaveLength(1)
    expect(contour.text('R0')[0].y).toBeGreaterThan(contour.layout.title.y + contour.layout.title.h)
  })
})
