// @vitest-environment node
/**
 * Task A4a (spec 5.4, 5.8, 5.11): markers, lines, fills, colour precedence and series options. Rewritten in fix
 * round 1 to assert painted geometry and colours (each case fails without its feature). Synthetic fixtures only.
 */
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { layoutChart } from '../src/charts/layout'
import { MARKER_AUTO_CYCLE, clampMarkerSize } from '../src/charts/markers'
import { resolveSeriesColor } from '../src/charts/style'
import { parseXlsx } from '../src/xlsx/parse'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { THEME_DEFAULTS } from '../src/styles/tokens'
import type { ChartModel } from '../src/charts/model'
import { chartWorkbook } from './helpers/xlsx-chart-fixture'

const parse = (xml: string) => parseXlsxChart(parseXmlOrdered(xml), undefined)!
const lit = (values: number[]) => `<numLit><ptCount val="${values.length}"/>${values.map((v, i) => `<pt idx="${i}"><v>${v}</v></pt>`).join('')}</numLit>`
const cats = (n: number) => `<cat><strLit><ptCount val="${n}"/>${Array.from({ length: n }, (_, i) => `<pt idx="${i}"><v>C${i}</v></pt>`).join('')}</strLit></cat>`
const ser = (i: number, values: number[], extra = '') => `<ser><idx val="${i}"/><order val="${i}"/><tx><v>S${i}</v></tx>${extra}${cats(values.length)}<val>${lit(values)}</val></ser>`
const AX = '<axId val="1"/><axId val="2"/>'
const AXES = '<catAx><axId val="1"/><crossAx val="2"/></catAx><valAx><axId val="2"/><crossAx val="1"/></valAx>'
const space = (plot: string, opts: { root?: string; chart?: string; legend?: string } = {}) => `<chartSpace>${opts.root ?? ''}<chart>${opts.chart ?? '<autoTitleDeleted val="1"/>'}<plotArea>${plot}</plotArea>${opts.legend ?? ''}</chart></chartSpace>`
const W = 400, H = 260
const rgba = (data: Uint8ClampedArray, x: number, y: number) => { const i = (Math.round(y) * W + Math.round(x)) * 4; return [data[i], data[i + 1], data[i + 2], data[i + 3]] }
const hexAt = (data: Uint8ClampedArray, x: number, y: number) => rgba(data, x, y).slice(0, 3).map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase()
/** The (antialiased) marker border colour 00FF00 dominates the pixel. */
const green = (data: Uint8ClampedArray, x: number, y: number) => { const [r, g, b] = rgba(data, x, y); return g > 200 && r < 140 && b < 140 }
const measure = (ctx: CanvasRenderingContext2D, model: ChartModel) => ({ width: (text: string, css: string) => { ctx.font = css; return ctx.measureText(text).width }, lineHeight: () => model.fontSizePt * 96 / 72 })
/** Paint, recording fills (rect and path) with their style kind, and line segments with the stroke style in force. */
function paint(model: ChartModel) {
  const canvas = createCanvas(W, H), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  const kind = (s: unknown) => typeof s === 'string' ? s.toUpperCase() : (s as object).constructor.name
  const rects: Array<{ x: number; y: number; w: number; h: number; style: string }> = []
  const segments: Array<{ x1: number; y1: number; x2: number; y2: number }> = []
  const strokes: Array<{ style: string; width: number; cap: string; dash: number[]; segments: typeof segments }> = []
  let path: typeof segments = [], pen: [number, number] = [0, 0]
  const o = { fillRect: ctx.fillRect.bind(ctx), moveTo: ctx.moveTo.bind(ctx), lineTo: ctx.lineTo.bind(ctx), beginPath: ctx.beginPath.bind(ctx), stroke: ctx.stroke.bind(ctx) }
  ctx.fillRect = (x, y, w, h) => { rects.push({ x, y, w, h, style: kind(ctx.fillStyle) }); o.fillRect(x, y, w, h) }
  ctx.beginPath = () => { path = []; o.beginPath() }
  ctx.moveTo = (x, y) => { pen = [x, y]; o.moveTo(x, y) }
  ctx.lineTo = (x, y) => { path.push({ x1: pen[0], y1: pen[1], x2: x, y2: y }); pen = [x, y]; o.lineTo(x, y) }
  ctx.stroke = (...a: []) => { strokes.push({ style: kind(ctx.strokeStyle), width: ctx.lineWidth, cap: ctx.lineCap, dash: ctx.getLineDash(), segments: path }); o.stroke(...a) }
  paintXlsxChart(model, ctx, W, H)
  const layout = layoutChart(model, W, H, 1, 0, measure(ctx, model))
  return { ctx, data: ctx.getImageData(0, 0, W, H).data, rects, strokes, layout }
}

describe('charts-style (Task A4a)', () => {
  it('marker symbols, auto cycle, size clamp 2-72, marker spPr', () => {
    expect(clampMarkerSize(1)).toBe(2)
    expect(clampMarkerSize(75)).toBe(72)
    expect(clampMarkerSize(10)).toBe(10)
    expect(MARKER_AUTO_CYCLE.slice(0, 3)).toEqual(['diamond', 'square', 'triangle'])
    // One point at a known spot: a marker of 12 pt (16 px) filled FF00AA, bordered 00FF00 (1 pt), no series line.
    const marker = (symbol: string, spPr = '<spPr><solidFill><srgbClr val="FF00AA"/></solidFill><ln w="12700"><solidFill><srgbClr val="00FF00"/></solidFill></ln></spPr>') =>
      parse(space(`<lineChart><grouping val="standard"/>${ser(0, [5, 5, 5], `<spPr><ln><noFill/></ln></spPr><marker><symbol val="${symbol}"/><size val="75"/>${spPr}</marker>`)}<marker val="1"/>${AX}</lineChart><catAx><axId val="1"/><crossAx val="2"/><delete val="1"/></catAx><valAx><axId val="2"/><crossAx val="1"/><delete val="1"/><min val="0"/><max val="10"/></valAx>`))
    const m = marker('diamond')
    expect(m.groups[0].series[0].markerStyle).toMatchObject({ symbol: 'diamond', size: 72, fill: { kind: 'solid', color: 'FF00AA' }, line: { color: '00FF00' } })
    const at = (model: ChartModel) => { const p = paint(model); const { plot } = p.layout; return { p, cx: plot.x + plot.w / 2, cy: plot.y + plot.h / 2 } }
    // Size 72 pt (clamped from 75) = 96 px: the centre point's marker reaches ~48 px out.
    const sq = at(marker('square'))
    expect(hexAt(sq.p.data, sq.cx, sq.cy)).toBe('FF00AA')
    expect(hexAt(sq.p.data, sq.cx + 40, sq.cy + 40)).toBe('FF00AA')
    expect(green(sq.p.data, sq.cx + 48, sq.cy)).toBe(true)
    // A diamond leaves the square's corners empty; a circle reaches its axis but not its corners.
    const di = at(marker('diamond'))
    expect(hexAt(di.p.data, di.cx, di.cy)).toBe('FF00AA')
    expect(hexAt(di.p.data, di.cx + 40, di.cy + 40)).not.toBe('FF00AA')
    expect(hexAt(di.p.data, di.cx + 40, di.cy)).toBe('FF00AA')
    const ci = at(marker('circle'))
    expect(hexAt(ci.p.data, ci.cx + 44, ci.cy)).toBe('FF00AA')
    expect(hexAt(ci.p.data, ci.cx + 40, ci.cy + 40)).not.toBe('FF00AA')
    // A marker noFill leaves the centre unpainted, its border still drawn.
    const hollow = at(marker('square', '<spPr><noFill/><ln w="12700"><solidFill><srgbClr val="00FF00"/></solidFill></ln></spPr>'))
    expect(hexAt(hollow.p.data, hollow.cx, hollow.cy)).toBe('FFFFFF')
    expect(green(hollow.p.data, hollow.cx + 48, hollow.cy)).toBe(true)
  })

  it('line dash, width, cap and noFill', () => {
    const model = parse(space(`<lineChart><grouping val="standard"/>
      ${ser(0, [10, 20], '<spPr><ln w="25400" cap="rnd"><noFill/><prstDash val="dash"/></ln></spPr><marker><symbol val="none"/></marker>')}
      ${ser(1, [15, 25], '<spPr><ln w="19050" cap="sq"><solidFill><srgbClr val="0000FF"/></solidFill><prstDash val="dot"/></ln></spPr><marker><symbol val="none"/></marker>')}
      <marker val="1"/>${AX}</lineChart>${AXES}`))
    expect(model.groups[0].series[0].line).toBe(false)
    const { strokes } = paint(model)
    const series = strokes.filter(s => s.segments.length && s.style !== '#D9D9D9' && s.style !== '#BFBFBF')
    // Only the second series strokes: 1.5 pt = 2 px, square caps, dot = [1, 3] line widths.
    expect(series).toEqual([expect.objectContaining({ style: '#0000FF', width: 2, cap: 'square', dash: [2, 6] })])
  })

  it('solid, gradient, pattern and noFill for series, points, plot area, chart area, legend and title', () => {
    const fills: Record<string, [string, string | undefined]> = {
      solid: ['<solidFill><srgbClr val="ABCDEF"/></solidFill>', '#ABCDEF'],
      gradient: ['<gradFill><gsLst><gs pos="0"><srgbClr val="FFFFFF"/></gs><gs pos="100000"><srgbClr val="000000"/></gs></gsLst><lin ang="5400000"/></gradFill>', 'CanvasGradient'],
      pattern: ['<pattFill prst="pct50"><fgClr><srgbClr val="000000"/></fgClr><bgClr><srgbClr val="FFFFFF"/></bgClr></pattFill>', 'CanvasPattern'],
      none: ['<noFill/>', undefined],
    }
    for (const [kind, [xml, expected]] of Object.entries(fills)) {
      const sp = `<spPr>${xml}</spPr>`
      const chart = (where: string) => parse(space(`<barChart><barDir val="col"/><grouping val="clustered"/>${ser(0, [3, 5],
        where === 'series' ? sp : where === 'point' ? `<spPr><solidFill><srgbClr val="CCCCCC"/></solidFill></spPr><dPt><idx val="1"/>${sp}</dPt>` : '<spPr><solidFill><srgbClr val="CCCCCC"/></solidFill></spPr>')}${AX}</barChart>${AXES}${where === 'plot' ? sp : ''}`, {
        root: where === 'chartArea' ? sp : '',
        chart: `<title><tx><rich><p><r><t>Title</t></r></p></rich></tx>${where === 'title' ? sp : ''}</title><autoTitleDeleted val="0"/>`,
        legend: `<legend><legendPos val="r"/>${where === 'legend' ? sp : ''}</legend>`,
      }))
      const at = (model: ChartModel, box: (p: ReturnType<typeof paint>) => { x: number; y: number; w: number; h: number }) => {
        const p = paint(model), b = box(p)
        return p.rects.filter(r => Math.abs(r.x - b.x) < .01 && Math.abs(r.y - b.y) < .01 && Math.abs(r.w - b.w) < .01 && Math.abs(r.h - b.h) < .01).map(r => r.style)
      }
      expect(at(chart('chartArea'), () => ({ x: 0, y: 0, w: W, h: H })), `chart area ${kind}`).toEqual(expected ? [expected] : [])
      expect(at(chart('plot'), p => p.layout.plot), `plot area ${kind}`).toEqual(expected ? [expected] : [])
      expect(at(chart('legend'), p => p.layout.legend), `legend ${kind}`).toEqual(expected ? [expected] : [])
      expect(at(chart('title'), p => p.layout.title), `title ${kind}`).toEqual(expected ? [expected] : [])
      const bars = (model: ChartModel) => { const p = paint(model), { plot } = p.layout; return p.rects.filter(r => r.x > plot.x && r.x + r.w < plot.x + plot.w && r.h > 2 && r.w > 2).map(r => r.style) }
      expect(bars(chart('series')), `series ${kind}`).toEqual(expected ? [expected, expected] : [])
      expect(bars(chart('point')), `point ${kind}`).toEqual(expected ? ['#CCCCCC', expected] : ['#CCCCCC'])
    }
  })

  it('chart area noFill is transparent; default via c:style', () => {
    const at = (root: string, partPath?: string) => {
      const xml = space(`<barChart><barDir val="col"/>${ser(0, [5])}${AX}</barChart>${AXES}`, { root })
      const model = partPath ? parseXlsxChart(parseXmlOrdered(xml), undefined, { pkg: {} as never, partPath, profile: 'desktop', workspace: undefined as never })! : parse(xml)
      return rgba(paint(model).data, 3, 3)
    }
    expect(at('<spPr><noFill/></spPr>')[3]).toBe(0)
    expect(at('')).toEqual([255, 255, 255, 255])
    expect(at('<style val="45"/>', 'ppt/charts/chart2.xml')).toEqual([0, 0, 0, 255])
    expect(at('<style val="45"/>', 'xl/charts/chart2.xml')).toEqual([255, 255, 255, 255])
  })

  it('colour precedence spPr, chartStyle, colorStyle, c:style, then accents with lumMod variants after six', () => {
    const parts = {
      color: { method: 'cycle', colors: [{ kind: 'srgb' as const, value: '778899', transforms: [] }], variations: [[], [{ name: 'lumMod', value: '50000' }]] },
      style: { dataPoint: { color: { kind: 'srgb' as const, value: '445566', transforms: [] } } },
    }
    expect(resolveSeriesColor({ index: 0, parts })).toBe('445566')
    expect(resolveSeriesColor({ index: 1, parts: { color: parts.color } })).toBe('3A444E')
    expect(resolveSeriesColor({ index: 0, styleId: 3, count: 1, parts: { color: parts.color } })).toBe('778899')
    expect(resolveSeriesColor({ index: 0, styleId: 1, count: 1 })).toBe('000000')
    expect(resolveSeriesColor({ index: 5 })).toBe('70AD47')
    expect(resolveSeriesColor({ index: 6 })).toBe('264478')
    // An explicit spPr colour is the series colour, whatever the parts say (parse).
    expect(parse(space(`<barChart>${ser(0, [1], '<spPr><solidFill><srgbClr val="112233"/></solidFill></spPr>')}${AX}</barChart>${AXES}`, { root: '<style val="3"/>' })).groups[0].series[0].color).toBe('112233')
  })

  it('invertIfNegative fills negative bars with the inverted fill', () => {
    const model = parse(space(`<barChart><barDir val="col"/>${ser(0, [10, -5], '<spPr><solidFill><srgbClr val="0000FF"/></solidFill></spPr><invertIfNegative val="1"/>')}${AX}</barChart>${AXES}`))
    const { rects, layout } = paint(model)
    const bars = rects.filter(r => r.x > layout.plot.x && r.w > 2 && r.h > 2).map(r => r.style)
    // Office paints an inverted solid negative white (no c14 invertSolidFillFmt here).
    expect(bars).toEqual(['#0000FF', '#FFFFFF'])
  })

  it('series and point explosion geometry', () => {
    const model = parse(space(`<pieChart><varyColors val="1"/>${ser(0, [50, 50], '<explosion val="20"/><dPt><idx val="0"/><explosion val="35"/></dPt>')}</pieChart>`))
    expect(model.groups[0].series[0].explosion).toBe(20)
    const canvas = createCanvas(300, 300), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    const moves: Array<[number, number]> = [], arcs: number[] = []
    const moveTo = ctx.moveTo.bind(ctx), arc = ctx.arc.bind(ctx)
    ctx.moveTo = (x, y) => { moves.push([x, y]); moveTo(x, y) }
    ctx.arc = (x, y, r, a, b, c) => { arcs.push(r); arc(x, y, r, a, b, c) }
    paintXlsxChart(model, ctx, 300, 300)
    const layout = layoutChart(model, 300, 300, 1, 0, measure(ctx, model)), area = layout.area
    const cx = area.x + area.w / 2, cy = area.y + area.h / 2, r = arcs[0]
    // The radius leaves room for the farthest slice (35 %); each slice moves out along its bisector by its share of r.
    expect(r).toBeCloseTo((Math.min(area.w, area.h) / 2 - 4) / 1.35, 6)
    expect(Math.hypot(moves[0][0] - cx, moves[0][1] - cy)).toBeCloseTo(0.35 * r, 6)
    expect(Math.hypot(moves[1][0] - cx, moves[1][1] - cy)).toBeCloseTo(0.20 * r, 6)
    // First slice (top, clockwise, half the pie): its bisector points right.
    expect(moves[0][0]).toBeGreaterThan(cx)
    expect(moves[0][1]).toBeCloseTo(cy, 6)
  })

  it('dPt overrides take precedence over series formatting', () => {
    const model = parse(space(`<barChart><barDir val="col"/>${ser(0, [10, 20, 15], '<spPr><solidFill><srgbClr val="00FF00"/></solidFill></spPr><dPt><idx val="0"/><spPr><solidFill><srgbClr val="FF0000"/></solidFill></spPr></dPt>')}${ser(1, [5, 6, 7], '<spPr><solidFill><srgbClr val="0000FF"/></solidFill></spPr><dPt><idx val="2"/><spPr><pattFill prst="pct50"><fgClr><srgbClr val="000000"/></fgClr><bgClr><srgbClr val="FFFFFF"/></bgClr></pattFill></spPr></dPt>')}${AX}</barChart>${AXES}`))
    const { rects, layout } = paint(model)
    // Point overrides win in multi-series groups too (series 0 point 0 red; series 1 point 2 patterned).
    expect(rects.filter(r => r.x > layout.plot.x && r.w > 2 && r.h > 2).map(r => r.style)).toEqual(['#FF0000', '#00FF00', '#00FF00', '#0000FF', '#0000FF', 'CanvasPattern'])
  })

  it('gapWidth, overlap and varyColors for existing families', () => {
    const geometry = (extra: string, n = 1) => {
      const model = parse(space(`<barChart><barDir val="col"/><grouping val="clustered"/>${extra}${[0, 1].slice(0, n).map(i => ser(i, [10, 20, 30])).join('')}${AX}</barChart>${AXES}`))
      const { rects, layout } = paint(model)
      return { bars: rects.filter(r => r.x > layout.plot.x && r.w > 2 && r.h > 2), band: layout.plot.w / 3 }
    }
    // gapWidth 200 %: a bar is band / (1 + 2) wide; 0 %: the full band.
    const gap = geometry('<gapWidth val="200"/>')
    expect(gap.bars[0].w).toBeCloseTo(gap.band / 3, 1)
    expect(geometry('<gapWidth val="0"/>').bars[0].w).toBeCloseTo(gap.band, 1)
    // overlap 50 % with two series: the second bar starts half a bar after the first.
    const over = geometry('<gapWidth val="100"/><overlap val="50"/>', 2)
    const [a, b] = [over.bars[0], over.bars[3]]
    expect(b.x - a.x).toBeCloseTo(a.w / 2, 1)
    // varyColors: a lone single series takes one accent per point.
    expect(geometry('<varyColors val="1"/>').bars.map(r => r.style)).toEqual(['#4472C4', '#ED7D31', '#A5A5A5'])
  })

  it('dropLines on line and area groups and serLines on stacked bars draw the expected segments', () => {
    // Line: one vertical drop per category, from the point down to the category axis.
    const line = parse(space(`<lineChart><grouping val="standard"/>${ser(0, [10, 20], '<marker><symbol val="none"/></marker>')}<dropLines><spPr><ln w="12700"><solidFill><srgbClr val="888888"/></solidFill></ln></spPr></dropLines>${AX}</lineChart>${AXES}`))
    const lp = paint(line), drops = lp.strokes.filter(s => s.style === '#888888').flatMap(s => s.segments)
    expect(drops).toHaveLength(2)
    for (const d of drops) { expect(d.x1).toBeCloseTo(d.x2, 6); expect(Math.max(d.y1, d.y2)).toBeCloseTo(lp.layout.plot.y + lp.layout.plot.h, 6) }
    // Stacked columns: series lines join each series' segment top across the gap between adjacent bars.
    const stacked = (lines: string) => parse(space(`<barChart><barDir val="col"/><grouping val="stacked"/><overlap val="100"/>${ser(0, [10, 20])}${ser(1, [15, 25])}${lines}${AX}</barChart>${AXES}`))
    const sp = paint(stacked('<serLines><spPr><ln w="12700"><solidFill><srgbClr val="555555"/></solidFill></ln></spPr></serLines>'))
    const bars = sp.rects.filter(r => r.x > sp.layout.plot.x && r.w > 2 && r.h > 2)
    // Bars in series order: s0 c0, s0 c1, s1 c0, s1 c1.
    const top = (r: typeof bars[number]) => r.y
    const segs = sp.strokes.filter(s => s.style === '#555555').flatMap(s => s.segments)
    expect(segs).toHaveLength(2)
    expect(segs[0]).toMatchObject({ x1: bars[0].x + bars[0].w, x2: bars[1].x })
    expect(segs[0].y1).toBeCloseTo(top(bars[0]), 6); expect(segs[0].y2).toBeCloseTo(top(bars[1]), 6)
    expect(segs[1].y1).toBeCloseTo(top(bars[2]), 6); expect(segs[1].y2).toBeCloseTo(top(bars[3]), 6)
    // Without serLines no such segment exists.
    const none = paint(stacked(''))
    expect(none.strokes.flatMap(s => s.segments).filter(g => Math.abs(g.x1 - segs[0].x1) < .01 && Math.abs(g.x2 - segs[0].x2) < .01)).toEqual([])
  })

  it('dark mode leaves chart pixels identical', async () => {
    // A transparent (noFill) chart area over the worksheet: the viewer's dark palette reaches the sheet painter,
    // and still no chart (or paper) pixel may change.
    const xml = `<chartSpace><spPr><noFill/></spPr><chart><autoTitleDeleted val="1"/><plotArea><barChart><barDir val="col"/>${ser(0, [3, 5], '<spPr><solidFill><srgbClr val="4472C4"/></solidFill></spPr>')}${AX}</barChart>${AXES}</plotArea></chart></chartSpace>`
    const sheet = (await parseXlsx(await chartWorkbook(xml))).sheets[0]
    expect(sheet.drawings?.[0].content?.kind).toBe('xlsxChart')
    const geometry = buildSheetGeometry(sheet)
    const render = (colors: Record<string, string>) => {
      const canvas = createCanvas(800, 400), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
      renderSheetViewport(sheet, ctx, geometry, { box: { x: 0, y: 0, width: 800, height: 400 }, zoom: 1, deviceScale: 1 } as never, undefined, { colors } as never)
      return Buffer.from(ctx.getImageData(0, 0, 800, 400).data)
    }
    const light = render(THEME_DEFAULTS.light), dark = render(THEME_DEFAULTS.dark)
    // The chart did paint (its series colour is present), and the two renders are byte-identical.
    let series = 0
    for (let i = 0; i < light.length; i += 4) if (light[i] === 0x44 && light[i + 1] === 0x72 && light[i + 2] === 0xc4) series++
    expect(series).toBeGreaterThan(100)
    expect(light.equals(dark)).toBe(true)
  })
})
