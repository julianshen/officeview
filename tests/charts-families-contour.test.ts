// @vitest-environment node
/**
 * Task A5d: contour charts (`surfaceChart`, the 2D surface) and the surface3D card (spec 4.1 contour row, 9 work
 * bounds). Fixtures are Office-shaped synthetic, never Office evidence.
 *
 * Geometry convention: the z grid is series (rows) x categories (columns); row 0 is the bottom row (the series
 * axis runs upward from the category axis), column 0 the left column. Bands are the intervals between the
 * resolved value axis's ticks (its min and major unit, shared with stage 1).
 */
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXmlOrdered } from '../src/core/xml'
import { parseThemeContext } from '../src/drawing/style'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { layoutChart, paintWorkspaceBytes } from '../src/charts/layout'
import { displayView } from '../src/charts/display'
import { prepareChartAxes } from '../src/charts/prepare'
import { documentChartContext } from '../src/charts/context'
import { parseDocx } from '../src/docx/parse'
import type { ChartModel } from '../src/charts/model'
import { buildChartDocx } from './helpers/chart-fixtures'
import { SURFACE3D_FAMILY, describeFamilyConformance } from './helpers/chart-family-conformance'

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const W = 200, H = 160
/** The engine's default accent cycle (no theme): accent 1..3. */
const ACCENT = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47']

const cats = (n: number) => `<cat><strLit><ptCount val="${n}"/>${Array.from({ length: n }, (_, i) => `<pt idx="${i}"><v>C${i}</v></pt>`).join('')}</strLit></cat>`
/** One series; `undefined` is a missing cell; `n` categories (default: the value count). */
const ser = (i: number, values: Array<number | undefined>, n = values.length) =>
  `<ser><idx val="${i}"/><order val="${i}"/><tx><v>R${i}</v></tx>${cats(n)}<val><numLit><ptCount val="${values.length}"/>${values.map((v, j) => v === undefined ? '' : `<pt idx="${j}"><v>${v}</v></pt>`).join('')}</numLit></val></ser>`
interface Options { wireframe?: boolean; bandFmts?: string; valAx?: string; legend?: boolean; deleteBase?: boolean }
/** A contour chart, Office's axis wiring (catAx/serAx/valAx), titles off. */
function contourXml(series: string, o: Options = {}): string {
  const del = o.deleteBase ? '<delete val="1"/>' : ''
  return `<chartSpace><chart><autoTitleDeleted val="1"/><plotArea><surfaceChart>${o.wireframe ? '<wireframe val="1"/>' : ''}${series}${o.bandFmts ?? ''}<axId val="1"/><axId val="2"/><axId val="3"/></surfaceChart>`
    + `<catAx><axId val="1"/>${del}<crossAx val="3"/></catAx><valAx><axId val="3"/>${o.valAx ?? '<majorUnit val="2"/>'}<crossAx val="1"/></valAx><serAx><axId val="2"/>${del}<crossAx val="3"/></serAx>`
    + `</plotArea>${o.legend ? '<legend><legendPos val="r"/></legend>' : ''}</chart></chartSpace>`
}
const parse = (xml: string, theme?: ReturnType<typeof parseThemeContext>): ChartModel => parseXlsxChart(parseXmlOrdered(xml), theme)!
const model = (series: string, o: Options = {}) => prepareChartAxes(parse(contourXml(series, o)))

const measureCtx = createCanvas(W, H).getContext('2d')
/** The plot box the paint uses (the same deterministic layout: size, dpr 1, font generation 0, default family). */
function plotOf(m: ChartModel) {
  return layoutChart(m, W, H, 1, 0, {
    width: (text, css) => { measureCtx.font = css; return measureCtx.measureText(text).width },
    lineHeight: () => m.fontSizePt * 96 / 72,
  }).plot
}
interface Fill { x: number; y: number; w: number; h: number; color: string }
interface Segment { x0: number; y0: number; x1: number; y1: number; color: string }
/** Paints and records every non-white fillRect, every moveTo/lineTo segment (with the stroke colour) and the text. */
function record(m: ChartModel) {
  const ctx = createCanvas(W, H).getContext('2d') as any
  const fills: Fill[] = [], segments: Segment[] = [], texts: Array<{ text: string; x: number; y: number }> = []
  let at: [number, number] = [0, 0]
  const fillRect = ctx.fillRect.bind(ctx), moveTo = ctx.moveTo.bind(ctx), lineTo = ctx.lineTo.bind(ctx), fillText = ctx.fillText.bind(ctx)
  ctx.fillRect = (x: number, y: number, w: number, h: number) => { if (ctx.fillStyle !== '#ffffff') fills.push({ x, y, w, h, color: ctx.fillStyle }); fillRect(x, y, w, h) }
  ctx.moveTo = (x: number, y: number) => { at = [x, y]; moveTo(x, y) }
  ctx.lineTo = (x: number, y: number) => { segments.push({ x0: at[0], y0: at[1], x1: x, y1: y, color: ctx.strokeStyle }); at = [x, y]; lineTo(x, y) }
  ctx.fillText = (text: string, x: number, y: number, ...rest: unknown[]) => { texts.push({ text, x, y }); fillText(text, x, y, ...rest) }
  paintXlsxChart(m, ctx, W, H)
  return { fills, segments, texts, ctx }
}
const near = (a: number, b: number) => Math.abs(a - b) < 1e-6
/** The fill colour painted for grid cell (row, col) of a rows x cols grid (row 0 at the bottom), or undefined. */
function cellColor(fills: Fill[], plot: { x: number; y: number; w: number; h: number }, rows: number, cols: number, row: number, col: number): string | undefined {
  const cw = plot.w / cols, ch = plot.h / rows
  const x = plot.x + col * cw, y = plot.y + plot.h - (row + 1) * ch
  return fills.find(f => near(f.x, x) && near(f.y, y) && near(f.w, cw) && near(f.h, ch))?.color
}
/** Fills inside the plot box (the band cells; legend swatches and the chart background are outside it). */
const plotFills = (fills: Fill[], plot: { x: number; y: number; w: number; h: number }) => fills.filter(f => f.x >= plot.x - 1e-6 && f.y >= plot.y - 1e-6 && f.x + f.w <= plot.x + plot.w + 1e-6 && f.y + f.h <= plot.y + plot.h + 1e-6)
/** Segments strictly inside the plot box (wireframe band boundaries; axis lines and tick marks sit on its edges). */
const interior = (segments: Segment[], plot: { x: number; y: number; w: number; h: number }) => segments.filter(s => [s.x0, s.x1].every(x => x > plot.x + 0.5 && x < plot.x + plot.w - 0.5) || [s.y0, s.y1].every(y => y > plot.y + 0.5 && y < plot.y + plot.h - 0.5))
  .filter(s => [s.x0, s.x1].every(x => x >= plot.x - 1e-6 && x <= plot.x + plot.w + 1e-6) && [s.y0, s.y1].every(y => y >= plot.y - 1e-6 && y <= plot.y + plot.h + 1e-6))

describe('charts-families-contour (Task A5d)', () => {
  it('contour bands from the major unit; grid at most 256x256', () => {
    // Data 1..6 with major unit 2 and no authored minimum: the shared value axis resolves to 0..6, ticks 0,2,4,6,
    // so the bands are [0,2) [2,4) [4,6] (Office band count 3) and 1 and 2 fall in different bands.
    const m = model(ser(0, [1, 2, 3, 4, 5, 6]))
    const value = m.preparedAxes!.axes.find(a => a.role === 'value')!
    expect(value.ticks).toEqual([0, 2, 4, 6])
    const { fills } = record(m), plot = plotOf(m)
    expect(plotFills(fills, plot)).toHaveLength(6)
    expect([0, 1, 2, 3, 4, 5].map(c => cellColor(fills, plot, 1, 6, 0, c))).toEqual([ACCENT[0], ACCENT[1], ACCENT[1], ACCENT[2], ACCENT[2], ACCENT[2]])
    // An authored minimum moves the grid: the shared axis resolves 1..6 with ticks 1,3,5, so the bands are
    // [1,3) [3,5) and the axis's last partial interval [5,6].
    const authored = model(ser(0, [1, 2, 3, 4, 5, 6]), { valAx: '<scaling><min val="1"/></scaling><majorUnit val="2"/>' })
    const authoredAxis = authored.preparedAxes!.axes.find(a => a.role === 'value')!
    expect(authoredAxis.ticks).toEqual([1, 3, 5])
    expect(authoredAxis.max).toBe(6)
    const b = record(authored), bp = plotOf(authored)
    expect([0, 1, 2, 3, 4, 5].map(c => cellColor(b.fills, bp, 1, 6, 0, c))).toEqual([ACCENT[0], ACCENT[0], ACCENT[1], ACCENT[1], ACCENT[2], ACCENT[2]])
  })

  it('bands use the zero-aligned shared value axis when min is not authored (re-review probe 1)', () => {
    const m = model(ser(0, [1, 2, 3, 4, 5, 6]))
    const { fills } = record(m), plot = plotOf(m)
    expect(cellColor(fills, plot, 1, 6, 0, 1)).not.toBe(cellColor(fills, plot, 1, 6, 0, 0))
  })

  it('without an authored major unit the bands follow the automatic axis unit, not a step of 1', () => {
    const m = model(ser(0, [0, 2500, 5000, 7500, 10000]), { valAx: '' })
    const value = m.preparedAxes!.axes.find(a => a.role === 'value')!
    expect(value.major).toBe(2000)
    expect(value.ticks).toEqual([0, 2000, 4000, 6000, 8000, 10000])
    const { fills } = record(m), plot = plotOf(m)
    // 0 -> band 0, 2500 -> 1, 5000 -> 2, 7500 -> 3, 10000 -> the last band (4).
    expect([0, 1, 2, 3, 4].map(c => cellColor(fills, plot, 1, 5, 0, c))).toEqual([ACCENT[0], ACCENT[1], ACCENT[2], ACCENT[3], ACCENT[4]])
  })

  it('rows are series from the bottom, columns are categories from the left', () => {
    const m = model(ser(0, [1, 3]) + ser(1, [5, 1]))
    const { fills } = record(m), plot = plotOf(m)
    expect(cellColor(fills, plot, 2, 2, 0, 0)).toBe(ACCENT[0])
    expect(cellColor(fills, plot, 2, 2, 0, 1)).toBe(ACCENT[1])
    expect(cellColor(fills, plot, 2, 2, 1, 0)).toBe(ACCENT[2])
    expect(cellColor(fills, plot, 2, 2, 1, 1)).toBe(ACCENT[0])
  })

  it('grid capped 256x256 with a diagnostic: columns clipped to 256 are diagnosed and drawn 256 wide', () => {
    const big = (i: number) => ser(i, Array.from({ length: 300 }, (_, k) => k + 1))
    const m = model(big(0) + big(1))
    expect(m.truncated).toBe(true)
    expect(m.diagnostics?.some(d => d.code === 'chart-contour-grid-limit')).toBe(true)
    const { fills } = record(m), plot = plotOf(m)
    const cells = plotFills(fills, plot)
    expect(cells).toHaveLength(2 * 256)
    for (const f of cells) expect(f.w).toBeCloseTo(plot.w / 256, 9)
    // The category axis has 256 slots (its labels sit on the drawn columns).
    expect(m.preparedAxes!.axes.find(a => a.role === 'category')!.slots).toBe(256)
  })

  it('category-only overflow is diagnosed when paint clips to 256 (re-review probe 3)', () => {
    const m = model(ser(0, [1], 300))
    expect(m.diagnostics?.some(d => d.code === 'chart-contour-grid-limit')).toBe(true)
    expect(m.truncated).toBe(true)
    const { fills } = record(m), plot = plotOf(m)
    const cells = plotFills(fills, plot)
    expect(cells).toHaveLength(1)
    expect(cells[0].w).toBeCloseTo(plot.w / 256, 9)
  })

  it('more than 64 series: rows clipped to 64 are diagnosed', () => {
    const m = model(Array.from({ length: 70 }, (_, i) => ser(i, [1, 2])).join(''))
    expect(m.truncated).toBe(true)
    expect(m.diagnostics?.some(d => d.code === 'chart-contour-grid-limit')).toBe(true)
    const { fills } = record(m), plot = plotOf(m)
    expect(plotFills(fills, plot)).toHaveLength(64 * 2)
  })

  it('paint admission covers the band grid at the dimensions the painter uses (re-review probe 4)', () => {
    // 64 series of one stored value each over 256 categories: the painter's grid is 64 x 256 (categories decide the
    // columns), one Int16 band per cell; admission must charge at least those 2-byte cells over a one-cell chart.
    const wide = model(Array.from({ length: 64 }, (_, i) => ser(i, [1], 256)).join(''))
    const one = model(ser(0, [1]))
    const bytes = (m: ChartModel) => paintWorkspaceBytes(m, displayView(m))
    const extraSeriesBytes = 63 * 256 // the per-series term the generic formula already charges for the 63 extra series
    expect(bytes(wide) - bytes(one) - extraSeriesBytes).toBeGreaterThanOrEqual((64 * 256 - 1) * 2)
  })

  it('no interpolation across missing cells', () => {
    const m = model(ser(0, [1, undefined, 5]) + ser(1, [4, 5, 6]))
    const { fills } = record(m), plot = plotOf(m)
    const cells = plotFills(fills, plot)
    expect(cells).toHaveLength(5)
    expect(cellColor(fills, plot, 2, 3, 0, 1)).toBeUndefined()
    expect(cellColor(fills, plot, 2, 3, 0, 0)).toBe(ACCENT[0])
    expect(cellColor(fills, plot, 2, 3, 0, 2)).toBe(ACCENT[2])
  })

  it('wireframe draws boundaries only; bandFmts colours', () => {
    // Row 0: [1, 1]; row 1: [1, 5]. Bands: 1 -> 0, 5 -> 2. Boundaries: the vertical edge between (1,0) and (1,1),
    // the horizontal edge between (0,1) and (1,1). Same-band pairs (0,0)-(0,1) and (0,0)-(1,0) draw nothing.
    const m = model(ser(0, [1, 1]) + ser(1, [1, 5]), { wireframe: true })
    expect(m.groups[0].wireframe).toBe(true)
    const { fills, segments } = record(m), plot = plotOf(m)
    expect(plotFills(fills, plot)).toEqual([])
    const lines = interior(segments, plot)
    const midX = plot.x + plot.w / 2, midY = plot.y + plot.h / 2
    expect(lines).toHaveLength(2)
    const vertical = lines.find(s => near(s.x0, midX) && near(s.x1, midX))!
    expect(vertical).toBeDefined()
    expect([vertical.y0, vertical.y1].sort((a, b) => a - b)).toEqual([plot.y, midY].map(v => expect.closeTo(v, 9)))
    const horizontal = lines.find(s => near(s.y0, midY) && near(s.y1, midY))!
    expect(horizontal).toBeDefined()
    expect([horizontal.x0, horizontal.x1].sort((a, b) => a - b)).toEqual([midX, plot.x + plot.w].map(v => expect.closeTo(v, 9)))
    // A boundary takes the higher band's colour (the band whose lower edge it is).
    expect(vertical.color).toBe(ACCENT[2])
    expect(horizontal.color).toBe(ACCENT[2])
    // Authored band colours (c:bandFmts, child c:idx per CT_BandFmt) colour the wireframe too.
    const red = model(ser(0, [1, 1]) + ser(1, [1, 5]), { wireframe: true, bandFmts: '<bandFmts><bandFmt><idx val="2"/><spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></spPr></bandFmt></bandFmts>' })
    expect(interior(record(red).segments, plotOf(red)).map(s => s.color)).toEqual(['#ff0000', '#ff0000'])
  })

  it('wireframe draws the edge with a missing neighbour in every direction, each pair once (re-review probe 2)', () => {
    for (const values of [[undefined, 1], [1, undefined]]) {
      const m = model(ser(0, values), { wireframe: true })
      const { segments } = record(m), plot = plotOf(m)
      const lines = interior(segments, plot)
      const midX = plot.x + plot.w / 2
      expect(lines, String(values)).toHaveLength(1)
      expect(near(lines[0].x0, midX) && near(lines[0].x1, midX), String(values)).toBe(true)
      expect(lines[0].color, String(values)).toBe(ACCENT[0])
    }
    // Missing below / above a populated cell (rows).
    for (const [a, b] of [[undefined, 1], [1, undefined]]) {
      const m = model(ser(0, [a]) + ser(1, [b]), { wireframe: true })
      const { segments } = record(m), plot = plotOf(m)
      const lines = interior(segments, plot)
      const midY = plot.y + plot.h / 2
      expect(lines, `${a},${b}`).toHaveLength(1)
      expect(near(lines[0].y0, midY) && near(lines[0].y1, midY), `${a},${b}`).toBe(true)
    }
    // Two missing neighbours share no boundary.
    const empty = model(ser(0, [undefined, undefined, 1]), { wireframe: true })
    expect(interior(record(empty).segments, plotOf(empty))).toHaveLength(1)
  })

  it('bandFmts colours; noFill bands stay unpainted; the attribute idx form is not CT_BandFmt', () => {
    const fmts = '<bandFmts><bandFmt><idx val="0"/><spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></spPr></bandFmt><bandFmt><idx val="1"/><spPr><a:noFill/></spPr></bandFmt></bandFmts>'
    const m = model(ser(0, [1, 3, 5]), { bandFmts: fmts })
    expect(m.groups[0].bandColors?.[0]).toBe('FF0000')
    expect(m.groups[0].bandColors?.[1]).toBeNull()
    const { fills } = record(m), plot = plotOf(m)
    expect(cellColor(fills, plot, 1, 3, 0, 0)).toBe('#ff0000')
    expect(cellColor(fills, plot, 1, 3, 0, 1)).toBeUndefined()
    expect(cellColor(fills, plot, 1, 3, 0, 2)).toBe(ACCENT[2])
    // <bandFmt idx="0"> is not the schema's form (idx is a child element): ignored, the accent cycle applies.
    const bogus = model(ser(0, [1, 3, 5]), { bandFmts: '<bandFmts><bandFmt idx="1"><spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></spPr></bandFmt></bandFmts>' })
    expect(bogus.groups[0].bandColors ?? []).toEqual([])
    const r = record(bogus)
    expect(cellColor(r.fills, plotOf(bogus), 1, 3, 0, 1)).toBe(ACCENT[1])
  })

  it('without bandFmts the bands take the document theme accent cycle', () => {
    const scheme = ['112233', '445566', '778899', 'AABBCC', 'DDEEFF', '010203']
    const theme = parseThemeContext(`<a:theme xmlns:a="${A}"><a:themeElements><a:clrScheme name="t"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="000000"/></a:dk2><a:lt2><a:srgbClr val="FFFFFF"/></a:lt2>${scheme.map((c, i) => `<a:accent${i + 1}><a:srgbClr val="${c}"/></a:accent${i + 1}>`).join('')}<a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>`)
    const m = prepareChartAxes(parse(contourXml(ser(0, [1, 3, 5])), theme))
    const { fills } = record(m), plot = plotOf(m)
    expect([0, 1, 2].map(c => cellColor(fills, plot, 1, 3, 0, c))).toEqual(['#112233', '#445566', '#778899'])
  })

  it('contour draws its category and series axes and a band legend', () => {
    const m = model(ser(0, [1, 3, 5]) + ser(1, [2, 4, 6]), { legend: true })
    const { texts, fills } = record(m), plot = plotOf(m)
    const drawn = texts.map(t => t.text)
    // Category labels under the plot, series names left of it (series 0 lowest).
    for (const c of ['C0', 'C1', 'C2']) expect(drawn).toContain(c)
    const c0 = texts.find(t => t.text === 'C0')!, c2 = texts.find(t => t.text === 'C2')!
    expect(c0.y).toBeGreaterThan(plot.y + plot.h)
    expect(c0.x).toBeLessThan(c2.x)
    const r0 = texts.find(t => t.text === 'R0')!, r1 = texts.find(t => t.text === 'R1')!
    expect(r0.x).toBeLessThanOrEqual(plot.x)
    expect(r0.y).toBeGreaterThan(r1.y)
    // The legend lists the bands (value-axis tick intervals), not the series, each with its band colour.
    expect(drawn).toEqual(expect.arrayContaining(['0-2', '2-4', '4-6']))
    const legendSwatches = fills.filter(f => f.x > plot.x + plot.w).map(f => f.color)
    expect(legendSwatches).toEqual([ACCENT[0], ACCENT[1], ACCENT[2]])
  })

  it('renders natively from a DOCX: native model, coverage chart, band geometry', async () => {
    const chartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${contourXml(ser(0, [1, 3]) + ser(1, [5, 1])).replace('<chartSpace>', `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}">`).replace('</chartSpace>', '</c:chartSpace>').replace(/<(\/?)(?!c:|a:)([A-Za-z]+)/g, '<$1c:$2')}`
    const doc = await parseDocx(await OfficePackage.load(buildChartDocx([{ chartXml }]).bytes), { chartContext: documentChartContext('desktop') })
    const charts = doc.sections.flatMap(s => s.paragraphs).flatMap(p => p.images ?? []).flatMap(i => i.drawing?.kind === 'xlsxChart' ? [i.drawing.chart] : [])
    expect(charts).toHaveLength(1)
    const chart = charts[0]
    expect(chart.unsupported).toBeUndefined()
    expect(chart.refused).toBeUndefined()
    expect(chart.groups[0].family).toBe('contour')
    const coverage = (doc.drawingCoverage ?? []).filter(e => e.scope === 'original')
    expect(coverage).toHaveLength(1)
    expect(coverage[0]).toMatchObject({ feature: 'chart', status: 'native', selectedRepresentation: 'chart' })
    const { fills } = record(chart), plot = plotOf(chart)
    expect(plotFills(fills, plot)).toHaveLength(4)
    expect([cellColor(fills, plot, 2, 2, 0, 0), cellColor(fills, plot, 2, 2, 0, 1), cellColor(fills, plot, 2, 2, 1, 0), cellColor(fills, plot, 2, 2, 1, 1)]).toEqual([ACCENT[0], ACCENT[1], ACCENT[2], ACCENT[0]])
  })

  it('surface3D uses the fallback picture, else a card', async () => {
    const surface3D = `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:plotArea><c:surface3DChart>${ser(0, [1, 2]).replace(/<(\/?)([A-Za-z]+)/g, '<$1c:$2')}<c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:surface3DChart><c:catAx><c:axId val="1"/><c:crossAx val="3"/></c:catAx><c:valAx><c:axId val="3"/><c:crossAx val="1"/></c:valAx><c:serAx><c:axId val="2"/><c:crossAx val="3"/></c:serAx></c:plotArea></c:chart></c:chartSpace>`
    // No fallback: a card, the only representation.
    const doc = await parseDocx(await OfficePackage.load(buildChartDocx([{ chartXml: surface3D }]).bytes), { chartContext: documentChartContext('desktop') })
    const contents = doc.sections.flatMap(s => s.paragraphs).flatMap(p => p.images ?? []).flatMap(i => i.drawing ? [i.drawing] : [])
    expect(contents).toHaveLength(1)
    const card = contents[0]
    if (card.kind !== 'xlsxChart') throw new Error('not a chart card')
    expect(card.chart.unsupported).toBe('3-D surface charts are shown as a placeholder')
    expect(card.chart.groups).toEqual([])
    expect((doc.drawingCoverage ?? []).filter(e => e.scope === 'original')[0]).toMatchObject({ feature: 'chart', selectedRepresentation: 'chart-placeholder' })
    // With an authored fallback picture (mc:AlternateContent): the picture, never a card.
    const picture = await surfaceWithFallback(surface3D)
    const images = picture.sections[0].paragraphs[0]?.images ?? []
    expect(images.filter(i => i.drawing)).toEqual([])
    expect(images.filter(i => !i.drawing && i.data.length)).toHaveLength(1)
    expect((picture.drawingCoverage ?? []).filter(e => e.scope === 'original')[0]).toMatchObject({ status: 'fallback', representation: 'fallback', selectedRepresentation: 'raster-fallback' })
  })

  it('conformance: empty, all blank, non-finite and negative values, single point', () => {
    // No series: a card with the reason, never an empty native chart.
    expect(parse(contourXml('')).unsupported).toBe('This chart has no data series')
    // All blank: native, nothing in the plot.
    const blank = model(ser(0, [undefined, undefined]))
    expect(blank.unsupported).toBeUndefined()
    expect(plotFills(record(blank).fills, plotOf(blank))).toEqual([])
    // A negative value extends the shared axis below zero: -5 -> band 0 of ticks -6,-4,...; non-numbers are gaps.
    const negative = model(`<ser><idx val="0"/><order val="0"/>${cats(3)}<val><numLit><ptCount val="3"/><pt idx="0"><v>NaN</v></pt><pt idx="1"><v>INF</v></pt><pt idx="2"><v>-5</v></pt></numLit></val></ser>`)
    expect(negative.preparedAxes!.axes.find(a => a.role === 'value')!.min).toBeLessThanOrEqual(-5)
    const n = record(negative), np = plotOf(negative)
    expect(plotFills(n.fills, np)).toHaveLength(1)
    expect(cellColor(n.fills, np, 1, 3, 0, 2)).toBe(ACCENT[0])
    // One point fills the whole plot.
    const single = model(ser(0, [7]))
    const s = record(single), sp = plotOf(single)
    expect(plotFills(s.fills, sp)).toHaveLength(1)
    const [only] = plotFills(s.fills, sp)
    for (const k of ['x', 'y', 'w', 'h'] as const) expect(only[k]).toBeCloseTo(sp[k], 9)
  })
})

// The A5a family harness: contour and wireframe contour are registered in CLASSIC_FAMILIES (run by the
// bubble/radar suite with native expectations); surface3D is a card in every case.
describeFamilyConformance(SURFACE3D_FAMILY)

/** A DOCX whose chart sits in mc:Choice with a fallback picture (the representation contract, spec 6.3). */
async function surfaceWithFallback(chart: string) {
  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006', PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
  const inline = (id: number, uri: string, payload: string) => `<w:drawing><wp:inline><wp:extent cx="1905000" cy="1524000"/><wp:docPr id="${id}" name="Object ${id}"/><a:graphic><a:graphicData uri="${uri}">${payload}</a:graphicData></a:graphic></wp:inline></w:drawing>`
  const body = `<mc:AlternateContent><mc:Choice Requires="c">${inline(1, C, '<c:chart r:id="rC"/>')}</mc:Choice><mc:Fallback>${inline(2, PIC, '<pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="Picture 2"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="im"/></pic:blipFill><pic:spPr/></pic:pic>')}</mc:Fallback></mc:AlternateContent>`
  const c = createCanvas(2, 2), g = c.getContext('2d'); g.fillStyle = '#00ff00'; g.fillRect(0, 0, 2, 2)
  const zip = new JSZip()
  zip.file('word/document.xml', `<w:document xmlns:w="${W_NS}" xmlns:wp="${WP}" xmlns:mc="${MC}" xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}" xmlns:pic="${PIC}"><w:body><w:p><w:r>${body}</w:r></w:p></w:body></w:document>`)
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rC" Type="${R}/chart" Target="charts/chart1.xml"/><Relationship Id="im" Type="${R}/image" Target="media/image1.png"/></Relationships>`)
  zip.file('word/charts/chart1.xml', chart)
  zip.file('word/media/image1.png', new Uint8Array(c.toBuffer('image/png')))
  return parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })))
}
