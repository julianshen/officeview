// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXmlOrdered } from '../src/core/xml'
import { parseDocx } from '../src/docx/parse'
import { parsePptx } from '../src/pptx/parse'
import { documentChartContext } from '../src/charts/context'
import { parseXlsxChart } from '../src/charts/parse'
import { CHART_STYLE_BYTES } from '../src/charts/own'
import { createChartWorkspace } from '../src/charts/workspace'
import { releaseChartModel } from '../src/charts/own'
import { prepareChartAxes } from '../src/charts/prepare'
import { paintXlsxChart } from '../src/charts/paint'
import { PointWriter } from '../src/charts/points'
import {
  ofPieSplit, ofPieLayout, ofPieSlices, ofPieConnectors, ofPieOtherIndex,
} from '../src/charts/families/ofPie'
import type { ChartModel } from '../src/charts/model'
import type { DrawingCoverageEntry } from '../src/drawing/coverage'
import { buildChartDocx, buildChartPptx } from './helpers/chart-fixtures'
import { describeFamilyConformance, OF_PIE_FAMILY } from './helpers/chart-family-conformance'

const parseChart = (xml: string): ChartModel => parseXlsxChart(parseXmlOrdered(xml), undefined)!

it('auto splits ceil(n/3) last points; pos, val, percent and cust splits', () => {
  const auto = (n: number) => ofPieSplit({ splitType: 'auto' }, n, () => 1)
  expect(auto(7).secondary).toEqual([false, false, false, false, true, true, true])
  expect(auto(6).secondary).toEqual([false, false, false, false, true, true])
  expect(auto(3).secondary).toEqual([false, false, true])
  expect(auto(1).secondary).toEqual([true])
  expect(auto(0).secondary).toEqual([])
  const pos = (n: number) => ofPieSplit({ splitType: 'pos', splitPos: 2 }, n, () => 1)
  expect(pos(5).secondary).toEqual([false, false, false, true, true])
  expect(ofPieSplit({ splitType: 'pos', splitPos: 9 }, 3, () => 1).secondary).toEqual([true, true, true]) // splitPos clamps to n
  expect(ofPieSplit({ splitType: 'val', splitPos: 3 }, 4, i => [5, 2, 3, 9][i]).secondary).toEqual([false, true, false, false])
  expect(ofPieSplit({ splitType: 'percent', splitPos: 50 }, 4, i => [10, 20, 30, 40][i]).secondary).toEqual([true, true, true, true]) // all four sit below 50 % of the 100 total
  expect(ofPieSplit({ splitType: 'percent', splitPos: 35 }, 4, i => [10, 20, 30, 40][i]).secondary).toEqual([true, true, true, false]) // the three below 35 % of the total
  expect(ofPieSplit({ splitType: 'cust', secondPiePt: [0, 4, 2] }, 4, () => 1).secondary).toEqual([true, false, true, false])
  // cust ignores indices past the point count
  expect(ofPieSplit({ splitType: 'cust', secondPiePt: [1, 9, -1] }, 2, () => 1).secondary).toEqual([false, true])
  // cust indices are source indices: the slots map to their sources before membership is decided
  expect(ofPieSplit({ splitType: 'cust', secondPiePt: [3] }, 3, () => 1, i => [0, 2, 3][i]).secondary).toEqual([false, false, true])
})

it('blank, error and zero points are never moved by the val and percent splits (zero sweep stays primary)', () => {
  // the display value of a blank, an error or an unavailable point is 0, and 0 is not below splitPos
  expect(ofPieSplit({ splitType: 'val', splitPos: 5 }, 4, i => [0, 5, 3, 1][i]).secondary).toEqual([false, false, true, true])
  expect(ofPieSplit({ splitType: 'percent', splitPos: 30 }, 4, i => [0, 10, 20, 70][i]).secondary).toEqual([false, true, true, false])
  // negative values read 0 and follow the pie convention (drawn as 0 sweep on the primary)
  expect(ofPieSplit({ splitType: 'val', splitPos: 5 }, 2, i => [-2, 8][i]).secondary).toEqual([false, false])
  expect(ofPieSplit({ splitType: 'percent', splitPos: 60 }, 2, i => [-2, 8][i]).secondary).toEqual([false, false])
})

it('regional totals conserving; the Other slice equals the secondary total', () => {
  const split = ofPieSplit({ splitType: 'auto' }, 4, i => [10, 20, 30, 40][i])
  expect(split.primaryTotal).toBe(30)
  expect(split.secondaryTotal).toBe(70)
  expect(split.primaryTotal + split.secondaryTotal).toBe(100)
})

it('the Other slice conserves totals and percent labels; original indices keep colours', () => {
  const xml = `<chartSpace><chart><plotArea><ofPieChart><ser><dLbls><numFmt formatCode="0.0%" sourceLinked="0"/><showVal val="0"/><showPercent val="1"/></dLbls><val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt><pt idx="2"><v>30</v></pt><pt idx="3"><v>40</v></pt></numLit></val></ser></ofPieChart></plotArea></chart></chartSpace>`
  const model = parseChart(xml)
  expect(model.unsupported).toBeUndefined()
  const prepared = prepareChartAxes(model).preparedAxes!
  const labels = prepared.dataLabels
  const text = (point: number) => labels.find(l => l.group === 0 && l.series === 0 && l.point === point)?.text
  // Primary percentages include the synthetic Other slice (10 % and 20 % of the conserved 100 total);
  // secondary percentages are of the secondary total only (30/70 = 42.9 %, 40/70 = 57.1 %).
  expect(text(0)).toBe('10.0%')
  expect(text(1)).toBe('20.0%')
  expect(text(2)).toBe('42.9%')
  expect(text(3)).toBe('57.1%')
  // the Other slice's label is the pseudo point n (the series' point count): its percentage reads
  // the region it conserves, and it tags itself `other`
  const other = labels.find(l => l.other)
  expect(other?.point).toBe(ofPieOtherIndex(4))
  expect(other?.text).toBe('70.0%')
  // a per-point dPt keeps the colour of its original index
  const colored = parseChart(`<chartSpace><chart><plotArea><ofPieChart><ser><dPt><idx val="3"/><spPr><solidFill><srgbClr val="00AA00"/></solidFill></spPr></dPt><val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt><pt idx="2"><v>30</v></pt><pt idx="3"><v>40</v></pt></numLit></val></ser></ofPieChart></plotArea></chart></chartSpace>`)
  expect(colored.unsupported).toBeUndefined()
  expect(colored.groups[0].series[0].dPt?.get(3)?.color).toBe('00AA00')
})

it('secondPieSize, serLines and bar-of-pie gapWidth', () => {
  // Geometry constants: horizontal margins 5 + 5, vertical radius cap (h - 8) / 2,
  // the gap is gapWidth % of r2 for both variants (pie-of-pie Office default gap width 50,
  // bar-of-pie 150); r2 = secondPieSize % of r1.
  const area = { x: 0, y: 0, w: 460, h: 308 }
  const pie = ofPieLayout(area, 'pie', 100, 50)
  expect(pie.primary).toEqual({ cx: 105, cy: 154, r: 100 })
  expect(pie.gap).toBe(50)
  expect(pie.secondary).toEqual({ cx: 355, cy: 154, r: 100 })
  // gapWidth scales the pie-of-pie gap off r2 (k = secondPieSize % × gapWidth %): 2 + 2s + k = 9
  const bigger = ofPieLayout(area, 'pie', 200, 150)
  expect(bigger.primary.r).toBe(50)
  expect(bigger.secondary.r).toBe(100)
  // clamped secondPieSize: 5..200
  expect(ofPieLayout(area, 'pie', 300, 50).secondary.r).toBe(ofPieLayout(area, 'pie', 200, 50).secondary.r)
  const bar = ofPieLayout(area, 'bar', 100, 50)
  expect(bar.primary).toEqual({ cx: 105, cy: 154, r: 100 })
  expect(bar.gap).toBe(50)
  expect(bar.bar).toEqual({ x: 305, y: 54, w: 100, h: 200 })
  expect(bar.secondary.r).toBe(100)
  // a larger gapWidth pushes the bar right
  const bar2 = ofPieLayout(area, 'bar', 100, 100)
  expect(bar2.bar!.x).toBeGreaterThan(bar.bar!.x)
  // serLines connectors: from the Other slice's two radius points to fixed secondary targets
  const slices = ofPieSlices({ firstSliceAngle: 0 }, [10, 20, 30, 40], ofPieSplit({ splitType: 'auto' }, 4, i => [10, 20, 30, 40][i]), pie)
  // C1 (gallery inspection): the secondary pie is a full circle of its own points (radians), the bar's segments
  // are height fractions. Auto split of 4 points moves the last ceil(4/3) = 2 (30 and 40): 3/7 and 4/7 of 2π.
  expect(slices.secondary.map(s => s.sweep)).toEqual([3 / 7 * Math.PI * 2, 4 / 7 * Math.PI * 2])
  expect(slices.secondary.reduce((n, s) => n + s.sweep, 0)).toBeCloseTo(Math.PI * 2, 12)
  // C1 phase 1.5 (visual review): Office ignores firstSliceAng for ofPie and centres the composite Other slice on
  // 3 o'clock (angle 0), facing the secondary plot; the primary points run clockwise before it and close the circle.
  // Other = 70 of 100 (the auto split moves 30 and 40): sweep 0.7 × 2π, from −0.35 × 2π to +0.35 × 2π.
  const turned = ofPieSlices({ firstSliceAngle: 120 }, [10, 20, 30, 40], ofPieSplit({ splitType: 'auto' }, 4, i => [10, 20, 30, 40][i]), pie)
  expect(turned).toEqual(slices)
  expect(slices.other.start).toBeCloseTo(-0.35 * Math.PI * 2, 12)
  expect(slices.other.start + slices.other.sweep / 2).toBeCloseTo(0, 12)
  expect(slices.primary.map(s => s.sweep)).toEqual([0.1 * Math.PI * 2, 0.2 * Math.PI * 2])
  expect(slices.primary[0].start + slices.primary[0].sweep + slices.primary[1].sweep).toBeCloseTo(slices.other.start, 12)
  expect(slices.primary[0].start + Math.PI * 2).toBeCloseTo(slices.other.start + slices.other.sweep, 12)
  // The secondary pie starts at 12 o'clock.
  expect(slices.secondary[0].start).toBeCloseTo(-Math.PI / 2, 12)
  // A typical split (Other = 30 %: the auto split of [40, 30, 20, 10] moves 20 and 10): the series lines leave the
  // Other slice's top and bottom rim points (±54°), both facing the secondary plot, and never cross the primary
  // disk: every point of each segment is at least r1 from the primary centre.
  const facing = ofPieSlices({ firstSliceAngle: 200 }, [40, 30, 20, 10], ofPieSplit({ splitType: 'auto' }, 4, i => [40, 30, 20, 10][i]), pie)
  const facingLines = ofPieConnectors(pie, facing.other)
  expect(facing.other.start).toBeCloseTo(-0.15 * Math.PI * 2, 12)
  expect(facingLines[0].x1).toBeCloseTo(pie.primary.cx + pie.primary.r * Math.cos(-0.15 * Math.PI * 2), 9)
  expect(facingLines[0].y1).toBeLessThan(pie.primary.cy)
  expect(facingLines[1].y1).toBeGreaterThan(pie.primary.cy)
  for (const line of facingLines) for (let t = 0; t <= 1; t += 1 / 64) {
    const x = line.x1 + (line.x2 - line.x1) * t, y = line.y1 + (line.y2 - line.y1) * t
    expect(Math.hypot(x - pie.primary.cx, y - pie.primary.cy)).toBeGreaterThanOrEqual(pie.primary.r - 1e-9)
  }
  const lines = ofPieConnectors(pie, slices.other)
  expect(lines).toHaveLength(2)
  for (const line of lines) expect(Math.hypot(line.x1 - pie.primary.cx, line.y1 - pie.primary.cy)).toBeCloseTo(100, 6)
  // pie targets sit at the secondary region's left rim, at a quarter and three quarter height
  expect(lines[0].x2).toBeCloseTo(355 - 100, 9)
  expect(lines[0].y2).toBeCloseTo(154 - 50, 9)
  expect(lines[1].y2).toBeCloseTo(154 + 50, 9)
  // bar targets: the bar's left edge at one quarter and three quarter of its height
  const barSlices = ofPieSlices({ firstSliceAngle: 0 }, [10, 20, 30, 40], ofPieSplit({ splitType: 'auto' }, 4, i => [10, 20, 30, 40][i]), bar)
  expect(barSlices.secondary.map(s => [s.start, s.sweep])).toEqual([[0, 3 / 7], [3 / 7, 4 / 7]])
  const barLines = ofPieConnectors(bar, barSlices.other)
  expect(barLines[0].x2).toBe(305)
  expect(barLines[0].y2).toBe(54 + 50)
  expect(barLines[1].y2).toBe(54 + 150)
})

it('parses ofPie group options and stops rejecting the family; combination pies stay cards', () => {
  const model = parseChart(`<chartSpace><chart><plotArea><ofPieChart><ofPieType val="bar"/><splitType val="percent"/><splitPos val="25"/><secondPieSize val="60"/><gapWidth val="40"/><ser><val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt></numLit></val></ser></ofPieChart></plotArea></chart></chartSpace>`)
  expect(model.unsupported).toBeUndefined()
  const group = model.groups[0]
  expect(group.family).toBe('ofPie')
  expect(group.binding).toEqual({ kind: 'none' })
  expect(group.ofPieType).toBe('bar')
  expect(group.splitType).toBe('percent')
  expect(group.splitPos).toBe(25)
  expect(group.secondPieSize).toBe(60)
  expect(group.gapWidth).toBe(40)
  // defaults and clamps
  const minimal = parseChart(`<chartSpace><chart><plotArea><ofPieChart><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></ofPieChart></plotArea></chart></chartSpace>`)
  expect(minimal.groups[0].splitType).toBeUndefined() // absent falls back to auto at the engine
  expect(minimal.groups[0].secondPieSize).toBeUndefined() // absent falls back to 75 at the engine
  // gapWidth defaults to the schema's CT_GapAmount value (150 %) for both variants
  expect(minimal.groups[0].gapWidth).toBe(150)
  expect(parseChart(`<chartSpace><chart><plotArea><ofPieChart><ofPieType val="bar"/><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></ofPieChart></plotArea></chart></chartSpace>`).groups[0].gapWidth).toBe(150)
  expect(parseChart(`<chartSpace><chart><plotArea><ofPieChart><secondPieSize val="500"/><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></ofPieChart></plotArea></chart></chartSpace>`).groups[0].secondPieSize).toBe(200)
  expect(parseChart(`<chartSpace><chart><plotArea><ofPieChart><secondPieSize val="1"/><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></ofPieChart></plotArea></chart></chartSpace>`).groups[0].secondPieSize).toBe(5)
  // cust split indices parse
  const cust = parseChart(`<chartSpace><chart><plotArea><ofPieChart><splitType val="cust"/><custSplit><secondPiePt idx="4"/><secondPiePt idx="1"/></custSplit><ser><val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt><pt idx="2"><v>30</v></pt></numLit></val></ser></ofPieChart></plotArea></chart></chartSpace>`)
  expect(cust.groups[0].splitType).toBe('cust')
  expect(cust.groups[0].secondPiePt).toEqual([4, 1])
  // the secondPiePt set carries one entry per index (no bound inflation): duplicates deduplicate and
  // indices at or past the point limit disappear
  const repeats = Array.from({ length: 300 }, () => '<secondPiePt idx="0"/>').join('')
  const deduped = parseChart(`<chartSpace><chart><plotArea><ofPieChart><splitType val="cust"/><custSplit>${repeats}<secondPiePt idx="5000"/></custSplit><ser><val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt></numLit></val></ser></ofPieChart></plotArea></chart></chartSpace>`)
  expect(deduped.groups[0].secondPiePt).toEqual([0])
  // combination pies keep the card rule
  expect(parseChart(`<chartSpace><chart><plotArea><pieChart><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></pieChart><barChart><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></barChart></plotArea></chart></chartSpace>`).unsupported).toBeDefined()
  expect(parseChart(`<chartSpace><chart><plotArea><ofPieChart><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></ofPieChart><barChart><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></barChart></plotArea></chart></chartSpace>`).unsupported).toBeDefined()
})

type Arc = { x: number; y: number; r: number; a0: number; a1: number }
type Event =
  | { kind: 'arc'; x: number; y: number; r: number; a0: number; a1: number }
  | { kind: 'fill'; style: string; rect?: { x: number; y: number; w: number; h: number }; arcs?: Arc[] }
  | { kind: 'stroke'; style: string; rect?: { x: number; y: number; w: number; h: number } }
  | { kind: 'text'; text: string; x: number; y: number }

type Events = Event[]
/** The records of one paint: every arc, path fill, stroke and text the painter emitted. */
function eventsOf(model: ChartModel, w = 460, h = 308): Events {
  const ctx = createCanvas(w, h).getContext('2d') as unknown as CanvasRenderingContext2D
  const events: Event[] = []
  let current: Arc[] = []
  const beginPath = ctx.beginPath.bind(ctx)
  ctx.beginPath = () => { current = []; beginPath() }
  const fillRect = ctx.fillRect.bind(ctx), arc = ctx.arc.bind(ctx), fill = ctx.fill.bind(ctx)
  const stroke = ctx.stroke.bind(ctx), strokeRect = ctx.strokeRect.bind(ctx), fillText = ctx.fillText.bind(ctx)
  ctx.fillRect = (x, y, rectW, rectH) => { events.push({ kind: 'fill', style: String(ctx.fillStyle), rect: { x, y, w: rectW, h: rectH } }); fillRect(x, y, rectW, rectH) }
  ctx.arc = (x, y, r, a0, a1) => { const e: Arc = { x, y, r, a0, a1 }; events.push({ ...e, kind: 'arc' }); current.push(e); arc(x, y, r, a0, a1) }
  ctx.fill = () => { events.push({ kind: 'fill', style: String(ctx.fillStyle), arcs: current }); fill() }
  ctx.stroke = () => { events.push({ kind: 'stroke', style: String(ctx.strokeStyle) }); stroke() }
  ctx.strokeRect = (x, y, rectW, rectH) => { events.push({ kind: 'stroke', style: String(ctx.strokeStyle), rect: { x, y, w: rectW, h: rectH } }); strokeRect(x, y, rectW, rectH) }
  ctx.fillText = ((t: string, x: number, y: number) => { events.push({ kind: 'text', text: String(t), x, y }); fillText(t, x, y) }) as typeof ctx.fillText
  paintXlsxChart(prepareChartAxes(model), ctx, w, h)
  ;(ctx as unknown as Record<string, unknown>).fillRect = fillRect
  ;(ctx as unknown as Record<string, unknown>).arc = arc
  ;(ctx as unknown as Record<string, unknown>).stroke = stroke
  return events
}
/** The path fills (one arc each): the painted slice colours in draw order. */
const arcFills = (events: Events, cx: number, r: number) =>
  events.filter((e): e is Extract<Event, { kind: 'fill' }> => e.kind === 'fill' && e.rect === undefined && e.arcs?.length === 1 && Math.abs(e.arcs[0].x - cx) < 1e-9 && Math.abs(e.arcs[0].r - r) < 1e-9)
const R = Math.PI * 2
const valuesXml = '<val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt><pt idx="2"><v>30</v></pt><pt idx="3"><v>40</v></pt></numLit></val>'

// The ofPie regions fit one span inside the paint area, which the shared layout insets by the 10px plot padding.
const AREA = { x: 10, y: 10, w: 460 - 20, h: 308 - 20 }
const CY = AREA.y + AREA.h / 2
// secondPieSize 100 and gapWidth 0 keep exact numbers: one radius reaches each plot edge.
const geo = () => {
  const r1 = Math.min((AREA.h - 8) / 2, (AREA.w - 10) / 4)
  return { r1, r2: r1, cx1: AREA.x + 5 + r1, cx2: AREA.x + 5 + 3 * r1 }
}

it('paints pie-of-pie: primary slices plus Other on the left, secondary on the right; serLines absent draws no connectors', () => {
  const model = parseChart(`<chartSpace><chart><plotArea><ofPieChart><gapWidth val="0"/><secondPieSize val="100"/><ser>${valuesXml}</ser></ofPieChart></plotArea></chart></chartSpace>`)
  expect(model.unsupported).toBeUndefined()
  const events = eventsOf(model)
  // auto split of 4 points puts the last 2 in the secondary: primary = 2 slices + Other = 3 arcs at r1,
  // secondary = 2 arcs at r2 at the right end of the span.
  const { r1, cx1, cx2 } = geo()
  const arcs = events.filter(e => e.kind === 'arc')
  const leftArcs = arcs.filter(a => a.kind === 'arc' && Math.abs(a.x - cx1) < 1e-9 && Math.abs(a.r - r1) < 1e-9)
  const rightArcs = arcs.filter(a => a.kind === 'arc' && Math.abs(a.x - cx2) < 1e-9 && Math.abs(a.r - r1) < 1e-9)
  expect(leftArcs).toHaveLength(3)
  expect(rightArcs).toHaveLength(2)
  // the Other slice is the last primary arc and sweeps 70/100 of the circle
  if (leftArcs[2].kind !== 'arc') throw new Error('arc')
  expect(leftArcs[2].a1 - leftArcs[2].a0).toBeCloseTo(0.7 * R, 9)
  // serLines absent: Office draws no connectors (the chart area's own grey frame is a strokeRect, not a path)
  const connectorStrokes = (events: Events, ...styles: string[]) =>
    events.filter(e => e.kind === 'stroke' && e.rect === undefined && styles.includes(e.style))
  expect(connectorStrokes(events, '#d9d9d9')).toHaveLength(0)
})

it('the serLines element gates the connectors: grey by default, the authored line wins, noFill draws none', () => {
  const base = `<gapWidth val="0"/><secondPieSize val="100"/><ser>${valuesXml}</ser>`
  const plain = parseChart(`<chartSpace><chart><plotArea><ofPieChart>${base}<serLines/></ofPieChart></plotArea></chart></chartSpace>`)
  const styled = parseChart(`<chartSpace><chart><plotArea><ofPieChart>${base}<serLines><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></c:spPr></serLines></ofPieChart></plotArea></chart></chartSpace>`)
  const none = parseChart(`<chartSpace><chart><plotArea><ofPieChart>${base}<serLines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></serLines></ofPieChart></plotArea></chart></chartSpace>`)
  const strokes = (model: ChartModel, ...styles: string[]) =>
    eventsOf(model).filter(e => e.kind === 'stroke' && e.rect === undefined && styles.includes(e.style))
  expect(strokes(plain, '#d9d9d9')).toHaveLength(2)
  expect(strokes(styled, '#0000ff')).toHaveLength(2)
  expect(strokes(styled, '#d9d9d9')).toHaveLength(0)
  expect(strokes(none, '#d9d9d9', '#0000ff')).toHaveLength(0)
})

it('the Other slice is the point n: it takes the next varied colour, its dPt paints it and source dPts keep their colours', () => {
  const { r1, cx1, cx2 } = geo()
  const varied = parseChart(`<chartSpace><chart><plotArea><ofPieChart><gapWidth val="0"/><secondPieSize val="100"/><ser>${valuesXml}</ser></ofPieChart></plotArea></chart></chartSpace>`)
  const varyFills = arcFills(eventsOf(varied), cx1, r1)
  expect(varyFills.map(e => e.style)).toEqual(['#4472c4', '#ed7d31', '#5b9bd5']) // slice 0, slice 1 (accent 2) and the Other slice (the point n = accent 5)
  expect(arcFills(eventsOf(varied), cx2, r1).map(e => e.style)).toEqual(['#a5a5a5', '#ffc000'])
  // a dPt at the point count styles the Other point, and a source-index dPt keeps its own colour
  const styled = parseChart(`<chartSpace><chart><plotArea><ofPieChart><gapWidth val="0"/><secondPieSize val="100"/><ser><dPt><idx val="4"/><spPr><solidFill><srgbClr val="FF0000"/></solidFill></spPr></dPt><dPt><idx val="3"/><spPr><solidFill><srgbClr val="00AA00"/></solidFill></spPr></dPt>${valuesXml}</ser></ofPieChart></plotArea></chart></chartSpace>`)
  const fills = eventsOf(styled)
  expect(arcFills(fills, cx1, r1).map(e => e.style)).toEqual(['#4472c4', '#ed7d31', '#ff0000'])
  expect(arcFills(fills, cx2, r1).map(e => e.style)).toEqual(['#a5a5a5', '#00aa00'])
})

it('bar-of-pie: the secondary slices stack in a bar moved by gapWidth', () => {
  const model = parseChart(`<chartSpace><chart><plotArea><ofPieChart><ofPieType val="bar"/><gapWidth val="0"/><secondPieSize val="100"/><ser>${valuesXml}</ser></ofPieChart></plotArea></chart></chartSpace>`)
  expect(model.unsupported).toBeUndefined()
  const events = eventsOf(model)
  const { r1 } = geo()
  const barX = geo().cx2 - r1 / 2, barY = CY - r1, barH = r1 * 2
  // the two segments are fillRects of the bar's width, heights 3/7 and 4/7 of the bar, stacked from the top
  const barRects = events.filter(e => e.kind === 'fill' && e.rect && Math.abs(e.rect.w - r1) < 1e-6 && Math.abs(e.rect.x - barX) < 1e-6)
  expect(barRects).toHaveLength(2)
  if (barRects[0].kind !== 'fill' || barRects[1].kind !== 'fill') throw new Error('rect')
  expect(barRects[0].rect!.h).toBeCloseTo(3 / 7 * barH, 6) // 30 of the 70 secondary total
  expect(barRects[1].rect!.h).toBeCloseTo(4 / 7 * barH, 6) // 40 of the 70
  expect(barRects[0].rect!.y).toBeCloseTo(barY, 6)
  expect(barRects[1].rect!.y).toBeCloseTo(barY + 3 / 7 * barH, 6)
  // unauthored bar segments draw a 1 px white border each (the Office default for bar-of-pie)
  const whiteBorders = events.filter((e): e is Extract<Event, { kind: 'stroke' }> => e.kind === 'stroke' && e.rect !== undefined && e.style.toLowerCase() === '#ffffff' && Math.abs(e.rect.x - barX) < 1e-6 && Math.abs(e.rect.w - r1) < 1e-6)
  expect(whiteBorders).toHaveLength(2)
  expect(whiteBorders[0].rect!.y).toBeCloseTo(barY, 6)
  expect(whiteBorders[1].rect!.y).toBeCloseTo(barY + 3 / 7 * barH, 6)
})

it('labels sit on their own slices: the plan is computed once per group and series, with slice/label agreement', () => {
  const model = parseChart(`<chartSpace><chart><plotArea><ofPieChart><gapWidth val="0"/><secondPieSize val="100"/><ser><dLbls><showVal val="1"/></dLbls>${valuesXml}</ser></ofPieChart></plotArea></chart></chartSpace>`)
  expect(model.unsupported).toBeUndefined()
  const prepared = prepareChartAxes(model)
  expect(prepared.preparedAxes!.dataLabels, 'four real labels plus the Other one').toHaveLength(5)
  const events = eventsOf(model)
  const texts = events.filter(e => e.kind === 'text').map(e => e.kind === 'text' ? [e.text, e.x, e.y] : null)
  // the Other slice's value label appears (its value is the secondary total): it is painted,
  // and it belongs to the primary region like the Other slice itself
  for (const t of ['10', '20', '30', '40', '70']) expect(texts.some(x => x?.[0] === t)).toBe(true)
  const other = texts.find(x => x?.[0] === '70')!
  expect(other![1] as number).toBeLessThan(AREA.x + AREA.w / 2)
  // every label lands on the primary region, and the secondary labels point at the secondary plot:
  // they keep to their region halves of the canvas
  const leftTexts = texts.filter(t => Array.isArray(t) && (t[1] as number) < AREA.x + AREA.w / 2)
  const rightTexts = texts.filter(t => Array.isArray(t) && (t[1] as number) >= AREA.x + AREA.w / 2)
  expect(leftTexts.length).toBeGreaterThanOrEqual(2)
  expect(rightTexts.length).toBeGreaterThanOrEqual(1)
})

it('hidden categories: slices follow source indices; percent labels agree with the drawn slices', () => {
  const model = parseChart(`<chartSpace><chart><plotArea><ofPieChart><gapWidth val="0"/><secondPieSize val="100"/><ser><dPt><idx val="3"/><spPr><solidFill><srgbClr val="00AA00"/></solidFill></spPr></dPt><dLbls><numFmt formatCode="0.0%" sourceLinked="0"/><showVal val="0"/><showPercent val="1"/></dLbls><val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt><pt idx="2"><v>30</v></pt><pt idx="3"><v>40</v></pt></numLit></val></ser></ofPieChart></plotArea></chart></chartSpace>`)
  const { r1, cx2 } = geo()
  expect(model.unsupported).toBeUndefined()
  // the hidden category (category 1) removes source 1 from the display: slots are sources 0, 2, 3
  model.categoryHidden = new Uint32Array([1])
  const shown = prepareChartAxes(model).preparedAxes!
  const text = (point: number) => shown.dataLabels.find(l => l.group === 0 && l.series === 0 && l.point === point)?.text
  // the auto split counts 3 display slots and moves the last (source 3) to the secondary; the
  // percentages read the drawn regions: 10 and 30 share the conserved 80 (12.5 % and 37.5 %), 40
  // carries its own secondary (100 %), and the Other slice is 40 of the 80
  expect(text(0)).toBe('12.5%')
  expect(text(2)).toBe('37.5%')
  expect(text(3)).toBe('100.0%')
  expect(shown.dataLabels.find(l => l.other)?.text).toBe('50.0%')
  // the secondary pie draws one arc at r2 at the right, painted with its dPt colour at the source index
  const events = eventsOf(model)
  const right = arcFills(events, cx2, r1)
  expect(right).toHaveLength(1)
  expect(right[0].style).toBe('#00aa00')
})

it('explosion: each shifted slice moves along its mid angle; the farthest slice (Other included) shrinks the radius', () => {
  // The Other slice's dPt explosion (idx 4, the series' point n) is 50 while every read point
  // explodes 0: the painted primary radius is r1 / 1.5 only if the Other slice joins the
  // farthest computation.
  const model = parseChart(`<chartSpace><chart><plotArea><ofPieChart><gapWidth val="0"/><secondPieSize val="100"/><ser><dPt><idx val="4"/><explosion val="50"/></dPt>${valuesXml}</ser></ofPieChart></plotArea></chart></chartSpace>`)
  expect(model.unsupported).toBeUndefined()
  const { r1, cx1 } = geo()
  const shrunk = r1 / 1.5
  const events = eventsOf(model)
  const left = events.filter(e => e.kind === 'fill' && e.rect === undefined && e.arcs?.length === 1 && Math.abs(e.arcs[0].r - shrunk) < 1e-9)
  expect(left, 'three primary-region slice arcs at the shrunk radius').toHaveLength(3)
  // no primary arc keeps the unshrunk radius; the secondary region has no explosion and keeps r1
  const unchanged = events.filter((e): e is Extract<Event, { kind: 'fill' }> => e.kind === 'fill' && e.rect === undefined && e.arcs?.length === 1 && Math.abs(e.arcs[0].r - r1) < 1e-9)
  expect(unchanged.filter(e => e.arcs![0].x < AREA.x + AREA.w / 2), 'no primary arc keeps the unshrunk radius').toHaveLength(0)
  // unexploded slices keep their centres; their radius is the shrunk one
  if (left[0].kind !== 'fill' || left[2].kind !== 'fill') throw new Error('fill')
  expect(left[0].arcs![0].x).toBeCloseTo(cx1, 6)
  expect(left[0].arcs![0].y).toBeCloseTo(CY, 6)
  expect(left[0].arcs![0].a1 - left[0].arcs![0].a0).toBeCloseTo(0.1 * R, 9) // 10 of the 100
  // the Other slice (sweep 0.7R, centred on 3 o'clock since C1 phase 1.5) carries the dPt explosion: offset r1·50 %
  // along its mid angle, 0, towards the secondary plot
  const otherArc = left[2].arcs![0]
  expect(otherArc.a1 - otherArc.a0).toBeCloseTo(0.7 * R, 9)
  expect(otherArc.a0).toBeCloseTo(-0.35 * R, 9)
  const otherMid = 0
  expect(otherArc.x).toBeCloseTo(cx1 + (shrunk / 2) * Math.cos(otherMid), 6)
  expect(otherArc.y).toBeCloseTo(CY + (shrunk / 2) * Math.sin(otherMid), 6)
})

it('secondPiePt charges one CHART_STYLE_BYTES.point each, per surviving entry', () => {
  const ws = createChartWorkspace('mobile')
  const charges: number[] = []
  const recorded = new Proxy(ws, {
    get(target, key, receiver) {
      if (key === 'reserve') return (bytes: number, tag: string) => { charges.push(bytes); return target.reserve(bytes, tag) }
      const value = Reflect.get(target, key, receiver)
      return typeof value === 'function' ? (value as (...a: unknown[]) => unknown).bind(target) : value
    },
  })
  const pointCharges = (xml: string) => {
    charges.length = 0
    const model = parseXlsxChart(parseXmlOrdered(xml), undefined, { profile: 'mobile', pkg: undefined as unknown as OfficePackage, partPath: 'xl/charts/chart1.xml', workspace: recorded })!
    releaseChartModel(model)
    return charges.filter(b => b === CHART_STYLE_BYTES.point).length
  }
  const values = '<ser><val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt><pt idx="2"><v>30</v></pt></numLit></val></ser>'
  const base = `<chartSpace><chart><plotArea><ofPieChart><splitType val="cust"/>${values}</ofPieChart></plotArea></chart></chartSpace>`
  expect(pointCharges(base), 'no custSplit leaves no secondPiePt charge').toBe(0)
  // idx 4 repeats (deduped) and 5000 is capped, so only idx 4 and 1 survive: two point charges
  const cust = `<chartSpace><chart><plotArea><ofPieChart><splitType val="cust"/><custSplit><secondPiePt idx="4"/><secondPiePt idx="1"/><secondPiePt idx="4"/><secondPiePt idx="5000"/></custSplit>${values}</ofPieChart></plotArea></chart></chartSpace>`
  expect(pointCharges(cust)).toBe(2)
})

it('conformance sweep: empty, all-blank, NaN, Infinity, negative and single-point series never throw', () => {
  for (const type of ['pie', 'bar'] as const) {
    const points = (n: number, f: (i: number) => { v?: number; state?: 'blank' | 'unavailable' | 'error'; e?: string } | undefined) => {
      const model = parseChart(`<chartSpace><chart><plotArea><ofPieChart><ofPieType val="${type}"/><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></ofPieChart></plotArea></chart></chartSpace>`)
      const writer = new PointWriter(n)
      for (let i = 0; i < n; i++) {
        const p = f(i) ?? { v: 1 }
        if (p.e) writer.error(i, p.e)
        else if (p.state === 'blank') { /* stays blank */ }
        else if (p.state === 'unavailable') writer.unavailable(i)
        else writer.value(i, p.v ?? 0)
      }
      model.groups[0].series[0].values = writer.done()
      return model
    }
    const datasets = [
      points(0, () => undefined), // empty
      points(2, i => (i === 0 ? { state: 'blank' } : { state: 'blank' })),
      points(1, () => ({ v: NaN })),
      points(2, i => (i === 0 ? { v: Infinity } : { v: 1 })),
      points(2, i => (i === 0 ? { v: -5 } : { v: 3 })),
      points(1, () => ({ v: 42 })),
      points(2, i => (i === 0 ? { e: '#N/A' } : { v: 8 })),
    ]
    for (const model of datasets) {
      expect(model.unsupported).toBeUndefined()
      expect(() => paintXlsxChart(prepareChartAxes(model), createCanvas(120, 90).getContext('2d') as unknown as CanvasRenderingContext2D, 120, 90)).not.toThrow()
    }
    // a chart with no series at all becomes the card, not a crash
    const none = parseChart(`<chartSpace><chart><plotArea><ofPieChart></ofPieChart></plotArea></chart></chartSpace>`)
    expect(none.unsupported).toBeDefined()
  }
})

it('ofPie is native through the DOCX and PPTX adapters, with chart coverage', async () => {
  for (const format of ['docx', 'pptx'] as const) {
    const chartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<chartSpace><chart><plotArea><ofPieChart><ser>${valuesXml}</ser></ofPieChart></plotArea></chart></chartSpace>`
    const fixture = format === 'docx' ? buildChartDocx([{ chartXml }]) : buildChartPptx([{ chartXml }])
    const pkg = await OfficePackage.load(fixture.bytes)
    const doc = documentChartContext('desktop')
    let charts: ChartModel[], coverage: DrawingCoverageEntry[]
    if (format === 'docx') {
      const parsed = await parseDocx(pkg, { chartContext: doc })
      charts = parsed.sections.flatMap(s => s.paragraphs).flatMap(p => p.images ?? []).flatMap(i => i.drawing?.kind === 'xlsxChart' ? [i.drawing.chart] : [])
      coverage = (parsed.drawingCoverage ?? []).filter(e => e.scope === 'original')
    } else {
      const parsed = await parsePptx(pkg, { chartContext: doc })
      charts = parsed.slides.flatMap(s => s.shapes).flatMap(s => s.content?.kind === 'xlsxChart' ? [s.content.chart] : [])
      coverage = (parsed.drawingCoverage ?? []).filter(e => e.scope === 'original')
    }
    expect(charts.map(c => [c.unsupported, c.groups[0]?.family]), format).toEqual([[undefined, 'ofPie']])
    expect(coverage.map(c => c.selectedRepresentation), format).toEqual(['chart'])
    // the drawn pie-of-pie: 2 primary slices + Other (3 arcs at r1) and 2 secondary arcs at r2
    const arcs = eventsOf(charts[0]).filter(e => e.kind === 'arc')
    expect(arcs, format).toHaveLength(5)
  }
})

// The family conformance harness (spec 7): ofPie with no series is a card; every other case is native.
describe('ofPie family conformance (spec 7)', () => describeFamilyConformance(OF_PIE_FAMILY))
