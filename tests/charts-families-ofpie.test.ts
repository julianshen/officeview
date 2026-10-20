// @vitest-environment node
import { expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import { prepareChartAxes } from '../src/charts/prepare'
import { paintXlsxChart } from '../src/charts/paint'
import { PointWriter } from '../src/charts/points'
import {
  ofPieSplit, ofPieLayout, ofPieSlices, ofPieConnectors,
} from '../src/charts/families/ofPie'
import type { ChartModel } from '../src/charts/model'

const parseChart = (xml: string): ChartModel => parseXlsxChart(parseXmlOrdered(xml), undefined)!

it('auto splits ceil(n/3) last points; pos, val, percent and cust splits', () => {
  const auto = (n: number) => ofPieSplit({ splitType: 'auto' }, n, () => 1)
  expect(auto(7).secondary).toEqual([false, false, false, false, true, true, true])
  expect(auto(6).secondary).toEqual([false, false, false, false, true, true])
  expect(auto(3).secondary).toEqual([false, true, true])
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
  // a per-point dPt keeps the colour of its original index
  const colored = parseChart(`<chartSpace><chart><plotArea><ofPieChart><ser><dPt><idx val="3"/><spPr><solidFill><srgbClr val="00AA00"/></solidFill></spPr></dPt><val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt><pt idx="2"><v>30</v></pt><pt idx="3"><v>40</v></pt></numLit></val></ser></ofPieChart></plotArea></chart></chartSpace>`)
  expect(colored.unsupported).toBeUndefined()
  expect(colored.groups[0].series[0].dPt?.get(3)?.color).toBe('00AA00')
})

it('secondPieSize, serLines and bar-of-pie gapWidth', () => {
  // Geometry constants: horizontal margins 5 + 5, vertical radius cap (h - 8) / 2,
  // pie gap 25 % of r1, bar gap == gapWidth % of r2; r2 = secondPieSize % of r1.
  const area = { x: 0, y: 0, w: 460, h: 308 }
  const pie = ofPieLayout(area, 'pie', 75, 150)
  expect(pie.primary).toEqual({ cx: 125, cy: 154, r: 120 })
  expect(pie.gap).toBe(30)
  expect(pie.secondary).toEqual({ cx: 365, cy: 154, r: 90 })
  const bigger = ofPieLayout(area, 'pie', 200, 150)
  expect(bigger.primary.r).toBe(72)
  expect(bigger.secondary.r).toBe(144)
  // clamped secondPieSize: 5..200
  expect(ofPieLayout(area, 'pie', 300, 150).secondary.r).toBe(ofPieLayout(area, 'pie', 200, 150).secondary.r)
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
  const lines = ofPieConnectors(pie, slices.other)
  expect(lines).toHaveLength(2)
  for (const line of lines) expect(Math.hypot(line.x1 - pie.primary.cx, line.y1 - pie.primary.cy)).toBeCloseTo(120, 6)
  // pie targets sit at the secondary region's left rim, at a quarter and three quarter height
  expect(lines[0].x2).toBeCloseTo(365 - 90, 9)
  expect(lines[0].y2).toBeCloseTo(154 - 45, 9)
  expect(lines[1].y2).toBeCloseTo(154 + 45, 9)
  // bar targets: the bar's left edge at one quarter and three quarter of its height
  const barSlices = ofPieSlices({ firstSliceAngle: 0 }, [10, 20, 30, 40], ofPieSplit({ splitType: 'auto' }, 4, i => [10, 20, 30, 40][i]), bar)
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
  expect(parseChart(`<chartSpace><chart><plotArea><ofPieChart><secondPieSize val="500"/><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></ofPieChart></plotArea></chart></chartSpace>`).groups[0].secondPieSize).toBe(200)
  expect(parseChart(`<chartSpace><chart><plotArea><ofPieChart><secondPieSize val="1"/><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></ofPieChart></plotArea></chart></chartSpace>`).groups[0].secondPieSize).toBe(5)
  // cust split indices parse
  const cust = parseChart(`<chartSpace><chart><plotArea><ofPieChart><splitType val="cust"/><custSplit><secondPiePt idx="4"/><secondPiePt idx="1"/></custSplit><ser><val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt><pt idx="2"><v>30</v></pt></numLit></val></ser></ofPieChart></plotArea></chart></chartSpace>`)
  expect(cust.groups[0].splitType).toBe('cust')
  expect(cust.groups[0].secondPiePt).toEqual([4, 1])
  // combination pies keep the card rule
  expect(parseChart(`<chartSpace><chart><plotArea><pieChart><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></pieChart><barChart><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></barChart></plotArea></chart></chartSpace>`).unsupported).toBeDefined()
  expect(parseChart(`<chartSpace><chart><plotArea><ofPieChart><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></ofPieChart><barChart><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></barChart></plotArea></chart></chartSpace>`).unsupported).toBeDefined()
})

type Event =
  | { kind: 'arc'; x: number; y: number; r: number; a0: number; a1: number }
  | { kind: 'fill'; style: string; rect?: { x: number; y: number; w: number; h: number } }
  | { kind: 'stroke'; style: string }

type Events = Event[]
/** The records of one paint: every arc, fill (via fillRect) and stroke the painter emitted. */
function eventsOf(model: ChartModel, w = 460, h = 308): Events {
  const ctx = createCanvas(w, h).getContext('2d') as unknown as CanvasRenderingContext2D
  const events: Event[] = []
  const fillRect = ctx.fillRect.bind(ctx), arc = ctx.arc.bind(ctx), stroke = ctx.stroke.bind(ctx)
  ctx.fillRect = (x, y, rectW, rectH) => { events.push({ kind: 'fill', style: String(ctx.fillStyle), rect: { x, y, w: rectW, h: rectH } }); fillRect(x, y, rectW, rectH) }
  ctx.arc = (x, y, r, a0, a1) => { events.push({ kind: 'arc', x, y, r, a0, a1 }); arc(x, y, r, a0, a1) }
  ctx.stroke = () => { events.push({ kind: 'stroke', style: String(ctx.strokeStyle) }); stroke() }
  paintXlsxChart(prepareChartAxes(model), ctx, w, h)
  ;(ctx as unknown as Record<string, unknown>).fillRect = fillRect
  ;(ctx as unknown as Record<string, unknown>).arc = arc
  ;(ctx as unknown as Record<string, unknown>).stroke = stroke
  return events
}
const R = Math.PI * 2
const valuesXml = '<val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt><pt idx="2"><v>30</v></pt><pt idx="3"><v>40</v></pt></numLit></val>'

// The ofPie regions fit one span inside the paint area, which the shared layout insets by the 10px plot padding.
const AREA = { x: 10, y: 10, w: 460 - 20, h: 308 - 20 }
const CY = AREA.y + AREA.h / 2

it('paints pie-of-pie: primary slices plus Other on the left, secondary on the right, serLines in grey', () => {
  const model = parseChart(`<chartSpace><chart><plotArea><ofPieChart><ser>${valuesXml}</ser></ofPieChart></plotArea></chart></chartSpace>`)
  expect(model.unsupported).toBeUndefined()
  const events = eventsOf(model)
  // auto split of 4 points puts the last 2 in the secondary: primary = 2 slices + Other = 3 arcs at (r1 120, gap 25 %),
  // secondary = 2 arcs a quarter smaller at the right end of the span.
  const r1 = Math.min((AREA.h - 8) / 2, (AREA.w - 10) / (2 + 2 * 0.75 + 0.25)), r2 = r1 * 0.75
  const cx1 = AREA.x + 5 + r1, cx2 = cx1 + r1 + r1 * 0.25 + r2
  const arcs = events.filter(e => e.kind === 'arc')
  const leftArcs = arcs.filter(a => a.kind === 'arc' && Math.abs(a.x - cx1) < 1e-9 && Math.abs(a.r - r1) < 1e-9)
  const rightArcs = arcs.filter(a => a.kind === 'arc' && Math.abs(a.x - cx2) < 1e-9 && Math.abs(a.r - r2) < 1e-9)
  expect(leftArcs).toHaveLength(3)
  expect(rightArcs).toHaveLength(2)
  // the Other slice is the last primary arc and sweeps 70/100 of the circle
  if (leftArcs[2].kind !== 'arc') throw new Error('arc')
  expect(leftArcs[2].a1 - leftArcs[2].a0).toBeCloseTo(0.7 * R, 9)
  // serLines: one grey stroke per connector (2 total); the white slice borders exclude by colour
  const grey = events.filter(e => e.kind === 'stroke' && e.style === '#d9d9d9')
  expect(grey).toHaveLength(2)
})

it('paints bar-of-pie: the secondary slices stack in a bar moved by gapWidth', () => {
  const model = parseChart(`<chartSpace><chart><plotArea><ofPieChart><ofPieType val="bar"/><gapWidth val="0"/><secondPieSize val="100"/><ser>${valuesXml}</ser></ofPieChart></plotArea></chart></chartSpace>`)
  expect(model.unsupported).toBeUndefined()
  const events = eventsOf(model)
  // secondPieSize 100 and gapWidth 0: r1 = (w - 10) / 4, and the bar (one radius wide) touches the primary pie.
  const r1 = Math.min((AREA.h - 8) / 2, (AREA.w - 10) / 4), r2 = r1
  const cx1 = AREA.x + 5 + r1, cx2 = cx1 + r1 + r2, barX = cx2 - r2 / 2, barY = CY - r2, barH = r2 * 2
  // the two segments are fillRects of the bar's width, heights 2/7 and 5/7 of the bar, stacked from the top
  const barRects = events.filter(e => e.kind === 'fill' && e.rect && Math.abs(e.rect.w - r2) < 1e-6 && Math.abs(e.rect.x - barX) < 1e-6)
  expect(barRects).toHaveLength(2)
  if (barRects[0].kind !== 'fill' || barRects[1].kind !== 'fill') throw new Error('rect')
  expect(barRects[0].rect!.h).toBeCloseTo(3 / 7 * barH, 6) // 30 of the 70 secondary total
  expect(barRects[1].rect!.h).toBeCloseTo(4 / 7 * barH, 6) // 40 of the 70
  expect(barRects[0].rect!.y).toBeCloseTo(barY, 6)
  expect(barRects[1].rect!.y).toBeCloseTo(barY + 3 / 7 * barH, 6)
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
