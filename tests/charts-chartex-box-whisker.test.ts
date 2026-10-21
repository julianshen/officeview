/**
 * Task B4: Box & Whisker charts (ChartEx, spec 4.3).
 *
 * Requirements (task-B4-brief.md verbatim):
 * - `it('exclusive and inclusive split halves with or without the median; odd and even n')`;
 * - `it('whiskers within 1.5 IQR; outliers when visible')`;
 * - `it('mean marker, mean line and inner points')`.
 */
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { boxMarkColor, computeBoxWhiskerStats } from '../src/charts/chartex/boxWhisker'
import type { ChartModel } from '../src/charts/model'
import { parseChartEx } from '../src/charts/chartex/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { parseXml } from '../src/core/xml'
import { chartParseContextFor } from '../src/charts/context'
import { OfficePackage } from '../src/core/zip'
import JSZip from 'jszip'

import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { parseXlsx } from '../src/xlsx/parse'
import { parseChartPart } from '../src/charts/dataPrecedence'
import { readChartPart } from '../src/charts/read'
import { BOX_WHISKER_FAMILY, describeFamilyConformance } from './helpers/chart-family-conformance'
import { CHART_ANALYTIC_POINT_LIMIT, CHART_POINT_LIMIT } from '../src/charts/parse'
import { CHART_TOO_MANY_POINTS } from '../src/charts/fill'
import { layoutChart } from '../src/charts/layout'
import { prepareChartAxes } from '../src/charts/prepare'
import { ChartRefusal } from '../src/charts/own'
import { createChartWorkspace } from '../src/charts/workspace'

const FIXTURES = join(__dirname, 'fixtures', 'chartex')
const fixture = (name: string) => new Uint8Array(readFileSync(join(FIXTURES, name)))
const pkgPromise = new JSZip().generateAsync({ type: 'uint8array' }).then(b => OfficePackage.load(b))
const testContext = async () => chartParseContextFor(await pkgPromise, 'xl/charts/chartEx1.xml')

/** A synthetic part read by the production bounded reader (real namespace dispatch) from a one-part package. */
async function readSynthetic(xml: string, path: string) {
  const zip = new JSZip()
  zip.file(path, xml)
  const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
  const context = chartParseContextFor(pkg, path)
  const read = await readChartPart(pkg, path, context)
  if (read?.kind !== 'tree') throw new Error(`${path}: ${read ? read.reason : 'missing'}`)
  return { root: read.root, context, release: read.release }
}
/** The production adapter (`parseChartPart`) on a DOCX (default) or PPTX part path. */
async function adapt(xml: string, path = 'word/charts/chart1.xml') {
  const part = await readSynthetic(xml, path)
  try { return await parseChartPart(part.root, undefined, part.context) } finally { part.release() }
}
/** A native model through the adapter: the DOCX/PPTX path runs stage 1 (prepare) once. */
async function nativeModel(xml: string, path?: string): Promise<ChartModel> {
  const result = await adapt(xml, path)
  if (result.kind !== 'native') throw new Error(`${result.kind}: ${result.reason}`)
  return result.model
}
/** The parsed model before stage 1 (no fill, no prepare). */
async function rawModel(xml: string): Promise<ChartModel> {
  const part = await readSynthetic(xml, 'xl/charts/chartEx1.xml')
  try { return parseChartEx(part.root, undefined, part.context) } finally { part.release() }
}
/** The plot box the painter uses (same layout key as paintXlsxChart at dpr 1). */
function plotOf(model: ChartModel, w = 400, h = 300) {
  const ctx = createCanvas(w, h).getContext('2d')
  return layoutChart(model, w, h, 1, 0, { width: (text: string, css: string) => { ctx.font = css; return ctx.measureText(text).width }, lineHeight: () => model.fontSizePt * 96 / 72 })
}
interface Segment { x1: number; y1: number; x2: number; y2: number; stroke: string }
/** Paints and records the filled rectangles inside the canvas, every moveTo/lineTo segment and every dot. */
function record(model: ChartModel, w = 400, h = 300) {
  const ctx = createCanvas(w, h).getContext('2d')
  const rects: Array<{ x: number; y: number; w: number; h: number }> = [], segments: Segment[] = [], dots: Array<{ x: number; y: number; r: number }> = []
  let card = false
  const fillRect = ctx.fillRect.bind(ctx), moveTo = ctx.moveTo.bind(ctx), lineTo = ctx.lineTo.bind(ctx), arc = ctx.arc.bind(ctx)
  ctx.fillRect = (x, y, rw, rh) => {
    if (String(ctx.fillStyle).toLowerCase() === '#f7f7f7' && rw === w && rh === h) card = true
    if (rw > 2 && rh > 1 && y > 0 && y + rh < h && !(rw === w && rh === h)) rects.push({ x, y, w: rw, h: rh })
    fillRect(x, y, rw, rh)
  }
  let lx = 0, ly = 0
  ctx.moveTo = (x, y) => { lx = x; ly = y; moveTo(x, y) }
  ctx.lineTo = (x, y) => { segments.push({ x1: lx, y1: ly, x2: x, y2: y, stroke: String(ctx.strokeStyle).toLowerCase() }); lx = x; ly = y; lineTo(x, y) }
  ctx.arc = (x, y, r, sa, ea) => { if (r <= 4) dots.push({ x, y, r }); arc(x, y, r, sa, ea) }
  paintXlsxChart(model, ctx as unknown as CanvasRenderingContext2D, w, h)
  return { rects, segments, dots, card }
}
interface SeriesSpec { values: number[]; cats?: Array<string | null>; catCount?: number; name?: string; hidden?: boolean; vis?: string; method?: 'inclusive' | 'exclusive' }
/** An Office-shaped synthetic box & whisker part (one data set per series; `null` category: no pt, i.e. empty). */
function bwXml(series: SeriesSpec[], scaling = ''): string {
  const data = series.map((s, i) => {
    const cat = s.cats ? `<cx:strDim type="cat"><cx:lvl ptCount="${s.catCount ?? s.cats.length}">${s.cats.map((c, j) => c === null ? '' : `<cx:pt idx="${j}">${c}</cx:pt>`).join('')}</cx:lvl></cx:strDim>` : ''
    return `<cx:data id="${i}">${cat}<cx:numDim type="val"><cx:lvl ptCount="${s.values.length}" formatCode="General">${s.values.map((v, j) => `<cx:pt idx="${j}">${v}</cx:pt>`).join('')}</cx:lvl></cx:numDim></cx:data>`
  }).join('')
  const nodes = series.map((s, i) => `<cx:series layoutId="boxWhisker"${s.hidden ? ' hidden="1"' : ''}>${s.name ? `<cx:tx><cx:txData><cx:v>${s.name}</cx:v></cx:txData></cx:tx>` : ''}<cx:dataId val="${i}"/><cx:layoutPr>${s.vis ? `<cx:visibility ${s.vis}/>` : ''}<cx:statistics quartileMethod="${s.method ?? 'exclusive'}"/></cx:layoutPr></cx:series>`).join('')
  return `<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cx:chartData>${data}</cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion>${nodes}</cx:plotAreaRegion><cx:axis id="0"><cx:catScaling gapWidth="1"/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling${scaling}/><cx:majorGridlines/><cx:tickLabels/></cx:axis></cx:plotArea></cx:chart></cx:chartSpace>`
}

describe('box & whisker (ChartEx, Task B4)', () => {
  it('exclusive and inclusive split halves with or without the median; odd and even n', () => {
    // --- 1. Odd n: n = 9 (from BW1 series 1: 3, 4, 4, 5, 6, 6, 7, 8, 9) ---
    // Median index = 4, median = 6.
    // Exclusive: lower half [3, 4, 4, 5] -> Q1 = 4. Upper half [6, 7, 8, 9] -> Q3 = 7.5.
    // Inclusive: lower half [3, 4, 4, 5, 6] -> Q1 = 4. Upper half [6, 6, 7, 8, 9] -> Q3 = 7.
    const odd9 = [6, 4, 7, 5, 6, 8, 9, 3, 4]
    const exc9 = computeBoxWhiskerStats(odd9, 'exclusive')!
    expect(exc9.median).toBe(6)
    expect(exc9.q1).toBe(4)
    expect(exc9.q3).toBe(7.5)
    expect(exc9.iqr).toBe(3.5)
    expect(exc9.mean).toBeCloseTo(52 / 9, 6)

    const inc9 = computeBoxWhiskerStats(odd9, 'inclusive')!
    expect(inc9.median).toBe(6)
    expect(inc9.q1).toBe(4)
    expect(inc9.q3).toBe(7)
    expect(inc9.iqr).toBe(3)
    expect(inc9.mean).toBeCloseTo(52 / 9, 6)

    // --- 2. Odd n: n = 7 (sorted: 10, 20, 30, 40, 50, 60, 70) ---
    // Median index = 3, median = 40.
    // Exclusive: lower half [10, 20, 30] -> Q1 = 20. Upper half [50, 60, 70] -> Q3 = 60.
    // Inclusive: lower half [10, 20, 30, 40] -> Q1 = 25. Upper half [40, 50, 60, 70] -> Q3 = 55.
    const odd7 = [50, 10, 70, 40, 20, 60, 30]
    const exc7 = computeBoxWhiskerStats(odd7, 'exclusive')!
    expect(exc7.median).toBe(40)
    expect(exc7.q1).toBe(20)
    expect(exc7.q3).toBe(60)
    expect(exc7.iqr).toBe(40)
    expect(exc7.mean).toBe(40)

    const inc7 = computeBoxWhiskerStats(odd7, 'inclusive')!
    expect(inc7.median).toBe(40)
    expect(inc7.q1).toBe(25)
    expect(inc7.q3).toBe(55)
    expect(inc7.iqr).toBe(30)
    expect(inc7.mean).toBe(40)

    // --- 3. Odd n: n = 3 (sorted: 10, 30, 50) ---
    // Median index = 1, median = 30.
    // Exclusive: lower half [10] -> Q1 = 10. Upper half [50] -> Q3 = 50.
    // Inclusive: lower half [10, 30] -> Q1 = 20. Upper half [30, 50] -> Q3 = 40.
    const odd3 = [30, 10, 50]
    const exc3 = computeBoxWhiskerStats(odd3, 'exclusive')!
    expect(exc3.median).toBe(30)
    expect(exc3.q1).toBe(10)
    expect(exc3.q3).toBe(50)
    expect(exc3.mean).toBe(30)

    const inc3 = computeBoxWhiskerStats(odd3, 'inclusive')!
    expect(inc3.median).toBe(30)
    expect(inc3.q1).toBe(20)
    expect(inc3.q3).toBe(40)
    expect(inc3.mean).toBe(30)

    // --- 4. Odd n: n = 1 (single value: 42) ---
    const odd1 = [42]
    const exc1 = computeBoxWhiskerStats(odd1, 'exclusive')!
    expect(exc1.median).toBe(42)
    expect(exc1.q1).toBe(42)
    expect(exc1.q3).toBe(42)
    expect(exc1.mean).toBe(42)

    const inc1 = computeBoxWhiskerStats(odd1, 'inclusive')!
    expect(inc1.median).toBe(42)
    expect(inc1.q1).toBe(42)
    expect(inc1.q3).toBe(42)
    expect(inc1.mean).toBe(42)

    // --- 5. Even n: n = 6 (sorted: 2, 4, 6, 8, 10, 12) ---
    // Median = (6 + 8) / 2 = 7.
    // Lower half: [2, 4, 6] -> Q1 = 4.
    // Upper half: [8, 10, 12] -> Q3 = 10.
    // Both exclusive and inclusive agree.
    const even6 = [8, 2, 12, 4, 10, 6]
    const exc6 = computeBoxWhiskerStats(even6, 'exclusive')!
    const inc6 = computeBoxWhiskerStats(even6, 'inclusive')!
    expect(exc6.median).toBe(7)
    expect(exc6.q1).toBe(4)
    expect(exc6.q3).toBe(10)
    expect(exc6.mean).toBe(7)
    expect(inc6.median).toBe(7)
    expect(inc6.q1).toBe(4)
    expect(inc6.q3).toBe(10)
    expect(inc6.mean).toBe(7)

    // --- 6. Even n: n = 4 (sorted: 10, 20, 30, 40) ---
    // Median = 25. Lower half [10, 20] -> Q1 = 15. Upper half [30, 40] -> Q3 = 35.
    const even4 = [40, 10, 30, 20]
    const exc4 = computeBoxWhiskerStats(even4, 'exclusive')!
    const inc4 = computeBoxWhiskerStats(even4, 'inclusive')!
    expect(exc4.median).toBe(25)
    expect(exc4.q1).toBe(15)
    expect(exc4.q3).toBe(35)
    expect(exc4.mean).toBe(25)
    expect(inc4.median).toBe(25)
    expect(inc4.q1).toBe(15)
    expect(inc4.q3).toBe(35)
    expect(inc4.mean).toBe(25)

    // --- 7. Even n: n = 2 (sorted: 10, 20) ---
    const even2 = [20, 10]
    const exc2 = computeBoxWhiskerStats(even2, 'exclusive')!
    const inc2 = computeBoxWhiskerStats(even2, 'inclusive')!
    expect(exc2.median).toBe(15)
    expect(exc2.q1).toBe(10)
    expect(exc2.q3).toBe(20)
    expect(exc2.mean).toBe(15)
    expect(inc2.median).toBe(15)
    expect(inc2.q1).toBe(10)
    expect(inc2.q3).toBe(20)
    expect(inc2.mean).toBe(15)
  })

  it('whiskers within 1.5 IQR; outliers when visible', async () => {
    // Dataset with both lower and upper outliers:
    // Values: [-20, 10, 12, 14, 16, 18, 20, 50] (n = 8, even).
    // Sorted: -20, 10, 12, 14, 16, 18, 20, 50
    // Median = (14 + 16) / 2 = 15
    // Lower half: [-20, 10, 12, 14] -> Q1 = (10 + 12) / 2 = 11
    // Upper half: [16, 18, 20, 50] -> Q3 = (18 + 20) / 2 = 19
    // IQR = 19 - 11 = 8
    // 1.5 * IQR = 12
    // Lower fence: 11 - 12 = -1. Points: -20 is < -1, so -20 is outlier.
    // Furthest inlier in lower range: 10. Lower whisker = 10.
    // Upper fence: 19 + 12 = 31. Points: 50 is > 31, so 50 is outlier.
    // Furthest inlier in upper range: 20. Upper whisker = 20.
    // Outliers: [-20, 50].
    // Inliers: [10, 12, 14, 16, 18, 20].
    const data = [14, 50, 10, 18, -20, 16, 20, 12]
    const stats = computeBoxWhiskerStats(data, 'exclusive', true)!
    expect(stats.q1).toBe(11)
    expect(stats.median).toBe(15)
    expect(stats.q3).toBe(19)
    expect(stats.iqr).toBe(8)
    expect(stats.mean).toBe(15)
    expect(stats.lowerWhisker).toBe(10)
    expect(stats.upperWhisker).toBe(20)
    expect(stats.outliers).toEqual([-20, 50])
    expect(stats.inliers).toEqual([10, 12, 14, 16, 18, 20])

    // Verify paint output with outliers visible vs invisible
    const makeXml = (outliersVis: '0' | '1') => `
      <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <cx:chartData>
          <cx:data id="0">
            <cx:numDim type="val">
              <cx:lvl ptCount="8" formatCode="General">
                <cx:pt idx="0">14</cx:pt><cx:pt idx="1">50</cx:pt><cx:pt idx="2">10</cx:pt><cx:pt idx="3">18</cx:pt>
                <cx:pt idx="4">-20</cx:pt><cx:pt idx="5">16</cx:pt><cx:pt idx="6">20</cx:pt><cx:pt idx="7">12</cx:pt>
              </cx:lvl>
            </cx:numDim>
          </cx:data>
        </cx:chartData>
        <cx:chart>
          <cx:plotArea>
            <cx:plotAreaRegion>
              <cx:series layoutId="boxWhisker">
                <cx:dataId val="0"/>
                <cx:layoutPr>
                  <cx:visibility meanLine="0" meanMarker="0" nonoutliers="0" outliers="${outliersVis}"/>
                  <cx:statistics quartileMethod="exclusive"/>
                </cx:layoutPr>
              </cx:series>
            </cx:plotAreaRegion>
            <cx:axis id="0"><cx:catScaling gapWidth="1"/><cx:tickLabels/></cx:axis>
            <cx:axis id="1"><cx:valScaling min="-30" max="60"/><cx:majorGridlines/><cx:tickLabels/></cx:axis>
          </cx:plotArea>
        </cx:chart>
      </cx:chartSpace>
    `

    const countDots = (model: ChartModel) => {
      const canvas = createCanvas(400, 300)
      const ctx = canvas.getContext('2d')
      let dots = 0
      const origArc = ctx.arc.bind(ctx)
      ctx.arc = (x, y, r, sa, ea) => {
        if (r <= 4) dots++
        origArc(x, y, r, sa, ea)
      }
      paintXlsxChart(model, ctx as any, 400, 300)
      return dots
    }

    // 1. With outliers="1": 2 outliers (-20 and 50) drawn as dots
    const mWith = await nativeModel(makeXml('1'))
    expect(countDots(mWith)).toBe(2)

    // 2. With outliers="0": 0 outlier dots drawn
    const mWithout = await nativeModel(makeXml('0'))
    expect(countDots(mWithout)).toBe(0)

    // 3. Pixel positions (M4). One box, no categories: one slot of the plot width (between), gapWidth 1 = 100%, so
    // the box is w/2 wide starting at plot.x + w/4; its centre is plot.x + w/2. The value axis is -30..60:
    // y(v) = plot.y + plot.h * (60 - v) / 90.
    const { plot } = plotOf(mWith)
    const y = (v: number) => plot.y + plot.h * (60 - v) / 90
    const cx = plot.x + plot.w / 2, x0 = plot.x + plot.w / 4, x1 = plot.x + plot.w * 3 / 4
    const drawn = record(mWith)
    expect(drawn.rects).toHaveLength(1)
    expect(drawn.rects[0].x).toBeCloseTo(x0, 6)
    expect(drawn.rects[0].w).toBeCloseTo(plot.w / 2, 6)
    expect(drawn.rects[0].y).toBeCloseTo(y(19), 6) // Q3
    expect(drawn.rects[0].h).toBeCloseTo(y(11) - y(19), 6) // Q1 - Q3
    const near = (sg: Segment, a: [number, number], b: [number, number]) =>
      Math.abs(sg.x1 - a[0]) < 1e-6 && Math.abs(sg.y1 - a[1]) < 1e-6 && Math.abs(sg.x2 - b[0]) < 1e-6 && Math.abs(sg.y2 - b[1]) < 1e-6
    // Lower whisker from the box bottom (Q1 = 11) to the furthest inlier 10; upper from Q3 = 19 to 20 (not the fences).
    expect(drawn.segments.filter(sg => near(sg, [cx, y(11)], [cx, y(10)]))).toHaveLength(1)
    expect(drawn.segments.filter(sg => near(sg, [cx, y(19)], [cx, y(20)]))).toHaveLength(1)
    // Median 15 across the box.
    expect(drawn.segments.filter(sg => near(sg, [x0, y(15)], [x1, y(15)]))).toHaveLength(1)
    // Outliers -20 and 50 on the centre line.
    expect(drawn.dots.map(d => [d.x, d.y])).toEqual([[cx, y(-20)], [cx, y(50)]].map(([a, b]) => [expect.closeTo(a, 6), expect.closeTo(b, 6)]))
  })

  it('mean marker, mean line and inner points', async () => {
    // Two series, no categories (one box per series):
    // Series 0: [10, 20, 30] -> mean 20, inliers [10, 20, 30]. Series 1: [40, 50, 60] -> mean 50, inliers [40, 50, 60].
    // Geometry: 2 slots of plot.w / 2 (between), gapWidth 100% -> each box plot.w / 4 wide, centred in its slot:
    // centre(k) = plot.x + k * plot.w / 2 + plot.w / 4. Value axis 0..70: y(v) = plot.y + plot.h * (70 - v) / 70.
    const xml = (v0: string, v1 = v0) => bwXml([
      { values: [10, 20, 30], vis: v0 },
      { values: [40, 50, 60], vis: v1 },
    ], ' min="0" max="70"')
    const vis = (meanMarker: 0 | 1, meanLine: 0 | 1, nonoutliers: 0 | 1) => `meanLine="${meanLine}" meanMarker="${meanMarker}" nonoutliers="${nonoutliers}" outliers="1"`
    const model0 = await nativeModel(xml(vis(1, 0, 0)))
    const { plot } = plotOf(model0)
    const y = (v: number) => plot.y + plot.h * (70 - v) / 70
    const centre = (k: number) => plot.x + k * plot.w / 2 + plot.w / 4
    const colors = model0.groups[0].series.map(s => `#${s.color.toLowerCase()}`)
    /** Mean-line segments: neither vertical (whiskers), the box width (medians), nor the plot width (gridlines). */
    const meanLines = (r: ReturnType<typeof record>) => r.segments.filter(sg => Math.abs(sg.x2 - sg.x1) > plot.w / 3 && Math.abs(sg.x2 - sg.x1) < plot.w * 2 / 3)
    /** Mean-marker strokes: the two 8 px diagonals of each ×. */
    const crosses = (r: ReturnType<typeof record>) => r.segments.filter(sg => Math.abs(Math.abs(sg.x2 - sg.x1) - 8) < 1e-6 && Math.abs(Math.abs(sg.y2 - sg.y1) - 8) < 1e-6)

    // 1. meanMarker on: one × per series, centred on (centre(k), y(mean)).
    const r1 = record(model0)
    const x1 = crosses(r1)
    expect(x1).toHaveLength(4)
    expect(x1.map(sg => [(sg.x1 + sg.x2) / 2, (sg.y1 + sg.y2) / 2])).toEqual([
      [centre(0), y(20)], [centre(0), y(20)], [centre(1), y(50)], [centre(1), y(50)],
    ].map(([a, b]) => [expect.closeTo(a, 6), expect.closeTo(b, 6)]))
    expect(meanLines(r1)).toHaveLength(0)
    expect(r1.dots).toHaveLength(0)

    // 2. meanLine on for both: exactly one segment from series 0's mean to series 1's mean, in series 1's colour.
    const r2 = record(await nativeModel(xml(vis(0, 1, 0))))
    expect(crosses(r2)).toHaveLength(0)
    expect(r2.dots).toHaveLength(0)
    const m2 = meanLines(r2)
    expect(m2).toHaveLength(1)
    expect(m2[0].x1).toBeCloseTo(centre(0), 6)
    expect(m2[0].y1).toBeCloseTo(y(20), 6)
    expect(m2[0].x2).toBeCloseTo(centre(1), 6)
    expect(m2[0].y2).toBeCloseTo(y(50), 6)
    expect(m2[0].stroke).toBe(colors[1])

    // 3. M1: the flag is per series. One side off: no join.
    expect(meanLines(record(await nativeModel(xml(vis(0, 1, 0), vis(0, 0, 0)))))).toHaveLength(0)
    expect(meanLines(record(await nativeModel(xml(vis(0, 0, 0), vis(0, 1, 0)))))).toHaveLength(0)

    // 4. nonoutliers on: the 3 + 3 inner points, on each box's centre line at their values.
    const r4 = record(await nativeModel(xml(vis(0, 0, 1))))
    expect(crosses(r4)).toHaveLength(0)
    expect(meanLines(r4)).toHaveLength(0)
    expect(r4.dots.map(d => [d.x, d.y])).toEqual([
      [centre(0), y(10)], [centre(0), y(20)], [centre(0), y(30)], [centre(1), y(40)], [centre(1), y(50)], [centre(1), y(60)],
    ].map(([a, b]) => [expect.closeTo(a, 6), expect.closeTo(b, 6)]))
  })

  it('in category mode the mean line joins each series\' means across categories (N2)', async () => {
    // 2 categories x 2 series, meanLine on, markers off.
    // s0: A [9, 10, 11] mean 10, B [19, 20, 21] mean 20. s1: A [39, 40, 41] mean 40, B [49, 50, 51] mean 50.
    // Clustered geometry: band = plot.w / 2, box = band / 3; start(k, s) = plot.x + k*band + band/6 + s*band/3, so the
    // centres are plot.x + plot.w * {1/6 (A,s0), 1/3 (A,s1), 2/3 (B,s0), 5/6 (B,s1)}. y(v) = plot.y + plot.h * (60 - v) / 60.
    // Right: (A,s0)->(B,s0) and (A,s1)->(B,s1). Wrong (box order): (A,s0)->(A,s1), (A,s1)->(B,s0), (B,s0)->(B,s1).
    const vis = 'meanLine="1" meanMarker="0" nonoutliers="0" outliers="1"'
    const cats = ['A', 'A', 'A', 'B', 'B', 'B']
    const model = await nativeModel(bwXml([
      { values: [9, 10, 11, 19, 20, 21], cats, vis },
      { values: [39, 40, 41, 49, 50, 51], cats, vis },
    ], ' min="0" max="60"'))
    const { plot } = plotOf(model)
    const y = (v: number) => plot.y + plot.h * (60 - v) / 60
    const x = (f: number) => plot.x + plot.w * f
    const colors = model.groups[0].series.map(s => `#${s.color.toLowerCase()}`)
    const r = record(model)
    // Box widths are plot.w / 6 (medians); gridlines are plot.w; whiskers are vertical: keep the slanted joins.
    const lines = r.segments.filter(sg => Math.abs(sg.y2 - sg.y1) > 1e-6 && Math.abs(sg.x2 - sg.x1) > 1e-6)
    expect(lines).toHaveLength(2)
    const want = [
      { x1: x(1 / 6), y1: y(10), x2: x(2 / 3), y2: y(20), stroke: colors[0] },
      { x1: x(1 / 3), y1: y(40), x2: x(5 / 6), y2: y(50), stroke: colors[1] },
    ]
    for (const [i, w] of want.entries()) {
      expect(lines[i].x1).toBeCloseTo(w.x1, 6)
      expect(lines[i].y1).toBeCloseTo(w.y1, 6)
      expect(lines[i].x2).toBeCloseTo(w.x2, 6)
      expect(lines[i].y2).toBeCloseTo(w.y2, 6)
      expect(lines[i].stroke).toBe(w.stroke)
    }
  })

  it('BW1, BW2a, BW2b native models, drawing coverage and geometry', async () => {
    // 1. BW1: boxWhisker.xlsx (3 series with 9 points each)
    const pkg1 = await OfficePackage.load(fixture('boxWhisker/boxWhisker.xlsx'))
    const doc1 = await parseXlsx(pkg1)
    const drawing1 = doc1.sheets[0]!.drawings![0]!
    expect(drawing1.content?.kind).toBe('xlsxChart')
    if (drawing1.content?.kind !== 'xlsxChart') throw new Error('not xlsxChart')
    const chart1 = drawing1.content.chart
    expect(chart1.unsupported).toBeUndefined()

    // Assert drawing coverage: selectedRepresentation is 'chart' and status is 'native'
    const cov1 = doc1.sheets[0]!.drawingCoverage?.find(c => c.feature === 'ChartEx' || c.feature === 'chart')
    expect(cov1).toMatchObject({ status: 'native', selectedRepresentation: 'chart' })

    // Measure geometry: 3 boxes, 3 median lines, 6 whiskers, 3 mean markers, 0 card
    const canvas1 = createCanvas(400, 300)
    const ctx1 = canvas1.getContext('2d')
    let rects1 = 0, card1 = false, meanMarkers1 = 0
    const origFillRect1 = ctx1.fillRect.bind(ctx1)
    ctx1.fillRect = (x, y, w, h) => {
      if (String(ctx1.fillStyle).toLowerCase() === '#f7f7f7' && w === 400 && h === 300) card1 = true
      if (w > 2 && h > 1 && y > 0 && y + h < 300) rects1++
      origFillRect1(x, y, w, h)
    }
    let lx1 = 0, ly1 = 0
    ctx1.moveTo = (x, y) => { lx1 = x; ly1 = y }
    ctx1.lineTo = (x, y) => {
      const dx = Math.abs(x - lx1), dy = Math.abs(y - ly1)
      if (dx > 2 && dx < 12 && dy > 2 && dy < 12) meanMarkers1++
      lx1 = x; ly1 = y
    }
    paintXlsxChart(chart1, ctx1 as any, 400, 300)
    expect(card1).toBe(false)
    expect(rects1).toBe(3) // 3 boxes
    expect(meanMarkers1).toBe(6) // 3 mean markers (each 2 diagonal strokes = 6)

    // 2. BW2a: box-whisker-gapwidth.xlsx chart 1 (gapWidth = 2.47)
    const pkg2 = await OfficePackage.load(fixture('box-whisker-gapwidth/box-whisker-gapwidth.xlsx'))
    const doc2 = await parseXlsx(pkg2)
    const drawing2a = doc2.sheets[0]!.drawings![0]!
    expect(drawing2a.content?.kind).toBe('xlsxChart')
    if (drawing2a.content?.kind !== 'xlsxChart') throw new Error('not xlsxChart')
    const chart2a = drawing2a.content.chart
    expect(chart2a.groups[0].gapWidth).toBe(247)

    const cov2a = doc2.sheets[0]!.drawingCoverage?.find(c => c.feature === 'ChartEx' || c.feature === 'chart')
    expect(cov2a).toMatchObject({ status: 'native', selectedRepresentation: 'chart' })

    const canvas2a = createCanvas(400, 300)
    const ctx2a = canvas2a.getContext('2d')
    let rects2a = 0, card2a = false, w2a = 0
    const origFillRect2a = ctx2a.fillRect.bind(ctx2a)
    ctx2a.fillRect = (x, y, w, h) => {
      if (String(ctx2a.fillStyle).toLowerCase() === '#f7f7f7' && w === 400 && h === 300) card2a = true
      if (w > 2 && h > 1 && y > 0 && y + h < 300) { rects2a++; w2a = w }
      origFillRect2a(x, y, w, h)
    }
    paintXlsxChart(chart2a, ctx2a as any, 400, 300)
    expect(card2a).toBe(false)
    expect(rects2a).toBe(3) // 3 boxes (Table1 has 3 series: B, C, D)

    // 3. BW2b: box-whisker-gapwidth.xlsx chart 2 (Table2 on Sheet1: D5:D13, gapWidth = 0.06)
    const drawing2b = doc2.sheets[1]!.drawings![0]!
    expect(drawing2b.content?.kind).toBe('xlsxChart')
    if (drawing2b.content?.kind !== 'xlsxChart') throw new Error('not xlsxChart')
    const chart2b = drawing2b.content.chart
    expect(chart2b.unsupported).toBeUndefined()
    expect(chart2b.groups[0].gapWidth).toBe(6)

    const canvas2b = createCanvas(400, 300)
    const ctx2b = canvas2b.getContext('2d')
    let rects2b = 0, card2b = false, w2b = 0
    const origFillRect2b = ctx2b.fillRect.bind(ctx2b)
    ctx2b.fillRect = (x, y, w, h) => {
      if (String(ctx2b.fillStyle).toLowerCase() === '#f7f7f7' && w === 400 && h === 300) card2b = true
      if (w > 2 && h > 1 && y > 0 && y + h < 300) { rects2b++; w2b = w }
      origFillRect2b(x, y, w, h)
    }
    paintXlsxChart(chart2b, ctx2b as any, 400, 300)
    expect(card2b).toBe(false)
    expect(rects2b).toBe(1) // 1 box for D5:D13

    // Gap width check (M5): BW2b gapWidth (0.06) produces a significantly wider box than BW2a gapWidth (2.47)
    expect(w2b).toBeGreaterThan(w2a * 2)

    // 4. Empty series: draws no box and remains supported
    const emptyXml = `
      <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="0"/></cx:numDim></cx:data></cx:chartData>
        <cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="boxWhisker"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion>
        <cx:axis id="0"><cx:catScaling gapWidth="1"/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis></cx:plotArea></cx:chart>
      </cx:chartSpace>
    `
    const emptyModel = await nativeModel(emptyXml)
    expect(emptyModel.unsupported).toBeUndefined()
    const canvasEmpty = createCanvas(400, 300)
    const ctxEmpty = canvasEmpty.getContext('2d')
    let emptyRects = 0, emptyCard = false
    const origEmpty = ctxEmpty.fillRect.bind(ctxEmpty)
    ctxEmpty.fillRect = (x, y, w, h) => {
      if (String(ctxEmpty.fillStyle).toLowerCase() === '#f7f7f7' && w === 400 && h === 300) emptyCard = true
      if (w > 2 && h > 1 && y > 0 && y + h < 300) emptyRects++
      origEmpty(x, y, w, h)
    }
    paintXlsxChart(emptyModel, ctxEmpty as any, 400, 300)
    expect(emptyCard).toBe(false)
    expect(emptyRects).toBe(0) // No boxes drawn for empty series
  })

  it('groups boxes by category when cat dimension is present; 2 categories x 2 series side-by-side', async () => {
    // s0: A [10, 11, 12, 13] B [20, 21, 22, 23]; s1: A [30, 31, 32, 33] B [40, 41, 42, 43] (n = 4: Q1/Q3 = halves' medians).
    // A s0: Q1 10.5, Q3 12.5; A s1: 30.5 / 32.5; B s0: 20.5 / 22.5; B s1: 40.5 / 42.5.
    // Clustered geometry (2 slots, between, gapWidth 100%, 2 series): band = plot.w / 2, box w = band / 3 = plot.w / 6,
    // start(k, s) = plot.x + k*band + band/6 + s*band/3 = plot.x + plot.w * {1/12, 1/4, 7/12, 3/4}.
    // y(v) = plot.y + plot.h * (60 - v) / 60; a box spans y(Q3)..y(Q1), h = plot.h * 2 / 60.
    const cats = ['Cat A', 'Cat A', 'Cat A', 'Cat A', 'Cat B', 'Cat B', 'Cat B', 'Cat B']
    const model = await nativeModel(bwXml([
      { values: [10, 11, 12, 13, 20, 21, 22, 23], cats },
      { values: [30, 31, 32, 33, 40, 41, 42, 43], cats },
    ], ' min="0" max="60"'))
    expect(model.unsupported).toBeUndefined()
    const layout = plotOf(model)
    const { plot } = layout
    const y = (v: number) => plot.y + plot.h * (60 - v) / 60
    expect(model.groups[0].gapWidth ?? 100).toBe(100)
    expect(model.preparedAxes?.axes.find(a => a.role === 'category')?.labels).toEqual(['Cat A', 'Cat B'])

    const boxes = record(model).rects.sort((p, q) => p.x - q.x)
    const want = [
      { f: 1 / 12, q3: 12.5 }, // Cat A, series 0
      { f: 1 / 4, q3: 32.5 }, // Cat A, series 1
      { f: 7 / 12, q3: 22.5 }, // Cat B, series 0
      { f: 3 / 4, q3: 42.5 }, // Cat B, series 1
    ]
    expect(boxes).toHaveLength(4)
    for (const [i, w] of want.entries()) {
      expect(boxes[i].x).toBeCloseTo(plot.x + plot.w * w.f, 6)
      expect(boxes[i].w).toBeCloseTo(plot.w / 6, 6)
      expect(boxes[i].y).toBeCloseTo(y(w.q3), 6)
      expect(boxes[i].h).toBeCloseTo(plot.h * 2 / 60, 6)
    }
  })

  it('refuses analytic input over 65,536 limit (never using a prefix)', async () => {
    // 65,537 points in a series
    const declared = CHART_ANALYTIC_POINT_LIMIT + 1
    const xml = `
      <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <cx:chartData>
          <cx:data id="0">
            <cx:numDim type="val">
              <cx:lvl ptCount="${declared}" formatCode="General">
                <cx:pt idx="0">10</cx:pt>
              </cx:lvl>
            </cx:numDim>
          </cx:data>
        </cx:chartData>
        <cx:chart>
          <cx:plotArea>
            <cx:plotAreaRegion>
              <cx:series layoutId="boxWhisker"><cx:dataId val="0"/></cx:series>
            </cx:plotAreaRegion>
            <cx:axis id="0"><cx:catScaling gapWidth="1"/><cx:tickLabels/></cx:axis>
            <cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis>
          </cx:plotArea>
        </cx:chart>
      </cx:chartSpace>
    `
    const model = parseChartEx(parseXml(xml), undefined, await testContext())
    expect(model.unsupported).toBe(CHART_TOO_MANY_POINTS)
    expect(model.refused).toBe(true)

    // Adapter returns kind: 'refused' (the part is read by the production reader: real namespace dispatch)
    const result = await adapt(xml, 'xl/charts/chartEx1.xml')
    expect(result).toMatchObject({ kind: 'refused', reason: CHART_TOO_MANY_POINTS })
  })

  it('DOCX and PPTX parts are prepared once in stage 1; paint never computes statistics (N1)', async () => {
    // [5, 1, 4, 2, 3] -> sorted 1..5, n = 5 (odd), exclusive: median 3, Q1 = median[1, 2] = 1.5, Q3 = median[4, 5] = 4.5,
    // IQR 3, fences -3 / 9 -> whiskers 1 / 5, no outliers, mean 3.
    for (const path of ['word/charts/chart1.xml', 'ppt/charts/chart1.xml']) {
      const model = await nativeModel(bwXml([{ values: [5, 1, 4, 2, 3] }], ' min="0" max="6"'), path)
      const prepared = model.preparedAxes?.boxWhisker
      expect(prepared, path).toHaveLength(1)
      expect(prepared![0].boxes).toHaveLength(1)
      expect(prepared![0].boxes[0].stats).toMatchObject({ n: 5, q1: 1.5, median: 3, q3: 4.5, iqr: 3, lowerWhisker: 1, upperWhisker: 5, mean: 3, outliers: [] })
      expect(record(model).rects, path).toHaveLength(1)
      // Without stage-1 records the painter draws no box (and no card): nothing is computed per frame.
      const unprepared = record({ ...model, preparedAxes: undefined })
      expect(unprepared.rects, path).toHaveLength(0)
      expect(unprepared.card, path).toBe(false)
    }
  })

  it('statistics scratch and retained point lists are charged through the workspace before allocation (I4)', async () => {
    // One series 1..1000 with inner points on: the value buffer (1000 x 8 B) is scratch, the inner-point list
    // (1000 x 8 B) and the box record (160 B) are retained; there are no outliers (fences far outside 1..1000).
    const n = 1000
    const values = Array.from({ length: n }, (_, i) => i + 1)
    const raw = await rawModel(bwXml([{ values, vis: 'meanLine="0" meanMarker="1" nonoutliers="1" outliers="1"' }]))
    expect(raw.preparedAxes).toBeUndefined()

    const workspace = createChartWorkspace('desktop')
    const reserves: Array<{ bytes: number; tag: string }> = [], transfers: Array<{ to: string; bytes: number }> = []
    const reserve = workspace.reserve.bind(workspace), transfer = workspace.transfer.bind(workspace)
    workspace.reserve = (bytes, tag) => { reserves.push({ bytes, tag }); return reserve(bytes, tag) }
    workspace.transfer = (from, to, bytes) => { transfers.push({ to, bytes }); transfer(from, to, bytes) }
    const prepared = prepareChartAxes(raw, { workspace, tag: 'model' })
    const scratch = reserves.filter(r => r.tag.endsWith(':box-scratch'))
    expect(Math.max(0, ...scratch.map(r => r.bytes))).toBeGreaterThanOrEqual(n * 8)
    expect(prepared.preparedAxes?.boxWhisker?.[0].boxes[0].stats.inliers).toHaveLength(n)
    const retained = transfers.find(t => t.to.endsWith(':retained'))!
    expect(retained.bytes).toBeGreaterThanOrEqual(160 + n * 8)
    // Scratch is refunded: only the retained record stays charged.
    expect(workspace.used).toBe(retained.bytes)

    // A workspace that refuses the scratch refuses the preparation (and refunds everything).
    const tight = createChartWorkspace('desktop')
    const tightReserve = tight.reserve.bind(tight)
    tight.reserve = (bytes, tag) => !tag.endsWith(':box-scratch') && tightReserve(bytes, tag)
    expect(() => prepareChartAxes(raw, { workspace: tight, tag: 'model' })).toThrow(ChartRefusal)
    expect(tight.used).toBe(0)
  })

  it('category grouping is linear per series and more than 4,096 distinct categories refuse (N3)', async () => {
    // 1. One series, 400 points over 40 categories. A per-category rescan reads 40 x 400 = 16,000 category slots in the
    // grouping alone; bucketing reads each slot a small constant number of times (distinct passes + two bucket passes).
    const n = 400
    const raw = await rawModel(bwXml([{ values: Array.from({ length: n }, (_, j) => j), cats: Array.from({ length: n }, (_, j) => `c${j % 40}`) }]))
    let reads = 0
    const counted = new Proxy(raw.categories as string[], { get(target, key, receiver) { if (typeof key === 'string' && /^\d+$/.test(key)) reads++; return Reflect.get(target, key, receiver) } })
    const group = raw.groups[0]
    const model: ChartModel = { ...raw, categories: counted, groups: [{ ...group, series: group.series.map(s => ({ ...s, categories: counted })) }] }
    const prepared = prepareChartAxes(model)
    expect(prepared.preparedAxes?.boxWhisker?.[0].boxes).toHaveLength(40)
    // Box k holds j = k, k + 40, ...: 10 points k..k+360 -> median k + 180.
    expect(prepared.preparedAxes!.boxWhisker![0].boxes.map(b => b.stats.median)).toEqual(Array.from({ length: 40 }, (_, k) => k + 180))
    expect(reads).toBeLessThanOrEqual(8 * n)

    // 2. 4,097 distinct categories (one point each) refuse the chart (never a prefix); 4,096 are drawn natively.
    const many = (k: number) => bwXml([{ values: Array.from({ length: k }, (_, j) => j), cats: Array.from({ length: k }, (_, j) => `k${j}`) }])
    const over = prepareChartAxes(await rawModel(many(CHART_POINT_LIMIT + 1)))
    expect(over).toMatchObject({ unsupported: CHART_TOO_MANY_POINTS, refused: true })
    expect(await adapt(many(CHART_POINT_LIMIT + 1))).toMatchObject({ kind: 'refused', reason: CHART_TOO_MANY_POINTS })
    const at = await nativeModel(many(CHART_POINT_LIMIT))
    expect(at.preparedAxes?.boxWhisker?.[0].boxes).toHaveLength(CHART_POINT_LIMIT)
    expect(at.preparedAxes?.axes.find(a => a.role === 'category')?.slots).toBe(CHART_POINT_LIMIT)
  })

  it('hidden series are skipped: no box, no slot, no label, no share of the value scale (M2)', async () => {
    // Alpha [10, 20, 30], Beta hidden [1000, 1001, 1002], Gamma [40, 50, 60]. Two visible series, no categories:
    // 2 slots, box w = plot.w / 4 at plot.x + plot.w * {1/8, 5/8}.
    const model = await nativeModel(bwXml([
      { name: 'Alpha', values: [10, 20, 30] },
      { name: 'Beta', values: [1000, 1001, 1002], hidden: true },
      { name: 'Gamma', values: [40, 50, 60] },
    ]))
    const prepared = model.preparedAxes!
    expect(prepared.boxWhisker?.[0].boxes.map(b => b.seriesIndex)).toEqual([0, 2])
    expect(prepared.axes.find(a => a.role === 'category')?.labels).toEqual(['Alpha', 'Gamma'])
    expect(prepared.axes.find(a => a.role === 'value')!.max).toBeLessThan(1000)
    const { plot } = plotOf(model)
    const rects = record(model).rects.sort((p, q) => p.x - q.x)
    expect(rects).toHaveLength(2)
    expect(rects[0].x).toBeCloseTo(plot.x + plot.w / 8, 6)
    expect(rects[1].x).toBeCloseTo(plot.x + plot.w * 5 / 8, 6)
    expect(rects[0].w).toBeCloseTo(plot.w / 4, 6)
  })

  it('empty categories form their own group; points without a category are diagnosed, never silently dropped', async () => {
    // 1. Categories A, '', A, '' over 1, 2, 3, 4: A = [1, 3] (median 2), '' = [2, 4] (median 3).
    const mixed = await nativeModel(bwXml([{ values: [1, 2, 3, 4], cats: ['A', null, 'A', null] }]))
    expect(mixed.preparedAxes?.boxWhisker?.[0].boxes.map(b => [b.slotIndex, b.stats.median])).toEqual([[0, 2], [1, 3]])
    expect(mixed.preparedAxes?.axes.find(a => a.role === 'category')?.labels).toEqual(['A', ''])
    // 2. Every category empty: no category dimension, one box per series over all four points (median 2.5).
    const blank = await nativeModel(bwXml([{ values: [1, 2, 3, 4], cats: [null, null, null, null] }]))
    expect(blank.preparedAxes?.boxWhisker?.[0].hasCategories).toBe(false)
    expect(blank.preparedAxes?.boxWhisker?.[0].boxes.map(b => b.stats.median)).toEqual([2.5])
    // 3. A category dimension shorter than the values: points 2 and 3 have no category; they are diagnosed.
    const short = await nativeModel(bwXml([{ values: [1, 2, 3, 4], cats: ['A', 'B'] }]))
    expect(short.preparedAxes?.boxWhisker?.[0].boxes.map(b => b.stats.median)).toEqual([1, 2])
    expect(short.preparedAxes?.diagnostics).toContainEqual({ code: 'chart-point-uncategorized', reason: '2 box & whisker points have no category and are not drawn' })
  })

  describeFamilyConformance(BOX_WHISKER_FAMILY)
})

describe('C1 phase 2b: the median line and mean marker stay visible on a filled box', () => {
  const pixel = (ctx: { getImageData(x: number, y: number, w: number, h: number): { data: Uint8ClampedArray } }, x: number, y: number) => [...ctx.getImageData(Math.round(x), Math.round(y), 1, 1).data].slice(0, 3).join(',')
  /** Paints on a real canvas; returns it, the series-coloured box rect, the median segment and the mean cross centre. */
  async function paintBox(xml: string) {
    const model = await nativeModel(xml)
    const w = 400, h = 300, ctx = createCanvas(w, h).getContext('2d')
    const rects: Array<{ x: number; y: number; w: number; h: number; fill: string }> = [], segments: Segment[] = []
    const fillRect = ctx.fillRect.bind(ctx), moveTo = ctx.moveTo.bind(ctx), lineTo = ctx.lineTo.bind(ctx)
    ctx.fillRect = (x, y, rw, rh) => { if (!(rw === w && rh === h)) rects.push({ x, y, w: rw, h: rh, fill: String(ctx.fillStyle).toLowerCase() }); fillRect(x, y, rw, rh) }
    let lx = 0, ly = 0
    ctx.moveTo = (x, y) => { lx = x; ly = y; moveTo(x, y) }
    ctx.lineTo = (x, y) => { segments.push({ x1: lx, y1: ly, x2: x, y2: y, stroke: String(ctx.strokeStyle).toLowerCase() }); lx = x; ly = y; lineTo(x, y) }
    paintXlsxChart(model, ctx as unknown as CanvasRenderingContext2D, w, h)
    const color = `#${model.groups[0].series[0].color.toLowerCase()}`
    const box = rects.filter(r => r.h > 10 && r.w > 10).sort((a, b) => a.w * a.h - b.w * b.h)[0]
    const median = segments.find(s => Math.abs(s.y1 - s.y2) < 1e-9 && s.y1 > box.y + 1 && s.y1 < box.y + box.h - 1 && Math.abs(s.x1 - box.x) < 1e-6 && Math.abs(s.x2 - (box.x + box.w)) < 1e-6)!
    const cross = segments.filter(s => Math.abs(Math.abs(s.x2 - s.x1) - 8) < 1e-6 && Math.abs(Math.abs(s.y2 - s.y1) - 8) < 1e-6)
    return { ctx, box, median, cross, color, model }
  }
  // Nine values, exclusive quartiles: median 5, mean 6.2 (both inside the box), 20 an outlier.
  const values = [1, 2, 3, 4, 5, 6, 7, 8, 20]

  it('default fill (the series colour): median and mean are drawn in white and differ from the box pixels', async () => {
    const { ctx, box, median, cross, color } = await paintBox(bwXml([{ values }]))
    expect(box.fill).toBe(color)
    expect(median, 'the median segment spans the box').toBeDefined()
    expect(median.stroke).toBe('#ffffff')
    expect(cross).toHaveLength(2)
    for (const s of cross) expect(s.stroke).toBe('#ffffff')
    // Pixels: the median line (1 px, at its y) and the cross centre differ from the box fill a few pixels away.
    const cx = box.x + box.w / 2
    const fillPixel = pixel(ctx, cx, box.y + 2)
    // A 1 px line at y covers the pixel rows around y (both, antialiased, when y is an integer): one of them differs.
    const rows = [Math.floor(median.y1 - 0.5), Math.floor(median.y1)].map(y => pixel(ctx, box.x + box.w / 4, y))
    expect(rows.some(p => p !== fillPixel), `median rows ${rows} vs fill ${fillPixel}`).toBe(true)
    const centre = { x: (cross[0].x1 + cross[0].x2) / 2, y: (cross[0].y1 + cross[0].y2) / 2 }
    expect(pixel(ctx, centre.x, centre.y)).not.toBe(fillPixel)
  })

  it('an explicit solid fill is the series colour too (white marks); an unfilled box keeps the series colour', async () => {
    // A ChartEx series spPr solid fill also sets the series colour, so the marks contrast with it in white.
    const solid = await paintBox(bwXml([{ values }]).replace('<cx:dataId val="0"/>', '<cx:spPr><a:solidFill><a:srgbClr val="FFC000"/></a:solidFill></cx:spPr><cx:dataId val="0"/>'))
    expect(solid.color).toBe('#ffc000')
    expect(solid.median.stroke).toBe('#ffffff')
    // noFill: nothing under the marks, so they keep the series colour.
    const model = await nativeModel(bwXml([{ values }]).replace('<cx:dataId val="0"/>', '<cx:spPr><a:noFill/></cx:spPr><cx:dataId val="0"/>'))
    const color = `#${model.groups[0].series[0].color.toLowerCase()}`
    const { segments } = record(model)
    const horizontal = segments.filter(seg => Math.abs(seg.y1 - seg.y2) < 1e-9 && seg.x2 - seg.x1 > 10)
    expect(horizontal.length).toBeGreaterThan(0)
    expect(horizontal.some(seg => seg.stroke === color)).toBe(true)
    expect(horizontal.some(seg => seg.stroke === '#ffffff')).toBe(false)
  })

  it('review: skewed mean stays visible outside its box', async () => {
    // [1..8, 100]: exclusive quartiles put the box at about 2.5..8, the mean (15) far above it, on the plot paper.
    const { ctx, box, median, cross, color } = await paintBox(bwXml([{ values: [1, 2, 3, 4, 5, 6, 7, 8, 100] }]))
    expect(median.stroke).toBe('#ffffff')
    expect(cross).toHaveLength(2)
    const centre = { x: (cross[0].x1 + cross[0].x2) / 2, y: (cross[0].y1 + cross[0].y2) / 2 }
    expect(centre.y < box.y || centre.y > box.y + box.h, 'the mean lies outside the box').toBe(true)
    for (const s of cross) expect(s.stroke).toBe(color)
    // An 11 x 11 sample enclosing the whole cross holds non-white (marker) pixels.
    const data = ctx.getImageData(Math.round(centre.x) - 5, Math.round(centre.y) - 5, 11, 11).data
    let ink = 0
    for (let i = 0; i < data.length; i += 4) if (data[i] < 200 || data[i + 1] < 200 || data[i + 2] < 200) ink++
    expect(ink).toBeGreaterThan(5)
  })

  it('a mean inside its filled box is drawn in white and differs from the fill', async () => {
    const { ctx, box, cross } = await paintBox(bwXml([{ values }]))
    const centre = { x: (cross[0].x1 + cross[0].x2) / 2, y: (cross[0].y1 + cross[0].y2) / 2 }
    expect(centre.y > box.y && centre.y < box.y + box.h).toBe(true)
    for (const s of cross) expect(s.stroke).toBe('#ffffff')
    expect(pixel(ctx, centre.x, centre.y)).not.toBe(pixel(ctx, box.x + box.w / 2, box.y + 2))
  })

  it('boxMarkColor: white only when the fill is the series colour', () => {
    expect(boxMarkColor(undefined, '4472C4')).toBe('FFFFFF')
    expect(boxMarkColor({ kind: 'solid', color: '4472c4' } as never, '4472C4')).toBe('FFFFFF')
    expect(boxMarkColor({ kind: 'solid', color: 'FFC000' } as never, '4472C4')).toBe('4472C4')
    expect(boxMarkColor({ kind: 'none' } as never, '4472C4')).toBe('4472C4')
    expect(boxMarkColor({ kind: 'gradient', stops: [] } as never, '4472C4')).toBe('4472C4')
  })
})
