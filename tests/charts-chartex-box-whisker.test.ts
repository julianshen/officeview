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
import { computeBoxWhiskerStats } from '../src/charts/chartex/boxWhisker'
import type { ChartModel } from '../src/charts/model'
import { parseChartEx } from '../src/charts/chartex/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { parseXml, parseXmlOrdered } from '../src/core/xml'
import { chartParseContextFor } from '../src/charts/context'
import { OfficePackage } from '../src/core/zip'
import JSZip from 'jszip'

import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { parseXlsx } from '../src/xlsx/parse'
import { parseChartPart } from '../src/charts/dataPrecedence'
import { markChartExRoot } from '../src/charts/read'
import { BOX_WHISKER_FAMILY, describeFamilyConformance } from './helpers/chart-family-conformance'
import { CHART_ANALYTIC_POINT_LIMIT } from '../src/charts/parse'
import { CHART_TOO_MANY_POINTS } from '../src/charts/fill'

const FIXTURES = join(__dirname, 'fixtures', 'chartex')
const fixture = (name: string) => new Uint8Array(readFileSync(join(FIXTURES, name)))
const pkgPromise = new JSZip().generateAsync({ type: 'uint8array' }).then(b => OfficePackage.load(b))
const testContext = async () => chartParseContextFor(await pkgPromise, 'xl/charts/chartEx1.xml')

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
    const mWith = parseChartEx(parseXml(makeXml('1')), undefined, await testContext())
    expect(countDots(mWith)).toBe(2)

    // 2. With outliers="0": 0 outlier dots drawn
    const mWithout = parseChartEx(parseXml(makeXml('0')), undefined, await testContext())
    expect(countDots(mWithout)).toBe(0)
  })

  it('mean marker, mean line and inner points', async () => {
    // 2 series:
    // Series 0: [10, 20, 30] -> mean = 20, inliers = [10, 20, 30], outliers = []
    // Series 1: [40, 50, 60] -> mean = 50, inliers = [40, 50, 60], outliers = []
    const makeXml = (meanMarker: '0' | '1', meanLine: '0' | '1', nonoutliers: '0' | '1') => `
      <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <cx:chartData>
          <cx:data id="0">
            <cx:numDim type="val">
              <cx:lvl ptCount="3"><cx:pt idx="0">10</cx:pt><cx:pt idx="1">20</cx:pt><cx:pt idx="2">30</cx:pt></cx:lvl>
            </cx:numDim>
          </cx:data>
          <cx:data id="1">
            <cx:numDim type="val">
              <cx:lvl ptCount="3"><cx:pt idx="0">40</cx:pt><cx:pt idx="1">50</cx:pt><cx:pt idx="2">60</cx:pt></cx:lvl>
            </cx:numDim>
          </cx:data>
        </cx:chartData>
        <cx:chart>
          <cx:plotArea>
            <cx:plotAreaRegion>
              <cx:series layoutId="boxWhisker">
                <cx:dataId val="0"/>
                <cx:layoutPr>
                  <cx:visibility meanLine="${meanLine}" meanMarker="${meanMarker}" nonoutliers="${nonoutliers}" outliers="1"/>
                  <cx:statistics quartileMethod="exclusive"/>
                </cx:layoutPr>
              </cx:series>
              <cx:series layoutId="boxWhisker">
                <cx:dataId val="1"/>
                <cx:layoutPr>
                  <cx:visibility meanLine="${meanLine}" meanMarker="${meanMarker}" nonoutliers="${nonoutliers}" outliers="1"/>
                  <cx:statistics quartileMethod="exclusive"/>
                </cx:layoutPr>
              </cx:series>
            </cx:plotAreaRegion>
            <cx:axis id="0"><cx:catScaling gapWidth="1"/><cx:tickLabels/></cx:axis>
            <cx:axis id="1"><cx:valScaling min="0" max="70"/><cx:majorGridlines/><cx:tickLabels/></cx:axis>
          </cx:plotArea>
        </cx:chart>
      </cx:chartSpace>
    `

    const probe = async (xml: string) => {
      const model = parseChartEx(parseXml(xml), undefined, await testContext())
      const canvas = createCanvas(400, 300)
      const ctx = canvas.getContext('2d')
      let dots = 0 // circle markers
      let diagonalSegments = 0
      let connectedMeanLines = 0
      const segments: Array<{ x1: number; y1: number; x2: number; y2: number }> = []

      const origArc = ctx.arc.bind(ctx)
      ctx.arc = (x, y, r, sa, ea) => {
        if (r <= 4) dots++
        origArc(x, y, r, sa, ea)
      }

      let lx = 0, ly = 0
      ctx.moveTo = (x, y) => { lx = x; ly = y }
      ctx.lineTo = (x, y) => {
        const dx = Math.abs(x - lx), dy = Math.abs(y - ly)
        segments.push({ x1: lx, y1: ly, x2: x, y2: y })
        // Mean marker has two diagonal strokes of radius r (~4px, so dx~8, dy~8)
        if (dx > 2 && dx < 12 && dy > 2 && dy < 12) diagonalSegments++
        // Mean line spans from box 0 center to box 1 center (dx > 50, dy > 20)
        if (dx > 50 && dy > 10) connectedMeanLines++
        lx = x; ly = y
      }

      paintXlsxChart(model, ctx as any, 400, 300)
      return { dots, diagonalSegments, connectedMeanLines, segments }
    }

    // 1. Base: meanMarker="1", meanLine="0", nonoutliers="0"
    const r1 = await probe(makeXml('1', '0', '0'))
    // 2 series -> 2 mean markers (each has 2 diagonal strokes -> 4 diagonal segments)
    expect(r1.diagonalSegments).toBe(4)
    expect(r1.connectedMeanLines).toBe(0)
    expect(r1.dots).toBe(0) // No inner points

    // 2. meanMarker="0", meanLine="1", nonoutliers="0"
    const r2 = await probe(makeXml('0', '1', '0'))
    expect(r2.diagonalSegments).toBe(0)
    expect(r2.connectedMeanLines).toBe(1) // 1 line connecting the 2 means
    expect(r2.dots).toBe(0)
    const meanLineSegment = r2.segments.find(s => Math.abs(s.x2 - s.x1) > 50 && Math.abs(s.y2 - s.y1) > 10)
    expect(meanLineSegment).toBeDefined()
    // Series 1 mean is 20, Series 2 mean is 40.
    // Higher value has lower canvas y, so y1 (mean 20) is strictly greater than y2 (mean 40).
    expect(meanLineSegment!.y1).toBeGreaterThan(meanLineSegment!.y2)

    // 3. meanMarker="0", meanLine="0", nonoutliers="1"
    const r3 = await probe(makeXml('0', '0', '1'))
    expect(r3.diagonalSegments).toBe(0)
    expect(r3.connectedMeanLines).toBe(0)
    // 3 inliers in series 0 + 3 inliers in series 1 = 6 inner dots
    expect(r3.dots).toBe(6)
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
    const emptyModel = parseChartEx(parseXml(emptyXml), undefined, await testContext())
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
    const xml = `
      <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <cx:chartData>
          <cx:data id="0">
            <cx:strDim type="cat">
              <cx:lvl ptCount="8">
                <cx:pt idx="0">Cat A</cx:pt>
                <cx:pt idx="1">Cat A</cx:pt>
                <cx:pt idx="2">Cat A</cx:pt>
                <cx:pt idx="3">Cat A</cx:pt>
                <cx:pt idx="4">Cat B</cx:pt>
                <cx:pt idx="5">Cat B</cx:pt>
                <cx:pt idx="6">Cat B</cx:pt>
                <cx:pt idx="7">Cat B</cx:pt>
              </cx:lvl>
            </cx:strDim>
            <cx:numDim type="val">
              <cx:lvl ptCount="8">
                <cx:pt idx="0">10</cx:pt>
                <cx:pt idx="1">11</cx:pt>
                <cx:pt idx="2">12</cx:pt>
                <cx:pt idx="3">13</cx:pt>
                <cx:pt idx="4">20</cx:pt>
                <cx:pt idx="5">21</cx:pt>
                <cx:pt idx="6">22</cx:pt>
                <cx:pt idx="7">23</cx:pt>
              </cx:lvl>
            </cx:numDim>
          </cx:data>
          <cx:data id="1">
            <cx:strDim type="cat">
              <cx:lvl ptCount="8">
                <cx:pt idx="0">Cat A</cx:pt>
                <cx:pt idx="1">Cat A</cx:pt>
                <cx:pt idx="2">Cat A</cx:pt>
                <cx:pt idx="3">Cat A</cx:pt>
                <cx:pt idx="4">Cat B</cx:pt>
                <cx:pt idx="5">Cat B</cx:pt>
                <cx:pt idx="6">Cat B</cx:pt>
                <cx:pt idx="7">Cat B</cx:pt>
              </cx:lvl>
            </cx:strDim>
            <cx:numDim type="val">
              <cx:lvl ptCount="8">
                <cx:pt idx="0">30</cx:pt>
                <cx:pt idx="1">31</cx:pt>
                <cx:pt idx="2">32</cx:pt>
                <cx:pt idx="3">33</cx:pt>
                <cx:pt idx="4">40</cx:pt>
                <cx:pt idx="5">41</cx:pt>
                <cx:pt idx="6">42</cx:pt>
                <cx:pt idx="7">43</cx:pt>
              </cx:lvl>
            </cx:numDim>
          </cx:data>
        </cx:chartData>
        <cx:chart>
          <cx:plotArea>
            <cx:plotAreaRegion>
              <cx:series layoutId="boxWhisker"><cx:dataId val="0"/></cx:series>
              <cx:series layoutId="boxWhisker"><cx:dataId val="1"/></cx:series>
            </cx:plotAreaRegion>
            <cx:axis id="0"><cx:catScaling gapWidth="1"/><cx:tickLabels/></cx:axis>
            <cx:axis id="1"><cx:valScaling min="0" max="60"/><cx:majorGridlines/><cx:tickLabels/></cx:axis>
          </cx:plotArea>
        </cx:chart>
      </cx:chartSpace>
    `
    const model = parseChartEx(parseXml(xml), undefined, await testContext())
    expect(model.unsupported).toBeUndefined()

    const boxes: Array<{ x: number; y: number; w: number; h: number }> = []
    const canvas = createCanvas(400, 300)
    const ctx = canvas.getContext('2d')
    const origFillRect = ctx.fillRect.bind(ctx)
    ctx.fillRect = (x, y, w, h) => {
      if (w > 2 && h > 1 && y > 0 && y + h < 300 && !(w === 400 && h === 300)) {
        boxes.push({ x, y, w, h })
      }
      origFillRect(x, y, w, h)
    }

    paintXlsxChart(model, ctx as any, 400, 300)

    // Expect 4 boxes total: Cat A (Series 1, Series 2) and Cat B (Series 1, Series 2)
    expect(boxes.length).toBe(4)

    // Sort by x
    boxes.sort((a, b) => a.x - b.x)

    // All boxes should have the same width
    expect(boxes[0].w).toBe(boxes[1].w)
    expect(boxes[1].w).toBe(boxes[2].w)
    expect(boxes[2].w).toBe(boxes[3].w)

    // Within Cat A: Series 1 and Series 2 sit side-by-side
    expect(boxes[1].x).toBeCloseTo(boxes[0].x + boxes[0].w, 1)

    // Cat B is strictly to the right of Cat A with gap
    expect(boxes[2].x).toBeGreaterThan(boxes[1].x + boxes[1].w)

    // Within Cat B: Series 1 and Series 2 sit side-by-side
    expect(boxes[3].x).toBeCloseTo(boxes[2].x + boxes[2].w, 1)
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

    // Adapter returns kind: 'refused'
    const orderedXml = parseXmlOrdered(xml)
    markChartExRoot(orderedXml)
    const result = await parseChartPart(orderedXml, undefined, await testContext())
    expect(result).toMatchObject({ kind: 'refused', reason: CHART_TOO_MANY_POINTS })
  })

  describeFamilyConformance(BOX_WHISKER_FAMILY)
})
