// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { PointWriter } from '../src/charts/points'
import { catmullRomToBezier, createLineTracer, traceSmoothCurve, type Point2D } from '../src/charts/smooth'

describe('charts-smooth (Task A4c)', () => {
  it('smooth passes through every point with Catmull-Rom tangents; end segments use the end point', () => {
    // 4 points with uniform unit x spacing
    const points: Point2D[] = [
      { x: 0, y: 10 },
      { x: 1, y: 20 },
      { x: 2, y: 15 },
      { x: 3, y: 30 },
    ]

    const segments = catmullRomToBezier(points)
    expect(segments).toHaveLength(3)

    // Segment 0: P0 to P1
    // Tangents: T0 = 0.5 * (P1 - P0) = (0.5, 5)
    // T1 = 0.5 * (P2 - P0) = (1, 2.5)
    // Control points: cp1 = P0 + T0 / 3 = (0 + 1/6, 10 + 5/3) = (1/6, 35/3)
    // cp2 = P1 - T1 / 3 = (1 - 1/3, 20 - 2.5/3) = (2/3, 115/6)
    expect(segments[0].p0).toEqual({ x: 0, y: 10 })
    expect(segments[0].p1).toEqual({ x: 1, y: 20 })
    expect(segments[0].cp1.x).toBeCloseTo(1 / 6, 5)
    expect(segments[0].cp1.y).toBeCloseTo(35 / 3, 5)
    expect(segments[0].cp2.x).toBeCloseTo(2 / 3, 5)
    expect(segments[0].cp2.y).toBeCloseTo(115 / 6, 5)

    // Segment 1: P1 to P2
    // T1 = 0.5 * (P2 - P0) = (1, 2.5)
    // T2 = 0.5 * (P3 - P1) = (1, 5)
    // cp1 = P1 + T1 / 3 = (1 + 1/3, 20 + 2.5/3) = (4/3, 125/6)
    // cp2 = P2 - T2 / 3 = (2 - 1/3, 15 - 5/3) = (5/3, 40/3)
    expect(segments[1].p0).toEqual({ x: 1, y: 20 })
    expect(segments[1].p1).toEqual({ x: 2, y: 15 })
    expect(segments[1].cp1.x).toBeCloseTo(4 / 3, 5)
    expect(segments[1].cp1.y).toBeCloseTo(125 / 6, 5)
    expect(segments[1].cp2.x).toBeCloseTo(5 / 3, 5)
    expect(segments[1].cp2.y).toBeCloseTo(40 / 3, 5)

    // Segment 2: P2 to P3
    // T2 = 0.5 * (P3 - P1) = (1, 5)
    // T3 = 0.5 * (P3 - P2) = (0.5, 7.5) [end segment uses P3 as its own neighbour]
    // cp1 = P2 + T2 / 3 = (2 + 1/3, 15 + 5/3) = (7/3, 50/3)
    // cp2 = P3 - T3 / 3 = (3 - 1/6, 30 - 2.5) = (17/6, 27.5)
    expect(segments[2].p0).toEqual({ x: 2, y: 15 })
    expect(segments[2].p1).toEqual({ x: 3, y: 30 })
    expect(segments[2].cp1.x).toBeCloseTo(7 / 3, 5)
    expect(segments[2].cp1.y).toBeCloseTo(50 / 3, 5)
    expect(segments[2].cp2.x).toBeCloseTo(17 / 6, 5)
    expect(segments[2].cp2.y).toBeCloseTo(27.5, 5)

    // Verify C^1 continuity at P1 and P2: arrival tangent = departure tangent
    const incomingT1 = {
      x: 3 * (segments[0].p1.x - segments[0].cp2.x),
      y: 3 * (segments[0].p1.y - segments[0].cp2.y),
    }
    const outgoingT1 = {
      x: 3 * (segments[1].cp1.x - segments[1].p0.x),
      y: 3 * (segments[1].cp1.y - segments[1].p0.y),
    }
    expect(incomingT1.x).toBeCloseTo(outgoingT1.x, 5)
    expect(incomingT1.y).toBeCloseTo(outgoingT1.y, 5)

    // Verify traceSmoothCurve issues moveTo and bezierCurveTo
    const canvas = createCanvas(100, 100)
    const ctx = canvas.getContext('2d')
    const calls: Array<{ type: string; args: number[] }> = []
    ctx.moveTo = (x, y) => { calls.push({ type: 'M', args: [x, y] }) }
    ctx.bezierCurveTo = (cp1x, cp1y, cp2x, cp2y, x, y) => {
      calls.push({ type: 'C', args: [cp1x, cp1y, cp2x, cp2y, x, y] })
    }

    traceSmoothCurve(ctx as any, points)
    expect(calls).toHaveLength(4)
    expect(calls[0].type).toBe('M')
    expect(calls[0].args).toEqual([0, 10])
    expect(calls[1].type).toBe('C')
    expect(calls[2].type).toBe('C')
    expect(calls[3].type).toBe('C')

    // Also verify paintXlsxChart with smoothed line chart issues Bézier curves (not lineTo)
    const xml = `
      <chartSpace>
        <chart>
          <plotArea>
            <lineChart>
              <smooth val="1"/>
              <ser>
                <spPr><ln><solidFill><srgbClr val="FF0000"/></solidFill></ln></spPr>
                <val>
                  <numLit>
                    <pt idx="0"><v>10</v></pt>
                    <pt idx="1"><v>20</v></pt>
                    <pt idx="2"><v>15</v></pt>
                    <pt idx="3"><v>30</v></pt>
                  </numLit>
                </val>
              </ser>
            </lineChart>
          </plotArea>
        </chart>
      </chartSpace>
    `
    const model = parseXlsxChart(parseXmlOrdered(xml), undefined)!
    expect(model.groups[0].series[0].smooth).toBe(true)

    const paintCanvas = createCanvas(400, 300)
    const paintCtx = paintCanvas.getContext('2d')
    const redOps: string[] = []
    const origMove = paintCtx.moveTo.bind(paintCtx)
    const origLine = paintCtx.lineTo.bind(paintCtx)
    const origBez = paintCtx.bezierCurveTo.bind(paintCtx)
    const isRed = () => paintCtx.strokeStyle === '#ff0000'
    paintCtx.moveTo = (x, y) => { if (isRed()) redOps.push('M'); origMove(x, y) }
    paintCtx.lineTo = (x, y) => { if (isRed()) redOps.push('L'); origLine(x, y) }
    paintCtx.bezierCurveTo = (cp1x, cp1y, cp2x, cp2y, x, y) => { if (isRed()) redOps.push('C'); origBez(cp1x, cp1y, cp2x, cp2y, x, y) }

    paintXlsxChart(model, paintCtx as any, 400, 300)
    // The red line series path should contain 'M' and three 'C' commands, NO 'L'
    expect(redOps).toEqual(['M', 'C', 'C', 'C'])
  })

  it('blanks break the curve per dispBlanksAs', () => {
    function paintWithBlanks(mode: 'gap' | 'span' | 'zero') {
      const xml = `
        <chartSpace>
          <chart>
            <dispBlanksAs val="${mode}"/>
            <plotArea>
              <lineChart>
                <smooth val="1"/>
                <ser>
                  <spPr><ln><solidFill><srgbClr val="FF0000"/></solidFill></ln></spPr>
                  <val>
                    <numLit>
                      <pt idx="0"><v>10</v></pt>
                      <pt idx="1"><v>20</v></pt>
                      <!-- idx 2 is blank -->
                      <pt idx="3"><v>15</v></pt>
                      <pt idx="4"><v>30</v></pt>
                    </numLit>
                  </val>
                </ser>
              </lineChart>
            </plotArea>
          </chart>
        </chartSpace>
      `
      const model = parseXlsxChart(parseXmlOrdered(xml), undefined)!
      const paintCanvas = createCanvas(400, 300)
      const paintCtx = paintCanvas.getContext('2d')
      const ops: string[] = []
      const isRed = () => paintCtx.strokeStyle === '#ff0000'
      paintCtx.moveTo = () => { if (isRed()) ops.push('M') }
      paintCtx.lineTo = () => { if (isRed()) ops.push('L') }
      paintCtx.bezierCurveTo = () => { if (isRed()) ops.push('C') }
      paintXlsxChart(model, paintCtx as any, 400, 300)
      return ops
    }

    // gap breaks into two separate continuous segments: M -> C, then M -> C
    const gapOps = paintWithBlanks('gap')
    expect(gapOps).toEqual(['M', 'C', 'M', 'C'])

    // span spans across the blank point: M -> C -> C -> C (single continuous curve of 4 points)
    const spanOps = paintWithBlanks('span')
    expect(spanOps).toEqual(['M', 'C', 'C', 'C'])

    // zero treats the blank as 0: M -> C -> C -> C -> C (5 points including the zero point)
    const zeroOps = paintWithBlanks('zero')
    expect(zeroOps).toEqual(['M', 'C', 'C', 'C', 'C'])
  })

  it('unequal x spacing for scatter', () => {
    // Scatter points with unequal x spacing:
    // P0 = (0, 0)
    // P1 = (2, 10)    [dx = 2]
    // P2 = (10, 4)    [dx = 8]
    // P3 = (20, 8)    [dx = 10]
    const scatterPoints: Point2D[] = [
      { x: 0, y: 0 },
      { x: 2, y: 10 },
      { x: 10, y: 4 },
      { x: 20, y: 8 },
    ]

    const segments = catmullRomToBezier(scatterPoints)
    expect(segments).toHaveLength(3)

    // Hand-computed oracle:
    // T0 = 0.5 * (P1 - P0) = (1, 5)
    // T1 = 0.5 * (P2 - P0) = (5, 2)
    // T2 = 0.5 * (P3 - P1) = (9, -1)
    // T3 = 0.5 * (P3 - P2) = (5, 2)
    //
    // Segment 0: P0 to P1
    // cp1 = P0 + T0 / 3 = (1/3, 5/3)
    // cp2 = P1 - T1 / 3 = (2 - 5/3, 10 - 2/3) = (1/3, 28/3)
    expect(segments[0].p0).toEqual({ x: 0, y: 0 })
    expect(segments[0].p1).toEqual({ x: 2, y: 10 })
    expect(segments[0].cp1.x).toBeCloseTo(1 / 3, 5)
    expect(segments[0].cp1.y).toBeCloseTo(5 / 3, 5)
    expect(segments[0].cp2.x).toBeCloseTo(1 / 3, 5)
    expect(segments[0].cp2.y).toBeCloseTo(28 / 3, 5)

    // Segment 1: P1 to P2
    // cp1 = P1 + T1 / 3 = (2 + 5/3, 10 + 2/3) = (11/3, 32/3)
    // cp2 = P2 - T2 / 3 = (10 - 3, 4 - (-1/3)) = (7, 13/3)
    expect(segments[1].p0).toEqual({ x: 2, y: 10 })
    expect(segments[1].p1).toEqual({ x: 10, y: 4 })
    expect(segments[1].cp1.x).toBeCloseTo(11 / 3, 5)
    expect(segments[1].cp1.y).toBeCloseTo(32 / 3, 5)
    expect(segments[1].cp2.x).toBeCloseTo(7, 5)
    expect(segments[1].cp2.y).toBeCloseTo(13 / 3, 5)

    // Segment 2: P2 to P3
    // cp1 = P2 + T2 / 3 = (10 + 3, 4 - 1/3) = (13, 11/3)
    // cp2 = P3 - T3 / 3 = (20 - 5/3, 8 - 2/3) = (55/3, 22/3)
    expect(segments[2].p0).toEqual({ x: 10, y: 4 })
    expect(segments[2].p1).toEqual({ x: 20, y: 8 })
    expect(segments[2].cp1.x).toBeCloseTo(13, 5)
    expect(segments[2].cp1.y).toBeCloseTo(11 / 3, 5)
    expect(segments[2].cp2.x).toBeCloseTo(55 / 3, 5)
    expect(segments[2].cp2.y).toBeCloseTo(22 / 3, 5)

    // End-to-end scatter chart paint test with smoothMarker style
    const xml = `
      <chartSpace>
        <chart>
          <plotArea>
            <scatterChart>
              <scatterStyle val="smoothMarker"/>
              <ser>
                <spPr><ln><solidFill><srgbClr val="FF0000"/></solidFill></ln></spPr>
                <xVal>
                  <numLit>
                    <pt idx="0"><v>0</v></pt>
                    <pt idx="1"><v>2</v></pt>
                    <pt idx="2"><v>10</v></pt>
                    <pt idx="3"><v>20</v></pt>
                  </numLit>
                </xVal>
                <yVal>
                  <numLit>
                    <pt idx="0"><v>0</v></pt>
                    <pt idx="1"><v>10</v></pt>
                    <pt idx="2"><v>4</v></pt>
                    <pt idx="3"><v>8</v></pt>
                  </numLit>
                </yVal>
              </ser>
            </scatterChart>
          </plotArea>
        </chart>
      </chartSpace>
    `
    const model = parseXlsxChart(parseXmlOrdered(xml), undefined)!
    expect(model.groups[0].series[0].smooth).toBe(true)

    const paintCanvas = createCanvas(400, 300)
    const paintCtx = paintCanvas.getContext('2d')
    const scatterOps: string[] = []
    const isRed = () => paintCtx.strokeStyle === '#ff0000'
    paintCtx.moveTo = () => { if (isRed()) scatterOps.push('M') }
    paintCtx.lineTo = () => { if (isRed()) scatterOps.push('L') }
    paintCtx.bezierCurveTo = () => { if (isRed()) scatterOps.push('C') }

    paintXlsxChart(model, paintCtx as any, 400, 300)
    // Series line should use Bézier curves (3 segments), not lineTo
    expect(scatterOps).toEqual(['M', 'C', 'C', 'C'])
  })

  it('smooth parse precedence: series override and scatterStyle defaults', () => {
    // 1. Line chart group smooth=1, but series has smooth=0 -> smooth is false
    const xml1 = `
      <chartSpace>
        <chart>
          <plotArea>
            <lineChart>
              <smooth val="1"/>
              <ser>
                <smooth val="0"/>
                <val><numLit><pt idx="0"><v>1</v></pt><pt idx="1"><v>2</v></pt></numLit></val>
              </ser>
            </lineChart>
          </plotArea>
        </chart>
      </chartSpace>
    `
    const model1 = parseXlsxChart(parseXmlOrdered(xml1), undefined)!
    expect(model1.groups[0].smooth).toBe(true)
    expect(model1.groups[0].series[0].smooth).toBe(false)

    // 2. Scatter with scatterStyle=line has smooth undefined/false
    const xml2 = `
      <chartSpace>
        <chart>
          <plotArea>
            <scatterChart>
              <scatterStyle val="line"/>
              <ser>
                <yVal><numLit><pt idx="0"><v>1</v></pt><pt idx="1"><v>2</v></pt></numLit></yVal>
              </ser>
            </scatterChart>
          </plotArea>
        </chart>
      </chartSpace>
    `
    const model2 = parseXlsxChart(parseXmlOrdered(xml2), undefined)!
    expect(model2.groups[0].smooth).toBeUndefined()
    expect(model2.groups[0].series[0].smooth).toBeUndefined()

    // 3. Scatter with ser smooth=1 enables smooth even when scatterStyle is line
    const xml3 = `
      <chartSpace>
        <chart>
          <plotArea>
            <scatterChart>
              <scatterStyle val="line"/>
              <ser>
                <smooth val="1"/>
                <yVal><numLit><pt idx="0"><v>1</v></pt><pt idx="1"><v>2</v></pt></numLit></yVal>
              </ser>
            </scatterChart>
          </plotArea>
        </chart>
      </chartSpace>
    `
    const model3 = parseXlsxChart(parseXmlOrdered(xml3), undefined)!
    expect(model3.groups[0].series[0].smooth).toBe(true)
  })

  it('degenerate cases: single point and empty series do not fail or draw invalid curves', () => {
    expect(catmullRomToBezier([])).toEqual([])
    expect(catmullRomToBezier([{ x: 10, y: 20 }])).toEqual([])

    const canvas = createCanvas(100, 100)
    const ctx = canvas.getContext('2d')
    const ops: string[] = []
    ctx.moveTo = () => { ops.push('M') }
    ctx.bezierCurveTo = () => { ops.push('C') }
    traceSmoothCurve(ctx as any, [{ x: 5, y: 5 }])
    expect(ops).toEqual(['M'])
  })

  // --- Regression tests for Fix Round 1 ---

  it('per-series smooth="0" override on a smooth="1" group paints with lineTo (straight lines)', () => {
    const xml = `
      <chartSpace>
        <chart>
          <plotArea>
            <lineChart>
              <smooth val="1"/>
              <ser>
                <smooth val="0"/>
                <spPr><ln><solidFill><srgbClr val="FF0000"/></solidFill></ln></spPr>
                <val>
                  <numLit>
                    <pt idx="0"><v>10</v></pt>
                    <pt idx="1"><v>20</v></pt>
                    <pt idx="2"><v>15</v></pt>
                    <pt idx="3"><v>30</v></pt>
                  </numLit>
                </val>
              </ser>
            </lineChart>
          </plotArea>
        </chart>
      </chartSpace>
    `
    const model = parseXlsxChart(parseXmlOrdered(xml), undefined)!
    expect(model.groups[0].smooth).toBe(true)
    expect(model.groups[0].series[0].smooth).toBe(false)

    const paintCanvas = createCanvas(400, 300)
    const paintCtx = paintCanvas.getContext('2d')
    const ops: string[] = []
    const isRed = () => paintCtx.strokeStyle === '#ff0000'
    paintCtx.moveTo = () => { if (isRed()) ops.push('M') }
    paintCtx.lineTo = () => { if (isRed()) ops.push('L') }
    paintCtx.bezierCurveTo = () => { if (isRed()) ops.push('C') }

    paintXlsxChart(model, paintCtx as any, 400, 300)
    // Must be straight lines: M followed by three L commands, ZERO C commands
    expect(ops).toEqual(['M', 'L', 'L', 'L'])
  })

  it('streaming createLineTracer streams Bézier curves with zero array allocations matching catmullRomToBezier', () => {
    const points: Point2D[] = [
      { x: 10, y: 50 },
      { x: 30, y: 90 },
      { x: 70, y: 40 },
      { x: 120, y: 110 },
      { x: 180, y: 60 },
    ]

    const expectedSegments = catmullRomToBezier(points)
    const canvas = createCanvas(200, 200)
    const ctx = canvas.getContext('2d')
    const streamed: Array<{ cp1x: number; cp1y: number; cp2x: number; cp2y: number; x: number; y: number }> = []
    let movePoint: [number, number] | undefined

    ctx.moveTo = (x, y) => { movePoint = [x, y] }
    ctx.bezierCurveTo = (cp1x, cp1y, cp2x, cp2y, x, y) => {
      streamed.push({ cp1x, cp1y, cp2x, cp2y, x, y })
    }

    const tracer = createLineTracer(ctx as any, true)
    for (const p of points) tracer.push(p.x, p.y)
    tracer.end()

    expect(movePoint).toEqual([10, 50])
    expect(streamed).toHaveLength(expectedSegments.length)
    for (let i = 0; i < expectedSegments.length; i++) {
      const exp = expectedSegments[i]
      const actual = streamed[i]
      expect(actual.cp1x).toBeCloseTo(exp.cp1.x, 5)
      expect(actual.cp1y).toBeCloseTo(exp.cp1.y, 5)
      expect(actual.cp2x).toBeCloseTo(exp.cp2.x, 5)
      expect(actual.cp2y).toBeCloseTo(exp.cp2.y, 5)
      expect(actual.x).toBeCloseTo(exp.p1.x, 5)
      expect(actual.y).toBeCloseTo(exp.p1.y, 5)
    }
  })

  it('scatter blank and error semantics with smoothing: #N/A spans and unavailable gaps', () => {
    function paintScatter(middleVal: string | null, unavailable = false) {
      const xml = `
        <chartSpace>
          <chart>
            <plotArea>
              <scatterChart>
                <scatterStyle val="smoothMarker"/>
                <ser>
                  <spPr><ln><solidFill><srgbClr val="FF0000"/></solidFill></ln></spPr>
                  <xVal>
                    <numLit>
                      <pt idx="0"><v>1</v></pt>
                      <pt idx="1"><v>2</v></pt>
                      <pt idx="2"><v>3</v></pt>
                    </numLit>
                  </xVal>
                  <yVal>
                    <numLit>
                      <pt idx="0"><v>10</v></pt>
                      ${middleVal !== null ? `<pt idx="1"><v>${middleVal}</v></pt>` : ''}
                      <pt idx="2"><v>20</v></pt>
                    </numLit>
                  </yVal>
                </ser>
              </scatterChart>
            </plotArea>
          </chart>
        </chartSpace>
      `
      const model = parseXlsxChart(parseXmlOrdered(xml), undefined)!
      if (unavailable) {
        const pw = new PointWriter(3)
        pw.value(0, 10)
        pw.unavailable(1)
        pw.value(2, 20)
        model.groups[0].series[0].values = pw.done()
      }
      const paintCanvas = createCanvas(400, 300)
      const paintCtx = paintCanvas.getContext('2d')
      const ops: string[] = []
      const isRed = () => paintCtx.strokeStyle === '#ff0000'
      paintCtx.moveTo = () => { if (isRed()) ops.push('M') }
      paintCtx.lineTo = () => { if (isRed()) ops.push('L') }
      paintCtx.bezierCurveTo = () => { if (isRed()) ops.push('C') }
      paintXlsxChart(model, paintCtx as any, 400, 300)
      return ops
    }

    // 1. Valid point: M -> C -> C (single continuous curve of 3 points)
    expect(paintScatter('15')).toEqual(['M', 'C', 'C'])

    // 2. #N/A: spans across the missing point -> M -> C (connects pt0 directly to pt2)
    expect(paintScatter('#N/A')).toEqual(['M', 'C'])

    // 3. Gap (absent point): breaks the curve into two disjoint runs (each with 1 point, so only M)
    expect(paintScatter(null)).toEqual(['M', 'M'])

    // 4. Unavailable point: breaks the curve into gaps
    expect(paintScatter(null, true)).toEqual(['M', 'M'])
  })

  it('exact geometric oracle for span across blank point asserts coordinates on painted canvas', () => {
    // 4 plottable points with a blank slot at index 2:
    // Slot 0: x=0, y=10
    // Slot 1: x=1, y=20
    // Slot 2: blank
    // Slot 3: x=3, y=15
    // Slot 4: x=4, y=30
    // Under dispBlanksAs="span", the sequence of points is:
    // P0(xAt(0), y10), P1(xAt(1), y20), P2(xAt(3), y15), P3(xAt(4), y30).
    const xml = `
      <chartSpace>
        <chart>
          <dispBlanksAs val="span"/>
          <plotArea>
            <lineChart>
              <smooth val="1"/>
              <ser>
                <spPr><ln><solidFill><srgbClr val="FF0000"/></solidFill></ln></spPr>
                <val>
                  <numLit>
                    <pt idx="0"><v>10</v></pt>
                    <pt idx="1"><v>20</v></pt>
                    <!-- idx 2 is blank -->
                    <pt idx="3"><v>15</v></pt>
                    <pt idx="4"><v>30</v></pt>
                  </numLit>
                </val>
              </ser>
            </lineChart>
          </plotArea>
        </chart>
      </chartSpace>
    `
    const model = parseXlsxChart(parseXmlOrdered(xml), undefined)!
    const paintCanvas = createCanvas(400, 300)
    const paintCtx = paintCanvas.getContext('2d')
    const captured: Array<{ cp1x: number; cp1y: number; cp2x: number; cp2y: number; x: number; y: number }> = []
    const isRed = () => paintCtx.strokeStyle === '#ff0000'
    const starts: Array<{ x: number; y: number }> = []
    const moveTo = paintCtx.moveTo.bind(paintCtx)
    paintCtx.moveTo = (x, y) => { if (isRed()) starts.push({ x, y }); moveTo(x, y) }
    paintCtx.bezierCurveTo = (cp1x, cp1y, cp2x, cp2y, x, y) => {
      if (isRed()) captured.push({ cp1x, cp1y, cp2x, cp2y, x, y })
    }

    paintXlsxChart(model, paintCtx as any, 400, 300)
    // There are 4 points connected by 3 Bézier segments across the span: one run, one moveTo.
    expect(captured).toHaveLength(3)
    expect(starts).toHaveLength(1)
    const P = [starts[0], ...captured.map(c => ({ x: c.x, y: c.y }))]

    // C1 (A4c N1): exact. Slots sit at equal steps on the category axis, so P2 (slot 3) is two steps from P1
    // (slot 1) and P3 one step from P2; values map linearly, so y offsets from P0 scale as 10 : 5 : 20.
    const step = P[1].x - P[0].x
    expect(step).toBeGreaterThan(0)
    expect(P[2].x - P[1].x).toBeCloseTo(2 * step, 9)
    expect(P[3].x - P[2].x).toBeCloseTo(step, 9)
    const unit = (P[1].y - P[0].y) / 10
    expect(unit).toBeLessThan(0)
    expect(P[2].y - P[0].y).toBeCloseTo(5 * unit, 9)
    expect(P[3].y - P[0].y).toBeCloseTo(20 * unit, 9)
    // Catmull-Rom (tension 1/2) control points: segment i from P[i] to P[i+1] has
    // cp1 = P[i] + (P[i+1] - P[i-1]) / 6 and cp2 = P[i+1] - (P[i+2] - P[i]) / 6, an end point its own neighbour.
    const at = (i: number) => P[Math.max(0, Math.min(P.length - 1, i))]
    captured.forEach((c, i) => {
      expect(c.cp1x, `segment ${i} cp1x`).toBeCloseTo(at(i).x + (at(i + 1).x - at(i - 1).x) / 6, 9)
      expect(c.cp1y, `segment ${i} cp1y`).toBeCloseTo(at(i).y + (at(i + 1).y - at(i - 1).y) / 6, 9)
      expect(c.cp2x, `segment ${i} cp2x`).toBeCloseTo(at(i + 1).x - (at(i + 2).x - at(i).x) / 6, 9)
      expect(c.cp2y, `segment ${i} cp2y`).toBeCloseTo(at(i + 1).y - (at(i + 2).y - at(i).y) / 6, 9)
    })
    // The bridge (slot 1 to slot 3) takes its tangents from slot 0 and slot 4, across the blank.
    expect(captured[1].cp1x).toBeCloseTo(P[1].x + (P[2].x - P[0].x) / 6, 9)
    expect(captured[1].cp2y).toBeCloseTo(P[2].y - (P[3].y - P[1].y) / 6, 9)
  })

  it('smoothed stacked lines accumulate base values', () => {
    const xml = `
      <chartSpace>
        <chart>
          <plotArea>
            <lineChart>
              <grouping val="stacked"/>
              <smooth val="1"/>
              <ser>
                <spPr><ln><solidFill><srgbClr val="FF0000"/></solidFill></ln></spPr>
                <val>
                  <numLit>
                    <pt idx="0"><v>10</v></pt>
                    <pt idx="1"><v>20</v></pt>
                  </numLit>
                </val>
              </ser>
              <ser>
                <spPr><ln><solidFill><srgbClr val="00FF00"/></solidFill></ln></spPr>
                <val>
                  <numLit>
                    <pt idx="0"><v>5</v></pt>
                    <pt idx="1"><v>15</v></pt>
                  </numLit>
                </val>
              </ser>
            </lineChart>
          </plotArea>
        </chart>
      </chartSpace>
    `
    const model = parseXlsxChart(parseXmlOrdered(xml), undefined)!
    expect(model.groups[0].grouping).toBe('stacked')

    const paintCanvas = createCanvas(400, 300)
    const paintCtx = paintCanvas.getContext('2d')
    const redBeziers: Array<[number, number]> = []
    const greenBeziers: Array<[number, number]> = []

    paintCtx.bezierCurveTo = (_cp1x, _cp1y, _cp2x, _cp2y, x, y) => {
      if (paintCtx.strokeStyle === '#ff0000') redBeziers.push([x, y])
      if (paintCtx.strokeStyle === '#00ff00') greenBeziers.push([x, y])
    }

    paintXlsxChart(model, paintCtx as any, 400, 300)
    expect(redBeziers).toHaveLength(1)
    expect(greenBeziers).toHaveLength(1)
    // Series 2 stacks on top of Series 1, so green end y is higher (closer to top = smaller y coordinate in canvas)
    expect(greenBeziers[0][1]).toBeLessThan(redBeziers[0][1])
  })
})
