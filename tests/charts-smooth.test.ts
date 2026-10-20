// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { catmullRomToBezier, traceSmoothCurve, type Point2D } from '../src/charts/smooth'

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
})
