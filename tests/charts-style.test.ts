// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { MARKER_AUTO_CYCLE, clampMarkerSize } from '../src/charts/markers'
import { resolveSeriesColor } from '../src/charts/style'
import type { ThemeContext } from '../src/drawing/style'

const parse = (xml: string, theme?: ThemeContext) => parseXlsxChart(parseXmlOrdered(xml), theme)!

describe('charts-style (Task A4a)', () => {
  it('marker symbols, auto cycle, size clamp 2-72, marker spPr', () => {
    // 1. Symbol parsing & auto cycle
    expect(MARKER_AUTO_CYCLE.length).toBeGreaterThanOrEqual(9)
    expect(MARKER_AUTO_CYCLE).toContain('square')
    expect(MARKER_AUTO_CYCLE).toContain('diamond')
    expect(MARKER_AUTO_CYCLE).toContain('triangle')
    expect(MARKER_AUTO_CYCLE).toContain('circle')

    // 2. Size clamp 2-72
    expect(clampMarkerSize(1)).toBe(2)
    expect(clampMarkerSize(75)).toBe(72)
    expect(clampMarkerSize(10)).toBe(10)
    expect(clampMarkerSize(undefined)).toBe(5)

    // 3. XML parse marker with size, symbol, spPr
    const xml = `<chartSpace><chart><plotArea><lineChart>
      <ser>
        <val><numLit><pt idx="0"><v>10</v></pt></numLit></val>
        <marker>
          <symbol val="diamond"/>
          <size val="12"/>
          <spPr>
            <solidFill><srgbClr val="FF00AA"/></solidFill>
            <ln w="12700"><solidFill><srgbClr val="00FF00"/></solidFill></ln>
          </spPr>
        </marker>
      </ser>
    </lineChart></plotArea></chart></chartSpace>`
    const model = parse(xml)
    const ser = model.groups[0].series[0]
    expect(ser.markerStyle?.symbol).toBe('diamond')
    expect(ser.markerStyle?.size).toBe(12)
    expect(ser.markerStyle?.fill?.kind).toBe('solid')

    // 4. Paint verification: marker draws with diamond geometry and fill/stroke
    const canvas = createCanvas(300, 200)
    const ctx = canvas.getContext('2d')
    let fillSeen = false
    let strokeSeen = false
    const origFill = ctx.fill.bind(ctx)
    const origStroke = ctx.stroke.bind(ctx)
    ctx.fill = () => { fillSeen = true; origFill() }
    ctx.stroke = () => { strokeSeen = true; origStroke() }
    paintXlsxChart(model, ctx as any, 300, 200)
    expect(fillSeen).toBe(true)
    expect(strokeSeen).toBe(true)
  })

  it('line dash, width, cap and noFill', () => {
    const xml = `<chartSpace><chart><plotArea><lineChart>
      <ser>
        <val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt></numLit></val>
        <spPr>
          <ln w="25400" cap="rnd">
            <noFill/>
            <prstDash val="dash"/>
          </ln>
        </spPr>
      </ser>
      <ser>
        <val><numLit><pt idx="0"><v>15</v></pt><pt idx="1"><v>25</v></pt></numLit></val>
        <spPr>
          <ln w="19050" cap="sq">
            <solidFill><srgbClr val="0000FF"/></solidFill>
            <prstDash val="dot"/>
          </ln>
        </spPr>
      </ser>
    </lineChart></plotArea></chart></chartSpace>`
    const model = parse(xml)
    const s0 = model.groups[0].series[0]
    const s1 = model.groups[0].series[1]
    expect(s0.lineStyle?.noFill).toBe(true)
    expect(s0.lineStyle?.dash).toBe('dash')
    expect(s0.lineStyle?.cap).toBe('round')

    expect(s1.lineStyle?.noFill).toBe(false)
    expect(s1.lineStyle?.dash).toBe('dot')
    expect(s1.lineStyle?.cap).toBe('square')
    expect(s1.lineStyle?.width).toBeCloseTo(2, 1)

    // Paint verification: s0 is noFill (not stroked), s1 is stroked with dashed line and round/sq cap
    const canvas = createCanvas(300, 200)
    const ctx = canvas.getContext('2d')
    const dashesSeen: number[][] = []
    const origSetLineDash = ctx.setLineDash.bind(ctx)
    ctx.setLineDash = (d) => { dashesSeen.push([...d]); origSetLineDash(d) }
    paintXlsxChart(model, ctx as any, 300, 200)
    expect(dashesSeen.some(d => d.length > 0)).toBe(true)
  })

  it('solid, gradient, pattern and noFill for series, points, plot area, chart area, legend and title', () => {
    const xml = `<chartSpace>
      <spPr><solidFill><srgbClr val="F0F0F0"/></solidFill></spPr>
      <chart>
        <title><spPr><noFill/></spPr><tx><v>Test Title</v></tx></title>
        <plotArea>
          <spPr>
            <gradFill rotWithShape="1">
              <gsLst>
                <gs pos="0"><srgbClr val="FFFFFF"/></gs>
                <gs pos="100000"><srgbClr val="CCCCCC"/></gs>
              </gsLst>
              <lin ang="5400000" scaled="0"/>
            </gradFill>
          </spPr>
          <barChart>
            <ser>
              <val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt></numLit></val>
              <spPr>
                <pattFill prst="pct50">
                  <fgClr><srgbClr val="000000"/></fgClr>
                  <bgClr><srgbClr val="FFFFFF"/></bgClr>
                </pattFill>
              </spPr>
              <dPt>
                <idx val="1"/>
                <spPr><solidFill><srgbClr val="FF0000"/></solidFill></spPr>
              </dPt>
            </ser>
          </barChart>
        </plotArea>
        <legend><spPr><solidFill><srgbClr val="E0E0E0"/></solidFill></spPr></legend>
      </chart>
    </chartSpace>`
    const model = parse(xml)
    expect(model.chartAreaFill?.kind).toBe('solid')
    expect(model.titleFill?.kind).toBe('none')
    expect(model.plotAreaFill?.kind).toBe('gradient')
    const s0 = model.groups[0].series[0]
    expect(s0.fill?.kind).toBe('pattern')
    expect(s0.pointFills?.[1]?.kind).toBe('solid')
    expect(model.legendFill?.kind).toBe('solid')

    // Paint to ensure gradient and pattern execute without throwing
    const canvas = createCanvas(300, 200)
    const ctx = canvas.getContext('2d')
    expect(() => paintXlsxChart(model, ctx as any, 300, 200)).not.toThrow()
  })

  it('chart area noFill is transparent; default via c:style', () => {
    // 1. Explicit noFill
    const xmlNoFill = `<chartSpace><spPr><noFill/></spPr><chart><plotArea><barChart><ser><val><numLit><pt idx="0"><v>5</v></pt></numLit></val></ser></barChart></plotArea></chart></chartSpace>`
    const mNoFill = parse(xmlNoFill)
    expect(mNoFill.chartAreaFill?.kind).toBe('none')

    const canvas1 = createCanvas(200, 200)
    const ctx1 = canvas1.getContext('2d')
    let fillRects: Array<{ x: number; y: number; w: number; h: number; style: any }> = []
    const origFillRect1 = ctx1.fillRect.bind(ctx1)
    ctx1.fillRect = (x, y, w, h) => { fillRects.push({ x, y, w, h, style: ctx1.fillStyle }); origFillRect1(x, y, w, h) }
    paintXlsxChart(mNoFill, ctx1 as any, 200, 200)
    // Entire canvas background (0,0,200,200) should NOT be filled with opaque white #FFFFFF when noFill
    const chartBg = fillRects.find(r => r.x === 0 && r.y === 0 && r.w === 200 && r.h === 200)
    expect(chartBg).toBeUndefined()

    // 2. Default Excel chart is white
    const xmlDefault = `<chartSpace><chart><plotArea><barChart><ser><val><numLit><pt idx="0"><v>5</v></pt></numLit></val></ser></barChart></plotArea></chart></chartSpace>`
    const mDefault = parse(xmlDefault)
    const canvas2 = createCanvas(200, 200)
    const ctx2 = canvas2.getContext('2d')
    fillRects = []
    const origFillRect2 = ctx2.fillRect.bind(ctx2)
    ctx2.fillRect = (x, y, w, h) => { fillRects.push({ x, y, w, h, style: ctx2.fillStyle }); origFillRect2(x, y, w, h) }
    paintXlsxChart(mDefault, ctx2 as any, 200, 200)
    expect(fillRects.some(r => r.x === 0 && r.y === 0 && r.w === 200 && r.h === 200 && String(r.style).toUpperCase() === '#FFFFFF')).toBe(true)

    // 3. c:style default
    const xmlStyle = `<chartSpace><chart><plotArea><barChart><ser><val><numLit><pt idx="0"><v>5</v></pt></numLit></val></ser></barChart></plotArea></chart><style val="2"/></chartSpace>`
    const mStyle = parse(xmlStyle)
    expect(mStyle.styleId).toBe(2)
  })

  it('colour precedence spPr, chartStyle, colorStyle, c:style, then accents with lumMod variants after six', () => {
    const dummyTheme: ThemeContext = {
      palette: {
        accent1: '#4472C4',
        accent2: '#ED7D31',
        accent3: '#A5A5A5',
        accent4: '#FFC000',
        accent5: '#5B9BD5',
        accent6: '#70AD47',
      },
      colors: {},
      colorMap: {},
      fonts: { major: { latin: 'Calibri', supplemental: {} }, minor: { latin: 'Calibri', supplemental: {} } },
      fillStyles: [],
      bgFillStyles: [],
      lineStyles: [],
      effectStyles: [],
      issues: [],
    }

    // Explicit spPr wins
    const colorSpPr = resolveSeriesColor({
      spPrColor: '112233',
      chartStyleColor: '445566',
      colorStyleColor: '778899',
      styleId: 2,
      seriesIndex: 0,
      theme: dummyTheme,
    })
    expect(colorSpPr).toBe('112233')

    // chartStyle part wins over colorStyle and c:style
    const colorChartStyle = resolveSeriesColor({
      chartStyleColor: '445566',
      colorStyleColor: '778899',
      styleId: 2,
      seriesIndex: 0,
      theme: dummyTheme,
    })
    expect(colorChartStyle).toBe('445566')

    // colorStyle part wins over c:style and accents
    const colorColorStyle = resolveSeriesColor({
      colorStyleColor: '778899',
      styleId: 2,
      seriesIndex: 0,
      theme: dummyTheme,
    })
    expect(colorColorStyle).toBe('778899')

    // Series 0-5 use theme accents
    expect(resolveSeriesColor({ seriesIndex: 0, theme: dummyTheme })).toBe('4472C4')
    expect(resolveSeriesColor({ seriesIndex: 1, theme: dummyTheme })).toBe('ED7D31')
    expect(resolveSeriesColor({ seriesIndex: 5, theme: dummyTheme })).toBe('70AD47')

    // Series 6+ uses lumMod variants after 6
    const ser6Color = resolveSeriesColor({ seriesIndex: 6, theme: dummyTheme })
    expect(ser6Color).not.toBe('4472C4') // modified by lumMod
    expect(ser6Color).toBeDefined()
  })

  it('invertIfNegative fills negative bars with the inverted fill', () => {
    const xml = `<chartSpace><chart><plotArea><barChart>
      <ser>
        <val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>-5</v></pt></numLit></val>
        <spPr><solidFill><srgbClr val="0000FF"/></solidFill></spPr>
        <invertIfNegative val="1"/>
      </ser>
    </barChart></plotArea></chart></chartSpace>`
    const model = parse(xml)
    const ser = model.groups[0].series[0]
    expect(ser.invertIfNegative).toBe(true)

    const canvas = createCanvas(300, 200)
    const ctx = canvas.getContext('2d')
    const filledColors: string[] = []
    const origFillRect = ctx.fillRect.bind(ctx)
    ctx.fillRect = (x, y, w, h) => {
      if (w > 5 && h > 5) filledColors.push(String(ctx.fillStyle).toUpperCase())
      origFillRect(x, y, w, h)
    }
    paintXlsxChart(model, ctx as any, 300, 200)
    // Must have drawn positive bar with #0000FF and negative bar with inverted color (not #0000FF)
    expect(filledColors).toContain('#0000FF')
    expect(filledColors.some(c => c !== '#0000FF' && c !== '#FFFFFF')).toBe(true)
  })

  it('series and point explosion geometry', () => {
    const xml = `<chartSpace><chart><plotArea><pieChart>
      <ser>
        <explosion val="20"/>
        <val><numLit><pt idx="0"><v>50</v></pt><pt idx="1"><v>50</v></pt></numLit></val>
        <dPt><idx val="0"/><explosion val="35"/></dPt>
      </ser>
    </pieChart></plotArea></chart></chartSpace>`
    const model = parse(xml)
    const ser = model.groups[0].series[0]
    expect(ser.explosion).toBe(20)
    expect(ser.dPt?.get(0)?.explosion).toBe(35)

    const canvas = createCanvas(300, 300)
    const ctx = canvas.getContext('2d')
    const moves: Array<[number, number]> = []
    const origMoveTo = ctx.moveTo.bind(ctx)
    ctx.moveTo = (x, y) => { moves.push([x, y]); origMoveTo(x, y) }
    paintXlsxChart(model, ctx as any, 300, 300)
    // Slices are exploded away from the center (150, 150)
    expect(moves.length).toBeGreaterThanOrEqual(2)
    const centerDist0 = Math.hypot(moves[0][0] - 150, moves[0][1] - 150)
    const centerDist1 = Math.hypot(moves[1][0] - 150, moves[1][1] - 150)
    expect(centerDist0).toBeGreaterThan(0)
    expect(centerDist1).toBeGreaterThan(0)
    expect(centerDist0).not.toBeCloseTo(centerDist1, 1) // 35% vs 20% explosion
  })

  it('dPt overrides take precedence over series formatting', () => {
    const xml = `<chartSpace><chart><plotArea><barChart>
      <ser>
        <val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt></numLit></val>
        <spPr><solidFill><srgbClr val="00FF00"/></solidFill></spPr>
        <dPt>
          <idx val="0"/>
          <spPr><solidFill><srgbClr val="FF0000"/></solidFill></spPr>
        </dPt>
      </ser>
    </barChart></plotArea></chart></chartSpace>`
    const model = parse(xml)
    const ser = model.groups[0].series[0]
    expect(ser.color).toBe('00FF00')
    expect(ser.pointColors?.[0]).toBe('FF0000')

    const canvas = createCanvas(300, 200)
    const ctx = canvas.getContext('2d')
    const filledColors: string[] = []
    const origFillRect = ctx.fillRect.bind(ctx)
    ctx.fillRect = (x, y, w, h) => {
      if (w > 5 && h > 5) filledColors.push(String(ctx.fillStyle).toUpperCase())
      origFillRect(x, y, w, h)
    }
    paintXlsxChart(model, ctx as any, 300, 200)
    expect(filledColors).toContain('#FF0000')
    expect(filledColors).toContain('#00FF00')
  })

  it('gapWidth, overlap and varyColors for existing families', () => {
    const xml = `<chartSpace><chart><plotArea><barChart>
      <gapWidth val="200"/>
      <overlap val="50"/>
      <varyColors val="1"/>
      <ser>
        <val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt><pt idx="2"><v>30</v></pt></numLit></val>
      </ser>
    </barChart></plotArea></chart></chartSpace>`
    const model = parse(xml)
    const group = model.groups[0]
    expect(group.gapWidth).toBe(200)
    expect(group.overlap).toBe(50)
    expect(group.varyColors).toBe(true)

    // With varyColors=1, points 0, 1, 2 have distinct pointColors
    const ser = group.series[0]
    expect(ser.pointColors?.[0]).toBeDefined()
    expect(ser.pointColors?.[1]).toBeDefined()
    expect(ser.pointColors?.[0]).not.toBe(ser.pointColors?.[1])
  })

  it('dropLines on line and area groups and serLines on stacked bars draw the expected segments', () => {
    // 1. Line chart with dropLines
    const xmlDrop = `<chartSpace><chart><plotArea><lineChart>
      <dropLines>
        <spPr><ln w="12700"><solidFill><srgbClr val="888888"/></solidFill></ln></spPr>
      </dropLines>
      <ser><val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt></numLit></val></ser>
    </lineChart></plotArea></chart></chartSpace>`
    const mDrop = parse(xmlDrop)
    expect(mDrop.groups[0].dropLines).toBeDefined()

    const canvas1 = createCanvas(300, 200)
    const ctx1 = canvas1.getContext('2d')
    let lineSegments: Array<[number, number, number, number]> = []
    let curX = 0, curY = 0
    ctx1.moveTo = (x, y) => { curX = x; curY = y }
    ctx1.lineTo = (x, y) => { lineSegments.push([curX, curY, x, y]); curX = x; curY = y }
    paintXlsxChart(mDrop, ctx1 as any, 300, 200)
    // Expect vertical segments dropped to crossing
    const verticalDrops = lineSegments.filter(([x1, y1, x2, y2]) => Math.abs(x1 - x2) < 1e-3 && Math.abs(y1 - y2) > 5)
    expect(verticalDrops.length).toBeGreaterThanOrEqual(2)

    // 2. Stacked bar chart with serLines
    const xmlSer = `<chartSpace><chart><plotArea><barChart>
      <grouping val="stacked"/>
      <serLines>
        <spPr><ln w="12700"><solidFill><srgbClr val="555555"/></solidFill></ln></spPr>
      </serLines>
      <ser><val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt></numLit></val></ser>
      <ser><val><numLit><pt idx="0"><v>15</v></pt><pt idx="1"><v>25</v></pt></numLit></val></ser>
    </barChart></plotArea></chart></chartSpace>`
    const mSer = parse(xmlSer)
    expect(mSer.groups[0].serLines).toBeDefined()

    const canvas2 = createCanvas(300, 200)
    const ctx2 = canvas2.getContext('2d')
    lineSegments = []
    ctx2.moveTo = (x, y) => { curX = x; curY = y }
    ctx2.lineTo = (x, y) => { lineSegments.push([curX, curY, x, y]); curX = x; curY = y }
    paintXlsxChart(mSer, ctx2 as any, 300, 200)
    // serLines connect adjacent categories across the series boundary
    const connectingLines = lineSegments.filter(([x1, , x2]) => Math.abs(x1 - x2) > 10)
    expect(connectingLines.length).toBeGreaterThan(0)
  })

  it('dark mode leaves chart pixels identical', () => {
    const xml = `<chartSpace><chart><plotArea><barChart>
      <ser><val><numLit><pt idx="0"><v>10</v></pt><pt idx="1"><v>20</v></pt></numLit></val></ser>
    </barChart></plotArea></chart></chartSpace>`
    const model = parse(xml)

    const cLight = createCanvas(300, 200)
    const ctxLight = cLight.getContext('2d')
    paintXlsxChart(model, ctxLight as any, 300, 200)
    const dataLight = ctxLight.getImageData(0, 0, 300, 200).data

    const cDark = createCanvas(300, 200)
    const ctxDark = cDark.getContext('2d')
    // Pre-fill canvas with dark background to simulate dark mode host environment
    ctxDark.fillStyle = '#1e1e1e'
    ctxDark.fillRect(0, 0, 300, 200)
    paintXlsxChart(model, ctxDark as any, 300, 200)
    const dataDark = ctxDark.getImageData(0, 0, 300, 200).data

    // Chart Area paints its paper background and pixels identically
    expect(Buffer.from(dataLight).equals(Buffer.from(dataDark))).toBe(true)
  })
})
