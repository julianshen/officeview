// @vitest-environment node
/**
 * Task A5b: Stock charts (spec 4.1, 5.1).
 * HLC, OHLC, volume-HLC, and volume-OHLC.
 *
 * Rules:
 * - Volume, when present, is the series of the bar group in the same plot.
 * - Stock series are, in order: Open (only when 4 series), High, Low, Close.
 * - Axes come from authored axId bindings, never assumed.
 * - Draw hiLowLines and upDownBars (up/down fills from c:upBars/c:downBars).
 * - A series-count/role mismatch becomes a card with a reason.
 */
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { prepareChartAxes } from '../src/charts/prepare'
import type { ChartModel } from '../src/charts/model'
import {
  CONFORMANCE_NS,
  describeFamilyConformance,
} from './helpers/chart-family-conformance'

const lit = (values: Array<number | string | undefined>) =>
  `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values
    .map((v, i) => (v === undefined ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`))
    .join('')}</c:numLit>`

const strLit = (labels: string[]) =>
  `<c:strLit><c:ptCount val="${labels.length}"/>${labels
    .map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)
    .join('')}</c:strLit>`

const space = (plot: string, options = '') =>
  `<c:chartSpace ${CONFORMANCE_NS}><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>${plot}</c:plotArea><c:plotVisOnly val="1"/>${options}</c:chart></c:chartSpace>`

const CAT_AX = (id: string, crossAx: string, deleted = false) =>
  `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${deleted ? '1' : '0'}"/><c:axPos val="b"/><c:crossAx val="${crossAx}"/></c:catAx>`

const VAL_AX = (id: string, crossAx: string, pos = 'l', title = '') =>
  `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:majorGridlines/><c:crossAx val="${crossAx}"/>${title ? `<c:title><c:tx><c:v>${title}</c:v></c:tx></c:title>` : ''}</c:valAx>`

const parse = (xml: string): ChartModel => parseXlsxChart(parseXmlOrdered(xml), undefined)!

describe('stock (spec 4.1, Task A5b)', () => {
  it('volume from the bar group; open, high, low, close by order; authored axes', () => {
    // Volume-OHLC combination chart:
    // Group 1: barChart (barDir="col") binding catAx 1 and valAx 2 (Volume: 0 to 1,000,000, pos="l")
    // Group 2: stockChart (4 series: Open, High, Low, Close) binding catAx 1 and valAx 3 (Price: 100 to 200, pos="r")
    const categories = ['2026-01-01', '2026-01-02']
    const volumeData = [500000, 800000]
    const openData = [120, 150]
    const highData = [160, 180]
    const lowData = [110, 130]
    const closeData = [150, 140]

    const xml = space(`
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/><c:tx><c:v>Volume</c:v></c:tx>
          <c:cat>${strLit(categories)}</c:cat>
          <c:val>${lit(volumeData)}</c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:stockChart>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/><c:tx><c:v>Open</c:v></c:tx>
          <c:cat>${strLit(categories)}</c:cat>
          <c:val>${lit(openData)}</c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/><c:order val="1"/><c:tx><c:v>High</c:v></c:tx>
          <c:cat>${strLit(categories)}</c:cat>
          <c:val>${lit(highData)}</c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2"/><c:order val="2"/><c:tx><c:v>Low</c:v></c:tx>
          <c:cat>${strLit(categories)}</c:cat>
          <c:val>${lit(lowData)}</c:val>
        </c:ser>
        <c:ser>
          <c:idx val="3"/><c:order val="3"/><c:tx><c:v>Close</c:v></c:tx>
          <c:cat>${strLit(categories)}</c:cat>
          <c:val>${lit(closeData)}</c:val>
        </c:ser>
        <c:hiLowLines/>
        <c:upDownBars/>
        <c:axId val="3"/>
        <c:axId val="4"/>
      </c:stockChart>
      ${CAT_AX('1', '2')}
      ${VAL_AX('2', '1', 'l', 'Volume')}
      ${CAT_AX('3', '4', true)}
      ${VAL_AX('4', '3', 'r', 'Price')}
    `)

    const model = parse(xml)
    expect(model.unsupported).toBeUndefined()
    expect(model.groups).toHaveLength(2)

    // Group 0: Volume (bar)
    expect(model.groups[0].family).toBe('bar')
    expect(model.groups[0].binding).toEqual({ kind: 'cartesian', category: '1', value: '2' })
    expect(model.groups[0].series[0].name).toBe('Volume')

    // Group 1: Stock
    expect(model.groups[1].family).toBe('stock')
    expect(model.groups[1].binding).toEqual({ kind: 'cartesian', category: '3', value: '4' })
    expect(model.groups[1].series).toHaveLength(4)
    expect(model.groups[1].series[0].name).toBe('Open')
    expect(model.groups[1].series[1].name).toBe('High')
    expect(model.groups[1].series[2].name).toBe('Low')
    expect(model.groups[1].series[3].name).toBe('Close')

    // Resolved axes domains:
    const prepared = prepareChartAxes(model)
    const axes = prepared.preparedAxes!.axes
    const volAxis = axes.find(a => a.id === '2')!
    const priceAxis = axes.find(a => a.id === '4')!

    expect(volAxis.role).toBe('value')
    expect(priceAxis.role).toBe('value')
    // Volume axis max must be >= 800,000 and min 0
    expect(volAxis.max).toBeGreaterThanOrEqual(800000)
    // Price axis domain spans ~100 to ~180 (never 800,000!)
    expect(priceAxis.min).toBeLessThanOrEqual(110)
    expect(priceAxis.max).toBeGreaterThanOrEqual(180)
    expect(priceAxis.max).toBeLessThan(1000)

    // Paint canvas oracle: verify both groups paint
    const canvas = createCanvas(500, 300)
    const ctx = canvas.getContext('2d')
    const filledRects: Array<{ x: number; y: number; w: number; h: number; style: string }> = []
    const lines: Array<{ x1: number; y1: number; x2: number; y2: number }> = []

    const fillRect = ctx.fillRect.bind(ctx)
    ctx.fillRect = (x, y, w, h) => {
      filledRects.push({ x, y, w, h, style: String(ctx.fillStyle) })
      fillRect(x, y, w, h)
    }
    const moveTo = ctx.moveTo.bind(ctx)
    let curX = 0, curY = 0
    ctx.moveTo = (x, y) => { curX = x; curY = y; moveTo(x, y) }
    const lineTo = ctx.lineTo.bind(ctx)
    ctx.lineTo = (x, y) => {
      lines.push({ x1: curX, y1: curY, x2: x, y2: y })
      curX = x; curY = y
      lineTo(x, y)
    }

    paintXlsxChart(prepared, ctx as any, 500, 300)
    // There are 2 categories:
    // Volume bars: 2 filled rects
    // Stock upDownBars: 2 filled rects (1 up bar, 1 down bar)
    // Plus hiLowLines: vertical lines connecting High and Low
    expect(filledRects.length).toBeGreaterThanOrEqual(4)
    // hiLowLines must be vertical (x1 == x2) and span high to low
    const verticalLines = lines.filter(l => Math.abs(l.x1 - l.x2) < 0.01 && Math.abs(l.y1 - l.y2) > 5)
    expect(verticalLines.length).toBeGreaterThanOrEqual(2)
  })

  it('HLC, OHLC, volume-HLC and volume-OHLC', () => {
    const H = [50, 60, 55], L = [30, 40, 35], C = [45, 42, 50], O = [35, 58, 38], V = [1000, 2000, 1500]

    const hlcXml = space(`
      <c:stockChart>
        <c:ser><c:idx val="0"/><c:val>${lit(H)}</c:val></c:ser>
        <c:ser><c:idx val="1"/><c:val>${lit(L)}</c:val></c:ser>
        <c:ser><c:idx val="2"/><c:val>${lit(C)}</c:val></c:ser>
        <c:hiLowLines/>
        <c:axId val="1"/><c:axId val="2"/>
      </c:stockChart>
      ${CAT_AX('1', '2')}${VAL_AX('2', '1')}
    `)

    const ohlcXml = space(`
      <c:stockChart>
        <c:ser><c:idx val="0"/><c:val>${lit(O)}</c:val></c:ser>
        <c:ser><c:idx val="1"/><c:val>${lit(H)}</c:val></c:ser>
        <c:ser><c:idx val="2"/><c:val>${lit(L)}</c:val></c:ser>
        <c:ser><c:idx val="3"/><c:val>${lit(C)}</c:val></c:ser>
        <c:hiLowLines/>
        <c:upDownBars/>
        <c:axId val="1"/><c:axId val="2"/>
      </c:stockChart>
      ${CAT_AX('1', '2')}${VAL_AX('2', '1')}
    `)

    const vHlcXml = space(`
      <c:barChart>
        <c:barDir val="col"/><c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:val>${lit(V)}</c:val></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:stockChart>
        <c:ser><c:idx val="0"/><c:val>${lit(H)}</c:val></c:ser>
        <c:ser><c:idx val="1"/><c:val>${lit(L)}</c:val></c:ser>
        <c:ser><c:idx val="2"/><c:val>${lit(C)}</c:val></c:ser>
        <c:hiLowLines/>
        <c:axId val="3"/><c:axId val="4"/>
      </c:stockChart>
      ${CAT_AX('1', '2')}${VAL_AX('2', '1')}${CAT_AX('3', '4', true)}${VAL_AX('4', '3', 'r')}
    `)

    const vOhlcXml = space(`
      <c:barChart>
        <c:barDir val="col"/><c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:val>${lit(V)}</c:val></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:stockChart>
        <c:ser><c:idx val="0"/><c:val>${lit(O)}</c:val></c:ser>
        <c:ser><c:idx val="1"/><c:val>${lit(H)}</c:val></c:ser>
        <c:ser><c:idx val="2"/><c:val>${lit(L)}</c:val></c:ser>
        <c:ser><c:idx val="3"/><c:val>${lit(C)}</c:val></c:ser>
        <c:hiLowLines/>
        <c:upDownBars/>
        <c:axId val="3"/><c:axId val="4"/>
      </c:stockChart>
      ${CAT_AX('1', '2')}${VAL_AX('2', '1')}${CAT_AX('3', '4', true)}${VAL_AX('4', '3', 'r')}
    `)

    const mHlc = parse(hlcXml)
    const mOhlc = parse(ohlcXml)
    const mVHlc = parse(vHlcXml)
    const mVOhlc = parse(vOhlcXml)

    expect(mHlc.unsupported).toBeUndefined()
    expect(mOhlc.unsupported).toBeUndefined()
    expect(mVHlc.unsupported).toBeUndefined()
    expect(mVOhlc.unsupported).toBeUndefined()

    // Test paint execution without card placeholder
    for (const model of [mHlc, mOhlc, mVHlc, mVOhlc]) {
      const canvas = createCanvas(400, 300)
      const ctx = canvas.getContext('2d')
      let card = false
      const fillRect = ctx.fillRect.bind(ctx)
      ctx.fillRect = (x, y, w, h) => {
        if (String(ctx.fillStyle).toLowerCase() === '#f7f7f7' && w === 400 && h === 300) card = true
        fillRect(x, y, w, h)
      }
      paintXlsxChart(model, ctx as any, 400, 300)
      expect(card, 'native stock charts must not paint the placeholder card').toBe(false)
    }
  })

  it('hiLowLines and upDownBars colours', () => {
    // 2 categories:
    // Pt 0: Open 100, High 150, Low 90, Close 140 (Close > Open => Up bar)
    // Pt 1: Open 140, High 160, Low 70, Close 80  (Close < Open => Down bar)
    const xml = space(`
      <c:stockChart>
        <c:ser><c:idx val="0"/><c:val>${lit([100, 140])}</c:val></c:ser>
        <c:ser><c:idx val="1"/><c:val>${lit([150, 160])}</c:val></c:ser>
        <c:ser><c:idx val="2"/><c:val>${lit([90, 70])}</c:val></c:ser>
        <c:ser><c:idx val="3"/><c:val>${lit([140, 80])}</c:val></c:ser>
        <c:hiLowLines>
          <c:spPr>
            <a:ln w="19050">
              <a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>
            </a:ln>
          </c:spPr>
        </c:hiLowLines>
        <c:upDownBars>
          <c:upBars>
            <c:spPr>
              <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
              <a:ln><a:solidFill><a:srgbClr val="003300"/></a:solidFill></a:ln>
            </c:spPr>
          </c:upBars>
          <c:downBars>
            <c:spPr>
              <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
              <a:ln><a:solidFill><a:srgbClr val="330000"/></a:solidFill></a:ln>
            </c:spPr>
          </c:downBars>
        </c:upDownBars>
        <c:axId val="1"/><c:axId val="2"/>
      </c:stockChart>
      ${CAT_AX('1', '2')}${VAL_AX('2', '1')}
    `)

    const model = parse(xml)
    expect(model.unsupported).toBeUndefined()

    const canvas = createCanvas(400, 300)
    const ctx = canvas.getContext('2d')
    const fills: Array<{ color: string; x: number; y: number; w: number; h: number }> = []
    const strokes: Array<{ color: string; width: number; x1: number; y1: number; x2: number; y2: number }> = []

    const fillRect = ctx.fillRect.bind(ctx)
    ctx.fillRect = (x, y, w, h) => {
      fills.push({ color: String(ctx.fillStyle).toLowerCase(), x, y, w, h })
      fillRect(x, y, w, h)
    }

    let cx = 0, cy = 0
    ctx.moveTo = (x, y) => { cx = x; cy = y }
    ctx.lineTo = (x, y) => {
      strokes.push({ color: String(ctx.strokeStyle).toLowerCase(), width: ctx.lineWidth, x1: cx, y1: cy, x2: x, y2: y })
      cx = x; cy = y
    }

    paintXlsxChart(model, ctx as any, 400, 300)

    // Up bar: #00ff00 fill
    const upFill = fills.find(f => f.color === '#00ff00')
    expect(upFill).toBeDefined()

    // Down bar: #ff0000 fill
    const downFill = fills.find(f => f.color === '#ff0000')
    expect(downFill).toBeDefined()

    // Up bar comes before down bar in x order
    expect(upFill!.x).toBeLessThan(downFill!.x)

    // hiLowLines: #0000ff stroke
    const hiLowStrokes = strokes.filter(s => s.color === '#0000ff' && Math.abs(s.x1 - s.x2) < 0.01)
    expect(hiLowStrokes).toHaveLength(2)
  })

  it('a role mismatch becomes a card', () => {
    // 1 series
    const m1 = parse(space(`
      <c:stockChart>
        <c:ser><c:idx val="0"/><c:val>${lit([10, 20])}</c:val></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:stockChart>
      ${CAT_AX('1', '2')}${VAL_AX('2', '1')}
    `))
    expect(m1.unsupported).toMatch(/series/i)

    // 2 series
    const m2 = parse(space(`
      <c:stockChart>
        <c:ser><c:idx val="0"/><c:val>${lit([10, 20])}</c:val></c:ser>
        <c:ser><c:idx val="1"/><c:val>${lit([15, 25])}</c:val></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:stockChart>
      ${CAT_AX('1', '2')}${VAL_AX('2', '1')}
    `))
    expect(m2.unsupported).toMatch(/series/i)

    // 5 series
    const m5 = parse(space(`
      <c:stockChart>
        <c:ser><c:idx val="0"/><c:val>${lit([10, 20])}</c:val></c:ser>
        <c:ser><c:idx val="1"/><c:val>${lit([15, 25])}</c:val></c:ser>
        <c:ser><c:idx val="2"/><c:val>${lit([5, 12])}</c:val></c:ser>
        <c:ser><c:idx val="3"/><c:val>${lit([12, 22])}</c:val></c:ser>
        <c:ser><c:idx val="4"/><c:val>${lit([14, 21])}</c:val></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:stockChart>
      ${CAT_AX('1', '2')}${VAL_AX('2', '1')}
    `))
    expect(m5.unsupported).toMatch(/series/i)

    // Painting a mismatch model renders placeholder card
    const canvas = createCanvas(400, 300)
    const ctx = canvas.getContext('2d')
    let card = false
    const fillRect = ctx.fillRect.bind(ctx)
    ctx.fillRect = (x, y, w, h) => {
      if (String(ctx.fillStyle).toLowerCase() === '#f7f7f7' && w === 400 && h === 300) card = true
      fillRect(x, y, w, h)
    }
    paintXlsxChart(m1, ctx as any, 400, 300)
    expect(card).toBe(true)
  })

  // Conformance harness
  describeFamilyConformance({
    name: 'stock',
    chart: c => {
      if (c === null) {
        return space(`<c:stockChart><c:hiLowLines/><c:axId val="1"/><c:axId val="2"/></c:stockChart>${CAT_AX('1', '2')}${VAL_AX('2', '1')}`)
      }
      return space(`
        <c:stockChart>
          <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>High</c:v></c:tx><c:val>${c}</c:val></c:ser>
          <c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:v>Low</c:v></c:tx><c:val>${c}</c:val></c:ser>
          <c:ser><c:idx val="2"/><c:order val="2"/><c:tx><c:v>Close</c:v></c:tx><c:val>${c}</c:val></c:ser>
          <c:hiLowLines/>
          <c:axId val="1"/><c:axId val="2"/>
        </c:stockChart>
        ${CAT_AX('1', '2')}${VAL_AX('2', '1')}
      `)
    },
    expect: {
      'no series': { kind: 'card' },
      'empty series': { kind: 'native' },
      'all blank': { kind: 'native' },
      'NaN and infinity': { kind: 'native' },
      'negative values': { kind: 'native' },
      'single point': { kind: 'native' },
      'unavailable point': { kind: 'native' },
    },
  })
})
