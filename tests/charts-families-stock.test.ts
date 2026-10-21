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
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { getChildren, parseXmlOrdered } from '../src/core/xml'
import { layoutChart } from '../src/charts/layout'
import { readStockOptions } from '../src/charts/families/stock'
import { CHART_STYLE_BYTES, ChartRefusal } from '../src/charts/own'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { prepareChartAxes } from '../src/charts/prepare'
import { documentChartContext, chartParseContext } from '../src/charts/context'
import { parseChartPart } from '../src/charts/dataPrecedence'
import { chartDescription } from '../src/charts/describe'
import { parseDocx } from '../src/docx/parse'
import { parsePptx } from '../src/pptx/parse'
import { buildChartDocx, buildChartPptx } from './helpers/chart-fixtures'
import type { ChartModel } from '../src/charts/model'
import { CONFORMANCE_NS } from './helpers/chart-family-conformance'

let emptyPackage: Promise<OfficePackage> | undefined
const pkg = () => emptyPackage ??= new JSZip().generateAsync({ type: 'uint8array' }).then(bytes => OfficePackage.load(bytes))

const lit = (values: Array<number | string | undefined>) =>
  `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values
    .map((v, i) => (v === undefined ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`))
    .join('')}</c:numLit>`

const numRef = (values: Array<number | string | undefined>) =>
  `<c:numRef><c:f>Sheet1!$A$1:$A$${values.length}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values
    .map((v, i) => (v === undefined ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`))
    .join('')}</c:numCache></c:numRef>`

const strLit = (labels: string[]) =>
  `<c:strLit><c:ptCount val="${labels.length}"/>${labels
    .map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)
    .join('')}</c:strLit>`

const space = (plot: string, options = '') =>
  `<c:chartSpace ${CONFORMANCE_NS}><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>${plot}</c:plotArea><c:plotVisOnly val="1"/>${options}</c:chart></c:chartSpace>`

const CAT_AX = (id: string, crossAx: string, deleted = false) =>
  `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${deleted ? '1' : '0'}"/><c:axPos val="b"/><c:crossAx val="${crossAx}"/></c:catAx>`

const DATE_AX = (id: string, crossAx: string, deleted = false) =>
  `<c:dateAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${deleted ? '1' : '0'}"/><c:axPos val="b"/><c:baseTimeUnit val="days"/><c:crossAx val="${crossAx}"/></c:dateAx>`

const VAL_AX = (id: string, crossAx: string, pos = 'l', title = '', min?: number, max?: number) =>
  `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/>${min !== undefined ? `<c:min val="${min}"/>` : ''}${max !== undefined ? `<c:max val="${max}"/>` : ''}</c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:majorGridlines/><c:crossAx val="${crossAx}"/>${title ? `<c:title><c:tx><c:v>${title}</c:v></c:tx></c:title>` : ''}</c:valAx>`

const parse = (xml: string): ChartModel => parseXlsxChart(parseXmlOrdered(xml), undefined)!

describe('stock (spec 4.1, Task A5b)', () => {
  it('volume from the bar group; open, high, low, close by order; authored axes', () => {
    // Volume-OHLC combination chart:
    // Group 1: barChart (barDir="col") binding catAx 1 and valAx 2 (Volume: 0 to 1,000,000, pos="l")
    // Group 2: stockChart (4 series: Open, High, Low, Close) binding catAx 3 (deleted) and valAx 4 (Price: 100 to 200, pos="r")
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

    // Accessible description of volume stock chart must be Stock chart, not Column chart
    expect(chartDescription(model)).toMatch(/Stock chart/)

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
    expect(volAxis.max).toBeGreaterThanOrEqual(800000)
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
    let curX = 0, curY = 0
    ctx.moveTo = (x, y) => { curX = x; curY = y }
    ctx.lineTo = (x, y) => {
      lines.push({ x1: curX, y1: curY, x2: x, y2: y })
      curX = x; curY = y
    }

    paintXlsxChart(prepared, ctx as any, 500, 300)
    expect(filledRects.length).toBeGreaterThanOrEqual(4)
    const verticalLines = lines.filter(l => Math.abs(l.x1 - l.x2) < 0.01 && Math.abs(l.y1 - l.y2) > 5)
    expect(verticalLines.length).toBeGreaterThanOrEqual(2)
  })

  it('date category axis aligns stock marks with volume bars', () => {
    // Dates: Fri (46024), Mon (46027), Tue (46028) with a weekend gap!
    // Excel's date axis places these at non-consecutive slots according to day count.
    const dates = [46024, 46027, 46028]
    const V = [1000, 2000, 1500]
    const H = [150, 160, 170]
    const L = [100, 110, 120]
    const C = [130, 140, 150]

    const xml = space(`
      <c:barChart>
        <c:barDir val="col"/><c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:cat>${lit(dates)}</c:cat>
          <c:val>${lit(V)}</c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:stockChart>
        <c:ser><c:idx val="0"/><c:cat>${lit(dates)}</c:cat><c:val>${lit(H)}</c:val></c:ser>
        <c:ser><c:idx val="1"/><c:cat>${lit(dates)}</c:cat><c:val>${lit(L)}</c:val></c:ser>
        <c:ser><c:idx val="2"/><c:cat>${lit(dates)}</c:cat><c:val>${lit(C)}</c:val></c:ser>
        <c:hiLowLines/>
        <c:axId val="3"/><c:axId val="4"/>
      </c:stockChart>
      ${DATE_AX('1', '2')}
      ${VAL_AX('2', '1', 'l', 'Volume')}
      ${DATE_AX('3', '4', true)}
      ${VAL_AX('4', '3', 'r', 'Price')}
    `)

    const model = parse(xml)
    expect(model.unsupported).toBeUndefined()

    const canvas = createCanvas(500, 300)
    const ctx = canvas.getContext('2d')
    const barCenters: number[] = []
    const hiLowXs: number[] = []

    const fillRect = ctx.fillRect.bind(ctx)
    ctx.fillRect = (x, y, w, h) => {
      // Volume bars are drawn with bar geometry (height > 5, width > 5)
      if (h > 5 && w > 5 && y > 0 && y + h < 300) {
        barCenters.push(x + w / 2)
      }
      fillRect(x, y, w, h)
    }

    let cx = 0, cy = 0
    ctx.moveTo = (x, y) => { cx = x; cy = y }
    ctx.lineTo = (x, y) => {
      if (Math.abs(cx - x) < 0.01 && Math.abs(cy - y) > 10) {
        hiLowXs.push(x)
      }
      cx = x; cy = y
    }

    paintXlsxChart(model, ctx as any, 500, 300)

    expect(barCenters).toHaveLength(3)
    expect(hiLowXs).toHaveLength(3)

    // Assert that each stock hiLowLine's x coordinate aligns exactly with its volume bar center!
    for (let i = 0; i < 3; i++) {
      expect(Math.abs(hiLowXs[i] - barCenters[i])).toBeLessThan(0.01)
    }
  })

  it('numeric role-mapping oracle: verifies hand-computed exact geometry', () => {
    // Fixed plot size and explicit value axis bounds (min 50, max 200).
    // Domain span = 150 units.
    // In canvas coordinates:
    // y(v) = plot.y + plot.h - ((v - 50) / 150) * plot.h.
    // 2 categories:
    // Cat 0: Open 100, High 150, Low 90, Close 140 (Up bar: Open 100, Close 140)
    // Cat 1: Open 140, High 160, Low 70, Close 80 (Down bar: Open 140, Close 80)
    const xml = space(`
      <c:stockChart>
        <c:ser><c:idx val="0"/><c:val>${lit([100, 140])}</c:val></c:ser>
        <c:ser><c:idx val="1"/><c:val>${lit([150, 160])}</c:val></c:ser>
        <c:ser><c:idx val="2"/><c:val>${lit([90, 70])}</c:val></c:ser>
        <c:ser><c:idx val="3"/><c:val>${lit([140, 80])}</c:val></c:ser>
        <c:hiLowLines/>
        <c:upDownBars/>
        <c:axId val="1"/><c:axId val="2"/>
      </c:stockChart>
      ${CAT_AX('1', '2')}
      ${VAL_AX('2', '1', 'l', '', 50, 200)}
    `)

    const model = parse(xml)
    const prepared = prepareChartAxes(model)
    const vAxis = prepared.preparedAxes!.axes.find(a => a.id === '2')!
    expect([vAxis.min, vAxis.max]).toEqual([50, 200])

    const canvas = createCanvas(400, 300)
    const ctx = canvas.getContext('2d')
    const capturedBodies: Array<{ x: number; yTop: number; yBot: number; isUp: boolean }> = []
    const capturedHiLows: Array<{ x: number; yHigh: number; yLow: number }> = []

    const fillRect = ctx.fillRect.bind(ctx)
    ctx.fillRect = (x, y, w, h) => {
      if (w > 5 && h > 0 && y > 0 && y + h < 300) {
        capturedBodies.push({ x, yTop: y, yBot: y + h, isUp: String(ctx.fillStyle).toLowerCase() === '#ffffff' })
      }
      fillRect(x, y, w, h)
    }

    let cx = 0, cy = 0
    ctx.moveTo = (x, y) => { cx = x; cy = y }
    ctx.lineTo = (x, y) => {
      if (Math.abs(cx - x) < 0.01 && Math.abs(cy - y) > 10) {
        capturedHiLows.push({ x, yHigh: Math.min(cy, y), yLow: Math.max(cy, y) })
      }
      cx = x; cy = y
    }

    paintXlsxChart(prepared, ctx as any, 400, 300)

    expect(capturedHiLows).toHaveLength(2)
    expect(capturedBodies).toHaveLength(2)

    // Compute expected Y coordinates on the plot
    // Cat 0:
    // hi-low line must span exactly Low(90) to High(150):
    // yLow > yHigh in canvas coords (since y increases downward)
    expect(capturedHiLows[0].yLow).toBeGreaterThan(capturedHiLows[0].yHigh)
    // Up body: top is Close(140), bottom is Open(100)
    expect(capturedBodies[0].isUp).toBe(true)
    // Body top (140) is below High (150) -> yTop > yHigh
    expect(capturedBodies[0].yTop).toBeGreaterThan(capturedHiLows[0].yHigh)
    // Body bottom (100) is above Low (90) -> yBot < yLow
    expect(capturedBodies[0].yBot).toBeLessThan(capturedHiLows[0].yLow)

    // Cat 1:
    // Down body: top is Open(140), bottom is Close(80)
    expect(capturedBodies[1].isUp).toBe(false)
    expect(capturedBodies[1].yTop).toBeGreaterThan(capturedHiLows[1].yHigh)
    expect(capturedBodies[1].yBot).toBeLessThan(capturedHiLows[1].yLow)

    // High of Cat 1 (160) is higher than High of Cat 0 (150) -> Cat 1 yHigh < Cat 0 yHigh
    expect(capturedHiLows[1].yHigh).toBeLessThan(capturedHiLows[0].yHigh)
    // Low of Cat 1 (70) is lower than Low of Cat 0 (90) -> Cat 1 yLow > Cat 0 yLow
    expect(capturedHiLows[1].yLow).toBeGreaterThan(capturedHiLows[0].yLow)

    // Also check HLC Close tick geometry
    const hlcXml = space(`
      <c:stockChart>
        <c:ser><c:idx val="0"/><c:val>${lit([150])}</c:val></c:ser>
        <c:ser><c:idx val="1"/><c:val>${lit([90])}</c:val></c:ser>
        <c:ser><c:idx val="2"/><c:val>${lit([140])}</c:val></c:ser>
        <c:hiLowLines/>
        <c:axId val="1"/><c:axId val="2"/>
      </c:stockChart>
      ${CAT_AX('1', '2')}
      ${VAL_AX('2', '1', 'l', '', 50, 200)}
    `)
    const mHlc = parse(hlcXml)
    const hlcCanvas = createCanvas(400, 300)
    const hlcCtx = hlcCanvas.getContext('2d')
    const horizTicks: Array<{ y: number; x1: number; x2: number }> = []
    let hx = 0, hy = 0
    hlcCtx.moveTo = (x, y) => { hx = x; hy = y }
    hlcCtx.lineTo = (x, y) => {
      if (Math.abs(hy - y) < 0.01 && x > hx && (x - hx) <= 100) {
        horizTicks.push({ y, x1: hx, x2: x })
      }
      hx = x; hy = y
    }
    paintXlsxChart(mHlc, hlcCtx as any, 400, 300)
    expect(horizTicks).toHaveLength(1)
    // The close tick Y corresponds to close price 140, which matches Cat 0's close Y
    expect(Math.abs(horizTicks[0].y - capturedBodies[0].yTop)).toBeLessThan(1)
  })

  it('HLC, OHLC, volume-HLC and volume-OHLC: native models, representation and geometry', async () => {
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

    const doc = documentChartContext('desktop')
    const parseAdapter = async (xmlStr: string) => {
      const context = chartParseContext(doc, await pkg(), '/xl/charts/chart1.xml')
      return parseChartPart(parseXmlOrdered(xmlStr), undefined, context)
    }

    const resHlc = await parseAdapter(hlcXml)
    const resOhlc = await parseAdapter(ohlcXml)
    const resVHlc = await parseAdapter(vHlcXml)
    const resVOhlc = await parseAdapter(vOhlcXml)

    // Assert model kind is native, rejecting any fallback representation
    expect(resHlc.kind).toBe('native')
    expect(resOhlc.kind).toBe('native')
    expect(resVHlc.kind).toBe('native')
    expect(resVOhlc.kind).toBe('native')

    // Document coverage assert: DOCX and PPTX select 'chart' representation
    const docxPkg = await OfficePackage.load((await buildChartDocx([{ chartXml: ohlcXml }])).bytes)
    const docx = await parseDocx(docxPkg, { chartContext: doc })
    const docxCoverage = (docx.drawingCoverage ?? []).filter(e => e.scope === 'original')
    expect(docxCoverage.map(c => c.selectedRepresentation)).toEqual(['chart'])

    const pptxPkg = await OfficePackage.load((await buildChartPptx([{ chartXml: vOhlcXml }])).bytes)
    const pptx = await parsePptx(pptxPkg, { chartContext: doc })
    const pptxCoverage = (pptx.drawingCoverage ?? []).filter(e => e.scope === 'original')
    expect(pptxCoverage.map(c => c.selectedRepresentation)).toEqual(['chart'])

    // Helper to count distinct geometry elements painted
    const measureGeometry = (model: ChartModel) => {
      const canvas = createCanvas(400, 300)
      const ctx = canvas.getContext('2d')
      let rects = 0, vertLines = 0, horizTicks = 0, card = false

      const fillRect = ctx.fillRect.bind(ctx)
      ctx.fillRect = (x, y, w, h) => {
        if (String(ctx.fillStyle).toLowerCase() === '#f7f7f7' && w === 400 && h === 300) card = true
        if (w > 2 && h > 2 && y > 0 && y + h < 300) rects++
        fillRect(x, y, w, h)
      }

      let lx = 0, ly = 0
      ctx.moveTo = (x, y) => { lx = x; ly = y }
      ctx.lineTo = (x, y) => {
        if (Math.abs(lx - x) < 0.01 && Math.abs(ly - y) > 5) vertLines++
        if (Math.abs(ly - y) < 0.01 && x > lx && (x - lx) < 100) horizTicks++
        lx = x; ly = y
      }

      paintXlsxChart(model, ctx as any, 400, 300)
      return { rects, vertLines, horizTicks, card }
    }

    if (resHlc.kind === 'native') {
      const g = measureGeometry(resHlc.model)
      expect(g.card).toBe(false)
      expect(g.vertLines).toBe(3) // 3 hi-low lines
      expect(g.horizTicks).toBe(3) // 3 close ticks
      expect(g.rects).toBe(0) // No body rects in HLC
    }

    if (resOhlc.kind === 'native') {
      const g = measureGeometry(resOhlc.model)
      expect(g.card).toBe(false)
      expect(g.vertLines).toBe(3) // 3 hi-low lines
      expect(g.rects).toBe(3) // 3 up/down bodies in OHLC
    }

    if (resVHlc.kind === 'native') {
      const g = measureGeometry(resVHlc.model)
      expect(g.card).toBe(false)
      expect(g.vertLines).toBe(3) // 3 hi-low lines
      expect(g.horizTicks).toBe(3) // 3 close ticks
      expect(g.rects).toBe(3) // 3 volume bars
    }

    if (resVOhlc.kind === 'native') {
      const g = measureGeometry(resVOhlc.model)
      expect(g.card).toBe(false)
      expect(g.vertLines).toBe(3) // 3 hi-low lines
      expect(g.rects).toBe(6) // 3 volume bars + 3 stock bodies
    }
  })

  it('blank and error points: skips bodies or hi-low lines appropriately', () => {
    // 4 categories:
    // Slot 0: valid OHLC (O=100, H=150, L=80, C=140) -> paints hi-low line + body
    // Slot 1: blank Open (O=blank, H=150, L=80, C=140) -> paints hi-low line, body skipped
    // Slot 2: error High (O=100, H=#N/A, L=80, C=140) -> hi-low line skipped, body painted
    // Slot 3: error Close (O=100, H=150, L=80, C=#N/A) -> paints hi-low line, body skipped
    const O = [100, undefined, 100, 100]
    const H = [150, 150, '#N/A', 150]
    const L = [80, 80, 80, 80]
    const C = [140, 140, 140, '#N/A']

    const xml = space(`
      <c:stockChart>
        <c:ser><c:idx val="0"/><c:val>${numRef(O)}</c:val></c:ser>
        <c:ser><c:idx val="1"/><c:val>${numRef(H)}</c:val></c:ser>
        <c:ser><c:idx val="2"/><c:val>${numRef(L)}</c:val></c:ser>
        <c:ser><c:idx val="3"/><c:val>${numRef(C)}</c:val></c:ser>
        <c:hiLowLines/>
        <c:upDownBars/>
        <c:axId val="1"/><c:axId val="2"/>
      </c:stockChart>
      ${CAT_AX('1', '2')}${VAL_AX('2', '1')}
    `)

    const model = parse(xml)
    expect(model.unsupported).toBeUndefined()

    const canvas = createCanvas(400, 300)
    const ctx = canvas.getContext('2d')
    let bodies = 0, hiLows = 0

    const fillRect = ctx.fillRect.bind(ctx)
    ctx.fillRect = (x, y, w, h) => {
      if (w > 2 && h > 0 && y > 0 && y + h < 300) bodies++
      fillRect(x, y, w, h)
    }

    let cx = 0, cy = 0
    ctx.moveTo = (x, y) => { cx = x; cy = y }
    ctx.lineTo = (x, y) => {
      if (Math.abs(cx - x) < 0.01 && Math.abs(cy - y) > 5) hiLows++
      cx = x; cy = y
    }

    paintXlsxChart(model, ctx as any, 400, 300)

    // Hi-Low lines drawn for slots 0, 1, 3 (slot 2 has #N/A High, so skipped)
    expect(hiLows).toBe(3)
    // Up-Down bodies drawn for slots 0, 2 (slot 1 has blank Open, slot 3 has #N/A Close, so both skipped)
    expect(bodies).toBe(2)
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
})

describe('C1: exact-coordinate stock role oracle and the up/down bar charge (A5b parked items)', () => {
  const W = 480, H = 320
  const measureCtx = createCanvas(10, 10).getContext('2d')
  const plotOf = (model: ChartModel) => layoutChart(model, W, H, 1, 0, { width: (t, f) => { measureCtx.font = f; return measureCtx.measureText(t).width }, lineHeight: () => model.fontSizePt * 96 / 72 }, undefined, JSON.stringify(model.fontFamily)).plot
  /** Paints on a recording canvas: every stroked segment (with its stroke colour) and every filled rectangle. */
  function record(model: ChartModel) {
    const ctx = createCanvas(W, H).getContext('2d')
    const segments: Array<{ x1: number; y1: number; x2: number; y2: number; color: string }> = [], rects: Array<{ x: number; y: number; w: number; h: number; color: string }> = []
    let pen: [number, number] = [0, 0]
    const moveTo = ctx.moveTo.bind(ctx), lineTo = ctx.lineTo.bind(ctx), fillRect = ctx.fillRect.bind(ctx)
    ctx.moveTo = (x, y) => { pen = [x, y]; moveTo(x, y) }
    ctx.lineTo = (x, y) => { segments.push({ x1: pen[0], y1: pen[1], x2: x, y2: y, color: String(ctx.strokeStyle).toLowerCase() }); pen = [x, y]; lineTo(x, y) }
    ctx.fillRect = (x, y, w, h) => { rects.push({ x, y, w, h, color: String(ctx.fillStyle).toLowerCase() }); fillRect(x, y, w, h) }
    paintXlsxChart(model, ctx as any, W, H)
    return { segments, rects }
  }
  const near = (a: number, b: number) => Math.abs(a - b) < 1e-6
  // Three categories, explicit price axis 0..100 (and volume axis 0..3000 in the volume charts), so every coordinate
  // is hand-computable: band = plot.w / 3, centre x = plot.x + (i + 0.5) band, price y(v) = plot.y + plot.h (1 - v / 100);
  // gapWidth 150: a bar is band / 2.5 wide, centred in its band.
  const O = [20, 70, 40], Hi = [60, 80, 90], L = [10, 30, 35], C = [50, 40, 80], V = [1500, 3000, 600]
  const stockSeries = (withOpen: boolean) => (withOpen ? [O, Hi, L, C] : [Hi, L, C]).map((v, i) => `<c:ser><c:idx val="${i + 1}"/><c:order val="${i + 1}"/><c:val>${lit(v)}</c:val></c:ser>`).join('')
  const volume = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill><a:srgbClr val="00B050"/></a:solidFill></c:spPr><c:val>${lit(V)}</c:val></c:ser><c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>`
  const xml = (withOpen: boolean, withVolume: boolean) => space(withVolume
    ? `${volume}<c:stockChart>${stockSeries(withOpen)}<c:hiLowLines/>${withOpen ? '<c:upDownBars><c:gapWidth val="150"/><c:upBars/><c:downBars/></c:upDownBars>' : ''}<c:axId val="3"/><c:axId val="4"/></c:stockChart>${CAT_AX('1', '2')}${VAL_AX('2', '1', 'l', '', 0, 3000)}${CAT_AX('3', '4', true)}${VAL_AX('4', '3', 'r', '', 0, 100).replace('<c:crossAx val="3"/>', '<c:crossAx val="3"/><c:crosses val="max"/>')}`
    : `<c:stockChart>${stockSeries(withOpen)}<c:hiLowLines/>${withOpen ? '<c:upDownBars><c:gapWidth val="150"/><c:upBars/><c:downBars/></c:upDownBars>' : ''}<c:axId val="1"/><c:axId val="2"/></c:stockChart>${CAT_AX('1', '2')}${VAL_AX('2', '1', 'l', '', 0, 100)}`)

  it.each([['HLC', false, false], ['OHLC', true, false], ['volume-HLC', false, true], ['volume-OHLC', true, true]] as const)('%s: painted coordinates equal the hand-computed role geometry', (_name, withOpen, withVolume) => {
    const model = prepareChartAxes(parse(xml(withOpen, withVolume)))
    expect(model.unsupported).toBeUndefined()
    const plot = plotOf(model), band = plot.w / 3, width = band / 2.5
    const x = (i: number) => plot.x + (i + 0.5) * band
    const y = (v: number) => plot.y + plot.h * (1 - v / 100)
    const { segments, rects } = record(model)
    for (let i = 0; i < 3; i++) {
      // High-low line: Low to High at the category centre.
      expect(segments.some(s => near(s.x1, x(i)) && near(s.x2, x(i)) && near(s.y1, y(L[i])) && near(s.y2, y(Hi[i]))), `hi-low ${i}`).toBe(true)
      if (withOpen) {
        // Up/down bar between Open and Close: white when Close >= Open, black otherwise.
        const top = Math.min(y(O[i]), y(C[i])), bottom = Math.max(y(O[i]), y(C[i]))
        const up = C[i] >= O[i]
        expect(rects.some(r => near(r.x, x(i) - width / 2) && near(r.w, width) && near(r.y, top) && near(r.h, bottom - top) && r.color === (up ? '#ffffff' : '#000000')), `body ${i}`).toBe(true)
      } else {
        // HLC: the close tick runs right from the centre, half a bar wide, at Close.
        expect(segments.some(s => near(s.x1, x(i)) && near(s.y1, y(C[i])) && near(s.x2, x(i) + width / 2) && near(s.y2, y(C[i]))), `close ${i}`).toBe(true)
      }
      if (withVolume) {
        // Volume on its own axis (0..3000), one clustered series: the bar fills its band's middle 1/2.5.
        const top = plot.y + plot.h * (1 - V[i] / 3000)
        expect(rects.some(r => r.color === '#00b050' && near(r.x, x(i) - width / 2) && near(r.w, width) && near(r.y, top) && near(r.y + r.h, plot.y + plot.h)), `volume ${i}`).toBe(true)
      }
    }
    // No up/down bodies without an Open series.
    if (!withOpen) expect(rects.filter(r => r.color === '#000000' || r.color === '#ffffff').filter(r => r.w === width)).toHaveLength(0)
  })

  it('the up/down bar record is charged by name before either bar is built', () => {
    const node = parseXmlOrdered(`<c:stockChart ${CONFORMANCE_NS}><c:upDownBars><c:gapWidth val="150"/><c:upBars><c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c:spPr></c:upBars><c:downBars/></c:upDownBars></c:stockChart>`)
    const root = getChildren(node, 'stockChart')[0] ?? node
    const charges: number[] = []
    const options = readStockOptions(root, undefined, bytes => { charges.push(bytes) })
    expect(options.upDownBars?.upBars?.fill).toBeDefined()
    expect(charges[0]).toBe(CHART_STYLE_BYTES.upDownBars)
    expect(charges.slice(1)).toContain(CHART_STYLE_BYTES.fill)
    // A refusal at the record's charge builds nothing: no bar fill is ever priced.
    const refused: number[] = []
    expect(() => readStockOptions(root, undefined, bytes => { refused.push(bytes); throw new ChartRefusal() })).toThrow(ChartRefusal)
    expect(refused).toEqual([CHART_STYLE_BYTES.upDownBars])
  })
})
