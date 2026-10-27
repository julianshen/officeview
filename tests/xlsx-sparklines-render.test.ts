// @vitest-environment node
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { getPaintables } from '../src/render/paint'
import { computeMetrics, computePrintMetrics, renderPrintPage, renderSheet } from '../src/xlsx/render'
import { parseXlsx } from '../src/xlsx/parse'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { registerSparklinePainter } from '../src/xlsx/sparklineLoader'
import { sparklinePrimitivesForViewport, translateSparkPrimitive } from '../src/xlsx/sparklines/paint'

const basic = () => readFileSync(fileURLToPath(new URL('./fixtures/xlsx-sparklines/line-basic.xlsx', import.meta.url)))
const parse = async () => parseXlsx(await OfficePackage.load(basic()))

describe('XLSX sparkline rendering', () => {
  it('a direct render before preload does not throw and records one unavailable diagnostic', async () => {
    const doc = await parse(), sheet = doc.sheets[0]
    const metrics = computeMetrics(sheet), canvas = createCanvas(metrics.widthPx, metrics.heightPx)
    expect(() => renderSheet(sheet, canvas.getContext('2d') as unknown as CanvasRenderingContext2D, metrics)).not.toThrow()
    expect(sheet.diagnostics?.filter(d => d.feature === 'sparkline-painter-unavailable')).toHaveLength(1)
  })

  it('getPaintables paints line-basic ink only within its host cell', async () => {
    const doc = await parse(), sheet = doc.sheets[0]
    const pages = await getPaintables(doc)
    try {
      const page = pages[0], canvas = createCanvas(page.spec.widthPx, page.spec.heightPx), ctx = canvas.getContext('2d')
      page.paint(ctx as unknown as CanvasRenderingContext2D)
      const m = computeMetrics(sheet), [row, cols] = [...sheet.resolvedSparklines!.entries()][0], [col, spark] = [...cols.entries()][0]
      const x = m.colWidthsPx.slice(0, col).reduce((a, b) => a + b, 0), y = m.rowHeightsPx.slice(0, row).reduce((a, b) => a + b, 0)
      const w = m.colWidthsPx[col], h = m.rowHeightsPx[row]
      const image = ctx.getImageData(0, 0, canvas.width, canvas.height).data
      let inside = 0, outside = 0
      const color = spark.group.colors.series?.slice(2).toUpperCase() ?? '000000'
      for (let py = 0; py < canvas.height; py++) for (let px = 0; px < canvas.width; px++) {
        const i = (py * canvas.width + px) * 4
        const match = [0, 1, 2].reduce((sum, c) => sum + Math.abs(image[i + c] - parseInt(color.slice(c * 2, c * 2 + 2), 16)), 0) < 100
          && image[i + 2] > image[i] + 16 && image[i + 1] > image[i] + 8 && image[i + 3] > 0
        if (!match) continue
        if (px >= x && px < x + w && py >= y && py < y + h) inside++
        else outside++
      }
      expect(inside).toBeGreaterThan(0)
      expect(outside).toBe(0)
    } finally { pages.dispose() }
  })

  it('the static print-page raster contains the sparkline for PDF/print export', async () => {
    const doc = await parse(), sheet = doc.sheets[0]
    const { preloadSparklines } = await import('../src/xlsx/sparklineLoader')
    await preloadSparklines()
    const m = computeMetrics(sheet), print = computePrintMetrics(sheet, m), canvas = createCanvas(print.paperPx.width, print.paperPx.height)
    renderPrintPage(sheet, canvas.getContext('2d') as unknown as CanvasRenderingContext2D, m, print)
    const [row, columns] = [...sheet.resolvedSparklines!.entries()][0], [col] = [...columns.entries()][0]
    const x = print.printable.x + m.colWidthsPx.slice(0, col).reduce((a, b) => a + b, 0) * print.scale
    const y = print.printable.y + m.rowHeightsPx.slice(0, row).reduce((a, b) => a + b, 0) * print.scale
    const host = canvas.getContext('2d').getImageData(Math.floor(x), Math.floor(y), Math.ceil(m.colWidthsPx[col] * print.scale), Math.ceil(m.rowHeightsPx[row] * print.scale)).data
    expect([...host].some((value, i) => i % 4 !== 3 && value < 240)).toBe(true)
  })

  it('renderSheet returns normally with the lazy painter loaded', async () => {
    const doc = await parse()
    const { preloadSparklines } = await import('../src/xlsx/sparklineLoader')
    await preloadSparklines()
    const m = computeMetrics(doc.sheets[0]), canvas = createCanvas(m.widthPx, m.heightPx)
    expect(() => renderSheet(doc.sheets[0], canvas.getContext('2d') as unknown as CanvasRenderingContext2D, m)).not.toThrow()
  })

  it('a viewport tile seam through a host matches one wide tile at every zoom and device scale', async () => {
    const doc = await parse(), sheet = doc.sheets[0]
    const { preloadSparklines } = await import('../src/xlsx/sparklineLoader')
    await preloadSparklines()
    // Give the host enough width that a 512-device-pixel tile can split it at every tested scale.
    sheet.cols = [...sheet.cols.filter(spec => spec.min > 9 || spec.max < 9), { min: 9, max: 9, widthChars: 180 }]
    const geometry = buildSheetGeometry(sheet)
    if (geometry.kind !== 'ready') throw new Error('expected ready geometry')
    const render = (x: number, logicalWidth: number, zoom: number, dpr: number) => {
      const scale = zoom * dpr, canvas = createCanvas(Math.ceil(logicalWidth * scale), Math.ceil(128 * scale)), ctx = canvas.getContext('2d')
      ctx.scale(scale, scale)
      renderSheetViewport(sheet, ctx as unknown as CanvasRenderingContext2D, geometry, { box: { x, y: 0, width: logicalWidth, height: 128 }, zoom, deviceScale: dpr })
      return ctx.getImageData(0, 0, canvas.width, canvas.height).data
    }
    const rasterResults: Array<{ zoom: number; dpr: number; maxDifference: number; where: string }> = []
    for (const zoom of [0.5, 1, 1.75]) for (const dpr of [1, 2]) {
      const scale = zoom * dpr, logicalTileWidth = 512 / scale
      const [row, columns] = [...sheet.resolvedSparklines!.entries()][0], [col] = [...columns.entries()][0]
      const host = geometry.metrics.sourceCellToLogical({ row, col })!
      // Half a logical px makes the tile origins fractional on the global device grid at several zoom/DPR pairs.
      const seam = host.x + host.width / 2 + 0.5, start = seam - logicalTileWidth
      const left = render(start, logicalTileWidth, zoom, dpr), right = render(seam, logicalTileWidth, zoom, dpr), whole = render(start, logicalTileWidth * 2, zoom, dpr)
      const equalSizeControl = render(start, logicalTileWidth, zoom, dpr)
      expect(equalSizeControl, `zoom=${zoom} dpr=${dpr}: repeated equal-size surface`).toEqual(left)
      const [sparkline] = [...sheet.resolvedSparklines!.get(row)!.values()]
      const scaleOptions = { zoom, dpr, output: 'device' as const }
      const leftOffset = -start * scale, rightOffset = -seam * scale
      const leftPrimitives = sparklinePrimitivesForViewport(sparkline, host, scaleOptions, leftOffset, 0)
      const rightPrimitives = sparklinePrimitivesForViewport(sparkline, host, scaleOptions, rightOffset, 0)
      const widePrimitives = sparklinePrimitivesForViewport(sparkline, host, scaleOptions, leftOffset, 0)
      const normalize = (primitives: typeof leftPrimitives, offset: number) => primitives.map(primitive => translateSparkPrimitive(primitive, -offset, 0))
      const globalLeft = normalize(leftPrimitives, leftOffset)
      expect(normalize(rightPrimitives, rightOffset), `zoom=${zoom} dpr=${dpr}: right primitives normalized to sheet device coordinates`).toEqual(globalLeft)
      expect(normalize(widePrimitives, leftOffset), `zoom=${zoom} dpr=${dpr}: wide primitives normalized to sheet device coordinates`).toEqual(globalLeft)
      const tileWidth = 512, wholeWidth = 1024, height = Math.ceil(128 * scale)
      let maxDifference = 0
      let where = ''
      // node-canvas/Cairo has extent-dependent antialiasing: identical primitives measured 10–15/255
      // off-seam between tile-sized and wide surfaces, while repeated equal-size surfaces differ by 0/255.
      for (let y = 0; y < height; y++) for (let x = 0; x < tileWidth; x++) for (let c = 0; c < 4; c++) {
        const tileIndex = (y * tileWidth + x) * 4 + c
        for (const [actual, expected, side] of [[left[tileIndex], whole[(y * wholeWidth + x) * 4 + c], 'L'], [right[tileIndex], whole[(y * wholeWidth + tileWidth + x) * 4 + c], 'R']] as const) {
          const difference = Math.abs(actual - expected)
          if (difference > maxDifference) { maxDifference = difference; where = `${side} ${x},${y},${c}=${actual}/${expected}` }
        }
      }
      rasterResults.push({ zoom, dpr, maxDifference, where })
    }
    for (const result of rasterResults) {
      expect(result.maxDifference, `zoom=${result.zoom} dpr=${result.dpr} ${result.where}`).toBeLessThanOrEqual(16)
    }
  })

  it('skips sparkline detail when the host is below MIN_DETAIL_DEVICE_PX', async () => {
    const doc = await parse(), sheet = doc.sheets[0]
    const { preloadSparklines } = await import('../src/xlsx/sparklineLoader')
    await preloadSparklines()
    const geometry = buildSheetGeometry(sheet)
    if (geometry.kind !== 'ready') throw new Error('expected ready geometry')
    const canvas = createCanvas(32, 8), ctx = canvas.getContext('2d')
    ctx.scale(0.04, 0.04)
    renderSheetViewport(sheet, ctx as unknown as CanvasRenderingContext2D, geometry, { box: { x: 560, y: 0, width: 800, height: 100 }, zoom: 0.04, deviceScale: 1 })
    const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data
    let blue = 0
    for (let i = 0; i < data.length; i += 4) if (data[i + 2] > data[i] + 16 && data[i + 1] > data[i] + 8) blue++
    expect(blue).toBe(0)
  })

  it('does no painter work for sheets without resolved sparklines', async () => {
    const original = (await import('../src/xlsx/sparklines/paint')).paintSparklinesIn
    let calls = 0
    registerSparklinePainter(((...args) => { calls++; return original(...args) }) as typeof original)
    try {
      const sheet = { name: 'plain', rows: [], cols: [], merges: [], mergeRanges: [] }
      const m = computeMetrics(sheet), canvas = createCanvas(m.widthPx, m.heightPx)
      renderSheet(sheet, canvas.getContext('2d') as unknown as CanvasRenderingContext2D, m)
      expect(calls).toBe(0)
    } finally { registerSparklinePainter(original) }
  })
})
