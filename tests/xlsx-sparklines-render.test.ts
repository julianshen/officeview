// @vitest-environment node
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { getPaintables } from '../src/render/paint'
import { computeMetrics, computePrintMetrics, renderPrintPage, renderSheet } from '../src/xlsx/render'
import { parseXlsx } from '../src/xlsx/parse'
import { buildSheetGeometry, planViewport } from '../src/xlsx/view/viewport'
import { createRowOrder } from '../src/xlsx/view/order'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { viewportSparklineCandidates } from '../src/xlsx/view/renderViewport'
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

  it('static export includes and paints a sparkline host beyond populated cells', async () => {
    const doc = await parse(), sheet = doc.sheets[0]
    const { preloadSparklines } = await import('../src/xlsx/sparklineLoader')
    await preloadSparklines()
    const sparkline = [...sheet.resolvedSparklines!.values()][0].values().next().value!
    const host = { row: 4, col: 25 }
    sheet.rows = [{ index: 0, cells: Array.from({ length: 8 }, (_, col) => ({
      ref: `${String.fromCharCode(65 + col)}1`, row: 0, col, value: col + 1, styleIndex: 0,
    })) }]
    sheet.mergeRanges = [{ minRow: 1, minCol: 1, maxRow: 7, maxCol: 6 }]
    sheet.resolvedSparklines = new Map([[host.row, new Map([[host.col, { ...sparkline, host }]])]])

    const source = createObjectWorkbookSource(doc, 'sparkline-host-extent')
    try {
      expect(source.store(source.sheets[0].sheetId).extent.used).toEqual({ r0: 0, r1: host.row, c0: 0, c1: host.col })
    } finally { await source.release() }

    const pages = await getPaintables(doc)
    try {
      const page = pages[0]
      expect(page.spec.widthPx).toBeGreaterThan(0)
      const canvas = createCanvas(page.spec.widthPx, page.spec.heightPx), ctx = canvas.getContext('2d')
      page.paint(ctx as unknown as CanvasRenderingContext2D)
      const metrics = computeMetrics(sheet)
      expect(metrics.colWidthsPx.length).toBeGreaterThan(host.col)
      expect(metrics.rowHeightsPx.length).toBeGreaterThan(host.row)
      const x = metrics.colWidthsPx.slice(0, host.col).reduce((sum, width) => sum + width, 0)
      const y = metrics.rowHeightsPx.slice(0, host.row).reduce((sum, height) => sum + height, 0)
      const width = metrics.colWidthsPx[host.col], height = metrics.rowHeightsPx[host.row]
      const pixels = ctx.getImageData(x, y, width, height).data
      const color = sparkline.group.colors.series?.slice(2).toUpperCase() ?? '000000'
      let ink = 0
      for (let i = 0; i < pixels.length; i += 4) {
        const matches = [0, 1, 2].reduce((sum, channel) => sum + Math.abs(pixels[i + channel] - parseInt(color.slice(channel * 2, channel * 2 + 2), 16)), 0) < 100
        if (matches && pixels[i + 2] > pixels[i] + 16 && pixels[i + 1] > pixels[i] + 8 && pixels[i + 3] > 0) ink++
      }
      expect(ink).toBeGreaterThan(0)
    } finally { pages.dispose() }
  })

  it('renderSheet returns normally with the lazy painter loaded', async () => {
    const doc = await parse()
    const { preloadSparklines } = await import('../src/xlsx/sparklineLoader')
    await preloadSparklines()
    const m = computeMetrics(doc.sheets[0]), canvas = createCanvas(m.widthPx, m.heightPx)
    expect(() => renderSheet(doc.sheets[0], canvas.getContext('2d') as unknown as CanvasRenderingContext2D, m)).not.toThrow()
  })

  it('derives sparkline zoom from the canvas transform without a public renderSheet parameter', async () => {
    const doc = await parse(), sheet = doc.sheets[0]
    const original = (await import('../src/xlsx/sparklines/paint')).paintSparklinesIn
    let seen: number | undefined
    registerSparklinePainter(((...args) => { seen = args[4].zoom; return true }) as typeof original)
    try {
      const m = computeMetrics(sheet), canvas = createCanvas(Math.ceil(m.widthPx * 0.75), Math.ceil(m.heightPx * 0.75)), ctx = canvas.getContext('2d')
      ctx.scale(0.75, 0.75)
      renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D, m)
      expect(seen).toBe(0.75)
      expect(renderSheet.length).toBe(6)
    } finally { registerSparklinePainter(original) }
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

  it('selects hosts by sorted display slots and maps them back to source rows', async () => {
    const doc = await parse(), sheet = doc.sheets[0]
    const [hostRow, columns] = [...sheet.resolvedSparklines!.entries()][0], [hostCol] = [...columns.keys()]
    const lastSlot = Math.max(2, hostRow)
    const displayOrder = Array.from({ length: lastSlot + 1 }, (_, row) => row).filter(row => row !== hostRow)
    displayOrder.push(hostRow)
    const geometry = buildSheetGeometry(sheet, { order: createRowOrder([{ r0: 0, r1: lastSlot, sourceAtOffset: new Uint32Array(displayOrder) }]) })
    if (geometry.kind !== 'ready') throw new Error('expected ready geometry')
    const height = geometry.metrics.rows.start(lastSlot)
    const plan = { kind: 'detail', merges: [] } as unknown as import('../src/xlsx/view/viewport').ViewportPlan
    const candidates = viewportSparklineCandidates(sheet, geometry, { box: { x: 0, y: 0, width: geometry.metrics.cols.totalCss(), height }, zoom: 1, deviceScale: 1 }, plan)
    expect(geometry.metrics.order.sourceToDisplay(hostRow)).toBe(lastSlot)
    expect(candidates.get(hostRow)?.has(hostCol) ?? false).toBe(false)
  })

  it('includes a merged host anchor in a tile that intersects only the far side of its merge', async () => {
    const doc = await parse(), sheet = doc.sheets[0]
    const [row, columns] = [...sheet.resolvedSparklines!.entries()][0], [col] = [...columns.keys()]
    sheet.mergeRanges.push({ minRow: row, minCol: col, maxRow: row, maxCol: col + 1 })
    const geometry = buildSheetGeometry(sheet)
    if (geometry.kind !== 'ready') throw new Error('expected ready geometry')
    const anchor = geometry.metrics.sourceCellToLogical({ row, col })!
    const farCell = geometry.metrics.sourceCellToLogical({ row, col: col + 1 })!
    const split = anchor.x + anchor.width
    const viewport = { box: { x: split, y: anchor.y, width: farCell.width, height: 100 }, zoom: 1, deviceScale: 1 }
    const plan = planViewport(geometry, viewport)
    expect(plan.kind).toBe('detail')
    if (plan.kind !== 'detail') throw new Error('expected detail plan')
    expect(plan.merges.some(merge => merge.range.minRow === row && merge.range.minCol === col)).toBe(true)
    const visible = viewportSparklineCandidates(sheet, geometry, viewport, plan)
    expect(visible.get(row)?.has(col)).toBe(true)
    const leftViewport = { box: { x: anchor.x, y: anchor.y, width: anchor.width, height: 100 }, zoom: 1, deviceScale: 1 }
    const leftPlan = planViewport(geometry, leftViewport)
    expect(leftPlan.kind).toBe('detail')
    if (leftPlan.kind !== 'detail') throw new Error('expected detail plan')
    expect(viewportSparklineCandidates(sheet, geometry, leftViewport, leftPlan).get(row)?.has(col)).toBe(true)
    let rectCalls = 0
    const painter = (await import('../src/xlsx/sparklines/paint')).paintSparklinesIn
    painter(createCanvas(128, 128).getContext('2d') as unknown as CanvasRenderingContext2D, sheet, () => { rectCalls++; return { x: -10, y: 0, width: 100, height: 20 } }, { r0: 99, c0: 99, r1: 100, c1: 100 }, { zoom: 1, dpr: 1 }, visible)
    expect(rectCalls).toBe(1)
  })

  it('omits a sparkline whose data crosses a viewer-filter-hidden row', async () => {
    const doc = await parse(), sheet = doc.sheets[0]
    const [row, columns] = [...sheet.resolvedSparklines!.entries()][0], [col, sparkline] = [...columns.entries()][0]
    const geometry = buildSheetGeometry(sheet)
    if (geometry.kind !== 'ready') throw new Error('expected ready geometry')
    const hiddenSourceRow = sparkline.sourceRows[0]
    const views = new Map([[sheet.name.toLowerCase(), {
      order: geometry.metrics.order,
      visibility: { visible: (slot: number) => slot !== hiddenSourceRow },
      baseHidden: { contains: () => false },
      rows: geometry.metrics.rows,
    }]])
    const plan = { kind: 'detail', merges: [] } as unknown as import('../src/xlsx/view/viewport').ViewportPlan
    const candidates = viewportSparklineCandidates(sheet, geometry, { box: { x: 0, y: 0, width: geometry.metrics.cols.totalCss(), height: geometry.metrics.rows.totalCss() }, zoom: 1, deviceScale: 1 }, plan, views)
    expect(candidates.get(row)?.has(col) ?? false).toBe(false)
    expect(sheet.diagnostics?.filter(item => item.feature === 'sparkline-filter-hidden')).toHaveLength(1)
    const authoredHidden = new Map([[sheet.name.toLowerCase(), {
      order: geometry.metrics.order, visibility: { visible: (slot: number) => slot !== hiddenSourceRow },
      baseHidden: { contains: (source: number) => source === hiddenSourceRow },
    }]])
    const authoredCandidate = viewportSparklineCandidates(sheet, geometry, { box: { x: 0, y: 0, width: geometry.metrics.cols.totalCss(), height: geometry.metrics.rows.totalCss() }, zoom: 1, deviceScale: 1 }, plan, authoredHidden)
    expect(authoredCandidate.get(row)?.has(col)).toBe(true)
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
