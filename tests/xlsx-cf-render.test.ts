// @vitest-environment node
import { createCanvas } from 'canvas'
import { expect, it } from 'vitest'
import { getPaintables } from '../src/render/paint'
import { computeMetrics } from '../src/xlsx/render'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import type { XlsxSheet } from '../src/xlsx/types'
export function cfSheet(row = 1, col = 1): XlsxSheet {
  return { name: 'CF', merges: [], mergeRanges: [], sheetId: '1', cols: [], rows: [{ index: 0, cells: Array.from({ length: 8 }, (_, col) => ({ row: 0, col, ref: '', value: col + 1, styleIndex: 0 })) }], conditionalFormatting: { diagnostics: [], rules: [{ id: 0, priority: 1, documentOrder: 0, stopIfTrue: false, ranges: [{ firstRow: row, lastRow: row, firstCol: col, lastCol: col }], writes: new Set(['fill']), dxf: { 'fill.pattern': 'solid', 'fill.bg': 'FF00FF00' }, definition: { kind: 'predicate', type: 'containsBlanks', dialect: 'base', formulas: [] } }] } }
}
it('CF-only unstored B2 paints and suppresses gridlines', async () => {
  const sheet = cfSheet(), pages = await getPaintables({ sheets: [sheet] })
  try {
    const canvas = createCanvas(pages[0].spec.widthPx, pages[0].spec.heightPx), ctx = canvas.getContext('2d')
    pages[0].paint(ctx as unknown as CanvasRenderingContext2D)
    expect([...ctx.getImageData(70, 25, 1, 1).data]).toEqual([0, 255, 0, 255])
    expect([...ctx.getImageData(64, 25, 1, 1).data]).toEqual([0, 255, 0, 255])
  } finally { pages.dispose() }
})
it('bounded Z5 expands static and interactive extent; no CF preserves metrics', async () => {
  const sheet = cfSheet(4, 25), before = computeMetrics({ ...sheet, conditionalFormatting: undefined })
  expect(before.rowHeightsPx).toHaveLength(1)
  expect(before.colWidthsPx).toHaveLength(8)
  const m = computeMetrics(sheet)
  expect(m.colWidthsPx).toHaveLength(26); expect(m.rowHeightsPx).toHaveLength(5)
  const pages = await getPaintables({ sheets: [sheet] })
  try {
    const ctx = createCanvas(pages[0].spec.widthPx, pages[0].spec.heightPx).getContext('2d')
    pages[0].paint(ctx as unknown as CanvasRenderingContext2D)
    expect([...ctx.getImageData(25 * 64 + 5, 4 * 20 + 5, 1, 1).data]).toEqual([0, 255, 0, 255])
  } finally { pages.dispose() }
  const source = createObjectWorkbookSource({ sheets: [sheet] }, 'cf-test')
  expect(source.store('1')?.extent.used).toMatchObject({ r1: 4, c1: 25 })
  const { buildSheetGeometry } = await import('../src/xlsx/view/viewport')
  const geometry = buildSheetGeometry(sheet)
  expect(geometry.kind).toBe('ready')
  if (geometry.kind === 'ready') {
    const box = geometry.metrics.sourceCellToLogical({ row: 4, col: 25 })!
    expect(box).not.toBeNull()
    expect(geometry.metrics.logicalPointToSource(box.x + 1, box.y + 1)).toMatchObject({ row: 4, col: 25 })
  }
})

it('a CF extent beyond static limits is explicitly refused, never clipped', async () => {
  const sheet = cfSheet(20000, 1)
  expect(computeMetrics(sheet).exportLimits?.some(x => x.resource === 'rows')).toBe(true)
})
it('viewport blank synthesis uses the prepared generation and geometry', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const { buildSheetGeometry } = await import('../src/xlsx/view/viewport')
  const { renderSheetViewport } = await import('../src/xlsx/view/renderViewport')
  const sheet = cfSheet(), batch = await prepareCfPaint(sheet), geometry = buildSheetGeometry(sheet)
  const canvas = createCanvas(256, 80), ctx = canvas.getContext('2d')
  try {
    renderSheetViewport(sheet, ctx as unknown as CanvasRenderingContext2D, geometry, { box: { x: 0, y: 0, width: 256, height: 80 }, zoom: 1, deviceScale: 1 }, { conditionalFormatting: batch })
    expect([...ctx.getImageData(70, 25, 1, 1).data]).toEqual([0, 255, 0, 255])
    expect([...ctx.getImageData(64, 25, 1, 1).data]).toEqual([0, 255, 0, 255])
  } finally { batch?.dispose() }
})
it('unmeasured merge formatting fails closed instead of guessing anchor semantics', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const sheet = cfSheet(0, 0); sheet.rows = []; sheet.mergeRanges = [{ minRow: 0, maxRow: 0, minCol: 0, maxCol: 2 }]
  const batch = await prepareCfPaint(sheet)
  try { expect(batch?.layer.styleAt(0, 0, undefined)?.fillColor).toBeUndefined() } finally { batch?.dispose() }
})
it('painting workspace is admitted before a batch becomes renderable and refunds on dispose', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const { createCfMemory } = await import('../src/xlsx/cf/limits')
  const memory = createCfMemory(), batch = await prepareCfPaint(cfSheet(), { memory })
  expect(batch).toBeDefined()
  // Includes >=2 KiB for cloned style, synthetic cell/placement and gridline
  // workspace, in addition to service-owned prepared state and layer framing.
  expect(batch!.layer.retainedBytes).toBeGreaterThanOrEqual(6000)
  batch!.dispose(); batch!.dispose(); expect(memory.currentBytes).toBe(0)
})
it('merged A1:C1 with hidden C spans visible A+B, and a fully hidden merge paints nothing (injected oracle)', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const { renderSheet } = await import('../src/xlsx/render')
  const { buildSheetGeometry } = await import('../src/xlsx/view/viewport')
  const { renderSheetViewport } = await import('../src/xlsx/view/renderViewport')
  for (const fullyHidden of [false, true]) {
    const sheet = cfSheet(0, 0); sheet.rows = []; sheet.mergeRanges = [{ minRow: 0, maxRow: 0, minCol: 0, maxCol: 2 }]
    sheet.cols = [{ min: fullyHidden ? 0 : 2, max: 2, hidden: true }]
    const batch = await prepareCfPaint(sheet, { evidence: { allows: () => true } })
    try { for (const interactive of [false, true]) {
      const canvas = createCanvas(192, 40), ctx = canvas.getContext('2d')
      if (interactive) renderSheetViewport(sheet, ctx as unknown as CanvasRenderingContext2D, buildSheetGeometry(sheet), { box: { x: 0, y: 0, width: 192, height: 40 }, zoom: 1, deviceScale: 1 }, { conditionalFormatting: batch })
      else renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D, undefined, undefined, { conditionalFormatting: batch })
      const pixel = [...ctx.getImageData(100, 5, 1, 1).data]
      if (fullyHidden) expect(pixel).not.toEqual([0, 255, 0, 255])
      else expect(pixel).toEqual([0, 255, 0, 255])
    } } finally { batch?.dispose() }
  }
})
it('tile and full render have exactly equal source-space fill primitives at six scales', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const { renderSheet } = await import('../src/xlsx/render')
  const { buildSheetGeometry } = await import('../src/xlsx/view/viewport')
  const { renderSheetViewport } = await import('../src/xlsx/view/renderViewport')
  const sheet = cfSheet(); sheet.cols = [{ min: 1, max: 1, widthChars: 400 }]
  const batch = await prepareCfPaint(sheet), geometry = buildSheetGeometry(sheet)
  try { for (const zoom of [.5, 1, 1.75]) for (const dpr of [1, 2]) {
    const scale = zoom * dpr, seam = 512 / scale
    const record = (x: number, width: number, staticPaint = false) => {
      const canvas = createCanvas(Math.ceil(width * scale), Math.ceil(60 * scale)), ctx = canvas.getContext('2d'); ctx.scale(scale, scale)
      const fills: number[][] = [], original = ctx.fillRect.bind(ctx)
      ctx.fillRect = (a, b, w, h) => {
        if (String(ctx.fillStyle).toLowerCase() === '#00ff00') { const t = ctx.getTransform(); fills.push([(t.a * a + t.e + Math.round(x * scale)) / scale, (t.d * b + t.f) / scale, t.a * w / scale, t.d * h / scale]) }
        original(a, b, w, h)
      }
      if (staticPaint) renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D, undefined, undefined, { conditionalFormatting: batch })
      else renderSheetViewport(sheet, ctx as unknown as CanvasRenderingContext2D, geometry, { box: { x, y: 0, width, height: 60 }, zoom, deviceScale: dpr }, { conditionalFormatting: batch })
      return fills
    }
    const wide = record(0, seam * 2)
    expect(wide).toHaveLength(1)
    expect(record(0, seam)).toEqual(wide); expect(record(seam, seam)).toEqual(wide)
    expect(record(0, seam * 2, true)).toEqual(wide)
  } } finally { batch?.dispose() }
}, 60_000)
it('layout cache keys include the prepared CF generation', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const { buildSheetGeometry } = await import('../src/xlsx/view/viewport')
  const { renderSheetViewport } = await import('../src/xlsx/view/renderViewport')
  const { LayoutCache } = await import('../src/xlsx/view/textLayout')
  const sheet = cfSheet(), cache = new LayoutCache(), geometry = buildSheetGeometry(sheet)
  const ctx = createCanvas(128, 40).getContext('2d') as unknown as CanvasRenderingContext2D
  const request = (cfGen: number) => ({ stamp: { owner: 'test', sheetId: '1', workbookValueGen: 0, cfGen, viewGen: 0, structureGen: 0, clockGen: 0 }, sourceRects: sheet.conditionalFormatting!.rules[0].ranges, signal: new AbortController().signal, grantBytes: 1_000_000 })
  const first = await prepareCfPaint(sheet, { request: request(1) }), second = await prepareCfPaint(sheet, { request: request(2) })
  try {
    const paint = (batch: typeof first) => renderSheetViewport(sheet, ctx, geometry, { box: { x: 0, y: 20, width: 128, height: 20 }, zoom: 1, deviceScale: 1 }, { conditionalFormatting: batch }, { cache })
    paint(first); const built = cache.stats.built
    paint(first); expect(cache.stats.built).toBe(built)
    paint(second); expect(cache.stats.built).toBeGreaterThan(built)
  } finally { first?.dispose(); second?.dispose() }
})
it('paged preparation performs reads before painting, and released batches cannot retain CF', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const { createCfMemory } = await import('../src/xlsx/cf/limits')
  const { renderSheet } = await import('../src/xlsx/render')
  const sheet = cfSheet(), stamp = { owner: 'paged', sheetId: '1', workbookValueGen: 1, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }
  let reads = 0, released = 0
  const memory = createCfMemory()
  const batch = await prepareCfPaint(sheet, { memory, request: { stamp, sourceRects: sheet.conditionalFormatting!.rules[0].ranges, signal: new AbortController().signal, grantBytes: 1_000_000 }, source: {
    stamp: () => stamp, bounds: async () => ({ firstRow: 0, lastRow: 1, firstCol: 0, lastCol: 7 }),
    async *pages() { reads++; yield { value: { rows: new Uint32Array(0), cols: new Uint16Array(0), values: [] }, bytes: 0, release() { released++ }, transfer() { return { kind: 'unresolved', reason: 'test' } } } },
  } })
  expect(reads).toBe(1); expect(released).toBe(1)
  const canvas = createCanvas(512, 40), ctx = canvas.getContext('2d')
  renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D, undefined, undefined, { conditionalFormatting: batch })
  expect([...ctx.getImageData(70, 25, 1, 1).data]).toEqual([0, 255, 0, 255]); expect(reads).toBe(1)
  batch?.dispose(); expect(memory.currentBytes).toBe(0)
  expect(batch?.layer.styleAt(1, 1, undefined)?.fillColor).toBeUndefined()
})
it('zero memory refuses before reading values, and whole-column declarations do not expand a raster', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const { createCfMemory } = await import('../src/xlsx/cf/limits')
  const sheet = cfSheet(); let reads = 0
  Object.defineProperty(sheet.rows[0].cells[0], 'value', { get() { reads++; return 1 } })
  expect(await prepareCfPaint(sheet, { memory: createCfMemory({ limitBytes: 0 }) })).toBeUndefined()
  expect(reads).toBe(0)
  const rule = sheet.conditionalFormatting!.rules[0]
  sheet.conditionalFormatting = { diagnostics: [], rules: [{ ...rule, ranges: [{ firstRow: 0, lastRow: 1048575, firstCol: 1, lastCol: 1 }] }] }
  expect(computeMetrics(sheet).rowHeightsPx).toHaveLength(1)
})
it('cold render bounds scans always receive a cooperative budget', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const sheet = cfSheet()
  const rule = sheet.conditionalFormatting!.rules[0]
  sheet.conditionalFormatting = { diagnostics: [], rules: [{ ...rule, ranges: [{ firstRow: 0, lastRow: 1048575, firstCol: 1, lastCol: 1 }] }] }
  let unbudgeted = 0
  const batch = await prepareCfPaint(sheet, { source: {
    stamp: () => ({ owner: 'static', sheetId: '1', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }),
    async bounds(_id, budget) { if (!budget) unbudgeted++; return { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 7 } },
    async *pages() {},
  } })
  try { expect(batch).toBeDefined(); expect(unbudgeted).toBe(0) } finally { batch?.dispose() }
})
it('CF gridline lookup does not index styles outside the admitted paint ranges', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const { renderSheet } = await import('../src/xlsx/render')
  const sheet = cfSheet(); let outsideStyleReads = 0
  for (let row = 100; row < 200; row++) {
    const cell = { row, col: 0, ref: '', value: 0, styleIndex: 0 }
    Object.defineProperty(cell, 'style', { get() { outsideStyleReads++; return { numFmtId: 0 } } })
    sheet.rows.push({ index: row, cells: [cell] })
  }
  const batch = await prepareCfPaint(sheet), ctx = createCanvas(64, 20).getContext('2d')
  try {
    renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D, undefined, undefined, { conditionalFormatting: batch }, { x: 64, y: 20, width: 64, height: 20 })
    expect(outsideStyleReads).toBe(0)
  } finally { batch?.dispose() }
})
it('finite CF hosts also respect the static raster cap', async () => {
  const sheet = cfSheet(1000, 1000)
  expect(computeMetrics(sheet).exportLimits?.some(x => x.resource === 'pixels' || x.resource === 'raster-bytes')).toBe(true)
  const { renderSheet } = await import('../src/xlsx/render')
  expect(() => renderSheet(sheet, createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D)).toThrow()
})
it('coarse gridline resolution preserves a CF fill over an own fill', async () => {
  const { regionGrid } = await import('../src/xlsx/regionGrid')
  const grid = regionGrid(cfSheet(), () => ({ numFmtId: 0, fillColor: 'FFFF0000' }), {
    retainedBytes: 0, rects: [{ minRow: 1, maxRow: 1, minCol: 1, maxCol: 1 }], styleAt: () => ({ numFmtId: 0, fillColor: 'FF00FF00' }),
  })!
  expect(grid.keyedLine({ row: 1, col: 1, own: { numFmtId: 0, fillColor: 'FFFF0000' }, rowKey: undefined, colKey: undefined }, { row: 1, col: 2, own: undefined, rowKey: undefined, colKey: undefined })).toBe('FF00FF00')
})
it('oversized paged rectangle requests refuse before copying caller arrays', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  let reads = 0
  const sourceRects = new Proxy(Array.from({ length: 16385 }, () => ({ firstRow: 1, lastRow: 1, firstCol: 1, lastCol: 1 })), {
    get(target, key, receiver) { if (typeof key === 'string' && /^\d+$/.test(key)) reads++; return Reflect.get(target, key, receiver) },
  })
  const batch = await prepareCfPaint(cfSheet(), { request: { sourceRects, signal: new AbortController().signal, grantBytes: 8 * 1024 ** 2,
    stamp: { owner: 'static', sheetId: '1', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 },
  } })
  batch?.dispose()
  expect(reads).toBe(0)
})
