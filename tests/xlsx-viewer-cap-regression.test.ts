import { describe, expect, test } from 'vitest'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { computeMetrics, computePrintMetrics, mergeEdges, renderPrintPage, renderSheet, renderSheetViewport, STATIC_SHEET_BUDGET, StaticExportLimitError } from '../src/xlsx/render'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import { getPaintables } from '../src/render/paint'
import { AXIS_ERROR_DIAGNOSTICS, buildSheetGeometry, planViewport, sheetBaseHidden, VIEWPORT_DETAIL_LIMITS, type SheetGeometry } from '../src/xlsx/view/viewport'
import { createRowOrder } from '../src/xlsx/view/order'
import { createIntervalSet, createRowVisibility, visibilityFromHidden } from '../src/xlsx/view/visibility'
import type { AxisMetricsErrorCode } from '../src/xlsx/view/axes'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import type { LogicalViewport } from '../src/xlsx/view/contracts'
import type { XlsxCell, XlsxColumnSpec, XlsxMergeRange, XlsxRow, XlsxSheet } from '../src/xlsx/types'

// Default 10pt text sits on y + 20 - (20 - 13.33)/2 inside a 20px row; strings start at x + 3.
const BASELINE = 20 - (20 - 10 * 96 / 72) / 2
const PAD = 3

function cell(row: number, col: number, value: XlsxCell['value']): XlsxCell {
  return { ref: `R${row}C${col}`, row, col, value, styleIndex: 0 }
}
function sheetOf(rows: XlsxRow[], cols: XlsxColumnSpec[] = []): XlsxSheet {
  return { name: 'Cap', rows, cols, merges: [], mergeRanges: [] }
}
const beyondCapSheet = () => sheetOf([
  { index: 0, cells: [cell(0, 0, 'Origin')] },
  { index: 16_385, cells: [cell(16_385, 0, 'BeyondCap')] },
])

interface Painted { text: string; x: number; y: number }
/** node-canvas context that records fillText in absolute output pixels. Tracks
 * translate/scale through save/restore itself instead of relying on getTransform. */
function recorder(width = 320, height = 200) {
  const canvas = createCanvas(width, height)
  const raw = canvas.getContext('2d')
  const texts: Painted[] = []
  let t = { tx: 0, ty: 0, sx: 1, sy: 1 }
  const stack: typeof t[] = []
  const ctx = new Proxy(raw, {
    get(target, property) {
      if (property === 'save') return () => { stack.push({ ...t }); target.save() }
      if (property === 'restore') return () => { t = stack.pop() ?? t; target.restore() }
      if (property === 'translate') return (x: number, y: number) => { t.tx += x * t.sx; t.ty += y * t.sy; target.translate(x, y) }
      if (property === 'scale') return (x: number, y: number) => { t.sx *= x; t.sy *= y; target.scale(x, y) }
      if (property === 'fillText') return (text: string, x: number, y: number) => { texts.push({ text, x: t.tx + x * t.sx, y: t.ty + y * t.sy }); target.fillText(text, x, y) }
      const value = Reflect.get(target, property, target)
      return typeof value === 'function' ? value.bind(target) : value
    },
    set(target, property, value) { return Reflect.set(target, property, value, target) },
  }) as unknown as CanvasRenderingContext2D
  const ink = () => {
    const data = raw.getImageData(0, 0, width, height).data
    let count = 0
    for (let i = 0; i < data.length; i += 4) if (data[i] < 200) count++
    return count
  }
  return { ctx, texts, ink, find: (text: string) => texts.filter(item => item.text === text) }
}
function view(x: number, y: number, width = 320, height = 200): LogicalViewport {
  return { box: { x, y, width, height }, zoom: 1, deviceScale: 1 }
}
function ready(sheet: XlsxSheet, options?: Parameters<typeof buildSheetGeometry>[1]): SheetGeometry {
  const geometry = buildSheetGeometry(sheet, options)
  if (geometry.kind !== 'ready') throw new Error(`geometry not ready: ${JSON.stringify(geometry.diagnostics)}`)
  return geometry
}

describe('static path never aliases a capped coordinate to the origin', () => {
  test('row 16385 is refused, not painted on the row-zero baseline', () => {
    const { ctx, find } = recorder()
    let error: unknown
    try {
      renderSheet(beyondCapSheet(), ctx)
    } catch (caught) { error = caught }
    expect(find('BeyondCap')).toEqual([])
    expect(error).toBeInstanceOf(StaticExportLimitError)
    expect(error).toMatchObject({ code: 'large-export-limit' })
  })
  test('computeMetrics reports the retained grid instead of silently clamping', () => {
    const m = computeMetrics(beyondCapSheet())
    expect(m.rowHeightsPx).toHaveLength(16_384)
    expect(m.exportLimits).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'large-export-limit', resource: 'rows', requested: 16_386, limit: 16_384 })]))
    const edge = computeMetrics(sheetOf([{ index: 0, cells: [cell(0, MAX_COLS - 1, 'EdgeCol')] }]))
    expect(edge.exportLimits).toEqual(expect.arrayContaining([expect.objectContaining({ resource: 'cols', requested: MAX_COLS, limit: 4096 })]))
    expect(computeMetrics(sheetOf([{ index: 0, cells: [cell(0, 0, 'A')] }])).exportLimits).toBeUndefined()
  })
  test('caller-supplied metrics that do not cover a cell are refused too', () => {
    const { ctx, find } = recorder()
    const sheet = sheetOf([{ index: 0, cells: [cell(0, 0, 'A'), cell(0, 3, 'Outside')] }])
    expect(() => renderSheet(sheet, ctx, { colWidthsPx: [64], rowHeightsPx: [20], widthPx: 64, heightPx: 20 })).toThrow(StaticExportLimitError)
    expect(find('Outside')).toEqual([])
  })
  test('getPaintables returns a typed refusal unit instead of a partial bitmap', async () => {
    const sheet = beyondCapSheet()
    const units = await getPaintables({ sheets: [sheet, sheetOf([{ index: 0, cells: [cell(0, 0, 'Fine')] }])] })
    try {
      expect(units[0].refusal).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'large-export-limit', resource: 'rows' })]))
      expect(units[1].refusal).toBeUndefined()
      const { ctx, find, texts } = recorder(units[0].spec.widthPx, units[0].spec.heightPx)
      units[0].paint(ctx)
      expect(find('BeyondCap')).toEqual([])
      expect(find('Origin')).toEqual([])
      expect(texts.some(item => /too large/i.test(item.text))).toBe(true)
      expect(sheet.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'large-export-limit', feature: 'large-export-limit' })]))
    } finally { units.dispose() }
  })
})

describe('interactive seam paints full logical coordinates', () => {
  test('BeyondCap is absent at the origin viewport and visible at its true offset', () => {
    const sheet = beyondCapSheet()
    const geometry = ready(sheet)
    const origin = recorder()
    expect(renderSheetViewport(sheet, origin.ctx, geometry, view(0, 0)).kind).toBe('detail')
    expect(origin.find('BeyondCap')).toEqual([])
    expect(origin.find('Origin')).toEqual([{ text: 'Origin', x: PAD, y: expect.closeTo(BASELINE, 6) }])

    const top = geometry.metrics.rows.start(16_385)
    expect(top).toBe(16_385 * 20)
    const far = recorder()
    renderSheetViewport(sheet, far.ctx, geometry, view(0, top - 40))
    expect(far.find('Origin')).toEqual([])
    expect(far.find('BeyondCap')).toEqual([{ text: 'BeyondCap', x: PAD, y: expect.closeTo(40 + BASELINE, 6) }])
  })
  test('a sorted band shows BeyondCap only at its display slot', () => {
    const sheet = beyondCapSheet()
    const permutation = new Uint32Array(16_386).map((_, slot) => slot)
    permutation[1] = 16_385
    permutation[16_385] = 1
    const geometry = ready(sheet, { order: createRowOrder([{ r0: 0, r1: 16_385, sourceAtOffset: permutation }]) })
    const origin = recorder()
    renderSheetViewport(sheet, origin.ctx, geometry, view(0, 0))
    expect(origin.find('BeyondCap')).toEqual([{ text: 'BeyondCap', x: PAD, y: expect.closeTo(20 + BASELINE, 6) }])
    const far = recorder()
    renderSheetViewport(sheet, far.ctx, geometry, view(0, 16_385 * 20 - 40))
    expect(far.find('BeyondCap')).toEqual([])
  })
  test('edge column XFD paints only at its own column offset', () => {
    const sheet = sheetOf([{ index: 0, cells: [cell(0, 0, 'A1'), cell(0, MAX_COLS - 1, 'EdgeCol')] }])
    const geometry = ready(sheet)
    const origin = recorder()
    renderSheetViewport(sheet, origin.ctx, geometry, view(0, 0))
    expect(origin.find('EdgeCol')).toEqual([])
    const left = geometry.metrics.cols.start(MAX_COLS - 1)
    expect(left).toBe((MAX_COLS - 1) * 64)
    const edge = recorder()
    renderSheetViewport(sheet, edge.ctx, geometry, view(left - 100, 0))
    expect(edge.find('A1')).toEqual([])
    expect(edge.find('EdgeCol')).toEqual([{ text: 'EdgeCol', x: 100 + PAD, y: expect.closeTo(BASELINE, 6) }])
  })
  test('edge cell XFD1048576 paints only at the far corner', () => {
    const sheet = sheetOf([
      { index: 0, cells: [cell(0, 0, 'A1')] },
      { index: MAX_ROWS - 1, cells: [cell(MAX_ROWS - 1, MAX_COLS - 1, 'Corner')] },
    ])
    const geometry = ready(sheet)
    const origin = recorder()
    renderSheetViewport(sheet, origin.ctx, geometry, view(0, 0))
    expect(origin.find('Corner')).toEqual([])
    const x = geometry.metrics.cols.start(MAX_COLS - 1), y = geometry.metrics.rows.start(MAX_ROWS - 1)
    expect([x + 64, y + 20]).toEqual([geometry.metrics.cols.totalCss(), geometry.metrics.rows.totalCss()])
    const corner = recorder()
    renderSheetViewport(sheet, corner.ctx, geometry, view(x - 200, y - 100))
    expect(corner.find('A1')).toEqual([])
    expect(corner.find('Corner')).toEqual([{ text: 'Corner', x: 200 + PAD, y: expect.closeTo(100 + BASELINE, 6) }])
  })
  test('no geometry is pending, never an origin paint', () => {
    const sheet = beyondCapSheet()
    expect(planViewport(undefined, view(0, 0))).toMatchObject({ kind: 'pending', diagnostics: [expect.objectContaining({ code: 'geometry-pending' })] })
    const out = recorder()
    const result = renderSheetViewport(sheet, out.ctx, undefined, view(0, 0))
    expect(result.kind).toBe('pending')
    expect(out.find('Origin')).toEqual([])
    expect(out.find('BeyondCap')).toEqual([])
  })
  test('a frame beyond the detail budget is coarse before any cell enumeration', () => {
    const sheet = beyondCapSheet()
    const geometry = ready(sheet)
    const tall = VIEWPORT_DETAIL_LIMITS.D.gridlines * 20 + 200
    expect(planViewport(geometry, view(0, 0, 320, tall))).toMatchObject({ kind: 'coarse', diagnostics: [expect.objectContaining({ code: 'detail-limit' })] })
    const mobile = VIEWPORT_DETAIL_LIMITS.M.gridlines * 20
    expect(planViewport(geometry, view(0, 0, 320, mobile), 'M').kind).toBe('coarse')
    expect(planViewport(geometry, view(0, 0, 320, 200), 'M').kind).toBe('detail')
  })
})

describe('typed AxisMetricsError codes map to diagnosed fallbacks at the renderer seam', () => {
  test('every code has a stable diagnostic', () => {
    const codes: AxisMetricsErrorCode[] = ['CONTRADICTORY_OVERRIDES', 'PRECISION_EXCEEDED', 'TOTAL_NOT_FINITE', 'INVALID_METADATA', 'INVALID_QUERY']
    const mapped = codes.map(code => AXIS_ERROR_DIAGNOSTICS[code])
    expect(new Set(mapped).size).toBe(codes.length)
    for (const code of mapped) expect(code).toMatch(/^axis-/)
  })
  const fallbackCases: Array<[string, XlsxSheet, string]> = [
    ['contradictory overlapping column widths', sheetOf([{ index: 0, cells: [cell(0, 1, 'B1')] }], [
      { min: 0, max: 2, widthChars: 10 }, { min: 1, max: 1, widthChars: 20 },
    ]), 'axis-contradictory-overrides'],
    ['an unrepresentable row total', sheetOf([
      { index: 0, heightPt: 1e308, customHeight: true, cells: [] },
      { index: 1, heightPt: 1e308, customHeight: true, cells: [cell(1, 1, 'B1')] },
    ]), 'axis-total-not-finite'],
    ['an imprecise row origin', sheetOf([
      { index: 0, heightPt: 0.75e18, customHeight: true, cells: [] },
      { index: 1, cells: [cell(1, 1, 'B1')] },
    ]), 'axis-precision-exceeded'],
  ]
  test.each(fallbackCases)('%s falls back to default sizes with a diagnostic', (_name, sheet, code) => {
    let geometry: ReturnType<typeof buildSheetGeometry> | undefined
    expect(() => { geometry = buildSheetGeometry(sheet) }).not.toThrow()
    expect(geometry!.kind).toBe('ready')
    expect(geometry!.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code })]))
    const out = recorder()
    const result = renderSheetViewport(sheet, out.ctx, geometry, view(0, 0))
    expect(result.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code })]))
    // Default geometry: the cell is painted where an unsized sheet would put it.
    const painted = out.find('B1')
    expect(painted).toHaveLength(1)
    expect(painted[0].x).toBe(64 + PAD)
  })
  test('invalid projection metadata is a diagnostic geometry, painted as a notice', () => {
    const sheet = beyondCapSheet()
    const geometry = buildSheetGeometry(sheet, { visibility: createRowVisibility(10) })
    expect(geometry).toMatchObject({ kind: 'diagnostic', diagnostics: [expect.objectContaining({ code: 'axis-invalid-metadata' })] })
    const out = recorder()
    let result: ReturnType<typeof renderSheetViewport> | undefined
    expect(() => { result = renderSheetViewport(sheet, out.ctx, geometry, view(0, 0)) }).not.toThrow()
    expect(result!.kind).toBe('diagnostic')
    expect(out.find('Origin')).toEqual([])
    expect(out.ink()).toBeGreaterThan(0)
  })
  test('an invalid query viewport is diagnosed, not thrown', () => {
    const sheet = beyondCapSheet()
    const geometry = ready(sheet)
    expect(planViewport(geometry, view(Number.NaN, 0))).toMatchObject({ kind: 'diagnostic', diagnostics: [expect.objectContaining({ code: 'axis-invalid-query' })] })
    const out = recorder()
    expect(renderSheetViewport(sheet, out.ctx, geometry, view(0, Number.POSITIVE_INFINITY)).kind).toBe('diagnostic')
  })
})

describe('authored-hidden rows and columns collapse unconditionally', () => {
  const hiddenSheet = () => sheetOf([
    { index: 0, cells: [cell(0, 0, 'A1'), cell(0, 1, 'HiddenCol'), cell(0, 2, 'C1')] },
    { index: 1, hidden: true, cells: [cell(1, 0, 'HiddenRow')] },
    { index: 2, cells: [cell(2, 0, 'A3')] },
  ], [{ min: 1, max: 1, widthChars: 10, hidden: true }])
  test('static metrics collapse hidden rows/cols on a sheet without drawings', () => {
    const sheet = hiddenSheet()
    const m = computeMetrics(sheet)
    expect(m.rowHeightsPx).toEqual([20, 0, 20])
    expect(m.colWidthsPx).toEqual([64, 0, 64])
    const out = recorder()
    renderSheet(sheet, out.ctx, m)
    expect(out.find('HiddenRow')).toEqual([])
    expect(out.find('HiddenCol')).toEqual([])
    expect(out.find('A3')).toEqual([{ text: 'A3', x: PAD, y: expect.closeTo(20 + BASELINE, 6) }])
    expect(out.find('C1')).toEqual([{ text: 'C1', x: 64 + PAD, y: expect.closeTo(BASELINE, 6) }])
  })
  test('the interactive seam collapses the same rows/cols', () => {
    const sheet = hiddenSheet()
    const geometry = ready(sheet)
    expect(geometry.metrics.rows.size(1)).toBe(0)
    expect(geometry.metrics.cols.size(1)).toBe(0)
    const out = recorder()
    renderSheetViewport(sheet, out.ctx, geometry, view(0, 0))
    expect(out.find('HiddenRow')).toEqual([])
    expect(out.find('HiddenCol')).toEqual([])
    expect(out.find('A3')).toEqual([{ text: 'A3', x: PAD, y: expect.closeTo(20 + BASELINE, 6) }])
    expect(out.find('C1')).toEqual([{ text: 'C1', x: 64 + PAD, y: expect.closeTo(BASELINE, 6) }])
  })
  test('conflicting duplicate row records keep the row visible and are diagnosed', () => {
    // Rows without r attributes currently all parse at index 0 (poi-59746_NoRowNums.xlsx):
    // one hidden record must not hide visible records sharing that index.
    const sheet = sheetOf([
      { index: 0, cells: [cell(0, 0, 'Visible')] },
      { index: 0, hidden: true, cells: [cell(0, 1, 'Duplicate')] },
    ])
    const m = computeMetrics(sheet)
    expect(m.rowHeightsPx).toEqual([20])
    expect(m.diagnostics).toEqual([expect.objectContaining({ code: 'conflicting-row-metadata' })])
    const geometry = ready(sheet)
    expect(geometry.metrics.rows.size(0)).toBe(20)
    expect(geometry.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'conflicting-row-metadata' })]))
  })
})

describe('review round 1 (Codex findings)', () => {
  async function parsedSheet(rows: Parameters<typeof buildXlsx>[0][0]['rows'], patch: (xml: string) => string = xml => xml, extra: Partial<Parameters<typeof buildXlsx>[0][0]> = {}): Promise<XlsxSheet> {
    const bytes = await buildXlsx([{ name: 'S', rows, ...extra }])
    const zip = await JSZip.loadAsync(bytes)
    const path = 'xl/worksheets/sheet1.xml'
    zip.file(path, patch(await zip.file(path)!.async('string')))
    return (await parseXlsx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))).sheets[0]
  }
  test('F1: a <row> without r keeps its cell at the explicit A16386 reference', async () => {
    const sheet = await parsedSheet([
      { r: 1, cells: [{ ref: 'A1', v: 1 }] },
      { r: 16386, cells: [{ ref: 'A16386', v: 7 }] },
    ], xml => xml.replace('<row r="16386">', '<row>'))
    expect(sheet.rows[1].index).toBe(0)
    expect(sheet.rows[1].cells[0].row).toBe(16_385)
    const geometry = ready(sheet)
    expect(geometry.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'inconsistent-row-coordinate' })]))
    const origin = recorder()
    renderSheetViewport(sheet, origin.ctx, geometry, view(0, 0))
    expect(origin.find('7')).toEqual([])
    const far = recorder()
    renderSheetViewport(sheet, far.ctx, geometry, view(0, 16_385 * 20 - 40))
    // Painted inside its own row band (the parsed default style is 11pt, not 10pt).
    expect(far.find('7')).toHaveLength(1)
    expect(far.find('7')[0].y).toBeGreaterThan(40)
    expect(far.find('7')[0].y).toBeLessThan(60)
    const m = computeMetrics(sheet)
    expect(m.exportLimits).toEqual(expect.arrayContaining([expect.objectContaining({ resource: 'rows', requested: 16_386 })]))
    expect(m.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'inconsistent-row-coordinate' })]))
  })
  test('F2 (corrected): only rasters the old path would have truncated are refused', async () => {
    // The legacy drawing clamp (16,384 px per axis, 16,777,216 px area) silently
    // cropped the bitmap; that request is now a typed refusal instead.
    const drawing = { ...sheetOf([{ index: 0, cells: [cell(0, 0, 'A1')] }]), drawings: [{ xEmu: 20_000 * 9525, yEmu: 0, widthEmu: 9525 * 10, heightEmu: 9525 * 10, geometry: 'rect' as const, fill: 'FF0000', source: { partPath: 'd', treePath: 't', element: 'sp', representation: 'native' as const } }] }
    const clamped = computeMetrics(drawing)
    expect(clamped.retainedDrawingBounds!.right).toBeLessThan(clamped.requestedDrawingBounds!.right)
    expect(clamped.exportLimits).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'large-export-limit', resource: 'pixels', limit: STATIC_SHEET_BUDGET.maxDrawingSidePx })]))
    expect(() => renderSheet(drawing, recorder().ctx)).toThrow(StaticExportLimitError)
    const units = await getPaintables({ sheets: [drawing] })
    try { expect(units[0].refusal).toEqual(expect.arrayContaining([expect.objectContaining({ resource: 'pixels' })])) } finally { units.dispose() }
    // A drawing-free sheet inside the legacy grid caps was always rendered in full
    // (spec §11.5: existing admitted static output is preserved), however large.
    const wide = sheetOf([{ index: 999, cells: [cell(999, 999, 'Far')] }])
    expect(computeMetrics(wide).exportLimits).toBeUndefined()
  })
  test('an ordinary 2,000-row sheet still paints completely through the static path', async () => {
    const rows: XlsxRow[] = Array.from({ length: 2000 }, (_, index) => ({ index, cells: [cell(index, 0, `Row ${index + 1}`), cell(index, 1, index)] }))
    const sheet = sheetOf(rows)
    const m = computeMetrics(sheet)
    expect(m.exportLimits).toBeUndefined()
    expect([m.widthPx, m.heightPx]).toEqual([128, 40_000])
    const units = await getPaintables({ sheets: [sheet] })
    try {
      expect(units[0].refusal).toBeUndefined()
      expect(units[0].spec).toEqual({ widthPx: 128, heightPx: 40_000 })
      // node-canvas caps a side at 32,767 px; fillText is recorded whether or not it lands on the surface.
      const out = recorder(128, 200)
      units[0].paint(out.ctx)
      expect(out.find('Row 1')).toEqual([{ text: 'Row 1', x: PAD, y: expect.closeTo(BASELINE, 6) }])
      expect(out.find('Row 2000')).toEqual([{ text: 'Row 2000', x: PAD, y: expect.closeTo(1999 * 20 + BASELINE, 6) }])
      expect(out.texts.filter(item => item.text.startsWith('Row ')).length).toBe(2000)
    } finally { units.dispose() }
  })
  test('F3: a canonical visibility mask can reveal an authored-hidden row', () => {
    const sheet = sheetOf([
      { index: 0, cells: [cell(0, 0, 'A1')] },
      { index: 1, hidden: true, cells: [cell(1, 0, 'Revealed')] },
    ])
    expect(sheetBaseHidden(sheet).intervals).toEqual([{ r0: 1, r1: 1 }])
    const revealed = ready(sheet, { visibility: visibilityFromHidden(createIntervalSet([]), MAX_ROWS) })
    expect(revealed.metrics.rows.size(1)).toBe(20)
    const out = recorder()
    renderSheetViewport(sheet, out.ctx, revealed, view(0, 0))
    expect(out.find('Revealed')).toEqual([{ text: 'Revealed', x: PAD, y: expect.closeTo(20 + BASELINE, 6) }])
    const snapshot = ready(sheet, { visibility: visibilityFromHidden(sheetBaseHidden(sheet), MAX_ROWS) })
    expect(snapshot.metrics.rows.size(1)).toBe(0)
  })
  test('F4: detail needs four device pixels per visible row/column; tiny positive sizes survive', () => {
    const onePx = sheetOf(Array.from({ length: 10 }, (_, index) => ({ index, heightPt: 0.75, customHeight: true, cells: [cell(index, 0, `r${index}`)] })))
    const geometry = ready(onePx)
    expect(geometry.metrics.rows.size(0)).toBe(1)
    const at = (zoom: number, deviceScale: number) => planViewport(geometry, { box: { x: 0, y: 0, width: 320, height: 200 }, zoom, deviceScale })
    expect(at(0.25, 1)).toMatchObject({ kind: 'coarse', diagnostics: [expect.objectContaining({ code: 'detail-limit' })] })
    expect(at(4, 1).kind).toBe('detail')
    const tiny = ready(sheetOf([{ index: 0, cells: [cell(0, 0, 'A1')] }, { index: 1, heightPt: 0.3, customHeight: true, cells: [cell(1, 0, 'Tiny')] }]))
    expect(tiny.metrics.rows.size(1)).toBeCloseTo(0.4, 12)
    expect(planViewport(tiny, { box: { x: 0, y: 0, width: 320, height: 200 }, zoom: 4, deviceScale: 3 }).kind).toBe('detail')
    expect(planViewport(tiny, { box: { x: 0, y: 0, width: 320, height: 200 }, zoom: 1, deviceScale: 1 }).kind).toBe('coarse')
    expect(planViewport(tiny, { box: { x: 0, y: 0, width: 320, height: 200 }, zoom: Number.NaN, deviceScale: 1 })).toMatchObject({ kind: 'diagnostic', diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'axis-invalid-query' })]) })
  })
  test('F5: an over-budget frame is refused before any cell content is read', () => {
    // value reads = content access; col reads = per-cell visits (index probes are logarithmic).
    let reads = 0, visits = 0
    const counted = (row: number, col: number): XlsxCell => {
      const base = cell(row, col, 'x')
      Object.defineProperty(base, 'value', { get() { reads++; return 'x' }, enumerable: true })
      return Object.defineProperty(base, 'col', { get() { visits++; return col }, enumerable: true }) as XlsxCell
    }
    const rows: XlsxRow[] = Array.from({ length: 100 }, (_, r) => ({ index: r, cells: Array.from({ length: 400 }, (_, c) => counted(r, c)) }))
    const sheet = sheetOf(rows)
    const geometry = ready(sheet)
    reads = 0
    visits = 0
    const plan = planViewport(geometry, view(0, 0, 400 * 64, 100 * 20))
    expect(plan).toMatchObject({ kind: 'coarse', diagnostics: [expect.objectContaining({ code: 'detail-limit' })] })
    expect(reads).toBe(0)
    // 100 rows x ~2 binary searches over 400 cells, never the 20,000 cells before refusal.
    expect(visits).toBeLessThan(100 * 40)
  })
  test('F5: a small pane paints only the drawings that intersect it', () => {
    let fills = 0
    const node = (xPx: number, yPx: number) => Object.defineProperty({ xEmu: xPx * 9525, yEmu: yPx * 9525, widthEmu: 20 * 9525, heightEmu: 10 * 9525, geometry: 'rect', source: { partPath: 'd', treePath: 't', element: 'sp', representation: 'native' } }, 'fill', { get() { fills++; return 'FF0000' }, enumerable: true })
    const drawings = [node(10, 5), ...Array.from({ length: 5000 }, (_, i) => node(100_000 + i * 30, 500_000))]
    const sheet = { ...sheetOf([{ index: 0, cells: [] }]), drawings } as XlsxSheet
    const geometry = ready(sheet)
    fills = 0
    const out = recorder(64, 20)
    renderSheetViewport(sheet, out.ctx, geometry, view(0, 0, 64, 20))
    expect(fills).toBeGreaterThan(0)
    expect(fills).toBeLessThan(10)
  })
  test('F6: boundary instances are coalesced and admitted at their real count', () => {
    const rows: XlsxRow[] = Array.from({ length: 4000 }, (_, index) => ({ index, hidden: index % 2 === 1, cells: [] }))
    const geometry = ready(sheetOf(rows))
    const plan = planViewport(geometry, view(0, 0, 320, 2000 * 20), 'M')
    if (plan.kind !== 'detail') throw new Error(`expected detail, got ${plan.kind}`)
    expect(new Set(plan.rowLines).size).toBe(plan.rowLines.length)
    expect(plan.rowLines.length + plan.colLines.length).toBeLessThanOrEqual(VIEWPORT_DETAIL_LIMITS.M.gridlines)
    const over = planViewport(geometry, view(0, 0, 320, 2050 * 20), 'M')
    expect(over.kind).toBe('coarse')
  })
  test('F7: merge extents are part of the static grid and refuse when uncovered', () => {
    const tall = { ...sheetOf([{ index: 0, cells: [cell(0, 0, 'Anchor')] }]), merges: ['A1:A20001'], mergeRanges: [{ minRow: 0, minCol: 0, maxRow: 20_000, maxCol: 0 }] }
    expect(computeMetrics(tall).exportLimits).toEqual(expect.arrayContaining([expect.objectContaining({ resource: 'rows', requested: 20_001 })]))
    const short = { ...sheetOf([{ index: 0, cells: [cell(0, 0, 'Anchor')] }]), merges: ['A1:B5'], mergeRanges: [{ minRow: 0, minCol: 0, maxRow: 4, maxCol: 1 }] }
    const m = computeMetrics(short)
    expect(m.exportLimits).toBeUndefined()
    expect([m.rowHeightsPx.length, m.colWidthsPx.length]).toEqual([5, 2])
    expect([m.widthPx, m.heightPx]).toEqual([128, 100])
  })
  test('F8: a bounded static viewport is admitted when it needs no uncovered coordinate', () => {
    const sheet = beyondCapSheet()
    const small = recorder(64, 20)
    expect(() => renderSheet(sheet, small.ctx, undefined, undefined, {}, { x: 0, y: 0, width: 64, height: 20 })).not.toThrow()
    expect(small.find('Origin')).toHaveLength(1)
    expect(small.find('BeyondCap')).toEqual([])
    const edge = recorder(64, 60)
    expect(() => renderSheet(sheet, edge.ctx, undefined, undefined, {}, { x: 0, y: 16_384 * 20 - 20, width: 64, height: 60 })).toThrow(StaticExportLimitError)
    expect(edge.find('BeyondCap')).toEqual([])
  })
  test('F9: text overflowing from a cell left of the pane is part of the crop', () => {
    const sheet = sheetOf([{ index: 0, cells: [cell(0, 0, 'A long label that overflows across several columns')] }, { index: 3, cells: [cell(3, 3, 'D4')] }])
    const m = computeMetrics(sheet)
    const full = createCanvas(m.widthPx + 400, 200)
    renderSheet(sheet, full.getContext('2d') as never, { ...m, widthPx: m.widthPx + 400, heightPx: 200 })
    const box = { x: 64, y: 0, width: 150, height: 60 }
    const pane = createCanvas(box.width, box.height)
    renderSheetViewport(sheet, pane.getContext('2d') as never, ready(sheet), { box, zoom: 1, deviceScale: 1 })
    const a = full.getContext('2d').getImageData(box.x, box.y, box.width, box.height).data
    const b = pane.getContext('2d').getImageData(0, 0, box.width, box.height).data
    let diff = 0
    for (let i = 0; i < a.length; i += 4) if (a[i] !== b[i] || a[i + 1] !== b[i + 1] || a[i + 2] !== b[i + 2]) diff++
    expect(diff).toBe(0)
  })
  test('F10: empty caller-supplied refusal metadata does not bypass coverage', () => {
    const sheet = beyondCapSheet()
    const m = computeMetrics(sheet)
    const out = recorder()
    expect(() => renderSheet(sheet, out.ctx, { ...m, exportLimits: [] })).toThrow(StaticExportLimitError)
    expect(out.find('BeyondCap')).toEqual([])
  })
})

describe('review round 2 (Codex findings)', () => {
  const source = { partPath: 'd', treePath: 't', element: 'sp', representation: 'native' as const }
  test('F4: overflow dependencies obey the four-device-pixel rule', () => {
    // Column A is 5px wide: 2.5 device px at 50%. B and the visible pane are wide.
    const sheet = sheetOf([{ index: 0, cells: [cell(0, 0, 'Overflow from tiny A'), cell(0, 1, 'B text'), cell(0, 3, 'D1')] }], [{ min: 0, max: 0, widthChars: 0 }])
    const geometry = ready(sheet)
    expect(geometry.metrics.cols.size(0)).toBe(5)
    const plan = planViewport(geometry, { box: { x: 5 + 64 * 2, y: 0, width: 200, height: 100 }, zoom: 0.5, deviceScale: 1 })
    if (plan.kind !== 'detail') throw new Error(`expected detail, got ${plan.kind}`)
    const cols = plan.cells.map(placed => placed.source.col)
    expect(cols).toContain(1)
    expect(cols).toContain(3)
    expect(cols).not.toContain(0)
  })
  test('F5: a group heavier than the record budget is coarse before any node is painted', () => {
    let fills = 0
    const child = (i: number) => Object.defineProperty({ xEmu: (i % 100) * 9525, yEmu: Math.floor(i / 100) * 9525, widthEmu: 9525, heightEmu: 9525, geometry: 'rect', source }, 'fill', { get() { fills++; return 'FF0000' }, enumerable: true })
    const group = { xEmu: 0, yEmu: 0, widthEmu: 100 * 9525, heightEmu: 300 * 9525, geometry: 'rect', source,
      group: { off: { x: 0, y: 0 }, ext: { width: 100 * 9525, height: 300 * 9525 }, chOff: { x: 0, y: 0 }, chExt: { width: 100 * 9525, height: 300 * 9525 } },
      children: Array.from({ length: 25_000 }, (_, i) => child(i)) }
    const sheet = { ...sheetOf([{ index: 0, cells: [] }]), drawings: [group] } as unknown as XlsxSheet
    const geometry = ready(sheet)
    fills = 0
    const out = recorder(64, 20)
    expect(renderSheetViewport(sheet, out.ctx, geometry, view(0, 0, 64, 20)).kind).toBe('coarse')
    expect(fills).toBe(0)
  })
  test('F5: drawing queries are visit-bounded; far drawings are pruned without coarse fallback', () => {
    const rect = (xPx: number) => ({ xEmu: xPx * 9525, yEmu: 0, widthEmu: 9525 * 10, heightEmu: 9525 * 10, geometry: 'rect', fill: 'FF0000', source })
    const far = { ...sheetOf([{ index: 0, cells: [] }]), drawings: Array.from({ length: 50_000 }, (_, i) => rect(100_000 + i * 30)) } as unknown as XlsxSheet
    expect(planViewport(ready(far), view(0, 0, 64, 20))).toMatchObject({ kind: 'detail', drawings: [] })
    // Same rows, alternating far left/right: no bounding box prunes them, so the
    // query exhausts its visit budget and the frame is coarse instead of scanning all.
    const straddle = { ...sheetOf([{ index: 0, cells: [] }]), drawings: Array.from({ length: 50_000 }, (_, i) => rect(i % 2 ? 100_000 + i * 30 : -100_000 - i * 30)) } as unknown as XlsxSheet
    expect(planViewport(ready(straddle), view(0, 0, 64, 20))).toMatchObject({ kind: 'coarse', diagnostics: [expect.objectContaining({ code: 'detail-limit' })] })
  })
  test('F7 (r3): a bounded request lays a grid-truncated merge out at its true extent', async () => {
    const tallMerge = (style?: XlsxCell['style']) => ({ ...sheetOf([{ index: 0, cells: [{ ...cell(0, 0, 'Anchor'), ...(style ? { style } : {}) }, cell(0, 1, 'B1')] }]), merges: ['A1:A20001'], mergeRanges: [{ minRow: 0, minCol: 0, maxRow: 20_000, maxCol: 0 }] })
    const sheet = tallMerge()
    const m = computeMetrics(sheet)
    const gridHeight = m.rowHeightsPx.length * 20
    // Upper region: allowed. Ordinary placement centres the text in the merged box,
    // whose true height is 20,001 rows; the truncated box would centre it at ~163,847.
    const trueCentre = (20_001 * 20 + 10 * 96 / 72) / 2
    const top = recorder(64, 200)
    expect(() => renderSheet(sheet, top.ctx, undefined, undefined, {}, { x: 0, y: 0, width: 64, height: 200 })).not.toThrow()
    expect(top.find('Anchor')).toEqual([{ text: 'Anchor', x: PAD, y: expect.closeTo(trueCentre, 6) }])
    const inWindow = (out: ReturnType<typeof recorder>, height: number) => out.find('Anchor').filter(item => item.y > 0 && item.y < height + 20)
    // Where the truncated layout would have put the text: nothing is painted there.
    const truncatedCentre = (gridHeight + 10 * 96 / 72) / 2
    const edge = recorder(64, 100)
    expect(() => renderSheet(sheet, edge.ctx, undefined, undefined, {}, { x: 0, y: truncatedCentre - 50, width: 64, height: 100 })).not.toThrow()
    expect(inWindow(edge, 100)).toEqual([])
    // Top-aligned anchor text paints at the merge top.
    const aligned = tallMerge({ numFmtId: 0, vertical: 'top' })
    const topAligned = recorder(64, 200)
    renderSheet(aligned, topAligned.ctx, undefined, undefined, {}, { x: 0, y: 0, width: 64, height: 200 })
    expect(topAligned.find('Anchor').length).toBeGreaterThan(0)
    expect(topAligned.find('Anchor')[0].y).toBeLessThan(20)
    // A window that needs rows past the grid is still refused.
    expect(() => renderSheet(sheet, recorder(64, 60).ctx, undefined, undefined, {}, { x: 0, y: gridHeight - 20, width: 64, height: 60 })).toThrow(StaticExportLimitError)
    const print = computePrintMetrics(sheet, m, 96)
    expect(() => renderPrintPage(sheet, recorder(64, 64).ctx, m, print, 0, 0)).not.toThrow()
  })

  test('r3 MINOR 2: per-frame merge lookup is not quadratic in records', () => {
    // 60 rows x 100 columns, every row covered by 50 two-column merges, all cells populated.
    let reads = 0
    const ranges: XlsxMergeRange[] = []
    for (let r = 0; r < 60; r++) for (let c = 0; c < 100; c += 2) {
      const range = { minRow: r, minCol: c, maxRow: r, maxCol: c + 1 }
      ranges.push(new Proxy(range, { get(target, key) { reads++; return Reflect.get(target, key) } }))
    }
    const rows: XlsxRow[] = Array.from({ length: 60 }, (_, r) => ({ index: r, cells: Array.from({ length: 100 }, (_, c) => cell(r, c, `${r}:${c}`)) }))
    const sheet = { ...sheetOf(rows), mergeRanges: ranges }
    const geometry = ready(sheet)
    reads = 0
    const plan = planViewport(geometry, view(0, 0, 100 * 64, 60 * 20))
    if (plan.kind !== 'detail') throw new Error(`expected detail, got ${plan.kind}`)
    expect(plan.cells).toHaveLength(3000)
    expect(plan.cells.every(placed => placed.merge !== undefined && placed.source.col % 2 === 0)).toBe(true)
    // A linear scan per cell reads ~6,000 x 3,000 range fields; an index is near-linear.
    expect(reads).toBeLessThan(500_000)
  })
  test('r3 NIT 4: exactly four device pixels admits detail', () => {
    const sheet = sheetOf([{ index: 0, cells: [cell(0, 0, 'A1')] }, { index: 1, heightPt: 0.3, customHeight: true, cells: [cell(1, 0, 'Tiny')] }])
    const geometry = ready(sheet)
    expect(geometry.metrics.rows.size(1) * 10).toBeLessThan(4)
    expect(planViewport(geometry, { box: { x: 0, y: 0, width: 64, height: 30 }, zoom: 10, deviceScale: 1 }).kind).toBe('detail')
    const fourPx = ready(sheetOf([{ index: 0, heightPt: 3, customHeight: true, cells: [cell(0, 0, 'A1')] }]))
    expect(fourPx.metrics.rows.size(0)).toBe(4)
    expect(planViewport(fourPx, { box: { x: 0, y: 0, width: 64, height: 4 }, zoom: 1, deviceScale: 1 }).kind).toBe('detail')
  })
})

describe('delta review (opencode)', () => {
  test('beyond-grid merge edges cost O(B + T log B), not T x B', () => {
    // 3,000 custom-height rows past the 16,384-row grid and 3,000 edge queries.
    const rows: XlsxRow[] = [{ index: 0, cells: [cell(0, 0, 'A1')] }]
    for (let i = 0; i < 3000; i++) rows.push({ index: 16_384 + i * 2, heightPt: 30, customHeight: true, cells: [] })
    const sheet = sheetOf(rows)
    const m = computeMetrics(sheet)
    expect(m.rowHeightsPx).toHaveLength(16_384)
    const edges = mergeEdges(sheet, m)
    let expected = 16_384 * 20
    let last = 16_384
    const results: number[] = []
    for (let t = 0; t < 3000; t++) results.push(edges.row(16_385 + t * 2))
    // Row 16_384 + 2k carries 40px; odd rows keep 20px.
    for (let t = 0; t < 3000; t++) {
      const index = 16_385 + t * 2
      while (last < index) { expected += (last - 16_384) % 2 === 0 ? 40 : 20; last++ }
      expect(results[t]).toBe(expected)
    }
    expect(edges.work()).toBeLessThan(3000 * 20)
  })
})
