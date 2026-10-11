import { describe, expect, test } from 'vitest'
import { createCanvas } from 'canvas'
import { computeMetrics, renderSheet, renderSheetViewport, StaticExportLimitError } from '../src/xlsx/render'
import { getPaintables } from '../src/render/paint'
import { AXIS_ERROR_DIAGNOSTICS, buildSheetGeometry, planViewport, VIEWPORT_DETAIL_LIMITS, type SheetGeometry } from '../src/xlsx/view/viewport'
import { createRowOrder } from '../src/xlsx/view/order'
import { createRowVisibility } from '../src/xlsx/view/visibility'
import type { AxisMetricsErrorCode } from '../src/xlsx/view/axes'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import type { LogicalViewport } from '../src/xlsx/view/contracts'
import type { XlsxCell, XlsxColumnSpec, XlsxRow, XlsxSheet } from '../src/xlsx/types'

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
      renderSheet(beyondCapSheet(), ctx, undefined, undefined, {}, { x: 0, y: 0, width: 320, height: 200 })
    } catch (caught) { error = caught }
    expect(find('BeyondCap')).toEqual([])
    expect(error).toBeInstanceOf(StaticExportLimitError)
    expect(error).toMatchObject({ code: 'large-export-limit' })
  })
  test('computeMetrics reports the retained grid instead of silently clamping', () => {
    const m = computeMetrics(beyondCapSheet())
    expect(m.rowHeightsPx).toHaveLength(16_384)
    expect(m.exportLimits).toEqual([expect.objectContaining({ code: 'large-export-limit', resource: 'rows', requested: 16_386, limit: 16_384 })])
    const edge = computeMetrics(sheetOf([{ index: 0, cells: [cell(0, MAX_COLS - 1, 'EdgeCol')] }]))
    expect(edge.exportLimits).toEqual([expect.objectContaining({ resource: 'cols', requested: MAX_COLS, limit: 4096 })])
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
      expect(units[0].refusal).toEqual([expect.objectContaining({ code: 'large-export-limit', resource: 'rows' })])
      expect(units[1].refusal).toBeUndefined()
      const { ctx, find, texts } = recorder(units[0].spec.widthPx, units[0].spec.heightPx)
      units[0].paint(ctx)
      expect(find('BeyondCap')).toEqual([])
      expect(find('Origin')).toEqual([])
      expect(texts.some(item => /too large/i.test(item.text))).toBe(true)
      expect(sheet.diagnostics).toEqual([expect.objectContaining({ kind: 'large-export-limit', feature: 'large-export-limit' })])
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
