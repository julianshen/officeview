import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cellTextRotation, renderSheet } from '../src/xlsx/render'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import type { XlsxCellStyle, XlsxSheet } from '../src/xlsx/types'
import { RECORD_TEXT, type LogicalTextRange } from '../src/core/text-recording'
import { buildTextIndex, findMatches } from '../src/core/search'
import { hitTest, textForRange } from '../src/core/selection'

function fixedSheetPaint(rotation: number | undefined, alignment: Partial<XlsxCellStyle> = {}, text: string | number = 'AB CD EF GH IJ', width = 128, height = 160, record = false, parsedSheet?: XlsxSheet) {
  const ctx = createCanvas(width, height).getContext('2d')
  Object.defineProperty(ctx.canvas, 'ownerDocument', { value: { createElement: () => ({ getContext: () => ({ font: '', measureText: (text: string) => ({ width: [...text].length * 10, actualBoundingBoxAscent: 12, actualBoundingBoxDescent: 4, fontBoundingBoxAscent: 12, fontBoundingBoxDescent: 4 }) }) }) } })
  ctx.measureText = text => ({ width: [...text].length * 10, actualBoundingBoxAscent: 12, actualBoundingBoxDescent: 4 }) as TextMetrics
  const calls: Array<{ text: string; x: number; y: number; font: string; matrix: number[]; logical?: LogicalTextRange; width: number }> = []
  const capture = (text: string, x: number, y: number, width = [...text].length * 10, logical?: LogicalTextRange) => {
    const m = ctx.getTransform(); calls.push({ text, x, y, width, font: ctx.font, matrix: [m.a, m.b, m.c, m.d, m.e, m.f], logical })
  }
  ctx.fillText = capture
  if (record) (ctx as unknown as Record<symbol, unknown>)[RECORD_TEXT] = capture
  const sheet: XlsxSheet = parsedSheet ?? { name: 'Fixed', cols: [{ min: 0, max: 0, widthChars: (width - 5) / 7 }], merges: [], mergeRanges: [], rows: [{ index: 0, heightPt: height * 72 / 96, customHeight: true, cells: [{ row: 0, col: 0, ref: 'A1', value: text, styleIndex: 0, style: { numFmtId: 0, fontSizePt: 11, textRotation: rotation, ...alignment } }] }] }
  renderSheet(sheet, ctx as never)
  return { calls, sheet }
}

test('Excel authored general/bottom/no-wrap defaults and manual rotation0 preserve ordinary placement', () => {
  for (const value of ['AB CD EF GH IJ', 123] as const) {
    const absent = fixedSheetPaint(undefined, {}, value, 64, 40, true)
    for (const style of [{ horizontal: 'general' }, { vertical: 'bottom' }, { wrapText: false }, { horizontal: 'general', vertical: 'bottom', wrapText: false }] as const) {
      expect(fixedSheetPaint(undefined, style, value, 64, 40, true).calls).toEqual(absent.calls)
      expect(fixedSheetPaint(0, style, value, 64, 40, true).calls).toEqual(absent.calls)
    }
    const genuinelyAuthored = [
      fixedSheetPaint(0, { vertical: 'top' }, value, 64, 40, true),
      fixedSheetPaint(0, { horizontal: 'center' }, value, 64, 40, true),
      fixedSheetPaint(0, { horizontal: 'general', vertical: 'bottom', wrapText: true }, value, 64, 80, true),
      fixedSheetPaint(90, { horizontal: 'general', vertical: 'bottom', wrapText: false }, value, 64, 40, true),
    ]
    for (const paint of genuinelyAuthored) expect(paint.calls.some(call => call.logical !== undefined)).toBe(true)
  }
})

const generalCases: Array<{ label: string; rotation: number; alignment: Partial<XlsxCellStyle>; xml: string }> = [
  { label: 'top', rotation: 0, alignment: { vertical: 'top' }, xml: ' vertical="top"' },
  { label: 'center', rotation: 0, alignment: { vertical: 'center' }, xml: ' vertical="center"' },
  { label: 'wrap', rotation: 0, alignment: { vertical: 'bottom', wrapText: true }, xml: ' vertical="bottom" wrapText="1"' },
  ...[45, 90, 255].map(rotation => ({ label: `rotation${rotation}`, rotation, alignment: { vertical: 'top', wrapText: true }, xml: ' vertical="top" wrapText="1"' })),
]
function paintedLeft(calls: ReturnType<typeof fixedSheetPaint>['calls']): number {
  const size = 11 * 96 / 72, ascent = Math.max(12, size * .85), descent = Math.max(4, size * .25)
  return Math.min(...calls.flatMap(call => {
    const [a, , c, , e] = call.matrix
    return [[call.x, call.y - ascent], [call.x + call.width, call.y - ascent], [call.x, call.y + descent], [call.x + call.width, call.y + descent]].map(([x, y]) => a * x + c * y + e)
  }))
}

test.each(generalCases)('numeric General+$label production fillText positions match absent/right and distinguish authored left/center/text', ({ rotation, alignment }) => {
  const general = fixedSheetPaint(rotation, { ...alignment, horizontal: 'general' }, 123, 64, 80)
  const absent = fixedSheetPaint(rotation, alignment, 123, 64, 80)
  const right = fixedSheetPaint(rotation, { ...alignment, horizontal: 'right' }, 123, 64, 80)
  const left = fixedSheetPaint(rotation, { ...alignment, horizontal: 'left' }, 123, 64, 80)
  const center = fixedSheetPaint(rotation, { ...alignment, horizontal: 'center' }, 123, 64, 80)
  expect(general.calls).toEqual(absent.calls); expect(general.calls).toEqual(right.calls)
  expect(paintedLeft(general.calls)).toBeGreaterThan(paintedLeft(left.calls))
  expect(paintedLeft(center.calls)).toBeGreaterThan(paintedLeft(left.calls)); expect(paintedLeft(center.calls)).toBeLessThan(paintedLeft(right.calls))
  const text = fixedSheetPaint(rotation, { ...alignment, horizontal: 'general' }, '123', 64, 80)
  expect(text.calls).toEqual(fixedSheetPaint(rotation, alignment, '123', 64, 80).calls)
  expect(text.calls).toEqual(fixedSheetPaint(rotation, { ...alignment, horizontal: 'left' }, '123', 64, 80).calls)
  expect(text.calls).toEqual(left.calls)
  expect(general.calls.map(call => call.text).join('')).toBe('123')
  if (rotation === 0) { expect(paintedLeft(general.calls)).toBeCloseTo(31); expect(paintedLeft(left.calls)).toBeCloseTo(3) }
})

test.each([0, 45, 90, 135, 180, 255])('Excel %s honors authored alignment and wraps fitting text in bounded local lines', rotation => {
  const left = fixedSheetPaint(rotation, { horizontal: 'left', vertical: 'top', wrapText: true })
  const right = fixedSheetPaint(rotation, { horizontal: 'right', vertical: 'bottom', wrapText: true })
  expect(left.calls).not.toEqual(right.calls)
  const wrapped = fixedSheetPaint(rotation, { horizontal: 'left', vertical: 'top', wrapText: true }, 'AB CD EF GH IJ KL MN OP QR ST')
  const noWrap = fixedSheetPaint(rotation, { horizontal: 'left', vertical: 'top', wrapText: false }, 'AB CD EF GH IJ KL MN OP QR ST')
  expect(wrapped.calls).not.toEqual(noWrap.calls)
  expect(wrapped.calls.map(call => call.text).join('')).toBe('AB CD EF GH IJ KL MN OP QR ST')
  for (const call of wrapped.calls) {
    const [a, b, c, d, e, f] = call.matrix
    const size = 11 * 96 / 72, ascent = Math.max(12, size * .85), descent = Math.max(4, size * .25)
    for (const [x, y] of [[call.x, call.y - ascent], [call.x + call.width, call.y - ascent], [call.x, call.y + descent], [call.x + call.width, call.y + descent]]) {
      expect(a * x + c * y + e).toBeGreaterThanOrEqual(-1e-8)
      expect(a * x + c * y + e).toBeLessThanOrEqual(128 + 1e-8)
      expect(b * x + d * y + f).toBeGreaterThanOrEqual(-1e-8)
      expect(b * x + d * y + f).toBeLessThanOrEqual(160 + 1e-8)
    }
  }
  const numeric = fixedSheetPaint(rotation, { horizontal: 'left', vertical: 'top', wrapText: true }, 12)
  expect(numeric.calls[0].matrix).toEqual(fixedSheetPaint(rotation, { horizontal: 'left', vertical: 'top', wrapText: true }, '12').calls[0].matrix)
  expect(numeric.calls[0].x).toEqual(fixedSheetPaint(rotation, { horizontal: 'left', vertical: 'top', wrapText: true }, '12').calls[0].x)
})

test('Excel255 uses seven em cells including authored space, valid CSS font, source offsets and fitting glyph bounds', () => {
  const { calls } = fixedSheetPaint(255, { horizontal: 'left', vertical: 'top', wrapText: true }, 'AB12 中文（甲）—乙', 64, 107, true)
  const firstX = calls[0].matrix[4]
  const right = calls.filter(call => Math.abs(call.matrix[4] - firstX) < 1e-8)
  const left = calls.filter(call => Math.abs(call.matrix[4] - firstX) > 1e-8)
  expect(right.map(call => call.text).join('')).toBe('AB12 中文')
  expect(left.map(call => call.text).join('')).toBe('（甲）—乙')
  expect(right[0].matrix[4]).toBeGreaterThan(left[0].matrix[4])
  expect(calls.map(call => [call.logical?.start, call.logical?.end])).toEqual(Array.from({ length: 12 }, (_, i) => [i, i + 1]))
  expect(right[4].logical).toMatchObject({ start: 4, end: 5 })
  expect(right[5].matrix[5] - right[4].matrix[5]).toBeCloseTo(11 * 96 / 72)
  for (const call of calls) {
    expect(call.font).toBe('11pt "Calibri"')
    expect(call.matrix[4] + call.x).toBeGreaterThanOrEqual(0)
    expect(call.matrix[4] + call.x + 10).toBeLessThanOrEqual(64)
    expect(call.matrix[5] - 11 * 96 / 72 * .85).toBeGreaterThanOrEqual(-1e-8)
    expect(call.matrix[5] + 4).toBeLessThanOrEqual(107 + 1e-8)
  }
})

test.each([45, 90, 135, 180, 255])('Excel %s records complete wrapped/merged source with in-cell bands and rejects following-row hits', async rotation => {
  const source = 'AB12 中文（甲）—乙 CD34 文句。第三行'
  const sheet: XlsxSheet = { name: 'Bounds', cols: [{ min: 0, max: 1, widthChars: (128 - 5) / 7 }], merges: ['A1:B2'], mergeRanges: [{ minRow: 0, maxRow: 1, minCol: 0, maxCol: 1 }], rows: [
    { index: 0, heightPt: 40, customHeight: true, cells: [{ row: 0, col: 0, ref: 'A1', value: source, styleIndex: 0, style: { numFmtId: 0, fontSizePt: 11, textRotation: rotation, horizontal: 'left', vertical: 'top', wrapText: true } }] },
    { index: 1, heightPt: 40, customHeight: true, cells: [] }, { index: 2, heightPt: 40, customHeight: true, cells: [] }
  ] }
  const index = await buildTextIndex([{ spec: { widthPx: 256, heightPx: 160 }, paint: ctx => renderSheet(sheet, ctx) }])
  const lines = index.pages[0].lines, last = lines.length - 1
  expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: last, charIndex: lines[last].text.length } })).toBe(source)
  const match = findMatches(index, source)[0]
  expect(match.logical?.source.text).toBe(source)
  for (const rect of match.rects) {
    expect(rect.x).toBeGreaterThanOrEqual(-1e-8); expect(rect.y).toBeGreaterThanOrEqual(-1e-8)
    expect(rect.x + rect.width).toBeLessThanOrEqual(256 + 1e-8)
    expect(rect.y + rect.height).toBeLessThanOrEqual(106 + 1e-8)
  }
  for (const strict of [true, false]) expect(hitTest(index, 0, 30, 125, { strict })).toBeUndefined()
  const span = lines[0].spans.find(span => span.text.length > 0)!, p = span.placement!, t = p.transform
  const localX = p.x + p.width / span.text.length * .75, localY = (p.top + p.bottom) / 2
  const x = t.a * localX + t.c * localY + t.e, y = t.b * localX + t.d * localY + t.f
  expect(hitTest(index, 0, x, y, { strict: true })?.charIndex).toBe(1)
})

test('Excel no-wrap overflow keeps authored source/font and clips all physical and recorded hits to its cell', async () => {
  const { sheet } = fixedSheetPaint(90, { horizontal: 'left', vertical: 'top', wrapText: false }, 'ABCDEFGHIJKLMNO', 64, 40)
  const index = await buildTextIndex([{ spec: { widthPx: 64, heightPx: 80 }, paint: ctx => renderSheet(sheet, ctx) }])
  const line = index.pages[0].lines[0]
  expect(line.text).toBe('ABCDEFGHIJKLMNO')
  expect(line.spans[0].fontSize).toBeCloseTo(11 * 96 / 72)
  const range = { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: 0, charIndex: line.text.length } }
  expect(textForRange(index, range)).toBe('ABCDEFGHIJKLMNO')
  expect(findMatches(index, 'ABCDEFGHIJKLMNO')[0].rects.every(rect => rect.y >= 0 && rect.y + rect.height <= 40 + 1e-8)).toBe(true)
  for (const strict of [true, false]) expect(hitTest(index, 0, 15, 55, { strict })).toBeUndefined()
})

const stylesXml = (fonts: string, xfs: string[], xfCount: number) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">${fonts}</fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="${xfCount}">
    ${xfs.join('\n    ')}
  </cellXfs>
</styleSheet>`
const font11 = '<font><sz val="11"/><name val="Calibri"/></font>'
const font6 = '<font><sz val="6"/><name val="Calibri"/></font>'
const xf = (rotation: number | undefined, fontId = 0, alignment = '') =>
  `<xf numFmtId="0" fontId="${fontId}" fillId="0" borderId="0" xfId="0">${rotation === undefined && !alignment ? '' : `<alignment${rotation === undefined ? '' : ` textRotation="${rotation}"`}${alignment} />`}</xf>`

async function fixture(styles: string, cells: Array<{ ref: string; v: string; style: number }>, rows = 1) {
  const base = await buildXlsx([{
    name: 'Rot',
    rows: Array.from({ length: rows }, (_, i) => ({
      r: i + 1, cells: cells.filter(c => c.ref === `A${i + 1}`).map(c => ({ ref: c.ref, v: c.v, style: c.style })),
    })),
  }])
  const zip = await JSZip.loadAsync(base)
  zip.file('xl/styles.xml', styles)
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
}
test.each(generalCases)('parsed authored numeric General+$label retains fallback through actual production paint', async ({ rotation, alignment, xml }) => {
  const doc = await parseXlsx(await fixture(stylesXml(font11, [xf(rotation, 0, ` horizontal="general"${xml}`)], 1), [{ ref: 'A1', v: '123', style: 0 }]))
  const sheet = doc.sheets[0], cell = sheet.rows[0].cells[0]
  expect(cell.value).toBe(123); expect(cell.style?.horizontal).toBe('general')
  sheet.rows[0].heightPt = 60; sheet.rows[0].customHeight = true
  const painted = fixedSheetPaint(rotation, {}, 123, 64, 80, false, sheet)
  expect(painted.calls).toEqual(fixedSheetPaint(rotation, alignment, 123, 64, 80).calls)
  expect(painted.calls).toEqual(fixedSheetPaint(rotation, { ...alignment, horizontal: 'right' }, 123, 64, 80).calls)
  expect(paintedLeft(painted.calls)).toBeGreaterThan(paintedLeft(fixedSheetPaint(rotation, { ...alignment, horizontal: 'left' }, 123, 64, 80).calls))
})

test.each(['0', 'false', '1', 'true', undefined])('parsed Excel255 wrapText=%s keeps explicit precedence through clipped source recording', async authored => {
  const doc = await parseXlsx(await fixture(stylesXml(font11, [xf(255, 0, authored === undefined ? '' : ` wrapText="${authored}"`)], 1), [{ ref: 'A1', v: 'ABCDEFGH', style: 0 }]))
  const sheet = doc.sheets[0], style = sheet.rows[0].cells[0].style!
  const explicitFalse = authored === '0' || authored === 'false'
  expect(style.wrapText).toBe(authored === undefined ? undefined : !explicitFalse)
  sheet.rows[0].heightPt = 33; sheet.rows[0].customHeight = true
  const index = await buildTextIndex([{ spec: { widthPx: 64, heightPx: 90 }, paint: ctx => renderSheet(sheet, ctx) }])
  const lines = index.pages[0].lines, last = lines.length - 1
  expect(lines.length).toBe(explicitFalse ? 1 : 3)
  expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: last, charIndex: lines[last].text.length } })).toBe('ABCDEFGH')
  expect(lines.flatMap(line => line.spans).every(span => Math.abs(span.fontSize - 11 * 96 / 72) < 1e-8)).toBe(true)
  expect(findMatches(index, 'ABCDEFGH')[0].rects.every(rect => rect.x >= -1e-8 && rect.y >= -1e-8 && rect.x + rect.width <= 64 + 1e-8 && rect.y + rect.height <= 44 + 1e-8)).toBe(true)
  if (explicitFalse) expect(findMatches(index, 'GH')[0].rects).toEqual([])
  for (const strict of [true, false]) expect(hitTest(index, 0, 20, 60, { strict })).toBeUndefined()
})

const rotationsStyles = (rotations: Array<number | undefined>) =>
  stylesXml(`${font11}${font6}`, rotations.map(r => xf(r)), rotations.length)

function inkBox(ctx: ReturnType<ReturnType<typeof createCanvas>['getContext']>, w: number, h: number) {
  const data = ctx.getImageData(0, 0, w, h).data
  let minX = w, maxX = -1, minY = h, maxY = -1, count = 0
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    if (data[(y * w + x) * 4] < 128) { count++; minX = Math.min(minX, x); maxX = Math.max(maxX, x); minY = Math.min(minY, y); maxY = Math.max(maxY, y) }
  }
  return { count, minX, maxX, minY, maxY, w: maxX - minX, h: maxY - minY }
}

describe('xlsx textRotation', () => {
  test('parses rotation values and normalizes horizontal/invalid to undefined', async () => {
    const rotations = [0, 45, 90, 135, 180, 255, 200, undefined]
    const doc = await parseXlsx(await fixture(rotationsStyles(rotations), rotations.map((_, i) => ({ ref: `A${i + 1}`, v: `v${i}`, style: i })), rotations.length))
    const got = doc.sheets[0].rows.map(row => row.cells[0].style?.textRotation)
    expect(got).toEqual([undefined, 45, 90, 135, 180, 255, undefined, undefined])
  })

  test('rotation angle mapping follows the OOXML encoding, not a naive rotation', () => {
    expect(cellTextRotation(undefined)).toBeUndefined()
    expect(cellTextRotation(0)).toBeUndefined()
    expect(cellTextRotation(45)).toEqual({ radians: -Math.PI / 4 })
    expect(cellTextRotation(90)).toEqual({ radians: -Math.PI / 2 })
    expect(cellTextRotation(135)).toEqual({ radians: Math.PI / 4 })
    // 180 means 90 degrees clockwise, not a half turn.
    expect(cellTextRotation(180)).toEqual({ radians: Math.PI / 2 })
    expect(cellTextRotation(255)).toEqual({ stacked: true })
    expect(cellTextRotation(200)).toBeUndefined()
  })

  test('parses alignment horizontal/vertical/wrap alongside rotation', async () => {
    const styles =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment horizontal="center" vertical="bottom" wrapText="1" textRotation="90"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
</styleSheet>`
    const base = await buildXlsx([{ name: 'Rot', rows: [{ r: 1, cells: [{ ref: 'A1', v: 'x', style: 0 }] }] }])
    const zip = await JSZip.loadAsync(base)
    zip.file('xl/styles.xml', styles)
    const doc = await parseXlsx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const [aligned, plain] = doc.sheets[0].rows[0].cells.map(c => c.style)
    expect(aligned).toMatchObject({ textRotation: 90, horizontal: 'center', vertical: 'bottom', wrapText: true })
    expect(plain?.horizontal).toBeUndefined()
  })

  test('90-degree text paints a tall narrow band where horizontal paints short wide', async () => {
    const box = async (rotation: number | undefined) => {
      const doc = await parseXlsx(await fixture(rotationsStyles([rotation]), [{ ref: 'A1', v: 'AB12 AB12 AB12', style: 0 }]))
      const sheet = doc.sheets[0]
      const canvas = createCanvas(400, 200), ctx = canvas.getContext('2d')
      ctx.fillStyle = '#ffffff'
      ctx.fillRect(0, 0, 400, 200)
      renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D)
      return inkBox(ctx, 400, 200)
    }
    const flat = await box(undefined)
    const up = await box(90)
    expect(flat.count).toBeGreaterThan(0)
    expect(up.count).toBeGreaterThan(0)
    expect(up.h / Math.max(1, up.w)).toBeGreaterThan(flat.h / Math.max(1, flat.w))
  })

  test('stacked 255 block anchors near the cell left with top-to-bottom order', async () => {
    // 6pt keeps AB in one upright column inside the default 20px row.
    const styles = stylesXml(`${font11}${font6}`, [xf(undefined), xf(255, 1)], 2)
    const doc = await parseXlsx(await fixture(styles, [{ ref: 'A1', v: 'AB', style: 1 }]))
    const sheet = doc.sheets[0]
    const canvas = createCanvas(400, 120), ctx = canvas.getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, 400, 120)
    renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D)
    const data = ctx.getImageData(0, 0, 400, 120).data
    const dark = (x: number, y: number) => data[(y * 400 + x) * 4] < 128
    let topA = -1, count = 0, inkLeft = 400, inkRight = -1
    for (let y = 0; y < 120; y++) for (let x = 0; x < 400; x++) {
      if (!dark(x, y)) continue
      count++
      inkLeft = Math.min(inkLeft, x); inkRight = Math.max(inkRight, x)
      if (topA < 0) topA = y
    }
    // Find the B row: second ink band below the first gap.
    let gap = false, topB = -1
    for (let y = topA; y < 120; y++) {
      let row = false
      for (let x = inkLeft; x <= inkRight; x++) if (dark(x, y)) { row = true; break }
      if (!row) gap = true
      else if (gap) { topB = y; break }
    }
    expect(count).toBeGreaterThan(0)
    expect(topA).toBeGreaterThanOrEqual(0)
    expect(topB).toBeGreaterThan(topA)
    // Block anchored near the cell left (default 64px column): ink starts
    // left of center. A right-anchored block would start past center.
    expect(inkLeft).toBeLessThan(32)
  })

  test('rotated text clips to its merged anchor box', async () => {
    const long = 'AB12 AB12 AB12 AB12 AB12 AB12'
    const base = await buildXlsx([{ name: 'Rot', rows: [{ r: 1, cells: [{ ref: 'A1', v: long, style: 1 }] }, { r: 2, cells: [] }], merges: ['A1:B2'] }])
    const zip = await JSZip.loadAsync(base)
    zip.file('xl/styles.xml', rotationsStyles([undefined, 90]))
    const doc = await parseXlsx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const canvas = createCanvas(400, 200), ctx = canvas.getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, 400, 200)
    renderSheet(doc.sheets[0], ctx as unknown as CanvasRenderingContext2D)
    const data = ctx.getImageData(0, 0, 400, 200).data
    let count = 0, outside = 0
    for (let y = 0; y < 200; y++) for (let x = 0; x < 400; x++) {
      if (data[(y * 400 + x) * 4] < 128) {
        count++
        // Merged A1:B2 with default metrics spans x[0,128) y[0,40).
        if (x >= 128 || y >= 40) outside++
      }
    }
    expect(count).toBeGreaterThan(0)
    expect(outside).toBe(0)
  })

  test('all five rotations paint visible ink inside fitting bounds with top/left alignment', async () => {
    // Merged A1:B2 (128x40 default metrics) fits every rotation of AB12@11pt.
    for (const rotation of [45, 90, 135, 180, 255]) {
      const styles = stylesXml(`${font11}${font6}`, [xf(rotation, 0, ' horizontal="left" vertical="top" wrapText="1"')], 1)
      const base = await buildXlsx([{ name: 'Rot', rows: [{ r: 1, cells: [{ ref: 'A1', v: 'AB12', style: 0 }] }, { r: 2, cells: [] }], merges: ['A1:B2'] }])
      const zip = await JSZip.loadAsync(base)
      zip.file('xl/styles.xml', styles)
      const doc = await parseXlsx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
      const canvas = createCanvas(400, 200), ctx = canvas.getContext('2d')
      ctx.fillStyle = '#ffffff'
      ctx.fillRect(0, 0, 400, 200)
      renderSheet(doc.sheets[0], ctx as unknown as CanvasRenderingContext2D)
      const data = ctx.getImageData(0, 0, 400, 200).data
      let count = 0
      for (let y = 0; y < 40; y++) for (let x = 0; x < 128; x++) {
        if (data[(y * 400 + x) * 4] < 128) count++
      }
      expect(count, `rotation ${rotation} visible`).toBeGreaterThan(0)
    }
  })

  test('stacked 255 wraps into RTL columns by available height with all glyphs kept', async () => {
    const styles = stylesXml(`${font11}${font6}`, [xf(255, 0, ' horizontal="left" vertical="top" wrapText="1"')], 1)
    const base = await buildXlsx([{ name: 'Rot', rows: [{ r: 1, cells: [{ ref: 'A1', v: 'AB12中文AB12中文', style: 0 }] }] }])
    const zip = await JSZip.loadAsync(base)
    zip.file('xl/styles.xml', styles)
    let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    sheet = sheet.replace('<row r="1">', '<row r="1" ht="80" customHeight="1">')
    zip.file('xl/worksheets/sheet1.xml', sheet)
    const doc = await parseXlsx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const canvas = createCanvas(400, 300), ctx = canvas.getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, 400, 300)
    renderSheet(doc.sheets[0], ctx as unknown as CanvasRenderingContext2D)
    const data = ctx.getImageData(0, 0, 400, 300).data
    const dark = (x: number, y: number) => data[(y * 400 + x) * 4] < 128
    // Column bands: vertical runs of ink separated by blank columns.
    const bands: Array<[number, number]> = []
    let open: number | undefined
    for (let x = 0; x < 400; x++) {
      let col = false
      for (let y = 0; y < 300; y++) if (dark(x, y)) { col = true; break }
      if (col && open === undefined) open = x
      if (!col && open !== undefined) {
        if (x - open > 2) bands.push([open, x])
        open = undefined
      }
    }
    if (open !== undefined && 400 - open > 2) bands.push([open, 400])
    expect(bands.length).toBeGreaterThanOrEqual(2)
    // Bands scan left-to-right, so index 1 is the first RTL column: it sits
    // right of the second column.
    const topOf = (band: [number, number]) => {
      for (let y = 0; y < 300; y++) for (let x = band[0]; x < band[1]; x++) if (dark(x, y)) return y
      return -1
    }
    expect(bands[1][0]).toBeGreaterThan(bands[0][0])
    expect(topOf(bands[0])).toBeGreaterThanOrEqual(0)
    expect(topOf(bands[1])).toBeGreaterThanOrEqual(0)
    // Block starts inside the cell (default 64px column), not off-canvas.
    expect(bands[1][0]).toBeLessThan(64)
  })

  test('parsed alignment moves rotated text between cell edges', async () => {
    const paint = async (horizontal: string, vertical: string) => {
      const styles =
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="20"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment horizontal="${horizontal}" vertical="${vertical}" textRotation="45"/></xf>
  </cellXfs>
</styleSheet>`
      // Merged A1:B4 (128x80 default metrics) fits the whole 45-degree box
      // under either alignment, so ordering assertions are geometric.
      const rows = [1, 2, 3, 4].map(r => ({ r, cells: r === 1 ? [{ ref: 'A1', v: 'AB', style: 0 }] : [] }))
      const base = await buildXlsx([{ name: 'Rot', rows, merges: ['A1:B4'] }])
      const zip = await JSZip.loadAsync(base)
      zip.file('xl/styles.xml', styles)
      const doc = await parseXlsx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
      const canvas = createCanvas(500, 200), ctx = canvas.getContext('2d')
      ctx.fillStyle = '#ffffff'
      ctx.fillRect(0, 0, 500, 200)
      renderSheet(doc.sheets[0], ctx as unknown as CanvasRenderingContext2D)
      return inkBox(ctx, 500, 200)
    }
    const leftTop = await paint('left', 'top')
    const rightBottom = await paint('right', 'bottom')
    expect(leftTop.count).toBeGreaterThan(0)
    expect(rightBottom.count).toBeGreaterThan(0)
    // Ordering only here; exact anchors are pinned by the model test above.
    expect(rightBottom.minX).toBeGreaterThan(leftTop.minX)
    expect(rightBottom.minY).toBeGreaterThan(leftTop.minY)
  })
})
