// @vitest-environment node
/**
 * PR #55 round 4 P2: an indented number is fitted against the indented content box
 * layoutCellText aligns it in. General fitting, the '#' overflow and shrink-to-fit
 * must decide exactly as for the same cell, unindented, that much narrower (the
 * existing non-indented rules; Excel's indented overflow is unmeasured), and the
 * painted text must stay inside the cell instead of losing its leading digits.
 */
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import { layoutCellBody, paintCellLayout, textMeasurer } from '../src/xlsx/render'
import type { XlsxCell, XlsxCellStyle, XlsxSheet } from '../src/xlsx/types'

const sheet: XlsxSheet = { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [], maxDigitWidthPx: 7 }
const H = 20, INDENT = 3, INDENT_PX = INDENT * 7

function lay(value: number, style: XlsxCellStyle, w: number) {
  const canvas = createCanvas(w + 40, H), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.fillStyle = '#FFFFFF'
  ctx.fillRect(0, 0, w + 40, H)
  const cell: XlsxCell = { ref: 'B1', row: 0, col: 1, value, styleIndex: 0, style }
  // The cell sits 20 px in, with room either side, and nothing limits overflow (empty neighbours).
  const body = layoutCellBody(ctx, sheet, cell, 20, 0, w, H, textMeasurer(ctx), { left: 1000, right: 1000 })
  paintCellLayout(ctx, body, { width: w + 40, height: H })
  const data = ctx.getImageData(0, 0, w + 40, H).data
  let left = -1, right = -1
  for (let x = 0; x < w + 40; x++) for (let y = 0; y < H; y++) {
    const i = (y * (w + 40) + x) * 4
    if (data[i] + data[i + 1] + data[i + 2] < 384) { if (left < 0) left = x; right = x }
  }
  return { text: body.formattedText, font: body.text?.font, left, right }
}

/** The narrowest cell width in which `value` shows `text` unindented. */
function fittingWidth(value: number, style: XlsxCellStyle, text: string): number {
  for (let w = 20; w < 400; w++) if (lay(value, style, w).text === text) return w
  throw new Error(`${text} never fits`)
}

describe('a right-aligned indented number is fitted to its indented box', () => {
  const right = (style: Partial<XlsxCellStyle>): XlsxCellStyle => ({ numFmtId: 0, fontSizePt: 11, horizontal: 'right', ...style })

  it('General drops decimals as it would in a cell narrower by the indent', () => {
    const value = 123456.789
    const w = fittingWidth(value, right({}), '123456.789')
    expect(lay(value, right({}), w).text).toBe('123456.789')
    const indented = lay(value, right({ indent: INDENT }), w)
    const narrower = lay(value, right({}), w - INDENT_PX)
    expect(narrower.text).not.toBe('123456.789')
    // 19c5c5c2: '123456.789', fitted to the full cell and pushed out of its left edge.
    expect(indented.text).toBe(narrower.text)
    expect(indented.left).toBeGreaterThanOrEqual(20)
    expect(indented.right).toBeLessThan(20 + w - INDENT_PX)
  })

  it("a fixed format shows '#' as it would in a cell narrower by the indent", () => {
    const value = 12345.67
    const style = right({ numFmtId: 164, formatCode: '0.00' })
    const w = fittingWidth(value, style, '12345.67')
    expect(lay(value, style, w).text).toBe('12345.67')
    const indented = lay(value, { ...style, indent: INDENT }, w)
    const narrower = lay(value, style, w - INDENT_PX)
    expect(narrower.text).toMatch(/^#+$/)
    expect(indented.text).toBe(narrower.text)
    expect(indented.left).toBeGreaterThanOrEqual(20)
  })

  it('shrink-to-fit shrinks to the indented width', () => {
    const value = 12345.67
    const style = right({ numFmtId: 164, formatCode: '0.00', shrinkToFit: true })
    const w = fittingWidth(value, { ...style, shrinkToFit: false }, '12345.67')
    const full = lay(value, style, w)
    expect(full.text).toBe('12345.67')
    const indented = lay(value, { ...style, indent: INDENT }, w)
    const narrower = lay(value, style, w - INDENT_PX)
    expect(narrower.font).not.toBe(full.font)
    expect(indented.text).toBe('12345.67')
    expect(indented.font).toBe(narrower.font)
    expect(indented.left).toBeGreaterThanOrEqual(20)
    expect(indented.right).toBeLessThan(20 + w - INDENT_PX)
  })
})
