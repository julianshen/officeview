// @vitest-environment node
/**
 * PR #55 round 6 P2: the '#' overflow accepts a value by Excel's Calibri advances only
 * for a Calibri cell (76999872 added that fit so a wider fallback face never shows
 * #### for a value Excel shows in Calibri). A cell in another authored face is
 * measured in that face: a fixed-format number that fits Calibri but not Courier New
 * shows '#', unshrunk.
 */
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import { layoutCellBody, textMeasurer } from '../src/xlsx/render'
import { cellFontString } from '../src/xlsx/cellFont'
import type { XlsxCell, XlsxCellStyle, XlsxSheet } from '../src/xlsx/types'

const sheet: XlsxSheet = { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [] }
const W = 64, H = 20, VALUE = 12345.67, TEXT = '12345.67'
/** Calibri 11 pt: seven digits at 1038 and '.' at 517 of 2048 units per em, 14.67 px per em. */
const CALIBRI_PX = (7 * 1038 + 517) / 2048 * 11 * 96 / 72

function lay(style: XlsxCellStyle) {
  const ctx = createCanvas(W, H).getContext('2d') as unknown as CanvasRenderingContext2D
  const cell: XlsxCell = { ref: 'A1', row: 0, col: 0, value: VALUE, styleIndex: 0, style }
  const body = layoutCellBody(ctx, sheet, cell, 0, 0, W, H, textMeasurer(ctx))
  return { text: body.formattedText, font: body.text?.font, measured: textMeasurer(ctx)(cellFontString(style, 11), TEXT) }
}

describe('Calibri metrics decide overflow only for Calibri cells', () => {
  const fixed = (fontName?: string): XlsxCellStyle => ({ numFmtId: 164, formatCode: '0.00', fontSizePt: 11, ...(fontName ? { fontName } : {}) })

  it("a Courier New number that fits Calibri but not its own face shows '#', unshrunk", () => {
    const inner = W - 6
    expect(CALIBRI_PX + 0.5).toBeLessThan(inner)
    const courier = lay(fixed('Courier New'))
    // Courier New (or its metric-compatible Liberation Mono) is wider than the cell.
    expect(courier.measured).toBeGreaterThan(inner)
    // 95809b9c: '12345.67', fitted by shrinking the font although shrinkToFit is off.
    expect(courier.text).toMatch(/^#+$/)
    expect(courier.font).toMatch(/^11pt /)
  })

  it('a Calibri or default-face cell keeps the Calibri fit (76999872)', () => {
    for (const style of [fixed('Calibri'), fixed()]) expect(lay(style).text).toBe(TEXT)
  })
})
