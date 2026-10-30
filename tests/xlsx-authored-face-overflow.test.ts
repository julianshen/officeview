// @vitest-environment node
/**
 * PR #55 round 6 P2 (and its fix round 2): Excel's Calibri advances stand in for the
 * '#' overflow decision whenever a cell is not measured in its own face (76999872: a
 * wider fallback face must not show #### for a value Excel shows). An authored face
 * that this host draws (Courier New) is measured in itself: a fixed-format number that
 * fits Calibri but not Courier New shows '#', unshrunk. An authored face the host lacks
 * (Aptos Narrow, Excel's current default, on most non-Office hosts) paints in the
 * fallback stack and keeps the Calibri leniency.
 */
import { createCanvas } from 'canvas'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { layoutCellBody, textMeasurer } from '../src/xlsx/render'
import * as cellFont from '../src/xlsx/cellFont'
import type { XlsxCell, XlsxCellStyle, XlsxSheet } from '../src/xlsx/types'

const sheet: XlsxSheet = { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [] }
const W = 64, H = 20, VALUE = 12345.67, TEXT = '12345.67'
/** Calibri 11 pt: seven digits at 1038 and '.' at 517 of 2048 units per em, 14.67 px per em. */
const CALIBRI_PX = (7 * 1038 + 517) / 2048 * 11 * 96 / 72

function lay(style: XlsxCellStyle) {
  const ctx = createCanvas(W, H).getContext('2d') as unknown as CanvasRenderingContext2D
  const cell: XlsxCell = { ref: 'A1', row: 0, col: 0, value: VALUE, styleIndex: 0, style }
  const body = layoutCellBody(ctx, sheet, cell, 0, 0, W, H, textMeasurer(ctx))
  return { text: body.formattedText, font: body.text?.font, measured: textMeasurer(ctx)(cellFont.cellFontString(style, 11), TEXT) }
}
const fixed = (fontName?: string): XlsxCellStyle => ({ numFmtId: 164, formatCode: '0.00', fontSizePt: 11, ...(fontName ? { fontName } : {}) })
/** Availability made explicit: these faces are (or are not) drawn as themselves. */
const available = (faces: Record<string, boolean>) => vi.spyOn(cellFont, 'cellFaceAvailable').mockImplementation((_ctx, name) => faces[name ?? ''] ?? false)

afterEach(() => vi.restoreAllMocks())

describe('Calibri metrics decide overflow unless the cell is measured in its own face', () => {
  it("an available Courier New number that fits Calibri but not its own face shows '#', unshrunk", () => {
    const spy = available({ 'Courier New': true })
    const inner = W - 6
    expect(CALIBRI_PX + 0.5).toBeLessThan(inner)
    const courier = lay(fixed('Courier New'))
    expect(spy).toHaveBeenCalled()
    // Courier New (or its metric-compatible Liberation Mono, or the fallback stack) is wider than the cell.
    expect(courier.measured).toBeGreaterThan(inner)
    // 95809b9c: '12345.67', fitted by shrinking the font although shrinkToFit is off.
    expect(courier.text).toMatch(/^#+$/)
    expect(courier.font).toMatch(/^11pt /)
  })

  it('a missing authored face (Aptos Narrow here) keeps the Calibri fit, shown in full as before', () => {
    available({ 'Aptos Narrow': false, 'Courier New': false })
    // c858c1d9: '####' — the fallback face's width is not Excel's.
    expect(lay(fixed('Aptos Narrow')).text).toBe(TEXT)
    expect(lay(fixed('Courier New')).text).toBe(TEXT)
  })

  it('a Calibri or default-face cell keeps the Calibri fit (76999872), without probing', () => {
    const spy = available({})
    for (const style of [fixed('Calibri'), fixed()]) expect(lay(style).text).toBe(TEXT)
    expect(spy).not.toHaveBeenCalled()
  })
})

describe('cellFaceAvailable probes the backend once per face', () => {
  it('a face that does not exist is unavailable; a common installed face is available', () => {
    const ctx = createCanvas(10, 10).getContext('2d') as unknown as CanvasRenderingContext2D
    ctx.font = '13px serif'
    expect(cellFont.cellFaceAvailable(ctx, 'No Such Face Xyzzy')).toBe(false)
    expect(ctx.font).toBe('13px serif')
    // One of the faces every supported host carries (macOS: Courier New/Arial; Linux CI: Liberation/DejaVu).
    expect(['Courier New', 'Arial', 'Liberation Sans', 'Liberation Mono', 'DejaVu Sans'].some(face => cellFont.cellFaceAvailable(ctx, face))).toBe(true)
  })

  it('a backend without a measurement signal reports unavailable (the Calibri leniency stays)', () => {
    const flat = { font: '', measureText: () => ({ width: 100 }) } as unknown as CanvasRenderingContext2D
    expect(cellFont.cellFaceAvailable(flat, 'Probe Flat Backend Face')).toBe(false)
  })
})
