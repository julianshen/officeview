/**
 * Wrapped, unrotated cell text tries its widest candidate box first; a
 * narrower box only wraps into more lines, so when the widest is already too
 * tall no candidate can fit and the rest are skipped. Before, a wrapped value
 * taller than its cell laid itself out once per candidate (two per line that
 * fits), which made layout quadratic in the text's line count.
 */
import { describe, expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import * as textLayout from '../src/drawing/text-layout'
import { layoutCellBody, textMeasurer } from '../src/xlsx/render'
import type { XlsxCell, XlsxSheet } from '../src/xlsx/types'

const sheet: XlsxSheet = { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [] }
const ctx = () => createCanvas(200, 200).getContext('2d') as unknown as CanvasRenderingContext2D
const wrapped = (value: string, textRotation?: number): XlsxCell => ({ ref: 'A1', row: 0, col: 0, value, styleIndex: 0, style: { numFmtId: 0, wrapText: true, ...(textRotation !== undefined ? { textRotation } : {}) } })

describe('wrapped cell layout candidates', () => {
  it('text taller than its box is laid out a bounded number of times, not once per candidate', () => {
    const spy = vi.spyOn(textLayout, 'layoutTextBody')
    try {
      const c = ctx()
      const body = layoutCellBody(c, sheet, wrapped('a\n'.repeat(300)), 0, 0, 64, 4000, textMeasurer(c))
      expect(body.text?.kind).toBe('layout')
      // Every line is still laid out (the overflow is clipped, never dropped).
      expect(body.text?.kind === 'layout' && body.text.lines).toBeGreaterThanOrEqual(300)
      // One layout: the widest candidate is the fallback and is not laid out twice.
      expect(spy).toHaveBeenCalledTimes(1)
    } finally { spy.mockRestore() }
  })

  it('text that fits keeps the widest fitting box, laid out once', () => {
    const spy = vi.spyOn(textLayout, 'layoutTextBody')
    try {
      const c = ctx()
      const body = layoutCellBody(c, sheet, wrapped('alpha beta gamma'), 0, 0, 64, 400, textMeasurer(c))
      expect(body.text?.kind === 'layout' && body.text.lines).toBe(3)
      expect(spy).toHaveBeenCalledTimes(1)
    } finally { spy.mockRestore() }
  })

  it('rotated wrapped text still searches its candidates for a fitting box', () => {
    const c = ctx()
    const body = layoutCellBody(c, sheet, wrapped('alpha beta gamma delta', 45), 0, 0, 64, 80, textMeasurer(c))
    expect(Math.abs(body.text?.kind === 'layout' ? body.text.rotation ?? 0 : 0)).toBeCloseTo(Math.PI / 4, 9)
  })
})
