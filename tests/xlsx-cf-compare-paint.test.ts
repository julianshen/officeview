// @vitest-environment node
import { createCanvas } from 'canvas'
import { afterEach, expect, it, vi } from 'vitest'
import * as compare from '../scripts/cf-compare'
import type { XlsxSheet } from '../src/xlsx/types'
afterEach(() => vi.restoreAllMocks())
const sheet = (): XlsxSheet => ({ name: 'S', cols: [], merges: [], mergeRanges: [], rows: [
  { index: 0, cells: [{ row: 0, col: 0, ref: 'A1', value: null, styleIndex: 0, style: { numFmtId: 0, fillColor: 'FFFF0000' } }] },
  { index: 1, cells: [{ row: 1, col: 1, ref: 'B2', value: null, styleIndex: 0, style: { numFmtId: 0, fillColor: 'FF00FF00' } }] },
] })
it('crops a B2 capture from B2, and print paints at 2x before density conversion', async () => {
  const scale = vi.spyOn(Object.getPrototypeOf(createCanvas(1, 1).getContext('2d')), 'scale')
  const result = await compare.paintCfCapture({ sheets: [sheet()] }, 0, 'B2:B2', 'print', 4 / 3, 64, 20)
  expect(scale).toHaveBeenCalledWith(2, 2)
  expect([...result.getContext('2d').getImageData(5, 5, 1, 1).data]).toEqual([0, 255, 0, 255])
})
it('reports static refusal before painting a refusal banner as reference pixels', async () => {
  const s = sheet(); s.rows.push({ index: 20000, cells: [{ row: 20000, col: 0, ref: 'A20001', value: 1, styleIndex: 0 }] })
  await expect(compare.paintCfCapture({ sheets: [s] }, 0, 'A1:A1', 'static', 4 / 3, 64, 20)).rejects.toThrow(/refus/i)
})
