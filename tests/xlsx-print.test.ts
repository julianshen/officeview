import { describe, expect, test } from 'vitest'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics, computePrintMetrics, renderPrintPage, PAPER_SIZES_IN, STATIC_SHEET_BUDGET, StaticExportLimitError } from '../src/xlsx/render'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import type { XlsxSheet } from '../src/xlsx/types'

async function printSheet(setup?: { paperSize?: number; orientation?: string; scale?: number; fitToWidth?: number; fitToHeight?: number; fitToPage?: boolean }) {
  const rows = []
  for (let r = 1; r <= 10; r++) {
    const cells = []
    for (let c = 0; c < 5; c++) cells.push({ ref: `${String.fromCharCode(65 + c)}${r}`, v: `R${r}C${c}` })
    rows.push({ r, cells })
  }
  const bytes = await buildXlsx([{ name: 'P', rows, pageSetup: setup, pageMargins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 } }])
  const doc = await parseXlsx(await OfficePackage.load(bytes))
  return doc.sheets[0]
}

describe('xlsx print setup', () => {
  test('parses pageSetup, fit flags and margins', async () => {
    const sheet = await printSheet({ paperSize: 9, orientation: 'landscape', scale: 50, fitToWidth: 1, fitToHeight: 0, fitToPage: true })
    expect(sheet.pageSetup).toMatchObject({ paperSizeId: 9, orientation: 'landscape', scale: 50, fitToWidth: 1, fitToHeight: 0, fitToPage: true })
    expect(sheet.pageMargins).toMatchObject({ left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 })
  })
  test('absent setup leaves print fields undefined', async () => {
    const sheet = await printSheet(undefined)
    expect(sheet.pageSetup).toBeUndefined()
    expect(sheet.pageMargins?.left).toBe(0.7)
  })
  test('letter landscape geometry yields the printable rect', async () => {
    const sheet = await printSheet({ paperSize: 1, orientation: 'landscape' })
    const m = computeMetrics(sheet)
    const print = computePrintMetrics(sheet, m, 96)
    // 11x8.5in at 96dpi, minus 0.7in side margins.
    expect([print.paperPx.width, print.paperPx.height]).toEqual([1056, 816])
    expect(print.printable.x).toBeCloseTo(0.7 * 96, 6)
    expect(print.printable.width).toBeCloseTo((11 - 1.4) * 96, 6)
    expect(print.scale).toBe(1)
  })
  test('explicit scale shrinks content into the printable area', async () => {
    const sheet = await printSheet({ scale: 50 })
    const print = computePrintMetrics(sheet, computeMetrics(sheet), 96)
    expect(print.scale).toBe(0.5)
    expect(print.pagesWide).toBe(1)
    expect(print.pagesTall).toBe(1)
  })
  test('fit-to-width distributes pages and rescales', async () => {
    const sheet = await printSheet({ fitToWidth: 1, fitToHeight: 0, fitToPage: true })
    const m = computeMetrics(sheet)
    const print = computePrintMetrics(sheet, m, 96)
    expect(print.scale).toBeLessThanOrEqual(1)
    expect(print.scale).toBeGreaterThan(0)
    expect(print.pagesWide).toBe(1)
  })
  test('print page renders paper-sized with white margins and scaled content', async () => {
    const sheet = await printSheet({ paperSize: 1, orientation: 'landscape', scale: 50 })
    const m = computeMetrics(sheet)
    const print = computePrintMetrics(sheet, m, 96)
    const canvas = createCanvas(print.paperPx.width, print.paperPx.height)
    renderPrintPage(sheet, canvas.getContext('2d') as never, m, print)
    const ctx = canvas.getContext('2d')
    const at = (x: number, y: number) => Array.from(ctx.getImageData(x, y, 1, 1).data).slice(0, 3)
    // Margin corner stays paper white; content area carries ink.
    expect(at(5, 5)).toEqual([255, 255, 255])
    const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data
    let ink = 0
    for (let i = 0; i < data.length; i += 4) if (data[i] < 128) ink++
    expect(ink).toBeGreaterThan(100)
  })
  describe('explicit static print budgets', () => {
    const handBuilt = (row: number, col: number): XlsxSheet => ({
      name: 'Huge', cols: [], merges: [], mergeRanges: [],
      rows: [{ index: 0, cells: [{ ref: 'A1', row: 0, col: 0, value: 'A1', styleIndex: 0 }] },
        { index: row, cells: [{ ref: 'Far', row, col, value: 'Far', styleIndex: 0 }] }],
    })
    const counting = () => {
      const canvas = createCanvas(64, 64), raw = canvas.getContext('2d')
      let calls = 0
      const ctx = new Proxy(raw, {
        get(target, property) {
          const value = Reflect.get(target, property, target)
          return typeof value === 'function' ? (...args: unknown[]) => { calls++; return value.apply(target, args) } : value
        },
        set(target, property, value) { return Reflect.set(target, property, value, target) },
      }) as unknown as CanvasRenderingContext2D
      return { ctx, calls: () => calls }
    }
    test('a full-run page quota is reported, but one bounded page is still admitted', () => {
      // Inside the 16,384 x 4,096 grid, but ~139k Letter pages at 100%.
      const sheet = handBuilt(16_383, 4_095)
      const m = computeMetrics(sheet)
      expect(m.exportLimits?.some(limit => limit.resource === 'rows' || limit.resource === 'cols')).toBeFalsy()
      const print = computePrintMetrics(sheet, m, 96)
      expect(print.pagesWide * print.pagesTall).toBeGreaterThan(STATIC_SHEET_BUDGET.maxPrintPages)
      expect(print.exportLimits).toEqual([expect.objectContaining({ code: 'large-export-limit', resource: 'pages', limit: STATIC_SHEET_BUDGET.maxPrintPages })])
      const out = counting()
      expect(() => renderPrintPage(sheet, out.ctx, m, print)).not.toThrow()
      expect(out.calls()).toBeGreaterThan(0)
    })
    test('a sheet beyond the static grid prints bounded pages but refuses the truncated edge page', () => {
      const sheet = handBuilt(16_385, 0)
      const m = computeMetrics(sheet)
      const print = computePrintMetrics(sheet, m, 96)
      expect(print.exportLimits).toEqual([expect.objectContaining({ resource: 'rows', requested: 16_386, limit: STATIC_SHEET_BUDGET.maxGridRows })])
      expect(() => renderPrintPage(sheet, counting().ctx, m, print, 0, 0)).not.toThrow()
      const out = counting()
      expect(() => renderPrintPage(sheet, out.ctx, m, print, 0, print.pagesTall - 1)).toThrow(StaticExportLimitError)
      expect(out.calls()).toBe(0)
    })
    test('a page outside the page grid is rejected instead of painting a blank page', async () => {
      const sheet = await printSheet({ scale: 50 })
      const m = computeMetrics(sheet)
      const print = computePrintMetrics(sheet, m, 96)
      expect(print.exportLimits).toBeUndefined()
      const out = counting()
      expect(() => renderPrintPage(sheet, out.ctx, m, print, print.pagesWide, 0)).toThrow(RangeError)
      expect(() => renderPrintPage(sheet, out.ctx, m, print, 0, -1)).toThrow(RangeError)
      expect(out.calls()).toBe(0)
      expect(() => renderPrintPage(sheet, out.ctx, m, print, 0, 0)).not.toThrow()
    })
  })
  test('paper size table covers common ids with letter default', () => {
    expect(PAPER_SIZES_IN[1]).toEqual([8.5, 11])
    expect(PAPER_SIZES_IN[9]).toEqual([8.27, 11.69])
    expect(PAPER_SIZES_IN[3]).toEqual([11, 17])
  })
})
