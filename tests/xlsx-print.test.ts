import { describe, expect, test } from 'vitest'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics, computePrintMetrics, renderPrintPage, PAPER_SIZES_IN } from '../src/xlsx/render'
import { buildXlsx } from '../src/testdata/ooxml-builders'

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
  test('paper size table covers common ids with letter default', () => {
    expect(PAPER_SIZES_IN[1]).toEqual([8.5, 11])
    expect(PAPER_SIZES_IN[9]).toEqual([8.27, 11.69])
    expect(PAPER_SIZES_IN[3]).toEqual([11, 17])
  })
})
