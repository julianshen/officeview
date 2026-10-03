import { describe, expect, test } from 'vitest'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { collectDocImages, layoutDocx } from '../src/docx/layout'
import { parsePptx } from '../src/pptx/parse'
import { renderSlide, slideMetrics } from '../src/pptx/render'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics, renderSheet } from '../src/xlsx/render'

const corpus = join(__dirname, '..', 'corpus')
const available = existsSync(join(corpus, 'poi-table_test.pptx'))
const load = async (name: string) => OfficePackage.load(readFileSync(join(corpus, name)))
const measure = (text: string, style: { fontSizePt: number }) => text.length * style.fontSizePt * 0.6 * 96 / 72

// Expectations come from the native Office PDFs retained in validation/office-reference.
describe.skipIf(!available)('native Office rendering regressions', () => {
  test('auto-width SDT table keeps short cells adjacent', async () => {
    const doc = await parseDocx(await load('poi-Bug66263-table.docx'))
    const page = layoutDocx(doc, measure)[0]
    expect(page.tables[0].widthPx).toBeLessThan(220)
    const [left, right] = page.tables[0].rows[0].cells
    expect(right.xPx - left.xPx).toBeLessThan(110)
    expect(page.lines).toHaveLength(2)
  })

  test('header relationship images are collected and placed with body images', async () => {
    const doc = await parseDocx(await load('pydocx-having-images.docx'))
    expect(doc.sections[0].header?.flatMap(p => p.images)).toHaveLength(1)
    expect(collectDocImages(doc)).toHaveLength(6)
    const page = layoutDocx(doc, measure)[0]
    expect(page.images).toHaveLength(6)
    expect(page.images.some(image => image.yPx === 48)).toBe(true)
  })

  test('image-only body paragraphs start at their flow position without a blank line', async () => {
    const doc = await parseDocx(await load('pydocx-having-images.docx'))
    const page = layoutDocx(doc, measure)[0]
    const body = page.images.filter(image => image.yPx >= 96)
    expect(body[0].yPx).toBe(96)
    expect(body[1].yPx - body[0].yPx - body[0].heightPx).toBeCloseTo(12700 / 9525)
    expect(body[4].yPx - body[3].yPx - body[3].heightPx).toBeCloseTo(12065 / 9525)
  })

  test('PowerPoint resolves real DrawingML table regions and theme tints', async () => {
    const doc = await parsePptx(await load('poi-table_test.pptx'))
    const table = doc.slides[0].shapes.find(shape => shape.table)!.table!
    expect(table.styleFills).toMatchObject({ firstRow: '#5B9BD5', band1: '#D2DEEF', wholeTable: '#EAEFF7' })
    const metrics = slideMetrics(doc)
    const ctx = createCanvas(metrics.widthPx, metrics.heightPx).getContext('2d')
    renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D, metrics)
    const rgb = (x: number, y: number) => [...ctx.getImageData(x, y, 1, 1).data].slice(0, 3)
    expect(rgb(210, 145)).toEqual([91, 155, 213])
    expect(rgb(210, 185)).toEqual([210, 222, 239])
    expect(rgb(210, 220)).toEqual([234, 239, 247])
    // White separators, including the thicker bottom edge of the header.
    expect(rgb(210, 166)).toEqual([255, 255, 255])
    expect(rgb(358, 185).every(channel => channel > 220)).toBe(true)
  })

  test('Excel explicit black borders survive worksheet grid painting', async () => {
    const doc = await parseXlsx(await load('poi-56295.xlsx'))
    const sheet = doc.sheets[0]
    const metrics = computeMetrics(sheet)
    const ctx = createCanvas(metrics.widthPx, metrics.heightPx).getContext('2d')
    renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D, metrics)
    expect([...ctx.getImageData(0, 30, 1, 1).data].slice(0, 3)).toEqual([0, 0, 0])
    expect([...ctx.getImageData(metrics.widthPx - 1, 30, 1, 1).data].slice(0, 3)).toEqual([0, 0, 0])
    expect([...ctx.getImageData(25, 60, 1, 1).data].slice(0, 3)).toEqual([0, 0, 0])
  })
})
