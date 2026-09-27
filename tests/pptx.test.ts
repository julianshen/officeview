import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import { slideMetrics, renderSlide } from '../src/pptx/render'
import { buildPptx } from '../src/testdata/ooxml-builders'

describe('pptx parse', () => {
  test('parses slide size, shapes, text runs', async () => {
    const buf = await buildPptx([
      {
        prst: 'rect',
        off: ['914400', '914400'],
        ext: ['3657600', '1828800'],
        fill: 'FFCC00',
        paragraphs: [
          { align: 'ctr', runs: [{ text: 'Slide Title', b: true, sz: '4400', color: '1F497D' }] },
          { runs: [{ text: 'Body line' }] },
        ],
      },
      { prst: 'ellipse', off: ['5486400', '914400'], ext: ['914400', '914400'] },
    ])
    const pkg = await OfficePackage.load(buf)
    const doc = await parsePptx(pkg)
    expect(doc.slideWidthEmu).toBe(9144000)
    expect(doc.slides).toHaveLength(1)
    const [title, ellipse] = doc.slides[0].shapes
    expect(title.xEmu).toBe(914400)
    expect(title.widthEmu).toBe(3657600)
    expect(title.fill).toBe('#FFCC00')
    expect(title.textBody?.paragraphs[0].align).toBe('center')
    expect(title.textBody?.paragraphs[0].runs[0]).toMatchObject({ text: 'Slide Title', bold: true, fontSizePt: 44, color: '#1F497D' })
    expect(ellipse.geometry).toBe('ellipse')
  })
})

describe('pptx render', () => {
  test('paints slide onto canvas with fill and text', async () => {
    const buf = await buildPptx([
      {
        prst: 'rect',
        off: ['914400', '914400'],
        ext: ['3657600', '1828800'],
        fill: 'FFCC00',
        paragraphs: [{ runs: [{ text: 'Hello', b: true, sz: '4400' }] }],
      },
    ])
    const doc = await parsePptx(await OfficePackage.load(buf))
    const slide = doc.slides[0]
    const m = slideMetrics(doc)
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(m.widthPx, m.heightPx)
    const ctx = canvas.getContext('2d')
    renderSlide(slide, ctx as unknown as CanvasRenderingContext2D, m)

    // fill color sample at center of shape (shape center: 914400+1828800 = 2743200 EMU x, y mid)
    const cx = Math.round((2743200 / 914400) * 96)
    const cy = Math.round((1828800 / 914400) * 96)
    const px = ctx.getImageData(cx, cy, 1, 1).data
    expect(px[0]).toBe(255) // R
    expect(px[1]).toBe(204) // G
    expect(px[2]).toBe(0)   // B

    // dark ink from the text somewhere in the shape
    const region = ctx.getImageData(Math.round(914400 / 9525), Math.round(914400 / 9525), 380, 190)
    let ink = 0
    for (let i = 0; i < region.data.length; i += 4) {
      if (region.data[i] < 100 && region.data[i + 1] < 100) ink++
    }
    expect(ink).toBeGreaterThan(50)
  })
})
