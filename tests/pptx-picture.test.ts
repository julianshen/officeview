import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import { slideMetrics, renderSlide } from '../src/pptx/render'
import { decodeImage } from '../src/core/images'
import { buildPptx } from '../src/testdata/ooxml-builders'

/** 40x40 png: left half red, right half green (for crop tests). */
async function halfRedHalfGreen(): Promise<Uint8Array> {
  const { createCanvas } = await import('canvas')
  const canvas = createCanvas(40, 40)
  const ctx = canvas.getContext('2d')!
  ctx.fillStyle = '#ff0000'
  ctx.fillRect(0, 0, 20, 40)
  ctx.fillStyle = '#00ff00'
  ctx.fillRect(20, 0, 20, 40)
  return new Uint8Array(canvas.toBuffer('image/png'))
}

async function parsePptxFixture(shapes: Parameters<typeof buildPptx>[0]) {
  return parsePptx(await OfficePackage.load(await buildPptx(shapes)))
}

describe('pptx pictures: parse', () => {
  test('resolves p:pic blip rels to image bytes with geometry', async () => {
    const doc = await parsePptxFixture([
      { image: { data: await halfRedHalfGreen() }, off: ['914400', '914400'], ext: ['1828800', '1828800'] },
    ])
    expect(doc.images).toHaveLength(1)
    expect(doc.images[0].mime).toBe('image/png')
    expect(doc.images[0].data.length).toBeGreaterThan(100)
    const pic = doc.slides[0].shapes[0]
    expect(pic.imageIndex).toBe(0)
    expect(pic.xEmu).toBe(914400)
    expect(pic.widthEmu).toBe(1828800)
    expect(pic.heightEmu).toBe(1828800)
  })

  test('parses a:srcRect crop as 0..1 fractions', async () => {
    // l=50000 => 50% cropped from the left
    const doc = await parsePptxFixture([
      { image: { data: await halfRedHalfGreen(), srcRect: { l: 50000 } }, off: ['0', '0'], ext: ['914400', '914400'] },
    ])
    expect(doc.slides[0].shapes[0].image?.srcRect).toEqual({ l: 0.5, t: 0, r: 0, b: 0 })
  })

  test('dedupes the same image across slides/shapes', async () => {
    const png = await halfRedHalfGreen()
    const doc = await parsePptxFixture([
      { image: { data: png }, off: ['0', '0'], ext: ['914400', '914400'] },
      { image: { data: png }, off: ['914400', '0'], ext: ['914400', '914400'] },
    ])
    expect(doc.images).toHaveLength(1)
    expect(doc.slides[0].shapes[0].imageIndex).toBe(0)
    expect(doc.slides[0].shapes[1].imageIndex).toBe(0)
  })
})

describe('pptx pictures: render', () => {
  test('draws picture pixels at the right position and size', async () => {
    const doc = await parsePptxFixture([
      { image: { data: await halfRedHalfGreen() }, off: ['914400', '914400'], ext: ['1828800', '1828800'] },
    ])
    const m = slideMetrics(doc)
    const images = await Promise.all(doc.images.map((img) => decodeImage(img.data, img.mime)))
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(m.widthPx, m.heightPx)
    const ctx = canvas.getContext('2d')!
    renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D, m, images)

    // shape: 1in..3in horizontally, 1in..3in vertically => 96..288 px
    const left = ctx.getImageData(120, 190, 1, 1).data
    const right = ctx.getImageData(260, 190, 1, 1).data
    expect([left[0], left[1], left[2]]).toEqual([255, 0, 0]) // source left = red
    expect([right[0], right[1], right[2]]).toEqual([0, 255, 0]) // source right = green
    // outside the shape is untouched slide background
    const outside = ctx.getImageData(600, 600, 1, 1).data
    expect([outside[0], outside[1], outside[2]]).toEqual([255, 255, 255])
  })

  test('srcRect crop shows only the remaining source region', async () => {
    const doc = await parsePptxFixture([
      { image: { data: await halfRedHalfGreen(), srcRect: { l: 50000 } }, off: ['914400', '914400'], ext: ['1828800', '914400'] },
    ])
    const m = slideMetrics(doc)
    const images = await Promise.all(doc.images.map((img) => decodeImage(img.data, img.mime)))
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(m.widthPx, m.heightPx)
    const ctx = canvas.getContext('2d')!
    renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D, m, images)

    // after cropping the left 50%, the visible content is all green
    const anyPixel = ctx.getImageData(150, 140, 1, 1).data
    expect([anyPixel[1], anyPixel[0]]).toEqual([255, 0])
  })

  test('picture is drawn over shape fill and beneath other shapes', async () => {
    const doc = await parsePptxFixture([
      { image: { data: await halfRedHalfGreen() }, off: ['914400', '914400'], ext: ['1828800', '1828800'] },
      { prst: 'ellipse', off: ['914400', '914400'], ext: ['914400', '914400'], fill: '0000FF' },
    ])
    const m = slideMetrics(doc)
    const images = await Promise.all(doc.images.map((img) => decodeImage(img.data, img.mime)))
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(m.widthPx, m.heightPx)
    const ctx = canvas.getContext('2d')!
    renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D, m, images)

    // later shape (blue ellipse) wins in the overlapping top-left quadrant
    const overlap = ctx.getImageData(140, 140, 1, 1).data
    expect([overlap[2]]).toEqual([255])
    // picture still visible outside the ellipse
    const picOnly = ctx.getImageData(260, 190, 1, 1).data
    expect(picOnly[1]).toBe(255)
  })
})
