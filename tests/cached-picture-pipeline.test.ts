import { expect, test, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { getPaintables } from '../src/render/paint'
import type { PptxDocument } from '../src/pptx/types'
import type { XlsxDocument } from '../src/xlsx/types'
import type { DrawingContent } from '../src/drawing/content'

test.each(['pptx', 'xlsx'] as const)('%s prepares and paints cached pictures nested in a diagram group', async format => {
  const raster = createCanvas(10, 10)
  const rasterCtx = raster.getContext('2d')
  rasterCtx.fillStyle = '#FF00FF'
  rasterCtx.fillRect(0, 0, 10, 10)
  const image = { data: new Uint8Array(raster.toBuffer('image/png')), mime: 'image/png' }
  const diagram: DrawingContent = { kind: 'diagram', shapes: [{
    xEmu: 0, yEmu: 0, widthEmu: 952500, heightEmu: 952500, geometry: 'group', paragraphs: [], fontFamily: '',
    group: { off: { x: 0, y: 0 }, ext: { width: 952500, height: 952500 }, chOff: { x: 0, y: 0 }, chExt: { width: 952500, height: 952500 } },
    children: [{ xEmu: 0, yEmu: 0, widthEmu: 952500, heightEmu: 952500, geometry: 'rect', paragraphs: [], fontFamily: '', image }],
  }] }
  const owner = { xEmu: 0, yEmu: 0, widthEmu: 952500, heightEmu: 952500, geometry: 'other' as const, content: diagram }
  const doc = format === 'pptx'
    ? { slideWidthEmu: 1905000, slideHeightEmu: 1905000, slides: [{ index: 0, widthEmu: 1905000, heightEmu: 1905000, shapes: [owner] }], images: [] } as unknown as PptxDocument
    : { sheets: [{ name: 'Cached', rows: [], cols: [], merges: [], mergeRanges: [], maxRow: 1, maxCol: 1, drawings: [owner] }], images: [] } as unknown as XlsxDocument
  const decodeImage = vi.fn(async () => raster as unknown as CanvasImageSource)
  const pages = await getPaintables(doc, { decodeImage })
  try {
    const ctx = createCanvas(pages[0].spec.widthPx, pages[0].spec.heightPx).getContext('2d')
    pages[0].paint(ctx as unknown as CanvasRenderingContext2D)
    expect(decodeImage).toHaveBeenCalledTimes(1)
    expect(Array.from(ctx.getImageData(50, 50, 1, 1).data)).toEqual([255, 0, 255, 255])
  } finally { pages.dispose() }
})
