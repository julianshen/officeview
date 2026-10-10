// @vitest-environment node
import { createCanvas } from 'canvas'
import { expect, test } from 'vitest'
import { collectDocImages } from '../src/docx/layout'
import type { DocxDocument, DocxDrawingShape, DocxImage } from '../src/docx/types'
import { getPaintables } from '../src/render/paint'

test('DOCX prepares and paints a cached picture inside nested diagram groups', async () => {
  const source = createCanvas(2, 2), sourceCtx = source.getContext('2d')
  sourceCtx.fillStyle = '#f00'; sourceCtx.fillRect(0, 0, 2, 2)
  const picture = { data: new Uint8Array(source.toBuffer('image/png')), mime: 'image/png' }
  const shape: DocxDrawingShape = { geometry: 'rect', xEmu: 0, yEmu: 0, widthEmu: 381000, heightEmu: 381000, paragraphs: [], fontFamily: 'Calibri' }
  const leaf = { ...shape, image: picture }
  const group: DocxDrawingShape = { ...shape, group: { off: { x: 0, y: 0 }, ext: { width: 381000, height: 381000 }, chOff: { x: 0, y: 0 }, chExt: { width: 381000, height: 381000 } }, children: [leaf] }
  const owner: DocxImage = { data: new Uint8Array(), widthEmu: 381000, heightEmu: 381000, drawing: { kind: 'diagram', shapes: [{ ...group, children: [group] }] } }
  const paragraph = { align: 'left' as const, runs: [], images: [owner] }
  const doc: DocxDocument = { defaultFontFamily: 'Calibri', defaultFontSizePt: 11, styleDefaults: new Map(), sections: [{ margins: { topTwips: 1440, bottomTwips: 1440, leftTwips: 1440, rightTwips: 1440, headerTwips: 720, footerTwips: 720, gutterTwips: 0 }, pageSize: { widthTwips: 12240, heightTwips: 15840, orientation: 'portrait' }, paragraphs: [paragraph], blocks: [{ kind: 'p', paragraph }] }] }
  expect(collectDocImages(doc)).toEqual([owner, picture])
  const pages = await getPaintables(doc)
  try {
    const ctx = createCanvas(816, 1056).getContext('2d')
    pages[0].paint(ctx as unknown as CanvasRenderingContext2D)
    expect([...ctx.getImageData(110, 110, 1, 1).data]).toEqual([255, 0, 0, 255])
  } finally { pages.dispose() }
  // Collection terminates on repeated references and cycles in handwritten models.
  group.children!.push(group)
  expect(collectDocImages(doc)).toEqual([owner, picture])
})
