import { describe, test, expect } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { layoutDocx, renderPages, type MeasureFn } from '../src/docx/layout'
import { collectDocImages } from '../src/docx/layout'
import { decodeImage } from '../src/core/images'
import { CT_TYPES, ROOT_RELS } from '../src/testdata/ooxml-builders'

const measureFixed: MeasureFn = (text, style) => text.length * style.fontSizePt * 0.6 * (96 / 72)

/** 24x16 red-over-blue PNG via node-canvas. */
async function tinyPng(): Promise<Uint8Array> {
  const { createCanvas } = await import('canvas')
  const canvas = createCanvas(24, 16)
  const ctx = canvas.getContext('2d')!
  ctx.fillStyle = '#0000ff'
  ctx.fillRect(0, 0, 24, 16)
  ctx.fillStyle = '#ff0000'
  ctx.fillRect(0, 0, 24, 8)
  return new Uint8Array(canvas.toBuffer('image/png'))
}

async function docxWithImage(): Promise<Uint8Array> {
  const png = await tinyPng()
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CT_TYPES)
  zip.file('_rels/.rels', ROOT_RELS)
  zip.file('word/media/image1.png', png)
  zip.file('word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`)
  zip.file('word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Before image</w:t></w:r></w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="914400" cy="609600"/>
            <wp:docPr id="1" name="Picture 1"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>
                  <pic:blipFill><a:blip r:embed="rIdImg1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                  <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="609600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p><w:r><w:t>After image</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`)
  return zip.generateAsync({ type: 'uint8array' })
}

describe('docx embedded images', () => {
  test('parses image with EMU extents from w:drawing + rels', async () => {
    const doc = await parseDocx(await OfficePackage.load(await docxWithImage()))
    const paras = doc.sections[0].paragraphs
    expect(paras).toHaveLength(3)
    expect(paras[1].images).toHaveLength(1)
    expect(paras[1].images[0].widthEmu).toBe(914400)
    expect(paras[1].images[0].heightEmu).toBe(609600)
    expect(paras[1].images[0].mime).toBe('image/png')
    expect(collectDocImages(doc)).toHaveLength(1)
  })

  test('image occupies flow space between text', async () => {
    const doc = await parseDocx(await OfficePackage.load(await docxWithImage()))
    const pages = layoutDocx(doc, measureFixed)
    expect(pages[0].images).toHaveLength(1)
    const box = pages[0].images[0]
    expect(box.widthPx).toBeCloseTo(914400 / 914400 * 96, 0) // 1 inch = 96px
    expect(box.heightPx).toBeCloseTo(64, 0)
    expect(box.imageIndex).toBe(0)
    // the "After image" line is below the image box
    const after = pages[0].lines.find((l) => l.segs.some((s) => s.text.includes('After')))
    expect(after!.yPx).toBeGreaterThanOrEqual(box.yPx + box.heightPx)
  })

  test('paints decoded image pixels onto canvas', async () => {
    const doc = await parseDocx(await OfficePackage.load(await docxWithImage()))
    const pages = layoutDocx(doc, measureFixed)
    const images = await Promise.all(collectDocImages(doc).map((img) => decodeImage(img.data, img.mime)))
    const page = pages[0]
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(Math.ceil(page.widthPx), Math.ceil(page.heightPx))
    const ctx = canvas.getContext('2d')!
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, canvas.width, canvas.height)
    renderPages(pages, ctx as unknown as CanvasRenderingContext2D, images)
    const box = pages[0].images[0]
    // top half red, bottom half blue (stretched 24x16 -> 96x64)
    const top = ctx.getImageData(Math.round(box.xPx + box.widthPx / 2), Math.round(box.yPx + 4), 1, 1).data
    expect([top[0], top[1], top[2]]).toEqual([255, 0, 0])
    const bottom = ctx.getImageData(Math.round(box.xPx + box.widthPx / 2), Math.round(box.yPx + box.heightPx - 4), 1, 1).data
    expect([bottom[0], bottom[1], bottom[2]]).toEqual([0, 0, 255])
  })
})
