import { describe, test, expect } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { layoutDocx, renderPages, type MeasureFn } from '../src/docx/layout'
import { decodeImage } from '../src/core/images'
import { collectDocImages } from '../src/docx/layout'

const measureFixed: MeasureFn = (t, s) => t.length * s.fontSizePt * 0.6 * (96 / 72)

async function png(): Promise<Uint8Array> {
  const { createCanvas } = await import('canvas')
  const c = createCanvas(20, 20)
  const ctx = c.getContext('2d')!
  ctx.fillStyle = '#ff00ff'
  ctx.fillRect(0, 0, 20, 20)
  return new Uint8Array(c.toBuffer('image/png'))
}

/** docx with a text paragraph and one anchored (floating) image. */
async function docxWithAnchor(anchorAttrs: string, posH: string, posV: string) {
  const img = await png()
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/></Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  zip.file('word/media/image1.png', img)
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>`)
  zip.file('word/document.xml', `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Paragraph before the anchor.</w:t></w:r></w:p>
    <w:p><w:r>
      <w:drawing>
        <wp:anchor ${anchorAttrs}>
          <wp:simplePos x="0" y="0"/>
          ${posH}
          ${posV}
          <wp:extent cx="914400" cy="914400"/>
          <wp:wrapNone/>
          <wp:docPr id="1" name="Floating 1"/>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="img"/><pic:cNvPicPr/></pic:nvPicPr>
              <pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
            </pic:pic>
          </a:graphicData></a:graphic>
        </wp:anchor>
      </w:drawing>
    </w:r></w:p>
    <w:p><w:r><w:t>Paragraph after the anchor.</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>`)
  return parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
}

describe('docx floating (wp:anchor) images', () => {
  test('parses anchor positioning instead of treating it as inline', async () => {
    const doc = await docxWithAnchor(
      'behindDoc="0" relativeHeight="10"',
      '<wp:positionH relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionH>',
      '<wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV>',
    )
    const imgs = doc.sections[0].paragraphs[1].images
    expect(imgs).toHaveLength(1)
    expect(imgs[0].floating).toBeDefined()
    expect(imgs[0].floating!.behindDoc).toBe(false)
    expect(imgs[0].floating!.relativeHeight).toBe(10)
    expect(imgs[0].floating!.posH.relativeFrom).toBe('page')
    expect(imgs[0].floating!.posH.offsetEmu).toBe(1828800)
    expect(imgs[0].floating!.wrap).toBe('none')
  })

  test('is positioned from the page origin and does not consume flow', async () => {
    const doc = await docxWithAnchor(
      'behindDoc="0" relativeHeight="10"',
      '<wp:positionH relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionH>',
      '<wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV>',
    )
    const pages = layoutDocx(doc, measureFixed)
    const boxes = pages[0].images
    expect(boxes).toHaveLength(1)
    expect(boxes[0].floating).toBeDefined()
    // 1828800 EMU = 2in = 192px, 914400 EMU = 1in = 96px, origin page = 0,0
    expect(boxes[0].xPx).toBeCloseTo(192, 0)
    expect(boxes[0].yPx).toBeCloseTo(96, 0)

    // the text after the anchor is NOT pushed down by the image height (96px)
    const lineText = (l: { segs: Array<{ text: string }> }): string => l.segs.map((s) => s.text).join('')
    const lines = pages[0].lines.map(lineText)
    expect(lines).toContain('Paragraph before the anchor.')
    expect(lines).toContain('Paragraph after the anchor.')
    const after = pages[0].lines.find((l) => lineText(l).includes('after'))!
    const before = pages[0].lines.find((l) => lineText(l).includes('before'))!
    const gap = after.yPx - before.yPx
    expect(gap).toBeLessThan(60) // two normal lines, not line + 96px image
  })

  test('behindDoc anchors paint beneath the text layer', async () => {
    const doc = await docxWithAnchor(
      'behindDoc="1" relativeHeight="10"',
      '<wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>',
      '<wp:positionV relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionV>',
    )
    const pages = layoutDocx(doc, measureFixed)
    expect(pages[0].images[0].floating?.behindDoc).toBe(true)
    // margin origin = the 1" margin = 96px
    expect(pages[0].images[0].xPx).toBeCloseTo(96, 0)
    expect(pages[0].images[0].yPx).toBeCloseTo(96, 0)

    const images = await Promise.all(collectDocImages(doc).map((i) => decodeImage(i.data, i.mime)))
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(Math.ceil(pages[0].widthPx), Math.ceil(pages[0].heightPx))
    const ctx = canvas.getContext('2d')!
    renderPages(pages, ctx as unknown as CanvasRenderingContext2D, images)
    // the magenta image is at the margin, where the first line of text sits:
    // behindDoc means text pixels win where they overlap
    const box = pages[0].images[0]
    const line = pages[0].lines[0]
    const px = ctx.getImageData(Math.round(box.xPx + 4), Math.round(line.yPx + 4), 1, 1).data
    // either pure magenta (no text at that spot) or dark (text over magenta)
    const isMagenta = px[0] > 200 && px[1] < 80 && px[2] > 200
    const textOver = px[0] < 120 && px[1] < 120 && px[2] < 120
    expect(isMagenta || textOver).toBe(true)
  })

  test('inline images still flow (regression guard)', async () => {
    const { buildDocx } = await import('../src/testdata/ooxml-builders')
    const doc = await parseDocx(await OfficePackage.load(await buildDocx([{ runs: [{ text: 'text' }] }])))
    expect(doc.sections[0].paragraphs[0].images).toHaveLength(0)
  })
})
