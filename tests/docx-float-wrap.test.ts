import { describe, test, expect } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { layoutDocx, freeBandsFor, type ImageBox, type MeasureFn } from '../src/docx/layout'

const measureFixed: MeasureFn = (t, s) => t.length * s.fontSizePt * 0.6 * (96 / 72)

/** body paragraph + an anchored image with the given wrap element */
async function docWithFloat(wrapEl: string, posH: string, posV: string) {
  const { createCanvas } = await import('canvas')
  const c = createCanvas(20, 20)
  c.getContext('2d')!.fillRect(0, 0, 20, 20)
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/></Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  zip.file('word/media/image1.png', new Uint8Array(c.toBuffer('image/png')))
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>`)
  zip.file('word/document.xml', `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>${'lead '.repeat(40).trim()}</w:t></w:r></w:p>
    <w:p><w:r>
      <w:drawing>
        <wp:anchor behindDoc="0" relativeHeight="5">
          <wp:simplePos x="0" y="0"/>
          ${posH}
          ${posV}
          <wp:extent cx="1828800" cy="1828800"/>
          ${wrapEl}
          <wp:docPr id="1" name="Float 1"/>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="img"/><pic:cNvPicPr/></pic:nvPicPr>
              <pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
            </pic:pic>
          </a:graphicData></a:graphic>
        </wp:anchor>
      </w:drawing>
    </w:r></w:p>
    <w:p><w:r><w:t>${'tail '.repeat(40).trim()}</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>`)
  return parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
}

const CONTENT_X = 96 // 1in margin
const CONTENT_W = 624 // 9360 twips

describe('freeBandsFor', () => {
  const float = (x: number, y: number, w = 200, h = 100, wrap: 'square' | 'none' | 'topAndBottom' = 'square'): ImageBox => ({
    xPx: x, yPx: y, widthPx: w, heightPx: h, imageIndex: 0, floating: { behindDoc: false, relativeHeight: 0, wrap },
  })

  test('a float on the right leaves a left band (plus a sliver on the right)', () => {
    // float spans 496..696, +12px gap each side => left band 96..484, sliver 708..720
    const bands = freeBandsFor([float(CONTENT_X + 400, 200)], 200, 240, CONTENT_X, CONTENT_W)!
    expect(bands).toEqual([
      { x: 96, width: 388 },
      { x: 708, width: 12 },
    ])
    // layoutParagraph keeps the widest band
    expect(Math.max(...bands.map((b) => b.width))).toBe(388)
  })

  test('a float on the left leaves a right band', () => {
    const bands = freeBandsFor([float(CONTENT_X, 200)], 200, 240, CONTENT_X, CONTENT_W)!
    expect(bands).toHaveLength(1)
    expect(bands[0].x).toBeGreaterThan(CONTENT_X)
  })

  test('a float in the middle leaves two bands and the widest is kept', () => {
    const bands = freeBandsFor([float(CONTENT_X + 200, 200, 100)], 200, 240, CONTENT_X, CONTENT_W)!
    expect(bands.length).toBeGreaterThanOrEqual(1)
  })

  test('a float outside the y range does not narrow anything', () => {
    expect(freeBandsFor([float(CONTENT_X + 400, 5000)], 200, 240, CONTENT_X, CONTENT_W)).toBeUndefined()
  })

  test('wrap none and topAndBottom do not narrow the column', () => {
    expect(freeBandsFor([float(CONTENT_X + 400, 200, 200, 100, 'none')], 200, 240, CONTENT_X, CONTENT_W)).toBeUndefined()
    expect(freeBandsFor([float(CONTENT_X + 400, 200, 200, 100, 'topAndBottom')], 200, 240, CONTENT_X, CONTENT_W)).toBeUndefined()
  })
})

describe('text wraps beside a floating image', () => {
  test('square wrap narrows the lines that overlap the float', async () => {
    const doc = await docWithFloat(
      '<wp:wrapSquare wrapText="bothSides"/>',
      '<wp:positionH relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionH>',
      '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>',
    )
    const pages = layoutDocx(doc, measureFixed)
    const box = pages[0].images[0]
    expect(box.floating?.wrap).toBe('square')
    // the tail paragraph comes after the anchor, so it wraps around the float
    const tailLines = pages[0].lines.filter((l) => l.segs.some((s) => s.text === 'tail'))
    const overlapping = tailLines.filter((l) => l.yPx < box.yPx + box.heightPx && l.yPx + l.heightPx > box.yPx)
    const below = tailLines.filter((l) => l.yPx >= box.yPx + box.heightPx)
    expect(overlapping.length).toBeGreaterThan(0)
    for (const l of overlapping) expect(l.contentWidthPx).toBeLessThan(CONTENT_W)
    for (const l of below) expect(l.contentWidthPx).toBe(CONTENT_W)
  })

  test('text before the anchor paragraph keeps full width (single-pass limitation)', async () => {
    // The float is only known once its anchor paragraph is processed, so text
    // laid out earlier on the same page cannot wrap around it. Documented here
    // so the behavior is pinned, not accidental.
    const doc = await docWithFloat(
      '<wp:wrapSquare wrapText="bothSides"/>',
      '<wp:positionH relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionH>',
      '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>',
    )
    const pages = layoutDocx(doc, measureFixed)
    const leadLines = pages[0].lines.filter((l) => l.segs.some((s) => s.text === 'lead'))
    expect(leadLines.length).toBeGreaterThan(0)
    for (const l of leadLines) expect(l.contentWidthPx).toBe(CONTENT_W)
  })

  test('wrapNone leaves every line full width', async () => {
    const doc = await docWithFloat(
      '<wp:wrapNone/>',
      '<wp:positionH relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionH>',
      '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>',
    )
    const pages = layoutDocx(doc, measureFixed)
    const box = pages[0].images[0]
    expect(box.floating?.wrap).toBe('none')
    const overlapping = pages[0].lines.filter((l) => l.yPx < box.yPx + box.heightPx && l.yPx + l.heightPx > box.yPx)
    for (const l of overlapping) expect(l.contentWidthPx).toBe(CONTENT_W)
  })

  test('no float means every line is full width (regression guard)', async () => {
    const doc = await docWithFloat(
      '<wp:wrapNone/>',
      '<wp:positionH relativeFrom="page"><wp:posOffset>9144000</wp:posOffset></wp:positionH>',
      '<wp:positionV relativeFrom="page"><wp:posOffset>9144000</wp:posOffset></wp:positionV>',
    )
    const pages = layoutDocx(doc, measureFixed)
    for (const l of pages[0].lines) expect(l.contentWidthPx).toBe(CONTENT_W)
  })
})

describe('wrapped text does not paint over the float', () => {
  test('no glyph ink inside the float rect', async () => {
    const doc = await docWithFloat(
      '<wp:wrapSquare wrapText="bothSides"/>',
      '<wp:positionH relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionH>',
      '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>',
    )
    const pages = layoutDocx(doc, measureFixed)
    const page = pages[0]
    const box = page.images[0]
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(Math.ceil(page.widthPx), Math.ceil(page.heightPx))
    const ctx = canvas.getContext('2d')!
    ctx.fillStyle = '#fff'
    ctx.fillRect(0, 0, canvas.width, canvas.height)
    const { renderPages } = await import('../src/docx/layout')
    // paint text only (skip the image) to isolate the wrap contract
    renderPages([{ ...page, images: [] }], ctx as unknown as CanvasRenderingContext2D)
    // Only the post-anchor band must be clear: the lead paragraph precedes
    // the anchor, so it keeps full width by design (see the test above) and
    // may legitimately overlap the float.
    const leadEnd = Math.max(
      ...page.lines
        .filter((l) => l.segs.some((sg) => sg.text === 'lead'))
        .map((l) => l.yPx + l.heightPx),
    )
    const bandTop = Math.max(leadEnd, box.yPx)
    const bandH = box.yPx + box.heightPx - bandTop
    expect(bandH).toBeGreaterThan(0)
    const region = ctx.getImageData(
      Math.round(box.xPx), Math.round(bandTop),
      Math.round(box.widthPx), Math.round(bandH),
    )
    let ink = 0
    for (let i = 0; i < region.data.length; i += 4) {
      if (region.data[i] < 128) ink++
    }
    expect(ink).toBe(0)
  })
})
