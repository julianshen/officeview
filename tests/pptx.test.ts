import { describe, test, expect } from 'vitest'
import JSZip from 'jszip'
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

const PPTX_NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
  + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const RELS_NS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/** buildPptx output plus extra parts (masters, themes) or a patched slide. */
async function pptxPatched(
  shapes: Parameters<typeof buildPptx>[0],
  patch: (zip: JSZip) => Promise<void> | void,
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await buildPptx(shapes))
  await patch(zip)
  return zip.generateAsync({ type: 'uint8array' })
}

async function withSlideBg(bgXml: string): Promise<Uint8Array> {
  return pptxPatched([], async (zip) => {
    const path = 'ppt/slides/slide1.xml'
    const xml = await zip.file(path)!.async('string')
    zip.file(path, xml.replace('<p:cSld>', `<p:cSld>${bgXml}`))
  })
}

describe('pptx slide background', () => {
  test('a direct bgPr solid fill resolves to the model', async () => {
    const buf = await withSlideBg('<p:bg><p:bgPr><a:solidFill><a:srgbClr val="0E0E1A"/></a:solidFill></p:bgPr></p:bg>')
    const doc = await parsePptx(await OfficePackage.load(buf))
    expect(doc.slides[0].background).toBe('#0E0E1A')
  })

  test('no background anywhere means white (undefined)', async () => {
    const doc = await parsePptx(await OfficePackage.load(await buildPptx([])))
    expect(doc.slides[0].background).toBeUndefined()
  })

  test('a gradient-only bgPr falls back gracefully instead of throwing', async () => {
    const buf = await withSlideBg(
      '<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs></a:gsLst></a:gradFill></p:bgPr></p:bg>',
    )
    const doc = await parsePptx(await OfficePackage.load(buf))
    expect(doc.slides[0].background).toBeUndefined()
  })

  test('a slide without bg inherits the master background', async () => {
    const buf = await pptxPatched([], async (zip) => {
      const relsPath = 'ppt/slides/_rels/slide1.xml.rels'
      const rels = await zip.file(relsPath)!.async('string')
      zip.file(relsPath, rels.replace('</Relationships>',
        `<Relationship Id="rIdLayout" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`))
      zip.file('ppt/slideLayouts/slideLayout1.xml',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout ${PPTX_NS}><p:cSld name="Title Slide"/></p:sldLayout>`)
      zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships ${RELS_NS}>`
        + `<Relationship Id="rIdMaster" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`)
      zip.file('ppt/slideMasters/slideMaster1.xml',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster ${PPTX_NS}><p:cSld>`
        + `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></p:bgPr></p:bg>`
        + `</p:cSld></p:sldMaster>`)
    })
    const doc = await parsePptx(await OfficePackage.load(buf))
    expect(doc.slides[0].background).toBe('#112233')
  })

  test('a bgRef resolves through the theme fill list', async () => {
    const buf = await pptxPatched([], async (zip) => {
      const slidePath = 'ppt/slides/slide1.xml'
      const slide = await zip.file(slidePath)!.async('string')
      zip.file(slidePath, slide.replace('<p:cSld>',
        '<p:cSld><p:bg><p:bgRef idx="1001"><a:scrgbClr r="0" g="0" b="0"/></p:bgRef></p:bg>'))
      const relsPath = 'ppt/slides/_rels/slide1.xml.rels'
      const rels = await zip.file(relsPath)!.async('string')
      zip.file(relsPath, rels.replace('</Relationships>',
        `<Relationship Id="rIdTheme" Type="${REL}/theme" Target="../theme/theme1.xml"/></Relationships>`))
      zip.file('ppt/theme/theme1.xml',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme ${PPTX_NS} name="Theme">`
        + `<a:themeElements><a:fmtScheme name="Office"><a:bgFillStyleLst>`
        + `<a:solidFill><a:srgbClr val="445566"/></a:solidFill>`
        + `</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`)
    })
    const doc = await parsePptx(await OfficePackage.load(buf))
    expect(doc.slides[0].background).toBe('#445566')
  })

  test('a dark background paints under light text', async () => {
    const buf = await pptxPatched(
      [{
        prst: 'rect',
        off: ['914400', '914400'],
        ext: ['3657600', '1828800'],
        paragraphs: [{ runs: [{ text: 'Light title', color: 'FFFFFF' }] }],
      }],
      async (zip) => {
        const path = 'ppt/slides/slide1.xml'
        const xml = await zip.file(path)!.async('string')
        zip.file(path, xml.replace('<p:cSld>',
          '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="0E0E1A"/></a:solidFill></p:bgPr></p:bg>'))
      },
    )
    const doc = await parsePptx(await OfficePackage.load(buf))
    const metrics = slideMetrics(doc)
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(metrics.widthPx, metrics.heightPx)
    const ctx = canvas.getContext('2d')!
    renderSlide(doc.slides[0], ctx as never, metrics)
    // base layer is the background, not white
    expect([...ctx.getImageData(5, 5, 1, 1).data].slice(0, 3)).toEqual([14, 14, 26])
    // light title text is legible ink, not blank canvas
    const region = ctx.getImageData(90, 90, 380, 190)
    let ink = 0
    for (let i = 0; i < region.data.length; i += 4) {
      if (region.data[i] > 200 && region.data[i + 1] > 200 && region.data[i + 2] > 200) ink++
    }
    expect(ink).toBeGreaterThan(50)
  })
})
