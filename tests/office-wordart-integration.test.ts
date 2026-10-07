import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import { renderSlide } from '../src/pptx/render'

const emu = (px: number) => px * 9525
const rels = (items: string) => `<Relationships>${items}</Relationships>`
const rel = (id: string, kind: string, target: string) =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}" Target="${target}"/>`

const THEME_XML = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="111111"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="FF2200"/></a:accent1>
      <a:accent2><a:srgbClr val="0033CC"/></a:accent2>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`

async function createPptxFixture(spBody: string) {
  const zip = new JSZip()
  zip.file(
    'ppt/presentation.xml',
    '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="2857500" cy="1905000"/><p:sldIdLst><p:sldId r:id="s1"/></p:sldIdLst></p:presentation>',
  )
  zip.file('ppt/_rels/presentation.xml.rels', rels(rel('s1', 'slide', 'slides/slide1.xml')))
  zip.file(
    'ppt/slides/slide1.xml',
    `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>${spBody}</p:spTree></p:cSld></p:sld>`,
  )
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    rels(rel('layout', 'slideLayout', '../slideLayouts/slideLayout1.xml')),
  )
  zip.file(
    'ppt/slideLayouts/slideLayout1.xml',
    '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:clrMapOvr><a:masterClrMapping xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></p:clrMapOvr></p:sldLayout>',
  )
  zip.file(
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    rels(rel('master', 'slideMaster', '../slideMasters/slideMaster1.xml')),
  )
  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    '<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:clrMap xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"/></p:sldMaster>',
  )
  zip.file(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    rels(rel('theme', 'theme', '../theme/theme1.xml')),
  )
  zip.file('ppt/theme/theme1.xml', THEME_XML)
  const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
  return parsePptx(pkg)
}

describe('Phase 3: Format Adapters Integration', () => {
  test('PPTX shapes parse and render WordArt text runs inheriting theme colors', async () => {
    const sp = `<p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="WordArt Shape"/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="${emu(10)}" y="${emu(10)}"/><a:ext cx="${emu(200)}" cy="${emu(80)}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr sz="2400">
              <a:gradFill>
                <a:gsLst>
                  <a:gs pos="0"><a:schemeClr val="accent1"/></a:gs>
                  <a:gs pos="100000"><a:schemeClr val="accent2"/></a:gs>
                </a:gsLst>
                <a:lin ang="5400000"/>
              </a:gradFill>
              <a:ln w="25400">
                <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
              </a:ln>
              <a:effectLst>
                <a:outerShdw blurRad="38100" dist="25400" dir="5400000">
                  <a:schemeClr val="dk1"/>
                </a:outerShdw>
              </a:effectLst>
            </a:rPr>
            <a:t>Theme WordArt</a:t>
          </a:r>
        </a:p>
      </p:txBody>
    </p:sp>`

    const doc = await createPptxFixture(sp)
    const shape = doc.slides[0].shapes[0]
    expect(shape).toBeDefined()
    expect(shape.textBody).toBeDefined()

    const run = shape.textBody!.paragraphs[0].runs[0]
    expect(run.text).toBe('Theme WordArt')
    // Theme scheme colors must be resolved:
    expect(run.textFill).toEqual({
      kind: 'gradient',
      angle: Math.PI / 2,
      stops: [
        { position: 0, color: '#FF2200' },
        { position: 1, color: '#0033CC' },
      ],
    })
    expect(run.textOutline).toBeDefined()
    expect(run.textOutline?.color).toBe('#FF2200')
    expect(run.textOutline?.widthPx).toBeCloseTo(2.67, 1)

    expect(run.textShadow).toBeDefined()
    expect(run.textShadow?.color).toBe('#111111')
    expect(run.textShadow?.blurPx).toBeCloseTo(4, 1)

    // Render onto canvas and verify stroke and fill executions
    const canvas = createCanvas(300, 200)
    const ctx = canvas.getContext('2d')
    const gradientStops: Array<{ offset: number; color: string }> = []
    const strokes: Array<{ text: string; lineWidth: number; strokeStyle: string }> = []
    const fills: Array<{ text: string; shadowBlur: number; shadowColor: string }> = []

    const origCreateLinearGradient = ctx.createLinearGradient.bind(ctx)
    ctx.createLinearGradient = (x0, y0, x1, y1) => {
      const grad = origCreateLinearGradient(x0, y0, x1, y1)
      const origAddColorStop = grad.addColorStop.bind(grad)
      grad.addColorStop = (offset, color) => {
        gradientStops.push({ offset, color })
        return origAddColorStop(offset, color)
      }
      return grad
    }

    const origStrokeText = ctx.strokeText.bind(ctx)
    ctx.strokeText = (text, x, y) => {
      strokes.push({ text: String(text), lineWidth: ctx.lineWidth, strokeStyle: String(ctx.strokeStyle) })
      return origStrokeText(text, x, y)
    }

    const origFillText = ctx.fillText.bind(ctx)
    ctx.fillText = (text, x, y) => {
      fills.push({ text: String(text), shadowBlur: ctx.shadowBlur, shadowColor: String(ctx.shadowColor) })
      return origFillText(text, x, y)
    }

    renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)

    // Gradient was constructed with resolved theme colors for both segments:
    expect(gradientStops).toEqual([
      { offset: 0, color: '#FF2200' },
      { offset: 1, color: '#0033CC' },
      { offset: 0, color: '#FF2200' },
      { offset: 1, color: '#0033CC' },
    ])
    // Outline stroke was called with resolved theme outline:
    expect(strokes.map(s => s.text)).toEqual(['Theme ', 'WordArt'])
    expect(strokes[0].strokeStyle).toBe('#ff2200')
    // Fill text was called with shadow parameters:
    expect(fills.map(f => f.text)).toEqual(['Theme ', 'WordArt'])
    expect(fills[0].shadowColor).toBe('#111111')
    expect(fills[0].shadowBlur).toBeCloseTo(4, 1)
  })

  test('DOCX DrawingML shapes (<wps:wsp>) parse and render WordArt gradient and outline text', async () => {
    const zip = new JSZip()
    const docxXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <w:body>
        <w:p>
          <w:r>
            <w:drawing>
              <wp:inline>
                <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
                <wp:docPr id="1" name="WordArt Shape 1"/>
                <a:graphic>
                  <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                    <wps:wsp>
                      <wps:spPr>
                        <a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm>
                        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                      </wps:spPr>
                      <wps:txBody>
                        <a:bodyPr/>
                        <a:p>
                          <a:r>
                            <a:rPr sz="2400">
                              <a:gradFill>
                                <a:gsLst>
                                  <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
                                  <a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs>
                                </a:gsLst>
                                <a:lin ang="5400000"/>
                              </a:gradFill>
                              <a:ln w="25400">
                                <a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>
                              </a:ln>
                            </a:rPr>
                            <a:t>Docx WordArt</a:t>
                          </a:r>
                        </a:p>
                      </wps:txBody>
                    </wps:wsp>
                  </a:graphicData>
                </a:graphic>
              </wp:inline>
            </w:drawing>
          </w:r>
        </w:p>
      </w:body>
    </w:document>`

    zip.file('word/document.xml', docxXml)
    zip.file('word/_rels/document.xml.rels', '<Relationships/>')
    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const doc = await parseDocx(pkg)

    const para = doc.sections[0].paragraphs[0]
    expect(para.images.length).toBe(1)
    const img = para.images[0]
    expect(img.drawing).toBeDefined()
    expect(img.drawing?.kind).toBe('diagram')
    if (img.drawing?.kind !== 'diagram') throw new Error('Expected diagram drawing')

    const shape = img.drawing.shapes[0]
    expect(shape).toBeDefined()
    expect(shape.textBody).toBeDefined()
    const run = shape.textBody!.paragraphs[0].runs[0]
    expect(run.text).toBe('Docx WordArt')
    expect(run.textFill).toEqual({
      kind: 'gradient',
      angle: Math.PI / 2,
      stops: [
        { position: 0, color: '#FF0000' },
        { position: 1, color: '#00FF00' },
      ],
    })
    expect(run.textOutline).toBeDefined()
    expect(run.textOutline?.color).toBe('#0000FF')

    // Render canvas
    const canvas = createCanvas(200, 100)
    const ctx = canvas.getContext('2d')
    const strokes: string[] = []
    const gradStops: Array<{ offset: number; color: string }> = []

    const origCreateGrad = ctx.createLinearGradient.bind(ctx)
    ctx.createLinearGradient = (x0, y0, x1, y1) => {
      const grad = origCreateGrad(x0, y0, x1, y1)
      const origAdd = grad.addColorStop.bind(grad)
      grad.addColorStop = (o, c) => {
        gradStops.push({ offset: o, color: c })
        return origAdd(o, c)
      }
      return grad
    }
    const origStroke = ctx.strokeText.bind(ctx)
    ctx.strokeText = (t, x, y) => {
      strokes.push(String(t))
      return origStroke(t, x, y)
    }

    paintDrawing(img.drawing, ctx as unknown as CanvasRenderingContext2D, 200, 100)
    expect(gradStops.length).toBeGreaterThanOrEqual(2)
    expect(gradStops[0].color).toBe('#FF0000')
    expect(gradStops[1].color).toBe('#00FF00')
    expect(strokes.length).toBeGreaterThan(0)
  })

  test('XLSX DrawingML shapes (<xdr:sp>) parse and render WordArt styled text runs', async () => {
    const { buildXlsx } = await import('../src/testdata/ooxml-builders')
    const { parseXlsx } = await import('../src/xlsx/parse')
    const { renderSheet } = await import('../src/xlsx/render')

    const xdrSp = `<xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="10" name="WordArt Shape 10"/></xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(180)}" cy="${emu(80)}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr sz="2200">
              <a:gradFill>
                <a:gsLst>
                  <a:gs pos="0"><a:srgbClr val="00AAFF"/></a:gs>
                  <a:gs pos="100000"><a:srgbClr val="FFAA00"/></a:gs>
                </a:gsLst>
                <a:lin ang="10800000"/>
              </a:gradFill>
              <a:ln w="12700">
                <a:solidFill><a:srgbClr val="112233"/></a:solidFill>
              </a:ln>
              <a:effectLst>
                <a:outerShdw blurRad="19050" dist="19050" dir="5400000">
                  <a:srgbClr val="445566"/>
                </a:outerShdw>
              </a:effectLst>
            </a:rPr>
            <a:t>Xlsx WordArt</a:t>
          </a:r>
        </a:p>
      </xdr:txBody>
    </xdr:sp>`

    const drawingXml = `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="${emu(180)}" cy="${emu(80)}"/>${xdrSp}<xdr:clientData/></xdr:absoluteAnchor>`

    const base = await buildXlsx([{ name: 'Sheet1', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }] }])
    const zip = await JSZip.loadAsync(base)
    let sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    sheetXml = sheetXml.replace('</worksheet>', '<drawing r:id="rDraw"/></worksheet>')
    zip.file('xl/worksheets/sheet1.xml', sheetXml)
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<Relationships><Relationship Id="rDraw" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>',
    )
    zip.file(
      'xl/drawings/drawing1.xml',
      `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${drawingXml}</xdr:wsDr>`,
    )

    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const doc = await parseXlsx(pkg)
    const sheet = doc.sheets[0]
    expect(sheet.drawings).toBeDefined()
    expect(sheet.drawings!.length).toBe(1)

    const shape = sheet.drawings![0]
    expect(shape.textBody).toBeDefined()
    const run = shape.textBody!.paragraphs[0].runs[0]
    expect(run.text).toBe('Xlsx WordArt')
    expect(run.textFill).toEqual({
      kind: 'gradient',
      angle: Math.PI,
      stops: [
        { position: 0, color: '#00AAFF' },
        { position: 1, color: '#FFAA00' },
      ],
    })
    expect(run.textOutline).toBeDefined()
    expect(run.textOutline?.color).toBe('#112233')
    expect(run.textOutline?.widthPx).toBeCloseTo(1.33, 1)
    expect(run.textShadow).toBeDefined()
    expect(run.textShadow?.color).toBe('#445566')
    expect(run.textShadow?.blurPx).toBeCloseTo(2, 1)

    // Render canvas
    const canvas = createCanvas(300, 200)
    const ctx = canvas.getContext('2d')
    const strokes: string[] = []
    const gradStops: Array<{ offset: number; color: string }> = []
    const shadows: string[] = []

    const origCreateGrad = ctx.createLinearGradient.bind(ctx)
    ctx.createLinearGradient = (x0, y0, x1, y1) => {
      const grad = origCreateGrad(x0, y0, x1, y1)
      const origAdd = grad.addColorStop.bind(grad)
      grad.addColorStop = (o, c) => {
        gradStops.push({ offset: o, color: c })
        return origAdd(o, c)
      }
      return grad
    }
    const origStroke = ctx.strokeText.bind(ctx)
    ctx.strokeText = (t, x, y) => {
      strokes.push(String(t))
      return origStroke(t, x, y)
    }
    const origFill = ctx.fillText.bind(ctx)
    ctx.fillText = (t, x, y) => {
      shadows.push(String(ctx.shadowColor))
      return origFill(t, x, y)
    }

    renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D)
    expect(gradStops.length).toBeGreaterThanOrEqual(2)
    expect(gradStops[0].color).toBe('#00AAFF')
    expect(gradStops[1].color).toBe('#FFAA00')
    expect(strokes.length).toBeGreaterThan(0)
    expect(shadows).toContain('#445566')
  })

  test('WordArt diagnostics flow into single body diagnostic channel across all three formats', async () => {
    const unsupportedPattXml = `<a:rPr>
      <a:pattFill prst="unsupportedPreset">
        <a:fgClr><a:srgbClr val="FF0000"/></a:fgClr>
        <a:bgClr><a:srgbClr val="0000FF"/></a:bgClr>
      </a:pattFill>
    </a:rPr><a:t>Bad Pattern</a:t>`

    // 1. PPTX
    const pptxSp = `<p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Bad Patt"/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(100)}" cy="${emu(50)}"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:p><a:r>${unsupportedPattXml}</a:r></a:p></p:txBody>
    </p:sp>`
    const pptxDoc = await createPptxFixture(pptxSp)
    const pptxBody = pptxDoc.slides[0].shapes[0].textBody
    const expectedDiag = {
      kind: 'unsupported-text-appearance',
      feature: 'pattFill:unsupportedPreset',
      message: 'WordArt pattern preset is deferred; using flat color',
    }
    expect(pptxBody?.diagnostics).toContainEqual(expectedDiag)

    // 2. DOCX
    const docxXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <w:body><w:p><w:r><w:drawing><wp:inline>
        <wp:extent cx="${emu(100)}" cy="${emu(50)}"/>
        <wp:docPr id="1" name="Docx Bad Patt"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(100)}" cy="${emu(50)}"/></a:xfrm></wps:spPr>
            <wps:txBody><a:bodyPr/><a:p><a:r>${unsupportedPattXml}</a:r></a:p></wps:txBody>
          </wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p></w:body>
    </w:document>`
    const zipDocx = new JSZip()
    zipDocx.file('word/document.xml', docxXml)
    zipDocx.file('word/_rels/document.xml.rels', '<Relationships/>')
    const { parseDocx } = await import('../src/docx/parse')
    const docxDoc = await parseDocx(await OfficePackage.load(await zipDocx.generateAsync({ type: 'uint8array' })))
    const docxDrawing = docxDoc.sections[0].paragraphs[0].images[0].drawing
    expect(docxDrawing?.kind).toBe('diagram')
    if (docxDrawing?.kind !== 'diagram') throw new Error('Expected diagram')
    const docxBody = docxDrawing.shapes[0].textBody
    expect(docxBody?.diagnostics).toContainEqual(expectedDiag)

    // 3. XLSX
    const { buildXlsx } = await import('../src/testdata/ooxml-builders')
    const { parseXlsx } = await import('../src/xlsx/parse')
    const xlsxSp = `<xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="2" name="Xlsx Bad Patt"/></xdr:nvSpPr>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(100)}" cy="${emu(50)}"/></a:xfrm></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r>${unsupportedPattXml}</a:r></a:p></xdr:txBody>
    </xdr:sp>`
    const drawingXml = `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="${emu(100)}" cy="${emu(50)}"/>${xlsxSp}<xdr:clientData/></xdr:absoluteAnchor>`
    const baseXlsx = await buildXlsx([{ name: 'Sheet1', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }] }])
    const zipXlsx = await JSZip.loadAsync(baseXlsx)
    let sheetXml = await zipXlsx.file('xl/worksheets/sheet1.xml')!.async('string')
    sheetXml = sheetXml.replace('</worksheet>', '<drawing r:id="rDraw"/></worksheet>')
    zipXlsx.file('xl/worksheets/sheet1.xml', sheetXml)
    zipXlsx.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<Relationships><Relationship Id="rDraw" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>',
    )
    zipXlsx.file(
      'xl/drawings/drawing1.xml',
      `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${drawingXml}</xdr:wsDr>`,
    )
    const xlsxDoc = await parseXlsx(await OfficePackage.load(await zipXlsx.generateAsync({ type: 'uint8array' })))
    const xlsxBody = xlsxDoc.sheets[0].drawings![0].textBody
    expect(xlsxBody?.diagnostics).toContainEqual(expectedDiag)
  })

  test('End-to-end multi-format fixture test parsing and painting WordArt shapes', async () => {
    // 1. PPTX E2E: WordArt shape + unstyled shape on same slide
    const pptxTree = `
      <p:sp>
        <p:nvSpPr><p:cNvPr id="1" name="WordArt 1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(150)}" cy="${emu(50)}"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:r>
          <a:rPr sz="2000">
            <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst></a:gradFill>
            <a:ln w="19050"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln>
            <a:effectLst><a:outerShdw blurRad="25400" dist="19050"><a:srgbClr val="111111"/></a:outerShdw></a:effectLst>
          </a:rPr>
          <a:t>PPTX WordArt</a:t>
        </a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Plain Text"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="${emu(60)}"/><a:ext cx="${emu(150)}" cy="${emu(50)}"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1600"/><a:t>Plain Body</a:t></a:r></a:p></p:txBody>
      </p:sp>
    `
    const pptxDoc = await createPptxFixture(pptxTree)
    const pptxCanvas = createCanvas(300, 200)
    const pptxCtx = pptxCanvas.getContext('2d')
    const pptxFillShadows: number[] = []
    const origPptxFill = pptxCtx.fillText.bind(pptxCtx)
    pptxCtx.fillText = (t, x, y) => {
      pptxFillShadows.push(pptxCtx.shadowBlur)
      return origPptxFill(t, x, y)
    }
    renderSlide(pptxDoc.slides[0], pptxCtx as unknown as CanvasRenderingContext2D)
    // First run (WordArt) had shadow blur > 0, second run (Plain) had shadow blur == 0
    expect(pptxFillShadows.length).toBeGreaterThanOrEqual(2)
    expect(pptxFillShadows[0]).toBeGreaterThan(0)
    expect(pptxFillShadows[pptxFillShadows.length - 1]).toBe(0)

    // 2. DOCX E2E: WordArt drawing
    const docxXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <w:body><w:p><w:r><w:drawing><wp:inline>
        <wp:extent cx="${emu(160)}" cy="${emu(80)}"/>
        <wp:docPr id="1" name="Docx WordArt E2E"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(160)}" cy="${emu(80)}"/></a:xfrm></wps:spPr>
            <wps:txBody><a:bodyPr/><a:p><a:r>
              <a:rPr sz="2200">
                <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF00FF"/></a:gs><a:gs pos="100000"><a:srgbClr val="00FFFF"/></a:gs></a:gsLst></a:gradFill>
                <a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
              </a:rPr>
              <a:t>Docx E2E</a:t>
            </a:r></a:p></wps:txBody>
          </wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p></w:body>
    </w:document>`
    const zipDocx = new JSZip()
    zipDocx.file('word/document.xml', docxXml)
    zipDocx.file('word/_rels/document.xml.rels', '<Relationships/>')
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const docxDoc = await parseDocx(await OfficePackage.load(await zipDocx.generateAsync({ type: 'uint8array' })))
    const docxDrawing = docxDoc.sections[0].paragraphs[0].images[0].drawing!
    const docxCanvas = createCanvas(160, 80)
    const docxCtx = docxCanvas.getContext('2d')
    let docxStroked = false
    const origDocxStroke = docxCtx.strokeText.bind(docxCtx)
    docxCtx.strokeText = (t, x, y) => {
      docxStroked = true
      return origDocxStroke(t, x, y)
    }
    paintDrawing(docxDrawing, docxCtx as unknown as CanvasRenderingContext2D, 160, 80)
    expect(docxStroked).toBe(true)

    // 3. XLSX E2E: WordArt drawing
    const { buildXlsx } = await import('../src/testdata/ooxml-builders')
    const { parseXlsx } = await import('../src/xlsx/parse')
    const { renderSheet } = await import('../src/xlsx/render')
    const xlsxSp = `<xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="1" name="Xlsx WordArt E2E"/></xdr:nvSpPr>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(160)}" cy="${emu(80)}"/></a:xfrm></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r>
        <a:rPr sz="2200">
          <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF5500"/></a:gs><a:gs pos="100000"><a:srgbClr val="0055FF"/></a:gs></a:gsLst></a:gradFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
        </a:rPr>
        <a:t>Xlsx E2E</a:t>
      </a:r></a:p></xdr:txBody>
    </xdr:sp>`
    const xlsxDrXml = `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="${emu(160)}" cy="${emu(80)}"/>${xlsxSp}<xdr:clientData/></xdr:absoluteAnchor>`
    const baseXlsx = await buildXlsx([{ name: 'Sheet1', rows: [{ r: 1, cells: [{ ref: 'A1', v: 42 }] }] }])
    const zipXlsx = await JSZip.loadAsync(baseXlsx)
    let sheetXml = await zipXlsx.file('xl/worksheets/sheet1.xml')!.async('string')
    sheetXml = sheetXml.replace('</worksheet>', '<drawing r:id="rDraw"/></worksheet>')
    zipXlsx.file('xl/worksheets/sheet1.xml', sheetXml)
    zipXlsx.file('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships><Relationship Id="rDraw" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>')
    zipXlsx.file('xl/drawings/drawing1.xml', `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${xlsxDrXml}</xdr:wsDr>`)
    const xlsxDoc = await parseXlsx(await OfficePackage.load(await zipXlsx.generateAsync({ type: 'uint8array' })))
    const xlsxCanvas = createCanvas(300, 200)
    const xlsxCtx = xlsxCanvas.getContext('2d')
    let xlsxStroked = false
    const origXlsxStroke = xlsxCtx.strokeText.bind(xlsxCtx)
    xlsxCtx.strokeText = (t, x, y) => {
      xlsxStroked = true
      return origXlsxStroke(t, x, y)
    }
    renderSheet(xlsxDoc.sheets[0], xlsxCtx as unknown as CanvasRenderingContext2D)
    expect(xlsxStroked).toBe(true)
  })
})
