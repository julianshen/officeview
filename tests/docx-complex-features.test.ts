import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { collectDocImages, layoutDocx, renderPages, type MeasureFn } from '../src/docx/layout'
import { decodeImage } from '../src/core/images'
import { CT_TYPES, ROOT_RELS } from '../src/testdata/ooxml-builders'
import { fontFamilyCss } from '../src/docx/styles'
import { paintDrawing } from '../src/docx/drawing'

const measure: MeasureFn = (text, style) => text.length * style.fontSizePt * 0.8
const ns =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
const picture = (id = 'img', w = 914400) =>
  `<w:drawing><wp:inline><wp:extent cx="${w}" cy="914400"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="${id}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`
const theme =
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1></a:clrScheme><a:fontScheme><a:minorFont><a:latin typeface="Cambria"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>'
async function fixture(body: string, parts: Record<string, string> = {}, sectionRefs = '') {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CT_TYPES)
  zip.file('_rels/.rels', ROOT_RELS)
  zip.file(
    'word/document.xml',
    `<w:document ${ns}><w:body>${body}<w:sectPr>${sectionRefs}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:left="1440" w:right="1440" w:bottom="1440"/></w:sectPr></w:body></w:document>`
  )
  zip.file('word/theme/theme1.xml', theme)
  const { createCanvas } = await import('canvas')
  const canvas = createCanvas(4, 4)
  const ctx = canvas.getContext('2d')
  ctx.fillStyle = '#ff0000'
  ctx.fillRect(0, 0, 4, 4)
  zip.file('word/media/img.png', canvas.toBuffer('image/png'))
  zip.file(
    'word/_rels/document.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="img" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/img.png"/><Relationship Id="theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="dm" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data.xml"/><Relationship Id="dg" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing.xml"/><Relationship Id="chart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart.xml"/><Relationship Id="ink" Type="http://schemas.microsoft.com/office/2007/relationships/ink" Target="ink/ink.xml"/></Relationships>'
  )
  for (const [name, xml] of Object.entries(parts)) zip.file(name, xml)
  return parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
}

describe('DOCX complex drawing and style regressions', () => {
  test('theme serif fonts retain a serif fallback when the font is unavailable', () => {
    expect(fontFamilyCss('Cambria')).toBe('"Cambria", serif')
    expect(fontFamilyCss('Calibri')).toBe('"Calibri", sans-serif')
  })
  test('paragraphs and tables remain interleaved in document order', async () => {
    const table =
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const doc = await fixture(`<w:p><w:r><w:t>before</w:t></w:r></w:p>${table}<w:p><w:r><w:t>after</w:t></w:r></w:p>`)
    expect(doc.sections[0].blocks.map((block) => block.kind)).toEqual(['p', 'table', 'p'])
    const lines = layoutDocx(doc, measure)[0].lines
    expect(lines.map((line) => line.segs.map((seg) => seg.text).join(''))).toEqual(['before', 'cell', 'after'])
  })
  test('field instructions and unused results do not enter mixed image flow', async () => {
    const doc = await fixture(
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:t>ignored</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r>${picture()}</w:r></w:p>`
    )
    const texts = layoutDocx(doc, measure)[0].lines.flatMap((line) => line.segs.map((seg) => seg.text))
    expect(texts).toEqual(['7'])
    doc.sections[0].header = [doc.sections[0].paragraphs[0]]
    doc.sections[0].blocks = []
    const { createCanvas } = await import('canvas')
    const ctx = createCanvas(816, 1056).getContext('2d')
    const calls: string[] = []
    ctx.fillText = (text) => {
      calls.push(text)
    }
    renderPages(layoutDocx(doc, measure), ctx as never, undefined, { pageNumberStart: 2 })
    expect(calls).toEqual(['2'])
  })
  test('table row height includes top and bottom cell padding', async () => {
    const doc = await fixture(
      '<w:tbl><w:tblPr><w:tblCellMar><w:top w:w="300"/><w:bottom w:w="150"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    )
    const page = layoutDocx(doc, measure)[0]
    const row = page.tables[0].rows[0]
    expect(row.heightPx).toBeCloseTo(page.lines[0].heightPx + 20 + 10)
  })
  test('vertical alignment uses the full merged cell height for text and pictures', async () => {
    const doc = await fixture(
      `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="1800" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>A</w:t>${picture()}</w:r></w:p></w:tc></w:tr><w:tr><w:trPr><w:trHeight w:val="1800" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`
    )
    const page = layoutDocx(doc, measure)[0]
    const cell = page.tables[0].rows[0].cells[0]
    expect(page.images[0].yPx).toBeCloseTo(page.tables[0].yPx + cell.yPx + (cell.heightPx - 96) / 2)
    expect(page.lines[0].yPx).toBeCloseTo(page.images[0].yPx)
  })
  test('empty paragraphs retain the font strut instead of treating points as pixels', async () => {
    const doc = await fixture('<w:p/><w:p><w:r><w:t>text</w:t></w:r></w:p>')
    const lines = layoutDocx(doc, measure)[0].lines
    expect(lines[0].heightPx / lines[1].heightPx).toBeGreaterThan(0.85)
  })
  test('inline pictures share a line and wrap as a group of line boxes', async () => {
    const doc = await fixture(`<w:p><w:r>${picture()}${picture()}${picture('img', 5486400)}${picture()}</w:r></w:p>`)
    const boxes = layoutDocx(doc, measure)[0].images
    expect(boxes.map((x) => [x.xPx, x.yPx])).toEqual([
      [96, 96],
      [192, 96],
      [96, 192],
      [96, 288]
    ])
  })
  test('text and pictures preserve their order within a run', async () => {
    const doc = await fixture(`<w:p><w:r><w:t>A</w:t>${picture()}<w:t>B</w:t></w:r></w:p>`)
    const page = layoutDocx(doc, measure)[0]
    const a = page.lines.flatMap((l) => l.segs).find((s) => s.text === 'A')!
    expect(a).toBeDefined()
    expect(page.images[0].xPx).toBeCloseTo(96 + a.widthPx)
    const b = page.lines.flatMap((l) => l.segs).find((s) => s.text === 'B')
    expect(b).toBeDefined()
    expect(page.images[0].yPx).toBe(96)
  })
  test('uses a compatibility fallback picture once, with the anchors placement', async () => {
    const doc = await fixture(
      '<w:p><w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:drawing><wp:anchor behindDoc="0"><wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/><wp:wrapNone/><a:graphic><a:graphicData uri="ink"/></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:imagedata r:id="img"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>'
    )
    expect(collectDocImages(doc)).toHaveLength(1)
    expect(layoutDocx(doc, measure)[0].images[0]).toMatchObject({ xPx: 192, yPx: 96, widthPx: 96, heightPx: 96 })
  })
  test('paints native compressed ink with its physical brush width', async () => {
    const doc = await fixture(
      '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"><w:contentPart r:id="ink"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
      {
        'word/ink/ink.xml': `<ink xmlns="http://www.w3.org/2003/InkML"><definitions><context xml:id="ctx"><inkSource><traceFormat><channel name="X" units="cm"/><channel name="Y" units="cm"/><channel name="F"/></traceFormat><channelProperties><channelProperty channel="X" name="resolution" value="1000"/><channelProperty channel="Y" name="resolution" value="1000"/></channelProperties></inkSource></context><brush xml:id="brush"><brushProperty name="width" value="0.035" units="cm"/></brush></definitions><trace contextRef="#ctx" brushRef="#brush">0 0 1,'1270'0'0,"-1270"1270"0,0 0 0</trace></ink>`
      }
    )
    const drawing = collectDocImages(doc)[0]?.drawing
    expect(drawing?.kind).toBe('ink')
    const { createCanvas } = await import('canvas')
    const ctx = createCanvas(816, 1056).getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, 816, 1056)
    renderPages(layoutDocx(doc, measure), ctx as never)
    expect(ctx.getImageData(120, 96, 1, 1).data[0]).toBeLessThan(128)
    expect(ctx.getImageData(120, 99, 1, 1).data[0]).toBe(255)
  })
  test('resolves theme fonts and conditional table styles without overriding explicit false', async () => {
    const doc = await fixture(
      '<w:tbl><w:tblPr><w:tblStyle w:val="accent"/><w:tblLook w:firstRow="1" w:firstColumn="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:shd w:fill="1F3F60"/></w:tcPr><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:rPr><w:b w:val="0"/><w:color w:val="FF0000"/></w:rPr><w:t>Explicit</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
      {
        'word/styles.xml':
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="table" w:styleId="base"><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="D6E3BC"/><w:insideV w:val="single" w:sz="4" w:color="D6E3BC"/></w:tblBorders></w:tblPr></w:style><w:style w:type="table" w:styleId="accent"><w:basedOn w:val="base"/><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tcPr><w:tcBorders><w:bottom w:val="single" w:sz="12" w:color="C2D69B"/></w:tcBorders></w:tcPr></w:tblStylePr></w:style></w:styles>'
      }
    )
    expect(doc.defaultFontFamily).toBe('Cambria')
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw Error('table')
    expect(block.table.borders?.top).toMatchObject({ color: 'D6E3BC' })
    expect(block.table.rows[0].cells[0].paragraphs[0].runs[0]).toMatchObject({ bold: true, color: 'FFFFFF' })
    expect(block.table.rows[0].cells[1].paragraphs[0].runs[0]).toMatchObject({ bold: false, color: 'FF0000' })
    const { createCanvas } = await import('canvas')
    const ctx = createCanvas(816, 1056).getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, 816, 1056)
    renderPages(layoutDocx(doc, measure), ctx as never)
    const pixel = ctx.getImageData(245, 96, 1, 1).data
    expect(pixel[1]).toBeGreaterThan(pixel[2])
  })
  test('renders a cached SmartArt drawing instead of losing its flow space and labels', async () => {
    const doc = await fixture(
      '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:effectExtent t="0" b="9525"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="dm"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:p><w:r><w:t>after</w:t></w:r></w:p>',
      {
        'word/diagrams/data.xml':
          '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="dg"/></dgm:ext></dgm:extLst></dgm:dataModel>',
        'word/diagrams/drawing.xml':
          '<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="ellipse"/><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></dsp:spPr><dsp:style><a:fontRef><a:schemeClr val="lt1"/></a:fontRef></dsp:style><dsp:txBody><a:bodyPr anchor="ctr"/><a:p><a:r><a:rPr sz="1400"/><a:t>Node</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>'
      }
    )
    expect(collectDocImages(doc)).toHaveLength(1)
    expect(layoutDocx(doc, measure)[0].images[0]).toMatchObject({ widthPx: 192, heightPx: 96 })
    expect(layoutDocx(doc, measure)[0].lines[0].yPx).toBe(193)
    const { createCanvas } = await import('canvas')
    const ctx = createCanvas(816, 1056).getContext('2d')
    renderPages(layoutDocx(doc, measure), ctx as never)
    expect([...ctx.getImageData(144, 120, 1, 1).data].slice(0, 3)).toEqual([79, 129, 189])
  })
  test('renders adjusted cached presets and ordered custom paths without fabricating unknown outlines', async () => {
    const drawing = `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree>
      <dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rightArrow"><a:avLst><a:gd name="adj1" fmla="val 60000"/><a:gd name="adj2" fmla="val 50000"/></a:avLst></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></dsp:spPr></dsp:sp>
      <dsp:sp><dsp:spPr><a:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:custGeom><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo><a:quadBezTo><a:pt x="100" y="100"/><a:pt x="50" y="100"/></a:quadBezTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></dsp:spPr></dsp:sp>
      <dsp:sp><dsp:spPr><a:xfrm><a:off x="1828800" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="unknownOfficeShape"/><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></dsp:spPr><dsp:txBody><a:p><a:r><a:t>Unknown</a:t></a:r></a:p></dsp:txBody></dsp:sp>
      <dsp:cxnSp><dsp:spPr><a:xfrm><a:off x="2743200" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></dsp:spPr></dsp:cxnSp>
    </dsp:spTree></dsp:drawing>`
    const doc = await fixture('<w:p><w:r><w:drawing><wp:inline><wp:extent cx="3657600" cy="914400"/><a:graphic><a:graphicData><dgm:relIds r:dm="dm"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>', {
      'word/diagrams/data.xml': '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="dg"/></dgm:ext></dgm:extLst></dgm:dataModel>',
      'word/diagrams/drawing.xml': drawing
    })
    const image = collectDocImages(doc)[0]
    expect(image.drawing?.kind).toBe('diagram')
    if (image.drawing?.kind !== 'diagram') throw Error('diagram missing')
    expect(image.drawing.shapes.map(s => s.geometry)).toEqual(['rightArrow', 'custom', 'unknownOfficeShape', 'rect'])
    expect(image.drawing.shapes[1].drawingGeometry?.paths[0].commands.map(command => command[0])).toEqual(['moveTo', 'lnTo', 'quadBezTo', 'lnTo', 'close'])
    expect(image.drawing.shapes[2].geometryIssues?.map(issue => issue.kind)).toContain('unknown-preset')
    const { createCanvas } = await import('canvas')
    const ctx = createCanvas(400, 100).getContext('2d')
    paintDrawing(image.drawing, ctx as never, 400, 100)
    const rgb = (x: number, y: number) => [...ctx.getImageData(x, y, 1, 1).data].slice(0, 3)
    expect(rgb(20, 22)).toEqual([255, 0, 0]) // adjusted shaft reaches y=20; default starts at y=25
    expect(rgb(110, 75)).toEqual([0, 0, 0]) // outside the custom triangle, no rectangle fallback
    expect(rgb(150, 25)).toEqual([0, 0, 255])
    expect(rgb(225, 50)).toEqual([0, 0, 0]) // unknown preset has no invented outline
    expect(rgb(325, 50)).toEqual([0, 255, 0]) // valid neighbor after unknown still paints
  })
  test('keeps valid cached paths when a sibling path is corrupt and resolves theme styles', async () => {
    const styledTheme = theme.replace('</a:themeElements>', '<a:fmtScheme><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme></a:themeElements>')
    const shapeXml = (x: number, directFill: string) => `<dsp:sp><dsp:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:custGeom><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="missingGuide" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo></a:path><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>${directFill}</dsp:spPr><dsp:style><a:fillRef idx="1"><a:schemeClr val="accent2"/></a:fillRef></dsp:style></dsp:sp>`
    const doc = await fixture('<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><a:graphic><a:graphicData><dgm:relIds r:dm="dm"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>', {
      'word/theme/theme1.xml': styledTheme,
      'word/diagrams/data.xml': '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="dg"/></dgm:ext></dgm:extLst></dgm:dataModel>',
      'word/diagrams/drawing.xml': `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree>${shapeXml(0, '')}${shapeXml(914400, '<a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>')}</dsp:spTree></dsp:drawing>`
    })
    const image = collectDocImages(doc)[0]
    if (image.drawing?.kind !== 'diagram') throw Error('diagram missing')
    expect(image.drawing.shapes[0].geometryIssues?.map(issue => issue.kind)).toContain('invalid-path')
    const { createCanvas } = await import('canvas')
    const ctx = createCanvas(200, 100).getContext('2d')
    paintDrawing(image.drawing, ctx as never, 200, 100)
    expect([...ctx.getImageData(50, 50, 1, 1).data].slice(0, 3)).toEqual([192, 80, 77])
    expect([...ctx.getImageData(150, 50, 1, 1).data].slice(0, 3)).toEqual([0, 0, 255])
  })
  test('body, header, and footer cached diagrams share the document theme palette and style matrix', async () => {
    const styledTheme = theme.replace('</a:themeElements>', '<a:fmtScheme><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme></a:themeElements>')
    const inline = '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><a:graphic><a:graphicData><dgm:relIds r:dm="dm"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    const cache = '<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></dsp:spPr></dsp:sp><dsp:sp><dsp:spPr><a:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></dsp:spPr><dsp:style><a:fillRef idx="1"><a:schemeClr val="accent2"/></a:fillRef></dsp:style></dsp:sp></dsp:spTree></dsp:drawing>'
    const data = '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="dg"/></dgm:ext></dgm:extLst></dgm:dataModel>'
    const rels = (extra = '') => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="dm" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data.xml"/><Relationship Id="dg" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing.xml"/>${extra}</Relationships>`
    const doc = await fixture(inline, {
      'word/theme/theme1.xml': styledTheme,
      'word/diagrams/data.xml': data,
      'word/diagrams/drawing.xml': cache,
      'word/header1.xml': `<w:hdr ${ns}>${inline}</w:hdr>`,
      'word/footer1.xml': `<w:ftr ${ns}>${inline}</w:ftr>`,
      'word/_rels/document.xml.rels': rels('<Relationship Id="theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="hdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="ftr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'),
      'word/_rels/header1.xml.rels': rels('<Relationship Id="badTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"/>'),
      'word/_rels/footer1.xml.rels': rels('<Relationship Id="missingTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/missing.xml"/>')
    }, '<w:headerReference w:type="default" r:id="hdr"/><w:footerReference w:type="default" r:id="ftr"/>')
    const parts = [doc.sections[0].paragraphs[0], doc.sections[0].header?.[0], doc.sections[0].footer?.[0]]
    const { createCanvas } = await import('canvas')
    for (const part of parts) {
      const drawing = part?.images[0]?.drawing
      expect(drawing?.kind).toBe('diagram')
      if (drawing?.kind !== 'diagram') continue
      const ctx = createCanvas(200, 100).getContext('2d')
      paintDrawing(drawing, ctx as never, 200, 100)
      expect([...ctx.getImageData(50, 50, 1, 1).data]).toEqual([79, 129, 189, 255])
      expect([...ctx.getImageData(150, 50, 1, 1).data]).toEqual([192, 80, 77, 255])
    }
  })
  test('empty, partial, and malformed local themes inherit document diagram styles', async () => {
    const styledTheme = theme.replace('</a:themeElements>', '<a:fmtScheme><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme></a:themeElements>')
    const inline = '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><a:graphic><a:graphicData><dgm:relIds r:dm="dm"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    const cache = '<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></dsp:spPr></dsp:sp><dsp:sp><dsp:spPr><a:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></dsp:spPr><dsp:style><a:fillRef idx="1"><a:schemeClr val="accent2"/></a:fillRef></dsp:style></dsp:sp></dsp:spTree></dsp:drawing>'
    const data = '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="dg"/></dgm:ext></dgm:extLst></dgm:dataModel>'
    const rels = (extra = '') => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="dm" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data.xml"/><Relationship Id="dg" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing.xml"/>${extra}</Relationships>`
    for (const [localTheme, directColor, referenceColor] of [
      ['<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>', [79, 129, 189], [192, 80, 77]],
      ['<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="00FF00"/></a:accent1></a:clrScheme></a:themeElements></a:theme>', [0, 255, 0], [192, 80, 77]],
      ['<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fmtScheme><a:fillStyleLst><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></a:fillStyleLst></a:fmtScheme></a:themeElements></a:theme>', [79, 129, 189], [255, 255, 0]],
      ['<a:theme>', [79, 129, 189], [192, 80, 77]]
    ] as const) {
      const doc = await fixture('', {
        'word/theme/theme1.xml': styledTheme,
        'word/theme/local.xml': localTheme,
        'word/diagrams/data.xml': data,
        'word/diagrams/drawing.xml': cache,
        'word/header1.xml': `<w:hdr ${ns}>${inline}</w:hdr>`,
        'word/_rels/document.xml.rels': rels('<Relationship Id="theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="hdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'),
        'word/_rels/header1.xml.rels': rels('<Relationship Id="localTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/local.xml"/>')
      }, '<w:headerReference w:type="default" r:id="hdr"/>')
      const drawing = doc.sections[0].header?.[0].images[0]?.drawing
      expect(drawing?.kind).toBe('diagram')
      if (drawing?.kind !== 'diagram') continue
      const { createCanvas } = await import('canvas')
      const ctx = createCanvas(200, 100).getContext('2d')
      paintDrawing(drawing, ctx as never, 200, 100)
      expect([...ctx.getImageData(50, 50, 1, 1).data]).toEqual([...directColor, 255])
      expect([...ctx.getImageData(150, 50, 1, 1).data]).toEqual([...referenceColor, 255])
    }
  })
  test('renders chart cached values with its categories and series', async () => {
    const doc = await fixture(
      '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="chart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
      {
        'word/charts/chart.xml':
          '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:overlap val="-27"/><c:ser><c:tx><c:v>Series</c:v></c:tx><c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr><c:cat><c:strLit><c:pt idx="0"><c:v>Category</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'
      }
    )
    expect(collectDocImages(doc)).toHaveLength(1)
    expect(collectDocImages(doc)[0]).toMatchObject({
      drawing: { kind: 'chart', series: [{ name: 'Series', values: [4] }], categories: ['Category'], overlap: -27 }
    })
    const { createCanvas } = await import('canvas')
    const ctx = createCanvas(816, 1056).getContext('2d')
    renderPages(layoutDocx(doc, measure), ctx as never)
    expect([...ctx.getImageData(380, 300, 1, 1).data].slice(0, 3)).toEqual([79, 129, 189])
  })
  test('keeps vertical text-box content from the supported Choice branch once', async () => {
    const doc = await fixture(
      '<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="1828800" cy="914400"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:wrapNone/><a:graphic><a:graphicData><wps:wsp><wps:spPr><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>你好</w:t></w:r></w:p><w:p><w:r><w:t>123</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr vert="eaVert"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>你好</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>'
    )
    expect(collectDocImages(doc)).toHaveLength(1)
    expect(collectDocImages(doc)[0]).toMatchObject({
      drawing: {
        kind: 'textbox',
        direction: 'eaVert',
        vertical: true,
        paragraphs: [{ runs: [{ text: '你好' }] }, { runs: [{ text: '123' }] }]
      }
    })
  })
  test('compatibility pictures remain paintable after decoding', async () => {
    const doc = await fixture(
      `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="unused"><unknown/></mc:Choice><mc:Fallback>${picture()}</mc:Fallback></mc:AlternateContent></w:r></w:p>`
    )
    const images = collectDocImages(doc)
    expect(images).toHaveLength(1)
    const decoded = await Promise.all(images.map((i) => decodeImage(i.data, i.mime)))
    const { createCanvas } = await import('canvas')
    const ctx = createCanvas(816, 1056).getContext('2d')
    renderPages(layoutDocx(doc, measure), ctx as never, decoded)
    expect([...ctx.getImageData(120, 120, 1, 1).data].slice(0, 3)).toEqual([255, 0, 0])
  })
})
