import { describe, expect, test } from 'vitest'
import { parseXmlOrdered } from '../src/core/xml'
import { parseTextBody } from '../src/drawing/text-parse'

const txBody = (rpr: string, text: string) =>
  parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:p><a:r>${rpr}<a:t>${text}</a:t></a:r></a:p></a:txBody>`)

describe('WordArt Extended Effects & Legacy VML Fallback - Phase 6', () => {
  test('emits diagnostic and falls back gracefully for text <a:glow> and <a:reflection>', () => {
    const glowRpr = `<a:rPr><a:effectLst><a:glow rad="63500"><a:srgbClr val="FFFF00"/></a:glow></a:effectLst></a:rPr>`
    const glowBody = parseTextBody(txBody(glowRpr, 'Glow Text'))
    expect(glowBody.paragraphs[0].runs[0].text).toBe('Glow Text')
    expect(glowBody.diagnostics).toContainEqual(
      expect.objectContaining({
        kind: 'unsupported-text-appearance',
        feature: 'glow',
      })
    )

    const reflRpr = `<a:rPr><a:effectLst><a:reflection blurRad="12700" stA="50000"/></a:effectLst></a:rPr>`
    const reflBody = parseTextBody(txBody(reflRpr, 'Reflection Text'))
    expect(reflBody.paragraphs[0].runs[0].text).toBe('Reflection Text')
    expect(reflBody.diagnostics).toContainEqual(
      expect.objectContaining({
        kind: 'unsupported-text-appearance',
        feature: 'reflection',
      })
    )
  })

  test('unified VML parser parses <v:shape><v:textpath> with string, font-family, font-size, alignment, fill, and stroke into DrawingTextBody', async () => {
    const { parseVmlWordArt } = await import('../src/drawing/vml')
    const vmlXml = parseXmlOrdered(`
      <xml xmlns:v="urn:schemas-microsoft-com:vml">
        <v:shape id="WordArt_1"
          style="position:absolute;width:300pt;height:60pt;"
          fillcolor="#FF0000" strokecolor="#0000FF" strokeweight="1.5pt">
          <v:fill color="#FF0000"/>
          <v:stroke color="#0000FF" weight="1.5pt"/>
          <v:textpath on="True"
            style="font-family:'Arial Black';font-size:36pt;font-weight:bold;font-style:italic;v-text-align:center;"
            string="Legacy WordArt"/>
        </v:shape>
      </xml>
    `)

    const result = parseVmlWordArt(vmlXml)
    expect(result).toBeDefined()
    expect(result?.shapeId).toBe('WordArt_1')
    expect(result?.widthPt).toBe(300)
    expect(result?.heightPt).toBe(60)

    const body = result!.textBody
    expect(body.paragraphs.length).toBe(1)
    expect(body.paragraphs[0].align).toBe('center')

    const run = body.paragraphs[0].runs[0]
    expect(run.text).toBe('Legacy WordArt')
    expect(run.fontFamily).toBe('Arial Black')
    expect(run.fontSizePt).toBe(36)
    expect(run.bold).toBe(true)
    expect(run.italic).toBe(true)
    expect(run.color).toBe('#FF0000')
    expect(run.textOutline).toBeDefined()
    expect(run.textOutline?.color).toBe('#0000FF')
    expect(run.textOutline?.widthPx).toBeCloseTo(1.5 * (96 / 72), 3)
  })

  test('DOCX drawing routes legacy VML WordArt through unified VML parser into canvas rendering', async () => {
    const JSZip = (await import('jszip')).default
    const { OfficePackage } = await import('../src/core/zip')
    const { parseDocx } = await import('../src/docx/parse')

    const zip = new JSZip()
    zip.file('[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`)
    zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
    zip.file('word/document.xml', `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    <w:p>
      <w:r>
        <w:pict>
          <v:shape id="WordArt_Docx" style="width:200pt;height:40pt;" fillcolor="#FF0000" strokecolor="#0000FF">
            <v:fill color="#FF0000"/>
            <v:stroke color="#0000FF" weight="1pt"/>
            <v:textpath on="True" style="font-family:'Arial';font-size:32pt;" string="DOCX WordArt"/>
          </v:shape>
        </w:pict>
      </w:r>
    </w:p>
  </w:body>
</w:document>`)

    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const doc = await parseDocx(pkg)
    expect(doc.sections.length).toBeGreaterThan(0)
    const paragraphs = doc.sections[0].paragraphs
    expect(paragraphs.length).toBeGreaterThan(0)
    const images = paragraphs[0].images
    expect(images.length).toBe(1)
    const img = images[0]
    expect(img.drawing).toBeDefined()
    expect(img.drawing?.kind).toBe('diagram')
    const shape = (img.drawing as any).shapes[0]
    expect(shape.textBody).toBeDefined()
    expect(shape.textBody.paragraphs[0].runs[0].text).toBe('DOCX WordArt')
  })

  test('XLSX drawing routes legacy VML WordArt through unified VML parser into canvas rendering', async () => {
    const JSZip = (await import('jszip')).default
    const { OfficePackage } = await import('../src/core/zip')
    const { parseXlsx } = await import('../src/xlsx/parse')

    const zip = new JSZip()
    zip.file('[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>`)
    zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`)
    zip.file('xl/workbook.xml', `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`)
    zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`)
    zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><v>100</v></c></row></sheetData><legacyDrawing r:id="rIdVml"/></worksheet>`)
    zip.file('xl/worksheets/_rels/sheet1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>`)
    zip.file('xl/drawings/vmlDrawing1.vml', `<xml xmlns:v="urn:schemas-microsoft-com:vml">
      <v:shape id="WordArt_Xlsx" style="position:absolute;width:250pt;height:50pt;" fillcolor="#00FF00" strokecolor="#000000">
        <v:fill color="#00FF00"/>
        <v:stroke color="#000000" weight="1pt"/>
        <v:textpath on="True" style="font-family:'Calibri';font-size:28pt;" string="XLSX WordArt"/>
      </v:shape>
    </xml>`)

    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const workbook = await parseXlsx(pkg)
    const sheet = workbook.sheets[0]
    expect(sheet.drawings).toBeDefined()
    expect(sheet.drawings?.length).toBeGreaterThan(0)
    const drawing = sheet.drawings!.find(d => d.textBody?.paragraphs[0].runs[0].text === 'XLSX WordArt')
    expect(drawing).toBeDefined()
    expect(drawing?.textBody?.paragraphs[0].runs[0].fontFamily).toBe('Calibri')
  })
})
