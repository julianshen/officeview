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

    // Canvas rendering integration
    const { paintDrawing } = await import('../src/docx/drawing')
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(300, 100)
    const ctx = canvas.getContext('2d')
    expect(() => paintDrawing(img.drawing!, ctx as any, 300, 100)).not.toThrow()
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

    // Canvas rendering integration
    const { renderSheet } = await import('../src/xlsx/render')
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(400, 200)
    const ctx = canvas.getContext('2d')
    expect(() => renderSheet(sheet, ctx as any)).not.toThrow()
  })

  test('VML parser handles ST_TrueFalse shorthand, margin positioning, font unquoting, and stroke outline wiring', async () => {
    const { parseXml } = await import('../src/core/xml')
    const { parseVmlWordArt } = await import('../src/drawing/vml')

    // 1. Textpath with on="f" should return undefined
    const disabledXml = `<v:shape id="s_off"><v:textpath on="f" string="Ignored"/></v:shape>`
    expect(parseVmlWordArt(parseXml(disabledXml))).toBeUndefined()

    // 2. Margin positioning, font unquoting, numeric weight 800, stroke outline with default 1pt
    const xml = `<v:shape id="s_pos" style="position:absolute;margin-left:120pt;margin-top:60pt;width:240pt;height:45pt;" strokecolor="#FF0000" filled="f">
      <v:fill on="f"/>
      <v:textpath on="t" style="font-family:'Arial Black', sans-serif;font-size:24pt;font-weight:800;" string="Outlined Heading"/>
    </v:shape>`
    const res = parseVmlWordArt(parseXml(xml))
    expect(res).toBeDefined()
    expect(res?.leftPt).toBe(120)
    expect(res?.topPt).toBe(60)
    expect(res?.widthPt).toBe(240)
    expect(res?.heightPt).toBe(45)

    const run = res!.textBody.paragraphs[0].runs[0]
    expect(run.text).toBe('Outlined Heading')
    expect(run.fontFamily).toBe('Arial Black')
    expect(run.bold).toBe(true)
    expect(run.noFill).toBe(true)
    expect(run.textOutline).toBeDefined()
    expect(run.textOutline?.color).toBe('#FF0000')
    expect(run.textOutline?.widthPx).toBeCloseTo(1 * (96 / 72), 2)

    // 3. Stroked="f" should suppress textOutline even if strokecolor is specified
    const strokedFalseXml = `<v:shape id="s_nostroke" strokecolor="#0000FF" stroked="f">
      <v:textpath on="true" string="No Outline"/>
    </v:shape>`
    const resNoStroke = parseVmlWordArt(parseXml(strokedFalseXml))
    expect(resNoStroke).toBeDefined()
    expect(resNoStroke!.textBody.paragraphs[0].runs[0].textOutline).toBeUndefined()

    // 4. Unsupported fill type (gradient) emits diagnostic
    const gradientXml = `<v:shape id="s_grad">
      <v:fill type="gradient"/>
      <v:textpath on="true" string="Gradient WordArt"/>
    </v:shape>`
    const resGrad = parseVmlWordArt(parseXml(gradientXml))
    expect(resGrad).toBeDefined()
    expect(resGrad?.diagnostics?.length).toBeGreaterThan(0)
    expect(resGrad?.diagnostics?.[0].kind).toBe('unsupported-fill')
  })

  test('DOCX pict and XLSX legacyDrawing map VML positions to xEmu / yEmu and preserve diagnostics', async () => {
    const JSZip = (await import('jszip')).default
    const { OfficePackage } = await import('../src/core/zip')
    const { parseDocx } = await import('../src/docx/parse')
    const { parseXlsx } = await import('../src/xlsx/parse')

    // DOCX positioning and multi-shape in pict
    const docxZip = new JSZip()
    docxZip.file('[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`)
    docxZip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
    docxZip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml">
      <w:body>
        <w:p>
          <w:r>
            <w:pict>
              <v:shape id="Shape1" style="position:absolute;margin-left:72pt;margin-top:36pt;width:150pt;height:40pt;">
                <v:textpath on="true" string="Shape One"/>
              </v:shape>
              <v:shape id="Shape2" style="position:absolute;margin-left:144pt;margin-top:72pt;width:150pt;height:40pt;">
                <v:textpath on="true" string="Shape Two"/>
              </v:shape>
            </w:pict>
          </w:r>
        </w:p>
      </w:body>
    </w:document>`)
    const docxPkg = await OfficePackage.load(await docxZip.generateAsync({ type: 'uint8array' }))
    const docxDoc = await parseDocx(docxPkg)
    const shapes = (docxDoc.sections[0].paragraphs[0].images[0] as any).drawing.shapes
    expect(shapes.length).toBe(2)
    expect(shapes[0].xEmu).toBe(Math.round(72 * 12700))
    expect(shapes[0].yEmu).toBe(Math.round(36 * 12700))
    expect(shapes[1].xEmu).toBe(Math.round(144 * 12700))
    expect(shapes[1].yEmu).toBe(Math.round(72 * 12700))

    // XLSX positioning and diagnostics
    const xlsxZip = new JSZip()
    xlsxZip.file('[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>`)
    xlsxZip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`)
    xlsxZip.file('xl/workbook.xml', `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`)
    xlsxZip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`)
    xlsxZip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><v>100</v></c></row></sheetData><legacyDrawing r:id="rIdVml"/></worksheet>`)
    xlsxZip.file('xl/worksheets/_rels/sheet1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>`)
    xlsxZip.file('xl/drawings/vmlDrawing1.vml', `<xml xmlns:v="urn:schemas-microsoft-com:vml">
      <v:shape id="WordArt_Positioned" style="position:absolute;margin-left:50pt;margin-top:25pt;width:200pt;height:40pt;">
        <v:fill type="gradient"/>
        <v:textpath on="True" string="Positioned WordArt"/>
      </v:shape>
    </xml>`)
    const xlsxPkg = await OfficePackage.load(await xlsxZip.generateAsync({ type: 'uint8array' }))
    const xlsxDoc = await parseXlsx(xlsxPkg)
    const xlsxSheet = xlsxDoc.sheets[0]
    const posDrawing = xlsxSheet.drawings?.find(d => d.textBody?.paragraphs[0].runs[0].text === 'Positioned WordArt')
    expect(posDrawing).toBeDefined()
    expect(posDrawing?.xEmu).toBe(Math.round(50 * 12700))
    expect(posDrawing?.yEmu).toBe(Math.round(25 * 12700))
    expect(xlsxSheet.drawingDiagnostics?.some(d => d.kind === 'unsupported-fill')).toBe(true)
  })
})
