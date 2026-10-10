import { describe, expect, test } from 'vitest'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { paintDrawing } from '../src/docx/drawing'

function emu(px: number): number {
  return Math.round(px * 9525)
}

async function createDocxWithTextbox(bodyPrAttrs: string, text: string = 'Hello World'): Promise<any> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `</Types>`)
  zip.file('_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships>`)

  const docxXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <w:body>
        <w:p>
          <w:r>
            <w:drawing>
              <wp:inline distT="0" distB="0" distL="0" distR="0">
                <wp:extent cx="${emu(200)}" cy="${emu(200)}"/>
                <wp:docPr id="1" name="Shape 1"/>
                <a:graphic>
                  <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                    <wps:wsp>
                      <wps:cNvSpPr txBox="1"/>
                      <wps:spPr>
                        <a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(200)}"/></a:xfrm>
                        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                      </wps:spPr>
                      <wps:txbx id="1">
                        <w:txbxContent>
                          <w:p>
                            <w:r>
                              <w:rPr><w:sz w:val="24"/></w:rPr>
                              <w:t>${text}</w:t>
                            </w:r>
                          </w:p>
                        </w:txbxContent>
                      </wps:txbx>
                      <wps:bodyPr ${bodyPrAttrs}/>
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
  return parseDocx(pkg)
}

describe('DOCX Textbox Positioning & Wrap Alignment', () => {
  test('parses wps:bodyPr anchor and wrap into textbox drawing model', async () => {
    const doc1 = await createDocxWithTextbox('anchor="ctr" wrap="none"')
    const tb1 = doc1.sections[0].paragraphs[0].images[0].drawing
    expect(tb1?.kind).toBe('textbox')
    expect((tb1 as any).anchor).toBe('ctr')
    expect((tb1 as any).wrap).toBe(false)

    const doc2 = await createDocxWithTextbox('anchor="b" wrap="square"')
    const tb2 = doc2.sections[0].paragraphs[0].images[0].drawing
    expect(tb2?.kind).toBe('textbox')
    expect((tb2 as any).anchor).toBe('b')
    expect((tb2 as any).wrap).toBe(true)

    const doc3 = await createDocxWithTextbox('')
    const tb3 = doc3.sections[0].paragraphs[0].images[0].drawing
    expect(tb3?.kind).toBe('textbox')
    expect((tb3 as any).anchor).toBeUndefined()
    expect((tb3 as any).wrap).toBe(true)

    const doc4 = await createDocxWithTextbox('anchor=" CTR " wrap=" NONE "')
    const tb4 = doc4.sections[0].paragraphs[0].images[0].drawing
    expect((tb4 as any).anchor).toBe('ctr')
    expect((tb4 as any).wrap).toBe(false)

    const doc5 = await createDocxWithTextbox('anchor="dist"')
    const tb5 = doc5.sections[0].paragraphs[0].images[0].drawing
    expect((tb5 as any).anchor).toBeUndefined()
    expect((tb5 as any).diagnostics?.some((d: any) => d.kind === 'unsupported-text-alignment' && d.feature === 'anchor-dist')).toBe(true)
  })

  test('vertical anchoring positions text baseline according to top, center, and bottom anchors', async () => {
    const docTop = await createDocxWithTextbox('anchor="t"')
    const docCtr = await createDocxWithTextbox('anchor="ctr"')
    const docBtm = await createDocxWithTextbox('anchor="b"')

    const getPaintedY = (drawing: any): number => {
      const canvas = createCanvas(200, 200)
      const ctx = canvas.getContext('2d')
      let filledY = -1
      const origFillText = ctx.fillText.bind(ctx)
      ctx.fillText = (text, x, y) => {
        filledY = Number(y)
        origFillText(text, x, y)
      }
      paintDrawing(drawing, ctx as never, 200, 200)
      return filledY
    }

    const yTop = getPaintedY(docTop.sections[0].paragraphs[0].images[0].drawing)
    const yCtr = getPaintedY(docCtr.sections[0].paragraphs[0].images[0].drawing)
    const yBtm = getPaintedY(docBtm.sections[0].paragraphs[0].images[0].drawing)

    expect(yTop).toBeGreaterThan(0)
    expect(yCtr).toBeGreaterThan(yTop + 50)
    expect(yBtm).toBeGreaterThan(yCtr + 50)
  })

  test('wrapping mode controls whether long text wraps into multiple lines', async () => {
    const longText = 'This is a very long text line that should wrap into multiple lines when square wrap is enabled but remain on one line when wrap is none.'
    const docWrapNone = await createDocxWithTextbox('wrap="none"', longText)
    const docWrapSquare = await createDocxWithTextbox('wrap="square"', longText)

    const countLines = (drawing: any): number => {
      const canvas = createCanvas(100, 200)
      const ctx = canvas.getContext('2d')
      let lines = 0
      const origFillText = ctx.fillText.bind(ctx)
      ctx.fillText = (text, x, y) => {
        lines++
        origFillText(text, x, y)
      }
      paintDrawing(drawing, ctx as never, 100, 200)
      return lines
    }

    const linesNone = countLines(docWrapNone.sections[0].paragraphs[0].images[0].drawing)
    const linesSquare = countLines(docWrapSquare.sections[0].paragraphs[0].images[0].drawing)

    expect(linesNone).toBe(1)
    expect(linesSquare).toBeGreaterThan(1)
  })
})
