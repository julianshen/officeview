import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { getChildren, type XmlNode } from '../src/core/xml'
import { prepareDrawingContent } from '../src/drawing/content'
import { drawingPartContext } from '../src/drawing/parts'
import { paintDrawingContent } from '../src/drawing/content-paint'
import { parseTextBody, textFontDefaults } from '../src/drawing/text-parse'
import { parsePptx } from '../src/pptx/parse'
import { parseDocx } from '../src/docx/parse'

const DIAGRAM_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
const DGM = 'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'

async function packageWithParts(parts: Record<string, string | Uint8Array>): Promise<OfficePackage> {
  const zip = new JSZip()
  for (const [path, content] of Object.entries(parts)) {
    zip.file(path, content)
  }
  const bytes = await zip.generateAsync({ type: 'uint8array' })
  return OfficePackage.load(bytes)
}

const relsXml = (items: Array<{ id: string; type: string; target: string }>) =>
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  items.map(i => `<Relationship Id="${i.id}" Type="${i.type}" Target="${i.target}"/>`).join('') +
  `</Relationships>`

const DIAGRAM_DATA_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData'

// Cacheless dataModel: point text but no dataModelExt (no cached dsp:drawing).
// The third point carries no text and must not produce an empty paragraph.
const cachelessDataXml =
  `<dgm:dataModel ${DGM}>` +
  `  <dgm:ptLst>` +
  `    <dgm:pt modelId="{D1}"><dgm:prSet/><dgm:t>Alpha &amp; Omega</dgm:t></dgm:pt>` +
  `    <dgm:pt modelId="{D2}"><dgm:prSet/><dgm:t>Beta &lt;gamma&gt;</dgm:t></dgm:pt>` +
  `    <dgm:pt modelId="{D3}"><dgm:prSet/></dgm:pt>` +
  `  </dgm:ptLst>` +
  `  <dgm:cxnLst/>` +
  `  <dgm:whole/>` +
  `</dgm:dataModel>`

const theme = { colors: new Map(), fonts: new Map() }
const drawingTheme: any = { colors: {}, palette: {}, colorMap: {}, fonts: { major: { supplemental: {} }, minor: { supplemental: {} } }, fillStyles: [], bgFillStyles: [], lineStyles: [], effectStyles: [], issues: [] }

const graphicNode = {
  '@attrs': { uri: DIAGRAM_URI },
  relIds: { '@attrs': { dm: 'rDm1' } },
}

function cachelessPackage(extra: Record<string, string | Uint8Array> = {}) {
  return packageWithParts({
    'word/document.xml': '<w:document/>',
    'word/_rels/document.xml.rels': relsXml([
      { id: 'rDm1', type: DIAGRAM_DATA_REL, target: 'diagrams/data1.xml' },
    ]),
    'word/diagrams/data1.xml': cachelessDataXml,
    ...extra,
  })
}

describe('Phase 23: cacheless SmartArt text-only fallback', () => {
  test('dataModel without cached drawing falls back to searchable text instead of losing all text', async () => {
    const pkg = await cachelessPackage()
    const content = await prepareDrawingContent(pkg, graphicNode, 'word/document.xml', theme, drawingTheme)
    expect(content).toBeDefined()
    expect(content?.kind).toBe('diagram')
    if (content?.kind === 'diagram') {
      expect(content.shapes).toHaveLength(1)
      const texts = content.shapes[0].paragraphs.map((p) => p.runs.map((r) => r.text).join(''))
      expect(texts).toEqual(['Alpha & Omega', 'Beta <gamma>'])
      expect(content.shapes[0].textBody).toBeUndefined()
    }
    const ctx = drawingPartContext(pkg)
    expect(ctx.diagnostics.some((d) => d.kind === 'unsupported-content' && d.reason === 'cacheless-smartart-text-fallback')).toBe(true)
  })

  test('fallback textBody parses through the docx adapter pipeline', async () => {
    const pkg = await cachelessPackage()
    const adapters = {
      parseDiagramText: (body: XmlNode, shape: XmlNode) =>
        parseTextBody(body, drawingTheme, undefined, textFontDefaults(getChildren(shape, 'style')[0], drawingTheme)),
    }
    const content = await prepareDrawingContent(pkg, graphicNode, 'word/document.xml', theme, drawingTheme, adapters)
    expect(content?.kind).toBe('diagram')
    if (content?.kind === 'diagram') {
      const shape = content.shapes[0]
      expect(shape.textBody).toBeDefined()
      const serialized = JSON.stringify(shape.textBody)
      expect(serialized).toContain('Alpha & Omega')
      expect(serialized).toContain('Beta <gamma>')
      expect(shape.paragraphs.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['Alpha & Omega', 'Beta <gamma>'])
    }
  })

  test('fallback textBody parses through the pptx adapter pipeline', async () => {
    const pkg = await cachelessPackage()
    const pptxTheme: any = { ...drawingTheme }
    const adapters = {
      parseDiagramText: (body: XmlNode, shape: XmlNode) =>
        parseTextBody(body, pptxTheme, undefined, { fontFamily: 'PptxProbe', ...textFontDefaults(getChildren(shape, 'style')[0], pptxTheme) }),
    }
    const content = await prepareDrawingContent(pkg, graphicNode, 'word/document.xml', theme, drawingTheme, adapters)
    expect(content?.kind).toBe('diagram')
    if (content?.kind === 'diagram') {
      const shape = content.shapes[0]
      expect(shape.textBody).toBeDefined()
      const textBodies = (shape.textBody as unknown as { paragraphs: Array<{ runs: Array<{ text: string }> }> }).paragraphs
      expect(textBodies).toHaveLength(2)
      expect(textBodies.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['Alpha & Omega', 'Beta <gamma>'])
      expect(JSON.stringify(shape.textBody)).toContain('PptxProbe')
    }
  })

  test('fallback delegates text parsing to adapters.parseDiagramText', async () => {
    const pkg = await cachelessPackage()
    const calls: XmlNode[] = []
    const marker = { marker: 'adapter-text-body' }
    const content = await prepareDrawingContent(
      pkg, graphicNode, 'word/document.xml', theme, drawingTheme,
      { parseDiagramText: (body: XmlNode) => { calls.push(body); return marker as never } },
    )
    expect(calls).toHaveLength(1)
    expect(JSON.stringify(calls[0])).toContain('Alpha & Omega')
    expect(content?.kind).toBe('diagram')
    if (content?.kind === 'diagram') expect(content.shapes[0].textBody).toEqual(marker)
  })

  test('broken cached drawing with dataModel text still falls back to text', async () => {
    // dataModelExt points at a cached drawing part that does not exist.
    const brokenExtDataXml =
      `<dgm:dataModel ${DGM}>` +
      `  <dgm:ptLst><dgm:pt modelId="{D1}"><dgm:prSet/><dgm:t>Rescued text</dgm:t></dgm:pt></dgm:ptLst>` +
      `  <dgm:cxnLst/>` +
      `  <dgm:whole/>` +
      `  <dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rDg1"/></dgm:ext></dgm:extLst>` +
      `</dgm:dataModel>`
    const pkg2 = await packageWithParts({
      'word/document.xml': '<w:document/>',
      'word/_rels/document.xml.rels': relsXml([{ id: 'rDm1', type: DIAGRAM_DATA_REL, target: 'diagrams/data1.xml' }]),
      'word/diagrams/data1.xml': brokenExtDataXml,
      'word/diagrams/_rels/data1.xml.rels': relsXml([
        { id: 'rDg1', type: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', target: 'drawing1.xml' },
      ]),
    })
    const content = await prepareDrawingContent(pkg2, graphicNode, 'word/document.xml', theme, drawingTheme)
    expect(content?.kind).toBe('diagram')
    if (content?.kind === 'diagram') {
      expect(content.shapes[0].paragraphs.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['Rescued text'])
    }
    const ctx = drawingPartContext(pkg2)
    expect(ctx.diagnostics.some((d) => d.kind === 'missing-part')).toBe(true)
  })

  test('textless cacheless dataModel still resolves to undefined without throwing', async () => {
    const pkg = await packageWithParts({
      'word/document.xml': '<w:document/>',
      'word/_rels/document.xml.rels': relsXml([{ id: 'rDm1', type: DIAGRAM_DATA_REL, target: 'diagrams/data1.xml' }]),
      'word/diagrams/data1.xml':
        `<dgm:dataModel ${DGM}><dgm:ptLst><dgm:pt modelId="{D1}"><dgm:prSet/></dgm:pt></dgm:ptLst><dgm:cxnLst/><dgm:whole/></dgm:dataModel>`,
    })
    const content = await prepareDrawingContent(pkg, graphicNode, 'word/document.xml', theme, drawingTheme)
    expect(content).toBeUndefined()
  })

  test('fallback diagram paints without throwing and exposes textBody to paintDiagramText', async () => {
    const pkg = await cachelessPackage()
    const adapters = {
      parseDiagramText: (body: XmlNode, shape: XmlNode) =>
        parseTextBody(body, drawingTheme, undefined, textFontDefaults(getChildren(shape, 'style')[0], drawingTheme)),
    }
    const content = await prepareDrawingContent(pkg, graphicNode, 'word/document.xml', theme, drawingTheme, adapters)
    expect(content?.kind).toBe('diagram')
    const canvas = createCanvas(200, 100)
    const ctx = canvas.getContext('2d')
    let seen: unknown
    expect(() => paintDrawingContent(content!, ctx as never, 200, 100, {
      paintDiagramText: (shape) => { seen = shape },
    })).not.toThrow()
    expect(seen).toBeDefined()
  })

  test('PPTX end-to-end: cacheless SmartArt reports text-only coverage', async () => {
    const zip = new JSZip()
    zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="952500" cy="952500"/><p:sldIdLst><p:sldId r:id="s1"/></p:sldIdLst></p:presentation>')
    zip.file('ppt/_rels/presentation.xml.rels', relsXml([{ id: 's1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', target: 'slides/slide1.xml' }]))
    zip.file('ppt/slides/slide1.xml',
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><p:cSld><p:spTree>` +
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="81" name="smartart1"/></p:nvGraphicFramePr>` +
      `<p:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></p:xfrm>` +
      `<a:graphic><a:graphicData uri="${DIAGRAM_URI}"><dgm:relIds dm="rDm1"/></a:graphicData></a:graphic>` +
      `</p:graphicFrame>` +
      `</p:spTree></p:cSld></p:sld>`)
    zip.file('ppt/slides/_rels/slide1.xml.rels', relsXml([{ id: 'rDm1', type: DIAGRAM_DATA_REL, target: '../diagrams/data1.xml' }]))
    zip.file('ppt/diagrams/data1.xml', cachelessDataXml)
    const doc = await parsePptx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const frame = doc.slides[0].shapes[0]
    expect(frame.content?.kind).toBe('diagram')
    if (frame.content?.kind === 'diagram') {
      expect(frame.content.shapes[0].paragraphs.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['Alpha & Omega', 'Beta <gamma>'])
    }
    expect(doc.drawingCoverage!.find((e) => e.id === '81')).toMatchObject({ selectedRepresentation: 'text-only' })
  })

  test('DOCX end-to-end: cacheless SmartArt reports text-only coverage', async () => {
    const zip = new JSZip()
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    zip.file('word/_rels/document.xml.rels', relsXml([{ id: 'rDm1', type: DIAGRAM_DATA_REL, target: 'diagrams/data1.xml' }]))
    zip.file('word/document.xml',
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r>` +
      `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="7" name="smartart1"/><a:graphic><a:graphicData uri="${DIAGRAM_URI}"><dgm:relIds dm="rDm1"/></a:graphicData></a:graphic></wp:inline></w:drawing>` +
      `</w:r></w:p></w:body></w:document>`)
    zip.file('word/diagrams/data1.xml', cachelessDataXml)
    const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const entry = doc.drawingCoverage!.find((e) => e.element === 'drawing' && e.feature === 'diagram')
    expect(entry).toBeDefined()
    expect(entry).toMatchObject({ selectedRepresentation: 'text-only' })
  })
})
