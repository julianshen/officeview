import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { prepareDrawingContent } from '../src/drawing/content'
import { drawingPartContext } from '../src/drawing/parts'
import { paintDrawingContent } from '../src/drawing/content-paint'

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

describe('Word cached-path parity in <dsp:spTree>', () => {
  test('cached leaf flipH and flipV in <dsp:spTree> compose into shape and paint', async () => {
    const drawingXml =
      `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <dsp:spTree>` +
      `    <dsp:sp>` +
      `      <dsp:spPr>` +
      `        <a:xfrm rot="5400000" flipH="1" flipV="1">` +
      `          <a:off x="100000" y="200000"/>` +
      `          <a:ext cx="300000" cy="400000"/>` +
      `        </a:xfrm>` +
      `        <a:prstGeom prst="rect"/>` +
      `        <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>` +
      `      </dsp:spPr>` +
      `    </dsp:sp>` +
      `  </dsp:spTree>` +
      `</dsp:drawing>`

    const dataXml =
      `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
      `  <dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rDg1"/></dgm:ext></dgm:extLst>` +
      `</dgm:dataModel>`

    const pkg = await packageWithParts({
      'word/document.xml': '<w:document/>',
      'word/_rels/document.xml.rels': relsXml([
        { id: 'rDm1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData', target: 'diagrams/data1.xml' },
      ]),
      'word/diagrams/data1.xml': dataXml,
      'word/diagrams/_rels/data1.xml.rels': relsXml([
        { id: 'rDg1', type: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', target: 'drawing1.xml' },
      ]),
      'word/diagrams/drawing1.xml': drawingXml,
    })

    const theme = { colors: new Map(), fonts: new Map() }
    const drawingTheme: any = { colors: {}, palette: {}, colorMap: {}, fonts: { major: { supplemental: {} }, minor: { supplemental: {} } }, fillStyles: [], bgFillStyles: [], lineStyles: [], effectStyles: [], issues: [] }

    const graphicNode = {
      '@attrs': { uri: 'http://schemas.openxmlformats.org/drawingml/2006/diagram' },
      relIds: { '@attrs': { dm: 'rDm1' } },
    }

    const content = await prepareDrawingContent(pkg, graphicNode, 'word/document.xml', theme, drawingTheme)
    expect(content).toBeDefined()
    expect(content?.kind).toBe('diagram')
    if (content?.kind === 'diagram') {
      expect(content.shapes.length).toBe(1)
      const s = content.shapes[0]
      expect(s.flipH).toBe(true)
      expect(s.flipV).toBe(true)
      expect(s.rotationDeg).toBe(90)
    }
  })

  test('cached leaf flip spellings t/True parse like 1/true', async () => {
    const drawingXml =
      `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <dsp:spTree>` +
      `    <dsp:sp>` +
      `      <dsp:spPr>` +
      `        <a:xfrm flipH="t" flipV="True">` +
      `          <a:off x="0" y="0"/>` +
      `          <a:ext cx="100000" cy="100000"/>` +
      `        </a:xfrm>` +
      `        <a:prstGeom prst="rect"/>` +
      `      </dsp:spPr>` +
      `    </dsp:sp>` +
      `  </dsp:spTree>` +
      `</dsp:drawing>`

    const dataXml =
      `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
      `  <dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rDg1"/></dgm:ext></dgm:extLst>` +
      `</dgm:dataModel>`

    const pkg = await packageWithParts({
      'word/document.xml': '<w:document/>',
      'word/_rels/document.xml.rels': relsXml([
        { id: 'rDm1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData', target: 'diagrams/data1.xml' },
      ]),
      'word/diagrams/data1.xml': dataXml,
      'word/diagrams/_rels/data1.xml.rels': relsXml([
        { id: 'rDg1', type: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', target: 'drawing1.xml' },
      ]),
      'word/diagrams/drawing1.xml': drawingXml,
    })

    const theme = { colors: new Map(), fonts: new Map() }
    const drawingTheme: any = { colors: {}, palette: {}, colorMap: {}, fonts: { major: { supplemental: {} }, minor: { supplemental: {} } }, fillStyles: [], bgFillStyles: [], lineStyles: [], effectStyles: [], issues: [] }
    const graphicNode = {
      '@attrs': { uri: 'http://schemas.openxmlformats.org/drawingml/2006/diagram' },
      relIds: { '@attrs': { dm: 'rDm1' } },
    }

    const content = await prepareDrawingContent(pkg, graphicNode, 'word/document.xml', theme, drawingTheme)
    expect(content?.kind).toBe('diagram')
    if (content?.kind === 'diagram') {
      expect(content.shapes.length).toBe(1)
      expect(content.shapes[0].flipH).toBe(true)
      expect(content.shapes[0].flipV).toBe(true)
    }
  })

  test('cached pic inside <dsp:spTree> resolves picture data and removes cached-picture-unsupported diagnostic', async () => {
    const fakePng = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    const drawingXml =
      `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `  <dsp:spTree>` +
      `    <dsp:pic>` +
      `      <dsp:nvPicPr>` +
      `        <dsp:cNvPr id="2" name="Image 1"/>` +
      `        <dsp:cNvPicPr/>` +
      `      </dsp:nvPicPr>` +
      `      <dsp:blipFill>` +
      `        <a:blip r:embed="rPic1"/>` +
      `      </dsp:blipFill>` +
      `      <dsp:spPr>` +
      `        <a:xfrm>` +
      `          <a:off x="0" y="0"/>` +
      `          <a:ext cx="914400" cy="914400"/>` +
      `        </a:xfrm>` +
      `        <a:prstGeom prst="rect"/>` +
      `      </dsp:spPr>` +
      `    </dsp:pic>` +
      `  </dsp:spTree>` +
      `</dsp:drawing>`

    const dataXml =
      `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
      `  <dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rDg1"/></dgm:ext></dgm:extLst>` +
      `</dgm:dataModel>`

    const pkg = await packageWithParts({
      'word/document.xml': '<w:document/>',
      'word/_rels/document.xml.rels': relsXml([
        { id: 'rDm1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData', target: 'diagrams/data1.xml' },
      ]),
      'word/diagrams/data1.xml': dataXml,
      'word/diagrams/_rels/data1.xml.rels': relsXml([
        { id: 'rDg1', type: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', target: 'drawing1.xml' },
      ]),
      'word/diagrams/drawing1.xml': drawingXml,
      'word/diagrams/_rels/drawing1.xml.rels': relsXml([
        { id: 'rPic1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', target: '../media/image1.png' },
      ]),
      'word/media/image1.png': fakePng,
    })

    const theme = { colors: new Map(), fonts: new Map() }
    const drawingTheme: any = { colors: {}, palette: {}, colorMap: {}, fonts: { major: { supplemental: {} }, minor: { supplemental: {} } }, fillStyles: [], bgFillStyles: [], lineStyles: [], effectStyles: [], issues: [] }

    const graphicNode = {
      '@attrs': { uri: 'http://schemas.openxmlformats.org/drawingml/2006/diagram' },
      relIds: { '@attrs': { dm: 'rDm1' } },
    }

    const content = await prepareDrawingContent(pkg, graphicNode, 'word/document.xml', theme, drawingTheme)
    expect(content).toBeDefined()
    expect(content?.kind).toBe('diagram')
    if (content?.kind === 'diagram') {
      expect(content.shapes.length).toBeGreaterThan(0)
      const picShape = content.shapes[0]
      expect(picShape.image).toBeDefined()
      expect(picShape.image?.data).toEqual(fakePng)
    }

    const ctx = drawingPartContext(pkg)
    expect(ctx.diagnostics.some(d => d.reason === 'cached-picture-unsupported')).toBe(false)
  })

  test('cached graphicFrame tables in <dsp:spTree> parse and paint with table grid geometry', async () => {
    const drawingXml =
      `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <dsp:spTree>` +
      `    <dsp:graphicFrame>` +
      `      <dsp:nvGraphicFramePr>` +
      `        <dsp:cNvPr id="5" name="Table 1"/>` +
      `        <dsp:cNvGraphicFramePr/>` +
      `      </dsp:nvGraphicFramePr>` +
      `      <dsp:xfrm>` +
      `        <a:off x="100000" y="200000"/>` +
      `        <a:ext cx="800000" cy="600000"/>` +
      `      </dsp:xfrm>` +
      `      <a:graphic>` +
      `        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
      `          <a:tbl>` +
      `            <a:tblGrid>` +
      `              <a:gridCol w="400000"/>` +
      `              <a:gridCol w="400000"/>` +
      `            </a:tblGrid>` +
      `            <a:tr h="300000">` +
      `              <a:tc>` +
      `                <a:txBody><a:bodyPr/><a:p><a:r><a:t>Cell A1</a:t></a:r></a:p></a:txBody>` +
      `                <a:tcPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:tcPr>` +
      `              </a:tc>` +
      `              <a:tc>` +
      `                <a:txBody><a:bodyPr/><a:p><a:r><a:t>Cell B1</a:t></a:r></a:p></a:txBody>` +
      `                <a:tcPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:tcPr>` +
      `              </a:tc>` +
      `            </a:tr>` +
      `            <a:tr h="300000">` +
      `              <a:tc>` +
      `                <a:txBody><a:bodyPr/><a:p><a:r><a:t>Cell A2</a:t></a:r></a:p></a:txBody>` +
      `                <a:tcPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:tcPr>` +
      `              </a:tc>` +
      `              <a:tc>` +
      `                <a:txBody><a:bodyPr/><a:p><a:r><a:t>Cell B2</a:t></a:r></a:p></a:txBody>` +
      `                <a:tcPr><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></a:tcPr>` +
      `              </a:tc>` +
      `            </a:tr>` +
      `          </a:tbl>` +
      `        </a:graphicData>` +
      `      </a:graphic>` +
      `    </dsp:graphicFrame>` +
      `  </dsp:spTree>` +
      `</dsp:drawing>`

    const dataXml =
      `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
      `  <dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rDg1"/></dgm:ext></dgm:extLst>` +
      `</dgm:dataModel>`

    const pkg = await packageWithParts({
      'word/document.xml': '<w:document/>',
      'word/_rels/document.xml.rels': relsXml([
        { id: 'rDm1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData', target: 'diagrams/data1.xml' },
      ]),
      'word/diagrams/data1.xml': dataXml,
      'word/diagrams/_rels/data1.xml.rels': relsXml([
        { id: 'rDg1', type: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', target: 'drawing1.xml' },
      ]),
      'word/diagrams/drawing1.xml': drawingXml,
    })

    const theme = { colors: new Map(), fonts: new Map() }
    const drawingTheme: any = { colors: {}, palette: {}, colorMap: {}, fonts: { major: { supplemental: {} }, minor: { supplemental: {} } }, fillStyles: [], bgFillStyles: [], lineStyles: [], effectStyles: [], issues: [] }

    const graphicNode = {
      '@attrs': { uri: 'http://schemas.openxmlformats.org/drawingml/2006/diagram' },
      relIds: { '@attrs': { dm: 'rDm1' } },
    }

    const content = await prepareDrawingContent(pkg, graphicNode, 'word/document.xml', theme, drawingTheme)
    expect(content).toBeDefined()
    expect(content?.kind).toBe('diagram')
    if (content?.kind === 'diagram') {
      expect(content.shapes.length).toBe(4)
      const [c1, c2, c3, c4] = content.shapes
      // Cell 1 at (100000, 200000)
      expect(c1.xEmu).toBe(100000)
      expect(c1.yEmu).toBe(200000)
      expect(c1.widthEmu).toBe(400000)
      expect(c1.heightEmu).toBe(300000)
      expect(c1.fill).toBe('FF0000')

      // Cell 2 at (500000, 200000)
      expect(c2.xEmu).toBe(500000)
      expect(c2.yEmu).toBe(200000)
      expect(c2.fill).toBe('00FF00')

      // Cell 3 at (100000, 500000)
      expect(c3.xEmu).toBe(100000)
      expect(c3.yEmu).toBe(500000)
      expect(c3.fill).toBe('0000FF')

      // Cell 4 at (500000, 500000)
      expect(c4.xEmu).toBe(500000)
      expect(c4.yEmu).toBe(500000)
      expect(c4.fill).toBe('FFFF00')
    }

    const ctx = drawingPartContext(pkg)
    expect(ctx.diagnostics.some(d => d.reason === 'no-supported-shapes')).toBe(false)
  })

  test('cached pic with missing relationship target emits missing-part diagnostic', async () => {
    const drawingXml =
      `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `  <dsp:spTree>` +
      `    <dsp:pic>` +
      `      <dsp:nvPicPr><dsp:cNvPr id="2" name="Missing Img"/><dsp:cNvPicPr/></dsp:nvPicPr>` +
      `      <dsp:blipFill><a:blip r:embed="rPicMissing"/></dsp:blipFill>` +
      `      <dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></a:xfrm></dsp:spPr>` +
      `    </dsp:pic>` +
      `  </dsp:spTree>` +
      `</dsp:drawing>`

    const dataXml =
      `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
      `  <dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rDg1"/></dgm:ext></dgm:extLst>` +
      `</dgm:dataModel>`

    const pkg = await packageWithParts({
      'word/document.xml': '<w:document/>',
      'word/_rels/document.xml.rels': relsXml([
        { id: 'rDm1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData', target: 'diagrams/data1.xml' },
      ]),
      'word/diagrams/data1.xml': dataXml,
      'word/diagrams/_rels/data1.xml.rels': relsXml([
        { id: 'rDg1', type: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', target: 'drawing1.xml' },
      ]),
      'word/diagrams/drawing1.xml': drawingXml,
      'word/diagrams/_rels/drawing1.xml.rels': relsXml([
        { id: 'rPicMissing', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', target: '../media/does-not-exist.png' },
      ]),
    })

    const theme = { colors: new Map(), fonts: new Map() }
    const drawingTheme: any = { colors: {}, palette: {}, colorMap: {}, fonts: { major: { supplemental: {} }, minor: { supplemental: {} } }, fillStyles: [], bgFillStyles: [], lineStyles: [], effectStyles: [], issues: [] }
    const graphicNode = {
      '@attrs': { uri: 'http://schemas.openxmlformats.org/drawingml/2006/diagram' },
      relIds: { '@attrs': { dm: 'rDm1' } },
    }

    const content = await prepareDrawingContent(pkg, graphicNode, 'word/document.xml', theme, drawingTheme)
    expect(content).toBeUndefined()
    const ctx = drawingPartContext(pkg)
    expect(ctx.diagnostics.some(d => d.kind === 'missing-part' && d.identity === 'rPicMissing')).toBe(true)
  })

  test('cached table paints onto canvas and handles missing grid columns gracefully', async () => {
    const drawingXml =
      `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <dsp:spTree>` +
      `    <dsp:graphicFrame>` +
      `      <dsp:xfrm><a:off x="0" y="0"/><a:ext cx="600000" cy="400000"/></dsp:xfrm>` +
      `      <a:graphic>` +
      `        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
      `          <a:tbl>` +
      `            <a:tblGrid/>` +
      `            <a:tr>` +
      `              <a:tc><a:tcPr><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill></a:tcPr></a:tc>` +
      `              <a:tc><a:tcPr><a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill></a:tcPr></a:tc>` +
      `            </a:tr>` +
      `          </a:tbl>` +
      `        </a:graphicData>` +
      `      </a:graphic>` +
      `    </dsp:graphicFrame>` +
      `  </dsp:spTree>` +
      `</dsp:drawing>`

    const dataXml =
      `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
      `  <dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rDg1"/></dgm:ext></dgm:extLst>` +
      `</dgm:dataModel>`

    const pkg = await packageWithParts({
      'word/document.xml': '<w:document/>',
      'word/_rels/document.xml.rels': relsXml([
        { id: 'rDm1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData', target: 'diagrams/data1.xml' },
      ]),
      'word/diagrams/data1.xml': dataXml,
      'word/diagrams/_rels/data1.xml.rels': relsXml([
        { id: 'rDg1', type: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', target: 'drawing1.xml' },
      ]),
      'word/diagrams/drawing1.xml': drawingXml,
    })

    const theme = { colors: new Map(), fonts: new Map() }
    const drawingTheme: any = { colors: {}, palette: {}, colorMap: {}, fonts: { major: { supplemental: {} }, minor: { supplemental: {} } }, fillStyles: [], bgFillStyles: [], lineStyles: [], effectStyles: [], issues: [] }
    const graphicNode = {
      '@attrs': { uri: 'http://schemas.openxmlformats.org/drawingml/2006/diagram' },
      relIds: { '@attrs': { dm: 'rDm1' } },
    }

    const content = await prepareDrawingContent(pkg, graphicNode, 'word/document.xml', theme, drawingTheme)
    expect(content).toBeDefined()
    if (content?.kind === 'diagram') {
      expect(content.shapes.length).toBe(2)
      expect(content.shapes[0].widthEmu).toBe(300000)
      expect(content.shapes[1].widthEmu).toBe(300000)
      expect(content.shapes[1].xEmu).toBe(300000)

      const canvas = createCanvas(100, 100)
      const ctx = canvas.getContext('2d')
      expect(() => paintDrawingContent(content, ctx as any, 100, 100)).not.toThrow()
    }
  })

  test('cached pic paints via assets.imageFor', async () => {
    const fakePng = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
    const drawingXml =
      `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `  <dsp:spTree>` +
      `    <dsp:pic>` +
      `      <dsp:nvPicPr><dsp:cNvPr id="2" name="Img"/><dsp:cNvPicPr/></dsp:nvPicPr>` +
      `      <dsp:blipFill><a:blip r:embed="rPic1"/></dsp:blipFill>` +
      `      <dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></dsp:spPr>` +
      `    </dsp:pic>` +
      `  </dsp:spTree>` +
      `</dsp:drawing>`

    const dataXml =
      `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
      `  <dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rDg1"/></dgm:ext></dgm:extLst>` +
      `</dgm:dataModel>`

    const pkg = await packageWithParts({
      'word/document.xml': '<w:document/>',
      'word/_rels/document.xml.rels': relsXml([
        { id: 'rDm1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData', target: 'diagrams/data1.xml' },
      ]),
      'word/diagrams/data1.xml': dataXml,
      'word/diagrams/_rels/data1.xml.rels': relsXml([
        { id: 'rDg1', type: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', target: 'drawing1.xml' },
      ]),
      'word/diagrams/drawing1.xml': drawingXml,
      'word/diagrams/_rels/drawing1.xml.rels': relsXml([
        { id: 'rPic1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', target: '../media/image1.png' },
      ]),
      'word/media/image1.png': fakePng,
    })

    const theme = { colors: new Map(), fonts: new Map() }
    const drawingTheme: any = { colors: {}, palette: {}, colorMap: {}, fonts: { major: { supplemental: {} }, minor: { supplemental: {} } }, fillStyles: [], bgFillStyles: [], lineStyles: [], effectStyles: [], issues: [] }
    const graphicNode = {
      '@attrs': { uri: 'http://schemas.openxmlformats.org/drawingml/2006/diagram' },
      relIds: { '@attrs': { dm: 'rDm1' } },
    }

    const content = await prepareDrawingContent(pkg, graphicNode, 'word/document.xml', theme, drawingTheme)
    expect(content).toBeDefined()
    let imageForCalled = false
    const canvas = createCanvas(100, 100)
    const ctx = canvas.getContext('2d')
    const fakeImageSource = createCanvas(10, 10)
    paintDrawingContent(content!, ctx as any, 100, 100, {
      assets: {
        imageFor: (img) => {
          imageForCalled = true
          expect(img.data).toEqual(fakePng)
          return fakeImageSource as any
        }
      }
    })
    expect(imageForCalled).toBe(true)
  })

  test('multi-picture diagram with broken first picture and valid second picture loads second picture', async () => {
    const fakePng = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
    const drawingXml =
      `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `  <dsp:spTree>` +
      `    <dsp:pic>` +
      `      <dsp:nvPicPr><dsp:cNvPr id="1" name="Broken"/><dsp:cNvPicPr/></dsp:nvPicPr>` +
      `      <dsp:blipFill><a:blip r:embed="rPicBroken"/></dsp:blipFill>` +
      `      <dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></a:xfrm></dsp:spPr>` +
      `    </dsp:pic>` +
      `    <dsp:pic>` +
      `      <dsp:nvPicPr><dsp:cNvPr id="2" name="Valid"/><dsp:cNvPicPr/></dsp:nvPicPr>` +
      `      <dsp:blipFill><a:blip r:embed="rPicValid"/></dsp:blipFill>` +
      `      <dsp:spPr><a:xfrm><a:off x="100000" y="0"/><a:ext cx="200000" cy="200000"/></a:xfrm></dsp:spPr>` +
      `    </dsp:pic>` +
      `  </dsp:spTree>` +
      `</dsp:drawing>`

    const dataXml =
      `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
      `  <dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rDg1"/></dgm:ext></dgm:extLst>` +
      `</dgm:dataModel>`

    const pkg = await packageWithParts({
      'word/document.xml': '<w:document/>',
      'word/_rels/document.xml.rels': relsXml([
        { id: 'rDm1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData', target: 'diagrams/data1.xml' },
      ]),
      'word/diagrams/data1.xml': dataXml,
      'word/diagrams/_rels/data1.xml.rels': relsXml([
        { id: 'rDg1', type: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', target: 'drawing1.xml' },
      ]),
      'word/diagrams/drawing1.xml': drawingXml,
      'word/diagrams/_rels/drawing1.xml.rels': relsXml([
        { id: 'rPicBroken', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', target: '../media/missing.png' },
        { id: 'rPicValid', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', target: '../media/valid.png' },
      ]),
      'word/media/valid.png': fakePng,
    })

    const theme = { colors: new Map(), fonts: new Map() }
    const drawingTheme: any = { colors: {}, palette: {}, colorMap: {}, fonts: { major: { supplemental: {} }, minor: { supplemental: {} } }, fillStyles: [], bgFillStyles: [], lineStyles: [], effectStyles: [], issues: [] }
    const graphicNode = {
      '@attrs': { uri: 'http://schemas.openxmlformats.org/drawingml/2006/diagram' },
      relIds: { '@attrs': { dm: 'rDm1' } },
    }

    const content = await prepareDrawingContent(pkg, graphicNode, 'word/document.xml', theme, drawingTheme)
    expect(content).toBeDefined()
    if (content?.kind === 'diagram') {
      expect(content.shapes.length).toBe(1)
      expect(content.shapes[0].xEmu).toBe(100000)
      expect(content.shapes[0].image?.data).toEqual(fakePng)
    }
    const ctx = drawingPartContext(pkg)
    expect(ctx.diagnostics.some(d => d.kind === 'missing-part' && d.identity === 'rPicBroken')).toBe(true)
  })

  test('cached table enforces DOCUMENT_DRAWING_NODE_LIMIT on cells', async () => {
    // Generate a table with 10001 cells
    const rows: string[] = []
    for (let r = 0; r < 101; r++) {
      const cells: string[] = []
      for (let c = 0; c < 100; c++) {
        cells.push(`<a:tc><a:tcPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:tcPr></a:tc>`)
      }
      rows.push(`<a:tr h="10000">${cells.join('')}</a:tr>`)
    }

    const drawingXml =
      `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <dsp:spTree>` +
      `    <dsp:graphicFrame>` +
      `      <dsp:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></dsp:xfrm>` +
      `      <a:graphic>` +
      `        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
      `          <a:tbl>${rows.join('')}</a:tbl>` +
      `        </a:graphicData>` +
      `      </a:graphic>` +
      `    </dsp:graphicFrame>` +
      `  </dsp:spTree>` +
      `</dsp:drawing>`

    const dataXml =
      `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
      `  <dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rDg1"/></dgm:ext></dgm:extLst>` +
      `</dgm:dataModel>`

    const pkg = await packageWithParts({
      'word/document.xml': '<w:document/>',
      'word/_rels/document.xml.rels': relsXml([
        { id: 'rDm1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData', target: 'diagrams/data1.xml' },
      ]),
      'word/diagrams/data1.xml': dataXml,
      'word/diagrams/_rels/data1.xml.rels': relsXml([
        { id: 'rDg1', type: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', target: 'drawing1.xml' },
      ]),
      'word/diagrams/drawing1.xml': drawingXml,
    })

    const theme = { colors: new Map(), fonts: new Map() }
    const drawingTheme: any = { colors: {}, palette: {}, colorMap: {}, fonts: { major: { supplemental: {} }, minor: { supplemental: {} } }, fillStyles: [], bgFillStyles: [], lineStyles: [], effectStyles: [], issues: [] }
    const graphicNode = {
      '@attrs': { uri: 'http://schemas.openxmlformats.org/drawingml/2006/diagram' },
      relIds: { '@attrs': { dm: 'rDm1' } },
    }

    const content = await prepareDrawingContent(pkg, graphicNode, 'word/document.xml', theme, drawingTheme)
    expect(content).toBeDefined()
    if (content?.kind === 'diagram') {
      expect(content.shapes.length).toBeLessThanOrEqual(10000)
    }
    const ctx = drawingPartContext(pkg)
    expect(ctx.diagnostics.some(d => d.kind === 'node-budget' && d.reason === 'source-node-limit')).toBe(true)
  })

  test('cached table resolves lnT and rowSpan geometry', async () => {
    const drawingXml =
      `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <dsp:spTree>` +
      `    <dsp:graphicFrame>` +
      `      <dsp:xfrm><a:off x="0" y="0"/><a:ext cx="400000" cy="600000"/></dsp:xfrm>` +
      `      <a:graphic>` +
      `        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
      `          <a:tbl>` +
      `            <a:tblGrid><a:gridCol w="400000"/></a:tblGrid>` +
      `            <a:tr h="300000">` +
      `              <a:tc rowSpan="2"><a:tcPr><a:lnT w="25400"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:lnT></a:tcPr></a:tc>` +
      `            </a:tr>` +
      `            <a:tr h="300000"/>` +
      `          </a:tbl>` +
      `        </a:graphicData>` +
      `      </a:graphic>` +
      `    </dsp:graphicFrame>` +
      `  </dsp:spTree>` +
      `</dsp:drawing>`

    const dataXml =
      `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
      `  <dgm:extLst><dgm:ext><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rDg1"/></dgm:ext></dgm:extLst>` +
      `</dgm:dataModel>`

    const pkg = await packageWithParts({
      'word/document.xml': '<w:document/>',
      'word/_rels/document.xml.rels': relsXml([
        { id: 'rDm1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData', target: 'diagrams/data1.xml' },
      ]),
      'word/diagrams/data1.xml': dataXml,
      'word/diagrams/_rels/data1.xml.rels': relsXml([
        { id: 'rDg1', type: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', target: 'drawing1.xml' },
      ]),
      'word/diagrams/drawing1.xml': drawingXml,
    })

    const theme = { colors: new Map(), fonts: new Map() }
    const drawingTheme: any = { colors: {}, palette: {}, colorMap: {}, fonts: { major: { supplemental: {} }, minor: { supplemental: {} } }, fillStyles: [], bgFillStyles: [], lineStyles: [], effectStyles: [], issues: [] }
    const graphicNode = {
      '@attrs': { uri: 'http://schemas.openxmlformats.org/drawingml/2006/diagram' },
      relIds: { '@attrs': { dm: 'rDm1' } },
    }

    const content = await prepareDrawingContent(pkg, graphicNode, 'word/document.xml', theme, drawingTheme)
    expect(content).toBeDefined()
    if (content?.kind === 'diagram') {
      expect(content.shapes.length).toBe(1)
      const cell = content.shapes[0]
      expect(cell.heightEmu).toBe(600000)
      expect(cell.line?.color).toBe('112233')
      expect(cell.line?.widthEmu).toBe(25400)
    }
  })
})
