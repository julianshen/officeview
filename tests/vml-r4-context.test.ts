import { describe, expect, test, vi } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'

const NS = 'xmlns:v="urn:schemas-microsoft-com:vml"'
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
const shape = (id: string, text: string, extra = '') =>
  `<v:shape ${id ? `id="${id}"` : ''} style="left:0pt;top:0pt;width:200pt;height:50pt;" ${extra}><v:textpath on="t" string="${text}"/></v:shape>`
const grad = (id: string, text: string) =>
  shape(id, text).replace('<v:textpath', '<v:fill type="gradient"/><v:textpath')
const pict = (inner: string) => `<w:p><w:r><w:pict>${inner}</w:pict></w:r></w:p>`

async function pack(body: string) {
  const z = new JSZip()
  z.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ${NS} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="${WPS}"><w:body>${body}</w:body></w:document>`)
  const bytes = await z.generateAsync({ type: 'uint8array' })
  return { bytes, pkg: await OfficePackage.load(bytes) }
}

describe('R4 F1 shared provenance across table clones', () => {
  test('2 cells + body pict get 3 distinct paths', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const cell = pict(grad('Same', 'One'))
    const table = `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${cell}</w:tc><w:tc>${cell}</w:tc></w:tr></w:tbl>`
    const { pkg } = await pack(table + pict(grad('Same', 'One')))
    const doc = await parseDocx(pkg)
    const diags = (doc.drawingCoverage ?? []).filter(d => d.scope === 'diagnostic')
    expect(diags.length).toBe(3)
    expect(new Set(diags.map(d => `${d.partPath}:${d.treePath}`)).size).toBe(3)
  })
})

describe('R4 F2 same-package reparse stays once per source', () => {
  test('1 pict parsed twice keeps count 1; 2 picts stays 2', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { pkg } = await pack(pict(grad('Same', 'One')))
    const first = await parseDocx(pkg)
    const n1 = (first.drawingCoverage ?? []).filter(d => d.scope === 'diagnostic').length
    expect(n1).toBe(1)
    const second = await parseDocx(pkg)
    const n2 = (second.drawingCoverage ?? []).filter(d => d.scope === 'diagnostic').length
    expect(n2).toBe(1)
    const two = pict(grad('Same', 'One')) + pict(grad('Same', 'Two'))
    const { pkg: pkg2 } = await pack(two)
    const r1 = await parseDocx(pkg2)
    expect((r1.drawingCoverage ?? []).filter(d => d.scope === 'diagnostic').length).toBe(2)
    const r2 = await parseDocx(pkg2)
    expect((r2.drawingCoverage ?? []).filter(d => d.scope === 'diagnostic').length).toBe(2)
  })
})

describe('R4 F3 nested textbox pict warns once with owner path', () => {
  test('inner Nested + outer AfterTextbox each warn distinctly', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const inner = pict(grad('', 'Nested'))
    const textbox = `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="3810000" cy="1714500"/><wp:docPr id="10" name="Review"/><a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1714500"/></a:xfrm></wps:spPr><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    const { pkg } = await pack(textbox + pict(grad('', 'AfterTextbox')))
    const doc = await parseDocx(pkg)
    const diags = (doc.drawingCoverage ?? []).filter(d => d.scope === 'diagnostic')
    expect(diags.length).toBe(2)
    expect(diags.every(d => d.partPath === 'word/document.xml')).toBe(true)
    expect(new Set(diags.map(d => d.treePath)).size).toBe(2)
  })
})

describe('R4 FYI inactive width suppressed when stroke off', () => {
  test('template bogus + local stroked=f emits no width warning', async () => {
    const { parseXmlOrdered: ord } = await import('../src/core/xml')
    const { parseVmlContainer } = await import('../src/drawing/vml')
    const xml = `<v:shapetype id="T" strokecolor="#008800" strokeweight="20bogus"><v:textpath on="t"/></v:shapetype>` + shape('Off', 'Off', 'type="#T" stroked="f"')
    const tree = parseVmlContainer(ord(`<xml ${NS}>${xml}</xml>`))
    const diags = tree.nodes.flatMap(n => n.kind === 'shape' ? (n.result.diagnostics ?? []) : [])
    expect(diags.some(d => d.feature === 'vml-stroke-width')).toBe(false)
  })
})

describe('V1 container-parse error handling', () => {
  test('DOCX: parseVmlContainer failure emits malformed-vml-container diagnostic', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const vmlModule = await import('../src/drawing/vml')
    const spy = vi.spyOn(vmlModule, 'parseVmlContainer').mockImplementationOnce(() => {
      throw new Error('corrupted VML container')
    })
    const { pkg } = await pack(pict('<v:shape id="Broken"/>'))
    const doc = await parseDocx(pkg)
    spy.mockRestore()
    const diags = (doc.drawingCoverage ?? []).filter(d => d.scope === 'diagnostic')
    expect(diags.some(d => d.element === 'malformed-vml-container' && d.reason?.includes('corrupted VML container'))).toBe(true)
  })

  test('XLSX: parseVmlContainer failure emits malformed-vml-container diagnostic', async () => {
    const { parseWorksheetDrawings } = await import('../src/xlsx/drawing')
    const vmlModule = await import('../src/drawing/vml')
    const spy = vi.spyOn(vmlModule, 'parseVmlContainer').mockImplementationOnce(() => {
      throw new Error('corrupted XLSX VML container')
    })
    const z = new JSZip()
    z.file('xl/drawings/vmlDrawing1.vml', '<xml><v:shape/></xml>')
    z.file('xl/worksheets/_rels/sheet1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>`)
    const bytes = await z.generateAsync({ type: 'uint8array' })
    const pkg = await OfficePackage.load(bytes)
    const { parseXmlOrdered: ord } = await import('../src/core/xml')
    const root = ord(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><legacyDrawing r:id="rId1"/></worksheet>`)
    const sheet: any = {}
    const theme: any = { colors: {}, palette: {}, colorMap: {}, fonts: { major: {}, minor: {} }, fillStyles: [], bgFillStyles: [], lineStyles: [], effectStyles: [], issues: [] }
    await parseWorksheetDrawings(pkg, sheet, 'xl/worksheets/sheet1.xml', root, theme)
    spy.mockRestore()
    expect(sheet.drawingDiagnostics?.some((d: any) => d.kind === 'malformed-vml-container' && d.message?.includes('corrupted XLSX VML container'))).toBe(true)
  })
})

describe('V2 non-finite EMU shapes emit unsupported-geometry diagnostic', () => {
  test('DOCX: shape with non-finite bounds emits unsupported-geometry diagnostic with shape identity', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const vmlModule = await import('../src/drawing/vml')
    const original = vmlModule.parseVmlContainer
    const spy = vi.spyOn(vmlModule, 'parseVmlContainer').mockImplementationOnce((node) => {
      const res = original(node)
      if (res.nodes.length > 0 && res.nodes[0].kind === 'shape') {
        res.nodes[0].result.widthPt = NaN
      }
      return res
    })
    const { pkg } = await pack(pict('<v:shape id="NonFiniteShape" style="left:0pt;top:0pt;width:200pt;height:50pt;"><v:textpath on="t" string="Bad"/></v:shape>'))
    const doc = await parseDocx(pkg)
    spy.mockRestore()
    const diags = (doc.drawingCoverage ?? []).filter(d => d.scope === 'diagnostic')
    expect(diags.some(d => d.element === 'unsupported-geometry' && d.id === 'NonFiniteShape')).toBe(true)
  })
})
