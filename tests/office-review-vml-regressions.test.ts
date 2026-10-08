import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'

const XML_WRAPPER = (bodyContent: string) =>
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
  `xmlns:v="urn:schemas-microsoft-com:vml" ` +
  `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ` +
  `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
  `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
  `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
  `<w:body>${bodyContent}</w:body></w:document>`


async function loadDoc(xml: string) {
  const z = new JSZip()
  z.file('word/document.xml', xml)
  const bytes = await z.generateAsync({ type: 'uint8array' })
  const pkg = await OfficePackage.load(bytes)
  return parseDocx(pkg)
}

describe('U1: empty VML groups must not suppress usable fallbacks', () => {
  it('U1.1: canonical paragraph MC control - empty group choice yields fallback text and zero images', async () => {
    const group = '<v:group style="width:200pt;height:100pt"><v:rect style="left:0;top:0;width:1000;height:1000" fillcolor="#FF0000"/></v:group>'
    const xml = XML_WRAPPER(
      `<w:p><mc:AlternateContent>` +
      `<mc:Choice Requires="w"><w:r><w:pict>${group}</w:pict></w:r></mc:Choice>` +
      `<mc:Fallback><w:r><w:t>Fallback Text</w:t></w:r></mc:Fallback>` +
      `</mc:AlternateContent></w:p>`
    )
    const doc = await loadDoc(xml)
    const p = doc.sections[0].paragraphs[0]
    expect(p.runs.map(r => r.text)).toEqual(['Fallback Text'])
    expect(p.images).toEqual([])
    const diagnostics = (doc.drawingCoverage ?? []).filter(e => e.scope === 'diagnostic')
    expect(diagnostics).toEqual([])
  })

  it('U1.2: nested empty VML groups do not claim usable pict and fallback is retained', async () => {
    const nestedGroup = '<v:group style="width:200pt;height:100pt"><v:group style="width:100pt;height:50pt"><v:rect style="left:0;top:0;width:500;height:500"/></v:group></v:group>'
    const xml = XML_WRAPPER(
      `<w:p><mc:AlternateContent>` +
      `<mc:Choice Requires="w"><w:r><w:pict>${nestedGroup}</w:pict></w:r></mc:Choice>` +
      `<mc:Fallback><w:r><w:t>Nested Fallback</w:t></w:r></mc:Fallback>` +
      `</mc:AlternateContent></w:p>`
    )
    const doc = await loadDoc(xml)
    const p = doc.sections[0].paragraphs[0]
    expect(p.runs.map(r => r.text)).toEqual(['Nested Fallback'])
    expect(p.images).toEqual([])
  })

  it('U1.3: VML group with supported-leaf sibling IS admitted and selects Choice', async () => {
    const groupWithLeaf =
      '<v:group style="width:200pt;height:100pt">' +
      '<v:rect style="left:0;top:0;width:1000;height:1000" fillcolor="#FF0000"/>' +
      '<v:shape id="Sibling" style="width:200pt;height:50pt"><v:textpath on="t" string="Supported Sibling"/></v:shape>' +
      '</v:group>'
    const xml = XML_WRAPPER(
      `<w:p><mc:AlternateContent>` +
      `<mc:Choice Requires="w"><w:r><w:pict>${groupWithLeaf}</w:pict></w:r></mc:Choice>` +
      `<mc:Fallback><w:r><w:t>Fallback</w:t></w:r></mc:Fallback>` +
      `</mc:AlternateContent></w:p>`
    )
    const doc = await loadDoc(xml)
    const p = doc.sections[0].paragraphs[0]
    expect(p.runs.map(r => r.text)).toEqual([])
    expect(p.images.length).toBe(1)
    const drawing = p.images[0].drawing
    expect(drawing?.kind).toBe('diagram')
  })

  it('U1.4: accepted empty choice blank semantics are preserved', async () => {
    const xml = XML_WRAPPER(
      `<w:p><mc:AlternateContent>` +
      `<mc:Choice Requires="w"/>` +
      `<mc:Fallback><w:r><w:t>Fallback</w:t></w:r></mc:Fallback>` +
      `</mc:AlternateContent></w:p>`
    )
    const doc = await loadDoc(xml)
    const p = doc.sections[0].paragraphs[0]
    expect(p.runs).toEqual([])
    expect(p.images).toEqual([])
    const coverage = (doc.drawingCoverage ?? []).find(e => e.feature === 'empty-choice')
    expect(coverage?.selectedRepresentation).toBe('blank')
  })
})

describe('U5: MC dry probing transactional provenance and diagnostics', () => {
  it('U5.1: selected Choice with VML diagnostic emits warning exactly once with stable pict[0]', async () => {
    const grad = '<v:shape id="S" style="width:200pt;height:50pt"><v:fill type="gradient"/><v:textpath on="t" string="S"/></v:shape>'
    const xml = XML_WRAPPER(
      `<w:p><mc:AlternateContent>` +
      `<mc:Choice Requires="w"><w:r><w:pict>${grad}</w:pict></w:r></mc:Choice>` +
      `<mc:Fallback><w:r><w:t>Fallback</w:t></w:r></mc:Fallback>` +
      `</mc:AlternateContent></w:p>`
    )
    const doc = await loadDoc(xml)
    const p = doc.sections[0].paragraphs[0]
    expect(p.images.length).toBe(1)
    const diags = (doc.drawingCoverage ?? []).filter(e => e.scope === 'diagnostic')
    expect(diags.map(d => d.treePath)).toEqual(['pict[0]/shape[0]'])
  })

  it('U5.2: skipped Choice rolls back provenance so subsequent pict gets sequence pict[0]', async () => {
    const emptyGroup = '<v:group style="width:200pt;height:100pt"><v:rect style="left:0;top:0;width:1000;height:1000"/></v:group>'
    const grad = '<v:shape id="S" style="width:200pt;height:50pt"><v:fill type="gradient"/><v:textpath on="t" string="S"/></v:shape>'
    const xml = XML_WRAPPER(
      `<w:p><mc:AlternateContent>` +
      `<mc:Choice Requires="w"><w:r><w:pict>${emptyGroup}</w:pict></w:r></mc:Choice>` +
      `<mc:Fallback><w:r><w:t>Fallback</w:t></w:r></mc:Fallback>` +
      `</mc:AlternateContent></w:p>` +
      `<w:p><w:r><w:pict>${grad}</w:pict></w:r></w:p>`
    )
    const doc = await loadDoc(xml)
    const diags = (doc.drawingCoverage ?? []).filter(e => e.scope === 'diagnostic')
    expect(diags.map(d => d.treePath)).toEqual(['pict[0]/shape[0]'])
  })

  it('U5.3: direct-context control preserves provenance identity and restores prior undefined state', async () => {
    const { parseParagraph } = await import('../src/docx/parse')
    const { parseXmlOrdered } = await import('../src/core/xml')
    const emptyGroup = '<v:group style="width:200pt;height:100pt"><v:rect style="left:0;top:0;width:1000;height:1000"/></v:group>'
    const pXml = parseXmlOrdered(
      `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:v="urn:schemas-microsoft-com:vml" ` +
      `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
      `<mc:AlternateContent>` +
      `<mc:Choice Requires="w"><w:r><w:pict>${emptyGroup}</w:pict></w:r></mc:Choice>` +
      `<mc:Fallback><w:r><w:t>Fallback</w:t></w:r></mc:Fallback>` +
      `</mc:AlternateContent></w:p>`
    )

    // Case A: Context with no preinitialized vmlProvenance
    const ctxWithoutProv: any = {
      styles: { defaultFont: 'Calibri', styles: new Map() },
      drawings: new Map(),
      reserveDrawing: () => true,
    }
    const paraA = parseParagraph(pXml, [], undefined, ctxWithoutProv)
    expect(paraA.runs.map(r => r.text)).toEqual(['Fallback'])
    expect(ctxWithoutProv.vmlProvenance).toBeUndefined()

    // Case B: Context with preinitialized vmlProvenance object
    const sharedProvenance = { nextPict: 5 }
    const ctxWithProv: any = {
      styles: { defaultFont: 'Calibri', styles: new Map() },
      drawings: new Map(),
      reserveDrawing: () => true,
      vmlProvenance: sharedProvenance,
    }
    const paraB = parseParagraph(pXml, [], undefined, ctxWithProv)
    expect(paraB.runs.map(r => r.text)).toEqual(['Fallback'])
    expect(ctxWithProv.vmlProvenance).toBe(sharedProvenance)
    expect(ctxWithProv.vmlProvenance.nextPict).toBe(5)
  })
})
