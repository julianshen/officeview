import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { RECORD_TEXT } from '../src/core/text-recording'
import { buildTextIndex, findMatches } from '../src/core/search'
import { textForRange } from '../src/core/selection'

async function docxPackageWithPict(inner: string) {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body><w:p><w:r><w:pict>${inner}</w:pict></w:r></w:p></w:body></w:document>`)
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
}

async function xlsxPackageWithVml(vml: string) {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`)
  zip.file('xl/workbook.xml', `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`)
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`)
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><legacyDrawing r:id="rIdVml"/></worksheet>`)
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/v1.vml"/></Relationships>`)
  zip.file('xl/drawings/v1.vml', vml)
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
}

function collectRecords(painter: (ctx: any) => void) {
  const ctx = createCanvas(750, 500).getContext('2d') as any
  const items: any[] = []
  const seen = new Map<object, number>()
  ctx[RECORD_TEXT] = (text: any, _x: any, _y: any, _w: any, logical: any) => {
    if (!seen.has(logical.source)) seen.set(logical.source, seen.size)
    items.push({ text, sourceId: seen.get(logical.source), source: logical.source.text, start: logical.start, end: logical.end })
  }
  painter(ctx)
  return items
}

describe('VML proposal adapters: grouped paint/search/copy', () => {
  const GROUP_INNER = `<v:group id="group" style="position:absolute;left:72pt;top:36pt;width:200pt;height:100pt;" coordorigin="100,200" coordsize="1000,2000"><v:shape id="Grouped" style="position:absolute;left:150;top:400;width:400;height:600;"><v:textpath on="t" style="font-family:Arial;font-size:24pt;" string="Grouped"/></v:shape></v:group>`
  test('DOCX grouped retains hierarchy, paints, records once, indexes', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const pkg = await docxPackageWithPict(GROUP_INNER)
    const doc = await parseDocx(pkg)
    const img = doc.sections[0].paragraphs[0].images[0]
    expect(img?.drawing?.kind).toBe('diagram')
    const shapes = (img?.drawing as any).shapes
    expect(shapes.length).toBe(1)
    expect(shapes[0].group).toBeDefined()
    expect(shapes[0].children?.length).toBe(1)
    expect(shapes[0].children[0].textBody.paragraphs[0].runs[0].text).toBe('Grouped')
    // Visible paint
    const canvas = createCanvas(750, 500)
    const ctx = canvas.getContext('2d') as any
    const fills: string[] = []
    const orig = ctx.fillText.bind(ctx)
    ctx.fillText = (t: string, x: number, y: number) => { fills.push(String(t)); return orig(t, x, y) }
    paintDrawing(img.drawing!, ctx, 750, 500)
    expect(fills).toContain('Grouped')
    // Logical records once-only
    const recs = collectRecords(c => paintDrawing(img.drawing!, c, 750, 500))
    expect(recs.length).toBeGreaterThan(0)
    const bySource = new Map<number, any[]>()
    for (const r of recs) { const g = bySource.get(r.sourceId) ?? []; g.push(r); bySource.set(r.sourceId, g) }
    expect([...bySource.values()].map(g => g[0].source)).toEqual(['Grouped'])
    // Public index/copy
    const index = await buildTextIndex([{ spec: { widthPx: 750, heightPx: 500 }, paint: (c: any) => paintDrawing(img.drawing!, c, 750, 500) }])
    expect(findMatches(index, 'Grouped').length).toBe(1)
    const lines = index.pages[0].lines
    const copy = textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: lines.length - 1, charIndex: lines.at(-1)!.text.length } })
    expect(copy).toBe('Grouped')
  })

  test('XLSX grouped retains hierarchy and paints', async () => {
    const { parseXlsx } = await import('../src/xlsx/parse')
    const { renderSheet } = await import('../src/xlsx/render')
    const vml = `<xml xmlns:v="urn:schemas-microsoft-com:vml">${GROUP_INNER}</xml>`
    const pkg = await xlsxPackageWithVml(vml)
    const wb = await parseXlsx(pkg)
    const sheet = wb.sheets[0]
    expect(sheet.drawings?.length).toBeGreaterThan(0)
    const top = sheet.drawings![0] as any
    expect(top.group).toBeDefined()
    expect(top.children?.length).toBe(1)
    const canvas = createCanvas(750, 500)
    const ctx = canvas.getContext('2d') as any
    const fills: string[] = []
    const orig = ctx.fillText.bind(ctx)
    ctx.fillText = (t: string, x: number, y: number) => { fills.push(String(t)); return orig(t, x, y) }
    renderSheet(sheet, ctx)
    expect(fills).toContain('Grouped')
    const recs = collectRecords(c => renderSheet(sheet, c))
    expect(recs.map(r => r.source)).toContain('Grouped')
  })

  test('DOCX nested-group hierarchy', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const inner = `<v:group id="group" style="position:absolute;left:72pt;top:36pt;width:200pt;height:100pt;" coordorigin="100,200" coordsize="1000,2000"><v:group id="inner" style="position:absolute;left:100;top:200;width:500;height:1000;" coordorigin="-100,-200" coordsize="500,1000"><v:shape id="Nested" style="position:absolute;left:0;top:0;width:100;height:200;"><v:textpath on="t" style="font-family:Arial;font-size:24pt;" string="Nested"/></v:shape></v:group></v:group>`
    const pkg = await docxPackageWithPict(inner)
    const doc = await parseDocx(pkg)
    const shapes = (doc.sections[0].paragraphs[0].images[0].drawing as any).shapes
    expect(shapes[0].group).toBeDefined()
    expect(shapes[0].children[0].group).toBeDefined()
    expect(shapes[0].children[0].children[0].textBody.paragraphs[0].runs[0].text).toBe('Nested')
  })

  test('DOCX mixed-source-order', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const inner = `<v:shape id="Before" style="position:absolute;left:0pt;top:0pt;width:160pt;height:40pt;"><v:textpath on="t" style="font-family:Arial;font-size:24pt;" string="Before"/></v:shape><v:group id="group" style="position:absolute;left:72pt;top:36pt;width:200pt;height:100pt;" coordorigin="100,200" coordsize="1000,2000"><v:shape id="Grouped" style="position:absolute;left:150;top:400;width:400;height:600;"><v:textpath on="t" string="Grouped"/></v:shape><v:group id="inner" style="position:absolute;left:100;top:200;width:500;height:1000;" coordorigin="-100,-200" coordsize="500,1000"><v:shape id="Nested" style="position:absolute;left:0;top:0;width:100;height:200;"><v:textpath on="t" string="Nested"/></v:shape></v:group></v:group><v:shape id="After" style="position:absolute;left:0pt;top:0pt;width:160pt;height:40pt;"><v:textpath on="t" string="After"/></v:shape>`
    const pkg = await docxPackageWithPict(inner)
    const doc = await parseDocx(pkg)
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing!
    const recs = collectRecords(c => paintDrawing(drawing, c, 750, 500))
    const sources = [...new Map(recs.map(r => [r.sourceId, r])).values()].map((r: any) => r.source)
    expect(sources).toEqual(['Before', 'Grouped', 'Nested', 'After'])
  })
})

describe('VML proposal adapters: template/typography/diagnostics', () => {
  test('partial-template-override', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const inner = `<v:shapetype id="st" coordsize="1000,1000" path="m0,500l1000,500e" fillcolor="#FF0000" strokecolor="#0000FF" strokeweight="2pt"><v:textpath on="t" style="font-family:Arial;font-size:36pt;font-weight:bold;font-style:italic;v-text-align:center"/></v:shapetype><v:shape id="Override" style="position:absolute;left:0pt;top:0pt;width:160pt;height:40pt;" type="#st" fillcolor="#00FF00" stroked="f"><v:textpath style="font-size:18pt;font-weight:normal;font-style:normal" string="Override"/></v:shape>`
    const pkg = await docxPackageWithPict(inner)
    const doc = await parseDocx(pkg)
    const shape = (doc.sections[0].paragraphs[0].images[0].drawing as any).shapes[0]
    const run = shape.textBody.paragraphs[0].runs[0]
    expect(run.text).toBe('Override')
    expect(run.fontFamily).toBe('Arial')
    expect(run.fontSizePt).toBe(18)
    expect(run.bold ?? false).toBe(false)
    expect(run.italic ?? false).toBe(false)
    expect(run.color).toBe('#00FF00')
    expect(run.textOutline).toBeUndefined()
    expect(shape.textBody.paragraphs[0].align).toBe('center')
  })

  test('group-local-template', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const inner = `<v:group id="group" style="position:absolute;left:72pt;top:36pt;width:200pt;height:100pt;" coordorigin="100,200" coordsize="1000,2000"><v:shapetype id="local" coordsize="1000,1000" path="m0,500l1000,500e" fillcolor="#FF0000" strokecolor="#0000FF" strokeweight="2pt"><v:textpath on="t" style="font-family:Arial;font-size:36pt;font-weight:bold;font-style:italic;v-text-align:center"/></v:shapetype><v:shape id="Local" style="position:absolute;left:150;top:400;width:400;height:600;" type="#local"><v:textpath style="font-size:18pt" string="Local"/></v:shape></v:group>`
    const pkg = await docxPackageWithPict(inner)
    const doc = await parseDocx(pkg)
    const group = (doc.sections[0].paragraphs[0].images[0].drawing as any).shapes[0]
    expect(group.group).toBeDefined()
    const run = group.children[0].textBody.paragraphs[0].runs[0]
    expect(run.fontFamily).toBe('Arial')
    expect(run.fontSizePt).toBe(18)
    expect(run.bold).toBe(true)
    expect(run.italic).toBe(true)
    expect(run.color).toBe('#FF0000')
    expect(run.textOutline?.color).toBe('#0000FF')
  })

  test('inherited-enabled-override suppresses Hidden only', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const inner = `<v:shapetype id="off" coordsize="1000,1000" path="m0,500l1000,500e" fillcolor="#FF0000" strokecolor="#0000FF" strokeweight="2pt"><v:textpath on="f" style="font-family:Arial;font-size:36pt;font-weight:bold;font-style:italic;v-text-align:center"/></v:shapetype><v:shape id="Hidden" style="position:absolute;left:0pt;top:0pt;width:160pt;height:40pt;" type="#off"><v:textpath style="" string="Hidden"/></v:shape><v:shape id="Visible" style="position:absolute;left:0pt;top:0pt;width:160pt;height:40pt;" type="#off"><v:textpath on="t" style="" string="Visible"/></v:shape>`
    const pkg = await docxPackageWithPict(inner)
    const doc = await parseDocx(pkg)
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing!
    const recs = collectRecords(c => paintDrawing(drawing, c, 750, 500))
    const sources = [...new Set(recs.map(r => r.source))]
    expect(sources).toEqual(['Visible'])
  })

  test('DOCX attributed-gradient diagnostic once with correct partPath', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const inner = `<v:shape id="Gradient" style="position:absolute;left:0pt;top:0pt;width:160pt;height:40pt;"><v:fill type="gradient" color="#FF0000" color2="#0000FF"/><v:textpath on="t" style="" string="Gradient"/></v:shape>`
    const pkg = await docxPackageWithPict(inner)
    const doc = await parseDocx(pkg)
    const grad = (doc.drawingCoverage ?? []).filter(e => e.id === 'Gradient' || e.treePath === 'Gradient')
    expect(grad.length).toBeGreaterThanOrEqual(1)
    const count = (doc.drawingCoverage ?? []).filter(e => (e.id === 'Gradient') && e.feature === 'vml-wordart').length
    // Exactly once for the actual shape (diagnostic scope) – allow placement + diagnostic? Check diagnostics via context:
    // Our diagnostics become scope diagnostic entries; ensure at least one and not duplicated per template.
    expect(count).toBe(1)
    expect(grad[0].partPath).toBe('word/document.xml')
  })

  test('XLSX attributed-gradient diagnostic', async () => {
    const { parseXlsx } = await import('../src/xlsx/parse')
    const vml = `<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="Gradient" style="position:absolute;left:0pt;top:0pt;width:160pt;height:40pt;"><v:fill type="gradient" color="#FF0000" color2="#0000FF"/><v:textpath on="t" string="Gradient"/></v:shape></xml>`
    const pkg = await xlsxPackageWithVml(vml)
    const wb = await parseXlsx(pkg)
    const diags = wb.sheets[0].drawingDiagnostics ?? []
    const grad = diags.filter(d => d.identity === 'Gradient' && d.kind === 'unsupported-fill')
    expect(grad.length).toBe(1)
    expect(grad[0].partPath).toBe('xl/drawings/v1.vml')
  })

  test('local-transform-effects rotation/flip/shadow/dash', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const inner = `<v:shape id="Effects" style="position:absolute;left:40pt;top:30pt;width:100pt;height:20pt;rotation:90;flip:x;"><v:shadow on="t" color="#00FF00" offset="10pt,5pt"/><v:stroke on="t" color="#0000FF" weight="2pt" dashstyle="dash"/><v:textpath on="t" style="font-family:Arial;font-size:24pt;" string="Effects"/></v:shape>`
    const pkg = await docxPackageWithPict(inner)
    const doc = await parseDocx(pkg)
    const shape = (doc.sections[0].paragraphs[0].images[0].drawing as any).shapes[0]
    expect(shape.rotationDeg).toBe(90)
    expect(shape.flipH).toBe(true)
    const run = shape.textBody.paragraphs[0].runs[0]
    // Shadow supported as textShadow
    expect(run.textShadow).toBeDefined()
    // Dash fallback: solid outline retained + diagnostic
    expect(run.textOutline?.color).toBe('#0000FF')
    const diags = (doc.drawingCoverage ?? []).filter(e => e.id === 'Effects')
    expect(diags.some(e => e.feature === 'vml-dash')).toBe(true)
  })

  test('path-fitting degradation retains text with diagnostic', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const inner = `<v:shapetype id="curve" coordsize="1000,1000" adj="250" path="m0@0c333,0,666,0,1000,@0e"><v:formulas><v:f eqn="val #0"/></v:formulas><v:textpath on="t" fitpath="t" xscale="f"/></v:shapetype><v:shape id="Fitted" type="#curve" style="position:absolute;width:160pt;height:40pt;"><v:textpath string="Fitted" style="font-family:Arial;font-size:24pt"/></v:shape>`
    const pkg = await docxPackageWithPict(inner)
    const doc = await parseDocx(pkg)
    const shape = (doc.sections[0].paragraphs[0].images[0].drawing as any).shapes[0]
    expect(shape.textBody.paragraphs[0].runs[0].text).toBe('Fitted')
    const diags = (doc.drawingCoverage ?? []).filter(e => e.id === 'Fitted')
    expect(diags.some(e => e.feature === 'vml-path')).toBe(true)
  })

  test('duplicate-independent-shapes control', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const inner = `<v:shape id="Repeat1" style="position:absolute;left:0pt;top:0pt;width:160pt;height:40pt;"><v:textpath on="t" style="font-family:Arial;font-size:24pt;" string="Repeat"/></v:shape><v:shape id="Repeat2" style="position:absolute;left:0pt;top:60pt;width:160pt;height:40pt;"><v:textpath on="t" style="font-family:Arial;font-size:24pt;" string="Repeat"/></v:shape>`
    const pkg = await docxPackageWithPict(inner)
    const doc = await parseDocx(pkg)
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing!
    const index = await buildTextIndex([{ spec: { widthPx: 750, heightPx: 500 }, paint: (c: any) => paintDrawing(drawing, c, 750, 500) }])
    expect(findMatches(index, 'Repeat').length).toBe(2)
  })
})
