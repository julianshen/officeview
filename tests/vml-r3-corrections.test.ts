import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'

const NS = 'xmlns:v="urn:schemas-microsoft-com:vml"'
const shape = (id: string, text: string, extra = '', style = 'left:0pt;top:0pt;width:200pt;height:50pt;') =>
  `<v:shape ${id ? `id="${id}"` : ''} style="${style}" ${extra}><v:textpath on="t" string="${text}"/></v:shape>`

async function docxPkg(inner: string) {
  const zip = new JSZip()
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ${NS}><w:body><w:p><w:r><w:pict>${inner}</w:pict></w:r></w:p></w:body></w:document>`)
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
}


describe('R3 F1 derived EMU/pixel safety', () => {
  test('finite 1e307pt width keeps shape finite with diagnostic', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const inner = shape('Overflow', 'Overflow', '', 'left:0pt;top:0pt;width:1e307pt;height:50pt;')
    const doc = await parseDocx(await docxPkg(inner))
    const images = doc.sections[0].paragraphs[0].images
    expect(images.length).toBeGreaterThan(0)
    const shp = (images[0].drawing as any).shapes[0]
    expect(Number.isFinite(shp.widthEmu)).toBe(true)
    const ctx: any = createCanvas(600, 400).getContext('2d')
    const fills: string[] = []
    const orig = ctx.fillText.bind(ctx)
    ctx.fillText = (t: any, x: any, y: any) => { fills.push(String(t)); return orig(t, x, y) }
    paintDrawing(images[0].drawing!, ctx, 600, 400)
    expect(fills).toContain('Overflow')
    expect((doc.drawingCoverage ?? []).length).toBeGreaterThanOrEqual(1)
  })
  test('font 1e308pt never reaches fillText as NaN', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const inner = shape('Font', 'Font').replace('on="t" string', 'on="t" style="font-size:1e308pt" string')
    const doc = await parseDocx(await docxPkg(inner))
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing!
    const ctx: any = createCanvas(600, 400).getContext('2d')
    const args: any[] = []
    const orig = ctx.fillText.bind(ctx)
    ctx.fillText = (t: any, x: any, y: any) => { args.push([x, y]); return orig(t, x, y) }
    paintDrawing(drawing, ctx, 600, 400)
    for (const [x, y] of args) {
      expect(Number.isFinite(x)).toBe(true)
      expect(Number.isFinite(y)).toBe(true)
    }
    expect((doc.drawingCoverage ?? []).length).toBeGreaterThanOrEqual(1)
  })
})

describe('R3 F2 thick outline clamp', () => {
  test('filled=f + 100pt stroke keeps diagnosed outline, thin stays', async () => {
    const { parseXmlOrdered: ord } = await import('../src/core/xml')
    const { parseVmlWordArt } = await import('../src/drawing/vml')
    const root = ord(`<xml ${NS}>${shape('Outline', 'Outline', 'filled="f" strokecolor="#008800" strokeweight="100pt"')}</xml>`)
    const r = parseVmlWordArt(root as any)
    expect(r?.textBody.paragraphs[0].runs[0].noFill).toBe(true)
    expect(r?.textBody.paragraphs[0].runs[0].textOutline).toBeDefined()
    expect(r?.textBody.paragraphs[0].runs[0].textOutline?.widthPx).toBeLessThanOrEqual(100)
    expect((r?.diagnostics ?? []).length).toBeGreaterThanOrEqual(1)
    const thin = ord(`<xml ${NS}>${shape('Thin', 'Thin', 'strokecolor="#008800" strokeweight="0.2pt"')}</xml>`)
    const t2 = parseVmlWordArt(thin as any)
    expect(t2?.textBody.paragraphs[0].runs[0].textOutline?.widthPx).toBeCloseTo(0.2 * 96 / 72, 4)
  })
})

describe('R3 F3 invalid template + valid local width', () => {
  test('local 2pt wins, no stale invalid warning; zero clears', async () => {
    const { parseXmlOrdered: ord } = await import('../src/core/xml')
    const { parseVmlContainer } = await import('../src/drawing/vml')
    const xml = `<v:shapetype id="T" strokecolor="#008800" strokeweight="20bogus"><v:textpath on="t"/></v:shapetype>` + shape('Override', 'Override', 'type="#T" strokeweight="2pt"')
    const tree = parseVmlContainer(ord(`<xml ${NS}>${xml}</xml>`))
    const diags = tree.nodes.flatMap(n => n.kind === 'shape' ? (n.result.diagnostics ?? []) : [])
    expect(diags.some(d => d.feature === 'vml-stroke-width')).toBe(false)
    const leaf = tree.nodes[0]
    const wPx = leaf.kind === 'shape' ? leaf.result.textBody.paragraphs[0].runs[0].textOutline?.widthPx : undefined
    expect(wPx).toBeCloseTo(2 * 96 / 72, 4)
  })
})

describe('R3 F4 XLSX coverage keeps both VML parts', () => {
  test('a.vml and b.vml Same/shape[0] both covered', async () => {
    const { parseXlsx } = await import('../src/xlsx/parse')
    const inner = shape('Same', 'One').replace('<v:textpath', '<v:fill type="gradient"/><v:textpath')
    const z = new JSZip()
    z.file('xl/workbook.xml', '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="s"/></sheets></workbook>')
    z.file('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
    z.file('xl/worksheets/sheet1.xml', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><legacyDrawing r:id="a"/><legacyDrawing r:id="b"/></worksheet>')
    z.file('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships><Relationship Id="a" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/a.vml"/><Relationship Id="b" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/b.vml"/></Relationships>')
    for (const p of ['a', 'b']) z.file(`xl/drawings/${p}.vml`, `<xml ${NS}>${inner}</xml>`)
    const sheet = (await parseXlsx(await OfficePackage.load(await z.generateAsync({ type: 'uint8array' })))).sheets[0]
    expect(sheet.drawingDiagnostics?.filter(d => d.kind === 'unsupported-fill').length).toBe(2)
    const cov = (sheet.drawingCoverage ?? []).filter(e => e.element === 'unsupported-fill')
    expect(cov.filter(e => e.partPath === 'xl/drawings/a.vml').length).toBe(1)
    expect(cov.filter(e => e.partPath === 'xl/drawings/b.vml').length).toBe(1)
  })
})

describe('R3 F5 stable pict provenance', () => {
  test('same bytes fresh packages share pict path; repeats and concurrency stable', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const inner = shape('Gradient', 'Gradient').replace('<v:textpath', '<v:fill type="gradient"/><v:textpath')
    const zip = new JSZip()
    zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ${NS}><w:body><w:p><w:r><w:pict>${inner}</w:pict></w:r></w:p></w:body></w:document>`)
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    const paths: string[][] = []
    for (let i = 0; i < 3; i++) {
      const doc = await parseDocx(await OfficePackage.load(bytes))
      paths.push((doc.drawingCoverage ?? []).filter(d => d.scope === 'diagnostic').map(d => d.treePath))
    }
    expect(paths[0]).toEqual(paths[1])
    expect(paths[1]).toEqual(paths[2])
    const samePkg = await OfficePackage.load(bytes)
    const a = await parseDocx(samePkg)
    const pa = (a.drawingCoverage ?? []).filter(d => d.scope === 'diagnostic').map(d => d.treePath)
    expect(pa.every(p => p === 'pict[0]/shape[0]')).toBe(true)
    const b = await parseDocx(samePkg)
    const pb = (b.drawingCoverage ?? []).filter(d => d.scope === 'diagnostic').map(d => d.treePath)
    expect(pb.every(p => p === 'pict[0]/shape[0]')).toBe(true)
    const [c1, c2] = await Promise.all([OfficePackage.load(bytes).then(p => parseDocx(p)), OfficePackage.load(bytes).then(p => parseDocx(p))])
    const pc1 = (c1.drawingCoverage ?? []).filter(d => d.scope === 'diagnostic').map(d => d.treePath)
    const pc2 = (c2.drawingCoverage ?? []).filter(d => d.scope === 'diagnostic').map(d => d.treePath)
    expect(pc1).toEqual(pc2)
  })
})
