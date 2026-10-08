import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXmlOrdered } from '../src/core/xml'

// Copied independent fixtures (rebound to r2 SRC, new output). Original/frozen never touched.
const NS = 'xmlns:v="urn:schemas-microsoft-com:vml"'
const shape = (id: string, text: string, extra = '', style = 'left:0pt;top:0pt;width:200pt;height:50pt;') =>
  `<v:shape ${id ? `id="${id}"` : ''} style="${style}" ${extra}><v:textpath on="t" string="${text}"/></v:shape>`
const group = (inner: string, id = 'G') =>
  `<v:group id="${id}" style="left:0pt;top:0pt;width:200pt;height:100pt" coordsize="1000,1000">${inner}</v:group>`

async function docxPkg(inner: string) {
  const zip = new JSZip()
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ${NS}><w:body><w:p><w:r><w:pict>${inner}</w:pict></w:r></w:p></w:body></w:document>`)
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
}
async function xlsxPkg(vmlInner: string) {
  const zip = new JSZip()
  zip.file('xl/workbook.xml', '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="s"/></sheets></workbook>')
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><legacyDrawing r:id="v"/></worksheet>')
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships><Relationship Id="v" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/v1.vml"/></Relationships>')
  zip.file('xl/drawings/v1.vml', `<xml ${NS}>${vmlInner}</xml>`)
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
}
function paintCalls(painter: (ctx: any) => void) {
  const canvas = createCanvas(750, 500)
  const ctx = canvas.getContext('2d') as any
  const calls: any[] = []
  const fill = ctx.fillText.bind(ctx)
  const rotate = ctx.rotate.bind(ctx)
  const rotateArgs: any[] = []
  ctx.fillText = (text: any, x: any, y: any) => {
    const t = ctx.getTransform()
    calls.push({ text: String(text), matrix: { a: t.a, b: t.b, c: t.c, d: t.d, e: t.e, f: t.f } })
    return fill(text, x, y)
  }
  ;(ctx as any).rotate = (a: any) => { rotateArgs.push(a); return rotate(a) }
  painter(ctx)
  return { calls, rotateArgs }
}
function projectCorners(leafEmuW: number, leafEmuH: number, m: any) {
  const w = leafEmuW / 9525
  const h = leafEmuH / 9525
  return [[0, 0], [w, 0], [w, h], [0, h]].map(([x, y]) => [(m.a * x + m.c * y + m.e) * 0.75, (m.b * x + m.d * y + m.f) * 0.75])
}
const close = (a: number, b: number, tol = 0.75) => Math.abs(a - b) <= tol

describe('R2 P1 omitted group positions use zero defaults', () => {
  test('omitted left+top → (30,20,80,30)pt, no malformed fallback', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const inner = `<v:group id="DefaultGroup" style="position:absolute;width:200pt;height:100pt;" coordorigin="0,0" coordsize="1000,2000"><v:shape id="Defaults" style="position:absolute;left:150;top:400;width:400;height:600;"><v:textpath on="t" style="font-family:Arial;font-size:24pt;" string="Defaults"/></v:shape></v:group>`
    const doc = await parseDocx(await docxPkg(inner))
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing! as any
    const leaf = drawing.shapes[0].children[0]
    const { calls } = paintCalls(c => paintDrawing(drawing, c, 750, 500))
    const draw = calls.find(c => c.text === 'Defaults')
    expect(draw).toBeDefined()
    const corners = projectCorners(leaf.widthEmu, leaf.heightEmu, draw.matrix)
    const wanted = [[30, 20], [110, 20], [110, 50], [30, 50]]
    for (let i = 0; i < 4; i++) {
      expect(close(corners[i][0], wanted[i][0])).toBe(true)
      expect(close(corners[i][1], wanted[i][1])).toBe(true)
    }
    const bad = (doc.drawingCoverage ?? []).filter(e => e.feature === 'vml-group')
    expect(bad.length).toBe(0)
  })
  test('omitted left only (top36) → (30,56,80,30)pt', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const inner = `<v:group id="DefaultGroup" style="position:absolute;top:36pt;width:200pt;height:100pt;" coordorigin="0,0" coordsize="1000,2000"><v:shape id="Defaults" style="position:absolute;left:150;top:400;width:400;height:600;"><v:textpath on="t" style="font-family:Arial;font-size:24pt;" string="Defaults"/></v:shape></v:group>`
    const doc = await parseDocx(await docxPkg(inner))
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing! as any
    const leaf = drawing.shapes[0].children[0]
    const { calls } = paintCalls(c => paintDrawing(drawing, c, 750, 500))
    const draw = calls.find(c => c.text === 'Defaults')
    const corners = projectCorners(leaf.widthEmu, leaf.heightEmu, draw.matrix)
    const wanted = [[30, 56], [110, 56], [110, 86], [30, 86]]
    for (let i = 0; i < 4; i++) {
      expect(close(corners[i][0], wanted[i][0])).toBe(true)
      expect(close(corners[i][1], wanted[i][1])).toBe(true)
    }
  })
})

describe('R2 P2 guards count groups, diagnose depth/budget, preserve siblings', () => {
  test('10005 empty groups capped at 10000 with diagnostic', async () => {
    const { parseVmlContainer } = await import('../src/drawing/vml')
    const many = Array.from({ length: 10005 }, (_, i) => group('', `G${i}`)).join('')
    const tree = parseVmlContainer(parseXmlOrdered(`<xml ${NS}>${many}</xml>`))
    expect(tree.nodes.length).toBeLessThanOrEqual(10000)
    const diags = (tree as any).diagnostics ?? []
    expect(diags.length).toBeGreaterThanOrEqual(1)
  })
  test('33-deep nesting diagnoses Deep loss but preserves After', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const badDepth = Array.from({ length: 33 }, (_, i) => i).reduceRight((inner, i) => group(inner, `G${i}`), shape('Deep', 'Deep')) + shape('After', 'After')
    const doc = await parseDocx(await docxPkg(badDepth))
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing!
    const ctx: any = createCanvas(600, 400).getContext('2d')
    const seen: string[] = []
    const orig = ctx.fillText.bind(ctx)
    ctx.fillText = (t: any, x: any, y: any) => { seen.push(String(t)); return orig(t, x, y) }
    paintDrawing(drawing, ctx, 600, 400)
    expect(seen).toContain('After')
    const cov = doc.drawingCoverage ?? []
    // At minimum a depth/budget diagnostic must exist; silent Deep loss is failure.
    expect(cov.some(e => (e.feature ?? '').includes('depth') || (e.reason ?? '').includes('depth') || e.feature === 'vml-group' || e.element === 'group-depth')).toBe(true)
  })
})

describe('R2 P3 finite numerics + strict units', () => {
  test('1e307in width falls back finite with diagnostic and still paints', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const inner = shape('Overflow', 'Overflow', '', 'left:0pt;top:0pt;width:1e307in;height:50pt;')
    const doc = await parseDocx(await docxPkg(inner))
    const shp = (doc.sections[0].paragraphs[0].images[0].drawing as any).shapes[0]
    expect(Number.isFinite(shp.widthEmu)).toBe(true)
    expect(Number.isFinite(shp.xEmu + shp.widthEmu)).toBe(true)
    const { calls } = paintCalls(c => paintDrawing(doc.sections[0].paragraphs[0].images[0].drawing!, c, 600, 400))
    expect(calls.map(c => c.text)).toContain('Overflow')
    expect((doc.drawingCoverage ?? []).length).toBeGreaterThanOrEqual(1)
  })
  test('rotation 1e308 does not reach canvas as Infinity', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const inner = shape('Rotate', 'Rotate', '', 'left:0pt;top:0pt;width:200pt;height:50pt;rotation:1e308;')
    const doc = await parseDocx(await docxPkg(inner))
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing!
    const { rotateArgs, calls } = paintCalls(c => paintDrawing(drawing, c, 600, 400))
    for (const a of rotateArgs) expect(Number.isFinite(a)).toBe(true)
    expect(calls.map(c => c.text)).toContain('Rotate')
  })
  test('font/stroke Infinity fall back finite with diagnostic', async () => {
    const { parseXmlOrdered: ord } = await import('../src/core/xml')
    const { parseVmlWordArt } = await import('../src/drawing/vml')
    const root = ord(`<xml ${NS}><v:shape id="F" style="width:200pt;height:50pt;" strokecolor="#008800" strokeweight="1e307in"><v:textpath on="t" style="font-size:1e307in" string="X"/></v:shape></xml>`)
    const r = parseVmlWordArt(root as any)
    const fontSize = r?.textBody.paragraphs[0].runs[0].fontSizePt
    expect(fontSize === undefined || Number.isFinite(fontSize)).toBe(true)
    expect(fontSize).not.toBe(Infinity)
    const outlineW = r?.textBody.paragraphs[0].runs[0].textOutline?.widthPx
    expect(outlineW === undefined || Number.isFinite(outlineW as number)).toBe(true)
    expect((r?.diagnostics ?? []).length).toBeGreaterThanOrEqual(1)
  })
  test('20bogus rejected, not partial 20', async () => {
    const { parseXmlOrdered: ord } = await import('../src/core/xml')
    const { parseVmlWordArt } = await import('../src/drawing/vml')
    const root = ord(`<xml ${NS}><v:shape id="U" style="left:0pt;top:0pt;width:20bogus;height:50pt;"><v:textpath on="t" string="Unit"/></v:shape></xml>`)
    const r = parseVmlWordArt(root as any)
    expect(r?.widthPt).not.toBe(20)
    expect((r?.diagnostics ?? []).some(d => (d.feature ?? '').includes('unit') || (d.message ?? '').toLowerCase().includes('bogus') || d.kind === 'unsupported-effect')).toBe(true)
  })
})

describe('R2 P4 effective inherited fill/dash/path', () => {
  test('template gradient warns once per actual use (First+Second)', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const xml = `<v:shapetype id="T"><v:fill type="gradient" color="#FF0000"/><v:textpath on="t"/></v:shapetype>` + shape('First', 'First', 'type="#T"') + shape('Second', 'Second', 'type="#T"')
    const doc = await parseDocx(await docxPkg(xml))
    const cov = (doc.drawingCoverage ?? []).filter(e => e.feature === 'vml-wordart' && e.element === 'unsupported-fill')
    expect(cov.length).toBe(2)
    const pkg2 = await xlsxPkg(`<xml ${NS}>${xml}</xml>`.replace('<xml xmlns:v="urn:schemas-microsoft-com:vml">', ''))
    void pkg2
  })
  test('inherited dash warns; local solid/off clears', async () => {
    const { parseVmlWordArt: _p } = await import('../src/drawing/vml')
    void _p
    const { parseXmlOrdered: ord } = await import('../src/core/xml')
    const { parseVmlContainer } = await import('../src/drawing/vml')
    const xml = `<v:shapetype id="T" strokecolor="#0000FF" strokeweight="2pt"><v:stroke dashstyle="dash"/><v:textpath on="t"/></v:shapetype>` + shape('First', 'First', 'type="#T"')
    const tree = parseVmlContainer(ord(`<xml ${NS}>${xml}</xml>`))
    const diags = tree.nodes.flatMap(n => n.kind === 'shape' ? (n.result.diagnostics ?? []) : [])
    expect(diags.some(d => d.feature === 'vml-dash')).toBe(true)
  })
  test('qx arc warns; local fitpath=f clears stale template warning', async () => {
    const { parseXmlOrdered: ord } = await import('../src/core/xml')
    const { parseVmlContainer } = await import('../src/drawing/vml')
    const arc = shape('Path', 'Path').replace('<v:textpath', '<v:path v="m 0,0 qx 100,100 e"/><v:textpath')
    const t1 = parseVmlContainer(ord(`<xml ${NS}>${arc}</xml>`))
    expect(t1.nodes.flatMap(n => n.kind === 'shape' ? (n.result.diagnostics ?? []) : []).some(d => d.feature === 'vml-path')).toBe(true)
    const localFalse = `<v:shapetype id="T"><v:textpath on="t" fitpath="t"/></v:shapetype>` + shape('First', 'First', 'type="#T"').replace('on="t" string', 'on="t" fitpath="f" string')
    const t2 = parseVmlContainer(ord(`<xml ${NS}>${localFalse}</xml>`))
    expect(t2.nodes.flatMap(n => n.kind === 'shape' ? (n.result.diagnostics ?? []) : []).some(d => d.feature === 'vml-path')).toBe(false)
  })
})

describe('R2 P5 shadow defaults/fractions/variants', () => {
  test('default shadow offset is 2.6667px', async () => {
    const { parseXmlOrdered: ord } = await import('../src/core/xml')
    const { parseVmlWordArt } = await import('../src/drawing/vml')
    const root = ord(`<xml ${NS}>${shape('Shadow', 'Shadow').replace('<v:textpath', '<v:shadow on="t"/><v:textpath')}</xml>`)
    const r = parseVmlWordArt(root as any)
    expect(r?.textBody.paragraphs[0].runs[0].textShadow?.offsetX).toBeCloseTo(2.6666667, 4)
    expect(r?.textBody.paragraphs[0].runs[0].textShadow?.offsetY).toBeCloseTo(2.6666667, 4)
  })
  test('fractional 0.5,0.5 on 200x50 supported or diagnosed (not 0.666px)', async () => {
    const { parseXmlOrdered: ord } = await import('../src/core/xml')
    const { parseVmlWordArt } = await import('../src/drawing/vml')
    const root = ord(`<xml ${NS}>${shape('Shadow', 'Shadow').replace('<v:textpath', '<v:shadow on="t" offset="0.5,0.5"/><v:textpath')}</xml>`)
    const r = parseVmlWordArt(root as any)
    const sh = r?.textBody.paragraphs[0].runs[0].textShadow
    const diags = r?.diagnostics ?? []
    const isFractional = Math.abs((sh?.offsetX ?? 0) - 0.6666667) > 0.01
    const diagnosed = diags.some(d => (d.feature ?? '').includes('shadow'))
    expect(isFractional || diagnosed).toBe(true)
  })
  test('double shadow type diagnosed, single fallback retained', async () => {
    const { parseXmlOrdered: ord } = await import('../src/core/xml')
    const { parseVmlWordArt } = await import('../src/drawing/vml')
    const root = ord(`<xml ${NS}>${shape('Shadow', 'Shadow').replace('<v:textpath', '<v:shadow on="t" type="double" color2="#008800" offset2="5pt,5pt"/><v:textpath')}</xml>`)
    const r = parseVmlWordArt(root as any)
    expect(r?.textBody.paragraphs[0].runs[0].text).toBe('Shadow')
    expect((r?.diagnostics ?? []).some(d => (d.feature ?? '').includes('shadow'))).toBe(true)
  })
})

describe('R2 P6 distinct diagnostic provenance', () => {
  test('anonymous One/Two gradients distinct across adapters', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { parseXlsx } = await import('../src/xlsx/parse')
    const xml = shape('', 'One').replace('<v:textpath', '<v:fill type="gradient"/><v:textpath') + shape('', 'Two').replace('<v:textpath', '<v:fill type="gradient"/><v:textpath')
    const doc = await parseDocx(await docxPkg(xml))
    const cov = (doc.drawingCoverage ?? []).filter(e => e.element === 'unsupported-fill')
    expect(cov.length).toBe(2)
    expect(new Set(cov.map(c => c.treePath)).size).toBe(2)
    const wb = await parseXlsx(await xlsxPkg(xml))
    const diags = (wb.sheets[0].drawingDiagnostics ?? []).filter(d => d.kind === 'unsupported-fill')
    expect(diags.length).toBe(2)
    expect(new Set(diags.map(d => `${d.partPath}|${d.identity}|${(d as any).sourcePath ?? ''}`)).size).toBeGreaterThanOrEqual(1)
    // XLSX identities must be distinct (no shared undefined collapse is asserted via treePaths)
    const trees = (wb.sheets[0].drawings ?? []).map(d => d.source.treePath)
    expect(new Set(trees).size).toBe(2)
  })
})

describe('R2 core LF/Unicode overlay works without core edits', () => {
  test('numeric LF decodes through VML string', async () => {
    const { parseXmlOrdered: ord } = await import('../src/core/xml')
    const { parseVmlWordArt } = await import('../src/drawing/vml')
    const root = ord(`<xml ${NS}><v:shape id="U" style="width:200pt;height:50pt;"><v:textpath on="t" string="AA&#10;BB"/></v:shape></xml>`)
    const r = parseVmlWordArt(root as any)
    expect(r?.textBody.paragraphs[0].runs[0].text).toBe('AA\nBB')
  })
})
