import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'

async function docxPkg(inner: string) {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`)
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><w:body><w:p><w:r><w:pict>${inner}</w:pict></w:r></w:p></w:body></w:document>`)
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
}

function paintCalls(painter: (ctx: any) => void) {
  const canvas = createCanvas(750, 500)
  const ctx = canvas.getContext('2d') as any
  const calls: any[] = []
  const fill = ctx.fillText.bind(ctx)
  ctx.fillText = (text: any, x: any, y: any) => {
    const t = ctx.getTransform()
    calls.push({ text: String(text), matrix: { a: t.a, b: t.b, c: t.c, d: t.d, e: t.e, f: t.f } })
    return fill(text, x, y)
  }
  painter(ctx)
  return calls
}

function projectLeaf(leafEmuW: number, leafEmuH: number, matrix: any) {
  const w = leafEmuW / 9525
  const h = leafEmuH / 9525
  return [[0, 0], [w, 0], [w, h], [0, h]].map(([x, y]) => [(matrix.a * x + matrix.c * y + matrix.e) * 0.75, (matrix.b * x + matrix.d * y + matrix.f) * 0.75])
}

const close = (a: number, b: number) => Math.abs(a - b) < 0.75

describe('VML proposal geometry: composed transforms', () => {
  test('unitless-group corners (82,46,80,30)pt', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const inner = `<v:group id="group" style="position:absolute;left:72pt;top:36pt;width:200pt;height:100pt;" coordorigin="100,200" coordsize="1000,2000"><v:shape id="Grouped" style="position:absolute;left:150;top:400;width:400;height:600;"><v:textpath on="t" style="font-family:Arial;font-size:24pt;" string="Grouped"/></v:shape></v:group>`
    const doc = await parseDocx(await docxPkg(inner))
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing! as any
    const leaf = drawing.shapes[0].children[0]
    const calls = paintCalls(c => paintDrawing(drawing, c, 750, 500))
    const draw = calls.find(c => c.text === 'Grouped')
    expect(draw).toBeDefined()
    const corners = projectLeaf(leaf.widthEmu, leaf.heightEmu, draw.matrix)
    const wanted = [[82, 46], [162, 46], [162, 76], [82, 76]]
    for (let i = 0; i < 4; i++) {
      expect(close(corners[i][0], wanted[i][0])).toBe(true)
      expect(close(corners[i][1], wanted[i][1])).toBe(true)
    }
  })

  test('nested-group corners (92,46,20,10)pt', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const inner = `<v:group id="group" style="position:absolute;left:72pt;top:36pt;width:200pt;height:100pt;" coordorigin="100,200" coordsize="1000,2000"><v:group id="inner" style="position:absolute;left:100;top:200;width:500;height:1000;" coordorigin="-100,-200" coordsize="500,1000"><v:shape id="Nested" style="position:absolute;left:0;top:0;width:100;height:200;"><v:textpath on="t" style="font-family:Arial;font-size:24pt;" string="Nested"/></v:shape></v:group></v:group>`
    const doc = await parseDocx(await docxPkg(inner))
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing! as any
    const leaf = drawing.shapes[0].children[0].children[0]
    const calls = paintCalls(c => paintDrawing(drawing, c, 750, 500))
    const draw = calls.find(c => c.text === 'Nested')
    const corners = projectLeaf(leaf.widthEmu, leaf.heightEmu, draw.matrix)
    const wanted = [[92, 46], [112, 46], [112, 56], [92, 56]]
    for (let i = 0; i < 4; i++) {
      expect(close(corners[i][0], wanted[i][0])).toBe(true)
      expect(close(corners[i][1], wanted[i][1])).toBe(true)
    }
  })

  test('rotated-group corners (212,80),(212,160),(182,160),(182,80)', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const inner = `<v:group id="group" style="position:absolute;left:72pt;top:120pt;width:200pt;height:100pt;rotation:90;" coordorigin="100,200" coordsize="1000,2000"><v:shape id="Grouped" style="position:absolute;left:150;top:400;width:400;height:600;"><v:textpath on="t" style="font-family:Arial;font-size:24pt;" string="Grouped"/></v:shape></v:group>`
    const doc = await parseDocx(await docxPkg(inner))
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing! as any
    expect(drawing.shapes[0].rotationDeg).toBe(90)
    const leaf = drawing.shapes[0].children[0]
    const calls = paintCalls(c => paintDrawing(drawing, c, 750, 500))
    const draw = calls.find(c => c.text === 'Grouped')
    const corners = projectLeaf(leaf.widthEmu, leaf.heightEmu, draw.matrix)
    const wanted = [[212, 80], [212, 160], [182, 160], [182, 80]]
    for (let i = 0; i < 4; i++) {
      expect(close(corners[i][0], wanted[i][0])).toBe(true)
      expect(close(corners[i][1], wanted[i][1])).toBe(true)
    }
  })

  test('cm-all-fields bounds and outline', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const inner = `<v:shape id="Centimeter" style="position:absolute;left:2.54cm;top:1.27cm;width:5.08cm;height:2.54cm;" strokecolor="#0000FF" strokeweight="0.1cm"><v:textpath on="t" style="font-family:Arial;font-size:1cm" string="Centimeter"/></v:shape>`
    const doc = await parseDocx(await docxPkg(inner))
    const shape = (doc.sections[0].paragraphs[0].images[0].drawing as any).shapes[0]
    expect(shape.xEmu).toBe(Math.round(72 * 12700))
    expect(shape.yEmu).toBe(Math.round(36 * 12700))
    expect(shape.widthEmu).toBe(Math.round(144 * 12700))
    expect(shape.heightEmu).toBe(Math.round(72 * 12700))
    expect(shape.textBody.paragraphs[0].runs[0].fontSizePt).toBeCloseTo(28.346456692913385, 5)
    expect(shape.textBody.paragraphs[0].runs[0].textOutline?.widthPx).toBeCloseTo(3.7795275590551185, 4)
  })

  test('font shorthand orders', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const mk = async (style: string, id: string) => {
      const doc = await parseDocx(await docxPkg(`<v:shape id="${id}" style="position:absolute;left:0pt;top:0pt;width:160pt;height:40pt;"><v:textpath on="t" style="${style}" string="${id}"/></v:shape>`))
      return (doc.sections[0].paragraphs[0].images[0].drawing as any).shapes[0].textBody.paragraphs[0].runs[0]
    }
    const r1 = await mk(`font:italic bold 24pt/1.2 &quot;Times New Roman&quot;, serif`, 'S')
    expect(r1.fontFamily).toBe('Times New Roman')
    expect(r1.fontSizePt).toBe(24)
    expect(r1.bold).toBe(true)
    expect(r1.italic).toBe(true)
    const r2 = await mk(`font:italic bold 24pt &quot;Times New Roman&quot;;font-size:12pt;font-weight:normal`, 'L')
    expect(r2.fontSizePt).toBe(12)
    expect(r2.bold ?? false).toBe(false)
    expect(r2.italic).toBe(true)
    expect(r2.fontFamily).toBe('Times New Roman')
    const r3 = await mk(`font-size:12pt;font-weight:normal;font:italic bold 24pt &quot;Times New Roman&quot;`, 'A')
    expect(r3.fontSizePt).toBe(24)
    expect(r3.bold).toBe(true)
  })
})
