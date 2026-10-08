/**
 * R1: transparent ordinary DOCX text must keep its parsed color/opacity.
 * Canonical control: w:color=008800 with w14 solid 008800 + w14:alpha.
 * Real package -> parse -> layout -> paint evidence (not model-only).
 */
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { createMeasurer, layoutDocx, renderPages } from '../src/docx/layout'
import { resolveColor, hexRgbToCss } from '../src/core/color'

const docXml = (alpha: number): string => `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"><w:body><w:p><w:r><w:rPr><w:color w:val="008800"/><w:sz w:val="48"/><w14:textFill><w14:solidFill><w14:srgbClr w14:val="008800"><w14:alpha w14:val="${alpha}"/></w14:srgbClr></w14:solidFill></w14:textFill></w:rPr><w:t>Transparent ordinary text</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`

async function paintDocXml(xml: string, W = 816, H = 1056, stylesXml?: string): Promise<{ doc: Awaited<ReturnType<typeof parseDocx>>; opaque: number; green: number; black: number }> {
  const zip = new JSZip()
  zip.file('word/document.xml', xml)
  if (stylesXml) zip.file('word/styles.xml', stylesXml)
  const bytes = await zip.generateAsync({ type: 'uint8array' })
  const doc = await parseDocx(await OfficePackage.load(bytes))
  const canvas = createCanvas(W, H)
  const ctx = canvas.getContext('2d') as any
  renderPages(layoutDocx(doc, createMeasurer(ctx as CanvasRenderingContext2D)), ctx)
  const pix = ctx.getImageData(0, 0, W, H).data
  let opaque = 0, green = 0, black = 0
  for (let i = 0; i < pix.length; i += 4) {
    if (pix[i + 3] > 0) {
      opaque++
      if (pix[i + 1] > pix[i] + 20 && pix[i + 1] > pix[i + 2] + 20) green++
      if (pix[i] < 10 && pix[i + 1] < 10 && pix[i + 2] < 10) black++
    }
  }
  return { doc, opaque, green, black }
}

describe('R1 inherited style and rotated runs', () => {
  it('paragraph-style w14 fill inherits the same rgba through paint', async () => {
    const styles = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:style w:type="paragraph" w:styleId="Green50"><w:rPr><w14:textFill><w14:solidFill><w14:srgbClr w14:val="008800"><w14:alpha w14:val="50000"/></w14:srgbClr></w14:solidFill></w14:textFill></w:rPr></w:style></w:styles>`
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p><w:pPr><w:pStyle w:val="Green50"/></w:pPr><w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t>Styled green text</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`
    const r = await paintDocXml(xml, 816, 1056, styles)
    expect(r.doc.sections[0].paragraphs[0].runs[0].color).toBe('rgba(0,136,0,0.5)')
    expect(r.opaque).toBeGreaterThan(0)
    expect(r.green).toBeGreaterThan(0)
    expect(r.black).toBe(0)
  })
  it('rotated table-cell run with underline/strike paints green, not black', async () => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="4000"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:textDirection w:val="tbRl"/><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:rPr><w:color w:val="008800"/><w:sz w:val="48"/><w:u w:val="single"/><w:strike/><w14:textFill><w14:solidFill><w14:srgbClr w14:val="008800"><w14:alpha w14:val="50000"/></w14:srgbClr></w14:solidFill></w14:textFill></w:rPr><w:t>Vertical green</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`
    const r = await paintDocXml(xml)
    expect(r.opaque).toBeGreaterThan(0)
    expect(r.green).toBeGreaterThan(0)
    expect(r.black).toBe(0)
  })
})
async function paintAlpha(alpha: number): Promise<{ parsedColor: unknown; calls: Array<{ text: string; color: unknown }>; opaque: number; green: number; black: number }> {
  const zip = new JSZip()
  zip.file('word/document.xml', docXml(alpha))
  const bytes = await zip.generateAsync({ type: 'uint8array' })
  const doc = await parseDocx(await OfficePackage.load(bytes))
  const canvas = createCanvas(816, 1056)
  const ctx = canvas.getContext('2d') as any
  const calls: Array<{ text: string; color: unknown }> = []
  const fill = ctx.fillText.bind(ctx)
  ctx.fillText = (text: string, x: number, y: number) => {
    calls.push({ text, color: ctx.fillStyle })
    return fill(text, x, y)
  }
  renderPages(layoutDocx(doc, createMeasurer(ctx as CanvasRenderingContext2D)), ctx)
  const pix = ctx.getImageData(0, 0, 816, 1056).data
  let opaque = 0, green = 0, black = 0
  for (let i = 0; i < pix.length; i += 4) {
    if (pix[i + 3] > 0) {
      opaque++
      if (pix[i + 1] > pix[i] + 20 && pix[i + 1] > pix[i + 2] + 20) green++
      if (pix[i] < 10 && pix[i + 1] < 10 && pix[i + 2] < 10) black++
    }
  }
  return { parsedColor: doc.sections[0].paragraphs[0].runs[0].color, calls, opaque, green, black }
}

describe('R1 transparent ordinary DOCX color/opacity', () => {
  // Canvas splits runs into word calls and serializes alpha with its own
  // precision, so painter colors are compared numerically, not as strings.
  const rgbaOf = (color: unknown): [number, number, number, number] | undefined => {
    const hex = /^#([0-9a-fA-F]{6})$/.exec(String(color).trim())
    if (hex) return [parseInt(hex[1].slice(0, 2), 16), parseInt(hex[1].slice(2, 4), 16), parseInt(hex[1].slice(4, 6), 16), 1]
    const m = /^\s*rgba?\(\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*(?:,\s*(\d+(?:\.\d+)?)\s*)?\)\s*$/.exec(String(color))
    if (!m) return undefined
    return [Number(m[1]), Number(m[2]), Number(m[3]), m[4] === undefined ? 1 : Number(m[4])]
  }
  it('parsed 50% color reaches paint as green, not black', async () => {
    const r = await paintAlpha(50000)
    expect(r.parsedColor).toBe('rgba(0,136,0,0.5)')
    expect(r.calls.length).toBeGreaterThan(0)
    for (const c of r.calls) expect(rgbaOf(c.color)).toEqual([0, 136, 0, 0.5])
    expect(r.opaque).toBeGreaterThan(0)
    expect(r.green).toBeGreaterThan(0)
    expect(r.black).toBe(0)
  })
  it('opaque 100% stays green', async () => {
    const r = await paintAlpha(100000)
    expect(r.parsedColor).toBe('#008800')
    expect(r.calls.length).toBeGreaterThan(0)
    for (const c of r.calls) expect(rgbaOf(c.color)).toEqual([0, 136, 0, 1])
    expect(r.opaque).toBeGreaterThan(0)
    expect(r.green).toBeGreaterThan(0)
    expect(r.black).toBe(0)
  })
  it('fully transparent 0% paints no visible ink', async () => {
    const r = await paintAlpha(0)
    expect(r.parsedColor).toBe('rgba(0,136,0,0)')
    expect(r.calls.length).toBeGreaterThan(0)
    for (const c of r.calls) expect(rgbaOf(c.color)).toEqual([0, 136, 0, 0])
    expect(r.opaque).toBe(0)
    expect(r.green).toBe(0)
    expect(r.black).toBe(0)
  })
})

describe('R1 shared CSS color helper semantics', () => {
  it('preserves existing hex/ARGB/auto/invalid behavior and passes valid CSS through', () => {
    expect(resolveColor(undefined)).toBe('#000000')
    expect(resolveColor('auto')).toBe('#000000')
    expect(resolveColor('FF0000')).toBe('#FF0000')
    expect(resolveColor('FF000000')).toBe('#000000')
    expect(resolveColor('not-a-color')).toBe('#000000')
    // hexRgbToCss keeps its hex/ARGB-only contract: CSS is rejected there and
    // handled narrowly by resolveColor instead.
    expect(hexRgbToCss('rgba(0,136,0,0.5)')).toBe(undefined)
    expect(hexRgbToCss('rgb(0,136,0)')).toBe(undefined)
    expect(resolveColor('rgba(0,136,0,0.5)')).toBe('rgba(0,136,0,0.5)')
    expect(resolveColor('rgba(0,136,0,0)')).toBe('rgba(0,136,0,0)')
    expect(resolveColor('rgb(0,136,0)')).toBe('rgb(0,136,0)')
    expect(resolveColor('red')).toBe('#000000')
  })
})
