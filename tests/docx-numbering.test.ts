import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { layoutDocx, renderPages, type MeasureFn } from '../src/docx/layout'
import { buildDocx, type DocxParaSpec } from '../src/testdata/ooxml-builders'

const measureFixed: MeasureFn = (text, style) => text.length * style.fontSizePt * 0.6 * (96 / 72)

const li = (text: string, numId: number, ilvl = 0): DocxParaSpec => ({ runs: [{ text }], numId, ilvl })

async function layout(paras: DocxParaSpec[]) {
  const doc = await parseDocx(await OfficePackage.load(await buildDocx(paras)))
  return { doc, pages: layoutDocx(doc, measureFixed) }
}

const markersOf = (pages: Awaited<ReturnType<typeof layout>>['pages']): string[] =>
  pages.flatMap((p) => p.lines.map((l) => l.marker?.text ?? ''))

describe('docx numbering', () => {
  test('decimal list numbers consecutive items', async () => {
    const { pages } = await layout([li('one', 1), li('two', 1), li('three', 1)])
    expect(markersOf(pages).filter(Boolean)).toEqual(['1.', '2.', '3.'])
  })

  test('bullet list uses the bullet glyph and does not count', async () => {
    const { pages } = await layout([li('alpha', 2), li('beta', 2), li('gamma', 2)])
    const markers = markersOf(pages).filter(Boolean)
    expect(markers).toHaveLength(3)
    expect(new Set(markers).size).toBe(1)
    expect(markers[0]).toBe('•')
  })

  test('alphabetic and roman formats', async () => {
    const alpha = await layout([li('a', 3), li('b', 3), li('c', 3)])
    expect(markersOf(alpha.pages).filter(Boolean)).toEqual(['a)', 'b)', 'c)'])
    const roman = await layout([li('i', 4), li('ii', 4), li('iii', 4)])
    expect(markersOf(roman.pages).filter(Boolean)).toEqual(['i.', 'ii.', 'iii.'])
  })

  test('nested levels keep independent counters', async () => {
    const { pages } = await layout([
      li('top one', 1, 0),
      li('sub a', 1, 1),
      li('sub b', 1, 1),
      li('top two', 1, 0),
      li('sub c', 1, 1),
    ])
    // level 1's lvlText is '%2.' — it counts independently and restarts
    // whenever the parent level ticks
    expect(markersOf(pages).filter(Boolean)).toEqual(['1.', '1.', '2.', '2.', '1.'])
  })

  test('separate numIds keep independent counters', async () => {
    const { pages } = await layout([li('a1', 1), li('b1', 3), li('a2', 1), li('b2', 3)])
    expect(markersOf(pages).filter(Boolean)).toEqual(['1.', 'a)', '2.', 'b)'])
  })

  test('list paragraphs carry the numbering level indent', async () => {
    const { doc } = await layout([li('nested', 1, 1)])
    expect(doc.sections[0].paragraphs[0].listLevel).toBe(1)
    // numbering.xml level 1 indent is 1440 twips
    expect(doc.sections[0].paragraphs[0].indentLeftTwips).toBe(1440)
  })

  test('marker is drawn in the gutter, text is indented past it', async () => {
    const { pages } = await layout([li('item text', 1)])
    const line = pages[0].lines[0]
    expect(line.marker?.text).toBe('1.')
    expect(line.marker!.widthPx).toBeGreaterThan(0)
    // the text starts to the right of the marker
    expect(line.xPx).toBeGreaterThanOrEqual(line.marker!.widthPx)
  })

  test('paints the marker glyph onto the canvas', async () => {
    const { pages } = await layout([li('visible', 1)])
    const page = pages[0]
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(Math.ceil(page.widthPx), Math.ceil(page.heightPx))
    const ctx = canvas.getContext('2d')!
    ctx.fillStyle = '#fff'
    ctx.fillRect(0, 0, canvas.width, canvas.height)
    renderPages(pages, ctx as unknown as CanvasRenderingContext2D)
    // ink appears left of the text (in the marker gutter)
    const line = page.lines[0]
    const gutter = ctx.getImageData(Math.round(line.xPx - line.marker!.widthPx), Math.round(line.yPx), Math.round(line.marker!.widthPx), Math.round(line.heightPx))
    let ink = 0
    for (let i = 0; i < gutter.data.length; i += 4) if (gutter.data[i] < 128) ink++
    expect(ink).toBeGreaterThan(3)
  })

  test('documents without numbering.xml are unaffected', async () => {
    const { pages } = await layout([{ runs: [{ text: 'plain' }] }])
    expect(markersOf(pages).filter(Boolean)).toHaveLength(0)
  })
})
