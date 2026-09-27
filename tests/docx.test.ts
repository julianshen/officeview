import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { layoutDocx, renderPages, type MeasureFn } from '../src/docx/layout'
import { buildDocx, type DocxParaSpec } from '../src/testdata/ooxml-builders'

async function parseFixture(paras: DocxParaSpec[]) {
  const pkg = await OfficePackage.load(await buildDocx(paras))
  return parseDocx(pkg)
}

const measureFixed: MeasureFn = (text, style) => {
  // deterministic mock measure: 0.6 * fontSize * charCount (pt -> px handled by caller)
  return text.length * style.fontSizePt * 0.6 * (96 / 72)
}

describe('docx parse', () => {
  test('parses paragraphs, runs, and section', async () => {
    const doc = await parseFixture([
      { align: 'center', runs: [{ text: 'Title', bold: true, size: 48 }] },
      { runs: [{ text: 'Body text here' }] },
    ])
    expect(doc.sections).toHaveLength(1)
    expect(doc.sections[0].pageSize.widthTwips).toBe(12240)
    expect(doc.sections[0].margins.topTwips).toBe(1440)
    const [p1, p2] = doc.sections[0].paragraphs
    expect(p1.align).toBe('center')
    expect(p1.runs[0]).toMatchObject({ text: 'Title', bold: true, fontSizePt: 24 })
    expect(p2.runs[0].text).toBe('Body text here')
  })
})

describe('docx layout', () => {
  test('breaks overflowing text across lines', async () => {
    const doc = await parseFixture([
      { runs: [{ text: 'aaaa '.repeat(20) }] },
    ])
    const pages = layoutDocx(doc, measureFixed)
    expect(pages.length).toBeGreaterThanOrEqual(1)
    const allLines = pages.flatMap((p) => p.lines)
    // 100 chars * 0.6 * 11pt * 4/3 = ~880px wide > content width (~9360twips=599px) — must wrap
    expect(allLines.length).toBeGreaterThan(1)
    // every line stays inside the content region: left margin 96px + 624px content width
    for (const line of allLines) {
      expect(line.xPx + line.widthPx).toBeLessThanOrEqual(720)
    }
  })

  test('respects page overflow into multiple pages', async () => {
    const doc = await parseFixture(
      Array.from({ length: 60 }, (_, i) => ({ runs: [{ text: `Line ${i}` }] })),
    )
    const pages = layoutDocx(doc, measureFixed)
    expect(pages.length).toBeGreaterThan(1)
    // pages ordered, lines within each page sorted by y
    for (const page of pages) {
      for (let i = 1; i < page.lines.length; i++) {
        expect(page.lines[i].yPx).toBeGreaterThanOrEqual(page.lines[i - 1].yPx)
      }
    }
  })
})

describe('docx render', () => {
  test('paints text pixels onto a real canvas', async () => {
    const doc = await parseFixture([
      { runs: [{ text: 'Hello OfficeView', bold: true, size: 32 }] },
    ])
    const pages = layoutDocx(doc, measureFixed)
    expect(pages).toHaveLength(1)
    const page = pages[0]

    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(Math.ceil(page.widthPx), Math.ceil(page.heightPx))
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, canvas.width, canvas.height)
    renderPages(pages, ctx as unknown as CanvasRenderingContext2D)

    // sample the ink: count non-white pixels in the text band
    const band = ctx.getImageData(0, Math.floor(page.lines[0].yPx), canvas.width, Math.ceil(page.lines[0].heightPx))
    let ink = 0
    for (let i = 0; i < band.data.length; i += 4) {
      if (band.data[i] < 128) ink++
    }
    expect(ink).toBeGreaterThan(100)
  })
})
