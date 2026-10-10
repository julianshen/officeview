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

describe('docx revision markup and symbol runs', () => {
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  async function parseBodyXml(bodyInner: string) {
    const JSZip = (await import('jszip')).default
    const zip = new JSZip()
    zip.file('word/document.xml', `<w:document ${NS}><w:body>${bodyInner}</w:body></w:document>`)
    zip.file('word/_rels/document.xml.rels', '<Relationships/>')
    return parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
  }

  test('w:sym parses to its codepoint with the symbol font', async () => {
    const doc = await parseBodyXml(
      `<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:sym w:font="Wingdings" w:char="F0B7"/></w:r></w:p>`)
    const run = doc.sections[0].paragraphs[0].runs[0]
    expect(run.text).toBe(String.fromCodePoint(0xF0B7))
    expect(run.fontFamily).toBe('Wingdings')
  })

  test('w:sym with invalid char is skipped safely', async () => {
    const doc = await parseBodyXml(
      `<w:p><w:r><w:t>A</w:t><w:sym w:font="Wingdings" w:char="ZZZ"/><w:sym w:font="Wingdings" w:char="110000"/></w:r></w:p>`)
    expect(doc.sections[0].paragraphs[0].runs[0].text).toBe('A')
  })

  test('mixed text and sym preserves order without overriding the run font', async () => {
    const doc = await parseBodyXml(
      `<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr><w:t>Item </w:t><w:sym w:font="Wingdings" w:char="F0A7"/></w:r></w:p>`)
    const run = doc.sections[0].paragraphs[0].runs[0]
    expect(run.text).toBe('Item ' + String.fromCodePoint(0xF0A7))
    expect(run.fontFamily).toBe('Calibri')
  })

  test('tracked insertions render while deletions stay hidden (Final view)', async () => {
    const doc = await parseBodyXml(
      `<w:p><w:r><w:t>Keep </w:t></w:r><w:ins><w:r><w:t>added</w:t></w:r></w:ins><w:del><w:r><w:t>removed</w:t></w:r></w:del></w:p>`)
    const text = doc.sections[0].paragraphs[0].runs.map(r => r.text).join('')
    expect(text).toContain('Keep ')
    expect(text).toContain('added')
    expect(text).not.toContain('removed')
  })

  test('symbols remain in order when a run also contains a drawing', async () => {
    const doc = await parseBodyXml('<w:p><w:r><w:t>A</w:t><w:sym w:font="Wingdings" w:char="F0B7"/><w:drawing/><w:t>B</w:t></w:r></w:p>')
    expect(doc.sections[0].paragraphs[0].runs.map(run => run.text).join('')).toBe('A' + String.fromCodePoint(0xF0B7) + 'B')
  })

  test('block-level tracked insertion collects inserted paragraphs', async () => {
    const doc = await parseBodyXml(
      `<w:p><w:r><w:t>First</w:t></w:r></w:p><w:ins><w:p><w:r><w:t>Inserted</w:t></w:r></w:p></w:ins><w:del><w:p><w:r><w:t>Deleted</w:t></w:r></w:p></w:del>`)
    const texts = doc.sections[0].paragraphs.map(p => p.runs.map(r => r.text).join(''))
    expect(texts).toEqual(['First', 'Inserted'])
  })
})
