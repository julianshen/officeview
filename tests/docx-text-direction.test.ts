import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { layoutDocx, renderPages, type MeasureFn, type PageLayout } from '../src/docx/layout'
import { graphemes, RECORD_TEXT, type TextRecordingContext } from '../src/core/text-recording'
import { CT_TYPES, ROOT_RELS } from '../src/testdata/ooxml-builders'
import { buildTextIndex, findMatches } from '../src/core/search'
import { hitTest, rectsForSelectionOnPage, textForRange } from '../src/core/selection'

const measure: MeasureFn = (text, style) => text.length * style.fontSizePt * 0.8
const ns =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const cell = (direction: string | undefined, texts: string | string[], w = 2000) =>
  `<w:tc><w:tcPr>${direction === undefined ? '' : `<w:textDirection w:val="${direction}"/>`}<w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${(Array.isArray(texts) ? texts : [texts]).map(text => `<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`).join('')}</w:tc>`
const table = (cells: string) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cells}</w:tr></w:tbl>`
const bottomUpCell = (texts: string[], properties = '', paragraphProperties = '') =>
  `<w:tc><w:tcPr><w:textDirection w:val="btLr"/><w:vAlign w:val="top"/><w:tcMar><w:top w:w="90"/><w:bottom w:w="150"/><w:left w:w="120"/><w:right w:w="180"/></w:tcMar>${properties}</w:tcPr>${texts.map(text => `<w:p><w:pPr>${paragraphProperties}</w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`).join('')}</w:tc>`
const heightRow = (contents: string, height: number, rule = 'exact') =>
  `<w:tr><w:trPr><w:trHeight w:val="${height}" w:hRule="${rule}"/></w:trPr>${contents}</w:tr>`
const rowsTable = (rows: string) =>
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>${rows}</w:tbl>`
async function fixture(body: string) {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CT_TYPES)
  zip.file('_rels/.rels', ROOT_RELS)
  zip.file(
    'word/document.xml',
    `<w:document ${ns}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:left="1440" w:right="1440" w:bottom="1440"/></w:sectPr></w:body></w:document>`
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>'
  )
  return parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
}
const cellBoxOf = (doc: Awaited<ReturnType<typeof fixture>>) => {
  // Cells are table-relative; lines are page-absolute.
  const page = layoutDocx(doc, measure)[0]
  const table = page.tables[0]
  const box = table.rows[0].cells[0]
  return { ...box, xPx: box.xPx + table.xPx, yPx: box.yPx + table.yPx }
}
const linesIn = (doc: Awaited<ReturnType<typeof fixture>>) => {
  const page = layoutDocx(doc, measure)[0]
  const box = cellBoxOf(doc)
  return page.lines.filter(l => l.xPx + l.widthPx > box.xPx && l.xPx < box.xPx + box.widthPx && l.yPx + l.heightPx > box.yPx && l.yPx < box.yPx + box.heightPx)
}
// Inspect glyph origins within an intact shaped run without requiring the
// renderer to paint each glyph separately. Fixed widths pin CCW progression.
const glyphsOf = (page: PageLayout) => page.lines.flatMap(line => line.segs.flatMap(seg => graphemes(seg.text).map(g => {
  const advance = measure(seg.text.slice(0, g.start), seg.style)
  const t = seg.transform
  return { ...seg, text: g.text, widthPx: measure(g.text, seg.style), transform: t ? { ...t, e: t.e + t.a * advance, f: t.f + t.b * advance } : undefined }
})))

describe('docx table cell text directions', () => {
  test.each(['lrTb', 'tbRl', 'btLr', 'lrTbV', 'tbRlV', 'tbLrV'] as const)('parses w:textDirection %s onto the cell', async (direction) => {
    const doc = await fixture(table(cell(direction, 'hi')))
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw Error('table')
    expect(block.table.rows[0].cells[0].textDirection).toBe(direction)
  })

  test('absent or invalid textDirection stays horizontal', async () => {
    for (const direction of [undefined, 'sideways']) {
      const doc = await fixture(table(cell(direction, 'hi')))
      const block = doc.sections[0].blocks[0]
      if (block.kind !== 'table') throw Error('table')
      expect(block.table.rows[0].cells[0].textDirection ?? 'lrTb').toBe('lrTb')
    }
  })

  test('tbRl rotates the whole line including CJK; tbRlV keeps mixed upright CJK', async () => {
    const audit = async (direction: string) => {
      const doc = await fixture(table(cell(direction, ['文A'])))
      const page = layoutDocx(doc, measure)[0]
      const segs = glyphsOf(page)
      const at = (text: string) => segs.find(s => s.text === text)?.transform
      return { wen: at('文'), a: at('A') }
    }
    const tbRl = await audit('tbRl')
    expect([tbRl.wen?.a, tbRl.wen?.b]).toEqual([0, 1])
    expect([tbRl.a?.a, tbRl.a?.b]).toEqual([0, 1])
    const mixed = await audit('tbRlV')
    expect([mixed.wen?.a, mixed.wen?.d]).toEqual([1, 1])
    expect([mixed.a?.a, mixed.a?.b]).toEqual([0, 1])
  })

  test('tbRl stacks columns right-to-left inside cell bounds', async () => {
    // Each paragraph starts a new column, as in native vertical text.
    const doc = await fixture(table(cell('tbRl', ['AAAA BBBB', 'CCCC DDDD'])))
    const lines = linesIn(doc)
    expect(lines.length).toBeGreaterThan(1)
    const box = cellBoxOf(doc)
    for (const line of lines) {
      expect(line.xPx).toBeGreaterThanOrEqual(box.xPx)
      expect(line.xPx + line.widthPx).toBeLessThanOrEqual(box.xPx + box.widthPx + 1)
    }
    expect(lines[0].xPx).toBeGreaterThan(lines[1].xPx)
  })

  test('tbLrV mirrors column progression left-to-right', async () => {
    const doc = await fixture(table(cell('tbLrV', ['AAAA BBBB', 'CCCC DDDD'])))
    const lines = linesIn(doc)
    expect(lines.length).toBeGreaterThan(1)
    expect(lines[0].xPx).toBeLessThan(lines[1].xPx)
    // The whole block anchors at the cell's left edge, not the right.
    const box = cellBoxOf(doc)
    expect(lines[0].xPx).toBeLessThan(box.xPx + box.widthPx / 2)
  })

  test('btLr runs bottom-to-top inside finite cell bounds with visible glyphs', async () => {
    const doc = await fixture(table(cell('btLr', 'AB')))
    const box = cellBoxOf(doc)
    const segs = glyphsOf(layoutDocx(doc, measure)[0]).filter(s => s.transform)
    expect(segs.length).toBeGreaterThan(0)
    const at = (text: string) => segs.find(s => s.text === text)?.transform
    expect(at('A')!.f).toBeGreaterThan(at('B')!.f)
    // Absolute origins: finite, inside the cell box, not sunk ~1e6px.
    for (const seg of segs) {
      const t = seg.transform!
      expect(Number.isFinite(t.e) && Number.isFinite(t.f)).toBe(true)
      expect(t.e).toBeGreaterThanOrEqual(box.xPx)
      expect(t.e).toBeLessThanOrEqual(box.xPx + box.widthPx)
      expect(t.f).toBeGreaterThanOrEqual(box.yPx)
      expect(t.f).toBeLessThanOrEqual(box.yPx + box.heightPx)
    }
    const ctx = createCanvas(816, 1056).getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, 816, 1056)
    renderPages(layoutDocx(doc, measure), ctx as never)
    const data = ctx.getImageData(0, 0, 816, 1056).data
    let count = 0
    for (let y = Math.floor(box.yPx); y < box.yPx + box.heightPx; y++) {
      for (let x = Math.floor(box.xPx); x < box.xPx + box.widthPx; x++) {
        if (data[(y * 816 + x) * 4] < 128) count++
      }
    }
    expect(count).toBeGreaterThan(0)
  })

  test('btLr starts each CCW paragraph at the final exact cell bottom and records matching carets', async () => {
    const doc = await fixture(rowsTable(heightRow(cell(undefined, 'heading'), 600) + heightRow(bottomUpCell(['AB中', 'CD']), 6000)))
    const page = layoutDocx(doc, measure)[0]
    const tableBox = page.tables[0]
    const row = tableBox.rows[1]
    const box = row.cells[0]
    const left = tableBox.xPx + box.xPx, top = tableBox.yPx + box.yPx
    const bottom = top + box.heightPx - 10
    const segs = glyphsOf(page).filter(seg => seg.transform)
    const first = segs.find(seg => seg.text === 'A')!, second = segs.find(seg => seg.text === 'C')!
    expect(tableBox.rows.map(row => row.heightPx)).toEqual([40, 400])
    // Different intrinsic paragraph lengths share the cell's actual bottom.
    expect(first.transform!.f).toBeCloseTo(bottom)
    expect(second.transform!.f).toBeCloseTo(bottom)
    expect(second.transform!.e).toBeGreaterThan(first.transform!.e)
    expect(first.transform!.e).toBeCloseTo(left + 8 + 13.2)
    for (const seg of segs) {
      const t = seg.transform!
      expect([t.a, t.b, t.c, t.d]).toEqual([0, -1, 1, 0])
      expect(t.f - seg.widthPx).toBeGreaterThanOrEqual(top + 6)
      expect(t.f).toBeLessThanOrEqual(bottom)
      expect(t.e - 9.6).toBeGreaterThanOrEqual(left + 8)
      expect(t.e + 2.4).toBeLessThanOrEqual(left + box.widthPx - 12)
    }
    expect(segs.find(seg => seg.text === 'B')!.transform!.f).toBeCloseTo(bottom - 9.6)
    expect(segs.find(seg => seg.text === '中')!.transform!.f).toBeCloseTo(bottom - 19.2)
    const index = await buildTextIndex([{ spec: { widthPx: page.widthPx, heightPx: page.heightPx }, paint: ctx => renderPages([page], ctx) }])
    const lineIndex = index.pages[0].lines.findIndex(line => line.text === 'AB中')
    expect(lineIndex).toBeGreaterThanOrEqual(0)
    const span = index.pages[0].lines[lineIndex].spans[0]
    expect(span.y).toBeCloseTo(bottom)
    expect(span.placement!.transform).toMatchObject(first.transform!)
    for (const strict of [false, true]) {
      expect(hitTest(index, 0, first.transform!.e - 4, bottom - span.width / span.text.length * .75, { strict })).toEqual({ pageIndex: 0, lineIndex, charIndex: 1 })
    }
    expect(hitTest(index, 0, first.transform!.e - 4, top + 12, { strict: true })).toBeUndefined()
    const range = { start: { pageIndex: 0, lineIndex, charIndex: 0 }, end: { pageIndex: 0, lineIndex, charIndex: 3 } }
    expect(textForRange(index, range)).toBe('AB中')
    expect(findMatches(index, 'AB中')).toHaveLength(1)
    for (const [i, rect] of rectsForSelectionOnPage(index, 0, range).entries()) {
      const recorded = index.pages[0].lines[lineIndex].spans[i]
      // Recording uses the paint font's measured widths; layout origins remain
      // fixed, so these relative bands are portable across available fonts.
      expect(rect.y).toBeCloseTo(recorded.y - recorded.width)
      expect(rect.y + rect.height).toBeCloseTo(recorded.y)
      expect(rect.y).toBeGreaterThanOrEqual(top + 6)
    }
    const ctx = createCanvas(816, 1056).getContext('2d')
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, 816, 1056)
    renderPages([page], ctx as never)
    const pixels = ctx.getImageData(Math.floor(left + 8), Math.floor(bottom - 40), 60, 40).data
    expect(pixels.some((value, i) => i % 4 === 0 && value < 128)).toBe(true)
  })

  test.each(['left', 'center', 'right'])('btLr composes paragraph %s alignment with the actual flow height', async align => {
    const doc = await fixture(rowsTable(heightRow(bottomUpCell(['AB'], '', `<w:jc w:val="${align}"/><w:ind w:left="120" w:right="180"/>`), 6000)))
    const page = layoutDocx(doc, measure)[0], box = cellBoxOf(doc)
    const a = glyphsOf(page).find(seg => seg.text === 'A')!.transform!
    const available = 400 - 6 - 10 - 8 - 12
    const offset = align === 'center' ? (available - 19.2) / 2 : align === 'right' ? available - 19.2 : 0
    expect(a.f).toBeCloseTo(box.yPx + 400 - 10 - 8 - offset)
    // Paragraph flow indents do not move the transverse baseline.
    expect(a.e).toBeCloseTo(box.xPx + 8 + 13.2)
  })

  test('btLr wraps using exact inner height and keeps every glyph in successive LTR columns', async () => {
    const text = 'ABCDEFGHIJ'
    const doc = await fixture(rowsTable(heightRow(bottomUpCell([text]), 900)))
    const page = layoutDocx(doc, measure)[0], box = cellBoxOf(doc)
    const segs = page.lines.flatMap(line => line.segs)
    expect(segs.map(seg => seg.text).join('')).toBe(text)
    expect(page.tables[0].rows[0].heightPx).toBe(60)
    expect(page.lines).toHaveLength(3)
    expect(page.lines[0].segs[0].transform!.e).toBeLessThan(page.lines[1].segs[0].transform!.e)
    for (const line of page.lines) {
      expect(line.segs[0].transform!.f).toBeCloseTo(box.yPx + 50)
      for (const seg of line.segs) expect(seg.transform!.f - seg.widthPx).toBeGreaterThanOrEqual(box.yPx + 6)
    }
  })

  test.each(['tbRl', 'tbRlV', 'tbLrV'])('%s also wraps against exact row height', async direction => {
    const doc = await fixture(rowsTable(heightRow(bottomUpCell(['ABCDEFGHIJ']).replace('w:val="btLr"', `w:val="${direction}"`), 900)))
    const page = layoutDocx(doc, measure)[0], box = cellBoxOf(doc)
    expect(page.tables[0].rows[0].heightPx).toBe(60)
    expect(page.lines).toHaveLength(3)
    expect(page.lines.flatMap(line => line.segs).map(seg => seg.text).join('')).toBe('ABCDEFGHIJ')
    for (const line of page.lines) {
      expect(line.segs[0].transform!.f).toBeCloseTo(box.yPx + 6)
      for (const seg of line.segs) expect(seg.transform!.f + seg.widthPx).toBeLessThanOrEqual(box.yPx + 50)
    }
  })

  test.each(['top', 'center', 'bottom'])('btLr cell %s alignment positions the transverse block', async align => {
    const doc = await fixture(rowsTable(heightRow(bottomUpCell(['AB']).replace('w:val="top"', `w:val="${align}"`), 6000)))
    const page = layoutDocx(doc, measure)[0], box = cellBoxOf(doc)
    const a = page.lines[0].segs[0].transform!
    const slack = box.widthPx - 8 - 12 - 19.2
    const shift = align === 'center' ? slack / 2 : align === 'bottom' ? slack : 0
    expect(a.e).toBeCloseTo(box.xPx + 8 + 13.2 + shift)
    expect(a.f).toBeCloseTo(box.yPx + 400 - 10)
  })

  test('btLr projects against the merged cell final height without shifting following flow', async () => {
    const doc = await fixture(rowsTable(heightRow(bottomUpCell(['AB', 'CD'], '<w:vMerge w:val="restart"/>'), 900) + heightRow('<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>', 1200)) + '<w:p><w:r><w:t>after</w:t></w:r></w:p>')
    const page = layoutDocx(doc, measure)[0], box = cellBoxOf(doc)
    expect(page.tables[0].rows.map(row => row.heightPx)).toEqual([60, 80])
    expect(box.heightPx).toBe(140)
    for (const text of ['A', 'C']) {
      expect(glyphsOf(page).find(seg => seg.text === text)!.transform!.f).toBeCloseTo(box.yPx + 130)
    }
    expect(page.lines.find(line => line.segs.some(seg => seg.text === 'after'))!.yPx).toBeGreaterThanOrEqual(box.yPx + 140)
  })

  test.each(['auto', 'atLeast'])('btLr %s height grows intrinsically and stays finite', async rule => {
    const contents = bottomUpCell(['ABCDEFGHIJ'])
    const row = rule === 'auto' ? `<w:tr>${contents}</w:tr>` : heightRow(contents, 600, rule)
    const doc = await fixture(rowsTable(row) + '<w:p><w:r><w:t>after</w:t></w:r></w:p>')
    const page = layoutDocx(doc, measure)[0], box = cellBoxOf(doc)
    expect(box.heightPx).toBeCloseTo(6 + 96 + 10)
    expect(page.lines[0].segs[0].transform!.f).toBeCloseTo(box.yPx + box.heightPx - 10)
    expect(page.lines.find(line => line.segs.some(seg => seg.text === 'after'))!.yPx).toBeGreaterThanOrEqual(box.yPx + box.heightPx)
  })

  test('btLr uses an atLeast minimum taller than its intrinsic content as the common bottom', async () => {
    const doc = await fixture(rowsTable(heightRow(bottomUpCell(['AB', 'C']), 6000, 'atLeast')))
    const page = layoutDocx(doc, measure)[0], box = cellBoxOf(doc)
    expect(box.heightPx).toBe(400)
    for (const text of ['A', 'C']) expect(glyphsOf(page).find(seg => seg.text === text)!.transform!.f).toBeCloseTo(box.yPx + 390)
  })

  test('btLr intrinsic height includes both writing-local paragraph indents', async () => {
    const doc = await fixture(rowsTable(`<w:tr>${bottomUpCell(['ABCDEFGHIJ'], '', '<w:ind w:left="120" w:right="180"/>')}</w:tr>`))
    const page = layoutDocx(doc, measure)[0], box = cellBoxOf(doc)
    expect(box.heightPx).toBeCloseTo(6 + 8 + 96 + 12 + 10)
    expect(page.lines).toHaveLength(1)
    expect(page.lines[0].segs[0].transform!.f).toBeCloseTo(box.yPx + box.heightPx - 10 - 8)
    const last = page.lines[0].segs.at(-1)!
    expect(last.transform!.f - last.widthPx).toBeCloseTo(box.yPx + 6 + 12)
  })

  test('btLr rewraps in a completed merge and preserves subsequent cell line slots', async () => {
    const doc = await fixture(rowsTable(
      heightRow(bottomUpCell(['ABCDEFGHIJ'], '<w:vMerge w:val="restart"/>'), 900) +
      heightRow('<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>', 1200) +
      heightRow(bottomUpCell(['XY']), 900)
    ))
    const page = layoutDocx(doc, measure)[0], first = cellBoxOf(doc)
    // The initial 44px flow would need three columns; the final merged 124px
    // flow fits the source in one column and moves the later cell's line slot.
    expect(page.lines).toHaveLength(2)
    expect(page.lines.map(line => line.segs.map(seg => seg.text).join(''))).toEqual(['ABCDEFGHIJ', 'XY'])
    expect(page.lines[0].segs[0].transform!.f).toBeCloseTo(first.yPx + 130)
    expect(page.lines[1].segs[0].transform!.f).toBeCloseTo(first.yPx + 190)
  })

  test('btLr can project into a merge whose anchor row has no inner flow height', async () => {
    const doc = await fixture(rowsTable(heightRow(bottomUpCell(['AB'], '<w:vMerge w:val="restart"/>'), 150) + heightRow('<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>', 1200)))
    const page = layoutDocx(doc, measure)[0], box = cellBoxOf(doc)
    expect(box.heightPx).toBe(90)
    expect(page.lines.flatMap(line => line.segs).map(seg => seg.text).join('')).toBe('AB')
    expect(page.lines[0].segs[0].transform!.f).toBeCloseTo(box.yPx + 80)
  })

  test.each(['tbRl', 'btLr', 'tbRlV', 'tbLrV', 'lrTbV'])('%s allocates ordered physical inline cell images and retains exact mixed source', async direction => {
    const doc = await fixture(rowsTable(heightRow(cell(direction, ['ABC', 'Following']), 6000)))
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw Error('table')
    block.table.cellMargins = { leftTwips: 0, rightTwips: 0, topTwips: 0, bottomTwips: 0 }
    for (const row of block.table.rows) for (const cell of row.cells) cell.margins = block.table.cellMargins
    const paragraphs = block.table.rows[0].cells[0].paragraphs
    const images = [24, 24].map((width, i) => ({ data: new Uint8Array([i + 1]), widthEmu: width * 9525, heightEmu: 24 * 9525 }))
    paragraphs[0].runs = ['A', 'B', 'C'].map(text => ({ text, fontSizePt: 12 }))
    paragraphs[0].images = images
    paragraphs[0].inline = [{ kind: 'text', run: paragraphs[0].runs[0] }, { kind: 'image', image: images[0] }, { kind: 'text', run: paragraphs[0].runs[1] }, { kind: 'image', image: images[1] }, { kind: 'text', run: paragraphs[0].runs[2] }]
    const fixed: MeasureFn = Object.assign((text: string) => [...text].length * 10, { metrics: () => ({ ascent: 12, descent: 4, normalHeight: 16 }) })
    const page = layoutDocx(doc, fixed)[0], table = page.tables[0], box = table.rows[0].cells[0]
    const [a, b] = page.images
    expect(a.imageIndex).toBe(0); expect(b.imageIndex).toBe(1)
    expect(Math.min(a.xPx + a.widthPx, b.xPx + b.widthPx) - Math.max(a.xPx, b.xPx) <= 0 || Math.min(a.yPx + a.heightPx, b.yPx + b.heightPx) - Math.max(a.yPx, b.yPx) <= 0).toBe(true)
    if (direction === 'lrTbV') expect(a.xPx + a.widthPx).toBeLessThanOrEqual(b.xPx)
    else if (direction === 'btLr') expect(a.yPx).toBeGreaterThanOrEqual(b.yPx + b.heightPx)
    else expect(a.yPx + a.heightPx).toBeLessThanOrEqual(b.yPx)
    for (const image of page.images) {
      expect(image.xPx).toBeGreaterThanOrEqual(table.xPx + box.xPx - 1e-8)
      expect(image.yPx).toBeGreaterThanOrEqual(table.yPx + box.yPx - 1e-8)
      expect(image.xPx + image.widthPx).toBeLessThanOrEqual(table.xPx + box.xPx + box.widthPx + 1e-8)
      expect(image.yPx + image.heightPx).toBeLessThanOrEqual(table.yPx + box.yPx + box.heightPx + 1e-8)
    }
    const decoded = images.map(() => createCanvas(2, 2)), painted: number[][] = []
    const index = await buildTextIndex([{ spec: { widthPx: page.widthPx, heightPx: page.heightPx }, paint: ctx => {
      const actual = ctx.drawImage.bind(ctx)
      ctx.drawImage = ((image: CanvasImageSource, x: number, y: number, w: number, h: number) => { painted.push([decoded.indexOf(image as never), x, y, w, h]); actual(image, x, y, w, h) }) as never
      renderPages([page], ctx, decoded as never)
    } }])
    expect(painted).toEqual(page.images.map(image => [image.imageIndex, image.xPx, image.yPx, image.widthPx, image.heightPx]))
    const firstIndexedLine = index.pages[0].lines[0]
    for (const [text, charIndex] of [['A', 1], ['B', 2]] as const) {
      const span = firstIndexedLine.spans.find(span => span.text === text)!, p = span.placement!, t = p.transform
      const x = t.a * (p.x + p.width * .75) + t.c * p.y + t.e, y = t.b * (p.x + p.width * .75) + t.d * p.y + t.f
      for (const strict of [true, false]) expect(hitTest(index, 0, x, y, { strict })).toEqual({ pageIndex: 0, lineIndex: 0, charIndex })
    }
    const lines = index.pages[0].lines, last = lines.length - 1
    expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: last, charIndex: lines[last].text.length } })).toBe('ABC\nFollowing')
    const before = page.lines.flatMap(line => line.segs).find(seg => seg.text === 'A')!, after = page.lines.flatMap(line => line.segs).find(seg => seg.text === 'B')!
    if (direction === 'lrTbV') { expect(before.transform!.e + before.widthPx).toBeCloseTo(a.xPx); expect(after.transform!.e).toBeCloseTo(a.xPx + a.widthPx) }
    else if (direction === 'btLr') { expect(before.transform!.f - before.widthPx).toBeCloseTo(a.yPx + a.heightPx); expect(after.transform!.f).toBeCloseTo(a.yPx) }
    else { expect(before.transform!.f + before.widthPx).toBeCloseTo(a.yPx); expect(after.transform!.f).toBeCloseTo(a.yPx + a.heightPx) }
  })

  test.each(['tbRl', 'btLr', 'tbRlV', 'tbLrV', 'lrTbV'])('%s reserves leading/consecutive/tall/image-only cell slots across row rules and completed merges', async direction => {
    for (const rule of ['auto', 'atLeast', 'exact', 'merged'] as const) {
      const rows = heightRow(cell(direction, ['A文B', '', 'Following']).replace('<w:tcPr>', `<w:tcPr>${rule === 'merged' ? '<w:vMerge w:val="restart"/>' : ''}`), rule === 'exact' ? 6000 : rule === 'merged' ? 0 : 600, rule === 'auto' ? 'atLeast' : rule === 'merged' ? 'exact' : rule) + (rule === 'merged' ? heightRow('<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>', 3000) : '')
      const doc = await fixture(rowsTable(rows)), block = doc.sections[0].blocks[0]
      if (block.kind !== 'table') throw Error('table')
      block.table.cellMargins = { leftTwips: 0, rightTwips: 0, topTwips: 0, bottomTwips: 0 }
      for (const row of block.table.rows) for (const cell of row.cells) cell.margins = block.table.cellMargins
      block.table.gridColsTwips = [6000]
      if (rule === 'auto') { delete block.table.rows[0].heightTwips; delete block.table.rows[0].heightRule }
      const paras = block.table.rows[0].cells[0].paragraphs
      const images = [[30, 60], [12, 20], [24, 24], [8, 16], [12, 20], [18, 12]].map(([w, h], i) => ({ data: new Uint8Array([i + 1]), widthEmu: w * 9525, heightEmu: h * 9525 }))
      paras[0].runs = [{ text: 'A', fontSizePt: 12 }, { text: '文B', fontSizePt: 12 }]
      paras[0].images = images.slice(0, 4)
      paras[0].inline = [{ kind: 'image', image: images[0] }, { kind: 'image', image: images[1] }, { kind: 'text', run: paras[0].runs[0] }, { kind: 'image', image: images[2] }, { kind: 'text', run: paras[0].runs[1] }, { kind: 'image', image: images[3] }]
      paras[1].images = images.slice(4); paras[1].inline = images.slice(4).map(image => ({ kind: 'image', image }))
      const fixed: MeasureFn = Object.assign((text: string) => [...text].length * 10, { metrics: () => ({ ascent: 12, descent: 4, normalHeight: 16 }) })
      const page = layoutDocx(doc, fixed)[0], table = page.tables[0], box = table.rows[0].cells[0]
      expect(page.images.map(image => image.imageIndex)).toEqual([0, 1, 2, 3, 4, 5])
      const nonoverlap = (a: { x: number; y: number; w: number; h: number }, b: { x: number; y: number; w: number; h: number }) => Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x) <= 1e-8 || Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y) <= 1e-8
      const imageRects = page.images.map(image => ({ x: image.xPx, y: image.yPx, w: image.widthPx, h: image.heightPx }))
      for (const [i, image] of imageRects.entries()) {
        expect(image.x).toBeGreaterThanOrEqual(table.xPx + box.xPx - 1e-8); expect(image.y).toBeGreaterThanOrEqual(table.yPx + box.yPx - 1e-8)
        expect(image.x + image.w).toBeLessThanOrEqual(table.xPx + box.xPx + box.widthPx + 1e-8); expect(image.y + image.h).toBeLessThanOrEqual(table.yPx + box.yPx + box.heightPx + 1e-8)
        for (const previous of imageRects.slice(0, i)) expect(nonoverlap(previous, image)).toBe(true)
      }
      for (const seg of page.lines.flatMap(line => line.segs).filter(seg => seg.logical?.source.text === 'Following')) {
        const t = seg.transform!, points = [[0, -12], [seg.widthPx, -12], [seg.widthPx, 4], [0, 4]].map(([x, y]) => ({ x: t.a * x + t.c * y + t.e, y: t.b * x + t.d * y + t.f }))
        const x = Math.min(...points.map(p => p.x)), y = Math.min(...points.map(p => p.y)), glyph = { x, y, w: Math.max(...points.map(p => p.x)) - x, h: Math.max(...points.map(p => p.y)) - y }
        for (const image of imageRects) expect(nonoverlap(glyph, image)).toBe(true)
      }
      const index = await buildTextIndex([{ spec: { widthPx: page.widthPx, heightPx: page.heightPx }, paint: ctx => renderPages([page], ctx) }]), lines = index.pages[0].lines, last = lines.length - 1
      expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: last, charIndex: lines[last].text.length } })).toBe('A文B\n\nFollowing')
    }
  })

  test('optional measured recording ink validates bounds, clips visible hits, and leaves later manual calls unchanged', async () => {
    const source = { text: 'ABCD' }
    const index = await buildTextIndex([{ spec: { widthPx: 100, heightPx: 100 }, paint: ctx => {
      ctx.font = '10pt "Fixed"'
      const record = (ctx as TextRecordingContext)[RECORD_TEXT]!
      ctx.save(); ctx.translate(10, 20)
      record('A', 0, 0, 10, { source, start: 0, end: 1, line: 0, ink: { ascent: 20, descent: 6 }, clip: { x: 0, y: -10, width: 10, height: 13 } })
      ctx.restore()
      record('B', 30, 30, 10, { source, start: 1, end: 2, line: 1, ink: { ascent: NaN, descent: 6 } })
      record('C', 50, 50, 10, { source, start: 2, end: 3, line: 2, ink: { ascent: 20, descent: -1 } })
      record('D', 70, 70, 10, { source, start: 3, end: 4, line: 3, ink: { ascent: Infinity, descent: 6 } })
      ctx.fillText('manual', 0, 90)
    } }])
    const spans = index.pages[0].lines.flatMap(line => line.spans), valid = spans.find(span => span.text === 'A')!
    expect(valid.placement!.top).toBe(-20); expect(valid.placement!.bottom).toBe(6)
    expect(findMatches(index, 'A')[0].rects).toEqual([{ x: 10, y: 10, width: 10, height: 13 }])
    expect(hitTest(index, 0, 15, 5, { strict: true })).toBeUndefined()
    for (const strict of [true, false]) { expect(hitTest(index, 0, 15, 5, { strict })?.lineIndex).not.toBe(0); expect(hitTest(index, 0, 15, 15, { strict })?.charIndex).toBe(1) }
    for (const [text, baseline] of [['B', 30], ['C', 50], ['D', 70], ['manual', 90]] as const) {
      const p = spans.find(span => span.text === text)!.placement!
      expect(p.top).toBeCloseTo(baseline - 10 * 96 / 72 * .85); expect(p.bottom).toBeCloseTo(baseline + 10 * 96 / 72 * .25)
      expect(p.clip).toBeUndefined(); expect(p.transform.e).toBe(0); expect(p.transform.f).toBe(0)
    }
  })

  test.each(['tbRlV', 'tbLrV'].flatMap(direction => [10, 12].flatMap(fontSizePt => ['uniform', 'glyphSpecific'].flatMap(metrics =>
    ['auto', 'atLeast', 'exact', 'merged'].map(rule => ({ direction, fontSizePt, metrics, rule }))))))('$direction $fontSizePt pt $metrics $rule reserves measured upright bands beside physical images', async ({ direction, fontSizePt, metrics, rule }) => {
    const rows = heightRow(cell(direction, ['文中B文', 'Following']).replace('<w:tcPr>', `<w:tcPr>${rule === 'merged' ? '<w:vMerge w:val="restart"/>' : ''}`), rule === 'exact' ? 900 : rule === 'merged' ? 0 : 600, rule === 'merged' ? 'exact' : rule === 'auto' ? 'atLeast' : rule)
      + (rule === 'merged' ? heightRow('<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>', 900) : '')
    const doc = await fixture(rowsTable(rows)), block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw Error('table')
    const tableSource = block.table
    tableSource.gridColsTwips = [6000]; tableSource.cellMargins = { leftTwips: 0, rightTwips: 0, topTwips: 0, bottomTwips: 0 }
    for (const row of tableSource.rows) for (const cell of row.cells) cell.margins = tableSource.cellMargins
    if (rule === 'auto') { delete tableSource.rows[0].heightTwips; delete tableSource.rows[0].heightRule }
    const para = tableSource.rows[0].cells[0].paragraphs[0]
    para.runs = ['文', '中', 'B', '文'].map((text, i) => ({ text, fontSizePt: i === 1 ? 12 : fontSizePt, fontFamily: i === 1 ? 'Other' : 'Fixed', bold: i === 1 }))
    const images = [24, 18, 12].map((width, i) => ({ data: new Uint8Array([i]), widthEmu: width * 9525, heightEmu: 24 * 9525 }))
    para.images = images
    para.inline = [{ kind: 'text', run: para.runs[0] }, { kind: 'image', image: images[0] }, { kind: 'text', run: para.runs[1] }, { kind: 'image', image: images[1] }, { kind: 'text', run: para.runs[2] }, { kind: 'image', image: images[2] }, { kind: 'text', run: para.runs[3] }]
    const glyphMetrics = (text: string, family: string) => metrics === 'glyphSpecific' && text === '文' ? { ascent: 20, descent: 6, width: 9 }
      : metrics === 'glyphSpecific' && family === 'Other' && text === '中' ? { ascent: 7, descent: 5, width: 18 }
      : { ascent: 12, descent: 4, width: [...text].length * 10 }
    const fixed: MeasureFn = Object.assign((text: string, style: Parameters<MeasureFn>[1]) => glyphMetrics(text, style.fontFamily).width,
      { metrics: (style: Parameters<MeasureFn>[1], text = 'Mg') => ({ ...glyphMetrics(text, style.fontFamily), normalHeight: 16 }) })
    const page = layoutDocx(doc, fixed)[0], table = page.tables[0], cellBox = table.rows[0].cells[0]
    expect(page.images.map(image => image.imageIndex)).toEqual([0, 1, 2])
    const paintedBands = page.lines.flatMap(line => line.segs).map(seg => {
      const m = glyphMetrics(seg.text, seg.style.fontFamily), t = seg.transform!
      const points = [[0, -m.ascent], [seg.widthPx, -m.ascent], [seg.widthPx, m.descent], [0, m.descent]].map(([x, y]) => ({ x: t.a * x + t.c * y + t.e, y: t.b * x + t.d * y + t.f }))
      return { text: seg.text, x: Math.min(...points.map(point => point.x)), y: Math.min(...points.map(point => point.y)), right: Math.max(...points.map(point => point.x)), bottom: Math.max(...points.map(point => point.y)) }
    })
    for (const [i, band] of paintedBands.entries()) {
      for (const other of paintedBands.slice(i + 1)) expect(Math.min(band.right, other.right) - Math.max(band.x, other.x) <= 1e-8 || Math.min(band.bottom, other.bottom) - Math.max(band.y, other.y) <= 1e-8, JSON.stringify({ band, other })).toBe(true)
      for (const image of page.images) expect(Math.min(band.right, image.xPx + image.widthPx) - Math.max(band.x, image.xPx) <= 1e-8 || Math.min(band.bottom, image.yPx + image.heightPx) - Math.max(band.y, image.yPx) <= 1e-8).toBe(true)
    }
    const paintedImages: number[][] = []
    const index = await buildTextIndex([{ spec: { widthPx: page.widthPx, heightPx: page.heightPx }, paint: ctx => {
      ctx.measureText = text => { const m = glyphMetrics(text, ctx.font.includes('Other') ? 'Other' : 'Fixed'); return { width: m.width, actualBoundingBoxAscent: m.ascent, actualBoundingBoxDescent: m.descent } as TextMetrics }
      ctx.drawImage = ((_image: CanvasImageSource, x: number, y: number, w: number, h: number) => { paintedImages.push([x, y, w, h]) }) as never
      renderPages([page], ctx, images.map(() => createCanvas(1, 1)) as never)
    } }])
    expect(paintedImages).toEqual(page.images.map(image => [image.xPx, image.yPx, image.widthPx, image.heightPx]))
    const lines = index.pages[0].lines, last = lines.length - 1
    expect(lines.flatMap(line => line.spans).map(span => span.text).join('')).toBe('文中B文Following')
    expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: last, charIndex: lines[last].text.length } })).toBe('文中B文\nFollowing')
    for (const [lineIndex, line] of lines.entries()) for (const span of line.spans.filter(span => span.text.length)) {
      const p = span.placement!, t = p.transform
      if (span.text === '文' || span.text === '中') {
        const m = glyphMetrics(span.text, span.text === '中' ? 'Other' : 'Fixed')
        expect(p.top).toBe(-m.ascent); expect(p.bottom).toBe(m.descent)
      }
      const points = [[p.x, p.top], [p.x + p.width, p.top], [p.x + p.width, p.bottom], [p.x, p.bottom]].map(([x, y]) => ({ x: t.a * x + t.c * y + t.e, y: t.b * x + t.d * y + t.f }))
      const x = Math.min(...points.map(point => point.x)), y = Math.min(...points.map(point => point.y)), right = Math.max(...points.map(point => point.x)), bottom = Math.max(...points.map(point => point.y))
      for (const image of page.images) expect(Math.min(right, image.xPx + image.widthPx) - Math.max(x, image.xPx) <= 1e-8 || Math.min(bottom, image.yPx + image.heightPx) - Math.max(y, image.yPx) <= 1e-8, JSON.stringify({ text: span.text, x, y, right, bottom, image })).toBe(true)
      expect(y).toBeGreaterThanOrEqual(table.yPx + cellBox.yPx - 1e-8); expect(bottom).toBeLessThanOrEqual(table.yPx + cellBox.yPx + cellBox.heightPx + 1e-8)
      const px = t.a * (p.x + p.width * .75) + t.c * ((p.top + p.bottom) / 2) + t.e, py = t.b * (p.x + p.width * .75) + t.d * ((p.top + p.bottom) / 2) + t.f
      for (const strict of [true, false]) {
        const hit = hitTest(index, 0, px, py, { strict }); expect(hit?.lineIndex).toBe(lineIndex)
        if (span.text === '文' || span.text === '中') expect(hit?.charIndex).toBe(line.spans.slice(0, line.spans.indexOf(span)).reduce((sum, previous) => sum + previous.text.length, 0) + 1)
      }
    }
    for (const image of page.images) {
      expect(image.widthPx).toBe(images[image.imageIndex].widthEmu / 9525); expect(image.heightPx).toBe(24)
      expect(image.clip).toEqual({ xPx: table.xPx + cellBox.xPx, yPx: table.yPx + cellBox.yPx, widthPx: cellBox.widthPx, heightPx: cellBox.heightPx })
    }
    expect(page.lines.flatMap(line => line.segs).find(seg => seg.text === '文')!.style.fontSizePt).toBe(fontSizePt)
    expect(page.lines.flatMap(line => line.segs).find(seg => seg.text === '中')!.style.fontFamily).toBe('Other')
    if (rule === 'exact' || rule === 'merged') expect(cellBox.heightPx).toBe(60)
  })

  test.each(['tbRl', 'btLr', 'tbRlV', 'tbLrV', 'lrTbV'])('%s keeps an authored image at an interrupted combining sequence without partial source carets', async direction => {
    const doc = await fixture(rowsTable(heightRow(cell(direction, 'A\u0301B'), 6000))), block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw Error('table')
    const para = block.table.rows[0].cells[0].paragraphs[0], image = { data: new Uint8Array([1]), widthEmu: 24 * 9525, heightEmu: 24 * 9525 }
    para.images = [image]; para.inline = [{ kind: 'text', run: { text: 'A', fontSizePt: 12 } }, { kind: 'image', image }, { kind: 'text', run: { text: '\u0301B', fontSizePt: 12 } }]
    const fixed: MeasureFn = Object.assign((text: string) => [...text].length * 10, { metrics: () => ({ ascent: 12, descent: 4, normalHeight: 16 }) }), page = layoutDocx(doc, fixed)[0]
    expect(page.images.map(image => image.imageIndex)).toEqual([0])
    const index = await buildTextIndex([{ spec: { widthPx: page.widthPx, heightPx: page.heightPx }, paint: ctx => renderPages([page], ctx) }]), lines = index.pages[0].lines, last = lines.length - 1
    expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: last, charIndex: lines[last].text.length } })).toBe('A\u0301B')
    for (const [lineIndex, line] of lines.entries()) for (const span of line.spans.filter(span => span.text.length)) {
      const p = span.placement!, t = p.transform, x = t.a * (p.x + p.width * .7) + t.c * p.y + t.e, y = t.b * (p.x + p.width * .7) + t.d * p.y + t.f
      const hit = hitTest(index, 0, x, y, { strict: true })
      expect(hit?.lineIndex).toBe(lineIndex)
      expect([0, 2, 3]).toContain(hit?.charIndex)
    }
  })

  test('genuinely oversized directional cell images retain size and clip to the final cell interior', async () => {
    const doc = await fixture(rowsTable(heightRow(cell('lrTbV', ''), 600))), block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw Error('table')
    block.table.gridColsTwips = [600]; block.table.cellMargins = { leftTwips: 0, rightTwips: 0, topTwips: 0, bottomTwips: 0 }; block.table.rows[0].cells[0].margins = block.table.cellMargins
    const image = { data: new Uint8Array([1]), widthEmu: 80 * 9525, heightEmu: 60 * 9525 }, para = block.table.rows[0].cells[0].paragraphs[0]
    para.images = [image]; para.inline = [{ kind: 'image', image }]
    const fixed: MeasureFn = Object.assign((text: string) => [...text].length * 10, { metrics: () => ({ ascent: 12, descent: 4, normalHeight: 16 }) }), page = layoutDocx(doc, fixed)[0]
    expect(page.images.map(image => [image.widthPx, image.heightPx])).toEqual([[80, 60]])
    expect(page.tables[0].rows[0].heightPx).toBe(40)
    const raster = createCanvas(80, 60), rasterCtx = raster.getContext('2d'); rasterCtx.fillStyle = '#ff0000'; rasterCtx.fillRect(0, 0, 80, 60)
    const ctx = createCanvas(page.widthPx, page.heightPx).getContext('2d'); ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, page.widthPx, page.heightPx)
    renderPages([page], ctx as never, [raster] as never)
    const table = page.tables[0]
    expect([...ctx.getImageData(table.xPx + 20, table.yPx + 20, 1, 1).data]).toEqual([255, 0, 0, 255])
    expect([...ctx.getImageData(table.xPx + 50, table.yPx + 20, 1, 1).data]).toEqual([255, 255, 255, 255])
    expect([...ctx.getImageData(table.xPx + 20, table.yPx + 50, 1, 1).data]).toEqual([255, 255, 255, 255])
  })

  test('lrTbV measures each rotated glyph band rather than using Latin Mg width for wrapping', async () => {
    const doc = await fixture(rowsTable(heightRow(cell('lrTbV', '文文'), 0, 'atLeast'))), block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw Error('table')
    block.table.gridColsTwips = [600]; block.table.cellMargins = { leftTwips: 0, rightTwips: 0, topTwips: 0, bottomTwips: 0 }; block.table.rows[0].cells[0].margins = block.table.cellMargins
    const fixed: MeasureFn = Object.assign((text: string) => [...text].length * 10, { metrics: (_style: unknown, text = 'Mg') => ({ ascent: text === '文' ? 18 : 12, descent: text === '文' ? 6 : 4, normalHeight: 16 }) })
    const page = layoutDocx(doc, fixed)[0], table = page.tables[0]
    expect(page.lines.map(line => line.segs.map(seg => seg.text).join(''))).toEqual(['文', '文'])
    for (const line of page.lines) {
      const han = line.segs[0], t = han.transform!
      expect(t.e - 18).toBeCloseTo(table.xPx)
      expect(t.e + 6).toBeCloseTo(table.xPx + 24)
      expect(t.f).toBeGreaterThanOrEqual(line.yPx)
      expect(t.f + han.widthPx).toBeLessThanOrEqual(line.yPx + line.heightPx)
    }
  })

  test('lrTbV wraps using projected CJK width and records bounded nonoverlapping lines and carets', async () => {
    const doc = await fixture(rowsTable(heightRow(cell('lrTbV', ['A文B文CD', 'Following']), 0, 'atLeast'))), block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw Error('table')
    block.table.cellMargins = { leftTwips: 0, rightTwips: 0, topTwips: 0, bottomTwips: 0 }
    for (const row of block.table.rows) for (const cell of row.cells) cell.margins = block.table.cellMargins
    block.table.gridColsTwips = [600]
    const fixed: MeasureFn = Object.assign((text: string) => [...text].length * 10, { metrics: () => ({ ascent: 12, descent: 4, normalHeight: 16 }) })
    const page = layoutDocx(doc, fixed)[0], table = page.tables[0], box = table.rows[0].cells[0]
    const first = page.lines.filter(line => line.logical?.source.text === 'A文B文CD')
    expect(first.map(line => line.segs.map(seg => seg.text).join(''))).toEqual(['A文B', '文CD'])
    for (const line of page.lines) for (const seg of line.segs) {
      const t = seg.transform!, points = [[0, -12], [seg.widthPx, -12], [seg.widthPx, 4], [0, 4]].map(([x, y]) => ({ x: t.a * x + t.c * y + t.e, y: t.b * x + t.d * y + t.f }))
      expect(Math.min(...points.map(p => p.x))).toBeGreaterThanOrEqual(table.xPx - 1e-8)
      expect(Math.max(...points.map(p => p.x))).toBeLessThanOrEqual(table.xPx + box.widthPx + 1e-8)
      expect(Math.min(...points.map(p => p.y))).toBeGreaterThanOrEqual(line.yPx - 1e-8)
      expect(Math.max(...points.map(p => p.y))).toBeLessThanOrEqual(line.yPx + line.heightPx + 1e-8)
    }
    const index = await buildTextIndex([{ spec: { widthPx: page.widthPx, heightPx: page.heightPx }, paint: ctx => renderPages([page], ctx) }]), lines = index.pages[0].lines, last = lines.length - 1
    expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: last, charIndex: lines[last].text.length } })).toBe('A文B文CD\nFollowing')
    const span = lines[0].spans.find(span => span.text === '文')!, p = span.placement!, t = p.transform, x = t.a * (p.x + p.width * .7) + t.c * p.y + t.e, y = t.b * (p.x + p.width * .7) + t.d * p.y + t.f
    for (const strict of [true, false]) expect(hitTest(index, 0, x, y, { strict })?.charIndex).toBe(2)
  })

  test.each(['auto', 'atLeast', 'exact'] as const)('lrTbV %s line metrics include the projected East Asian band with authored spacing', async rule => {
    const doc = await fixture(rowsTable(heightRow(cell('lrTbV', ['A文', 'BC']), 0, 'atLeast')))
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw Error('table')
    block.table.cellMargins = { leftTwips: 0, rightTwips: 0, topTwips: 0, bottomTwips: 0 }
    for (const row of block.table.rows) for (const cell of row.cells) cell.margins = block.table.cellMargins
    const paras = block.table.rows[0].cells[0].paragraphs
    paras[0].lineSpacing = { rule, value: rule === 'auto' ? 240 : 120 }
    paras[0].spacingAfterTwips = 150; paras[1].spacingBeforeTwips = 60
    const fixed: MeasureFn = Object.assign((text: string) => [...text].reduce((sum, char) => sum + (char === '文' ? 42 : 10), 0), { metrics: () => ({ ascent: 12, descent: 4, normalHeight: 16 }) })
    const page = layoutDocx(doc, fixed)[0]
    const han = page.lines[0].segs.find(seg => seg.text === '文')!, next = page.lines[1].segs[0]
    const nextTop = next.transform!.f - 12
    if (rule === 'exact') { expect(page.lines[0].heightPx).toBe(8); expect(nextTop).toBeCloseTo(page.tables[0].yPx + 8 + 10 + 4) }
    else {
      expect(page.lines[0].heightPx).toBeGreaterThanOrEqual(42)
      expect(han.transform!.f).toBeGreaterThanOrEqual(page.lines[0].yPx)
      expect(han.transform!.f + han.widthPx).toBeLessThanOrEqual(page.lines[0].yPx + page.lines[0].heightPx)
      expect(nextTop).toBeGreaterThanOrEqual(han.transform!.f + han.widthPx + 14)
    }
    expect(page.tables[0].rows[0].heightPx).toBeGreaterThanOrEqual(page.lines[1].yPx + page.lines[1].heightPx - page.tables[0].yPx)
  })

  test.each(['tbRl', 'btLr', 'tbRlV', 'tbLrV', 'lrTbV'])('%s preserves inline picture payload and paint order through final merged projection', async direction => {
    const doc = await fixture(rowsTable(heightRow(bottomUpCell(['BeforeAfter', ''], '<w:vMerge w:val="restart"/>').replace('w:val="btLr"', `w:val="${direction}"`), 150) + heightRow('<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>', 3000)))
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw Error('table')
    const paragraphs = block.table.rows[0].cells[0].paragraphs
    const images = [24, 36, 12].map((width, i) => ({ data: new Uint8Array([i + 1]), widthEmu: width * 9525, heightEmu: (12 + i) * 9525 }))
    paragraphs[0].images = images.slice(0, 2)
    paragraphs[0].inline = [{ kind: 'text', run: { text: 'Before' } }, { kind: 'image', image: images[0] }, { kind: 'image', image: images[1] }, { kind: 'text', run: { text: 'After' } }]
    paragraphs[1].images = [images[2]]
    paragraphs[1].inline = [{ kind: 'image', image: images[2] }]
    const page = layoutDocx(doc, measure)[0]
    expect(page.images.map(image => image.imageIndex)).toEqual([0, 1, 2])
    expect(page.images.map(image => [image.widthPx, image.heightPx])).toEqual([[24, 12], [36, 13], [12, 14]])
    expect(page.images.every(image => Number.isFinite(image.xPx) && Number.isFinite(image.yPx))).toBe(true)
    const decoded = images.map(() => createCanvas(2, 2))
    const ctx = createCanvas(page.widthPx, page.heightPx).getContext('2d')
    const painted: unknown[] = []
    ctx.drawImage = ((image: unknown) => painted.push(image)) as never
    renderPages([page], ctx as never, decoded as never)
    expect(painted).toEqual(decoded)
  })

  test.each(['tbRl', 'btLr', 'tbRlV', 'tbLrV', 'lrTbV'])('%s retains later paragraph source styles in layout and actual recording', async direction => {
    const doc = await fixture(rowsTable(heightRow(cell(direction, ['A', 'CD']), 6000)))
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw Error('table')
    Object.assign(block.table.rows[0].cells[0].paragraphs[0].runs[0], { fontSizePt: 10, fontFamily: 'First', color: 'FF0000', bold: true })
    const second = block.table.rows[0].cells[0].paragraphs[1].runs[0]
    Object.assign(second, { fontSizePt: 20, fontFamily: 'Second', color: '0000FF', italic: true, underline: true, highlight: 'yellow' })
    const page = layoutDocx(doc, measure)[0]
    const segs = page.lines.flatMap(line => line.segs).filter(seg => /[CD]/.test(seg.text))
    expect(segs.map(seg => seg.text).join('')).toBe('CD')
    for (const seg of segs) {
      expect(seg.run).toBe(second)
      expect(seg.style).toMatchObject({ fontSizePt: 20, fontFamily: 'Second', italic: true, bold: false })
      expect(seg.run).toMatchObject({ color: '0000FF', underline: true, highlight: 'yellow' })
    }
    const index = await buildTextIndex([{ spec: { widthPx: page.widthPx, heightPx: page.heightPx }, paint: ctx => renderPages([page], ctx) }])
    const recorded = index.pages[0].lines.flatMap(line => line.spans).filter(span => /[CD]/.test(span.text))
    expect(recorded.map(span => span.text).join('')).toBe('CD')
    for (const span of recorded) expect(span.fontSize).toBeCloseTo(20 * 96 / 72)
  })

  test.each(['tbRl', 'btLr', 'tbRlV', 'tbLrV', 'lrTbV'])('%s resolves paragraph spacing and line rules including an empty paragraph', async direction => {
    const doc = await fixture(rowsTable(heightRow(cell(direction, ['A', '', 'B']), 6000)))
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw Error('table')
    const paras = block.table.rows[0].cells[0].paragraphs
    Object.assign(paras[0], { spacingBeforeTwips: 60, spacingAfterTwips: 150, lineSpacing: { rule: 'auto', value: 480 } })
    Object.assign(paras[1], { spacingBeforeTwips: 30, spacingAfterTwips: 90, lineSpacing: { rule: 'exact', value: 300 } })
    Object.assign(paras[2], { spacingBeforeTwips: 120, lineSpacing: { rule: 'atLeast', value: 120 } })
    const controlled: MeasureFn = Object.assign(measure, { metrics: () => ({ ascent: 12, descent: 4, normalHeight: 16 }) })
    try {
      const page = layoutDocx(doc, controlled)[0]
      expect(page.lines).toHaveLength(3)
      const a = page.lines[0].segs[0].transform!, b = page.lines[2].segs[0].transform!
      // First advance32, after10, empty before2/advance20/after6, next before8;
      // baseline difference also composes the two lines' own leading halves.
      const gap = direction === 'lrTbV' ? b.f - a.f : Math.abs(b.e - a.e)
      expect(gap).toBeCloseTo(32 + 10 + 2 + 20 + 6 + 8 - 8)
    } finally { delete controlled.metrics }
  })

  test('single-column tbLrV anchors its upright glyph at the left content edge', async () => {
    const doc = await fixture(rowsTable(heightRow(cell('tbLrV', '文'), 6000)))
    const page = layoutDocx(doc, measure)[0], box = cellBoxOf(doc)
    const wen = page.lines[0].segs[0].transform!
    expect([wen.a, wen.d]).toEqual([1, 1])
    expect(wen.e).toBeCloseTo(box.xPx + 7.2)
  })

  test.each(['tbRl', 'tbRlV', 'tbLrV'])('%s rewraps against a completed merge, including a zero-height anchor', async direction => {
    for (const height of [0, 600]) {
      const doc = await fixture(rowsTable(heightRow(cell(direction, 'ABCDEFGHIJ').replace('<w:tcPr>', '<w:tcPr><w:vMerge w:val="restart"/>'), height) + heightRow('<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>', 5400) + heightRow(cell(direction, 'Z'), 600)) + '<w:p><w:r><w:t>after</w:t></w:r></w:p>')
      const page = layoutDocx(doc, measure)[0], box = cellBoxOf(doc)
      expect(page.lines.map(line => line.segs.map(seg => seg.text).join(''))).toEqual(['ABCDEFGHIJ', 'Z', 'after'])
      expect(page.lines[0].segs[0].transform!.f).toBeCloseTo(box.yPx)
      expect(page.lines[1].segs[0].transform!.f).toBeCloseTo(box.yPx + box.heightPx)
    }
  })

  test.each(['tbRl', 'btLr', 'lrTbV'])('%s preserves shaped authored cell runs in actual paint', async direction => {
    for (const text of ['مرحبا', 'नमस्ते', 'AV']) {
      const doc = await fixture(rowsTable(heightRow(cell(direction, text), 6000)))
      const page = layoutDocx(doc, measure)[0]
      const ctx = createCanvas(816, 1056).getContext('2d')
      const calls: string[] = []
      ctx.fillText = text => { calls.push(text) }
      renderPages([page], ctx as never)
      expect(calls).toEqual([text])
    }
  })

  test.each(['tbRl', 'btLr', 'tbRlV', 'tbLrV', 'lrTbV'])('%s records authored paragraph/run source order through mixed rotations and unequal lengths', async direction => {
    const doc = await fixture(rowsTable(heightRow(cell(direction, ['A文', 'BCDE']), 6000)))
    const page = layoutDocx(doc, measure)[0]
    const index = await buildTextIndex([{ spec: { widthPx: page.widthPx, heightPx: page.heightPx }, paint: ctx => renderPages([page], ctx) }])
    expect(index.pages[0].lines.map(line => line.text)).toEqual(['A文', 'BCDE'])
    expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: 1, charIndex: 4 } })).toBe('A文\nBCDE')
    expect(findMatches(index, 'A文')).toHaveLength(1)
    const source = index.pages[0].lines[0].spans[0].logical?.source
    expect(source?.text).toBe('A文')
    expect(index.pages[0].lines[0].spans.every(span => span.logical?.source === source)).toBe(true)
  })

  test('lrTbV rotates East Asian clusters but keeps Latin horizontal', async () => {
    const doc = await fixture(table(cell('lrTbV', 'A中')))
    const lines = linesIn(doc)
    const segs = lines.flatMap(l => l.segs)
    const latin = segs.find(s => s.text === 'A')!
    const han = segs.find(s => s.text === '中')!
    // Every cluster carries an absolute origin: identity for Latin kept in
    // horizontal flow, 90 CW for upright-form East Asian.
    expect(latin.transform).toMatchObject({ a: 1, b: 0, c: 0, d: 1 })
    expect(han.transform?.a ?? 99).toBeCloseTo(0)
    expect(han.transform?.b ?? 99).toBeCloseTo(1)
    expect(han.transform?.c ?? 99).toBeCloseTo(-1)
    expect(han.transform?.d ?? 99).toBeCloseTo(0)
  })

  test('lrTbV keeps justified gaps gap-exact through affine projection', async () => {
    const justified = `<w:tc><w:tcPr><w:textDirection w:val="lrTbV"/><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>AA BB CC DD EE FF GG</w:t></w:r></w:p></w:tc>`
    const ragged = justified.replace('w:val="both"', 'w:val="left"')
  const rightEdge = async (body: string) => {
      const doc = await fixture(table(body))
      const page = layoutDocx(doc, measure)[0]
      const box = cellBoxOf(doc)
      const right = box.xPx + box.widthPx
      let max = -Infinity, lines = 0
      for (const line of page.lines) {
        if (line.segs.length === 0) continue
        let edge = -Infinity
        for (const seg of line.segs) if (seg.transform) edge = Math.max(edge, seg.transform.e + seg.widthPx)
        if (edge > -Infinity) { lines++; max = Math.max(max, edge) }
      }
      return { lines, max, right }
    }
    const just = await rightEdge(justified)
    const rag = await rightEdge(ragged)
    expect(just.lines).toBeGreaterThan(1)
    // Justified non-end lines fill the content width (108-twips right margin
    // is 7.2px); left-aligned do not.
    expect(Math.abs(just.max - (just.right - 7.2))).toBeLessThan(3)
    expect(rag.max).toBeLessThan(just.max - 5)
  })

  test('later paragraphs keep their own indents in vertical cells', async () => {
    const indented = (direction: string, leftTwips: string | undefined) =>
      `<w:tc><w:tcPr><w:textDirection w:val="${direction}"/><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>` +
      `<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>AA</w:t></w:r></w:p>` +
      `<w:p>${leftTwips === undefined ? '' : `<w:pPr><w:ind w:left="${leftTwips}"/></w:pPr>`}<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>BB</w:t></w:r></w:p></w:tc>`
    {
      // Horizontal flow: the second paragraph shifts right by its own indent.
      const doc = await fixture(table(indented('lrTbV', '720')))
      const page = layoutDocx(doc, measure)[0]
      const first = (text: string) => glyphsOf(page).find(s => s.text === text && s.transform)?.transform
      const a = first('A')!, b = first('B')!
      // 720 twips is 48px; the second paragraph shifts by exactly that.
      expect(b.e - a.e).toBeCloseTo(48, 0)
    }
    {
      // Vertical RTL flow: the transverse indent that moves columns is the
      // right indent; compare against an unindented control cell.
      const one = async (ind: string | undefined) => {
        const doc = await fixture(table(
          `<w:tc><w:tcPr><w:textDirection w:val="tbRl"/><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>` +
          `<w:p>${ind === undefined ? '' : `<w:pPr><w:ind w:right="${ind}"/></w:pPr>`}<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>BB</w:t></w:r></w:p></w:tc>`
        ))
        const page = layoutDocx(doc, measure)[0]
        return glyphsOf(page).find(s => s.text === 'B' && s.transform)?.transform!.e
      }
      const plain = (await one(undefined))!
      const shifted = (await one('720'))!
      expect(plain - shifted).toBeCloseTo(48, 0)
    }
  })

  test('vertical cell content grows merged row height and following flow', async () => {
    const doc = await fixture(`${table(cell('tbRl', 'AAAA BBBB CCCC DDDD EEEE FFFF'))}<w:p><w:r><w:t>after</w:t></w:r></w:p>`)
    const page = layoutDocx(doc, measure)[0]
    const row = page.tables[0].rows[0]
    expect(row.heightPx).toBeGreaterThan(40)
    const after = page.lines.find(l => l.segs.some(s => s.text === 'after'))!
    expect(after.yPx).toBeGreaterThanOrEqual(row.yPx + row.heightPx)
  })

  test('tbRl paints ink stacked inside the cell box', async () => {
    const doc = await fixture(table(cell('tbRl', '中文测试文字内容')))
    const box = cellBoxOf(doc)
    const ctx = createCanvas(816, 1056).getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, 816, 1056)
    renderPages(layoutDocx(doc, measure), ctx as never)
    const data = ctx.getImageData(0, 0, 816, 1056).data
    let count = 0, top = 1056, bottom = -1
    for (let y = 0; y < 1056; y++) for (let x = 0; x < 816; x++) {
      if (data[(y * 816 + x) * 4] < 128 && x >= box.xPx && x < box.xPx + box.widthPx && y >= box.yPx && y < box.yPx + box.heightPx) {
        count++
        top = Math.min(top, y); bottom = Math.max(bottom, y)
      }
    }
    expect(count).toBeGreaterThan(0)
    expect(bottom - top).toBeGreaterThan(box.heightPx / 2)
  })
})
