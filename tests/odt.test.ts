/**
 * ODT (OpenDocument text) adapter tests: Writer content parsed into the
 * shared flow model, exercising detection through loadOfficeFile.
 */
import { describe, test, expect } from 'vitest'
import { loadOfficeFile } from '../src/components/OfficeFile'
import { parseOdt } from '../src/odt/parse'
import { layoutDocx } from '../src/docx/layout'
import type { MeasureFn } from '../src/docx/layout'
import { OfficePackage } from '../src/core/zip'
import { buildOdt, odfBulletListStyle, odfDecimalListStyle } from '../src/testdata/odf-builders'
import type { DocxDocument } from '../src/docx/types'

async function odtDoc(opts: Parameters<typeof buildOdt>[0]): Promise<DocxDocument> {
  return parseOdt(await OfficePackage.load(await buildOdt(opts)))
}

function textOf(doc: DocxDocument): string {
  return doc.sections.flatMap((s) => s.blocks).map((b) =>
    b.kind === 'p' ? b.paragraph.runs.map((r) => r.text).join('') : '[table]',
  ).join('\n')
}

describe('odt detection and paragraphs', () => {
  test('hello world through loadOfficeFile', async () => {
    const doc = await loadOfficeFile(await buildOdt({ paras: [{ text: 'hello odt' }] }))
    expect('sections' in doc).toBe(true)
    const odt = doc as DocxDocument
    expect(odt.sections).toHaveLength(1)
    expect(textOf(odt)).toBe('hello odt')
  })

  test('a spreadsheet ODF is detected but refused with a clear seam message', async () => {
    const { default: JSZip } = await import('jszip')
    const zip = await JSZip.loadAsync(await buildOdt({ paras: [{ text: 'x' }] }))
    zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet')
    const buf = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
    await expect(loadOfficeFile(buf)).rejects.toThrow('only word-processing (.odt)')
  })

  test('mixed-content run order survives (bold mid-sentence)', async () => {
    const doc = await odtDoc({
      paras: [{ runs: ['hello ', { text: 'bold', bold: true }, ' world'] }],
    })
    const para = doc.sections[0].blocks[0]
    expect(para.kind).toBe('p')
    if (para.kind !== 'p') return
    expect(para.paragraph.runs.map((r) => r.text)).toEqual(['hello ', 'bold', ' world'])
    expect(para.paragraph.runs[1].bold).toBe(true)
    expect(para.paragraph.runs[0].bold).toBeUndefined()
  })

  test('run formatting maps (italic, size, color, background)', async () => {
    const doc = await odtDoc({
      paras: [{ runs: [{ text: 'fancy', italic: true, sizePt: 18, color: 'FF0000', background: 'FFFF00' }] }],
    })
    const run = (doc.sections[0].blocks[0] as { kind: 'p'; paragraph: { runs: Array<{ italic?: boolean; fontSizePt?: number; color?: string; highlight?: string }> } }).paragraph.runs[0]
    expect(run.italic).toBe(true)
    expect(run.fontSizePt).toBe(18)
    expect(run.color).toBe('FF0000')
    expect(run.highlight).toBe('#FFFF00')
  })

  test('paragraph alignment and spacing map', async () => {
    const doc = await odtDoc({
      paras: [{ text: 'centered', align: 'center', spacingBeforePt: 6, spacingAfterPt: 12 }],
    })
    const para = (doc.sections[0].blocks[0] as { kind: 'p'; paragraph: import('../src/docx/types').DocxParagraph }).paragraph
    expect(para.align).toBe('center')
    expect(para.spacingBeforeTwips).toBeCloseTo(120, 0)
    expect(para.spacingAfterTwips).toBeCloseTo(240, 0)
  })

  test('headings carry outline levels', async () => {
    const doc = await odtDoc({ paras: [{ text: 'Title', heading: 1 }, { text: 'Sub', heading: 2 }] })
    const paras = doc.sections[0].paragraphs
    expect(paras.map((p) => p.outlineLevel)).toEqual([1, 2])
    expect(textOf(doc)).toBe('Title\nSub')
  })

  test('named parent styles chain (alignment inherited)', async () => {
    const doc = await odtDoc({
      extraCommonStyles:
        '<style:style style:name="Base" style:family="paragraph">'
        + '<style:paragraph-properties fo:text-align="right"/></style:style>'
        + '<style:style style:name="Child" style:family="paragraph" style:parent-style-name="Base">'
        + '<style:text-properties fo:font-weight="bold"/></style:style>',
      paras: [{ runs: [{ text: 'chained' }], style: 'Child' }],
    })
    const para = (doc.sections[0].blocks[0] as { kind: 'p'; paragraph: import('../src/docx/types').DocxParagraph }).paragraph
    expect(para.align).toBe('right')
    expect(para.runs[0].bold).toBe(true)
  })
})

describe('odt lists', () => {
  test('bullet list markers per level', async () => {
    const doc = await odtDoc({
      extraAutoStyles: odfBulletListStyle('BL'),
      lists: [{
        styleName: 'BL',
        items: [{ text: 'one' }, { text: 'two', nested: { styleName: 'BL', items: [{ text: 'nested' }] } }],
      }],
    })
    const markers = doc.sections[0].paragraphs.map((p) => [p.listMarker, p.listLevel])
    expect(markers).toEqual([['•', 0], ['•', 0], ['–', 1]])
  })

  test('decimal counters advance and restart per list', async () => {
    const doc = await odtDoc({
      extraAutoStyles: odfDecimalListStyle('DL'),
      lists: [
        { styleName: 'DL', items: [{ text: 'a' }, { text: 'b' }] },
        { styleName: 'DL', items: [{ text: 'c' }] },
      ],
    })
    expect(doc.sections[0].paragraphs.map((p) => p.listMarker)).toEqual(['1.', '2.', '1.'])
  })

  test('continue-numbering carries the counter across lists', async () => {
    const doc = await odtDoc({
      extraAutoStyles: odfDecimalListStyle('DL'),
      lists: [
        { styleName: 'DL', items: [{ text: 'a' }, { text: 'b' }] },
        { styleName: 'DL', continueNumbering: true, items: [{ text: 'c' }] },
      ],
    })
    expect(doc.sections[0].paragraphs.map((p) => p.listMarker)).toEqual(['1.', '2.', '3.'])
  })
})

describe('odt fields and whitespace', () => {
  test('page fields tag runs with cached text', async () => {
    const doc = await odtDoc({ paras: [{ runs: ['p', { text: '7', field: 'PAGE' }, ' of ', { text: '12', field: 'NUMPAGES' }] }] })
    const runs = (doc.sections[0].blocks[0] as { kind: 'p'; paragraph: { runs: Array<{ text: string; field?: string }> } }).paragraph.runs
    expect(runs.map((r) => [r.text, r.field])).toEqual([['p', undefined], ['7', 'PAGE'], [' of ', undefined], ['12', 'NUMPAGES']])
  })

  test('tabs, extra spaces and line breaks survive', async () => {
    const doc = await odtDoc({
      paras: [{ runs: ['a', { text: '', raw: '<text:tab/>' }, 'b', { text: '', raw: '<text:s text:c="2"/>' }, 'c'] }],
    })
    const runs = (doc.sections[0].blocks[0] as { kind: 'p'; paragraph: { runs: Array<{ text: string }> } }).paragraph.runs
    expect(runs.map((r) => r.text)).toEqual(['a', '\t', 'b', '  ', 'c'])
  })

  test('a line break flags the next run', async () => {
    const doc = await odtDoc({ paras: [{ runs: ['one', { text: '', raw: '<text:line-break/>' }, 'two'] }] })
    const runs = (doc.sections[0].blocks[0] as { kind: 'p'; paragraph: { runs: Array<{ text: string; breakBefore?: boolean }> } }).paragraph.runs
    expect(runs.map((r) => r.text)).toEqual(['one', 'two'])
    expect(runs[1].breakBefore).toBe(true)
  })
})

describe('odt tables', () => {
  test('grid, fills, alignment and borders map', async () => {
    const doc = await odtDoc({
      tables: [{
        colWidths: ['3cm', '2cm'],
        tableBorders: 'fo:border-top="0.5pt solid #000000"',
        cellBorders: 'fo:border-left="0.5pt solid #112233"',
        rows: [{
          cells: [
            { text: 'A1', fill: 'FFCC00', vAlign: 'middle' },
            { text: 'B1', gridSpan: 1 },
          ],
        }],
      }],
    })
    const block = doc.sections[0].blocks[0]
    expect(block.kind).toBe('table')
    if (block.kind !== 'table') return
    expect(block.table.gridColsTwips.map(Math.round)).toEqual([1701, 1134])
    expect(block.table.borders?.top).toMatchObject({ style: 'single', color: '000000' })
    const [a1, b1] = block.table.rows[0].cells
    expect(a1.fill).toBe('FFCC00')
    expect(a1.vAlign).toBe('center')
    expect(a1.borders?.left).toMatchObject({ style: 'single', color: '112233' })
    expect(b1.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('')).toBe('B1')
  })

  test('spans, covered cells and header rows map', async () => {
    const doc = await odtDoc({
      tables: [{
        colWidths: ['2cm', '2cm', '2cm'],
        headerRows: [{ cells: [{ text: 'H1' }, { text: 'H2' }, { text: 'H3' }] }],
        rows: [
          { cells: [{ text: 'wide', gridSpan: 2 }, { text: 'C1' }] },
          { cells: [{ text: 'tall' }, { covered: true }, { text: 'C2' }] },
        ],
      }],
    })
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw new Error('expected a table')
    expect(block.table.rows[0].isHeader).toBe(true)
    expect(block.table.rows[1].cells[0].gridSpan).toBe(2)
    const [, covered] = block.table.rows[2].cells
    expect(covered.vMerge).toBe('continue')
  })
})

describe('odt images', () => {
  async function pngBytes(): Promise<Uint8Array> {
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(24, 16)
    const ctx = canvas.getContext('2d')!
    ctx.fillStyle = '#ff0000'
    ctx.fillRect(0, 0, 24, 16)
    return new Uint8Array(canvas.toBuffer('image/png'))
  }

  test('as-char frames stay inline with EMU extents', async () => {
    const png = await pngBytes()
    const doc = await odtDoc({
      pictures: { 'Pictures/img1.png': png },
      paras: [{ text: 'see', frames: [{ picture: 'Pictures/img1.png', widthCm: 2, heightCm: 1 }] }],
    })
    const para = (doc.sections[0].blocks[0] as { kind: 'p'; paragraph: { images: Array<{ widthEmu: number; heightEmu: number; floating?: unknown }> } }).paragraph
    expect(para.images).toHaveLength(1)
    expect(para.images[0].widthEmu).toBeCloseTo(720000, 0)
    expect(para.images[0].heightEmu).toBeCloseTo(360000, 0)
    expect(para.images[0].floating).toBeUndefined()
  })

  test('paragraph-anchored frames float with offsets and z-order', async () => {
    const png = await pngBytes()
    const doc = await odtDoc({
      pictures: { 'Pictures/img1.png': png },
      paras: [{
        text: 'anchor',
        frames: [{ picture: 'Pictures/img1.png', widthCm: 2, heightCm: 1, anchor: 'paragraph', xCm: 3, yCm: 4, wrap: 'parallel', zIndex: 5 }],
      }],
    })
    const images = (doc.sections[0].blocks[0] as { kind: 'p'; paragraph: { images: Array<{ floating?: { wrap: string; behindDoc: boolean; relativeHeight: number; posH: { offsetEmu: number }; posV: { offsetEmu: number } } }> } }).paragraph.images
    expect(images).toHaveLength(1)
    expect(images[0].floating).toMatchObject({ wrap: 'square', behindDoc: false, relativeHeight: 5 })
    expect(images[0].floating?.posH.offsetEmu).toBeCloseTo(1080000, 0)
    expect(images[0].floating?.posV.offsetEmu).toBeCloseTo(1440000, 0)
  })

  test('run-through background floats behind the text', async () => {
    const png = await pngBytes()
    const doc = await odtDoc({
      pictures: { 'Pictures/img1.png': png },
      paras: [{
        text: 'anchor',
        frames: [{ picture: 'Pictures/img1.png', widthCm: 2, heightCm: 1, anchor: 'paragraph', wrap: 'run-through', behind: true }],
      }],
    })
    const images = (doc.sections[0].blocks[0] as { kind: 'p'; paragraph: { images: Array<{ floating?: { wrap: string; behindDoc: boolean } }> } }).paragraph.images
    expect(images[0].floating).toMatchObject({ wrap: 'through', behindDoc: true })
  })
})

describe('odt sections and headers', () => {
  test('a master-page change starts a section with its layout', async () => {
    const doc = await odtDoc({
      pageLayouts: { L1: { widthCm: 15, heightCm: 20, marginCm: 1 } },
      masters: { M1: { layout: 'L1', header: [{ text: 'HD' }], footer: [{ text: 'FT' }] } },
      paras: [{ text: 'first' }, { text: 'second', masterPage: 'M1' }],
    })
    expect(doc.sections).toHaveLength(2)
    const [plain, styled] = doc.sections
    expect(plain.pageSize.widthTwips).toBeCloseTo(11906, 0)
    expect(styled.pageSize.widthTwips).toBeCloseTo(8504, 0)
    expect(styled.margins.leftTwips).toBeCloseTo(567, 0)
    expect(styled.header?.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['HD'])
    expect(styled.footer?.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['FT'])
    expect(plain.header).toBeUndefined()
  })
})

describe('odt outline numbering', () => {
  const OUTLINE =
    '<text:outline-style>'
    + '<text:outline-level-style text:level="1" style:num-format="1" style:num-suffix="."/>'
    + '<text:outline-level-style text:level="2" style:num-format="1" style:num-suffix="." style:display-levels="2"/>'
    + '</text:outline-style>'

  test('headings number continuously with joined display levels', async () => {
    const doc = await odtDoc({
      extraCommonStyles: OUTLINE,
      paras: [{ text: 'A', heading: 1 }, { text: 'B', heading: 2 }, { text: 'C', heading: 2 }],
    })
    expect(doc.sections[0].paragraphs.map((p) => p.listMarker)).toEqual(['1.', '1.1.', '1.2.'])
  })
})

describe('odt end-to-end render', () => {
  test('a mixed document paints ink on one page', async () => {
    const { createCanvas } = await import('canvas')
    const picture = createCanvas(24, 16)
    const pctx = picture.getContext('2d')!
    pctx.fillStyle = '#ff0000'
    pctx.fillRect(0, 0, 24, 16)
    const png = new Uint8Array(picture.toBuffer('image/png'))
    const { getPaintables } = await import('../src/render/paint')
    const { renderPaintables } = await import('../src/test/pixel-diff')
    const doc = await loadOfficeFile(await buildOdt({
      pictures: { 'Pictures/img1.png': png },
      paras: [
        { text: 'Rendered title', heading: 1 },
        { runs: ['body text with ', { text: 'emphasis', bold: true }] },
        { text: 'caption', frames: [{ picture: 'Pictures/img1.png', widthCm: 2, heightCm: 1 }] },
      ],
      tables: [{ colWidths: ['4cm', '4cm'], rows: [{ cells: [{ text: 'cell one' }, { text: 'cell two' }] }] }],
    }))
    const paintables = await getPaintables(doc as never)
    expect(paintables.length).toBeGreaterThan(0)
    const bitmaps = await renderPaintables(paintables as never)
    let ink = 0
    let pixels = 0
    for (const bm of bitmaps) {
      for (let i = 0; i < bm.data.length; i += 4) {
        if (bm.data[i] < 235 || bm.data[i + 1] < 235 || bm.data[i + 2] < 235) ink++
      }
      pixels += bm.width * bm.height
    }
    expect(ink / pixels).toBeGreaterThan(0.001)
  })
})

describe('odt master persistence', () => {
  test('a custom first master survives ordinary paragraphs', async () => {
    const doc = await odtDoc({
      pageLayouts: { L1: { widthCm: 15, heightCm: 20, marginCm: 1 } },
      masters: { M1: { layout: 'L1', header: [{ text: 'HD' }] } },
      paras: [{ text: 'first', masterPage: 'M1' }, { text: 'second' }, { text: 'third' }],
    })
    // one section, not three: undefined masters inherit the active one
    expect(doc.sections).toHaveLength(1)
    const [section] = doc.sections
    expect(section.pageSize.widthTwips).toBeCloseTo(8504, 0)
    expect(section.header?.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['HD'])
    expect(section.paragraphs.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['first', 'second', 'third'])
  })

  test('an explicit later master still seals a new section', async () => {
    const doc = await odtDoc({
      pageLayouts: { L1: { widthCm: 15, heightCm: 20 } },
      masters: { M1: { layout: 'L1' } },
      paras: [{ text: 'first' }, { text: 'second', masterPage: 'M1' }],
    })
    expect(doc.sections).toHaveLength(2)
    expect(doc.sections[0].pageSize.widthTwips).toBeCloseTo(11906, 0)
    expect(doc.sections[1].pageSize.widthTwips).toBeCloseTo(8504, 0)
  })
})

describe('odt covered cells', () => {
  test('a horizontally covered cell is consumed by its span', async () => {
    const doc = await odtDoc({
      tables: [{
        colWidths: ['2cm', '2cm', '2cm'],
        rows: [{ cells: [{ text: 'merged', gridSpan: 2 }, { covered: true }, { text: 'last' }] }],
      }],
    })
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw new Error('expected a table')
    // 3 grid columns, 2 model cells — the covered placeholder adds no column
    expect(block.table.gridColsTwips).toHaveLength(3)
    expect(block.table.rows[0].cells.map((c) => c.gridSpan)).toEqual([2, 1])
  })

  test('merged spans land on the grid when laid out', async () => {
    const measureFixed: MeasureFn = (text, style) => text.length * style.fontSizePt * 0.6 * (96 / 72)
    const doc = await odtDoc({
      tables: [{
        colWidths: ['2cm', '2cm', '2cm'],
        rows: [{ cells: [{ text: 'merged', gridSpan: 2 }, { covered: true }, { text: 'last' }] }],
      }],
    })
    const pages = layoutDocx(doc, measureFixed)
    expect(pages.length).toBeGreaterThan(0)
    const table = pages[0].tables[0]
    const colPx = 2 * 1440 / 2.54 / 15 // 2cm column at 96dpi
    expect(table.rows[0].cells).toHaveLength(2)
    expect(table.rows[0].cells[0].widthPx).toBeCloseTo(2 * colPx, 0)
    expect(table.rows[0].cells[1].xPx).toBeCloseTo(2 * colPx, 0)
  })

  test('combined row+column spans keep vertical leftovers', async () => {
    const doc = await odtDoc({
      extraBlocks: [
        '<table:table>'
        + '<table:table-column table:number-columns-repeated="2"/>'
        + '<table:table-row><table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="2"><text:p>block</text:p></table:table-cell><table:covered-table-cell/></table:table-row>'
        + '<table:table-row><table:covered-table-cell/><table:covered-table-cell/></table:table-row>'
        + '</table:table>',
      ],
    })
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw new Error('expected a table')
    expect(block.table.rows[0].cells.map((c) => [c.gridSpan, c.vMerge])).toEqual([[2, 'restart']])
    expect(block.table.rows[1].cells.map((c) => [c.gridSpan, c.vMerge])).toEqual([[1, 'continue'], [1, 'continue']])
  })
})

describe('odt repetition', () => {
  test('content rows repeat faithfully', async () => {
    const doc = await odtDoc({
      tables: [{ colWidths: ['3cm'], rows: [{ cells: [{ text: 'again' }], repeat: 2 }] }],
    })
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw new Error('expected a table')
    expect(block.table.rows).toHaveLength(2)
    expect(block.table.rows.map((r) => r.cells[0].paragraphs[0].runs[0].text)).toEqual(['again', 'again'])
  })

  test('content cells repeat faithfully', async () => {
    const doc = await odtDoc({
      tables: [{ colWidths: ['2cm', '2cm', '2cm'], rows: [{ cells: [{ text: 'x', repeat: 3 }] }] }],
    })
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw new Error('expected a table')
    const cells = block.table.rows[0].cells
    expect(cells).toHaveLength(3)
    expect(cells.map((c) => c.paragraphs[0].runs[0].text)).toEqual(['x', 'x', 'x'])
  })

  test('absurd repetition counts are capped', async () => {
    const doc = await odtDoc({
      tables: [{ colWidths: ['2cm'], rows: [{ cells: [{ text: 'x' }], repeat: 5000 }] }],
    })
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw new Error('expected a table')
    expect(block.table.rows.length).toBeLessThanOrEqual(1024)
    expect(block.table.rows.length).toBeGreaterThan(1)
  })
})

describe('odt fonts', () => {
  test('font-name resolves through font-face-decls on paras and spans', async () => {
    const doc = await odtDoc({
      paras: [{ style: 'P', runs: [{ text: 'hello', spanStyle: 'T' }] }],
      extraAutoStyles:
        '<style:style style:name="P" style:family="paragraph"><style:text-properties style:font-name="Arial"/></style:style>'
        + '<style:style style:name="T" style:family="text"><style:text-properties style:font-name="Courier New"/></style:style>',
    })
    const run = (doc.sections[0].blocks[0] as { kind: 'p'; paragraph: { runs: Array<{ fontFamily?: string }> } }).paragraph.runs[0]
    expect(run.fontFamily).toBe('Courier New')
  })

  test('fo:font-family works and the resolved font reaches layout', async () => {
    const measureFixed: MeasureFn = (text, style) => text.length * style.fontSizePt * 0.6 * (96 / 72)
    const doc = await odtDoc({
      paras: [{ style: 'P', runs: [{ text: 'sized' }] }],
      extraAutoStyles:
        '<style:style style:name="P" style:family="paragraph"><style:text-properties fo:font-family="Georgia" fo:font-size="14pt"/></style:style>',
    })
    const para = (doc.sections[0].blocks[0] as { kind: 'p'; paragraph: { runs: Array<{ fontFamily?: string; fontSizePt?: number }> } }).paragraph
    expect(para.runs[0].fontFamily).toBe('Georgia')
    expect(para.runs[0].fontSizePt).toBe(14)
    const pages = layoutDocx(doc, measureFixed)
    expect(pages[0].lines[0].segs[0].style.fontFamily).toBe('Georgia')
  })
})

describe('odt nested lists', () => {
  test('a nested list preserves the parent counter', async () => {
    const doc = await odtDoc({
      extraAutoStyles: odfDecimalListStyle('DL'),
      lists: [{
        styleName: 'DL',
        items: [{ text: 'one' }, { text: 'two', nested: { styleName: 'DL', items: [{ text: 'sub' }] } }, { text: 'three' }],
      }],
    })
    expect(doc.sections[0].paragraphs.map((p) => p.listMarker)).toEqual(['1.', '2.', 'a)', '3.'])
  })

  test('a nested list without a style inherits the surrounding one', async () => {
    const doc = await odtDoc({
      extraAutoStyles: odfDecimalListStyle('DL'),
      lists: [{ styleName: 'DL', items: [{ text: 'one', nested: { items: [{ text: 'sub' }] } }] }],
    })
    expect(doc.sections[0].paragraphs.map((p) => p.listMarker)).toEqual(['1.', 'a)'])
  })
})

describe('odt uppercase numbering', () => {
  const STYLES =
    '<text:list-style style:name="U"><text:list-level-style-number text:level="1" style:num-format="A" style:num-suffix="."/></text:list-style>'
    + '<text:list-style style:name="R"><text:list-level-style-number text:level="1" style:num-format="I" style:num-suffix="."/></text:list-style>'
  test('uppercase alpha and Roman markers render', async () => {
    const doc = await odtDoc({
      extraCommonStyles: STYLES,
      lists: [
        { styleName: 'U', items: [{ text: 'a' }, { text: 'b' }] },
        { styleName: 'R', items: [{ text: 'i' }, { text: 'ii' }] },
      ],
    })
    expect(doc.sections[0].paragraphs.map((p) => p.listMarker)).toEqual(['A.', 'B.', 'I.', 'II.'])
  })
})

describe('odt nested display-level joins', () => {
  const JOINED =
    '<text:list-style style:name="DJ">'
    + '<text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."><style:list-level-properties fo:margin-left="1.2cm" fo:text-indent="-0.6cm"/></text:list-level-style-number>'
    + '<text:list-level-style-number text:level="2" style:num-format="1" style:num-suffix="." style:display-levels="2"><style:list-level-properties fo:margin-left="1.8cm" fo:text-indent="-0.6cm"/></text:list-level-style-number>'
    + '</text:list-style>'

  test('a style-less nested sequence joins ancestor counters', async () => {
    const doc = await odtDoc({
      extraCommonStyles: JOINED,
      lists: [{ styleName: 'DJ', items: [{ text: 'one' }, { text: 'two', nested: { items: [{ text: 'sub' }] } }] }],
    })
    expect(doc.sections[0].paragraphs.map((p) => p.listMarker)).toEqual(['1.', '2.', '2.1.'])
  })

  test('an explicitly same-styled nested sequence joins too', async () => {
    const doc = await odtDoc({
      extraCommonStyles: JOINED,
      lists: [{ styleName: 'DJ', items: [{ text: 'one' }, { text: 'two', nested: { styleName: 'DJ', items: [{ text: 'sub' }] } }] }],
    })
    expect(doc.sections[0].paragraphs.map((p) => p.listMarker)).toEqual(['1.', '2.', '2.1.'])
  })
})

describe('odt repeated covered cells', () => {
  test('repeated row-span leftovers stay vertical without shifting the row', async () => {
    const doc = await odtDoc({
      tables: [{
        colWidths: ['2cm', '2cm', '2cm'],
        rows: [
          { cells: [{ text: 'merged', rowSpan: 2, repeat: 2 }, { text: 'last' }] },
          { cells: [{ covered: true, repeat: 2 }, { text: 'last' }] },
        ],
      }],
    })
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw new Error('expected a table')
    expect(block.table.rows[0].cells.map((c) => [c.gridSpan, c.vMerge])).toEqual([[1, 'restart'], [1, 'restart'], [1, undefined]])
    expect(block.table.rows[1].cells.map((c) => [c.gridSpan, c.vMerge])).toEqual([[1, 'continue'], [1, 'continue'], [1, undefined]])
    expect(block.table.rows[1].cells.map((c) => c.paragraphs[0].runs.map((r) => r.text).join(''))).toEqual(['', '', 'last'])
  })

  test('a gridSpan=3 consumes both repeated placeholders', async () => {
    const doc = await odtDoc({
      tables: [{
        colWidths: ['2cm', '2cm', '2cm', '2cm'],
        rows: [{ cells: [{ text: 'wide', gridSpan: 3 }, { covered: true, repeat: 2 }, { text: 'last' }] }],
      }],
    })
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw new Error('expected a table')
    expect(block.table.rows[0].cells.map((c) => c.gridSpan)).toEqual([3, 1])
  })

  test('combined horizontal and vertical spans keep both axes', async () => {
    const doc = await odtDoc({
      tables: [{
        colWidths: ['2cm', '2cm', '2cm'],
        rows: [
          { cells: [{ text: 'block', gridSpan: 2, rowSpan: 2 }, { covered: true }, { text: 'side' }] },
          { cells: [{ covered: true, repeat: 2 }, { text: 'below' }] },
        ],
      }],
    })
    const block = doc.sections[0].blocks[0]
    if (block.kind !== 'table') throw new Error('expected a table')
    expect(block.table.rows[0].cells.map((c) => [c.gridSpan, c.vMerge])).toEqual([[2, 'restart'], [1, undefined]])
    expect(block.table.rows[1].cells.map((c) => [c.gridSpan, c.vMerge])).toEqual([[1, 'continue'], [1, 'continue'], [1, undefined]])
  })
})
