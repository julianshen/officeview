/**
 * ODT (OpenDocument text) adapter tests: Writer content parsed into the
 * shared flow model, exercising detection through loadOfficeFile.
 */
import { describe, test, expect } from 'vitest'
import { loadOfficeFile } from '../src/components/OfficeFile'
import { parseOdt } from '../src/odt/parse'
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
