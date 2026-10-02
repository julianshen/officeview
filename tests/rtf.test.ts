import { describe, test, expect } from 'vitest'
import { parseRtf, isRtf } from '../src/rtf/parse'
import { loadOfficeFile } from '../src/components/OfficeFile'
import type { DocxParagraph, DocxTable } from '../src/docx/types'

/**
 * Assertions target the parsed model, never pixels: RTF parsing is font
 * independent, and these stay meaningful across rasterizers.
 */

const textOf = (p: DocxParagraph): string => p.runs.map((r) => r.text).join('')
const paragraphs = (rtf: string): DocxParagraph[] => parseRtf(rtf).sections[0].paragraphs
const firstTable = (rtf: string): DocxTable | undefined => {
  const block = parseRtf(rtf).sections[0].blocks.find((b) => b.kind === 'table')
  return block && block.kind === 'table' ? block.table : undefined
}

/** A Word-shaped preamble so tests exercise real-world structure. */
const preamble =
  '{\\rtf1\\ansi\\ansicpg1252\\deff0\\deflang1033' +
  '{\\fonttbl{\\f0\\froman\\fcharset0 Times New Roman;}{\\f1\\fswiss\\fcharset0 Calibri;}{\\f2\\fnil\\fcharset2 Symbol;}}' +
  '{\\colortbl ;\\red255\\green0\\blue0;\\red0\\green0\\blue255;\\red255\\green255\\blue0;}' +
  '{\\*\\generator Word 16.0;}' +
  '{\\info{\\title Should never render}{\\author Nor this}}'

describe('isRtf', () => {
  test('recognises the RTF magic', () => {
    expect(isRtf(new TextEncoder().encode('{\\rtf1\\ansi hi}'))).toBe(true)
  })

  test('tolerates leading whitespace and a BOM', () => {
    expect(isRtf(new TextEncoder().encode('\n  {\\rtf1 x}'))).toBe(true)
    expect(isRtf(new Uint8Array([0xef, 0xbb, 0xbf, 0x7b, 0x5c, 0x72, 0x74, 0x66, 0x31]))).toBe(true)
  })

  test('rejects a zip container', () => {
    expect(isRtf(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0, 0]))).toBe(false)
    expect(isRtf(new Uint8Array(0))).toBe(false)
  })
})

describe('RTF text and character formatting', () => {
  test('extracts plain text', () => {
    const paras = paragraphs('{\\rtf1\\ansi Hello world\\par}')
    expect(paras).toHaveLength(1)
    expect(textOf(paras[0])).toBe('Hello world')
  })

  test('bold, italic, underline and strike map to runs', () => {
    const [p] = paragraphs('{\\rtf1\\ansi \\b bold\\b0  \\i italic\\i0  \\ul under\\ulnone  \\strike gone\\strike0\\par}')
    const byText = new Map(p.runs.map((r) => [r.text, r]))
    expect(byText.get('bold')?.bold).toBe(true)
    expect(byText.get('italic')?.italic).toBe(true)
    expect(byText.get('under')?.underline).toBe(true)
    expect(byText.get('gone')?.strike).toBe(true)
    expect(byText.get(' ')?.bold).toBeUndefined()
  })

  test('font size comes from \\fs in half-points', () => {
    const [p] = paragraphs('{\\rtf1\\ansi \\fs40 big\\fs24 small\\par}')
    const big = p.runs.find((r) => r.text === 'big')
    expect(big?.fontSizePt).toBe(20)
  })

  test('font and colour tables resolve through \\fN and \\cfN', () => {
    const [p] = paragraphs(preamble + '\\f1 Calibri\\f0  \\cf1 red\\cf0  \\cf2 blue\\cf0 \\par}')
    expect(p.runs.find((r) => r.text === 'Calibri')?.fontFamily).toBe('Calibri')
    expect(p.runs.find((r) => r.text === 'red')?.color).toBe('#ff0000')
    expect(p.runs.find((r) => r.text === 'blue')?.color).toBe('#0000ff')
  })

  test('colour index 0 means automatic, so no colour is emitted', () => {
    const [p] = paragraphs(preamble + '\\cf0 plain\\par}')
    expect(p.runs[0].color).toBeUndefined()
  })

  test('ignorable destinations never leak their text', () => {
    const paras = paragraphs(preamble + 'visible\\par')
    expect(textOf(paras[0])).toBe('visible')
    expect(paras.map(textOf).join('')).not.toContain('Should never render')
    expect(paras.map(textOf).join('')).not.toContain('Nor this')
  })

  test('escaped braces and backslashes render literally', () => {
    const [p] = paragraphs('{\\rtf1\\ansi a\\{b\\}c\\\\d\\par}')
    expect(textOf(p)).toBe('a{b}c\\d')
  })

  test("hex escapes decode through the code page, not as mojibake", () => {
    const [p] = paragraphs("{\\rtf1\\ansi it\\'92s\\par}")
    expect(textOf(p)).toBe('it\u2019s')
  })

  test('\\uN produces unicode and drops the fallback character', () => {
    // each \uN is followed by \uc1 fallback characters that must not be rendered
    const [p] = paragraphs('{\\rtf1\\ansi\\uc1 \\u9731 ?snow \\u8217 ?quote\\par}')
    expect(textOf(p)).toBe('\u2603snow \u2019quote')
  })

  test('negative \\uN recovers the BMP code point', () => {
    // Astral characters arrive as a surrogate pair: two \uN sequences.
    const [p] = paragraphs('{\\rtf1\\ansi\\uc1 \\u-10179 ?\\u-8704 ?\\par}')
    expect(textOf(p)).toBe('\uD83D\uDE00')
  })

  test('a tab becomes a tab in the run text', () => {
    const [p] = paragraphs('{\\rtf1\\ansi a\\tab b\\par}')
    expect(textOf(p)).toBe('a\tb')
  })
})

describe('RTF paragraphs and page setup', () => {
  test('each \\par starts a new paragraph', () => {
    const paras = paragraphs('{\\rtf1\\ansi one\\par two\\par three\\par}')
    expect(paras.map(textOf).filter((t) => t.length > 0)).toEqual(['one', 'two', 'three'])
  })

  test('an empty \\par is preserved as an empty paragraph', () => {
    const paras = paragraphs('{\\rtf1\\ansi a\\par\\par b\\par}')
    expect(paras.map(textOf)).toEqual(['a', '', 'b'])
  })

  test('alignment control words map to the four alignments', () => {
    const paras = paragraphs('{\\rtf1\\ansi \\ql l\\par \\qr r\\par \\qc c\\par \\qj j\\par}')
    const nonEmpty = paras.filter((p) => textOf(p).length > 0)
    expect(nonEmpty.map((p) => p.align)).toEqual(['left', 'right', 'center', 'justify'])
  })

  test('indents and spacing map onto the paragraph', () => {
    const [p] = paragraphs('{\\rtf1\\ansi \\li720\\ri360\\fi-360\\sb120\\sa240 text\\par}')
    expect(p.indentLeftTwips).toBe(720)
    expect(p.indentRightTwips).toBe(360)
    expect(p.indentFirstLineTwips).toBe(-360)
    expect(p.spacingBeforeTwips).toBe(120)
    expect(p.spacingAfterTwips).toBe(240)
  })

  test('page size and margins come from \\paperw and \\marg*', () => {
    const section = parseRtf('{\\rtf1\\ansi\\paperw11906\\paperh16838\\margl1701\\margr1701\\margt1134\\margb1134 x\\par}').sections[0]
    expect(section.pageSize.widthTwips).toBe(11906) // A4
    expect(section.pageSize.heightTwips).toBe(16838)
    expect(section.margins.leftTwips).toBe(1701)
    expect(section.margins.bottomTwips).toBe(1134)
  })

  test('\\landscape swaps the page dimensions', () => {
    const section = parseRtf('{\\rtf1\\ansi\\paperw12240\\paperh15840\\lndscpsxn\\sectdscpsxn x\\par}').sections[0]
    expect(section.pageSize.orientation).toBe('landscape')
    expect(section.pageSize.widthTwips).toBe(15840)
  })

  test('\\sect starts a new section', () => {
    const doc = parseRtf('{\\rtf1\\ansi first\\par\\sect\\paperw8000 second\\par}')
    expect(doc.sections.length).toBeGreaterThanOrEqual(2)
    expect(textOf(doc.sections[0].paragraphs[0])).toBe('first')
    expect(textOf(doc.sections[1].paragraphs[0])).toBe('second')
  })

  test('\\line breaks the line without ending the paragraph', () => {
    const [p] = paragraphs('{\\rtf1\\ansi first\\line second\\par}')
    expect(p.runs.some((r) => r.breakBefore)).toBe(true)
    expect(textOf(p)).toBe('firstsecond')
  })
})

describe('RTF tables', () => {
  const tableRtf =
    '{\\rtf1\\ansi' +
    '\\trowd\\trgaph108\\trleft0' +
    '\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrb\\brdrs\\clbrdrr\\brdrs\\cellx2000' +
    '\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrb\\brdrs\\clbrdrr\\brdrs\\cellx4000' +
    '\\intbl A\\cell B\\cell\\row' +
    '\\trowd\\trgaph108\\trleft0\\cellx2000\\cellx4000' +
    '\\intbl C\\cell D\\cell\\row}'

  test('rows and cells are built in document order', () => {
    const table = firstTable(tableRtf)
    expect(table).toBeDefined()
    expect(table!.rows).toHaveLength(2)
    expect(table!.rows[0].cells).toHaveLength(2)
    expect(table!.rows[0].cells.map((c) => textOf(c.paragraphs[0]))).toEqual(['A', 'B'])
    expect(table!.rows[1].cells.map((c) => textOf(c.paragraphs[0]))).toEqual(['C', 'D'])
  })

  test('\\cellx becomes column widths', () => {
    const table = firstTable(tableRtf)!
    expect(table.rows[0].cells[0].widthTwips).toBe(2000)
    expect(table.rows[0].cells[1].widthTwips).toBe(4000)
    // widths are derived from consecutive right edges
    expect(table.gridColsTwips).toEqual([2000, 2000])
  })

  test('\\clbrdr* sets cell borders', () => {
    const cell = firstTable(tableRtf)!.rows[0].cells[0]
    expect(cell.borders?.top?.style).toBe('single')
    expect(cell.borders?.left?.style).toBe('single')
  })

  test('a cell keeps text without a trailing \\par', () => {
    const table = firstTable('{\\rtf1\\ansi\\trowd\\cellx2000\\intbl solo\\cell\\row}')
    expect(textOf(table!.rows[0].cells[0].paragraphs[0])).toBe('solo')
  })

  test('a second row may reuse the previous row definitions', () => {
    const table = firstTable(
      '{\\rtf1\\ansi\\trowd\\cellx2000\\cellx4000\\intbl A\\cell B\\cell\\row \\intbl C\\cell D\\cell\\row}',
    )!
    expect(table.rows).toHaveLength(2)
    expect(table.rows[1].cells.map((c) => textOf(c.paragraphs[0]))).toEqual(['C', 'D'])
  })

  test('text after the table returns to the body', () => {
    const paras = paragraphs('{\\rtf1\\ansi before\\par \\trowd\\cellx2000\\intbl cell\\cell\\row after\\par}')
    expect(paras.map(textOf)).toEqual(['before', 'after'])
  })
})

describe('RTF images', () => {
  // A real 1x1 PNG, hex encoded exactly as RTF stores it.
  const PNG_HEX =
    '89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489' +
    '0000000a49444154789c6360000002000100ffff03000006000557bfabd4000000' +
    '0049454e44ae426082'
  const JPEG_HEX = 'ffd8ffe000104a46494600010100000100010000ffdb004300ffd9'

  test('\\pngblip decodes to PNG bytes', () => {
    const [p] = paragraphs(`{\\rtf1\\ansi {\\pict\\pngblip\\picw1\\pich1 ${PNG_HEX}}\\par}`)
    expect(p.images).toHaveLength(1)
    expect(p.images[0].mime).toBe('image/png')
    expect(Array.from(p.images[0].data.subarray(0, 4))).toEqual([0x89, 0x50, 0x4e, 0x47])
  })

  test('\\picwgoal sizes the image from twips', () => {
    const [p] = paragraphs(`{\\rtf1\\ansi {\\pict\\pngblip\\picw100\\pich50\\picwgoal1440\\pichgoal720 ${PNG_HEX}}\\par}`)
    expect(p.images[0].widthEmu).toBe(1440 * 635)
    expect(p.images[0].heightEmu).toBe(720 * 635)
  })

  test('without \\picwgoal the pixel size is used at 96dpi', () => {
    const [p] = paragraphs(`{\\rtf1\\ansi {\\pict\\pngblip\\picw96\\pich48 ${PNG_HEX}}\\par}`)
    expect(p.images[0].widthEmu).toBe(Math.round(96 * (914400 / 96)))
    expect(p.images[0].heightEmu).toBe(Math.round(48 * (914400 / 96)))
  })

  test('\\jpegblip decodes to JPEG bytes', () => {
    const [p] = paragraphs(`{\\rtf1\\ansi {\\pict\\jpegblip\\picwgoal1440\\pichgoal720 ${JPEG_HEX}}\\par}`)
    expect(p.images[0].mime).toBe('image/jpeg')
    expect(Array.from(p.images[0].data.subarray(0, 2))).toEqual([0xff, 0xd8])
  })

  test('an image is dropped rather than rendered as hex text', () => {
    const [p] = paragraphs(`{\\rtf1\\ansi a{\\pict\\pngblip\\picwgoal1440\\pichgoal720 ${PNG_HEX}}b\\par}`)
    expect(textOf(p)).toBe('ab')
    expect(p.images).toHaveLength(1)
  })

  test('unsupported metafiles are skipped, not half-rendered', () => {
    const [p] = paragraphs('{\\rtf1\\ansi a{\\pict\\emfblip\\picwgoal1440 0102030405}b\\par}')
    expect(p.images).toHaveLength(0)
    expect(textOf(p)).toBe('ab')
  })
})

describe('RTF lists', () => {
  test('\\listtext becomes the resolved marker and \\ilvl the level', () => {
    const [p] = paragraphs(
      '{\\rtf1\\ansi\\li720\\fi-360{\\*\\pn\\pnf2\\pnindent720\\pnstart1\\pndec}' +
        '{\\listtext\\f2 1.\\tab}\\ilvl1\\pntext First item\\par}',
    )
    expect(textOf(p)).toBe('First item')
    expect(p.listMarker).toBe('1.')
    expect(p.listLevel).toBe(1)
    expect(p.indentLeftTwips).toBe(720)
    expect(p.indentFirstLineTwips).toBe(-360)
  })

  test('a bullet marker survives the Symbol-font escape', () => {
    const [p] = paragraphs('{\\rtf1\\ansi{\\listtext\\f2\\u-3913 ?\\tab}\\pntext Item\\par}')
    expect(p.listMarker).toBe('\u2022')
  })

  test('\\pntext without \\listtext still marks the paragraph as a list', () => {
    const [p] = paragraphs('{\\rtf1\\ansi\\pntext dangling\\par}')
    expect(p.listMarker).toBe('\u2022')
  })

  test('a list paragraph does not leak the marker text into the content', () => {
    const [p] = paragraphs('{\\rtf1\\ansi{\\listtext\\f2 1.\\tab}\\pntext Item\\par}')
    expect(textOf(p)).toBe('Item')
  })
})

describe('RTF robustness', () => {
  test('malformed input does not hang or throw', () => {
    expect(() => parseRtf('{\\rtf1\\ansi {unclosed group')).not.toThrow()
    expect(() => parseRtf('{\\rtf1')).not.toThrow()
    expect(() => parseRtf('{\\rtf1\\ansi \\')).not.toThrow()
  })

  test('empty and non-RTF input yield an empty document', () => {
    expect(parseRtf('').sections).toHaveLength(0)
    expect(parseRtf('not rtf at all').sections).toHaveLength(0)
  })

  test('deep nesting is bounded rather than overflowing the stack', () => {
    const deep = '{\\rtf1\\ansi ' + '{'.repeat(5000) + 'x' + '}'.repeat(5000) + '}'
    expect(() => parseRtf(deep)).not.toThrow()
  })

  test('parsing bytes and parsing text agree', () => {
    const rtf = '{\\rtf1\\ansi\\b byte path\\b0\\par}'
    const fromBytes = parseRtf(new TextEncoder().encode(rtf))
    const fromText = parseRtf(rtf)
    expect(fromBytes.sections[0].paragraphs[0].runs).toEqual(fromText.sections[0].paragraphs[0].runs)
  })

  test('grouped formatting does not leak past its closing brace', () => {
    const [p] = paragraphs('{\\rtf1\\ansi {\\b bold}\\b0 plain\\par}')
    expect(p.runs.find((r) => r.text === 'bold')?.bold).toBe(true)
    expect(p.runs.find((r) => r.text === 'plain')?.bold).toBeUndefined()
  })
})

describe('RTF end-to-end rendering', () => {
  // The model tests above are font-independent; this one proves a parsed RTF
  // actually flows through the shared layout/paint pipeline, which is the whole
  // point of emitting DocxDocument rather than building a second renderer.
  test('an RTF document lays out into pages with lines, a table and a list', async () => {
    const { createCanvas } = await import('canvas')
    const { createMeasurer, layoutDocx } = await import('../src/docx/layout')
    const measure = createMeasurer(createCanvas(10, 10).getContext('2d') as unknown as CanvasRenderingContext2D)
    const rtf =
      '{\\rtf1\\ansi\\deff0' +
      '{\\fonttbl{\\f0\\froman Times New Roman;}{\\f1\\fswiss Calibri;}}' +
      '{\\colortbl ;\\red255\\green0\\blue0;}' +
      '\\paperw12240\\paperh15840\\margl1440\\margr1440\\margt1440\\margb1440' +
      '\\qc\\b RTF title\\b0\\par' +
      '\\pard Body text that should wrap onto a second line once it fills the page width.\\par' +
      '\\trowd\\cellx3000\\cellx6000\\intbl Cell A\\cell Cell B\\cell\\row' +
      '\\pard\\li720\\fi-360{\\listtext\\f2 1.\\tab}\\ilvl0\\pntext A list item\\par' +
      '}'
    const pages = layoutDocx(parseRtf(rtf), measure)
    expect(pages.length).toBeGreaterThan(0)
    expect(pages[0].tables).toHaveLength(1)
    expect(pages[0].lines.length).toBeGreaterThan(2)
    const painted = pages[0].lines.map((l) => l.segs.map((s) => s.text).join('')).join(' ')
    expect(painted).toContain('RTF title')
    expect(painted).toContain('Body text')
    expect(painted).toContain('A list item')
  })

  test('page size follows the RTF section', async () => {
    const { createCanvas } = await import('canvas')
    const { createMeasurer, layoutDocx } = await import('../src/docx/layout')
    const measure = createMeasurer(createCanvas(10, 10).getContext('2d') as unknown as CanvasRenderingContext2D)
    const pages = layoutDocx(
      parseRtf('{\\rtf1\\ansi\\paperw11906\\paperh16838\\margl1440\\margr1440\\margt1440\\margb1440 x\\par}'),
      measure,
    )
    // A4 is 210x297mm; at 96dpi that is roughly 794x1123 css px.
    expect(pages[0].widthPx).toBeCloseTo(794, 0)
    expect(pages[0].heightPx).toBeCloseTo(1123, 0)
  })
})

describe('RTF through loadOfficeFile', () => {
  test('rtf bytes are detected and parsed without touching the zip path', async () => {
    const bytes = new TextEncoder().encode('{\\rtf1\\ansi\\b detected\\b0\\par}')
    const doc = await loadOfficeFile(bytes)
    const paras = (doc as { sections: Array<{ paragraphs: DocxParagraph[] }> }).sections[0].paragraphs
    expect(textOf(paras[0])).toBe('detected')
  })

  test('an rtf string still streams through as bytes', async () => {
    const bytes = new TextEncoder().encode('{\\rtf1\\ansi streamed\\par}')
    const stream = new ReadableStream<Uint8Array>({
      start(c) {
        c.enqueue(bytes.subarray(0, 5))
        c.enqueue(bytes.subarray(5))
        c.close()
      },
    })
    const doc = await loadOfficeFile(stream)
    const paras = (doc as { sections: Array<{ paragraphs: DocxParagraph[] }> }).sections[0].paragraphs
    expect(textOf(paras[0])).toBe('streamed')
  })
})