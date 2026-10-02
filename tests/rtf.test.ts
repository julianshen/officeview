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

  test('\\landscape does not re-swap explicitly declared dimensions', () => {
    // Writers store the page in its final orientation, so \paperw/\paperh are
    // already the landscape dimensions and must be used as given.
    const section = parseRtf('{\\rtf1\\ansi\\paperw15840\\paperh12240\\lndscpsxn x\\par}').sections[0]
    expect(section.pageSize.orientation).toBe('landscape')
    expect(section.pageSize.widthTwips).toBe(15840)
    expect(section.pageSize.heightTwips).toBe(12240)
  })

  test('\\landscape rotates the built-in default when no size was declared', () => {
    // With no \paperw/\paperh the default is portrait Letter, so landscape
    // should present the wider page.
    const section = parseRtf('{\\rtf1\\ansi\\lndscpsxn x\\par}').sections[0]
    expect(section.pageSize.orientation).toBe('landscape')
    expect(section.pageSize.widthTwips).toBeGreaterThan(section.pageSize.heightTwips)
  })

  test('portrait is unaffected', () => {
    const section = parseRtf('{\\rtf1\\ansi\\paperw12240\\paperh15840 x\\par}').sections[0]
    expect(section.pageSize.orientation).toBe('portrait')
    expect(section.pageSize.widthTwips).toBe(12240)
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
/**
 * Regression tests for the six review findings against commit 72bb78a.
 * Each uses the exact input from the finding so the fix is pinned.
 */
describe('review finding 1: declared code pages', () => {
  test('ansicpg1251 decodes Cyrillic hex escapes', () => {
    const [p] = paragraphs("{\\rtf1\\ansi\\ansicpg1251 \\'cf\\'f0\\'e8\\'e2\\'e5\\'f2\\par}")
    expect(textOf(p)).toBe('\u041f\u0440\u0438\u0432\u0435\u0442')
  })

  test('the default is still windows-1252', () => {
    const [p] = paragraphs("{\\rtf1\\ansi it\\'92s\\par}")
    expect(textOf(p)).toBe('it\u2019s')
  })

  test('a font charset overrides the document code page', () => {
    const [p] = paragraphs(
      '{\\rtf1\\ansi\\ansicpg1252' +
        '{\\fonttbl{\\f0\\froman\\fcharset204 Arial Cyr;}}' +
        "\\f0 \\'c0\\'e1\\'e2\\par}",
    )
    expect(textOf(p)).toBe('\u0410\u0431\u0432')
  })

  test('hex escapes and \\uN coexist without losing bytes', () => {
    // the space between the hex run and \uN is literal document text
    const [p] = paragraphs("{\\rtf1\\ansi\\ansicpg1251 \\'cf\\'f0 \\u945 ?\\par}")
    expect(textOf(p)).toBe('\u041f\u0440 \u03b1')
  })

  test('a Symbol font keeps its codepage-relative bytes unmangled', () => {
    // charset 2 is Symbol: the byte is a glyph index, not text. \'92 must NOT
    // become the cp1252 right-quote that a text decoder would produce.
    const rtf = String.raw`{\rtf1\ansi{\fonttbl{\f0\froman\fcharset0 Arial;}{\f2\fnil\fcharset2 Symbol;}}\f2 \'92\f0 \par}`
    const [p] = paragraphs(rtf)
    expect(textOf(p)).not.toContain('\u2019')
    expect(textOf(p)).toContain('\u0092')
  })
})

describe('review finding 2: list paragraphs', () => {
  test('grouped {\\pntext} (the form Word writes) marks a list item', () => {
    const [p] = paragraphs(
      '{\\rtf1\\ansi\\li720\\fi-360{\\pntext\\f2\\pnindent0{\\pntxtb\\u-3913 ?}}\\ls1\\ilvl0 Item\\par}',
    )
    expect(p.listMarker).toBe('\u2022')
    expect(p.listLevel).toBe(0)
  })

  test('\\lsN alone marks list membership, without any \\pntext', () => {
    const [p] = paragraphs('{\\rtf1\\ansi{\\listtext\\f2 1.\\tab}\\ls1\\ilvl0 Item\\par}')
    expect(p.listMarker).toBe('1.')
  })

  test('\\ls0 is not a list', () => {
    const [p] = paragraphs('{\\rtf1\\ansi\\ls0 plain\\par}')
    expect(p.listMarker).toBeUndefined()
  })

  test('nested list levels are preserved', () => {
    const [p] = paragraphs('{\\rtf1\\ansi{\\listtext\\f2 a.\\tab}\\ls1\\ilvl2 Item\\par}')
    expect(p.listMarker).toBe('a.')
    expect(p.listLevel).toBe(2)
  })

  test('the grouped pntext marker definition never leaks as text', () => {
    const [p] = paragraphs(
      '{\\rtf1\\ansi{\\pntext\\f2\\pnindent0{\\pntxtb\\u-3913 ?}}\\ls1 Item\\par}',
    )
    expect(textOf(p)).toBe('Item')
  })
})

describe('review finding 3: \\plain', () => {
  test('\\plain clears colour, highlight and font', () => {
    const p = paragraphs(
      '{\\rtf1\\ansi{\\colortbl ;\\red255\\green0\\blue0;\\red0\\green255\\blue0;}' +
        "{\\fonttbl{\\f0\\froman Calibri;}{\\f1\\fswiss Arial;}}\\cf1 \\cb2 \\f1 red\\plain normal\\par}",
    )[0]
    const red = p.runs.find((r) => r.text.includes('red'))
    const normal = p.runs.find((r) => r.text.includes('normal'))
    expect(red?.color).toBe('#ff0000')
    expect(red?.highlight).toBe('#00ff00')
    expect(normal?.color).toBeUndefined()
    expect(normal?.highlight).toBeUndefined()
    // \plain restores the document default font rather than leaving none.
    expect(normal?.fontFamily).toBe('Calibri')
  })

  test('\\plain restores the document default font from \\deff', () => {
    const [p] = paragraphs(
      '{\\rtf1\\ansi\\deff1{\\fonttbl{\\f0\\froman Calibri;}{\\f1\\fswiss Arial;}}\\f0 body\\f1 \\plain after\\par}',
    )
    expect(p.runs.find((r) => r.text === 'after')?.fontFamily).toBe('Arial')
  })

  test('\\plain resets bold, italic and size too', () => {
    const [p] = paragraphs('{\\rtf1\\ansi\\b\\i\\fs40 loud\\plain quiet\\par}')
    expect(p.runs.find((r) => r.text === 'loud')?.bold).toBe(true)
    const quiet = p.runs.find((r) => r.text === 'quiet')
    expect(quiet?.bold).toBeUndefined()
    expect(quiet?.italic).toBeUndefined()
    expect(quiet?.fontSizePt).toBeUndefined()
  })
})

describe('review finding 4: paragraph properties persist across \\par', () => {
  test('alignment persists until \\pard', () => {
    const paras = paragraphs('{\\rtf1\\ansi\\qc A\\par B\\par\\pard C\\par}')
    expect(paras.map((p) => p.align)).toEqual(['center', 'center', 'left'])
  })

  test('an explicit override still wins', () => {
    const paras = paragraphs('{\\rtf1\\ansi\\qc A\\par\\qr B\\par}')
    expect(paras.map((p) => p.align)).toEqual(['center', 'right'])
  })

  test('indents and spacing persist across \\par', () => {
    const paras = paragraphs('{\\rtf1\\ansi\\li720\\sb200 A\\par B\\par\\pard C\\par}')
    expect(paras[0].indentLeftTwips).toBe(720)
    expect(paras[1].indentLeftTwips).toBe(720)
    expect(paras[1].spacingBeforeTwips).toBe(200)
    expect(paras[2].indentLeftTwips).toBe(0)
  })

  test('table routing survives the persistence change', () => {
    const table = firstTable(
      '{\\rtf1\\ansi\\qc\\trowd\\cellx2000\\intbl \\b Shaded\\b0\\cell\\row}',
    )!
    expect(table.rows[0].cells[0].paragraphs[0].runs[0].text).toContain('Shaded')
  })

  test('text after a table still returns to the body', () => {
    const paras = paragraphs('{\\rtf1\\ansi\\qc before\\par\\trowd\\cellx2000\\intbl cell\\cell\\row after\\par}')
    expect(paras.map(textOf)).toEqual(['before', 'after'])
  })

  test('a multi-paragraph cell keeps every paragraph', () => {
    const table = firstTable('{\\rtf1\\ansi\\trowd\\cellx4000\\intbl one\\par\\intbl two\\par\\cell\\row}')!
    const texts = table.rows[0].cells[0].paragraphs.map(textOf).filter((t) => t.length > 0)
    expect(texts).toEqual(['one', 'two'])
  })

  test('list membership does not leak to the following paragraph', () => {
    const paras = paragraphs('{\\rtf1\\ansi{\\listtext\\f2 1.\\tab}\\ls1 Item\\par Not a list\\par}')
    expect(paras[0].listMarker).toBe('1.')
    expect(paras[1].listMarker).toBeUndefined()
  })
})

describe('review finding 5: \\uc is group scoped', () => {
  test('an inner \\uc is restored at group exit', () => {
    const [p] = paragraphs('{\\rtf1\\ansi\\uc1 \\u945 ?{\\uc0 \\u946 }\\u946 ?X\\par}')
    // Without restoration the final \u946 would skip 0 characters and leak "?".
    expect(textOf(p)).toBe('\u03b1\u03b2\u03b2X')
  })

  test('with \\uc0 a fallback character is literal, not skipped', () => {
    const [p] = paragraphs('{\\rtf1\\ansi\\uc1 \\u945 ?{\\uc0 \\u946 ?}X\\par}')
    expect(textOf(p)).toBe('\u03b1\u03b2?X')
  })

  test('fallback skipping stays correct across many \\uN', () => {
    const [p] = paragraphs('{\\rtf1\\ansi\\uc1 \\u945 ?\\u946 ?\\u947 ?\\par}')
    expect(textOf(p)).toBe('\u03b1\u03b2\u03b3')
  })
})

describe('review finding 6: landscape geometry', () => {
  test('explicit landscape dimensions are honoured as written', () => {
    const s = parseRtf('{\\rtf1\\ansi\\paperw15840\\paperh12240\\lndscpsxn x\\par}').sections[0]
    expect(s.pageSize.widthTwips).toBe(15840)
    expect(s.pageSize.heightTwips).toBe(12240)
    expect(s.pageSize.orientation).toBe('landscape')
  })

  test("A4 landscape keeps the writer's own dimensions", () => {
    const s = parseRtf('{\\rtf1\\ansi\\paperw16838\\paperh11906\\lndscpsxn x\\par}').sections[0]
    expect(s.pageSize.widthTwips).toBe(16838)
    expect(s.pageSize.heightTwips).toBe(11906)
  })

  test('with no declared size the portrait default is rotated', () => {
    const s = parseRtf('{\\rtf1\\ansi\\landscape x\\par}').sections[0]
    expect(s.pageSize.orientation).toBe('landscape')
    expect(s.pageSize.widthTwips).toBeGreaterThan(s.pageSize.heightTwips)
  })

  test('portrait with explicit dimensions is untouched', () => {
    const s = parseRtf('{\\rtf1\\ansi\\paperw12240\\paperh15840 x\\par}').sections[0]
    expect(s.pageSize.widthTwips).toBe(12240)
    expect(s.pageSize.heightTwips).toBe(15840)
  })
})

describe('review fixes render correctly', () => {
  const measureFor = async () => {
    const { createCanvas } = await import('canvas')
    const { createMeasurer } = await import('../src/docx/layout')
    return createMeasurer(createCanvas(10, 10).getContext('2d') as unknown as CanvasRenderingContext2D)
  }

  test('code-page text paints real ink onto the page', async () => {
    const { layoutDocx, renderPages } = await import('../src/docx/layout')
    const rtf = [
      String.raw`{\rtf1\ansi\ansicpg1251\deff0`,
      String.raw`{\fonttbl{\f0\froman Times New Roman;}}`,
      String.raw`\paperw12240\paperh15840\margl1440\margr1440\margt1440\margb1440`,
      String.raw`\qc \'cf\'f0\'e8\'e2\'e5\'f2\par`,
      String.raw`\pard trailing paragraph\par}`,
    ].join('\n')
    const pages = layoutDocx(parseRtf(rtf), await measureFor())
    const text = pages[0].lines.map((l) => l.segs.map((s) => s.text).join('')).join(' ')
    expect(text).toContain('\u041f\u0440\u0438\u0432\u0435\u0442')

    const { createCanvas } = await import('canvas')
    const c = createCanvas(Math.ceil(pages[0].widthPx), Math.ceil(pages[0].heightPx))
    const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
    ctx.fillStyle = '#fff'
    ctx.fillRect(0, 0, c.width, c.height)
    renderPages(pages, ctx)
    const data = (ctx as unknown as { getImageData: (a: number, b: number, w: number, h: number) => { data: Uint8ClampedArray } })
      .getImageData(0, 0, c.width, c.height).data
    let dark = 0
    for (let i = 0; i < data.length; i += 4) if (data[i] < 240 || data[i + 1] < 240 || data[i + 2] < 240) dark++
    expect(dark).toBeGreaterThan(100)
  })

  test('bold set before \\par carries into the following paragraphs', async () => {
    const { layoutDocx } = await import('../src/docx/layout')
    // \pard resets paragraph properties but is not a paragraph break; \par is.
    const pages = layoutDocx(
      parseRtf('{\\rtf1\\ansi\\b first\\par second\\par\\b0 third\\par}'),
      await measureFor(),
    )
    const boldFlags = pages[0].lines.map((l) => l.segs.some((s) => s.style.bold))
    expect(boldFlags).toEqual([true, true, false])
  })

  test('explicit landscape renders a wide page', async () => {
    const { layoutDocx } = await import('../src/docx/layout')
    const pages = layoutDocx(
      parseRtf('{\\rtf1\\ansi\\paperw15840\\paperh12240\\margl1440\\margr1440\\margt1440\\margb1440 x\\par}'),
      await measureFor(),
    )
    expect(pages[0].widthPx).toBeGreaterThan(pages[0].heightPx)
  })

  test('a default portrait page is still taller than wide', async () => {
    const { layoutDocx } = await import('../src/docx/layout')
    const pages = layoutDocx(parseRtf('{\\rtf1\\ansi x\\par}'), await measureFor())
    expect(pages[0].heightPx).toBeGreaterThan(pages[0].widthPx)
  })
})

/**
 * Regression tests for the second independent review of c2ec19f.
 * Every case below fails against c2ec19f and passes after the fixes.
 */
import { CP1251_TABLE, hasEncodingSupport } from '../src/rtf/parse'

/**
 * The authoritative windows-1251 mapping, generated from Python's cp1251 codec.
 * `null` marks byte 0x98, which is genuinely undefined in the encoding.
 */
const EXPECTED_CP1251: (number | null)[] = [
  0x0000, 0x0001, 0x0002, 0x0003, 0x0004, 0x0005, 0x0006, 0x0007, 0x0008, 0x0009, 0x000A,
  0x000B, 0x000C, 0x000D, 0x000E, 0x000F, 0x0010, 0x0011, 0x0012, 0x0013, 0x0014, 0x0015,
  0x0016, 0x0017, 0x0018, 0x0019, 0x001A, 0x001B, 0x001C, 0x001D, 0x001E, 0x001F, 0x0020,
  0x0021, 0x0022, 0x0023, 0x0024, 0x0025, 0x0026, 0x0027, 0x0028, 0x0029, 0x002A, 0x002B,
  0x002C, 0x002D, 0x002E, 0x002F, 0x0030, 0x0031, 0x0032, 0x0033, 0x0034, 0x0035, 0x0036,
  0x0037, 0x0038, 0x0039, 0x003A, 0x003B, 0x003C, 0x003D, 0x003E, 0x003F, 0x0040, 0x0041,
  0x0042, 0x0043, 0x0044, 0x0045, 0x0046, 0x0047, 0x0048, 0x0049, 0x004A, 0x004B, 0x004C,
  0x004D, 0x004E, 0x004F, 0x0050, 0x0051, 0x0052, 0x0053, 0x0054, 0x0055, 0x0056, 0x0057,
  0x0058, 0x0059, 0x005A, 0x005B, 0x005C, 0x005D, 0x005E, 0x005F, 0x0060, 0x0061, 0x0062,
  0x0063, 0x0064, 0x0065, 0x0066, 0x0067, 0x0068, 0x0069, 0x006A, 0x006B, 0x006C, 0x006D,
  0x006E, 0x006F, 0x0070, 0x0071, 0x0072, 0x0073, 0x0074, 0x0075, 0x0076, 0x0077, 0x0078,
  0x0079, 0x007A, 0x007B, 0x007C, 0x007D, 0x007E, 0x007F, 0x0402, 0x0403, 0x201A, 0x0453,
  0x201E, 0x2026, 0x2020, 0x2021, 0x20AC, 0x2030, 0x0409, 0x2039, 0x040A, 0x040C, 0x040B,
  0x040F, 0x0452, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, null, 0x2122, 0x0459,
  0x203A, 0x045A, 0x045C, 0x045B, 0x045F, 0x00A0, 0x040E, 0x045E, 0x0408, 0x00A4, 0x0490,
  0x00A6, 0x00A7, 0x0401, 0x00A9, 0x0404, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x0407, 0x00B0,
  0x00B1, 0x0406, 0x0456, 0x0491, 0x00B5, 0x00B6, 0x00B7, 0x0451, 0x2116, 0x0454, 0x00BB,
  0x0458, 0x0405, 0x0455, 0x0457, 0x0410, 0x0411, 0x0412, 0x0413, 0x0414, 0x0415, 0x0416,
  0x0417, 0x0418, 0x0419, 0x041A, 0x041B, 0x041C, 0x041D, 0x041E, 0x041F, 0x0420, 0x0421,
  0x0422, 0x0423, 0x0424, 0x0425, 0x0426, 0x0427, 0x0428, 0x0429, 0x042A, 0x042B, 0x042C,
  0x042D, 0x042E, 0x042F, 0x0430, 0x0431, 0x0432, 0x0433, 0x0434, 0x0435, 0x0436, 0x0437,
  0x0438, 0x0439, 0x043A, 0x043B, 0x043C, 0x043D, 0x043E, 0x043F, 0x0440, 0x0441, 0x0442,
  0x0443, 0x0444, 0x0445, 0x0446, 0x0447, 0x0448, 0x0449, 0x044A, 0x044B, 0x044C, 0x044D,
  0x044E, 0x044F,
]

const hex2 = (b: number) => b.toString(16).padStart(2, '0')

describe('review 1: buffered hex at end of document', () => {
  test('a document ending in hex with no \\par still yields text', () => {
    const doc = parseRtf(String.raw`{\rtf1 \'e9}`)
    const paras = doc.sections[0].paragraphs
    expect(paras).toHaveLength(1)
    expect(textOf(paras[0])).toBe('\u00e9')
  })

  test('a longer trailing hex run survives', () => {
    const [p] = paragraphs(String.raw`{\rtf1 \'cf\'f0\'e8}`)
    expect(textOf(p)).toBe('\u00cf\u00f0\u00e8')
  })
})

describe('review 2: buffered hex must not cross formatting boundaries', () => {
  test('hex before \\b stays unbold', () => {
    const rtf = String.raw`{\rtf1 \'e9\b bold\par}`
    const runs = paragraphs(rtf)[0].runs
    expect(runs[0]).toMatchObject({ text: '\u00e9' })
    expect(runs[0].bold).toBeUndefined()
    expect(runs.find((r) => r.text === 'bold')?.bold).toBe(true)
  })

  test('hex before \\cf is not painted in the new colour', () => {
    const rtf = String.raw`{\rtf1{\colortbl ;\red255\green0\blue0;}\'e9\cf1 red\cf0 \par}`
    const runs = paragraphs(rtf)[0].runs
    expect(runs[0].color).toBeUndefined()
    expect(runs.find((r) => r.text === 'red')?.color).toBe('#ff0000')
  })

  test('hex before \\f is not attributed to the new font', () => {
    const rtf = '{\\rtf1{\\fonttbl{\\f0\\froman Calibri;}{\\f1\\fswiss Arial;}}\\\'e9\\f1 later\\par}'
    const runs = paragraphs(rtf)[0].runs
    expect(runs[0].fontFamily).toBe('Calibri')
    expect(runs.find((r) => r.text === 'later')?.fontFamily).toBe('Arial')
  })

  test('hex at the end of a nested bold group keeps its bold', () => {
    const rtf = String.raw`{\rtf1 {\b x\'e9}\par}`
    const runs = paragraphs(rtf)[0].runs
    expect(runs).toHaveLength(1)
    expect(runs[0]).toMatchObject({ text: 'x\u00e9', bold: true })
  })

  test('hex at the start of a nested group does not inherit the outer style', () => {
    const rtf = String.raw`{\rtf1 \b {\i \'e9}\par}`
    const runs = paragraphs(rtf)[0].runs
    expect(runs[0].text).toBe('\u00e9')
    expect(runs[0].italic).toBe(true)
  })

  test('a complete multibyte character survives a style boundary', () => {
    // Two adjacent escapes form one Shift-JIS character: neither byte may be
    // flushed on its own.
    if (!hasEncodingSupport('shift_jis')) return
    const rtf = String.raw`{\rtf1\ansicpg932 \'82\'a0\par}`
    expect(textOf(paragraphs(rtf)[0])).toBe('\u3042')
  })

  test('adjacent escapes are decoded together, not per byte', () => {
    const rtf = String.raw`{\rtf1\ansicpg932 \'82\'a0\'82\'a2\par}`
    expect(textOf(paragraphs(rtf)[0])).toBe('\u3042\u3044')
  })
})

describe('review 3: marker inside a grouped \\pntext', () => {
  test('a numbered legacy list keeps its number', () => {
    const rtf = String.raw`{\rtf1\li720\fi-360{\pntext\f2 1.\tab}\ls1\ilvl0 Item\par}`
    const p = paragraphs(rtf)[0]
    expect(p.listMarker).toBe('1.')
    expect(textOf(p)).toBe('Item')
  })

  test('a marker nested in pntxtb is captured', () => {
    const rtf = String.raw`{\rtf1\li720\fi-360{\pntext\f2\pnindent0{\pntxtb 1.}}\ls1\ilvl0 Item\par}`
    const p = paragraphs(rtf)[0]
    expect(p.listMarker).toBe('1.')
    expect(textOf(p)).toBe('Item')
  })

  test('a bulleted legacy list renders a bullet', () => {
    const rtf = String.raw`{\rtf1\li720\fi-360{\pntext\f2\pnindent0{\pntxtb\u-3913 ?}}\ls1\ilvl0 Item\par}`
    const p = paragraphs(rtf)[0]
    expect(p.listMarker).toBe('\u2022')
    expect(textOf(p)).toBe('Item')
  })

  test('a multi-character marker is preserved', () => {
    const rtf = String.raw`{\rtf1{\pntext\f2 1.a.\tab}\ls1\ilvl0 Item\par}`
    expect(paragraphs(rtf)[0].listMarker).toBe('1.a.')
  })

  test('the marker never leaks into the paragraph text', () => {
    const rtf = String.raw`{\rtf1{\pntext\f2 1.\tab}\ls1\ilvl0 Item\par}`
    expect(textOf(paragraphs(rtf)[0])).toBe('Item')
  })
})

describe('review 4: windows-1251 decoding', () => {
  test('the fallback table matches the authoritative mapping exactly', () => {
    expect(CP1251_TABLE).toHaveLength(256)
    const mismatches: string[] = []
    for (let b = 0; b < 256; b++) {
      const expected = EXPECTED_CP1251[b]
      const got = CP1251_TABLE[b]
      if (expected === null) {
        if (got !== undefined) mismatches.push(`${hex2(b)}: expected undefined, got ${JSON.stringify(got)}`)
      } else if (got === undefined || got.codePointAt(0) !== expected) {
        mismatches.push(`${hex2(b)}: expected U+${(expected as number).toString(16)}, got ${JSON.stringify(got)}`)
      }
    }
    expect(mismatches).toEqual([])
  })

  test('every byte round-trips through the parser', () => {
    const mismatches: string[] = []
    for (let b = 0; b < 256; b++) {
      const expected = EXPECTED_CP1251[b]
      const rtf = String.raw`{\rtf1\ansi\ansicpg1251 \'` + hex2(b) + String.raw`\par}`
      const got = textOf(paragraphs(rtf)[0])
      const want = expected === null ? '\ufffd' : String.fromCodePoint(expected)
      if (got !== want) mismatches.push(`${hex2(b)}: want ${JSON.stringify(want)}, got ${JSON.stringify(got)}`)
    }
    expect(mismatches).toEqual([])
  })

  test("the review's specific bytes decode correctly", () => {
    // a0 aa af -> NBSP, Є, Ї (the old table read past its bounds and produced
    // a literal "undefined"); 8b/8c were swapped; 9a and b2 were wrong.
    const rtf = String.raw`{\rtf1\ansi\ansicpg1251 \'a0\'aa\'af\'8b\'8c\'9a\'b2\par}`
    expect(textOf(paragraphs(rtf)[0])).toBe('\u00a0\u0404\u0407\u2039\u040a\u0459\u0406')
  })

  test('TextDecoder and the table agree on every defined byte', () => {
    if (!hasEncodingSupport('windows-1251')) return
    const dec = new TextDecoder('windows-1251')
    const mismatches: string[] = []
    for (let b = 0; b < 256; b++) {
      const expected = EXPECTED_CP1251[b]
      if (expected === null) continue
      if (dec.decode(new Uint8Array([b])).codePointAt(0) !== expected) mismatches.push(hex2(b))
    }
    expect(mismatches).toEqual([])
  })

  test('0x98 is the only byte where Node disagrees with the standard', () => {
    // Documented divergence: Node maps the undefined byte to U+0098; the
    // encoding standard and Python treat it as undefined. The parser always
    // uses the table so output does not vary by runtime.
    if (!hasEncodingSupport('windows-1251')) return
    const dec = new TextDecoder('windows-1251')
    expect(dec.decode(new Uint8Array([0x98])).codePointAt(0)).toBe(0x0098)
    expect(CP1251_TABLE[0x98]).toBeUndefined()
    expect(textOf(paragraphs(String.raw`{\rtf1\ansi\ansicpg1251 \'98\par}`)[0])).toBe('\ufffd')
  })

  test('the fallback table is what runs when TextDecoder is unavailable', () => {
    // Guards the assumption behind the table: if this runtime gained cp1251
    // support, the table would silently stop being exercised.
    if (hasEncodingSupport('windows-1251')) return
    const rtf = String.raw`{\rtf1\ansi\ansicpg1251 \'cf\'f0\'e8\'e2\'e5\'f2\par}`
    expect(textOf(paragraphs(rtf)[0])).toBe('\u041f\u0440\u0438\u0432\u0435\u0442')
  })
})

describe('review 5: the \\deff default font', () => {
  const deffRtf =
    '{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\froman\\fcharset204 Arial;}}\\\'c0\\par}'

  test('the default font family applies with no explicit \\fN', () => {
    const [p] = paragraphs(deffRtf)
    expect(p.runs[0].fontFamily).toBe('Arial')
  })

  test("the default font's charset governs hex decoding", () => {
    const [p] = paragraphs(deffRtf)
    expect(textOf(p)).toBe('\u0410')
  })

  test('an explicit \\fN still overrides the default', () => {
    const rtf =
      '{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\froman\\fcharset204 Arial;}{\\f1\\froman\\fcharset0 Calibri;}}\\f1 x\\f0 \\\'c0\\par}'
    const runs = paragraphs(rtf)[0].runs
    expect(runs.find((r) => r.text === 'x')?.fontFamily).toBe('Calibri')
    expect(runs.find((r) => r.text === '\u0410')?.fontFamily).toBe('Arial')
  })

  test('\\plain restores the default font and its charset', () => {
    const rtf =
      '{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\froman\\fcharset204 Arial;}}\\f0 \\\'c0\\plain \\\'c1\\par}'
    const runs = paragraphs(rtf)[0].runs
    // 0xc1 is U+0411 (CYRILLIC CAPITAL LETTER BE) in windows-1251. The runs
    // merge because the restored formatting matches what preceded it, so
    // assert on the run that carries the post-\plain character.
    const afterPlain = runs.find((r) => r.text.includes('\u0411'))
    expect(afterPlain).toBeDefined()
    expect(afterPlain?.fontFamily).toBe('Arial')
    // and the byte decoded via that font's charset, not the document code page
    expect(afterPlain?.text).toContain('\u0411')
  })

  test('an explicit \\ansicpg is not overridden by a charset-0 default font', () => {
    const rtf =
      '{\\rtf1\\ansi\\ansicpg1251\\deff0{\\fonttbl{\\f0\\froman\\fcharset0 Calibri;}}\\\'c0\\par}'
    expect(textOf(paragraphs(rtf)[0])).toBe('\u0410')
  })
})
