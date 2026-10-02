/**
 * RTF (Rich Text Format) parsing.
 *
 * RTF is not a zip like the other Office formats, so it never goes through
 * OfficePackage. This parser reads the byte stream directly and emits the same
 * intermediate model the DOCX parser produces (DocxDocument), which means
 * pagination, painting, the text index, search and selection all come from the
 * existing pipeline rather than a second renderer.
 *
 * Structure of the format: literal text mixed with control words (`\b`, `\fs24`),
 * control symbols (`\'92`, `\\`) and groups (`{ }`). A group is a scope -- a
 * property set inside it is discarded at the matching `}` -- so the parser
 * recurses with a *copy* of the current state and lets normal control words fall
 * out of scope for free. Destinations marked `{\*\name ...}` are ignorable and
 * skipped, except the two we must read (font and colour tables) and the two we
 * render (`\pict`, `\listtext`).
 *
 * Honest scope: \emfblip / \wmetafile vector metafiles, drawing shapes, fields
 * and embedded objects are skipped rather than mis-rendered. Nested tables
 * (\itap > 1) are flattened into the outer table. Documented in the README.
 */

import type {
  DocxBlock,
  DocxDocument,
  DocxImage,
  DocxPageMargins,
  DocxParagraph,
  DocxSection,
  DocxTable,
  DocxTableBorders,
  DocxTableCell,
  DocxTableRow,
  DocxTextRun,
  ParagraphAlign,
  TableCellBorder,
} from '../docx/types'

/** RTF stores font size in half-points; 24 = 12pt. */
const DEFAULT_FONT_SIZE_HALF_POINTS = 24
/** 1 twip = 635 EMU, the unit DocxImage uses. */
const EMU_PER_TWIP = 635
/** 96dpi expressed as EMU per pixel, for \picw without \picwgoal. */
const EMU_PER_PX = 914400 / 96
/** Guards against pathological nesting in malformed input. */
const MAX_GROUP_DEPTH = 200

/**
 * windows-1252 for 0x80-0x9f, which differs from latin-1. RTF is a 7-bit format,
 * so every non-ASCII character arrives as a `\'hh` escape and is decoded through
 * this table -- which is why an RTF curly quote survives instead of mojibake.
 */
const CP1252_HIGH = [
  '€', '', '‚', 'ƒ', '„', '…', '†', '‡',
  'ˆ', '‰', 'Š', '‹', 'Œ', '', 'Ž', '',
  '', '‘', '’', '“', '”', '•', '–', '—',
  '˜', '™', 'š', '›', 'œ', '', 'ž', 'Ÿ',
]

/**
 * Charset (`\fcharsetN`) to encoding label. RTF declares the document code page
 * with `\ansicpgN` and each font may override it with `\fcharsetN`; `\'hh`
 * escapes are bytes in *that* encoding, not always windows-1252.
 */
const CHARSET_LABELS: Record<number, string> = {
  77: 'macintosh', 128: 'shift_jis', 129: 'euc-kr', 130: 'euc-kr',
  134: 'gbk', 136: 'big5', 161: 'iso-8859-7', 162: 'iso-8859-9',
  163: 'windows-1255', 177: 'windows-1255', 178: 'windows-1256',
  186: 'iso-8859-4', 204: 'windows-1251', 222: 'windows-874',
  238: 'windows-1250', 254: 'ibm866',
}

/** Code page number (`\ansicpgN`) to encoding label. */
function codePageLabel(cp: number): string {
  if (cp >= 1250 && cp <= 1258) return `windows-${cp}`
  switch (cp) {
    case 874: return 'windows-874'
    case 932: return 'shift_jis'
    case 936: return 'gbk'
    case 949: return 'euc-kr'
    case 950: return 'big5'
    case 10000: return 'macintosh'
    case 65000: return 'utf-7'
    case 65001: return 'utf-8'
    default: return 'windows-1252'
  }
}

const decoderCache = new Map<string, TextDecoder | null>()

/** Exported for tests: which encoding a label resolves to, or null if unsupported. */
export function hasEncodingSupport(label: string): boolean {
  return decoderFor(label) !== null
}

/**
 * Built-in single-byte fallback for windows-1251.
 *
 * The TextDecoder path is the real one — every browser supports windows-1251.
 * But lean server runtimes (bun, or a small-ICU node build) reject the label,
 * and silently falling back to cp1252 would turn Russian into mojibake instead
 * of surfacing the gap. So the Cyrillic block is encoded directly here.
 */
export const CP1251_TABLE: readonly (string | undefined)[] = [
  '\u0000', '\u0001', '\u0002', '\u0003', '\u0004', '\u0005', '\u0006', '\u0007',
  '\u0008', '\u0009', '\u000A', '\u000B', '\u000C', '\u000D', '\u000E', '\u000F',
  '\u0010', '\u0011', '\u0012', '\u0013', '\u0014', '\u0015', '\u0016', '\u0017',
  '\u0018', '\u0019', '\u001A', '\u001B', '\u001C', '\u001D', '\u001E', '\u001F',
  '\u0020', '\u0021', '\u0022', '\u0023', '\u0024', '\u0025', '\u0026', '\u0027',
  '\u0028', '\u0029', '\u002A', '\u002B', '\u002C', '\u002D', '\u002E', '\u002F',
  '\u0030', '\u0031', '\u0032', '\u0033', '\u0034', '\u0035', '\u0036', '\u0037',
  '\u0038', '\u0039', '\u003A', '\u003B', '\u003C', '\u003D', '\u003E', '\u003F',
  '\u0040', '\u0041', '\u0042', '\u0043', '\u0044', '\u0045', '\u0046', '\u0047',
  '\u0048', '\u0049', '\u004A', '\u004B', '\u004C', '\u004D', '\u004E', '\u004F',
  '\u0050', '\u0051', '\u0052', '\u0053', '\u0054', '\u0055', '\u0056', '\u0057',
  '\u0058', '\u0059', '\u005A', '\u005B', '\u005C', '\u005D', '\u005E', '\u005F',
  '\u0060', '\u0061', '\u0062', '\u0063', '\u0064', '\u0065', '\u0066', '\u0067',
  '\u0068', '\u0069', '\u006A', '\u006B', '\u006C', '\u006D', '\u006E', '\u006F',
  '\u0070', '\u0071', '\u0072', '\u0073', '\u0074', '\u0075', '\u0076', '\u0077',
  '\u0078', '\u0079', '\u007A', '\u007B', '\u007C', '\u007D', '\u007E', '\u007F',
  '\u0402', '\u0403', '\u201A', '\u0453', '\u201E', '\u2026', '\u2020', '\u2021',
  '\u20AC', '\u2030', '\u0409', '\u2039', '\u040A', '\u040C', '\u040B', '\u040F',
  '\u0452', '\u2018', '\u2019', '\u201C', '\u201D', '\u2022', '\u2013', '\u2014',
  undefined,           '\u2122', '\u0459', '\u203A', '\u045A', '\u045C', '\u045B', '\u045F',
  '\u00A0', '\u040E', '\u045E', '\u0408', '\u00A4', '\u0490', '\u00A6', '\u00A7',
  '\u0401', '\u00A9', '\u0404', '\u00AB', '\u00AC', '\u00AD', '\u00AE', '\u0407',
  '\u00B0', '\u00B1', '\u0406', '\u0456', '\u0491', '\u00B5', '\u00B6', '\u00B7',
  '\u0451', '\u2116', '\u0454', '\u00BB', '\u0458', '\u0405', '\u0455', '\u0457',
  '\u0410', '\u0411', '\u0412', '\u0413', '\u0414', '\u0415', '\u0416', '\u0417',
  '\u0418', '\u0419', '\u041A', '\u041B', '\u041C', '\u041D', '\u041E', '\u041F',
  '\u0420', '\u0421', '\u0422', '\u0423', '\u0424', '\u0425', '\u0426', '\u0427',
  '\u0428', '\u0429', '\u042A', '\u042B', '\u042C', '\u042D', '\u042E', '\u042F',
  '\u0430', '\u0431', '\u0432', '\u0433', '\u0434', '\u0435', '\u0436', '\u0437',
  '\u0438', '\u0439', '\u043A', '\u043B', '\u043C', '\u043D', '\u043E', '\u043F',
  '\u0440', '\u0441', '\u0442', '\u0443', '\u0444', '\u0445', '\u0446', '\u0447',
  '\u0448', '\u0449', '\u044A', '\u044B', '\u044C', '\u044D', '\u044E', '\u044F',
]

/**
 * Decode windows-1251 from the complete table above, generated from Python's
 * cp1251 codec so it matches the standard exactly (byte 0x98 is genuinely
 * undefined there). The earlier hand-written approximation read past the end of
 * a 32-entry table and had several entries wrong -- 0x8b/0x8c swapped,
 * 0xaa/0xaf/0xb2/0x9a incorrect -- which is what produced literal "undefined".
 */
function decodeWithCp1251(bytes: number[]): string {
  let out = ''
  for (const b of bytes) {
    const mapped = CP1251_TABLE[b]
    // 0x98 is undefined in windows-1251; U+FFFD is the honest rendering.
    out += mapped ?? '\ufffd'
  }
  return out
}

/** A cached TextDecoder, or null when the runtime lacks that encoding. */
function decoderFor(label: string): TextDecoder | null {
  const hit = decoderCache.get(label)
  if (hit !== undefined) return hit
  let made: TextDecoder | null = null
  try {
    made = new TextDecoder(label)
  } catch {
    made = null
  }
  decoderCache.set(label, made)
  return made
}

/**
 * Decode a run of `\'hh` bytes in the given encoding. The whole run is decoded
 * at once so multibyte encodings (Shift-JIS, GBK, Big5, EUC-KR) see a complete
 * lead byte + trail byte pair rather than two independent single-byte decodes.
 */
function decodeBytesIn(bytes: number[], label: string): string {
  // windows-1251 always uses the table, even where TextDecoder supports it.
  // Node decodes the genuinely-undefined byte 0x98 as U+0098 while the WHATWG
  // encoding standard and Python's cp1251 codec treat it as undefined; using
  // the table everywhere keeps output identical across runtimes.
  if (label === 'windows-1251') return decodeWithCp1251(bytes)
  const dec = decoderFor(label)
  if (!dec) {
    // Runtime lacks this encoding. Fall back to a built-in table where we have
    // one, otherwise to cp1252 rather than dropping the characters.
    let out = ''
    for (const b of bytes) out += decodeCp1252(b)
    return out
  }
  try {
    return dec.decode(new Uint8Array(bytes))
  } catch {
    let out = ''
    for (const b of bytes) out += decodeCp1252(b)
    return out
  }
}

let cpDecoder: TextDecoder | null | undefined

function decodeCp1252(byte: number): string {
  if (cpDecoder === undefined) {
    try {
      cpDecoder = new TextDecoder('windows-1252')
    } catch {
      cpDecoder = null // small-ICU runtimes may not know the label
    }
  }
  if (cpDecoder) return cpDecoder.decode(new Uint8Array([byte]))
  if (byte >= 0x80 && byte < 0xa0) return CP1252_HIGH[byte - 0x80]
  return String.fromCharCode(byte)
}

function hexToBytes(hex: string): Uint8Array {
  const clean = hex.length % 2 === 1 ? hex.slice(0, -1) : hex
  const out = new Uint8Array(clean.length / 2)
  for (let k = 0; k < out.length; k++) out[k] = parseInt(clean.substr(k * 2, 2), 16)
  return out
}

function rgb(r: number, g: number, b: number): string {
  const h = (n: number) => Math.max(0, Math.min(255, n)).toString(16).padStart(2, '0')
  return `#${h(r)}${h(g)}${h(b)}`
}

interface CharProps {
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  fontFamily?: string
  /** Font table index from \fN, used to pick the encoding for \'hh escapes. */
  fontIndex?: number
  fontSizePt: number
  color?: string
  highlight?: string
}

interface ParaProps {
  align: ParagraphAlign
  indentLeftTwips: number
  indentRightTwips: number
  indentFirstLineTwips: number
  spacingBeforeTwips: number
  spacingAfterTwips: number
  lineSpacing?: { rule: 'auto' | 'exact' | 'atLeast'; value: number }
  inTable: boolean
  isListPara: boolean
  listMarker?: string
  listLevel: number
}

interface SectionProps {
  paperWidthTwips: number
  paperHeightTwips: number
  leftTwips: number
  rightTwips: number
  topTwips: number
  bottomTwips: number
  headerTwips: number
  footerTwips: number
  gutterTwips: number
  landscape: boolean
  /** True once \paperw/\paperh were supplied for this section. */
  explicitPageSize: boolean
}

interface CellDef {
  rightTwips: number
  borders: DocxTableBorders
  fill?: string
  vAlign: DocxTableCell['vAlign']
}

function defaultChar(): CharProps {
  return {
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    fontSizePt: DEFAULT_FONT_SIZE_HALF_POINTS / 2,
  }
}

function defaultPara(): ParaProps {
  return {
    align: 'left',
    indentLeftTwips: 0,
    indentRightTwips: 0,
    indentFirstLineTwips: 0,
    spacingBeforeTwips: 0,
    spacingAfterTwips: 0,
    inTable: false,
    isListPara: false,
    listLevel: 0,
  }
}

function defaultSection(): SectionProps {
  // US Letter at 1in margins -- the most common Word default.
  return {
    paperWidthTwips: 12240,
    paperHeightTwips: 15840,
    leftTwips: 1440,
    rightTwips: 1440,
    topTwips: 1440,
    bottomTwips: 1440,
    headerTwips: 720,
    footerTwips: 720,
    gutterTwips: 0,
    landscape: false,
    explicitPageSize: false,
  }
}

/**
 * Destinations whose bodies carry no rendered content. Skipping them is both
 * correct and necessary: their bodies contain literal text that must not be
 * painted as document content.
 */
const IGNORABLE = new Set([
  'stylesheet', 'info', 'listtable', 'listoverridetable', 'pgptbl', 'rsidtbl',
  'filetbl', 'themedata', 'colorschememapping', 'datastore', 'latentstyles',
  'xmlnstbl', 'mmathPr', 'upr', 'panose', 'falt', 'listname', 'pntextleveltext',
  'pnseclvl', 'template', 'operator', 'userprops', 'bkmkstart', 'bkmkend',
  'bkmkcolparent', 'field', 'fldinst', 'annotation', 'atrfstart', 'atrfend',
  'do', 'shp', 'shpinst', 'shptxt', 'nonshppict', 'shppict', 'field',
  'header', 'footer', 'headerl', 'headerr', 'headerf', 'footerl', 'footerr',
  'footerf', 'ftnsep', 'ftnsepc', 'ftncn', 'aftnsep', 'aftnsepc', 'aftncn',
  'listpicture', 'xmlopen', 'datafield', 'private', 'pntxtb', 'pnf', 'pnindent',
])

/**
 * The Symbol and Wingdings fonts store glyphs in the Unicode private use area
 * (U+F000-U+F0FF). Word emits bullets as `\u-3913` (U+F0B7), which is
 * meaningless without this mapping, so translate the common private-use glyphs
 * back to real Unicode. Unknown glyphs pass through unchanged.
 */
const SYMBOL_MAP: Record<number, string> = {
  0xa0: '\u2020', 0xa1: '\u2190', 0xa2: '\u2191', 0xa3: '\u2192', 0xa4: '\u2193',
  0xa5: '\u00b7', 0xa6: '\u03b1', 0xa7: '\u00a7', 0xa8: '\u2022', 0xa9: '\u00b6',
  0xaa: '\u00a4', 0xab: '\u00a2', 0xac: '\u221e', 0xad: '\u00b1', 0xae: '\u2264',
  0xaf: '\u2265', 0xb0: '\u2248', 0xb1: '\u0393', 0xb2: '\u03a0', 0xb3: '\u03a3',
  0xb4: '\u03c0', 0xb5: '\u03bc', 0xb6: '\u03c4', 0xb7: '\u2022', 0xb8: '\u03a9',
  0xb9: '\u03b4', 0xba: '\u222e', 0xbb: '\u00f7', 0xbc: '\u2260', 0xbd: '\u2261',
  0xbe: '\u00b1', 0xbf: '\u00d7', 0xc0: '\u221a', 0xc1: '\u222b', 0xc2: '\u2229',
  0xc3: '\u2261', 0xc4: '\u00b0', 0xc5: '\u00b1', 0xc6: '\u00d7', 0xc7: '\u00f7',
  0xc8: '\u2218', 0xc9: '\u2219', 0xca: '\u221a', 0xcb: '\u2205', 0xcc: '\u2208',
  0xcd: '\u2229', 0xce: '\u2261', 0xcf: '\u00b1', 0xd0: '\u03b1', 0xd1: '\u03b2',
  0xd2: '\u03b3', 0xd3: '\u03b4', 0xd4: '\u03b5', 0xd5: '\u03b6', 0xd6: '\u03b7',
  0xd7: '\u03b8', 0xd8: '\u03b9', 0xd9: '\u03ba', 0xda: '\u03bb', 0xdb: '\u00b5',
  0xdc: '\u03bd', 0xdd: '\u03be', 0xde: '\u03bf', 0xdf: '\u03c0', 0xe0: '\u03c1',
  0xe1: '\u03c3', 0xe2: '\u03c2', 0xe3: '\u03c4', 0xe4: '\u03c6', 0xe5: '\u03c7',
  0xe6: '\u03c8', 0xe7: '\u03d1', 0xe8: '\u03c9', 0xe9: '\u03d2', 0xea: '\u03d5',
  0xeb: '\u03d6', 0xec: '\u03dc', 0xed: '\u03dd', 0xee: '\u25cf', 0xef: '\u03d7',
  0xf0: '\u03d8', 0xf1: '\u2261', 0xf2: '\u2265', 0xf3: '\u00b7', 0xf4: '\u2219',
  0xf5: '\u00b0', 0xf6: '\u221a', 0xf7: '\u2248', 0xf8: '\u2261', 0xf9: '\u03b1',
  0xfa: '\u03c0', 0xfb: '\u03c3', 0xfc: '\u03c4', 0xfd: '\u03c6', 0xfe: '\u03c7',
  0xff: '\u03c8',
}

/** Translate Symbol/Wingdings private-use glyphs into real Unicode. */
function mapSymbolText(text: string): string {
  let out = ''
  for (const ch of text) {
    const code = ch.codePointAt(0) ?? 0
    if (code >= 0xf000 && code <= 0xf0ff) {
      out += SYMBOL_MAP[code - 0xf000] ?? ch
    } else {
      out += ch
    }
  }
  return out
}

class RtfParser {
  private readonly s: string
  private i = 0
  private depth = 0

  private fonts: (string | undefined)[] = []
  private colors: (string | undefined)[] = []
  /** Per-font `\fcharsetN`, used to pick the encoding for `\'hh` escapes. */
  private fontCharsets: (number | undefined)[] = []
  /** `\ansicpgN`, the document code page (default windows-1252). */
  private codePage = 1252
  /** `\deffN`, the default font index applied by `\plain`. */
  private defaultFontIndex = 0
  /**
   * Pending `\'hh` bytes. Buffered rather than decoded one at a time so
   * multibyte encodings receive a complete character.
   */
  private hexBuf: number[] = []
  /**
   * Encoding in force when the buffered bytes were read. Captured eagerly:
   * bytes are decoded at flush time, by which point a later `\fN` may have
   * changed the font, which would decode them with the wrong code page.
   */
  private hexLabel = 'windows-1252'
  /** Fallback characters to drop after \uN -- the \ucN count. */
  private skipChars = 0
  private unicodeSkipCount = 1

  private blocks: DocxBlock[] = []
  private sections: DocxSection[] = []
  private section = defaultSection()

  private runs: DocxTextRun[] = []
  private images: DocxImage[] = []
  private textBuffer = ''
  private bufferStyleKey = ''
  /**
   * The formatting that applies to the text already in `textBuffer`. This must
   * be a snapshot rather than reading `this.char` at flush time: when `\b0`
   * turns bold off, the bold text that came *before* it still needs to be
   * flushed as bold, not as the new unbold state.
   */
  private bufferStyle: CharProps = defaultChar()
  private char: CharProps = defaultChar()
  private para: ParaProps = defaultPara()
  /** Set by \line so the next run starts on a fresh line. */
  private pendingBreak = false

  private rowDefs: CellDef[] = []
  private cellDefInProgress: CellDef | null = null
  private rowCells: DocxTableCell[] = []
  private cellParas: DocxParagraph[] = []
  private openTable: DocxTable | null = null

  constructor(source: string) {
    this.s = source
  }

  parse(): DocxDocument {
    // Any leading bytes before the first group are a BOM or stray text.
    const start = this.s.indexOf('{')
    if (start < 0) return this.emptyDocument()
    this.i = start
    this.parseGroup()
    // Hex bytes may still be buffered: a document can end with `\'e9` and no
    // \par at all. Detect content only after they become text.
    this.flushHex()
    // Only flush if something is actually pending: `text\par` ends a paragraph,
    // and emitting a trailing empty one would add a phantom blank line.
    if (this.textBuffer.length > 0 || this.runs.length > 0 || this.images.length > 0) {
      this.flushParagraph()
    }
    this.closeTable()
    if (this.blocks.length > 0 || this.sections.length === 0) {
      this.sections.push(this.finalizeSection(this.blocks))
    }
    return {
      sections: this.sections,
      defaultFontFamily: 'Calibri',
      defaultFontSizePt: DEFAULT_FONT_SIZE_HALF_POINTS / 2,
      styleDefaults: new Map(),
    }
  }

  private emptyDocument(): DocxDocument {
    return {
      sections: [],
      defaultFontFamily: 'Calibri',
      defaultFontSizePt: DEFAULT_FONT_SIZE_HALF_POINTS / 2,
      styleDefaults: new Map(),
    }
  }

  // ---- output plumbing -------------------------------------------------

  private finalizeSection(blocks: DocxBlock[]): DocxSection {
    const s = this.section
    const margins: DocxPageMargins = {
      topTwips: s.topTwips,
      rightTwips: s.rightTwips,
      bottomTwips: s.bottomTwips,
      leftTwips: s.leftTwips,
      headerTwips: s.headerTwips,
      footerTwips: s.footerTwips,
      gutterTwips: s.gutterTwips,
    }
    const paragraphs = blocks.filter((b) => b.kind === 'p').map((b) => (b as { paragraph: DocxParagraph }).paragraph)
    return {
      margins,
      pageSize: {
        // Writers that set \landscape already store the page in its FINAL
        // orientation (\paperw15840 \paperh12240), so swapping again would
        // render portrait. Honour explicit dimensions as given; only rotate the
        // built-in portrait default when the section never declared a size.
        widthTwips: s.landscape && !s.explicitPageSize ? s.paperHeightTwips : s.paperWidthTwips,
        heightTwips: s.landscape && !s.explicitPageSize ? s.paperWidthTwips : s.paperHeightTwips,
        orientation: s.landscape ? 'landscape' : 'portrait',
      },
      paragraphs,
      blocks,
    }
  }

  private currentParagraph(): DocxParagraph {
    this.flushRun()
    const p: DocxParagraph = {
      runs: this.runs,
      images: this.images,
      align: this.para.align,
      indentLeftTwips: this.para.indentLeftTwips,
      indentRightTwips: this.para.indentRightTwips,
      indentFirstLineTwips: this.para.indentFirstLineTwips,
      spacingBeforeTwips: this.para.spacingBeforeTwips,
      spacingAfterTwips: this.para.spacingAfterTwips,
      lineSpacing: this.para.lineSpacing,
    }
    if (this.para.isListPara) {
      p.listMarker = this.para.listMarker ?? '•'
      p.listLevel = this.para.listLevel
    }
    return p
  }

  private flushRun(): void {
    if (this.textBuffer.length === 0) return
    const style = this.bufferStyle
    const run: DocxTextRun = { text: this.textBuffer }
    if (style.bold) run.bold = true
    if (style.italic) run.italic = true
    if (style.underline) run.underline = true
    if (style.strike) run.strike = true
    if (style.fontFamily) run.fontFamily = style.fontFamily
    if (style.fontSizePt !== DEFAULT_FONT_SIZE_HALF_POINTS / 2) run.fontSizePt = style.fontSizePt
    if (style.color) run.color = style.color
    if (style.highlight) run.highlight = style.highlight
    if (this.pendingBreak) run.breakBefore = true
    this.pendingBreak = false
    this.runs.push(run)
    this.textBuffer = ''
    this.bufferStyleKey = ''
  }

  /**
   * The encoding label for the bytes currently being read: the active font's
   * charset when it declares one, otherwise the document code page.
 */
  private activeEncodingLabel(): string {
    // Fall back to the \deff font when no \fN has been applied yet (or a group
    // exit restored that state), so the document's declared default charset
    // governs \'hh bytes from the very first character.
    const fontIndex = this.char.fontIndex ?? (this.fonts[this.defaultFontIndex] ? this.defaultFontIndex : undefined)
    return this.encodingLabelFor(fontIndex)
  }

  /**
   * The encoding that `\'hh` escapes in the given font decode with. Shared with
   * marker capture so a hex bullet inside {\pntext} or {\listtext} resolves the
   * same way as hex in ordinary body text.
   */
  private encodingLabelFor(fontIndex: number | undefined, codePage: number = this.codePage): string {
    if (fontIndex !== undefined) {
      const charset = this.fontCharsets[fontIndex]
      // charset 0 (ANSI) and 1 (Default) defer to \ansicpg; 2 is Symbol, whose
      // private-use glyphs must not be run through a text decoder at all.
      if (charset === 2) return 'symbol'
      if (charset !== undefined && charset !== 0 && charset !== 1) {
        const label = CHARSET_LABELS[charset]
        if (label) return label
      }
    }
    return codePageLabel(codePage)
  }

  /**
   * Adopt the \deff default font unless an explicit \fN has already been seen.
   * Called once the font table is readable, because \deff normally appears
   * BEFORE it in the header.
   */
  private applyDefaultFont(): void {
    if (this.char.fontIndex !== undefined) return
    const name = this.fonts[this.defaultFontIndex]
    if (name === undefined) return
    this.char.fontIndex = this.defaultFontIndex
    this.char.fontFamily = name
  }

  /**
   * Decode a run of `\'hh` bytes. Multibyte encodings need the whole run so a
   * lead byte meets its trail byte; Symbol charset bytes are glyph indices, so
   * they go through the Symbol table (0xB7 is a bullet, not a middot).
   */
  private static decodeHexBytes(bytes: number[], label: string): string {
    if (label === 'symbol') {
      let out = ''
      for (const b of bytes) out += SYMBOL_MAP[b] ?? String.fromCharCode(b)
      return out
    }
    return decodeBytesIn(bytes, label)
  }

  /** Decode and emit any buffered `\'hh` bytes. */
  private flushHex(): void {
    if (this.hexBuf.length === 0) return
    const bytes = this.hexBuf
    this.hexBuf = []
    this.emit(RtfParser.decodeHexBytes(bytes, this.hexLabel))
  }

  /** Emit a literal, flushing any buffered hex bytes first so order is kept. */
private emitLiteral(text: string): void {
    this.flushHex()
    this.emit(text)
  }

  private emit(text: string): void {
    if (text.length === 0) return
    if (this.skipChars > 0) {
      // \uN is followed by \ucN fallback characters that must not be rendered.
      let drop = this.skipChars
      let kept = ''
      for (const ch of text) {
        if (drop > 0) {
          drop--
          continue
        }
        kept += ch
      }
      this.skipChars = drop
      if (kept.length === 0) return
      text = kept
    }
    // Start a new run only when the character formatting actually changes, so
    // a plain sentence stays one run instead of one run per letter.
    text = mapSymbolText(text)
    const key = this.styleKey()
    if (this.textBuffer.length > 0 && key !== this.bufferStyleKey) this.flushRun()
    if (this.textBuffer.length === 0) this.bufferStyle = { ...this.char }
    this.textBuffer += text
    this.bufferStyleKey = key
  }

  private styleKey(): string {
    const c = this.char
    return `${c.bold ? 1 : 0}${c.italic ? 1 : 0}${c.underline ? 1 : 0}${c.strike ? 1 : 0}|${c.fontFamily ?? ''}|${c.fontSizePt}|${c.color ?? ''}|${c.highlight ?? ''}`
  }

  /** End the current paragraph, routing it into a table cell or the body. */
  private flushParagraph(): void {
    this.flushHex()
    this.flushRun()
    const hadContent = this.runs.length > 0 || this.images.length > 0
    const p: DocxParagraph = hadContent ? this.currentParagraph() : {
      runs: [],
      images: [],
      align: this.para.align,
    }
    const inTable = this.para.inTable
    const isListPara = this.para.isListPara
    if (inTable) {
      this.cellParas.push(p)
    } else {
      if (hadContent) this.closeTable()
      this.blocks.push({ kind: 'p', paragraph: p })
    }
    this.runs = []
    this.images = []
    this.textBuffer = ''
    this.pendingBreak = false
    // Per the RTF spec paragraph properties (alignment, indents, spacing) PERSIST
    // until \pard or an explicit override -- resetting them here made
    // \qc A\par B\par centre only the first paragraph. Keep the visual
    // properties; only per-paragraph list membership is cleared, since that is
    // re-declared on every list paragraph. \intbl is kept so a multi-paragraph
    // cell stays in the cell (it is cleared at \row).
    this.para.isListPara = false
    this.para.listMarker = undefined
    this.para.listLevel = 0
    void isListPara
  }

  // ---- tables ----------------------------------------------------------

  private ensureTable(): DocxTable {
    if (!this.openTable) {
      this.openTable = {
        gridColsTwips: [],
        rows: [],
        cellMargins: { topTwips: 0, rightTwips: 108, bottomTwips: 0, leftTwips: 108 },
      }
    }
    return this.openTable
  }

  private closeTable(): void {
    if (!this.openTable) return
    const table = this.openTable
    this.openTable = null
    this.rowDefs = []
    if (table.rows.length === 0) return
    // Column widths come from the cell right edges of the first row.
    const cols: number[] = []
    let prev = 0
    for (const cell of table.rows[0].cells) {
      const w = cell.widthTwips ?? 1440
      cols.push(Math.max(1, w - prev))
      prev = w
    }
    table.gridColsTwips = cols.length > 0 ? cols : [1440]
    this.blocks.push({ kind: 'table', table })
  }

  private cellDef(): CellDef {
    if (!this.cellDefInProgress) {
      this.cellDefInProgress = { rightTwips: 0, borders: {}, vAlign: 'top' }
    }
    return this.cellDefInProgress
  }

  private endCell(): void {
    // A cell's content does not need a trailing \par -- \cell ends it. Flush any
    // pending paragraph first or the cell renders empty.
    if (this.textBuffer.length > 0 || this.runs.length > 0 || this.images.length > 0) {
      this.para.inTable = true
      this.flushParagraph()
    }
    const cell: DocxTableCell = {
      gridSpan: 1,
      paragraphs: this.cellParas.length > 0 ? this.cellParas : [{ runs: [], images: [], align: 'left' }],
    }
    const def = this.rowDefs[this.rowCells.length]
    if (def) {
      if (def.rightTwips > 0) cell.widthTwips = def.rightTwips
      if (Object.keys(def.borders).length > 0) cell.borders = def.borders
      if (def.fill) cell.fill = def.fill
      if (def.vAlign !== 'top') cell.vAlign = def.vAlign
    }
    this.rowCells.push(cell)
    this.cellParas = []
    this.runs = []
    this.images = []
    this.textBuffer = ''
  }

  private endRow(): void {
    const table = this.ensureTable()
    // Pad short rows so every row matches its cell definitions.
    const target = Math.max(this.rowDefs.length, this.rowCells.length)
    while (this.rowCells.length < target) {
      const def = this.rowDefs[this.rowCells.length]
      this.rowCells.push({
        gridSpan: 1,
        paragraphs: [{ runs: [], images: [], align: 'left' }],
        ...(def?.rightTwips ? { widthTwips: def.rightTwips } : {}),
      })
    }
    table.rows.push({ cells: this.rowCells } as DocxTableRow)
    this.rowCells = []
    this.cellParas = []
    // Leaving table context: paragraph properties now persist across \par, so
    // without this the next body paragraph would still be treated as in-cell.
    this.para.inTable = false
    // Row definitions persist: a following row may omit \trowd and \cellx.
  }

  // ---- pictures --------------------------------------------------------

  /** Parse a `{\pict ...}` destination: hex payload plus sizing control words. */
  private parsePicture(): DocxImage | undefined {
    let kind = ''
    let unsupported = false
    let picw = 0
    let pich = 0
    let goalW = 0
    let goalH = 0
    let hex = ''
    let depth = 1
    while (this.i < this.s.length && depth > 0) {
      const c = this.s[this.i]
      if (c === '{') {
        depth++
        this.i++
        continue
      }
      if (c === '}') {
        depth--
        this.i++
        continue
      }
      if (c === '\\') {
        const w = this.readControlWord()
        if (w) {
          if (w.name === 'pngblip') kind = 'png'
          else if (w.name === 'jpegblip' || w.name === 'jpgblip') kind = 'jpeg'
          else if (w.name === 'emfblip' || w.name === 'wmetafile' || w.name === 'dibitmap' || w.name === 'wbitmap') unsupported = true
          else if (w.name === 'picw') picw = w.param ?? 0
          else if (w.name === 'pich') pich = w.param ?? 0
          else if (w.name === 'picwgoal') goalW = w.param ?? 0
          else if (w.name === 'pichgoal') goalH = w.param ?? 0
          continue
        }
        this.i++
        continue
      }
      if ((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F')) hex += c
      this.i++
    }
    if (unsupported || !kind) return undefined
    const data = hexToBytes(hex)
    if (data.length === 0) return undefined
    const widthEmu = goalW > 0 ? goalW * EMU_PER_TWIP : Math.round((picw || 100) * EMU_PER_PX)
    const heightEmu = goalH > 0 ? goalH * EMU_PER_TWIP : Math.round((pich || 100) * EMU_PER_PX)
    return { data, mime: kind === 'png' ? 'image/png' : 'image/jpeg', widthEmu, heightEmu }
  }

  // ---- tokenizer -------------------------------------------------------

  /**
   * Read a control word. `this.i` may point at the leading backslash or at the
   * first letter -- both call sites exist, and normalising here keeps them from
   * silently eating the first character (which is how an earlier version turned
   * every `\b` into a literal "b").
   */
  private readControlWord(): { name: string; param?: number } | null {
    let p = this.i
    if (this.s[p] === '\\') p++
    const first = this.s[p]
    if (first === undefined || !/[A-Za-z]/.test(first)) return null
    let name = ''
    while (p < this.s.length && /[A-Za-z]/.test(this.s[p])) {
      name += this.s[p]
      p++
    }
    let param: number | undefined
    if (this.s[p] === '-') {
      p++
      let digits = ''
      while (p < this.s.length && this.s[p] >= '0' && this.s[p] <= '9') {
        digits += this.s[p]
        p++
      }
      param = -parseInt(digits || '0', 10)
    } else if (this.s[p] !== undefined && this.s[p] >= '0' && this.s[p] <= '9') {
      let digits = ''
      while (p < this.s.length && this.s[p] >= '0' && this.s[p] <= '9') {
        digits += this.s[p]
        p++
      }
      param = parseInt(digits, 10)
    }
    if (this.s[p] === ' ') p++
    this.i = p
    return { name, param }
  }

  /**
   * Parse one group through its matching `}`.
   *
   * `this.char` / `this.para` are the single source of truth for the current
   * state. Nested groups save and restore them rather than being handed copies:
   * `emit` and `flushRun` also read these fields, so passing clones (as an
   * earlier version did) formatted a copy that nothing ever looked at.
   */
  private parseGroup(): void {
    if (++this.depth > MAX_GROUP_DEPTH) {
      this.depth--
      this.skipRestOfGroup()
      return
    }
    while (this.i < this.s.length) {
      const c = this.s[this.i]

      if (c === '}') {
        this.i++
        break
      }

      if (c === '{') {
        this.i++
        let innerStar = false
        if (this.s[this.i] === '\\' && this.s[this.i + 1] === '*') {
          innerStar = true
          this.i += 2
        }
        const save = this.i
        const w = this.readControlWord()
        if (w) {
          if (w.name === 'fonttbl') {
            this.parseFontTable()
            continue
          }
          if (w.name === 'colortbl') {
            this.parseColorTable()
            continue
          }
          if (w.name === 'pict') {
            const img = this.parsePicture()
            if (img) this.images.push(img)
            continue
          }
          if (w.name === 'pntext') {
            // Grouped {\pntext\f2\pnindent0{\pntxtb 1.}} is the form Word
            // actually writes, and it carries the marker itself. Capture it
            // rather than skipping the group, or a numbered legacy list falls
            // back to a bullet. Unicode/Symbol glyphs are translated the same
            // way \listtext markers are.
            this.para.isListPara = true
            const marker = this.captureGroupText()
            if (marker) this.para.listMarker = marker
            continue
          }
          if (w.name === 'listtext') {
            const marker = this.captureGroupText()
            if (marker) this.para.listMarker = marker
            continue
          }
          if (innerStar || IGNORABLE.has(w.name)) {
            this.skipRestOfGroup()
            continue
          }
          this.i = save // normal group: let the loop inside handle it
        } else {
          this.i = save
        }
        // Group scope: properties revert at the closing brace, but accumulated
        // text, runs and blocks deliberately do not.
        const savedChar = { ...this.char }
        const savedPara = { ...this.para }
        const savedSkipCount = this.unicodeSkipCount
        this.parseGroup()
        // Inner-group bytes belong to the inner formatting: decode them before
        // restoring the enclosing scope.
        this.flushHex()
        this.char = savedChar
        this.para = savedPara
        // \uc is document-scoped but overridden within a group, so it must be
        // restored too or an inner \uc0 leaks out and desynchronises the
        // fallback skipping for every later \uN.
        this.unicodeSkipCount = savedSkipCount
        continue
      }

      if (c === '\\') {
        this.handleBackslash(this.char, this.para)
        continue
      }

      // CR/LF are insignificant in RTF.
      if (c === '\r' || c === '\n') {
        this.i++
        continue
      }
      this.emitLiteral(c)
      this.i++
    }
    this.depth--
  }

  /**
   * Read a group's literal text, used for `{\listtext ...}` markers. Unlike
   * `emit`, this runs outside the paragraph's formatting, so it tracks the font
   * itself to translate Symbol glyphs, and handles \uN plus its \ucN fallback.
   */
  private captureGroupText(): string {
    let out = ''
    let depth = 1
    let skipChars = 0
    let done = false
    // Marker groups carry their own \fN and \uc, and \'hh escapes must decode
    // through the same code-page logic as body text -- a Symbol bullet written
    // as \'b7 previously fell through as the literal text "b7".
    let fontIndex: number | undefined
    let uc = this.unicodeSkipCount
    // A marker group may switch code page; keep it local so bytes decode the
    // same way they would in body text, without leaking back out.
    let codePage = this.codePage
    let hex: number[] = []
    let hexLabel = this.encodingLabelFor(this.char.fontIndex, codePage)

    const flushHex = (): void => {
      if (hex.length === 0) return
      const bytes = hex
      hex = []
      if (!done) out += RtfParser.decodeHexBytes(bytes, hexLabel)
    }
    const pushByte = (byte: number): void => {
      const label = this.encodingLabelFor(fontIndex ?? this.char.fontIndex, codePage)
      if (hex.length > 0 && hexLabel !== label) flushHex()
      if (hex.length === 0) hexLabel = label
      hex.push(byte)
    }

    while (this.i < this.s.length && depth > 0) {
      const c = this.s[this.i]
      if (c === '{') {
        depth++
        this.i++
        continue
      }
      if (c === '}') {
        depth--
        this.i++
        continue
      }
      if (c === '\\') {
        // 'hh must be recognised before readControlWord, which would otherwise
        // decline it and leave the two hex digits as literal marker text.
        if (this.s[this.i + 1] === "'") {
          const byte = parseInt(this.s.substr(this.i + 2, 2), 16)
          this.i += 4
          if (!Number.isNaN(byte)) pushByte(byte)
          continue
        }
        const w = this.readControlWord()
        if (w) {
          if (w.name === 'u' && w.param !== undefined) {
            flushHex()
            const code = w.param < 0 ? w.param + 65536 : w.param
            if (code >= 0 && code <= 0x10ffff && !done) {
              out += mapSymbolText(String.fromCodePoint(code))
            }
            skipChars = uc
          } else if (w.name === 'uc') {
            // \uc is group scoped: honour it here without leaking it back out.
            uc = w.param ?? 1
          } else if (w.name === 'f') {
            fontIndex = w.param
          } else if (w.name === 'ansicpg' && w.param !== undefined) {
            flushHex()
            codePage = w.param
          } else if (w.name === 'tab' || w.name === 'cell') {
            // The marker ends at the tab, but we must keep consuming through the
            // group's closing brace -- bailing out here would leave the `}` to be
            // read as the end of the *enclosing* group.
            flushHex()
            done = true
          }
          continue
        }
        this.i++
        const sym = this.s[this.i]
        this.i++
        flushHex()
        if (sym === '\\' || sym === '{' || sym === '}') out += sym
        continue
      }
      if (c === '\r' || c === '\n') {
        this.i++
        continue
      }
      if (done) {
        this.i++
        continue
      }
      if (skipChars > 0) {
        skipChars--
        this.i++
        continue
      }
      // Plain text must not jump ahead of buffered hex bytes.
      flushHex()
      out += mapSymbolText(c)
      this.i++
    }
    flushHex()
    return out.replace(/\s+$/, '').replace(/^[ \t]+/, '')
  }

  /** Consume through the `}` that closes the group we are inside. */
  private skipRestOfGroup(): void {
    let depth = 1
    while (this.i < this.s.length && depth > 0) {
      const c = this.s[this.i]
      if (c === '\\') {
        this.i++
        if (this.s[this.i] === "'") {
          this.i += 3
          continue
        }
        this.i++
        continue
      }
      if (c === '{') depth++
      else if (c === '}') depth--
      this.i++
    }
  }

  private handleBackslash(char: CharProps, para: ParaProps): void {
    const c = this.s[this.i + 1]
    if (c === undefined) {
      this.i++
      return
    }

    if (c === "'") {
      const byte = parseInt(this.s.substr(this.i + 2, 2), 16)
      this.i += 4
      if (!Number.isNaN(byte)) {
        const label = this.activeEncodingLabel()
        // If the code page changed since the previous byte, decode what we have
        // rather than mixing encodings inside one character.
        if (this.hexBuf.length > 0 && this.hexLabel !== label) this.flushHex()
        if (this.hexBuf.length === 0) this.hexLabel = label
        this.hexBuf.push(byte)
      }
      return
    }
    if (!/[A-Za-z]/.test(c)) {
      this.i += 2
      this.flushHex()
      switch (c) {
        case '\\':
          this.emit('\\')
          break
        case '{':
          this.emit('{')
          break
        case '}':
          this.emit('}')
          break
        case '_':
          this.emit('-') // non-breaking hyphen
          break
        case '~':
          this.emit('\u00a0')
          break
        default:
          break
      }
      return
    }
    const w = this.readControlWord()
    if (w) this.applyControlWord(w.name, w.param, char, para)
  }

  private applyControlWord(name: string, param: number | undefined, char: CharProps, para: ParaProps): void {
    // Buffered \'hh bytes were read under the CURRENT formatting, so they must
    // become text before any control word can change that formatting. Without
    // this, `\'e9\b bold` folded the accent into the following bold run, and a
    // nested group ending in hex lost its style when the group exited.
    this.flushHex()
    switch (name) {
      // ---- character formatting ----
      case 'b': char.bold = param !== 0; return
      case 'i': char.italic = param !== 0; return
      case 'ul': char.underline = param !== 0; return
      case 'ulnone': char.underline = false; return
      case 'uldb': char.underline = true; return
      case 'strike': char.strike = param !== 0; return
      case 'plain': {
        // Object.assign against defaultChar() would NOT clear these: the default
        // object simply has no such keys, so the old values survive. Reset them
        // explicitly and fall back to the document's default font (\deff).
        const defaults = defaultChar()
        char.bold = defaults.bold
        char.italic = defaults.italic
        char.underline = defaults.underline
        char.strike = defaults.strike
        char.fontSizePt = defaults.fontSizePt
        char.color = undefined
        char.highlight = undefined
        char.fontIndex = undefined
        this.applyDefaultFont()
        return
      }
      case 'fs':
        if (param !== undefined) char.fontSizePt = param / 2
        return
      case 'f':
        this.flushHex()
        char.fontIndex = param
        char.fontFamily = param !== undefined ? this.fonts[param] : undefined
        return
      case 'ansicpg':
        this.flushHex()
        this.codePage = param ?? this.codePage
        return
      case 'deff':
        this.defaultFontIndex = param ?? 0
        this.applyDefaultFont()
        return
      case 'cf':
        char.color = param !== undefined ? this.colors[param] : undefined
        return
      case 'cb':
      case 'highlight':
        char.highlight = param !== undefined ? this.colors[param] : undefined
        return
      case 'uc':
        this.unicodeSkipCount = param ?? 1
        return
      case 'u': {
        if (param === undefined) return
        // Negative parameters are 16-bit signed; recover the BMP code point.
        const code = param < 0 ? param + 65536 : param
        if (code >= 0 && code <= 0x10ffff) this.emitLiteral(String.fromCodePoint(code))
        this.skipChars = this.unicodeSkipCount
        return
      }

      // ---- paragraph structure ----
      case 'par': this.flushHex(); this.flushParagraph(); return
      case 'line':
        this.flushHex()
        this.flushRun()
        this.pendingBreak = true
        return
      case 'pard':
        this.para = defaultPara()
        return
      case 'ql': para.align = 'left'; return
      case 'qr': para.align = 'right'; return
      case 'qc': para.align = 'center'; return
      case 'qj': para.align = 'justify'; return
      case 'li': para.indentLeftTwips = param ?? 0; return
      case 'ri': para.indentRightTwips = param ?? 0; return
      case 'fi': para.indentFirstLineTwips = param ?? 0; return
      case 'sb': para.spacingBeforeTwips = param ?? 0; return
      case 'sa': para.spacingAfterTwips = param ?? 0; return
      case 'sl':
        if (param === undefined || param === 0) para.lineSpacing = undefined
        else if (param < 0) para.lineSpacing = { rule: 'exact', value: -param }
        else para.lineSpacing = { rule: 'atLeast', value: param }
        return
      case 'tx': return // absolute tab stops: the model uses \t in run text

      // ---- tables ----
      case 'trowd': this.rowDefs = []; this.cellDefInProgress = null; return
      case 'cellx': {
        const def = this.cellDef()
        def.rightTwips = param ?? 0
        this.rowDefs.push(def)
        this.cellDefInProgress = null
        return
      }
      case 'intbl': para.inTable = true; return
      case 'cell': this.flushHex(); this.endCell(); return
      case 'row': this.endRow(); return
      case 'clvertalc': this.cellDef().vAlign = 'center'; return
      case 'clvertalb': this.cellDef().vAlign = 'bottom'; return
      case 'clcbpat':
      case 'clcfpat':
        this.cellDef().fill = param !== undefined ? this.colors[param] : undefined
        return
      case 'clbrdrt': this.cellDef().borders.top = borderStyle(param); return
      case 'clbrdrl': this.cellDef().borders.left = borderStyle(param); return
      case 'clbrdrb': this.cellDef().borders.bottom = borderStyle(param); return
      case 'clbrdrr': this.cellDef().borders.right = borderStyle(param); return

      // ---- lists ----
      case 'pntext': para.isListPara = true; return
      case 'ilvl': para.listLevel = param ?? 0; return
      case 'ls':
        // \lsN is the list identifier. Word usually repeats it on every list
        // paragraph, so treat a non-zero id as list membership on its own.
        if (param !== undefined && param !== 0) this.para.isListPara = true
        return

      // ---- section / page setup ----
      case 'sectd': this.section = defaultSection(); return
      case 'sect':
        this.flushParagraph()
        this.closeTable()
        if (this.blocks.length > 0) {
          this.sections.push(this.finalizeSection(this.blocks))
          this.blocks = []
        }
        return
      case 'paperw':
        this.section.paperWidthTwips = param ?? this.section.paperWidthTwips
        this.section.explicitPageSize = true
        return
      case 'paperh':
        this.section.paperHeightTwips = param ?? this.section.paperHeightTwips
        this.section.explicitPageSize = true
        return
      case 'margl': this.section.leftTwips = param ?? this.section.leftTwips; return
      case 'margr': this.section.rightTwips = param ?? this.section.rightTwips; return
      case 'margt': this.section.topTwips = param ?? this.section.topTwips; return
      case 'margb': this.section.bottomTwips = param ?? this.section.bottomTwips; return
      case 'gutter': this.section.gutterTwips = param ?? 0; return
      case 'headery': this.section.headerTwips = param ?? this.section.headerTwips; return
      case 'footery': this.section.footerTwips = param ?? this.section.footerTwips; return
      case 'landscape':
      case 'lndscpsxn': this.section.landscape = true; return

      // ---- literal characters ----
      case 'tab': this.emitLiteral('\t'); return
      case 'emdash': this.emitLiteral('\u2014'); return
      case 'endash': this.emitLiteral('\u2013'); return
      case 'lquote': this.emitLiteral('\u2018'); return
      case 'rquote': this.emitLiteral('\u2019'); return
      case 'ldblquote': this.emitLiteral('\u201c'); return
      case 'rdblquote': this.emitLiteral('\u201d'); return
      case 'bullet': this.emitLiteral('\u2022'); return
      case 'enspace':
      case 'emspace':
      case 'qmspace': this.emitLiteral(' '); return
      default:
        return
    }
  }

  // ---- destination tables ---------------------------------------------

  private parseFontTable(): void {
    let index = 0
    let charset: number | undefined
    let name = ''
    let depth = 1
    const commit = (): void => {
      const trimmed = name.trim()
      if (trimmed.length > 0 || this.fonts[index] === undefined) this.fonts[index] = trimmed
      // Only write when a charset was actually seen. The unconditional write
      // clobbered a charset recorded by the preceding entry, because the final
      // commit() below runs with charset already reset to undefined.
      if (charset !== undefined) this.fontCharsets[index] = charset
      name = ''
      charset = undefined
    }
    while (this.i < this.s.length && depth > 0) {
      const c = this.s[this.i]
      if (c === '{') {
        depth++
        this.i++
        continue
      }
      if (c === '}') {
        depth--
        this.i++
        if (depth === 0) break
        continue
      }
      if (c === '\\') {
        const w = this.readControlWord()
        if (w) {
          if (w.name === 'f' && w.param !== undefined) {
            commit()
            index = w.param
          } else if (w.name === 'fcharset' && w.param !== undefined) {
            charset = w.param
          }
          continue
        }
        this.i++
        const sym = this.s[this.i]
        this.i++
        if (sym === ';') commit()
        continue
      }
      if (c === ';') {
        commit()
        this.i++
        continue
      }
      if (c !== '\r' && c !== '\n') name += c
      this.i++
    }
    commit()
    // Now that font names are known, adopt \deff if no \fN has been applied.
    this.applyDefaultFont()
  }

  private parseColorTable(): void {
    let red = 0
    let green = 0
    let blue = 0
    let seen = false
    let depth = 1
    while (this.i < this.s.length && depth > 0) {
      const c = this.s[this.i]
      if (c === '{') {
        depth++
        this.i++
        continue
      }
      if (c === '}') {
        depth--
        this.i++
        if (depth === 0) break
        continue
      }
      if (c === '\\') {
        const w = this.readControlWord()
        if (w) {
          const v = w.param ?? 0
          if (w.name === 'red') { red = v; seen = true }
          else if (w.name === 'green') { green = v; seen = true }
          else if (w.name === 'blue') { blue = v; seen = true }
          continue
        }
        this.i++
        this.i++
        continue
      }
      if (c === ';') {
        // An entry with no colour components means "automatic".
        this.colors.push(seen ? rgb(red, green, blue) : undefined)
        red = 0
        green = 0
        blue = 0
        seen = false
        this.i++
        continue
      }
      this.i++
    }
  }
}

function borderStyle(kind?: number): TableCellBorder {
  // \brdrnil none, \brdrth thick, \brdrdb double, everything else a hairline.
  if (kind === 0) return { style: 'none' }
  if (kind === 5) return { style: 'thick' }
  if (kind === 6) return { style: 'double' }
  return { style: 'single' }
}

/**
 * True when the bytes are an RTF document. RTF is plain text, not a zip, so it
 * must be recognised before anything reaches OfficePackage.load, which would
 * throw on it. Leading whitespace and a UTF-8 BOM are tolerated because Word
 * emits them.
 */
export function isRtf(data: Uint8Array): boolean {
  let i = 0
  if (data[0] === 0xef && data[1] === 0xbb && data[2] === 0xbf) i = 3
  while (i < data.length) {
    const b = data[i]
    if (b === 0x20 || b === 0x09 || b === 0x0a || b === 0x0d) {
      i++
      continue
    }
    break
  }
  // "{\rtf" — 0x7b 0x5c 0x72 0x74 0x66
  return data[i] === 0x7b && data[i + 1] === 0x5c && data[i + 2] === 0x72 && data[i + 3] === 0x74 && data[i + 4] === 0x66
}

/**
 * Parse an RTF document. Accepts raw bytes (the usual case, straight off a
 * fetch) or a string.
 */
export function parseRtf(input: Uint8Array | string): DocxDocument {
  // Deliberately NOT utf-8: RTF carries high bytes as \'hh escapes, and a
  // UTF-8 pass would mangle them. Map bytes one-to-one instead.
  const text = typeof input === 'string' ? input : decodeBytesLatin1(input)
  return new RtfParser(text).parse()
}

function decodeBytesLatin1(bytes: Uint8Array): string {
  let out = ''
  for (let i = 0; i < bytes.length; i += 8192) {
    const slice = bytes.subarray(i, Math.min(i + 8192, bytes.length))
    for (const b of slice) out += String.fromCharCode(b)
  }
  return out
}