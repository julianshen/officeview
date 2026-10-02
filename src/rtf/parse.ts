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
  'listpicture', 'pntext', 'xmlopen', 'datafield', 'private', 'themedata',
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
        widthTwips: s.landscape ? s.paperHeightTwips : s.paperWidthTwips,
        heightTwips: s.landscape ? s.paperWidthTwips : s.paperHeightTwips,
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
    // \pard resets paragraph properties; table membership and list flags are
    // re-established by the control words that follow it.
    this.para = defaultPara()
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
        this.parseGroup()
        this.char = savedChar
        this.para = savedPara
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
      this.emit(c)
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
          if (w.name === 'u' && w.param !== undefined) {
            const code = w.param < 0 ? w.param + 65536 : w.param
            if (code >= 0 && code <= 0x10ffff && !done) {
              const ch = String.fromCodePoint(code)
              out += mapSymbolText(ch)
            }
            skipChars = this.unicodeSkipCount
          } else if (w.name === "tab" || w.name === 'cell') {
            // The marker ends at the tab, but we must keep consuming through the
            // group's closing brace -- bailing out here would leave the `}` to be
            // read as the end of the *enclosing* group.
            done = true
          }
          continue
        }
        this.i++
        const sym = this.s[this.i]
        this.i++
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
      out += mapSymbolText(c)
      this.i++
    }
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
      if (!Number.isNaN(byte)) this.emit(decodeCp1252(byte))
      return
    }
    if (!/[A-Za-z]/.test(c)) {
      this.i += 2
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
    switch (name) {
      // ---- character formatting ----
      case 'b': char.bold = param !== 0; return
      case 'i': char.italic = param !== 0; return
      case 'ul': char.underline = param !== 0; return
      case 'ulnone': char.underline = false; return
      case 'uldb': char.underline = true; return
      case 'strike': char.strike = param !== 0; return
      case 'plain':
        Object.assign(char, defaultChar())
        return
      case 'fs':
        if (param !== undefined) char.fontSizePt = param / 2
        return
      case 'f':
        char.fontFamily = param !== undefined ? this.fonts[param] : undefined
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
        if (code >= 0 && code <= 0x10ffff) this.emit(String.fromCodePoint(code))
        this.skipChars = this.unicodeSkipCount
        return
      }

      // ---- paragraph structure ----
      case 'par': this.flushParagraph(); return
      case 'line':
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
      case 'cell': this.endCell(); return
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
      case 'ls': return

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
      case 'paperw': this.section.paperWidthTwips = param ?? this.section.paperWidthTwips; return
      case 'paperh': this.section.paperHeightTwips = param ?? this.section.paperHeightTwips; return
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
      case 'tab': this.emit('\t'); return
      case 'emdash': this.emit('\u2014'); return
      case 'endash': this.emit('\u2013'); return
      case 'lquote': this.emit('\u2018'); return
      case 'rquote': this.emit('\u2019'); return
      case 'ldblquote': this.emit('\u201c'); return
      case 'rdblquote': this.emit('\u201d'); return
      case 'bullet': this.emit('\u2022'); return
      case 'enspace':
      case 'emspace':
      case 'qmspace': this.emit(' '); return
      default:
        return
    }
  }

  // ---- destination tables ---------------------------------------------

  private parseFontTable(): void {
    let index = 0
    let name = ''
    let depth = 1
    const commit = (): void => {
      const trimmed = name.trim()
      if (trimmed.length > 0 || this.fonts[index] === undefined) this.fonts[index] = trimmed
      name = ''
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