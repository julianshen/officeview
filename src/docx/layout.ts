/**
 * DOCX layout + paint. Coordinates are CSS px (96dpi): 1pt = 96/72 px.
 * Text measurement uses a real Canvas2D measureText (via a provided measure
 * function) so widths match what we paint.
 */
import type { DocxDocument, DocxImage, DocxParagraph, DocxSection, DocxTable, DocxTextRun } from './types'
import { emuToPx } from '../core/geometry'
import { twipsToPx } from '../core/geometry'
import { resolveColor } from '../core/color'

export interface RunStyle {
  fontFamily: string
  fontSizePt: number
  bold: boolean
  italic: boolean
}

export type MeasureFn = (text: string, style: RunStyle) => number

export interface Segment {
  text: string
  run: DocxTextRun
  style: RunStyle
  widthPx: number
}

export interface LineBox {
  yPx: number
  xPx: number
  widthPx: number
  segs: Segment[]
  align: DocxParagraph['align']
  /** true for the last line of its paragraph (no justification stretch). */
  isParagraphEnd: boolean
  heightPx: number
  /** Width of the content area this line flows in (page column or table cell). */
  contentWidthPx: number
}

export interface PageLayout {
  widthPx: number
  heightPx: number
  lines: LineBox[]
  /** Tables painted beneath the text lines. */
  tables: TableBox[]
  /** Inline images in flow order. */
  images: ImageBox[]
}

/** Image placed on a page. Coordinates are page-relative px. */
export interface ImageBox {
  xPx: number
  yPx: number
  widthPx: number
  heightPx: number
  /** Index into the document's decoded image list (set by the caller). */
  imageIndex: number
}

/** Rect geometry for one table cell (page-relative, after layout). */
export interface TableCellBox {
  xPx: number
  yPx: number
  widthPx: number
  heightPx: number
  fill?: string
  borders?: { left?: string; right?: string; top?: string; bottom?: string }
}

export interface TableRowBox {
  yPx: number
  heightPx: number
  cells: TableCellBox[]
}

export interface TableBox {
  xPx: number
  yPx: number
  widthPx: number
  rows: TableRowBox[]
}

const LINE_HEIGHT_FACTOR = 1.35
const TAB_STOP_PX = twipsToPx(720) // 0.5 inch

/** Build a measure function on any 2D context, memoized by style key. */
export function createMeasurer(ctx: CanvasRenderingContext2D): MeasureFn {
  const cache = new Map<string, number>()
  return (text, style) => {
    const key = `${text}\u0000${style.fontFamily}|${style.fontSizePt}|${style.bold ? 1 : 0}|${style.italic ? 1 : 0}`
    const hit = cache.get(key)
    if (hit !== undefined) return hit
    ctx.font = fontCss(style)
    const w = ctx.measureText(text).width
    cache.set(key, w)
    return w
  }
}

export function fontCss(s: RunStyle): string {
  const parts: string[] = []
  if (s.italic) parts.push('italic')
  if (s.bold) parts.push('bold')
  parts.push(`${s.fontSizePt}pt`)
  parts.push(`"${s.fontFamily}"`)
  return parts.join(' ')
}

function runStyleOf(run: DocxTextRun, defaults: { fontFamily: string; fontSizePt: number }): RunStyle {
  return {
    fontFamily: run.fontFamily ?? defaults.fontFamily,
    fontSizePt: run.fontSizePt ?? defaults.fontSizePt,
    bold: !!run.bold,
    italic: !!run.italic,
  }
}

interface Token {
  kind: 'text' | 'break' | 'tab'
  text: string
  run: DocxTextRun
  style: RunStyle
}

function tokenize(para: DocxParagraph, defaults: { fontFamily: string; fontSizePt: number }): Token[] {
  const tokens: Token[] = []
  for (const run of para.runs) {
    const style = runStyleOf(run, defaults)
    if (run.breakBefore) tokens.push({ kind: 'break', text: '', run, style })
    if (run.text.length === 0) continue
    const chunks = run.text.split(/(\t)/)
    for (const chunk of chunks) {
      if (chunk === '\t') tokens.push({ kind: 'tab', text: '', run, style })
      else if (chunk.length > 0) tokens.push({ kind: 'text', text: chunk, run, style })
    }
  }
  return tokens
}

/**
 * Lay one paragraph into lines. Returns laid lines and the y after the
 * paragraph (including after-spacing).
 */
function layoutParagraph(
  para: DocxParagraph,
  measure: MeasureFn,
  opts: {
    contentX: number
    contentWidth: number
    startY: number
    defaults: { fontFamily: string; fontSizePt: number }
  },
): { lines: LineBox[]; endY: number } {
  const lines: LineBox[] = []
  const { contentX, contentWidth, startY, defaults } = opts
  const indentLeft = twipsToPx(para.indentLeftTwips ?? 0)
  const indentRight = twipsToPx(para.indentRightTwips ?? 0)
  const firstLineIndent = twipsToPx(para.indentFirstLineTwips ?? 0)
  const usable = contentWidth - indentLeft - indentRight
  if (usable <= 0) return { lines, endY: startY }

  let y = startY
  let segs: Segment[] = []
  let width = 0
  let lineHeight = 0

  const paragraphLineHeight = (): number => {
    if (para.lineSpacing?.rule === 'exact') return twipsToPx(para.lineSpacing.value)
    if (para.lineSpacing?.rule === 'atLeast') return twipsToPx(para.lineSpacing.value)
    const mult = para.lineSpacing?.rule === 'auto' ? para.lineSpacing.value / 240 : 1
    const base = (maxRunFontSize(para, defaults) || defaults.fontSizePt) * LINE_HEIGHT_FACTOR
    return base * mult
  }

  const flush = (isParagraphEnd: boolean) => {
    const h = para.lineSpacing?.rule === 'exact' ? twipsToPx(para.lineSpacing.value) : Math.max(lineHeight, paragraphLineHeight() * 0.9)
    const lineIndent = indentLeft + (firstLine ? firstLineIndent : 0)
    if (segs.length === 0) {
      lines.push({ yPx: y, xPx: contentX + lineIndent, widthPx: 0, segs: [], align: para.align, isParagraphEnd, heightPx: h, contentWidthPx: usable })
    } else {
      lines.push({ yPx: y, xPx: contentX + lineIndent, widthPx: width, segs, align: para.align, isParagraphEnd, heightPx: h, contentWidthPx: usable })
    }
    y += h
    if (para.lineSpacing?.rule === 'atLeast') {
      y = Math.max(y, startY + twipsToPx(para.lineSpacing.value))
    }
    segs = []
    width = 0
    lineHeight = 0
  }

  const pushSeg = (text: string, run: DocxTextRun, style: RunStyle, w: number) => {
    segs.push({ text, run, style, widthPx: w })
    width += w
    lineHeight = Math.max(lineHeight, style.fontSizePt * LINE_HEIGHT_FACTOR * (96 / 72))
  }

  const pushWord = (word: string, run: DocxTextRun, style: RunStyle): boolean => {
    const w = measure(word, style)
    if (width > 0 && width + w > usable) return false // needs new line
    pushSeg(word, run, style, w)
    return true
  }

  const tokens = tokenize(para, defaults)
  let firstLine = true

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i]
    if (token.kind === 'break') {
      flush(false)
      firstLine = false
      continue
    }
    if (token.kind === 'tab') {
      // advance to next 0.5in tab stop relative to indent
      const tabWidth = TAB_STOP_PX
      const cur = (firstLine ? indentLeft + twipsToPx(para.indentFirstLineTwips ?? 0) : indentLeft) + width
      const next = Math.floor(cur / tabWidth) * tabWidth + tabWidth
      const w = Math.max(0, next - cur)
      segs.push({ text: ' ', run: token.run, style: token.style, widthPx: w })
      width += w
      continue
    }
    // text: split into words on spaces
    let word = ''
    for (let c = 0; c < token.text.length; c++) {
      const ch = token.text[c]
      if (ch === ' ') {
        if (word) {
          if (!pushWord(word, token.run, token.style)) {
            flush(false)
            firstLine = false
            pushWord(word, token.run, token.style)
          }
          word = ''
        }
        // spaces collapse at line start; measured inside line
        if (width > 0) {
          const w = measure(' ', token.style)
          if (width + w <= usable) {
            segs.push({ text: ' ', run: token.run, style: token.style, widthPx: w })
            width += w
          }
          // else drop trailing space at wrap point
        }
      } else {
        word += ch
      }
    }
    if (word) {
      if (!pushWord(word, token.run, token.style)) {
        flush(false)
        firstLine = false
        pushWord(word, token.run, token.style)
      }
    }
  }
  flush(true)

  return { lines, endY: y }
}

function maxRunFontSize(para: DocxParagraph, defaults: { fontSizePt: number }): number {
  let max = 0
  for (const r of para.runs) max = Math.max(max, r.fontSizePt ?? 0)
  return max > 0 ? max : defaults.fontSizePt
}

function sectionMargins(section: DocxSection) {
  return {
    top: twipsToPx(section.margins.topTwips),
    bottom: twipsToPx(section.margins.bottomTwips),
    left: twipsToPx(section.margins.leftTwips),
    right: twipsToPx(section.margins.rightTwips),
  }
}

/** Lay out a whole document into page boxes. */
/** Unique images in document order — also the index order used by ImageBox. */
export function collectDocImages(document: DocxDocument): DocxImage[] {
  const out: DocxImage[] = []
  const seen = new Set<DocxImage>()
  for (const section of document.sections) {
    for (const block of section.blocks ?? []) {
      if (block.kind === 'p') {
        for (const image of block.paragraph.images ?? []) {
          if (!seen.has(image)) {
            seen.add(image)
            out.push(image)
          }
        }
      }
    }
  }
  return out
}

export function layoutDocx(document: DocxDocument, measure: MeasureFn): PageLayout[] {
  const pages: PageLayout[] = []
  const defaults = { fontFamily: document.defaultFontFamily, fontSizePt: document.defaultFontSizePt }
  const docImages = collectDocImages(document)
  const imageIndex = new Map(docImages.map((img, i) => [img, i]))
  for (const section of document.sections) {
    const widthPx = twipsToPx(section.pageSize.widthTwips)
    const heightPx = twipsToPx(section.pageSize.heightTwips)
    const m = sectionMargins(section)
    const contentX = m.left
    const contentWidth = widthPx - m.left - m.right
    const contentBottom = heightPx - m.bottom
    let page: PageLayout = { widthPx, heightPx, lines: [], tables: [], images: [] }
    let y = m.top

    const commitPage = () => {
      if (page.lines.length > 0 || page.tables.length > 0 || page.images.length > 0) pages.push(page)
      page = { widthPx, heightPx, lines: [], tables: [], images: [] }
      y = m.top
    }

    const blocks = section.blocks ?? section.paragraphs.map((paragraph) => ({ kind: 'p' as const, paragraph }))
    for (const block of blocks) {
      if (block.kind === 'p') {
        const para = block.paragraph
        y += twipsToPx(para.spacingBeforeTwips ?? 0)
        const { lines, endY } = layoutParagraph(para, measure, {
          contentX,
          contentWidth,
          startY: y,
          defaults,
        })
        // Lines carry absolute y measured from the section start. Any line
        // past contentBottom rolls onto a fresh page, rebased to the top
        // margin; the same shift applies to the paragraph's remaining lines
        // and to the running flow position.
        let shift = 0
        for (const line of lines) {
          if (line.yPx + shift > contentBottom && page.lines.length > 0) {
            commitPage()
            shift = m.top - line.yPx
          }
          page.lines.push({ ...line, yPx: line.yPx + shift })
        }
        y = endY + shift + twipsToPx(para.spacingAfterTwips ?? 0)
        // inline images after the paragraph's text
        for (const image of para.images ?? []) {
          const w = emuToPx(image.widthEmu)
          const h = emuToPx(image.heightEmu)
          if (w <= 0 || h <= 0) continue
          if (y + h > contentBottom && (page.lines.length > 0 || page.images.length > 0)) commitPage()
          page.images.push({ xPx: contentX, yPx: y, widthPx: w, heightPx: h, imageIndex: imageIndex.get(image) ?? -1 })
          y += h
        }
        continue
      }
      // Table block: whole-table page placement (a table taller than a page
      // overflows the bottom — acceptable baseline).
      const table = block.table
      const laid = layoutTable(table, measure, defaults)
      let tableY = y
      if (tableY + laid.heightPx > contentBottom && (page.lines.length > 0 || page.tables.length > 0)) {
        commitPage()
        tableY = y
      }
      const tableX = m.left
      for (const line of laid.lines) {
        page.lines.push({ ...line, xPx: line.xPx + tableX, yPx: line.yPx + tableY })
      }
      page.tables.push({ ...laid.box, xPx: tableX, yPx: tableY })
      y = tableY + laid.heightPx
    }
    if (page.lines.length > 0 || page.tables.length > 0 || page.images.length > 0 || pages.length === 0) pages.push(page)
  }
  return pages
}

const HIGHLIGHT_CSS: Record<string, string> = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff',
  blue: '#0000ff', red: '#ff0000', darkBlue: '#00008b', darkCyan: '#008b8b',
  darkGreen: '#006400', darkMagenta: '#8b008b', darkRed: '#8b0000',
  darkYellow: '#808000', darkGray: '#a9a9a9', lightGray: '#d3d3d3',
  black: '#000000', white: '#ffffff',
}

/** Paint laid pages onto a 2D context already scaled so 1 unit = 1 px. */
export function renderPages(
  pages: PageLayout[],
  ctx: CanvasRenderingContext2D,
  images?: Array<CanvasImageSource | undefined>,
  _defaults?: { fontFamily: string },
): void {
  ctx.save()
  ctx.textBaseline = 'alphabetic'
  ctx.fillStyle = '#000000'
  let lastFont = ''
  for (const page of pages) {
    paintTables(page.tables, ctx)
    paintImages(page.images, ctx, images)
    for (const line of page.lines) {
      let extraSpacePerGap = 0
      if (line.align === 'justify' && !line.isParagraphEnd && line.segs.length > 1) {
        const gaps = countGaps(line.segs)
        if (gaps > 0) extraSpacePerGap = (line.contentWidthPx - line.widthPx) / gaps
      }
      let offset = 0
      if (line.align === 'center') offset = (line.contentWidthPx - line.widthPx) / 2
      else if (line.align === 'right') offset = line.contentWidthPx - line.widthPx
      let x = line.xPx + offset
      let maxAscent = 0
      for (const seg of line.segs) maxAscent = Math.max(maxAscent, seg.style.fontSizePt * LINE_HEIGHT_FACTOR * 0.8)
      const baseline = line.yPx + maxAscent
      for (const seg of line.segs) {
        const fontCssStr = fontCss(seg.style)
        if (fontCssStr !== lastFont) {
          ctx.font = fontCssStr
          lastFont = fontCssStr
        }
        if (seg.run.highlight) {
          const hl = HIGHLIGHT_CSS[seg.run.highlight] ?? seg.run.highlight
          ctx.fillStyle = hl
          ctx.fillRect(x, line.yPx, seg.widthPx, line.heightPx)
        }
        ctx.fillStyle = seg.run.color ? resolveColor(seg.run.color) : '#000000'
        ctx.fillText(seg.text, x, baseline)
        if (seg.run.underline || seg.run.strike) {
          ctx.strokeStyle = ctx.fillStyle
          ctx.lineWidth = Math.max(1, seg.style.fontSizePt * 0.06)
          ctx.beginPath()
          const yy = seg.run.underline ? baseline + seg.style.fontSizePt * 0.15 : baseline - seg.style.fontSizePt * 0.3
          ctx.moveTo(x, yy)
          ctx.lineTo(x + seg.widthPx, yy)
          ctx.stroke()
        }
        x += seg.widthPx + (seg.text === ' ' ? extraSpacePerGap : 0)
      }
    }
  }
  ctx.restore()
}

function countGaps(segs: Segment[]): number {
  let n = 0
  for (const s of segs) if (s.text === ' ') n++
  return n
}

const BORDER_WIDTH: Record<string, number> = {
  thin: 1,
  thick: 2.5,
  double: 1,
  single: 1,
  dashed: 1,
  dotted: 1,
}

/** Draw inline images (after tables, beneath text). */
function paintImages(
  boxes: ImageBox[],
  ctx: CanvasRenderingContext2D,
  images: Array<CanvasImageSource | undefined> | undefined,
): void {
  if (!images) return
  for (const box of boxes) {
    const img = images[box.imageIndex]
    if (!img) continue
    ctx.drawImage(img as CanvasImageSource, box.xPx, box.yPx, box.widthPx, box.heightPx)
  }
}

/** Paint table cell fills and borders (beneath text). Cell coords are
 * table-relative; box carries the page position. */
function paintTables(tables: TableBox[], ctx: CanvasRenderingContext2D): void {
  for (const table of tables) {
    for (const row of table.rows) {
      for (const cell of row.cells) {
        if (cell.fill) {
          ctx.fillStyle = resolveColor(cell.fill)
          ctx.fillRect(table.xPx + cell.xPx, table.yPx + cell.yPx, cell.widthPx, cell.heightPx)
        }
      }
    }
    // borders after fills so they sit on top of shading
    for (const row of table.rows) {
      for (const cell of row.cells) {
        const b = cell.borders
        if (!b) continue
        ctx.strokeStyle = '#000000'
        const cx = table.xPx + cell.xPx
        const cy = table.yPx + cell.yPx
        const draw = (w: number, x1: number, y1: number, x2: number, y2: number) => {
          ctx.lineWidth = w
          ctx.beginPath()
          ctx.moveTo(x1 + 0.5, y1 + 0.5)
          ctx.lineTo(x2 + 0.5, y2 + 0.5)
          ctx.stroke()
        }
        const lw = (s: string | undefined) => Math.max(1, BORDER_WIDTH[s ?? 'thin'] ?? 1)
        if (b.left) draw(lw(b.left), cx, cy, cx, cy + cell.heightPx)
        if (b.right) draw(lw(b.right), cx + cell.widthPx, cy, cx + cell.widthPx, cy + cell.heightPx)
        if (b.top) draw(lw(b.top), cx, cy, cx + cell.widthPx, cy)
        if (b.bottom) draw(lw(b.bottom), cx, cy + cell.heightPx, cx + cell.widthPx, cy + cell.heightPx)
      }
    }
  }
}

// ---------- table layout ----------

interface CellBorderCss { style?: string; color?: string }

function borderCss(b: CellBorderCss | undefined): string | undefined {
  if (!b || !b.style) return undefined
  if (b.style === 'single' || b.style === 'thin') return 'thin'
  return b.style
}

/**
 * Lay out a table into cell rects + text lines relative to (0,0) at the
 * table origin; caller rebases y onto the page.
 */
function layoutTable(
  table: DocxTable,
  measure: MeasureFn,
  defaults: { fontFamily: string; fontSizePt: number },
): { lines: LineBox[]; box: TableBox; heightPx: number } {
  const lines: LineBox[] = []
  const gridWidthPx = table.gridColsTwips.reduce((a, b) => a + b, 0)
  const colOffsets = prefixSum(table.gridColsTwips)
  const colWidths = table.gridColsTwips.map((t) => t) // twips; convert at use
  const box: TableBox = { xPx: 0, yPx: 0, widthPx: twipsToPx(gridWidthPx), rows: [] }
  const margins = {
    top: twipsToPx(table.cellMargins.topTwips),
    bottom: twipsToPx(table.cellMargins.bottomTwips),
    left: twipsToPx(table.cellMargins.leftTwips),
    right: twipsToPx(table.cellMargins.rightTwips),
  }

  let yRel = 0
  for (const row of table.rows) {
    const rowCells: TableCellBox[] = []
    let rowContentH = 0
    let col = 0
    for (const cell of row.cells) {
      if (cell.vMerge === 'continue') {
        col += cell.gridSpan
        continue
      }
      const span = Math.max(1, cell.gridSpan)
      let cellW = 0
      for (let i = col; i < col + span && i < colWidths.length; i++) cellW += colWidths[i]
      col += span
      const cellXPx = twipsToPx(colOffsets[Math.max(0, col - span)] ?? 0)
      const innerW = twipsToPx(cellW) - margins.left - margins.right
      if (innerW <= 0) continue
      let cellLines: LineBox[] = []
      let cy = yRel + margins.top
      for (const para of cell.paragraphs) {
        const laid = layoutParagraph(para, measure, {
          contentX: cellXPx + margins.left,
          contentWidth: innerW,
          startY: cy,
          defaults,
        })
        cellLines = cellLines.concat(laid.lines)
        cy = laid.endY + twipsToPx(para.spacingAfterTwips ?? 0)
      }
      const contentH = cellLines.length > 0
        ? cellLines[cellLines.length - 1].yPx - yRel - margins.top + cellLines[cellLines.length - 1].heightPx + margins.bottom
        : margins.top + margins.bottom + defaults.fontSizePt * LINE_HEIGHT_FACTOR * (96 / 72) * 0.5
      rowContentH = Math.max(rowContentH, contentH)
      for (const line of cellLines) lines.push(line)
      // borders: cell overrides, falling back to the table's outside/
      // inside border definitions
      const cb = cell.borders
      const tb = table.borders
      const pick = (side: 'left' | 'right' | 'top' | 'bottom'): string | undefined => {
        if (cb && side in cb) return borderCss(cb[side])
        if (tb) return borderCss(tb[side] ?? (side === 'top' || side === 'bottom' ? tb.insideH : tb.insideV))
        return undefined
      }
      rowCells.push({
        xPx: cellXPx,
        yPx: yRel,
        widthPx: twipsToPx(cellW),
        heightPx: 0, // filled after row height known
        fill: cell.fill ?? table.fill,
        borders: {
          left: pick('left'),
          right: pick('right'),
          top: pick('top'),
          bottom: pick('bottom'),
        },
      })
    }
    // explicit row height (atLeast semantics)
    let rowH = rowContentH
    if (row.heightTwips !== undefined) {
      const hPx = twipsToPx(row.heightTwips)
      if (row.heightRule === 'exact') rowH = hPx
      else rowH = Math.max(rowH, hPx)
    }
    rowH = Math.max(rowH, defaults.fontSizePt * LINE_HEIGHT_FACTOR * (96 / 72))
    for (const cellBox of rowCells) cellBox.heightPx = rowH
    box.rows.push({ yPx: yRel, heightPx: rowH, cells: rowCells })
    yRel += rowH
  }
  return { lines, box, heightPx: yRel }
}

function prefixSum(widths: number[]): number[] {
  const out = new Array<number>(widths.length + 1)
  out[0] = 0
  for (let i = 0; i < widths.length; i++) out[i + 1] = out[i] + widths[i]
  return out
}
