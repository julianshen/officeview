/**
 * DOCX layout + paint. Coordinates are CSS px (96dpi): 1pt = 96/72 px.
 * Text measurement uses a real Canvas2D measureText (via a provided measure
 * function) so widths match what we paint.
 */
import type { DocxDocument, DocxParagraph, DocxSection, DocxTextRun } from './types'
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
}

export interface PageLayout {
  widthPx: number
  heightPx: number
  lines: LineBox[]
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
      lines.push({ yPx: y, xPx: contentX + lineIndent, widthPx: 0, segs: [], align: para.align, isParagraphEnd, heightPx: h })
    } else {
      lines.push({ yPx: y, xPx: contentX + lineIndent, widthPx: width, segs, align: para.align, isParagraphEnd, heightPx: h })
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
export function layoutDocx(document: DocxDocument, measure: MeasureFn): PageLayout[] {
  const pages: PageLayout[] = []
  const defaults = { fontFamily: document.defaultFontFamily, fontSizePt: document.defaultFontSizePt }
  for (const section of document.sections) {
    const widthPx = twipsToPx(section.pageSize.widthTwips)
    const heightPx = twipsToPx(section.pageSize.heightTwips)
    const m = sectionMargins(section)
    const contentX = m.left
    const contentWidth = widthPx - m.left - m.right
    const contentBottom = heightPx - m.bottom
    let page: PageLayout = { widthPx, heightPx, lines: [] }
    let y = m.top

    const commitPage = () => {
      if (page.lines.length > 0) pages.push(page)
      page = { widthPx, heightPx, lines: [] }
      y = m.top
    }

    for (const para of section.paragraphs) {
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
    }
    if (page.lines.length > 0 || pages.length === 0) pages.push(page)
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
export function renderPages(pages: PageLayout[], ctx: CanvasRenderingContext2D, _defaults?: { fontFamily: string }): void {
  ctx.save()
  ctx.textBaseline = 'alphabetic'
  ctx.fillStyle = '#000000'
  let lastFont = ''
  for (const page of pages) {
    for (const line of page.lines) {
      let extraSpacePerGap = 0
      if (line.align === 'justify' && !line.isParagraphEnd && line.segs.length > 1) {
        const gaps = countGaps(line.segs)
        if (gaps > 0) extraSpacePerGap = (page.widthPx - marginsApprox(line) - line.widthPx) / gaps
      }
      let offset = 0
      if (line.align === 'center') offset = (usableWidth(page, line) - line.widthPx) / 2
      else if (line.align === 'right') offset = usableWidth(page, line) - line.widthPx
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

function marginsApprox(line: LineBox): number {
  // x offset within the page (left margin)
  return line.xPx
}

function usableWidth(page: PageLayout, line: LineBox): number {
  return page.widthPx - line.xPx - twipsToPx(1440)
}
