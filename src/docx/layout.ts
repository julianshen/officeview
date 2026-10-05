/**
 * DOCX layout + paint. Coordinates are CSS px (96dpi): 1pt = 96/72 px.
 * Text measurement uses a real Canvas2D measureText (via a provided measure
 * function) so widths match what we paint.
 */
import type { DocxBlock, DocxDocument, DocxImage, DocxParagraph, DocxSection, DocxTable, DocxTableCell, DocxTableRow, DocxTextRun, TableCellBorder } from './types'
import { paintDrawing } from './drawing'
import type { ContentPaintAssets } from '../drawing/content'
import { fontFamilyCss } from './styles'
import { withFallbackFonts } from '../core/fonts/fallback'
import { emuToPx } from '../core/geometry'
import { twipsToPx } from '../core/geometry'
import { resolveColor } from '../core/color'
import { paintWatermark, type ResolvedWatermark, type WatermarkOptions } from '../core/watermark'
import { layoutTextBody } from '../drawing/text-layout'
import { clusterOrientation } from '../drawing/vertical-orientation'
import { graphemes, RECORD_TEXT, type LogicalTextRange, type TextRecordingContext } from '../core/text-recording'
import type { DrawingTextParagraph, LocalAffine, TextDirection } from '../drawing/text'

export interface RunStyle {
  fontFamily: string
  fontSizePt: number
  bold: boolean
  italic: boolean
}

export type MeasureFn = ((text: string, style: RunStyle) => number) & {
  metrics?: (style: RunStyle, text?: string) => { ascent: number; descent: number; normalHeight?: number }
}

export interface Segment {
  text: string
  run: DocxTextRun
  style: RunStyle
  widthPx: number
  /**
   * Pen advance (px) accumulated before this segment, including any inline
   * image widths. Relative to `line.xPx + alignOffset`; lets the painter place
   * text after inline images correctly.
   */
  penOffset?: number
  /** Absolute page-space affine for direction-rotated glyphs. Absent means ordinary horizontal flow paint at the accumulated pen. */
  transform?: LocalAffine
  logical?: LogicalTextRange
}

function shiftedImageClip(clip: ImageBox['clip'], x: number, y: number): ImageBox['clip'] {
  return clip ? { ...clip, xPx: clip.xPx + x, yPx: clip.yPx + y } : undefined
}

/** Shift a line vertically, keeping direction-rotated segment placements glued to it. */
function shiftLineY(line: LineBox, dy: number): LineBox {
  if (dy === 0) return line
  return {
    ...line,
    yPx: line.yPx + dy,
    baselinePx: line.baselinePx === undefined ? undefined : line.baselinePx + dy,
    inlineImages: line.inlineImages?.map((box) => ({ ...box, yPx: box.yPx + dy, clip: shiftedImageClip(box.clip, 0, dy) })),
    segs: line.segs.map((seg) =>
      seg.transform ? { ...seg, transform: { ...seg.transform, f: seg.transform.f + dy } } : seg
    ),
  }
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
  /** List marker drawn in the left gutter (first line only). */
  marker?: { text: string; widthPx: number }
  inlineImages?: Array<{ image: DocxImage; xPx: number; yPx: number; widthPx: number; heightPx: number; clip?: ImageBox['clip'] }>
  baselinePx?: number
  logical?: LogicalTextRange
  defaultStyle?: RunStyle
}

export interface PageLayout {
  widthPx: number
  heightPx: number
  lines: LineBox[]
  /** Tables painted beneath the text lines. */
  tables: TableBox[]
  /** Inline images in flow order. */
  images: ImageBox[]
  header?: HFBlock
  footer?: HFBlock
  /** Repeating-content overflow or bounded field convergence limitations. */
  diagnostics?: Array<'repeated-content-overflow' | 'field-layout-nonconvergence'>
}

/** Page header or footer content, positioned in page-relative px. */
export interface HFBlock {
  paragraphs: DocxParagraph[]
  blocks?: DocxBlock[]
  footer?: boolean
  /** Physical section margins, independent of story/paragraph/cell origins. */
  margins?: { top: number; bottom: number; left: number; right: number }
  clipTopPx?: number
  clipBottomPx?: number
  imageIndex?: Map<DocxImage, number>
  yPx: number
  xPx: number
  widthPx: number
  defaults: { fontFamily: string; fontSizePt: number }
}

/** Image placed on a page. Coordinates are page-relative px. */
export interface ImageBox {
  xPx: number
  yPx: number
  widthPx: number
  heightPx: number
  /** Index into the document's decoded image list (set by the caller). */
  imageIndex: number
  drawing?: DocxImage['drawing']
  /** Physical cell viewport for directional inline images that exceed its bounds. */
  clip?: { xPx: number; yPx: number; widthPx: number; heightPx: number }
  /** Set for floating (wp:anchor) images: drawn behind text and out of flow. */
  repeated?: 'header' | 'footer'
  floating?: { behindDoc: boolean; relativeHeight: number; wrap: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom' }
}

/** Rect geometry for one table cell (page-relative, after layout). */
export interface TableCellBox {
  xPx: number
  yPx: number
  widthPx: number
  heightPx: number
  fill?: string
  borders?: { left?: string; right?: string; top?: string; bottom?: string }
  borderSpecs?: Partial<Record<'left' | 'right' | 'top' | 'bottom', TableCellBorder>>
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
export function createMeasurer(ctx: CanvasRenderingContext2D, familyCss: (family: string) => string = fontFamilyCss): MeasureFn {
  const cache = new Map<string, number>()
  const measure: MeasureFn = (text, style) => {
    const key = `${text}\u0000${style.fontFamily}|${style.fontSizePt}|${style.bold ? 1 : 0}|${style.italic ? 1 : 0}`
    const hit = cache.get(key)
    if (hit !== undefined) return hit
    ctx.font = fontCss(style, familyCss)
    const w = ctx.measureText(text).width
    cache.set(key, w)
    return w
  }
  measure.metrics = (style, text = 'Mg') => {
    ctx.font = fontCss(style, familyCss)
    const m = ctx.measureText(text)
    return { ascent: m.actualBoundingBoxAscent, descent: m.actualBoundingBoxDescent,
      normalHeight: (m.fontBoundingBoxAscent ?? (m as TextMetrics & { emHeightAscent?: number }).emHeightAscent ?? m.actualBoundingBoxAscent) + (m.fontBoundingBoxDescent ?? (m as TextMetrics & { emHeightDescent?: number }).emHeightDescent ?? m.actualBoundingBoxDescent) }
  }
  return measure
}

export function fontCss(s: RunStyle, familyCss: (family: string) => string = fontFamilyCss): string {
  const parts: string[] = []
  if (s.italic) parts.push('italic')
  if (s.bold) parts.push('bold')
  parts.push(`${s.fontSizePt}pt`)
  parts.push(familyCss(s.fontFamily))
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
  kind: 'text' | 'break' | 'tab' | 'image'
  text: string
  run: DocxTextRun
  style: RunStyle
  image?: DocxImage
}

function tokenize(para: DocxParagraph, defaults: { fontFamily: string; fontSizePt: number }): Token[] {
  const tokens: Token[] = []
  const inline = para.inline ?? [
    ...para.runs.map((run) => ({ kind: 'text' as const, run })),
    ...para.images.map((image) => ({ kind: 'image' as const, image }))
  ]
  for (const item of inline) {
    if (item.kind === 'image') {
      if (!item.image.floating)
        tokens.push({
          kind: 'image',
          image: item.image,
          text: '',
          run: { text: '' },
          style: runStyleOf({ text: '' }, defaults)
        })
      continue
    }
    const run = item.run
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
    /**
     * Free horizontal bands at a given y range, after subtracting floating
     * objects. Undefined/empty means "no float here, use the full column".
     */
    floatBands?: (yTop: number, yBottom: number) => Array<{ x: number; width: number }> | undefined
  }
): { lines: LineBox[]; endY: number } {
  const lines: LineBox[] = []
  const { contentX, contentWidth, startY, defaults } = opts
  const indentLeft = twipsToPx(para.indentLeftTwips ?? 0)
  const indentRight = twipsToPx(para.indentRightTwips ?? 0)
  const firstLineIndent = twipsToPx(para.indentFirstLineTwips ?? 0)
  // A list marker sits in the gutter: reserve its width and shift the text so
  // wrapped lines align with the first line (Word's hanging indent).
  const markerStyle: RunStyle = {
    fontFamily: para.runs[0]?.fontFamily ?? defaults.fontFamily,
    fontSizePt: para.runs[0]?.fontSizePt ?? defaults.fontSizePt,
    bold: !!para.runs[0]?.bold,
    italic: !!para.runs[0]?.italic
  }
  const markerGutter = para.listMarker ? measure(`${para.listMarker} `, markerStyle) : 0
  const fullUsable = contentWidth - indentLeft - indentRight - markerGutter
  if (fullUsable <= 0) return { lines, endY: startY }

  // Per-line band: a floating image narrows the column on the lines it spans.
  // Recomputed at each line start (Word snaps wrapping to line granularity).
  const lineHeightEstimate = (): number => {
    if (para.lineSpacing?.rule === 'exact' || para.lineSpacing?.rule === 'atLeast')
      return twipsToPx(para.lineSpacing.value)
    const mult = para.lineSpacing?.rule === 'auto' ? para.lineSpacing.value / 240 : 1
    return (maxRunFontSize(para, defaults) || defaults.fontSizePt) * LINE_HEIGHT_FACTOR * mult
  }
  let lineX = contentX
  let lineUsable = fullUsable
  const refreshBand = (): void => {
    if (segs.length > 0 || inlineImages.length > 0) return // this line is already positioned
    const bands = opts.floatBands?.(y, y + lineHeightEstimate())
    let best: { x: number; width: number } | undefined
    if (bands) {
      for (const b of bands) if (!best || b.width > best.width) best = b
    }
    // no bands at all, or nothing left to write into: fall back to the column
    // rather than dropping the paragraph's text
    if (!best || best.width <= 0) {
      lineX = contentX
      lineUsable = fullUsable
      return
    }
    lineX = best.x
    lineUsable = Math.min(fullUsable, best.width - indentLeft - indentRight - markerGutter)
    if (lineUsable <= 0) {
      lineX = contentX
      lineUsable = fullUsable
    }
  }

  let y = startY
  let segs: Segment[] = []
  let width = 0
  let lineHeight = 0
  let inlineImages: Array<{
    image: DocxImage
    offset: number
    width: number
    height: number
    top: number
    bottom: number
  }> = []

  const paragraphLineHeight = (): number => {
    if (para.lineSpacing?.rule === 'exact') return twipsToPx(para.lineSpacing.value)
    if (para.lineSpacing?.rule === 'atLeast') return twipsToPx(para.lineSpacing.value)
    const mult = para.lineSpacing?.rule === 'auto' ? para.lineSpacing.value / 240 : 1
    const base = (maxRunFontSize(para, defaults) || defaults.fontSizePt) * (96 / 72) * LINE_HEIGHT_FACTOR
    return base * mult
  }

  const flush = (isParagraphEnd: boolean) => {
    const h =
      para.lineSpacing?.rule === 'exact'
        ? twipsToPx(para.lineSpacing.value)
        : Math.max(lineHeight, inlineImages.length && !segs.length ? 0 : paragraphLineHeight() * 0.9)
    const lineIndent = indentLeft + (firstLine ? firstLineIndent : 0) + markerGutter
    const marker = firstLine && para.listMarker ? { text: para.listMarker, widthPx: markerGutter } : undefined
    const alignOffset =
      para.align === 'center' ? (lineUsable - width) / 2 : para.align === 'right' ? lineUsable - width : 0
    // Non-final justified lines expand EVERY gap. Images and text must resolve to
    // the SAME expanded pen, so neither can overlap the other.
    const gaps = segs.reduce((n, s) => n + (s.text === ' ' ? 1 : 0), 0)
    const extra = para.align === 'justify' && !isParagraphEnd && gaps > 0 ? (lineUsable - width) / gaps : 0
    const gapsBefore = (offset: number): number => {
      let n = 0
      for (const s of segs) if (s.text === ' ' && (s.penOffset ?? 0) < offset) n++
      return n
    }
    const boxes = inlineImages.map((i) => ({
      image: i.image,
      xPx: lineX + lineIndent + alignOffset + i.offset + extra * gapsBefore(i.offset),
      yPx: y + (h - (i.height + i.bottom)),
      widthPx: i.width,
      heightPx: i.height
    }))
    if (extra !== 0) {
      let gapsSoFar = 0
      for (const seg of segs) {
        if (seg.penOffset !== undefined) seg.penOffset += extra * gapsSoFar
        if (seg.text === ' ') gapsSoFar++
      }
    }
    lines.push({
      yPx: y,
      xPx: lineX + lineIndent,
      widthPx: extra > 0 ? lineUsable : width,
      segs,
      align: para.align,
      isParagraphEnd,
      heightPx: h,
      contentWidthPx: lineUsable,
      marker,
      ...(boxes.length ? { inlineImages: boxes, baselinePx: y + h } : {})
    })
    y += h
    if (para.lineSpacing?.rule === 'atLeast') {
      y = Math.max(y, startY + twipsToPx(para.lineSpacing.value))
    }
    segs = []
    width = 0
    lineHeight = 0
    inlineImages = []
    lineX = contentX
    lineUsable = fullUsable
  }

  const pushSeg = (text: string, run: DocxTextRun, style: RunStyle, w: number) => {
    segs.push({ text, run, style, widthPx: w, penOffset: width })
    width += w
    lineHeight = Math.max(lineHeight, style.fontSizePt * LINE_HEIGHT_FACTOR * (96 / 72))
  }

  const pushWord = (word: string, run: DocxTextRun, style: RunStyle): boolean => {
    const w = measure(word, style)
    refreshBand()
    if (width > 0 && width + w > lineUsable) return false // needs new line
    pushSeg(word, run, style, w)
    return true
  }

  const tokens = tokenize(para, defaults)
  let firstLine = true

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i]
    if (token.kind === 'image' && token.image) {
      const image = token.image,
        w = emuToPx(image.widthEmu),
        h = emuToPx(image.heightEmu)
      if (w <= 0 || h <= 0) continue
      refreshBand()
      if (width > 0 && width + w > lineUsable + 0.01) {
        flush(false)
        firstLine = false
        refreshBand()
      }
      const top = emuToPx(image.effectExtentEmu?.top ?? 0),
        bottom = emuToPx(image.effectExtentEmu?.bottom ?? 0)
      inlineImages.push({ image, offset: width, width: w, height: h, top, bottom })
      width += w
      lineHeight = Math.max(lineHeight, top + h + bottom)
      continue
    }
    if (token.kind === 'break') {
      flush(false)
      firstLine = false
      continue
    }
    if (token.kind === 'tab') {
      // advance to next 0.5in tab stop relative to indent
      const tabWidth = TAB_STOP_PX
      const cur =
        lineX - contentX + (firstLine ? indentLeft + twipsToPx(para.indentFirstLineTwips ?? 0) : indentLeft) + width
      const next = Math.floor(cur / tabWidth) * tabWidth + tabWidth
      const w = Math.max(0, next - cur)
      segs.push({ text: ' ', run: token.run, style: token.style, widthPx: w, penOffset: width })
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
          if (width + w <= lineUsable) {
            segs.push({ text: ' ', run: token.run, style: token.style, widthPx: w, penOffset: width })
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
  for (const r of para.runs) if (r.text) max = Math.max(max, r.fontSizePt ?? 0)
  return max > 0 ? max : para.paragraphMark?.fontSizePt ?? defaults.fontSizePt
}

function sectionMargins(section: DocxSection) {
  return {
    top: twipsToPx(section.margins.topTwips),
    bottom: twipsToPx(section.margins.bottomTwips),
    left: twipsToPx(section.margins.leftTwips),
    right: twipsToPx(section.margins.rightTwips),
  }
}

/** Ordered models are authoritative, including explicit empty arrays. */
function repeatedBlocks(section: DocxSection, kind: 'header' | 'footer' | 'firstHeader' | 'firstFooter'): DocxBlock[] | undefined {
  return section[`${kind}Blocks`] ?? section[kind]?.map(paragraph => ({ kind: 'p', paragraph }))
}
function blockParagraphs(blocks: DocxBlock[]): DocxParagraph[] {
  return blocks.flatMap(block => block.kind === 'p' ? [block.paragraph] : block.table.rows.flatMap(row => row.cells.flatMap(cell => cell.paragraphs)))
}
/** Unique images in document order; body indices precede all repeated content.
 * Vector drawings that carry textbox paragraphs contribute their nested
 * images depth-first (cycles impossible through the seen set), so a textbox
 * asset keeps its source position directly after its owning placement. */
export function collectDocImages(document: DocxDocument): DocxImage[] {
  const out: DocxImage[] = [], seen = new Set<DocxImage>()
  const push = (image: DocxImage): void => {
    if (seen.has(image)) return
    seen.add(image)
    out.push(image)
    const drawing = image.drawing
    if (drawing?.kind === 'textbox') for (const paragraph of drawing.paragraphs) for (const nested of paragraph.images ?? []) push(nested)
  }
  const collect = (blocks: DocxBlock[]) => {
    for (const paragraph of blockParagraphs(blocks)) for (const image of paragraph.images ?? []) push(image)
  }
  for (const section of document.sections) collect(section.blocks ?? section.paragraphs.map(paragraph => ({ kind: 'p', paragraph })))
  for (const section of document.sections) for (const kind of ['header', 'footer', 'firstHeader', 'firstFooter'] as const) collect(repeatedBlocks(section, kind) ?? [])
  return out
}

/** NUMPAGES may change wrapping. Keep conservative bounds across at most eight passes. */
export function layoutDocx(document: DocxDocument, measure: MeasureFn): PageLayout[] {
  const reservations = new Map<string, { top: number; bottom: number }>()
  let total = 1, previous = '', pages: PageLayout[] = []
  for (let pass = 0; pass < 8; pass++) {
    pages = layoutDocxPass(document, measure, total, reservations)
    const signature = JSON.stringify([pages.length, [...reservations]])
    if (pages.length === total && signature === previous) return pages
    previous = signature
    total = pages.length
  }
  // Final actual field values paint in the fixed conservative viewport.
  for (const page of pages) (page.diagnostics ??= []).push('field-layout-nonconvergence')
  refreshRepeatedImages(pages, measure, pages.length)
  return pages
}

function layoutDocxPass(document: DocxDocument, measure: MeasureFn, totalPages: number, reservations: Map<string, { top: number; bottom: number }>): PageLayout[] {
  const pages: PageLayout[] = []
  const defaults = { fontFamily: document.defaultFontFamily, fontSizePt: document.defaultFontSizePt }
  const docImages = collectDocImages(document)
  const imageIndex = new Map(docImages.map((img, i) => [img, i]))
  let previousEndY = 0
  for (const [sectionIndex, section] of document.sections.entries()) {
    const widthPx = twipsToPx(section.pageSize.widthTwips)
    const heightPx = twipsToPx(section.pageSize.heightTwips)
    const m = sectionMargins(section)
    // OOXML type belongs to this section's ending sectPr and governs its
    // start relative to the previous section, not the following section.
    const previousPage = pages.at(-1)
    const continuousPage = section.type === 'continuous' && previousPage?.widthPx === widthPx && previousPage.heightPx === heightPx ? pages.pop() : undefined
    const pagesBefore = pages.length
    const contentX = m.left
    const contentWidth = widthPx - m.left - m.right
    let contentTop = m.top, contentBottom = heightPx - m.bottom
    let page: PageLayout
    let y = contentTop
    const makePage = (continued?: PageLayout): PageLayout => {
      const localIndex = pages.length - pagesBefore
      const first = localIndex === 0 && section.titlePg
      const makeHF = (kind: 'header' | 'footer'): HFBlock | undefined => {
        const variant = kind === 'header' ? 'firstHeader' : 'firstFooter'
        const blocks = (first ? repeatedBlocks(section, variant) : undefined) ?? repeatedBlocks(section, kind)
        return blocks === undefined ? undefined : { blocks, paragraphs: blocks.flatMap(block => block.kind === 'p' ? [block.paragraph] : []), footer: kind === 'footer', yPx: kind === 'header' ? twipsToPx(section.margins.headerTwips) : heightPx - twipsToPx(section.margins.footerTwips), xPx: m.left, widthPx: contentWidth, defaults, imageIndex, margins: m }
      }
      // A continuous section retains the current page's repeating stories;
      // its own references become applicable when it produces a new page.
      const result: PageLayout = continued ?? { widthPx, heightPx, lines: [], tables: [], images: [], header: makeHF('header'), footer: makeHF('footer') }
      const header = result.header && layoutRepeated(result.header, measure, pages.length + 1, totalPages, result)
      const footer = result.footer && layoutRepeated(result.footer, measure, pages.length + 1, totalPages, result)
      const key = `${sectionIndex}:${localIndex}`
      const old = reservations.get(key)
      const top = Math.max(m.top, header?.occupied ? header.bottom : 0, old?.top ?? 0)
      const bottom = Math.max(m.bottom, footer?.occupied ? heightPx - footer.top : 0, old?.bottom ?? 0)
      reservations.set(key, { top, bottom })
      // A repeated header/footer cannot starve ordinary flow. One ordinary
      // line is the general minimum; the first body item may overflow it.
      const minimumBand = Math.min(heightPx, defaults.fontSizePt * (96 / 72) * LINE_HEIGHT_FACTOR)
      contentTop = Math.min(top, Math.max(0, heightPx - minimumBand))
      contentBottom = Math.max(contentTop + minimumBand, heightPx - bottom)
      contentBottom = Math.min(heightPx, contentBottom)
      if (top + bottom > heightPx - minimumBand || (header?.occupied && (header.top < 0 || header.bottom > heightPx)) || (footer?.occupied && (footer.top < 0 || footer.bottom > heightPx))) result.diagnostics = ['repeated-content-overflow']
      if (result.header) { result.header.clipTopPx = 0; result.header.clipBottomPx = contentTop }
      if (result.footer) { result.footer.clipTopPx = contentBottom; result.footer.clipBottomPx = heightPx }
      y = continued ? Math.max(previousEndY, contentTop) : contentTop
      return result
    }
    page = makePage(continuousPage)
    const commitPage = () => {
      if (page.lines.length > 0 || page.tables.length > 0 || page.images.length > 0) pages.push(page)
      page = makePage()
    }

    const blocks = section.blocks ?? section.paragraphs.map((paragraph) => ({ kind: 'p' as const, paragraph }))
    for (const block of blocks) {
      if (block.kind === 'p') {
        const para = block.paragraph
        y += twipsToPx(para.spacingBeforeTwips ?? 0)
        // Floating images anchor to this paragraph: place them first so the
        // paragraph's own text (and everything after it) can wrap around them.
        // They stay on the page where the anchor starts, even if the text
        // itself flows onward; paragraph/line-relative anchors use this y.
        for (const image of para.images ?? []) {
          if (!image.floating) continue
          const w = emuToPx(image.widthEmu)
          const h = emuToPx(image.heightEmu)
          if (w <= 0 || h <= 0) continue
          const place = resolveFloating(image, { pageWidthPx: widthPx, pageHeightPx: heightPx, contentX, contentWidth, flowY: y, m })
          page.images.push({
            xPx: place.x,
            yPx: place.y,
            widthPx: w,
            heightPx: h,
            imageIndex: imageIndex.get(image) ?? -1,
            drawing: image.drawing,
            floating: {
              behindDoc: image.floating.behindDoc,
              relativeHeight: image.floating.relativeHeight,
              wrap: image.floating.wrap
            }
          })
        }
        const { lines, endY } = layoutParagraph(para, measure, {
          contentX,
          contentWidth,
          startY: y,
          defaults,
          // body text flows beside floating objects that wrap square/tight
          floatBands: (top, bottom) => freeBandsFor(page.images, top, bottom, contentX, contentWidth)
        })
        // Lines carry absolute y measured from the section start. Any line
        // past contentBottom rolls onto a fresh page, rebased to the top
        // margin; the same shift applies to the paragraph's remaining lines
        // and to the running flow position.
        let shift = 0
        for (const line of lines) {
          if (
            line.yPx + shift + line.heightPx > contentBottom &&
            (page.lines.length > 0 || page.tables.length > 0 || page.images.length > 0)
          ) {
            commitPage()
            shift = contentTop - line.yPx
          }
          for (const box of line.inlineImages ?? [])
            page.images.push({
              ...box,
              yPx: box.yPx + shift,
              clip: shiftedImageClip(box.clip, 0, shift),
              imageIndex: imageIndex.get(box.image) ?? -1,
              drawing: box.image.drawing
            })
          if (line.segs.length || !line.inlineImages?.length || line.logical)
            page.lines.push({
              ...line,
              yPx: line.yPx + shift,
              baselinePx: line.baselinePx === undefined ? undefined : line.baselinePx + shift
            })
        }
        y = endY + shift
        y += twipsToPx(para.spacingAfterTwips ?? 0)
        continue
      }
      // Table block: split at row boundaries across pages, repeating header
      // rows (w:tblHeader) on each continuation page.
      const table = autoWidthTable(block.table, measure, defaults, contentWidth)
      const tableX = m.left
      const hasHeader = table.rows.some((r) => r.isHeader)
      const pageHasContent = () => page.lines.length > 0 || page.tables.length > 0 || page.images.length > 0
      let fromRow = 0
      let firstChunk = true
      while (fromRow < table.rows.length) {
        let chunk = layoutTableRows(table, measure, defaults, {
          fromRow,
          repeatHeader: !firstChunk && hasHeader,
          maxHeightPx: contentBottom - y,
          allowFirstRowOverflow: !pageHasContent()
        })
        if (chunk.consumedRows === 0 && pageHasContent()) {
          // not even one row fits in the remaining space — start a new page
          commitPage()
          chunk = layoutTableRows(table, measure, defaults, {
            fromRow,
            repeatHeader: !firstChunk && hasHeader,
            maxHeightPx: contentBottom - y,
            allowFirstRowOverflow: true
          })
        }
        for (const line of chunk.lines) {
          for (const box of line.inlineImages ?? [])
            page.images.push({
              ...box,
              xPx: box.xPx + tableX,
              yPx: box.yPx + y,
              clip: shiftedImageClip(box.clip, tableX, y),
              imageIndex: imageIndex.get(box.image) ?? -1,
              drawing: box.image.drawing
            })
          if (line.segs.length || !line.inlineImages?.length || line.logical)
            page.lines.push({
              ...line,
              xPx: line.xPx + tableX,
              yPx: line.yPx + y,
              baselinePx: line.baselinePx === undefined ? undefined : line.baselinePx + y,
              segs: line.segs.map((seg) =>
                seg.transform ? { ...seg, transform: { ...seg.transform, e: seg.transform.e + tableX, f: seg.transform.f + y } } : seg
              )
            })
        }
        page.tables.push({ xPx: tableX, yPx: y, widthPx: chunk.widthPx, rows: chunk.rows })
        y += chunk.heightPx
        fromRow += chunk.consumedRows
        firstChunk = false
        if (fromRow < table.rows.length) commitPage()
      }
    }
    // A modeled section is real, including an explicitly finalized empty
    // nextPage section. Continuous sections have reused their existing page.
    pages.push(page)
    previousEndY = y

  }
  refreshRepeatedImages(pages, measure, totalPages)
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
export interface RenderPagesOptions {
  /** First page's number for PAGE field substitution (default 1). */
  pageNumberStart?: number
  /** Total for NUMPAGES substitution (default pages.length). */
  totalPages?: number
  /** Watermark stamped onto every page. Drawn under the content, not over it. */
  watermark?: WatermarkOptions | ResolvedWatermark
  /** Optional decoded-asset lookup for nested picture paint. */
  assets?: ContentPaintAssets
}

export function renderPages(
  pages: PageLayout[],
  ctx: CanvasRenderingContext2D,
  images?: Array<CanvasImageSource | undefined>,
  options?: RenderPagesOptions,
): void {
  ctx.save()
  ctx.textBaseline = 'alphabetic'
  ctx.fillStyle = '#000000'
  let lastFont = ''
  const resolveFamily = withFallbackFonts(fontFamilyCss, options?.assets?.fallbackFonts)
  const measure = createMeasurer(ctx, resolveFamily)
  const firstNumber = options?.pageNumberStart ?? 1
  const totalPages = options?.totalPages ?? pages.length
  pages.forEach((page, pageIndex) => {
    // The watermark goes down first, straight onto the page background, so the
    // document's own text is drawn on top of it rather than being washed out.
    if (options?.watermark) {
      paintWatermark(ctx, { widthPx: page.widthPx, heightPx: page.heightPx }, options.watermark)
    }
    const paintLines = (allLines: LineBox[]) => {
      for (const line of allLines) {
        let offset = 0
        if (line.align === 'center') offset = (line.contentWidthPx - line.widthPx) / 2
        else if (line.align === 'right') offset = line.contentWidthPx - line.widthPx
        let x = line.xPx + offset
        let maxAscent = 0
        for (const seg of line.segs) maxAscent = Math.max(maxAscent, seg.style.fontSizePt * LINE_HEIGHT_FACTOR * 0.8)
        const baseline = line.baselinePx ?? line.yPx + maxAscent
        const record = (ctx as TextRecordingContext)[RECORD_TEXT]
        if (!line.segs.length && line.logical && record) {
          if (line.defaultStyle) { ctx.font = fontCss(line.defaultStyle, resolveFamily); lastFont = ctx.font }
          record('', line.xPx, baseline, 0, line.logical)
        }
        // list marker sits left of the (indented) text
        if (line.marker) {
          const markerStyle: RunStyle = {
            fontFamily: seg0Font(line),
            fontSizePt: seg0Size(line),
            bold: !!line.segs[0]?.run.bold,
            italic: !!line.segs[0]?.run.italic,
          }
          const font = fontCss(markerStyle, resolveFamily)
          if (font !== lastFont) {
            ctx.font = font
            lastFont = font
          }
          ctx.fillStyle = line.segs[0]?.run.color ? resolveColor(line.segs[0].run.color) : '#000000'
          ctx.fillText(line.marker.text.trimEnd(), line.xPx - line.marker.widthPx, baseline)
        }
        for (const seg of line.segs) {
          const fontCssStr = fontCss(seg.style, resolveFamily)
          if (fontCssStr !== lastFont) {
            ctx.font = fontCssStr
            lastFont = fontCssStr
          }
          if (seg.transform) {
            // Direction-rotated glyphs carry absolute page placement; the
            // shared layout engine already resolved rotation and position.
            const t = seg.transform
            const size = seg.style.fontSizePt
            ctx.save()
            try {
              ctx.transform(t.a, t.b, t.c, t.d, t.e, t.f)
              if (seg.run.highlight) {
                const hl = HIGHLIGHT_CSS[seg.run.highlight] ?? seg.run.highlight
                ctx.fillStyle = hl
                ctx.fillRect(0, -maxAscent, seg.widthPx, line.heightPx)
              }
              ctx.fillStyle = seg.run.color ? resolveColor(seg.run.color) : '#000000'
              if (record && seg.logical) record(seg.text, 0, 0, seg.widthPx, seg.logical)
              else ctx.fillText(seg.text, 0, 0)
              if (seg.run.underline || seg.run.strike) {
                ctx.strokeStyle = ctx.fillStyle
                ctx.lineWidth = Math.max(1, size * 0.06)
                ctx.beginPath()
                const yy = seg.run.underline ? size * 0.15 : -size * 0.3
                ctx.moveTo(0, yy)
                ctx.lineTo(seg.widthPx, yy)
                ctx.stroke()
              }
            } finally {
              ctx.restore()
            }
            continue
          }
          const segX = seg.penOffset !== undefined ? line.xPx + offset + seg.penOffset : x
          if (seg.run.highlight) {
            const hl = HIGHLIGHT_CSS[seg.run.highlight] ?? seg.run.highlight
            ctx.fillStyle = hl
            ctx.fillRect(segX, line.yPx, seg.widthPx, line.heightPx)
          }
          ctx.fillStyle = seg.run.color ? resolveColor(seg.run.color) : '#000000'
          ctx.fillText(seg.text, segX, baseline)
          if (seg.run.underline || seg.run.strike) {
            ctx.strokeStyle = ctx.fillStyle
            ctx.lineWidth = Math.max(1, seg.style.fontSizePt * 0.06)
            ctx.beginPath()
            const yy = seg.run.underline ? baseline + seg.style.fontSizePt * 0.15 : baseline - seg.style.fontSizePt * 0.3
            ctx.moveTo(segX, yy)
            ctx.lineTo(segX + seg.widthPx, yy)
            ctx.stroke()
          }
          x = segX + seg.widthPx
        }
      }
    }
    paintTables(page.tables, ctx)
    const bodyImages = page.images.filter(image => !image.repeated)
    paintImages(bodyImages, ctx, images, 'behind', options?.assets)
    paintImages(bodyImages, ctx, images, 'front', options?.assets)
    paintLines(page.lines)
    for (const hf of [page.header, page.footer]) {
      if (!hf) continue
      const repeated = layoutRepeated(hf, measure, firstNumber + pageIndex, totalPages, page)
      ctx.save()
      ctx.beginPath()
      ctx.rect(0, hf.clipTopPx ?? 0, page.widthPx, Math.max(0, (hf.clipBottomPx ?? page.heightPx) - (hf.clipTopPx ?? 0)))
      ctx.clip()
      paintTables(repeated.tables, ctx)
      paintImages(repeated.images, ctx, images, 'behind', options?.assets)
      paintImages(repeated.images, ctx, images, 'front', options?.assets)
      lastFont = ''
      paintLines(repeated.lines)
      ctx.restore()
      lastFont = ''
    }
  })
  ctx.restore()
}

/** Replace field runs (PAGE, NUMPAGES) with this page's values. */
function substituteFields(paragraphs: DocxParagraph[], pageNumber: number, totalPages: number): DocxParagraph[] {
  let needsSub = false
  for (const p of paragraphs) {
    for (const r of p.runs)
      if (r.field) {
        needsSub = true
        break
      }
    if (needsSub) break
  }
  if (!needsSub) return paragraphs
  return paragraphs.map((p) => {
    const substituted = new Map<DocxTextRun, DocxTextRun>()
    const runs = p.runs.map((r) => {
      if (!r.field) return r
      const value = r.field === 'PAGE' ? String(pageNumber) : r.field === 'NUMPAGES' ? String(totalPages) : ''
      const run = { ...r, text: value }
      substituted.set(r, run)
      return run
    })
    return {
      ...p,
      runs,
      inline: p.inline?.map((item) =>
        item.kind === 'text' ? { ...item, run: substituted.get(item.run) ?? item.run } : item
      )
    }
  })
}

interface FloatingCellAnchor { image: DocxImage; contentX: number; contentWidth: number; flowY: number }
interface RepeatedLayout { lines: LineBox[]; tables: TableBox[]; images: ImageBox[]; top: number; bottom: number; occupied: boolean }
/** Fixed-page block layout used for bounds, field measurement and painting. */
function layoutRepeated(hf: HFBlock, measure: MeasureFn, pageNumber: number, totalPages: number, page: Pick<PageLayout, 'widthPx' | 'heightPx'>): RepeatedLayout {
  const out: RepeatedLayout = { lines: [], tables: [], images: [], top: hf.yPx, bottom: hf.yPx, occupied: false }
  let y = 0
  const floating: FloatingCellAnchor[] = []
  const addLines = (lines: LineBox[], x = 0, dy = 0) => {
    for (const line of lines) {
      out.lines.push({ ...line, xPx: line.xPx + x, yPx: line.yPx + dy, baselinePx: line.baselinePx === undefined ? undefined : line.baselinePx + dy })
      for (const box of line.inlineImages ?? []) out.images.push({ ...box, xPx: box.xPx + x, yPx: box.yPx + dy, clip: shiftedImageClip(box.clip, x, dy), imageIndex: hf.imageIndex?.get(box.image) ?? -1, drawing: box.image.drawing })
    }
  }
  for (const block of hf.blocks ?? hf.paragraphs.map(paragraph => ({ kind: 'p' as const, paragraph }))) {
    if (block.kind === 'p') {
      const para = substituteFields([block.paragraph], pageNumber, totalPages)[0]
      y += twipsToPx(para.spacingBeforeTwips ?? 0)
      const laid = layoutParagraph(para, measure, { contentX: hf.xPx, contentWidth: hf.widthPx, startY: y, defaults: hf.defaults })
      addLines(laid.lines)
      for (const image of para.images ?? []) if (image.floating) floating.push({ image, contentX: hf.xPx, contentWidth: hf.widthPx, flowY: y })
      y = laid.endY + twipsToPx(para.spacingAfterTwips ?? 0)
    } else {
      const substituted: DocxTable = { ...block.table, rows: block.table.rows.map(row => ({ ...row, cells: row.cells.map(cell => ({ ...cell, paragraphs: substituteFields(cell.paragraphs, pageNumber, totalPages) })) })) }
      const table = autoWidthTable(substituted, measure, hf.defaults, hf.widthPx)
      const laid = layoutTableRows(table, measure, hf.defaults, { fromRow: 0, repeatHeader: false, maxHeightPx: Infinity, allowFirstRowOverflow: true })
      out.tables.push({ xPx: hf.xPx, yPx: y, widthPx: laid.widthPx, rows: laid.rows })
      addLines(laid.lines, hf.xPx, y)
      floating.push(...laid.floatingAnchors.map(anchor => ({ ...anchor, contentX: anchor.contentX + hf.xPx, flowY: anchor.flowY + y })))
      y += laid.heightPx
    }
  }
  const offset = hf.footer ? hf.yPx - y : hf.yPx
  for (const line of out.lines) { line.yPx += offset; if (line.baselinePx !== undefined) line.baselinePx += offset }
  for (const table of out.tables) table.yPx += offset
  for (const image of out.images) { image.yPx += offset; image.clip = shiftedImageClip(image.clip, 0, offset) }
  for (const { image, flowY, contentX, contentWidth } of floating) {
    const pos = resolveFloating(image, { pageWidthPx: page.widthPx, pageHeightPx: page.heightPx, contentX, contentWidth, flowY: offset + flowY, m: hf.margins ?? { top: offset, bottom: page.heightPx - offset, left: hf.xPx, right: page.widthPx - hf.xPx - hf.widthPx } })
    out.images.push({ xPx: pos.x, yPx: pos.y, widthPx: emuToPx(image.widthEmu), heightPx: emuToPx(image.heightEmu), imageIndex: hf.imageIndex?.get(image) ?? -1, drawing: image.drawing, floating: { ...image.floating! } })
  }
  out.occupied = y > 0 || out.lines.length > 0 || out.images.length > 0
  out.top = Math.min(offset, ...out.images.map(image => image.yPx))
  out.bottom = Math.max(offset + y, ...out.images.map(image => image.yPx + image.heightPx), ...out.lines.map(line => line.yPx + line.heightPx))
  return out
}
function refreshRepeatedImages(pages: PageLayout[], measure: MeasureFn, totalPages: number): void {
  for (const [i, page] of pages.entries()) {
    page.images = page.images.filter(image => !image.repeated)
    for (const kind of ['header', 'footer'] as const) if (page[kind]) {
      const laid = layoutRepeated(page[kind], measure, i + 1, totalPages, page)
      page.images.push(...laid.images.map(image => ({ ...image, repeated: kind })))
    }
  }
}

function seg0Font(line: LineBox): string {
  return line.segs[0]?.style.fontFamily ?? 'Calibri'
}

function seg0Size(line: LineBox): number {
  return line.segs[0]?.style.fontSizePt ?? 11
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
/**
 * Draw images. `pass` selects behind-document anchors (drawn before the text
 * layer) or everything else; the two calls bracket the text paint so a
 * w:behindDoc image sits under it.
 */
function paintImages(
  boxes: ImageBox[],
  ctx: CanvasRenderingContext2D,
  images: Array<CanvasImageSource | undefined> | undefined,
  pass: 'behind' | 'front',
  assets?: ContentPaintAssets
): void {
  for (const box of boxes) {
    if (pass === 'behind' ? !box.floating?.behindDoc : box.floating?.behindDoc) continue
    if (box.clip) {
      ctx.save(); ctx.beginPath(); ctx.rect(box.clip.xPx, box.clip.yPx, box.clip.widthPx, box.clip.heightPx); ctx.clip()
    }
    try {
      if (box.drawing) {
        ctx.save()
        try { ctx.translate(box.xPx, box.yPx); paintDrawing(box.drawing, ctx, box.widthPx, box.heightPx, assets) }
        finally { ctx.restore() }
      } else {
        const img = images?.[box.imageIndex]
        if (img) ctx.drawImage(img as CanvasImageSource, box.xPx, box.yPx, box.widthPx, box.heightPx)
      }
    } finally { if (box.clip) ctx.restore() }
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
        const cx = table.xPx + cell.xPx
        const cy = table.yPx + cell.yPx
        const draw = (side: 'left' | 'right' | 'top' | 'bottom', x1: number, y1: number, x2: number, y2: number) => {
          const spec = cell.borderSpecs?.[side]
          ctx.strokeStyle = spec?.color && spec.color !== 'auto' ? resolveColor(spec.color) : '#000000'
          ctx.lineWidth =
            spec?.widthPt !== undefined ? (spec.widthPt * 4) / 3 : Math.max(1, BORDER_WIDTH[b[side] ?? 'thin'] ?? 1)
          const offset = spec?.widthPt !== undefined ? 0 : 0.5
          ctx.beginPath()
          ctx.moveTo(x1 + offset, y1 + offset)
          ctx.lineTo(x2 + offset, y2 + offset)
          ctx.stroke()
        }
        if (b.left) draw('left', cx, cy, cx, cy + cell.heightPx)
        if (b.right) draw('right', cx + cell.widthPx, cy, cx + cell.widthPx, cy + cell.heightPx)
        if (b.top) draw('top', cx, cy, cx + cell.widthPx, cy)
        if (b.bottom) draw('bottom', cx, cy + cell.heightPx, cx + cell.widthPx, cy + cell.heightPx)
      }
    }
  }
}

/** Origin/size reference for a floating image, in page coordinates. */
function floatingOrigin(
  relativeFrom: string,
  ctx: { pageWidthPx: number; pageHeightPx: number; contentX: number; contentWidth: number; flowY: number; m: { top: number; left: number; right: number; bottom: number } },
): { x: number; y: number; width: number; height: number } {
  const pageWidth = ctx.pageWidthPx
  const pageHeight = ctx.pageHeightPx
  switch (relativeFrom) {
    case 'page':
      return { x: 0, y: 0, width: pageWidth, height: pageHeight }
    case 'margin':
      return { x: ctx.m.left, y: ctx.m.top, width: Math.max(0, pageWidth - ctx.m.left - ctx.m.right), height: Math.max(0, pageHeight - ctx.m.top - ctx.m.bottom) }
    case 'paragraph':
    case 'line':
      return { x: ctx.contentX, y: ctx.flowY, width: ctx.contentWidth, height: 0 }
    default:
      // 'column' and anything else behave like the text column
      return { x: ctx.contentX, y: ctx.flowY, width: ctx.contentWidth, height: 0 }
  }
}

/** Gap left around a wrapped float (Word's square-wrap default is ~0.13in). */
const WRAP_GAP_PX = 12

/**
 * Free horizontal bands on a page between `yTop` and `yBottom` after
 * subtracting every floating object that overlaps that band and declares a
 * wrap mode that excludes text from its side.
 *
 * `wrapTopAndBottom` is parsed but deliberately not narrowed here: pushing a
 * whole block above/below a float is a layout decision we do not make yet.
 */
export function freeBandsFor(
  floats: ImageBox[],
  yTop: number,
  yBottom: number,
  contentX: number,
  contentWidth: number,
): Array<{ x: number; width: number }> | undefined {
  const blocking = floats.filter(
    (f) =>
      f.floating &&
      (f.floating.wrap === 'square' || f.floating.wrap === 'tight') &&
      f.yPx < yBottom &&
      f.yPx + f.heightPx > yTop,
  )
  if (blocking.length === 0) return undefined
  let bands: Array<{ x: number; width: number }> = [{ x: contentX, width: contentWidth }]
  for (const f of blocking) {
    const rx1 = f.xPx - WRAP_GAP_PX
    const rx2 = f.xPx + f.widthPx + WRAP_GAP_PX
    const next: Array<{ x: number; width: number }> = []
    for (const b of bands) {
      const bx1 = b.x
      const bx2 = b.x + b.width
      if (rx2 <= bx1 || rx1 >= bx2) {
        next.push(b) // no overlap
        continue
      }
      if (rx1 > bx1) next.push({ x: bx1, width: rx1 - bx1 }) // left remainder
      if (rx2 < bx2) next.push({ x: rx2, width: bx2 - rx2 }) // right remainder
    }
    bands = next
  }
  return bands.filter((b) => b.width > 0)
}

/** Resolve a floating image's page position from its anchor. */
function resolveFloating(
  image: {
    widthEmu: number
    heightEmu: number
    floating?: {
      posH: { relativeFrom: string; offsetEmu: number; align?: string }
      posV: { relativeFrom: string; offsetEmu: number; align?: string }
    }
  },
  ctx: { pageWidthPx: number; pageHeightPx: number; contentX: number; contentWidth: number; flowY: number; m: { top: number; left: number; right: number; bottom: number } },
): { x: number; y: number } {
  const w = emuToPx(image.widthEmu)
  const h = emuToPx(image.heightEmu)
  const f = image.floating
  if (!f) return { x: ctx.contentX, y: ctx.flowY }
  const ox = floatingOrigin(f.posH.relativeFrom, ctx)
  const oy = floatingOrigin(f.posV.relativeFrom, ctx)
  let x = ox.x + emuToPx(f.posH.offsetEmu)
  let y = oy.y + emuToPx(f.posV.offsetEmu)
  const alignH = f.posH.align
  if (alignH === 'center') x = ox.x + Math.max(0, (ox.width - w) / 2)
  else if (alignH === 'right') x = ox.x + Math.max(0, ox.width - w)
  const alignV = f.posV.align
  if (alignV === 'center') y = oy.y + Math.max(0, (oy.height - h) / 2)
  else if (alignV === 'bottom') y = oy.y + Math.max(0, oy.height - h)
  return { x, y }
}

// ---------- table layout ----------

interface CellBorderCss { style?: string; color?: string }

/** A vertical merge: one cell spanning rows startRow..endRow in a column. */
interface MergeRegion {
  startRow: number
  endRow: number
  startCol: number
  colSpan: number
  fill?: string
  borders?: { left?: string; right?: string; top?: string; bottom?: string }
  borderSpecs?: TableCellBox['borderSpecs']
}

interface CellPosition {
  cell: DocxTableCell
  col: number
  span: number
}

/** Column index of every cell, mirroring the layout loop's grid accounting. */
function rowCellPositions(table: DocxTable): CellPosition[][] {
  return table.rows.map((row) => {
    const out: CellPosition[] = []
    let col = row.gridBefore ?? 0
    for (const cell of row.cells) {
      out.push({ cell, col, span: Math.max(1, cell.gridSpan) })
      col += Math.max(1, cell.gridSpan)
    }
    return out
  })
}

/**
 * Resolve vertical merges into regions keyed by `${row}:${col}`. A
 * `vMerge` restart opens a region; following `continue` cells close it.
 */
function computeMergeRegions(table: DocxTable, positions: CellPosition[][]): Map<string, MergeRegion> {
  const regions = new Map<string, MergeRegion>()
  const key = (r: number, c: number) => `${r}:${c}`
  for (let r = 0; r < table.rows.length; r++) {
    for (const { cell, col, span } of positions[r]) {
      if (cell.vMerge !== 'restart') continue
      const tb = table.borders
      const spec = (side: 'left' | 'right' | 'top' | 'bottom'): TableCellBorder | undefined => {
        const cb = cell.borders
        if (cb && side in cb) return cb[side]
        if (tb) return tb[side] ?? (side === 'top' || side === 'bottom' ? tb.insideH : tb.insideV)
        return undefined
      }
      const region: MergeRegion = {
        startRow: r,
        endRow: r,
        startCol: col,
        colSpan: span,
        fill: 'fill' in cell ? cell.fill : table.fill,
        borders: {
          left: borderCss(spec('left')),
          right: borderCss(spec('right')),
          top: borderCss(spec('top')),
          bottom: borderCss(spec('bottom'))
        },
        borderSpecs: { left: spec('left'), right: spec('right'), top: spec('top'), bottom: spec('bottom') }
      }
      // extend over following continue cells
      for (let rr = r + 1; rr < table.rows.length; rr++) {
        const cont = positions[rr].find((p) => p.col === col)
        if (!cont || cont.cell.vMerge !== 'continue') break
        region.endRow = rr
        regions.set(key(rr, col), region)
      }
      regions.set(key(r, col), region)
    }
  }
  return regions
}

function borderCss(b: CellBorderCss | undefined): string | undefined {
  if (!b || !b.style) return undefined
  if (b.style === 'single' || b.style === 'thin') return 'thin'
  return b.style
}

/** Derive preferred widths and shrink wrappable columns to the page width. */
function autoWidthTable(
  table: DocxTable,
  measure: MeasureFn,
  defaults: { fontFamily: string; fontSizePt: number },
  availableWidth: number,
): DocxTable {
  if (!table.autoWidth) return table
  const padding = twipsToPx(table.cellMargins.leftTwips + table.cellMargins.rightTwips)
  const widths = table.gridColsTwips.map(() => Math.max(padding + 1, 1))
  const minimums = [...widths]
  const expand = (columns: number[], column: number, span: number, needed: number) => {
    const current = columns.slice(column, column + span).reduce((a, b) => a + b, 0)
    const extra = Math.max(0, needed - current) / span
    for (let c = column; c < column + span && c < columns.length; c++) columns[c] += extra
  }
  for (const row of table.rows) {
    let column = row.gridBefore ?? 0
    for (const cell of row.cells) {
      const span = Math.max(1, cell.gridSpan)
      const cm = cell.margins ?? table.cellMargins
      const padding = twipsToPx(cm.leftTwips + cm.rightTwips)
      const textWidth = Math.max(0, ...cell.paragraphs.map(paragraph =>
        // Match pushWord's measurements, including spaces. Kerning across a
        // whole string can otherwise make an intrinsically sized cell wrap.
        paragraph.runs.reduce((width, run) => width + run.text.split(/( )/)
          .reduce((sum, chunk) => sum + measure(chunk, runStyleOf(run, defaults)), 0), 0)
          + twipsToPx((paragraph.indentLeftTwips ?? 0) + (paragraph.indentRightTwips ?? 0))))
      const wordWidth = Math.max(0, ...cell.paragraphs.flatMap(paragraph =>
        paragraph.runs.flatMap(run => run.text.split(/\s+/).map(word => measure(word, runStyleOf(run, defaults))))))
      const imageWidth = Math.max(0, ...cell.paragraphs.flatMap(paragraph =>
        paragraph.images.filter(image => !image.floating).map(image => emuToPx(image.widthEmu))))
      expand(widths, column, span, Math.max(textWidth, imageWidth) + padding + 0.01)
      expand(minimums, column, span, Math.max(wordWidth, imageWidth) + padding)
      column += span
    }
  }
  const total = widths.reduce((a, b) => a + b, 0)
  if (total > availableWidth) {
    const minimum = minimums.reduce((a, b) => a + b, 0)
    for (let i = 0; i < widths.length; i++) {
      widths[i] = minimum >= availableWidth ? minimums[i] * availableWidth / minimum
        : minimums[i] + (widths[i] - minimums[i]) * (availableWidth - minimum) / (total - minimum)
    }
  }
  return { ...table, gridColsTwips: widths.map(width => width * 15) }
}

/**
 * Lay out a horizontal slice of a table — cell rects + text lines relative to
 * (0,0) at the table origin. Starts at row `fromRow`, optionally repeats the
 * header rows, and stops before exceeding `maxHeightPx` so the caller can
 * split the table across pages. `consumedRows` counts source rows placed.
 */
/** Word cell textDirection mapped onto the shared text engine. lrTb keeps the
 * ordinary horizontal path. Word tbRl
 * rotates the whole line (shared vert); tbRlV/tbLrV keep mixed upright CJK
 * (shared eaVert/mongolianVert frames for RTL/LTR column progression). */
const CELL_TEXT_DIRECTIONS = { tbRl: 'vert', tbRlV: 'eaVert', btLr: 'vert270', tbLrV: 'mongolianVert', lrTbV: 'horz' } as const
type VerticalCellMode = keyof typeof CELL_TEXT_DIRECTIONS

/** Lay a cell paragraph in its local direction bounds, projecting shared
 * placed segments into LineBoxes with absolute page-space transforms. The
 * engine resolves wrapping, glyph orientation and column order; the cell
 * keeps margins, alignment slack, merged height and document flow. */
function layoutVerticalCellParagraphs(
  paras: DocxParagraph[],
  measure: MeasureFn,
  mode: VerticalCellMode,
  contentX: number,
  startY: number,
  innerW: number,
  defaults: { fontFamily: string; fontSizePt: number },
  innerH?: number,
  cellAlign: 'top' | 'center' | 'bottom' = 'top'
): { lines: LineBox[]; endY: number; perPara: Array<{ lines: LineBox[]; endY: number }> } {
  // Preserve shaped authored runs and paragraph-local source provenance.
  // The shared engine handles grapheme alignment across style boundaries. All paragraphs
  // share one engine call so consecutive paragraphs continue into fresh
  // transverse columns instead of restarting on the same one.
  const sources: DocxTextRun[][] = paras.map(() => [])
  const scope = {}
  const paragraphSources = paras.map((para, order) => ({ text: (para.inline?.filter(item => item.kind === 'text').map(item => item.run) ?? para.runs).map(run => run.text).join(''), scope, order }))
  const images: DocxImage[] = []
  const bodyParagraphs: Array<DrawingTextParagraph & { paraIndex: number }> = []
  for (const [paraIndex, para] of paras.entries()) {
    const runs: Array<{ text: string; fontSizePt: number; fontFamily: string; color?: string; bold?: boolean; italic?: boolean }> = []
    const inlineSlots: NonNullable<DrawingTextParagraph['inlineSlots']> = []
    let sourceOffset = 0
    const items = para.inline ?? [...para.runs.map(run => ({ kind: 'text' as const, run })), ...para.images.map(image => ({ kind: 'image' as const, image }))]
    for (const item of items) {
      if (item.kind === 'image') {
        if (item.image.floating) continue
        const width = emuToPx(item.image.widthEmu), height = emuToPx(item.image.heightEmu)
        if (width > 0 && height > 0 && Number.isFinite(width + height)) inlineSlots.push({ id: images.push(item.image) - 1, sourceOffset, width, height })
        continue
      }
      const run = item.run
      if (!run.text) continue
      const style = {
        fontSizePt: run.fontSizePt ?? defaults.fontSizePt,
        fontFamily: run.fontFamily ?? defaults.fontFamily,
        color: run.color,
        bold: run.bold,
        italic: run.italic
      }
      sources[paraIndex].push(run)
      runs.push({ ...style, text: run.text }); sourceOffset += run.text.length
    }
    bodyParagraphs.push({ runs, inlineSlots, align: para.align, level: 0, paraIndex,
      defaultProperties: { fontFamily: defaults.fontFamily, fontSizePt: defaults.fontSizePt },
      wordLineSpacing: para.lineSpacing ?? { rule: 'auto', value: 240 },
      spaceBefore: { kind: 'points', value: (para.spacingBeforeTwips ?? 0) / 20 },
      spaceAfter: { kind: 'points', value: (para.spacingAfterTwips ?? 0) / 20 } })
  }
  if (!bodyParagraphs.length) return { lines: [], endY: startY, perPara: paras.map(() => ({ lines: [], endY: startY })) }
  // Paragraph indents ride on the engine paragraphs (twips to EMU), so every
  // paragraph keeps its resolved left/right/first-line properties instead of
  // only the first paragraph's. Multi-paragraph fixed-measure tests pin this.
  const EMU_PER_TWIP = 635
  for (const [bodyIndex, bodyPara] of bodyParagraphs.entries()) {
    const source = paras[bodyPara.paraIndex]
    const withIndents = {
      ...bodyPara,
      marginLeftEmu: (source.indentLeftTwips ?? 0) * EMU_PER_TWIP,
      marginRightEmu: (source.indentRightTwips ?? 0) * EMU_PER_TWIP,
      indentEmu: (source.indentFirstLineTwips ?? 0) * EMU_PER_TWIP
    }
    bodyParagraphs[bodyIndex] = withIndents
  }
  const flowX = contentX
  const flowW = innerW
  if (flowW <= 0) return { lines: [], endY: startY, perPara: paras.map(() => ({ lines: [], endY: startY })) }
  const direction = CELL_TEXT_DIRECTIONS[mode] as TextDirection
  // Auto/atLeast rows first measure their natural writing-local flow length.
  // Exact rows and the final projection use the actual cell interior instead.
  // In particular vert270's bottom origin must be a cell bound, not a sentinel.
  const localMetrics = (style: RunStyle, text?: string) => measure.metrics?.(style, text) ?? { ascent: style.fontSizePt * .8, descent: style.fontSizePt * .2, normalHeight: style.fontSizePt * 96 / 72 * 1.2 }
  const naturalFlowH = Math.max(1, ...bodyParagraphs.map(bodyPara => {
    const para = paras[bodyPara.paraIndex]
    const advance = bodyPara.runs.reduce((sum, run) => {
      const size = run.fontSizePt ?? defaults.fontSizePt
      const width = measure(run.text, { fontFamily: run.fontFamily ?? defaults.fontFamily, fontSizePt: size, bold: !!run.bold, italic: !!run.italic })
      const advance = direction === 'eaVert' || mode === 'tbLrV' ? graphemes(run.text).reduce((total, cluster) => {
        const orientation = clusterOrientation(cluster.text)
        const clusterWidth = measure(cluster.text, { fontFamily: run.fontFamily ?? defaults.fontFamily, fontSizePt: size, bold: !!run.bold, italic: !!run.italic })
        const ink = localMetrics({ fontFamily: run.fontFamily ?? defaults.fontFamily, fontSizePt: size, bold: !!run.bold, italic: !!run.italic }, cluster.text)
        return total + (orientation === 'U' || orientation === 'Tu' ? Math.max(size * 96 / 72, clusterWidth, ink.ascent + ink.descent) : clusterWidth)
      }, 0) : width
      return sum + advance
    }, 0)
    return advance + (bodyPara.inlineSlots ?? []).reduce((sum, slot) => sum + (mode === 'lrTbV' ? slot.width : slot.height), 0) + twipsToPx((para.indentLeftTwips ?? 0) + (para.indentRightTwips ?? 0) + Math.max(0, para.indentFirstLineTwips ?? 0))
  }))
  const flowH = innerH === undefined ? naturalFlowH : Math.max(0, innerH)
  const isRotated = (text: string) => { const orientation = clusterOrientation(text); return orientation === 'U' || orientation === 'Tu' }
  const projectedAdvance = (text: string, style: RunStyle): number => {
    if (mode !== 'lrTbV') return measure(text, style)
    let width = 0, shaped = ''
    for (const cluster of graphemes(text)) {
      if (isRotated(cluster.text)) { const ink = localMetrics(style, cluster.text); width += measure(shaped, style) + ink.ascent + ink.descent; shaped = '' }
      else shaped += cluster.text
    }
    return width + measure(shaped, style)
  }
  const laid = layoutTextBody(
    { direction, paragraphs: bodyParagraphs, anchor: 't', wrap: true,
      insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0 },
    flowW, flowH, (text, style) => {
      const runStyle = { fontFamily: style.fontFamily ?? defaults.fontFamily, fontSizePt: style.fontSizePt ?? defaults.fontSizePt,
        bold: !!style.bold, italic: !!style.italic }
      const metrics = localMetrics(runStyle, text)
      const projectedHeight = mode === 'lrTbV' ? Math.max(0, ...graphemes(text).filter(cluster => isRotated(cluster.text)).map(cluster => measure(cluster.text, runStyle))) : 0
      return { width: projectedAdvance(text, runStyle), ...metrics, normalHeight: Math.max(metrics.normalHeight ?? metrics.ascent + metrics.descent, projectedHeight) }
    }
  )
  // Group placed lines back to source paragraphs; placed paragraphIndex
  // counts body paragraphs, mapped through paraIndex. Origins are built
  // box-local in e; the cy walk below assigns absolute f per paragraph
  // (shared flow origin for vertical modes, stacked with spacing for lrTbV).
  type BuiltSeg = { seg: Segment; owner: number; flowAdvance: number; size: number; upright: boolean }
  type BuiltLine = { line: (typeof laid.lines)[number]; owner: number; para: DocxParagraph; segs: BuiltSeg[]; minE: number; maxE: number; maxSize: number; isParagraphEnd: boolean }
  const built: BuiltLine[] = []
  for (const [lineIndex, line] of laid.lines.entries()) {
    const owner = bodyParagraphs[line.paragraphIndex]?.paraIndex ?? 0
    const para = paras[owner]
    const isParagraphEnd = lineIndex === laid.lines.length - 1 || laid.lines[lineIndex + 1].paragraphIndex !== line.paragraphIndex
    const out: BuiltLine = { line, owner, para, segs: [], minE: Infinity, maxE: -Infinity, maxSize: 0, isParagraphEnd }
    if (mode === 'lrTbV') {
      // Replay measured horizontal positions, including authored indents,
      // justification and physical slots. Latin stays shaped; East Asian
      // glyphs occupy their measured CW band, centered in the resolved line.
      type Bit = { text: string; origin: DocxTextRun; style: RunStyle; advance: number; inkWidth: number; x: number; upright: boolean; start: number; end: number; runIndex: number }
      const bits: Bit[] = []
      for (const seg of line.segments) {
        const origin = sources[owner]?.[seg.runIndex]
        if (!origin) continue
        const style: RunStyle = {
          fontFamily: origin.fontFamily ?? defaults.fontFamily,
          fontSizePt: origin.fontSizePt ?? defaults.fontSizePt,
          bold: !!origin.bold,
          italic: !!origin.italic
        }
        for (const cluster of graphemes(seg.text)) {
          const orientation = clusterOrientation(cluster.text)
          const upright = orientation === 'U' || orientation === 'Tu'
          const x = seg.x + projectedAdvance(seg.text.slice(0, cluster.start), style)
          const previous = bits[bits.length - 1]
          if (!upright && previous && !previous.upright && previous.origin === origin && previous.end === seg.sourceStart + cluster.start && Math.abs(previous.x + previous.advance - x) < 1e-8 && (para.align !== 'justify' || cluster.text !== ' ' && previous.text !== ' ')) {
            previous.text += cluster.text; previous.advance = previous.inkWidth = measure(previous.text, style); previous.end = seg.sourceStart + cluster.start + cluster.text.length
          } else bits.push({ text: cluster.text, origin, style, advance: projectedAdvance(cluster.text, style), inkWidth: measure(cluster.text, style), x, upright,
            start: seg.sourceStart + cluster.start, end: seg.sourceStart + cluster.start + cluster.text.length, runIndex: seg.runIndex })
        }
      }
      for (const bit of bits) {
        out.maxSize = Math.max(out.maxSize, bit.style.fontSizePt)
        // Box-local origin (the second pass adds contentX/cy); paragraph
        // indents already ride the engine paragraphs measured above.
        const metrics = localMetrics(bit.style, bit.text)
        const e = bit.x + (bit.upright ? metrics.ascent : 0)
        const f = bit.upright ? line.y + (line.height - bit.inkWidth) / 2 : line.baseline
        out.segs.push({
          seg: {
            text: bit.text,
            run: bit.origin,
            style: bit.style,
            widthPx: bit.inkWidth,
            logical: { source: paragraphSources[owner], start: bit.start, end: bit.end, run: bit.runIndex, line: lineIndex },
            transform: bit.upright
              ? { a: 0, b: 1, c: -1, d: 0, e, f }
              : { a: 1, b: 0, c: 0, d: 1, e, f }
          },
          owner,
          flowAdvance: bit.upright ? bit.advance : bit.style.fontSizePt * 0.2,
          size: bit.style.fontSizePt,
          upright: bit.upright
        })
        out.minE = Math.min(out.minE, e)
        out.maxE = Math.max(out.maxE, e)
      }
    } else {
      for (const seg of line.segments) {
        const origin = sources[owner]?.[seg.runIndex]
        if (!origin) continue
        const style: RunStyle = {
          fontFamily: origin.fontFamily ?? defaults.fontFamily,
          fontSizePt: origin.fontSizePt ?? defaults.fontSizePt,
          bold: !!origin.bold,
          italic: !!origin.italic
        }
        out.maxSize = Math.max(out.maxSize, style.fontSizePt)
        const widthPx = measure(seg.text, style)
        const placed = seg.transform
        if (placed) {
          out.segs.push({
            seg: { text: seg.text, run: origin, style, widthPx, transform: { ...placed }, logical: { source: paragraphSources[owner], start: seg.sourceStart, end: seg.sourceEnd,
              run: seg.runIndex, line: lineIndex, flow: 'vertical', graphemeBoundaries: seg.graphemeBoundaries,
              ...(seg.orientation === 'upright' ? { ink: { ascent: localMetrics(style, seg.text).ascent, descent: localMetrics(style, seg.text).descent } } : {}) } },
            owner,
            flowAdvance: placed.b === 1 && placed.a === 0 && placed.c === -1 && placed.d === 0 ? widthPx
              : seg.orientation === 'upright' ? localMetrics(style, seg.text).descent : style.fontSizePt * 0.2,
            size: style.fontSizePt,
            upright: false
          })
          out.minE = Math.min(out.minE, placed.e)
          out.maxE = Math.max(out.maxE, placed.e)
        } else {
          out.segs.push({ seg: { text: seg.text, run: origin, style, widthPx }, owner, flowAdvance: 0, size: style.fontSizePt, upright: false })
        }
      }
    }
    for (const slot of line.inlineSlots ?? []) { out.minE = Math.min(out.minE, slot.x); out.maxE = Math.max(out.maxE, slot.x + slot.width) }
    built.push(out)
  }
  // Assign absolute origins: vertical modes share one flow origin (each
  // paragraph already continued into fresh transverse columns); lrTbV stacks
  // paragraphs with spacing like horizontal flow.
  const lines: LineBox[] = []
  const perPara: Array<{ lines: LineBox[]; endY: number }> = paras.map(() => ({ lines: [], endY: startY }))
  const horizontalSlack = mode === 'lrTbV' && innerH !== undefined ? Math.max(0, innerH - laid.height) : 0
  const cy = startY + (cellAlign === 'center' ? horizontalSlack / 2 : cellAlign === 'bottom' ? horizontalSlack : 0)
  // A rotated Word cell's vertical alignment is transverse to its paragraph
  // flow: top/center/bottom become left/center/right for bottom-to-top text.
  const transverseSlack = Math.max(0, innerW - laid.height)
  const cellShift = cellAlign === 'center' ? transverseSlack / 2 : cellAlign === 'bottom' ? transverseSlack : 0
  for (const [builtIndex, item] of built.entries()) {
    const finalSegs: Segment[] = []
    let lineFlow = 0
    // Transverse paragraph margins anchor vertical columns: the shared
    // engine narrows wrapping but always anchors at the body edge, so shift
    // RTL columns left by the right indent. btLr indents are already resolved
    // along its writing-local flow; cell alignment positions its column block.
    // lrTbV placed x already carries engine paragraph offsets.
    const ownerPara = paras[item.owner]
    const transverseShift = mode === 'tbRl' || mode === 'tbRlV'
      ? -twipsToPx(ownerPara.indentRightTwips ?? 0) - cellShift
      : mode === 'btLr' ? cellShift : 0
    for (const builtSeg of item.segs) {
      const absolute: LocalAffine | undefined = builtSeg.seg.transform
        ? { ...builtSeg.seg.transform, e: builtSeg.seg.transform.e + contentX + transverseShift, f: builtSeg.seg.transform.f + cy }
        : undefined
      finalSegs.push(absolute ? { ...builtSeg.seg, transform: absolute } : builtSeg.seg)
      if (absolute) lineFlow = Math.max(lineFlow, mode === 'btLr'
        ? flowH - builtSeg.seg.transform!.f + builtSeg.seg.widthPx
        : builtSeg.seg.transform!.f + builtSeg.flowAdvance)
    }
    const inlineImages = item.line.inlineSlots?.map(slot => ({ image: images[slot.id], xPx: contentX + transverseShift + slot.x, yPx: cy + slot.y, widthPx: slot.width, heightPx: slot.height,
      clip: innerH === undefined ? undefined : { xPx: contentX, yPx: startY, widthPx: innerW, heightPx: Math.max(0, innerH) } }))
    for (const image of inlineImages ?? []) lineFlow = Math.max(lineFlow, image.yPx + image.heightPx - cy)
    const flowTop = mode === 'btLr' && finalSegs.length ? Math.min(...finalSegs.map(seg => seg.transform!.f - seg.widthPx)) : cy + item.line.y
    const flowBottom = mode === 'btLr' && finalSegs.length ? Math.max(...finalSegs.map(seg => seg.transform!.f)) : flowTop + Math.max(item.line.height, lineFlow)
    const minE = (item.minE === Infinity ? item.line.x : item.minE) + contentX + transverseShift
    const extent = item.maxE === -Infinity ? 0 : item.maxE - item.minE
    const box: LineBox = {
      xPx: mode === 'lrTbV' ? flowX : minE,
      yPx: flowTop,
      widthPx: mode === 'lrTbV' ? innerW : Math.max(1, extent) + item.maxSize,
      heightPx: mode === 'btLr' ? flowBottom - flowTop : mode === 'lrTbV' ? item.line.height : Math.max(item.line.height, lineFlow),
      segs: finalSegs, inlineImages,
      align: mode === 'lrTbV' ? item.para.align : 'left',
      isParagraphEnd: item.isParagraphEnd,
      contentWidthPx: innerW,
      baselinePx: cy + item.line.baseline,
      logical: { source: paragraphSources[item.owner], start: item.line.inlineSlots?.[0]?.sourceOffset ?? paragraphSources[item.owner].text.length, end: item.line.inlineSlots?.[0]?.sourceOffset ?? paragraphSources[item.owner].text.length,
        line: builtIndex, ...(mode === 'lrTbV' ? {} : { flow: 'vertical' as const }) },
      defaultStyle: { fontFamily: paras[item.owner].runs[0]?.fontFamily ?? defaults.fontFamily, fontSizePt: paras[item.owner].runs[0]?.fontSizePt ?? defaults.fontSizePt, bold: !!paras[item.owner].runs[0]?.bold, italic: !!paras[item.owner].runs[0]?.italic }
    }
    lines.push(box)
    perPara[item.owner].lines.push(box)
    const endY = mode === 'lrTbV' ? cy + item.line.y + item.line.height : cy + lineFlow
    perPara[item.owner].endY = Math.max(perPara[item.owner].endY, endY)
  }
  if (mode === 'tbLrV' && lines.length) {
    // The shared mixed LTR frame preserves each column's physical pitch,
    // including images wider than the text. Anchor the complete block left.
    const leftmost = Math.min(...lines.map(line => line.xPx).filter(Number.isFinite))
    const anchor = contentX + cellShift - leftmost
    if (anchor !== 0) {
      for (const line of lines) {
        line.xPx += anchor
        for (const seg of line.segs) if (seg.transform) seg.transform = { ...seg.transform, e: seg.transform.e + anchor }
        for (const image of line.inlineImages ?? []) image.xPx += anchor
      }
    }
  }
  // Intrinsic row sizing includes the trailing paragraph margin as well as
  // glyph advances. Dropping it would make final projection wrap a row that
  // was just measured as one column.
  const endY = mode === 'btLr' && innerH === undefined ? startY + naturalFlowH : mode === 'lrTbV' ? cy + laid.height : Math.max(startY, ...perPara.map(p => p.endY))
  return { lines, endY, perPara }
}
function layoutTableRows(
  table: DocxTable,
  measure: MeasureFn,
  defaults: { fontFamily: string; fontSizePt: number },
  opts: { fromRow: number; repeatHeader: boolean; maxHeightPx: number; allowFirstRowOverflow: boolean }
): { lines: LineBox[]; rows: TableRowBox[]; floatingAnchors: FloatingCellAnchor[]; widthPx: number; heightPx: number; consumedRows: number } {
  const lines: LineBox[] = []
  const floatingAnchors: FloatingCellAnchor[] = []
  const colOffsets = prefixSum(table.gridColsTwips)
  const colWidths = table.gridColsTwips
  const widthPx = twipsToPx(colOffsets[colOffsets.length - 1] ?? 0)


  const committedRows: TableRowBox[] = []
  const positions = rowCellPositions(table)
  const regions = computeMergeRegions(table, positions)
  const regionKey = (r: number, c: number) => `${r}:${c}`
  // smallest source row this chunk lays out; merges starting before it were
  // opened on an earlier page and must be continued as plain boxes here
  const minRowIndex = Math.min(
    opts.fromRow,
    ...(opts.repeatHeader ? table.rows.map((r, i) => (r.isHeader ? i : Infinity)) : [Infinity])
  )

  const headerRows = opts.repeatHeader ? table.rows.filter((r) => r.isHeader) : []
  const headerIndices = table.rows.map((r, i) => (r.isHeader ? i : -1)).filter((i) => i >= 0)
  const queue: Array<{ row: DocxTableRow; rowIndex: number; counts: boolean }> = [
    ...headerRows.map((row, i) => ({ row, rowIndex: headerIndices[i] ?? 0, counts: false })),
    ...table.rows.map((row, i) => ({ row, rowIndex: i, counts: true })).slice(opts.fromRow)
  ]

  let yRel = 0
  let consumedRows = 0
  let contentPlaced = 0

  /** merge regions whose box is still open in this chunk */
  const openMerges = new Map<string, TableCellBox>()
  const mergedAlignments: Array<{ box: TableCellBox; lines: LineBox[]; anchors: FloatingCellAnchor[]; initialHeight: number; factor: number }> = []
  const finalDirectionLayouts: Array<{ box: TableCellBox; start: number; count: number; project: (height: number) => LineBox[] }> = []

  for (const { row, rowIndex, counts } of queue) {
    // lay the row into temporary coordinates (relative to its own top)
    const entries: Array<{
      cellBox: TableCellBox
      lines: LineBox[]
      anchors: FloatingCellAnchor[]
      contentH: number
      vAlign: 'top' | 'center' | 'bottom'
      col: number
      project?: (height: number) => LineBox[]
    }> = []
    let rowContentH = 0
    for (const { cell, col, span } of positions[rowIndex]) {
      if (cell.vMerge === 'continue') {
        continue
      }
      let cellW = 0
      for (let i = col; i < col + span && i < colWidths.length; i++) cellW += colWidths[i]
      const cellXPx = twipsToPx(colOffsets[col] ?? 0)
      const cm = cell.margins ?? table.cellMargins
      const margins = { top: twipsToPx(cm.topTwips), bottom: twipsToPx(cm.bottomTwips), left: twipsToPx(cm.leftTwips), right: twipsToPx(cm.rightTwips) }
      const innerW = twipsToPx(cellW) - margins.left - margins.right
      if (innerW <= 0) continue
      let cellLines: LineBox[] = []
      let project: ((height: number) => LineBox[]) | undefined
      const cellAnchors: FloatingCellAnchor[] = []
      let cy = margins.top
      const verticalMode = cell.textDirection && cell.textDirection !== 'lrTb' ? cell.textDirection : undefined
      if (verticalMode && cell.paragraphs.length) {
        // One shared engine call for the whole cell so consecutive paragraphs
        // continue into fresh transverse columns with spacing accounted.
        const layoutCell = (innerHeight?: number) => layoutVerticalCellParagraphs(cell.paragraphs, measure, verticalMode, cellXPx + margins.left, margins.top, innerW, defaults, innerHeight, cell.vAlign)
        const exactInnerH = row.heightRule === 'exact' && row.heightTwips !== undefined ? twipsToPx(row.heightTwips) - margins.top - margins.bottom : undefined
        const laid = layoutCell(exactInnerH)
        for (const para of cell.paragraphs) {
          for (const image of para.images ?? []) if (image.floating)
            cellAnchors.push({ image, contentX: cellXPx + margins.left, contentWidth: innerW, flowY: cy })
        }
        cellLines = laid.lines
        cy = laid.endY
        if (verticalMode) {
          // Row heights (including vertical merges) are known only after the
          // chunk is committed. Re-layout in that final writing-local frame;
          // paragraphs share its bottom, while wrapping uses its full height.
          project = height => {
            const final = layoutCell(height - margins.top - margins.bottom)
            return final.lines
          }
        }
      } else for (const para of cell.paragraphs) {
        cy += twipsToPx(para.spacingBeforeTwips ?? 0)
        for (const image of para.images ?? []) if (image.floating)
          cellAnchors.push({ image, contentX: cellXPx + margins.left, contentWidth: innerW, flowY: cy })
        const laid = layoutParagraph(para, measure, {
          contentX: cellXPx + margins.left,
          contentWidth: innerW,
          startY: cy,
          defaults
        })
        cellLines = cellLines.concat(laid.lines)
        cy = laid.endY + twipsToPx(para.spacingAfterTwips ?? 0)
      }
      const contentH =
        cellLines.length > 0
          ? cy + margins.bottom
          : margins.top + margins.bottom + defaults.fontSizePt * LINE_HEIGHT_FACTOR * (96 / 72) * 0.5
      rowContentH = Math.max(rowContentH, contentH)
      // borders: cell overrides, falling back to the table's outside/inside
      // border definitions
      const cb = cell.borders
      const tb = table.borders
      const spec = (side: 'left' | 'right' | 'top' | 'bottom'): TableCellBorder | undefined => {
        if (cb && side in cb) return cb[side]
        const exterior =
          side === 'left'
            ? col === 0
            : side === 'right'
              ? col + span === colWidths.length
              : side === 'top'
                ? rowIndex === 0
                : rowIndex === table.rows.length - 1
        if (tb) return exterior ? tb[side] : tb[side === 'left' || side === 'right' ? 'insideV' : 'insideH']
        return undefined
      }
      entries.push({
        cellBox: {
          xPx: cellXPx,
          yPx: yRel, // replaced when the row is committed
          widthPx: twipsToPx(cellW),
          heightPx: 0,
          fill: 'fill' in cell ? cell.fill : table.fill,
          borders: {
            left: borderCss(spec('left')),
            right: borderCss(spec('right')),
            top: borderCss(spec('top')),
            bottom: borderCss(spec('bottom'))
          },
          borderSpecs: { left: spec('left'), right: spec('right'), top: spec('top'), bottom: spec('bottom') }
        },
        lines: cellLines,
        anchors: cellAnchors,
        contentH,
        vAlign: cell.vAlign ?? 'top',
        col,
        project
      })
    }
    // explicit row height (atLeast semantics)
    let rowH = rowContentH
    if (row.heightTwips !== undefined) {
      const hPx = twipsToPx(row.heightTwips)
      if (row.heightRule === 'exact') rowH = hPx
      else rowH = Math.max(rowH, hPx)
    }
    if (row.heightRule !== 'exact') rowH = Math.max(rowH, defaults.fontSizePt * LINE_HEIGHT_FACTOR * (96 / 72))

    // stop before overflowing the available height, unless this is the first
    // content row of the chunk and the caller allows it
    const isFirstContent = counts && contentPlaced === 0
    // A repeated body-table header cannot consume the entire new page and
    // prevent its first source row from progressing. Both may overflow.
    const overflowingHeader = !counts && contentPlaced === 0 && opts.allowFirstRowOverflow
    if (!isFirstContent && !overflowingHeader && yRel + rowH > opts.maxHeightPx) break
    if (isFirstContent && !opts.allowFirstRowOverflow && yRel + rowH > opts.maxHeightPx) break

    // commit the row: vertical alignment shifts text within the final height
    const rowCells: TableCellBox[] = []

    // merge cells: an anchor owns a box that grows over its region; a region
    // that started on an earlier page is continued here as a plain box
    for (const { col, span } of positions[rowIndex]) {
      const region = regions.get(regionKey(rowIndex, col))
      if (!region) continue
      const key = regionKey(region.startRow, region.startCol)
      if (rowIndex === region.startRow) {
        const entry = entries.find((e) => e.col === col)
        if (entry) openMerges.set(key, entry.cellBox)
        continue
      }
      if (openMerges.has(key)) continue // covered by the anchor's tall box
      if (region.startRow >= minRowIndex) continue // shouldn't happen
      // opened on a previous page: draw the continuation without text
      let cellW = 0
      for (let i = col; i < col + span && i < colWidths.length; i++) cellW += colWidths[i]
      const cont: TableCellBox = {
        xPx: twipsToPx(colOffsets[col] ?? 0),
        yPx: yRel,
        widthPx: twipsToPx(cellW),
        heightPx: 0, // finalized when the region ends (or the chunk does)
        fill: region.fill,
        borders: region.borders,
        borderSpecs: region.borderSpecs
      }
      rowCells.push(cont)
      openMerges.set(key, cont)
    }

    for (const entry of entries) {
      const slack = rowH - entry.contentH
      let entryLines = entry.lines
      const anchorShift = !entry.project && slack > 0 && entry.vAlign !== 'top' ? (entry.vAlign === 'center' ? slack / 2 : slack) : 0
      if (!entry.project && slack > 0 && entry.vAlign !== 'top') {
        const shift = entry.vAlign === 'center' ? slack / 2 : slack
        entryLines = entryLines.map((l) => shiftLineY(l, shift))
      }
      const placedLines = entryLines.map((line) => shiftLineY(line, yRel))
      const placedAnchors = entry.anchors.map(anchor => ({ ...anchor, flowY: anchor.flowY + anchorShift + yRel }))
      floatingAnchors.push(...placedAnchors)
      lines.push(...placedLines)
      entry.cellBox.yPx = yRel
      entry.cellBox.heightPx = rowH
      rowCells.push(entry.cellBox)
      if (entry.project) finalDirectionLayouts.push({ box: entry.cellBox, start: lines.length - placedLines.length, count: placedLines.length, project: entry.project })
      const region = regions.get(regionKey(rowIndex, entry.col))
      if (region && region.endRow > rowIndex && entry.vAlign !== 'top' && !entry.project) {
        mergedAlignments.push({
          box: entry.cellBox,
          lines: placedLines,
          anchors: placedAnchors,
          initialHeight: rowH,
          factor: entry.vAlign === 'center' ? 0.5 : 1
        })
      }
    }
    committedRows.push({ yPx: yRel, heightPx: rowH, cells: rowCells })
    yRel += rowH
    // close merge regions whose last row this was
    for (const [key, box] of [...openMerges]) {
      const region = [...regions.values()].find((r) => regionKey(r.startRow, r.startCol) === key)
      if (region && region.endRow === rowIndex) {
        box.heightPx = yRel - box.yPx
        openMerges.delete(key)
      }
    }
    if (counts) {
      consumedRows++
      contentPlaced++
    }
  }
  // regions still open reach the bottom of this chunk (they continue on the
  // next page, where a continuation box is emitted for them)
  for (const box of openMerges.values()) {
    box.heightPx = yRel - box.yPx
  }
  // A merge grows after its anchor row is placed. Align content against the
  // final height of this page's merged box, including its inline pictures.
  for (const entry of mergedAlignments) {
    const shift = (entry.box.heightPx - entry.initialHeight) * entry.factor
    for (const anchor of entry.anchors) anchor.flowY += shift
    for (const line of entry.lines) {
      line.yPx += shift
      if (line.baselinePx !== undefined) line.baselinePx += shift
      for (const image of line.inlineImages ?? []) image.yPx += shift
      for (const seg of line.segs) if (seg.transform) seg.transform = { ...seg.transform, f: seg.transform.f + shift }
    }
  }
  // Replace from the end so changed wrap counts do not disturb earlier slots.
  for (const entry of finalDirectionLayouts.reverse()) {
    const final = entry.project(entry.box.heightPx).map(line => shiftLineY(line, entry.box.yPx))
    lines.splice(entry.start, entry.count, ...final)
  }

  return { lines, rows: committedRows, floatingAnchors, widthPx, heightPx: yRel, consumedRows }
}

function prefixSum(widths: number[]): number[] {
  const out = new Array<number>(widths.length + 1)
  out[0] = 0
  for (let i = 0; i < widths.length; i++) out[i + 1] = out[i] + widths[i]
  return out
}
