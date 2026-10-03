/**
 * DOCX layout + paint. Coordinates are CSS px (96dpi): 1pt = 96/72 px.
 * Text measurement uses a real Canvas2D measureText (via a provided measure
 * function) so widths match what we paint.
 */
import type { DocxDocument, DocxImage, DocxParagraph, DocxSection, DocxTable, DocxTableCell, DocxTableRow, DocxTextRun, TableCellBorder } from './types'
import { paintDrawing } from './drawing'
import { fontFamilyCss } from './styles'
import { emuToPx } from '../core/geometry'
import { twipsToPx } from '../core/geometry'
import { resolveColor } from '../core/color'
import { paintWatermark, type ResolvedWatermark, type WatermarkOptions } from '../core/watermark'

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
  /** List marker drawn in the left gutter (first line only). */
  marker?: { text: string; widthPx: number }
  inlineImages?: Array<{ image: DocxImage; xPx: number; yPx: number; widthPx: number; heightPx: number }>
  baselinePx?: number
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
}

/** Page header or footer content, positioned in page-relative px. */
export interface HFBlock {
  paragraphs: DocxParagraph[]
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
  /** Set for floating (wp:anchor) images: drawn behind text and out of flow. */
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
  parts.push(fontFamilyCss(s.fontFamily))
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
    const boxes = inlineImages.map((i) => ({
      image: i.image,
      xPx: lineX + lineIndent + alignOffset + i.offset,
      yPx: y + (h - (i.height + i.bottom)),
      widthPx: i.width,
      heightPx: i.height
    }))
    lines.push({
      yPx: y,
      xPx: lineX + lineIndent,
      widthPx: width,
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
    segs.push({ text, run, style, widthPx: w })
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
          if (width + w <= lineUsable) {
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
      const paragraphs =
        block.kind === 'p'
          ? [block.paragraph]
          : block.table.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs))
      for (const paragraph of paragraphs) {
        for (const image of paragraph.images ?? []) {
          if (!seen.has(image)) {
            seen.add(image)
            out.push(image)
          }
        }
      }
    }
  }
  for (const section of document.sections) {
    for (const paragraphs of [section.header, section.footer, section.firstHeader, section.firstFooter]) {
      for (const paragraph of paragraphs ?? []) {
        for (const image of paragraph.images) {
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
    const pagesBefore = pages.length
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
        // Floating images anchor to this paragraph: place them first so the
        // paragraph's own text (and everything after it) can wrap around them.
        // They stay on the page where the anchor starts, even if the text
        // itself flows onward; paragraph/line-relative anchors use this y.
        for (const image of para.images ?? []) {
          if (!image.floating) continue
          const w = emuToPx(image.widthEmu)
          const h = emuToPx(image.heightEmu)
          if (w <= 0 || h <= 0) continue
          const place = resolveFloating(image, { pageWidthPx: widthPx, contentX, contentWidth, flowY: y, m })
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
            shift = m.top - line.yPx
          }
          for (const box of line.inlineImages ?? [])
            page.images.push({
              ...box,
              yPx: box.yPx + shift,
              imageIndex: imageIndex.get(box.image) ?? -1,
              drawing: box.image.drawing
            })
          if (line.segs.length || !line.inlineImages?.length)
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
              imageIndex: imageIndex.get(box.image) ?? -1,
              drawing: box.image.drawing
            })
          if (line.segs.length || !line.inlineImages?.length)
            page.lines.push({
              ...line,
              xPx: line.xPx + tableX,
              yPx: line.yPx + y,
              baselinePx: line.baselinePx === undefined ? undefined : line.baselinePx + y
            })
        }
        page.tables.push({ xPx: tableX, yPx: y, widthPx: chunk.widthPx, rows: chunk.rows })
        y += chunk.heightPx
        fromRow += chunk.consumedRows
        firstChunk = false
        if (fromRow < table.rows.length) commitPage()
      }
    }
    if (page.lines.length > 0 || page.tables.length > 0 || page.images.length > 0 || pages.length === 0)
      pages.push(page)

    // attach this section's header/footer to every page it produced; with
    // w:titlePg the first page gets the type="first" variants instead
    if (section.header || section.footer || section.firstHeader || section.firstFooter) {
      const block = (paragraphs: DocxParagraph[] | undefined, yPx: number): HFBlock | undefined =>
        paragraphs ? { paragraphs, yPx, xPx: m.left, widthPx: contentWidth, defaults } : undefined
      const headerY = twipsToPx(section.margins.headerTwips)
      const footerY = heightPx - twipsToPx(section.margins.footerTwips)
      const make = (firstPage: boolean) => ({
        header: block(firstPage && section.titlePg ? (section.firstHeader ?? section.header) : section.header, headerY),
        footer: block(firstPage && section.titlePg ? (section.firstFooter ?? section.footer) : section.footer, footerY)
      })
      pages.slice(pagesBefore).forEach((p, i) => {
        const { header, footer } = make(i === 0)
        if (header) p.header = header
        if (footer) p.footer = footer
        for (const hf of [header, footer]) {
          if (!hf) continue
          let imageY = hf.yPx
          for (const paragraph of hf.paragraphs) {
            imageY += twipsToPx(paragraph.spacingBeforeTwips ?? 0)
            const laid = layoutParagraph(paragraph, measure, {
              contentX: hf.xPx,
              contentWidth: hf.widthPx,
              startY: imageY,
              defaults
            })
            imageY = laid.endY
            for (const line of laid.lines)
              for (const box of line.inlineImages ?? [])
                p.images.push({ ...box, imageIndex: imageIndex.get(box.image) ?? -1, drawing: box.image.drawing })
            for (const image of paragraph.images) {
              const width = emuToPx(image.widthEmu)
              const height = emuToPx(image.heightEmu)
              if (width <= 0 || height <= 0) continue
              if (image.floating) {
                const place = resolveFloating(image, {
                  pageWidthPx: widthPx,
                  contentX,
                  contentWidth,
                  flowY: imageY,
                  m
                })
                p.images.push({
                  xPx: place.x,
                  yPx: place.y,
                  widthPx: width,
                  heightPx: height,
                  imageIndex: imageIndex.get(image) ?? -1,
                  drawing: image.drawing,
                  floating: { ...image.floating }
                })
              }
            }
            imageY += twipsToPx(paragraph.spacingAfterTwips ?? 0)
          }
        }
      })
    }
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
export interface RenderPagesOptions {
  /** First page's number for PAGE field substitution (default 1). */
  pageNumberStart?: number
  /** Total for NUMPAGES substitution (default pages.length). */
  totalPages?: number
  /** Watermark stamped onto every page. Drawn under the content, not over it. */
  watermark?: WatermarkOptions | ResolvedWatermark
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
  const measure = createMeasurer(ctx)
  const firstNumber = options?.pageNumberStart ?? 1
  const totalPages = options?.totalPages ?? pages.length
  pages.forEach((page, pageIndex) => {
    // The watermark goes down first, straight onto the page background, so the
    // document's own text is drawn on top of it rather than being washed out.
    if (options?.watermark) {
      paintWatermark(ctx, { widthPx: page.widthPx, heightPx: page.heightPx }, options.watermark)
    }
    paintTables(page.tables, ctx)
    paintImages(page.images, ctx, images, 'behind')
    paintImages(page.images, ctx, images, 'front')
    const hfLines = layoutHeaderFooter(page, firstNumber + pageIndex, totalPages, measure)
    lastFont = '' // header measurement can change ctx.font independently of the paint cache
    const allLines = hfLines.length > 0 ? [...hfLines, ...page.lines] : page.lines
    for (const line of allLines) {
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
      const baseline = line.baselinePx ?? line.yPx + maxAscent
      // list marker sits left of the (indented) text
      if (line.marker) {
        const markerStyle: RunStyle = {
          fontFamily: seg0Font(line),
          fontSizePt: seg0Size(line),
          bold: !!line.segs[0]?.run.bold,
          italic: !!line.segs[0]?.run.italic,
        }
        const font = fontCss(markerStyle)
        if (font !== lastFont) {
          ctx.font = font
          lastFont = font
        }
        ctx.fillStyle = line.segs[0]?.run.color ? resolveColor(line.segs[0].run.color) : '#000000'
        ctx.fillText(line.marker.text.trimEnd(), line.xPx - line.marker.widthPx, baseline)
      }
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

/** Lay out a page's header and footer into positioned lines. */
function layoutHeaderFooter(
  page: PageLayout,
  pageNumber: number,
  totalPages: number,
  measure: MeasureFn,
): LineBox[] {
  const out: LineBox[] = []
  for (const block of [page.header, page.footer]) {
    if (!block) continue
    let y = block.yPx
    for (const para of substituteFields(block.paragraphs, pageNumber, totalPages)) {
      y += twipsToPx(para.spacingBeforeTwips ?? 0)
      const laid = layoutParagraph(para, measure, {
        contentX: block.xPx,
        contentWidth: block.widthPx,
        startY: y,
        defaults: block.defaults,
      })
      out.push(...laid.lines)
      y = laid.endY + twipsToPx(para.spacingAfterTwips ?? 0)
    }
  }
  return out
}

function seg0Font(line: LineBox): string {
  return line.segs[0]?.style.fontFamily ?? 'Calibri'
}

function seg0Size(line: LineBox): number {
  return line.segs[0]?.style.fontSizePt ?? 11
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
/**
 * Draw images. `pass` selects behind-document anchors (drawn before the text
 * layer) or everything else; the two calls bracket the text paint so a
 * w:behindDoc image sits under it.
 */
function paintImages(
  boxes: ImageBox[],
  ctx: CanvasRenderingContext2D,
  images: Array<CanvasImageSource | undefined> | undefined,
  pass: 'behind' | 'front'
): void {
  for (const box of boxes) {
    if (pass === 'behind' ? !box.floating?.behindDoc : box.floating?.behindDoc) continue
    if (box.drawing) {
      ctx.save()
      ctx.translate(box.xPx, box.yPx)
      paintDrawing(box.drawing, ctx, box.widthPx, box.heightPx)
      ctx.restore()
      continue
    }
    const img = images?.[box.imageIndex]
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
  ctx: { pageWidthPx: number; contentX: number; contentWidth: number; flowY: number; m: { top: number; left: number; right: number; bottom: number } },
): { x: number; y: number; width: number; height: number } {
  const pageWidth = ctx.pageWidthPx
  const pageHeight = pageWidth === 0 ? 0 : ctx.pageWidthPx
  void pageHeight
  const contentRight = ctx.contentX + ctx.contentWidth
  switch (relativeFrom) {
    case 'page':
      return { x: 0, y: 0, width: pageWidth, height: 0 }
    case 'margin':
      return { x: ctx.m.left, y: ctx.m.top, width: ctx.contentWidth, height: 0 }
    case 'paragraph':
    case 'line':
      return { x: ctx.contentX, y: ctx.flowY, width: ctx.contentWidth, height: 0 }
    default:
      // 'column' and anything else behave like the text column
      return { x: ctx.contentX, y: ctx.flowY, width: ctx.contentWidth, height: 0 }
  }
  void contentRight
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
  ctx: { pageWidthPx: number; contentX: number; contentWidth: number; flowY: number; m: { top: number; left: number; right: number; bottom: number } },
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
  if (alignV === 'center') y = oy.y
  else if (alignV === 'bottom') y = oy.y
  void h
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
    let col = 0
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
        fill: cell.fill ?? table.fill,
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
    let column = 0
    for (const cell of row.cells) {
      const span = Math.max(1, cell.gridSpan)
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
function layoutTableRows(
  table: DocxTable,
  measure: MeasureFn,
  defaults: { fontFamily: string; fontSizePt: number },
  opts: { fromRow: number; repeatHeader: boolean; maxHeightPx: number; allowFirstRowOverflow: boolean }
): { lines: LineBox[]; rows: TableRowBox[]; widthPx: number; heightPx: number; consumedRows: number } {
  const lines: LineBox[] = []
  const colOffsets = prefixSum(table.gridColsTwips)
  const colWidths = table.gridColsTwips
  const widthPx = twipsToPx(colOffsets[colOffsets.length - 1] ?? 0)
  const margins = {
    top: twipsToPx(table.cellMargins.topTwips),
    bottom: twipsToPx(table.cellMargins.bottomTwips),
    left: twipsToPx(table.cellMargins.leftTwips),
    right: twipsToPx(table.cellMargins.rightTwips)
  }

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
  const mergedAlignments: Array<{ box: TableCellBox; lines: LineBox[]; initialHeight: number; factor: number }> = []

  for (const { row, rowIndex, counts } of queue) {
    // lay the row into temporary coordinates (relative to its own top)
    const entries: Array<{
      cellBox: TableCellBox
      lines: LineBox[]
      contentH: number
      vAlign: 'top' | 'center' | 'bottom'
      col: number
    }> = []
    let rowContentH = 0
    for (const { cell, col, span } of positions[rowIndex]) {
      if (cell.vMerge === 'continue') {
        continue
      }
      let cellW = 0
      for (let i = col; i < col + span && i < colWidths.length; i++) cellW += colWidths[i]
      const cellXPx = twipsToPx(colOffsets[col] ?? 0)
      const innerW = twipsToPx(cellW) - margins.left - margins.right
      if (innerW <= 0) continue
      let cellLines: LineBox[] = []
      let cy = margins.top
      for (const para of cell.paragraphs) {
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
          ? cellLines[cellLines.length - 1].yPx + cellLines[cellLines.length - 1].heightPx + margins.bottom
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
          fill: cell.fill ?? table.fill,
          borders: {
            left: borderCss(spec('left')),
            right: borderCss(spec('right')),
            top: borderCss(spec('top')),
            bottom: borderCss(spec('bottom'))
          },
          borderSpecs: { left: spec('left'), right: spec('right'), top: spec('top'), bottom: spec('bottom') }
        },
        lines: cellLines,
        contentH,
        vAlign: cell.vAlign ?? 'top',
        col
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

    // stop before overflowing the available height, unless this is the first
    // content row of the chunk and the caller allows it
    const isFirstContent = counts && contentPlaced === 0
    if (!isFirstContent && yRel + rowH > opts.maxHeightPx) break
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
      if (slack > 0 && entry.vAlign !== 'top') {
        const shift = entry.vAlign === 'center' ? slack / 2 : slack
        entryLines = entryLines.map((l) => ({
          ...l,
          yPx: l.yPx + shift,
          baselinePx: l.baselinePx === undefined ? undefined : l.baselinePx + shift,
          inlineImages: l.inlineImages?.map((box) => ({ ...box, yPx: box.yPx + shift }))
        }))
      }
      const placedLines = entryLines.map((line) => ({
        ...line,
        yPx: line.yPx + yRel,
        baselinePx: line.baselinePx === undefined ? undefined : line.baselinePx + yRel,
        inlineImages: line.inlineImages?.map((box) => ({ ...box, yPx: box.yPx + yRel }))
      }))
      lines.push(...placedLines)
      entry.cellBox.yPx = yRel
      entry.cellBox.heightPx = rowH
      rowCells.push(entry.cellBox)
      const region = regions.get(regionKey(rowIndex, entry.col))
      if (region && region.endRow > rowIndex && entry.vAlign !== 'top') {
        mergedAlignments.push({
          box: entry.cellBox,
          lines: placedLines,
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
    for (const line of entry.lines) {
      line.yPx += shift
      if (line.baselinePx !== undefined) line.baselinePx += shift
      for (const image of line.inlineImages ?? []) image.yPx += shift
    }
  }

  return { lines, rows: committedRows, widthPx, heightPx: yRel, consumedRows }
}

function prefixSum(widths: number[]): number[] {
  const out = new Array<number>(widths.length + 1)
  out[0] = 0
  for (let i = 0; i < widths.length; i++) out[i + 1] = out[i] + widths[i]
  return out
}
