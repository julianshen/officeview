/**
 * Parse ODT (OpenDocument text) into the shared DocxDocument flow model.
 *
 * ODF keeps formatting in styles (resolved by ../odf/styles) while the
 * content holds structure; this adapter maps that structure onto paragraphs,
 * tables and images the existing layout already knows how to paginate and
 * paint. Mixed content order (text interleaved with spans) is preserved via
 * parseXmlOrdered — the default parser joins those segments and would
 * scramble run order.
 */
import type { OfficePackage } from '../core/zip'
import { sniffImageMime } from '../core/images'
import { attrs, getChildren, orderedChildren, textOf, type XmlNode } from '../core/xml'
import type { DocxBlock, DocxDocument, DocxParagraph, DocxSection, DocxTable, DocxTableCell, DocxTableRow, DocxTextRun } from '../docx/types'
import { lengthEmu } from '../odf/units'
import { OdfStyles, type OdfListStyle, type OdfTextProps } from '../odf/styles'
import type { OdtImageEntry, OdtListCounters } from './types'

const ROMAN: Array<[number, string]> = [
  [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'],
  [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i'],
]

function toRoman(n: number): string {
  let out = ''
  for (const [value, sym] of ROMAN) {
    while (n >= value) {
      out += sym
      n -= value
    }
  }
  return out
}

function formatOdtCounter(n: number, numFormat: string): string {
  switch (numFormat) {
    case 'decimal': return String(n)
    case 'lowerLetter': return String.fromCharCode(96 + ((n - 1) % 26) + 1)
    case 'upperLetter': return String.fromCharCode(64 + ((n - 1) % 26) + 1)
    case 'lowerRoman': return toRoman(n)
    case 'upperRoman': return toRoman(n).toUpperCase()
    case 'bullet': return ''
    default: return String(n)
  }
}

interface OdtContext {
  styles: OdfStyles
  images: OdtImageEntry[]
  /** Numbering snapshots per list style, for continue-numbering resume. */
  listStates: OdtListCounters
  /** Outline (heading) numbering never restarts within a document. */
  outlineCounters: number[]
}

/** LibreOffice-ish fallback page: A4 with 2cm-class margins (twips). */
const FALLBACK_PAGE = {
  widthTwips: 11906,
  heightTwips: 16838,
  margins: { topTwips: 1134, rightTwips: 1134, bottomTwips: 1134, leftTwips: 1134, headerTwips: 720, footerTwips: 720, gutterTwips: 0 },
}

/** style:master-page-name lives on the paragraph element itself (page breaks
 * almost always use this form) or, rarely, on its paragraph style. */
function paraMaster(node: XmlNode, styles: OdfStyles): string | undefined {
  const direct = attrs(node)['master-page-name'] as string | undefined
  if (direct) return direct
  return styles.paraProps(attrs(node)['style-name'] as string | undefined).masterPageName
}

function blankSection(): DocxSection {
  return {
    margins: { ...FALLBACK_PAGE.margins },
    pageSize: { widthTwips: FALLBACK_PAGE.widthTwips, heightTwips: FALLBACK_PAGE.heightTwips, orientation: 'portrait' },
    paragraphs: [],
    blocks: [],
  }
}

function runFromProps(text: string, props: OdfTextProps): DocxTextRun {
  const run: DocxTextRun = { text }
  if (props.bold !== undefined) run.bold = props.bold
  if (props.italic !== undefined) run.italic = props.italic
  if (props.underline !== undefined) run.underline = props.underline
  if (props.strike !== undefined) run.strike = props.strike
  if (props.fontFamily) run.fontFamily = props.fontFamily
  if (props.fontSizePt !== undefined) run.fontSizePt = props.fontSizePt
  if (props.color) run.color = props.color
  // the renderer resolves named highlights and passes the rest through as
  // CSS, so an ODF background lands as a hex fill directly
  if (props.background) run.highlight = `#${props.background}`
  return run
}

function isBlankSegment(text: string): boolean {
  return text.length > 0 && text.trim().length === 0 && text.includes('\n')
}

interface InlineOut {
  runs: DocxTextRun[]
  images: DocxParagraph['images']
}

/** Walk mixed paragraph content in document order (spans, fields, frames…). */
function walkInline(children: Array<[string, XmlNode]>, inherited: OdfTextProps, ctx: OdtContext, out: InlineOut): void {
  let pendingBreak = false
  const pushRun = (text: string, props: OdfTextProps): void => {
    if (text.length === 0 && !pendingBreak) return
    const run = runFromProps(text, props)
    if (pendingBreak) {
      run.breakBefore = true
      pendingBreak = false
    }
    if (run.text.length === 0) return
    out.runs.push(run)
  }
  const walk = (items: Array<[string, XmlNode]>, props: OdfTextProps): void => {
    for (const [name, child] of items) {
      if (name === '#text') {
        const text = textOf(child)
        if (!isBlankSegment(text)) pushRun(text, props)
        continue
      }
      if (name === 'span') {
        const merged = { ...props, ...ctx.styles.textProps(attrs(child)['style-name'] as string | undefined, props.fontSizePt) }
        walk(orderedChildren(child), merged)
        continue
      }
      if (name === 'a') {
        // hyperlinks keep their text; the target has no model equivalent
        walk(orderedChildren(child), props)
        continue
      }
      if (name === 's') {
        const extra = parseInt((attrs(child).c as string | undefined) ?? '1', 10)
        const count = Number.isFinite(extra) && extra > 0 ? extra : 1
        pushRun(' '.repeat(count), props)
        continue
      }
      if (name === 'tab') {
        pushRun('\t', props)
        continue
      }
      if (name === 'line-break') {
        pendingBreak = true
        continue
      }
      if (name === 'page-number') {
        const run = runFromProps(textOf(child), props)
        run.field = 'PAGE'
        out.runs.push(run)
        continue
      }
      if (name === 'page-count') {
        const run = runFromProps(textOf(child), props)
        run.field = 'NUMPAGES'
        out.runs.push(run)
        continue
      }
      if (name === 'frame') {
        const image = parseFrame(child, ctx)
        if (image) out.images.push(image)
        continue
      }
      if (name === 'annotation' || name === 'deletion') continue // dropped content
      if (name === 'insertion' || name === 'change') {
        walk(orderedChildren(child), props) // accept-all semantics
        continue
      }
      if (name === 'change-start' || name === 'change-end') continue
      // Unknown inline: recurse into text:* (keeps ruby text etc.), drop
      // drawing/object subtrees whose text is not flow content.
      if (name === 'text-box' || name === 'object' || name === 'object-ole' || name === 'plugin' || name === 'applet' || name === 'page-thumbnail' || name === 'custom-shape' || name === 'rect' || name === 'line' || name === 'ellipse' || name === 'caption') continue
      walk(orderedChildren(child), props)
    }
  }
  walk(children, inherited)
}

function matchImage(images: OdtImageEntry[], href: string): OdtImageEntry | undefined {
  let decoded = href
  try {
    decoded = decodeURIComponent(href)
  } catch {
    // keep the raw href
  }
  const base = (p: string): string => p.split('/').pop() ?? p
  return (
    images.find((img) => img.path === decoded)
    ?? images.find((img) => base(img.path) === base(decoded))
  )
}

/**
 * draw:frame → an image (inline for as-char anchors) or a floating anchor.
 * Frames without image content (text boxes, shapes, OLE objects) have no
 * model equivalent and are skipped.
 */
function parseFrame(frame: XmlNode, ctx: OdtContext): DocxParagraph['images'][number] | undefined {
  const a = attrs(frame)
  const imageNode = getChildren(frame, 'image')[0]
  if (!imageNode) return undefined
  const href = (attrs(imageNode).href as string | undefined) ?? ''
  const entry = matchImage(ctx.images, href)
  if (!entry) return undefined
  const widthEmu = lengthEmu(a.width as string | undefined)
  const heightEmu = lengthEmu(a.height as string | undefined)
  if (widthEmu === undefined || heightEmu === undefined || widthEmu <= 0 || heightEmu <= 0) return undefined
  const base = { data: entry.data, mime: entry.mime, widthEmu, heightEmu }
  const anchor = (a['anchor-type'] as string | undefined) ?? 'as-char'
  if (anchor === 'as-char') return base
  const graphic = ctx.styles.graphicProps(a['style-name'] as string | undefined)
  const zIndex = parseInt((a['z-index'] as string | undefined) ?? '', 10)
  const hOff = lengthEmu(a.x as string | undefined) ?? 0
  const vOff = lengthEmu(a.y as string | undefined) ?? 0
  const relativeFrom = anchor === 'page' ? 'page' : 'paragraph'
  const posH = graphic.hPos === 'center'
    ? { relativeFrom, offsetEmu: 0, align: 'center' as const }
    : graphic.hPos === 'right'
      ? { relativeFrom, offsetEmu: 0, align: 'right' as const }
      : { relativeFrom, offsetEmu: hOff }
  const posV = graphic.vPos === 'center' || graphic.vPos === 'bottom'
    ? { relativeFrom, offsetEmu: 0, align: graphic.vPos as 'center' | 'bottom' }
    : { relativeFrom, offsetEmu: vOff }
  const wrap = graphic.wrap === 'none' ? 'none' : graphic.wrap === 'through' ? 'through' : 'square'
  return {
    ...base,
    floating: {
      behindDoc: graphic.behindDoc,
      relativeHeight: Number.isFinite(zIndex) ? zIndex : 0,
      wrap,
      posH,
      posV,
    },
  }
}

function markerFor(style: OdfListStyle, level: number, counters: number[]): string | undefined {
  const lvl = style.levels[level]
  if (!lvl) return undefined
  if (lvl.numFormat === 'bullet') return lvl.bulletChar ?? '•'
  const shown = Math.max(1, Math.min(lvl.displayLevels, level + 1))
  const parts: string[] = []
  for (let i = level - shown + 1; i <= level; i++) {
    const ancestor = style.levels[i]
    parts.push(formatOdtCounter(counters[i] ?? 1, ancestor?.numFormat ?? 'decimal'))
  }
  return `${lvl.prefix}${parts.join('.')}${lvl.suffix}`
}

/** Advance counters for one list item; deeper levels restart. Returns the marker. */
function nextMarker(style: OdfListStyle, level: number, counters: number[], startValue?: number): string | undefined {
  const lvl = style.levels[level]
  const start = startValue ?? lvl?.start ?? 1
  counters[level] = (counters[level] ?? start - 1) + 1
  for (let deeper = level + 1; deeper < 10; deeper++) counters[deeper] = undefined as unknown as number
  return markerFor(style, level, counters)
}

interface InheritedList {
  styleName?: string
  /** The enclosing list's live counters, for continue-numbering into the same style. */
  active?: { key: string; counters: number[] }
}

function parseOdtParagraph(node: XmlNode, ctx: OdtContext, opts?: { outlineLevel?: number }): DocxParagraph {
  const a = attrs(node)
  const resolved = ctx.styles.paraProps(a['style-name'] as string | undefined)
  const paragraph: DocxParagraph = { runs: [], images: [], align: resolved.align ?? 'left' }
  if (resolved.indentLeftTwips !== undefined) paragraph.indentLeftTwips = resolved.indentLeftTwips
  if (resolved.indentRightTwips !== undefined) paragraph.indentRightTwips = resolved.indentRightTwips
  if (resolved.firstLineTwips !== undefined) paragraph.indentFirstLineTwips = resolved.firstLineTwips
  if (resolved.spaceBeforeTwips !== undefined) paragraph.spacingBeforeTwips = resolved.spaceBeforeTwips
  if (resolved.spaceAfterTwips !== undefined) paragraph.spacingAfterTwips = resolved.spaceAfterTwips
  if (resolved.lineSpacing) paragraph.lineSpacing = resolved.lineSpacing
  if (opts?.outlineLevel !== undefined) paragraph.outlineLevel = opts.outlineLevel
  const basePt = resolved.text.fontSizePt ?? ctx.styles.defaultFontSizePt
  const inherited: OdfTextProps = {
    ...resolved.text,
    fontFamily: resolved.text.fontFamily ?? ctx.styles.defaultFontFamily,
    fontSizePt: basePt,
  }
  const out: InlineOut = { runs: [], images: [] }
  walkInline(orderedChildren(node), inherited, ctx, out)
  paragraph.runs = out.runs
  paragraph.images = out.images
  return paragraph
}

/** A text:list / text:continue-list → marked paragraphs (nested lists deepen).
 *
 * Each list element owns its numbering sequence: entering a nested list must
 * not disturb the parent's counters, so sequences live in per-element arrays
 * (snapshotted for later continue-numbering) rather than one shared entry
 * per style. A nested list without its own style inherits the surrounding one.
 */
function parseOdtList(
  node: XmlNode,
  ctx: OdtContext,
  level: number,
  blocks: DocxBlock[],
  paragraphs: DocxParagraph[],
  inherited?: InheritedList,
): void {
  const a = attrs(node)
  const ownStyle = a['style-name'] as string | undefined
  const styleName = ownStyle ?? inherited?.styleName
  const style = ctx.styles.listStyle(styleName)
  const continuing = (a['continue-numbering'] as string | undefined)?.toLowerCase() === 'true'
  let counters: number[]
  let key: string | undefined
  if (styleName) {
    key = styleName
    if (continuing && inherited?.active?.key === styleName) {
      counters = inherited.active.counters
    } else if (continuing) {
      const saved = ctx.listStates.get(styleName)
      counters = saved ? [...saved] : []
    } else if (inherited?.active) {
      // Fresh nested sequence: preserve the ancestors' counters so
      // display-level joins (e.g. "2.1.") see their context, but reset our
      // own level and deeper so numbering restarts here. The copy keeps the
      // parent's live array untouched.
      counters = [...inherited.active.counters]
      for (let i = level; i < 10; i++) counters[i] = undefined as unknown as number
    } else {
      counters = []
    }
  } else {
    counters = []
  }
  const active = key ? { key, counters } : inherited?.active
  const startValue = parseInt((a['start-value'] as string | undefined) ?? '', 10)
  const start = Number.isFinite(startValue) ? startValue : undefined
  for (const [name, child] of orderedChildren(node)) {
    if (name !== 'list-item' && name !== 'list-header') continue
    let marked = false
    for (const [itemName, itemNode] of orderedChildren(child)) {
      if (itemName === 'p' || itemName === 'h') {
        const outline = itemName === 'h' ? parseInt((attrs(itemNode)['outline-level'] as string | undefined) ?? '', 10) : undefined
        const para = parseOdtParagraph(itemNode, ctx, outline !== undefined && Number.isFinite(outline) ? { outlineLevel: outline } : undefined)
        if (!marked && style) {
          const marker = nextMarker(style, Math.min(level, 9), counters, marked ? undefined : start)
          if (marker) {
            para.listMarker = marker
            para.listLevel = Math.min(level, 9)
            const indent = style.levels[Math.min(level, 9)]?.indentLeftTwips
            if (indent !== undefined && para.indentLeftTwips === undefined) para.indentLeftTwips = indent
            const first = style.levels[Math.min(level, 9)]?.firstLineTwips
            if (first !== undefined && para.indentFirstLineTwips === undefined) para.indentFirstLineTwips = first
          }
          marked = true
        } else if (!marked) {
          marked = true
        }
        paragraphs.push(para)
        blocks.push({ kind: 'p', paragraph: para })
      } else if (itemName === 'list' || itemName === 'continue-list') {
        parseOdtList(itemNode, ctx, level + 1, blocks, paragraphs, { styleName, active })
      }
    }
  }
  if (key) ctx.listStates.set(key, counters)
}

/** Flow blocks (paragraphs, headings, lists, tables) in document order. */
function parseBlocks(nodes: Array<[string, XmlNode]>, ctx: OdtContext, blocks: DocxBlock[], paragraphs: DocxParagraph[]): void {
  for (const [name, node] of nodes) {
    if (name === 'p') {
      const para = parseOdtParagraph(node, ctx)
      paragraphs.push(para)
      blocks.push({ kind: 'p', paragraph: para })
    } else if (name === 'h') {
      const outline = parseInt((attrs(node)['outline-level'] as string | undefined) ?? '', 10)
      const para = parseOdtParagraph(node, ctx, Number.isFinite(outline) ? { outlineLevel: outline } : undefined)
      const outlineStyle = ctx.styles.listStyle('\0outline')
      if (outlineStyle && Number.isFinite(outline)) {
        const marker = nextMarker(outlineStyle, Math.min(Math.max(outline - 1, 0), 9), ctx.outlineCounters)
        if (marker) {
          para.listMarker = marker
          para.listLevel = Math.min(Math.max(outline - 1, 0), 9)
        }
      }
      paragraphs.push(para)
      blocks.push({ kind: 'p', paragraph: para })
    } else if (name === 'list' || name === 'continue-list') {
      parseOdtList(node, ctx, 0, blocks, paragraphs)
    } else if (name === 'table') {
      const table = parseOdtTable(node, ctx)
      if (table) blocks.push({ kind: 'table', table })
    } else if (name === 'section') {
      // text:section is a region wrapper, not a page section — flatten it
      parseBlocks(orderedChildren(node), ctx, blocks, paragraphs)
    } else if (name === 'table-of-content' || name === 'illustration-index' || name === 'table-index' || name === 'object-index' || name === 'user-index' || name === 'alphabetical-index' || name === 'bibliography') {
      // index bodies duplicate headed content; the entries are real visible
      // text (like Word's TOC), so include them as plain blocks
      for (const [indexName, indexNode] of orderedChildren(node)) {
        if (indexName === 'index-body') parseBlocks(orderedChildren(indexNode), ctx, blocks, paragraphs)
      }
    }
    // soft-page-break, change markers at block level, metadata: ignored
  }
}

function emptyCell(vMerge?: 'restart' | 'continue'): DocxTableCell {
  return {
    paragraphs: [{ runs: [], images: [], align: 'left' }],
    gridSpan: 1,
    vMerge,
  }
}

/**
 * Bounds for table:number-*-repeated. Repetition is faithful content, so it
 * expands — but an unbounded count is a memory-exhaustion vector in a hostile
 * file, hence the cap (documented, generous: real templates repeat dozens).
 */
const MAX_TABLE_REPEAT = 1024

function cappedRepeat(raw: number): number {
  if (!Number.isFinite(raw) || raw < 1) return 1
  return Math.min(Math.floor(raw), MAX_TABLE_REPEAT)
}

/** table:table → grid model (spans, covered cells, fills, borders, heights). */
function parseOdtTable(node: XmlNode, ctx: OdtContext): DocxTable | undefined {
  const a = attrs(node)
  const tprops = ctx.styles.tableProps(a['style-name'] as string | undefined)
  // column track: absolute widths stick, relative weights share the mean
  const colAbs: Array<number | undefined> = []
  const readColumns = (parent: XmlNode): void => {
    for (const [name, child] of orderedChildren(parent)) {
      if (name === 'table-column-group') {
        readColumns(child)
        continue
      }
      if (name !== 'table-column') continue
      const ca = attrs(child)
      const repeat = cappedRepeat(parseInt((ca['number-columns-repeated'] as string | undefined) ?? '1', 10))
      const props = ctx.styles.columnProps(ca['style-name'] as string | undefined)
      for (let i = 0; i < repeat; i++) colAbs.push(props.widthTwips)
    }
  }
  readColumns(node)
  const defined = colAbs.filter((w) => w !== undefined)
  const absMean = defined.length > 0 ? defined.reduce((sum, w) => sum + (w ?? 0), 0) / defined.length : 1440
  let gridColsTwips = colAbs.map((w) => Math.round(w ?? absMean))
  const rows: DocxTableRow[] = []
  let firstCellPadding: { top?: number; right?: number; bottom?: number; left?: number } | undefined
  const parseCellContent = (cellNode: XmlNode): DocxParagraph[] => {
    const paras: DocxParagraph[] = []
    for (const [contentName, contentNode] of orderedChildren(cellNode)) {
      if (contentName === 'p' || contentName === 'h') paras.push(parseOdtParagraph(contentNode, ctx))
      else if (contentName === 'list' || contentName === 'continue-list') {
        const cellBlocks: DocxBlock[] = []
        parseOdtList(contentNode, ctx, 0, cellBlocks, paras)
        void cellBlocks
      }
      // nested tables: out of scope, skipped
    }
    if (paras.length === 0) paras.push({ runs: [], images: [], align: 'left' })
    return paras
  }
  const parseRowCells = (rowNode: XmlNode): DocxTableCell[] => {
    const cells: DocxTableCell[] = []
    // Positions already covered by a horizontal span in this row: a covered
    // placeholder there extends the span, it is not a grid column of its
    // own. Covered cells outside such spans are vertical leftovers and keep
    // their vMerge-continue entry (including combined row+column spans,
    // whose lower rows arrive as their own covered cells).
    let hCover = 0
    for (const [cellName, cellNode] of orderedChildren(rowNode)) {
      if (cellName === 'covered-table-cell') {
        // Covered placeholders repeat like cells; each logical placeholder
        // consumes one horizontal span position before counting as vertical.
        const coverRepeat = cappedRepeat(parseInt((attrs(cellNode)['number-columns-repeated'] as string | undefined) ?? '1', 10))
        for (let i = 0; i < coverRepeat; i++) {
          if (hCover > 0) {
            hCover -= 1
            continue
          }
          cells.push(emptyCell('continue'))
        }
        continue
      }
      if (cellName !== 'table-cell') continue
      const cellA = attrs(cellNode)
      // number-columns-repeated stamps identical cells; content re-parses per
      // copy so list counters advance exactly as if written out
      const cellRepeat = cappedRepeat(parseInt((cellA['number-columns-repeated'] as string | undefined) ?? '1', 10))
      for (let c = 0; c < cellRepeat; c++) {
        const cprops = ctx.styles.cellProps(cellA['style-name'] as string | undefined)
        if (!firstCellPadding && cprops.paddingTwips) firstCellPadding = cprops.paddingTwips
        const span = Math.max(1, parseInt((cellA['number-columns-spanned'] as string | undefined) ?? '1', 10) || 1)
        const rowSpan = Math.max(1, parseInt((cellA['number-rows-spanned'] as string | undefined) ?? '1', 10) || 1)
        cells.push({
          paragraphs: parseCellContent(cellNode),
          gridSpan: span,
          vMerge: rowSpan > 1 ? 'restart' : undefined,
          fill: cprops.fill,
          borders: cprops.borders,
          vAlign: cprops.vAlign,
        })
        if (span > 1) hCover += span - 1
      }
    }
    return cells
  }
  const readRows = (parent: XmlNode, header: boolean): void => {
    for (const [name, child] of orderedChildren(parent)) {
      if (name === 'table-row-group' || name === 'table-header-rows' || name === 'table-rows') {
        readRows(child, header || name === 'table-header-rows')
        continue
      }
      if (name !== 'table-row') continue
      const ra = attrs(child)
      const repeat = cappedRepeat(parseInt((ra['number-rows-repeated'] as string | undefined) ?? '1', 10))
      const rprops = ctx.styles.rowProps(ra['style-name'] as string | undefined)
      // Repeated rows re-parse from the element, so list counters and fields
      // advance per copy exactly as if the rows were written out.
      for (let r = 0; r < repeat; r++) {
        rows.push({
          cells: parseRowCells(child),
          heightTwips: rprops.heightTwips,
          heightRule: rprops.heightRule,
          isHeader: header || undefined,
        })
      }
    }
  }
  readRows(node, false)
  if (rows.length === 0) return undefined
  if (gridColsTwips.length === 0) {
    // no column definitions: split a default gutter across the widest row
    const widest = Math.max(...rows.map((r) => r.cells.reduce((n, c) => n + c.gridSpan, 0)))
    gridColsTwips = new Array(Math.max(widest, 1)).fill(1440)
  }
  return {
    gridColsTwips,
    rows,
    fill: tprops.fill,
    borders: tprops.borders,
    cellMargins: {
      topTwips: firstCellPadding?.top ?? 0,
      rightTwips: firstCellPadding?.right ?? 0,
      bottomTwips: firstCellPadding?.bottom ?? 0,
      leftTwips: firstCellPadding?.left ?? 0,
    },
  }
}

function masterBlocks(nodes: Array<[string, XmlNode]>, ctx: OdtContext): DocxParagraph[] {
  const blocks: DocxBlock[] = []
  const paragraphs: DocxParagraph[] = []
  // headers/footers hold paragraphs (lists and tables are out of scope here)
  for (const [name, node] of nodes) {
    if (name !== 'p' && name !== 'h') continue
    const para = parseOdtParagraph(node, ctx)
    paragraphs.push(para)
    blocks.push({ kind: 'p', paragraph: para })
  }
  void blocks
  return paragraphs
}

export async function parseOdt(pkg: OfficePackage): Promise<DocxDocument> {
  const styles = await OdfStyles.load(pkg)
  const content = await pkg.xmlOrdered('content.xml')
  if (!content) throw new Error('content.xml missing — not a valid ODT?')
  const ctx: OdtContext = { styles, images: await loadOdtImages(pkg), listStates: new Map(), outlineCounters: [] }
  const body = getChildren(content, 'body')[0]
  const text = body ? getChildren(body, 'text')[0] : undefined
  const sections: DocxSection[] = []
  let current = blankSection()
  let currentMaster: string | undefined
  const push = (): void => {
    if (current.paragraphs.length > 0 || current.blocks.length > 0 || sections.length === 0) {
      sections.push(current)
    }
  }
  const applyMaster = (masterName: string | undefined): void => {
    // An ordinary paragraph carries no master: it inherits the active one
    // rather than resetting to the fallback layout. Only an explicit change
    // seals the current section — this keeps a custom first master (plus its
    // header and geometry) across the paragraphs that follow it.
    if (masterName === undefined || masterName === currentMaster) return
    if (current.paragraphs.length > 0 || current.blocks.length > 0) {
      // content accumulated under another master: seal it and start fresh.
      // An empty current section is reused instead so a leading master does
      // not leave a spurious empty section behind.
      push()
      current = blankSection()
    }
    currentMaster = masterName
    const master = styles.masterPage(masterName)
    const layout = styles.pageLayout(master?.layoutName)
    if (layout) {
      current.pageSize = { widthTwips: layout.widthTwips, heightTwips: layout.heightTwips, orientation: layout.orientation }
      current.margins = {
        ...current.margins,
        topTwips: layout.margins.topTwips,
        rightTwips: layout.margins.rightTwips,
        bottomTwips: layout.margins.bottomTwips,
        leftTwips: layout.margins.leftTwips,
      }
      if (layout.headerTwips !== undefined) current.margins.headerTwips = layout.headerTwips
      if (layout.footerTwips !== undefined) current.margins.footerTwips = layout.footerTwips
    }
    if (master) {
      if (master.header.length > 0) {
        const header = masterBlocks(master.header, ctx)
        if (header.length > 0) current.header = header
      }
      if (master.footer.length > 0) {
        const footer = masterBlocks(master.footer, ctx)
        if (footer.length > 0) current.footer = footer
      }
    }
  }
  if (text) {
    const items = orderedChildren(text)
    for (const [name, node] of items) {
      if (name === 'p' || name === 'h') {
        applyMaster(paraMaster(node, styles))
      }
      parseBlocks([[name, node]], ctx, current.blocks, current.paragraphs)
    }
  }
  push()
  return {
    sections,
    defaultFontFamily: styles.defaultFontFamily,
    defaultFontSizePt: styles.defaultFontSizePt,
    styleDefaults: new Map(),
  }
}

async function loadOdtImages(pkg: OfficePackage): Promise<OdtImageEntry[]> {
  const out: OdtImageEntry[] = []
  const paths = pkg.paths((p) => p.startsWith('Pictures/')).sort()
  for (const path of paths) {
    const data = await pkg.bytes(path)
    if (!data) continue
    let decoded = path
    try {
      decoded = decodeURIComponent(path)
    } catch {
      // keep the raw path
    }
    out.push({ path: decoded, data, mime: sniffImageMime(data) })
  }
  return out
}
