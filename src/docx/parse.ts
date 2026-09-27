/**
 * Parse DOCX parts (document.xml, styles.xml) into the DocxDocument model.
 */
import type { OfficePackage } from '../core/zip'
import { sniffImageMime } from '../core/images'
import { attrs, elementChildren, getChildren, textOf, type XmlNode } from '../core/xml'
import type { DocxBlock, DocxDocument, DocxImage, DocxParagraph, DocxSection, DocxTable, DocxTableCell, DocxTableCellMargins, DocxTableBorders, DocxTableRow, DocxTextRun, ParagraphAlign } from './types'

function alignOf(pPr: XmlNode | undefined): ParagraphAlign {
  const jc = pPr ? getChildren(pPr, 'jc')[0] : undefined
  const v = attrs(jc).val as string | undefined
  switch (v) {
    case 'center': return 'center'
    case 'right': return 'right'
    case 'both': return 'justify'
    default: return 'left'
  }
}

export function twips(v: string | number | undefined): number | undefined {
  if (v === undefined) return undefined
  const n = typeof v === 'number' ? v : parseFloat(v)
  return Number.isFinite(n) ? n : undefined
}

export function halfPointToPt(v: string | number | undefined): number | undefined {
  if (v === undefined) return undefined
  const n = typeof v === 'number' ? v : parseFloat(v)
  return Number.isFinite(n) ? n / 2 : undefined
}

function boolAttr(v: string | undefined): boolean {
  return v === '1' || v === 'true'
}

interface FieldAwareRun extends DocxTextRun {
  _fldChar?: string
  _instr?: string
}

function parseRun(r: XmlNode, inherited?: Partial<DocxTextRun>): FieldAwareRun {
  const rPr = getChildren(r, 'rPr')[0]
  let run: FieldAwareRun = { text: '', ...inherited }
  const fld = getChildren(r, 'fldChar')[0]
  if (fld) run._fldChar = attrs(fld).fldCharType as string
  const instr = getChildren(r, 'instrText')[0]
  if (instr) run._instr = textOf(instr).trim()
  if (rPr) {
    const rFonts = rPr['rFonts'] ? getChildren(rPr, 'rFonts')[0] : undefined
    const fam = attrs(rFonts).ascii as string | undefined
    if (fam) run.fontFamily = fam
    const sz = attrs(rPr['sz'] as XmlNode | undefined).val
    const szPt = halfPointToPt(sz)
    if (szPt !== undefined) run.fontSizePt = szPt
    if (rPr['b'] !== undefined) run.bold = !boolAttr(attrs(rPr['b'] as XmlNode).val)
    if (rPr['i'] !== undefined) run.italic = !boolAttr(attrs(rPr['i'] as XmlNode).val)
    if (rPr['u'] !== undefined) {
      const uv = attrs(rPr['u'] as XmlNode).val
      run.underline = uv !== 'none'
    }
    if (rPr['strike'] !== undefined) run.strike = !boolAttr(attrs(rPr['strike'] as XmlNode).val)
    const color = attrs(rPr['color'] as XmlNode | undefined).val as string | undefined
    if (color) run.color = color === 'auto' ? undefined : color
    const highlight = attrs(rPr['highlight'] as XmlNode | undefined).val as string | undefined
    if (highlight) run.highlight = highlight
  }
  // Runs may contain text fragments plus tabs/breaks
  for (const [name, child] of elementChildren(r)) {
    if (name === 't') {
      run.text += textOf(child)
    } else if (name === 'tab') {
      run.text += '\t'
    } else if (name === 'br') {
      run.breakBefore = true
    }
  }
  return run
}

export function parseParagraph(p: XmlNode, images?: DocxImage[]): DocxParagraph {
  const pPr = getChildren(p, 'pPr')[0]
  const paragraph: DocxParagraph = {
    runs: [],
    images: [],
    align: alignOf(pPr),
  }
  if (pPr) {
    const ind = attrs(pPr['ind'] as XmlNode | undefined)
    paragraph.indentLeftTwips = twips(ind.left) ?? twips(ind.start)
    paragraph.indentRightTwips = twips(ind.right) ?? twips(ind.end)
    paragraph.indentFirstLineTwips = twips(ind.firstLine)
    const spacing = attrs(pPr['spacing'] as XmlNode | undefined)
    paragraph.spacingBeforeTwips = twips(spacing.before)
    paragraph.spacingAfterTwips = twips(spacing.after)
    const line = twips(spacing.line)
    const lineRule = spacing.lineRule as string | undefined
    if (line !== undefined) {
      paragraph.lineSpacing = { rule: (lineRule as 'auto' | 'exact' | 'atLeast') ?? 'auto', value: line }
    }
    const outline = attrs(pPr['outlineLvl'] as XmlNode | undefined).val
    if (outline !== undefined) paragraph.outlineLevel = parseInt(outline, 10)
  }
  // Runs and hyperlinks (treat hyperlink content as runs too)
  for (const child of elementChildren(p)) {
    const [name, node] = child
    if (name === 'r') {
      paragraph.runs.push(parseRun(node))
      for (const [iname, inode] of elementChildren(node)) {
        if (iname === 'drawing' && inode) {
          const image = parseDrawing(inode, images)
          if (image) paragraph.images.push(image)
        }
      }
    } else if (name === 'hyperlink') {
      for (const [iname, inode] of elementChildren(node)) {
        if (iname === 'r' && inode) {
          paragraph.runs.push(parseRun(inode))
        }
      }
    } else if (name === 'drawing' && node) {
      const image = parseDrawing(node, images)
      if (image) paragraph.images.push(image)
    }
    // other children (bookmarks, proofErr, etc.) ignored
  }
  resolveFields(paragraph)
  return paragraph
}

/**
 * Collapse w:fldChar/w:instrText field sequences into runs tagged with
 * `field` (e.g. PAGE, NUMPAGES), keeping the cached result text.
 *   begin -> instrText " PAGE " -> separate -> "1" -> end
 */
function resolveFields(paragraph: DocxParagraph): void {
  const runs = paragraph.runs as FieldAwareRun[]
  let pendingInstr: string | null = null
  let inField = false
  let inResult = false
  const out: DocxTextRun[] = []
  for (const run of runs) {
    if (run._fldChar === 'begin') {
      inField = true
      inResult = false
      pendingInstr = null
      continue
    }
    if (inField && run._instr !== undefined) {
      pendingInstr = run._instr.toUpperCase().split(/\s+/)[0]
      continue
    }
    if (run._fldChar === 'separate') {
      inResult = true
      continue
    }
    if (run._fldChar === 'end') {
      inField = false
      inResult = false
      pendingInstr = null
      continue
    }
    if (inField) {
      if (inResult && pendingInstr) {
        // cached field result: mark the run so paint can substitute the value
        const { _fldChar, _instr, ...rest } = run
        void _fldChar
        void _instr
        out.push({ ...rest, field: pendingInstr })
        pendingInstr = null
      }
      continue
    }
    const { _fldChar, _instr, ...rest } = run
    void _fldChar
    void _instr
    out.push(rest)
  }
  paragraph.runs = out
}

/** w:drawing -> wp:inline|wp:anchor -> a:graphic -> pic:pic -> a:blip r:embed */
function parseDrawing(drawing: XmlNode, images: DocxImage[] | undefined): DocxImage | undefined {
  const wp = getChildren(drawing, 'inline')[0] ?? getChildren(drawing, 'anchor')[0]
  if (!wp) return undefined
  const extent = getChildren(wp, 'extent')[0]
  const ea = attrs(extent)
  const widthEmu = parseFloat(ea.cx as string) || 0
  const heightEmu = parseFloat(ea.cy as string) || 0
  const graphic = getChildren(wp, 'graphic')[0]
  // pic:pic lives inside a:graphicData, not directly under a:graphic
  const graphicData = graphic ? getChildren(graphic, 'graphicData')[0] : undefined
  const pic = graphicData ? getChildren(graphicData, 'pic')[0] : undefined
  const blip = pic ? getChildren(getChildren(pic, 'blipFill')[0], 'blip')[0] : undefined
  const rid = (attrs(blip)['r:embed'] ?? attrs(blip).embed) as string | undefined
  if (!rid) return undefined
  if (!images) return undefined
  const data = images.find((img) => img && (img as DocxImage & { relId?: string }).relId === rid)
  if (!data) return undefined
  return { data: data.data, mime: data.mime, widthEmu, heightEmu }
}

/** Resolve word/_rels/document.xml.rels into embedded image bytes keyed by rId. */
/** rId -> part path for word/document.xml.rels. */
async function loadDocRels(pkg: OfficePackage): Promise<Map<string, { type: string; target: string }>> {
  const out = new Map<string, { type: string; target: string }>()
  const rels = await pkg.xml('word/_rels/document.xml.rels')
  if (!rels) return out
  for (const rel of getChildren(rels, 'Relationship')) {
    const a = attrs(rel)
    if (a.Id && a.Target) {
      out.set(a.Id, { type: (a.Type as string) ?? '', target: a.Target })
    }
  }
  return out
}

/** Resolve a sectPr's header/footer references to paragraph lists. */
async function loadHeaderFooter(
  sectPr: XmlNode,
  rels: Map<string, { type: string; target: string }>,
  pkg: OfficePackage,
): Promise<{ header?: DocxParagraph[]; footer?: DocxParagraph[] }> {
  const out: { header?: DocxParagraph[]; footer?: DocxParagraph[] } = {}
  for (const [kind, tag] of [['header', 'headerReference'], ['footer', 'footerReference']] as const) {
    const refs = getChildren(sectPr, tag)
    if (refs.length === 0) continue
    // prefer the default type, else the first reference
    const chosen = refs.find((r) => attrs(r).type === 'default') ?? refs[0]
    const rid = (attrs(chosen)['r:id'] ?? attrs(chosen).id) as string | undefined
    if (!rid) continue
    const rel = rels.get(rid)
    if (!rel) continue
    const path = rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target.replace(/^\.\.\//, '')}`
    const part = await pkg.xml(path)
    if (!part) continue
    const paragraphs: DocxParagraph[] = []
    for (const [name, node] of elementChildren(part)) {
      if (name === 'p') paragraphs.push(parseParagraph(node))
    }
    if (paragraphs.length > 0) out[kind] = paragraphs
  }
  return out
}

async function loadDocImages(pkg: OfficePackage): Promise<DocxImage[]> {
  const rels = await pkg.xml('word/_rels/document.xml.rels')
  if (!rels) return []
  const images: Array<DocxImage & { relId?: string }> = []
  for (const rel of getChildren(rels, 'Relationship')) {
    const a = attrs(rel)
    const type = a.Type as string | undefined
    const target = (a.Target as string | undefined) ?? ''
    if (!type || !type.includes('/image') || !target) continue
    const path = target.startsWith('/') ? target.slice(1) : `word/${target.replace(/^\.\.\//, '')}`
    const data = await pkg.bytes(path)
    if (!data) continue
    images.push({ data, mime: sniffImageMime(data), widthEmu: 0, heightEmu: 0, relId: a.Id })
  }
  return images
}

export async function parseDocx(pkg: OfficePackage): Promise<DocxDocument> {
  const doc = await pkg.xml('word/document.xml')
  if (!doc) throw new Error('word/document.xml missing — not a valid docx?')
  const stylesXml = await pkg.xml('word/styles.xml')
  let defaultFontFamily = 'Calibri'
  let defaultFontSizePt = 11
  const styleDefaults = new Map<string, { fontFamily?: string; fontSizePt?: number }>()
  if (stylesXml) {
    const docDefaults = stylesXml['docDefaults']
    if (docDefaults && typeof docDefaults === 'object') {
      const rPrDefault = getChildren(docDefaults as XmlNode, 'rPrDefault')[0]
      const rPr = rPrDefault ? getChildren(rPrDefault, 'rPr')[0] : undefined
      const rFonts = rPr ? getChildren(rPr, 'rFonts')[0] : undefined
      const fam = attrs(rFonts).ascii as string | undefined
      const sz = attrs(rPr?.['sz'] as XmlNode | undefined).val
      const szPt = halfPointToPt(sz)
      if (fam) {
        defaultFontFamily = fam
        styleDefaults.set('__default__', { fontFamily: fam, fontSizePt: szPt })
      }
      if (szPt !== undefined) defaultFontSizePt = szPt
    }
  }
  const body = doc?.['body'] as XmlNode | undefined
  const docImages = (await loadDocImages(pkg)) as DocxImage[]
  const docRels = await loadDocRels(pkg)
  const sections: DocxSection[] = []
  let current: DocxSection = {
    margins: { topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, headerTwips: 720, footerTwips: 720, gutterTwips: 0 },
    pageSize: { widthTwips: 12240, heightTwips: 15840, orientation: 'portrait' },
    paragraphs: [],
    blocks: [],
  }
  const push = () => { if (current.paragraphs.length > 0 || sections.length === 0) sections.push(current) }
  for (const [name, node] of elementChildren(body as XmlNode)) {
    if (name === 'p') {
      const para = parseParagraph(node, docImages)
      current.paragraphs.push(para)
      current.blocks.push({ kind: 'p', paragraph: para })
    } else if (name === 'tbl' && node) {
      const table = parseTable(node)
      const block: DocxBlock = { kind: 'table', table }
      current.blocks.push(block)
    } else if (name === 'sectPr') {
      // section properties at body level — finalize current section
      const hf = await loadHeaderFooter(node, docRels, pkg)
      if (hf.header) current.header = hf.header
      if (hf.footer) current.footer = hf.footer
      const pgMar = getChildren(node, 'pgMar')[0]
      const pgSz = getChildren(node, 'pgSz')[0]
      if (pgMar) {
        const a = attrs(pgMar)
        current.margins = {
          topTwips: twips(a.top) ?? current.margins.topTwips,
          rightTwips: twips(a.right) ?? current.margins.rightTwips,
          bottomTwips: twips(a.bottom) ?? current.margins.bottomTwips,
          leftTwips: twips(a.left) ?? current.margins.leftTwips,
          headerTwips: twips(a.header) ?? current.margins.headerTwips,
          footerTwips: twips(a.footer) ?? current.margins.footerTwips,
          gutterTwips: twips(a.gutter) ?? 0,
        }
      }
      if (pgSz) {
        const a = attrs(pgSz)
        const w = twips(a.w) ?? current.pageSize.widthTwips
        const h = twips(a.h) ?? current.pageSize.heightTwips
        const orient = (a.orient as 'landscape' | 'portrait') ?? 'portrait'
        current.pageSize = { widthTwips: w, heightTwips: h, orientation: orient }
      }
      const carried = {
        header: current.header,
        footer: current.footer,
      }
      push()
      current = {
        margins: current.margins,
        pageSize: current.pageSize,
        paragraphs: [],
        blocks: [],
        ...carried,
      }
    }
  }
  if (current.paragraphs.length > 0) sections.push(current)
  return {
    sections,
    defaultFontFamily,
    defaultFontSizePt,
    styleDefaults,
  }
}

type TableCellBorder = DocxTableBorders['top']

function parseSideBorder(node: XmlNode | undefined): TableCellBorder {
  if (!node) return undefined
  const a = attrs(node)
  if (a.val === 'nil' || a.val === 'none') return undefined
  return { style: a.val as string, color: a.color as string | undefined }
}

function parseBorders(parent: XmlNode): DocxTableBorders {
  const el = getChildren(parent, 'tblBorders')[0] ?? getChildren(parent, 'tcBorders')[0]
  if (!el) return {}
  return {
    top: parseSideBorder(getChildren(el, 'top')[0]),
    bottom: parseSideBorder(getChildren(el, 'bottom')[0]),
    left: parseSideBorder(getChildren(el, 'left')[0]),
    right: parseSideBorder(getChildren(el, 'right')[0]),
    insideH: parseSideBorder(getChildren(el, 'insideH')[0]),
    insideV: parseSideBorder(getChildren(el, 'insideV')[0]),
  }
}

const DEFAULT_CELL_MARGINS: DocxTableCellMargins = { topTwips: 0, rightTwips: 108, bottomTwips: 0, leftTwips: 108 }

function parseCellMargins(tblPr: XmlNode): DocxTableCellMargins {
  const mar = getChildren(tblPr, 'tblCellMar')[0]
  if (!mar) return DEFAULT_CELL_MARGINS
  const side = (n: string, dflt: number): number => {
    const node = getChildren(mar, n)[0]
    return node ? (twips(attrs(node).w) ?? dflt) : dflt
  }
  return {
    topTwips: side('top', 0),
    rightTwips: side('right', 108),
    bottomTwips: side('bottom', 0),
    leftTwips: side('left', 108),
  }
}

function parseTableCell(tc: XmlNode): DocxTableCell {
  const tcPr = getChildren(tc, 'tcPr')[0]
  const cell: DocxTableCell = { paragraphs: [], gridSpan: 1 }
  if (tcPr) {
    const a = attrs(tcPr['gridSpan'] as XmlNode | undefined)
    if (a.val !== undefined) cell.gridSpan = parseInt(a.val as string, 10) || 1
    const vMerge = attrs(tcPr['vMerge'] as XmlNode | undefined).val as string | undefined
    if (vMerge === 'restart') cell.vMerge = 'restart'
    else if (vMerge !== undefined || tcPr['vMerge'] !== undefined) cell.vMerge = 'continue'
    const shd = getChildren(tcPr, 'shd')[0]
    const shdAttrs = attrs(shd)
    if (shd && shdAttrs.val !== 'nil') cell.fill = shdAttrs.fill as string | undefined
    cell.borders = parseBorders(tcPr)
    const vAlign = attrs(tcPr['vAlign'] as XmlNode | undefined).val as string | undefined
    if (vAlign === 'center' || vAlign === 'bottom' || vAlign === 'top') cell.vAlign = vAlign
  }
  for (const [name, node] of elementChildren(tc)) {
    if (name === 'p') cell.paragraphs.push(parseParagraph(node))
  }
  return cell
}

export function parseTable(tbl: XmlNode): DocxTable {
  const tblPr = getChildren(tbl, 'tblPr')[0]
  const table: DocxTable = {
    gridColsTwips: [],
    rows: [],
    cellMargins: tblPr ? parseCellMargins(tblPr) : DEFAULT_CELL_MARGINS,
    borders: tblPr ? parseBorders(tblPr) : undefined,
  }
  if (tblPr) {
    const shd = getChildren(tblPr, 'shd')[0]
    const shdAttrs = attrs(shd)
    if (shd && shdAttrs.val !== 'nil') table.fill = shdAttrs.fill as string | undefined
  }
  const grid = getChildren(tbl, 'tblGrid')[0]
  if (grid) {
    for (const col of getChildren(grid, 'gridCol')) {
      table.gridColsTwips.push(twips(attrs(col).w) ?? 0)
    }
  }
  for (const [name, node] of elementChildren(tbl)) {
    if (name !== 'tr' || !node) continue
    const trPr = getChildren(node, 'trPr')[0]
    const row: DocxTableRow = { cells: [] }
    if (trPr) {
      const trHeight = getChildren(trPr, 'trHeight')[0]
      if (trHeight) {
        const a = attrs(trHeight)
        row.heightTwips = twips(a.val)
        const rule = a.hRule as string | undefined
        if (rule === 'exact' || rule === 'atLeast' || rule === 'auto') row.heightRule = rule
      }
      if (trPr['tblHeader'] !== undefined) row.isHeader = true
    }
    for (const child of elementChildren(node)) {
      if (child[0] === 'tc' && child[1]) row.cells.push(parseTableCell(child[1]))
    }
    table.rows.push(row)
  }
  return table
}

export type { TableCellBorder }
