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

/** One w:lvl definition from numbering.xml. */
interface NumberingLevel {
  numFmt: string
  /** e.g. "%1." or "\u2022" */
  lvlText: string
  start: number
  indentLeftTwips?: number
  hangingTwips?: number
}

interface NumberingState {
  /** numId -> abstractNumId */
  numToAbstract: Map<string, string>
  /** abstractNumId -> level index -> level */
  abstracts: Map<string, NumberingLevel[]>
  /** numId -> level index -> current counter */
  counters: Map<string, number[]>
}

function parseNumbering(root: XmlNode): NumberingState {
  const abstracts = new Map<string, NumberingLevel[]>()
  for (const abs of getChildren(root, 'abstractNum')) {
    const a = attrs(abs)
    const id = a.abstractNumId as string | undefined
    if (!id) continue
    const levels: NumberingLevel[] = []
    for (const lvl of getChildren(abs, 'lvl')) {
      const la = attrs(lvl)
      const lvlText = attrs(getChildren(lvl, 'lvlText')[0]).val as string | undefined
      const startVal = attrs(getChildren(lvl, 'start')[0]).val
      const ind = attrs(getChildren(getChildren(lvl, 'pPr')[0], 'ind')[0])
      levels[Number(la.ilvl ?? levels.length)] = {
        numFmt: (attrs(getChildren(lvl, 'numFmt')[0]).val as string) ?? 'decimal',
        lvlText: lvlText ?? '%1.',
        start: startVal !== undefined ? parseInt(startVal, 10) || 1 : 1,
        indentLeftTwips: twips(ind.left ?? ind.start),
        hangingTwips: twips(ind.hanging),
      }
    }
    abstracts.set(id, levels)
  }
  const numToAbstract = new Map<string, string>()
  for (const num of getChildren(root, 'num')) {
    const a = attrs(num)
    const numId = a.numId as string | undefined
    const absId = attrs(getChildren(num, 'abstractNumId')[0]).val as string | undefined
    if (numId && absId) numToAbstract.set(numId, absId)
  }
  return { numToAbstract, abstracts, counters: new Map() }
}

const ROMAN: [number, string][] = [
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

function formatCounter(n: number, numFmt: string): string {
  switch (numFmt) {
    case 'decimal': return String(n)
    case 'lowerLetter': return String.fromCharCode(96 + ((n - 1) % 26) + 1)
    case 'upperLetter': return String.fromCharCode(64 + ((n - 1) % 26) + 1)
    case 'lowerRoman': return toRoman(n)
    case 'upperRoman': return toRoman(n).toUpperCase()
    case 'none': return ''
    default: return String(n)
  }
}

/** Fill %1..%9 in lvlText using each referenced level's own numFmt. */
function renderLvlText(lvlText: string, counters: number[], levels: NumberingLevel[] | undefined): string {
  return lvlText.replace(/%([1-9])/g, (_, d: string) => {
    const idx = Number(d) - 1
    const fmt = levels?.[idx]?.numFmt ?? 'decimal'
    return formatCounter(counters[idx] ?? 1, fmt)
  })
}

/**
 * Resolve a paragraph's w:numPr into a marker string, advancing the list
 * counters. Counters live on the parse-time state so a marker is stable even
 * if layout re-runs (e.g. table chunking).
 */
function resolveListMarker(pPr: XmlNode | undefined, state: NumberingState | undefined): { marker?: string; level?: number; indentTwips?: number } {
  if (!pPr || !state) return {}
  const numPr = getChildren(pPr, 'numPr')[0]
  if (!numPr) return {}
  const numId = attrs(getChildren(numPr, 'numId')[0]).val as string | undefined
  if (!numId) return {}
  const level = Number(attrs(getChildren(numPr, 'ilvl')[0]).val ?? 0) || 0
  const abstractId = state.numToAbstract.get(numId)
  const levels = abstractId ? state.abstracts.get(abstractId) : undefined
  const lvl = levels?.[level]
  if (!lvl) return {}
  let counters = state.counters.get(numId)
  if (!counters) {
    counters = []
    state.counters.set(numId, counters)
  }
  // deeper levels restart whenever this level ticks
  counters[level] = (counters[level] ?? lvl.start - 1) + 1
  for (let deeper = level + 1; deeper < (levels?.length ?? 0); deeper++) counters[deeper] = undefined as unknown as number
  const marker = lvl.numFmt === 'bullet' ? lvl.lvlText : renderLvlText(lvl.lvlText, counters, levels)
  return {
    marker: marker || undefined,
    level,
    indentTwips: lvl.indentLeftTwips,
  }
}

export function parseParagraph(p: XmlNode, images?: DocxImage[], numbering?: NumberingState): DocxParagraph {
  const pPr = getChildren(p, 'pPr')[0]
  const paragraph: DocxParagraph = {
    runs: [],
    images: [],
    align: alignOf(pPr),
  }
  const list = resolveListMarker(pPr, numbering)
  if (list.marker) {
    paragraph.listMarker = list.marker
    paragraph.listLevel = list.level
    // numbering levels carry their own indent; it wins over the paragraph's
    if (list.indentTwips !== undefined && paragraph.indentLeftTwips === undefined) {
      paragraph.indentLeftTwips = list.indentTwips
    }
  }
  if (pPr) {
    const ind = attrs(pPr['ind'] as XmlNode | undefined)
    // fall back to whatever the numbering level already supplied
    paragraph.indentLeftTwips = twips(ind.left) ?? twips(ind.start) ?? paragraph.indentLeftTwips
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
): Promise<{
  header?: DocxParagraph[]
  footer?: DocxParagraph[]
  firstHeader?: DocxParagraph[]
  firstFooter?: DocxParagraph[]
}> {
  const out: { header?: DocxParagraph[]; footer?: DocxParagraph[]; firstHeader?: DocxParagraph[]; firstFooter?: DocxParagraph[] } = {}
  for (const [kind, tag] of [['header', 'headerReference'], ['footer', 'footerReference']] as const) {
    const refs = getChildren(sectPr, tag)
    if (refs.length === 0) continue
    // w:type="first" overrides only the section's first page
    const firstRef = refs.find((r) => attrs(r).type === 'first')
    if (firstRef) {
      const first = await loadPart(firstRef, rels, pkg)
      if (first) out[kind === 'header' ? 'firstHeader' : 'firstFooter'] = first
    }
    // prefer the default type, else the first non-first reference
    const chosen = refs.find((r) => attrs(r).type === 'default') ?? refs.find((r) => attrs(r).type !== 'first') ?? refs[0]
    if (!chosen || attrs(chosen).type === 'first') {
      // only a first-page reference exists
      if (!firstRef) continue
      const only = await loadPart(firstRef, rels, pkg)
      if (only) out[kind] = only
      continue
    }
    const paragraphs = await loadPart(chosen, rels, pkg)
    if (paragraphs) out[kind] = paragraphs
  }
  return out
}

/** Read a header/footer part referenced by rId into paragraphs. */
async function loadPart(
  ref: XmlNode,
  rels: Map<string, { type: string; target: string }>,
  pkg: OfficePackage,
): Promise<DocxParagraph[] | undefined> {
  const rid = (attrs(ref)['r:id'] ?? attrs(ref).id) as string | undefined
  if (!rid) return undefined
  const rel = rels.get(rid)
  if (!rel) return undefined
  const path = rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target.replace(/^\.\.\//, '')}`
  const part = await pkg.xml(path)
  if (!part) return undefined
  const paragraphs: DocxParagraph[] = []
  // lists in a header/footer start from their own counters
  const partNumbering = await pkg.xml('word/numbering.xml')
  const partState = partNumbering ? parseNumbering(partNumbering) : undefined
  for (const [name, node] of elementChildren(part)) {
    if (name === 'p') paragraphs.push(parseParagraph(node, undefined, partState))
  }
  return paragraphs.length > 0 ? paragraphs : undefined
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

/**
 * Word wraps content in structured document tags (`w:sdt`) far more often than
 * fixtures suggest — around whole table rows, individual cells, and body-level
 * paragraphs/tables. Content inside `w:sdtContent` must still be collected, so
 * these loops descend through the wrapper.
 */
function unwrapContentControls(node: XmlNode | undefined): Array<[string, XmlNode]> {
  const out: Array<[string, XmlNode]> = []
  if (!node) return out
  for (const [name, child] of elementChildren(node)) {
    if (name === 'sdt') {
      out.push(...unwrapContentControls(getChildren(child, 'sdtContent')[0]))
    } else {
      out.push([name, child])
    }
  }
  return out
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
  const numberingRoot = await pkg.xml('word/numbering.xml')
  const numbering = numberingRoot ? parseNumbering(numberingRoot) : undefined
  const sections: DocxSection[] = []
  let current: DocxSection = {
    margins: { topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, headerTwips: 720, footerTwips: 720, gutterTwips: 0 },
    pageSize: { widthTwips: 12240, heightTwips: 15840, orientation: 'portrait' },
    paragraphs: [],
    blocks: [],
  }
  const push = () => { if (current.paragraphs.length > 0 || sections.length === 0) sections.push(current) }
  for (const [name, node] of unwrapContentControls(body as XmlNode)) {
    if (name === 'p') {
      const para = parseParagraph(node, docImages, numbering)
      current.paragraphs.push(para)
      current.blocks.push({ kind: 'p', paragraph: para })
    } else if (name === 'tbl' && node) {
      const table = parseTable(node, numbering)
      const block: DocxBlock = { kind: 'table', table }
      current.blocks.push(block)
    } else if (name === 'sectPr') {
      // section properties at body level — finalize current section
      const hf = await loadHeaderFooter(node, docRels, pkg)
      if (hf.header) current.header = hf.header
      if (hf.footer) current.footer = hf.footer
      if (hf.firstHeader) current.firstHeader = hf.firstHeader
      if (hf.firstFooter) current.firstFooter = hf.firstFooter
      if (node['titlePg'] !== undefined) current.titlePg = true
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
  // a body containing only a table (or only images) is still real content
  if (current.blocks.length > 0 || current.paragraphs.length > 0) sections.push(current)
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

function parseTableCell(tc: XmlNode, cellNumbering?: NumberingState): DocxTableCell {
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
    const tcW = getChildren(tcPr, 'tcW')[0]
    if (tcW) {
      const wa = attrs(tcW)
      // only dxa carries a usable width (pct/auto need the page width)
      if ((wa.type as string | undefined) === 'dxa' || wa.type === undefined) {
        cell.widthTwips = twips(wa.w)
      }
    }
    const vAlign = attrs(tcPr['vAlign'] as XmlNode | undefined).val as string | undefined
    if (vAlign === 'center' || vAlign === 'bottom' || vAlign === 'top') cell.vAlign = vAlign
  }
  for (const [name, node] of unwrapContentControls(tc)) {
    if (name === 'p') cell.paragraphs.push(parseParagraph(node, undefined, cellNumbering))
  }
  return cell
}

export function parseTable(tbl: XmlNode, tableNumbering?: NumberingState): DocxTable {
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
  for (const [name, node] of unwrapContentControls(tbl)) {
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
    for (const child of unwrapContentControls(node)) {
      if (child[0] === 'tc' && child[1]) row.cells.push(parseTableCell(child[1], tableNumbering))
    }
    table.rows.push(row)
  }
  // Word may omit w:tblGrid entirely (tblW type=auto). Fall back to the cells'
  // own w:tcW values, then to an even split, so those tables still lay out.
  if (table.gridColsTwips.length === 0) {
    const maxCols = table.rows.reduce((max, r) => {
      let n = 0
      for (const c of r.cells) n += Math.max(1, c.gridSpan)
      return Math.max(max, n)
    }, 0)
    if (maxCols > 0) {
      const widths = new Array<number>(maxCols).fill(0)
      const seen = new Array<boolean>(maxCols).fill(false)
      for (const row of table.rows) {
        let ci = 0
        for (const cell of row.cells) {
          const span = Math.max(1, cell.gridSpan)
          if (cell.widthTwips !== undefined) {
            // spread the declared width across the spanned columns
            const each = cell.widthTwips / span
            for (let k = 0; k < span && ci + k < maxCols; k++) {
              widths[ci + k] = Math.max(widths[ci + k], each)
              seen[ci + k] = true
            }
          }
          ci += span
        }
      }
      const anyDeclared = seen.some(Boolean)
      if (anyDeclared) {
        for (let i = 0; i < maxCols; i++) if (!seen[i]) widths[i] = 2880 // 2in default
      } else {
        for (let i = 0; i < maxCols; i++) widths[i] = Math.floor(9360 / maxCols)
      }
      table.gridColsTwips = widths.map((w) => Math.round(w))
    }
  }
  return table
}

export type { TableCellBorder }
