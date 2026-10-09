/**
 * Parse DOCX parts (document.xml, styles.xml) into the DocxDocument model.
 */
import type { OfficePackage } from '../core/zip'
import { sniffImageMime } from '../core/images'
import { attrs, elementChildren, getChildren, orderedChildren, textOf, type XmlNode } from '../core/xml'
import { applyParagraphDefaults, authoredCategories, issueCategory, paragraphRunDefaults, readRunProperties, readTheme, styleChain, styleContext, type DocxStyleContext, type ParagraphStyleLayers } from './styles'
import { loadDrawingParts, wordDrawingSelections } from './drawing'
import { findSvgBlip, looksLikeSvg, scanSelfContainedSvg, svgCandidate, svgStateCandidate, type ImageSelection, type SvgCandidate, type SvgVerdict } from '../core/svg'
import { attemptedMalformedRelationshipIssue, malformedRelationshipAttempt, reserveDrawingContent } from '../drawing/content'
import { contentRepresentation, coverageIssueMatchesEntry, supportedChoiceRequirements, type DrawingCoverageEntry } from '../drawing/coverage'
import { DOCUMENT_DRAWING_NODE_LIMIT, drawingPartContext, partRelationshipNodes, reserveDrawingNode } from '../drawing/parts'
import { parseVmlContainer, type VmlNode } from '../drawing/vml'
import type { DocxBlock, DocxDocument, DocxDrawing, DocxDrawingShape, DocxFloating, DocxImage, DocxParagraph, DocxSection, DocxTable, DocxTableCell, DocxTableCellMargins, DocxTableBorders, DocxTableRow, DocxTextRun, ParagraphAlign } from './types'

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

export interface VmlProvenance { nextPict: number }
interface ParagraphContext { styles: DocxStyleContext; drawings: Map<XmlNode, DocxImage>; tableLayers?: ParagraphStyleLayers; reserveDrawing?: (drawing?: DocxDrawing) => boolean; drawingCoverage?: DrawingCoverageEntry[]; drawingPaths?: WeakMap<XmlNode, string>; unselectedReferenceIds?: Set<string>; partPath?: string; pkg?: OfficePackage; representation?: DrawingCoverageEntry['representation']; imageSelections?: Map<string, ImageSelection>; imageExternal?: Set<string>; vmlProvenance?: VmlProvenance }
/** Share one selection record per candidate pair so aliases and coverage agree. */
function selectionFor(context: ParagraphContext | undefined, key: string): ImageSelection {
  const map = context?.imageSelections
  if (!map) return { phase: 'pending', representation: 'none' }
  let selection = map.get(key)
  if (!selection) { selection = { phase: 'pending', representation: 'none' }; map.set(key, selection) }
  return selection
}
function drawingReferenceIds(root: XmlNode): string[] {
  const ids: string[] = []
  const walk = (node: XmlNode, depth: number) => {
    if (depth >= 128) return
    for (const [name, child] of orderedChildren(node)) {
      if (['blip', 'chart', 'relIds', 'contentPart', 'imagedata'].includes(name)) {
        const a = attrs(child)
        for (const key of ['id', 'embed', 'dm', 'lo', 'qs', 'cs']) if (a[key]) ids.push(a[key])
      }
      if (name !== '#text') walk(child, depth + 1)
    }
  }
  walk(root, 0)
  return ids
}
function sourceDrawingPaths(root: XmlNode, rootName = 'document'): WeakMap<XmlNode, string> {
  const paths = new WeakMap<XmlNode, string>()
  const walk = (node: XmlNode, path: string, depth: number) => {
    if (depth >= 128) return
    for (const [index, [name, child]] of orderedChildren(node).entries()) {
      if (name === '#text') continue
      const childPath = `${path}/${name}[${index}]`
      if (name === 'drawing' || name === 'AlternateContent') paths.set(child, childPath)
      walk(child, childPath, depth + 1)
    }
  }
  walk(root, rootName, 0)
  return paths
}
function drawingReservation(pkg: OfficePackage, part: string): (drawing?: DocxDrawing) => boolean {
  return drawing => drawing ? reserveDrawingContent(pkg, drawing, part) : reserveDrawingNode(drawingPartContext(pkg), part)
}

interface FieldAwareRun extends DocxTextRun {
  _fldChar?: string
  _instr?: string
}

function parseRun(r: XmlNode, inherited?: Partial<DocxTextRun>, context?: ParagraphContext, issues?: import('../drawing/text-parse').TextAppearanceIssue[]): FieldAwareRun {
  const rPr = getChildren(r, 'rPr')[0]
  let run: FieldAwareRun = { text: '', ...inherited }
  const fld = getChildren(r, 'fldChar')[0]
  if (fld) run._fldChar = attrs(fld).fldCharType as string
  const instr = getChildren(r, 'instrText')[0]
  if (instr) run._instr = textOf(instr).trim()
  Object.assign(run, readRunProperties(rPr, context?.styles.theme, issues))
  // Runs may contain text fragments plus tabs/breaks
  for (const [name, child] of orderedChildren(r)) {
    if (name === 't') {
      run.text += textOf(child)
    } else if (name === 'tab') {
      run.text += '\t'
    } else if (name === 'br') {
      run.text += '\n'
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

export function parseParagraph(
  p: XmlNode,
  images?: DocxImage[],
  numbering?: NumberingState,
  context?: ParagraphContext
): DocxParagraph {
  const pPr = getChildren(p, 'pPr')[0]
  const paragraph: DocxParagraph = {
    runs: [],
    images: [],
    align: alignOf(pPr)
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
  const appearanceIssues: import('../drawing/text-parse').TextAppearanceIssue[] = []
  const inline: NonNullable<DocxParagraph['inline']> = []
  const textRunAuthored: Array<Set<import('./styles').AppearanceCategory>> = []
  const inheritedIssues: import('../drawing/text-parse').TextAppearanceIssue[] = []
  const inherited = paragraphRunDefaults(p, context?.styles, context?.tableLayers, inheritedIssues)
  paragraph.paragraphMark = { ...inherited, ...readRunProperties(getChildren(pPr, 'rPr')[0], context?.styles.theme, appearanceIssues) }
  const addRun = (run: FieldAwareRun) => {
    paragraph.runs.push(run)
    inline.push({ kind: 'text', run })
  }
  const addImage = (image: DocxImage | undefined) => {
    if (image) {
      paragraph.images.push(image)
      inline.push({ kind: 'image', image })
    }
  }
  const walk = (parent: XmlNode): void => {
    for (const [name, node] of orderedChildren(parent)) {
      if (name === 'r') {
        const rPr = getChildren(node, 'rPr')[0]
        const hasText = getChildren(node, 't').length > 0 || getChildren(node, 'tab').length > 0 || getChildren(node, 'br').length > 0
        if (hasText) {
          textRunAuthored.push(authoredCategories(rPr))
        }
        const run = parseRun(node, inherited, context, appearanceIssues)
        if (!node.drawing && !node.AlternateContent && !node.pict) addRun(run)
        else {
          // Split a mixed run at drawing boundaries without losing the styles.
          for (const [tag, child] of orderedChildren(node)) {
            if (tag === 't') addRun({ ...run, text: textOf(child), breakBefore: undefined })
            else if (tag === 'tab') addRun({ ...run, text: '\t', breakBefore: undefined })
            else if (tag === 'br') addRun({ ...run, text: '\n', breakBefore: undefined })
            else if (tag === 'drawing') addImage(parseDrawing(child, images, context))
            else if (tag === 'pict') addImage(parsePict(child, context))
            else if (tag === 'AlternateContent') alternate(child)
          }
        }
      } else if (name === 'hyperlink' || name === 'sdtContent') walk(node)
      else if (name === 'drawing') addImage(parseDrawing(node, images, context))
      else if (name === 'pict') addImage(parsePict(node, context))
      else if (name === 'AlternateContent') alternate(node)
      else if (name === 'sdt') {
        const content = getChildren(node, 'sdtContent')[0]
        if (content) walk(content)
      }
    }
  }
  const alternate = (node: XmlNode): void => {
    const skipped: XmlNode[] = []
    const coverageBefore = context?.drawingCoverage?.length ?? 0
    let attemptedChoice: XmlNode | undefined
    const markSkipped = () => { for (const choice of skipped) for (const id of drawingReferenceIds(choice)) context?.unselectedReferenceIds?.add(id) }
    const markUnselected = (selected: XmlNode) => {
      for (const branch of [...getChildren(node, 'Choice'), ...getChildren(node, 'Fallback')]) {
        if (branch === selected) continue
        for (const id of drawingReferenceIds(branch)) context?.unselectedReferenceIds?.add(id)
      }
    }
    const hasSelectedBlank = (parent: XmlNode | undefined, depth = 0): boolean => {
      if (!parent || depth >= 128) return false
      for (const [name, child] of orderedChildren(parent)) {
        if (name === 'drawing' && !context?.drawings.get(child)?.drawing && wordDrawingSelections(context?.drawings, child).some(selection => orderedChildren(selection.node).every(([childName]) => childName === '#text'))) return true
        if (name !== '#text' && hasSelectedBlank(child, depth + 1)) return true
      }
      return false
    }
    for (const choice of getChildren(node, 'Choice')) {
      const understood = supportedChoiceRequirements(choice, ['w', 'a', 'c', 'dgm', 'dsp', 'ink', 'wpi', 'wps', 'wpg', 'wpc', 'wp', 'pic']).supported
      if (!understood) { skipped.push(choice); continue }
      attemptedChoice ??= choice
      if (orderedChildren(choice).every(([name]) => name === '#text')) {
        markUnselected(choice)
        if (context?.reserveDrawing && context.drawingCoverage) context.drawingCoverage.push({
          partPath: context.partPath ?? 'word/document.xml', treePath: context.drawingPaths?.get(node) ?? `AlternateContent[${context.drawingCoverage.length}]`,
          element: 'AlternateContent', feature: 'empty-choice', status: 'native', selectedRepresentation: 'blank', representation: 'choice',
          reason: 'Supported empty Choice selected', scope: 'original',
        })
        return
      }
      // Probe without spending the document drawing budget. Word can supply
      // several Choices; only the first one that yields visible inline content
      // is selected and parsed again with placement reservations enabled.
      if (context?.reserveDrawing) {
        const reserve = context.reserveDrawing
        const inlineBefore = inline.length, runsBefore = paragraph.runs.length, imagesBefore = paragraph.images.length
        const hadProvenance = context ? 'vmlProvenance' in context && context.vmlProvenance !== undefined : false
        const nextPictBefore = context?.vmlProvenance?.nextPict
        const pkgDiagnostics = context.pkg ? drawingPartContext(context.pkg).diagnostics : undefined
        const pkgDiagBefore = pkgDiagnostics ? pkgDiagnostics.length : 0
        context.reserveDrawing = undefined
        let usable = false
        try {
          walk(choice)
          usable = inline.length > inlineBefore || hasSelectedBlank(choice)
        } finally {
          context.reserveDrawing = reserve
          inline.length = inlineBefore
          paragraph.runs.length = runsBefore
          paragraph.images.length = imagesBefore
          if (context) {
            if (!hadProvenance) {
              delete context.vmlProvenance
            } else if (context.vmlProvenance && nextPictBefore !== undefined) {
              context.vmlProvenance.nextPict = nextPictBefore
            }
          }
          if (pkgDiagnostics && pkgDiagnostics.length > pkgDiagBefore) {
            pkgDiagnostics.length = pkgDiagBefore
          }
        }
        if (!usable) { skipped.push(choice); continue }
      }
      const before = inline.length
      const prior = context?.representation
      if (context) context.representation = 'choice'
      try { walk(choice) } finally { if (context) context.representation = prior }
      if (inline.length > before || hasSelectedBlank(choice)) { markUnselected(choice); return }
      skipped.push(choice)
    }
    const fallback = getChildren(node, 'Fallback')[0]
    const fallbackImage = findDescendant(fallback, 'imagedata')
    const rid = attrs(fallbackImage).id
    const data = images?.find((img) => (img as DocxImage & { relId?: string }).relId === rid)
    const drawing = findDescendant(getChildren(node, 'Choice')[0], 'drawing')
    const wp = getChildren(drawing, 'anchor')[0]
    if (data && wp && (!context?.reserveDrawing || context.reserveDrawing())) {
      const extent = attrs(getChildren(wp, 'extent')[0])
      addImage({
        ...data,
        widthEmu: Number(extent.cx) || 0,
        heightEmu: Number(extent.cy) || 0,
        floating: parseAnchor(wp)
      })
      if (fallback) markUnselected(fallback)
      const docPr = attrs(getChildren(wp, 'docPr')[0])
      context?.drawingCoverage?.push({ partPath: context.partPath ?? 'word/document.xml', treePath: drawing ? context.drawingPaths?.get(drawing) ?? `drawing[${context.drawingCoverage.length}]` : `drawing[${context.drawingCoverage.length}]`, element: 'drawing', id: docPr.id, name: docPr.name,
        feature: 'picture', status: 'fallback', selectedRepresentation: 'raster-fallback', representation: 'fallback', reason: 'Selected compatibility fallback image', scope: 'original' })
    } else if (fallback) {
      const prior = context?.representation, before = inline.length
      if (context) context.representation = 'fallback'
      try { walk(fallback) } finally { if (context) context.representation = prior }
      if (inline.length > before) markUnselected(fallback)
    }
    // A failed compatibility wrapper still owns a source placement. Probing
    // disables audit writes, so retain its identity when no branch yielded a
    // selected drawing or fallback; the failed relationship diagnostic can
    // then attach to this object instead of becoming an orphan entry.
    if (context?.reserveDrawing && context.drawingCoverage?.length === coverageBefore) {
      const candidate = attemptedChoice ?? getChildren(node, 'Choice')[0]
      const sourceDrawings: XmlNode[] = []
      const collect = (parent: XmlNode | undefined, depth = 0): void => {
        if (!parent || depth >= 128) return
        for (const [name, child] of orderedChildren(parent)) {
          if (name === 'drawing') sourceDrawings.push(child)
          else if (name !== '#text') collect(child, depth + 1)
        }
      }
      collect(candidate)
      for (const sourceDrawing of sourceDrawings) {
        const wp = getChildren(sourceDrawing, 'anchor')[0] ?? getChildren(sourceDrawing, 'inline')[0]
        const docPr = attrs(getChildren(wp, 'docPr')[0])
        const graphicData = getChildren(getChildren(wp, 'graphic')[0], 'graphicData')[0]
        const feature = getChildren(graphicData, 'chart').length ? 'chart' : getChildren(graphicData, 'relIds').length ? 'diagram' : getChildren(graphicData, 'contentPart').length ? 'ink' : getChildren(graphicData, 'pic').length ? 'picture' : 'graphicData'
        context.drawingCoverage.push({ partPath: context.partPath ?? 'word/document.xml', treePath: context.drawingPaths?.get(sourceDrawing) ?? context.drawingPaths?.get(node) ?? `AlternateContent[${coverageBefore}]`,
          element: 'drawing', id: docPr.id, name: docPr.name, referenceId: attemptedChoice ? drawingReferenceIds(sourceDrawing)[0] : undefined,
          feature, status: 'unsupported', selectedRepresentation: 'none', representation: 'native',
          reason: attemptedChoice ? 'No selected static representation' : 'Unsupported compatibility requirement', scope: 'original' })
      }
      if (sourceDrawings.length) markSkipped()
    }
  }
  walk(p)
  resolveFields(paragraph)
  if (paragraph.images.length > 0) {
    const retained = new Set(paragraph.runs)
    paragraph.inline = inline.filter((item) => item.kind === 'image' || retained.has(item.run))
  }
  applyParagraphDefaults(paragraph, p, context?.styles, context?.tableLayers)

  for (const issue of inheritedIssues) {
    const cat = issueCategory(issue.feature)
    if (textRunAuthored.length > 0) {
      if (textRunAuthored.every(auth => auth.has(cat))) continue
    } else {
      const pPrAuthored = authoredCategories(getChildren(pPr, 'rPr')[0])
      if (pPrAuthored.has(cat)) continue
    }
    if (!appearanceIssues.some(d => d.kind === issue.kind && d.feature === issue.feature)) {
      appearanceIssues.push(issue)
    }
  }

  if (appearanceIssues.length > 0) {
    paragraph.diagnostics ??= []
    for (const issue of appearanceIssues) {
      if (!paragraph.diagnostics.some(d => d.kind === issue.kind && d.feature === issue.feature)) {
        paragraph.diagnostics.push(issue)
      }
    }
  }
  return paragraph
}

function findDescendant(node: XmlNode | undefined, tag: string): XmlNode | undefined {
  for (const [name, child] of elementChildren(node)) {
    if (name === tag) return child
    const found = findDescendant(child,tag)
    if (found) return found
  }
  return undefined
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
        delete run._fldChar
        delete run._instr
        run.field = pendingInstr
        out.push(run)
        pendingInstr = null
      }
      continue
    }
    delete run._fldChar
    delete run._instr
    out.push(run)
  }
  paragraph.runs = out
}

/** w:drawing -> wp:inline|wp:anchor -> a:graphic -> pic:pic -> a:blip r:embed */
function parseDrawing(
  drawing: XmlNode,
  images: DocxImage[] | undefined,
  context?: ParagraphContext
): DocxImage | undefined {
  const vector = context?.drawings.get(drawing)
  const nestedSelection = wordDrawingSelections(context?.drawings, drawing).at(-1)
  const selectedRepresentation = nestedSelection?.representation ?? context?.representation ?? 'native'
  const selectedEmpty = !vector?.drawing && nestedSelection && orderedChildren(nestedSelection.node).every(([name]) => name === '#text')
  const allowed = !context?.reserveDrawing || context.reserveDrawing(vector?.drawing)
  const anchor = getChildren(drawing, 'anchor')[0]
  const wp = anchor ?? getChildren(drawing, 'inline')[0]
  const docPr = attrs(getChildren(wp, 'docPr')[0])
  const graphicData = getChildren(getChildren(wp, 'graphic')[0], 'graphicData')[0]
  const graphicUri = attrs(graphicData).uri?.toLowerCase() ?? ''
  const feature = vector?.drawing?.kind ?? (getChildren(graphicData, 'pic').length ? 'picture' : getChildren(graphicData, 'chart').length ? 'chart' : getChildren(graphicData, 'relIds').length ? 'diagram' : getChildren(graphicData, 'contentPart').length ? 'ink' : graphicUri.includes('chartex') ? 'ChartEx' : graphicData ? 'graphicData' : 'drawing')
  const audit = (status: DrawingCoverageEntry['status'], reason?: string, referenceId?: string, consumed = true, limit?: number, imageSelection?: ImageSelection) => {
    if (!context?.reserveDrawing || !context.drawingCoverage) return
    context.drawingCoverage.push({ partPath: context.partPath ?? 'word/document.xml', treePath: context.drawingPaths?.get(drawing) ?? `drawing[${context.drawingCoverage.length}]`, element: 'drawing', id: docPr.id, name: docPr.name,
      referenceId, feature, status: status === 'native' && selectedRepresentation === 'fallback' ? 'fallback' : status,
      selectedRepresentation: !consumed ? 'none' : vector?.drawing ? contentRepresentation(vector.drawing) : status === 'native' ? selectedRepresentation === 'fallback' ? 'raster-fallback' : 'picture' : 'none',
      representation: selectedRepresentation, reason: reason ?? nestedSelection?.reason, scope: 'original', limit, ...(imageSelection ? { imageSelection } : {}) })
  }
  if (!allowed) { audit('unsupported', 'drawing node budget exceeded', undefined, false, DOCUMENT_DRAWING_NODE_LIMIT); return undefined }
  if (!wp) { audit('malformed', 'missing inline or anchor placement'); return undefined }
  if (selectedEmpty) {
    if (context?.reserveDrawing && context.drawingCoverage) context.drawingCoverage.push({ partPath: context.partPath ?? 'word/document.xml', treePath: context.drawingPaths?.get(drawing) ?? `drawing[${context.drawingCoverage.length}]`, element: 'drawing', id: docPr.id, name: docPr.name,
      feature: 'empty-choice', status: 'native', selectedRepresentation: 'blank', representation: selectedRepresentation, reason: 'Supported empty Choice selected', scope: 'original' })
    return undefined
  }
  const extent = getChildren(wp, 'extent')[0]
  const ea = attrs(extent)
  const widthEmu = parseFloat(ea.cx as string) || 0
  const heightEmu = parseFloat(ea.cy as string) || 0
  const effect = getChildren(wp, 'effectExtent')[0]
  const effectAttrs = attrs(effect)
  const placement = {
    effectExtentEmu: effect
      ? {
          top: parseFloat(effectAttrs.t as string) || 0,
          bottom: parseFloat(effectAttrs.b as string) || 0
        }
      : undefined,
    floating: anchor ? parseAnchor(anchor) : undefined
  }
  if (vector) {
    const nestedTextboxImage = vector.drawing?.kind === 'textbox' && !!findDescendant(findDescendant(drawing, 'txbxContent'), 'blip')
    audit(nestedTextboxImage ? 'unsupported' : 'native', nestedTextboxImage ? 'Textbox text retained; nested picture flow is not painted' : undefined, vector.referenceId)
    return { ...vector, ...placement }
  }
  // pic:pic lives inside a:graphicData, not directly under a:graphic
  const pic = graphicData ? getChildren(graphicData, 'pic')[0] : undefined
  const blip = pic ? getChildren(getChildren(pic, 'blipFill')[0], 'blip')[0] : undefined
  const rid = (attrs(blip)['r:embed'] ?? attrs(blip).embed) as string | undefined
  const svgBlip = findSvgBlip(blip)
  const svgRid = svgBlip?.embed
  if (!rid && !svgRid) { audit(feature === 'picture' || feature === 'chart' || feature === 'diagram' || feature === 'ink' ? 'malformed' : graphicData ? 'unsupported' : 'malformed', feature === 'picture' ? 'missing image relationship' : feature === 'graphicData' ? 'No supported selected payload' : 'missing selected payload part', attrs(getChildren(graphicData, 'chart')[0]).id ?? attrs(getChildren(graphicData, 'relIds')[0]).dm ?? attrs(getChildren(graphicData, 'contentPart')[0]).id); return undefined }
  if (context?.pkg && context.partPath) {
    if (rid) malformedRelationshipAttempt(context.pkg, context.partPath, rid, 'image')
    if (svgRid) malformedRelationshipAttempt(context.pkg, context.partPath, svgRid, 'image')
  }
  if (!images) { audit('malformed', 'missing image collection', rid ?? svgRid); return undefined }
  const raster = rid ? images.find((img) => img && (img as DocxImage & { relId?: string }).relId === rid) : undefined
  const svgEntry = svgRid ? images.find((img) => img && (img as DocxImage & { relId?: string }).relId === svgRid) : undefined
  if (!raster && !svgEntry) { audit('malformed', 'missing image part', rid ?? svgRid); return undefined }
  const selection = selectionFor(context, `${context?.partPath ?? ''}\u0000${rid ?? ''}\u0000${svgRid ?? ''}`)
  let svg: SvgCandidate | undefined
  let primarySvgVerdict: SvgVerdict | undefined
  if (svgEntry) {
    // Retain the candidate even when rejected, so distinct pairs never collapse
    // and an ordinary raster never inherits this source's rejection reason.
    const candidate = svgCandidate(svgEntry.data, (svgEntry as DocxImage & { partPath?: string }).partPath, svgRid)
    svg = candidate
    selection.representation = candidate.verdict.ok ? 'svg' : (raster ? 'raster' : 'none')
    selection.reason = candidate.verdict.ok ? undefined : candidate.verdict.reason
  } else if (svgBlip?.link) {
    selection.representation = raster ? 'raster' : 'none'; selection.reason = 'svg:link-only'
    svg = svgStateCandidate('svg:link-only')
  } else if (svgRid) {
    const reason = context?.imageExternal?.has(svgRid) ? 'svg:external' : 'svg:unresolved'
    selection.representation = raster ? 'raster' : 'none'; selection.reason = reason
    svg = svgStateCandidate(reason, undefined, svgRid)
  } else if (raster && looksLikeSvg(raster.data, raster.mime, (raster as DocxImage & { partPath?: string }).partPath)) {
    const verdict = scanSelfContainedSvg(raster.data)
    primarySvgVerdict = verdict
    selection.representation = verdict.ok ? 'svg' : 'none'
    selection.reason = verdict.ok ? undefined : verdict.reason
  } else {
    selection.representation = 'raster'
  }
  audit('native', undefined, rid ?? svgRid, true, undefined, selection)
  return {
    data: raster?.data ?? svgEntry!.data,
    mime: raster?.mime ?? svgEntry?.mime,
    pathHint: (raster as DocxImage & { partPath?: string } | undefined)?.partPath ?? (svgEntry as DocxImage & { partPath?: string } | undefined)?.partPath,
    widthEmu,
    heightEmu,
    ...placement,
    hasRaster: !!raster,
    ...(svg ? { svg } : {}),
    ...(primarySvgVerdict ? { primarySvgVerdict } : {}),
    imageSelection: selection
  }
}

function vmlNodeToDocxShape(node: VmlNode): DocxDrawingShape | undefined {
  if (node.kind === 'shape') {
    const vml = node.result
    const widthPt = vml.widthPt ?? 200
    const heightPt = vml.heightPt ?? 50
    const leftPt = vml.leftPt ?? 0
    const topPt = vml.topPt ?? 0
    const widthEmu = Math.round(widthPt * 12700)
    const heightEmu = Math.round(heightPt * 12700)
    const xEmu = Math.round(leftPt * 12700)
    const yEmu = Math.round(topPt * 12700)
    if (![xEmu, yEmu, widthEmu, heightEmu].every(Number.isFinite)) {
      vml.diagnostics ??= []
      if (!vml.diagnostics.some(d => d.kind === 'unsupported-geometry')) {
        vml.diagnostics.push({
          kind: 'unsupported-geometry',
          feature: 'vml-wordart',
          message: 'VML shape has non-finite bounds',
        })
      }
      return undefined
    }
    return {
      xEmu,
      yEmu,
      widthEmu,
      heightEmu,
      geometry: 'rect',
      fontFamily: vml.textBody.paragraphs[0]?.runs[0]?.fontFamily ?? 'Calibri',
      textBody: vml.textBody,
      paragraphs: [],
      ...(vml.rotationDeg !== undefined ? { rotationDeg: vml.rotationDeg } : {}),
      ...(vml.flipH ? { flipH: true } : {}),
      ...(vml.flipV ? { flipV: true } : {}),
    }
  }
  const children: DocxDrawingShape[] = []
  for (const c of node.children) {
    const s = vmlNodeToDocxShape(c)
    if (s) children.push(s)
  }
  // Retain empty groups as structural nodes so guards never silently drop siblings.
  const group = {
    off: { x: node.xEmu, y: node.yEmu },
    ext: { width: node.widthEmu, height: node.heightEmu },
    chOff: { x: Math.round(node.coordorigin.x * 12700), y: Math.round(node.coordorigin.y * 12700) },
    chExt: { width: Math.round(node.coordsize.width * 12700), height: Math.round(node.coordsize.height * 12700) },
  }
  if (![node.xEmu, node.yEmu, node.widthEmu, node.heightEmu, group.chOff.x, group.chOff.y, group.chExt.width, group.chExt.height].every(Number.isFinite)) {
    node.diagnostics ??= []
    if (!node.diagnostics.some(d => d.kind === 'unsupported-geometry')) {
      node.diagnostics.push({
        kind: 'unsupported-geometry',
        feature: 'vml-wordart',
        message: 'VML group has non-finite bounds',
      })
    }
    return undefined
  }
  return {
    xEmu: node.xEmu,
    yEmu: node.yEmu,
    widthEmu: node.widthEmu,
    heightEmu: node.heightEmu,
    geometry: 'group',
    fontFamily: '',
    paragraphs: [],
    group,
    children,
    ...(node.rotationDeg !== undefined ? { rotationDeg: node.rotationDeg } : {}),
    ...(node.flipH ? { flipH: true } : {}),
    ...(node.flipV ? { flipV: true } : {}),
  }
}

function sameVmlDiagnostic(
  list: { kind: string; partPath: string; feature?: string; identity?: string; sourcePath?: string; message: string }[],
  candidate: { kind: string; partPath: string; feature?: string; identity?: string; sourcePath?: string; message: string },
): boolean {
  return list.some(
    e =>
      e.kind === candidate.kind &&
      e.partPath === candidate.partPath &&
      (e.feature ?? '') === (candidate.feature ?? '') &&
      (e.identity ?? '') === (candidate.identity ?? '') &&
      (e.sourcePath ?? '') === (candidate.sourcePath ?? '') &&
      e.message === candidate.message,
  )
}

function pushVmlDiagnostics(
  nodes: VmlNode[],
  ctx: ParagraphContext | undefined,
  containerDiagnostics: import('../drawing/vml').VmlWordArtDiagnostic[] | undefined,
  pictIndex: number,
): void {
  if (!ctx?.pkg) return
  const context = drawingPartContext(ctx.pkg)
  const partPath = ctx.partPath ?? 'word/document.xml'
  const pictPrefix = `pict[${pictIndex}]`
  const pushOne = (d: import('../drawing/vml').VmlWordArtDiagnostic, sourcePath: string, authoredId: string | undefined): void => {
    const fullPath = `${pictPrefix}/${sourcePath}`
    const candidate = {
      kind: d.kind,
      partPath,
      feature: d.feature ?? 'vml-wordart',
      identity: authoredId ?? fullPath,
      sourcePath: fullPath,
      message: authoredId ? `${d.message} (shape ${authoredId})` : d.message,
    }
    if (!sameVmlDiagnostic(context.diagnostics, candidate)) context.diagnostics.push(candidate)
  }
  const walk = (n: VmlNode): void => {
    if (n.kind === 'shape') {
      for (const d of n.result.diagnostics ?? []) pushOne(d, n.sourcePath, n.result.shapeId)
    } else {
      for (const d of n.diagnostics ?? []) pushOne(d, n.sourcePath, n.groupId)
      for (const c of n.children) walk(c)
    }
  }
  for (const n of nodes) walk(n)
  for (const d of containerDiagnostics ?? []) {
    const fullPath = d.sourcePath ? `${pictPrefix}/${d.sourcePath}` : pictPrefix
    const candidate = {
      kind: d.kind,
      partPath,
      feature: d.feature ?? 'vml-wordart',
      identity: d.identity ?? fullPath,
      sourcePath: fullPath,
      message: d.message,
    }
    if (!sameVmlDiagnostic(context.diagnostics, candidate)) context.diagnostics.push(candidate)
  }
}

function shapeHasRenderableLeaf(shape: DocxDrawingShape): boolean {
  if (shape.geometry !== 'group') return true
  return shape.children?.some(shapeHasRenderableLeaf) ?? false
}

function parsePict(node: XmlNode, context?: ParagraphContext): DocxImage | undefined {
  if (context && !context.vmlProvenance) context.vmlProvenance = { nextPict: 0 }
  const pictIndex = context?.vmlProvenance?.nextPict ?? 0
  let container
  try {
    container = parseVmlContainer(node)
  } catch (err) {
    if (context?.vmlProvenance) context.vmlProvenance.nextPict = pictIndex + 1
    if (context?.pkg) {
      const partCtx = drawingPartContext(context.pkg)
      const partPath = context.partPath ?? 'word/document.xml'
      const pictPrefix = `pict[${pictIndex}]`
      const candidate = {
        kind: 'malformed-vml-container' as const,
        partPath,
        feature: 'vml-wordart',
        identity: pictPrefix,
        sourcePath: pictPrefix,
        message: err instanceof Error ? err.message : String(err),
      }
      if (!sameVmlDiagnostic(partCtx.diagnostics, candidate)) {
        partCtx.diagnostics.push(candidate)
      }
    }
    return undefined
  }
  if (context?.vmlProvenance) context.vmlProvenance.nextPict = pictIndex + 1
  const shapes: DocxDrawingShape[] = []
  for (const n of container.nodes) {
    const s = vmlNodeToDocxShape(n)
    if (s && [s.xEmu, s.yEmu, s.widthEmu, s.heightEmu].every(Number.isFinite)) {
      shapes.push(s)
    } else {
      if (n.kind === 'shape') {
        n.result.diagnostics ??= []
        if (!n.result.diagnostics.some(d => d.kind === 'unsupported-geometry')) {
          n.result.diagnostics.push({
            kind: 'unsupported-geometry',
            feature: 'vml-wordart',
            message: 'VML shape has non-finite bounds',
          })
        }
      } else {
        n.diagnostics ??= []
        if (!n.diagnostics.some(d => d.kind === 'unsupported-geometry')) {
          n.diagnostics.push({
            kind: 'unsupported-geometry',
            feature: 'vml-wordart',
            message: 'VML group has non-finite bounds',
          })
        }
      }
    }
  }
  pushVmlDiagnostics(container.nodes, context, container.diagnostics, pictIndex)
  if (container.nodes.length === 0 && !(container.diagnostics && container.diagnostics.length > 0)) return undefined
  if (shapes.length === 0 || !shapes.some(shapeHasRenderableLeaf)) return undefined
  let maxXEmu = 0
  let maxYEmu = 0
  for (const s of shapes) {
    maxXEmu = Math.max(maxXEmu, s.xEmu + s.widthEmu)
    maxYEmu = Math.max(maxYEmu, s.yEmu + s.heightEmu)
  }
  const drawing: DocxDrawing = {
    kind: 'diagram',
    shapes,
  }
  return {
    data: new Uint8Array(),
    widthEmu: maxXEmu,
    heightEmu: maxYEmu,
    drawing,
  }
}

/** wp:positionH / wp:positionV -> a from/offset pair. */
function parsePosition(node: XmlNode | undefined): { relativeFrom: string; offsetEmu: number; align?: string } | undefined {
  if (!node) return undefined
  const relativeFrom = (attrs(node).relativeFrom as string) ?? 'column'
  const posOffset = getChildren(node, 'posOffset')[0]
  if (posOffset) {
    const value = textOf(posOffset).trim()
    const parsed = parseFloat(value)
    return { relativeFrom, offsetEmu: Number.isFinite(parsed) ? parsed : 0 }
  }
  const align = getChildren(node, 'align')[0]
  if (align) return { relativeFrom, offsetEmu: 0, align: textOf(align).trim() }
  return { relativeFrom, offsetEmu: 0 }
}

/** Floating positioning from a wp:anchor. */
function parseAnchor(anchor: XmlNode): DocxFloating {
  const a = attrs(anchor)
  const wrap =
    getChildren(anchor, 'wrapNone')[0] !== undefined
      ? 'none'
      : getChildren(anchor, 'wrapSquare')[0] !== undefined
        ? 'square'
        : getChildren(anchor, 'wrapTight')[0] !== undefined
          ? 'tight'
          : getChildren(anchor, 'wrapThrough')[0] !== undefined
            ? 'through'
            : 'topAndBottom'
  return {
    behindDoc: a.behindDoc === '1' || a.behindDoc === 'true',
    relativeHeight: parseInt(a.relativeHeight as string, 10) || 0,
    wrap,
    posH: parsePosition(getChildren(anchor, 'positionH')[0]) ?? { relativeFrom: 'column', offsetEmu: 0 },
    posV: parsePosition(getChildren(anchor, 'positionV')[0]) ?? { relativeFrom: 'paragraph', offsetEmu: 0 },
  }
}

/** Resolve word/_rels/document.xml.rels into embedded image bytes keyed by rId. */
/** rId -> part path for word/document.xml.rels. */
async function loadDocRels(pkg: OfficePackage): Promise<Map<string, { type: string; target: string }>> {
  const out = new Map<string, { type: string; target: string }>()
  for (const rel of await partRelationshipNodes(pkg, 'word/document.xml')) {
    const a = attrs(rel)
    if (a.Id && a.Target) {
      out.set(a.Id, { type: (a.Type as string) ?? '', target: a.Target })
    }
  }
  return out
}

type RepeatedContent = Pick<DocxSection, 'header' | 'footer' | 'firstHeader' | 'firstFooter' | 'headerBlocks' | 'footerBlocks' | 'firstHeaderBlocks' | 'firstFooterBlocks'>
/** Missing references inherit; a resolved empty part is explicitly empty. */
async function loadHeaderFooter(sectPr: XmlNode, rels: Map<string, { type: string; target: string }>, pkg: OfficePackage, styles: DocxStyleContext, coverage?: DrawingCoverageEntry[], unselectedReferenceIds?: Set<string>): Promise<RepeatedContent> {
  const out: RepeatedContent = {}
  for (const [kind, tag] of [['header', 'headerReference'], ['footer', 'footerReference']] as const) {
    for (const type of ['default', 'first'] as const) {
      const ref = getChildren(sectPr, tag).find(r => attrs(r).type === type || (type === 'default' && !attrs(r).type))
      if (!ref) continue
      const blocks = await loadPart(ref, rels, pkg, styles, coverage, unselectedReferenceIds)
      if (blocks === undefined) continue
      const key = type === 'first' ? (kind === 'header' ? 'firstHeader' : 'firstFooter') : kind
      out[key] = blocks.flatMap(block => block.kind === 'p' ? [block.paragraph] : [])
      out[`${key}Blocks`] = blocks
    }
  }
  return out
}
async function loadPart(ref: XmlNode, rels: Map<string, { type: string; target: string }>, pkg: OfficePackage, styles: DocxStyleContext, coverage?: DrawingCoverageEntry[], unselectedReferenceIds?: Set<string>): Promise<DocxBlock[] | undefined> {
  const rel = rels.get(attrs(ref).id)
  if (!rel) return undefined
  const segments = rel.target.startsWith('/') ? [] : ['word']
  for (const segment of rel.target.split('/')) {
    if (segment === '..') segments.pop()
    else if (segment && segment !== '.') segments.push(segment)
  }
  const path = segments.join('/')
  const part = await pkg.xmlOrdered(path)
  if (!part) return undefined
  const numbering = await pkg.xml('word/numbering.xml')
  const state = numbering ? parseNumbering(numbering) : undefined
  const { images, external: imageExternal } = await loadDocImages(pkg, path)
  const imageSelections = new Map<string, ImageSelection>()
  const vmlProvenance: VmlProvenance = { nextPict: 0 }
  const context: ParagraphContext = { styles, drawings: await loadDrawingParts(pkg, part, path, styles.theme, p => parseParagraph(p, images, state, { styles, drawings: new Map(), imageSelections, imageExternal, pkg, partPath: path, vmlProvenance }), false), reserveDrawing: drawingReservation(pkg, path), drawingCoverage: coverage, drawingPaths: sourceDrawingPaths(part, path.includes('/header') ? 'header' : 'footer'), unselectedReferenceIds, partPath: path, pkg, imageSelections, imageExternal, vmlProvenance }
  return unwrapContentControls(part).flatMap(([name, node]): DocxBlock[] => name === 'p' ? [{ kind: 'p', paragraph: parseParagraph(node, images, state, context) }] : name === 'tbl' ? [{ kind: 'table', table: parseTable(node, state, context, images) }] : [])
}

async function loadDocImages(pkg: OfficePackage, partPath = 'word/document.xml'): Promise<{ images: DocxImage[]; external: Set<string> }> {
  const slash = partPath.lastIndexOf('/')
  const directory = partPath.slice(0, slash)
  const images: Array<DocxImage & { relId?: string; partPath?: string }> = []
  const external = new Set<string>()
  for (const rel of await partRelationshipNodes(pkg, partPath)) {
    const a = attrs(rel)
    const type = a.Type as string | undefined
    const target = (a.Target as string | undefined) ?? ''
    if (!type || !type.includes('/image') || !target) continue
    if (a.TargetMode === 'External') { if (a.Id) external.add(a.Id); continue }
    const segments = target.startsWith('/') ? [] : directory.split('/')
    for (const segment of target.split('/')) {
      if (segment === '..') segments.pop()
      else if (segment && segment !== '.') segments.push(segment)
    }
    const path = segments.join('/')
    const data = await pkg.bytes(path)
    if (!data) continue
    images.push({ data, mime: sniffImageMime(data), widthEmu: 0, heightEmu: 0, relId: a.Id, partPath: path })
  }
  return { images, external }
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
  for (const [name, child] of orderedChildren(node)) {
    if (name === 'sdt') {
      out.push(...unwrapContentControls(getChildren(child, 'sdtContent')[0]))
    } else if (name === 'sdtContent') {
      out.push(...unwrapContentControls(child))
    } else {
      out.push([name, child])
    }
  }
  return out
}

export async function parseDocx(pkg: OfficePackage): Promise<DocxDocument> {
  const doc = await pkg.xmlOrdered('word/document.xml')
  if (!doc) throw new Error('word/document.xml missing — not a valid docx?')
  const stylesXml = await pkg.xml('word/styles.xml')
  const docRels = await loadDocRels(pkg)
  const themeRel = [...docRels.values()].find(rel => rel.type.endsWith('/theme'))
  const themePath = themeRel?.target.startsWith('/') ? themeRel.target.slice(1) : themeRel ? `word/${themeRel.target}` : undefined
  const styles = styleContext(stylesXml,readTheme(themePath ? await pkg.xml(themePath) : undefined))
  let defaultFontFamily = 'Calibri'
  let defaultFontSizePt = 11
  const styleDefaults = new Map<string, { fontFamily?: string; fontSizePt?: number }>()
  if (stylesXml) {
    const docDefaults = stylesXml['docDefaults']
    if (docDefaults && typeof docDefaults === 'object') {
      const rPrDefault = getChildren(docDefaults as XmlNode, 'rPrDefault')[0]
      const rPr = rPrDefault ? getChildren(rPrDefault, 'rPr')[0] : undefined
      const fam = readRunProperties(rPr,styles.theme).fontFamily
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
  const { images: docImages, external: imageExternal } = await loadDocImages(pkg)
  const numberingRoot = await pkg.xml('word/numbering.xml')
  const numbering = numberingRoot ? parseNumbering(numberingRoot) : undefined
  const drawingCoverage: DrawingCoverageEntry[] = []
  const unselectedReferenceIds = new Set<string>()
  const imageSelections = new Map<string, ImageSelection>()
  const vmlProvenance: VmlProvenance = { nextPict: 0 }
  const context: ParagraphContext = {styles,drawings:await loadDrawingParts(pkg,doc,'word/document.xml',styles.theme,p => parseParagraph(p,docImages,numbering,{styles,drawings:new Map(),imageSelections,imageExternal,pkg,partPath:'word/document.xml',vmlProvenance}),false),reserveDrawing:drawingReservation(pkg,'word/document.xml'),drawingCoverage,drawingPaths:sourceDrawingPaths(doc),unselectedReferenceIds,partPath:'word/document.xml',pkg,imageSelections,imageExternal,vmlProvenance}
  const sections: DocxSection[] = []
  let current: DocxSection = {
    margins: { topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, headerTwips: 720, footerTwips: 720, gutterTwips: 0 },
    pageSize: { widthTwips: 12240, heightTwips: 15840, orientation: 'portrait' },
    paragraphs: [],
    blocks: [],
  }
  const bodyChildren = unwrapContentControls(body as XmlNode).flatMap(([name, node]): Array<[string, XmlNode]> => {
    const sect = name === 'p' ? getChildren(getChildren(node, 'pPr')[0], 'sectPr')[0] : undefined
    return sect ? [[name, node], ['sectPr', sect]] : [[name, node]]
  })
  for (const [name, node] of bodyChildren) {
    if (name === 'p') {
      const para = parseParagraph(node, docImages, numbering,context)
      current.paragraphs.push(para)
      current.blocks.push({ kind: 'p', paragraph: para })
    } else if (name === 'tbl' && node) {
      const table = parseTable(node, numbering,context,docImages)
      const block: DocxBlock = { kind: 'table', table }
      current.blocks.push(block)
    } else if (name === 'sectPr') {
      // section properties at body level — finalize current section
      const hf = await loadHeaderFooter(node, docRels, pkg,styles,drawingCoverage,unselectedReferenceIds)
      Object.assign(current, hf)
      const sectionType = attrs(getChildren(node, 'type')[0]).val
      current.type = sectionType === 'continuous' || sectionType === 'evenPage' || sectionType === 'oddPage' || sectionType === 'nextColumn' ? sectionType : 'nextPage'
      if (node['titlePg'] !== undefined) current.titlePg = !['0', 'false', 'off'].includes(attrs(getChildren(node, 'titlePg')[0]).val)
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
        firstHeader: current.firstHeader, firstFooter: current.firstFooter,
        headerBlocks: current.headerBlocks, footerBlocks: current.footerBlocks,
        firstHeaderBlocks: current.firstHeaderBlocks, firstFooterBlocks: current.firstFooterBlocks,
      }
      // A source sectPr explicitly finalizes this section even when it has
      // no body blocks. The fresh trailing builder is retained only if used.
      sections.push(current)
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
  for (const issue of drawingPartContext(pkg).diagnostics) {
    if (attemptedMalformedRelationshipIssue(pkg, issue)) continue
    if (issue.kind === 'node-budget' && issue.reason === 'document-budget' && drawingCoverage.some(entry =>
      entry.partPath === issue.partPath && entry.reason === 'drawing node budget exceeded' && entry.limit === issue.limit)) continue
    if ((issue.sourceReferenceId ?? issue.identity) && unselectedReferenceIds.has(issue.sourceReferenceId ?? issue.identity!) && !drawingCoverage.some(entry => coverageIssueMatchesEntry(issue, entry))) continue
    const matched = drawingCoverage.filter(entry => coverageIssueMatchesEntry(issue, entry))
    if (matched.length) {
      for (const entry of matched) {
        if (issue.kind === 'missing-part' || issue.kind === 'malformed-part') entry.status = 'malformed'
        else if (['external-reference', 'unsupported-content', 'content-cycle', 'content-depth', 'group-depth', 'node-budget'].includes(issue.kind)) entry.status = 'unsupported'
        entry.reason = issue.reason ?? issue.message
        entry.limit = issue.limit
        // A text-only entry already reports rendered fallback text (its reason
        // names the fallback); resetting it to none would un-report content
        // the entry proves is retained.
        if (entry.selectedRepresentation !== 'raster-fallback' && entry.selectedRepresentation !== 'text-only' && !['document-budget', 'source-node-limit'].includes(issue.reason ?? '') && !entry.reason?.includes('cached-picture')) entry.selectedRepresentation = 'none'
      }
    } else drawingCoverage.push({ partPath: issue.partPath, treePath: issue.sourcePath ?? issue.identity ?? issue.partPath, element: issue.kind, id: issue.identity,
      feature: issue.feature ?? issue.kind, status: issue.kind === 'missing-part' || issue.kind === 'malformed-part' ? 'malformed' : 'unsupported',
      representation: 'native', selectedRepresentation: 'none', reason: issue.reason ?? issue.message, scope: 'diagnostic', limit: issue.limit })
  }
  // Header/footer parts may be resolved for several sections. A source
  // placement belongs to the part once even when Word repeats it on pages.
  const seenPlacements = new Set<string>()
  const uniqueCoverage = drawingCoverage.filter(entry => {
    if (entry.scope === 'diagnostic') return true
    const key = `${entry.partPath}\u0000${entry.treePath}`
    if (seenPlacements.has(key)) return false
    seenPlacements.add(key)
    return true
  })
  return {
    drawingCoverage: uniqueCoverage,
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
  return { ...(a.val !== undefined ? { style: a.val } : {}), ...(a.color !== undefined ? { color: a.color } : {}), ...(a.sz !== undefined ? { widthPt: Number(a.sz) / 8 } : {}) }
}

function parseBorders(parent: XmlNode): DocxTableBorders {
  const el = getChildren(parent, 'tblBorders')[0] ?? getChildren(parent, 'tcBorders')[0]
  if (!el) return {}
  const out: DocxTableBorders = {}
  for (const side of ['top','bottom','left','right','insideH','insideV'] as const) {
    const node = getChildren(el,side)[0]
    if (node) out[side] = parseSideBorder(node)
  }
  return out
}

/** Border sides merge individually; explicit nil/none clears an inherited side. */
function mergeBorders(...sources: Array<DocxTableBorders | undefined>): DocxTableBorders {
  const out: DocxTableBorders = {}
  for (const source of sources) for (const side of ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'] as const) {
    if (!source || !(side in source)) continue
    out[side] = source[side] === undefined ? undefined : { ...out[side], ...source[side] }
  }
  return out
}

const DEFAULT_CELL_MARGINS: DocxTableCellMargins = { topTwips: 0, rightTwips: 108, bottomTwips: 0, leftTwips: 108 }

function parseCellMargins(parent: XmlNode | undefined): Partial<DocxTableCellMargins> {
  const mar = getChildren(parent, 'tblCellMar')[0] ?? getChildren(parent, 'tcMar')[0]
  const out: Partial<DocxTableCellMargins> = {}
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const node = getChildren(mar, side)[0] ?? getChildren(mar, side === 'left' ? 'start' : side === 'right' ? 'end' : side)[0]
    if (node) out[`${side}Twips`] = attrs(node).type === 'nil' ? 0 : twips(attrs(node).w) ?? 0
  }
  return out
}
function shading(parent: XmlNode | undefined): Partial<DocxTableCell> {
  const shd = getChildren(parent, 'shd')[0]
  return shd ? { fill: attrs(shd).val === 'nil' || attrs(shd).fill === 'auto' ? undefined : attrs(shd).fill } : {}
}
function parseTableCell(tc: XmlNode, cellNumbering?: NumberingState, context?: ParagraphContext, images?: DocxImage[], inherited: Partial<DocxTableCell> = {}): DocxTableCell {
  const tcPr = getChildren(tc, 'tcPr')[0]
  const cell: DocxTableCell = { ...inherited, paragraphs: [], gridSpan: 1 }
  if (tcPr) {
    const a = attrs(tcPr['gridSpan'] as XmlNode | undefined)
    if (a.val !== undefined) cell.gridSpan = parseInt(a.val as string, 10) || 1
    const vMerge = attrs(tcPr['vMerge'] as XmlNode | undefined).val as string | undefined
    if (vMerge === 'restart') cell.vMerge = 'restart'
    else if (vMerge !== undefined || tcPr['vMerge'] !== undefined) cell.vMerge = 'continue'
    Object.assign(cell, shading(tcPr))
    cell.borders = mergeBorders(inherited.borders, parseBorders(tcPr))
    cell.margins = { ...DEFAULT_CELL_MARGINS, ...inherited.margins, ...parseCellMargins(tcPr) }
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
    const textDirection = attrs(getChildren(tcPr, 'textDirection')[0]).val as DocxTableCell['textDirection']
    if (textDirection === 'lrTb' || textDirection === 'tbRl' || textDirection === 'btLr' || textDirection === 'lrTbV' || textDirection === 'tbRlV' || textDirection === 'tbLrV') cell.textDirection = textDirection
  }
  for (const [name, node] of unwrapContentControls(tc)) {
    if (name === 'p') cell.paragraphs.push(parseParagraph(node, images, cellNumbering,context))
  }
  // Automatic foreground follows the resolved cell background. Explicit
  // paragraph/run colors, including inherited style colors, retain priority.
  if (cell.fill && /^[\da-f]{6}$/i.test(cell.fill)) {
    const rgb = [0, 2, 4].map(i => parseInt(cell.fill!.slice(i, i + 2), 16))
    if (rgb[0] * 0.299 + rgb[1] * 0.587 + rgb[2] * 0.114 < 128)
      for (const paragraph of cell.paragraphs) for (const run of paragraph.runs) run.color ??= 'FFFFFF'
  }
  return cell
}

export function parseTable(tbl: XmlNode, tableNumbering?: NumberingState, context?: ParagraphContext, images?: DocxImage[]): DocxTable {
  const tblPr = getChildren(tbl, 'tblPr')[0]
  const chain = context ? styleChain(attrs(getChildren(tblPr, 'tblStyle')[0]).val ?? context.styles.defaultTable, context.styles) : []
  const table: DocxTable = {
    gridColsTwips: [],
    rows: [],
    cellMargins: { ...DEFAULT_CELL_MARGINS, ...Object.assign({}, ...chain.map(style => parseCellMargins(getChildren(style, 'tblPr')[0]))), ...parseCellMargins(tblPr) },
    borders: mergeBorders(...chain.map(style => parseBorders(getChildren(style, 'tblPr')[0] ?? {})), tblPr ? parseBorders(tblPr) : {}),
  }
  Object.assign(table, ...chain.map(style => shading(getChildren(style, 'tblPr')[0])), shading(tblPr))
  const grid = getChildren(tbl, 'tblGrid')[0]
  if (grid) {
    for (const col of getChildren(grid, 'gridCol')) {
      table.gridColsTwips.push(twips(attrs(col).w) ?? 0)
    }
  }
  const sourceRows = unwrapContentControls(tbl).filter(([name]) => name === 'tr').map(([, node]) => node)
  const columnCount = Math.max(table.gridColsTwips.length, ...sourceRows.map(node => Number(attrs(getChildren(getChildren(node, 'trPr')[0], 'gridBefore')[0]).val ?? 0) + Number(attrs(getChildren(getChildren(node, 'trPr')[0], 'gridAfter')[0]).val ?? 0) + unwrapContentControls(node).filter(([name]) => name === 'tc').reduce((n, [, tc]) => n + Math.max(1, Number(attrs(getChildren(getChildren(tc, 'tcPr')[0], 'gridSpan')[0]).val ?? 1)), 0)))
  for (const [ri, node] of sourceRows.entries()) {
    const trPr = getChildren(node, 'trPr')[0]
    const row: DocxTableRow = { cells: [], gridBefore: Number(attrs(getChildren(trPr, 'gridBefore')[0]).val ?? 0), gridAfter: Number(attrs(getChildren(trPr, 'gridAfter')[0]).val ?? 0) }
    if (trPr) {
      const trHeight = getChildren(trPr, 'trHeight')[0]
      if (trHeight) {
        const a = attrs(trHeight)
        row.heightTwips = twips(a.val)
        const rule = a.hRule as string | undefined
        if (rule === 'exact' || rule === 'atLeast' || rule === 'auto') row.heightRule = rule
      }
      if (trPr['tblHeader'] !== undefined) row.isHeader = !['0', 'false', 'off'].includes(attrs(getChildren(trPr, 'tblHeader')[0]).val)
    }
    let col = row.gridBefore ?? 0
    for (const [name, tc] of unwrapContentControls(node)) {
      if (name !== 'tc') continue
      const span = Math.max(1, Number(attrs(getChildren(getChildren(tc, 'tcPr')[0], 'gridSpan')[0]).val ?? 1))
      const layers = tableCellLayers(chain, tblPr, ri, sourceRows.length, col, span, columnCount)
      const inherited: Partial<DocxTableCell> = { margins: { ...table.cellMargins }, borders: {} }
      for (const layer of layers) {
        const tcPr = getChildren(layer, 'tcPr')[0]
        Object.assign(inherited, shading(tcPr))
        inherited.borders = mergeBorders(inherited.borders, parseBorders(tcPr ?? {}))
        inherited.margins = { ...inherited.margins!, ...parseCellMargins(tcPr) }
        const align = attrs(getChildren(tcPr, 'vAlign')[0]).val
        if (align === 'top' || align === 'center' || align === 'bottom') inherited.vAlign = align
      }
      const cellContext = context ? { ...context, tableLayers: { pPr: layers.map(layer => getChildren(layer, 'pPr')[0]), rPr: layers.map(layer => getChildren(layer, 'rPr')[0]) } } : undefined
      row.cells.push(parseTableCell(tc, tableNumbering, cellContext, images, inherited))
      col += span
    }
    table.rows.push(row)
  }
  // Word may omit w:tblGrid entirely (tblW type=auto). Fall back to the cells'
  // own w:tcW values, then to an even split, so those tables still lay out.
  if (table.gridColsTwips.length === 0) {
    const maxCols = table.rows.reduce((max, r) => {
      let n = (r.gridBefore ?? 0) + (r.gridAfter ?? 0)
      for (const c of r.cells) n += Math.max(1, c.gridSpan)
      return Math.max(max, n)
    }, 0)
    if (maxCols > 0) {
      const widths = new Array<number>(maxCols).fill(0)
      const seen = new Array<boolean>(maxCols).fill(false)
      for (const row of table.rows) {
        let ci = row.gridBefore ?? 0
        for (const cell of row.cells) {
          const span = Math.max(1, cell.gridSpan)
          if (cell.widthTwips !== undefined && cell.widthTwips > 0) {
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
        const preferred = attrs(getChildren(tblPr, 'tblW')[0])
        const preferredWidth = preferred.type === 'dxa' ? twips(preferred.w) : undefined
        const fixed = attrs(getChildren(tblPr, 'tblLayout')[0]).type === 'fixed'
        table.autoWidth = !fixed && !(preferredWidth && preferredWidth > 0)
        for (let i = 0; i < maxCols; i++) widths[i] = Math.floor((preferredWidth || 9360) / maxCols)
      }
      table.gridColsTwips = widths.map((w) => Math.round(w))
    }
  }
  return table
}

/** Conditional priority is independent of XML region order. Base chain is immutable. */
function tableCellLayers(chain: XmlNode[], tblPr: XmlNode | undefined, row: number, rowCount: number, col: number, span: number, columns: number): XmlNode[] {
  const look = Object.assign({}, ...chain.map(style => attrs(getChildren(getChildren(style, 'tblPr')[0], 'tblLook')[0])), attrs(getChildren(tblPr, 'tblLook')[0]))
  const bits = parseInt(look.val ?? '0', 16)
  const enabled = (name: string, bit: number) => look[name] !== undefined ? ['1', 'true', 'on'].includes(look[name]) : !!(bits & bit)
  const firstRow = enabled('firstRow', 0x20), lastRow = enabled('lastRow', 0x40), firstCol = enabled('firstColumn', 0x80), lastCol = enabled('lastColumn', 0x100)
  const top = firstRow && row === 0, bottom = lastRow && row === rowCount - 1, left = firstCol && col === 0, right = lastCol && col + span === columns
  const bandSize = (tag: string) => Math.max(1, Number(attrs(getChildren(tblPr, tag)[0]).val ?? chain.reduce((value, style) => attrs(getChildren(getChildren(style, 'tblPr')[0], tag)[0]).val ?? value, '1')))
  const regions = ['wholeTable']
  if (!enabled('noVBand', 0x400) && !left && !right) regions.push(Math.floor(Math.max(0, col - (firstCol ? 1 : 0)) / bandSize('tblStyleColBandSize')) % 2 ? 'band2Vert' : 'band1Vert')
  if (!enabled('noHBand', 0x200) && !top && !bottom) regions.push(Math.floor(Math.max(0, row - (firstRow ? 1 : 0)) / bandSize('tblStyleRowBandSize')) % 2 ? 'band2Horz' : 'band1Horz')
  if (left) regions.push('firstCol')
  if (right) regions.push('lastCol')
  if (top) regions.push('firstRow')
  if (bottom) regions.push('lastRow')
  if (top && left) regions.push('nwCell')
  if (top && right) regions.push('neCell')
  if (bottom && left) regions.push('swCell')
  if (bottom && right) regions.push('seCell')
  return [...chain, ...regions.flatMap(type => chain.flatMap(style => getChildren(style, 'tblStylePr').filter(region => attrs(region).type === type)))]
}

export type { TableCellBorder }
