/**
 * SpreadsheetML token-stream consumers (design note §3.4, §4.3, §7.7).
 *
 * S1.1b provides the admission counter used by the verification pass (gate
 * G2): it sees every token of every XML part exactly once and keeps only
 * counters and a bounded formula-text buffer, so nothing proportional to the
 * part is retained. The paged record builders (S1.2) will live next to it.
 */
import { sheetFormat } from '../columnWidth'
import { frozenPaneFrom, type XlsxFrozenPane } from '../frozenPane'
import { PackageError } from '../../core/package-error'
import { localName, XstringDecoder, type XmlEvents } from '../../core/xml-stream'
import { isLoadEvaluated, MAX_LOAD_FORMULA_CHARS } from '../view/values'
import { ANCHOR_ELEMENTS, DRAWING_ROOT, type AnchorParts } from '../anchorRecords'
import { registerNamespaceAware, type XmlNode } from '../../core/xml'
import type { SheetColorContext } from '../sheetColors'
import type { SparklineDiagnostic, XlsxSparklineGroup } from '../sparklines/model'
import { MAX_POINTS_PER_SHEET, MAX_POINTS_PER_SPARKLINE, MAX_SPARKLINES_PER_GROUP, MAX_SPARKLINE_GROUPS_PER_SHEET, SPARKLINE_EXTENSION_URI, SPREADSHEETML_NS, X14_SPARKLINE_NS, XM_SPARKLINE_NS, parseSparklineGroups } from '../sparklines/parseGroups'
import { parseA1Area } from '../ranges'

export interface AdmissionCounters {
  expandedBytes: number
  maxPartBytes: number
  elements: number
  attributes: number
  textChars: number
  rows: number
  cells: number
  uniqueStrings: number
  longestString: number
  merges: number
  styleRecords: number
  maxAttributesPerElement: number
  loadEvaluatedFormulas: number
  fullCalcOnLoad: boolean
  /** `mc:AlternateContent` subtrees seen in workbook data parts (each one a counted diagnostic). */
  alternateContentSubtrees: number
  /** Structural domains whose metadata sat inside AlternateContent (never a known-empty result, §3.4). */
  structuralIncomplete: Array<'merges' | 'tables' | 'drawings' | 'columns' | 'rows'>
  /** Parts above 1 MiB declaring more than maxRatio:1 (diagnostic only). */
  highRatioParts: string[]
  /** Parts with sheetData/row/c or si inside AlternateContent (recorded in 'record' mode). */
  alternateContentData: string[]
}

export function newAdmissionCounters(): AdmissionCounters {
  return {
    expandedBytes: 0, maxPartBytes: 0, elements: 0, attributes: 0, textChars: 0, rows: 0, cells: 0, uniqueStrings: 0,
    longestString: 0, merges: 0, styleRecords: 0, maxAttributesPerElement: 0, loadEvaluatedFormulas: 0, fullCalcOnLoad: false,
    alternateContentSubtrees: 0, structuralIncomplete: [], highRatioParts: [], alternateContentData: [],
  }
}

/** Formula text kept for the volatile-function scan (values.ts owns the rule). */
const MAX_FORMULA_CHARS = MAX_LOAD_FORMULA_CHARS

const STYLE_RECORDS = new Set(['xf', 'font', 'fill', 'border', 'numFmt', 'dxf'])
const STRUCTURE: Record<string, AdmissionCounters['structuralIncomplete'][number]> = {
  mergeCells: 'merges', mergeCell: 'merges', tableParts: 'tables', tablePart: 'tables',
  drawing: 'drawings', legacyDrawing: 'drawings', cols: 'columns', col: 'columns',
}
const TRUE = (v: string | undefined) => v === '1' || v === 'true'

function attr(names: readonly string[], values: readonly string[], count: number, wanted: string): string | undefined {
  for (let i = 0; i < count; i++) if (localName(names[i]) === wanted) return values[i]
  return undefined
}

export interface CountingOptions {
  /** Called after every event that can grow a counter; throws to stop the pass. */
  check(counters: AdmissionCounters, pendingStringChars: number): void
  /**
   * Data inside AlternateContent: 'refuse' (large/paged mode, default) throws a
   * typed `unsupported`; 'record' (small object files, which keep today's
   * rendering) lists the part so the loader can attach a visible diagnostic.
   */
  alternateContentData?: 'refuse' | 'record'
}

/**
 * Count one XML part. The part kind is taken from its root element, never
 * from its name. Data (`sheetData`/`row`/`c`, `si`) inside AlternateContent is
 * dropped by the legacy tree walk: large mode refuses it with a typed
 * `unsupported` (§3.4, spec §4.3.8); small files record it (decision 5).
 */
export function countingEvents(counters: AdmissionCounters, partId: string, options: CountingOptions): XmlEvents {
  type Kind = 'worksheet' | 'sst' | 'styleSheet' | 'workbook' | 'other'
  let kind: Kind | undefined
  const stack: string[] = []
  let mcDepth = 0
  let inSheetData = false
  // cell state
  let inCell = false
  // cached value: non-whitespace <v> text, an inline string, or t="str" with
  // any <v> (a cached empty string). parse.ts also treats whitespace-only
  // <v> text as cached (measured: fast-xml-parser keeps " "); counting it as
  // uncached here only over-reserves, which is the fail-closed direction.
  let cellHasValue = false
  let cellSawV = false
  let cellHasInline = false
  let cellIsStr = false
  let cellCalcAlways = false
  let cellHasFormula = false
  let formulaText = ''
  let formulaVolatile = false
  let stringChars = 0
  let inString: 'si' | 'is' | 'v' | undefined
  let inPhonetic = 0

  const finishCell = (): void => {
    const cached = cellHasValue || cellHasInline || (cellIsStr && cellSawV)
    if (cellHasFormula && (formulaVolatile || isLoadEvaluated({ cached, calcAlways: cellCalcAlways, text: formulaText }, counters.fullCalcOnLoad))) {
      counters.loadEvaluatedFormulas++
    }
    inCell = false
  }
  const endString = (): void => {
    if (stringChars > counters.longestString) counters.longestString = stringChars
    stringChars = 0
    inString = undefined
  }

  const open = (local: string, names: readonly string[], values: readonly string[], count: number, selfClosing: boolean): void => {
    if (kind === undefined) {
      kind = local === 'worksheet' || local === 'sst' || local === 'styleSheet' || local === 'workbook' ? local : 'other'
    }
    if (local === 'AlternateContent') {
      if (mcDepth === 0) counters.alternateContentSubtrees++
      if (!selfClosing) mcDepth++
    } else if (mcDepth > 0) {
      if ((kind === 'worksheet' && (local === 'sheetData' || local === 'row' || local === 'c')) || (kind === 'sst' && local === 'si')) {
        if (options.alternateContentData !== 'record') {
          throw new PackageError('unsupported', `${local} inside mc:AlternateContent is not supported; the sheet would load blank`, partId)
        }
        if (!counters.alternateContentData.includes(partId)) counters.alternateContentData.push(partId)
      }
      const domain = kind === 'worksheet' ? STRUCTURE[local] : undefined
      if (domain !== undefined && !counters.structuralIncomplete.includes(domain)) counters.structuralIncomplete.push(domain)
    }
    switch (kind) {
      case 'worksheet':
        if (local === 'sheetData') inSheetData = !selfClosing
        else if (local === 'row' && inSheetData) counters.rows++
        else if (local === 'c' && inSheetData) {
          counters.cells++
          inCell = true
          cellHasValue = false
          cellSawV = false
          cellHasInline = false
          cellIsStr = attr(names, values, count, 't') === 'str'
          cellHasFormula = false
          formulaText = ''
          formulaVolatile = false
          cellCalcAlways = TRUE(attr(names, values, count, 'ca'))
          if (selfClosing) finishCell()
        } else if (inCell && local === 'f') {
          cellHasFormula = true
          if (TRUE(attr(names, values, count, 'ca'))) cellCalcAlways = true
        } else if (inCell && local === 'v') {
          cellSawV = true
          if (!selfClosing) inString = 'v'
        } else if (inCell && local === 'is') {
          cellHasInline = true
          if (!selfClosing) inString = 'is'
        } else if (local === 'mergeCell') counters.merges++
        else if (local === 'rPh' && inString === 'is' && !selfClosing) inPhonetic++
        break
      case 'sst':
        if (local === 'si') {
          counters.uniqueStrings++
          if (!selfClosing) inString = 'si'
          else endString()
        } else if (local === 'rPh' && inString === 'si' && !selfClosing) inPhonetic++
        break
      case 'styleSheet':
        if (STYLE_RECORDS.has(local)) counters.styleRecords++
        break
      case 'workbook':
        if (local === 'calcPr' && TRUE(attr(names, values, count, 'fullCalcOnLoad'))) counters.fullCalcOnLoad = true
        break
      default:
        break
    }
    if (!selfClosing) stack.push(local)
  }

  const close = (local: string): void => {
    stack.pop()
    if (local === 'AlternateContent' && mcDepth > 0) mcDepth--
    else if (local === 'rPh' && inPhonetic > 0) inPhonetic--
    if (kind === 'worksheet') {
      if (local === 'sheetData') inSheetData = false
      else if (local === 'c' && inCell) finishCell()
      else if ((local === 'v' && inString === 'v') || (local === 'is' && inString === 'is')) endString()
    } else if (kind === 'sst' && local === 'si' && inString === 'si') endString()
  }

  return {
    open(qname, names, values, count, selfClosing) {
      counters.elements++
      counters.attributes += count
      if (count > counters.maxAttributesPerElement) counters.maxAttributesPerElement = count
      open(localName(qname), names, values, count, selfClosing)
      options.check(counters, stringChars)
    },
    close(qname) {
      close(localName(qname))
      options.check(counters, stringChars)
    },
    text(text) {
      counters.textChars += text.length
      const top = stack[stack.length - 1]
      if (inCell && cellHasFormula && top === 'f') {
        if (formulaText.length + text.length > MAX_FORMULA_CHARS) formulaVolatile = true
        else formulaText += text
      }
      if (inCell && top === 'v' && /\S/.test(text)) cellHasValue = true
      // shared/inline strings: t and r/t text, never phonetic runs (parse.ts sharedStrings)
      if (inString !== undefined && inPhonetic === 0 && (inString === 'v' ? top === 'v' : top === 't')) stringChars += text.length
      options.check(counters, stringChars)
    },
  }
}

// ---------------------------------------------------------------------------
// Paged record builders (S1.2): shared strings and worksheet records straight
// from the token stream into bounded pages, never through a tree or a
// whole-part string. Cell value normalization mirrors parse.ts parseSheet, so
// both adapters answer identically under one value policy (spec §4.2).

const FORMULA_TEXT_LIMIT = 1 << 20

/**
 * Collects `si` strings and hands each out once complete. Mirrors parse.ts
 * sharedStrings: every direct `t` is concatenated and decoded as one text,
 * then the first `t` of each `r` run is decoded and appended, in that order
 * whatever the document order; phonetic runs never count.
 */
export function sharedStringEvents(partId: string, onString: (text: string) => void, maxStringChars: number): XmlEvents {
  const stack: string[] = []
  let mcDepth = 0
  let inSi = false
  let inPhonetic = 0
  let collecting: 'direct' | 'run' | undefined
  let direct = ''
  let runs = ''
  let directDecoder = new XstringDecoder()
  let runDecoder = new XstringDecoder()
  let runTSeen = false
  return {
    open(qname, _names, _values, _count, selfClosing) {
      const local = localName(qname)
      if (local === 'AlternateContent') { if (!selfClosing) mcDepth++ }
      else if (local === 'si') {
        if (mcDepth > 0) throw new PackageError('unsupported', 'si inside mc:AlternateContent is not supported; shared strings would be lost', partId)
        if (selfClosing) { onString(''); return }
        inSi = true; direct = ''; runs = ''; directDecoder = new XstringDecoder()
      } else if (inSi && local === 'rPh' && !selfClosing) inPhonetic++
      else if (inSi && local === 'r') runTSeen = false
      else if (inSi && inPhonetic === 0 && local === 't' && !selfClosing) {
        const parent = stack[stack.length - 1]
        if (parent === 'si') collecting = 'direct'
        else if (parent === 'r' && !runTSeen) {
          runTSeen = true
          collecting = 'run'
          runDecoder = new XstringDecoder()
        }
      }
      if (!selfClosing) stack.push(local)
    },
    close(qname) {
      const local = localName(qname)
      stack.pop()
      if (local === 'AlternateContent' && mcDepth > 0) mcDepth--
      else if (local === 'rPh' && inPhonetic > 0) inPhonetic--
      else if (local === 't' && collecting) {
        if (collecting === 'run') runs += runDecoder.push('', true)
        collecting = undefined
      } else if (local === 'si' && inSi) { inSi = false; onString(direct + directDecoder.push('', true) + runs) }
    },
    text(text) {
      if (!collecting) return
      if (collecting === 'direct') direct += directDecoder.push(text)
      else runs += runDecoder.push(text)
      if (direct.length + runs.length > maxStringChars) throw new PackageError('resource-limit', `a shared string exceeds ${maxStringChars} characters`, partId)
    },
  }
}

/** Value-kind codes stored in pages (mirrors sourcePages.ts K_*). */
export interface CellRecord {
  row: number
  col: number
  kind: number
  styleIndex: number
  num: number
  text: string
  cached: boolean
  valueIsError: boolean
  meta?: {
    formula?: string
    sharedFormula?: { si: number; ref?: string }
    arrayRef?: string
    formulaType?: string
    dynamicArray?: true
    ca?: true
  }
}
export interface SheetMetadata {
  rowHeights: { row: number; heightPt: number }[]
  hiddenRows: { r0: number; r1: number }[]
  cols: { min: number; max: number; widthChars?: number; hidden: boolean }[]
  /** `<sheetFormatPr>` column defaults (Excel column-width model). */
  format?: { baseColWidth?: number; defaultColWidth?: number }
  merges: string[]
  tablePartIds: (string | undefined)[]
  drawingRefs: number
  drawingIds: (string | undefined)[]
  /** `<legacyDrawing r:id>` (VML) references, painted by the drawing content
   * reader (S13.2); they carry no anchors the structural query reports. */
  legacyDrawingIds?: (string | undefined)[]
  structuralIncomplete: Array<'merges' | 'tables' | 'drawings' | 'columns' | 'rows'>
  /** Cells whose coordinates fall outside the sheet (invisible in both adapters). */
  skippedCells: number
  /** Authored frozen panes of the first sheet view. */
  frozen?: XlsxFrozenPane
  /** Metadata-only rows arrived out of order: rowHeights/hiddenRows need normalizeRowMetadata. */
  unorderedRows?: boolean
  /** Bounded sparkline records parsed from worksheet extension parts. */
  sparklineGroups?: XlsxSparklineGroup[]
  sparklineDiagnostics?: SparklineDiagnostic[]
}
/** Stable sort by row (as the legacy geometry sorts rows) and re-merge hidden intervals. */
export function normalizeRowMetadata(meta: SheetMetadata): void {
  if (!meta.unorderedRows) return
  meta.rowHeights = meta.rowHeights.map((h, i) => ({ h, i })).sort((a, b) => a.h.row - b.h.row || a.i - b.i).map(x => x.h)
  const hidden = [...meta.hiddenRows].sort((a, b) => a.r0 - b.r0)
  const merged: { r0: number; r1: number }[] = []
  for (const run of hidden) {
    const last = merged[merged.length - 1]
    if (last && run.r0 <= last.r1 + 1) last.r1 = Math.max(last.r1, run.r1)
    else merged.push({ ...run })
  }
  meta.hiddenRows = merged
  meta.unorderedRows = false
}
export interface SheetRecordOptions {
  partId: string
  sharedStringCount: number
  dynamicIndices: ReadonlySet<number>
  maxTextChars: number
  /** Cap on merges, tableParts and drawing references each (typed resource-limit). */
  maxStructuralRecords?: number
  sparklineColorContext?: SheetColorContext
  onSparklineDiagnostic?(diagnostic: SparklineDiagnostic): void
  onCell(record: CellRecord): void
  /** Called after every complete row; the importer cuts pages at row boundaries. */
  onRowEnd(row: number): void
  declared?: {
    autoFilter(attrs: Record<string, string>): { addColumn(input: { attrs: Record<string, string>; ruleCount: number; firstRule?: { tag: string; attrs: Record<string, string>; filterCount: number; dateGroupItemCount: number; customFilters: Record<string, string>[]; customFilterCount: number } }): void } | undefined
    sortState(attrs: Record<string, string>, placement: 'direct' | 'nested'): { addCondition(attrs: Record<string, string>): void } | undefined
  }
}

const K_EMPTY = 0, K_NUMBER = 1, K_SST = 2, K_TEXT = 3, K_BOOL = 4, K_ERROR = 5

function refOf(ref: string): [number, number] | undefined {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)
  if (!m) return undefined
  let col = 0
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64)
  return [parseInt(m[2], 10) - 1, col - 1]
}
function attrMap(names: readonly string[], values: readonly string[], count: number): Record<string, string> {
  const out: Record<string, string> = {}
  for (let i = 0; i < count; i++) out[localName(names[i])] = values[i]
  return out
}
const MAX_COORD_ROW = 1_048_575, MAX_COORD_COL = 16_383

/**
 * Worksheet records: cells (value normalization as parse.ts), row heights and
 * hidden rows, column specs, merges, tablePart ids, drawing references and
 * declared autoFilter/sortState (fed to the S8.3 builders). Rows and cells
 * must ascend strictly; anything else is a typed `unsupported` (decision 10),
 * as is cell data inside AlternateContent. Structural metadata inside
 * AlternateContent is recorded as incomplete, never as known-empty (§3.4).
 */
export function sheetRecordEvents(meta: SheetMetadata, options: SheetRecordOptions): XmlEvents {
  const { partId } = options
  const stack: string[] = []
  type NsFrame = { local: string; uri?: string; scope: ReadonlyMap<string, string>; extensionUri?: string }
  type SparkGroupInfo = { sparkCount: number; incomplete: boolean }
  type SparkFrame = { node: XmlNode; local: string; uri?: string; scope: ReadonlyMap<string, string>; children: Map<string, XmlNode[]>; group?: SparkGroupInfo; textOverflow?: boolean }
  const nsStack: NsFrame[] = []
  const aggregateChildren = new Map<string, XmlNode[]>()
  let aggregateRoot: XmlNode | undefined
  let aggregateScope: ReadonlyMap<string, string> | undefined
  let captureRootDepth = -1
  let captureFrames: Array<SparkFrame | undefined> = []
  let sparkGroupCount = 0
  let groupLimitExceeded = false
  let sparkCountExceeded = false
  let pointLimitExceeded = false
  let sparkPoints = 0
  let stopSparkCapture = false
  const sparkDiagnostics: SparklineDiagnostic[] = []
  const sparkAttrNames = new Set(['type', 'lineWeight', 'displayEmptyCellsAs', 'markers', 'high', 'low', 'first', 'last', 'negative', 'displayXAxis', 'displayHidden', 'dateAxis', 'rightToLeft', 'minAxisType', 'maxAxisType', 'manualMin', 'manualMax'])
  const colorAttrNames = new Set(['rgb', 'theme', 'indexed', 'tint', 'auto'])
  /** The eight colour elements parseSparklineGroups reads; any other `color*` name is never captured (PR #49 review). */
  const sparkColorNames = new Set(['colorSeries', 'colorNegative', 'colorAxis', 'colorMarkers', 'colorFirst', 'colorLast', 'colorHigh', 'colorLow'])
  const sparkDiagnostic = (feature: string, message: string) => {
    if (!sparkDiagnostics.some(item => item.feature === feature && item.message === message)) sparkDiagnostics.push({ kind: 'unsupported', feature, message })
  }
  const childAllowed = (parent: SparkFrame, local: string, uri: string | undefined): boolean => {
    if (parent.local === 'sparklineGroups') return local === 'sparklineGroup' && uri === X14_SPARKLINE_NS
    if (parent.local === 'sparklineGroup') return (uri === X14_SPARKLINE_NS && (local === 'sparklines' || sparkColorNames.has(local))) || (uri === XM_SPARKLINE_NS && local === 'f')
    if (parent.local === 'sparklines') return local === 'sparkline' && uri === X14_SPARKLINE_NS
    if (parent.local === 'sparkline') return uri === XM_SPARKLINE_NS && (local === 'f' || local === 'sqref')
    return false
  }
  const singletonSparkChild = (parent: SparkFrame, local: string, uri: string | undefined): boolean =>
    (parent.local === 'sparklineGroup' && (local === 'sparklines' || sparkColorNames.has(local) || (uri === XM_SPARKLINE_NS && local === 'f')))
    || (parent.local === 'sparkline' && uri === XM_SPARKLINE_NS && (local === 'f' || local === 'sqref'))
  const attachSparkChild = (parent: SparkFrame, child: SparkFrame) => {
    const list = parent.children.get(child.local) ?? []
    list.push(child.node)
    parent.children.set(child.local, list)
    parent.node[child.local] = list.length === 1 ? child.node : list
  }
  const finishSparkFrame = (frame: SparkFrame) => registerNamespaceAware(frame.node, frame.uri, frame.scope, frame.children)
  const finalizeSparklines = () => {
    if (!aggregateRoot) return
    const groups = parseSparklineGroups(aggregateRoot, options.sparklineColorContext ?? {}, sparkDiagnostics)
    for (const group of groups) {
      const info = sparkGroupInfos[groupIndex++]
      if (info?.incomplete) group.incomplete = true
    }
    if (groupLimitExceeded) sparkDiagnostic('sparkline-group-limit', `Worksheet exceeds ${MAX_SPARKLINE_GROUPS_PER_SHEET} sparkline groups`)
    if (sparkCountExceeded) sparkDiagnostic('sparkline-count-limit', `Sparkline group exceeds ${MAX_SPARKLINES_PER_GROUP} entries`)
    if (pointLimitExceeded) sparkDiagnostic('sparkline-sheet-point-limit', `Worksheet exceeds ${MAX_POINTS_PER_SHEET} sparkline source points`)
    if (groups.length) meta.sparklineGroups = groups
    if (sparkDiagnostics.length) {
      meta.sparklineDiagnostics = [...sparkDiagnostics]
      for (const diagnostic of sparkDiagnostics) options.onSparklineDiagnostic?.(diagnostic)
    }
  }
  const sparkGroupInfos: SparkGroupInfo[] = []
  let groupIndex = 0
  let mcDepth = 0
  let paneSeen = false, sheetViews = 0
  let inSheetData = false
  let rowIndex = -1
  let lastRow = -1
  let lastCol = -1
  let rowOpen = false
  let rowBehind = false
  // cell state
  let inCell = false
  let cRow = 0, cCol = 0, cT: string | undefined, cS = 0, cCm: string | undefined, cCa = false
  let fSeen = false, fText = '', fAttrs: Record<string, string> = {}
  let vSeen = false, vText = ''
  let isSeen = false, isText = '', isTDirectSeen = false, runTSeen = false
  let collect: 'f' | 'v' | 'is' | undefined
  let decoder = new XstringDecoder()
  let inPhonetic = 0
  let cellSkipped = false
  // declared filters
  type Auto = NonNullable<ReturnType<NonNullable<SheetRecordOptions['declared']>['autoFilter']>>
  type Sort = NonNullable<ReturnType<NonNullable<SheetRecordOptions['declared']>['sortState']>>
  let autoDepth = -1, auto: Auto | undefined, autoSkip = false
  let columnAttrs: Record<string, string> | undefined, columnRules = 0
  let firstRule: { tag: string; attrs: Record<string, string>; filterCount: number; dateGroupItemCount: number; customFilters: Record<string, string>[]; customFilterCount: number } | undefined
  let ruleDepth = -1
  let sort: Sort | undefined, sortDepth = -1

  const cap = options.maxStructuralRecords ?? Number.MAX_SAFE_INTEGER
  const bounded = (list: unknown[], what: string) => {
    if (list.length >= cap) throw new PackageError('resource-limit', `more than ${cap} ${what} in one sheet`, partId)
  }
  const structural = (domain: SheetMetadata['structuralIncomplete'][number]) => {
    if (!meta.structuralIncomplete.includes(domain)) meta.structuralIncomplete.push(domain)
  }
  const finishCell = (): void => {
    inCell = false
    if (cellSkipped) return
    const t = cT
    let kind = K_EMPTY, num = 0, text = ''
    if (t === 's') {
      const idx = vSeen ? parseInt(vText, 10) : NaN
      if (Number.isFinite(idx) && idx >= 0 && idx < options.sharedStringCount) { kind = K_SST; num = idx } else kind = K_TEXT
    } else if (t === 'inlineStr') {
      kind = K_TEXT; text = isSeen ? isText : ''
    } else if (t === 'b') {
      kind = K_BOOL; num = vSeen && vText === '1' ? 1 : 0
    } else if (t === 'str') {
      kind = K_TEXT; text = vSeen ? vText : ''
    } else if (t === 'e') {
      if (vSeen) { kind = K_ERROR; text = vText }
    } else {
      const raw = vSeen ? vText : ''
      const n = parseFloat(raw)
      if (raw !== '' && Number.isFinite(n)) { kind = K_NUMBER; num = n } else if (raw !== '') { kind = K_TEXT; text = raw }
    }
    const cached = (vSeen && vText !== '') || isSeen || (t === 'str' && vSeen && vText === '')
    let record: CellRecord['meta']
    if (fSeen || cCa) {
      record = {}
      if (fSeen) {
        if (fText !== '') record.formula = fText
        const ft = fAttrs.t
        if (typeof ft === 'string' && ft !== '') record.formulaType = ft
        if (fAttrs.ca === '1' || fAttrs.ca === 'true') cCa = true
        if (ft === 'array' && fAttrs.ref) record.arrayRef = fAttrs.ref
        if (ft === 'shared') {
          const si = parseInt(fAttrs.si ?? '0', 10)
          record.sharedFormula = { si: Number.isFinite(si) ? si : 0, ...(fAttrs.ref ? { ref: fAttrs.ref } : {}) }
        }
      }
      if (/^\d+$/.test(cCm ?? '') && options.dynamicIndices.has(Number(cCm)) && record.formulaType === 'array') record.dynamicArray = true
      if (cCa) record.ca = true
    }
    options.onCell({ row: cRow, col: cCol, kind, styleIndex: cS, num, text, cached, valueIsError: t === 'e', meta: record })
  }

  return {
    open(qname, names, values, count, selfClosing) {
      const local = localName(qname)
      const nsDepth = nsStack.length
      const parentScope = nsStack[nsDepth - 1]?.scope ?? new Map<string, string>()
      let scope: ReadonlyMap<string, string> = parentScope
      for (let i = 0; i < count; i++) {
        const attr = names[i]
        if (attr !== 'xmlns' && !attr.startsWith('xmlns:')) continue
        if (scope === parentScope) scope = new Map(parentScope)
        ;(scope as Map<string, string>).set(attr === 'xmlns' ? '' : attr.slice(6), values[i])
      }
      const colon = qname.indexOf(':')
      const uri = scope.get(colon < 0 ? '' : qname.slice(0, colon))
      const extensionUri = local === 'ext' ? attrMap(names, values, count).uri : undefined
      const nsParent = nsStack[nsDepth - 1]
      const nsGrandparent = nsStack[nsDepth - 2]
      const isSparklineExtension = local === 'sparklineGroups' && uri === X14_SPARKLINE_NS
        && nsParent?.local === 'ext' && nsParent.uri === SPREADSHEETML_NS && nsParent.extensionUri === SPARKLINE_EXTENSION_URI
        && nsGrandparent?.local === 'extLst' && nsGrandparent.uri === SPREADSHEETML_NS
        && nsStack[0]?.local === 'worksheet' && nsStack[0]?.uri === SPREADSHEETML_NS
        && nsStack[1] === nsGrandparent

      const nsFrame: NsFrame = { local, uri, scope, ...(extensionUri !== undefined ? { extensionUri } : {}) }
      if (isSparklineExtension && captureRootDepth < 0) {
        aggregateRoot ??= { '@attrs': {} }
        aggregateScope ??= scope
        captureRootDepth = nsDepth
        const root: SparkFrame = { node: { '@attrs': {} }, local, uri, scope, children: new Map() }
        captureFrames = [root]
        if (selfClosing) {
          finishSparkFrame(root)
          for (const [name, children] of root.children) {
            const list = aggregateChildren.get(name) ?? []
            list.push(...children); aggregateChildren.set(name, list)
            ;(aggregateRoot as XmlNode)[name] = list.length === 1 ? list[0] : list
          }
          captureRootDepth = -1; captureFrames = []
        }
      } else if (captureRootDepth >= 0 && nsDepth > captureRootDepth) {
        const parent = captureFrames[captureFrames.length - 1]
        let frame: SparkFrame | undefined
        if (parent && childAllowed(parent, local, uri)) {
          // These children are read with first-match semantics by the resident
          // parser. Ignore later copies before node allocation so uncharged
          // duplicate payload cannot bypass the sparkline/point caps.
          let skip = singletonSparkChild(parent, local, uri) && parent.children.has(local)
          let group = parent.group
          if (parent.local === 'sparklineGroups' && local === 'sparklineGroup') {
            if (sparkGroupCount >= MAX_SPARKLINE_GROUPS_PER_SHEET) groupLimitExceeded = true
            else {
              group = { sparkCount: 0, incomplete: false }
              sparkGroupCount++
              sparkGroupInfos.push(group)
            }
          }
          if (parent.local === 'sparklines' && local === 'sparkline' && group) {
            group.sparkCount++
            if (group.sparkCount > MAX_SPARKLINES_PER_GROUP) {
              group.incomplete = true; sparkCountExceeded = true
              skip = true
            }
            if (stopSparkCapture) { group.incomplete = true; skip = true }
          }
          if (!skip && !(parent.local === 'sparklineGroups' && local === 'sparklineGroup' && !group)) {
            const node: XmlNode = { '@attrs': {} }
            const allowedAttrs = local === 'sparklineGroup' ? sparkAttrNames : sparkColorNames.has(local) ? colorAttrNames : undefined
            if (allowedAttrs) {
              const attrs = node['@attrs'] as Record<string, string>
              for (let i = 0; i < count; i++) {
                const key = localName(names[i])
                if (!allowedAttrs.has(key)) continue
                if (values[i].length <= 512) attrs[key] = values[i]
                else if (local === 'sparklineGroup') {
                  if (group) group.incomplete = true
                  sparkDiagnostic('sparkline-attribute-limit', 'Sparkline group attribute exceeds the streaming text limit')
                }
              }
            }
            frame = { node, local, uri, scope, children: new Map(), group }
            attachSparkChild(parent, frame)
            if (local === 'sparklineGroups') aggregateRoot = node
          }
        }
        if (!selfClosing) captureFrames.push(frame)
        else if (frame) finishSparkFrame(frame)
      }
      if (!selfClosing) nsStack.push(nsFrame)

      const depth = stack.length
      if (local === 'AlternateContent') { if (!selfClosing) mcDepth++; return }
      if (mcDepth > 0) {
        if (local === 'sheetData' || local === 'row' || local === 'c') {
          throw new PackageError('unsupported', `${local} inside mc:AlternateContent is not supported; the sheet would load blank`, partId)
        }
        const domain = STRUCTURE[local]
        if (domain !== undefined) structural(domain)
        if (!selfClosing) stack.push(local)
        return
      }
      if (local === 'sheetData' && depth === 1) inSheetData = !selfClosing
      else if (inSheetData && local === 'row') {
        const a = attrMap(names, values, count)
        const r = a.r === undefined ? NaN : parseInt(a.r, 10) - 1
        if (!Number.isSafeInteger(r) || r < 0) throw new PackageError('unsupported', 'a worksheet row has no valid row number', partId)
        // A row at or before the last one may only carry row metadata (the
        // legacy model keeps it); a cell in it is refused below.
        rowBehind = r <= lastRow
        if (rowBehind) meta.unorderedRows = true
        rowIndex = r
        if (!rowBehind) lastRow = r
        lastCol = -1
        if (r <= MAX_COORD_ROW) {
          const ht = a.ht !== undefined ? parseFloat(a.ht) : undefined
          if (ht !== undefined && Number.isFinite(ht) && ht >= 0) meta.rowHeights.push({ row: r, heightPt: ht })
          if (a.hidden === '1' || a.hidden === 'true') {
            const last = meta.hiddenRows[meta.hiddenRows.length - 1]
            if (last && last.r1 + 1 === r) last.r1 = r
            else meta.hiddenRows.push({ r0: r, r1: r })
          }
        }
        rowOpen = !selfClosing
        if (selfClosing && !rowBehind) options.onRowEnd(r)
      } else if (rowOpen && local === 'c' && stack[stack.length - 1] === 'row') {
        const a = attrMap(names, values, count)
        const ref = a.r === undefined ? undefined : refOf(a.r)
        if (ref === undefined) throw new PackageError('unsupported', `a cell has no valid reference in row ${rowIndex + 1}`, partId)
        if (rowBehind) throw new PackageError('unsupported', `worksheet rows are out of order or duplicated at row ${rowIndex + 1}`, partId)
        if (ref[0] !== rowIndex) throw new PackageError('unsupported', `cell ${a.r} is not in its row ${rowIndex + 1}`, partId)
        if (ref[1] <= lastCol) throw new PackageError('unsupported', `cells are out of order or duplicated at ${a.r}`, partId)
        lastCol = ref[1]
        cRow = ref[0]; cCol = ref[1]
        cellSkipped = cRow > MAX_COORD_ROW || cCol > MAX_COORD_COL || cCol < 0
        if (cellSkipped) meta.skippedCells++
        cT = a.t
        const s = parseInt(a.s ?? '0', 10)
        cS = Number.isFinite(s) ? s : 0
        cCm = a.cm
        cCa = a.ca === '1' || a.ca === 'true'
        fSeen = false; fText = ''; fAttrs = {}
        vSeen = false; vText = ''
        isSeen = false; isText = ''; isTDirectSeen = false
        collect = undefined
        inPhonetic = 0
        inCell = true
        if (selfClosing) finishCell()
      } else if (inCell) {
        const parent = stack[stack.length - 1]
        if (local === 'f' && parent === 'c' && !fSeen) {
          fSeen = true; fAttrs = attrMap(names, values, count)
          if (!selfClosing) collect = 'f'
        } else if (local === 'v' && parent === 'c' && !vSeen) {
          vSeen = true
          if (!selfClosing) { collect = 'v'; decoder = new XstringDecoder() }
        } else if (local === 'is' && parent === 'c' && !isSeen) {
          isSeen = true
        } else if (local === 'rPh' && !selfClosing) inPhonetic++
        else if (local === 'r' && parent === 'is') runTSeen = false
        else if (local === 't' && isSeen && inPhonetic === 0 && !selfClosing) {
          if (parent === 'is' ? !isTDirectSeen : (parent === 'r' && stack[stack.length - 2] === 'is' && !runTSeen)) {
            if (parent === 'is') isTDirectSeen = true
            else runTSeen = true
            collect = 'is'
            decoder = new XstringDecoder()
          }
        }
      } else if (depth === 1 && local === 'sheetFormatPr') {
        const format = sheetFormat(attrMap(names, values, count))
        if (format.baseColWidth !== undefined || format.defaultColWidth !== undefined) meta.format = format
      } else if (depth === 1 && local === 'cols') {
        // columns are collected below
      } else if (depth === 2 && local === 'col' && stack[1] === 'cols') {
        const a = attrMap(names, values, count)
        if (meta.cols.length >= 16_384) throw new PackageError('resource-limit', 'more than 16,384 column definitions', partId)
        meta.cols.push({
          min: parseInt(a.min ?? '1', 10) - 1, max: parseInt(a.max ?? '1', 10) - 1,
          widthChars: a.width !== undefined ? parseFloat(a.width) : undefined, hidden: a.hidden === '1' || a.hidden === 'true',
        })
      } else if (depth === 2 && local === 'mergeCell' && stack[1] === 'mergeCells') {
        const ref = attrMap(names, values, count).ref
        if (ref) {
          bounded(meta.merges, 'merges')
          meta.merges.push(ref)
        }
      } else if (depth === 2 && local === 'tablePart' && stack[1] === 'tableParts') {
        const id = attrMap(names, values, count).id
        bounded(meta.tablePartIds, 'table parts')
        meta.tablePartIds.push(typeof id === 'string' && id !== '' ? id : undefined)
      } else if (depth === 2 && local === 'sheetView' && stack[1] === 'sheetViews') {
        // Only the first sheet view counts (as the resident parser reads it), even when it has no pane.
        sheetViews++
      } else if (depth === 3 && local === 'pane' && stack[1] === 'sheetViews' && stack[2] === 'sheetView' && sheetViews === 1 && !paneSeen) {
        paneSeen = true
        const frozen = frozenPaneFrom(attrMap(names, values, count))
        if (frozen) meta.frozen = frozen
      } else if (depth === 1 && local === 'drawing') {
        // legacyDrawing (comment VML) carries no anchors the structural query reports.
        bounded(meta.drawingIds, 'drawing references')
        meta.drawingRefs++
        const id = attrMap(names, values, count).id
        meta.drawingIds.push(typeof id === 'string' && id !== '' ? id : undefined)
      } else if (depth === 1 && local === 'legacyDrawing') {
        const ids = meta.legacyDrawingIds ??= []
        bounded(ids, 'legacy drawing references')
        const id = attrMap(names, values, count).id
        ids.push(typeof id === 'string' && id !== '' ? id : undefined)
      } else if (options.declared && depth === 1 && local === 'autoFilter') {
        auto = options.declared.autoFilter(attrMap(names, values, count))
        autoSkip = auto === undefined
        if (!selfClosing) autoDepth = depth
        else { auto = undefined; autoSkip = false }
      } else if (autoDepth >= 0 && depth === autoDepth + 1 && local === 'filterColumn') {
        columnAttrs = attrMap(names, values, count); columnRules = 0; firstRule = undefined
        if (selfClosing) { auto?.addColumn({ attrs: columnAttrs, ruleCount: 0 }); columnAttrs = undefined }
      } else if (columnAttrs && depth === autoDepth + 2) {
        if (local !== 'extLst') {
          columnRules++
          if (columnRules === 1) {
            firstRule = { tag: local, attrs: attrMap(names, values, count), filterCount: 0, dateGroupItemCount: 0, customFilters: [], customFilterCount: 0 }
            if (!selfClosing) ruleDepth = depth
          }
        }
      } else if (firstRule && ruleDepth >= 0 && depth === ruleDepth + 1) {
        if (local === 'filter') firstRule.filterCount++
        else if (local === 'dateGroupItem') firstRule.dateGroupItemCount++
        else if (local === 'customFilter') {
          firstRule.customFilterCount++
          if (firstRule.customFilters.length < 2) firstRule.customFilters.push(attrMap(names, values, count))
        }
      } else if (options.declared && local === 'sortState' && (depth === 1 || (autoDepth >= 0 && depth === autoDepth + 1))) {
        const nested = depth !== 1
        sort = nested && autoSkip ? undefined : options.declared.sortState(attrMap(names, values, count), nested ? 'nested' : 'direct')
        if (!selfClosing) sortDepth = depth
        else sort = undefined
      } else if (sortDepth >= 0 && depth === sortDepth + 1 && local === 'sortCondition') {
        sort?.addCondition(attrMap(names, values, count))
      }
      if (!selfClosing) stack.push(local)
    },
    close(qname) {
      const local = localName(qname)
      const nsDepth = nsStack.length - 1
      if (captureRootDepth >= 0) {
        if (nsDepth > captureRootDepth) {
          const frame = captureFrames.pop()
          if (frame) {
            finishSparkFrame(frame)
            if (frame.local === 'sparkline') {
              const fNode = frame.node.f as XmlNode | undefined
              const formula = typeof fNode?.['#text'] === 'string' ? fNode['#text'] as string : undefined
              let quoted = false, validFormula = true
              if (formula) for (let i = 0; i < formula.length; i++) {
                if (formula[i] === "'") {
                  if (quoted && formula[i + 1] === "'") { i++; continue }
                  quoted = !quoted
                } else if (!quoted && /\s/.test(formula[i])) { validFormula = false; break }
              }
              const area = formula && validFormula && !quoted ? parseA1Area(formula) : undefined
              if (area && !area.definedName && !area.sheet?.startsWith('[') && !(area.r0 !== area.r1 && area.c0 !== area.c1)) {
                const points = (area.r1 - area.r0 + 1) * (area.c1 - area.c0 + 1)
                if (points <= MAX_POINTS_PER_SPARKLINE) {
                  if (sparkPoints + points > MAX_POINTS_PER_SHEET) {
                    pointLimitExceeded = true; stopSparkCapture = true
                    if (frame.group) frame.group.incomplete = true
                  } else sparkPoints += points
                }
              }
            }
          }
        } else if (nsDepth === captureRootDepth) {
          const root = captureFrames.pop()
          if (root) finishSparkFrame(root)
          for (const [name, children] of root?.children ?? []) {
            const list = aggregateChildren.get(name) ?? []
            list.push(...children); aggregateChildren.set(name, list)
            if (aggregateRoot) aggregateRoot[name] = list.length === 1 ? list[0] : list
          }
          captureRootDepth = -1; captureFrames = []
        }
      }
      if (nsStack.length) nsStack.pop()
      if (local === 'AlternateContent' && mcDepth > 0) { mcDepth--; return }
      stack.pop()
      if (mcDepth > 0) return
      const depth = stack.length
      if (local === 'sheetData' && depth === 1) inSheetData = false
      else if (local === 'row' && rowOpen) { rowOpen = false; if (!rowBehind) options.onRowEnd(rowIndex) }
      else if (inCell && local === 'c' && depth === stack.length && stack[stack.length - 1] === 'row') finishCell()
      else if (inCell) {
        if (local === 'f' && collect === 'f') collect = undefined
        else if (local === 'v' && collect === 'v') { vText += decoderTail(); collect = undefined }
        else if (local === 't' && collect === 'is') { isText += decoder.push('', true); collect = undefined }
        else if (local === 'rPh' && inPhonetic > 0) inPhonetic--
      } else if (local === 'autoFilter' && depth === autoDepth) { auto = undefined; autoSkip = false; autoDepth = -1 }
      else if (local === 'filterColumn' && columnAttrs && depth === autoDepth + 1) {
        auto?.addColumn({ attrs: columnAttrs, ruleCount: columnRules, ...(firstRule ? { firstRule } : {}) })
        columnAttrs = undefined; firstRule = undefined; ruleDepth = -1
      } else if (ruleDepth >= 0 && depth === ruleDepth) ruleDepth = -1
      else if (local === 'sortState' && depth === sortDepth) { sort = undefined; sortDepth = -1 }
      if (local === 'worksheet' && nsDepth === 0) {
        if (aggregateRoot && aggregateScope) registerNamespaceAware(aggregateRoot, X14_SPARKLINE_NS, aggregateScope, aggregateChildren)
        finalizeSparklines()
      }
    },
    text(text) {
      const sparkFrame = captureFrames[captureFrames.length - 1]
      if (sparkFrame && sparkFrame.uri === XM_SPARKLINE_NS && (sparkFrame.local === 'f' || sparkFrame.local === 'sqref') && !sparkFrame.textOverflow) {
        const current = typeof sparkFrame.node['#text'] === 'string' ? sparkFrame.node['#text'] as string : ''
        if (current.length + text.length <= 256) sparkFrame.node['#text'] = current + text
        else {
          sparkFrame.node['#text'] = ''
          sparkFrame.textOverflow = true
          if (sparkFrame.group) sparkFrame.group.incomplete = true
          sparkDiagnostic('sparkline-data-reference', 'Sparkline reference exceeds the streaming text limit')
        }
      }
      if (!inCell || collect === undefined) return
      // str/inline text is decoded (_xHHHH_); numeric, boolean and error <v> text is not (parse.ts).
      if (collect === 'f') {
        fText += text
        if (fText.length > FORMULA_TEXT_LIMIT) throw new PackageError('resource-limit', 'a formula exceeds 1,048,576 characters', partId)
      } else if (collect === 'v') {
        vText += cT === 'str' ? decoder.push(text) : text
        if (vText.length > options.maxTextChars) throw new PackageError('resource-limit', `a cell value exceeds ${options.maxTextChars} characters`, partId)
      } else {
        isText += decoder.push(text)
        if (isText.length > options.maxTextChars) throw new PackageError('resource-limit', `a cell value exceeds ${options.maxTextChars} characters`, partId)
      }
    },
  }
  function decoderTail(): string { return cT === 'str' ? decoder.push('', true) : '' }
}

export function newSheetMetadata(): SheetMetadata {
  return { rowHeights: [], hiddenRows: [], cols: [], merges: [], tablePartIds: [], drawingRefs: 0, drawingIds: [], structuralIncomplete: [], skippedCells: 0 }
}

/** Streams a drawing part's top-level anchors into placement-only parts (S3.3):
 * first xdr:from/xdr:to/xdr:ext/xdr:pos child of each anchor, nothing else is
 * retained. A top-level mc:AlternateContent, or a document element other than
 * wsDr, reports `onUnreadable` (fail-closed), matching the object parser. */
/** Internal admission seam: absolute conservative ownership, before retention.
 * Strings count as 2 bytes per UTF-16 unit plus a header, including borrowed
 * names/attributes. This models logical ownership, not an engine's ropes/heap. */
export interface AnchorConstructionOptions {
  need?(bytes: number): void
}
const ANCHOR_CONSUMER_BYTES = 256
const ANCHOR_HOLDER_BYTES = 64
const ANCHOR_PROPERTY_BYTES = 48
const anchorStringBytes = (chars: number) => 48 + 2 * chars

export function anchorRecordEvents(onAnchor: (name: string, parts: AnchorParts) => void, onUnreadable: () => void, options: AnchorConstructionOptions = {}): XmlEvents & { dispose(): void } {
  options.need?.(ANCHOR_CONSUMER_BYTES)
  let held = ANCHOR_CONSUMER_BYTES
  let disposed = false
  let depth = 0
  let anchor: { name: string; parts: AnchorParts } | undefined
  /** A non-wsDr document element: one unreadable report, nothing beneath it is collected. */
  let foreign = false
  let section: Record<string, string> | undefined
  let field: string | undefined
  let text = ''
  let textBytes = 0
  const need = (bytes: number) => {
    if (disposed) throw new PackageError('cancelled', 'drawing marker consumer already disposed')
    options.need?.(bytes)
  }
  const grow = (bytes: number) => { need(held + bytes); held += bytes }
  const clear = () => { anchor = undefined; section = undefined; field = undefined; text = ''; textBytes = 0 }
  const finish = () => {
    try { onAnchor(anchor!.name, anchor!.parts) }
    finally {
      // The callback performs Number conversion while every lexical field is
      // still charged. Drop the last consumer references before returning bytes.
      clear()
      held = ANCHOR_CONSUMER_BYTES
      need(held)
    }
  }
  return {
    dispose() { clear(); disposed = true },
    open(qname, names, values, count, selfClosing) {
      need(held)
      const local = localName(qname)
      if (depth === 0) { foreign = local !== DRAWING_ROOT; if (foreign) onUnreadable() }
      if (foreign) { if (!selfClosing) depth++; return }
      if (depth === 1) {
        if (ANCHOR_ELEMENTS.has(local)) {
          grow(2 * ANCHOR_HOLDER_BYTES + anchorStringBytes(local.length))
          anchor = { name: local, parts: {} }
        } else if (local === 'AlternateContent') onUnreadable()
        if (selfClosing && anchor) finish()
      } else if (anchor && depth === 2) {
        if ((local === 'from' || local === 'to') && !anchor.parts[local]) {
          grow(ANCHOR_HOLDER_BYTES + ANCHOR_PROPERTY_BYTES + anchorStringBytes(local.length))
          section = anchor.parts[local] = {}
        } else if ((local === 'ext' || local === 'pos') && !anchor.parts[local]) {
          let bytes = ANCHOR_HOLDER_BYTES + ANCHOR_PROPERTY_BYTES + anchorStringBytes(local.length)
          for (let i = 0; i < count; i++) bytes += ANCHOR_PROPERTY_BYTES + anchorStringBytes(localName(names[i]).length) + anchorStringBytes(values[i].length)
          grow(bytes)
          anchor.parts[local] = attrMap(names, values, count)
        }
      } else if (section && depth === 3 && !selfClosing && !(local in section)) {
        grow(anchorStringBytes(local.length) + anchorStringBytes(0))
        field = local
        text = ''
        textBytes = anchorStringBytes(0)
      }
      if (!selfClosing) depth++
    },
    close() {
      need(held)
      depth--
      if (depth === 3 && field !== undefined && section) {
        grow(ANCHOR_PROPERTY_BYTES)
        section[field] = text
        field = undefined
        text = ''
        textBytes = 0
      } else if (depth === 2) section = undefined
      else if (depth === 1 && anchor) finish()
    },
    text(chunk) {
      need(held)
      if (field === undefined) return
      const nextBytes = anchorStringBytes(text.length + chunk.length)
      // Admit both the old continuation and its replacement before concat.
      need(held + nextBytes)
      text += chunk
      held += nextBytes - textBytes
      textBytes = nextBytes
      need(held)
    },
  }
}
