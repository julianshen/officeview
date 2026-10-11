/**
 * SpreadsheetML token-stream consumers (design note §3.4, §4.3, §7.7).
 *
 * S1.1b provides the admission counter used by the verification pass (gate
 * G2): it sees every token of every XML part exactly once and keeps only
 * counters and a bounded formula-text buffer, so nothing proportional to the
 * part is retained. The paged record builders (S1.2) will live next to it.
 */
import { PackageError } from '../../core/package-error'
import { localName, XstringDecoder, type XmlEvents } from '../../core/xml-stream'

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

/** Functions that force evaluation on load whatever the cache says (design note §4.3 rule 1). */
const VOLATILE = /(^|[^A-Za-z0-9_.])(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|OFFSET|INDIRECT|CELL|INFO)\s*\(/i
/** Formula text kept for the volatile-function scan. Excel caps formulas at
 * 8,192 characters; twice that is kept as margin, and anything longer is
 * treated as volatile (fail-closed). */
const MAX_FORMULA_CHARS = 16_384

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
    if (cellHasFormula && (counters.fullCalcOnLoad || !cached || cellCalcAlways || formulaVolatile || VOLATILE.test(formulaText))) {
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
  merges: string[]
  tablePartIds: (string | undefined)[]
  drawingRefs: number
  drawingIds: (string | undefined)[]
  /** Anchors found in the referenced drawing parts; unreadable parts count as one (fail-closed). */
  drawingAnchors: number
  structuralIncomplete: Array<'merges' | 'tables' | 'drawings' | 'columns' | 'rows'>
  /** Cells whose coordinates fall outside the sheet (invisible in both adapters). */
  skippedCells: number
  /** Metadata-only rows arrived out of order: rowHeights/hiddenRows need normalizeRowMetadata. */
  unorderedRows?: boolean
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
  let mcDepth = 0
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
      } else if (depth === 1 && local === 'drawing') {
        // legacyDrawing (comment VML) carries no anchors the structural query reports.
        bounded(meta.drawingIds, 'drawing references')
        meta.drawingRefs++
        const id = attrMap(names, values, count).id
        meta.drawingIds.push(typeof id === 'string' && id !== '' ? id : undefined)
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
    },
    text(text) {
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
  return { rowHeights: [], hiddenRows: [], cols: [], merges: [], tablePartIds: [], drawingRefs: 0, drawingIds: [], drawingAnchors: 0, structuralIncomplete: [], skippedCells: 0 }
}

const ANCHORS = new Set(['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'])
/** Counts top-level anchors of a drawing part without retaining it. */
export function anchorCountEvents(onAnchor: () => void): XmlEvents {
  let depth = 0
  return {
    open(qname, _n, _v, _c, selfClosing) {
      if (depth === 1 && ANCHORS.has(localName(qname))) onAnchor()
      if (!selfClosing) depth++
    },
    close() { depth-- },
    text() {},
  }
}
