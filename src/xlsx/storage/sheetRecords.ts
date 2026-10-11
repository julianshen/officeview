/**
 * SpreadsheetML token-stream consumers (design note §3.4, §4.3, §7.7).
 *
 * S1.1b provides the admission counter used by the verification pass (gate
 * G2): it sees every token of every XML part exactly once and keeps only
 * counters and a bounded formula-text buffer, so nothing proportional to the
 * part is retained. The paged record builders (S1.2) will live next to it.
 */
import { PackageError } from '../../core/package-error'
import { localName, type XmlEvents } from '../../core/xml-stream'

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
}

export function newAdmissionCounters(): AdmissionCounters {
  return {
    expandedBytes: 0, maxPartBytes: 0, elements: 0, attributes: 0, textChars: 0, rows: 0, cells: 0, uniqueStrings: 0,
    longestString: 0, merges: 0, styleRecords: 0, maxAttributesPerElement: 0, loadEvaluatedFormulas: 0, fullCalcOnLoad: false,
    alternateContentSubtrees: 0, structuralIncomplete: [],
  }
}

/** Functions that force evaluation on load whatever the cache says (design note §4.3 rule 1). */
const VOLATILE = /(^|[^A-Za-z0-9_.])(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|OFFSET|INDIRECT|CELL|INFO)\s*\(/i
/** Excel formulas are at most 8,192 characters; anything longer is treated as volatile. */
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
}

/**
 * Count one XML part. The part kind is taken from its root element, never
 * from its name. Data (`sheetData`/`row`/`c`, `si`) inside AlternateContent is
 * refused with a typed `unsupported`, because the legacy tree walk would drop
 * it silently and publish a blank sheet (§3.4, spec §4.3.8).
 */
export function countingEvents(counters: AdmissionCounters, partId: string, options: CountingOptions): XmlEvents {
  type Kind = 'worksheet' | 'sst' | 'styleSheet' | 'workbook' | 'other'
  let kind: Kind | undefined
  const stack: string[] = []
  let mcDepth = 0
  let inSheetData = false
  // cell state
  let inCell = false
  let cellHasValue = false
  let cellCalcAlways = false
  let cellHasFormula = false
  let formulaText = ''
  let formulaVolatile = false
  let stringChars = 0
  let inString: 'si' | 'is' | 'v' | undefined
  let inPhonetic = 0

  const finishCell = (): void => {
    if (cellHasFormula && (counters.fullCalcOnLoad || !cellHasValue || cellCalcAlways || formulaVolatile || VOLATILE.test(formulaText))) {
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
        throw new PackageError('unsupported', `${local} inside mc:AlternateContent is not supported; the sheet would load blank`, partId)
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
          cellHasFormula = false
          formulaText = ''
          formulaVolatile = false
          cellCalcAlways = TRUE(attr(names, values, count, 'ca'))
          if (selfClosing) finishCell()
        } else if (inCell && local === 'f') {
          cellHasFormula = true
          if (TRUE(attr(names, values, count, 'ca'))) cellCalcAlways = true
        } else if (inCell && local === 'v') {
          cellHasValue = true
          if (!selfClosing) inString = 'v'
        } else if (inCell && local === 'is') {
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
      // shared/inline strings: t and r/t text, never phonetic runs (parse.ts sharedStrings)
      if (inString !== undefined && inPhonetic === 0 && (inString === 'v' ? top === 'v' : top === 't')) stringChars += text.length
      options.check(counters, stringChars)
    },
  }
}
