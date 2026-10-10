/**
 * C4 phase 1 — calculation intent/policy and Unicode settings.
 *
 * SCOPE (bounded, phase 1 only):
 *  - Calculation intent: file-load (default) / explicit-recalc / force.
 *  - calcMode auto / manual / autoNoTable. autoNoTable does not itself force
 *    recalculation: the established file-load cache policy is identical for
 *    auto and autoNoTable. The mode only excludes the actual `<f
 *    t="dataTable">` what-if case — never structured-table refs (which are
 *    ordinary formulas recalculated on an explicit request, with calcChain
 *    part availability irrelevant).
 *  - Unicode settings precedence (explicit option > verified metadata >
 *    parsed-absent 1 / standalone 2) with tagged provenance, and version-
 *    consistent LEN/LEFT/RIGHT/MID code-unit vs code-point slicing.
 *
 * EXPLICITLY NOT CLAIMED HERE (phase 2, held):
 *  - Iterative SCC evaluation initialization/order/convergence metric. Root's
 *    current native read-only facts are iteration OFF (app-level toggle; no
 *    global mutation permitted). No zero/cache seed or Jacobi/Gauss-Seidel
 *    order is invented; existing cycle behavior is preserved unchanged.
 *
 * Expectation provenance:
 *  - MEASURED/PRIMARY : Microsoft compatibility-version facts — Version 1
 *    counts a surrogate pair as 2, Version 2 as 1; variation selectors and
 *    modifiers are separate code points (no grapheme clustering).
 *    LEN("😀")=1 is Version-2 evidence only.
 *  - POLICY           : file-load/explicit-recalc/force/autoNoTable intent
 *    rules and what-if data-table retention (reference-contract-final).
 *  - MODEL            : synthetic parsable formula text on a what-if cell
 *    (the policy, not Excel's TABLE() computation, is under test).
 */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import type { CalcSettings, TableMetadata } from '../src/xlsx/formula/types'
import type { XlsxCell, XlsxDocument } from '../src/xlsx/types'

const EMOJI = '😀' // U+1F600: one code point, one surrogate pair (two code units)

function calc(mode: CalcSettings['calcMode'], overrides: Partial<CalcSettings> = {}): CalcSettings {
  return { calcMode: mode, fullCalcOnLoad: false, iterate: false, iterateCount: 100, iterateDelta: 0.001, ...overrides }
}

/** Build one sheet from a flat cell list (row grouping is mechanical). */
function sheetOf(name: string, cells: XlsxCell[], workbookIndex = 0) {
  const byRow = new Map<number, XlsxCell[]>()
  for (const cell of cells) {
    const row = byRow.get(cell.row) ?? []
    row.push(cell)
    byRow.set(cell.row, row)
  }
  return {
    name,
    sheetId: name.toLowerCase(),
    workbookIndex,
    cols: [],
    merges: [],
    mergeRanges: [],
    rows: [...byRow.entries()].sort((a, b) => a[0] - b[0]).map(([index, rowCells]) => ({ index, cells: rowCells })),
  }
}

function docOf(cells: XlsxCell[], extra: Partial<XlsxDocument> = {}): XlsxDocument {
  return { sheets: [sheetOf('S', cells) as never], definedNames: [], tables: [], ...extra }
}

function cachedFormula(ref: string, col: number, row: number, formula: string, value: XlsxCell['value']): XlsxCell {
  return { ref, col, row, value, styleIndex: 0, hasCachedValue: true, formula }
}

function formulaCell(ref: string, col: number, row: number, formula: string): XlsxCell {
  return { ref, col, row, value: null, styleIndex: 0, hasCachedValue: false, formula }
}

// ---------------------------------------------------------------------------
// Unicode settings: precedence, provenance, version-consistent slicing
// ---------------------------------------------------------------------------
describe('C4 Unicode settings precedence and version-consistent text functions', () => {
  it('standalone (no workbook metadata) default is version 2 — code points', () => {
    const doc = docOf([formulaCell('A1', 0, 0, `LEN("${EMOJI}")`)])
    evaluateWorkbookFormulas(doc)
    expect(doc.sheets[0]!.rows[0]!.cells[0]!.value).toBe(1)
  })

  it('parsed workbook absent metadata (model version 1) drives the workbook context', () => {
    const doc = docOf([formulaCell('A1', 0, 0, `LEN("${EMOJI}")`)], {
      semantics: { dateSystem: '1900', unicode: { version: 1, source: 'workbook-default' }, locale: 'en-US', timeZone: 'UTC', epochNowMs: 0 },
    })
    evaluateWorkbookFormulas(doc)
    // Version 1 counts the surrogate pair as two UTF-16 code units.
    expect(doc.sheets[0]!.rows[0]!.cells[0]!.value).toBe(2)
  })

  it('explicit option wins over verified workbook metadata (version 1 over metadata 2)', () => {
    const doc = docOf([formulaCell('A1', 0, 0, `LEN("${EMOJI}")`)], {
      semantics: { dateSystem: '1900', unicode: { version: 2, source: 'verified-workbook' }, locale: 'en-US', timeZone: 'UTC', epochNowMs: 0 },
    })
    evaluateWorkbookFormulas(doc, { unicodeVersion: 1 })
    expect(doc.sheets[0]!.rows[0]!.cells[0]!.value).toBe(2)
  })

  it('explicit option provenance is tagged on the resolved semantics', () => {
    const doc = docOf([formulaCell('A1', 0, 0, `LEN("${EMOJI}")`)])
    evaluateWorkbookFormulas(doc, { unicodeVersion: 1 })
    expect(doc.semantics?.unicode).toEqual({ version: 1, source: 'explicit' })
  })

  it('parsed-absent metadata keeps its workbook-default provenance tag', () => {
    const doc = docOf([formulaCell('A1', 0, 0, '1+1')], {
      semantics: { dateSystem: '1900', unicode: { version: 1, source: 'workbook-default' }, locale: 'en-US', timeZone: 'UTC', epochNowMs: 0 },
    })
    evaluateWorkbookFormulas(doc)
    expect(doc.semantics?.unicode).toEqual({ version: 1, source: 'workbook-default' })
  })

  it('LEN/MID honor compatibility version while LEFT/RIGHT retain surrogate pairs', () => {
    const doc2 = docOf([
      formulaCell('A1', 0, 0, `LEFT("${EMOJI}x",1)`),
      formulaCell('B1', 1, 0, `RIGHT("x${EMOJI}",1)`),
      formulaCell('C1', 2, 0, `MID("a${EMOJI}b",2,1)`),
    ])
    evaluateWorkbookFormulas(doc2, { unicodeVersion: 2 })
    const [left2, right2, mid2] = doc2.sheets[0]!.rows[0]!.cells
    expect(left2!.value).toBe(EMOJI)
    expect(right2!.value).toBe(EMOJI)
    expect(mid2!.value).toBe(EMOJI)

    const doc1 = docOf([
      formulaCell('A1', 0, 0, `LEFT("${EMOJI}x",1)`),
      formulaCell('B1', 1, 0, `RIGHT("x${EMOJI}",1)`),
      formulaCell('C1', 2, 0, `MID("a${EMOJI}b",2,1)`),
    ])
    evaluateWorkbookFormulas(doc1, { unicodeVersion: 1 })
    const [left1, right1, mid1] = doc1.sheets[0]!.rows[0]!.cells
    expect(left1!.value).toBe(EMOJI)
    expect(right1!.value).toBe(EMOJI)
    expect(mid1!.value).toBe('\uD83D')
  })

  it('combining marks and variation selectors are separate code points, never graphemes', () => {
    const doc = docOf([
      formulaCell('A1', 0, 0, 'LEN("e\u0301")'), // e + combining acute
      formulaCell('B1', 1, 0, 'LEN("\u2764\uFE0F")'), // heart + variation selector-16
    ])
    evaluateWorkbookFormulas(doc, { unicodeVersion: 2 })
    const [combining, variation] = doc.sheets[0]!.rows[0]!.cells
    expect(combining!.value).toBe(2)
    expect(variation!.value).toBe(2)
  })
})

// ---------------------------------------------------------------------------
// Calculation intent / calcMode policy
// ---------------------------------------------------------------------------
describe('C4 calculation intent and calcMode cache policy', () => {
  it('file-load (default) preserves genuine cached false / empty-string / error identity', () => {
    const doc = docOf([
      { ref: 'A1', col: 0, row: 0, value: false, styleIndex: 0, hasCachedValue: true, formula: '1=2' },
      { ref: 'B1', col: 1, row: 0, value: '', styleIndex: 0, hasCachedValue: true, formula: 'IF(TRUE,"","x")' },
      { ref: 'C1', col: 2, row: 0, value: '#DIV/0!', styleIndex: 0, hasCachedValue: true, valueIsError: true, formula: '1/0' },
    ])
    evaluateWorkbookFormulas(doc, { intent: 'file-load' })
    const [falsy, empty, error] = doc.sheets[0]!.rows[0]!.cells
    expect(falsy!.value).toBe(false)
    expect(empty!.value).toBe('')
    expect(error!.value).toBe('#DIV/0!')
    expect(doc.diagnostics ?? []).toEqual([])
  })

  it('file-load does not sample the injected clock for a retained volatile cache', () => {
    let calls = 0
    const doc = docOf([cachedFormula('A1', 0, 0, 'NOW()', 99)])
    evaluateWorkbookFormulas(doc, { now: () => { calls++; return 1768478400000 }, timeZone: 'UTC' })
    expect(doc.sheets[0]!.rows[0]!.cells[0]!.value).toBe(99)
    expect(calls).toBe(0)
  })

  it('explicit-recalc recalculates a supported cached formula', () => {
    const doc = docOf([cachedFormula('A1', 0, 0, '1+1', 99)])
    evaluateWorkbookFormulas(doc, { intent: 'explicit-recalc' })
    expect(doc.sheets[0]!.rows[0]!.cells[0]!.value).toBe(2)
  })

  it('explicit-recalc samples the injected clock exactly once for a volatile cache', () => {
    let calls = 0
    const doc = docOf([cachedFormula('A1', 0, 0, 'NOW()', 99)])
    evaluateWorkbookFormulas(doc, { intent: 'explicit-recalc', now: () => { calls++; return 1768478400000 }, timeZone: 'UTC' })
    expect(calls).toBe(1)
    expect(doc.sheets[0]!.rows[0]!.cells[0]!.value).not.toBe(99)
  })

  it('forceRecalc overrides a manual calcMode for supported formulas', () => {
    const doc = docOf([cachedFormula('A1', 0, 0, '1+1', 99)], { calc: calc('manual') })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(doc.sheets[0]!.rows[0]!.cells[0]!.value).toBe(2)
  })

  it('manual calcMode retains a cached formula with no forced warning (no calc request blocked)', () => {
    const doc = docOf([cachedFormula('A1', 0, 0, '1+1', 99)], { calc: calc('manual') })
    evaluateWorkbookFormulas(doc)
    expect(doc.sheets[0]!.rows[0]!.cells[0]!.value).toBe(99)
    // A valid manual cached file-load is app policy, not a blocked request: no
    // forced warning is emitted.
    expect(doc.diagnostics ?? []).toEqual([])
  })

  it('manual calcMode + an explicit-recalc request recalculates supported formulas', () => {
    const doc = docOf([cachedFormula('A1', 0, 0, '1+1', 99)], { calc: calc('manual') })
    evaluateWorkbookFormulas(doc, { intent: 'explicit-recalc' })
    expect(doc.sheets[0]!.rows[0]!.cells[0]!.value).toBe(2)
  })

  it('fullCalcOnLoad recalculates cached ordinary formulas', () => {
    const doc = docOf([cachedFormula('A1', 0, 0, '1+1', 99)], { calc: calc('auto', { fullCalcOnLoad: true }) })
    evaluateWorkbookFormulas(doc)
    expect(doc.sheets[0]!.rows[0]!.cells[0]!.value).toBe(2)
  })

  it('auto and autoNoTable share the same file-load cache policy (mode itself forces nothing)', () => {
    const auto = docOf([cachedFormula('A1', 0, 0, '1+1', 99)], { calc: calc('auto') })
    evaluateWorkbookFormulas(auto)
    expect(auto.sheets[0]!.rows[0]!.cells[0]!.value).toBe(99)

    const autoNoTable = docOf([cachedFormula('A1', 0, 0, '1+1', 99)], { calc: calc('autoNoTable') })
    evaluateWorkbookFormulas(autoNoTable)
    expect(autoNoTable.sheets[0]!.rows[0]!.cells[0]!.value).toBe(99)
  })

  it('explicit-recalc recomputes an ordinary cache under either auto or autoNoTable', () => {
    const auto = docOf([cachedFormula('A1', 0, 0, '1+1', 99)], { calc: calc('auto') })
    evaluateWorkbookFormulas(auto, { intent: 'explicit-recalc' })
    expect(auto.sheets[0]!.rows[0]!.cells[0]!.value).toBe(2)

    const autoNoTable = docOf([cachedFormula('A1', 0, 0, '1+1', 99)], { calc: calc('autoNoTable') })
    evaluateWorkbookFormulas(autoNoTable, { intent: 'explicit-recalc' })
    expect(autoNoTable.sheets[0]!.rows[0]!.cells[0]!.value).toBe(2)
  })

  it('autoNoTable recalculates structured-table-referencing formulas on an explicit request (calcChain irrelevant)', () => {
    const table: TableMetadata = {
      id: 'table', name: 'Table1', displayName: 'Table1', sheetId: 't', partPath: 'table.xml',
      extent: { sheetId: 't', firstCol: 0, firstRow: 0, cols: 2, rows: 3 },
      headerRowCount: 1, totalsRowCount: 0,
      columns: [{ id: 'a', name: 'Amount', index: 0 }, { id: 'b', name: 'Other', index: 1 }],
    }
    const tableCells: XlsxCell[] = [
      { ref: 'A1', col: 0, row: 0, value: 'Amount', styleIndex: 0 },
      { ref: 'B1', col: 1, row: 0, value: 'Other', styleIndex: 0 },
      { ref: 'A2', col: 0, row: 1, value: 10, styleIndex: 0 },
      { ref: 'A3', col: 0, row: 2, value: 20, styleIndex: 0 },
    ]
    const target = cachedFormula('C1', 2, 0, 'SUM(Table1[Amount])', 1)
    const doc: XlsxDocument = {
      sheets: [sheetOf('T', tableCells, 0) as never, sheetOf('S', [target], 1) as never],
      definedNames: [],
      tables: [table],
      calc: calc('autoNoTable'),
    }
    evaluateWorkbookFormulas(doc, { intent: 'explicit-recalc' })
    expect(target.value).toBe(30)
  })

  it('autoNoTable retains a what-if data-table cache and diagnoses it', () => {
    const whatIf: XlsxCell = { ...cachedFormula('A1', 0, 0, '1+1', 77), formulaType: 'dataTable', ca: true }
    const doc = docOf([whatIf], { calc: calc('autoNoTable') })
    evaluateWorkbookFormulas(doc)
    expect(doc.sheets[0]!.rows[0]!.cells[0]!.value).toBe(77)
    expect(doc.diagnostics?.some(d => d.feature === 'what-if-data-table-unverified')).toBe(true)
  })

  it('force/explicit-recalc cannot recompute what-if data tables', () => {
    const whatIf: XlsxCell = { ...cachedFormula('A1', 0, 0, '1+1', 77), formulaType: 'dataTable' }
    const forced = docOf([whatIf])
    evaluateWorkbookFormulas(forced, { forceRecalc: true })
    expect(forced.sheets[0]!.rows[0]!.cells[0]!.value).toBe(77)
    expect(forced.diagnostics?.some(d => d.feature === 'what-if-data-table-unverified')).toBe(true)

    const explicit: XlsxCell = { ...cachedFormula('A1', 0, 0, '1+1', 77), formulaType: 'dataTable' }
    const doc = docOf([explicit])
    evaluateWorkbookFormulas(doc, { intent: 'explicit-recalc' })
    expect(doc.sheets[0]!.rows[0]!.cells[0]!.value).toBe(77)
  })

  it('an uncached what-if data table publishes no fabricated value', () => {
    const whatIf: XlsxCell = { ref: 'A1', col: 0, row: 0, value: null, styleIndex: 0, hasCachedValue: false, formula: '1+1', formulaType: 'dataTable' }
    const doc = docOf([whatIf], { calc: calc('autoNoTable') })
    evaluateWorkbookFormulas(doc)
    const cell = doc.sheets[0]!.rows[0]!.cells[0]!
    expect(cell.value).toBe(null)
    expect(cell.hasCachedValue).toBe(false)
    expect(doc.diagnostics?.some(d => d.feature === 'what-if-data-table-unverified')).toBe(true)
  })

  it('phase 2: an uncached cyclic cell has no trusted seed and is retained, never assumed zero', () => {
    const doc = docOf([formulaCell('A1', 0, 0, 'A1+1')], {
      calc: calc('auto', { iterate: true, iterateCount: 100, iterateDelta: 0.001 }),
    })
    evaluateWorkbookFormulas(doc)
    const cell = doc.sheets[0]!.rows[0]!.cells[0]!
    // Missing cache is UNVERIFIED: the placeholder is not invented as 0 and the
    // (absent) cache/flags are preserved.
    expect(cell.value).toBe(null)
    expect(cell.hasCachedValue).toBe(false)
    expect(doc.diagnostics?.some(d => d.feature === 'calc-iterative-seed-unverified')).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// manual calcMode gates the automatic flags (root manual-boundaries probe)
// ---------------------------------------------------------------------------
describe('C4 manual calcMode gates automatic recalculation before ca/fullCalcOnLoad/uncached', () => {
  function manualNow(cache: boolean, ca = false, full = false): XlsxDocument {
    const cell: XlsxCell = { ref: 'A1', col: 0, row: 0, value: cache ? 77 : null, styleIndex: 0, hasCachedValue: cache, formula: 'NOW()', ca }
    return docOf([cell], { calc: calc('manual', { fullCalcOnLoad: full }) })
  }

  it.each([[true, true, false], [true, false, true], [false, false, false]])(
    'manual file-load avoids volatile execution (cache=%s ca=%s fullCalcOnLoad=%s)',
    (cache, ca, full) => {
      const doc = manualNow(cache as boolean, ca as boolean, full as boolean)
      let calls = 0
      evaluateWorkbookFormulas(doc, { now: () => { calls++; return Date.UTC(2024, 0, 1) } })
      const cell = doc.sheets[0]!.rows[0]!.cells[0]!
      expect(calls).toBe(0)
      expect(cell.value).toBe(cache ? 77 : null)
      expect(cell.hasCachedValue).toBe(cache)
    },
  )

  it('manual leaves an uncached formula uncomputed with an attributable deferred-calc diagnostic', () => {
    const doc = manualNow(false)
    evaluateWorkbookFormulas(doc, { now: () => 0 })
    expect(doc.diagnostics?.some(d => d.feature === 'calc-manual-deferred')).toBe(true)
  })

  it('manual file-load preserves prior owned dynamic followers and their cached reader', () => {
    const doc: XlsxDocument = {
      sheets: [sheetOf('S', [
        { ref: 'A1', col: 0, row: 0, value: null, styleIndex: 0, hasCachedValue: false, formula: 'SEQUENCE(2,2)' },
        { ref: 'D1', col: 3, row: 0, value: null, styleIndex: 0, hasCachedValue: false, formula: 'B2' },
      ]) as never],
      definedNames: [], tables: [],
    }
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    const cell = (ref: string) => doc.sheets[0]!.rows.flatMap(r => r.cells).find(x => x.ref === ref)
    expect(cell('B2')?.value).toBe(4)
    doc.calc = calc('manual')
    evaluateWorkbookFormulas(doc, { arrayMode: 'dynamic' })
    expect(cell('B2')?.value).toBe(4)
    expect(cell('D1')?.value).toBe(4)
  })

  it('a follow-up force after a manual no-op still cleans up the preserved registry (shrink)', () => {
    const doc: XlsxDocument = {
      sheets: [sheetOf('S', [
        { ref: 'A1', col: 0, row: 0, value: null, styleIndex: 0, hasCachedValue: false, formula: 'SEQUENCE(2,2)' },
      ]) as never],
      definedNames: [], tables: [],
    }
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    const cell = (ref: string) => doc.sheets[0]!.rows.flatMap(r => r.cells).find(x => x.ref === ref)
    expect(cell('B2')?.value).toBe(4)

    doc.calc = calc('manual')
    evaluateWorkbookFormulas(doc, { arrayMode: 'dynamic' }) // no-op preserves B2
    expect(cell('B2')?.value).toBe(4)

    // Shrink the spill on an explicit force: the preserved registry must still
    // allow the own-only cleanup to remove the stale follower.
    doc.calc = calc('auto')
    cell('A1')!.formula = 'SEQUENCE(1,1)'
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cell('A1')?.value).toBe(1)
    expect(cell('B2')).toBeUndefined()
  })
})

// ---------------------------------------------------------------------------
// C4 r2 — a formula-less what-if data table is the ONLY formula-ish cell
//
// Empty `<f t="dataTable" ref=… dt2D=… dtr=… r1=… r2=…/>` (no formula text)
// is the normal OOXML form of a what-if data table. The cheap `hasFormulas`
// early scan used to look only at `cell.formula` / `cell.sharedFormula`, so a
// workbook whose sole formula construct is such a table took the early exit
// and the what-if diagnosis branch was unreachable. These tests pin the
// repaired behavior at both the model and real-bytes levels.
// ---------------------------------------------------------------------------

/** Minimal real .xlsx bytes around one `<sheetData>` block (no calcChain part). */
async function buildRawXlsx(sheetDataXml: string): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`)
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`)
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>
</workbook>`)
  zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetDataXml}</sheetData></worksheet>`)
  zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`)
  return zip.generateAsync({ type: 'uint8array' })
}

describe('C4 r2 — a formula-less what-if data table is diagnosed even as the only formula-ish cell', () => {
  it('model: an empty dataTable cell (no formula text) keeps its cache and is diagnosed', () => {
    const whatIf: XlsxCell = { ref: 'A1', col: 0, row: 0, value: 77, styleIndex: 0, hasCachedValue: true, formulaType: 'dataTable' }
    const doc = docOf([whatIf])
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    const cell = doc.sheets[0]!.rows[0]!.cells[0]!
    expect(cell.value).toBe(77)
    expect(cell.hasCachedValue).toBe(true)
    expect(doc.diagnostics?.some(d => d.feature === 'what-if-data-table-unverified')).toBe(true)
  })

  it('parse: real xlsx bytes with an empty <f t="dataTable"/> keep the cache and diagnose', async () => {
    const bytes = await buildRawXlsx('<row r="1"><c r="A1"><f t="dataTable" ref="A1:B2"/><v>77</v></c></row>')
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const cell = doc.sheets[0]!.rows.flatMap(r => r.cells)[0]!
    expect(cell.formulaType).toBe('dataTable')
    expect(cell.formula).toBeUndefined()
    expect(cell.value).toBe(77)
    expect(doc.diagnostics?.some(d => d.feature === 'what-if-data-table-unverified')).toBe(true)
  })

  it('control: an ordinary formula-less, type-less workbook still takes the cheap early exit', () => {
    const doc = docOf([{ ref: 'A1', col: 0, row: 0, value: 5, styleIndex: 0 }])
    evaluateWorkbookFormulas(doc)
    // The early exit still fires for ordinary all-non-formula content: no
    // semantics write-back and no diagnostics.
    expect(doc.semantics).toBeUndefined()
    expect(doc.diagnostics ?? []).toEqual([])
  })
})

// ---------------------------------------------------------------------------
// C4 phase 2 shipping RED — cache7 cap3 and cross-row in-place order
//
// Expectations are the exact native-established values (ITERATION-FACTS.json,
// root-native-iteration). These are SHIPPING RED first: phase 1 resolves cycles
// to 0, so both fail until the bounded iterative executor lands.
// ---------------------------------------------------------------------------
describe('C4 phase 2 shipping RED — cache7 cap and cross-row order', () => {
  function cellOf(doc: XlsxDocument, ref: string): XlsxCell {
    return doc.sheets[0]!.rows.flatMap(r => r.cells).find(c => c.ref === ref)!
  }

  it('cache7 cap3: A1=A1+1 seed7 / B1=A1*2 → 10/20, next force recalculates from the prior result → 13/26', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'A1+1', 7),
      cachedFormula('B1', 1, 0, 'A1*2', 14),
    ], { calc: calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 }) })

    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(cellOf(doc, 'A1').value).toBe(10)
    expect(cellOf(doc, 'B1').value).toBe(20)

    // Next calculation seeds from the previous result, not the original file cache.
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(cellOf(doc, 'A1').value).toBe(13)
    expect(cellOf(doc, 'B1').value).toBe(26)
  })

  it('cross-row: B1=A2+1 / C1=A2+B1 / A2=B1+1 cache0 cap3 delta0 → 5/9/6 (row/column in-place, never 11)', () => {
    const doc = docOf([
      cachedFormula('B1', 1, 0, 'A2+1', 0),
      cachedFormula('C1', 2, 0, 'A2+B1', 0),
      cachedFormula('A2', 0, 1, 'B1+1', 0),
    ], { calc: calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 }) })

    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(cellOf(doc, 'B1').value).toBe(5)
    expect(cellOf(doc, 'C1').value).toBe(9)
    expect(cellOf(doc, 'A2').value).toBe(6)
    // A final post-pass dependent recompute would publish 11 — explicitly wrong.
    expect(cellOf(doc, 'C1').value).not.toBe(11)
  })

  it('iterate:false control: the original cycle result (0) is retained with no iteration gate', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'A1+1', 7),
    ], { calc: calc('auto', { iterate: false, iterateCount: 3, iterateDelta: 0 }) })

    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(cellOf(doc, 'A1').value).toBe(0)
    expect(doc.diagnostics?.some(d => d.feature.startsWith('calc-iterative')) ?? false).toBe(false)
  })

  it('reached volatile cyclic profile is gated and the valid cache retained', () => {
    const doc = docOf([cachedFormula('A1', 0, 0, 'A1+NOW()', 7)], {
      calc: calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 }),
    })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, now: () => Date.UTC(2024, 0, 1), timeZone: 'UTC' })
    expect(cellOf(doc, 'A1').value).toBe(7)
    expect(doc.diagnostics?.some(d => d.feature.includes('calc-iterative') && d.message.toLowerCase().includes('volatile'))).toBe(true)
  })

  it('dead volatile branch does not gate an established numerical cycle', () => {
    const doc = docOf([cachedFormula('A1', 0, 0, 'IF(TRUE,A1+1,NOW())', 7)], {
      calc: calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 }),
    })
    let clocks = 0
    evaluateWorkbookFormulas(doc, { forceRecalc: true, now: () => { clocks++; return Date.UTC(2024, 0, 1) } })
    expect(cellOf(doc, 'A1').value).toBe(10)
    expect(clocks).toBe(0)
    expect(doc.diagnostics?.some(d => d.message.toLowerCase().includes('volatile')) ?? false).toBe(false)
  })

  it('changed actual dependency edge gates the adaptive cyclic profile', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'IF(A1<1,A1+1,B1+1)', 0),
      cachedFormula('B1', 1, 0, 'A1+1', 0),
    ], { calc: calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 }) })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(cellOf(doc, 'A1').value).toBe(0)
    expect(cellOf(doc, 'B1').value).toBe(0)
    expect(doc.diagnostics?.some(d => d.feature.includes('calc-iterative') && d.message.toLowerCase().includes('dependenc'))).toBe(true)
  })

  it('discovery never publishes a placeholder over a cyclic cache before an unrelated clock callback', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'A1+1', 7),
      cachedFormula('B1', 1, 0, 'NOW()', 0),
    ], { calc: calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 }) })
    let observed: unknown
    evaluateWorkbookFormulas(doc, { forceRecalc: true, now: () => { observed = cellOf(doc, 'A1').value; return Date.UTC(2024, 0, 1) } })
    expect(observed).toBe(7)
    expect(cellOf(doc, 'A1').value).toBe(10)
  })

  it('an actually shipped calcChain gates the unverified pass-order provenance', () => {
    const doc = docOf([cachedFormula('A1', 0, 0, 'A1+1', 7)], {
      calc: calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 }),
      calcChainPresent: true,
    })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(cellOf(doc, 'A1').value).toBe(7)
    expect(doc.diagnostics?.some(d => d.feature.includes('calc-iterative') && d.message.toLowerCase().includes('calcchain'))).toBe(true)
  })

  it('a non-integer iteration count is unavailable rather than silently rounded', () => {
    const doc = docOf([cachedFormula('A1', 0, 0, 'A1+1', 7)], {
      calc: calc('auto', { iterate: true, iterateCount: 1.5, iterateDelta: 0 }),
    })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(cellOf(doc, 'A1').value).toBe(7)
    expect(doc.diagnostics?.some(d => d.feature === 'calc-iterative-limit-unverified')).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// C4 r2 — ROOT COUNTEREXAMPLES (SHIPPING RED FIRST)
//
// C4/r1/root-phase2-preview/iteration-gates.test.ts last two cases. Both fail
// against the frozen r1 candidate and are added BEFORE the narrow repair:
//  1. an untrusted cycle back-edge invented a provisional 0, which decided
//     `IF(A1=0, …)` and sampled the clock once; the missing seed must instead
//     keep the cell null / hasCachedValue=false and emit
//     `calc-iterative-seed-unverified` with ZERO clock samples.
//  2. the adaptive guard compared read sets already filtered to formula keys,
//     so a plain INPUT dependency edge (B1) appearing on the third pass was
//     invisible; the actual read sets (inputs included) must gate with
//     `calc-iterative-dependency-unverified` and retain the shipped caches.
// ---------------------------------------------------------------------------
describe('C4 r2 — untrusted cycle seed and plain-input adaptive edge (root counterexamples)', () => {
  function cellOf(doc: XlsxDocument, ref: string): XlsxCell {
    return doc.sheets[0]!.rows.flatMap(r => r.cells).find(c => c.ref === ref)!
  }

  it('missing seed never invents zero to choose and execute a volatile branch', () => {
    const doc = docOf([formulaCell('A1', 0, 0, 'IF(A1=0,NOW(),A1+1)')], {
      calc: calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 }),
    })
    const a = doc.sheets[0]!.rows[0]!.cells[0]!
    a.value = null
    a.hasCachedValue = false
    let clocks = 0
    evaluateWorkbookFormulas(doc, { forceRecalc: true, now: () => { clocks++; return Date.UTC(2024, 0, 1) } })
    expect(clocks).toBe(0)
    expect(a.value).toBe(null)
    expect(a.hasCachedValue).toBe(false)
    expect(doc.diagnostics?.some(d => d.feature === 'calc-iterative-seed-unverified')).toBe(true)
  })

  it('changing a plain-input dependency is an actual adaptive edge change', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'IF(A1<2,A1+1,B1+1)', 0),
      { ref: 'B1', col: 1, row: 0, value: 5, styleIndex: 0, hasCachedValue: true },
    ], { calc: calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 }) })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(cellOf(doc, 'A1').value).toBe(0)
    expect(cellOf(doc, 'B1').value).toBe(5)
    expect(doc.diagnostics?.some(d => d.feature === 'calc-iterative-dependency-unverified')).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// C4 r2 followup — the unavailable-seed control is NOT a catchable native error
//
// Returning a typed #NAME? gate from the untrusted back-edge was still visible
// to IFERROR, which then executed its fallback (NOW()) and sampled the clock.
// The missing-seed signal must be a PRIVATE workbook driver control that no
// Core formula function can observe. Paired positives pin the boundary: a
// trusted seed keeps an IFERROR-wrapped cycle, a genuine acyclic #N/A still
// runs its fallback, and an opaque caller-thrown error keeps its identity.
// ---------------------------------------------------------------------------
describe('C4 r2 followup — the unavailable-seed control is not a catchable native error', () => {
  function cellOf(doc: XlsxDocument, ref: string): XlsxCell {
    return doc.sheets[0]!.rows.flatMap(r => r.cells).find(c => c.ref === ref)!
  }
  const ITER = () => calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 })

  it('IFERROR cannot turn an unverified missing cycle seed into a volatile fallback execution', () => {
    const doc = docOf([formulaCell('A1', 0, 0, 'IFERROR(A1,NOW())')], { calc: ITER() })
    const a = doc.sheets[0]!.rows[0]!.cells[0]!
    a.value = null
    a.hasCachedValue = false
    let clocks = 0
    evaluateWorkbookFormulas(doc, { forceRecalc: true, now: () => { clocks++; return Date.UTC(2024, 0, 1) } })
    expect(clocks).toBe(0)
    expect(a.value).toBe(null)
    expect(a.hasCachedValue).toBe(false)
    expect(doc.diagnostics?.some(d => d.feature === 'calc-iterative-seed-unverified')).toBe(true)
  })

  it('a trusted seed keeps an IFERROR-wrapped cycle and never samples the clock', () => {
    const doc = docOf([cachedFormula('A1', 0, 0, 'IFERROR(A1,NOW())', 7)], { calc: ITER() })
    let clocks = 0
    evaluateWorkbookFormulas(doc, { forceRecalc: true, now: () => { clocks++; return Date.UTC(2024, 0, 1) } })
    expect(cellOf(doc, 'A1').value).toBe(7)
    expect(clocks).toBe(0)
  })

  it('a genuine acyclic #N/A still runs its IFERROR fallback', () => {
    const doc = docOf([formulaCell('A1', 0, 0, 'IFERROR(NA(),5)')], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(cellOf(doc, 'A1').value).toBe(5)
  })

  it('an opaque caller-thrown error propagates with identity', () => {
    const doc = docOf([formulaCell('A1', 0, 0, 'NOW()')], { calc: ITER() })
    const sentinel = new Error('caller-boom')
    let thrown: unknown
    try {
      evaluateWorkbookFormulas(doc, { forceRecalc: true, now: () => { throw sentinel } })
    } catch (error) {
      thrown = error
    }
    expect(thrown).toBe(sentinel)
  })
})

// ---------------------------------------------------------------------------
// C4 r2 — ALL NINE root-native numeric iteration profiles as SHIPPING tests
//
// Expectations are the exact measured native values in
// C4/r1/root-native-iteration/ITERATION-FACTS.json (fixture SHA retained).
// Only two of the nine were previously in the shipping inventory; the rest
// lived only in a root tmp probe, so the repo CI could pass while a native
// profile regressed. These are model reconstructions of the exact fixture
// formulas / caches / settings, asserted at the established public 15-digit
// tolerance (toBeCloseTo(…, 11), same as the root replay).
// ---------------------------------------------------------------------------
describe('C4 r2 — all nine root-native numeric iteration profiles', () => {
  function cellOf(doc: XlsxDocument, ref: string): XlsxCell {
    return doc.sheets[0]!.rows.flatMap(r => r.cells).find(c => c.ref === ref)!
  }
  function num(doc: XlsxDocument, ref: string, expected: number): void {
    expect(cellOf(doc, ref).value, `native ${ref}`).toBeCloseTo(expected, 11)
  }

  it('cap3-cache7-native-r2: A1=A1+1 seed7 / B1=A1*2 seed14, cap3 delta0 → 10/20', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'A1+1', 7),
      cachedFormula('B1', 1, 0, 'A1*2', 14),
    ], { calc: calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 }) })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    num(doc, 'A1', 10)
    num(doc, 'B1', 20)
  })

  it('cap3-cache7-recalculate: next force seeds the prior result → 13/26', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'A1+1', 7),
      cachedFormula('B1', 1, 0, 'A1*2', 14),
    ], { calc: calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 }) })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    num(doc, 'A1', 13)
    num(doc, 'B1', 26)
  })

  function profilesCap3(): XlsxDocument {
    return docOf([
      cachedFormula('A1', 0, 0, 'A1+1', 7),
      cachedFormula('B1', 1, 0, 'A1*2', 14),
      cachedFormula('C1', 2, 0, 'C1+1', 0),
      cachedFormula('D1', 3, 0, 'C1*2', 0),
      formulaCell('E1', 4, 0, 'E1+1'),
      formulaCell('F1', 5, 0, 'E1*2'),
      cachedFormula('G1', 6, 0, '(G1+1)/2', 0),
      cachedFormula('H1', 7, 0, 'G1*2', 0),
      cachedFormula('I1', 8, 0, 'J1+1', 0),
      cachedFormula('J1', 9, 0, 'I1+1', 0),
      cachedFormula('K1', 10, 0, 'I1+J1', 0),
      cachedFormula('L1', 11, 0, '(M1+1)/2', 0),
      cachedFormula('M1', 12, 0, '(L1+1)/2', 0),
      cachedFormula('N1', 13, 0, 'L1+M1', 0),
    ], { calc: calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 }) })
  }

  it('profiles-cap3-native: seed7 / seed0 / missing / halvers / two-cell cycle exact', () => {
    const doc = profilesCap3()
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    num(doc, 'A1', 10)
    num(doc, 'B1', 20)
    num(doc, 'C1', 3)
    num(doc, 'D1', 6)
    // E1/F1 ship NO cache: unavailable, retained, never invented as 0.
    expect(cellOf(doc, 'E1').value).toBe(null)
    expect(cellOf(doc, 'E1').hasCachedValue).toBe(false)
    expect(cellOf(doc, 'F1').value).toBe(null)
    expect(cellOf(doc, 'F1').hasCachedValue).toBe(false)
    expect(doc.diagnostics?.some(d => d.feature === 'calc-iterative-seed-unverified')).toBe(true)
    num(doc, 'G1', 0.875)
    num(doc, 'H1', 1.75)
    num(doc, 'I1', 5)
    num(doc, 'J1', 6)
    num(doc, 'K1', 11)
    num(doc, 'L1', 0.96875)
    num(doc, 'M1', 0.984375)
    num(doc, 'N1', 1.953125)
  })

  it('profiles-cap3-recalculate: next force seeds prior results exact', () => {
    const doc = profilesCap3()
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    num(doc, 'A1', 13)
    num(doc, 'B1', 26)
    num(doc, 'C1', 6)
    num(doc, 'D1', 12)
    expect(cellOf(doc, 'E1').value).toBe(null)
    expect(cellOf(doc, 'F1').value).toBe(null)
    num(doc, 'G1', 0.984375)
    num(doc, 'H1', 1.96875)
    num(doc, 'I1', 11)
    num(doc, 'J1', 12)
    num(doc, 'K1', 23)
    num(doc, 'L1', 0.99951171875)
    num(doc, 'M1', 0.999755859375)
    num(doc, 'N1', 1.999267578125)
  })

  it('converge-single-native: A1=(A1+1)/2 seed0, cap100 delta.001 → .9990234375 (10 passes)', () => {
    const doc = docOf([cachedFormula('A1', 0, 0, '(A1+1)/2', 0)], {
      calc: calc('auto', { iterate: true, iterateCount: 100, iterateDelta: 0.001 }),
    })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    num(doc, 'A1', 0.9990234375)
  })

  it('converge100-native: halver + dependent 2x, cap100 delta.001 → .99951171875/1.9990234375 (11 passes)', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, '(A1+1)/2', 0),
      cachedFormula('B1', 1, 0, 'A1*2', 0),
    ], { calc: calc('auto', { iterate: true, iterateCount: 100, iterateDelta: 0.001 }) })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    num(doc, 'A1', 0.99951171875)
    num(doc, 'B1', 1.9990234375)
  })

  it('converge-cohort-native: two halvers, dependents 2x/1000x, cap100 delta.001 → 20 passes', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, '(A1+1)/2', 0),
      cachedFormula('B1', 1, 0, 'A1*2', 0),
      cachedFormula('C1', 2, 0, '(C1+1)/2', 0),
      cachedFormula('D1', 3, 0, 'C1*1000', 0),
    ], { calc: calc('auto', { iterate: true, iterateCount: 100, iterateDelta: 0.001 }) })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    num(doc, 'A1', 0.99999904632568359)
    num(doc, 'B1', 1.9999980926513672)
    num(doc, 'C1', 0.99999904632568359)
    num(doc, 'D1', 999.99904632568359)
  })

  it('row-column-order-native: B1=A2+1 / C1=A2+B1 / A2=B1+1 cache0 cap3 → 5/9/6', () => {
    const doc = docOf([
      cachedFormula('B1', 1, 0, 'A2+1', 0),
      cachedFormula('C1', 2, 0, 'A2+B1', 0),
      cachedFormula('A2', 0, 1, 'B1+1', 0),
    ], { calc: calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 }) })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    num(doc, 'B1', 5)
    num(doc, 'C1', 9)
    num(doc, 'A2', 6)
  })

  it('threshold-half-native: strict `<` delta .5 stops after two passes → .75', () => {
    const doc = docOf([cachedFormula('A1', 0, 0, '(A1+1)/2', 0)], {
      calc: calc('auto', { iterate: true, iterateCount: 100, iterateDelta: 0.5 }),
    })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    num(doc, 'A1', 0.75)
  })
})

// ---------------------------------------------------------------------------
// C4 r3 — a cyclic/provisional dynamic matrix is gated BEFORE commit
//
// Root QUALITY (C4/r2/root-review/cyclic-spill-publication.test.ts, extended
// cyclic-spill-reader-r1.log: 2 FAIL | 1 PASS) proves the discovery shadow
// PUBLISHED a cyclic owner's generated followers before the cohort was
// verified: A1=A1+SEQUENCE(1,2) restored the anchor cache 7 but left a
// generated B1=9, and a cached reader D1=B1 lost 99 for that 9. An unverified
// cyclic/provisional matrix must be gated before commit so no follower is ever
// written; an acyclic dynamic array in an iterate-enabled workbook still
// publishes normally. These are SHIPPING RED first (added before the repair).
// ---------------------------------------------------------------------------
describe('C4 r3 — unverified cyclic dynamic matrix never commits followers', () => {
  function cellOf(doc: XlsxDocument, ref: string): XlsxCell | undefined {
    return doc.sheets[0]!.rows.flatMap(r => r.cells).find(c => c.ref === ref)
  }
  const ITER = () => calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 })

  it('A1=A1+SEQUENCE(1,2) cache7 dynamic: anchor cache 7 retained and no generated B1', () => {
    const doc = docOf([cachedFormula('A1', 0, 0, 'A1+SEQUENCE(1,2)', 7)], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'A1')!.value).toBe(7)
    expect(cellOf(doc, 'B1')).toBeUndefined()
    expect(doc.diagnostics?.some(d => d.feature === 'calc-iterative-spill-unverified')).toBe(true)
  })

  it('a cached reader of the gated cyclic owner keeps its valid value (D1=B1, 99)', () => {
    const doc = docOf([
      cachedFormula('D1', 3, 0, 'B1', 99),
      cachedFormula('A1', 0, 0, 'A1+SEQUENCE(1,2)', 7),
    ], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'D1')!.value).toBe(99)
    expect(cellOf(doc, 'B1')).toBeUndefined()
    expect(doc.diagnostics?.some(d => d.feature.includes('iterative') || d.feature.includes('spill'))).toBe(true)
  })

  it('parent/descendant: numeric cycle A1 + B1=SEQUENCE(1,2,A1) retains both caches, no C1', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'A1+1', 5),
      cachedFormula('B1', 1, 0, 'SEQUENCE(1,2,A1)', 10),
    ], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'A1')!.value).toBe(5)
    expect(cellOf(doc, 'B1')!.value).toBe(10)
    expect(cellOf(doc, 'C1')).toBeUndefined()
    expect(cellOf(doc, 'A1')!.hasCachedValue).toBe(true)
    expect(cellOf(doc, 'B1')!.hasCachedValue).toBe(true)
    expect(doc.diagnostics?.some(d => d.feature === 'calc-iterative-spill-unverified')).toBe(true)
  })

  it('an acyclic dynamic array in an iterate-enabled workbook still publishes its followers', () => {
    const doc = docOf([cachedFormula('A1', 0, 0, 'SEQUENCE(1,2)', 7)], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'A1')!.value).toBe(1)
    expect(cellOf(doc, 'B1')!.value).toBe(2)
    expect(doc.diagnostics?.some(d => d.feature.startsWith('calc-iterative')) ?? false).toBe(false)
  })

  it('gating never deletes or rewrites a pre-existing styled/blank cell in the potential footprint', () => {
    const styled: XlsxCell = { ref: 'B1', col: 1, row: 0, value: null, styleIndex: 3, hasCachedValue: false }
    const input: XlsxCell = { ref: 'C5', col: 2, row: 4, value: 'keep', styleIndex: 0, hasCachedValue: true }
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'A1+SEQUENCE(1,2)', 7),
      styled,
      input,
    ], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'A1')!.value).toBe(7)
    const b1 = cellOf(doc, 'B1')!
    expect(b1.value).toBe(null)
    expect(b1.styleIndex).toBe(3)
    expect(b1.hasCachedValue).toBe(false)
    expect(cellOf(doc, 'C5')!.value).toBe('keep')
  })
})

describe('C4 r4 — sparse REGION read over an unavailable cyclic spill keeps its valid cache', () => {
  function cellOf(doc: XlsxDocument, ref: string): XlsxCell | undefined {
    return doc.sheets[0]!.rows.flatMap(r => r.cells).find(c => c.ref === ref)
  }
  const ITER = () => calc('auto', { iterate: true, iterateCount: 3, iterateDelta: 0 })

  it('RED: F1=SUM(B1:C1) cached 42 over the gated cyclic owner footprint is retained, never a guessed 0', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'A1+SEQUENCE(1,2)', 7),
      cachedFormula('F1', 5, 0, 'SUM(B1:C1)', 42),
    ], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'B1')).toBeUndefined()
    expect(cellOf(doc, 'F1')!.value).toBe(42)
    expect(doc.diagnostics?.some(d => d.feature.includes('iterative') || d.feature.includes('spill'))).toBe(true)
  })

  it('RED: G1=COUNTIF(B1:C1,"") cached 55 over the gated cyclic owner footprint is retained, never a guessed blank count', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'A1+SEQUENCE(1,2)', 7),
      cachedFormula('G1', 6, 0, 'COUNTIF(B1:C1,"")', 55),
    ], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'B1')).toBeUndefined()
    expect(cellOf(doc, 'G1')!.value).toBe(55)
    expect(doc.diagnostics?.some(d => d.feature.includes('iterative') || d.feature.includes('spill'))).toBe(true)
  })

  it('POSITIVE: an ordinary ownerless empty region still computes SUM(B1:C1) = 0', () => {
    const doc = docOf([cachedFormula('F1', 5, 0, 'SUM(B1:C1)', 42)], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'F1')!.value).toBe(0)
  })

  it('COUNTIFS blank criterion over the same gated footprint retains its cache', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'A1+SEQUENCE(1,2)', 7),
      cachedFormula('G1', 6, 0, 'COUNTIFS(B1:C1,"")', 55),
    ], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'B1')).toBeUndefined()
    expect(cellOf(doc, 'G1')!.value).toBe(55)
  })

  it('COUNT over the same gated footprint retains its cache', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'A1+SEQUENCE(1,2)', 7),
      cachedFormula('G1', 6, 0, 'COUNT(B1:C1)', 7),
    ], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'G1')!.value).toBe(7)
  })

  it('POSITIVE: ordinary ownerless blank-criterion structural count is unchanged (COUNTIF = 2)', () => {
    const doc = docOf([cachedFormula('G1', 6, 0, 'COUNTIF(B1:C1,"")', 55)], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'G1')!.value).toBe(2)
  })

  it('POSITIVE: ordinary ownerless COUNT is unchanged (0)', () => {
    const doc = docOf([cachedFormula('G1', 6, 0, 'COUNT(B1:C1)', 7)], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'G1')!.value).toBe(0)
  })

  it('POSITIVE: an acyclic spill in an iterate-enabled workbook still publishes and reads normally', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'SEQUENCE(1,2)', 7),
      cachedFormula('F1', 5, 0, 'SUM(B1:C1)', 42),
    ], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'A1')!.value).toBe(1)
    expect(cellOf(doc, 'B1')!.value).toBe(2)
    expect(cellOf(doc, 'F1')!.value).toBe(2)
    expect(doc.diagnostics?.some(d => d.feature.startsWith('calc-iterative')) ?? false).toBe(false)
  })

  it('direct cached single-cell read of the gated footprint keeps 99 alongside a region reader', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'A1+SEQUENCE(1,2)', 7),
      cachedFormula('D1', 3, 0, 'B1', 99),
      cachedFormula('F1', 5, 0, 'SUM(B1:C1)', 42),
    ], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'B1')).toBeUndefined()
    expect(cellOf(doc, 'D1')!.value).toBe(99)
    expect(cellOf(doc, 'F1')!.value).toBe(42)
  })

  it('an unknown # footprint reader retains its valid cache (never a guessed 0)', () => {
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'A1+SEQUENCE(1,2)', 7),
      cachedFormula('F1', 5, 0, 'SUM(A1#)', 42),
    ], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'F1')!.value).toBe(42)
  })

  it('retained reader snapshots preserve the shipped error-flag value exactly (false / empty / error-looking text)', () => {
    const cases: Array<{ value: XlsxCell['value']; valueIsError?: boolean }> = [
      { value: false },
      { value: '' },
      { value: '#N/A', valueIsError: true },
      { value: '#N/A', valueIsError: false },
    ]
    for (const c of cases) {
      const f1: XlsxCell = { ref: 'F1', col: 5, row: 0, value: c.value, styleIndex: 0, hasCachedValue: true, formula: 'SUM(B1:C1)', ...(c.valueIsError === undefined ? {} : { valueIsError: c.valueIsError }) }
      const doc = docOf([cachedFormula('A1', 0, 0, 'A1+SEQUENCE(1,2)', 7), f1], { calc: ITER() })
      evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
      const out = cellOf(doc, 'F1')!
      expect(out.value).toBe(c.value)
      expect(out.valueIsError).toBe(c.valueIsError)
      expect(out.hasCachedValue).toBe(true)
    }
  })

  it('gating a region read never rewrites or deletes a pre-existing blank/input cell', () => {
    const blank: XlsxCell = { ref: 'B1', col: 1, row: 0, value: null, styleIndex: 3, hasCachedValue: false }
    const input: XlsxCell = { ref: 'C1', col: 2, row: 0, value: 'keep', styleIndex: 0, hasCachedValue: true }
    const doc = docOf([
      cachedFormula('A1', 0, 0, 'A1+SEQUENCE(1,2)', 7),
      cachedFormula('F1', 5, 0, 'SUM(B1:C1)', 42),
      blank,
      input,
    ], { calc: ITER() })
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect(cellOf(doc, 'F1')!.value).toBe(42)
    expect(cellOf(doc, 'B1')).toBe(blank)
    expect(blank.value).toBe(null)
    expect(blank.styleIndex).toBe(3)
    expect(blank.hasCachedValue).toBe(false)
    expect(cellOf(doc, 'C1')).toBe(input)
    expect(input.value).toBe('keep')
  })
})
