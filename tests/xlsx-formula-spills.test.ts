/**
 * C2 phase2 — workbook spill/legacy lifecycle (bounded static RED batch).
 *
 * Official primary facts (root primary-spill-root-r1/FACTS.json, not fresh
 * native export):
 *  - blocked output range, merged cells, array formula inside an Excel Table,
 *    and sheet-edge overrun are documented #SPILL! causes;
 *  - only the top-left anchor is editable, followers are derived;
 *  - dynamic A1# (SEQ10 spilled A2:A11, summed via A2#) tracks the owner;
 *  - the spilled-range-operator page's #REF! note is a CLOSED external-workbook
 *    limitation, NOT proof of a failed/nonspilling-anchor error — that exact
 *    error is explicitly UNVERIFIED and stays gated;
 *  - array-mode ROW(refRange) returns row numbers vertically; scalar ROW(ref)
 *    returns the first row; array-mode COLUMN(refRange) returns column numbers
 *    horizontally; scalar COLUMN(ref) returns the leftmost column.
 *
 * Explicit dynamic provenance is supplied by the caller (arrayMode:'dynamic');
 * no modern CM/VM metadata is guessed, no competing-owner arbitration or
 * generic legacy padding/scalar repetition is invented.
 */
import { describe, it, expect } from 'vitest'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import type { XlsxCell, XlsxDocument, XlsxRow, XlsxSheet } from '../src/xlsx/types'
import type { TableMetadata } from '../src/xlsx/formula/types'

// ---------------------------------------------------------------------------
// Synthetic document helpers
// ---------------------------------------------------------------------------

function mkCell(ref: string, col: number, row: number, extra: Partial<XlsxCell> = {}): XlsxCell {
  return { ref, col, row, value: null, styleIndex: 0, hasCachedValue: false, ...extra }
}

function mkSheet(name: string, cells: XlsxCell[], extra: Partial<XlsxSheet> = {}): XlsxSheet {
  const byRow = new Map<number, XlsxCell[]>()
  for (const c of cells) {
    const list = byRow.get(c.row)
    if (list) list.push(c)
    else byRow.set(c.row, [c])
  }
  const rows: XlsxRow[] = [...byRow.keys()].sort((a, b) => a - b).map((index) => ({ index, cells: byRow.get(index)! }))
  return {
    name,
    sheetId: name.toLowerCase(),
    workbookIndex: 0,
    cols: [],
    merges: [],
    mergeRanges: [],
    rows,
    ...extra,
  }
}

function docOf(sheets: XlsxSheet[], extra: Partial<XlsxDocument> = {}): XlsxDocument {
  return { sheets, definedNames: [], tables: [], ...extra }
}

function cellAt(doc: XlsxDocument, sheetName: string, col: number, row: number): XlsxCell | undefined {
  const sheet = doc.sheets.find((s) => s.name.toLowerCase() === sheetName.toLowerCase())
  if (!sheet) return undefined
  const rowObj = sheet.rows.find((r) => r.index === row)
  return rowObj?.cells.find((c) => c.col === col)
}

function valueAt(doc: XlsxDocument, sheetName: string, col: number, row: number): unknown {
  const cell = cellAt(doc, sheetName, col, row)
  return cell ? cell.value : undefined
}

function hasFeature(doc: XlsxDocument, feature: string): boolean {
  return (doc.diagnostics ?? []).some((d) => d.feature === feature)
}

function featureMessage(doc: XlsxDocument, feature: string): string {
  return (doc.diagnostics ?? []).filter((d) => d.feature === feature).map((d) => d.message).join(' | ')
}

const DYNAMIC = { forceRecalc: true, arrayMode: 'dynamic' as const }

// ---------------------------------------------------------------------------
// 1. Single owner: order-independent discovery, A1#, region reads
// ---------------------------------------------------------------------------

describe('C2 phase2 spill — single dynamic owner and order independence', () => {
  function build(order: 'anchor-first' | 'follower-first'): XlsxDocument {
    const a1 = mkCell('A1', 0, 0, { formula: 'SEQUENCE(2,3)' })
    const d1 = mkCell('D1', 3, 0, { formula: 'C2' })
    const e1 = mkCell('E1', 4, 0, { formula: 'SUM(A1:C2)' })
    const e2 = mkCell('E2', 4, 1, { formula: 'ROWS(A1#)' })
    const e3 = mkCell('E3', 4, 2, { formula: 'COLUMNS(A1#)' })
    const e4 = mkCell('E4', 4, 3, { formula: 'INDEX(A1#,2,3)' })
    const row0 = order === 'follower-first' ? [d1, a1, e1] : [a1, d1, e1]
    return docOf([
      mkSheet('S', [
        ...row0,
        e2, e3, e4,
        mkCell('A5', 0, 4, { value: 7, hasCachedValue: true }),
      ]),
    ])
  }

  it('D1=C2 evaluated before A1=SEQUENCE(2,3) still sees the committed follower', () => {
    const doc = build('follower-first')
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe(1)
    expect(valueAt(doc, 'S', 1, 0)).toBe(2)
    expect(valueAt(doc, 'S', 2, 0)).toBe(3)
    expect(valueAt(doc, 'S', 0, 1)).toBe(4)
    expect(valueAt(doc, 'S', 1, 1)).toBe(5)
    expect(valueAt(doc, 'S', 2, 1)).toBe(6)
    expect(valueAt(doc, 'S', 3, 0)).toBe(6) // D1 = C2
    expect(valueAt(doc, 'S', 4, 0)).toBe(21) // SUM(A1:C2)
    expect(valueAt(doc, 'S', 4, 1)).toBe(2) // ROWS(A1#)
    expect(valueAt(doc, 'S', 4, 2)).toBe(3) // COLUMNS(A1#)
    expect(valueAt(doc, 'S', 4, 3)).toBe(6) // INDEX(A1#,2,3)
  })

  it('reverse input iteration produces identical results', () => {
    const first = build('follower-first')
    const second = build('anchor-first')
    evaluateWorkbookFormulas(first, DYNAMIC)
    evaluateWorkbookFormulas(second, DYNAMIC)
    for (const [col, row] of [[0, 0], [1, 0], [2, 0], [0, 1], [1, 1], [2, 1], [3, 0], [4, 0], [4, 1], [4, 2], [4, 3]] as const) {
      expect(valueAt(second, 'S', col, row)).toBe(valueAt(first, 'S', col, row))
    }
  })

  it('the anchor keeps its original formula text; followers carry primitives', () => {
    const doc = build('follower-first')
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(cellAt(doc, 'S', 0, 0)?.formula).toBe('SEQUENCE(2,3)')
    expect(cellAt(doc, 'S', 1, 0)?.value).toBe(2)
    expect(cellAt(doc, 'S', 2, 1)?.valueIsError).toBe(false)
    expect(cellAt(doc, 'S', 0, 0)?.valueIsError).toBe(false)
  })
})

// ---------------------------------------------------------------------------
// 2. Collisions: verified #SPILL! family + reason, no partial publication
// ---------------------------------------------------------------------------

describe('C2 phase2 spill — collision outcomes (#SPILL!)', () => {
  it('an original input cell in the footprint blocks with reason input-cell', () => {
    const doc = docOf([
      mkSheet('S', [
        mkCell('A1', 0, 0, { formula: 'SEQUENCE(2,3)' }),
        mkCell('B1', 1, 0, { value: 99, hasCachedValue: true }),
      ]),
    ])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe('#SPILL!')
    expect(cellAt(doc, 'S', 0, 0)?.valueIsError).toBe(true)
    expect(valueAt(doc, 'S', 1, 0)).toBe(99) // original input stable
    expect(cellAt(doc, 'S', 2, 0)).toBeUndefined() // no partial followers
    expect(cellAt(doc, 'S', 0, 1)).toBeUndefined()
    expect(cellAt(doc, 'S', 1, 1)).toBeUndefined()
    expect(hasFeature(doc, 'spill-collision')).toBe(true)
    expect(featureMessage(doc, 'spill-collision')).toContain('input-cell')
  })

  it('another formula cell in the footprint blocks with reason formula-cell', () => {
    const doc = docOf([
      mkSheet('S', [
        mkCell('A1', 0, 0, { formula: 'SEQUENCE(2,3)' }),
        mkCell('B1', 1, 0, { formula: '5' }),
      ]),
    ])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe('#SPILL!')
    expect(valueAt(doc, 'S', 1, 0)).toBe(5) // other formula still computed
    expect(cellAt(doc, 'S', 2, 0)).toBeUndefined()
    expect(featureMessage(doc, 'spill-collision')).toContain('formula-cell')
  })

  it('a merged cell in the footprint blocks with reason merged-cell', () => {
    const doc = docOf([
      mkSheet('S', [mkCell('A1', 0, 0, { formula: 'SEQUENCE(2,3)' })], {
        merges: ['B1:B1'],
        mergeRanges: [{ minRow: 0, minCol: 1, maxRow: 0, maxCol: 1 }],
      }),
    ])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe('#SPILL!')
    expect(cellAt(doc, 'S', 1, 0)).toBeUndefined()
    expect(featureMessage(doc, 'spill-collision')).toContain('merged-cell')
  })

  it('a structured-table region blocks with reason table', () => {
    const table: TableMetadata = {
      id: 't1', name: 'T', displayName: 'T', sheetId: 's', partPath: 'xl/tables/table1.xml',
      extent: { sheetId: 's', firstCol: 1, firstRow: 0, cols: 2, rows: 2 },
      headerRowCount: 1, totalsRowCount: 0,
      columns: [{ id: 'c1', name: 'X', index: 0 }, { id: 'c2', name: 'Y', index: 1 }],
    }
    const doc = docOf([
      mkSheet('S', [
        mkCell('A1', 0, 0, { formula: 'SEQUENCE(2,3)' }),
        mkCell('B1', 1, 0, { value: 'X', hasCachedValue: true }),
        mkCell('C1', 2, 0, { value: 'Y', hasCachedValue: true }),
        mkCell('B2', 1, 1, { value: 1, hasCachedValue: true }),
        mkCell('C2', 2, 1, { value: 2, hasCachedValue: true }),
      ]),
    ], { tables: [table] })
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe('#SPILL!')
    expect(featureMessage(doc, 'spill-collision')).toContain('table')
    expect(cellAt(doc, 'S', 2, 0)?.value).toBe('Y') // table inputs stable
  })

  it('a sheet-edge overrun blocks with reason sheet-edge', () => {
    const lastCol = 16383
    const doc = docOf([
      mkSheet('S', [mkCell('XFD1', lastCol, 0, { formula: 'SEQUENCE(1,2)' })]),
    ])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', lastCol, 0)).toBe('#SPILL!')
    expect(featureMessage(doc, 'spill-collision')).toContain('sheet-edge')
  })

  it('does not clear or rewrite the colliding original input style/formula', () => {
    const input = mkCell('B1', 1, 0, { value: 99, hasCachedValue: true, styleIndex: 4 })
    const doc = docOf([mkSheet('S', [mkCell('A1', 0, 0, { formula: 'SEQUENCE(2,3)' }), input])])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(input.value).toBe(99)
    expect(input.styleIndex).toBe(4)
    expect(input.formula).toBeUndefined()
  })
})

// ---------------------------------------------------------------------------
// 3. Own-only shrink / replacement / failure
// ---------------------------------------------------------------------------

describe('C2 phase2 spill — own-only shrink and replacement', () => {
  it('shrinking SEQUENCE(2,1) -> SEQUENCE(1,1) removes only the owned tail and recalcs the dependent to absent 0', () => {
    const a1 = mkCell('A1', 0, 0, { formula: 'SEQUENCE(2,1)' })
    const b5 = mkCell('B5', 1, 4, { formula: 'A2' })
    const doc = docOf([mkSheet('S', [a1, b5])])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe(1)
    expect(valueAt(doc, 'S', 0, 1)).toBe(2)
    expect(valueAt(doc, 'S', 1, 4)).toBe(2)

    a1.formula = 'SEQUENCE(1,1)'
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe(1)
    expect(cellAt(doc, 'S', 0, 1)).toBeUndefined() // owned tail cleared
    expect(valueAt(doc, 'S', 1, 4)).toBe(0) // tail dependent -> absent 0
    expect(a1.formula).toBe('SEQUENCE(1,1)') // anchor formula never cleared
  })

  it('replacing a generated follower with a new input object never clears it (collision instead)', () => {
    const a1 = mkCell('A1', 0, 0, { formula: 'SEQUENCE(2,1)' })
    const doc = docOf([mkSheet('S', [a1])])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 1)).toBe(2)

    const sheet = doc.sheets[0]
    const row = sheet.rows.find((r) => r.index === 1)!
    const idx = row.cells.findIndex((c) => c.col === 0)
    row.cells[idx] = mkCell('A2', 0, 1, { value: 99, hasCachedValue: true })

    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 1)).toBe(99) // preserved input, not cleared
    expect(valueAt(doc, 'S', 0, 0)).toBe('#SPILL!') // owner collides
    expect(featureMessage(doc, 'spill-collision')).toContain('input-cell')
  })

  it('a generated follower whose value was replaced in place is preserved, not cleared', () => {
    const a1 = mkCell('A1', 0, 0, { formula: 'SEQUENCE(2,1)' })
    const doc = docOf([mkSheet('S', [a1])])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    const generated = cellAt(doc, 'S', 0, 1)!
    generated.value = 99 // same object, now user data

    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 1)).toBe(99)
    expect(valueAt(doc, 'S', 0, 0)).toBe('#SPILL!')
  })
})

// ---------------------------------------------------------------------------
// 4. Typed outputs: genuine error vs quoted text; atomic publication
// ---------------------------------------------------------------------------

describe('C2 phase2 spill — typed follower values', () => {
  it('a genuine #N/A element and a quoted "#N/A" text keep their distinct flags', () => {
    const doc = docOf([
      mkSheet('S', [mkCell('A1', 0, 0, { formula: '{1,#N/A;2,"#N/A"}' })]),
    ])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe(1)
    expect(valueAt(doc, 'S', 1, 0)).toBe('#N/A')
    expect(cellAt(doc, 'S', 1, 0)?.valueIsError).toBe(true) // genuine error
    expect(valueAt(doc, 'S', 0, 1)).toBe(2)
    expect(valueAt(doc, 'S', 1, 1)).toBe('#N/A')
    expect(cellAt(doc, 'S', 1, 1)?.valueIsError).toBe(false) // quoted text lookalike
  })

  it('never stores a matrix object as an XlsxCell value', () => {
    const doc = docOf([mkSheet('S', [mkCell('A1', 0, 0, { formula: 'SEQUENCE(2,3)' })])])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    for (const row of doc.sheets[0].rows) {
      for (const cell of row.cells) {
        expect(typeof cell.value === 'object' && cell.value !== null).toBe(false)
      }
    }
  })

  it('a blocked owner publishes no partial followers even though the matrix is complete', () => {
    const doc = docOf([
      mkSheet('S', [
        mkCell('A1', 0, 0, { formula: 'SEQUENCE(3,3)' }),
        mkCell('C1', 2, 0, { value: 'block', hasCachedValue: true }),
      ]),
    ])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe('#SPILL!')
    expect(cellAt(doc, 'S', 1, 0)).toBeUndefined()
    expect(cellAt(doc, 'S', 0, 1)).toBeUndefined()
    expect(cellAt(doc, 'S', 1, 1)).toBeUndefined()
  })
})

// ---------------------------------------------------------------------------
// 5. Legacy fixed f.t=array / f.ref projection
// ---------------------------------------------------------------------------

describe('C2 phase2 spill — legacy fixed array ranges', () => {
  it('maps a 2x2 legacy array literal into its declared f.ref rectangle', () => {
    const a1 = mkCell('A1', 0, 0, { formula: '{1,2;3,4}', arrayRef: 'A1:B2' })
    const doc = docOf([mkSheet('S', [a1])])
    evaluateWorkbookFormulas(doc)
    expect(valueAt(doc, 'S', 0, 0)).toBe(1)
    expect(valueAt(doc, 'S', 1, 0)).toBe(2)
    expect(valueAt(doc, 'S', 0, 1)).toBe(3)
    expect(valueAt(doc, 'S', 1, 1)).toBe(4)
    expect(a1.formula).toBe('{1,2;3,4}') // canonical source retained
    expect(a1.arrayRef).toBe('A1:B2')
  })

  it('evaluates the legacy formula once at the anchor context (references resolve there)', () => {
    const doc = docOf([
      mkSheet('S', [
        mkCell('A1', 0, 0, { formula: 'TRANSPOSE(A5:A6)', arrayRef: 'A1:B1' }),
        mkCell('A5', 0, 4, { value: 7, hasCachedValue: true }),
        mkCell('A6', 0, 5, { value: 8, hasCachedValue: true }),
      ]),
    ])
    evaluateWorkbookFormulas(doc)
    expect(valueAt(doc, 'S', 0, 0)).toBe(7)
    expect(valueAt(doc, 'S', 1, 0)).toBe(8)
  })

  it('gates a declared output that does not match the evaluated matrix instead of inventing an error', () => {
    const a1 = mkCell('A1', 0, 0, { formula: '{1,2;3,4}', arrayRef: 'A1:C3' })
    const doc = docOf([mkSheet('S', [a1])])
    evaluateWorkbookFormulas(doc)
    expect(hasFeature(doc, 'legacy-array-output-mismatch')).toBe(true)
    // no invented #VALUE! and no invented padding into the declared rectangle
    expect(valueAt(doc, 'S', 0, 0)).toBe(1)
    expect(cellAt(doc, 'S', 2, 0)).toBeUndefined()
    expect(cellAt(doc, 'S', 0, 2)).toBeUndefined()
  })
})

// ---------------------------------------------------------------------------
// 6. Twelve non-overlapping owners must not exhaust the unstable retry budget
// ---------------------------------------------------------------------------

describe('C2 phase2 spill — many non-overlapping owners', () => {
  it('A1..L1 SEQUENCE(2,1) all spill and M1=SUM(A2:L2)=24 without an unstable-retry diagnostic', () => {
    const cells: XlsxCell[] = []
    for (let i = 0; i < 12; i++) {
      cells.push(mkCell(`${String.fromCharCode(65 + i)}1`, i, 0, { formula: 'SEQUENCE(2,1)' }))
    }
    cells.push(mkCell('M1', 12, 0, { formula: 'SUM(A2:L2)' }))
    const doc = docOf([mkSheet('S', cells)])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    for (let i = 0; i < 12; i++) {
      expect(valueAt(doc, 'S', i, 0)).toBe(1)
      expect(valueAt(doc, 'S', i, 1)).toBe(2)
    }
    expect(valueAt(doc, 'S', 12, 0)).toBe(24)
    expect(hasFeature(doc, 'reference-generation-unstable')).toBe(false)
    expect(hasFeature(doc, 'spill-collision')).toBe(false)
  })
})

// ---------------------------------------------------------------------------
// ROW / COLUMN documented array-mode geometry vectors
// ---------------------------------------------------------------------------

describe('C2 phase2 spill — ROW/COLUMN documented array-mode vectors', () => {
  it('ROW(C5:C7) spills the vertical vector [[5],[6],[7]] in array mode', () => {
    const doc = docOf([mkSheet('S', [mkCell('A1', 0, 0, { formula: 'ROW(C5:C7)' })])])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe(5)
    expect(valueAt(doc, 'S', 0, 1)).toBe(6)
    expect(valueAt(doc, 'S', 0, 2)).toBe(7)
    expect(cellAt(doc, 'S', 1, 0)).toBeUndefined()
  })

  it('COLUMN(B3:D9) spills the horizontal vector [[2,3,4]] in array mode', () => {
    const doc = docOf([mkSheet('S', [mkCell('A1', 0, 0, { formula: 'COLUMN(B3:D9)' })])])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe(2)
    expect(valueAt(doc, 'S', 1, 0)).toBe(3)
    expect(valueAt(doc, 'S', 2, 0)).toBe(4)
    expect(cellAt(doc, 'S', 0, 1)).toBeUndefined()
  })

  it('scalar ROW/COLUMN keep the first-row / leftmost-column result', () => {
    const doc = docOf([
      mkSheet('S', [
        mkCell('A1', 0, 0, { formula: 'ROW(C5:C7)' }),
        mkCell('B1', 1, 0, { formula: 'COLUMN(B3:D9)' }),
      ]),
    ])
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(valueAt(doc, 'S', 0, 0)).toBe(5)
    expect(valueAt(doc, 'S', 1, 0)).toBe(2)
  })
})

// ---------------------------------------------------------------------------
// Value-dependent discovery: FILTER shape, selected IF, reference-valued names
// ---------------------------------------------------------------------------

describe('C2 phase2 spill — value-dependent ownership discovery', () => {
  it('FILTER shape follows the actual include values (suspends and resumes cooperatively)', () => {
    const doc = docOf([
      mkSheet('S', [
        mkCell('A1', 0, 0, { formula: 'FILTER(A5:A7,B5:B7>1)' }),
        mkCell('A5', 0, 4, { value: 10, hasCachedValue: true }),
        mkCell('A6', 0, 5, { value: 20, hasCachedValue: true }),
        mkCell('A7', 0, 6, { value: 30, hasCachedValue: true }),
        mkCell('B5', 1, 4, { value: 1, hasCachedValue: true }),
        mkCell('B6', 1, 5, { value: 2, hasCachedValue: true }),
        mkCell('B7', 1, 6, { value: 3, hasCachedValue: true }),
      ]),
    ])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe(20)
    expect(valueAt(doc, 'S', 0, 1)).toBe(30)
    expect(cellAt(doc, 'S', 0, 2)).toBeUndefined()
  })

  it('FILTER over formula-valued criteria suspends on the dependent cells', () => {
    const doc = docOf([
      mkSheet('S', [
        mkCell('A1', 0, 0, { formula: 'FILTER(A5:A7,B5:B7>1)' }),
        mkCell('A5', 0, 4, { value: 10, hasCachedValue: true }),
        mkCell('A6', 0, 5, { value: 20, hasCachedValue: true }),
        mkCell('A7', 0, 6, { value: 30, hasCachedValue: true }),
        mkCell('B5', 1, 4, { formula: '0+1' }),
        mkCell('B6', 1, 5, { formula: '1+1' }),
        mkCell('B7', 1, 6, { formula: '1+2' }),
      ]),
    ])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe(20)
    expect(valueAt(doc, 'S', 0, 1)).toBe(30)
  })

  it('IF(FALSE, SEQUENCE(2,3), 8) is a scalar outcome: no spill, no follower', () => {
    const doc = docOf([mkSheet('S', [mkCell('A1', 0, 0, { formula: 'IF(FALSE,SEQUENCE(2,3),8)' })])])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe(8)
    expect(cellAt(doc, 'S', 1, 0)).toBeUndefined()
    expect(cellAt(doc, 'S', 0, 1)).toBeUndefined()
    expect(hasFeature(doc, 'spill-collision')).toBe(false)
  })

  it('an untaken IF branch with an unknown function is never evaluated', () => {
    const doc = docOf([mkSheet('S', [mkCell('A1', 0, 0, { formula: 'IF(FALSE,NOSUCHFUNC(1),8)' })])])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe(8)
    expect(hasFeature(doc, 'NOSUCHFUNC')).toBe(false)
    expect(hasFeature(doc, 'formula-syntax')).toBe(false)
  })

  it('SEQUENCE(1,1) is a real 1x1 array (A1# resolves) while IF(FALSE,...) is scalar (A1# gated)', () => {
    const arrayDoc = docOf([
      mkSheet('S', [
        mkCell('A1', 0, 0, { formula: 'SEQUENCE(1,1)' }),
        mkCell('B1', 1, 0, { formula: 'ROWS(A1#)' }),
      ]),
    ])
    evaluateWorkbookFormulas(arrayDoc, DYNAMIC)
    expect(valueAt(arrayDoc, 'S', 0, 0)).toBe(1)
    expect(valueAt(arrayDoc, 'S', 1, 0)).toBe(1)

    const scalarDoc = docOf([
      mkSheet('S', [
        mkCell('A1', 0, 0, { formula: 'IF(FALSE,SEQUENCE(2,3),8)' }),
        mkCell('B1', 1, 0, { formula: 'ROWS(A1#)' }),
      ]),
    ])
    evaluateWorkbookFormulas(scalarDoc, DYNAMIC)
    expect(valueAt(scalarDoc, 'S', 1, 0)).toBe('#NAME?')
    expect(hasFeature(scalarDoc, 'spill-provenance')).toBe(true)
  })

  it('an array-valued defined name spills through the name node', () => {
    const doc = docOf(
      [mkSheet('S', [mkCell('A1', 0, 0, { formula: 'Seq' })])],
      { definedNames: [{ name: 'Seq', source: 'SEQUENCE(2,3)', baseProvenance: 'explicit' }] },
    )
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe(1)
    expect(valueAt(doc, 'S', 1, 0)).toBe(2)
    expect(valueAt(doc, 'S', 2, 1)).toBe(6)
  })

  it('a reference-valued defined name spills its full rectangle', () => {
    const doc = docOf(
      [
        mkSheet('S', [
          mkCell('A1', 0, 0, { formula: 'Rng' }),
          mkCell('A5', 0, 4, { value: 1, hasCachedValue: true }),
          mkCell('B5', 1, 4, { value: 2, hasCachedValue: true }),
          mkCell('A6', 0, 5, { value: 3, hasCachedValue: true }),
          mkCell('B6', 1, 5, { value: 4, hasCachedValue: true }),
        ]),
      ],
      { definedNames: [{ name: 'Rng', source: '$A$5:$B$6', baseProvenance: 'explicit' }] },
    )
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe(1)
    expect(valueAt(doc, 'S', 1, 0)).toBe(2)
    expect(valueAt(doc, 'S', 0, 1)).toBe(3)
    expect(valueAt(doc, 'S', 1, 1)).toBe(4)
  })
})

// ---------------------------------------------------------------------------
// Pending / hook cooperation and cached provenance
// ---------------------------------------------------------------------------

describe('C2 phase2 spill — pending cooperation and cached provenance', () => {
  it('an owner that materializes a reference suspends on formula dependencies and resumes', () => {
    const doc = docOf([
      mkSheet('S', [
        mkCell('A1', 0, 0, { formula: 'A5:A6' }),
        mkCell('A5', 0, 4, { formula: '1+1' }),
        mkCell('A6', 0, 5, { formula: '2+2' }),
      ]),
    ])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe(2)
    expect(valueAt(doc, 'S', 0, 1)).toBe(4)
  })

  it('a consumer expression argument is evaluated in array mode through the hook', () => {
    const doc = docOf([
      mkSheet('S', [
        mkCell('A1', 0, 0, { formula: 'TRANSPOSE(A5:A6+1)' }),
        mkCell('A5', 0, 4, { value: 7, hasCachedValue: true }),
        mkCell('A6', 0, 5, { value: 8, hasCachedValue: true }),
      ]),
    ])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe(8)
    expect(valueAt(doc, 'S', 1, 0)).toBe(9)
  })

  it('from-file mode keeps a cached dynamic-looking formula scalar with no invented owner', () => {
    const doc = docOf([
      mkSheet('S', [mkCell('A1', 0, 0, { formula: 'SEQUENCE(2,3)', value: 1, hasCachedValue: true })]),
    ])
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(valueAt(doc, 'S', 0, 0)).toBe(1)
    expect(cellAt(doc, 'S', 1, 0)).toBeUndefined()
    expect(hasFeature(doc, 'spill-collision')).toBe(false)
  })

  it('from-file mode still projects a legacy fixed array range', () => {
    const doc = docOf([mkSheet('S', [mkCell('A1', 0, 0, { formula: '{1,2;3,4}', arrayRef: 'A1:B2' })])])
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(valueAt(doc, 'S', 0, 0)).toBe(1)
    expect(valueAt(doc, 'S', 1, 1)).toBe(4)
  })
})

// ---------------------------------------------------------------------------
// Arbitration gate, generation invalidation and cursor cleanup
// ---------------------------------------------------------------------------

describe('C2 phase2 spill — arbitration gate and generation cleanup', () => {
  it('competing owners with non-anchor overlap gate both (unavailable arbitration, caches retained, no followers)', () => {
    // A2=SEQUENCE(2,4) -> A2:D3 and C1=SEQUENCE(3,2) -> C1:D3 intersect at
    // C2:D3 with NEITHER anchor inside the other's footprint: genuinely unknown
    // arbitration, not a known formula blockage.
    function build(forceA2First: boolean): XlsxDocument {
      const a2 = mkCell('A2', 0, 1, { formula: 'SEQUENCE(2,4)', value: 77, hasCachedValue: true })
      const c1 = mkCell('C1', 2, 0, { formula: 'SEQUENCE(3,2)', value: 88, hasCachedValue: true })
      const cells = forceA2First
        ? [mkCell('B1', 1, 0, { formula: 'A3' }), c1, a2]
        : [c1, a2]
      return docOf([mkSheet('S', cells)])
    }
    for (const forceA2First of [false, true]) {
      const doc = build(forceA2First)
      evaluateWorkbookFormulas(doc, DYNAMIC)
      // Unknown arbitration is an explicit UNAVAILABLE gate (internal #NAME? +
      // frame marks), never a fabricated native #SPILL! and never a winner by
      // publication order.
      expect(hasFeature(doc, 'spill-arbitration')).toBe(true)
      expect(valueAt(doc, 'S', 0, 1)).toBe(77) // A2 valid cache retained
      expect(valueAt(doc, 'S', 2, 0)).toBe(88) // C1 valid cache retained
      expect(valueAt(doc, 'S', 0, 1)).not.toBe('#SPILL!')
      expect(valueAt(doc, 'S', 2, 0)).not.toBe('#SPILL!')
      // no generated followers from either anchor
      expect(cellAt(doc, 'S', 1, 1)).toBeUndefined()
      expect(cellAt(doc, 'S', 3, 1)).toBeUndefined()
      expect(cellAt(doc, 'S', 2, 2)).toBeUndefined()
      expect(cellAt(doc, 'S', 3, 2)).toBeUndefined()
    }
  })

  it('a formula anchor inside another owner footprint is a known blockage, never arbitration', () => {
    function build(order: 'a-first' | 'b-first'): XlsxDocument {
      const a1 = mkCell('A1', 0, 0, { formula: 'SEQUENCE(2,2)' })
      const b1 = mkCell('B1', 1, 0, { formula: 'SEQUENCE(2,2)' })
      return docOf([mkSheet('S', order === 'b-first' ? [b1, a1] : [a1, b1])])
    }
    for (const order of ['a-first', 'b-first'] as const) {
      const doc = build(order)
      evaluateWorkbookFormulas(doc, DYNAMIC)
      expect(hasFeature(doc, 'spill-arbitration')).toBe(false)
      expect(valueAt(doc, 'S', 0, 0)).toBe('#SPILL!') // A1 genuine known blockage
      expect(valueAt(doc, 'S', 1, 0)).toBe(1) // B1 still spills 1..4
      expect(valueAt(doc, 'S', 2, 0)).toBe(2)
      expect(valueAt(doc, 'S', 1, 1)).toBe(3)
      expect(valueAt(doc, 'S', 2, 1)).toBe(4)
    }
  })

  it('shrinking an owner invalidates a SUM(A1#) consumer and clears the old tail cursor', () => {
    const a1 = mkCell('A1', 0, 0, { formula: 'SEQUENCE(3,1)' })
    const b1 = mkCell('B1', 1, 0, { formula: 'SUM(A1#)' })
    const doc = docOf([mkSheet('S', [a1, b1])])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 1, 0)).toBe(6)
    expect(valueAt(doc, 'S', 0, 2)).toBe(3)

    a1.formula = 'SEQUENCE(2,1)'
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 1, 0)).toBe(3)
    expect(cellAt(doc, 'S', 0, 2)).toBeUndefined()
  })

  it('repeated recalculation is stable and publishes no duplicate followers', () => {
    const doc = docOf([
      mkSheet('S', [
        mkCell('A1', 0, 0, { formula: 'SEQUENCE(2,3)' }),
        mkCell('D1', 3, 0, { formula: 'C2' }),
      ]),
    ])
    for (let i = 0; i < 3; i++) evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(valueAt(doc, 'S', 0, 0)).toBe(1)
    expect(valueAt(doc, 'S', 2, 1)).toBe(6)
    expect(valueAt(doc, 'S', 3, 0)).toBe(6)
    const row1 = doc.sheets[0].rows.find((r) => r.index === 1)!
    expect(row1.cells.filter((c) => c.col === 0).length).toBe(1)
  })
})

// ---------------------------------------------------------------------------
// Discovery/read boundary (root root-phase2-preview/discovery-boundaries)
// ---------------------------------------------------------------------------

describe('C2 phase2 spill — discovery/read boundary', () => {
  it('unresolved self-owned shape discovery never turns an absent possible follower into scalar zero', () => {
    const anchor = mkCell('A1', 0, 0, { formula: 'SEQUENCE(B2,2)', value: 77, hasCachedValue: true })
    const doc = docOf([mkSheet('S', [anchor])])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(anchor.value).toBe(77)
    expect((doc.diagnostics ?? []).some((x) => x.kind === 'spill-discovery' || (x.feature ?? '').includes('discovery'))).toBe(true)
    expect(doc.sheets[0].rows.flatMap((r) => r.cells).filter((x) => x.ref !== 'A1')).toHaveLength(0)
  })

  it('an unavailable arbitrated footprint does not become a memoized blank for a cached follower reader', () => {
    const a = mkCell('A2', 0, 1, { formula: 'SEQUENCE(2,4)', value: 77, hasCachedValue: true })
    const b = mkCell('C1', 2, 0, { formula: 'SEQUENCE(3,2)', value: 88, hasCachedValue: true })
    const reader = mkCell('E1', 4, 0, { formula: 'C2', value: 99, hasCachedValue: true })
    const doc = docOf([mkSheet('S', [b, reader, a])])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(a.value).toBe(77)
    expect(b.value).toBe(88)
    expect(reader.value).toBe(99)
    expect((doc.diagnostics ?? []).some((x) => x.feature === 'spill-arbitration')).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// Own-only cleanup must preserve original blank cells/rows and user edits
// (root root-phase2-preview/cleanup-content)
// ---------------------------------------------------------------------------

describe('C2 phase2 spill — own-only cleanup content preservation', () => {
  function setup(extra: XlsxCell[] = []): { d: XlsxDocument; a: XlsxCell } {
    const a: XlsxCell = { ref: 'A1', col: 0, row: 0, formula: 'SEQUENCE(2,2)', value: null, styleIndex: 0 }
    return {
      a,
      d: docOf([mkSheet('S', [a], {
        rows: [
          { index: 0, cells: [a] },
          { index: 1, heightPt: 27, customHeight: true, cells: extra },
        ],
      })]),
    }
  }

  it('shrinking clears generated values while retaining an original blank styled cell', () => {
    const blank: XlsxCell = { ref: 'B2', col: 1, row: 1, value: null, styleIndex: 7, hasCachedValue: false }
    const { d, a } = setup([blank])
    evaluateWorkbookFormulas(d, DYNAMIC)
    expect(cellAt(d, 'S', 1, 1)?.value).toBe(4)
    expect(cellAt(d, 'S', 1, 1)?.styleIndex).toBe(7)
    a.formula = 'SEQUENCE(1,1)'
    evaluateWorkbookFormulas(d, DYNAMIC)
    expect(cellAt(d, 'S', 1, 1)).toBeDefined()
    expect(cellAt(d, 'S', 1, 1)?.value).toBe(null)
    expect(cellAt(d, 'S', 1, 1)?.styleIndex).toBe(7)
  })

  it('cleanup preserves original formatted empty row metadata', () => {
    const { d, a } = setup()
    evaluateWorkbookFormulas(d, DYNAMIC)
    a.formula = 'SEQUENCE(1,1)'
    evaluateWorkbookFormulas(d, DYNAMIC)
    expect(d.sheets[0].rows.find((r) => r.index === 1)?.heightPt).toBe(27)
    expect(d.sheets[0].rows.find((r) => r.index === 1)?.customHeight).toBe(true)
  })

  it('a generated cell changed into a formula is user input even if its cached value is unchanged', () => {
    const { d, a } = setup()
    evaluateWorkbookFormulas(d, DYNAMIC)
    const b = cellAt(d, 'S', 1, 1)!
    expect(b.value).toBe(4)
    b.formula = '99'
    a.formula = 'SEQUENCE(1,1)'
    evaluateWorkbookFormulas(d, DYNAMIC)
    expect(cellAt(d, 'S', 1, 1)?.formula).toBe('99')
    expect(cellAt(d, 'S', 1, 1)?.value).toBe(99)
  })
})
