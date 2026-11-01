// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { evaluateWorkbookFormulas, recomputeDependents } from '../src/xlsx/formula/workbook'
import { registerFormulaFunction } from '../src/xlsx/formula/functions'
import { markPagedDependentsStale, noteWorkbookCellsChanged, restoreDependentFormulaValues, savedDependentCells } from '../src/xlsx/formula/workbook'
import { mergeOwedRects, type OwedRect } from '../src/xlsx/pivots/session/dependentRects'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../src/xlsx/types'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { readFileSync } from 'node:fs'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { loadAll } from './helpers/xlsx-pivots-load'
import { openPivotSession } from '../src/xlsx/pivots/session/session'
import { sheetCellReader } from '../src/xlsx/pivots/engine/faithful'
import { getPivotData, registerGetPivotData } from '../src/xlsx/pivots/session/getpivotdata'
import type { ChartModel } from '../src/charts/model'
import { existsSync } from 'node:fs'
import { join } from 'node:path'
import { eagerModules, type ModuleMap } from './helpers/eager-graph'

const FIXTURES = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))

function sheet(name: string, cells: Array<{ ref: string; col: number; row: number; value: string | number | null; formula?: string }>): XlsxSheet {
  const rows = new Map<number, XlsxSheet['rows'][number]>()
  for (const cell of cells) {
    let row = rows.get(cell.row)
    if (!row) rows.set(cell.row, row = { index: cell.row, cells: [] })
    row.cells.push({ ...cell, styleIndex: 0, hasCachedValue: cell.value !== null })
  }
  return { name, sheetId: name, workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows: [...rows.values()] }
}

function dependentDoc(valuePolicy?: XlsxDocument['valuePolicy']): XlsxDocument {
  return {
    valuePolicy,
    sheets: [
      sheet('Pivot', [{ ref: 'A1', col: 0, row: 0, value: 10 }]),
      sheet('Calc', [
        { ref: 'A1', col: 0, row: 0, value: 0, formula: "='Pivot'!A1*2" },
        { ref: 'A2', col: 0, row: 1, value: 0, formula: '=A1+1' },
        { ref: 'B1', col: 1, row: 0, value: 77, formula: '=40+37' },
      ]),
    ], definedNames: [], tables: [],
  }
}

describe('pivot dependent formulas (Task 31)', () => {
  it('recomputes the transitive readers of changed pivot cells and leaves unrelated formulas alone', () => {
    const doc = dependentDoc()
    doc.sheets[0]!.rows[0]!.cells.push({ ref: 'B1', col: 1, row: 0, value: 5, styleIndex: 0, hasCachedValue: true })
    doc.sheets[1]!.rows.find(row => row.index === 1)!.cells.push({ ref: 'B2', col: 1, row: 1, value: 0, formula: "='Pivot'!B1", styleIndex: 0, hasCachedValue: true })
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    const calc = doc.sheets[1]!
    expect(calc.rows[0]!.cells[0]!.value).toBe(20)
    expect(calc.rows[1]!.cells[0]!.value).toBe(21)

    const result = recomputeDependents(doc, [{ sheetId: 0, rect: { minRow: 0, minCol: 0, maxRow: 0, maxCol: 1 } }], {
      forceRecalc: true, readCellValue: (name, col, row) => name === 'Pivot' && col === 0 && row === 0 ? 15 : name === 'Pivot' && col === 1 && row === 0 ? null : undefined,
    })
    expect(result.recomputed).toBe(3)
    expect(calc.rows[0]!.cells[0]!.value).toBe(30)
    expect(calc.rows[1]!.cells[0]!.value).toBe(31)
    expect(calc.rows[0]!.cells[1]!.value).toBe(77)
    // A direct reference to an empty pivot cell evaluates as zero in Excel.
    expect(calc.rows[1]!.cells[1]!.value).toBe(0)
  })

  it('throws under the saved-values policy', () => {
    const doc = dependentDoc('cached')
    expect(() => recomputeDependents(doc, [], {})).toThrow(/saved values only/)
  })

  it('supports a lazy formula-function registration hook', () => {
    registerFormulaFunction('TASK31_DOUBLE', (args, _ctx, evalNode) => {
      const value = evalNode(args[0]!)
      return typeof value === 'number' ? value * 2 : value
    })
    const doc: XlsxDocument = {
      sheets: [sheet('S', [{ ref: 'A1', col: 0, row: 0, value: 0, formula: '=TASK31_DOUBLE(6)' }])],
      definedNames: [], tables: [],
    }
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(doc.sheets[0]!.rows[0]!.cells[0]!.value).toBe(12)
  })

  it('reports charts whose referenced ranges include a recomputed formula cell', () => {
    const doc = dependentDoc()
    const chart = { categoryRef: { sheet: 'Calc', r0: 0, c0: 0, r1: 1, c1: 0 }, groups: [] } as unknown as ChartModel
    doc.sheets[1]!.drawings = [{ content: { kind: 'xlsxChart', chart } } as NonNullable<XlsxSheet['drawings']>[number]]
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    doc.sheets[0]!.rows[0]!.cells[0]!.value = 11
    const result = recomputeDependents(doc, [{ sheetId: 0, rect: { minRow: 0, minCol: 0, maxRow: 0, maxCol: 0 } }], { forceRecalc: true })
    expect(result.charts).toEqual([0])
  })

  it('marks paged dependent sheets once, and ignores formulas inside pivots', () => {
    const calc = sheet('S', [
      { ref: 'A1', col: 0, row: 0, value: 1, formula: '=1+0' },
      { ref: 'A2', col: 0, row: 1, value: 2, formula: '=1+1' },
    ])
    const pivot = [{ minRow: 0, minCol: 0, maxRow: 0, maxCol: 0 }]
    const first = markPagedDependentsStale(new Set(), 2, calc, pivot)
    expect(first).toEqual(new Set([2]))
    expect(markPagedDependentsStale(first, 2, calc, pivot)).toBe(first)
    expect(markPagedDependentsStale(new Set(), 3, calc, [{ minRow: 0, minCol: 0, maxRow: 1, maxCol: 0 }])).toEqual(new Set())
  })

  it('keeps a saved GETPIVOTDATA value until a pivot on its sheet changes', async () => {
    registerGetPivotData()
    const file = resolve(FIXTURES, 'recipes/getpivotdata-seed.xlsx')
    const bytes = new Uint8Array(readFileSync(file))
    const pkg = await OfficePackage.load(bytes)
    const doc = await parseXlsx(pkg)
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    const formulas = doc.sheets.find(sheet => sheet.name === 'Formulas')!
    const total = formulas.rows.flatMap(row => row.cells).find(cell => cell.formula?.startsWith('GETPIVOTDATA'))!
    expect(total.value).toBe(222603.36000000004)
    expect(total.valueIsError).not.toBe(true)
  })

  it('matches every saved GETPIVOTDATA oracle case against its live visible layout', async () => {
    registerGetPivotData()
    const names = ['getpivotdata-seed', 'gpd-baseline', 'gpd-hide-item', 'gpd-collapse-item', 'gpd-filter-out']
    for (const name of names) {
      const folder = name === 'getpivotdata-seed' ? 'recipes' : 'interaction'
      const file = resolve(FIXTURES, folder, `${name}.xlsx`)
      const bytes = new Uint8Array(readFileSync(file))
      const pkg = await OfficePackage.load(bytes)
      const doc = await parseXlsx(pkg)
      const loaded = (await loadAll(file))[0]!
      const saved = sheetCellReader(loaded.sheet)
      const opened = openPivotSession({ def: loaded.def, cache: loaded.cache, store: loaded.store, date1904: loaded.date1904, saved,
        occupied: () => false, linkedCharts: [] })
      expect(opened.ok, `${name} session opens`).toBe(true)
      if (!opened.ok) continue
      const layout = opened.value.layout
      const formulaSheet = doc.sheets.find(sheet => sheet.name === 'Formulas')!
      const expected = formulaSheet.rows.flatMap(row => row.cells).filter(cell => cell.formula?.toUpperCase().startsWith('GETPIVOTDATA('))
      const expectedValues = expected.map(cell => ({ row: cell.row, col: cell.col, value: cell.value, error: cell.valueIsError }))
      for (const cell of expected) cell.value = -123456789 // prove the live evaluator replaces the saved cache
      evaluateWorkbookFormulas(doc, { forceRecalc: true }) // unknown GPD retains the Excel cache
      const result = recomputeDependents(doc, [{ sheetId: doc.sheets.findIndex(sheet => sheet.name === loaded.sheet.name), rect: {
        minRow: layout.rect.minRow, minCol: layout.rect.minCol, maxRow: layout.rect.maxRow, maxCol: layout.rect.maxCol,
      } }], { forceRecalc: true, getPivotData: (args, evaluate) =>
        getPivotData([{ sheet: loaded.sheet.name, def: opened.value.def, cache: loaded.cache, layout, date1904: loaded.date1904 }], args, evaluate, 'Formulas') })
      expect(result.recomputed, name).toBeGreaterThan(0)
      for (const cell of expectedValues) {
        const actual = formulaSheet.rows.find(row => row.index === cell.row)?.cells.find(item => item.col === cell.col)
        if (cell.error) expect(actual?.value).toBe(cell.value)
        else if (typeof cell.value === 'number') expect(actual?.value as number).toBeCloseTo(cell.value, 9)
        else expect(actual?.value).toBe(cell.value)
      }
      opened.value.release()
    }
  })

  it('the built eager formula graph contains no pivot implementation module', () => {
    const moduleMapFile = join(process.cwd(), 'dist/.module-map.json')
    if (!existsSync(moduleMapFile)) return // The required post-build check-eager gate verifies this in CI.
    const map = JSON.parse(readFileSync(moduleMapFile, 'utf8')) as ModuleMap
    const pivotModules = [...eagerModules(map)].map(id => id.replace(/\\/g, '/')).filter(id =>
      /src\/xlsx\/pivots\/(?!stub\.ts(?:$|\?))/.test(id))
    expect(pivotModules).toEqual([])
  })
})

describe('stale readers and edited cells', () => {
  const staleCell = (ref: string, col: number, value: number, formula?: string) => ({ ref, col, row: 0, value, ...(formula ? { formula } : {}), styleIndex: 0, hasCachedValue: true })
  function staleDocument(formula: string): XlsxDocument {
    return { sheets: [
      { name: 'Pivot', sheetId: 'P', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [staleCell('A1', 0, 1)] }] },
      { name: 'Calc', sheetId: 'C', workbookIndex: 1, cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [staleCell('A1', 0, 77, formula), staleCell('B1', 1, 154, 'A1*2')] }] },
    ], definedNames: [], tables: [] }
  }

  it('S3-T: a formula reading a retained stale formula is also stale and keeps its own cached value', () => {
    const doc = staleDocument('INDIRECT("Pivot!A1")')
    evaluateWorkbookFormulas(doc)
    const result = recomputeDependents(doc, [{ sheetId: 0, rect: { minRow: 0, minCol: 0, maxRow: 0, maxCol: 0 } }], { readCellValue: () => 5 })
    expect(result.stale).toEqual(expect.arrayContaining(['calc!0:0', 'calc!1:0']))
    expect(doc.sheets[1]!.rows[0]!.cells.map(item => item.value)).toEqual([77, 154])
  })

  it('S3-AJ: an announced in-place formula edit rebuilds the syntactic and runtime cell indexes', () => {
    const doc = staleDocument('Pivot!A1*2')
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    const formula = doc.sheets[1]!.rows[0]!.cells[0]!
    formula.formula = 'Pivot!B1*2'
    doc.sheets[0]!.rows[0]!.cells.push(staleCell('B1', 1, 9))
    // T31 r4 M-1: a direct edit of the cell model is announced; the commit path no longer walks every cell to find it.
    noteWorkbookCellsChanged(doc)
    const result = recomputeDependents(doc, [{ sheetId: 0, rect: { minRow: 0, minCol: 1, maxRow: 0, maxCol: 1 } }])
    expect(result.recomputed).toBe(2)
    expect(formula.value).toBe(18)
  })
})

describe('the rectangles a pivot owes its dependents', () => {
  const owedKey = (rect: OwedRect) => `${rect.r0}:${rect.r1}:${rect.c0}:${rect.c1}`

  it('S3-AE: 20 consecutive over-budget commits leave the owed rectangles merged, with no duplicates', () => {
    const saved = { r0: 2, r1: 23, c0: 0, c1: 8 }, collapsed = { r0: 2, r1: 18, c0: 0, c1: 8 }, grown = { r0: 2, r1: 30, c0: 0, c1: 8 }
    let owed: OwedRect[] = []
    // Each commit owes saved ∪ before ∪ current; nothing ran, so the owed set carries into the next commit.
    for (let commit = 0; commit < 20; commit++) {
      const before = commit % 2 ? collapsed : grown, current = commit % 2 ? grown : collapsed
      owed = mergeOwedRects(owed, [saved, before, current])
      expect(new Set(owed.map(owedKey)).size).toBe(owed.length)
      expect(owed.length).toBeLessThanOrEqual(3)
    }
    // The grown area contains the other two.
    expect(owed).toEqual([grown])
  })

  it('S3-AE: disjoint owed areas are all kept, once each', () => {
    const a = { r0: 0, r1: 4, c0: 0, c1: 4 }, b = { r0: 10, r1: 12, c0: 10, c1: 12 }
    let owed: OwedRect[] = []
    for (let commit = 0; commit < 20; commit++) owed = mergeOwedRects(owed, [a, b, { ...a }])
    expect(owed.map(owedKey).sort()).toEqual([owedKey(a), owedKey(b)].sort())
  })
})

describe('a dependent spill that grows or shrinks (Ruling S3-AL)', () => {
  const spillCell = (ref: string, col: number, row: number, extra: Partial<XlsxCell> = {}): XlsxCell => ({ ref, col, row, value: null, styleIndex: 0, hasCachedValue: false, ...extra })
  /** Pivot!A1 = `rows`; Calc!A1 = SEQUENCE(Pivot!A1,1), spilling down column A; Calc!C1 reads the spill's third cell. */
  function spillDoc(rows: number): XlsxDocument {
    const pivot: XlsxSheet = { name: 'Pivot', sheetId: 'Pivot', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [spillCell('A1', 0, 0, { value: rows, hasCachedValue: true })] }] }
    const calc: XlsxSheet = { name: 'Calc', sheetId: 'Calc', workbookIndex: 1, cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [
      spillCell('A1', 0, 0, { formula: 'SEQUENCE(Pivot!A1,1)', dynamicArray: true }), spillCell('B1', 1, 0, { value: 'beside', hasCachedValue: true }),
    ] }] }
    const doc: XlsxDocument = { sheets: [pivot, calc], definedNames: [], tables: [] }
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    return doc
  }
  const column = (doc: XlsxDocument) => doc.sheets[1]!.rows.flatMap(row => row.cells.filter(cell => cell.col === 0).map(cell => [cell.ref, cell.value]))
  const commit = (doc: XlsxDocument, rows: number) => recomputeDependents(doc, [{ sheetId: 0, rect: { minRow: 0, maxRow: 0, minCol: 0, maxCol: 0 } }],
    { arrayMode: 'dynamic', readCellValue: (name, col, row) => name === 'Pivot' && row === 0 && col === 0 ? rows : undefined })
  /** What the print copy shows of Calc's column A (the viewer's withSavedCells, by position). */
  const printed = (doc: XlsxDocument) => {
    const calc = doc.sheets[1]!, saved = savedDependentCells(doc)?.get(calc)
    const at = new Map<string, unknown>()
    for (const row of calc.rows) for (const cell of row.cells) if (cell.col === 0) at.set(cell.ref, cell.value)
    for (const [row, cols] of saved ?? []) for (const [col, cell] of cols) if (col === 0) { const ref = `A${row + 1}`; if (cell) at.set(ref, cell.value); else at.delete(ref) }
    return [...at].sort(([a], [b]) => a.localeCompare(b, undefined, { numeric: true }))
  }

  it('grown: print and the restore show the saved two cells, not the live third; a later commit spills again', () => {
    const doc = spillDoc(2)
    const saved = column(doc)
    expect(saved).toEqual([['A1', 1], ['A2', 2]])
    commit(doc, 3)
    expect(column(doc)).toEqual([['A1', 1], ['A2', 2], ['A3', 3]])
    expect(printed(doc)).toEqual(saved)
    expect(savedDependentCells(doc)!.get(doc.sheets[1]!)!.get(2)!.get(0)).toBeNull()
    restoreDependentFormulaValues(doc)
    expect(column(doc)).toEqual(saved)
    expect(doc.sheets[1]!.rows.map(row => row.index)).toEqual([0, 1])   // the row the spill created goes with it
    expect(savedDependentCells(doc)).toBeUndefined()
    // The restored cells are the saved spill's own again: the next commit grows it, unblocked.
    commit(doc, 4)
    expect(column(doc)).toEqual([['A1', 1], ['A2', 2], ['A3', 3], ['A4', 4]])
    restoreDependentFormulaValues(doc)
    expect(column(doc)).toEqual(saved)
  })

  it('shrunk: print and the restore put back the cleared saved cells; a later commit shrinks it again', () => {
    const doc = spillDoc(3)
    const saved = column(doc)
    expect(saved).toEqual([['A1', 1], ['A2', 2], ['A3', 3]])
    commit(doc, 1)
    expect(column(doc)).toEqual([['A1', 1]])
    expect(printed(doc)).toEqual(saved)
    expect(savedDependentCells(doc)!.get(doc.sheets[1]!)!.get(2)!.get(0)).toMatchObject({ ref: 'A3', value: 3 })
    restoreDependentFormulaValues(doc)
    expect(column(doc)).toEqual(saved)
    expect(doc.sheets[1]!.rows[0]!.cells.find(cell => cell.col === 1)!.value).toBe('beside')
    commit(doc, 2)
    expect(column(doc)).toEqual([['A1', 1], ['A2', 2]])
    commit(doc, 3)   // a later commit of the same change: the saved state is still the first one
    expect(printed(doc)).toEqual(saved)
    restoreDependentFormulaValues(doc)
    expect(column(doc)).toEqual(saved)
  })
})

describe('readers of a dependent spill\'s cells, with numeric sheet ids', () => {
  const at = (ref: string, col: number, row: number, extra: Partial<XlsxCell> = {}): XlsxCell => ({ ref, col, row, value: null, styleIndex: 0, hasCachedValue: false, ...extra })
  /** Excel's sheet ids are numbers: the spill engine keys owners by them, the reader queue by the sheet's name. */
  function numericDoc(extraCalc: XlsxCell[] = []): XlsxDocument {
    const pivot: XlsxSheet = { name: 'Pivot', sheetId: '1', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [at('A1', 0, 0, { value: 10, hasCachedValue: true })] }] }
    const calc: XlsxSheet = { name: 'Calc', sheetId: '2', workbookIndex: 1, cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [
      at('A1', 0, 0, { formula: 'SEQUENCE(2,1,Pivot!A1)', dynamicArray: true }), at('C1', 2, 0, { formula: 'A2*2' }), ...extraCalc,
    ] }] }
    const doc: XlsxDocument = { sheets: [pivot, calc], definedNames: [], tables: [] }
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    return doc
  }
  const value = (doc: XlsxDocument, ref: string) => doc.sheets[1]!.rows.flatMap(row => row.cells).find(cell => cell.ref === ref)?.value
  const commit = (doc: XlsxDocument, pivotA1: number) => recomputeDependents(doc, [{ sheetId: 0, rect: { minRow: 0, maxRow: 0, minCol: 0, maxCol: 0 } }],
    { arrayMode: 'dynamic', readCellValue: (name, col, row) => name === 'Pivot' && row === 0 && col === 0 ? pivotA1 : undefined })

  it('a formula reading a spilled cell recomputes with the spill', () => {
    const doc = numericDoc()
    expect([value(doc, 'A1'), value(doc, 'A2'), value(doc, 'C1')]).toEqual([10, 11, 22])
    const result = commit(doc, 20)
    expect([value(doc, 'A1'), value(doc, 'A2'), value(doc, 'C1')]).toEqual([20, 21, 42])
    expect(result.readers).toEqual(expect.arrayContaining(['calc!2:0']))
    // A later commit too (the spill's cells are now the recompute's own).
    commit(doc, 30)
    expect(value(doc, 'C1')).toBe(62)
  })

  it('a formula reading a cell the spill grows into, or leaves, recomputes too', () => {
    const doc = numericDoc([at('D1', 3, 0, { formula: 'A3*2' })])
    doc.sheets[1]!.rows[0]!.cells[0]!.formula = 'SEQUENCE(Pivot!A1/5,1)'   // 2 rows at 10, 4 at 20
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    expect([value(doc, 'A2'), value(doc, 'A3'), value(doc, 'D1')]).toEqual([2, undefined, 0])
    commit(doc, 20)
    expect([value(doc, 'A3'), value(doc, 'A4'), value(doc, 'D1')]).toEqual([3, 4, 6])
    commit(doc, 10)
    expect([value(doc, 'A3'), value(doc, 'D1')]).toEqual([undefined, 0])
  })

  it('a chart over a spilled cell only (not its anchor) is refilled', () => {
    const doc = numericDoc()
    const chartOnA2 = { groups: [{ series: [{ refs: { values: { sheet: 'Calc', r0: 1, c0: 0, r1: 1, c1: 0 } } }] }] }
    doc.sheets[1]!.drawings = [{ content: { kind: 'xlsxChart', chart: chartOnA2 } } as never]
    expect(commit(doc, 20).charts).toEqual([0])
  })
})

describe('a dependent spill and the live pivot cells it would cover', () => {
  const at = (ref: string, col: number, row: number, extra: Partial<XlsxCell> = {}): XlsxCell => ({ ref, col, row, value: null, styleIndex: 0, hasCachedValue: false, ...extra })
  /** Pivot!B3 is a pivot value (1); A6 = SEQUENCE(1,B3,10) spills right; D1 reads B6. `b6` is the saved B6 (blank by default). */
  function blockerDoc(b6?: number): XlsxDocument {
    const pivot: XlsxSheet = { name: 'Pivot', sheetId: '1', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows: [
      { index: 0, cells: [at('D1', 3, 0, { formula: 'B6*2' })] },
      { index: 2, cells: [at('B3', 1, 2, { value: 1, hasCachedValue: true })] },
      { index: 5, cells: [at('A6', 0, 5, { formula: 'SEQUENCE(1,B3,10)', dynamicArray: true }), ...(b6 === undefined ? [] : [at('B6', 1, 5, { value: b6, hasCachedValue: true })])] },
    ] }
    const doc: XlsxDocument = { sheets: [pivot], definedNames: [], tables: [] }
    evaluateWorkbookFormulas(doc, { forceRecalc: true, arrayMode: 'dynamic' })
    return doc
  }
  const cell = (doc: XlsxDocument, ref: string) => doc.sheets[0]!.rows.flatMap(row => row.cells).find(candidate => candidate.ref === ref)
  /** The commit: B3 becomes 2 and B6 is a pivot cell showing `b6` (null: a cell the pivot vacated). */
  const commit = (doc: XlsxDocument, b6: number | null) => recomputeDependents(doc, [{ sheetId: 0, rect: { minRow: 2, maxRow: 5, minCol: 1, maxCol: 1 } }], {
    arrayMode: 'dynamic',
    readCellValue: (name, col, row) => name !== 'Pivot' || col !== 1 ? undefined : row === 2 ? 2 : row === 5 ? b6 : row > 2 && row < 5 ? null : undefined,
  })

  it('a live pivot value blocks the spill: the anchor shows #SPILL! and a reader of the cell sees the pivot\'s value', () => {
    const doc = blockerDoc()
    expect([cell(doc, 'A6')?.value, cell(doc, 'D1')?.value]).toEqual([10, 0])
    commit(doc, 99)
    expect(cell(doc, 'A6')).toMatchObject({ value: '#SPILL!', valueIsError: true })
    expect(cell(doc, 'B6')?.value ?? null).toBeNull()   // the spill claimed nothing; B6 shows the pivot's 99 through the overlay
    expect(cell(doc, 'D1')?.value).toBe(198)
    // Print and the restore still give the saved cells.
    const saved = savedDependentCells(doc)!.get(doc.sheets[0]!)!
    expect(saved.get(5)!.get(0)).toMatchObject({ value: 10 })
    expect(saved.get(0)!.get(3)).toMatchObject({ value: 0 })
    restoreDependentFormulaValues(doc)
    expect([cell(doc, 'A6')?.value, cell(doc, 'A6')?.valueIsError ?? false, cell(doc, 'D1')?.value]).toEqual([10, false, 0])
  })

  it('a cell the pivot vacated does not block the spill, even where the saved sheet still holds the old pivot output', () => {
    const doc = blockerDoc(5)
    expect(cell(doc, 'D1')?.value).toBe(10)
    commit(doc, null)
    expect([cell(doc, 'A6')?.value, cell(doc, 'B6')?.value, cell(doc, 'D1')?.value]).toEqual([10, 11, 22])
    restoreDependentFormulaValues(doc)
    expect([cell(doc, 'A6')?.value, cell(doc, 'B6')?.value, cell(doc, 'D1')?.value]).toEqual([10, 5, 10])
  })
})
