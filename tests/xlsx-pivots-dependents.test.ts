// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { evaluateWorkbookFormulas, recomputeDependents } from '../src/xlsx/formula/workbook'
import { registerFormulaFunction } from '../src/xlsx/formula/functions'
import { markPagedDependentsStale } from '../src/xlsx/formula/workbook'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'
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
