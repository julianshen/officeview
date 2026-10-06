import { evaluateFormula, makeCellKey } from './evaluator'
import { translateSharedFormula } from './shared'
import type { EvaluationContext, FormulaValue } from './types'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../types'

export interface WorkbookEvaluationOptions {
  /** If true, recalculates all formula cells even if a cached <v> is present. */
  fullCalcOnLoad?: boolean
  /** Force recalculation of every formula cell. */
  forceRecalc?: boolean
}

/**
 * Resolves shared formulas within a sheet by translating relative offsets
 * from the master cell (`cell.formula` defined with `si`) to follower cells.
 */
export function resolveSheetSharedFormulas(sheet: XlsxSheet): void {
  const masterBySi = new Map<number, XlsxCell>()

  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      if (cell.sharedFormula?.si !== undefined && cell.formula) {
        masterBySi.set(cell.sharedFormula.si, cell)
      }
    }
  }

  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      if (cell.sharedFormula?.si !== undefined && !cell.formula) {
        const master = masterBySi.get(cell.sharedFormula.si)
        if (master && master.formula) {
          const dCol = cell.col - master.col
          const dRow = cell.row - master.row
          const { formula } = translateSharedFormula(master.formula, dCol, dRow)
          cell.formula = formula
        }
      }
    }
  }
}

/**
 * Evaluates all formulas across an XLSX workbook using a single shared
 * EvaluationContext and per-cell memoization, supporting cross-sheet references,
 * shared formula groups, and circular reference cycle poisoning.
 */
export function evaluateWorkbookFormulas(
  doc: XlsxDocument,
  options?: WorkbookEvaluationOptions
): void {
  const sheetsByName = new Map<string, XlsxSheet>()
  const cellMap = new Map<string, XlsxCell>()

  // 1. Resolve shared formulas on all sheets and index sheets and cells
  for (const sheet of doc.sheets) {
    resolveSheetSharedFormulas(sheet)
    sheetsByName.set(sheet.name.toLowerCase(), sheet)

    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        cellMap.set(makeCellKey(sheet.name, cell.col, cell.row), cell)
      }
    }
  }

  // 2. Shared context and memoization structures
  const memo = new Map<string, FormulaValue>()
  const visited = new Set<string>()

  const ctx: EvaluationContext = {
    visited,
    getCellValue: (sheetName, col, row) => {
      const targetSheetName = sheetName ?? ctx.currentSheet
      if (!targetSheetName) return '#REF!'

      const resolvedSheet = sheetsByName.get(targetSheetName.toLowerCase())
      if (!resolvedSheet) return '#REF!'

      const key = makeCellKey(resolvedSheet.name, col, row)
      if (memo.has(key)) return memo.get(key)!

      const cell = cellMap.get(key)
      if (!cell) {
        memo.set(key, null)
        return null
      }

      if (!cell.formula) {
        memo.set(key, cell.value)
        return cell.value
      }

      const shouldCalc =
        options?.fullCalcOnLoad ||
        options?.forceRecalc ||
        cell.value === null ||
        cell.value === undefined

      if (!shouldCalc) {
        memo.set(key, cell.value)
        return cell.value
      }

      const prevSheet = ctx.currentSheet
      const prevCell = ctx.currentCell
      ctx.currentSheet = resolvedSheet.name
      ctx.currentCell = {
        sheet: resolvedSheet.name,
        col,
        row,
        absCol: false,
        absRow: false,
      }

      const val = evaluateFormula(cell.formula, ctx)

      ctx.currentSheet = prevSheet
      ctx.currentCell = prevCell

      cell.value = val
      memo.set(key, val)
      return val
    },

    getRangeValues: (sheetName, from, to) => {
      const targetSheetName = sheetName ?? ctx.currentSheet
      if (!targetSheetName) return [['#REF!']]

      const resolvedSheet = sheetsByName.get(targetSheetName.toLowerCase())
      if (!resolvedSheet) return [['#REF!']]

      const minCol = Math.min(from.col, to.col)
      const maxCol = Math.max(from.col, to.col)
      const minRow = Math.min(from.row, to.row)
      const maxRow = Math.max(from.row, to.row)

      const matrix: FormulaValue[][] = []
      for (let r = minRow; r <= maxRow; r++) {
        const rowVals: FormulaValue[] = []
        for (let c = minCol; c <= maxCol; c++) {
          rowVals.push(ctx.getCellValue!(resolvedSheet.name, c, r))
        }
        matrix.push(rowVals)
      }
      return matrix
    },
  }

  // 3. Evaluate each formula cell across all sheets
  for (const sheet of doc.sheets) {
    ctx.currentSheet = sheet.name
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        if (cell.formula) {
          ctx.getCellValue!(sheet.name, cell.col, cell.row)
        }
      }
    }
  }
}
