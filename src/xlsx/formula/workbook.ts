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
  const sheetMaxCols = new Map<string, number>()
  const sheetMaxRows = new Map<string, number>()

  // 1. Resolve shared formulas on all sheets and index sheets, cells, and bounds
  for (const sheet of doc.sheets) {
    resolveSheetSharedFormulas(sheet)
    sheetsByName.set(sheet.name.toLowerCase(), sheet)

    let maxCol = -1
    let maxRow = -1
    for (const row of sheet.rows) {
      if (row.index > maxRow) maxRow = row.index
      for (const cell of row.cells) {
        if (cell.col > maxCol) maxCol = cell.col
        cellMap.set(makeCellKey(sheet.name, cell.col, cell.row), cell)
      }
    }
    sheetMaxCols.set(sheet.name.toLowerCase(), Math.max(0, maxCol))
    sheetMaxRows.set(sheet.name.toLowerCase(), Math.max(0, maxRow))
  }

  // 2. Shared context, memoization, and cycle-tracking structures
  const memo = new Map<string, FormulaValue>()
  const visited = new Set<string>()
  const activePath = new Set<string>()
  const cycleMembers = new Set<string>()

  const ctx: EvaluationContext = {
    visited,
    getCellValue: (sheetName, col, row) => {
      const targetSheetName = sheetName ?? ctx.currentSheet
      if (!targetSheetName) return '#REF!'

      const resolvedSheet = sheetsByName.get(targetSheetName.toLowerCase())
      if (!resolvedSheet) return '#REF!'

      const key = makeCellKey(resolvedSheet.name, col, row)

      // Cycle pre-check: if cell is currently on the active evaluation path, cycle detected
      if (activePath.has(key)) {
        ctx.hasCycle = true
        cycleMembers.add(key)
        return 0
      }

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
        cell.ca ||
        cell.value === null ||
        cell.value === undefined

      if (!shouldCalc) {
        memo.set(key, cell.value)
        return cell.value
      }

      activePath.add(key)
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

      let val: FormulaValue
      try {
        val = evaluateFormula(cell.formula, ctx)
      } finally {
        ctx.currentSheet = prevSheet
        ctx.currentCell = prevCell
        activePath.delete(key)
      }

      // If a cycle occurred, poison the result and all cycle members to 0
      if (ctx.hasCycle || cycleMembers.has(key)) {
        cycleMembers.add(key)
        val = 0
      }

      cell.value = val
      memo.set(key, val)

      // Once the root of the active chain unwinds, overwrite any partial memo values for all cycle members
      if (activePath.size === 0 && cycleMembers.size > 0) {
        for (const memberKey of cycleMembers) {
          const mCell = cellMap.get(memberKey)
          if (mCell) mCell.value = 0
          memo.set(memberKey, 0)
        }
        cycleMembers.clear()
        ctx.hasCycle = false
      }

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

      const usedMaxCol = sheetMaxCols.get(resolvedSheet.name.toLowerCase()) ?? 0
      const usedMaxRow = sheetMaxRows.get(resolvedSheet.name.toLowerCase()) ?? 0

      // Clamp range to used sheet bounds (empty/blank cells contribute null/0)
      const clampedMaxCol = Math.min(maxCol, Math.max(minCol, usedMaxCol))
      const clampedMaxRow = Math.min(maxRow, Math.max(minRow, usedMaxRow))

      const totalCells = (clampedMaxCol - minCol + 1) * (clampedMaxRow - minRow + 1)
      if (totalCells > 100000) {
        return [['#NUM!']]
      }

      const matrix: FormulaValue[][] = []
      for (let r = minRow; r <= clampedMaxRow; r++) {
        const rowVals: FormulaValue[] = []
        for (let c = minCol; c <= clampedMaxCol; c++) {
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
