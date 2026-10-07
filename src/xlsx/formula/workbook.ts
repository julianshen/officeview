import { evaluateFormulaInternal, makeCellKey, formulaError, isFormulaError, isEvaluationError, publicFormulaValue } from './evaluator'
import { parseFormula } from './parser'
import { translateSharedFormula } from './shared'
import type { AstNode, EvaluationContext, EvaluationValue } from './types'
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
  // Cheap early-exit: skip evaluation work entirely when no cell has a formula.
  let hasFormulas = false
  for (const sheet of doc.sheets) {
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        if (cell.formula !== undefined || cell.sharedFormula !== undefined) {
          hasFormulas = true
          break
        }
      }
      if (hasFormulas) break
    }
    if (hasFormulas) break
  }
  if (!hasFormulas) return

  const sheetsByName = new Map<string, XlsxSheet>()
  const cellMap = new Map<string, XlsxCell>()
  const cellSheets = new Map<string, string>()
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
        if (cell.row > maxRow) maxRow = cell.row
        const key = makeCellKey(sheet.name, cell.col, cell.row)
        cellMap.set(key, cell)
        cellSheets.set(key, sheet.name)
      }
    }
    sheetMaxCols.set(sheet.name.toLowerCase(), Math.max(0, maxCol))
    sheetMaxRows.set(sheet.name.toLowerCase(), Math.max(0, maxRow))
  }

  // Evaluation frames suspend only when a dynamically requested dependency is
  // unresolved. No recursive cell calls or static traversal of lazy branches.
  const memo = new Map<string, EvaluationValue>()
  const activeStack: string[] = []
  const activeSet = new Set<string>()
  const cycleMembers = new Set<string>()
  const formulas = new Map<string, AstNode>()
  for (const [key, cell] of cellMap) {
    const shouldCalc = options?.fullCalcOnLoad || options?.forceRecalc || cell.ca ||
      (cell.hasCachedValue === undefined ? cell.value == null : !cell.hasCachedValue)
    if (cell.formula && shouldCalc) formulas.set(key, parseFormula(cell.formula))
  }

  class PendingDependency {
    constructor(readonly key: string) {}
  }

  interface RangeProgress { matrix: EvaluationValue[][]; next: number }
  interface Frame {
    key: string
    nodeValues: WeakMap<AstNode, EvaluationValue>
    flatArgs: NonNullable<EvaluationContext['flatArgs']>
    ranges: Map<string, RangeProgress>
  }
  let currentFrame: Frame | undefined

  const ctx: EvaluationContext = {
    typedValues: true,
    getCellValue: (sheetName, col, row) => {
      const targetSheetName = sheetName ?? ctx.currentSheet
      if (!targetSheetName) return formulaError('#REF!')
      const resolvedSheet = sheetsByName.get(targetSheetName.toLowerCase())
      if (!resolvedSheet) return formulaError('#REF!')
      const key = makeCellKey(resolvedSheet.name, col, row)
      if (activeSet.has(key)) {
        const index = activeStack.indexOf(key)
        for (let i = index; i < activeStack.length; i++) cycleMembers.add(activeStack[i])
        return 0
      }
      if (memo.has(key)) return memo.get(key)!
      const cell = cellMap.get(key)
      if (!cell) {
        memo.set(key, null)
        return null
      }
      if (formulas.has(key)) throw new PendingDependency(key)
      // Parsed cells explicitly distinguish errors and strings. Legacy hand-built
      // models retain their previous canonical-error-string convention for all cells.
      const error = cell.valueIsError ?? isFormulaError(cell.value)
      const value = error && isFormulaError(cell.value) ? formulaError(cell.value) : cell.value
      memo.set(key, value)
      return value
    },

    getRangeValues: (sheetName, from, to) => {
      const targetSheetName = sheetName ?? ctx.currentSheet
      if (!targetSheetName) return [[formulaError('#REF!')]]

      const resolvedSheet = sheetsByName.get(targetSheetName.toLowerCase())
      if (!resolvedSheet) return [[formulaError('#REF!')]]

      const minCol = Math.min(from.col, to.col)
      const maxCol = Math.max(from.col, to.col)
      const minRow = Math.min(from.row, to.row)
      const maxRow = Math.max(from.row, to.row)

      const usedMaxCol = sheetMaxCols.get(resolvedSheet.name.toLowerCase()) ?? 0
      const usedMaxRow = sheetMaxRows.get(resolvedSheet.name.toLowerCase()) ?? 0

      // Only clamp large/unbounded ranges (> 10k cells) to used bounds
      let clampedMaxCol = maxCol
      let clampedMaxRow = maxRow
      if ((maxCol - minCol + 1) * (maxRow - minRow + 1) > 10000) {
        clampedMaxCol = Math.min(maxCol, Math.max(minCol, usedMaxCol))
        clampedMaxRow = Math.min(maxRow, Math.max(minRow, usedMaxRow))
      }

      const totalCells = (clampedMaxCol - minCol + 1) * (clampedMaxRow - minRow + 1)
      if (totalCells > 100000) {
        return [[formulaError('#NUM!')]]
      }

      // Retain only completed range cells when a dependency suspends this frame.
      // Resuming SUM over N uncached cells needs O(N) reads, not O(N squared).
      const rangeKey = `${resolvedSheet.name}!${minCol}:${minRow}:${clampedMaxCol}:${clampedMaxRow}`
      let progress = currentFrame?.ranges.get(rangeKey)
      if (!progress) {
        progress = { matrix: [], next: 0 }
        currentFrame?.ranges.set(rangeKey, progress)
      }
      const width = clampedMaxCol - minCol + 1
      while (progress.next < totalCells) {
        const rowIndex = Math.floor(progress.next / width)
        const colIndex = progress.next % width
        const value = ctx.getCellValue!(resolvedSheet.name, minCol + colIndex, minRow + rowIndex)
        const row = progress.matrix[rowIndex] ?? (progress.matrix[rowIndex] = [])
        row.push(value)
        progress.next++
      }
      return progress.matrix
    },
  }

  function evaluateCell(rootKey: string): void {
    if (memo.has(rootKey) || !formulas.has(rootKey)) return
    const frames: Frame[] = []
    const push = (key: string) => {
      frames.push({ key, nodeValues: new WeakMap(), flatArgs: new WeakMap(), ranges: new Map() })
      activeStack.push(key)
      activeSet.add(key)
    }
    push(rootKey)
    while (frames.length > 0) {
      const frame = frames[frames.length - 1]
      currentFrame = frame
      const cell = cellMap.get(frame.key)!
      ctx.currentSheet = cellSheets.get(frame.key)!
      ctx.currentCell = { col: cell.col, row: cell.row, absCol: false, absRow: false }
      ctx.nodeValues = frame.nodeValues
      ctx.flatArgs = frame.flatArgs
      let value: EvaluationValue
      try {
        value = evaluateFormulaInternal(formulas.get(frame.key)!, ctx)
      } catch (error) {
        if (!(error instanceof PendingDependency)) throw error
        push(error.key)
        continue
      }
      if (cycleMembers.has(frame.key)) value = 0
      cell.value = publicFormulaValue(value)
      cell.valueIsError = isEvaluationError(value)
      cell.hasCachedValue = true
      memo.set(frame.key, value)
      frames.pop()
      activeStack.pop()
      activeSet.delete(frame.key)
    }
    cycleMembers.clear()
    currentFrame = undefined
    ctx.nodeValues = undefined
    ctx.flatArgs = undefined
    ctx.currentCell = undefined
  }

  for (const key of formulas.keys()) evaluateCell(key)
}
