import type { AstNode, EvaluationContext, FormulaError, FormulaValue } from './types'
import { isFormulaError, round15 } from './evaluator'

export type EvaluatorFn = (node: AstNode, ctx?: EvaluationContext) => FormulaValue

export type FunctionHandler = (
  args: AstNode[],
  ctx: EvaluationContext | undefined,
  evalNode: EvaluatorFn,
) => FormulaValue

export function coerceToNumber(val: FormulaValue): number | FormulaError {
  if (isFormulaError(val)) return val
  if (val === null) return 0
  if (typeof val === 'boolean') return val ? 1 : 0
  if (typeof val === 'number') return val
  if (typeof val === 'string') {
    const trimmed = val.trim()
    if (trimmed === '') return 0
    const num = Number(trimmed)
    if (Number.isNaN(num)) return '#VALUE!'
    return num
  }
  return '#VALUE!'
}

export function coerceToBoolean(val: FormulaValue): boolean | FormulaError {
  if (isFormulaError(val)) return val
  if (val === null) return false
  if (typeof val === 'boolean') return val
  if (typeof val === 'number') return val !== 0
  if (typeof val === 'string') {
    const upper = val.trim().toUpperCase()
    if (upper === 'TRUE') return true
    if (upper === 'FALSE') return false
    return '#VALUE!'
  }
  return '#VALUE!'
}

export function flattenArgs(
  args: AstNode[],
  ctx: EvaluationContext | undefined,
  evalNode: EvaluatorFn,
): FormulaValue[] {
  const values: FormulaValue[] = []

  for (const arg of args) {
    if (arg.type === 'range') {
      const minRow = Math.min(arg.ref.from.row, arg.ref.to.row)
      const maxRow = Math.max(arg.ref.from.row, arg.ref.to.row)
      const minCol = Math.min(arg.ref.from.col, arg.ref.to.col)
      const maxCol = Math.max(arg.ref.from.col, arg.ref.to.col)

      if (ctx?.getRangeValues) {
        const grid = ctx.getRangeValues(arg.ref.sheet, arg.ref.from, arg.ref.to)
        for (const row of grid) {
          for (const cellVal of row) {
            values.push(cellVal)
          }
        }
      } else if (ctx?.getCellValue) {
        for (let r = minRow; r <= maxRow; r++) {
          for (let c = minCol; c <= maxCol; c++) {
            values.push(ctx.getCellValue(arg.ref.sheet, c, r))
          }
        }
      }
    } else {
      values.push(evalNode(arg, ctx))
    }
  }

  return values
}

export const FUNCTIONS: Record<string, FunctionHandler> = {
  SUM: (args, ctx, evalNode) => {
    const vals = flattenArgs(args, ctx, evalNode)
    let sum = 0
    for (const v of vals) {
      if (isFormulaError(v)) return v
      if (typeof v === 'number') {
        sum += v
      }
    }
    return round15(sum)
  },

  AVERAGE: (args, ctx, evalNode) => {
    const vals = flattenArgs(args, ctx, evalNode)
    let sum = 0
    let count = 0
    for (const v of vals) {
      if (isFormulaError(v)) return v
      if (typeof v === 'number') {
        sum += v
        count++
      }
    }
    if (count === 0) return '#DIV/0!'
    return round15(sum / count)
  },

  MIN: (args, ctx, evalNode) => {
    const vals = flattenArgs(args, ctx, evalNode)
    let min: number | undefined
    for (const v of vals) {
      if (isFormulaError(v)) return v
      if (typeof v === 'number') {
        min = min === undefined ? v : Math.min(min, v)
      }
    }
    return min === undefined ? 0 : round15(min)
  },

  MAX: (args, ctx, evalNode) => {
    const vals = flattenArgs(args, ctx, evalNode)
    let max: number | undefined
    for (const v of vals) {
      if (isFormulaError(v)) return v
      if (typeof v === 'number') {
        max = max === undefined ? v : Math.max(max, v)
      }
    }
    return max === undefined ? 0 : round15(max)
  },

  COUNT: (args, ctx, evalNode) => {
    const vals = flattenArgs(args, ctx, evalNode)
    let count = 0
    for (const v of vals) {
      if (isFormulaError(v)) return v
      if (typeof v === 'number') {
        count++
      }
    }
    return count
  },

  COUNTA: (args, ctx, evalNode) => {
    const vals = flattenArgs(args, ctx, evalNode)
    let count = 0
    for (const v of vals) {
      if (isFormulaError(v)) {
        count++
        continue
      }
      if (v !== null && v !== '') {
        count++
      }
    }
    return count
  },

  ABS: (args, ctx, evalNode) => {
    if (args.length !== 1) return '#VALUE!'
    const val = evalNode(args[0], ctx)
    if (isFormulaError(val)) return val
    const num = coerceToNumber(val)
    if (isFormulaError(num)) return num
    return round15(Math.abs(num))
  },

  ROUND: (args, ctx, evalNode) => {
    if (args.length !== 2) return '#VALUE!'
    const val = evalNode(args[0], ctx)
    if (isFormulaError(val)) return val
    const num = coerceToNumber(val)
    if (isFormulaError(num)) return num

    const digitsVal = evalNode(args[1], ctx)
    if (isFormulaError(digitsVal)) return digitsVal
    const digits = coerceToNumber(digitsVal)
    if (isFormulaError(digits)) return digits

    const factor = Math.pow(10, digits)
    return round15(Math.round(num * factor) / factor)
  },

  INT: (args, ctx, evalNode) => {
    if (args.length !== 1) return '#VALUE!'
    const val = evalNode(args[0], ctx)
    if (isFormulaError(val)) return val
    const num = coerceToNumber(val)
    if (isFormulaError(num)) return num
    return Math.floor(num)
  },

  MOD: (args, ctx, evalNode) => {
    if (args.length !== 2) return '#VALUE!'
    const val = evalNode(args[0], ctx)
    if (isFormulaError(val)) return val
    const num = coerceToNumber(val)
    if (isFormulaError(num)) return num

    const divVal = evalNode(args[1], ctx)
    if (isFormulaError(divVal)) return divVal
    const div = coerceToNumber(divVal)
    if (isFormulaError(div)) return div

    if (div === 0) return '#DIV/0!'
    const res = num - div * Math.floor(num / div)
    return round15(res)
  },

  PRODUCT: (args, ctx, evalNode) => {
    const vals = flattenArgs(args, ctx, evalNode)
    let prod: number | undefined
    for (const v of vals) {
      if (isFormulaError(v)) return v
      if (typeof v === 'number') {
        prod = prod === undefined ? v : prod * v
      }
    }
    return prod === undefined ? 0 : round15(prod)
  },

  IF: (args, ctx, evalNode) => {
    if (args.length < 2 || args.length > 3) return '#VALUE!'
    const condVal = evalNode(args[0], ctx)
    if (isFormulaError(condVal)) return condVal
    const bool = coerceToBoolean(condVal)
    if (isFormulaError(bool)) return bool

    if (bool) {
      return evalNode(args[1], ctx)
    } else {
      if (args.length === 3) {
        return evalNode(args[2], ctx)
      }
      return false
    }
  },

  IFERROR: (args, ctx, evalNode) => {
    if (args.length !== 2) return '#VALUE!'
    const val = evalNode(args[0], ctx)
    if (isFormulaError(val)) {
      return evalNode(args[1], ctx)
    }
    return val
  },

  AND: (args, ctx, evalNode) => {
    if (args.length === 0) return '#VALUE!'
    for (const arg of args) {
      const v = evalNode(arg, ctx)
      if (isFormulaError(v)) return v
      const b = coerceToBoolean(v)
      if (isFormulaError(b)) return b
      if (!b) return false
    }
    return true
  },

  OR: (args, ctx, evalNode) => {
    if (args.length === 0) return '#VALUE!'
    for (const arg of args) {
      const v = evalNode(arg, ctx)
      if (isFormulaError(v)) return v
      const b = coerceToBoolean(v)
      if (isFormulaError(b)) return b
      if (b) return true
    }
    return false
  },

  NOT: (args, ctx, evalNode) => {
    if (args.length !== 1) return '#VALUE!'
    const v = evalNode(args[0], ctx)
    if (isFormulaError(v)) return v
    const b = coerceToBoolean(v)
    if (isFormulaError(b)) return b
    return !b
  },
}
