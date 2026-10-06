import type { AstNode, EvaluationContext, FormulaValue } from './types'
import {
  coerceToBoolean,
  coerceToNumber,
  coerceToString,
  finiteOrNum,
  isFormulaError,
  makeCellKey,
  round15,
} from './evaluator'

export { coerceToBoolean, coerceToNumber, coerceToString }

export type EvaluatorFn = (node: AstNode, ctx?: EvaluationContext) => FormulaValue

export type FunctionHandler = (
  args: AstNode[],
  ctx: EvaluationContext | undefined,
  evalNode: EvaluatorFn,
) => FormulaValue

export interface FlatArgItem {
  value: FormulaValue
  fromRef: boolean
}

export function flattenArgs(
  args: AstNode[],
  ctx: EvaluationContext | undefined,
  evalNode: EvaluatorFn,
): FlatArgItem[] {
  const items: FlatArgItem[] = []

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
            items.push({ value: cellVal === undefined ? null : cellVal, fromRef: true })
          }
        }
      } else if (ctx?.getCellValue) {
        for (let r = minRow; r <= maxRow; r++) {
          for (let c = minCol; c <= maxCol; c++) {
            const sheet = arg.ref.sheet ?? ctx.currentSheet
            const key = makeCellKey(sheet, c, r)
            if (ctx.visited?.has(key)) {
              if (ctx) ctx.hasCycle = true
              items.push({ value: 0, fromRef: true })
              continue
            }
            ctx.visited?.add(key)
            try {
              const cellVal = ctx.getCellValue(sheet, c, r)
              items.push({ value: cellVal === undefined ? null : cellVal, fromRef: true })
            } finally {
              ctx.visited?.delete(key)
            }
          }
        }
      }
    } else if (arg.type === 'cell') {
      const val = evalNode(arg, ctx)
      items.push({ value: val, fromRef: true })
    } else {
      const val = evalNode(arg, ctx)
      items.push({ value: val, fromRef: false })
    }
  }

  return items
}

export const FUNCTIONS: Record<string, FunctionHandler> = {
  SUM: (args, ctx, evalNode) => {
    const items = flattenArgs(args, ctx, evalNode)
    let sum = 0
    for (const item of items) {
      const v = item.value
      if (isFormulaError(v)) return v
      if (item.fromRef) {
        if (typeof v === 'number') {
          sum += v
        }
      } else {
        if (v === null || v === undefined) continue
        const n = coerceToNumber(v)
        if (isFormulaError(n)) return n
        sum += n
      }
    }
    return finiteOrNum(sum)
  },

  AVERAGE: (args, ctx, evalNode) => {
    const items = flattenArgs(args, ctx, evalNode)
    let sum = 0
    let count = 0
    for (const item of items) {
      const v = item.value
      if (isFormulaError(v)) return v
      if (item.fromRef) {
        if (typeof v === 'number') {
          sum += v
          count++
        }
      } else {
        if (v === null || v === undefined) continue
        const n = coerceToNumber(v)
        if (isFormulaError(n)) return n
        sum += n
        count++
      }
    }
    if (count === 0) return '#DIV/0!'
    return finiteOrNum(sum / count)
  },

  MIN: (args, ctx, evalNode) => {
    const items = flattenArgs(args, ctx, evalNode)
    let min: number | undefined
    for (const item of items) {
      const v = item.value
      if (isFormulaError(v)) return v
      let num: number | undefined
      if (item.fromRef) {
        if (typeof v === 'number') num = v
      } else {
        if (v === null || v === undefined) continue
        const n = coerceToNumber(v)
        if (isFormulaError(n)) return n
        num = n
      }
      if (num !== undefined) {
        min = min === undefined ? num : Math.min(min, num)
      }
    }
    return min === undefined ? 0 : finiteOrNum(min)
  },

  MAX: (args, ctx, evalNode) => {
    const items = flattenArgs(args, ctx, evalNode)
    let max: number | undefined
    for (const item of items) {
      const v = item.value
      if (isFormulaError(v)) return v
      let num: number | undefined
      if (item.fromRef) {
        if (typeof v === 'number') num = v
      } else {
        if (v === null || v === undefined) continue
        const n = coerceToNumber(v)
        if (isFormulaError(n)) return n
        num = n
      }
      if (num !== undefined) {
        max = max === undefined ? num : Math.max(max, num)
      }
    }
    return max === undefined ? 0 : finiteOrNum(max)
  },

  COUNT: (args, ctx, evalNode) => {
    const items = flattenArgs(args, ctx, evalNode)
    let count = 0
    for (const item of items) {
      const v = item.value
      if (isFormulaError(v)) return v
      if (item.fromRef) {
        if (typeof v === 'number') {
          count++
        }
      } else {
        if (v === null || v === undefined) continue
        if (typeof v === 'number') {
          count++
        } else if (typeof v === 'boolean') {
          count++
        } else if (typeof v === 'string') {
          const n = coerceToNumber(v)
          if (!isFormulaError(n)) {
            count++
          }
        }
      }
    }
    return count
  },

  COUNTA: (args, ctx, evalNode) => {
    const items = flattenArgs(args, ctx, evalNode)
    let count = 0
    for (const item of items) {
      const v = item.value
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
    return finiteOrNum(Math.abs(num))
  },

  ROUND: (args, ctx, evalNode) => {
    if (args.length !== 2) return '#VALUE!'
    const val = evalNode(args[0], ctx)
    if (isFormulaError(val)) return val
    const num = coerceToNumber(val)
    if (isFormulaError(num)) return num

    const digitsVal = evalNode(args[1], ctx)
    if (isFormulaError(digitsVal)) return digitsVal
    const digitsNum = coerceToNumber(digitsVal)
    if (isFormulaError(digitsNum)) return digitsNum

    if (!Number.isFinite(num) || !Number.isFinite(digitsNum)) return '#NUM!'

    const digits = Math.trunc(digitsNum)
    if (digits > 308 || digits < -308) return '#NUM!'

    const factor = Math.pow(10, digits)
    if (!Number.isFinite(factor) || factor === 0) return '#NUM!'

    const sign = num < 0 ? -1 : 1
    const absRounded = Math.round(Math.abs(num) * factor) / factor
    const res = sign * absRounded
    if (!Number.isFinite(res)) return '#NUM!'
    return round15(res)
  },

  INT: (args, ctx, evalNode) => {
    if (args.length !== 1) return '#VALUE!'
    const val = evalNode(args[0], ctx)
    if (isFormulaError(val)) return val
    const num = coerceToNumber(val)
    if (isFormulaError(num)) return num
    return finiteOrNum(Math.floor(round15(num)))
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
    return finiteOrNum(res)
  },

  PRODUCT: (args, ctx, evalNode) => {
    const items = flattenArgs(args, ctx, evalNode)
    let prod: number | undefined
    for (const item of items) {
      const v = item.value
      if (isFormulaError(v)) return v
      let num: number | undefined
      if (item.fromRef) {
        if (typeof v === 'number') num = v
      } else {
        if (v === null || v === undefined) continue
        const n = coerceToNumber(v)
        if (isFormulaError(n)) return n
        num = n
      }
      if (num !== undefined) {
        prod = prod === undefined ? num : prod * num
      }
    }
    return prod === undefined ? 0 : finiteOrNum(prod)
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

  CONCAT: (args, ctx, evalNode) => {
    const items = flattenArgs(args, ctx, evalNode)
    let res = ''
    for (const item of items) {
      const v = item.value
      if (isFormulaError(v)) return v
      res += coerceToString(v)
    }
    return res
  },

  LEFT: (args, ctx, evalNode) => {
    if (args.length < 1 || args.length > 2) return '#VALUE!'
    const val = evalNode(args[0], ctx)
    if (isFormulaError(val)) return val
    const str = coerceToString(val)

    let num = 1
    if (args.length === 2) {
      const numVal = evalNode(args[1], ctx)
      if (isFormulaError(numVal)) return numVal
      const n = coerceToNumber(numVal)
      if (isFormulaError(n)) return n
      num = Math.floor(n)
    }
    if (num < 0) return '#VALUE!'
    return str.slice(0, num)
  },

  RIGHT: (args, ctx, evalNode) => {
    if (args.length < 1 || args.length > 2) return '#VALUE!'
    const val = evalNode(args[0], ctx)
    if (isFormulaError(val)) return val
    const str = coerceToString(val)

    let num = 1
    if (args.length === 2) {
      const numVal = evalNode(args[1], ctx)
      if (isFormulaError(numVal)) return numVal
      const n = coerceToNumber(numVal)
      if (isFormulaError(n)) return n
      num = Math.floor(n)
    }
    if (num < 0) return '#VALUE!'
    if (num === 0) return ''
    return str.slice(-num)
  },

  MID: (args, ctx, evalNode) => {
    if (args.length !== 3) return '#VALUE!'
    const val = evalNode(args[0], ctx)
    if (isFormulaError(val)) return val
    const str = coerceToString(val)

    const startVal = evalNode(args[1], ctx)
    if (isFormulaError(startVal)) return startVal
    const startNum = coerceToNumber(startVal)
    if (isFormulaError(startNum)) return startNum

    const numVal = evalNode(args[2], ctx)
    if (isFormulaError(numVal)) return numVal
    const numChars = coerceToNumber(numVal)
    if (isFormulaError(numChars)) return numChars

    const start = Math.floor(startNum)
    const len = Math.floor(numChars)
    if (start < 1 || len < 0) return '#VALUE!'
    return str.slice(start - 1, start - 1 + len)
  },

  LEN: (args, ctx, evalNode) => {
    if (args.length !== 1) return '#VALUE!'
    const val = evalNode(args[0], ctx)
    if (isFormulaError(val)) return val
    const str = coerceToString(val)
    return str.length
  },

  TRIM: (args, ctx, evalNode) => {
    if (args.length !== 1) return '#VALUE!'
    const val = evalNode(args[0], ctx)
    if (isFormulaError(val)) return val
    const str = coerceToString(val)
    return str.replace(/ +/g, ' ').replace(/^ +| +$/g, '')
  },

  UPPER: (args, ctx, evalNode) => {
    if (args.length !== 1) return '#VALUE!'
    const val = evalNode(args[0], ctx)
    if (isFormulaError(val)) return val
    return coerceToString(val).toUpperCase()
  },

  LOWER: (args, ctx, evalNode) => {
    if (args.length !== 1) return '#VALUE!'
    const val = evalNode(args[0], ctx)
    if (isFormulaError(val)) return val
    return coerceToString(val).toLowerCase()
  },
}
