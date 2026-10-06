import type { AstNode, EvaluationContext, FormulaError, FormulaValue } from './types'
import { parseFormula } from './parser'

export const CANONICAL_ERRORS = new Set<FormulaError>([
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#NULL!',
])

export function isFormulaError(val: unknown): val is FormulaError {
  return typeof val === 'string' && CANONICAL_ERRORS.has(val as FormulaError)
}

/**
 * Rounds a floating-point number to 15 significant digits (Excel standard precision)
 * to eliminate floating-point artifacts like 0.1 + 0.2 = 0.30000000000000004.
 */
export function round15(val: number): number {
  if (!Number.isFinite(val)) return val
  return parseFloat(val.toPrecision(15))
}

function coerceToNumber(val: FormulaValue): number | FormulaError {
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

export function evaluateNode(node: AstNode, ctx?: EvaluationContext): FormulaValue {
  switch (node.type) {
    case 'number':
      return round15(node.value)

    case 'string':
      return node.value

    case 'boolean':
      return node.value

    case 'empty':
      return null

    case 'error': {
      if (isFormulaError(node.error)) {
        return node.error
      }
      return '#NAME?'
    }

    case 'cell': {
      if (!ctx || !ctx.getCellValue) {
        return 0
      }
      return ctx.getCellValue(node.ref.sheet, node.ref.col, node.ref.row)
    }

    case 'range': {
      return '#VALUE!'
    }

    case 'unary': {
      const inner = evaluateNode(node.expr, ctx)
      if (isFormulaError(inner)) return inner

      const num = coerceToNumber(inner)
      if (isFormulaError(num)) return num

      if (node.op === '+') {
        return round15(num)
      }
      if (node.op === '-') {
        return round15(-num)
      }
      if (node.op === '%') {
        return round15(num / 100)
      }
      return '#VALUE!'
    }

    case 'binary': {
      const leftVal = evaluateNode(node.left, ctx)
      if (isFormulaError(leftVal)) return leftVal

      const rightVal = evaluateNode(node.right, ctx)
      if (isFormulaError(rightVal)) return rightVal

      if (node.op === '&') {
        const leftStr = leftVal === null ? '' : String(leftVal)
        const rightStr = rightVal === null ? '' : String(rightVal)
        return leftStr + rightStr
      }

      // Arithmetic operations
      const n1 = coerceToNumber(leftVal)
      if (isFormulaError(n1)) return n1
      const n2 = coerceToNumber(rightVal)
      if (isFormulaError(n2)) return n2

      switch (node.op) {
        case '+':
          return round15(n1 + n2)

        case '-':
          return round15(n1 - n2)

        case '*':
          return round15(n1 * n2)

        case '/': {
          if (n2 === 0) return '#DIV/0!'
          return round15(n1 / n2)
        }

        case '^': {
          if (n1 === 0 && n2 === 0) return '#NUM!'
          if (n1 < 0 && !Number.isInteger(n2)) return '#NUM!'
          const res = Math.pow(n1, n2)
          if (!Number.isFinite(res) || Number.isNaN(res)) return '#NUM!'
          return round15(res)
        }

        default:
          return '#VALUE!'
      }
    }

    case 'call': {
      return '#NAME?'
    }

    default:
      return '#VALUE!'
  }
}

/**
 * Parses (if string) and evaluates an Excel formula, returning a computed FormulaValue.
 */
export function evaluateFormula(input: string | AstNode, ctx?: EvaluationContext): FormulaValue {
  const node = typeof input === 'string' ? parseFormula(input) : input
  return evaluateNode(node, ctx)
}
