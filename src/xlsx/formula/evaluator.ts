import type { AstNode, EvaluationContext, FormulaError, FormulaValue, EvaluationError, EvaluationValue } from './types'
import { parseFormula } from './parser'
import { FUNCTIONS } from './functions'

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

export function formulaError(code: FormulaError): EvaluationError {
  return { kind: 'formula-error', code }
}

export function isEvaluationError(value: unknown): value is EvaluationError {
  return typeof value === 'object' && value !== null && 'kind' in value && value.kind === 'formula-error'
}

export function publicFormulaValue(value: EvaluationValue): FormulaValue {
  return isEvaluationError(value) ? value.code : value
}

/** Normalize callback boundaries only; formula strings always remain ordinary text. */
export function contextValue(value: EvaluationValue, ctx?: EvaluationContext): EvaluationValue {
  return !ctx?.typedValues && isFormulaError(value) ? formulaError(value) : value
}

export const MAX_EVAL_DEPTH = 512

export function makeCellKey(sheet: string | undefined, col: number, row: number): string {
  const s = (sheet ?? '').trim().toLowerCase()
  return `${s}!${col}:${row}`
}

/**
 * Rounds a floating-point number to 15 significant digits (Excel standard precision)
 * to eliminate floating-point artifacts like 0.1 + 0.2 = 0.30000000000000004.
 */
export function round15(val: number): number {
  if (!Number.isFinite(val)) return val
  const rounded = parseFloat(val.toPrecision(15))
  return rounded === 0 ? 0 : rounded
}

export function finiteOrNum(res: number): number | EvaluationError {
  if (!Number.isFinite(res) || Number.isNaN(res)) {
    return formulaError('#NUM!')
  }
  return round15(res)
}

export function coerceToNumber(val: EvaluationValue): number | EvaluationError {
  if (isEvaluationError(val)) return val
  if (val === null) return 0
  if (typeof val === 'boolean') return val ? 1 : 0
  if (typeof val === 'number') return val
  if (typeof val === 'string') {
    const trimmed = val.trim()
    if (trimmed === '') return 0
    const num = Number(trimmed)
    if (Number.isNaN(num)) return formulaError('#VALUE!')
    return num
  }
  return formulaError('#VALUE!')
}

export function coerceToString(val: EvaluationValue): string {
  if (val === null) return ''
  if (typeof val === 'boolean') return val ? 'TRUE' : 'FALSE'
  if (typeof val === 'number') return String(round15(val))
  return isEvaluationError(val) ? val.code : String(val)
}

export function coerceToBoolean(val: EvaluationValue): boolean | EvaluationError {
  if (isEvaluationError(val)) return val
  if (val === null) return false
  if (typeof val === 'boolean') return val
  if (typeof val === 'number') return val !== 0
  if (typeof val === 'string') {
    const upper = val.trim().toUpperCase()
    if (upper === 'TRUE') return true
    if (upper === 'FALSE') return false
    return formulaError('#VALUE!')
  }
  return formulaError('#VALUE!')
}

function getTypeRank(val: EvaluationValue): number {
  if (typeof val === 'number') return 1
  if (typeof val === 'string') return 2
  if (typeof val === 'boolean') return 3
  return 0 // null
}

function compareValues(v1: EvaluationValue, v2: EvaluationValue): number {
  if (v1 === null && v2 === null) return 0

  let a = v1
  let b = v2
  if (a === null) {
    if (typeof b === 'number') a = 0
    else if (typeof b === 'string') a = ''
    else if (typeof b === 'boolean') a = 0
  }
  if (b === null) {
    if (typeof a === 'number') b = 0
    else if (typeof a === 'string') b = ''
    else if (typeof a === 'boolean') b = 0
  }

  const rankA = getTypeRank(a)
  const rankB = getTypeRank(b)

  if (rankA !== rankB) {
    return rankA < rankB ? -1 : 1
  }

  if (typeof a === 'number' && typeof b === 'number') {
    const rA = round15(a)
    const rB = round15(b)
    if (rA < rB) return -1
    if (rA > rB) return 1
    return 0
  }

  if (typeof a === 'string' && typeof b === 'string') {
    const sA = a.toUpperCase()
    const sB = b.toUpperCase()
    if (sA < sB) return -1
    if (sA > sB) return 1
    return 0
  }

  if (typeof a === 'boolean' && typeof b === 'boolean') {
    if (a === b) return 0
    return a ? 1 : -1
  }

  return 0
}

export function evaluateNode(node?: AstNode, ctx?: EvaluationContext): EvaluationValue {
  if (node && ctx?.nodeValues?.has(node)) return ctx.nodeValues.get(node)!
  const value = evaluateUncachedNode(node, ctx)
  if (node) ctx?.nodeValues?.set(node, value)
  return value
}

function evaluateUncachedNode(node?: AstNode, ctx?: EvaluationContext): EvaluationValue {
  if (!node) return null
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
        return formulaError(node.error)
      }
      return formulaError('#NAME?')
    }

    case 'cell': {
      if (!ctx || !ctx.getCellValue) {
        return null
      }
      const sheet = node.ref.sheet ?? ctx.currentSheet
      const key = makeCellKey(sheet, node.ref.col, node.ref.row)
      if (!ctx.currentCell && ctx.visited?.has(key)) {
        ctx.hasCycle = true
        return 0
      }

      ctx.visited?.add(key)
      try {
        const val = ctx.getCellValue(sheet, node.ref.col, node.ref.row)
        return val === undefined ? null : contextValue(val, ctx)
      } finally {
        ctx.visited?.delete(key)
      }
    }

    case 'range': {
      return formulaError('#VALUE!')
    }

    case 'unary': {
      const inner = evaluateNode(node.expr, ctx)
      if (isEvaluationError(inner)) return inner

      const num = coerceToNumber(inner)
      if (isEvaluationError(num)) return num

      if (node.op === '+') {
        return finiteOrNum(num)
      }
      if (node.op === '-') {
        return finiteOrNum(-num)
      }
      if (node.op === '%') {
        return finiteOrNum(num / 100)
      }
      return formulaError('#VALUE!')
    }

    case 'binary': {
      const leftVal = evaluateNode(node.left, ctx)
      if (isEvaluationError(leftVal)) return leftVal

      const rightVal = evaluateNode(node.right, ctx)
      if (isEvaluationError(rightVal)) return rightVal

      if (node.op === '&') {
        return coerceToString(leftVal) + coerceToString(rightVal)
      }

      if (
        node.op === '=' ||
        node.op === '<>' ||
        node.op === '<' ||
        node.op === '<=' ||
        node.op === '>' ||
        node.op === '>='
      ) {
        const cmp = compareValues(leftVal, rightVal)
        switch (node.op) {
          case '=':
            return cmp === 0
          case '<>':
            return cmp !== 0
          case '<':
            return cmp < 0
          case '<=':
            return cmp <= 0
          case '>':
            return cmp > 0
          case '>=':
            return cmp >= 0
        }
      }

      // Arithmetic operations
      const n1 = coerceToNumber(leftVal)
      if (isEvaluationError(n1)) return n1
      const n2 = coerceToNumber(rightVal)
      if (isEvaluationError(n2)) return n2

      switch (node.op) {
        case '+':
          return finiteOrNum(n1 + n2)

        case '-':
          return finiteOrNum(n1 - n2)

        case '*':
          return finiteOrNum(n1 * n2)

        case '/': {
          if (n2 === 0) return formulaError('#DIV/0!')
          return finiteOrNum(n1 / n2)
        }

        case '^': {
          if (n1 < 0 && !Number.isInteger(n2)) return formulaError('#NUM!')
          const res = Math.pow(n1, n2)
          return finiteOrNum(res)
        }

        default:
          return formulaError('#VALUE!')
      }
    }

    case 'call': {
      const handler = FUNCTIONS[node.name.toUpperCase()]
      if (!handler) {
        return formulaError('#NAME?')
      }
      return handler(node.args, ctx, evaluateNode)
    }

    default:
      return formulaError('#VALUE!')
  }
}

/**
 * Parses (if string) and evaluates an Excel formula, returning a computed EvaluationValue.
 */
export function evaluateFormulaInternal(input: string | AstNode | undefined | null, ctx?: EvaluationContext): EvaluationValue {
  if (input == null) return null
  const node = typeof input === 'string' ? parseFormula(input) : input
  if (!node) return null
  if (!ctx) {
    return evaluateNode(node)
  }

  if (!ctx.visited) {
    ctx.visited = new Set<string>()
  }

  const currentDepth = ctx.evalDepth ?? 0
  if (currentDepth >= MAX_EVAL_DEPTH) {
    return formulaError('#NUM!')
  }
  ctx.evalDepth = currentDepth + 1

  try {
    const res = evaluateNode(node, ctx)
    if (currentDepth === 0 && ctx.hasCycle) {
      if (!ctx.currentCell) {
        return 0
      }
    }
    return res
  } finally {
    ctx.evalDepth = currentDepth
    if (currentDepth === 0) {
      ctx.hasCycle = false
    }
  }
}

/** Public compatibility boundary: never returns an internal error object. */
export function evaluateFormula(input: string | AstNode | undefined | null, ctx?: EvaluationContext): FormulaValue {
  return publicFormulaValue(evaluateFormulaInternal(input, ctx))
}
