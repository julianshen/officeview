import { parseFormula } from './parser'
import type { AstNode, BinaryOp, CellRef } from './types'

/**
 * Converts a 0-based column index to Excel column letters (0 -> 'A', 25 -> 'Z', 26 -> 'AA').
 */
export function colToName(col: number): string {
  let name = ''
  let c = col
  while (c >= 0) {
    name = String.fromCharCode(65 + (c % 26)) + name
    c = Math.floor(c / 26) - 1
  }
  return name
}

/**
 * Formats a CellRef object into standard Excel notation.
 */
export function formatCellRef(ref: CellRef): string {
  const colStr = (ref.absCol ? '$' : '') + colToName(ref.col)
  const rowStr = (ref.absRow ? '$' : '') + (ref.row + 1)
  const sheetStr = ref.sheet
    ? /^[A-Za-z0-9_]+$/.test(ref.sheet)
      ? `${ref.sheet}!`
      : `'${ref.sheet.replace(/'/g, "''")}'!`
    : ''
  return `${sheetStr}${colStr}${rowStr}`
}

const OP_PRECEDENCE: Record<BinaryOp, number> = {
  '^': 5,
  '*': 4,
  '/': 4,
  '+': 3,
  '-': 3,
  '&': 2,
  '=': 1,
  '<>': 1,
  '<': 1,
  '<=': 1,
  '>': 1,
  '>=': 1,
}

function getNodePrecedence(node: AstNode): number {
  if (node.type === 'binary') return OP_PRECEDENCE[node.op] ?? 0
  if (node.type === 'unary') return 6
  return 100 // atoms and function calls have highest precedence
}

/**
 * Formats an AstNode back into standard Excel formula text (without leading '=').
 */
export function formatFormula(node: AstNode): string {
  switch (node.type) {
    case 'number':
      return String(node.value)
    case 'string':
      return `"${node.value.replace(/"/g, '""')}"`
    case 'boolean':
      return node.value ? 'TRUE' : 'FALSE'
    case 'empty':
      return ''
    case 'error':
      return node.error
    case 'cell':
      return formatCellRef(node.ref)
    case 'range': {
      const sheetStr = node.ref.sheet
        ? /^[A-Za-z0-9_]+$/.test(node.ref.sheet)
          ? `${node.ref.sheet}!`
          : `'${node.ref.sheet.replace(/'/g, "''")}'!`
        : ''
      const fromStr =
        (node.ref.from.absCol ? '$' : '') +
        colToName(node.ref.from.col) +
        (node.ref.from.absRow ? '$' : '') +
        (node.ref.from.row + 1)
      const toStr =
        (node.ref.to.absCol ? '$' : '') +
        colToName(node.ref.to.col) +
        (node.ref.to.absRow ? '$' : '') +
        (node.ref.to.row + 1)
      return `${sheetStr}${fromStr}:${toStr}`
    }
    case 'unary': {
      if (node.op === '%') {
        const inner = formatFormula(node.expr)
        return getNodePrecedence(node.expr) < 6 ? `(${inner})%` : `${inner}%`
      }
      const inner = formatFormula(node.expr)
      return getNodePrecedence(node.expr) < 6 ? `${node.op}(${inner})` : `${node.op}${inner}`
    }
    case 'binary': {
      const parentPrec = OP_PRECEDENCE[node.op] ?? 0
      const leftPrec = getNodePrecedence(node.left)
      const rightPrec = getNodePrecedence(node.right)

      const leftStr = leftPrec < parentPrec ? `(${formatFormula(node.left)})` : formatFormula(node.left)
      const rightStr =
        rightPrec < parentPrec || (rightPrec === parentPrec && (node.op === '-' || node.op === '/' || node.op === '^'))
          ? `(${formatFormula(node.right)})`
          : formatFormula(node.right)

      return `${leftStr}${node.op}${rightStr}`
    }
    case 'call': {
      const argsStr = node.args.map(formatFormula).join(',')
      return `${node.name}(${argsStr})`
    }
  }
}

/**
 * Translates all relative CellRef and RangeRef references in an AST by (dCol, dRow).
 * Out-of-bounds coordinates (<0 or >Excel limits) become '#REF!'.
 */
export function translateAst(node: AstNode, dCol: number, dRow: number): AstNode {
  switch (node.type) {
    case 'cell': {
      const col = node.ref.absCol ? node.ref.col : node.ref.col + dCol
      const row = node.ref.absRow ? node.ref.row : node.ref.row + dRow
      if (col < 0 || row < 0 || col > 16383 || row > 1048575) {
        return { type: 'error', error: '#REF!' }
      }
      return {
        type: 'cell',
        ref: {
          ...node.ref,
          col,
          row,
        },
      }
    }
    case 'range': {
      const fromCol = node.ref.from.absCol ? node.ref.from.col : node.ref.from.col + dCol
      const fromRow = node.ref.from.absRow ? node.ref.from.row : node.ref.from.row + dRow
      const toCol = node.ref.to.absCol ? node.ref.to.col : node.ref.to.col + dCol
      const toRow = node.ref.to.absRow ? node.ref.to.row : node.ref.to.row + dRow
      if (
        fromCol < 0 ||
        fromRow < 0 ||
        fromCol > 16383 ||
        fromRow > 1048575 ||
        toCol < 0 ||
        toRow < 0 ||
        toCol > 16383 ||
        toRow > 1048575
      ) {
        return { type: 'error', error: '#REF!' }
      }
      return {
        type: 'range',
        ref: {
          sheet: node.ref.sheet,
          from: { ...node.ref.from, col: fromCol, row: fromRow },
          to: { ...node.ref.to, col: toCol, row: toRow },
        },
      }
    }
    case 'unary':
      return {
        type: 'unary',
        op: node.op,
        expr: translateAst(node.expr, dCol, dRow),
      }
    case 'binary':
      return {
        type: 'binary',
        op: node.op,
        left: translateAst(node.left, dCol, dRow),
        right: translateAst(node.right, dCol, dRow),
      }
    case 'call':
      return {
        type: 'call',
        name: node.name,
        args: node.args.map(arg => translateAst(arg, dCol, dRow)),
      }
    default:
      return node
  }
}

/**
 * Translates a shared master formula for a dependent cell located at (dCol, dRow) offset.
 */
export function translateSharedFormula(
  masterFormula: string | AstNode,
  dCol: number,
  dRow: number
): { ast: AstNode; formula: string } {
  const ast = typeof masterFormula === 'string' ? parseFormula(masterFormula) : masterFormula
  const translatedAst = translateAst(ast, dCol, dRow)
  const formula = formatFormula(translatedAst)
  return { ast: translatedAst, formula }
}
