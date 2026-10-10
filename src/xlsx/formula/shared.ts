import { parseFormula } from './parser'
import { decodeTableSyntax } from './tables'
import type { AstNode, BinaryOp, CellRef, GridReference, ReferenceNode, WholeColSyntax, WholeRowSyntax } from './types'

function isReferenceLike(n: AstNode): n is ReferenceNode {
  return (
    n.type === 'cell' ||
    n.type === 'range' ||
    n.type === 'wholeCol' ||
    n.type === 'wholeRow' ||
    n.type === 'name' ||
    n.type === 'table' ||
    n.type === 'ref3d' ||
    n.type === 'union' ||
    n.type === 'intersect' ||
    n.type === 'spill'
  )
}

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
/** Sheet-name quoting per ECMA-376 §18.2.19: unquoted references are only
 * valid for identifier-safe names; digit-leading names REQUIRE quotes
 * ('2024'!A1, never 2024!A1 — the unquoted form lexes as a number). */
export function sheetNameSafe(sheet: string): boolean {
  return /^[A-Za-z_\p{L}][A-Za-z0-9_\p{L}]*$/u.test(sheet)
}

function quotedSheet(sheet: string): string {
  return sheetNameSafe(sheet) ? `${sheet}!` : `'${sheet.replace(/'/g, "''")}'!`
}

export function formatCellRef(ref: CellRef): string {
  const colStr = (ref.absCol ? '$' : '') + colToName(ref.col)
  const rowStr = (ref.absRow ? '$' : '') + (ref.row + 1)
  const sheetStr = ref.sheet ? quotedSheet(ref.sheet) : ''
  return `${sheetStr}${colStr}${rowStr}`
}

function formatAxisEndpoint(index: number, absolute: boolean, isCol: boolean): string {
  const base = isCol ? colToName(index) : String(index + 1)
  return (absolute ? '$' : '') + base
}

function formatWholeCol(ref: WholeColSyntax): string {
  const sheetStr = ref.sheet ? quotedSheet(ref.sheet) : ''
  return `${sheetStr}${formatAxisEndpoint(ref.from.index, ref.from.absolute, true)}:${formatAxisEndpoint(ref.to.index, ref.to.absolute, true)}`
}

function formatWholeRow(ref: WholeRowSyntax): string {
  const sheetStr = ref.sheet ? quotedSheet(ref.sheet) : ''
  return `${sheetStr}${formatAxisEndpoint(ref.from.index, ref.from.absolute, false)}:${formatAxisEndpoint(ref.to.index, ref.to.absolute, false)}`
}

function formatGridReference(target: GridReference): string {
  switch (target.type) {
    case 'cell':
      return formatCellRef(target.ref)
    case 'range':
      return formatFormula({ type: 'range', ref: target.ref })
    case 'wholeCol':
      return formatWholeCol(target.ref)
    case 'wholeRow':
      return formatWholeRow(target.ref)
  }
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
      const sheetStr = node.ref.sheet ? quotedSheet(node.ref.sheet) : ''
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
    case 'wholeCol':
      return formatWholeCol(node.ref)
    case 'wholeRow':
      return formatWholeRow(node.ref)
    case 'name':
      return node.ref.sheet ? `${quotedSheet(node.ref.sheet)}${node.ref.name}` : node.ref.name
    case 'table':
      return node.ref.raw
    case 'ref3d': {
      // Group 8: a 3D run formats as ONE quoted unit covering both sheet
      // endpoints (quoted-run spelling preserved: spaces and escaped
      // apostrophes stay joined inside the run); safe identifier names stay
      // unquoted.
      const target = formatGridReference(node.target)
      const from = node.sheets.fromSheet
      const to = node.sheets.toSheet
      if (sheetNameSafe(from) && sheetNameSafe(to) && from.indexOf(':') < 0) {
        return `${from}:${to}!${target}`
      }
      // Joined quoted run: escape the DECODED interior apostrophes once via
      // the existing quotedSheet rule across the whole joined `${from}:${to}`
      // unit (from/to endpoint spellings are not re-quoted separately).
      const joined = `${from}:${to}`
      return `'${joined.replace(/'/g, "''")}'!${target}`
    }
    case 'union': {
      // A discrete structured selector has multiple table operands sharing
      // one raw expression. Preserve it only when decoding produces these
      // exact operands; ordinary duplicate table unions must remain duplicates.
      const first = node.refs[0]
      if (first?.type === 'table' && node.refs.every(ref => ref.type === 'table' && ref.ref.raw === first.ref.raw)) {
        const decoded = decodeTableSyntax(first.ref.table, first.ref.raw)
        if (decoded.kind === 'union' && decoded.items.length === node.refs.length && decoded.items.every((item, i) => {
          const ref = node.refs[i]
          return ref.type === 'table' && ref.ref.table === item.table && JSON.stringify(ref.ref.items) === JSON.stringify(item.items) && JSON.stringify(ref.ref.columns) === JSON.stringify(item.columns)
        })) return first.ref.raw
      }
      return `(${node.refs.map((r) => formatFormula(r)).join(',')})`
    }
    case 'intersect':
      return `${formatFormula(node.left)} ${formatFormula(node.right)}`
    case 'spill':
      return `${formatCellRef(node.anchor)}#`
    case 'implicitIntersect': {
      const operand = formatFormula(node.expr)
      // @ consumes one primary. Group compound expressions and sheet-run
      // syntax so generated shared text retains the exact operand boundary.
      const group = node.expr.type === 'intersect' || node.expr.type === 'binary' || node.expr.type === 'unary' || node.expr.type === 'ref3d' || node.expr.type === 'implicitIntersect'
      return group ? `@(${operand})` : `@${operand}`
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
        rightPrec < parentPrec || (rightPrec === parentPrec)
          ? `(${formatFormula(node.right)})`
          : formatFormula(node.right)

      return `${leftStr}${node.op}${rightStr}`
    }
    case 'call': {
      const argsStr = node.args.map(formatFormula).join(',')
      return `${node.name}(${argsStr})`
    }
    case 'arrayConst': {
      const rows = node.rows.map((row) => row.map((element) => {
        if (typeof element === 'number') return String(element)
        if (typeof element === 'string') return `"${element.replace(/"/g, '""')}"`
        if (typeof element === 'boolean') return element ? 'TRUE' : 'FALSE'
        return element.code
      }).join(','))
      return `{${rows.join(';')}}`
    }
  }
}

function translateAxisEndpoint(
  ep: { index: number; absolute: boolean },
  delta: number,
  limit: number,
): { index: number; absolute: boolean } | null {
  const index = ep.absolute ? ep.index : ep.index + delta
  if (index < 0 || index > limit) return null
  return { index, absolute: ep.absolute }
}

function translateWholeCol(ref: WholeColSyntax, dCol: number): WholeColSyntax | null {
  const from = translateAxisEndpoint(ref.from, dCol, 16383)
  const to = translateAxisEndpoint(ref.to, dCol, 16383)
  if (!from || !to) return null
  return { ...ref, from, to }
}

function translateWholeRow(ref: WholeRowSyntax, dRow: number): WholeRowSyntax | null {
  const from = translateAxisEndpoint(ref.from, dRow, 1048575)
  const to = translateAxisEndpoint(ref.to, dRow, 1048575)
  if (!from || !to) return null
  return { ...ref, from, to }
}

function translateGrid(target: GridReference, dCol: number, dRow: number): GridReference | null {
  switch (target.type) {
    case 'cell': {
      const moved = translateAst({ type: 'cell', ref: target.ref }, dCol, dRow)
      return moved.type === 'cell' ? moved : null
    }
    case 'range': {
      const moved = translateAst({ type: 'range', ref: target.ref }, dCol, dRow)
      return moved.type === 'range' ? moved : null
    }
    case 'wholeCol': {
      const moved = translateWholeCol(target.ref, dCol)
      return moved ? { type: 'wholeCol', ref: moved } : null
    }
    case 'wholeRow': {
      const moved = translateWholeRow(target.ref, dRow)
      return moved ? { type: 'wholeRow', ref: moved } : null
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
    case 'wholeCol': {
      const moved = translateWholeCol(node.ref, dCol)
      return moved ? { type: 'wholeCol', ref: moved } : { type: 'error', error: '#REF!' }
    }
    case 'wholeRow': {
      const moved = translateWholeRow(node.ref, dRow)
      return moved ? { type: 'wholeRow', ref: moved } : { type: 'error', error: '#REF!' }
    }
    // Shared copies preserve identifiers; binding occurs at evaluation.
    // Table identifiers remain source expressions (never rewritten).
    case 'name':
    case 'table':
      return node
    case 'ref3d': {
      const moved = translateGrid(node.target, dCol, dRow)
      return moved ? { ...node, target: moved } : { type: 'error', error: '#REF!' }
    }
    case 'union': {
      const refs: ReferenceNode[] = []
      for (const r of node.refs) {
        const moved = translateAst(r, dCol, dRow)
        if (moved.type === 'error') return moved
        if (!isReferenceLike(moved)) return { type: 'error', error: '#REF!' }
        refs.push(moved)
      }
      return { type: 'union', refs }
    }
    case 'intersect': {
      const left = translateAst(node.left, dCol, dRow)
      const right = translateAst(node.right, dCol, dRow)
      if (left.type === 'error') return left
      if (right.type === 'error') return right
      if (!isReferenceLike(left) || !isReferenceLike(right)) {
        return { type: 'error', error: '#REF!' }
      }
      return { type: 'intersect', left, right }
    }
    case 'spill': {
      const moved = translateAst({ type: 'cell', ref: node.anchor }, dCol, dRow)
      return moved.type === 'cell' ? { type: 'spill', anchor: moved.ref } : moved
    }
    case 'implicitIntersect':
      return { type: 'implicitIntersect', expr: translateAst(node.expr, dCol, dRow) }
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
    // Literal matrices carry no references; they translate verbatim.
    case 'arrayConst':
      return node
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
