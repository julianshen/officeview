/**
 * C2 phase1 array runtime: rectangular literal matrices, explicit scalar/array
 * evaluation modes, the public array API, and the six finite array functions.
 *
 * Matrices are internal only (MatrixValue); XlsxCell.value never receives one.
 * Public evaluateFormula stays primitive top-left; evaluateArrayFormula returns
 * a rectangular MatrixValue or a tagged root error. Mode-sensitive evaluation
 * keeps separate scalar/array stores, so a scalar/top-left memo is never reused
 * for an array argument. Reference arguments read full sparse geometry through
 * the approved reference services; a non-singleton dimension mismatch is an
 * honest capability gate, never an invented #VALUE! or zero/#N/A padding.
 *
 * Documented profiles come from primary Microsoft pages (root FACTS.json):
 * FILTER Boolean-convertible include / empty #CALC! / structured sources;
 * SORT same-shape, sort_index default1, order1/-1, by_col rows/columns, table
 * sources; SORTBY one-row/one-column keys, order1/-1, multiple key/order
 * pairs, table sources; UNIQUE by_col/exactly_once flags, no approximate mode.
 * Tie order and mixed-type sort profiles stay unsettled and are not claimed.
 */
import type { FunctionHandler, EvaluatorFn } from './functions'
import type {
  AstNode,
  ElementValue,
  EvaluationContext,
  EvaluationError,
  EvaluationValue,
  MatrixValue,
  ResolvedRef,
} from './types'
import { parseFormula } from './parser'
import { isReferenceNode } from './refs'
import {
  coerceToBoolean,
  coerceToNumber,
  compareValues,
  evaluateNode,
  formulaError,
  isEvaluationError,
  isMatrixValue,
  markUnsupported,
  materializeResolved,
  matrixOf,
  scalarProjection,
} from './evaluator'

const DEFAULT_MAX_ARRAY_CELLS = 1048576
const DEFAULT_MAX_ARRAY_SIDE = 1048576

function limits(ctx: EvaluationContext | undefined): { maxCells: number; maxSide: number } {
  return {
    maxCells: ctx?.maxArrayCells ?? DEFAULT_MAX_ARRAY_CELLS,
    maxSide: ctx?.maxArraySide ?? DEFAULT_MAX_ARRAY_SIDE,
  }
}

/** Capability gate: a specific diagnostic + internal #NAME? shape (never a
 * fabricated native error code). */
function gate(ctx: EvaluationContext | undefined, feature: string, message: string): EvaluationError {
  markUnsupported(ctx, feature)
  ctx?.reportFormulaIssue?.({ kind: 'unsupported-reference', feature, message })
  return formulaError('#NAME?')
}

function resourceGate(ctx: EvaluationContext | undefined, rows: number, cols: number, maxCells: number, maxSide: number): EvaluationError {
  markUnsupported(ctx, 'array-resource-limit')
  ctx?.reportFormulaIssue?.({
    kind: 'unsupported-reference',
    feature: 'array-resource-limit',
    message: `Array materialization of ${rows}x${cols} exceeds the application policy (maxArrayCells ${maxCells}, maxArraySide ${maxSide}); this is a resource policy, not a native shape error`,
  })
  return formulaError('#NAME?')
}

function materialize(
  rows: number,
  cols: number,
  read: (row: number, col: number) => ElementValue,
  ctx: EvaluationContext | undefined,
): MatrixValue | EvaluationError {
  const { maxCells, maxSide } = limits(ctx)
  if (!Number.isSafeInteger(rows) || !Number.isSafeInteger(cols) || rows < 1 || cols < 1) {
    return gate(ctx, 'array-shape-unavailable', `Array result ${rows}x${cols} is not a nonempty rectangle`)
  }
  if (rows > maxSide || cols > maxSide || rows * cols > maxCells) {
    return resourceGate(ctx, rows, cols, maxCells, maxSide)
  }
  const values: ElementValue[][] = []
  for (let row = 0; row < rows; row++) {
    const out: ElementValue[] = []
    for (let col = 0; col < cols; col++) out.push(read(row, col))
    values.push(out)
  }
  return matrixOf(rows, cols, values)
}

function arrayConstMatrix(node: Extract<AstNode, { type: 'arrayConst' }>): MatrixValue {
  return matrixOf(node.rows.length, node.rows[0].length, node.rows as readonly (readonly ElementValue[])[])
}

function resolvedMatrix(ref: ResolvedRef, ctx: EvaluationContext): MatrixValue | EvaluationError {
  // Shared with the driver path: full prepared rectangle + generation restart.
  const value = materializeResolved(ref, ctx)
  return isEvaluationError(value) ? value : (value as MatrixValue)
}

/** Value/literal argument matrix: an array constant, an already-materialized
 * matrix, or a cooperatively evaluated value (e.g. a name whose binding is an
 * array). */
function valueMatrix(arg: AstNode, ctx: EvaluationContext | undefined, evalNode: EvaluatorFn): MatrixValue | EvaluationError {
  if (arg.type === 'arrayConst') return arrayConstMatrix(arg)
  // Array consumers evaluate their expression argument in ARRAY mode through
  // the approved evaluateInMode hook (same engine, separate stores), never by
  // reusing a scalar/top-left memo.
  const value = ctx?.evaluateInMode ? ctx.evaluateInMode(arg, 'array') : evalNode(arg, ctx)
  if (isEvaluationError(value)) return value
  if (isMatrixValue(value)) return value
  return matrixOf(1, 1, [[value]])
}

/**
 * Full-geometry argument matrix: reference operands read the sparse geometry
 * (or, for a value-denoting name, evaluate through their OWN name node), and
 * every other operand is evaluated through the cooperative driver. A pending
 * child suspends rather than recursing.
 */
function argMatrix(arg: AstNode, ctx: EvaluationContext | undefined, evalNode: EvaluatorFn): MatrixValue | EvaluationError {
  if (isReferenceNode(arg) && ctx?.references) {
    const resolved = ctx.references.resolve(arg, ctx)
    if (resolved === undefined) return valueMatrix(arg, ctx, evalNode)
    if (isEvaluationError(resolved)) return resolved
    ctx.references.prepare(resolved, ctx)
    return resolvedMatrix(resolved, ctx)
  }
  return valueMatrix(arg, ctx, evalNode)
}

function numberArg(
  arg: AstNode | undefined,
  ctx: EvaluationContext | undefined,
  evalNode: EvaluatorFn,
  fallback?: number,
): number | EvaluationError {
  if (!arg || arg.type === 'empty') return fallback === undefined ? formulaError('#VALUE!') : fallback
  const value = evalNode(arg, ctx)
  if (isEvaluationError(value)) return value
  const num = coerceToNumber(scalarProjection(value))
  if (isEvaluationError(num)) return num
  if (!Number.isFinite(num)) return formulaError('#VALUE!')
  return num
}

function integerArg(arg: AstNode | undefined, ctx: EvaluationContext | undefined, evalNode: EvaluatorFn, fallback?: number): number | EvaluationError {
  const num = numberArg(arg, ctx, evalNode, fallback)
  if (isEvaluationError(num)) return num
  return Math.trunc(num)
}

function booleanArg(arg: AstNode | undefined, ctx: EvaluationContext | undefined, evalNode: EvaluatorFn, fallback: boolean): boolean | EvaluationError {
  if (!arg || arg.type === 'empty') return fallback
  const value = evalNode(arg, ctx)
  if (isEvaluationError(value)) return value
  return coerceToBoolean(scalarProjection(value))
}

function elementKey(value: ElementValue): string {
  if (isEvaluationError(value)) return `e:${value.code}`
  if (value === null) return 'z:'
  if (typeof value === 'number') return `n:${String(value)}`
  if (typeof value === 'boolean') return `b:${String(value)}`
  return `s:${value.toUpperCase()}`
}

function columnOf(matrix: MatrixValue, index: number): ElementValue[] {
  return matrix.values.map((row) => row[index])
}

export function sequenceHandler(): FunctionHandler {
  return (args, ctx, evalNode) => {
    if (args.length < 1 || args.length > 4) return formulaError('#VALUE!')
    const rows = integerArg(args[0], ctx, evalNode)
    if (isEvaluationError(rows)) return rows
    const cols = integerArg(args[1], ctx, evalNode, 1)
    if (isEvaluationError(cols)) return cols
    const start = numberArg(args[2], ctx, evalNode, 1)
    if (isEvaluationError(start)) return start
    const step = numberArg(args[3], ctx, evalNode, 1)
    if (isEvaluationError(step)) return step
    if (rows < 1 || cols < 1) return formulaError('#VALUE!')
    return materialize(rows, cols, (row, col) => start + (row * cols + col) * step, ctx)
  }
}

export function transposeHandler(): FunctionHandler {
  return (args, ctx, evalNode) => {
    if (args.length !== 1) return formulaError('#VALUE!')
    const matrix = argMatrix(args[0], ctx, evalNode)
    if (isEvaluationError(matrix)) return matrix
    return materialize(matrix.cols, matrix.rows, (row, col) => matrix.values[col][row], ctx)
  }
}

export function uniqueHandler(): FunctionHandler {
  return (args, ctx, evalNode) => {
    if (args.length < 1 || args.length > 3) return formulaError('#VALUE!')
    const matrix = argMatrix(args[0], ctx, evalNode)
    if (isEvaluationError(matrix)) return matrix
    const byCol = booleanArg(args[1], ctx, evalNode, false)
    if (isEvaluationError(byCol)) return byCol
    const exactlyOnce = booleanArg(args[2], ctx, evalNode, false)
    if (isEvaluationError(exactlyOnce)) return exactlyOnce
    const count = byCol ? matrix.cols : matrix.rows
    const keys: string[] = []
    for (let i = 0; i < count; i++) keys.push((byCol ? columnOf(matrix, i) : matrix.values[i]).map(elementKey).join('|'))
    const occurrences = new Map<string, number>()
    for (const key of keys) occurrences.set(key, (occurrences.get(key) ?? 0) + 1)
    const emitted = new Set<string>()
    const keep: number[] = []
    for (let i = 0; i < count; i++) {
      const key = keys[i]
      if (exactlyOnce) {
        if (occurrences.get(key) === 1) keep.push(i)
      } else if (!emitted.has(key)) {
        emitted.add(key)
        keep.push(i)
      }
    }
    if (keep.length === 0) return formulaError('#CALC!')
    if (byCol) return materialize(matrix.rows, keep.length, (row, col) => matrix.values[row][keep[col]], ctx)
    return materialize(keep.length, matrix.cols, (row, col) => matrix.values[keep[row]][col], ctx)
  }
}

export function sortHandler(): FunctionHandler {
  return (args, ctx, evalNode) => {
    if (args.length < 1 || args.length > 4) return formulaError('#VALUE!')
    const matrix = argMatrix(args[0], ctx, evalNode)
    if (isEvaluationError(matrix)) return matrix
    const byCol = booleanArg(args[3], ctx, evalNode, false)
    if (isEvaluationError(byCol)) return byCol
    const sortIndex = integerArg(args[1], ctx, evalNode, 1)
    if (isEvaluationError(sortIndex)) return sortIndex
    const order = numberArg(args[2], ctx, evalNode, 1)
    if (isEvaluationError(order)) return order
    if (order !== 1 && order !== -1) return formulaError('#VALUE!')
    const length = byCol ? matrix.cols : matrix.rows
    const axis = byCol ? matrix.rows : matrix.cols
    if (sortIndex < 1 || sortIndex > axis) return formulaError('#VALUE!')
    const indices = Array.from({ length }, (_, i) => i)
    indices.sort((a, b) => {
      const keyA = byCol ? matrix.values[sortIndex - 1][a] : matrix.values[a][sortIndex - 1]
      const keyB = byCol ? matrix.values[sortIndex - 1][b] : matrix.values[b][sortIndex - 1]
      return order * compareValues(keyA, keyB)
    })
    if (byCol) return materialize(matrix.rows, length, (row, col) => matrix.values[row][indices[col]], ctx)
    return materialize(length, matrix.cols, (row, col) => matrix.values[indices[row]][col], ctx)
  }
}

export function sortByHandler(): FunctionHandler {
  return (args, ctx, evalNode) => {
    if (args.length < 2) return formulaError('#VALUE!')
    const matrix = argMatrix(args[0], ctx, evalNode)
    if (isEvaluationError(matrix)) return matrix
    const keys: Array<{ values: readonly ElementValue[]; order: number }> = []
    let orientation: 'row' | 'col' | undefined
    for (let i = 1; i < args.length; i += 2) {
      const by = argMatrix(args[i], ctx, evalNode)
      if (isEvaluationError(by)) return by
      const order = numberArg(args[i + 1], ctx, evalNode, 1)
      if (isEvaluationError(order)) return order
      if (order !== 1 && order !== -1) return formulaError('#VALUE!')
      let vector: readonly ElementValue[]
      let current: 'row' | 'col'
      if (by.cols === 1 && by.rows === matrix.rows) {
        vector = by.values.map((row) => row[0])
        current = 'row'
      } else if (by.rows === 1 && by.cols === matrix.cols) {
        vector = [...by.values[0]]
        current = 'col'
      } else {
        return formulaError('#VALUE!')
      }
      if (orientation !== undefined && orientation !== current) return formulaError('#VALUE!')
      orientation = current
      keys.push({ values: vector, order })
    }
    const orient = orientation ?? 'row'
    const length = orient === 'row' ? matrix.rows : matrix.cols
    const indices = Array.from({ length }, (_, i) => i)
    indices.sort((a, b) => {
      for (const key of keys) {
        const cmp = compareValues(key.values[a], key.values[b])
        if (cmp !== 0) return key.order * cmp
      }
      return 0
    })
    if (orient === 'row') return materialize(length, matrix.cols, (row, col) => matrix.values[indices[row]][col], ctx)
    return materialize(matrix.rows, length, (row, col) => matrix.values[row][indices[col]], ctx)
  }
}

export function filterHandler(): FunctionHandler {
  return (args, ctx, evalNode) => {
    if (args.length < 2 || args.length > 3) return formulaError('#VALUE!')
    const matrix = argMatrix(args[0], ctx, evalNode)
    if (isEvaluationError(matrix)) return matrix
    const include = argMatrix(args[1], ctx, evalNode)
    if (isEvaluationError(include)) return include
    let orientation: 'row' | 'col'
    let mask: readonly ElementValue[]
    if (include.cols === 1 && include.rows === matrix.rows) {
      orientation = 'row'
      mask = include.values.map((row) => row[0])
    } else if (include.rows === 1 && include.cols === matrix.cols) {
      orientation = 'col'
      mask = include.values[0]
    } else {
      return formulaError('#VALUE!')
    }
    const keep: number[] = []
    for (let i = 0; i < mask.length; i++) {
      const value = mask[i]
      if (isEvaluationError(value)) return value
      const flag = coerceToBoolean(value)
      if (isEvaluationError(flag)) return flag
      if (flag) keep.push(i)
    }
    if (keep.length === 0) {
      if (args.length === 3 && args[2].type !== 'empty') {
        const fallback = argMatrix(args[2], ctx, evalNode)
        if (isEvaluationError(fallback)) return fallback
        return fallback
      }
      return formulaError('#CALC!')
    }
    if (orientation === 'row') return materialize(keep.length, matrix.cols, (row, col) => matrix.values[keep[row]][col], ctx)
    return materialize(matrix.rows, keep.length, (row, col) => matrix.values[row][keep[col]], ctx)
  }
}

/**
 * Public array API. Evaluates in explicit array mode using a private array
 * store pair (or the caller's supplied modeStores) and restores the caller
 * context on every exit — success, root error, or a Pending suspension thrown
 * by a workbook frame. Intentionally supplied stores are reused, never cleared.
 */
export function evaluateArrayFormula(input: string | AstNode, ctx?: EvaluationContext): MatrixValue | EvaluationError {
  if (!input) return formulaError('#VALUE!')
  const node = typeof input === 'string' ? parseFormula(input) : input
  const saved = ctx
    ? { mode: ctx.mode, nodeValues: ctx.nodeValues, flatArgs: ctx.flatArgs, functionWork: ctx.functionWork }
    : undefined
  let result: EvaluationValue = null
  try {
    if (!ctx) {
      const local: EvaluationContext = {
        mode: 'array',
        nodeValues: new WeakMap(),
        flatArgs: new WeakMap(),
        functionWork: new WeakMap(),
      }
      result = evaluateNode(node, local)
    } else {
      const stores = ctx.modeStores
      ctx.mode = 'array'
      if (stores) {
        ctx.nodeValues = stores.array.nodeValues
        ctx.flatArgs = stores.array.flatArgs
        ctx.functionWork = stores.array.functionWork
      } else {
        // API-created array stores: leave nodeValues unset so evaluateNode
        // creates and OWNS the memo and installs the private generation
        // restart scope. A mid-evaluation footprint change then restarts
        // bounded at the existing budget instead of leaking the private
        // ReferenceGenerationChanged error. Caller memos stay untouched.
        ctx.nodeValues = undefined
        ctx.flatArgs = new WeakMap()
        ctx.functionWork = new WeakMap()
      }
      result = evaluateNode(node, ctx)
    }
  } finally {
    if (ctx && saved) {
      ctx.mode = saved.mode
      ctx.nodeValues = saved.nodeValues
      ctx.flatArgs = saved.flatArgs
      ctx.functionWork = saved.functionWork
    }
  }
  if (isEvaluationError(result)) return result
  if (isMatrixValue(result)) return result
  return matrixOf(1, 1, [[result]])
}
