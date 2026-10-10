import { LOOKUP_FUNCTIONS } from './functions/lookup'
import { CONDITIONAL_FUNCTIONS } from './functions/cond'
import { DATETIME_FUNCTIONS } from './functions/datetime'
import { TEXT_FUNCTIONS } from './functions/text'
import { filterHandler, sequenceHandler, sortByHandler, sortHandler, transposeHandler, uniqueHandler } from './arrays'
import { currentReferenceGeneration, isReferenceGenerationChanged, referenceGenerationChanged, MAX_REFERENCE_GENERATION_RESTARTS, ownedMemoEpoch, requestOwnedMemoRestart } from './reference-generation'
import type {AstNode, ElementValue, EvaluationContext, EvaluationError, EvaluationValue, MatrixValue, RefArea, RefCursor, ReferenceNode, ResolvedRef } from './types'
import {
  areaCells,
  gridAreas,
  intersectAreas,
  isReferenceNode,
  syntacticAreas,
} from './refs'
import {
  coerceToBoolean,
  coerceToNumber,
  coerceToString,
  decimalRound,
  finiteOrNum,
  isEvaluationError,
  isMatrixValue,
  matrixOf,
  formulaError,
  contextValue,
  makeCellKey,
  textChars,
  markUnsupported,
} from './evaluator'

export { coerceToBoolean, coerceToNumber, coerceToString }

export type EvaluatorFn = (node: AstNode, ctx?: EvaluationContext) => EvaluationValue

export type FunctionHandler = (
  args: AstNode[],
  ctx: EvaluationContext | undefined,
  evalNode: EvaluatorFn,
) => EvaluationValue

export interface FlatArgItem {
  value: EvaluationValue
  fromRef: boolean
  /** Explicit omitted argument (e.g. the empty slot in MIN(5,)): participates
   * as 0 for SUM/MIN/MAX/AVERAGE/COUNT, counts for COUNTA, ignored by PRODUCT. */
  omitted?: boolean
}

interface RefDrainProgress {
  items: FlatArgItem[]
  refCursor?: { ref: ResolvedRef; cursor: RefCursor; drainStartLen: number } | undefined
}

/** Read one legacy cell with cycle tracking, mirroring the range branch. */
function readLegacyCell(
  sheet: string | undefined,
  col: number,
  row: number,
  ctx: EvaluationContext | undefined,
): FlatArgItem {
  const key = makeCellKey(sheet, col, row)
  if (ctx?.visited?.has(key)) {
    if (ctx) ctx.hasCycle = true
    return { value: 0, fromRef: true }
  }
  ctx?.visited?.add(key)
  try {
    const cellVal = ctx?.getCellValue?.(sheet, col, row)
    return { value: cellVal === undefined ? null : contextValue(cellVal, ctx), fromRef: true }
  } finally {
    ctx?.visited?.delete(key)
  }
}

function expandLegacyGrid(
  area: RefArea,
  sheet: string | undefined,
  ctx: EvaluationContext | undefined,
): FlatArgItem[] {
  const items: FlatArgItem[] = []
  for (let r = area.firstRow; r < area.firstRow + area.rows; r++) {
    for (let c = area.firstCol; c < area.firstCol + area.cols; c++) {
      items.push(readLegacyCell(sheet, c, r, ctx))
    }
  }
  return items
}

/**
 * Standalone legacy expansion of unions/intersections/whole axes without
 * reference services: syntactic geometry read through the legacy cell
 * callback. Registry operands stay unavailable errors; empty intersections
 * are typed #NULL! (never silent zeros).
 */
function expandLegacyReference(
  arg: ReferenceNode,
  ctx: EvaluationContext | undefined,
  evalNode: EvaluatorFn,
): FlatArgItem[] {
  void evalNode
  if (arg.type !== 'union' && arg.type !== 'intersect' && arg.type !== 'wholeCol' && arg.type !== 'wholeRow') {
    return [{ value: formulaError('#VALUE!'), fromRef: true }]
  }
  const areas = syntacticAreas(arg, ctx?.currentSheet)
  if (isEvaluationError(areas)) return [{ value: areas, fromRef: true }]
  if (arg.type === 'intersect' && areas.length === 0) {
    return [{ value: formulaError('#NULL!'), fromRef: true }]
  }
  const items: FlatArgItem[] = []
  for (const area of areas) {
    // Group 9 (exact scope): FULL physical axis geometry (a whole column or
    // whole row — rows 1048576 or cols 16384) cannot be enumerated by
    // standalone legacy expansion; the specific unavailable-capability
    // diagnostic replaces the scan. Finite (even >100000) ranges keep the
    // reviewed legacy policy unchanged.
    if (area.rows === 1048576 || area.cols === 16384) {
      markUnsupported(ctx, 'sparse-population-unavailable')
      ctx?.reportFormulaIssue?.({
        kind: 'unsupported-reference',
        feature: 'sparse-population-unavailable',
        message: `Standalone full-axis area of ${areaCells(area)} cells needs sparse population capability that is unavailable here`,
      })
      items.push({ value: formulaError('#NAME?'), fromRef: true })
      continue
    }
    // Standalone areas carry the qualifier (or empty) as sheet identity;
    // empty restores an undefined sheet for the legacy cell callback.
    items.push(...expandLegacyGrid(area, area.sheetId === '' ? undefined : area.sheetId, ctx))
  }
  return items
}

/**
 * Drain one reference argument through sparse services into progress items.
 * Each peeked entry commits before advancing, so a suspension (pending
 * dependency) resumes on the same entry: no skips, duplicates or double
 * counts. The live cursor persists on the frame-owned progress across the
 * suspension; completion clears it and the caller advances next.
 */
function drainReferenceArg(
  arg: ReferenceNode,
  ctx: EvaluationContext,
  evalNode: EvaluatorFn,
  progress: RefDrainProgress,
): void {
  const services = ctx.references!
  // Generation invalidation propagates to the whole argument owner. A
  // sibling's already committed items cannot survive a changed footprint.
  let state = progress.refCursor
  for (;;) {
    if (!state) {
      const resolved = services.resolve(arg, ctx)
      if (resolved === undefined) {
        // A name denoting a value evaluates as a direct (non-ref) argument.
        if (arg.type !== 'name') {
          progress.items.push({ value: formulaError('#VALUE!'), fromRef: true })
          return
        }
        const binding = services.bindName(arg.ref, ctx)
        if (isEvaluationError(binding)) {
          progress.items.push({ value: binding, fromRef: true })
          return
        }
        // Group 3 (owner arg): the request rides the OWN name node so the
        // driver's name machinery owns traversal identity; binding.ast bypasses
        // name ownership and can loop the same call AST in a hidden cycle.
        progress.items.push({ value: evalNode(arg, ctx), fromRef: false })
        return
      }
      if (isEvaluationError(resolved)) {
        // Reference-shape failures (#REF!/#NULL!/#NAME?) propagate as errors.
        progress.items.push({ value: resolved, fromRef: true })
        return
      }
      // Group 10 fix: call the approved optional preparation hook before
      // opening/enumerating/counting/reading the reference. Pending exceptions
      // (suspensions) propagate uncaught to the cooperative scheduler.
      services.prepare(resolved, ctx)
      state = {
        ref: resolved,
        cursor: resolved.openCursor('populated'),
        drainStartLen: progress.items.length, // scoped drop point for this drain
      }
      progress.refCursor = state
    }
    const entry = state.ref.peek(state.cursor)
    if (entry === undefined) break
    progress.items.push({ value: entry.value, fromRef: true })
    state.ref.advance(state.cursor)
  }
  progress.refCursor = undefined
}

// Generation/retry state is private and owned by one frame's argument progress;
// no EvaluationContext or ReferenceServices fields and no caller memo resets.
const argumentGenerations = new WeakMap<object, { generation: number | undefined; restarts: number; owned: ReturnType<typeof ownedMemoEpoch> }>()

export function flattenArgs(
  args: AstNode[],
  ctx: EvaluationContext | undefined,
  evalNode: EvaluatorFn,
): FlatArgItem[] {
  let progress = ctx?.flatArgs?.get(args)
  if (!progress) {
    progress = { next: 0, items: [] }
    ctx?.flatArgs?.set(args, progress)
  }

  const owner = progress
  const epoch = argumentGenerations.get(owner) ?? {
    generation: ctx?.references ? currentReferenceGeneration(ctx.references) : undefined,
    restarts: 0,
    owned: ownedMemoEpoch(ctx),
  }
  argumentGenerations.set(owner, epoch)
  const resetProgress = (): void => {
    owner.next = 0
    owner.items.length = 0
    owner.refCursor = undefined
    epoch.generation = ctx?.references ? currentReferenceGeneration(ctx.references) : undefined
    epoch.owned = ownedMemoEpoch(ctx)
  }
  for (;;) {
    try {
      const currentOwned = ownedMemoEpoch(ctx)
      if (currentOwned && (epoch.owned?.owner !== currentOwned.owner || epoch.owned.epoch !== currentOwned.epoch)) {
        if (currentOwned.epoch > 0) {
          // An actual internal generation restart invalidates completed nested
          // progress, including wrappers without a private generation getter.
          resetProgress()
        } else {
          // Ordinary Pending/user unwind creates a fresh local memo but keeps
          // same-generation argument/cursor progress under its original owner.
          epoch.owned = currentOwned
        }
      }
      const current = ctx?.references ? currentReferenceGeneration(ctx.references) : undefined
      if (epoch.generation !== undefined && current !== undefined && epoch.generation !== current) {
        throw referenceGenerationChanged(epoch.generation, current)
      }
      // A later dependency can suspend the function. Keep completed arguments so
      // resumed calls do not re-scan ranges or scalar argument prefixes.
      while (progress.next < args.length) {
        const arg = args[progress.next]
        // Reference arguments resolve through sparse services when present:
        // full geometry, populated-only traversal, suspension-safe commit.
        // Single cells keep the legacy direct path; everything else reference-
        // shaped drains through a frame-persisted cursor.
        if (ctx?.references && isReferenceNode(arg) && arg.type !== 'cell') {
          drainReferenceArg(arg, ctx, evalNode, progress)
          progress.next++
          continue
        }
        const items: FlatArgItem[] = []
        if (arg.type === 'range') {
          const minRow = Math.min(arg.ref.from.row, arg.ref.to.row)
          const maxRow = Math.max(arg.ref.from.row, arg.ref.to.row)
          const minCol = Math.min(arg.ref.from.col, arg.ref.to.col)
          const maxCol = Math.max(arg.ref.from.col, arg.ref.to.col)

          const cellCount = (maxRow - minRow + 1) * (maxCol - minCol + 1)
          // Group 9 (exact scope): ranges spanning EXACTLY one full physical
          // axis (rows === 1048576 or cols === 16384) are whole-axis population
          // shapes — standalone legacy expansion cannot enumerate them: emit the
          // specific sparse-population-unavailable diagnostic with the internal
          // convention result (public evaluateFormula surfaces '#NAME?').
          if ((maxRow - minRow + 1) === 1048576 || (maxCol - minCol + 1) === 16384) {
            const c0 = (maxRow - minRow + 1) === 1048576 ? 1048576 : (maxCol - minCol + 1)
            markUnsupported(ctx, 'sparse-population-unavailable')
            ctx?.reportFormulaIssue?.({
              kind: 'unsupported-reference',
              feature: 'sparse-population-unavailable',
              message: `Standalone full-axis range of ${c0} cells needs sparse population capability that is unavailable here`,
            })
            items.push({ value: formulaError('#NAME?'), fromRef: true })
          } else if (cellCount > 100000 && !ctx?.getRangeValues) {
            // Reviewed legacy >100000 finite-range policy (unchanged).
            items.push({ value: formulaError('#NUM!'), fromRef: true })
          } else if (ctx?.getRangeValues) {
            const grid = ctx.getRangeValues(arg.ref.sheet, arg.ref.from, arg.ref.to)
            for (const row of grid) {
              for (const cellVal of row) {
                items.push({ value: cellVal === undefined ? null : contextValue(cellVal, ctx), fromRef: true })
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
                  items.push({ value: cellVal === undefined ? null : contextValue(cellVal, ctx), fromRef: true })
                } finally {
                  ctx.visited?.delete(key)
                }
              }
            }
          }
        } else if (arg.type === 'cell') {
          const val = evalNode(arg, ctx)
          items.push({ value: val, fromRef: true })
        } else if (arg.type === 'union' || arg.type === 'intersect' || arg.type === 'wholeCol' || arg.type === 'wholeRow') {
          // Standalone legacy expansion (no services): syntactic geometry only.
          // Registry operands (names/tables) stay unavailable errors.
          for (const item of expandLegacyReference(arg, ctx, evalNode)) items.push(item)
        } else if (arg.type === 'name' || arg.type === 'table') {
          items.push({ value: formulaError('#NAME?'), fromRef: true })
        } else if (arg.type === 'ref3d' || arg.type === 'spill') {
          items.push({ value: formulaError('#REF!'), fromRef: true })
        } else if (arg.type === 'empty') {
          // Context-specific omission: NOT an ordinary blank reference.
          items.push({ value: null, fromRef: false, omitted: true })
        } else {
          // An array-capable expression argument (array constant / function
          // call / operator expression) is evaluated in ARRAY mode through the
          // same heap hook. A deliberately supplied/completed memo entry for
          // the argument is honored first (B1 owned-memo invariant); primitive
          // direct args and Core lazy/error rules keep the direct path.
          const arrayCapable = arg.type === 'arrayConst' || arg.type === 'call' || arg.type === 'binary' || arg.type === 'unary' || arg.type === 'implicitIntersect'
          const val = arrayCapable && ctx?.evaluateInMode && !ctx.nodeValues?.has(arg)
            ? ctx.evaluateInMode(arg, 'array')
            : evalNode(arg, ctx)
          if (isMatrixValue(val)) {
            // C2: a returned array/lookup matrix flattens row-major into the
            // aggregate's argument items. Elements keep reference-array
            // semantics (SUM ignores text/logical; errors still propagate),
            // which is what the documented SUM(IF(ISERROR(Data),"",Data))
            // example relies on.
            for (const row of val.values) for (const cell of row) items.push({ value: cell, fromRef: true })
          } else {
            items.push({ value: val, fromRef: false })
          }
        }
        // Commit only a completed argument; suspension must not append partial data.
        for (const item of items) progress.items.push(item)
        progress.next++
      }

      const currentAfter = ctx?.references ? currentReferenceGeneration(ctx.references) : undefined
      if (epoch.generation !== undefined && currentAfter !== undefined && epoch.generation !== currentAfter) {
        throw referenceGenerationChanged(epoch.generation, currentAfter)
      }
      epoch.restarts = 0
      return progress.items
    } catch (error) {
      if (!isReferenceGenerationChanged(error)) throw error
      resetProgress()
      if (++epoch.restarts > MAX_REFERENCE_GENERATION_RESTARTS) {
        // Application retry budget, not an invented Excel error/result.
        markUnsupported(ctx, 'reference-generation-unstable')
        ctx?.reportFormulaIssue?.({
          kind: 'unsupported-reference',
          feature: 'reference-generation-unstable',
          message: `Reference footprint changed beyond the ${MAX_REFERENCE_GENERATION_RESTARTS}-restart application budget; calculation unavailable`,
        })
        progress.items.push({ value: formulaError('#NAME?'), fromRef: true })
        progress.next = args.length
        return progress.items
      }
      requestOwnedMemoRestart(ctx)
    }
  }
}

/** Reference text and blanks do not participate; all arguments are evaluated so
 * an early false/true cannot hide a later error. IF/IFERROR remain lazy. */
function logicalArgs(args: AstNode[], ctx: EvaluationContext | undefined, evalNode: EvaluatorFn, and: boolean): EvaluationValue {
  const items = flattenArgs(args, ctx, evalNode)
  let result = and
  let count = 0
  for (const { value, fromRef } of items) {
    if (isEvaluationError(value)) return value
    if (fromRef && (value === null || typeof value === 'string')) continue
    const boolean = coerceToBoolean(value)
    if (isEvaluationError(boolean)) return boolean
    count++
    result = and ? result && boolean : result || boolean
  }
  return count === 0 ? formulaError('#VALUE!') : result
}

/**
 * Reference areas for the finite ROW/COLUMN/ROWS/COLUMNS helpers.
 * Syntactic grids resolve without services; names/tables/3D/unions/
 * intersections use the workbook reference services when present.
 * A value-denoting name is not a reference: helpers report #VALUE!.
 */
function helperAreas(
  arg: AstNode | undefined,
  ctx: EvaluationContext | undefined,
  evalNode: EvaluatorFn,
  allowMatrix: boolean = false,
): RefArea[] | EvaluationError {
  if (!arg || arg.type === 'empty') {
    const cell = ctx?.currentCell
    if (!cell) return formulaError('#VALUE!')
    return [{ sheetId: '', firstCol: cell.col, firstRow: cell.row, cols: 1, rows: 1 }]
  }
  if (arg.type === 'cell' || arg.type === 'range' || arg.type === 'wholeCol' || arg.type === 'wholeRow') {
    // Group 12 fix: with reference services present, helper geometry resolves
    // through the approved services path so sheet identity and coordinate
    // domains are validated (no needless value reads). Without services the
    // legacy syntactic path is kept for standalone/older callers.
    const services = ctx?.references
    if (services && ctx) {
      const resolved = services.resolve(arg, ctx)
      if (resolved === undefined) return formulaError('#VALUE!')
      if (isEvaluationError(resolved)) return resolved
      return [...resolved.areas]
    }
    const sheet = arg.ref.sheet ?? ctx?.currentSheet ?? ''
    return gridAreas(arg, sheet)
  }
  if ((arg.type === 'union' || arg.type === 'intersect') && !ctx?.references) {
    // Standalone syntactic composition; registry operands stay unavailable.
    const parts = arg.type === 'union' ? [...arg.refs] : [arg.left, arg.right]
    const lists: RefArea[][] = []
    for (const part of parts) {
      if (part.type !== 'cell' && part.type !== 'range' && part.type !== 'wholeCol' && part.type !== 'wholeRow') {
        return formulaError('#NAME?')
      }
      lists.push(gridAreas(part, part.ref.sheet ?? ctx?.currentSheet ?? ''))
    }
    if (arg.type === 'union') return lists.flat()
    return intersectAreas(lists[0], lists[1])
  }
  // Array/array-formula input for the dimension helpers: evaluate in array mode
  // and use the in-memory matrix dimensions (no value reads, no location guess).
  if (allowMatrix && !isReferenceNode(arg)) {
    const value = ctx?.evaluateInMode ? ctx.evaluateInMode(arg, 'array') : evalNode(arg, ctx)
    if (isEvaluationError(value)) return value
    if (isMatrixValue(value)) return [{ sheetId: '', firstCol: 0, firstRow: 0, cols: value.cols, rows: value.rows }]
    return [{ sheetId: '', firstCol: 0, firstRow: 0, cols: 1, rows: 1 }]
  }
  const services = ctx?.references
  if (!services || !ctx) {
    if (arg.type === 'name' || arg.type === 'table') return formulaError('#NAME?')
    return formulaError('#REF!')
  }
  if (!isReferenceNode(arg)) return formulaError('#VALUE!')
  const resolved = services.resolve(arg, ctx)
  if (resolved === undefined) return formulaError('#VALUE!')
  if (isEvaluationError(resolved)) return resolved
  return [...resolved.areas]
}

function firstArea(areas: RefArea[] | EvaluationError): RefArea | EvaluationError {
  if (isEvaluationError(areas)) return areas
  if (areas.length !== 1) return formulaError('#VALUE!')
  return areas[0]
}

/** Broadcast index into a matrix, or undefined when the position is not
 * covered. Undefined is a shape gate, never a padded null/zero. */
function broadcastAt(value: EvaluationValue, row: number, col: number): ElementValue | undefined {
  if (!isMatrixValue(value)) return value
  const r = value.rows === 1 ? 0 : row
  const c = value.cols === 1 ? 0 : col
  if (r >= value.rows || c >= value.cols) return undefined
  return value.values[r][c]
}

/** Approved contract: an unequal non-singleton conditional shape is an honest
 * capability gate (specific diagnostic + internal #NAME?), never a fabricated
 * null/zero/#N/A padded matrix and never an invented native #VALUE!. */
function shapeGate(ctx: EvaluationContext | undefined, message: string): EvaluationError {
  markUnsupported(ctx, 'array-shape-mismatch-unverified')
  ctx?.reportFormulaIssue?.({ kind: 'unsupported-reference', feature: 'array-shape-mismatch-unverified', message })
  return formulaError('#NAME?')
}

/** Documented elementwise IF: an array condition selects the branch per cell.
 * Both branches are evaluated because an array IF fundamentally needs both
 * arrays; scalar branches broadcast. Scalar conditions never reach here. A
 * selected cell outside its branch array is a shape gate (no null padding). */
function selectElementwise(
  mask: MatrixValue,
  args: AstNode[],
  ctx: EvaluationContext | undefined,
  evalNode: EvaluatorFn,
): EvaluationValue {
  const thenValue: EvaluationValue = args[1].type === 'empty' ? 0 : evalNode(args[1], ctx)
  const elseValue: EvaluationValue = args.length === 3
    ? (args[2].type === 'empty' ? 0 : evalNode(args[2], ctx))
    : false
  const values: ElementValue[][] = []
  for (let row = 0; row < mask.rows; row++) {
    const out: ElementValue[] = []
    for (let col = 0; col < mask.cols; col++) {
      const flag = coerceToBoolean(mask.values[row][col])
      if (isEvaluationError(flag)) return flag
      const element = broadcastAt(flag ? thenValue : elseValue, row, col)
      if (element === undefined) return shapeGate(ctx, 'IF branch array dimensions do not cover the selected cells; no null/zero padding is invented')
      out.push(element)
    }
    values.push(out)
  }
  return matrixOf(mask.rows, mask.cols, values)
}

export const FUNCTIONS: Record<string, FunctionHandler> = {
  ...CONDITIONAL_FUNCTIONS,
  ...LOOKUP_FUNCTIONS,
  ...DATETIME_FUNCTIONS,
  ...TEXT_FUNCTIONS,
  SEQUENCE: sequenceHandler(),
  TRANSPOSE: transposeHandler(),
  UNIQUE: uniqueHandler(),
  SORT: sortHandler(),
  SORTBY: sortByHandler(),
  FILTER: filterHandler(),
  SUM: (args, ctx, evalNode) => {    const items = flattenArgs(args, ctx, evalNode)
    let sum = 0
    for (const item of items) {
      const v = item.value
      if (isEvaluationError(v)) return v
      if (item.fromRef) {
        if (typeof v === 'number') {
          sum += v
        }
      } else {
        if (v === null || v === undefined) continue
        const n = coerceToNumber(v)
        if (isEvaluationError(n)) return n
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
      if (isEvaluationError(v)) return v
      if (item.fromRef) {
        if (typeof v === 'number') {
          sum += v
          count++
        }
      } else if (item.omitted) {
        // Explicit omitted argument participates as 0 (native AVERAGE(,)=0).
        count++
      } else {
        if (v === null || v === undefined) continue
        const n = coerceToNumber(v)
        if (isEvaluationError(n)) return n
        sum += n
        count++
      }
    }
    if (count === 0) return formulaError('#DIV/0!')
    return finiteOrNum(sum / count)
  },

  MIN: (args, ctx, evalNode) => {
    const items = flattenArgs(args, ctx, evalNode)
    let min: number | undefined
    for (const item of items) {
      const v = item.value
      if (isEvaluationError(v)) return v
      let num: number | undefined
      if (item.fromRef) {
        if (typeof v === 'number') num = v
      } else if (item.omitted) {
        num = 0 // explicit omitted argument participates as 0 (native MIN(5,)=0)
      } else {
        if (v === null || v === undefined) continue
        const n = coerceToNumber(v)
        if (isEvaluationError(n)) return n
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
      if (isEvaluationError(v)) return v
      let num: number | undefined
      if (item.fromRef) {
        if (typeof v === 'number') num = v
      } else if (item.omitted) {
        num = 0 // explicit omitted argument participates as 0 (native MAX(-5,)=0)
      } else {
        if (v === null || v === undefined) continue
        const n = coerceToNumber(v)
        if (isEvaluationError(n)) return n
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
      // Native COUNT never propagates or counts errors (COUNT(1,#N/A)=1).
      if (isEvaluationError(v)) continue
      if (item.omitted) { count++; continue }
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
          // COUNT counts only Excel-numeric text ("10" yes; ""/" "/"0x10" no).
          const n = coerceToNumber(v)
          if (!isEvaluationError(n)) {
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
      if (item.omitted) { count++; continue } // native COUNTA(,)=2
      const v = item.value
      if (isEvaluationError(v)) {
        count++
        continue
      }
      if (item.fromRef) {
        // References: blanks do not count; a referenced EMPTY STRING does (native).
        if (v !== null && v !== undefined) count++
      } else {
        // Direct arguments count even when empty strings (native COUNTA("")=1).
        count++
      }
    }
    return count
  },

  ABS: (args, ctx, evalNode) => {
    if (args.length !== 1) return formulaError('#VALUE!')
    const val = evalNode(args[0], ctx)
    if (isEvaluationError(val)) return val
    const num = coerceToNumber(val)
    if (isEvaluationError(num)) return num
    return finiteOrNum(Math.abs(num))
  },

  ROUND: (args, ctx, evalNode) => {
    if (args.length !== 2) return formulaError('#VALUE!')
    const val = evalNode(args[0], ctx)
    if (isEvaluationError(val)) return val
    const num = coerceToNumber(val)
    if (isEvaluationError(num)) return num

    const digitsVal = evalNode(args[1], ctx)
    if (isEvaluationError(digitsVal)) return digitsVal
    const digitsNum = coerceToNumber(digitsVal)
    if (isEvaluationError(digitsNum)) return digitsNum

    if (!Number.isFinite(num) || !Number.isFinite(digitsNum)) return formulaError('#NUM!')

    const digits = Math.trunc(digitsNum)
    return decimalRound(num, digits)
  },

  INT: (args, ctx, evalNode) => {
    if (args.length !== 1) return formulaError('#VALUE!')
    const val = evalNode(args[0], ctx)
    if (isEvaluationError(val)) return val
    const num = coerceToNumber(val)
    if (isEvaluationError(num)) return num
    // Native INT truncates the ACTUAL value (no 15-digit pre-rounding).
    return finiteOrNum(Math.floor(num))
  },

  MOD: (args, ctx, evalNode) => {
    if (args.length !== 2) return formulaError('#VALUE!')
    const val = evalNode(args[0], ctx)
    if (isEvaluationError(val)) return val
    const num = coerceToNumber(val)
    if (isEvaluationError(num)) return num

    const divVal = evalNode(args[1], ctx)
    if (isEvaluationError(divVal)) return divVal
    const div = coerceToNumber(divVal)
    if (isEvaluationError(div)) return div

    if (div === 0) return formulaError('#DIV/0!')
    // Native remainder (sign follows the divisor), one final normalization.
    let rem = num % div
    if (rem !== 0 && (rem < 0) !== (div < 0)) rem += div
    return finiteOrNum(rem)
  },

  PRODUCT: (args, ctx, evalNode) => {
    const items = flattenArgs(args, ctx, evalNode)
    let prod: number | undefined
    for (const item of items) {
      const v = item.value
      if (isEvaluationError(v)) return v
      if (item.omitted) continue // native PRODUCT(5,)=5 (omitted args ignored)
      let num: number | undefined
      if (item.fromRef) {
        if (typeof v === 'number') num = v
      } else {
        if (v === null || v === undefined) continue
        const n = coerceToNumber(v)
        if (isEvaluationError(n)) return n
        num = n
      }
      if (num !== undefined) {
        prod = prod === undefined ? num : prod * num
      }
    }
    return prod === undefined ? 0 : finiteOrNum(prod)
  },

  IF: (args, ctx, evalNode) => {
    if (args.length < 2 || args.length > 3) return formulaError('#VALUE!')
    const condVal = evalNode(args[0], ctx)
    if (isEvaluationError(condVal)) return condVal
    // Documented elementwise IF (Microsoft array guidelines): an array condition
    // selects the branch per cell. Scalar conditions keep the existing lazy path.
    if (isMatrixValue(condVal)) return selectElementwise(condVal, args, ctx, evalNode)
    const bool = coerceToBoolean(condVal)
    if (isEvaluationError(bool)) return bool

    if (bool) {
      // Selected omitted arg or blank reference scalarizes to 0 (native
      // IF(TRUE,,7)=0, IF(TRUE,D1,7)=0 with blank D1).
      if (args[1].type === 'empty') return 0
      const v = evalNode(args[1], ctx)
      return isEvaluationError(v) ? v : (v === null ? 0 : v)
    } else {
      if (args.length === 3) {
        if (args[2].type === 'empty') return 0
        const v = evalNode(args[2], ctx)
        return isEvaluationError(v) ? v : (v === null ? 0 : v)
      }
      return false
    }
  },

  IFERROR: (args, ctx, evalNode) => {
    if (args.length !== 2) return formulaError('#VALUE!')
    const val = evalNode(args[0], ctx)
    if (isEvaluationError(val)) {
      return evalNode(args[1], ctx)
    }
    // Documented IFERROR array output replaces only genuine error cells; the
    // fallback is evaluated only when an error is actually present (lazy).
    if (isMatrixValue(val)) {
      const hasError = val.values.some((row) => row.some(isEvaluationError))
      if (!hasError) return val
      const fallback = evalNode(args[1], ctx)
      const fallbackMatrix = isMatrixValue(fallback) ? fallback : undefined
      let gated = false
      const values = val.values.map((row, r) => row.map((element, c): ElementValue => {
        if (!isEvaluationError(element)) return element
        if (fallbackMatrix) {
          const picked = broadcastAt(fallbackMatrix, r, c)
          if (picked === undefined) { gated = true; return null }
          return picked
        }
        return fallback as ElementValue
      }))
      if (gated) return shapeGate(ctx, 'IFERROR fallback array dimensions do not cover the error cells; no null/zero padding is invented')
      return matrixOf(val.rows, val.cols, values)
    }
    return val
  },

  AND: (args, ctx, evalNode) => logicalArgs(args, ctx, evalNode, true),

  OR: (args, ctx, evalNode) => logicalArgs(args, ctx, evalNode, false),

  NOT: (args, ctx, evalNode) => {
    if (args.length !== 1) return formulaError('#VALUE!')
    const v = evalNode(args[0], ctx)
    if (isEvaluationError(v)) return v
    const b = coerceToBoolean(v)
    if (isEvaluationError(b)) return b
    return !b
  },

  CONCAT: (args, ctx, evalNode) => {
    const items = flattenArgs(args, ctx, evalNode)
    let res = ''
    for (const item of items) {
      const v = item.value
      if (isEvaluationError(v)) return v
      res += coerceToString(v)
    }
    return res
  },

  LEFT: (args, ctx, evalNode) => {
    if (args.length < 1 || args.length > 2) return formulaError('#VALUE!')
    const val = evalNode(args[0], ctx)
    if (isEvaluationError(val)) return val
    const str = coerceToString(val)

    let num = 1
    if (args.length === 2) {
      const numVal = evalNode(args[1], ctx)
      if (isEvaluationError(numVal)) return numVal
      const n = coerceToNumber(numVal)
      if (isEvaluationError(n)) return n
      num = Math.floor(n)
    }
    if (num < 0) return formulaError('#VALUE!')
    // LEFT/RIGHT already count surrogate pairs as one character in native
    // Excel, independently of the compatibility version used by LEN/MID.
    const chars = Array.from(str)
    return chars.slice(0, num).join('')
  },

  RIGHT: (args, ctx, evalNode) => {
    if (args.length < 1 || args.length > 2) return formulaError('#VALUE!')
    const val = evalNode(args[0], ctx)
    if (isEvaluationError(val)) return val
    const str = coerceToString(val)

    let num = 1
    if (args.length === 2) {
      const numVal = evalNode(args[1], ctx)
      if (isEvaluationError(numVal)) return numVal
      const n = coerceToNumber(numVal)
      if (isEvaluationError(n)) return n
      num = Math.floor(n)
    }
    if (num < 0) return formulaError('#VALUE!')
    if (num === 0) return ''
    const chars = Array.from(str)
    return chars.slice(-num).join('')
  },

  MID: (args, ctx, evalNode) => {
    if (args.length !== 3) return formulaError('#VALUE!')
    const val = evalNode(args[0], ctx)
    if (isEvaluationError(val)) return val
    const str = coerceToString(val)

    const startVal = evalNode(args[1], ctx)
    if (isEvaluationError(startVal)) return startVal
    const startNum = coerceToNumber(startVal)
    if (isEvaluationError(startNum)) return startNum

    const numVal = evalNode(args[2], ctx)
    if (isEvaluationError(numVal)) return numVal
    const numChars = coerceToNumber(numVal)
    if (isEvaluationError(numChars)) return numChars

    const start = Math.floor(startNum)
    const len = Math.floor(numChars)
    if (start < 1 || len < 0) return formulaError('#VALUE!')
    const chars = textChars(str, ctx)
    return chars.slice(start - 1, start - 1 + len).join('')
  },

  LEN: (args, ctx, evalNode) => {
    if (args.length !== 1) return formulaError('#VALUE!')
    const val = evalNode(args[0], ctx)
    if (isEvaluationError(val)) return val
    const str = coerceToString(val)
    return textChars(str, ctx).length
  },

  TRIM: (args, ctx, evalNode) => {
    if (args.length !== 1) return formulaError('#VALUE!')
    const val = evalNode(args[0], ctx)
    if (isEvaluationError(val)) return val
    const str = coerceToString(val)
    return str.replace(/ +/g, ' ').replace(/^ +| +$/g, '')
  },

  UPPER: (args, ctx, evalNode) => {
    if (args.length !== 1) return formulaError('#VALUE!')
    const val = evalNode(args[0], ctx)
    if (isEvaluationError(val)) return val
    return coerceToString(val).toUpperCase()
  },

  LOWER: (args, ctx, evalNode) => {
    if (args.length !== 1) return formulaError('#VALUE!')
    const val = evalNode(args[0], ctx)
    if (isEvaluationError(val)) return val
    return coerceToString(val).toLowerCase()
  },

  // Finite B1 reference helpers (not added catalog scope): NA plus the
  // ROW/COLUMN/ROWS/COLUMNS geometry helpers that reference families need.
  NA: (args) => {
    if (args.length !== 0) return formulaError('#VALUE!')
    return formulaError('#N/A')
  },

  ROW: (args, ctx, evalNode) => {
    if (args.length > 1) return formulaError('#VALUE!')
    const area = firstArea(helperAreas(args[0], ctx, evalNode))
    if (isEvaluationError(area)) return area
    if (area.rows <= 0) return formulaError('#VALUE!')
    // Documented array-mode profile: ROW(refRange) is the vertical row vector.
    if (ctx?.mode === 'array' && area.rows > 1) {
      const values: ElementValue[][] = []
      for (let i = 0; i < area.rows; i++) values.push([area.firstRow + 1 + i])
      return matrixOf(area.rows, 1, values)
    }
    return area.firstRow + 1
  },

  COLUMN: (args, ctx, evalNode) => {
    if (args.length > 1) return formulaError('#VALUE!')
    const area = firstArea(helperAreas(args[0], ctx, evalNode))
    if (isEvaluationError(area)) return area
    if (area.cols <= 0) return formulaError('#VALUE!')
    // Documented array-mode profile: COLUMN(refRange) is the horizontal vector.
    if (ctx?.mode === 'array' && area.cols > 1) {
      const values: ElementValue[] = []
      for (let i = 0; i < area.cols; i++) values.push(area.firstCol + 1 + i)
      return matrixOf(1, area.cols, [values])
    }
    return area.firstCol + 1
  },

  ROWS: (args, ctx, evalNode) => {
    if (args.length !== 1) return formulaError('#VALUE!')
    const area = firstArea(helperAreas(args[0], ctx, evalNode, true))
    if (isEvaluationError(area)) return area
    return area.rows
  },

  COLUMNS: (args, ctx, evalNode) => {
    if (args.length !== 1) return formulaError('#VALUE!')
    const area = firstArea(helperAreas(args[0], ctx, evalNode, true))
    if (isEvaluationError(area)) return area
    return area.cols
  },
}
