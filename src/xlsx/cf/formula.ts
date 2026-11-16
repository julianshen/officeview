/**
 * CF formula compilation and admission (Stage 5, Task 32; Rulings S5-B, S5-SUMIF-SHAPE, S5-L, S5-M).
 *
 * This module is the compile and admission half only: it parses a CF formula once, walks the AST statically against a
 * positive allow-list, records the facts later tasks need (atoms, invariance, clock use, families, root kinds) and names
 * the capabilities a formula needs. It evaluates nothing and reads no cell. It also owns the evaluation-context contract
 * (calculation mode, value policy, clock) that every preparation route forwards.
 *
 * Admission is static on the AST, never lazy: an untaken `IF` branch is still walked, a function name is looked up in a
 * frozen table (a later `registerFormulaFunction` cannot widen it), and every node type outside the allow-list refuses.
 * The CF engine stays out of the eager and paged-worker closures; this file imports the formula engine modules directly
 * (never the `../formula` barrel) and no eager site imports it.
 */
import type { AstNode, CellRef, ElementValue, EvaluationContext, EvaluationValue, FormulaError, MaterializationBudget, MatrixValue, ModeFrameStores, ReferenceIssue, ResolvedSemantics } from '../formula/types'
import { parseFormula } from '../formula/parser'
// functions.ts before evaluator.ts and cond.ts: the engine modules import each other, and functions.ts must finish
// evaluating before the conditional table is read.
import { FUNCTIONS } from '../formula/functions'
import { coerceToNumber, coerceToString, evaluateNode, isEvaluationError, isMatrixValue } from '../formula/evaluator'
import { CONDITIONAL_MAX_CRITERION_LENGTH } from '../formula/functions/cond'
import { DEFAULT_MAX_DP_CELLS, tokenizeWildcardPattern } from '../formula/functions/wildcard'
import { createReferenceServices } from '../formula/refs'
import { epochToSerial } from '../formula/serial'
import type { XlsxDocument } from '../types'
import type { ValuePolicy } from '../view/contracts'
import {
  CF_CHAR_SLOT_BYTES, CF_FLAT_ITEM_BYTES, CF_FORMULA_ATOM_BYTES, CF_FORMULA_CALL_BYTES, CF_FORMULA_FIXED_BYTES, CF_VIEW_WORKSPACE_BYTES, CF_FORMULA_NODE_BYTES,
  CF_FORMULA_TOKEN_BYTES, CF_LIMITS, CF_MATCHER_CELL_BYTES, CF_POPULATED_SLOT_BYTES, CF_SLOT_BYTES, CF_TRAVERSAL_BYTES, CF_VIEW_KEY_BYTES,
  CF_WORK_REFUSED, type CfMemory, type CfUnitCeiling, type CfWorkBudget,
} from './limits'
import { MEASURE_ENTRY_BYTES } from '../view/payloadBytes'
import { measurePayloadBytes } from '../view/payloadBytes'
import type { CfFormula, CfFormulaRef, CfLease, CfLiteral, CfOutcome, CfRuleType, CfTruthKind, SourceRect } from './model'
import type { CfDependencySnapshot, CfDependencyView } from './dependencies'
import { EDGE_UNMEASURED, shiftPair, translateCfFormula, type CfEdgeMode } from './translate'

export type { CfFormula, CfFormulaRef, CfTruthKind }

/** Crude lexical token count (no evaluation): identifiers, numbers, strings, operators. Stops past the token cap. */
export function countFormulaTokens(text: string): number {
  let tokens = 0
  let i = 0
  const push = () => {
    tokens++
    return tokens > CF_LIMITS.formulaTokens
  }
  while (i < text.length) {
    const ch = text[i]!
    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      i++
      continue
    }
    if (ch === '"') {
      let j = i + 1
      while (j < text.length) {
        if (text[j] === '"') {
          if (text[j + 1] === '"') j += 2
          else break
        } else j++
      }
      i = Math.min(text.length, j + 1)
      if (push()) return tokens
      continue
    }
    if (/[A-Za-z_]/.test(ch)) {
      let j = i + 1
      while (j < text.length && /[A-Za-z0-9_.]/.test(text[j]!)) j++
      i = j
      if (push()) return tokens
      continue
    }
    if (/[0-9.]/.test(ch)) {
      let j = i + 1
      while (j < text.length && /[0-9.]/.test(text[j]!)) j++
      i = j
      if (push()) return tokens
      continue
    }
    i++
    if (push()) return tokens
  }
  return tokens
}

/* ------------------------------------------------------------------ */
/* The frozen allow-list (OD-13, decided 2026-11-16).                  */
/* ------------------------------------------------------------------ */

type Slot = 'scalar' | 'range'
/** Effective read geometry of one argument (Ruling S5-SUMIF-SHAPE): its own reference, or resized to another argument's. */
export type CfFormulaGeometry = 'literal' | Readonly<{ resizedTo: number }>
/**
 * One admitted function. `args` are the leading slots; `rest` is a repeating cycle for the arguments beyond them (absent:
 * the function takes no more than `args.length`). A `range` slot admits a multi-cell reference; a `scalar` slot does not
 * (Ruling S5-B round-2 #4). `geometry` has one entry per leading argument; every repeated argument reads its own reference.
 */
export interface CfFormulaFunction {
  readonly family: string
  readonly args: readonly Slot[]
  readonly rest?: readonly Slot[]
  readonly text?: true
  readonly geometry: readonly CfFormulaGeometry[]
}

const literals = (n: number): readonly CfFormulaGeometry[] => Array.from({ length: n }, () => 'literal' as const)
const fn = (family: string, args: readonly Slot[], extra: { rest?: readonly Slot[]; text?: true; geometry?: readonly CfFormulaGeometry[] } = {}): CfFormulaFunction =>
  Object.freeze({
    family, args: Object.freeze([...args]),
    ...(extra.rest ? { rest: Object.freeze([...extra.rest]) } : {}),
    ...(extra.text ? { text: true as const } : {}),
    geometry: Object.freeze([...(extra.geometry ?? literals(args.length))]),
  })
const S: Slot = 'scalar', R: Slot = 'range'
const TABLE: readonly (readonly [string, CfFormulaFunction])[] = [
  ['AND', fn('logical', [], { rest: [S] })], ['OR', fn('logical', [], { rest: [S] })], ['NOT', fn('logical', [S])],
  ['IF', fn('logical', [S, S, S])], ['IFERROR', fn('logical', [S, S])], ['IFNA', fn('logical', [S, S])],
  ['ISBLANK', fn('information', [S])], ['ISERROR', fn('information', [S])], ['ISNA', fn('information', [S])],
  ['ISNUMBER', fn('information', [S])], ['ISTEXT', fn('information', [S])], ['ISLOGICAL', fn('information', [S])],
  ['LEN', fn('text', [S], { text: true })], ['TRIM', fn('text', [S], { text: true })], ['UPPER', fn('text', [S], { text: true })],
  ['LOWER', fn('text', [S], { text: true })], ['LEFT', fn('text', [S, S], { text: true })], ['RIGHT', fn('text', [S, S], { text: true })],
  ['MID', fn('text', [S, S, S], { text: true })],
  // CONCAT takes scalar arguments only: its output size would depend on the payload of a range (S5-B round-2 #1).
  ['CONCAT', fn('text', [], { rest: [S], text: true })],
  ['ABS', fn('math', [S])], ['INT', fn('math', [S])], ['MOD', fn('math', [S, S])], ['ROUND', fn('math', [S, S])],
  ['ROW', fn('row-column', [S])], ['COLUMN', fn('row-column', [S])],
  ['SUM', fn('aggregate', [], { rest: [R] })], ['COUNT', fn('aggregate', [], { rest: [R] })], ['COUNTA', fn('aggregate', [], { rest: [R] })],
  ['AVERAGE', fn('aggregate', [], { rest: [R] })], ['MIN', fn('aggregate', [], { rest: [R] })], ['MAX', fn('aggregate', [], { rest: [R] })],
  ['COUNTIF', fn('conditional-aggregate', [R, S])],
  ['COUNTIFS', fn('conditional-aggregate', [], { rest: [R, S] })],
  // SUMIF/AVERAGEIF resize the sum range to the criteria range's shape (cond.ts resizeTarget): admitted only when the two
  // shapes already agree, so the cells read are exactly the literal atoms (Ruling S5-SUMIF-SHAPE).
  ['SUMIF', fn('conditional-aggregate', [R, S, R], { geometry: ['literal', 'literal', { resizedTo: 0 }] })],
  ['AVERAGEIF', fn('conditional-aggregate', [R, S, R], { geometry: ['literal', 'literal', { resizedTo: 0 }] })],
  ['SUMIFS', fn('conditional-aggregate', [R], { rest: [R, S] })], ['AVERAGEIFS', fn('conditional-aggregate', [R], { rest: [R, S] })],
  ['DATE', fn('date', [S, S, S])], ['YEAR', fn('date', [S])], ['MONTH', fn('date', [S])], ['DAY', fn('date', [S])], ['WEEKDAY', fn('date', [S, S])],
  // TODAY reads the preparation clock (Ruling S5-M); NOW is withheld (OD-14).
  ['TODAY', fn('today', [])],
]
/** NAME to family and argument profile. Frozen: a name outside it never compiles, whatever the engine registers later. */
export const CF_FORMULA_FUNCTIONS: ReadonlyMap<string, CfFormulaFunction> = (() => {
  const map = new Map(TABLE)
  for (const [name] of TABLE) if (typeof FUNCTIONS[name] !== 'function') throw new Error(`CF allow-list names an unregistered function: ${name}`)
  return Object.freeze(map) as ReadonlyMap<string, CfFormulaFunction>
})()
/** The handlers present when this module loaded: a call compiles only while the engine still holds the same handler. */
const BOUND_HANDLERS: ReadonlyMap<string, unknown> = new Map(TABLE.map(([name]) => [name, FUNCTIONS[name]]))

export type CfFormulaRefusal = Readonly<{ refused: 'cf-unsupported-rule' | 'cf-resource-limit'; reason: string }>
const unsupported = (reason: string): CfFormulaRefusal => ({ refused: 'cf-unsupported-rule', reason })
const limited = (reason: string): CfFormulaRefusal => ({ refused: 'cf-resource-limit', reason })

/* ------------------------------------------------------------------ */
/* Compile: parse once, walk statically.                               */
/* ------------------------------------------------------------------ */

const MAX_ROW = 1048576, MAX_COL = 16384
/** Object nesting of the AST (a call's `args` array is a level of its own). A resident document is structured-cloned to
 * workers, and the clone recurses per level: V8 throws past about 1,800 levels here, so the compile refuses at 1,024, well
 * inside it. A 4,096-node formula is still admitted when it is not one operator chain (`1*1^1+1*1^1+...` is about 700 deep);
 * a single 4,096-operand chain is not, and refuses as `cf-resource-limit` like any other cap. */
const MAX_AST_DEPTH = 1024
const COMPARISONS: ReadonlySet<string> = new Set(['=', '<>', '<', '<=', '>', '>='])
const BOOLEAN_FUNCTIONS: ReadonlySet<string> = new Set(['AND', 'OR', 'NOT', 'ISBLANK', 'ISERROR', 'ISNA', 'ISNUMBER', 'ISTEXT', 'ISLOGICAL'])
const TEXT_FUNCTIONS: ReadonlySet<string> = new Set(['TRIM', 'UPPER', 'LOWER', 'LEFT', 'RIGHT', 'MID', 'CONCAT'])
const ALL_VALUE_KINDS: readonly CfTruthKind[] = ['number', 'boolean', 'text', 'blank', 'error']

type RefNode = Extract<AstNode, { type: 'cell' | 'range' | 'wholeCol' | 'wholeRow' }>
const isRefNode = (node: AstNode): node is RefNode => node.type === 'cell' || node.type === 'range' || node.type === 'wholeCol' || node.type === 'wholeRow'
type Frame = { node: AstNode; slot: 'root' | 'scalar' | 'range'; omittable: boolean; depth: number }

/** `[` in a sheet qualifier is an external workbook (`[1]Sheet1`). */
const externalQualifier = (sheet: string | undefined): boolean => sheet !== undefined && sheet.includes('[')
/** A sheet name that needs quoting in a formula (anything but a plain identifier). */
const needsQuoting = (sheet: string): boolean => !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet)

/** The atom of a reference node, or a refusal (mixed `$` pins across the endpoints of one range cannot be one atom). */
function atomOf(node: RefNode): CfFormulaRef | CfFormulaRefusal {
  switch (node.type) {
    case 'cell': {
      const r = node.ref
      return { ...(r.sheet !== undefined ? { sheet: r.sheet } : {}), firstRow: r.row, lastRow: r.row, firstCol: r.col, lastCol: r.col, absRow: r.absRow, absCol: r.absCol }
    }
    case 'range': {
      const { from, to } = node.ref
      if (from.absRow !== to.absRow || from.absCol !== to.absCol) return unsupported('a range whose endpoints disagree on `$` pinning')
      const sheet = node.ref.sheet ?? from.sheet
      return {
        ...(sheet !== undefined ? { sheet } : {}),
        firstRow: Math.min(from.row, to.row), lastRow: Math.max(from.row, to.row), firstCol: Math.min(from.col, to.col), lastCol: Math.max(from.col, to.col),
        absRow: from.absRow, absCol: from.absCol,
      }
    }
    case 'wholeCol': {
      const { from, to } = node.ref
      if (from.absolute !== to.absolute) return unsupported('a column span whose endpoints disagree on `$` pinning')
      return {
        ...(node.ref.sheet !== undefined ? { sheet: node.ref.sheet } : {}),
        firstRow: 0, lastRow: MAX_ROW - 1, firstCol: Math.min(from.index, to.index), lastCol: Math.max(from.index, to.index), absRow: true, absCol: from.absolute, axis: 'col',
      }
    }
    case 'wholeRow': {
      const { from, to } = node.ref
      if (from.absolute !== to.absolute) return unsupported('a row span whose endpoints disagree on `$` pinning')
      return {
        ...(node.ref.sheet !== undefined ? { sheet: node.ref.sheet } : {}),
        firstRow: Math.min(from.index, to.index), lastRow: Math.max(from.index, to.index), firstCol: 0, lastCol: MAX_COL - 1, absRow: from.absolute, absCol: true, axis: 'row',
      }
    }
  }
}
const isRefusal = (v: unknown): v is CfFormulaRefusal => typeof v === 'object' && v !== null && 'refused' in v
const shapeOf = (a: CfFormulaRef): string => `${a.lastRow - a.firstRow + 1}x${a.lastCol - a.firstCol + 1}`
const isMultiCell = (a: CfFormulaRef): boolean => a.lastRow !== a.firstRow || a.lastCol !== a.firstCol
const finiteRef = (a: CfFormulaRef): boolean => [a.firstRow, a.lastRow, a.firstCol, a.lastCol].every(Number.isSafeInteger)
  && a.firstRow >= 0 && a.lastRow < MAX_ROW && a.firstCol >= 0 && a.lastCol < MAX_COL

/** Static result kinds of a node, from its operator or function; a reference or an omitted argument can be any value. */
function kindsOf(node: AstNode, into: Set<CfTruthKind>): void {
  switch (node.type) {
    case 'number': into.add('number'); return
    case 'string': into.add('text'); return
    case 'boolean': into.add('boolean'); return
    case 'empty': into.add('blank'); return
    case 'cell': for (const k of ALL_VALUE_KINDS) into.add(k); return
    case 'unary': into.add('number'); return
    case 'binary':
      into.add(COMPARISONS.has(node.op) ? 'boolean' : node.op === '&' ? 'text' : 'number')
      return
    case 'call': {
      const name = node.name.toUpperCase()
      if (name === 'IF') {
        kindsOf(node.args[1] ?? { type: 'boolean', value: true }, into)
        kindsOf(node.args[2] ?? { type: 'boolean', value: false }, into)
      } else if (name === 'IFERROR' || name === 'IFNA') {
        for (const a of node.args) kindsOf(a, into)
      } else if (BOOLEAN_FUNCTIONS.has(name)) into.add('boolean')
      else if (TEXT_FUNCTIONS.has(name)) into.add('text')
      else into.add('number')
      return
    }
    default: for (const k of ALL_VALUE_KINDS) into.add(k)
  }
}

/** The pure compile step over a parsed AST: refuses statically, never reads a cell. */
function admit(ast: AstNode, anchor: { row: number; col: number } | undefined): CfFormula | CfFormulaRefusal {
  if (ast.type === 'empty' || ast.type === 'error') return unsupported(ast.type === 'error' ? `unparseable formula (${ast.error})` : 'empty formula')
  const atoms: CfFormulaRef[] = []
  const sheets: string[] = [], sheetKeys = new Set<string>()
  const families = new Set<string>()
  let nodes = 0, usesClock = false, variant = false, pinned = false
  const rootArea = isRefNode(ast) && ast.type !== 'cell'
  const stack: Frame[] = [{ node: ast, slot: 'root', omittable: false, depth: 1 }]
  while (stack.length > 0) {
    const { node, slot, omittable, depth } = stack.pop()!
    if (++nodes > CF_LIMITS.formulaTokens) return limited(`formula has more than ${CF_LIMITS.formulaTokens} nodes`)
    if (depth > MAX_AST_DEPTH) return limited('formula nesting exceeds the structured-clone depth bound')
    switch (node.type) {
      case 'number':
        if (!Number.isFinite(node.value)) return unsupported('non-finite number literal')
        break
      case 'string': case 'boolean': break
      case 'empty':
        if (!omittable) return unsupported('omitted argument outside a scalar function slot')
        break
      case 'cell': case 'range': case 'wholeCol': case 'wholeRow': {
        const sheet = node.type === 'range' ? node.ref.sheet ?? node.ref.from.sheet : node.ref.sheet
        if (externalQualifier(sheet)) return unsupported('external workbook reference')
        const atom = atomOf(node)
        if (isRefusal(atom)) return atom
        if (!finiteRef(atom)) return unsupported('reference outside the sheet grid')
        if (isMultiCell(atom) && slot === 'scalar') return unsupported('multi-cell reference where a single value is read')
        if (atoms.length >= CF_LIMITS.formulaAtoms) return limited(`formula has more than ${CF_LIMITS.formulaAtoms} reference atoms`)
        atoms.push(atom)
        if (!atom.absRow || !atom.absCol) variant = true
        const pinnedAxes = node.type === 'cell' ? node.ref.absRow || node.ref.absCol
          : node.type === 'range' ? node.ref.from.absRow || node.ref.from.absCol
          : node.ref.from.absolute
        if (pinnedAxes) pinned = true
        if (atom.sheet !== undefined) {
          const key = atom.sheet.toLowerCase()
          if (!sheetKeys.has(key)) {
            if (sheets.length >= CF_LIMITS.referencedSheets) return limited(`formula references more than ${CF_LIMITS.referencedSheets} sheets`)
            sheetKeys.add(key); sheets.push(atom.sheet)
          }
        }
        break
      }
      case 'unary': stack.push({ node: node.expr, slot: 'scalar', omittable: false, depth: depth + 1 }); break
      case 'binary':
        if (node.op === '&') families.add('text')
        stack.push({ node: node.right, slot: 'scalar', omittable: false, depth: depth + 1 }, { node: node.left, slot: 'scalar', omittable: false, depth: depth + 1 })
        break
      case 'call': {
        const name = node.name.toUpperCase()
        const entry = CF_FORMULA_FUNCTIONS.get(name)
        if (!entry || FUNCTIONS[name] !== BOUND_HANDLERS.get(name)) return unsupported(`function ${name} is not on the CF allow-list`)
        if (node.args.length > entry.args.length && !entry.rest) return unsupported(`${name} called with more arguments than its admitted profile`)
        families.add(entry.family)
        if (name === 'TODAY') usesClock = true
        if ((name === 'ROW' || name === 'COLUMN') && node.args.length === 0) variant = true
        const slots: Slot[] = []
        for (let i = 0; i < node.args.length; i++) {
          slots.push(i < entry.args.length ? entry.args[i]! : entry.rest![(i - entry.args.length) % entry.rest!.length]!)
        }
        // Effective read geometry: a function that resizes an argument to another's shape (SUMIF's sum range) is admitted
        // only when both are references of the identical shape, so its reads equal the literal atoms (S5-SUMIF-SHAPE).
        for (let i = 0; i < Math.min(entry.geometry.length, node.args.length); i++) {
          const g = entry.geometry[i]!
          if (g === 'literal') continue
          const mine = node.args[i]!, target = node.args[g.resizedTo]
          if (!isRefNode(mine) || !target || !isRefNode(target)) return unsupported(`${name} resized argument must be a reference matching its target`)
          const a = atomOf(mine), b = atomOf(target)
          if (isRefusal(a)) return a
          if (isRefusal(b)) return b
          if (shapeOf(a) !== shapeOf(b)) return unsupported(`${name} sum range shape differs from its criteria range`)
        }
        for (let i = node.args.length - 1; i >= 0; i--) {
          const arg = node.args[i]!
          const rangeSlot = slots[i] === 'range'
          stack.push({ node: arg, slot: rangeSlot && isRefNode(arg) ? 'range' : 'scalar', omittable: !rangeSlot, depth: depth + 2 })
        }
        break
      }
      default:
        // name, table, spill, arrayConst, union, intersect, ref3d, implicitIntersect, error: refused by default.
        return unsupported(`${node.type} is outside the CF formula grammar`)
    }
  }
  const kinds = new Set<CfTruthKind>()
  if (rootArea) kinds.add('range'); else kindsOf(ast, kinds)
  const formula: CfFormula = {
    ast, nodes, atoms, invariant: !variant, usesClock, families: [...families].sort(), dependencySheets: sheets,
    ...(anchor ? { anchor: { row: anchor.row, col: anchor.col } } : {}),
    pinned, rootArea, truth: [...kinds].sort(),
  }
  return formula
}

/** `compileCfFormulaWith`'s reservation hook: the absolute bytes the compile now needs held; false when it cannot be. */
export type CfFormulaReserve = (absoluteBytes: number) => boolean

/** Code units priced per source character: the lexer's token value and `raw`, the AST's literal or qualifier, an atom's
 * sheet and the dependency list can each hold a copy, and a qualified range copies its sheet into every endpoint node. */
const COPIES_PER_SOURCE_UNIT = 10
/** Bytes a measuring walk may admit as workspace: one visited-set slot per distinct object, doubling in chunks. */
const workspaceBound = (tokens: number): number => MEASURE_ENTRY_BYTES * (2 * (2 * tokens + 2 * CF_LIMITS.formulaAtoms + 16) + 16)
/** An upper bound on the tokens the lexer builds, with no allocation and valid for every input, malformed ones included:
 * every token consumes at least one non-whitespace character, and the stream ends in one `eof`. (`countFormulaTokens` is a
 * lexical estimate for the policy cap only: `.1.1.1...` counts as one token but lexes as one token per number.) */
function lexerTokenBound(source: string): number {
  let n = 1
  for (let i = 0; i < source.length; i++) {
    const c = source.charCodeAt(i)
    if (c !== 32 && c !== 9 && c !== 10 && c !== 13) n++
  }
  return n
}
/** What the compile may build, from the input alone: lexer tokens, the AST's nodes (a node needs at least one token),
 * string payloads, atoms and the wrapper. */
const retainedBound = (lexTokens: number, units: number): number =>
  lexTokens * (CF_FORMULA_NODE_BYTES + CF_FORMULA_TOKEN_BYTES) + units * 2 * COPIES_PER_SOURCE_UNIT + CF_FORMULA_FIXED_BYTES + Math.min(lexTokens, CF_LIMITS.formulaAtoms) * CF_FORMULA_ATOM_BYTES
/** The measuring walk only ever sees an admitted AST (at most `formulaTokens` nodes). */
const measureWorkspace = (lexTokens: number): number => workspaceBound(Math.min(lexTokens, CF_LIMITS.formulaTokens))

/**
 * The bytes reserved before `parseFormula` runs: a bound derived from the input alone that covers every allocation the
 * compile makes (lexer tokens, the AST with its string payloads, atoms, the wrapper) and the measuring workspace. Nothing
 * built may exceed it, so a refusal can never come after the allocation it was meant to prevent. Text is capped at
 * `CF_LIMITS.textUnits`, which keeps the bound finite for any input.
 */
export function cfFormulaReserveBound(text: string): number {
  const source = text.trim().replace(/^=/, '').trim()
  const lexTokens = lexerTokenBound(source)
  return retainedBound(lexTokens, source.length) + measureWorkspace(lexTokens)
}

/**
 * Compile with the caller owning the reservation. The input-derived bound is requested BEFORE `parseFormula` runs and
 * nothing is requested afterwards: the AST, the atoms and the measuring walk all fit inside it or the compile refuses.
 * Returns the formula, the bytes the caller now holds (`bytes`, the whole bound) and what the formula retains (`retained`,
 * its exact measured size, which the caller keeps after releasing the rest).
 */
export function compileCfFormulaWith(text: string, anchor: { row: number; col: number } | undefined, reserve: CfFormulaReserve): { formula: CfFormula; bytes: number; retained: number } | CfFormulaRefusal {
  if (typeof text !== 'string') return unsupported('formula text is not a string')
  const source = text.trim().replace(/^=/, '').trim()
  if (source === '') return unsupported('empty formula')
  if (source.length > CF_LIMITS.textUnits) return limited(`formula text exceeds ${CF_LIMITS.textUnits} units`)
  if (anchor !== undefined && !(Number.isSafeInteger(anchor.row) && Number.isSafeInteger(anchor.col) && anchor.row >= 0 && anchor.col >= 0 && anchor.row < MAX_ROW && anchor.col < MAX_COL)) {
    return unsupported('anchor outside the sheet grid')
  }
  const tokens = countFormulaTokens(source)
  if (tokens > CF_LIMITS.formulaTokens) return limited(`formula exceeds ${CF_LIMITS.formulaTokens} tokens`)
  const lexTokens = lexerTokenBound(source)
  const retainedCap = retainedBound(lexTokens, source.length), workspaceCap = measureWorkspace(lexTokens)
  const reserved = retainedCap + workspaceCap
  if (!reserve(reserved)) return limited('compiled formula reservation refused')
  let ast: AstNode
  try {
    ast = parseFormula(source)
  } catch {
    return limited('formula could not be parsed within the engine stack')
  }
  const result = admit(ast, anchor)
  if (isRefusal(result)) return result
  // The measuring walk admits its own workspace out of the reserved share, in chunks, before it grows.
  let workspace = 0
  const measured = measurePayloadBytes(result, extra => (workspace += extra) <= workspaceCap)
  if (measured === undefined) return limited('compiled formula could not be measured within its reservation')
  if (measured > retainedCap) return limited('compiled formula exceeds its reserved bound')
  return { formula: result, bytes: reserved, retained: measured }
}

/**
 * Compile a CF formula, admitted on `memory`. The returned lease owns the reservation and releases with the rule;
 * a refusal holds nothing. Never reads a cell and never evaluates.
 */
export function compileCfFormula(text: string, anchor: { row: number; col: number } | undefined, memory: CfMemory): CfLease<CfFormula> | CfFormulaRefusal {
  const hold = memory.reserve(0)
  if (hold.kind !== 'reserved') return limited('compiled formula reservation refused')
  try {
    const compiled = compileCfFormulaWith(text, anchor, bytes => hold.grow(bytes))
    if (isRefusal(compiled)) { hold.release(); return compiled }
    // The workspace and the unused share go back: the lease keeps exactly what the formula retains.
    return hold.adopt(compiled.formula, compiled.retained)
  } catch (error) {
    hold.release()
    throw error
  }
}
export const isCfFormulaRefusal = isRefusal

/* ------------------------------------------------------------------ */
/* Capabilities: the parse gate and the evaluation gate (round-1 #12).  */
/* ------------------------------------------------------------------ */

export type CfFormulaUse = 'expression' | 'operand' | 'cfvo' | 'activation'
export interface CfFormulaRuleContext { readonly dialect: 'base' | 'x14'; readonly rangeCount: number; readonly sourceType: CfRuleType }

/**
 * The capabilities one compiled formula needs. `parse` gates whether the rule becomes a formula definition at all
 * (grammar facts: dialect, anchor, pinning, cross-sheet spelling); `evaluation` gates what the evaluator may then trust
 * (function families, the truth reading of the root shape, the operand and cfvo kinds). Neither set is ever empty.
 */
export function formulaRequiredCapabilities(f: CfFormula, use: CfFormulaUse, rule: CfFormulaRuleContext): Readonly<{ parse: ReadonlySet<string>; evaluation: ReadonlySet<string> }> {
  const parse = new Set<string>([`formula:dialect-${rule.dialect}`, 'formula:anchor-single'])
  if (rule.rangeCount > 1) parse.add('formula:anchor-multi-range')
  if (f.pinned) parse.add('formula:anchor-mixed-abs')
  if (f.dependencySheets.length > 0) {
    parse.add('formula:cross-sheet-reference')
    if (f.dependencySheets.some(needsQuoting)) parse.add('formula:cross-sheet-quoted-name')
  }
  const evaluation = new Set<string>()
  for (const family of f.families) evaluation.add(`formula:fn-${family}`)
  const crossSheet = f.dependencySheets.length > 0
  const relative = !f.invariant
  // A multi-cell result is only ever a range result (Ruling S5-D), whatever the use: its projection needs its own reading.
  if (f.rootArea) evaluation.add('formula:truth-range')
  switch (use) {
    case 'expression':
    case 'activation':
      for (const kind of f.truth) evaluation.add(`formula:truth-${kind}`)
      if (use === 'activation') {
        evaluation.add(rule.sourceType === 'colorScale' ? 'formula:active-present-scale' : rule.sourceType === 'dataBar' ? 'formula:active-present-bar' : 'formula:active-present-icon')
      } else if (rule.sourceType === 'timePeriod') evaluation.add('formula:predicate-timeperiod-divergent')
      else if (rule.sourceType !== 'expression') evaluation.add('formula:predicate-divergent')
      break
    case 'operand': {
      const reference = f.ast.type === 'cell'
      evaluation.add(reference ? 'formula:cellis-operand-reference' : 'formula:cellis-operand-expression')
      if (relative) evaluation.add('formula:cellis-operand-relative')
      if (crossSheet) evaluation.add('formula:cellis-operand-cross-sheet')
      break
    }
    case 'cfvo':
      evaluation.add(f.ast.type === 'cell' ? 'formula:cfvo-reference' : 'formula:cfvo-expression')
      if (relative) evaluation.add('formula:cfvo-relative')
      if (rule.rangeCount > 1) evaluation.add('formula:cfvo-scope-multi-range')
      break
  }
  return { parse, evaluation }
}

/* ------------------------------------------------------------------ */
/* The evaluation context contract (Ruling S5-L, S5-M; round-2 #3).     */
/* ------------------------------------------------------------------ */

/** Built once per document: the policy every preparation of that document shares. */
export type CfDocumentEvaluation = Readonly<{ calcMode: 'auto' | 'manual' | 'autoNoTable'; valuePolicy: ValuePolicy; dateSystem: '1900' | '1904'; timeZone: string }>
/** One preparation's context: the document policy plus one clock snapshot (never shared between preparations). */
export type CfEvaluationContext = CfDocumentEvaluation & Readonly<{ clock: Readonly<{ epochNowMs: number; clockGen: number }> }>

/**
 * The document policy of a resident document: an undefined `valuePolicy` resolves `'engine'` and an absent `calc`
 * resolves `'auto'`, exactly as the workbook evaluator treats them. A document with no resolved semantics has no known
 * time zone (`''`): `TODAY()` then has no clock to read and the rule withholds rather than guessing a zone.
 */
export function cfDocumentEvaluationOf(doc: Pick<XlsxDocument, 'calc' | 'valuePolicy' | 'semantics'>): CfDocumentEvaluation {
  return retainedPolicy(doc, () => Object.freeze({
    calcMode: doc.calc?.calcMode ?? 'auto',
    // `workbookValuePolicy` (formula/value-policy.ts) restated: only a cached document carries the flag.
    valuePolicy: doc.valuePolicy === 'cached' ? 'cached' as const : 'engine' as const,
    dateSystem: doc.semantics?.dateSystem ?? '1900',
    timeZone: doc.semantics?.timeZone ?? '',
  }))
}
/** One frozen policy per document or workbook, built the first time it is asked for and reused: the policy is never
 * re-sampled (Ruling S5-L), so a later change to the document or the ambient zone cannot make two preparations disagree. */
const documentPolicies = new WeakMap<object, CfDocumentEvaluation>()
function retainedPolicy(owner: object, build: () => CfDocumentEvaluation): CfDocumentEvaluation {
  let policy = documentPolicies.get(owner)
  if (!policy) documentPolicies.set(owner, policy = build())
  return policy
}
/** The worker twin: a paged workbook is always under the `'cached'` value policy (the adapters' policy). The zone is the
 * same ambient zone the paged CF date context takes. */
export function cfDocumentEvaluationOfPaged(workbook: Readonly<{ calcMode?: 'auto' | 'manual' | 'autoNoTable'; formatting?: Readonly<{ dateSystem?: '1900' | '1904' }> }>): CfDocumentEvaluation {
  // The ambient zone is read once, when this workbook's policy is first built.
  return retainedPolicy(workbook, () => {
    let timeZone = 'UTC'
    try {
      const resolved = Intl.DateTimeFormat().resolvedOptions().timeZone
      if (typeof resolved === 'string' && resolved !== '') timeZone = resolved
    } catch { /* keep UTC */ }
    return Object.freeze({ calcMode: workbook.calcMode ?? 'auto', valuePolicy: 'cached' as const, dateSystem: workbook.formatting?.dateSystem ?? '1900', timeZone })
  })
}
/** One clock snapshot from the injected `now`; `clockGen` is the local day serial of that sample (the interactive
 * midnight tick keys on it). A zone the engine cannot resolve falls back to the UTC day so the generation still advances. */
export function cfEvaluationContextFor(policy: CfDocumentEvaluation, now: () => number): CfEvaluationContext {
  const epochNowMs = now()
  const serial = epochToSerial(epochNowMs, policy.timeZone, policy.dateSystem)
  const clockGen = typeof serial === 'number' ? Math.floor(serial) : Number.isFinite(epochNowMs) ? Math.floor(epochNowMs / 86_400_000) : 0
  return Object.freeze({ ...policy, clock: Object.freeze({ epochNowMs, clockGen }) })
}

/* ================================================================== */
/* Task 34: translation, footprint, truth, values, hoists, cost model, */
/* and the bounded per-cell evaluation (Rulings S5-C/D/E/N/HOIST).     */
/* ================================================================== */

// Translation (Ruling S5-C) lives in ./translate, its own module, so the evaluator's call to it is a module boundary.
export { translateCfFormula, type CfEdgeMode } from './translate'

/** The atom of an admitted reference node (admission already refused every shape `atomOf` refuses). */
function admittedAtom(node: RefNode): CfFormulaRef {
  const atom = atomOf(node)
  if (isRefusal(atom)) throw new Error('CF formula holds an unadmitted reference')
  return atom
}
/** The on-sheet spans `[first + dMin, last + dMax]` covers, plus under `wrap` the part shifted past either edge. */
function shiftedSpans(first: number, last: number, pinned: boolean, dMin: number, dMax: number, max: number, edge: CfEdgeMode | undefined): Array<[number, number]> {
  if (pinned) return [[first, last]]
  const lo = first + dMin, hi = last + dMax, out: Array<[number, number]> = []
  if (Math.min(hi, max - 1) >= Math.max(lo, 0)) out.push([Math.max(lo, 0), Math.min(hi, max - 1)])
  if (edge === 'wrap') {
    if (hi >= max) out.push([0, Math.min(hi - max, max - 1)])
    if (lo < 0) out.push([Math.max(lo + max, 0), max - 1])
  }
  return out
}
const rectArea = (r: SourceRect): number => (r.lastRow - r.firstRow + 1) * (r.lastCol - r.firstCol + 1)
const contains = (outer: SourceRect, inner: SourceRect): boolean =>
  outer.firstRow <= inner.firstRow && outer.lastRow >= inner.lastRow && outer.firstCol <= inner.firstCol && outer.lastCol >= inner.lastCol
export type CfFootprintEntry = Readonly<{ sheet?: string; rect: SourceRect }>

/**
 * The cells a formula reads over `targets` (Ruling S5-N caps table): each atom shifted by every target's offset from the
 * anchor (a relative axis spans the targets' range, a pinned axis stays), clipped to the sheet (plus the wrapped part
 * under `wrap`), whole-column and whole-row atoms clamped to `options.bounds` along their axis (the referenced sheet's
 * populated bounds, as statistics clamp). More than `CF_LIMITS.dependencyCells` cells is `'over-cap'`. With
 * `options.hoists` the atoms are the per-cell remainder's: hoisted ranges are scanned by the hoist, not snapshotted.
 */
export function cfFormulaFootprint(f: CfFormula, targets: readonly SourceRect[], anchor: Readonly<{ row: number; col: number }>, edge: CfEdgeMode | undefined,
  options: Readonly<{ bounds?: SourceRect; hoists?: readonly CfHoist[]; uncapped?: boolean }> = {}): readonly CfFootprintEntry[] | 'over-cap' {
  const out: CfFootprintEntry[] = []
  const owner: number[] = [] // the atom each entry came from
  // One entry per atom (the brief's shape: `{A1:A50, A2:A9}` keeps both); an atom's own shifts merge when one contains
  // another, and an identical rect from any atom is kept once. The cell count is the sum (an overlap overcounts: safe).
  const push = (entry: CfFootprintEntry, atomIndex: number): void => {
    for (let i = 0; i < out.length; i++) {
      const other = out[i]!
      if ((other.sheet ?? '').toLowerCase() !== (entry.sheet ?? '').toLowerCase()) continue
      const same = owner[i] === atomIndex
      if (same ? contains(other.rect, entry.rect) : sameRect(other.rect, entry.rect)) return
      if (same && contains(entry.rect, other.rect)) { out.splice(i, 1); owner.splice(i, 1); i-- }
    }
    out.push(entry); owner.push(atomIndex)
  }
  let atomIndex = -1
  // The references the per-cell call reads, from the same shape its view is built from: a hoisted call's range is
  // excluded, but a frequency hoist's criterion cell is still read per cell and stays in the footprint.
  for (const node of cfCallShape(f, options.hoists ?? []).refs) {
    const atom = admittedAtom(node)
    atomIndex++
    for (const t of targets) {
      const rows = shiftedSpans(atom.firstRow, atom.lastRow, atom.absRow, t.firstRow - anchor.row, t.lastRow - anchor.row, MAX_ROW, edge)
      const cols = shiftedSpans(atom.firstCol, atom.lastCol, atom.absCol, t.firstCol - anchor.col, t.lastCol - anchor.col, MAX_COL, edge)
      for (const [r0, r1] of rows) for (const [c0, c1] of cols) {
        let rect: SourceRect = { firstRow: r0, lastRow: r1, firstCol: c0, lastCol: c1 }
        const b = options.bounds
        if (b && atom.axis === 'col') rect = { ...rect, firstRow: Math.max(r0, b.firstRow), lastRow: Math.min(r1, b.lastRow) }
        if (b && atom.axis === 'row') rect = { ...rect, firstCol: Math.max(c0, b.firstCol), lastCol: Math.min(c1, b.lastCol) }
        if (rect.firstRow > rect.lastRow || rect.firstCol > rect.lastCol) continue
        push(atom.sheet !== undefined ? { sheet: atom.sheet, rect } : { rect }, atomIndex)
      }
    }
  }
  if (options.uncapped) return out // declared geometry (the caller caps the populated clamp)
  let cells = 0
  for (const entry of out) { cells += rectArea(entry.rect); if (cells > CF_LIMITS.dependencyCells) return 'over-cap' }
  return out
}

/* ------------------------------------------------------------------ */
/* Result shape and truth readings (Ruling S5-D).                      */
/* ------------------------------------------------------------------ */

/** The readings, one observable outcome each (S5-D round-1 #11): indistinguishable readings are already merged. */
export type CfTruthReading = 'number-nonzero-true' | 'boolean' | 'text-no-match' | 'text-true' | 'blank-no-match'
  | 'error-no-match' | 'error-halts' | 'range-implicit' | 'range-top-left' | 'range-no-match'
/** A formula's classified result. A range result keeps its matrix and, for a reference root, its area (implicit
 * intersection needs it); an engine matrix has no area, so only area-free readings can apply to it. */
export type CfResult = Readonly<{ kind: 'number' | 'boolean' | 'text' | 'blank' | 'error'; value: ElementValue } | { kind: 'range'; matrix: MatrixValue; area?: SourceRect }>
const unresolved = <T>(reason: string): CfOutcome<T> => ({ kind: 'unresolved', reason })
const RANGE_READINGS: readonly CfTruthReading[] = ['range-implicit', 'range-top-left', 'range-no-match']

/**
 * The CF truth of one result under the selected readings (S5-D). Each kind needs its own reading and evidence for one
 * kind never authorises another; two readings for the same kind are ambiguous and unresolved. A range result is projected
 * by its reading, and the projected value then needs its own reading too. `coerceToBoolean` is never consulted.
 */
export function cfTruth(result: CfResult, at: Readonly<{ row: number; col: number }>, readings: ReadonlySet<CfTruthReading>): CfOutcome<boolean | 'halt'> {
  const one = (a: CfTruthReading, b?: CfTruthReading): CfTruthReading | 'none' | 'ambiguous' => {
    const hasA = readings.has(a), hasB = b !== undefined && readings.has(b)
    return hasA && hasB ? 'ambiguous' : hasA ? a : hasB ? b! : 'none'
  }
  switch (result.kind) {
    case 'number': {
      const v = result.value
      if (typeof v !== 'number' || !Number.isFinite(v)) return unresolved('cf-formula-nonfinite')
      return readings.has('number-nonzero-true') ? { kind: 'ready', value: v !== 0 } : unresolved('cf-unmeasured-capability')
    }
    case 'boolean':
      return readings.has('boolean') && typeof result.value === 'boolean' ? { kind: 'ready', value: result.value } : unresolved('cf-unmeasured-capability')
    case 'blank': return readings.has('blank-no-match') ? { kind: 'ready', value: false } : unresolved('cf-unmeasured-capability')
    case 'text': {
      const r = one('text-true', 'text-no-match')
      return r === 'text-true' ? { kind: 'ready', value: true } : r === 'text-no-match' ? { kind: 'ready', value: false } : unresolved('cf-unmeasured-capability')
    }
    case 'error': {
      const r = one('error-halts', 'error-no-match')
      return r === 'error-halts' ? { kind: 'ready', value: 'halt' } : r === 'error-no-match' ? { kind: 'ready', value: false } : unresolved('cf-unmeasured-capability')
    }
    case 'range': {
      const picked = RANGE_READINGS.filter(x => readings.has(x))
      if (picked.length !== 1) return unresolved('cf-unmeasured-capability')
      if (picked[0] === 'range-no-match') return { kind: 'ready', value: false }
      const m = result.matrix
      let element: ElementValue | undefined
      if (picked[0] === 'range-top-left') element = m.values[0]?.[0]
      else {
        const area = result.area
        if (!area) return unresolved('cf-unmeasured-capability')
        if (m.rows === 1 && m.cols === 1) element = m.values[0]![0]
        else if (m.cols === 1 && at.row >= area.firstRow && at.row <= area.lastRow) element = m.values[at.row - area.firstRow]![0]
        else if (m.rows === 1 && at.col >= area.firstCol && at.col <= area.lastCol) element = m.values[0]![at.col - area.firstCol]
        // No intersection (or a 2-D area): Excel's outcome is unmeasured.
        else return unresolved('cf-unmeasured-capability')
      }
      if (element === undefined) return unresolved('cf-unmeasured-capability')
      const projected = classifyValue(element)
      if (projected.kind !== 'ready' || projected.value.kind === 'range') return projected.kind === 'ready' ? unresolved('cf-unmeasured-capability') : projected
      return cfTruth(projected.value, at, readings)
    }
  }
}

/* ------------------------------------------------------------------ */
/* The values a formula reads (Ruling S5-E).                           */
/* ------------------------------------------------------------------ */

const KNOWN_ERRORS: ReadonlySet<string> = new Set<FormulaError>(['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#NULL!', '#SPILL!', '#CALC!'])
const ERROR_ELEMENTS: ReadonlyMap<string, ElementValue> = new Map([...KNOWN_ERRORS].map(code => [code, Object.freeze({ kind: 'formula-error', code: code as FormulaError })]))
/**
 * The engine value of the host literal a cell shows (S5-E): blank is `null`, an error its tagged error, a date serial
 * its number. An error code the engine does not know is unrepresentable (undefined): the cell is unavailable, never a
 * guessed error.
 */
export function cfLiteralToElement(v: CfLiteral): ElementValue | undefined {
  switch (v.kind) {
    case 'blank': return null
    case 'number': return Number.isFinite(v.value) ? v.value : undefined
    case 'string': return v.value
    case 'boolean': return v.value
    case 'error': return ERROR_ELEMENTS.get(v.value)
  }
}
/** The host literal of an engine scalar (a non-finite number or a matrix has none). */
export function cfElementToLiteral(v: EvaluationValue): CfLiteral | undefined {
  if (v === null) return { kind: 'blank' }
  if (typeof v === 'number') return Number.isFinite(v) ? { kind: 'number', value: v } : undefined
  if (typeof v === 'string') return { kind: 'string', value: v }
  if (typeof v === 'boolean') return { kind: 'boolean', value: v }
  if (isMatrixValue(v)) return undefined
  return isEvaluationError(v) && KNOWN_ERRORS.has(v.code) ? { kind: 'error', value: v.code } : undefined
}
function classifyValue(v: EvaluationValue): CfOutcome<CfResult> {
  if (v === null) return { kind: 'ready', value: { kind: 'blank', value: null } }
  if (typeof v === 'number') return Number.isFinite(v) ? { kind: 'ready', value: { kind: 'number', value: v } } : unresolved('cf-formula-nonfinite')
  if (typeof v === 'string') return { kind: 'ready', value: { kind: 'text', value: v } }
  if (typeof v === 'boolean') return { kind: 'ready', value: { kind: 'boolean', value: v } }
  if (isMatrixValue(v)) return { kind: 'ready', value: { kind: 'range', matrix: v } }
  if (isEvaluationError(v)) return { kind: 'ready', value: { kind: 'error', value: v } }
  return unresolved('cf-formula-unsupported')
}

/* ------------------------------------------------------------------ */
/* Hoists (Ruling S5-HOIST).                                           */
/* ------------------------------------------------------------------ */

export type CfHoist = Readonly<{ node: AstNode; kind: 'scalar' | 'countif-frequency'; fn: string; range: Readonly<{ sheet?: string; rect: SourceRect }>; criterion?: AstNode }>
/** Hoisted results keyed by the compiled formula's own nodes. `longest` is the longest string a frequency scan met, for
 * the engine's matcher DP policy per criterion (S5-HOIST "refused per cell"). */
export type CfHoisted = Readonly<{ scalars: ReadonlyMap<AstNode, ElementValue>; frequencies: ReadonlyMap<AstNode, ReadonlyMap<string, number>>; longest: ReadonlyMap<AstNode, number> }>
const SCALAR_AGGREGATES: ReadonlySet<string> = new Set(['SUM', 'COUNT', 'COUNTA', 'AVERAGE', 'MIN', 'MAX'])
/** Every axis `$`-pinned, on the host sheet (a cross-sheet range is scanned through its own store in Task 36). */
function absoluteArea(node: AstNode): Readonly<{ rect: SourceRect }> | undefined {
  if (node.type !== 'range' && node.type !== 'wholeCol' && node.type !== 'wholeRow') return undefined
  const atom = admittedAtom(node)
  if (!atom.absRow || !atom.absCol || atom.sheet !== undefined) return undefined
  return { rect: { firstRow: atom.firstRow, lastRow: atom.lastRow, firstCol: atom.firstCol, lastCol: atom.lastCol } }
}
const invariantCriterion = (node: AstNode | undefined): boolean => node !== undefined && (node.type === 'number' || node.type === 'string'
  || (node.type === 'cell' && node.ref.absRow && node.ref.absCol && node.ref.sheet === undefined))
const relativeCellCriterion = (node: AstNode | undefined): boolean => node !== undefined && node.type === 'cell' && node.ref.sheet === undefined && (!node.ref.absRow || !node.ref.absCol)
const sameRect = (a: SourceRect, b: SourceRect): boolean => a.firstRow === b.firstRow && a.lastRow === b.lastRow && a.firstCol === b.firstCol && a.lastCol === b.lastCol

/**
 * The calls to hoist out of the per-cell evaluation (Ruling S5-HOIST): a scalar aggregate over one absolute range,
 * `COUNTIF`/`SUMIF` over an absolute range with an invariant criterion (one scalar per rule), and
 * `COUNTIF(<absolute range>, <relative cell>)` (one frequency map per rule). Only the outermost such call is hoisted.
 */
export function planCfHoist(f: CfFormula): readonly CfHoist[] {
  const out: CfHoist[] = []
  const stack: AstNode[] = [f.ast]
  while (stack.length) {
    const node = stack.pop()!
    if (node.type === 'unary') { stack.push(node.expr); continue }
    if (node.type === 'binary') { stack.push(node.right, node.left); continue }
    if (node.type !== 'call') continue
    const fnName = node.name.toUpperCase()
    const range = node.args[0] ? absoluteArea(node.args[0]) : undefined
    if (range && SCALAR_AGGREGATES.has(fnName) && node.args.length === 1) { out.push({ node, kind: 'scalar', fn: fnName, range }); continue }
    if (range && fnName === 'COUNTIF' && node.args.length === 2) {
      if (invariantCriterion(node.args[1])) { out.push({ node, kind: 'scalar', fn: fnName, range, criterion: node.args[1]! }); continue }
      if (relativeCellCriterion(node.args[1])) { out.push({ node, kind: 'countif-frequency', fn: fnName, range, criterion: node.args[1]! }); continue }
    }
    if (range && fnName === 'SUMIF' && invariantCriterion(node.args[1])) {
      const sum = node.args[2] ? absoluteArea(node.args[2]) : undefined
      if (node.args.length === 2 || (node.args.length === 3 && sum && sameRect(sum.rect, range.rect))) { out.push({ node, kind: 'scalar', fn: fnName, range, criterion: node.args[1]! }); continue }
    }
    for (let i = node.args.length - 1; i >= 0; i--) stack.push(node.args[i]!)
  }
  return out
}

/** `criterionOf`'s numeric-criterion pattern (functions/cond.ts, `criterionOf`), cited rather than restated in meaning. */
const ENGINE_NUMERIC_CRITERION = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/
const ENGINE_COMPARISON_PREFIX = /^(<=|>=|<>|=|<|>)/
const isAscii = (s: string): boolean => { for (let i = 0; i < s.length; i++) if (s.charCodeAt(i) >= 0x80) return false; return true }
/** The engine matcher's fold: A-Z only (functions/wildcard.ts `fold`). */
const asciiLower = (s: string): string => s.replace(/[A-Z]/g, ch => String.fromCharCode(ch.charCodeAt(0) + 32))
const numericLooking = (s: string): boolean => ENGINE_NUMERIC_CRITERION.test(s) || !isEvaluationError(coerceToNumber(s))

/**
 * The plain-equality COUNTIF key of a criterion (Ruling S5-HOIST), never `cfValueKey` (which merges numeric text with
 * numbers). A number keys as itself; text keys as its ASCII-folded spelling, and only when the engine would read it as a
 * plain equality criterion matched with no wildcard token. Every other criterion is `'refuse'`: wildcard characters,
 * a comparison prefix, error-literal text (`#…`), text past the engine's criterion limit, non-ASCII or numeric-looking
 * text (by the engine's own pattern or `coerceToNumber`), the empty string, a blank, a boolean or an error.
 */
export function cfCountifKey(value: CfLiteral): string | 'refuse' {
  if (value.kind === 'number') return Number.isFinite(value.value) ? `n:${value.value}` : 'refuse'
  if (value.kind !== 'string') return 'refuse'
  const s = value.value
  if (s === '' || s.length > CONDITIONAL_MAX_CRITERION_LENGTH || !isAscii(s) || /[*?~]/.test(s) || ENGINE_COMPARISON_PREFIX.test(s) || s.startsWith('#') || numericLooking(s)) return 'refuse'
  return `s:${asciiLower(s)}`
}
/** The key a range cell contributes to a frequency map: a key, `undefined` (it can never equal an admitted criterion:
 * blank, error, empty text), or `'refuse-rule'` (a value class whose engine treatment the map does not reproduce:
 * numeric-looking text, a boolean, non-ASCII text). */
export function cfCountifRangeKey(value: CfLiteral): string | undefined | 'refuse-rule' {
  switch (value.kind) {
    case 'number': return Number.isFinite(value.value) ? `n:${value.value}` : 'refuse-rule'
    case 'boolean': return 'refuse-rule'
    case 'blank': case 'error': return undefined
    case 'string': {
      const s = value.value
      if (s === '') return undefined
      if (!isAscii(s) || numericLooking(s)) return 'refuse-rule'
      return `s:${asciiLower(s)}`
    }
  }
}
/** Whether the engine's matcher would refuse this admitted text criterion against the longest string of the range
 * (DEFAULT_MAX_DP_CELLS, checked at functions/wildcard.ts before allocation; a plain criterion has one token per char). */
export function cfCountifDpRefused(criterion: CfLiteral, longestString: number): boolean {
  return criterion.kind === 'string' && (criterion.value.length + 1) * (longestString + 1) > DEFAULT_MAX_DP_CELLS
}

/* ------------------------------------------------------------------ */
/* The per-call shape and its static work bound (Ruling S5-N).         */
/* ------------------------------------------------------------------ */

/** Measured while a call's view is built: the view's cardinality and, per read reference in the call's shape order, the
 * stored cells of its area, their total string units and the longest string. */
export type CfPayloadMetrics = Readonly<{ viewCells: number; index?: number; atoms: readonly Readonly<{ cells: number; stringUnits: number; longestString: number
  /** Stored string cells, and those not all-ASCII (round 3, finding 4); absent, a string atom is taken as non-ASCII. */
  strings?: number; nonAscii?: number }>[] }>
const CONDITIONAL: ReadonlySet<string> = new Set(['COUNTIF', 'COUNTIFS', 'SUMIF', 'SUMIFS', 'AVERAGEIF', 'AVERAGEIFS'])
/** Argument roles of a conditional aggregate: `t` is a sum/average target range, `r` a criteria range, `c` a criterion. */
function conditionalRoles(name: string, count: number): Array<'t' | 'r' | 'c'> {
  if (name === 'COUNTIF') return ['r', 'c'].slice(0, count) as Array<'r' | 'c'>
  if (name === 'SUMIF' || name === 'AVERAGEIF') return (['r', 'c', 't'] as const).slice(0, count) as Array<'t' | 'r' | 'c'>
  if (name === 'COUNTIFS') return Array.from({ length: count }, (_, i) => (i % 2 === 0 ? 'r' : 'c'))
  return Array.from({ length: count }, (_, i) => (i === 0 ? 't' : i % 2 === 1 ? 'r' : 'c'))
}
type ReadKind = 'cell' | 'traversal' | 'root-area'
interface Conditional { refs: number[]; pairs: Array<{ range: number | undefined; criterion: AstNode }> }
/** The shape one per-cell call evaluates: the compiled tree with hoisted calls replaced by one node each. `refs` is the
 * order every per-call list (effective areas, payload atoms) follows. */
export interface CfCallShape {
  readonly nodes: number
  readonly refs: readonly RefNode[]
  readonly reads: readonly ReadKind[]
  /** Plain aggregate traversals (each one occurrence). */
  readonly aggregates: readonly number[]
  readonly conditionals: readonly Conditional[]
  /** Frequency-hoist lookups (each also reads its criterion cell, listed in `refs`). */
  readonly lookups: number
  /** The tree the text model walks (round 3, finding 4), with the nodes that enter the call as values. */
  readonly root: AstNode
  readonly hoisted: ReadonlyMap<AstNode, CfHoist>
  readonly memoised: ReadonlyMap<AstNode, ElementValue>
  /** A single-cell read's index in `refs`, by its node. */
  readonly cellRefs: ReadonlyMap<AstNode, number>
  /** Criterion text parsed per item, in shape order: a literal's length, or a cell read's index; with its pass count. */
  readonly criteria: readonly Readonly<{ literal?: number; ref?: number; passes: number }>[]
}
export function cfCallShape(f: CfFormula, hoisted: readonly CfHoist[] = [], memoised: ReadonlyMap<AstNode, ElementValue> = new Map()): CfCallShape {
  const hoists = new Map(hoisted.map(h => [h.node, h]))
  const refs: RefNode[] = [], reads: ReadKind[] = [], aggregates: number[] = [], conditionals: Conditional[] = []
  const cellRefs = new Map<AstNode, number>()
  const criteria: Array<{ literal?: number; ref?: number; passes: number }> = []
  let nodes = 0, lookups = 0
  const addRef = (node: RefNode, kind: ReadKind): number => { refs.push(node); reads.push(kind); if (kind === 'cell') cellRefs.set(node, refs.length - 1); return refs.length - 1 }
  const visit = (node: AstNode, root = false): void => {
    nodes++
    const hoist = hoists.get(node)
    if (hoist) {
      if (hoist.kind === 'countif-frequency' && hoist.criterion && isRefNode(hoist.criterion)) {
        criteria.push({ ref: addRef(hoist.criterion, 'cell'), passes: CF_FREQUENCY_KEY_PASSES }); lookups++
      }
      return
    }
    if (memoised.has(node)) return // enters the call as one literal (finding 8)
    switch (node.type) {
      case 'cell': addRef(node, 'cell'); return
      case 'range': case 'wholeCol': case 'wholeRow': addRef(node, root ? 'root-area' : 'traversal'); return
      case 'unary': visit(node.expr); return
      case 'binary': visit(node.left); visit(node.right); return
      case 'call': {
        const name = node.name.toUpperCase()
        if (CONDITIONAL.has(name)) {
          const roles = conditionalRoles(name, node.args.length)
          const c: Conditional = { refs: [], pairs: [] }
          let lastRange: number | undefined
          node.args.forEach((arg, i) => {
            const role = roles[i] ?? 'c'
            if (role !== 'c' && isRefNode(arg)) {
              nodes++
              const index = addRef(arg, 'traversal')
              c.refs.push(index)
              if (role === 'r') lastRange = index
            } else if (role === 'c') {
              c.pairs.push({ range: lastRange, criterion: arg }); lastRange = undefined
              // The criterion's text is parsed whatever the formula's other work (cond.ts criterionOf): charged always.
              const before = refs.length
              visit(arg)
              if (arg.type === 'string') criteria.push({ literal: arg.value.length, passes: CF_CRITERION_PASSES })
              else if (arg.type === 'cell' && refs.length === before + 1) criteria.push({ ref: before, passes: CF_CRITERION_PASSES })
            } else visit(arg)
          })
          conditionals.push(c)
          return
        }
        if (SCALAR_AGGREGATES.has(name)) {
          for (const arg of node.args) {
            if (arg.type === 'range' || arg.type === 'wholeCol' || arg.type === 'wholeRow') { nodes++; aggregates.push(addRef(arg, 'traversal')) } else visit(arg)
          }
          return
        }
        for (const arg of node.args) visit(arg)
        return
      }
      default: return
    }
  }
  visit(f.ast, true)
  return { nodes, refs, reads, aggregates, conditionals, lookups, root: f.ast, hoisted: hoists, memoised, cellRefs, criteria }
}

/** Components of one call's static bound, each charged before the work happens (S5-N). */
export type CfCallWorkBreakdown = Readonly<{ nodes: number; viewBuild: number; index: number; enumeration: number; reads: number; sort: number; comparisons: number; matcher: number; text: number
  /** Criterion text parsed (conditional criteria and frequency keys), spent on the meter before the parse. */
  criterion: number; lookups: number; total: number }>
const sortUnits = (v: number): number => (v <= 1 ? 0 : v * Math.ceil(Math.log2(v)))
/**
 * The static per-call work bound (Ruling S5-N, round-2 #2, round-4 #1): nodes + view build (v) + per traversal
 * occurrence (v enumeration visits + its matched reads + v x ceil(log2 v) sort) + single-cell reads + text units +
 * conditional comparisons and matcher transitions + frequency lookups, with v = `payload.viewCells`. A conditional
 * aggregate reads at most (its ranges) x P cells, P being the stored cells of its ranges; each criterion pair compares
 * (t + 1) x (P + the range's string units) where t is the criterion's wildcard token count (0 for a non-text criterion:
 * one unit per range cell, plus the numeric-text check over string units). `criterionLengths` gives t per pair in shape
 * order; an undefined entry (the criterion is not a literal or a snapshot cell) makes the bound `'unbounded'`.
 */
export function cfCallWorkBreakdown(shape: CfCallShape, payload: CfPayloadMetrics, criterionLengths: readonly (number | undefined)[]): CfCallWorkBreakdown | 'unbounded' {
  const v = payload.viewCells
  const atom = (i: number) => payload.atoms[i] ?? { cells: 0, stringUnits: 0, longestString: 0 }
  let enumeration = 0, reads = 0, sort = 0, comparisons = 0, matcher = 0, text = 0
  const traverse = (): void => { enumeration += v; sort += sortUnits(v) }
  shape.reads.forEach((kind, i) => {
    if (kind === 'cell') reads++
    else if (kind === 'root-area') { const a = admittedAtom(shape.refs[i]!); reads += (a.lastRow - a.firstRow + 1) * (a.lastCol - a.firstCol + 1) }
  })
  for (const i of shape.aggregates) { traverse(); reads += atom(i).cells }
  let pair = 0
  for (const c of shape.conditionals) {
    let positions = 0
    for (const i of c.refs) { traverse(); positions += atom(i).cells }
    reads += c.refs.length * positions
    for (const p of c.pairs) {
      const t = criterionLengths[pair++]
      if (t === undefined || !Number.isSafeInteger(t) || t < 0) return 'unbounded'
      const units = p.range === undefined ? 0 : atom(p.range).stringUnits
      const amount = (t + 1) * (positions + units)
      if (t === 0) comparisons += amount
      // A text criterion is validated once by `matchWildcard(body, '')` (cond.ts criterionOf) whatever the range holds:
      // one DP row per token plus the initial one, over the empty text (round 4, finding 4).
      else matcher += amount + (t + 1)
    }
  }
  const textUnits = cfTextUnits(shape, payload)
  text = textUnits.text
  const criterion = textUnits.criterion
  const lookups = shape.lookups
  // The empty index steps the view build took (finding 1), measured and charged as they happened.
  const index = payload.index ?? 0
  const total = shape.nodes + v + index + enumeration + reads + sort + comparisons + matcher + text + criterion + lookups
  return { nodes: shape.nodes, viewBuild: v, index, enumeration, reads, sort, comparisons, matcher, text, criterion, lookups, total }
}
/* ------------------------------------------------------------------ */
/* The per-node text model (round 3, finding 4).                       */
/* ------------------------------------------------------------------ */

/** The longest text the engine renders from a non-string value (`coerceToString`: `String(round15(v))` for a finite
 * number is at most 23 units, e.g. `-0.00000123456789012345`; `FALSE`; an error code). Asserted by test. */
export const CF_TEXT_RENDER_UNITS = 24
/** The most a case mapping lengthens one code unit (`'ΐ'.toUpperCase()` is three units); ASCII never grows. */
const CASE_GROWTH = 3
/** [policy] Characters of string work per unit. A unit prices one examined source cell or one engine node; a native
 * string pass over a few characters is cheaper than either, so text is charged per CF_TEXT_UNIT_CHARS characters
 * (rounded up per call). Reviewed with the caps table; the timing evidence is in the report. */
export const CF_TEXT_UNIT_CHARS = 4
/** Passes over a string operand, per engine handler (formula/evaluator.ts, functions.ts): `coerceToNumber` (trim, two
 * replaces, the numeric pattern, a suffix test, slice, search, Number), `coerceToBoolean` (trim and a case fold, then
 * the compare), a string comparison (a case fold of each side, then the compare), `textChars` (Array.from) and a join. */
const COERCE_NUMBER_PASSES = 8
/** Passes a conditional criterion's parse makes (prefix strip, numeric pattern, wildcard tokens), and an admitted
 * frequency key's costliest path (`cfCountifKey`: the ASCII test, the wildcard test, the numeric pattern, then
 * `coerceToNumber`'s trim, thousands replace and numeric test, which rejects there; the fold's replace and its output;
 * and the map lookup's hash: 9. The anchored prefix and `#` tests read O(1)). */
export const CF_CRITERION_PASSES = 3
export const CF_FREQUENCY_KEY_PASSES = 9
/** Each function's text profile: how its operands are consumed and what it returns. Exhaustive over the allow-list
 * (asserted at load), so a function added to it is never priced as free. */
type TextProfile = 'concat' | 'len' | 'trim' | 'case' | 'slice' | 'logical' | 'if' | 'pass' | 'is' | 'number-of-args' | 'position' | 'aggregate' | 'conditional'
const TEXT_PROFILE: Readonly<Record<string, TextProfile>> = {
  CONCAT: 'concat', LEN: 'len', TRIM: 'trim', UPPER: 'case', LOWER: 'case', LEFT: 'slice', RIGHT: 'slice', MID: 'slice',
  AND: 'logical', OR: 'logical', NOT: 'logical', IF: 'if', IFERROR: 'pass', IFNA: 'pass',
  ISBLANK: 'is', ISERROR: 'is', ISNA: 'is', ISNUMBER: 'is', ISTEXT: 'is', ISLOGICAL: 'is',
  ABS: 'number-of-args', INT: 'number-of-args', MOD: 'number-of-args', ROUND: 'number-of-args',
  DATE: 'number-of-args', YEAR: 'number-of-args', MONTH: 'number-of-args', DAY: 'number-of-args', WEEKDAY: 'number-of-args',
  ROW: 'position', COLUMN: 'position', TODAY: 'position',
  SUM: 'aggregate', COUNT: 'aggregate', COUNTA: 'aggregate', AVERAGE: 'aggregate', MIN: 'aggregate', MAX: 'aggregate',
  COUNTIF: 'conditional', COUNTIFS: 'conditional', SUMIF: 'conditional', SUMIFS: 'conditional', AVERAGEIF: 'conditional', AVERAGEIFS: 'conditional',
}
for (const name of CF_FORMULA_FUNCTIONS.keys()) if (!TEXT_PROFILE[name]) throw new Error(`CF text model has no profile for ${name}`)
/** One node's value as text: the most characters it can render to, whether it may be a string, and whether a string
 * it holds is ASCII (non-strings render ASCII). */
export type TextValue = Readonly<{ len: number; text: boolean; ascii: boolean }>
const NUMBER_VALUE: TextValue = { len: CF_TEXT_RENDER_UNITS, text: false, ascii: true }
const BOOLEAN_VALUE: TextValue = { len: 5, text: false, ascii: true }
const BLANK_VALUE: TextValue = { len: 0, text: false, ascii: true }
const literalValue = (v: ElementValue): TextValue => {
  if (typeof v === 'string') return { len: v.length, text: true, ascii: isAscii(v) }
  if (v === null) return BLANK_VALUE
  return { len: Math.min(coerceToString(v).length, CF_TEXT_RENDER_UNITS), text: false, ascii: true }
}
const union = (a: TextValue, b: TextValue): TextValue => ({ len: Math.max(a.len, b.len), text: a.text || b.text, ascii: a.ascii && b.ascii })
const growth = (x: TextValue): number => (x.ascii ? 1 : CASE_GROWTH)
/** Characters a coercion to a number walks (only a string is parsed). */
const numberPasses = (x: TextValue): number => (x.text ? x.len * COERCE_NUMBER_PASSES : 0)
/** Characters a coercion to a boolean walks: trim and case fold (which may grow), then the compare. */
const booleanPasses = (x: TextValue): number => (x.text ? x.len * (1 + 2 * growth(x)) : 0)
/** Characters one side of a string comparison walks: its case fold and the compare over the folded text. */
const comparePasses = (x: TextValue): number => (x.text ? x.len * (1 + growth(x)) : 0)
const charsToUnits = (chars: number): number => Math.min(Math.ceil(chars / CF_TEXT_UNIT_CHARS), Number.MAX_SAFE_INTEGER)
/** The units a criterion's parse costs: spent on the call meter before the parse, counted in the bound, per item. */
export const cfCriterionUnits = (length: number, passes = CF_CRITERION_PASSES): number => charsToUnits(length * passes)

/**
 * Text work of one call (round 3, finding 4). Every node gets its value's text bound (`TextValue`): literals and memoised
 * values exactly, a single-cell read from the view's measured cell (a string's length and ASCII flag, recorded during
 * the charged snapshot build; a number, boolean or error renders at most CF_TEXT_RENDER_UNITS; a blank renders empty),
 * a hoisted value as a number, and each operation from its operands. Each operation is charged its OWN passes over its
 * own operands and output: `&` and CONCAT render their operands and write the sum; LEN renders and splits its operand;
 * TRIM makes two replaces and a copy; UPPER/LOWER read their operand and write a result that only a non-ASCII operand
 * can lengthen (x3); LEFT/RIGHT/MID split, slice and join; comparisons fold and compare each string side; arithmetic,
 * math and date functions and directly supplied aggregate operands parse a string side; AND/OR/NOT and IF's condition
 * trim and fold it. Non-strings are never charged as text by a consumer. `intermediate` is the longest string any
 * node can hold (the scratch strings are sized from the charged characters).
 */
export function cfTextUnits(shape: CfCallShape, payload: CfPayloadMetrics, trace?: Map<AstNode, TextValue>): Readonly<{ text: number; criterion: number; intermediate: number; chars: number }> {
  const atom = (i: number) => payload.atoms[i] ?? { cells: 0, stringUnits: 0, longestString: 0 }
  let chars = 0, intermediate = 0
  const out = (v: TextValue): TextValue => { if (v.text) intermediate = Math.max(intermediate, v.len); return v }
  const cellValue = (i: number): TextValue => {
    const a = atom(i)
    if (a.cells === 0) return BLANK_VALUE
    const strings = a.strings ?? (a.stringUnits > 0 ? a.cells : 0)
    if (strings === 0) return NUMBER_VALUE
    // An atom without measured ASCII flags is conservatively non-ASCII.
    const nonAscii = a.nonAscii ?? strings
    return { len: a.stringUnits, text: true, ascii: nonAscii === 0 }
  }
  /** `trace` (tests) records every node's bound. */
  const walk = (node: AstNode): TextValue => {
    const v = walkNode(node)
    trace?.set(node, v)
    return v
  }
  const walkNode = (node: AstNode): TextValue => {
    if (shape.hoisted.has(node)) return NUMBER_VALUE
    const memo = shape.memoised.get(node)
    if (memo !== undefined) return out(literalValue(memo))
    switch (node.type) {
      case 'string': return out(literalValue(node.value))
      case 'number': return literalValue(node.value)
      case 'boolean': return BOOLEAN_VALUE
      case 'error': return { len: node.error.length, text: false, ascii: true }
      case 'empty': return NUMBER_VALUE // an omitted argument reads as 0 or blank depending on the handler: bounded as a number
      case 'cell': { const i = shape.cellRefs.get(node); return out(i === undefined ? NUMBER_VALUE : cellValue(i)) }
      case 'range': case 'wholeCol': case 'wholeRow': return BLANK_VALUE // traversed, never rendered (scalar slots refuse ranges)
      case 'unary': { chars += numberPasses(walk(node.expr)); return NUMBER_VALUE }
      case 'binary': {
        const l = walk(node.left), r = walk(node.right)
        if (node.op === '&') {
          const len = l.len + r.len
          chars += l.len + r.len + len
          return out({ len, text: true, ascii: l.ascii && r.ascii })
        }
        if (COMPARISONS.has(node.op)) { chars += comparePasses(l) + comparePasses(r); return BOOLEAN_VALUE }
        chars += numberPasses(l) + numberPasses(r)
        return NUMBER_VALUE
      }
      case 'call': {
        const name = node.name.toUpperCase()
        const profile = TEXT_PROFILE[name]
        if (profile === 'conditional') {
          // Ranges are traversed (priced by the comparison and matcher terms); a criterion expression is walked for its
          // own work, and its parse is the criterion term.
          for (const arg of node.args) if (!isRefNode(arg)) walk(arg)
          return NUMBER_VALUE
        }
        if (profile === 'aggregate') {
          for (const arg of node.args) if (arg.type !== 'range' && arg.type !== 'wholeCol' && arg.type !== 'wholeRow') chars += numberPasses(walk(arg))
          return NUMBER_VALUE
        }
        const args = node.args.map(walk)
        switch (profile) {
          case 'concat': { const len = args.reduce((n, a) => n + a.len, 0); chars += len * 2; return out({ len, text: true, ascii: args.every(a => a.ascii) }) }
          case 'len': chars += 2 * (args[0]?.len ?? 0); return NUMBER_VALUE
          case 'trim': { const x = args[0] ?? BLANK_VALUE; chars += 3 * x.len; return out({ len: x.len, text: true, ascii: x.ascii }) }
          case 'case': { const x = args[0] ?? BLANK_VALUE; const len = x.len * growth(x); chars += x.len + len; return out({ len, text: true, ascii: x.ascii }) }
          case 'slice': {
            const x = args[0] ?? BLANK_VALUE
            for (const n of args.slice(1)) chars += numberPasses(n)
            chars += 3 * x.len // split into characters, slice, join
            return out({ len: x.len, text: true, ascii: x.ascii })
          }
          case 'logical': for (const a of args) chars += booleanPasses(a); return BOOLEAN_VALUE
          case 'if': chars += booleanPasses(args[0] ?? BLANK_VALUE); return out(union(args[1] ?? BOOLEAN_VALUE, args[2] ?? BOOLEAN_VALUE))
          case 'pass': return out(args.reduce(union, BLANK_VALUE))
          case 'is': return BOOLEAN_VALUE
          case 'number-of-args': for (const a of args) chars += numberPasses(a); return NUMBER_VALUE
          case 'position': return NUMBER_VALUE
          default: throw new Error(`CF text model has no profile for ${name}`)
        }
      }
      default: return NUMBER_VALUE
    }
  }
  walk(shape.root)
  let criterion = 0
  for (const c of shape.criteria) criterion += cfCriterionUnits(c.literal ?? atom(c.ref!).stringUnits, c.passes)
  return { text: charsToUnits(chars), criterion, intermediate, chars }
}
/** The static bound of one call (or `'unbounded'`); a call is evaluated only at or under `CF_LIMITS.checkpointUnits`. */
export function cfCallWorkUnits(f: CfFormula, payload: CfPayloadMetrics, criterionLengths: readonly (number | undefined)[], hoisted: readonly CfHoist[] = planCfHoist(f)): number | 'unbounded' {
  const b = cfCallWorkBreakdown(cfCallShape(f, hoisted), payload, criterionLengths)
  return b === 'unbounded' ? b : b.total
}
/**
 * The evaluator scratch one call may allocate, reserved before invocation (S5-N, Lesson 15): engine bookkeeping, one memo
 * slot per node, the view's key list, per traversal its resolved reference, populated slots and flattened items, a root
 * area's matrix, the text intermediates (string and character array per text op plus one, for the atoms' and literals'
 * string units: S5-B round-2 #1) and the matcher's DP rows.
 */
export function cfEvaluatorScratchBytes(f: CfFormula, payload: CfPayloadMetrics, hoisted: readonly CfHoist[] = planCfHoist(f), criterionLengths: readonly (number | undefined)[] = [],
  memoised: ReadonlyMap<AstNode, ElementValue> = new Map()): number {
  const shape = cfCallShape(f, hoisted, memoised)
  const atom = (i: number) => payload.atoms[i] ?? { cells: 0, stringUnits: 0, longestString: 0 }
  let bytes = CF_FORMULA_CALL_BYTES + CF_VIEW_WORKSPACE_BYTES + shape.nodes * CF_SLOT_BYTES + payload.viewCells * CF_VIEW_KEY_BYTES
  shape.reads.forEach((kind, i) => {
    if (kind === 'traversal') bytes += CF_TRAVERSAL_BYTES + atom(i).cells * (CF_POPULATED_SLOT_BYTES + CF_FLAT_ITEM_BYTES)
    else if (kind === 'root-area') { const a = admittedAtom(shape.refs[i]!); bytes += (a.lastRow - a.firstRow + 1) * (a.lastCol - a.firstCol + 1) * CF_FLAT_ITEM_BYTES }
  })
  // Every string and character array the text operations allocate is bounded by the characters they are charged for.
  bytes += cfTextUnits(shape, payload).chars * (2 + CF_CHAR_SLOT_BYTES)
  const work = cfCallWorkBreakdown(shape, payload, criterionLengths)
  if (work !== 'unbounded') bytes += work.matcher * CF_MATCHER_CELL_BYTES
  return bytes
}

/**
 * The view-engine tree guard (Task 34 fix 2): the translated per-call tree must read exactly the areas the call's view
 * was built and priced from. Every reference node of `tree` (after hoist substitution) is resolved to its rect, and the
 * set of those rects must equal the set of `expected`; a node kind the guard cannot resolve (a name, a union, a 3-D or
 * table reference, an intersection, a spill or an implicit intersection) is a mismatch. Fail closed: on `false` the call
 * never runs.
 */
export function cfTreeReadsMatchView(tree: AstNode, expected: readonly SourceRect[]): boolean {
  const key = (sheet: string | undefined, r: SourceRect): string => `${(sheet ?? '').toLowerCase()}!${r.firstRow}:${r.lastRow}:${r.firstCol}:${r.lastCol}`
  const read = new Set<string>()
  const stack: AstNode[] = [tree]
  while (stack.length) {
    const node = stack.pop()!
    switch (node.type) {
      case 'cell': case 'range': case 'wholeCol': case 'wholeRow': {
        const atom = atomOf(node)
        if (isRefusal(atom)) return false
        read.add(key(atom.sheet, atom))
        break
      }
      case 'unary': stack.push(node.expr); break
      case 'binary': stack.push(node.left, node.right); break
      case 'call': stack.push(...node.args); break
      case 'number': case 'string': case 'boolean': case 'empty': case 'error': case 'arrayConst': break
      default: return false
    }
  }
  const want = new Set(expected.map(r => key(undefined, r)))
  if (want.size !== read.size) return false
  for (const k of read) if (!want.has(k)) return false
  return true
}

/* ------------------------------------------------------------------ */
/* One bounded per-cell evaluation (Rulings S5-D, S5-E, S5-N).         */
/* ------------------------------------------------------------------ */

/** A charged per-preparation memo of invariant formulas' results (S5-N caps table "Invariant results"). */
export interface CfInvariantMemo {
  get(f: CfFormula): CfResult | undefined; set(f: CfFormula, result: CfResult): void
  /** Invariant subtrees (finding 8): a node's scalar value, memoised after the first clean call that computed it. */
  getNode(node: AstNode): ElementValue | undefined; setNode(node: AstNode, value: ElementValue): void
  dispose(): void
}
export function createCfInvariantMemo(memory: CfMemory): CfInvariantMemo {
  const entries = new Map<CfFormula, { result: CfResult; release(): void }>()
  const nodes = new Map<AstNode, { value: ElementValue; release(): void }>()
  return {
    get: f => entries.get(f)?.result,
    getNode: node => nodes.get(node)?.value,
    setNode(node, value) {
      if (nodes.has(node)) return
      const bytes = measurePayloadBytes(value)
      if (bytes === undefined) return
      // Charged on the preparation's memory, owned by this memo (disposed with the preparation).
      const hold = memory.reserve(CF_SLOT_BYTES + bytes)
      if (hold.kind !== 'reserved') return // not memoised: later calls evaluate the subtree again, within their bound
      nodes.set(node, { value, release: () => hold.release() })
    },
    set(f, result) {
      if (entries.has(f)) return
      const bytes = measurePayloadBytes(result)
      if (bytes === undefined) return
      const hold = memory.reserve(CF_SLOT_BYTES + bytes)
      if (hold.kind !== 'reserved') return // not memoised: the next cell evaluates again, within its own bound
      entries.set(f, { result, release: () => hold.release() })
    },
    dispose() { for (const e of entries.values()) e.release(); entries.clear(); for (const e of nodes.values()) e.release(); nodes.clear() },
  }
}

const NO_HOISTS: readonly CfHoist[] = Object.freeze([])
/** Calls whose result may be one of their operands as read (a reference's value keeps its `fromRef` origin): never a
 * memo root, since a literal in their place could read differently. */
const PASS_THROUGH: ReadonlySet<string> = new Set(['IF', 'IFERROR', 'IFNA'])
/**
 * The maximal invariant subtrees of a formula below its root (finding 8, S5-N): the rule behind `f.invariant` applied per
 * subtree (every reference `$`-pinned on both axes, no zero-argument ROW/COLUMN; compile already refuses a multi-cell
 * reference in a scalar position, so no implicit intersection is involved), holding a reference or a call, rooted at an
 * operator or at a call that computes its value (not a pass-through), and containing no hoisted node (its node identity
 * must survive translation unchanged).
 */
const subtreeCache = new WeakMap<CfFormula, { hoisted: readonly CfHoist[]; out: readonly AstNode[] }>()
export function cfInvariantSubtrees(f: CfFormula, hoisted: readonly CfHoist[] = []): readonly AstNode[] {
  // Memoised per formula and plan (a preparation passes one plan array per rule throughout).
  const cached = subtreeCache.get(f)
  if (cached && cached.hoisted === hoisted) return cached.out
  const out = invariantSubtreesOf(f, hoisted)
  subtreeCache.set(f, { hoisted, out })
  return out
}
function invariantSubtreesOf(f: CfFormula, hoisted: readonly CfHoist[]): readonly AstNode[] {
  const hoistNodes = new Set<AstNode>(hoisted.map(h => h.node))
  const info = new Map<AstNode, { invariant: boolean; work: boolean; hoist: boolean }>()
  const facts = (node: AstNode): { invariant: boolean; work: boolean; hoist: boolean } => {
    const known = info.get(node)
    if (known) return known
    let out: { invariant: boolean; work: boolean; hoist: boolean }
    if (hoistNodes.has(node)) out = { invariant: false, work: true, hoist: true }
    else switch (node.type) {
      case 'cell': out = { invariant: node.ref.absRow && node.ref.absCol && node.ref.sheet === undefined, work: true, hoist: false }; break
      case 'range': case 'wholeCol': case 'wholeRow': { const a = admittedAtom(node); out = { invariant: a.absRow && a.absCol && a.sheet === undefined, work: true, hoist: false }; break }
      case 'unary': { const e = facts(node.expr); out = { ...e }; break }
      case 'binary': { const l = facts(node.left), r = facts(node.right); out = { invariant: l.invariant && r.invariant, work: l.work || r.work, hoist: l.hoist || r.hoist }; break }
      case 'call': {
        const name = node.name.toUpperCase()
        const args = node.args.map(facts)
        out = { invariant: args.every(a => a.invariant) && !((name === 'ROW' || name === 'COLUMN') && node.args.length === 0), work: true, hoist: args.some(a => a.hoist) }
        break
      }
      default: out = { invariant: true, work: false, hoist: false }
    }
    info.set(node, out)
    return out
  }
  const eligible = (node: AstNode): boolean => {
    const x = facts(node)
    if (!x.invariant || !x.work || x.hoist) return false
    if (node.type === 'unary' || node.type === 'binary') return true
    return node.type === 'call' && node.args.length > 0 && !PASS_THROUGH.has(node.name.toUpperCase())
  }
  const out: AstNode[] = []
  const stack: AstNode[] = []
  const children = (node: AstNode): readonly AstNode[] => node.type === 'unary' ? [node.expr] : node.type === 'binary' ? [node.left, node.right] : node.type === 'call' ? node.args : []
  stack.push(...children(f.ast))
  while (stack.length) {
    const node = stack.pop()!
    if (eligible(node)) { out.push(node); continue }
    stack.push(...children(node))
  }
  return out
}
/** Pacing between calls: a call never yields inside the engine, so it records the budget's `'yield'` for the loop. */
export interface CfFormulaPacing { yieldPending: boolean }
/** Test-only observation point inside a read (simulates an engine-internal recording the grammar cannot reach). */
export type CfFormulaProbe = (ctx: EvaluationContext) => void
export type CfFormulaContext = Readonly<{
  sheetName: string; evaluation: CfEvaluationContext; semantics: ResolvedSemantics; edge: CfEdgeMode | undefined
  materialization: MaterializationBudget; budget: CfWorkBudget; memory: CfMemory; invariants: CfInvariantMemo
  /** The preparation ceiling every call draws its static bound from (non-final refusal). */
  ceiling: CfUnitCeiling
  /** The rule's hoist plan and hoisted values; absent, nothing is hoisted (a hoistable call then pays per cell). */
  hoists?: Readonly<{ plan: readonly CfHoist[]; values: CfHoisted }>
  pacing?: CfFormulaPacing
  probe?: CfFormulaProbe
  /** Observes each admitted call's static breakdown (tests, calibration). */
  onCall?: (breakdown: CfCallWorkBreakdown) => void
  /** Observes the store work a call actually metered (view build, reads, enumeration, sort, lookups) against its bound. */
  onSettled?: (spent: number, breakdown: CfCallWorkBreakdown) => void
}>
/** The semantics a CF formula reads: the document's date system and zone, and this preparation's clock (S5-M). */
export function cfFormulaSemantics(evaluation: CfEvaluationContext): ResolvedSemantics {
  return Object.freeze({ dateSystem: evaluation.dateSystem, timeZone: evaluation.timeZone, epochNowMs: evaluation.clock.epochNowMs, locale: 'en-US', unicode: Object.freeze({ version: 2 as const, source: 'standalone-default' as const }) })
}
const literalNode = (v: ElementValue): AstNode => {
  if (typeof v === 'number') return { type: 'number', value: v }
  if (typeof v === 'string') return { type: 'string', value: v }
  if (typeof v === 'boolean') return { type: 'boolean', value: v }
  if (v === null) return { type: 'empty' }
  return { type: 'error', error: v.code }
}
/** The criterion's wildcard token count the bound needs (0 for a non-text criterion), or undefined when unknown. */
function criterionTokens(value: EvaluationValue | undefined): number | undefined | 'long' {
  if (value === undefined) return undefined
  if (typeof value !== 'string') return 0
  if (value.length > CONDITIONAL_MAX_CRITERION_LENGTH) return 'long'
  // `criterionOf` (functions/cond.ts): a comparison prefix is stripped, and a numeric body is a number criterion.
  const body = value.replace(ENGINE_COMPARISON_PREFIX, '')
  if (ENGINE_NUMERIC_CRITERION.test(body)) return 0
  const toks = tokenizeWildcardPattern(body)
  return Array.isArray(toks) ? toks.length : body.length
}

/**
 * Evaluate one compiled formula at one target over the snapshot (Rulings S5-C, S5-D, S5-E, S5-N). Never through
 * `evaluateFormulaInternal`: the root is classified first (a root cell reads the snapshot directly and keeps a blank as
 * blank; a root range is a range result built from the snapshot, never the engine's scalar projection), and any other
 * root runs `evaluateNode` with a fresh context over the call's restricted view. Before invocation the call's static
 * bound must fit both the 128-unit per-call cap and the preparation ceiling, and its scratch must be reserved; a refusal
 * means the call never runs. The result is discarded as unresolved on any recorded unsupported feature or issue, a cycle
 * flag, a work refusal or a taint, even when `IFERROR`/`IFNA` masked it.
 */
export function evaluateCfFormula(f: CfFormula, at: Readonly<{ row: number; col: number }>, snapshot: CfDependencySnapshot, context: CfFormulaContext): CfOutcome<CfResult> {
  if (f.dependencySheets.length > 0) return unresolved('cf-formula-dependency') // cross-sheet sources: Task 36 (S5-K)
  if (f.usesClock && context.evaluation.timeZone === '') return unresolved('cf-evaluation-context-missing')
  const anchor = f.anchor
  if (!anchor) return unresolved('cf-unmeasured-capability') // multi-range anchor (S5-C)
  const cached = f.invariant ? context.invariants.get(f) : undefined
  if (cached) return { kind: 'ready', value: cached }
  // A fresh evaluation-local taint scope before any read (the view, criterion reads, frequency lookups, the call):
  // a taint left by an earlier call on the shared snapshot never reaches this one, and this call's own preliminary
  // reads stay in scope through invocation.
  snapshot.reset()
  const dRow = at.row - anchor.row, dCol = at.col - anchor.col
  const plan = context.hoists?.plan ?? NO_HOISTS
  // Invariant subtrees already computed in this preparation enter as literals (finding 8).
  const subtrees = cfInvariantSubtrees(f, context.hoists?.plan ?? NO_HOISTS)
  const memoised = new Map<AstNode, ElementValue>()
  for (const node of subtrees) { const v = context.invariants.getNode(node); if (v !== undefined) memoised.set(node, v) }
  const shape = cfCallShape(f, plan, memoised)
  // Per-call meter: the dynamic spends (view build, reads, enumeration, sort) are refused past this call's own bound.
  // Discovery (view build, index steps, criterion preflight) is bounded by every account it will be charged to.
  let limit = Math.min(CF_LIMITS.checkpointUnits, context.ceiling.remaining, context.budget.remaining), used = 0, admitted = false
  const meter: CfWorkBudget = {
    get units() { return used },
    get remaining() { return limit - used },
    spend(n) { if (!Number.isSafeInteger(n) || n < 0 || n > limit - used) return 'refused'; used += n; return 'continue' },
    resume() {},
  }
  // Effective read areas from the exact per-call shape (S5-SUMIF-SHAPE: admitted geometry is literal geometry).
  const translateRef = (node: RefNode): SourceRect | 'ref' | 'unmeasured' => {
    const a = admittedAtom(node)
    try {
      const [r0, r1] = shiftPair(a.firstRow, a.lastRow, node.type === 'wholeCol' ? 0 : dRow, a.absRow, MAX_ROW, context.edge)
      const [c0, c1] = shiftPair(a.firstCol, a.lastCol, node.type === 'wholeRow' ? 0 : dCol, a.absCol, MAX_COL, context.edge)
      if (r0 === 'ref' || r1 === 'ref' || c0 === 'ref' || c1 === 'ref') return 'ref'
      return { firstRow: r0, lastRow: r1, firstCol: c0, lastCol: c1 }
    } catch (error) { if (error === EDGE_UNMEASURED) return 'unmeasured'; throw error }
  }
  const areas: Array<SourceRect | 'ref'> = []
  for (const node of shape.refs) {
    const area = translateRef(node)
    if (area === 'unmeasured') return unresolved('cf-formula-edge-unmeasured')
    areas.push(area)
  }
  // Frequency lookups, scalar hoists and memoised subtrees become literal nodes in the per-call tree (no handler runs).
  const substitutes = new Map<AstNode, AstNode>()
  for (const [node, value] of memoised) substitutes.set(node, literalNode(value))
  const live: number[] = [] // indexes of refs with an on-sheet area, in shape order
  areas.forEach((a, i) => { if (a !== 'ref') live.push(i) })
  // The nodes alone past the cap refuse before any view work (the area dedupe is quadratic in references).
  if (shape.nodes > CF_LIMITS.checkpointUnits) return unresolved('cf-resource-limit')
  // The per-call bookkeeping, the view's workspace included, is reserved BEFORE the view is built (finding 1); it grows
  // to the full scratch bound once the payload is measured.
  const hold = context.memory.reserve(CF_FORMULA_CALL_BYTES + CF_VIEW_WORKSPACE_BYTES)
  if (hold.kind !== 'reserved') return unresolved('cf-resource-limit')
  try {
  const view = snapshot.viewFor(live.map(i => ({ rect: areas[i] as SourceRect })), meter, hold)
  if (view.kind !== 'ready') return view.kind === 'unresolved' ? view : unresolved('cf-formula-dependency')
  const metrics: CfPayloadMetrics = { viewCells: view.value.metrics.viewCells, index: view.value.metrics.index ?? 0, atoms: shape.refs.map((_, i) => {
    const j = live.indexOf(i)
    return j < 0 ? { cells: 0, stringUnits: 0, longestString: 0 } : view.value.metrics.atoms[j]!
  }) }
  const readAt = (i: number): ElementValue | 'ref' => {
    const a = areas[i]!
    return a === 'ref' ? 'ref' : view.value.read(a.firstRow, a.firstCol)
  }
  for (const hoist of plan) {
    if (hoist.kind === 'scalar') {
      const value = context.hoists!.values.scalars.get(hoist.node)
      if (value === undefined) return unresolved('cf-resource-limit') // the hoist was refused: never paid per cell
      substitutes.set(hoist.node, literalNode(value))
    }
  }
  // Criterion token counts for the bound, read from literals and the snapshot (uncharged here: the engine's own read of
  // the same cell is charged).
  const criterionLengths: Array<number | undefined> = []
  for (const c of shape.conditionals) for (const p of c.pairs) {
    const node = p.criterion
    let value: EvaluationValue | undefined
    if (node.type === 'number' || node.type === 'boolean') value = 0
    else if (node.type === 'string') value = node.value
    else if (node.type === 'cell') {
      const i = shape.refs.indexOf(node)
      const a = i < 0 ? 'ref' : areas[i]!
      value = a === 'ref' ? null : snapshot.read(undefined, a.firstRow, a.firstCol)
    }
    // The criterion's parse (prefix strip, numeric pattern, wildcard tokens) is charged on the meter BEFORE it runs, at
    // the units the bound's text term counts for it (round 3, finding 2): a long criterion is refused unparsed.
    if (typeof value === 'string') {
      if (value.length > CONDITIONAL_MAX_CRITERION_LENGTH) return unresolved('cf-formula-unsupported') // conditional-long-criterion-unverified
      if (meter.spend(cfCriterionUnits(value.length)) === 'refused') return unresolved('cf-resource-limit')
    }
    const t = criterionTokens(value)
    if (t === 'long') return unresolved('cf-formula-unsupported') // the engine records conditional-long-criterion-unverified
    criterionLengths.push(t)
  }
  const breakdown = cfCallWorkBreakdown(shape, metrics, criterionLengths)
  // Preflight against every account the admission draws from: a call that does not fit refuses without spending, so a
  // latching budget is never poisoned for the calls that follow.
  if (breakdown === 'unbounded' || breakdown.total > CF_LIMITS.checkpointUnits || breakdown.total > context.ceiling.remaining
    || breakdown.total > context.budget.remaining) return unresolved('cf-resource-limit')
  const scratch = cfEvaluatorScratchBytes(f, metrics, plan, criterionLengths, memoised)
  if (scratch > CF_LIMITS.formulaScratchBytes) return unresolved('cf-resource-limit')
  if (!hold.grow(scratch)) return unresolved('cf-resource-limit')
  {
    // Frequency lookups: the criterion read and the lookup are within the bound (one unit each).
    for (const hoist of plan) {
      if (hoist.kind !== 'countif-frequency') continue
      const i = shape.refs.indexOf(hoist.criterion as RefNode)
      const raw = i < 0 ? 'ref' : readAt(i)
      if (snapshot.tainted) return unresolved('cf-formula-dependency')
      const literal = raw === 'ref' ? undefined : cfElementToLiteral(raw)
      // The key's processing (ASCII, wildcard and numeric checks, the fold, the hash) is spent on the meter BEFORE it
      // runs (round 4, finding 5), at the units the bound's criterion term counts for it: a refused key keeps the charge
      // (the finally below), an admitted call's bound already includes it once.
      if (literal?.kind === 'string' && meter.spend(cfCriterionUnits(literal.value.length, CF_FREQUENCY_KEY_PASSES)) === 'refused') return unresolved('cf-resource-limit')
      const key = literal ? cfCountifKey(literal) : 'refuse'
      if (key === 'refuse' || meter.spend(1) === 'refused') return unresolved(key === 'refuse' ? 'cf-formula-unsupported' : 'cf-resource-limit')
      if (cfCountifDpRefused(literal!, context.hoists!.values.longest.get(hoist.node) ?? 0)) return unresolved('cf-formula-unsupported')
      const counts = context.hoists!.values.frequencies.get(hoist.node)
      if (!counts) return unresolved('cf-resource-limit')
      substitutes.set(hoist.node, { type: 'number', value: counts.get(key) ?? 0 })
    }
    const tree = translateCfFormula(f, dRow, dCol, context.edge, context.memory, node => substitutes.get(node))
    if (tree.kind !== 'ready') return tree.kind === 'unresolved' ? tree : unresolved('cf-resource-limit')
    try {
      // The tree the engine walks must read exactly the view's areas, except the frequency criteria read above (their
      // calls are already literal nodes). A mismatch never runs: the view would not cover, or not price, what it reads.
      const lookupRefs = new Set(plan.filter(h => h.kind === 'countif-frequency').map(h => shape.refs.indexOf(h.criterion as RefNode)))
      const engineAreas = areas.filter((a, i): a is SourceRect => a !== 'ref' && !lookupRefs.has(i))
      if (!cfTreeReadsMatchView(tree.value.value, engineAreas)) return unresolved('cf-formula-dependency')
      // Admission: the whole static bound is charged before the call, from the pacing budget then the ceiling.
      const spent = context.budget.spend(breakdown.total)
      if (spent === 'refused') return unresolved('cf-resource-limit')
      if (spent === 'yield' && context.pacing) context.pacing.yieldPending = true
      if (!context.ceiling.take(breakdown.total)) return unresolved('cf-resource-limit')
      admitted = true // the bound charged covers the discovery already metered
      limit = breakdown.total
      context.onCall?.(breakdown)
      const outcome = invoke(f, tree.value.value, at, shape, areas, view.value, snapshot, context, stores => {
        // A clean call: memoise each invariant subtree it computed, in whichever mode the engine evaluated it (a scalar
        // value; a blank is skipped, since an empty literal could read as an omitted argument; a node computed in both
        // modes to different values is skipped). Its node is shared by every translation, so the value the engine
        // stored for it is the subtree's own, and the literal evaluates to the same value in either mode.
        for (const node of subtrees) {
          if (memoised.has(node)) continue
          const inScalar = stores.scalar.nodeValues.has(node), inArray = stores.array.nodeValues.has(node)
          const sv = stores.scalar.nodeValues.get(node), av = stores.array.nodeValues.get(node)
          if (inScalar && inArray && !Object.is(sv, av) && !(isEvaluationError(sv) && isEvaluationError(av) && sv.code === av.code)) continue
          const v = inScalar ? sv : av
          if (v === undefined || v === null || isMatrixValue(v) || (typeof v === 'number' && !Number.isFinite(v))) continue
          if (typeof v === 'object' && !isEvaluationError(v)) continue
          context.invariants.setNode(node, v as ElementValue)
        }
      })
      context.onSettled?.(used, breakdown)
      if (outcome.kind === 'ready' && f.invariant) context.invariants.set(f, outcome.value)
      return outcome
    } finally { tree.value.release() }
  }
  } finally {
    hold.release()
    // A call refused after discovery still pays for it (round 3, finding 1): the metered units go to the shared budget
    // and the ceiling, and a yield they earn is honoured between cells.
    if (!admitted && used > 0) {
      if (context.budget.spend(used) === 'yield' && context.pacing) context.pacing.yieldPending = true
      context.ceiling.take(Math.min(used, context.ceiling.remaining))
    }
  }
}

function invoke(f: CfFormula, tree: AstNode, at: Readonly<{ row: number; col: number }>, shape: CfCallShape, areas: readonly (SourceRect | 'ref')[],
  view: CfDependencyView, snapshot: CfDependencySnapshot, context: CfFormulaContext, onClean?: (stores: ModeFrameStores) => void): CfOutcome<CfResult> {
  try {
    // Shape before projection (S5-D round-1 #7): a root reference never reaches the engine's scalar projection.
    if (tree.type === 'error' && (f.ast.type === 'cell' || f.rootArea)) return { kind: 'ready', value: { kind: 'error', value: ERROR_ELEMENTS.get(tree.error) ?? ERROR_ELEMENTS.get('#REF!')! } }
    if (f.ast.type === 'cell') {
      const a = areas[0]!
      if (a === 'ref') return { kind: 'ready', value: { kind: 'error', value: ERROR_ELEMENTS.get('#REF!')! } }
      const value = view.read(a.firstRow, a.firstCol)
      return snapshot.tainted ? unresolved('cf-formula-dependency') : classifyValue(value)
    }
    if (f.rootArea) {
      const a = areas[0]!
      if (a === 'ref') return { kind: 'ready', value: { kind: 'error', value: ERROR_ELEMENTS.get('#REF!')! } }
      const values: ElementValue[][] = []
      for (let r = a.firstRow; r <= a.lastRow; r++) { const row: ElementValue[] = []; for (let c = a.firstCol; c <= a.lastCol; c++) row.push(view.read(r, c)); values.push(row) }
      if (snapshot.tainted) return unresolved('cf-formula-dependency')
      return { kind: 'ready', value: { kind: 'range', matrix: { kind: 'formula-matrix', rows: values.length, cols: values[0]!.length, values }, area: a } }
    }
    void shape
    const lower = context.sheetName.toLowerCase()
    const references = createReferenceServices({
      generation: () => 0, sheets: [{ sheetId: snapshot.sheetId, name: context.sheetName, workbookIndex: 0 }],
      sheetIdOfName: name => (name.toLowerCase() === lower ? snapshot.sheetId : undefined),
      sheetNameOfId: id => (id === snapshot.sheetId ? context.sheetName : undefined),
      store: view.store, definedNames: [], tables: undefined,
    })
    const features = new Set<string>()
    const issues: ReferenceIssue[] = []
    // Caller-owned stores for both evaluation modes (round 3, finding 5): an argument the engine evaluates in array mode
    // (SUM, AND, ...) leaves its value in the array store, where the subtree memo can read it.
    const stores: ModeFrameStores = {
      scalar: { nodeValues: new WeakMap(), flatArgs: new WeakMap(), functionWork: new WeakMap() },
      array: { nodeValues: new WeakMap(), flatArgs: new WeakMap(), functionWork: new WeakMap() },
    }
    const currentCell: CellRef = { row: at.row, col: at.col, absRow: false, absCol: false }
    const ctx: EvaluationContext = {
      typedValues: true, currentSheet: context.sheetName, currentCell, semantics: context.semantics, references,
      unsupportedFeatures: features, reportFormulaIssue: issue => { issues.push(issue) }, hasCycle: false,
      mode: 'scalar', modeStores: stores, nodeValues: stores.scalar.nodeValues, flatArgs: stores.scalar.flatArgs, functionWork: stores.scalar.functionWork,
      materializationBudget: context.materialization,
      // Engine argument order is (sheet, col, row); a single-cell read goes through the call's view (charged, area-checked).
      getCellValue: (sheet, col, row) => {
        if (sheet !== undefined && sheet.toLowerCase() !== lower) { features.add('cf-cross-sheet-read'); return ERROR_ELEMENTS.get('#N/A')! }
        const value = view.read(row, col)
        context.probe?.(ctx)
        return value
      },
    }
    const value = evaluateNode(tree, ctx)
    if (snapshot.tainted) return unresolved('cf-formula-dependency')
    if (features.has('array-resource-limit')) return unresolved('cf-resource-limit')
    if (features.size > 0 || issues.length > 0 || ctx.hasCycle) return unresolved('cf-formula-unsupported')
    const classified = classifyValue(value)
    if (classified.kind === 'ready') onClean?.(stores)
    return classified
  } catch (error) {
    if (error === CF_WORK_REFUSED) return unresolved('cf-resource-limit')
    return unresolved('cf-formula-unsupported') // an engine failure is never a result
  }
}
