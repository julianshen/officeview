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
import type { AstNode } from '../formula/types'
import { parseFormula } from '../formula/parser'
import { FUNCTIONS } from '../formula/functions'
import { epochToSerial } from '../formula/serial'
import type { XlsxDocument } from '../types'
import type { ValuePolicy } from '../view/contracts'
import { CF_FORMULA_NODE_BYTES, CF_LIMITS, type CfMemory } from './limits'
import { measurePayloadBytes } from '../view/payloadBytes'
import type { CfFormula, CfFormulaRef, CfLease, CfRuleType, CfTruthKind } from './model'

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

/**
 * Compile with the caller owning the reservation. The reservation for the worst case (`tokens x CF_FORMULA_NODE_BYTES`)
 * is requested BEFORE `parseFormula` runs, so a refused reservation constructs no AST; after parsing the exact measured
 * bytes are requested too (the token count does not bound the string payload or the node count of a range). Returns the
 * formula, the bytes the caller now holds for it (`bytes`) and its exact measured size (`retained`).
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
  const worst = tokens * CF_FORMULA_NODE_BYTES
  if (!reserve(worst)) return limited('compiled formula reservation refused')
  let ast: AstNode
  try {
    ast = parseFormula(source)
  } catch {
    return limited('formula could not be parsed within the engine stack')
  }
  const result = admit(ast, anchor)
  if (isRefusal(result)) return result
  const measured = measurePayloadBytes(result)
  if (measured === undefined) return limited('compiled formula could not be measured')
  const bytes = Math.max(worst, measured)
  if (bytes > worst && !reserve(bytes)) return limited('compiled formula exceeds its reservation')
  return { formula: result, bytes, retained: measured }
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
    return hold.adopt(compiled.formula, compiled.bytes)
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
      } else if (rule.sourceType !== 'expression') evaluation.add('formula:predicate-divergent')
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
  return Object.freeze({
    calcMode: doc.calc?.calcMode ?? 'auto',
    // `workbookValuePolicy` (formula/value-policy.ts) restated: only a cached document carries the flag.
    valuePolicy: doc.valuePolicy === 'cached' ? 'cached' as const : 'engine' as const,
    dateSystem: doc.semantics?.dateSystem ?? '1900',
    timeZone: doc.semantics?.timeZone ?? '',
  })
}
/** The worker twin: a paged workbook is always under the `'cached'` value policy (the adapters' policy). The zone is the
 * same ambient zone the paged CF date context takes. */
export function cfDocumentEvaluationOfPaged(workbook: Readonly<{ calcMode?: 'auto' | 'manual' | 'autoNoTable'; formatting?: Readonly<{ dateSystem?: '1900' | '1904' }> }>): CfDocumentEvaluation {
  let timeZone = 'UTC'
  try {
    const resolved = Intl.DateTimeFormat().resolvedOptions().timeZone
    if (typeof resolved === 'string' && resolved !== '') timeZone = resolved
  } catch { /* keep UTC */ }
  return Object.freeze({ calcMode: workbook.calcMode ?? 'auto', valuePolicy: 'cached' as const, dateSystem: workbook.formatting?.dateSystem ?? '1900', timeZone })
}
/** One clock snapshot from the injected `now`; `clockGen` is the local day serial of that sample (the interactive
 * midnight tick keys on it). A zone the engine cannot resolve falls back to the UTC day so the generation still advances. */
export function cfEvaluationContextFor(policy: CfDocumentEvaluation, now: () => number): CfEvaluationContext {
  const epochNowMs = now()
  const serial = epochToSerial(epochNowMs, policy.timeZone, policy.dateSystem)
  const clockGen = typeof serial === 'number' ? Math.floor(serial) : Number.isFinite(epochNowMs) ? Math.floor(epochNowMs / 86_400_000) : 0
  return Object.freeze({ ...policy, clock: Object.freeze({ epochNowMs, clockGen }) })
}
