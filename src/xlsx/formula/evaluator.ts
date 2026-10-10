import type { ArgumentProgress, AstNode, BinaryOp, ElementValue, EvaluationContext, EvaluationMode, FormulaError, FormulaValue, EvaluationError, EvaluationValue, MatrixValue, ReferenceNode, ResolvedRef } from './types'
import { parseFormula } from './parser'
import { FUNCTIONS } from './functions'
import { scalarOfImplicitIntersect, scalarOfReference, scalarOfResolved, syntacticAreas } from './refs'
import { installActiveNameScope } from './name-state'
import { translateAst } from './shared'
import { registerRelativeNodeGate, relativeNodeGate } from './ref-discovery'
import { inheritOwnedMemoScope, installOwnedMemoScope, isReferenceGenerationChanged, MAX_REFERENCE_GENERATION_RESTARTS, requestOwnedMemoRestart } from './reference-generation'
import type { NameBinding } from './types'

class NodeMemoSuspension {
  constructor(readonly node: AstNode) {}
}

/** A function handler requesting that a child be evaluated in the OTHER
 * explicit mode using that mode's store. Caught only by the requesting call
 * frame, which schedules the child on the SAME heap/frame engine. */
class NodeModeSuspension {
  constructor(readonly node: AstNode, readonly mode: EvaluationMode) {}
}

/** Group 3 private adapter state, keyed by the existing nodeValues memo
 * ownership (survives workbook replay); no ABI change. */
const bindingsByMemo = new WeakMap<
  WeakMap<AstNode, EvaluationValue>,
  WeakMap<AstNode, NameBinding | EvaluationError>
>()

/** Stable evaluated-body clones follow the existing per-cell memo lifetime.
 * Original bound SOURCE ASTs remain the active-name/cycle identity. */
const valueNameSourcesByMemo = new WeakMap<WeakMap<AstNode, EvaluationValue>, WeakMap<AstNode, Map<string, AstNode>>>()

function relativeAtom(node: AstNode): boolean {
  switch (node.type) {
    case 'cell': return !node.ref.absCol || !node.ref.absRow
    case 'range': return !node.ref.from.absCol || !node.ref.from.absRow || !node.ref.to.absCol || !node.ref.to.absRow
    case 'wholeCol': case 'wholeRow': return !node.ref.from.absolute || !node.ref.to.absolute
    case 'ref3d': return relativeAtom(node.target)
    case 'spill': return !node.anchor.absCol || !node.anchor.absRow
    default: return false
  }
}

function visitSourceNodes(root: AstNode, visit: (node: AstNode) => void): void {
  const pending = [root]
  while (pending.length > 0) {
    const node = pending.pop()!
    visit(node)
    switch (node.type) {
      case 'call': for (let i = node.args.length - 1; i >= 0; i--) pending.push(node.args[i]); break
      case 'union': for (let i = node.refs.length - 1; i >= 0; i--) pending.push(node.refs[i]); break
      case 'binary': case 'intersect': pending.push(node.right); pending.push(node.left); break
      case 'unary': case 'implicitIntersect': pending.push(node.expr); break
    }
  }
}

function valueNameSource(binding: NameBinding, ctx: EvaluationContext, memo: WeakMap<AstNode, EvaluationValue>): AstNode {
  let relative = false
  visitSourceNodes(binding.ast, node => { if (relativeAtom(node)) relative = true })
  if (!relative) return binding.ast
  let sources = valueNameSourcesByMemo.get(memo)
  if (!sources) { sources = new WeakMap(); valueNameSourcesByMemo.set(memo, sources) }
  let copies = sources.get(binding.ast)
  if (!copies) { copies = new Map(); sources.set(binding.ast, copies) }
  const useSite = ctx.currentAddress ?? ctx.currentCell
  const key = `${ctx.currentAddress?.sheetId ?? ctx.currentSheet ?? ''}:${useSite?.col}:${useSite?.row}`
  const cached = copies.get(key)
  if (cached) return cached
  const base = binding.relativeBase
  let source: AstNode
  if (base && binding.baseProvenance !== 'unknown' && useSite) {
    source = translateAst(binding.ast, useSite.col - base.col, useSite.row - base.row)
  } else {
    // Clone without shifting, then attach capability metadata to relative
    // atoms only. Reference unions keep their order and original name nodes,
    // so an earlier active binding still wins over a later unavailable base.
    source = translateAst(binding.ast, 0, 0)
    const feature = base && binding.baseProvenance !== 'unknown' ? 'relative-name-context-missing' : 'relative-name'
    const message = feature === 'relative-name'
      ? `Relative defined name "${binding.name}" has unknown base provenance; offsets not guessed`
      : `Relative defined name "${binding.name}" needs a supplied use-site address`
    visitSourceNodes(source, node => { if (relativeAtom(node)) registerRelativeNodeGate(node, { feature, message }) })
  }
  copies.set(key, source)
  return source
}

export const CANONICAL_ERRORS = new Set<FormulaError>([
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#NULL!',
  '#SPILL!',
  '#CALC!',
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
  const scalar = scalarProjection(value)
  return isEvaluationError(scalar) ? scalar.code : scalar
}

/** Internal matrix predicate; matrices never escape into XlsxCell.value. */
export function isMatrixValue(value: unknown): value is MatrixValue {
  return typeof value === 'object' && value !== null && (value as { kind?: string }).kind === 'formula-matrix'
}

/** Scalar projection: a matrix projects to its nonempty top-left element. */
export function topLeftOfMatrix(matrix: MatrixValue): ElementValue {
  return matrix.values[0]?.[0] ?? null
}

export function scalarProjection(value: EvaluationValue): ElementValue {
  return isMatrixValue(value) ? topLeftOfMatrix(value) : value
}

export function matrixOf(rows: number, cols: number, values: readonly (readonly ElementValue[])[]): MatrixValue {
  return { kind: 'formula-matrix', rows, cols, values }
}

export function matrixFromArrayConst(node: Extract<AstNode, { type: 'arrayConst' }>): MatrixValue {
  return matrixOf(node.rows.length, node.rows[0].length, node.rows as readonly (readonly ElementValue[])[])
}

const DEFAULT_ARRAY_CELLS = 1048576
const DEFAULT_ARRAY_SIDE = 1048576

function arrayPolicy(ctx: EvaluationContext | undefined): { maxCells: number; maxSide: number } {
  return {
    maxCells: ctx?.maxArrayCells ?? DEFAULT_ARRAY_CELLS,
    maxSide: ctx?.maxArraySide ?? DEFAULT_ARRAY_SIDE,
  }
}

function arrayShapeGate(ctx: EvaluationContext | undefined, feature: string, message: string): EvaluationError {
  markUnsupported(ctx, feature)
  ctx?.reportFormulaIssue?.({ kind: 'unsupported-reference', feature, message })
  return formulaError('#NAME?')
}

function arrayResourceGate(ctx: EvaluationContext | undefined, rows: number, cols: number): EvaluationError {
  const { maxCells, maxSide } = arrayPolicy(ctx)
  markUnsupported(ctx, 'array-resource-limit')
  ctx?.reportFormulaIssue?.({
    kind: 'unsupported-reference',
    feature: 'array-resource-limit',
    message: `Array-mode reference ${rows}x${cols} exceeds the application policy (maxArrayCells ${maxCells}, maxArraySide ${maxSide}); resource policy, not a native shape error`,
  })
  return formulaError('#NAME?')
}

function boundedRectangle(rows: number, cols: number, ctx: EvaluationContext | undefined): EvaluationError | undefined {
  const { maxCells, maxSide } = arrayPolicy(ctx)
  if (!Number.isSafeInteger(rows) || !Number.isSafeInteger(cols) || rows < 1 || cols < 1) {
    return arrayShapeGate(ctx, 'array-shape-unavailable', `Array-mode reference ${rows}x${cols} is not a nonempty rectangle`)
  }
  if (rows > maxSide || cols > maxSide || rows * cols > maxCells) return arrayResourceGate(ctx, rows, cols)
  return undefined
}

/**
 * Array-mode materialization of a resolved reference: the full prepared,
 * bounded, rectangular area read positionally. Multi-area unions/3D runs are an
 * honest shape gate (a union has no single rectangle). `readAt` may suspend a
 * Pending workbook dependency, which propagates cooperatively to the frame.
 */
export function materializeResolved(ref: ResolvedRef, ctx: EvaluationContext | undefined): EvaluationValue {
  if (ref.areas.length !== 1) {
    return arrayShapeGate(ctx, 'array-reference-shape-unavailable', `Array mode needs exactly one rectangular reference area; this reference has ${ref.areas.length}`)
  }
  const area = ref.areas[0]
  const bad = boundedRectangle(area.rows, area.cols, ctx)
  if (bad) return bad
  const values: ElementValue[][] = []
  try {
    for (let row = 0; row < area.rows; row++) {
      const out: ElementValue[] = []
      for (let col = 0; col < area.cols; col++) out.push(ref.readAt(0, row, col))
      values.push(out)
    }
  } catch (error) {
    // A mid-materialization footprint change must restart the WHOLE array
    // evaluation, not leak the private ReferenceGenerationChanged error. The
    // owned scope (API-created array stores) turns this into the private
    // restart token; caller-owned stores keep the original error identity.
    if (!isReferenceGenerationChanged(error)) throw error
    requestOwnedMemoRestart(ctx)
    throw error
  }
  return matrixOf(area.rows, area.cols, values)
}

/** Standalone array-mode materialization through the legacy cell callback. */
function legacyMaterializeReference(node: ReferenceNode, ctx: EvaluationContext | undefined): EvaluationValue {
  const areas = syntacticAreas(node, ctx?.currentSheet)
  if (isEvaluationError(areas)) return areas
  if (areas.length !== 1) {
    return arrayShapeGate(ctx, 'array-reference-shape-unavailable', `Array mode needs exactly one rectangular reference area; this reference has ${areas.length}`)
  }
  const area = areas[0]
  const bad = boundedRectangle(area.rows, area.cols, ctx)
  if (bad) return bad
  const sheet = area.sheetId === '' ? undefined : area.sheetId
  const values: ElementValue[][] = []
  for (let row = 0; row < area.rows; row++) {
    const out: ElementValue[] = []
    for (let col = 0; col < area.cols; col++) {
      const value = ctx?.getCellValue?.(sheet, area.firstCol + col, area.firstRow + row)
      out.push(value === undefined ? null : scalarProjection(contextValue(value, ctx)))
    }
    values.push(out)
  }
  return matrixOf(area.rows, area.cols, values)
}

/** Mode-aware grid-reference value: array mode keeps full geometry; scalar mode
 * keeps the existing implicit-intersection context projection. */
function referenceArrayValue(
  node: ReferenceNode,
  ctx: EvaluationContext | undefined,
  evalNode: (node: AstNode, ctx?: EvaluationContext) => EvaluationValue,
): EvaluationValue {
  const services = ctx?.references
  if (services && ctx) {
    const resolved = services.resolve(node, ctx)
    if (resolved === undefined) {
      if (node.type !== 'name') return formulaError('#VALUE!')
      return evalNode(node, ctx)
    }
    if (isEvaluationError(resolved)) return resolved
    services.prepare(resolved, ctx)
    return materializeResolved(resolved, ctx)
  }
  // No cell source at all: keep the existing diagnosed API-context boundary
  // rather than materializing a rectangle of invented nulls.
  if (!ctx?.getCellValue) return scalarOfReference(node, ctx, evalNode)
  return legacyMaterializeReference(node, ctx)
}

/** Normalize callback boundaries only; formula strings always remain ordinary text. */
export function contextValue(value: EvaluationValue, ctx?: EvaluationContext): EvaluationValue {
  return !ctx?.typedValues && isFormulaError(value) ? formulaError(value) : value
}

export const MAX_EVAL_DEPTH = 512

export function makeCellKey(sheet: string | undefined, col: number, row: number): string {
  // Exact sheet identity: names are matched case-insensitively (Excel sheet-name
  // rule) but NEVER trimmed — ' Data ' and 'Data' are distinct workbooks sheets.
  const s = (sheet ?? '').toLowerCase()
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

/** Native Excel decimal ROUND per root guidance: shortest-decimal significand
 * + BigInt half-away rounding, no 15-digit pre-rounding, no huge allocations.
 * d >= #fractionalDigits returns the input; very negative d rounds to 0. */
export function decimalRound(num: number, digits: number): number | EvaluationError {
  if (!Number.isFinite(num)) return formulaError('#NUM!')
  const neg = num < 0
  const abs = Math.abs(num)
  if (abs === 0) return num
  const s = String(abs)
  const eIdx = s.search(/e/i)
  let mant = s, explicitExp = 0
  if (eIdx >= 0) { mant = s.slice(0, eIdx); explicitExp = parseInt(s.slice(eIdx + 1), 10) }
  const dot = mant.indexOf('.')
  const fracDigits = dot >= 0 ? mant.length - dot - 1 : 0
  const significand = BigInt((dot >= 0 ? mant.slice(0, dot) + mant.slice(dot + 1) : mant).replace(/^0+/, '') || '0')
  const e = explicitExp - fracDigits
  const shift = e + digits
  if (shift >= 0) return num
  const places = -shift
  const len = significand.toString().replace(/^0+$/, '0').length
  if (places > len + 16) return neg ? -0 : 0
  const div = 10n ** BigInt(places)
  let q = significand / div
  const r = significand % div
  if (2n * r >= div) q += 1n
  const out = Number(`${q.toString()}e${-digits}`)
  if (!Number.isFinite(out)) return formulaError('#NUM!')
  return neg ? -out : out
}

/** Function-result guard: NO inline 15-digit normalization — native chains
 * (INT(SUM(1.999999999999999))=1) must see the raw computed value; the single
 * 15-digit normalization happens at the public/cell boundary. */
export function finiteOrNum(res: number): number | EvaluationError {
  if (!Number.isFinite(res) || Number.isNaN(res)) {
    return formulaError('#NUM!')
  }
  return res
}

/**
 * Excel numeric TEXT grammar: base-10 number with an optional percent suffix.
 * JS stops ("0x10", "0b1", "0o7", "Infinity", "", "  ") never coerce — native
 * Excel 16.106.1 returns #VALUE! for ""+1 / "  "+1 / "0x10"+1 and 1.1 for "10%"+1.
 */
const EXCEL_NUMERIC_TEXT = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?%?$/

export function coerceToNumber(val: EvaluationValue): number | EvaluationError {
  if (isEvaluationError(val)) return val
  if (val === null) return 0
  if (typeof val === 'boolean') return val ? 1 : 0
  if (typeof val === 'number') return val
  if (typeof val === 'string') {
    // Surrounding whitespace trims FIRST (native " 5"+1 / " $5 "+1 coerce);
    // after trimming the empty/allspace/hex text still rejects. Optional
    // currency prefix + thousands groups of three remain (native "1,234,567");
    // percent suffix scales. A parsed text that is NOT FINITE (overflow,
    // including overflow percentages) is #VALUE!, never #NUM! (native oracle).
    const raw = val.trim().replace(/^\$/, '').replace(/,(\d{3})(?!\d)/g, '$1')
    if (!EXCEL_NUMERIC_TEXT.test(raw)) return formulaError('#VALUE!')
    const percent = raw.endsWith('%')
    let numText = raw
    // Percent scales by 10^−2. FUSE that exponent into the DECIMAL exponent of
    // the mantissa BEFORE floating parse, so "1e309%" = 1e307 (finite — root
    // native counterexample Core-Revision4/1); a TRULY scaled-overflow text
    // ("1e400%") must still reject as #VALUE!.
    if (percent) {
      numText = raw.slice(0, -1)
      const eIdx = numText.search(/e/i)
      if (eIdx >= 0) {
        const exp = parseInt(numText.slice(eIdx + 1), 10)
        if (!Number.isFinite(exp)) return formulaError('#VALUE!')
        numText = `${numText.slice(0, eIdx)}e${exp - 2}`
      }
    }
    const num = Number(numText)
    if (Number.isNaN(num) || !Number.isFinite(num)) return formulaError('#VALUE!')
    const scaled = percent && !/e/i.test(raw) ? num / 100 : num
    return Number.isFinite(scaled) ? scaled : formulaError('#VALUE!')
  }
  return formulaError('#VALUE!')
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

export function coerceToString(val: EvaluationValue): string {
  if (val === null) return ''
  if (typeof val === 'boolean') return val ? 'TRUE' : 'FALSE'
  if (typeof val === 'number') return String(round15(val))
  return isEvaluationError(val) ? val.code : String(val)
}

function getTypeRank(val: EvaluationValue): number {
  if (typeof val === 'number') return 1
  if (typeof val === 'string') return 2
  if (typeof val === 'boolean') return 3
  return 0 // null
}

/** Excel 15-significant-digit storage TRUNCATION (no rounding) used for
 * numeric comparisons, operating on the value's SHORTEST round-trip decimal
 * (native: 1.999999999999999=2 FALSE while 1.000000000000001=1 TRUE and
 * 0.1+0.2=0.3 TRUE). */
export function trunc15(v: number): number {
  if (!Number.isFinite(v) || v === 0) return v
  const s = String(v)
  const eIdx = s.search(/e/i)
  let mant = s, exp = 0
  if (eIdx >= 0) { mant = s.slice(0, eIdx); exp = parseInt(s.slice(eIdx + 1), 10) }
  const dot = mant.indexOf('.')
  const intPart = dot >= 0 ? mant.slice(0, dot) : mant
  const fracPart = dot >= 0 ? mant.slice(dot + 1) : ''
  let digits: string
  let place0: number
  if (intPart === '0' || intPart === '-0') {
    const neg = intPart.startsWith('-')
    const frac = neg ? fracPart : fracPart
    let zeros = 0
    while (zeros < frac.length && frac[zeros] === '0') zeros++
    digits = frac.slice(zeros) || '0'
    place0 = exp - 1 - zeros
    if (neg) digits = '-' + digits
  } else {
    const stripped = intPart.replace(/^-/, '').replace(/^0+/, '') || '0'
    digits = (intPart.startsWith('-') ? '-' : '') + stripped + fracPart
    place0 = exp + stripped.length - 1
  }
  if (digits.replace(/^-/, '').length <= 15) return v
  const neg = digits.startsWith('-')
  const mag = digits.replace(/^-/, '')
  const keep = mag.slice(0, 15)
  const reassembled = parseFloat(`${keep[0]}.${keep.slice(1)}e${place0}`)
  return neg ? -reassembled : reassembled
}

export function compareValues(v1: EvaluationValue, v2: EvaluationValue): number {
  if (v1 === null && v2 === null) return 0

  let a = v1
  let b = v2
  if (a === null) {
    if (typeof b === 'number') a = 0
    else if (typeof b === 'string') a = ''
    else if (typeof b === 'boolean') a = false
  }
  if (b === null) {
    if (typeof a === 'number') b = 0
    else if (typeof a === 'string') b = ''
    else if (typeof a === 'boolean') b = false
  }

  const rankA = getTypeRank(a)
  const rankB = getTypeRank(b)

  if (rankA !== rankB) {
    return rankA < rankB ? -1 : 1
  }

  if (typeof a === 'number' && typeof b === 'number') {
    // Excel numeric comparisons run on 15-significant-digit storage values,
    // which TRUNCATE entered/computed values at the 15th digit (no rounding):
    // 1.999999999999999 keeps '1.99999999999999' -> =2 is FALSE, while
    // 1.000000000000001 truncates to 1 -> =1 is TRUE (native-verified).
    const rA = trunc15(a)
    const rB = trunc15(b)
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


/** Record an actually-evaluated unsupported feature (dead IF branches never reach us). */
export function markUnsupported(ctx: EvaluationContext | undefined, feature: string): void {
  if (!ctx) return
  const record = ctx as EvaluationContext & { unsupportedFeatures?: Set<string> }
  if (!record.unsupportedFeatures) record.unsupportedFeatures = new Set()
  record.unsupportedFeatures.add(feature)
}

/** Text iteration per the workbook text-length semantics:
 * version 1 (legacy code units) or 2 (current standalone Excel code points). */
export function textChars(text: string, ctx?: EvaluationContext): string[] {
  const version = (ctx as (EvaluationContext & { excelTextLengthVersion?: 1 | 2 }) | undefined)?.excelTextLengthVersion
  if (version === 1) return Array.from({ length: text.length }, (_, i) => text[i])
  return Array.from(text)
}

function evaluateCellRef(node: Extract<AstNode, { type: 'cell' }>, ctx: EvaluationContext | undefined): EvaluationValue {
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

/** Combine one binary operation from already-computed child values. */
/** Intermediate arithmetic keeps FULL native precision (chained formulas must
 * reproduce native float controls like (0.1+0.2)-0.3 = 5.55e-17); the 15-digit
 * Excel value normalization happens ONCE at the final public/cell boundary. */
function finiteOrRaw(res: number): number | EvaluationError {
  if (!Number.isFinite(res) || Number.isNaN(res)) {
    return formulaError('#NUM!')
  }
  return res
}

export function combineBinary(node: { op: BinaryOp }, leftVal: EvaluationValue, rightVal: EvaluationValue): EvaluationValue {
  if (isEvaluationError(leftVal)) return leftVal
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
      return finiteOrRaw(n1 + n2)

    case '-':
      return finiteOrRaw(n1 - n2)

    case '*':
      return finiteOrRaw(n1 * n2)

    case '/': {
      if (n2 === 0) return formulaError('#DIV/0!')
      return finiteOrNum(n1 / n2)
    }

    case '^': {
      if (n1 === 0 && n2 === 0) return formulaError('#NUM!')
      if (n1 < 0 && !Number.isInteger(n2)) return formulaError('#NUM!')
      const res = Math.pow(n1, n2)
      return finiteOrRaw(res)
    }

    default:
      return formulaError('#VALUE!')
  }
}

/** Combine one unary operation from its computed child value. */
export function combineUnary(node: { op: '+' | '-' | '%' }, inner: EvaluationValue): EvaluationValue {
  if (isEvaluationError(inner)) return inner

  const num = coerceToNumber(inner)
  if (isEvaluationError(num)) return num

  if (node.op === '+') {
    return finiteOrRaw(num)
  }
  if (node.op === '-') {
    return finiteOrRaw(-num)
  }
  if (node.op === '%') {
    return finiteOrRaw(num / 100)
  }
  return formulaError('#VALUE!')
}

/**
 * Mode-aware binary combine. Scalar mode projects matrices to their top-left;
 * array mode broadcasts compatible singleton axes elementwise. A non-singleton
 * dimension mismatch is an honest capability gate (native outcome unmeasured),
 * never an invented #VALUE! or zero/#N/A padding. Scalar known-right errors keep
 * the tagged root error shape rather than guessing an array-of-errors extent.
 */
function combineBinaryNode(
  op: BinaryOp,
  left: EvaluationValue,
  right: EvaluationValue,
  ctx: EvaluationContext | undefined,
): EvaluationValue {
  if (isEvaluationError(right)) return right
  if (ctx?.mode === 'array' && (isMatrixValue(left) || isMatrixValue(right))) {
    return broadcastCombine(op, left, right, ctx)
  }
  return combineBinary({ op }, scalarProjection(left), scalarProjection(right))
}

function combineUnaryNode(op: '+' | '-' | '%', inner: EvaluationValue, ctx: EvaluationContext | undefined): EvaluationValue {
  if (ctx?.mode === 'array' && isMatrixValue(inner)) {
    const values = inner.values.map((row) => row.map((element) => combineUnary({ op }, element) as ElementValue))
    return matrixOf(inner.rows, inner.cols, values)
  }
  return combineUnary({ op }, scalarProjection(inner))
}

function broadcastCombine(
  op: BinaryOp,
  left: EvaluationValue,
  right: EvaluationValue,
  ctx: EvaluationContext | undefined,
): EvaluationValue {
  const lm = isMatrixValue(left) ? left : undefined
  const rm = isMatrixValue(right) ? right : undefined
  const lr = lm ? lm.rows : 1
  const lc = lm ? lm.cols : 1
  const rr = rm ? rm.rows : 1
  const rc = rm ? rm.cols : 1
  const rows = lr === rr ? lr : lr === 1 ? rr : rr === 1 ? lr : -1
  const cols = lc === rc ? lc : lc === 1 ? rc : rc === 1 ? lc : -1
  if (rows < 0 || cols < 0) {
    markUnsupported(ctx, 'array-shape-mismatch-unverified')
    ctx?.reportFormulaIssue?.({
      kind: 'unsupported-reference',
      feature: 'array-shape-mismatch-unverified',
      message: `Array operands ${lr}x${lc} and ${rr}x${rc} have no measured non-singleton broadcast outcome; no VALUE/padding is invented`,
    })
    return formulaError('#NAME?')
  }
  const at = (value: EvaluationValue, row: number, col: number): ElementValue => {
    if (!isMatrixValue(value)) return value
    return value.values[value.rows === 1 ? 0 : row][value.cols === 1 ? 0 : col]
  }
  const values: ElementValue[][] = []
  for (let row = 0; row < rows; row++) {
    const out: ElementValue[] = []
    for (let col = 0; col < cols; col++) out.push(combineBinary({ op }, at(left, row, col), at(right, row, col)) as ElementValue)
    values.push(out)
  }
  if (rows === 1 && cols === 1) return values[0][0]
  return matrixOf(rows, cols, values)
}

/** One driver frame: the node plus its child-scheduling state. Results live
 * in the AstNode memo, so a child completes into the memo and a parent frame
 * simply re-reads memo(child) — no sibling-ownership scans, naturally lazy and
 * resumable (PendingDependency throws out; the workbook re-runs the root and
 * already-computed nodes memo-hit). */
interface EvalFrame {
  node: AstNode
  phase: 'enter' | 'combine'
  /** Explicit evaluation mode for this frame (scalar/array store view). */
  mode: EvaluationMode
  /** Group 3 private adapter fields: dedicated name frame bookkeeping. */
  nameChild?: AstNode
  nameIdentity?: AstNode
}

/**
 * Bounded, non-recursive AST evaluation. Unary/binary chains are driven by an
 * explicit heap stack, scheduling ONE needed child at a time (Node 22
 * computes/diagnoses 8000 nested unaries and 4000 addends without RangeError);
 * function handlers keep their exact public `evalNode` callback contract (lazy
 * IF/IFERROR preserved) and re-enter the same driver through the shared memo.
 */
export function evaluateNode(node?: AstNode, ctx?: EvaluationContext): EvaluationValue {
  if (!node) return null
  const provided = ctx?.nodeValues
  // A deliberately supplied memo is caller-owned. Nested drivers borrowing the
  // current scoped memo retain its existing registration and unwind to its owner.
  if (provided) return evaluateNodeDriven(node, ctx, provided)
  if (!ctx) {
    // Standalone: a private ephemeral context gives the mode engine and its
    // evaluateInMode hook a home WITHOUT inventing currentCell/Sheet/data.
    return evaluateNode(node, { flatArgs: new WeakMap(), functionWork: new WeakMap() })
  }

  let memo = new WeakMap<AstNode, EvaluationValue>()
  let epoch = 0
  const restartToken = {} // invocation identity: never match error name/message
  try {
    for (;;) {
      ctx.nodeValues = memo
      const restoreOwnership = installOwnedMemoScope(ctx, memo, restartToken, () => epoch, () => { throw restartToken })
      try {
        return evaluateNodeDriven(node, ctx, memo)
      } catch (error) {
        if (error !== restartToken) throw error
        // Driver finally has restored active name scopes before the root retries.
        // Fresh memo ownership also refreshes its binding/body-clone weak caches.
        if (++epoch > MAX_REFERENCE_GENERATION_RESTARTS) {
          markUnsupported(ctx, 'reference-generation-unstable')
          ctx.reportFormulaIssue?.({
            kind: 'unsupported-reference',
            feature: 'reference-generation-unstable',
            message: `Reference footprint changed beyond the ${MAX_REFERENCE_GENERATION_RESTARTS}-restart application budget; owned calculation unavailable`,
          })
          return formulaError('#NAME?')
        }
        memo = new WeakMap<AstNode, EvaluationValue>()
      } finally {
        restoreOwnership()
      }
    }
  } finally {
    ctx.nodeValues = provided
  }
}

function evaluateNodeDriven(node: AstNode, ctx: EvaluationContext | undefined, memo: WeakMap<AstNode, EvaluationValue>): EvaluationValue {
  if (memo.has(node)) return memo.get(node)!

  // Mode-owned stores: the scalar and array views never share a memo, so a
  // scalar/top-left result is never reused for an array argument. Caller
  // supplied modeStores are used as-is; otherwise the primary store is the
  // passed memo and the other mode gets a private lazily-created store.
  const initialMode: EvaluationMode = ctx?.mode ?? 'scalar'
  const stores = ctx?.modeStores
  // Immutable primary memo: the loop rebinds `memo` per frame, so `memoFor`
  // must never close over the mutable binding (aliasing would let a scalar
  // frame read/write the array store).
  const primaryMemo = memo
  const primaryFlatArgs = ctx?.flatArgs
  const primaryFunctionWork = ctx?.functionWork
  const secondaryMemo = new WeakMap<AstNode, EvaluationValue>()
  const secondaryFlatArgs = new WeakMap<AstNode[], ArgumentProgress>()
  const secondaryFunctionWork = new WeakMap<AstNode[], unknown>()
  const memoFor = (mode: EvaluationMode): WeakMap<AstNode, EvaluationValue> =>
    stores ? stores[mode].nodeValues : mode === initialMode ? primaryMemo : secondaryMemo
  const flatArgsFor = (mode: EvaluationMode): EvaluationContext['flatArgs'] =>
    stores ? stores[mode].flatArgs : mode === initialMode ? primaryFlatArgs : secondaryFlatArgs
  const functionWorkFor = (mode: EvaluationMode): WeakMap<AstNode[], unknown> | undefined =>
    stores ? stores[mode].functionWork : mode === initialMode ? primaryFunctionWork : secondaryFunctionWork
  const bindMode = (mode: EvaluationMode): void => {
    if (!ctx) return
    ctx.mode = mode
    ctx.nodeValues = memoFor(mode)
    ctx.flatArgs = flatArgsFor(mode)
    ctx.functionWork = functionWorkFor(mode)
  }
  const savedCtx = ctx
    ? { mode: ctx.mode, nodeValues: ctx.nodeValues, flatArgs: ctx.flatArgs, functionWork: ctx.functionWork, evaluateInMode: ctx.evaluateInMode }
    : undefined
  // A privately created secondary view participates in the SAME root owned
  // scope (owner/epoch/matching restart). Caller-supplied stores are never
  // registered as owned and never cleared.
  const restoreSecondaryOwnership = ctx && !stores ? inheritOwnedMemoScope(ctx, primaryMemo, secondaryMemo) : () => {}
  // Approved ABI hook: request a child in the other mode through the same
  // engine. A memo hit returns directly; otherwise the requesting call frame
  // schedules it (never a recursive evaluateArrayFormula/driver call).
  const evaluateInMode = (child: AstNode, mode: EvaluationMode): EvaluationValue => {
    const target = memoFor(mode)
    if (target.has(child)) return target.get(child)!
    throw new NodeModeSuspension(child, mode)
  }
  // Honor a caller-provided hook; install the internal default only when absent.
  const providedHook = ctx?.evaluateInMode
  if (ctx && !providedHook) ctx.evaluateInMode = evaluateInMode

  const stack: EvalFrame[] = [{ node, phase: 'enter', mode: initialMode }]
  const push = (child: AstNode, mode: EvaluationMode): void => { stack.push({ node: child, phase: 'enter', mode }) }
  // Group 3: invocation-local active name identities (released on combine,
  // dropped on unwind; never stored across root replay). The set is
  // registered privately under this exact ctx for the resolver's hidden-
  // cycle bridge and unregistered in finally (save/restore previous).
  const activeNames = new Set<AstNode>()
  const restoreActiveNames = ctx ? installActiveNameScope(ctx, activeNames) : () => {}
  const requestNode = (child: AstNode, requestedCtx: EvaluationContext | undefined = ctx): EvaluationValue => {
    if (requestedCtx !== ctx) return evaluateNode(child, requestedCtx)
    if (memo.has(child)) return memo.get(child)!
    throw new NodeMemoSuspension(child)
  }
  try {
  while (stack.length > 0) {
    const frame = stack[stack.length - 1]
    const n = frame.node
    memo = memoFor(frame.mode)
    bindMode(frame.mode)
    if (frame.phase === 'enter') {
      if (memo.has(n)) { stack.pop(); continue }
      const relativeGate = relativeNodeGate(n)
      if (relativeGate) {
        markUnsupported(ctx, relativeGate.feature)
        ctx?.reportFormulaIssue?.({ kind: 'unsupported-reference', ...relativeGate })
        memo.set(n, formulaError('#NAME?'))
        stack.pop()
        continue
      }
      switch (n.type) {
        case 'number':
          // Literal values are ingested as parsed (no 15-digit pre-rounding:
          // INT(1.999999999999999) must truncate the ACTUAL value to 1).
          memo.set(n, n.value)
          stack.pop()
          continue
        case 'string':
          memo.set(n, n.value)
          stack.pop()
          continue
        case 'boolean':
          memo.set(n, n.value)
          stack.pop()
          continue
        case 'arrayConst':
          memo.set(n, matrixFromArrayConst(n))
          stack.pop()
          continue
        case 'empty':
          memo.set(n, null)
          stack.pop()
          continue
        case 'error': {
          if (isFormulaError(n.error)) {
            memo.set(n, formulaError(n.error))
          } else {
            // Parser-sentence error: the construct was ACTUALLY reached.
            markUnsupported(ctx, 'formula-syntax')
            memo.set(n, formulaError('#NAME?'))
          }
          stack.pop()
          continue
        }
        case 'cell':
          memo.set(n, evaluateCellRef(n, ctx))
          stack.pop()
          continue
        case 'range':
        case 'wholeCol':
        case 'wholeRow':
        case 'table':
        case 'ref3d':
        case 'union':
        case 'intersect':
        case 'spill':
          memo.set(n, ctx?.mode === 'array' ? referenceArrayValue(n, ctx, evaluateNode) : scalarOfReference(n, ctx, evaluateNode))
          stack.pop()
          continue

        case 'name': {
          // Group 3 (DESIGN verbatim-adapted): COOPERATIVE defined-name
          // evaluation. Bindings cached ONCE per AST node per the existing
          // nodeValues lifetime; the cached binding SOURCE AST is the cycle
          // identity; active identities release on combine. Optional-services
          // guard first (standalone contexts keep their scalar path).
          if (!ctx?.references) {
            memo.set(n, scalarOfReference(n, ctx, evaluateNode))
            stack.pop()
            continue
          }
          const memoBindings = bindingsByMemo.get(memo) ?? (() => {
            const m = new WeakMap<AstNode, NameBinding | EvaluationError>()
            bindingsByMemo.set(memo, m)
            return m
          })()
          let binding = memoBindings.get(n)
          if (binding === undefined) {
            binding = ctx.references.bindName(n.ref, ctx)
            memoBindings.set(n, binding)
          }
          if (isEvaluationError(binding)) {
            memo.set(n, binding)
            stack.pop()
            continue
          }
          const identity = binding.ast
          if (activeNames.has(identity)) {
            // markedNameCycle: internal typed #NAME? + 'name-cycle' marker +
            // specific issue (markUnsupported = the EXISTING evaluator helper);
            // native circular initialization/value UNMEASURED — workbook.ts
            // retains valid cached 77 under the current #NAME? policy.
            markUnsupported(ctx, 'name-cycle')
            ctx?.reportFormulaIssue?.({
              kind: 'unsupported-reference',
              feature: 'name-cycle',
              message: `Defined name "${binding.name}" is part of a cycle; bounded capability gate (native circular initialization/value unmeasured)`,
            })
            memo.set(n, formulaError('#NAME?'))
            stack.pop()
            continue
          }
          // Preserve the prior SERVICE path for ref-shaped binding targets
          // WITHOUT acquiring cycle ownership: reference/error exits below do
          // NOT appear active to the private bridge.
          const resolved = ctx.references.resolve(n, ctx)
          if (isEvaluationError(resolved)) {
            memo.set(n, resolved)
            stack.pop()
            continue
          }
          if (resolved) {
            // Group 10: the prepare hook runs before the value read. Array mode
            // keeps the full prepared rectangle; scalar mode keeps context projection.
            ctx.references.prepare(resolved, ctx)
            memo.set(n, ctx.mode === 'array' ? materializeResolved(resolved, ctx) : scalarOfResolved(resolved, ctx.currentCell, ctx))
            stack.pop()
            continue
          }
          // CRITICAL (UNION-CYCLE.md): ownership is acquired ONLY now — after
          // the initial resolve said this name is a VALUE — so the resolver's
          // own binding lookup never sees the owner as active (no premature
          // self-cycle for named SUM/union shapes).
          activeNames.add(identity)
          frame.nameIdentity = identity
          frame.nameChild = valueNameSource(binding, ctx, memo)
          frame.phase = 'combine'
          push(frame.nameChild, frame.mode)
          continue
        }

        case 'implicitIntersect': {
          // Reference operands project without scheduling a child whose own
          // evaluation would mistarget (a bare range evaluates to #VALUE!);
          // value operands evaluate normally, then project by identity.
          const inner = n.expr
          if (
            inner.type === 'cell' ||
            inner.type === 'range' ||
            inner.type === 'wholeCol' ||
            inner.type === 'wholeRow' ||
            inner.type === 'name' ||
            inner.type === 'table' ||
            inner.type === 'ref3d' ||
            inner.type === 'union' ||
            inner.type === 'intersect' ||
            inner.type === 'spill'
          ) {
            try {
              memo.set(n, scalarOfImplicitIntersect(inner, ctx, requestNode))
              stack.pop()
            } catch (error) {
              if (!(error instanceof NodeMemoSuspension)) throw error
              push(error.node, frame.mode)
            }
            continue
          }
          frame.phase = 'combine'
          push(n.expr, frame.mode)
          continue
        }
        case 'unary': {
          frame.phase = 'combine'
          push(n.expr, frame.mode)
          continue
        }
        case 'binary': {
          frame.phase = 'combine'
          push(n.left, frame.mode)
          continue
        }
        case 'call': {
          const handler = FUNCTIONS[n.name.toUpperCase()]
          if (!handler) {
            markUnsupported(ctx, n.name.toUpperCase())
            memo.set(n, formulaError('#NAME?'))
            stack.pop()
            continue
          }
          // Group 3: callback cooperation uses the SAME heap driver. The
          // callback requests a child via requestNode: memoized children
          // return directly; a pending child throws a PRIVATE request caught
          // ONLY by this call frame, which pushes the requested child on the
          // SAME stack and retries the handler. External PendingDependency,
          // prepare-hook and arbitrary user callback exceptions keep their
          // exact identity (never caught or converted).
          try {
            memo.set(n, handler(n.args, ctx, requestNode))
            stack.pop() // success: the call frame completes
            continue
          } catch (e) {
            if (e instanceof NodeModeSuspension) { push(e.node, e.mode); continue }
            if (!(e instanceof NodeMemoSuspension)) throw e
            // Requested child: keep the call frame alive (retry after the
            // child completes) — continue IMMEDIATELY so the unconditional
            // pop below never removes the just-pushed child.
            push(e.node, frame.mode)
            continue
          }
        }
        default:
          memo.set(n, formulaError('#VALUE!'))
          stack.pop()
          continue
      }
    }
    // combine phase: schedule one missing child at a time; when both children
    // are memoized, combine and record.
    if (n.type === 'name') {
      // Group 3 name combine (DESIGN): null is a legitimate completed value;
      // the active identity releases when the child completes.
      if (!memo.has(frame.nameChild!)) {
        push(frame.nameChild!, frame.mode)
        continue
      }
      activeNames.delete(frame.nameIdentity!)
      memo.set(n, memo.get(frame.nameChild!)!)
      stack.pop()
      continue
    }
    if (n.type === 'implicitIntersect') {
      const inner = memo.get(n.expr)
      if (inner === undefined) { push(n.expr, frame.mode); continue }
      // Value operands project by scalar identity; matrices project to their
      // top-left (explicit @ is always a scalar projection, even in array mode).
      memo.set(n, scalarProjection(inner))
      stack.pop()
      continue
    }
    if (n.type === 'unary') {
      const inner = memo.get(n.expr)
      if (inner === undefined) { push(n.expr, frame.mode); continue }
      memo.set(n, combineUnaryNode(n.op, inner, ctx))
      stack.pop()
      continue
    }
    if (n.type === 'binary') {
      const left = memo.get(n.left)
      if (left === undefined) { push(n.left, frame.mode); continue }
      // Native L→R: a known left TYPED error finishes the node — the right
      // child is never scheduled or read (no RHS callbacks, no cycle poisoning
      // through avoidable suspension). This scalar invariant is unchanged.
      if (isEvaluationError(left)) { memo.set(n, left); stack.pop(); continue }
      const right = memo.get(n.right)
      if (right === undefined) { push(n.right, frame.mode); continue }
      memo.set(n, combineBinaryNode(n.op, left, right, ctx))
      stack.pop()
      continue
    }
    memo.set(n, formulaError('#VALUE!'))
    stack.pop()
  }

  return memo.get(node)!
  } finally {
    restoreSecondaryOwnership()
    restoreActiveNames()
    if (ctx && savedCtx) {
      ctx.mode = savedCtx.mode
      ctx.nodeValues = savedCtx.nodeValues
      ctx.flatArgs = savedCtx.flatArgs
      ctx.functionWork = savedCtx.functionWork
      ctx.evaluateInMode = savedCtx.evaluateInMode
    }
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
    const value = evaluateNode(node)
    return isMatrixValue(value) ? topLeftOfMatrix(value) : value
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
    let res = evaluateNode(node, ctx)
    // Public scalar boundary: an array-mode-free evaluation projects a matrix
    // to its top-left element. Array mode (evaluateArrayFormula) preserves it.
    if (ctx.mode !== 'array' && isMatrixValue(res)) res = topLeftOfMatrix(res)
    if (currentDepth === 0 && ctx.hasCycle) {
      if (!ctx.currentCell) {
        return 0
      }
    }
    // Public typed-context boundary: a KNOWN blank (getCellValue defined and
    // returned null) evaluates to native 0; unbound data (no getCellValue)
    // stays null. Referenced-blank INPUT nulls inside formulas are untouched.
    if (res === null && currentDepth === 0 && ctx.getCellValue) {
      res = 0
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
  const value = evaluateFormulaInternal(input, ctx)
  return typeof value === 'number' ? round15(value) : publicFormulaValue(value)
}
