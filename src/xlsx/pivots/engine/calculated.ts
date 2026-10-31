/**
 * Pivot calculated fields (spec §5.6): a cache field with a `formula` and no record column. Its formula combines the
 * SUMS of other fields and is evaluated per cell after aggregation (subtotal and grand-total cells included) through
 * the workbook formula engine, with each field name bound as a defined name whose value is that cell's sum. Lazy
 * chunk `pivots` (Stage 2: library only).
 *
 * The formula text is the cache's, quoted as written (`'Field A'*20`, `SUM('Field A', 'Field B')`); the lexer's
 * `quotedNames` mode (AGY-8) reads a quoted run not followed by `!` as a name.
 *
 * A name that is another calculated field re-applies that field's formula over the same cell's sums (Task 15
 * calc-references-calculated: Double Price = 'Price per Unit'*2 shows (48118.96/623)*2 for East, #DIV/0! where Units
 * sum to 0); `fields` lists the record fields the whole chain sums.
 *
 * Refused as `pivot-calculated-field` (the engine cannot evaluate it): a parse error, a cell/range/table/array or
 * other reference construct, a function the engine lacks or one whose value depends on the cell, the clock, chance or
 * the workbook's date system, a name that is no cache field (or more than one), and a cycle of calculated fields.
 * A formula over no field fills every value cell, cells without records included (calc-constant: `=1` shows 1 in
 * every cell). Results are not rounded to 15 digits: Excel saves calculated cells as full doubles (calc-show-as).
 */
import { tokenize } from '../../formula/lexer'
import { parseFormula } from '../../formula/parser'
import { evaluateFormulaInternal, formulaError, isFormulaError, publicFormulaValue } from '../../formula/evaluator'
import { FUNCTIONS } from '../../formula/functions'
import { DATETIME_FUNCTIONS } from '../../formula/functions/datetime'
import type { AstNode, EvaluationContext, NameBinding, ReferenceServices } from '../../formula/types'
import type { PivotCacheDefinition, PivotDiagnostic } from '../model'
import type { EngineResult } from './axes'
import type { PivotCellValue } from './place'

/** A calculated field's formula, parsed once, with each name it uses mapped to a cache field: `names` the record
 * fields (bound to their sums), `nested` the calculated fields (bound to their own formula), `fields` every record
 * field the chain sums. */
export interface CalculatedField {
  readonly ast: AstNode; readonly names: ReadonlyMap<string, number>; readonly fields: readonly number[]
  /** A referenced calculated field's parsed formula, by upper-cased name (every name the chain reaches was validated). */
  readonly nested: { get(name: string): AstNode | undefined }
}

const fail = (feature: string, message: string): { ok: false; diagnostic: PivotDiagnostic } => ({ ok: false, diagnostic: { kind: 'unsupported', feature, message } })

/** Functions that read the cell, the sheet, the clock or chance, or return arrays: not a pivot cell's scalar. */
const CONTEXT_FUNCTIONS = new Set([
  'RAND', 'RANDBETWEEN', 'RANDARRAY', 'NOW', 'TODAY', 'ROW', 'COLUMN', 'ROWS', 'COLUMNS', 'OFFSET', 'INDIRECT', 'CELL', 'INFO',
  'ISREF', 'ISFORMULA', 'FORMULATEXT', 'SHEET', 'SHEETS', 'LET', 'LAMBDA',
  'SEQUENCE', 'TRANSPOSE', 'UNIQUE', 'SORT', 'SORTBY', 'FILTER',
])

const key = (name: string) => name.toUpperCase()

/**
 * Parse a calculated field's formula and bind its names; `self` is the calculated field's cache index.
 * Amended contract (ruling S2-V): the brief's `calculatedValue(formula, sums)` is two stages, because binding names
 * needs the cache: `compileCalculatedField(formula, cache, self)` once, then `calculatedValue(compiled, sums)` per
 * cell. Both are engine-internal (values.ts); they are not exported from the pivots entry.
 */
/** Per cache: the name → field map, each formula's parse, and each compiled field (perf audit item 7: a chain of C
 * calculated fields used to parse C²/2 formulas, and value filters and sorts compiled again). */
interface CompileMemo { byName: Map<string, number>; parsed: Map<string, AstNode>; compiled: Map<string, EngineResult<CalculatedField>>; open: Set<string> }
const memos = new WeakMap<PivotCacheDefinition, CompileMemo>()
function memoOf(cache: PivotCacheDefinition): CompileMemo {
  let memo = memos.get(cache)
  if (!memo) {
    const byName = new Map<string, number>()
    cache.fields.forEach((field, index) => {
      const k = key(field.name ?? '')
      byName.set(k, byName.has(k) ? -1 : index)
    })
    memos.set(cache, memo = { byName, parsed: new Map(), compiled: new Map(), open: new Set() })
  }
  return memo
}
const parsedFormula = (memo: CompileMemo, formula: string): AstNode => {
  let ast = memo.parsed.get(formula)
  if (!ast) memo.parsed.set(formula, ast = parseFormula(tokenize(formula, { quotedNames: true })))
  return ast
}

export function compileCalculatedField(formula: string, cache: PivotCacheDefinition, self: number): EngineResult<CalculatedField> {
  const memo = memoOf(cache), memoKey = `${self}\u0000${formula}`
  let result = memo.compiled.get(memoKey)
  if (!result) {
    memo.open.add(memoKey)
    try { result = compileUncached(formula, cache, self, memo) } finally { memo.open.delete(memoKey) }
    memo.compiled.set(memoKey, result)
  }
  return result
}

function compileUncached(formula: string, cache: PivotCacheDefinition, self: number, memo: CompileMemo): EngineResult<CalculatedField> {
  const byName = memo.byName
  const names = new Map<string, number>(), nested = new Map<string, AstNode>(), fields = new Set<number>()
  const ownFormula = cache.fields[self]?.formula === formula
  let unsupported: string | undefined
  // Each calculated field of the chain is parsed once; `active` holds the chain being walked (a cycle check).
  const active = new Set<number>([self])
  const walk = (ast: AstNode): void => {
    const pending: AstNode[] = [ast]
    while (pending.length) {
      const node = pending.pop()!
      switch (node.type) {
        case 'number': case 'string': case 'boolean': case 'empty': break
        case 'unary': pending.push(node.expr); break
        case 'binary': pending.push(node.left, node.right); break
        case 'call':
          // Date functions need the workbook's date system, which the value pass does not carry (ruling S2-J).
          if (!FUNCTIONS[node.name] || CONTEXT_FUNCTIONS.has(node.name) || DATETIME_FUNCTIONS[node.name]) unsupported ??= `function ${node.name} is not evaluated in a pivot`
          pending.push(...node.args)
          break
        case 'name': {
          const field = node.ref.sheet === undefined ? byName.get(key(node.ref.name)) : undefined
          if (field === undefined || field < 0) { unsupported ??= `name ${node.ref.name} is not one cache field`; break }
          const cacheField = cache.fields[field]!
          if (field === self) unsupported ??= `the formula refers to its own field ${node.ref.name}`
          else if (cacheField.formula !== undefined) {
            if (active.has(field)) { unsupported ??= `calculated field ${node.ref.name} refers back to itself`; break }
            if (nested.has(key(node.ref.name))) break
            // A nested field that compiles alone binds exactly what walking it here would: compiling alone, a field on a
            // cycle fails (the cycle reaches it again), so an ok result reaches neither this field nor the chain being
            // walked. Reuse it: the chain compiles in linear time, not quadratic. A field still being compiled (a
            // cycle in progress) is walked as before, so the refusal and its message are unchanged.
            // (Only when this field compiles its own cache formula: the argument is about the cache's graph.)
            const reuse = !ownFormula || memo.open.has(`${field}\u0000${cacheField.formula}`) ? undefined : compileCalculatedField(cacheField.formula, cache, field)
            if (reuse?.ok) {
              nested.set(key(node.ref.name), reuse.value.ast)
              for (const [name, index] of reuse.value.names) names.set(name, index)
              for (const index of reuse.value.fields) fields.add(index)
              break
            }
            // calc-references-calculated: the referenced formula over the same sums.
            const inner = parsedFormula(memo, cacheField.formula)
            nested.set(key(node.ref.name), inner)
            active.add(field)
            walk(inner)
            active.delete(field)
          } else if (!cacheField.databaseField) unsupported ??= `field ${node.ref.name} has no records`
          else { names.set(key(node.ref.name), field); fields.add(field) }
          break
        }
        case 'error': unsupported ??= `the formula does not parse: ${node.error}`; break
        default: unsupported ??= `${node.type} is not evaluated in a pivot formula`
      }
    }
  }
  const ast = parsedFormula(memo, formula)
  walk(ast)
  if (unsupported) return fail('pivot-calculated-field', `calculated field ${self}: ${unsupported}`)
  // A formula over no field (=1) fills every value cell, those without records included (Task 22 calc-constant R22C4).
  // Nested formulas resolve through the cache memo: every calculated name the chain reaches was validated above.
  const resolve = {
    get: (name: string) => {
      const index = byName.get(name), formula = index !== undefined && index >= 0 ? cache.fields[index]?.formula : undefined
      return formula === undefined ? undefined : parsedFormula(memo, formula)
    },
  }
  return { ok: true, value: { ast, names, nested: resolve, fields: [...fields].sort((a, b) => a - b) } }
}

/**
 * One cell's value: the formula over `sums` (a cache field index → that field's sum in the cell). Errors and text are
 * returned as values for the caller's gate; 'unsupported' when the formula engine reports a construct it lacks.
 */
export function calculatedValue(field: CalculatedField, sums: (field: number) => number): PivotCellValue | 'unsupported' {
  const references: ReferenceServices = {
    resolve: node => node.type === 'name' ? undefined : formulaError('#REF!'),
    bindName: (ref): NameBinding | ReturnType<typeof formulaError> => {
      const inner = field.nested.get(key(ref.name))
      if (inner) return { name: ref.name, source: ref.name, ast: inner, baseProvenance: 'explicit' }
      const index = field.names.get(key(ref.name))
      if (index === undefined) return formulaError('#NAME?')
      return { name: ref.name, source: ref.name, ast: { type: 'number', value: sums(index) }, baseProvenance: 'explicit' }
    },
    prepare: () => {},
  }
  // A fresh context per cell: the evaluator caches name bindings per evaluation memo.
  const ctx: EvaluationContext = { references, unsupportedFeatures: new Set() }
  // Unrounded: Excel saves calculated cells as full doubles (calc-show-as R6C7 holds 2 × 22019.200000000004, and 64
  // of its 76 cells have more than 15 digits), and percent of total divides those unrounded results. The worksheet
  // boundary's 15-digit rounding (evaluateFormula) is not applied to pivot cells.
  const raw = evaluateFormulaInternal(field.ast, ctx)
  const value = typeof raw === 'number' ? raw : publicFormulaValue(raw)
  if (ctx.unsupportedFeatures!.size) return 'unsupported'
  if (typeof value === 'number') return Number.isFinite(value) ? { kind: 'number', value } : { kind: 'error', value: '#NUM!' }
  if (isFormulaError(value)) return { kind: 'error', value }
  if (typeof value === 'string') return { kind: 'text', value }
  if (typeof value === 'boolean') return { kind: 'text', value: value ? 'TRUE' : 'FALSE' }
  return 'unsupported'
}
