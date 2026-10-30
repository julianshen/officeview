/**
 * Pivot calculated fields (spec §5.6): a cache field with a `formula` and no record column. Its formula combines the
 * SUMS of other fields and is evaluated per cell after aggregation (subtotal and grand-total cells included) through
 * the workbook formula engine, with each field name bound as a defined name whose value is that cell's sum. Lazy
 * chunk `pivots` (Stage 2: library only).
 *
 * The formula text is the cache's, quoted as written (`'Field A'*20`, `SUM('Field A', 'Field B')`); the lexer's
 * `quotedNames` mode (AGY-8) reads a quoted run not followed by `!` as a name.
 *
 * Refused as `pivot-calculated-field` (the engine cannot evaluate it): a parse error, a cell/range/table/array or
 * other reference construct, a function the engine lacks or one whose value depends on the cell, the clock, chance or
 * the workbook's date system, and a name that is no cache field (or more than one). Refused as `pivot-engine-unmeasured` (ruling S2-M, no
 * Excel-written oracle): a reference to another calculated field (whether Excel re-applies its formula over the sums
 * or reads its cell value), and a formula over no field (whether Excel fills cells without records).
 */
import { tokenize } from '../../formula/lexer'
import { parseFormula } from '../../formula/parser'
import { evaluateFormula, formulaError, isFormulaError } from '../../formula/evaluator'
import { FUNCTIONS } from '../../formula/functions'
import { DATETIME_FUNCTIONS } from '../../formula/functions/datetime'
import type { AstNode, EvaluationContext, NameBinding, ReferenceServices } from '../../formula/types'
import type { PivotCacheDefinition, PivotDiagnostic } from '../model'
import type { EngineResult } from './axes'
import type { PivotCellValue } from './place'

/** A calculated field's formula, parsed once, with each name it uses mapped to a cache field. */
export interface CalculatedField { readonly ast: AstNode; readonly names: ReadonlyMap<string, number>; readonly fields: readonly number[] }

const fail = (feature: string, message: string): { ok: false; diagnostic: PivotDiagnostic } => ({ ok: false, diagnostic: { kind: 'unsupported', feature, message } })

/** Functions that read the cell, the sheet, the clock or chance, or return arrays: not a pivot cell's scalar. */
const CONTEXT_FUNCTIONS = new Set([
  'RAND', 'RANDBETWEEN', 'RANDARRAY', 'NOW', 'TODAY', 'ROW', 'COLUMN', 'ROWS', 'COLUMNS', 'OFFSET', 'INDIRECT', 'CELL', 'INFO',
  'ISREF', 'ISFORMULA', 'FORMULATEXT', 'SHEET', 'SHEETS', 'LET', 'LAMBDA',
  'SEQUENCE', 'TRANSPOSE', 'UNIQUE', 'SORT', 'SORTBY', 'FILTER',
])

const key = (name: string) => name.toUpperCase()

/** Parse a calculated field's formula and bind its names; `self` is the calculated field's cache index. */
export function compileCalculatedField(formula: string, cache: PivotCacheDefinition, self: number): EngineResult<CalculatedField> {
  const byName = new Map<string, number>()
  cache.fields.forEach((field, index) => {
    const k = key(field.name ?? '')
    byName.set(k, byName.has(k) ? -1 : index)
  })
  const ast = parseFormula(tokenize(formula, { quotedNames: true }))
  const names = new Map<string, number>(), fields = new Set<number>()
  let unsupported: string | undefined, unmeasured: string | undefined
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
        else if (cacheField.formula !== undefined) unmeasured ??= `a reference to calculated field ${node.ref.name} is not measured`
        else if (!cacheField.databaseField) unsupported ??= `field ${node.ref.name} has no records`
        else { names.set(key(node.ref.name), field); fields.add(field) }
        break
      }
      case 'error': unsupported ??= `the formula does not parse: ${node.error}`; break
      default: unsupported ??= `${node.type} is not evaluated in a pivot formula`
    }
  }
  if (unsupported) return fail('pivot-calculated-field', `calculated field ${self}: ${unsupported}`)
  if (unmeasured) return fail('pivot-engine-unmeasured', `calculated field ${self}: ${unmeasured}`)
  if (!fields.size) return fail('pivot-engine-unmeasured', `calculated field ${self} uses no field; its cells without records are not measured`)
  return { ok: true, value: { ast, names, fields: [...fields].sort((a, b) => a - b) } }
}

/**
 * One cell's value: the formula over `sums` (a cache field index → that field's sum in the cell). Errors and text are
 * returned as values for the caller's gate; 'unsupported' when the formula engine reports a construct it lacks.
 */
export function calculatedValue(field: CalculatedField, sums: (field: number) => number): PivotCellValue | 'unsupported' {
  const references: ReferenceServices = {
    resolve: node => node.type === 'name' ? undefined : formulaError('#REF!'),
    bindName: (ref): NameBinding | ReturnType<typeof formulaError> => {
      const index = field.names.get(key(ref.name))
      if (index === undefined) return formulaError('#NAME?')
      return { name: ref.name, source: ref.name, ast: { type: 'number', value: sums(index) }, baseProvenance: 'explicit' }
    },
    prepare: () => {},
  }
  // A fresh context per cell: the evaluator caches name bindings per evaluation memo.
  const ctx: EvaluationContext = { references, unsupportedFeatures: new Set() }
  const value = evaluateFormula(field.ast, ctx)
  if (ctx.unsupportedFeatures!.size) return 'unsupported'
  if (typeof value === 'number') return Number.isFinite(value) ? { kind: 'number', value } : { kind: 'error', value: '#NUM!' }
  if (isFormulaError(value)) return { kind: 'error', value }
  if (typeof value === 'string') return { kind: 'text', value }
  if (typeof value === 'boolean') return { kind: 'text', value: value ? 'TRUE' : 'FALSE' }
  return 'unsupported'
}
