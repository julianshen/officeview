/** Lazy GETPIVOTDATA adapter. This module belongs to the pivots chunk; the
 * eager formula engine exposes only registerFormulaFunction/getPivotData. */
import type { AstNode, EvaluationValue } from '../../formula/types'
import { formulaError, markUnsupported } from '../../formula/evaluator'
import { registerFormulaFunction } from '../../formula/functions'
import { dataFieldSource, PIVOT_VALUES_FIELD, type PivotCacheDefinition, type PivotTableDefinition } from '../model'
import type { PivotLayout } from '../engine/place'
import { pivotItemLabel } from '../engine/place'

export interface LivePivotDataContext {
  sheet: string
  def: PivotTableDefinition
  cache: PivotCacheDefinition
  layout: PivotLayout
  date1904: boolean
}

const functions: Record<string, string> = { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count', stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp' }
const text = (value: EvaluationValue): string | undefined => typeof value === 'string' ? value : undefined

function fieldByName(def: PivotTableDefinition, cache: PivotCacheDefinition, name: string): number {
  return def.fields.findIndex((field, index) => (field.name || cache.fields[index]?.name || '').toLocaleLowerCase() === name.toLocaleLowerCase())
}

function dataByName(context: LivePivotDataContext, name: string): number {
  return context.def.dataFields.findIndex(data => {
    const source = dataFieldSource(data)
    const fieldName = context.def.fields[source]?.name ?? context.cache.fields[source]?.name ?? ''
    const label = data.name || `${functions[data.subtotal] ?? data.subtotal} of ${fieldName}`
    return label.toLocaleLowerCase() === name.toLocaleLowerCase()
  })
}

function labelMatches(context: LivePivotDataContext, field: number, item: number, requested: EvaluationValue): boolean {
  const actual = pivotItemLabel(context.def, context.cache, field, item, context.date1904).value
  if (actual.kind === 'empty') return requested === null || requested === ''
  if (actual.kind === 'number') return typeof requested === 'number' && actual.value === requested
  if (actual.kind === 'text') return typeof requested === 'string' && actual.value.toLocaleLowerCase() === requested.toLocaleLowerCase()
  return false
}

function lineMatches(context: LivePivotDataContext, axis: readonly number[], line: { items: number[]; type: string }, pairs: ReadonlyMap<number, EvaluationValue>): boolean {
  if (line.type === 'blank') return false
  for (const [field, requested] of pairs) {
    const level = axis.indexOf(field)
    if (level < 0 || line.items[level] === undefined || line.items[level]! < 0 || !labelMatches(context, field, line.items[level]!, requested)) return false
  }
  return true
}

function pivotValue(context: LivePivotDataContext, dataIndex: number, pairs: ReadonlyMap<number, EvaluationValue>): EvaluationValue {
  const { def, layout } = context
  const matches = (fields: readonly number[], lines: PivotLayout['axes']['rows']) => {
    const axisPairs = new Map([...pairs].filter(([field]) => fields.includes(field)))
    if (fields.includes(PIVOT_VALUES_FIELD)) return lines.map((line, index) => ({ line, index }))
      .filter(({ line }) => line.data === dataIndex && (axisPairs.size ? lineMatches(context, fields, line, axisPairs) : line.type === 'grand'))
    const candidates = lines.map((line, index) => ({ line, index })).filter(({ line }) => axisPairs.size
      ? lineMatches(context, fields, line, axisPairs) : line.type === 'grand')
    // An omitted axis field means its visible grand-total line; never sum or
    // synthesize a value that is not a cell in the current layout.
    if (!axisPairs.size) return candidates
    // Prefer the shallowest visible total/item line for the specified prefix.
    const depth = Math.min(...candidates.map(({ line }) => line.items.filter(item => item >= 0).length))
    return Number.isFinite(depth) ? candidates.filter(({ line }) => line.items.filter(item => item >= 0).length === depth) : []
  }
  const rows = matches(def.rowFields, layout.axes.rows)
  const cols = matches(def.colFields, layout.axes.cols)
  const candidates: Array<{ row: number; col: number }> = []
  for (const row of rows) for (const col of cols) {
    const sheetRow = layout.rect.minRow + def.location.firstDataRow + row.index
    const sheetCol = layout.rect.minCol + def.location.firstDataCol + col.index
    candidates.push({ row: sheetRow, col: sheetCol })
  }
  // If Values is on an axis, its line index already selects the data field;
  // otherwise the cell grid's data dimension follows the column lines.
  const visible = candidates.map(position => layout.cells.find(cell => cell.row === position.row && cell.col === position.col)).filter(Boolean)
  if (visible.length !== 1) return formulaError('#REF!')
  const value = visible[0]!.value
  if (value.kind === 'error') return formulaError(value.value as Parameters<typeof formulaError>[0])
  if (value.kind === 'number' || value.kind === 'text') return value.value
  return null
}

export function getPivotData(contexts: readonly LivePivotDataContext[], args: readonly AstNode[], evaluate: (node: AstNode) => EvaluationValue, currentSheet?: string): EvaluationValue {
  if (args.length < 2 || args.length % 2 !== 0) return formulaError('#VALUE!')
  const dataName = text(evaluate(args[0]!))
  const anchor = args[1]
  if (dataName === undefined || anchor?.type !== 'cell') return formulaError('#REF!')
  const sheet = anchor.ref.sheet ?? currentSheet
  const context = contexts.find(candidate => candidate.sheet.toLocaleLowerCase() === sheet?.toLocaleLowerCase()
    && candidate.layout.rect.minRow <= anchor.ref.row && candidate.layout.rect.maxRow >= anchor.ref.row
    && candidate.layout.rect.minCol <= anchor.ref.col && candidate.layout.rect.maxCol >= anchor.ref.col)
  if (!context) return formulaError('#REF!')
  const dataIndex = dataByName(context, dataName)
  if (dataIndex < 0) return formulaError('#REF!')
  const pairs = new Map<number, EvaluationValue>()
  for (let index = 2; index < args.length; index += 2) {
    const fieldName = text(evaluate(args[index]!)), value = evaluate(args[index + 1]!)
    if (fieldName === undefined) return formulaError('#VALUE!')
    const field = fieldByName(context.def, context.cache, fieldName)
    if (field < 0 || pairs.has(field)) return formulaError('#REF!')
    pairs.set(field, value)
  }
  // Page fields are visible only when the requested item is the current page
  // selection; the saved/canonical definition represents its current selection.
  for (const page of context.def.pageFields) {
    const requested = pairs.get(page.fld)
    if (requested === undefined && !pairs.has(page.fld)) continue
    const index = page.item
    if (index === undefined || !labelMatches(context, page.fld, index, requested!)) return formulaError('#REF!')
  }
  return pivotValue(context, dataIndex, pairs)
}

export function registerGetPivotData(): void {
  registerFormulaFunction('GETPIVOTDATA', (args, ctx, evaluate) => {
    if (!ctx?.getPivotData) {
      // Until a pivot on the sheet has changed, treat this just like an
      // uninstalled function so the formula engine keeps Excel's shipped cache.
      markUnsupported(ctx, 'GETPIVOTDATA')
      return formulaError('#NAME?')
    }
    return ctx.getPivotData(args, evaluate)
  })
}

