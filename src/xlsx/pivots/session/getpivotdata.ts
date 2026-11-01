/** Lazy GETPIVOTDATA adapter. This module belongs to the pivots chunk; the
 * eager formula engine exposes only registerFormulaFunction/getPivotData. */
import type { AstNode, EvaluationValue } from '../../formula/types'
import { formulaError, markUnsupported } from '../../formula/evaluator'
import { registerFormulaFunction } from '../../formula/functions'
import { dataFieldSource, PIVOT_VALUES_FIELD, type PivotCacheDefinition, type PivotTableDefinition } from '../model'
import type { PivotLayout } from '../engine/place'
import { pivotItemLabel } from '../engine/place'
import { PIVOT_FUNCTION_NAMES } from './functionNames'

export interface LivePivotDataContext {
  sheet: string
  def: PivotTableDefinition
  cache: PivotCacheDefinition
  layout: PivotLayout
  date1904: boolean
  /** The layout's cells by `${row}:${col}` (the viewer builds it once per commit); absent, one is built per layout. */
  layoutCells?: ReadonlyMap<string, PivotLayout['cells'][number]>
}

/** A layout's cells by position, built once per layout (final review minor 10: a lookup, not a scan per formula). */
const cellIndexes = new WeakMap<PivotLayout, ReadonlyMap<string, PivotLayout['cells'][number]>>()
function cellsOf(context: LivePivotDataContext): ReadonlyMap<string, PivotLayout['cells'][number]> {
  if (context.layoutCells) return context.layoutCells
  let index = cellIndexes.get(context.layout)
  if (!index) cellIndexes.set(context.layout, index = new Map(context.layout.cells.map(cell => [`${cell.row}:${cell.col}`, cell])))
  return index
}

const text = (value: EvaluationValue): string | undefined => typeof value === 'string' ? value : undefined

function fieldByName(def: PivotTableDefinition, cache: PivotCacheDefinition, name: string): number {
  return def.fields.findIndex((field, index) => (field.name || cache.fields[index]?.name || '').toLocaleLowerCase() === name.toLocaleLowerCase())
}

function dataByName(context: LivePivotDataContext, name: string): number {
  return context.def.dataFields.findIndex(data => {
    const source = dataFieldSource(data)
    const fieldName = context.def.fields[source]?.name ?? context.cache.fields[source]?.name ?? ''
    const label = data.name || `${PIVOT_FUNCTION_NAMES[data.subtotal] ?? data.subtotal} of ${fieldName}`
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

/**
 * What GETPIVOTDATA resolves to (Ruling S3-T): the one layout cell it reads; #REF! only for the visible-layout rules
 * Excel was measured on; or unresolved, which keeps the formula's saved value and marks it stale. Every failure that is
 * not one of the measured rules is unresolved, so a shape nobody measured never invents a #REF!.
 */
type Resolution =
  | { kind: 'cell'; value: EvaluationValue }
  /** Measured #REF! (gpd-hide-item, gpd-collapse-item, gpd-filter-out, the saved "Cherries" case): `not-visible`, the
   * requested items are on no line of the layout (hidden, filtered out, or under a collapsed parent); `inner-without-line`,
   * an inner item asked for without its parent has no line of its own (it repeats under several parents). */
  | { kind: 'ref'; rule: 'not-visible' | 'inner-without-line' }
  | { kind: 'unresolved'; reason: string }
type AxisLines = { kind: 'lines'; lines: Array<{ line: PivotLayout['axes']['rows'][number]; index: number }> }

function resolvePivotData(context: LivePivotDataContext, dataIndex: number, pairs: ReadonlyMap<number, EvaluationValue>): Resolution {
  const { def, layout } = context
  // Ruling S3-AO: a spilled pivot's GETPIVOTDATA was never measured; the formula keeps its saved value (S3-T).
  if (def.spill) return { kind: 'unresolved', reason: 'spilled pivot' }
  const axis = (fields: readonly number[], lines: PivotLayout['axes']['rows']): AxisLines | Exclude<Resolution, { kind: 'cell' }> => {
    const all = lines.map((line, index) => ({ line, index }))
    // An axis with no fields (a single data field and no column fields, or no row fields) lays out one line holding
    // every value: its sole data line is the cell, as the measured Values-only column axis (I4) generalises.
    if (!fields.length) return lines.length === 1 && lines[0]!.data === dataIndex ? { kind: 'lines', lines: all } : { kind: 'unresolved', reason: 'axis without fields' }
    const valuesLevel = fields.indexOf(PIVOT_VALUES_FIELD)
    // A line that holds values of the requested data field: on an axis carrying Values, its own data field, and not a
    // label-only parent line (a data line above the Values level only names its items; its subtotal is a total line).
    const holdsValues = (line: PivotLayout['axes']['rows'][number]) => line.type !== 'blank'
      && (valuesLevel < 0 || (line.data === dataIndex && !(line.type === 'data' && line.items.length <= valuesLevel)))
    const depth = (line: PivotLayout['axes']['rows'][number]) => line.items.filter((item, level) => item >= 0 && level !== valuesLevel).length
    const axisPairs = new Map([...pairs].filter(([field]) => fields.includes(field)))
    if (!axisPairs.size) {
      // No item of this axis asked for: its grand-total line (on an axis of Values alone, the data field's line).
      const found = all.filter(({ line }) => holdsValues(line) && (valuesLevel >= 0 && fields.length === 1 ? true : line.type === 'grand'))
      return found.length === 1 ? { kind: 'lines', lines: found } : { kind: 'unresolved', reason: found.length ? 'several total lines' : 'no total line' }
    }
    const carrying = all.filter(({ line }) => line.type !== 'blank' && lineMatches(context, fields, line, axisPairs))
    if (!carrying.length) return { kind: 'ref', rule: 'not-visible' }
    const candidates = carrying.filter(({ line }) => holdsValues(line))
    if (!candidates.length) return { kind: 'unresolved', reason: 'no line holds values' }
    // The shallowest visible line for the requested prefix: a parent's subtotal, else its item line.
    const shallowest = Math.min(...candidates.map(({ line }) => depth(line)))
    const at = candidates.filter(({ line }) => depth(line) === shallowest)
    const innerWithoutParent = [...axisPairs.keys()].some(field => {
      const level = fields.indexOf(field)
      return fields.slice(0, level).some(parent => parent !== PIVOT_VALUES_FIELD && !pairs.has(parent))
    })
    if (at.length === 1) return innerWithoutParent ? { kind: 'unresolved', reason: 'inner item with one parent' } : { kind: 'lines', lines: at }
    // Measured only on an axis without Values: an inner item repeated under several parents is #REF!.
    return innerWithoutParent && valuesLevel < 0 ? { kind: 'ref', rule: 'inner-without-line' } : { kind: 'unresolved', reason: 'several lines' }
  }
  const rows = axis(def.rowFields, layout.axes.rows)
  if (rows.kind !== 'lines') return rows
  const cols = axis(def.colFields, layout.axes.cols)
  if (cols.kind !== 'lines') return cols
  const row = layout.rect.minRow + def.location.firstDataRow + rows.lines[0]!.index
  const col = layout.rect.minCol + def.location.firstDataCol + cols.lines[0]!.index
  const cell = cellsOf(context).get(`${row}:${col}`)
  if (!cell) return { kind: 'unresolved', reason: 'no cell at the lines' }
  const value = cell.value
  return { kind: 'cell', value: value.kind === 'error' ? formulaError(value.value as Parameters<typeof formulaError>[0]) : value.kind === 'number' || value.kind === 'text' ? value.value : null }
}

export function getPivotData(contexts: readonly LivePivotDataContext[], args: readonly AstNode[], evaluate: (node: AstNode) => EvaluationValue, currentSheet?: string, markStale?: (feature: string) => void): EvaluationValue {
  if (args.length < 2 || args.length % 2 !== 0) return formulaError('#NAME?')
  const evaluatedName = evaluate(args[0]!)
  if (typeof evaluatedName === 'object' && evaluatedName !== null && 'kind' in evaluatedName && evaluatedName.kind === 'formula-error') return evaluatedName
  const dataName = text(evaluatedName)
  const anchor = args[1]
  if (dataName === undefined || anchor?.type !== 'cell') return formulaError('#NAME?')
  const sheet = anchor.ref.sheet ?? currentSheet
  const context = contexts.find(candidate => candidate.sheet.toLocaleLowerCase() === sheet?.toLocaleLowerCase()
    && candidate.layout.rect.minRow <= anchor.ref.row && candidate.layout.rect.maxRow >= anchor.ref.row
    && candidate.layout.rect.minCol <= anchor.ref.col && candidate.layout.rect.maxCol >= anchor.ref.col)
  if (!context) return formulaError('#NAME?')
  const dataIndex = dataByName(context, dataName)
  if (dataIndex < 0) return formulaError('#NAME?')
  const pairs = new Map<number, EvaluationValue>()
  for (let index = 2; index < args.length; index += 2) {
    const fieldName = text(evaluate(args[index]!)), value = evaluate(args[index + 1]!)
    if (fieldName === undefined) { markStale?.('GETPIVOTDATA'); return formulaError('#NAME?') }
    const field = fieldByName(context.def, context.cache, fieldName)
    if (field < 0 || pairs.has(field)) return formulaError('#NAME?')
    if (!context.def.rowFields.includes(field) && !context.def.colFields.includes(field) && !context.def.pageFields.some(page => page.fld === field)) return formulaError('#NAME?')
    if (typeof value === 'object' && value !== null && 'kind' in value && value.kind === 'formula-error') return value
    const items = (context.def.fields[field]?.items ?? []).flatMap(item => typeof item.x === 'number' ? [item.x] : [])
    const labels = items.map(item => pivotItemLabel(context.def, context.cache, field, item, context.date1904).value)
    if (!items.some(item => labelMatches(context, field, item, value))
      && ((typeof value === 'string' && labels.some(label => label.kind === 'number')) || (typeof value === 'number' && labels.some(label => label.kind === 'text')))) return formulaError('#NAME?')
    pairs.set(field, value)
  }
  // Page fields are visible only when the requested item is the current page
  // selection; the saved/canonical definition represents its current selection.
  for (const page of context.def.pageFields) {
    const requested = pairs.get(page.fld)
    if (requested === undefined && !pairs.has(page.fld)) continue
    const index = page.item
    if (index === undefined || !labelMatches(context, page.fld, index, requested!)) return formulaError('#NAME?')
  }
  const resolved = resolvePivotData(context, dataIndex, pairs)
  if (resolved.kind === 'cell') return resolved.value
  if (resolved.kind === 'ref') return formulaError('#REF!')
  markStale?.('GETPIVOTDATA')
  return formulaError('#NAME?')
}

export function registerGetPivotData(): void {
  registerFormulaFunction('GETPIVOTDATA', (args, ctx, evaluate) => {
    if (!ctx?.getPivotData) {
      // Until a pivot on the sheet has changed, treat this just like an
      // uninstalled function so the formula engine keeps Excel's shipped cache.
      markUnsupported(ctx, 'GETPIVOTDATA')
      return formulaError('#NAME?')
    }
    const result = ctx.getPivotData(args, evaluate, ctx.currentSheet ?? '', feature => markUnsupported(ctx, feature))
    if (typeof result === 'object' && result !== null && 'kind' in result && result.kind === 'formula-error' && result.code === '#NAME?') markUnsupported(ctx, 'GETPIVOTDATA')
    return result
  })
}

