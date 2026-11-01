/**
 * The interaction oracle's comparison (plan Tasks 25 and 32): a session's live definition and layout against an
 * Excel-saved case, by checkSavedLayout's rules. Shared by the interaction oracle and the refresh tests.
 */
import { savedPivotLayout } from '../../src/xlsx/pivots/savedLayout'
import { sheetCellReader, type SavedCell } from '../../src/xlsx/pivots/engine/faithful'
import type { PivotLayout, PivotLayoutCell } from '../../src/xlsx/pivots/engine/place'
import { PIVOT_VALUES_FIELD, type PivotTableDefinition } from '../../src/xlsx/pivots/model'
import type { XlsxSheet } from '../../src/xlsx/types'

export type Outcome = { status: 'ok' } | { status: 'refused'; feature: string; message: string } | { status: 'differs'; reason: string }
export interface SavedCase { def: PivotTableDefinition; sheet: Pick<XlsxSheet, 'rows'> }
const same15 = (a: number, b: number) => a === b || Number(a.toPrecision(15)) === Number(b.toPrecision(15))
/** checkSavedLayout's comparison (faithful.ts `differs`), with implicit captions compared for presence. */
function differs(cell: PivotLayoutCell | undefined, saved: SavedCell | undefined): string | undefined {
  const blank = !cell || cell.value.kind === 'empty' || (cell.value.kind === 'text' && cell.value.value === '')
  if (blank) return saved === undefined ? undefined : `saved ${JSON.stringify(saved)}, computed blank`
  const value = cell!.value
  if (saved === undefined) return `computed ${JSON.stringify(value)}, saved blank`
  if (cell!.caption === 'implicit') return saved.kind === 'text' ? undefined : 'implicit caption over a non-text cell'
  if (value.kind === 'number') return saved.kind === 'number' && same15(value.value, saved.value) ? undefined : `computed ${value.value}, saved ${JSON.stringify(saved)}`
  if (value.kind === 'text' || value.kind === 'error') return saved.kind === value.kind && saved.value === value.value ? undefined : `computed ${JSON.stringify(value)}, saved ${JSON.stringify(saved)}`
  return 'the value differs'
}
const at = (row: number, col: number) => `R${row + 1}C${col + 1}`

/**
 * A definition's `rowItems`/`colItems` with each item index read through its field's `items` list (its cache item).
 * When Excel sorts a field it saves that field's `<items>` in the new order and indexes the lines into it; the live
 * definition keeps the saved list and indexes the lines into that. Both name the same items in the same order.
 */
/** Each row, column and page field's hidden and collapsed data items, by cache item (h/sd as Excel saves them). */
function itemFlags(def: PivotTableDefinition): string {
  const fields = new Set([...def.rowFields, ...def.colFields, ...def.pageFields.map(page => page.fld)])
  return JSON.stringify([...fields].filter(field => field !== PIVOT_VALUES_FIELD).sort((a, b) => a - b).map(field => {
    const items = (def.fields[field]?.items ?? []).filter(item => item.type === 'data')
    const by = (pick: (item: typeof items[number]) => boolean) => items.filter(pick).map(item => item.x).sort((a, b) => a! - b!)
    return { field, hidden: by(item => item.hidden), collapsed: by(item => !item.showDetails) }
  }))
}
function lines(def: PivotTableDefinition, axis: 'rowItems' | 'colItems'): string {
  const fields = axis === 'rowItems' ? def.rowFields : def.colFields
  return JSON.stringify((def[axis] ?? []).map(line => ({
    type: line.type, r: line.r, i: line.i,
    x: line.type === 'grand' ? line.x : line.x.map((x, k) => {
      const field = fields[line.r + k]
      return field === undefined || field === PIVOT_VALUES_FIELD ? `d${x}` : def.fields[field]?.items?.[x]?.x ?? x
    }),
  })))
}
export function compareCells(layout: PivotLayout, target: SavedCase): Outcome {
  const saved = sheetCellReader(target.sheet)
  const rects = savedPivotLayout(target.def)?.rects
  if (!rects) {
    // Excel's measured empty-pivot case omits all row/column/page/data fields and all populated pivot cells.
    if (!target.def.rowFields.length && !target.def.colFields.length && !target.def.pageFields.length && !target.def.dataFields.length && !layout.cells.length) return { status: 'ok' }
    return { status: 'differs', reason: 'the case has no saved layout' }
  }
  const cells = new Map(layout.cells.map(cell => [`${cell.row},${cell.col}`, cell]))
  const seen = new Set<string>()
  for (const rect of rects) for (let row = rect.minRow; row <= rect.maxRow; row++) for (let col = rect.minCol; col <= rect.maxCol; col++) {
    const key = `${row},${col}`
    if (seen.has(key)) continue
    seen.add(key)
    const reason = differs(cells.get(key), saved(row, col))
    if (reason) return { status: 'differs', reason: `${at(row, col)}: ${reason}` }
  }
  for (const cell of layout.cells) if (!seen.has(`${cell.row},${cell.col}`) && differs(cell, undefined)) return { status: 'differs', reason: `${at(cell.row, cell.col)} is outside the saved pivot area` }
  return { status: 'ok' }
}
export const describeOutcome = (outcome: Outcome) => outcome.status === 'ok' ? 'ok' : outcome.status === 'refused' ? `refused ${outcome.feature}: ${outcome.message}` : `differs: ${outcome.reason}`


/** The live definition's lines (through items), location and item flags, then every saved cell. */
export function compareWithCase(live: { def: PivotTableDefinition; layout: PivotLayout }, target: SavedCase): Outcome {
  for (const axis of ['rowItems', 'colItems'] as const) {
    const ours = lines(live.def, axis), theirs = lines(target.def, axis)
    if (ours !== theirs) return { status: 'differs', reason: `${axis}: live ${ours}, saved ${theirs}` }
  }
  if (JSON.stringify(live.def.location) !== JSON.stringify(target.def.location)) return { status: 'differs', reason: `location: live ${JSON.stringify(live.def.location)}, saved ${JSON.stringify(target.def.location)}` }
  const flags = itemFlags(live.def), saved = itemFlags(target.def)
  if (flags !== saved) return { status: 'differs', reason: `item flags (hidden, collapsed): live ${flags}, saved ${saved}` }
  return compareCells(live.layout, target)
}
