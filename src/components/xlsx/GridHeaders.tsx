import type { MouseEvent, ReactElement } from 'react'
import type { GridSelection, ScrollWindow, SortSpec, FilterSpec, ColumnRule } from '../../xlsx/view/contracts'
import type { ProjectedSheetMetrics } from '../../xlsx/view/projection'
import { columnToLetters } from '../../xlsx/view/coordinates'
import { logicalToViewport, type ViewportSize } from '../../xlsx/view/scrollWindow'
import { rowContains, boundedSelectionText } from '../../xlsx/view/selection'

export const GRID_HEADER_SIZE = 44
export interface HeaderSelectionOptions { toggle: boolean; extend: boolean }
export interface GridHeadersProps {
  projection: ProjectedSheetMetrics
  window: ScrollWindow
  viewport: ViewportSize
  coarse?: boolean
  columns: readonly number[]
  selection: GridSelection
  sort: SortSpec | null
  controlMode?: boolean
  onSelect?: (kind: 'row' | 'column', index: number, options: HeaderSelectionOptions, event: MouseEvent<HTMLButtonElement>) => void
  onSelectAll?: () => void
  controlColumn?: number | null
  filter?: FilterSpec | null
  onSortCycle?: (col: number, shift: boolean, invoker: HTMLButtonElement) => void
  onFilter?: (col: number, invoker: HTMLButtonElement) => void
  onSortDialog?: (col: number, invoker: HTMLButtonElement) => void
  onNotice?: (message: string) => void
}
/** Viewport CSS geometry, shared by both strips; DPR is applied by the painter,
 * not again to DOM CSS coordinates. ScrollWindow already includes recentering. */
export function gridHeaderBox(projection: ProjectedSheetMetrics, window: ScrollWindow, kind: 'row' | 'column', index: number) {
  const axis = kind === 'row' ? projection.rows : projection.cols
  const point = logicalToViewport(window, kind === 'row' ? { x: 0, y: axis.start(index) } : { x: axis.start(index), y: 0 })
  return kind === 'row' ? { x: 0, y: point.y, width: GRID_HEADER_SIZE, height: axis.size(index) * window.zoom }
    : { x: point.x, y: 0, width: axis.size(index) * window.zoom, height: GRID_HEADER_SIZE }
}
/** DOM geometry stays in the bounded pane even for a giant authored cell. */
export function viewportHeaderBox(projection: ProjectedSheetMetrics, window: ScrollWindow, viewport: ViewportSize, kind: 'row' | 'column', index: number) {
  const full = gridHeaderBox(projection, window, kind, index), length = kind === 'row' ? viewport.height : viewport.width
  const near = kind === 'row' ? full.y : full.x, size = kind === 'row' ? full.height : full.width
  const start = Math.min(length, Math.max(0, near)), end = Math.min(length, Math.max(start, near + size))
  return kind === 'row' ? { ...full, y: start, height: end - start } : { ...full, x: start, width: end - start }
}
function rowCount(rows: GridSelection['ranges'][number]['rows']): number {
  return rows.kind === 'intervals' ? rows.intervals.reduce((n, r) => n + r.r1 - r.r0 + 1, 0) : rows.kind === 'list' ? rows.rows.length : rows.count
}
function selectedColumn(props: GridHeadersProps, col: number): boolean {
  return props.selection.ranges.reduce((n, part) => n + (part.cols.some(c => col >= c.c0 && col <= c.c1) ? rowCount(part.rows) : 0), 0) === props.projection.rows.logicalCount
}
export function columnRuleDescription(rule: ColumnRule): string {
  if (rule.kind === 'compound') return boundedSelectionText(`${columnRuleDescription(rule.first)} ${rule.connective.toUpperCase()} ${columnRuleDescription(rule.second)}`)
  if (rule.kind === 'values') return rule.selection.mode === 'allExcept' ? `All except ${rule.selection.excluded.count} values` : rule.selection.mode === 'include' ? `${rule.selection.keys.count} included values` : 'Complete value-query selection'
  if (rule.kind === 'datePreset') return `Date period: ${rule.preset}`
  const op: Record<string, string> = { gt: 'Greater than', gte: 'Greater than or equal', lt: 'Less than', lte: 'Less than or equal', eq: 'Equals', neq: 'Does not equal', between: 'Between', contains: 'Contains', 'not-contains': 'Does not contain', begins: 'Begins with', ends: 'Ends with', before: 'Before', after: 'After' }
  const value = typeof rule.first === 'object' ? `${rule.first.year}-${rule.first.month}-${rule.first.day}` : String(rule.first)
  return boundedSelectionText(`${op[rule.op]} ${value}${'second' in rule && rule.second !== undefined ? ` through ${typeof rule.second === 'object' ? `${rule.second.year}-${rule.second.month}-${rule.second.day}` : rule.second}` : ''}`)
}
export function GridHeaders(props: GridHeadersProps): ReactElement {
  return <div role="row" aria-rowindex={1} data-officeview-ui="grid-headers" onPointerDown={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()} style={{ position: 'absolute', top: 0, left: 0, height: GRID_HEADER_SIZE, width: '100%', zIndex: 2, background: '#f3f4f6' }}>
    <div role="columnheader" aria-colindex={1} aria-label="Worksheet corner" style={{ position: 'absolute', left: 0, width: GRID_HEADER_SIZE, height: GRID_HEADER_SIZE, zIndex: 1 }}>
      <button type="button" aria-label="Select worksheet" tabIndex={props.controlMode ? 0 : -1} onClick={props.onSelectAll} style={{ width: GRID_HEADER_SIZE, height: GRID_HEADER_SIZE }}>↖</button>
    </div>
    {props.columns.map(col => {
      const box = viewportHeaderBox(props.projection, props.window, props.viewport, 'column', col), priority = props.sort?.keys.findIndex(key => key.col === col) ?? -1
      const direction = priority < 0 ? undefined : props.sort!.keys[priority].order === 'asc' ? 'ascending' : 'descending'
      if (props.coarse) box.width = Math.min(GRID_HEADER_SIZE, props.viewport.width - box.x)
      const selected = selectedColumn(props, col), label = `Column ${columnToLetters(col)}${direction ? `; sort priority ${priority + 1}, ${direction}` : ''}`
      const rule = props.filter?.columns.get(col), description = rule ? columnRuleDescription(rule) : 'No viewer filter'
      const actions = Boolean(props.onSortCycle || props.onFilter || props.onSortDialog)
      return <div key={col} role="columnheader" aria-colindex={col + 2} aria-label={label} aria-selected={selected} aria-sort={priority === 0 ? direction : undefined}
        style={{ position: 'absolute', left: GRID_HEADER_SIZE + box.x, width: box.width, height: box.height }}>
        <button type="button" aria-label={`Select ${label}`} data-header-column={col} tabIndex={props.controlMode ? 0 : -1}
          aria-description={actions ? `${description}; focus for column tools` : undefined} title={`${label}; ${description}`}
          onClick={event => { props.onNotice?.(`${label}; ${description}`); props.onSelect?.('column', col, { toggle: event.ctrlKey || event.metaKey, extend: event.shiftKey }, event) }}
          style={{ minWidth: GRID_HEADER_SIZE, minHeight: GRID_HEADER_SIZE, width: '100%', border: 0, background: 'transparent', padding: 0 }}><span style={{ display: 'block', width: box.width, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{columnToLetters(col)}</span></button>
        {actions && <div hidden={props.controlMode === false || (props.controlColumn !== undefined && props.controlColumn !== col)} style={{ position: 'absolute', top: GRID_HEADER_SIZE, left: 0, width: Math.max(GRID_HEADER_SIZE, box.width), background: '#fff', border: '1px solid #888', zIndex: 3 }}>
          <button type="button" data-action-column={col} tabIndex={props.controlMode ? 0 : -1} style={{ minHeight: 44, minWidth: 44, width: '100%' }} disabled={!props.onSortCycle} aria-label={`Cycle sort column ${columnToLetters(col)}`} aria-description={direction ?? 'unsorted'} onClick={event => props.onSortCycle?.(col, event.shiftKey, event.currentTarget)}>{direction ?? 'Sort'}</button>
          <button type="button" data-action-column={col} tabIndex={props.controlMode ? 0 : -1} style={{ minHeight: 44, minWidth: 44, width: '100%' }} disabled={!props.onFilter} aria-label={`Filter column ${columnToLetters(col)} — ${description}`} aria-pressed={Boolean(rule)} onClick={event => { props.onNotice?.(`${label}; ${description}`); props.onFilter?.(col, event.currentTarget) }}>Filter</button>
          {props.onSortDialog && <button type="button" data-action-column={col} tabIndex={props.controlMode ? 0 : -1} style={{ minHeight: 44, minWidth: 44, width: '100%' }} onClick={event => props.onSortDialog?.(col, event.currentTarget)}>More sort options for {columnToLetters(col)}</button>}
        </div>}
      </div>
    })}
  </div>
}
export function GridRowHeader({ props, slot }: { props: GridHeadersProps; slot: number }): ReactElement {
  const row = props.projection.order.displayToSource(slot), box = viewportHeaderBox(props.projection, props.window, props.viewport, 'row', slot)
  const count = props.selection.ranges.reduce((n, part) => n + (rowContains(part.rows, row) ? part.cols.reduce((sum, c) => sum + c.c1 - c.c0 + 1, 0) : 0), 0)
  return <div role="rowheader" aria-colindex={1} aria-label={`Row ${slot + 1}`} aria-selected={count === props.projection.cols.logicalCount}
    data-officeview-ui="row-header" onPointerDown={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()} style={{ position: 'absolute', width: GRID_HEADER_SIZE, height: '100%', pointerEvents: 'auto', background: '#f3f4f6' }}>
    <button type="button" aria-label={`Select row ${slot + 1}`} data-header-row={slot} tabIndex={props.controlMode ? 0 : -1}
      onClick={event => props.onSelect?.('row', row, { toggle: event.ctrlKey || event.metaKey, extend: event.shiftKey }, event)}
      style={{ position: 'absolute', top: '50%', transform: 'translateY(-50%)', minWidth: GRID_HEADER_SIZE, minHeight: GRID_HEADER_SIZE, width: GRID_HEADER_SIZE, border: 0, padding: 0, background: 'transparent' }}><span style={{ display: 'block', height: box.height, overflow: 'hidden', fontSize: Math.min(14, box.height), lineHeight: `${box.height}px` }}>{slot + 1}</span></button>
  </div>
}
