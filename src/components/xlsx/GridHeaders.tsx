import type { MouseEvent, ReactElement } from 'react'
import type { GridSelection, ScrollWindow, SortSpec } from '../../xlsx/view/contracts'
import type { ProjectedSheetMetrics } from '../../xlsx/view/projection'
import { columnToLetters } from '../../xlsx/view/coordinates'
import { logicalToViewport, type ViewportSize } from '../../xlsx/view/scrollWindow'
import { rowContains } from '../../xlsx/view/selection'

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
type SelectionPart = GridSelection['ranges'][number]
/** Sorted, merged copy; S6.1 parts may overlap (a Shift extension keeps its
 * Ctrl prefix), so header fullness is a union, never a sum. */
function mergeIntervals(intervals: readonly { lo: number; hi: number }[]): { lo: number; hi: number }[] {
  const out: { lo: number; hi: number }[] = []
  for (const next of [...intervals].sort((a, b) => a.lo - b.lo)) {
    const previous = out[out.length - 1]
    if (previous && next.lo <= previous.hi + 1) previous.hi = Math.max(previous.hi, next.hi)
    else out.push({ ...next })
  }
  return out
}
/** Worker `ref` row sets need a resolver this shell does not have, so their
 * membership is unknown (undefined: aria-selected is omitted, not false). */
export function cellSelected(ranges: readonly SelectionPart[], at: { row: number; col: number }): boolean | undefined {
  let unknown = false
  for (const part of ranges) {
    if (!part.cols.some(c => at.col >= c.c0 && at.col <= c.c1)) continue
    if (part.rows.kind === 'ref') unknown = true
    else if (rowContains(part.rows, at.row)) return true
  }
  return unknown ? undefined : false
}
function selectedColumn(props: GridHeadersProps, col: number): boolean | undefined {
  const parts = props.selection.ranges.filter(part => part.cols.some(c => col >= c.c0 && col <= c.c1))
  const last = props.projection.rows.logicalCount - 1, intervals: { lo: number; hi: number }[] = []
  let unknown = false
  for (const part of parts) {
    if (part.rows.kind === 'intervals') for (const r of part.rows.intervals) intervals.push({ lo: r.r0, hi: r.r1 })
    else if (part.rows.kind === 'ref') unknown = true
  }
  const merged = mergeIntervals(intervals)
  if (merged.length === 1 && merged[0].lo <= 0 && merged[0].hi >= last) return true
  const sparse = parts.filter(part => part.rows.kind === 'list' || part.rows.kind === 'bitset')
  if (sparse.length) {
    // Fill only the gaps the intervals leave; bounded sets cannot cover more
    // rows than they hold, so a larger gap is decided without scanning it.
    const capacity = sparse.reduce((n, part) => n + (part.rows.kind === 'list' ? part.rows.rows.length : part.rows.kind === 'bitset' ? part.rows.count : 0), 0)
    let next = 0, gap = 0
    const gaps: { lo: number; hi: number }[] = []
    for (const r of merged) { if (r.lo > next) gaps.push({ lo: next, hi: Math.min(last, r.lo - 1) }); next = Math.max(next, r.hi + 1) }
    if (next <= last) gaps.push({ lo: next, hi: last })
    for (const g of gaps) gap += g.hi - g.lo + 1
    if (gap <= capacity && gaps.every(g => { for (let row = g.lo; row <= g.hi; row++) if (!sparse.some(part => rowContains(part.rows, row))) return false; return true })) return true
  }
  return unknown ? undefined : false
}
function selectedRow(props: GridHeadersProps, row: number): boolean | undefined {
  const cols: { lo: number; hi: number }[] = []
  let unknown = false
  for (const part of props.selection.ranges) {
    if (part.rows.kind === 'ref') unknown = true
    else if (rowContains(part.rows, row)) for (const c of part.cols) cols.push({ lo: c.c0, hi: c.c1 })
  }
  const merged = mergeIntervals(cols)
  if (merged.length === 1 && merged[0].lo <= 0 && merged[0].hi >= props.projection.cols.logicalCount - 1) return true
  return unknown ? undefined : false
}
// A mouse press must not move focus onto a header control outside control mode.
const keepGridFocus = (event: MouseEvent<HTMLButtonElement>) => { if (event.currentTarget.tabIndex < 0) event.preventDefault() }
export function GridHeaders(props: GridHeadersProps): ReactElement {
  return <div role="row" aria-rowindex={1} data-officeview-ui="grid-headers" onPointerDown={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()} style={{ position: 'absolute', top: 0, left: 0, height: GRID_HEADER_SIZE, width: '100%', zIndex: 2, background: '#f3f4f6' }}>
    <div role="columnheader" aria-colindex={1} aria-label="Worksheet corner" style={{ position: 'absolute', left: 0, width: GRID_HEADER_SIZE, height: GRID_HEADER_SIZE, zIndex: 1 }}>
      <button type="button" aria-label="Select worksheet" tabIndex={props.controlMode ? 0 : -1} onMouseDown={keepGridFocus} onClick={props.onSelectAll} style={{ width: GRID_HEADER_SIZE, height: GRID_HEADER_SIZE }}>↖</button>
    </div>
    {props.columns.map(col => {
      const box = viewportHeaderBox(props.projection, props.window, props.viewport, 'column', col), priority = props.sort?.keys.findIndex(key => key.col === col) ?? -1
      const direction = priority < 0 ? undefined : props.sort!.keys[priority].order === 'asc' ? 'ascending' : 'descending'
      if (props.coarse) box.width = Math.min(GRID_HEADER_SIZE, props.viewport.width - box.x)
      const selected = selectedColumn(props, col), label = `Column ${columnToLetters(col)}${direction ? `; sort priority ${priority + 1}, ${direction}` : ''}`
      return <div key={col} role="columnheader" aria-colindex={col + 2} aria-label={label} aria-selected={selected} aria-sort={priority === 0 ? direction : undefined}
        style={{ position: 'absolute', left: GRID_HEADER_SIZE + box.x, width: box.width, height: box.height }}>
        <button type="button" aria-label={`Select ${label}`} data-header-column={col} tabIndex={props.controlMode ? 0 : -1} onMouseDown={keepGridFocus}
          onClick={event => props.onSelect?.('column', col, { toggle: event.ctrlKey || event.metaKey, extend: event.shiftKey }, event)}
          style={{ minWidth: GRID_HEADER_SIZE, minHeight: GRID_HEADER_SIZE, width: '100%', border: 0, background: 'transparent', padding: 0 }}><span style={{ display: 'block', width: box.width, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{columnToLetters(col)}</span></button>
      </div>
    })}
  </div>
}
export function GridRowHeader({ props, slot }: { props: GridHeadersProps; slot: number }): ReactElement {
  const row = props.projection.order.displayToSource(slot), box = viewportHeaderBox(props.projection, props.window, props.viewport, 'row', slot)
  return <div role="rowheader" aria-colindex={1} aria-label={`Row ${slot + 1}`} aria-selected={selectedRow(props, row)}
    data-officeview-ui="row-header" onPointerDown={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()} style={{ position: 'absolute', width: GRID_HEADER_SIZE, height: '100%', pointerEvents: 'auto', background: '#f3f4f6' }}>
    <button type="button" aria-label={`Select row ${slot + 1}`} data-header-row={slot} tabIndex={props.controlMode ? 0 : -1} onMouseDown={keepGridFocus}
      onClick={event => props.onSelect?.('row', row, { toggle: event.ctrlKey || event.metaKey, extend: event.shiftKey }, event)}
      style={{ position: 'absolute', top: '50%', transform: 'translateY(-50%)', minWidth: GRID_HEADER_SIZE, minHeight: GRID_HEADER_SIZE, width: GRID_HEADER_SIZE, border: 0, padding: 0, background: 'transparent' }}><span style={{ display: 'block', height: box.height, overflow: 'hidden', fontSize: Math.min(14, box.height), lineHeight: `${box.height}px` }}>{slot + 1}</span></button>
  </div>
}
