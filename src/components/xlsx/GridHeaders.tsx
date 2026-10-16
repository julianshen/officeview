import { useEffect, useId, useRef, useState, type CSSProperties, type MouseEvent, type ReactElement } from 'react'
import type { GridSelection, ScrollWindow, SortSpec, FilterSpec, ColumnRule } from '../../xlsx/view/contracts'
import type { ProjectedSheetMetrics } from '../../xlsx/view/projection'
import { columnToLetters } from '../../xlsx/view/coordinates'
import { logicalToViewport, type ViewportSize } from '../../xlsx/view/scrollWindow'
import { rowContains, boundedSelectionText } from '../../xlsx/view/selection'
import { useControlMedia } from './controlMedia'

export const GRID_HEADER_SIZE = 44
/** Header strip palette (WCAG AA on `background`: text/action/active >= 4.5:1, track >= 3:1). */
export const HEADER_COLORS = Object.freeze({ background: '#f3f4f6', text: '#1f2937', action: '#374151', active: '#1459a6', track: '#858b94' })
/** Desktop header layout: label font, per-letter estimate and one action glyph's width (CSS px). */
const LABEL_FONT_PX = 12, LABEL_LETTER_PX = 7, LABEL_PAD_PX = 4, ACTION_PX = 14
/** Row-number label font and per-digit width (tabular figures). */
const ROW_FONT_PX = 12, ROW_DIGIT_PX = 7.2, ROW_PAD_PX = 10
/** Row-header gutter wide enough for the largest visible row number (seven digits at row 1048576). */
export function rowGutterWidth(lastLabel: number): number {
  return Math.max(GRID_HEADER_SIZE, Math.ceil(String(Math.max(1, Math.floor(lastLabel))).length * ROW_DIGIT_PX + ROW_PAD_PX))
}
/** Visually hidden yet focusable and named (a compact header's actions). */
const visuallyHidden: CSSProperties = { position: 'absolute', width: 1, height: 1, minWidth: 0, minHeight: 0, margin: -1, padding: 0, border: 0, overflow: 'hidden', clip: 'rect(0 0 0 0)', whiteSpace: 'nowrap' }
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
  touchControls?: boolean
  onSelect?: (kind: 'row' | 'column', index: number, options: HeaderSelectionOptions, event: MouseEvent<HTMLButtonElement>) => void
  onSelectAll?: () => void
  controlColumn?: number | null
  filter?: FilterSpec | null
  onSortCycle?: (col: number, shift: boolean, invoker: HTMLButtonElement) => void
  pendingSortColumn?: number | null
  onFilter?: (col: number, invoker: HTMLButtonElement) => void
  onSortDialog?: (col: number, invoker: HTMLButtonElement) => void
  onNotice?: (message: string) => void
  /** Row-header gutter width (CSS px); default GRID_HEADER_SIZE. */
  gutter?: number
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
export function columnRuleDescription(rule: ColumnRule): string {
  if (rule.kind === 'compound') return boundedSelectionText(`${columnRuleDescription(rule.first)} ${rule.connective.toUpperCase()} ${columnRuleDescription(rule.second)}`)
  if (rule.kind === 'values') return rule.selection.mode === 'allExcept' ? `All except ${rule.selection.excluded.count} values` : rule.selection.mode === 'include' ? `${rule.selection.keys.count} included values` : 'Complete value-query selection'
  if (rule.kind === 'datePreset') return `Date period: ${rule.preset}`
  const op: Record<string, string> = { gt: 'Greater than', gte: 'Greater than or equal', lt: 'Less than', lte: 'Less than or equal', eq: 'Equals', neq: 'Does not equal', between: 'Between', contains: 'Contains', 'not-contains': 'Does not contain', begins: 'Begins with', ends: 'Ends with', before: 'Before', after: 'After' }
  const value = (input: typeof rule.first) => typeof input === 'object'
    ? `${String(input.year).padStart(4, '0')}-${String(input.month).padStart(2, '0')}-${String(input.day).padStart(2, '0')}` : String(input)
  return boundedSelectionText(`${op[rule.op]} ${value(rule.first)}${'second' in rule && rule.second !== undefined ? ` through ${value(rule.second)}` : ''}`)
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
  const coarsePointer = useControlMedia('(pointer: coarse)'), touch = props.touchControls ?? coarsePointer
  const actions = Boolean(props.onSortCycle || props.onFilter || props.onSortDialog)
  const height = columnHeaderHeight(touch, actions, Boolean(props.onSortDialog))
  const gutter = props.gutter ?? GRID_HEADER_SIZE
  const target = { minHeight: GRID_HEADER_SIZE, minWidth: touch ? GRID_HEADER_SIZE : 0, overflow: 'hidden', padding: 0, border: 0, background: 'transparent', color: HEADER_COLORS.text }
  // Desktop: a column too narrow for its label plus the action glyphs (or any
  // column at <= 50% zoom) shows its label only; its actions stay named and
  // focusable (control mode), and focus reveals them over the neighbours.
  const [focusCol, setFocusCol] = useState<number | null>(null)
  // Spec §9.3: 44px minimum control targets on mobile, with deterministic
  // ownership. A column narrower than 44px cannot hold 44px targets without
  // overlapping a neighbour's, so on a coarse pointer the active cell's narrow
  // column gets a corner-lane disclosure ("Actions for column B", under
  // "Select worksheet", in a lane no column owns) that opens its select, sort,
  // filter and more-sort buttons as native ≥44x44 targets. The disclosure is
  // a semantic button (aria-expanded/controls); Escape, an action or a tap
  // outside closes it, and focus returns to it in control mode.
  const narrowTouch = (col: number) => touch && viewportHeaderBox(props.projection, props.window, props.viewport, 'column', col).width < GRID_HEADER_SIZE
  const activeCol = props.selection.active?.col
  const laneCol = touch && (actions || props.onSelect) && activeCol !== undefined && props.columns.includes(activeCol) && narrowTouch(activeCol) ? activeCol : undefined
  const [openLane, setOpenLane] = useState<number | null>(null)
  // The group belongs to one lane column: another active column closes it, and
  // returning to the first one does not reopen it (state adjusted during render).
  const [laneFor, setLaneFor] = useState(laneCol)
  if (laneFor !== laneCol) { setLaneFor(laneCol); setOpenLane(null) }
  const laneOpen = laneCol !== undefined && laneFor === laneCol && openLane === laneCol
  const laneButton = useRef<HTMLButtonElement | null>(null), laneGroup = useRef<HTMLDivElement | null>(null)
  const laneId = `${useId()}-lane`
  const closeLane = (refocus: boolean) => {
    setOpenLane(null)
    if (refocus && props.controlMode) laneButton.current?.focus()
  }
  useEffect(() => {
    if (!laneOpen) return
    // Control mode moves focus into the opened group; pointer use leaves it on the grid.
    if (props.controlMode) laneGroup.current?.querySelector('button')?.focus()
    const outside = (event: Event) => {
      const target = event.target as Node | null
      if (target && !laneGroup.current?.contains(target) && !laneButton.current?.contains(target)) setOpenLane(null)
    }
    document.addEventListener('pointerdown', outside, true)
    return () => document.removeEventListener('pointerdown', outside, true)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [laneOpen])
  const laneTarget = (index: number): CSSProperties => ({ ...target, position: 'absolute', left: 0, top: GRID_HEADER_SIZE * index, width: GRID_HEADER_SIZE, height: GRID_HEADER_SIZE, background: '#ffffff' })
  return <div role="row" aria-rowindex={1} data-officeview-ui="grid-headers" onPointerDown={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()} style={{ position: 'absolute', top: 0, left: 0, height, width: '100%', zIndex: 2, background: HEADER_COLORS.background, boxShadow: `inset 0 -1px ${HEADER_COLORS.track}` }}>
    <div role="columnheader" aria-colindex={1} aria-label="Worksheet corner" style={{ position: 'absolute', left: 0, width: gutter, height: laneCol === undefined ? GRID_HEADER_SIZE : height, zIndex: 1 }}>
      <button type="button" aria-label="Select worksheet" tabIndex={props.controlMode ? 0 : -1} onMouseDown={keepGridFocus} onClick={props.onSelectAll} style={{ width: gutter, height: GRID_HEADER_SIZE, color: HEADER_COLORS.text }}>↖</button>
      {laneCol !== undefined && <>
        <button ref={laneButton} type="button" data-header-lane={laneCol} tabIndex={props.controlMode ? 0 : -1} onMouseDown={keepGridFocus}
          aria-label={`Actions for column ${columnToLetters(laneCol)}`} aria-expanded={laneOpen} aria-controls={laneId}
          onClick={() => setOpenLane(laneOpen ? null : laneCol)} onKeyDown={event => { if (event.key === 'Escape' && laneOpen) { event.preventDefault(); event.stopPropagation(); closeLane(true) } }}
          style={{ ...target, position: 'absolute', left: 0, top: GRID_HEADER_SIZE, width: GRID_HEADER_SIZE, height: GRID_HEADER_SIZE }}>{columnToLetters(laneCol)}⋯</button>
        {laneOpen && <div ref={laneGroup} id={laneId} role="group" aria-label={`Column ${columnToLetters(laneCol)} actions`} data-officeview-ui="header-lane"
          onKeyDown={event => { if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); closeLane(true) } }}
          onClickCapture={() => { setOpenLane(null) }}
          style={{ position: 'absolute', left: 0, top: GRID_HEADER_SIZE * 2, width: GRID_HEADER_SIZE, zIndex: 3, boxShadow: '0 2px 6px rgba(0,0,0,.25)' }}>
          {props.onSelect && <button type="button" data-header-column={laneCol} data-header-lane={laneCol} tabIndex={props.controlMode ? 0 : -1} onMouseDown={keepGridFocus}
            aria-label={`Select whole column ${columnToLetters(laneCol)}`} style={laneTarget(0)}
            onClick={event => props.onSelect?.('column', laneCol, { toggle: event.ctrlKey || event.metaKey, extend: event.shiftKey }, event)}>{columnToLetters(laneCol)}</button>}
          {actions && columnActions(props, laneCol, index => laneTarget(index + (props.onSelect ? 1 : 0)))}
        </div>}
      </>}
    </div>
    {props.columns.map(col => {
      const box = viewportHeaderBox(props.projection, props.window, props.viewport, 'column', col), priority = props.sort?.keys.findIndex(key => key.col === col) ?? -1
      const direction = priority < 0 ? undefined : props.sort!.keys[priority].order === 'asc' ? 'ascending' : 'descending'
      if (props.coarse) box.width = Math.min(GRID_HEADER_SIZE, props.viewport.width - box.x)
      const selected = selectedColumn(props, col), label = `Column ${columnToLetters(col)}${direction ? `; sort priority ${priority + 1}, ${direction}` : ''}`
      const rule = props.filter?.columns.get(col), description = rule ? columnRuleDescription(rule) : 'No viewer filter'
      const actions = Boolean(props.onSortCycle || props.onFilter || props.onSortDialog)
      // A touch target keeps its 44px floor only while the column is that wide;
      // narrower, it is clamped (and the header clips) so a neighbour's control
      // never overlaps and steals the hit (spec §9.3 deterministic ownership).
      const letters = columnToLetters(col), labelPx = letters.length * LABEL_LETTER_PX + LABEL_PAD_PX
      const actionsPx = (props.onSortDialog ? 3 : 2) * ACTION_PX
      const compact = !touch && actions && (props.window.zoom <= 0.5 || box.width < labelPx + actionsPx)
      const revealed = compact && focusCol === col
      const width = revealed ? Math.max(box.width, labelPx + actionsPx) : box.width
      // A label wider than its column shrinks instead of truncating ("XFD", not "X..").
      const labelFont = touch || box.width >= labelPx ? LABEL_FONT_PX : Math.max(6, (box.width - 2) / (letters.length * LABEL_LETTER_PX / LABEL_FONT_PX))
      const cellTarget = touch ? { ...target, minWidth: Math.min(GRID_HEADER_SIZE, box.width), width: box.width, flex: 'none' } : { ...target, flex: '1 1 auto', minWidth: 0 }
      const actionTarget: CSSProperties = touch ? cellTarget : compact && !revealed ? { ...target, ...visuallyHidden } : { ...target, flex: 'none', width: ACTION_PX, color: HEADER_COLORS.action }
      const iconClip = { display: 'block', width: touch ? box.width : undefined, overflow: 'hidden', textAlign: 'center' as const }
      return <div key={col} role="columnheader" aria-colindex={col + 2} aria-label={label} aria-selected={selected} aria-sort={priority === 0 ? direction : undefined}
        onFocus={() => setFocusCol(col)} onBlur={event => { if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setFocusCol(current => current === col ? null : current) }}
        style={{ position: 'absolute', left: 0, transform: `translateX(${gutter + box.x}px)`, width, height, overflow: 'hidden', display: 'flex', flexDirection: touch ? 'column' : 'row', alignItems: 'stretch', boxSizing: 'border-box', boxShadow: `inset -1px 0 ${HEADER_COLORS.track}`, ...(revealed ? { zIndex: 3, background: HEADER_COLORS.background } : {}) }}>
        <button type="button" aria-label={`Select ${label}`} data-header-column={col} tabIndex={props.controlMode ? 0 : -1} onMouseDown={keepGridFocus}
          aria-description={actions ? description : undefined} title={`${label}; ${description}`}
          onClick={event => { props.onNotice?.(`${label}; ${description}`); props.onSelect?.('column', col, { toggle: event.ctrlKey || event.metaKey, extend: event.shiftKey }, event) }}
          style={cellTarget}><span style={{ ...iconClip, whiteSpace: 'nowrap', fontSize: labelFont }}>{letters}</span></button>
        {actions && !narrowTouch(col) && columnActions(props, col, () => actionTarget, iconClip)}
      </div>
    })}
  </div>
}
/** Sort / filter / more-sort buttons of one column; `style(i)` places the i-th. */
function columnActions(props: GridHeadersProps, col: number, style: (index: number) => CSSProperties, iconClip: CSSProperties = { display: 'block', overflow: 'hidden', textAlign: 'center' }): ReactElement {
  const priority = props.sort?.keys.findIndex(key => key.col === col) ?? -1
  const direction = priority < 0 ? undefined : props.sort!.keys[priority].order === 'asc' ? 'ascending' : 'descending'
  const label = `Column ${columnToLetters(col)}${direction ? `; sort priority ${priority + 1}, ${direction}` : ''}`
  const rule = props.filter?.columns.get(col), description = rule ? columnRuleDescription(rule) : 'No viewer filter'
  return <>
    <button type="button" data-action-column={col} tabIndex={props.controlMode ? 0 : -1} onMouseDown={keepGridFocus} style={{ ...style(0), ...(direction ? { color: HEADER_COLORS.active } : {}) }} disabled={!props.onSortCycle} aria-label={`Cycle sort column ${columnToLetters(col)}`} aria-busy={props.pendingSortColumn === col ? true : undefined} aria-description={props.pendingSortColumn === col ? 'sorting' : direction ?? 'unsorted'} aria-pressed={Boolean(direction)} title={`${label}; ${props.pendingSortColumn === col ? 'sorting' : direction ?? 'unsorted'}`} onClick={event => props.onSortCycle?.(col, event.shiftKey, event.currentTarget)}><span style={iconClip}>{props.pendingSortColumn === col ? '…' : direction === 'ascending' ? '↑' : direction === 'descending' ? '↓' : '↕'}</span></button>
    <button type="button" data-action-column={col} tabIndex={props.controlMode ? 0 : -1} onMouseDown={keepGridFocus} style={{ ...style(1), ...(rule ? { color: HEADER_COLORS.active } : {}) }} disabled={!props.onFilter} aria-label={`Filter column ${columnToLetters(col)} — ${description}`} aria-pressed={Boolean(rule)} title={description} onClick={event => { props.onNotice?.(`${label}; ${description}`); props.onFilter?.(col, event.currentTarget) }}><span style={iconClip}><svg aria-hidden="true" width="12" height="14" viewBox="0 0 12 14"><path d="M1 1h10L7 6v6l-2-1V6z" fill={rule ? 'currentColor' : 'none'} stroke="currentColor" /></svg></span></button>
    {props.onSortDialog && <button type="button" data-action-column={col} tabIndex={props.controlMode ? 0 : -1} onMouseDown={keepGridFocus} style={style(2)} aria-label={`More sort options for ${columnToLetters(col)}`} onClick={event => props.onSortDialog?.(col, event.currentTarget)}><span style={iconClip}>⋯</span></button>}
  </>
}
export function columnHeaderHeight(touch: boolean, actions: boolean, more: boolean): number {
  return GRID_HEADER_SIZE * (touch && actions ? more ? 4 : 3 : 1)
}
/** `cancelX`: the row's horizontal transform, undone so the header stays in the gutter. */
export function GridRowHeader({ props, slot, cancelX = 0 }: { props: GridHeadersProps; slot: number; cancelX?: number }): ReactElement {
  const row = props.projection.order.displayToSource(slot), box = viewportHeaderBox(props.projection, props.window, props.viewport, 'row', slot)
  const gutter = props.gutter ?? GRID_HEADER_SIZE
  // A row cut by the pane edge keeps its label where the whole row would put it
  // (no re-centring as it scrolls in: that would be a layout shift); a row taller
  // than the pane centres in its visible part.
  const full = gridHeaderBox(props.projection, props.window, 'row', slot)
  const label = full.height <= props.viewport.height ? { top: full.y - box.y, height: full.height } : { top: 0, height: box.height }
  return <div role="rowheader" aria-colindex={1} aria-label={`Row ${slot + 1}`} aria-selected={selectedRow(props, row)}
    data-officeview-ui="row-header" onPointerDown={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()} style={{ position: 'absolute', width: gutter, height: '100%', pointerEvents: 'auto', background: HEADER_COLORS.background, boxShadow: `inset -1px 0 ${HEADER_COLORS.track}`, ...(cancelX ? { transform: `translateX(${-cancelX}px)` } : {}) }}>
    <button type="button" aria-label={`Select row ${slot + 1}`} data-header-row={slot} tabIndex={props.controlMode ? 0 : -1} onMouseDown={keepGridFocus}
      onClick={event => props.onSelect?.('row', row, { toggle: event.ctrlKey || event.metaKey, extend: event.shiftKey }, event)}
      style={{ position: 'absolute', top: label.top + label.height / 2, transform: 'translateY(-50%)', minWidth: GRID_HEADER_SIZE, minHeight: GRID_HEADER_SIZE, width: gutter, border: 0, padding: 0, background: 'transparent', color: HEADER_COLORS.text }}><span style={{ display: 'block', height: label.height, overflow: 'hidden', whiteSpace: 'nowrap', fontVariantNumeric: 'tabular-nums', fontSize: Math.min(ROW_FONT_PX, label.height), lineHeight: `${label.height}px` }}>{slot + 1}</span></button>
  </div>
}
