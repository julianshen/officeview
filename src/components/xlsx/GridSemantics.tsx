import { useId, useLayoutEffect, useRef, useState, type KeyboardEvent, type MouseEvent, type ReactElement, type ReactNode } from 'react'
import type { CacheIdentity, GridSelection, SheetView, ScrollWindow, SourceRowInterval, ProfileId } from '../../xlsx/view/contracts'
import type { ProjectedSheetMetrics } from '../../xlsx/view/projection'
import type { SourceCell } from '../../xlsx/view/coordinates'
import { visibleSpans } from '../../xlsx/view/frozenPanes'
import type { AxisWindow } from '../../xlsx/view/contracts'
import { navigateKey, navigateTo, hitTestViewport, type NavigationKey, type ViewportSize } from '../../xlsx/view/scrollWindow'
import { MAX_ROWS, MAX_COLS, cellId, columnToLetters, formatA1 } from '../../xlsx/view/coordinates'
import { activeAnnouncement, boundedSelectionText, reconcileSelection, recordCounts, emptyGridSelection, selectCell, selectHeader } from '../../xlsx/view/selection'
import { belowDetail } from '../../xlsx/view/viewport'
import { GridHeaders, GridRowHeader, GRID_HEADER_SIZE, cellSelected, viewportHeaderBox, columnHeaderHeight } from './GridHeaders'
import { ViewerStatus, type ViewerStatusProps, type HiddenCellInspection } from './ViewerStatus'
import { useControlMedia } from './controlMedia'

export interface GridCellPresentation { text: string; state?: 'loading' | 'error' | 'unavailable' }
export interface GridTransition { selection: GridSelection; window: ScrollWindow; announcement: string }
export interface GridPopup { content: ReactNode; invoker: HTMLElement | null; close: () => void; /** Accessible name of the popup (default "Column filter"). */ label?: string }
export interface GridSemanticsProps {
  owner: CacheIdentity
  projection: ProjectedSheetMetrics
  view: Pick<SheetView, 'visibility' | 'baseHidden' | 'filter' | 'sort'>
  selection: GridSelection
  window: ScrollWindow
  viewport: ViewportSize
  body: SourceRowInterval | null
  onChange: (transition: GridTransition) => void
  cellPresentation?: (at: SourceCell) => GridCellPresentation
  announcement?: string
  popup?: GridPopup
  onFilter?: (col: number, invoker: HTMLElement) => void
  children?: ReactNode
  label?: string
  onSortCycle?: (col: number, shift: boolean, invoker: HTMLElement) => void
  /** Column whose header sort is pending: the grid is busy and its control shows it. */
  pendingSortColumn?: number | null
  onSortDialog?: (col: number, invoker: HTMLElement) => void
  onNotice?: (message: string) => void
  status?: Pick<ViewerStatusProps, 'records' | 'countMode' | 'loadReport' | 'inspection'>
  profile?: ProfileId
  renderMode?: 'detail' | 'coarse'
  /** Exact read-only inspection of a hidden Go To/find target (grid focus stays put). */
  inspection?: HiddenCellInspection
}
const navigationKeys = new Set<string>(['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Enter', 'Home', 'End', 'PageUp', 'PageDown'])
/** The parent commits viewport and selection together, so the new descendant
 * is represented in the same React publication as its ID. */
export function gridNavigationTransition(props: GridSemanticsProps, key: NavigationKey, options: { ctrl?: boolean; extend?: boolean } = {}): GridTransition {
  const selection = reconcileSelection(props.selection, props.projection, props.view).selection
  const result = navigateKey(props.window, { metrics: props.projection, viewport: props.viewport, view: { ...props.view, selection } }, key, selection.active, options)
  if (result.kind !== 'visible') return { window: result.window, selection, announcement: 'No visible records.' }
  const next = selectCell(selection, props.projection, result.target, { extend: options.extend })
  return { window: result.window, selection: next, announcement: targetAnnouncement(props.projection, next, result.target, props.cellPresentation?.(result.target).text, options.extend) }
}
/** The active-cell announcement; an extension also names what is selected
 * (spec §9.2: the live region reports selection truthfully), from the gesture
 * kind: "selected B2:C3" for cells, "selected columns A:C" for a column-header
 * extension, "selected rows 3:5" for rows, plus any additional (Ctrl) ranges. */
function targetAnnouncement(projection: ProjectedSheetMetrics, selection: GridSelection, target: SourceCell, text: string | undefined, extend?: boolean): string {
  const base = activeAnnouncement(projection, target, text)
  const anchor = selection.gesture?.anchor ?? selection.anchor
  if (!extend || !anchor) return base
  const kind = selection.gesture?.kind ?? 'cells'
  const span = (a: number, b: number, label: (n: number) => string) => a === b ? label(a) : `${label(Math.min(a, b))}:${label(Math.max(a, b))}`
  const extra = selection.ranges.length > 1 ? ` and ${selection.ranges.length - 1} more range${selection.ranges.length > 2 ? 's' : ''}` : ''
  if (kind === 'column') {
    const one = anchor.col === target.col
    return `${base}; selected column${one ? '' : 's'} ${span(anchor.col, target.col, columnToLetters)}${extra}`
  }
  if (kind === 'row') {
    const a = projection.order.sourceToDisplay(anchor.row) + 1, b = projection.order.sourceToDisplay(target.row) + 1
    return `${base}; selected row${a === b ? '' : 's'} ${span(a, b, String)}${extra}`
  }
  if (anchor.row === target.row && anchor.col === target.col) return base
  const display = (at: SourceCell) => formatA1({ row: projection.order.sourceToDisplay(at.row), col: at.col })
  return `${base}; selected ${display(anchor)}:${display(target)}${extra}`
}
/** More than one cell selected: the cells then describe themselves as selected
 * (a lone active cell is always aria-selected, so saying it would be noise). */
function multiCellSelection(selection: GridSelection): boolean {
  if (selection.ranges.length !== 1) return selection.ranges.length > 1
  const part = selection.ranges[0], rows = part.rows
  if (part.cols.length !== 1 || part.cols[0].c0 !== part.cols[0].c1) return true
  // Cardinality from each representation's own count, never by enumeration.
  const count = rows.kind === 'intervals' ? rows.intervals.reduce((n, i) => n + (i.r1 - i.r0 + 1), 0)
    : rows.kind === 'list' ? rows.rows.length : rows.count
  return count > 1
}
// Spec §12.3 frame/gridline admission; S5 passes the same profile/render mode.
const FRAME_LIMITS = { D: { cells: 20000, lines: 4096 }, M: { cells: 10000, lines: 2048 } } as const
/** FNV-1a: a short, stable, ID-safe token for arbitrary doc/sheet IDs. */
function shortHash(text: string): string {
  let hash = 2166136261
  for (let i = 0; i < text.length; i++) hash = Math.imul(hash ^ text.charCodeAt(i), 16777619)
  return (hash >>> 0).toString(36)
}
const hiddenText = { position: 'absolute', width: 1, height: 1, overflow: 'hidden', clipPath: 'inset(50%)', whiteSpace: 'nowrap' } as const
/** S5 supplies current presentation snapshots and mounts its bounded canvas in
 * children. This shell never fetches values or renders legacy worksheet paint. */
/** Visible runs of an axis over one or two disjoint logical spans (frozen strip, body). */
function spanWindow(axis: { visibleBetween(start: number, end: number): AxisWindow }, spans: ReadonlyArray<[number, number]>): AxisWindow {
  if (spans.length === 1) return axis.visibleBetween(spans[0][0], spans[0][1])
  const windows = spans.map(([start, end]) => axis.visibleBetween(start, end))
  return { runs: windows.flatMap(window => [...window.runs]), count: windows.reduce((sum, window) => sum + window.count, 0) }
}
export function GridSemantics(props: GridSemanticsProps): ReactElement {
  const instance = useId(), { projection, window: scroll, viewport, view } = props
  const { docId, sheetId, valueGen, fontGen, viewGen } = props.owner
  const touchControls = useControlMedia('(pointer: coarse)')
  const headerHeight = columnHeaderHeight(touchControls, Boolean(props.onSortCycle || props.onFilter || props.onSortDialog), Boolean(props.onSortDialog))
  // Value/font republish keeps the same cells, so IDs omit those generations;
  // the reset key still closes popups/control mode on any owner change.
  const prefix = `${instance}-${shortHash(docId)}-${shortHash(sheetId)}-${viewGen}`
  const ownerKey = JSON.stringify([docId, sheetId, valueGen, fontGen, viewGen])
  const wrapperRef = useRef<HTMLDivElement>(null), gridRef = useRef<HTMLDivElement>(null), popupRef = useRef<HTMLDivElement>(null), previousPopup = useRef<{ popup: GridPopup; owner: string } | null>(null)
  const previousOwner = useRef(ownerKey)
  const [controlMode, setControlModeState] = useState(false), [notice, setNotice] = useState('')
  // Focus handlers run before React re-renders, so they read the mode from a ref.
  const controlModeRef = useRef(false)
  const setControlMode = (on: boolean) => { controlModeRef.current = on; setControlModeState(on) }
  useLayoutEffect(() => {
    if (previousOwner.current !== ownerKey) {
      if (gridRef.current?.contains(document.activeElement) || popupRef.current?.contains(document.activeElement)) gridRef.current?.focus()
      props.popup?.close()
      setControlMode(false)
      setNotice('')
      previousOwner.current = ownerKey
    }
    // A removed popup leaves focus on body; focus the user already moved
    // outside the viewer is theirs to keep.
    const focus = document.activeElement
    const focusOwned = !focus || focus === document.body || Boolean(wrapperRef.current?.contains(focus))
    if (previousPopup.current && !props.popup && focusOwned) {
      const { popup, owner } = previousPopup.current, invoker = popup.invoker
      if (owner === ownerKey && invoker?.isConnected && gridRef.current?.contains(invoker)) {
        // A header invoker is a native control, so restoring it re-enters control mode.
        if (invoker !== gridRef.current) setControlMode(true)
        invoker.focus()
      } else gridRef.current?.focus()
    }
    previousPopup.current = props.popup ? { popup: props.popup, owner: previousPopup.current?.owner ?? ownerKey } : null
  }, [ownerKey, props.popup])
  const current = reconcileSelection(props.selection, projection, view), selection = current.selection
  const multiSelected = multiCellSelection(selection)
  // Frozen panes: the pinned strip and the scrolling body are separate logical spans.
  const spans = visibleSpans(scroll.frozen, scroll.logicalOrigin, scroll.zoom, viewport)
  const rowWindow = spanWindow(projection.rows, spans.y), colWindow = spanWindow(projection.cols, spans.x)
  const limits = FRAME_LIMITS[props.profile ?? 'D']
  let coarse = props.renderMode === 'coarse' || rowWindow.count * colWindow.count >= limits.cells || rowWindow.count + colWindow.count >= limits.lines
  let slots: number[] = [], columns: number[] = []
  let subDetail = false
  if (!coarse) {
    // A line below the painter's detail rule (belowDetail: 4 device px) is
    // painted as a coarse pixel group (§11.1) and owns no row/cell here;
    // every other line keeps its owned semantics. One
    // tiny authored row must not strip the whole viewport (A11Y-01). The window
    // counts above already bound how many lines this can visit.
    const scale = scroll.zoom * scroll.deviceScale // the painter's scale expression (viewport.ts detailScale)
    const detail = (size: number) => !belowDetail(size, scale)
    for (const run of rowWindow.runs) for (let slot = run.start; slot <= run.end; slot++) if (detail(projection.rows.size(slot))) slots.push(slot); else subDetail = true
    for (const run of colWindow.runs) for (let col = run.start; col <= run.end; col++) if (detail(projection.cols.size(col))) columns.push(col); else subDetail = true
    if (!slots.length || !columns.length) { coarse = true; slots = []; columns = [] }
  }
  const active = selection.active && projection.sourceCellToLogical(selection.active) ? selection.active : null
  const activeSlot = active ? projection.order.sourceToDisplay(active.row) : null
  const transient = activeSlot !== null && !slots.includes(activeSlot)
  if (transient) slots.push(activeSlot!)
  slots.sort((a, b) => a - b)
  const publish = (action: () => GridTransition) => {
    try { props.onChange(action()); setNotice('') }
    catch (error) { setNotice(error instanceof Error ? boundedSelectionText(error.message) : 'Navigation unavailable. Please retry.') }
  }
  const keyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') return
    const grid = gridRef.current
    if (event.key === 'Escape' && props.popup && (grid?.contains(document.activeElement) || popupRef.current?.contains(document.activeElement))) {
      event.preventDefault(); event.stopPropagation()
      props.popup.close()
      return
    }
    if (event.target !== grid) {
      if (event.key === 'Escape' && grid?.contains(event.target as Node)) {
        event.preventDefault(); setControlMode(false); grid.focus()
      }
      event.stopPropagation()
      return
    }
    if (document.activeElement !== grid) return
    if (event.key === 'Escape') {
      if (!selection.ranges.length && !selection.active && !selection.anchor) return
      event.preventDefault(); event.stopPropagation()
      publish(() => ({ selection: emptyGridSelection(), window: scroll, announcement: 'Selection cleared.' }))
    } else if (event.key === 'F2') {
      const buttons = [...grid.querySelectorAll<HTMLButtonElement>('[data-header-column]')]
      const button = buttons.find(button => Number(button.dataset.headerColumn) === active?.col) ?? buttons[0]
      if (button) { event.preventDefault(); event.stopPropagation(); setControlMode(true); button.focus() }
    } else if (event.altKey && event.key === 'ArrowDown' && active && props.onFilter) {
      event.preventDefault(); event.stopPropagation(); props.onFilter(active.col, grid)
    } else if (navigationKeys.has(event.key) && !event.altKey) {
      event.preventDefault(); event.stopPropagation()
      publish(() => gridNavigationTransition(props, event.key as NavigationKey, { ctrl: event.ctrlKey || event.metaKey, extend: event.shiftKey }))
    }
  }
  const activeBox = active && projection.sourceCellToLogical(active)
  const activeColumnInPane = activeBox && spans.x.some(([x0, x1]) => activeBox.x < x1 && activeBox.x + activeBox.width > x0)
  const headerColumns = coarse && active && activeColumnInPane ? [active.col] : columns
  const headerProps = { projection, window: scroll, viewport, coarse, columns: headerColumns, selection, sort: view.sort, filter: view.filter, controlMode, touchControls, onSortCycle: props.onSortCycle, pendingSortColumn: props.pendingSortColumn ?? null, onFilter: props.onFilter, onSortDialog: props.onSortDialog, onNotice: props.onNotice,
    onSelect: (kind: 'row' | 'column', index: number, options: { toggle: boolean; extend: boolean }, event: MouseEvent<HTMLButtonElement>) => {
      // A corner-lane button (S13.2) names its column; only strip buttons are re-hit-tested.
      if (event.detail > 0 && !coarse && gridRef.current && event.currentTarget.dataset.headerLane === undefined) {
        const bounds = gridRef.current.getBoundingClientRect()
        const point = { x: event.clientX - bounds.left - GRID_HEADER_SIZE, y: event.clientY - bounds.top - headerHeight }
        if (kind === 'row' ? point.y < 0 || point.y >= viewport.height : point.x < 0 || point.x >= viewport.width) return
        const target = hitTestViewport(scroll, projection, viewport, { x: kind === 'row' ? 0 : point.x, y: kind === 'row' ? point.y : 0 })
        let hit: number | undefined
        if (target) hit = kind === 'row' ? projection.order.sourceToDisplay(target.row) : target.col
        else {
          // A header can still select when its orthogonal data axis is all hidden.
          hit = (kind === 'row' ? slots : headerColumns).find(candidate => {
            const box = viewportHeaderBox(projection, scroll, viewport, kind, candidate)
            return kind === 'row' ? point.y >= box.y && point.y < box.y + box.height : point.x >= box.x && point.x < box.x + box.width
          })
        }
        if (hit === undefined) return
        index = kind === 'row' ? projection.order.displayToSource(hit) : hit
        // Pointer selection stays in navigation mode: arrows continue from the grid.
        gridRef.current.focus()
      }
      publish(() => {
        const selected = selectHeader(selection, projection, kind, index, options.toggle, options.extend)
        const result = selected.active ? navigateTo(scroll, { metrics: projection, viewport, view: { ...view, selection: selected } }, { addressSpace: 'source', at: selected.active }) : null
        return { selection: selected, window: result?.window ?? scroll, announcement: selected.active ? activeAnnouncement(projection, selected.active) : 'No visible records.' }
      })
    },
    onSelectAll: () => publish(() => ({ selection: { active, anchor: active, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: projection.rows.logicalCount - 1 }] }, cols: [{ c0: 0, c1: projection.cols.logicalCount - 1 }] }] }, window: scroll, announcement: 'Worksheet selected.' })),
  }
  const status = [props.announcement, current.announcement, notice, coarse || subDetail ? 'Zoom in or inspect a cell for detail.' : ''].filter(Boolean).join(' ')
  return <div ref={wrapperRef} onKeyDown={keyDown} onBlur={event => { if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setControlMode(false) }} style={{ position: 'relative', width: viewport.width + GRID_HEADER_SIZE, height: viewport.height + headerHeight }}>
    <div ref={gridRef} role="grid" aria-label={props.label ?? 'Worksheet'} aria-busy={props.pendingSortColumn != null ? true : undefined} aria-rowcount={MAX_ROWS + 1} aria-colcount={MAX_COLS + 1}
      aria-multiselectable="true" aria-activedescendant={active ? `${prefix}-cell-${cellId(active.row, active.col)}` : undefined}
      aria-description={`Worksheet capacity ${MAX_ROWS} rows by ${MAX_COLS} columns.${coarse ? ' Coarse view. Zoom in or inspect a cell for detail.' : ''}`}
      tabIndex={controlMode ? -1 : 0} data-officeview-grid=""
      onFocusCapture={event => {
        // Only F2 or a restored header invoker enters control mode; incidental
        // focus (mousedown, screen-reader cursor) returns to the single grid stop.
        if (event.target === event.currentTarget) { if (controlModeRef.current) setControlMode(false) }
        else if (!controlModeRef.current) event.currentTarget.focus()
      }}
      onPointerDown={event => { event.stopPropagation(); if (event.target === event.currentTarget) event.currentTarget.focus() }}
      onDoubleClick={event => event.stopPropagation()}
      onClick={event => {
        if (event.defaultPrevented || event.detail > 1 || (event.target as Element).closest('[data-officeview-ui]')) return
        const bounds = event.currentTarget.getBoundingClientRect()
        const target = hitTestViewport(scroll, projection, viewport, { x: event.clientX - bounds.left - GRID_HEADER_SIZE, y: event.clientY - bounds.top - headerHeight })
        if (!target) return
        event.currentTarget.focus()
        publish(() => {
          const next = selectCell(selection, projection, target, { extend: event.shiftKey, toggle: event.ctrlKey || event.metaKey })
          return { selection: next, window: scroll, announcement: targetAnnouncement(projection, next, target, props.cellPresentation?.(target).text, event.shiftKey) }
        })
      }} style={{ position: 'relative', width: '100%', height: '100%', overflow: 'hidden' }}>
      <div aria-hidden="true" style={{ position: 'absolute', left: GRID_HEADER_SIZE, top: headerHeight, width: viewport.width, height: viewport.height }}>{props.children}</div>
      <GridHeaders {...headerProps} />
      {slots.map(slot => {
        const row = projection.order.displayToSource(slot), extra = transient && slot === activeSlot
        const cells = extra ? [active!.col] : [...columns]
        if (!extra && slot === activeSlot && active && !cells.includes(active.col)) cells.push(active.col)
        const box = viewportHeaderBox(projection, scroll, viewport, 'row', slot)
        // Rows and header controls follow the scroll window by transform: a
        // transform is not a layout shift, so the viewer adds no CLS to its host.
        // Cells keep `left` inside their row (per-cell transforms would make every
        // cell its own paint layer, ~1 KB each; their moves register no shift).
        return <div key={slot} role="row" aria-rowindex={slot + 2} data-transient-active={extra ? 'true' : undefined}
          style={extra ? { ...hiddenText, top: 0, left: 0, pointerEvents: 'none' } : { position: 'absolute', top: 0, left: 0, transform: `translateY(${headerHeight + box.y}px)`, height: box.height, width: '100%', pointerEvents: 'none' }}>
          <GridRowHeader props={headerProps} slot={slot} />
          {cells.map(col => {
            const at = { row, col }, value = props.cellPresentation?.(at) ?? { text: '', state: 'loading' as const }
            const column = viewportHeaderBox(projection, scroll, viewport, 'column', col)
            const offscreen = extra || (slot === activeSlot && active?.col === col && !columns.includes(col))
            const isSelected = cellSelected(selection.ranges, at)
            // Name from contents: the visually hidden text IS the name, so VoiceOver
            // reads the value once (no label + contents) and never "blank" (S13 V1/V5).
            return <div key={col} role="gridcell" id={`${prefix}-cell-${cellId(row, col)}`} aria-colindex={col + 2}
              aria-selected={isSelected} aria-description={[value.state, multiSelected && isSelected ? 'selected' : ''].filter(Boolean).join(', ') || undefined}
              style={offscreen ? { ...hiddenText, top: 0, left: 0 } : { position: 'absolute', top: 0, left: GRID_HEADER_SIZE + column.x, width: column.width, height: box.height }}><span style={hiddenText}>{activeAnnouncement(projection, at, value.text)}</span></div>
          })}
        </div>
      })}
    </div>
    {/* Bounded by the pane (overflow:hidden): a taller popup scrolls instead of losing its top controls.
        Anchored at the top: content that arrives later (a loading value list) grows the popup downward,
        never moving what is already shown (no layout shift for the host page). */}
    {props.popup && <div ref={popupRef} role="dialog" aria-label={props.popup.label ?? 'Column filter'} data-officeview-ui="grid-popup"
      onPointerDown={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()}
      style={{ position: 'absolute', zIndex: 3, top: 0, maxWidth: '100%', maxHeight: '100%', overflowY: 'auto', boxSizing: 'border-box', paddingBottom: 'env(safe-area-inset-bottom)' }}>{props.popup.content}</div>}
    <ViewerStatus {...props.status} records={props.status?.records ?? recordCounts(projection, props.body)} announcement={status} inspection={props.inspection ?? props.status?.inspection} />
  </div>
}
