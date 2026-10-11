import { useId, useLayoutEffect, useRef, useState, type KeyboardEvent, type MouseEvent, type ReactElement, type ReactNode } from 'react'
import type { CacheIdentity, GridSelection, SheetView, ScrollWindow, SourceRowInterval, ProfileId } from '../../xlsx/view/contracts'
import type { ProjectedSheetMetrics } from '../../xlsx/view/projection'
import type { SourceCell } from '../../xlsx/view/coordinates'
import { visibleLogicalViewport, navigateKey, navigateTo, hitTestViewport, type NavigationKey, type ViewportSize } from '../../xlsx/view/scrollWindow'
import { MAX_ROWS, MAX_COLS, cellId } from '../../xlsx/view/coordinates'
import { activeAnnouncement, boundedSelectionText, reconcileSelection, recordCounts, SourceSelection, emptyGridSelection, selectCell, selectHeader } from '../../xlsx/view/selection'
import { GridHeaders, GridRowHeader, GRID_HEADER_SIZE, viewportHeaderBox } from './GridHeaders'
import { ViewerStatus, type ViewerStatusProps } from './ViewerStatus'

export interface GridCellPresentation { text: string; state?: 'loading' | 'error' | 'unavailable' }
export interface GridTransition { selection: GridSelection; window: ScrollWindow; announcement: string }
export interface GridPopup { content: ReactNode; invoker: HTMLElement | null; close: () => void }
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
  onSortDialog?: (col: number, invoker: HTMLElement) => void
  onNotice?: (message: string) => void
  status?: Pick<ViewerStatusProps, 'records' | 'countMode' | 'loadReport' | 'inspection'>
  profile?: ProfileId
  renderMode?: 'detail' | 'coarse'
}
const navigationKeys = new Set<string>(['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Enter', 'Home', 'End', 'PageUp', 'PageDown'])
/** The parent commits viewport and selection together, so the new descendant
 * is represented in the same React publication as its ID. */
export function gridNavigationTransition(props: GridSemanticsProps, key: NavigationKey, options: { ctrl?: boolean; extend?: boolean } = {}): GridTransition {
  const selection = reconcileSelection(props.selection, props.projection, props.view).selection
  const result = navigateKey(props.window, { metrics: props.projection, viewport: props.viewport, view: { ...props.view, selection } }, key, selection.active, options)
  if (result.kind !== 'visible') return { window: result.window, selection, announcement: 'No visible records.' }
  return { window: result.window, selection: selectCell(selection, props.projection, result.target, { extend: options.extend }),
    announcement: activeAnnouncement(props.projection, result.target, props.cellPresentation?.(result.target).text) }
}
// Spec §12.3 frame/gridline admission; S5 passes the same profile/render mode.
const FRAME_LIMITS = { D: { cells: 20000, lines: 4096 }, M: { cells: 10000, lines: 2048 } } as const
const hiddenText = { position: 'absolute', width: 1, height: 1, overflow: 'hidden', clipPath: 'inset(50%)', whiteSpace: 'nowrap' } as const
/** S5 supplies current presentation snapshots and mounts its bounded canvas in
 * children. This shell never fetches values or renders legacy worksheet paint. */
export function GridSemantics(props: GridSemanticsProps): ReactElement {
  const instance = useId(), { projection, window, viewport, view } = props
  const { docId, sheetId, valueGen, fontGen, viewGen } = props.owner
  const prefix = `${instance}-${encodeURIComponent(JSON.stringify([docId, sheetId, valueGen, fontGen, viewGen]))}`
  const gridRef = useRef<HTMLDivElement>(null), popupRef = useRef<HTMLDivElement>(null), previousPopup = useRef<{ popup: GridPopup; owner: string } | null>(null)
  const previousOwner = useRef(prefix)
  const [controlMode, setControlMode] = useState(false), [controlColumn, setControlColumn] = useState<number | null>(null), [notice, setNotice] = useState('')
  useLayoutEffect(() => {
    if (previousOwner.current !== prefix) {
      if (gridRef.current?.contains(document.activeElement) || popupRef.current?.contains(document.activeElement)) gridRef.current?.focus()
      props.popup?.close()
      setControlMode(false)
      setNotice('')
      previousOwner.current = prefix
    }
    if (previousPopup.current && !props.popup) {
      const { popup, owner } = previousPopup.current, invoker = popup.invoker
      if (owner === prefix && invoker?.isConnected && gridRef.current?.contains(invoker)) invoker.focus()
      else gridRef.current?.focus()
    }
    previousPopup.current = props.popup ? { popup: props.popup, owner: previousPopup.current?.owner ?? prefix } : null
  }, [prefix, props.popup])
  const current = reconcileSelection(props.selection, projection, view), selection = current.selection
  const membership = new SourceSelection(selection.ranges), pane = visibleLogicalViewport(window, viewport)
  const rowWindow = projection.rows.visibleBetween(pane.y, pane.y + pane.height), colWindow = projection.cols.visibleBetween(pane.x, pane.x + pane.width)
  const limits = FRAME_LIMITS[props.profile ?? 'D']
  let coarse = props.renderMode === 'coarse' || rowWindow.count * colWindow.count >= limits.cells || rowWindow.count + colWindow.count >= limits.lines
  let slots: number[] = [], columns: number[] = []
  if (!coarse) {
    for (const run of rowWindow.runs) for (let slot = run.start; slot <= run.end; slot++) slots.push(slot)
    for (const run of colWindow.runs) for (let col = run.start; col <= run.end; col++) columns.push(col)
    coarse = slots.some(slot => projection.rows.size(slot) * window.zoom * window.deviceScale < 4)
      || columns.some(col => projection.cols.size(col) * window.zoom * window.deviceScale < 4)
    if (coarse) { slots = []; columns = [] }
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
      if (event.key === 'Escape' && controlMode && grid?.contains(event.target as Node)) {
        event.preventDefault(); setControlMode(false); grid.focus()
      }
      event.stopPropagation()
      return
    }
    if (document.activeElement !== grid) return
    if (event.key === 'Escape') {
      if (!selection.ranges.length && !selection.active && !selection.anchor) return
      event.preventDefault(); event.stopPropagation()
      publish(() => ({ selection: emptyGridSelection(), window, announcement: 'Selection cleared.' }))
    } else if (event.key === 'F2') {
      const buttons = [...grid.querySelectorAll<HTMLButtonElement>('[data-header-column]')]
      const button = buttons.find(button => Number(button.dataset.headerColumn) === active?.col) ?? buttons[0]
      if (button) { event.preventDefault(); event.stopPropagation(); button.focus() }
    } else if (event.altKey && event.key === 'ArrowDown' && active && props.onFilter) {
      event.preventDefault(); event.stopPropagation(); props.onFilter(active.col, grid)
    } else if (navigationKeys.has(event.key) && !event.altKey) {
      event.preventDefault(); event.stopPropagation()
      publish(() => gridNavigationTransition(props, event.key as NavigationKey, { ctrl: event.ctrlKey || event.metaKey, extend: event.shiftKey }))
    }
  }
  const activeBox = active && projection.sourceCellToLogical(active)
  const activeColumnInPane = activeBox && activeBox.x < pane.x + pane.width && activeBox.x + activeBox.width > pane.x
  const headerColumns = coarse && active && activeColumnInPane ? [active.col] : columns
  const headerProps = { projection, window, viewport, coarse, columns: headerColumns, selection, sort: view.sort, filter: view.filter, controlMode, controlColumn, onSortCycle: props.onSortCycle, onFilter: props.onFilter, onSortDialog: props.onSortDialog, onNotice: props.onNotice,
    onSelect: (kind: 'row' | 'column', index: number, options: { toggle: boolean; extend: boolean }, event: MouseEvent<HTMLButtonElement>) => {
      if (event.detail > 0 && !coarse && gridRef.current) {
        const bounds = gridRef.current.getBoundingClientRect()
        const point = { x: event.clientX - bounds.left - GRID_HEADER_SIZE, y: event.clientY - bounds.top - GRID_HEADER_SIZE }
        if (kind === 'row' ? point.y < 0 || point.y >= viewport.height : point.x < 0 || point.x >= viewport.width) return
        const target = hitTestViewport(window, projection, viewport, { x: kind === 'row' ? 0 : point.x, y: kind === 'row' ? point.y : 0 })
        let hit: number | undefined
        if (target) hit = kind === 'row' ? projection.order.sourceToDisplay(target.row) : target.col
        else {
          // A header can still select when its orthogonal data axis is all hidden.
          hit = (kind === 'row' ? slots : headerColumns).find(candidate => {
            const box = viewportHeaderBox(projection, window, viewport, kind, candidate)
            return kind === 'row' ? point.y >= box.y && point.y < box.y + box.height : point.x >= box.x && point.x < box.x + box.width
          })
        }
        if (hit === undefined) return
        index = kind === 'row' ? projection.order.displayToSource(hit) : hit
        const buttons = [...gridRef.current.querySelectorAll<HTMLButtonElement>(kind === 'row' ? '[data-header-row]' : '[data-header-column]')]
        buttons.find(button => Number(kind === 'row' ? button.dataset.headerRow : button.dataset.headerColumn) === hit)?.focus()
      }
      publish(() => {
        const selected = selectHeader(selection, projection, kind, index, options.toggle, options.extend)
        const result = selected.active ? navigateTo(window, { metrics: projection, viewport, view: { ...view, selection: selected } }, { addressSpace: 'source', at: selected.active }) : null
        return { selection: selected, window: result?.window ?? window, announcement: selected.active ? activeAnnouncement(projection, selected.active) : 'No visible records.' }
      })
    },
    onSelectAll: () => publish(() => ({ selection: { active, anchor: active, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: projection.rows.logicalCount - 1 }] }, cols: [{ c0: 0, c1: projection.cols.logicalCount - 1 }] }] }, window, announcement: 'Worksheet selected.' })),
  }
  const status = [props.announcement, current.announcement, notice, coarse ? 'Zoom in or inspect a cell for detail.' : ''].filter(Boolean).join(' ')
  return <div onKeyDown={keyDown} onBlur={event => { if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setControlMode(false) }} style={{ position: 'relative', width: viewport.width + GRID_HEADER_SIZE, height: viewport.height + GRID_HEADER_SIZE }}>
    <div ref={gridRef} role="grid" aria-label={props.label ?? 'Worksheet'} aria-rowcount={MAX_ROWS + 1} aria-colcount={MAX_COLS + 1}
      aria-multiselectable="true" aria-activedescendant={active ? `${prefix}-cell-${cellId(active.row, active.col)}` : undefined}
      aria-description={`Worksheet capacity ${MAX_ROWS} rows by ${MAX_COLS} columns.${coarse ? ' Coarse view. Zoom in or inspect a cell for detail.' : ''}`}
      tabIndex={controlMode ? -1 : 0} data-officeview-grid=""
      onFocusCapture={event => {
        setControlMode(event.target !== event.currentTarget)
        const target = event.target as HTMLElement, header = target.closest<HTMLElement>('[data-header-column], [data-action-column]')
        setControlColumn(header ? Number(header.dataset.headerColumn ?? header.dataset.actionColumn) : null)
      }}
      onPointerDown={event => { event.stopPropagation(); if (event.target === event.currentTarget) event.currentTarget.focus() }}
      onDoubleClick={event => event.stopPropagation()}
      onClick={event => {
        if (event.defaultPrevented || event.detail > 1 || (event.target as Element).closest('[data-officeview-ui]')) return
        const bounds = event.currentTarget.getBoundingClientRect()
        const target = hitTestViewport(window, projection, viewport, { x: event.clientX - bounds.left - GRID_HEADER_SIZE, y: event.clientY - bounds.top - GRID_HEADER_SIZE })
        if (!target) return
        event.currentTarget.focus()
        publish(() => ({ selection: selectCell(selection, projection, target, { extend: event.shiftKey, toggle: event.ctrlKey || event.metaKey }), window, announcement: activeAnnouncement(projection, target, props.cellPresentation?.(target).text) }))
      }} style={{ position: 'relative', width: '100%', height: '100%', overflow: 'hidden' }}>
      <div aria-hidden="true" style={{ position: 'absolute', left: GRID_HEADER_SIZE, top: GRID_HEADER_SIZE, width: viewport.width, height: viewport.height }}>{props.children}</div>
      <GridHeaders {...headerProps} />
      {slots.map(slot => {
        const row = projection.order.displayToSource(slot), extra = transient && slot === activeSlot
        const cells = extra ? [active!.col] : [...columns]
        if (!extra && slot === activeSlot && active && !cells.includes(active.col)) cells.push(active.col)
        const box = viewportHeaderBox(projection, window, viewport, 'row', slot)
        return <div key={slot} role="row" aria-rowindex={slot + 2} data-transient-active={extra ? 'true' : undefined}
          style={extra ? { ...hiddenText, top: 0, left: 0, pointerEvents: 'none' } : { position: 'absolute', top: GRID_HEADER_SIZE + box.y, left: 0, height: box.height, width: '100%', pointerEvents: 'none' }}>
          <GridRowHeader props={headerProps} slot={slot} />
          {cells.map(col => {
            const at = { row, col }, value = props.cellPresentation?.(at) ?? { text: '', state: 'loading' as const }
            const column = viewportHeaderBox(projection, window, viewport, 'column', col)
            const offscreen = extra || (slot === activeSlot && active?.col === col && !columns.includes(col))
            return <div key={col} role="gridcell" tabIndex={-1} id={`${prefix}-cell-${cellId(row, col)}`} aria-colindex={col + 2}
              aria-label={activeAnnouncement(projection, at, value.text)} aria-selected={membership.contains(at)} aria-description={value.state}
              style={offscreen ? { ...hiddenText, top: 0, left: 0 } : { position: 'absolute', top: 0, left: GRID_HEADER_SIZE + column.x, width: column.width, height: box.height }}><span style={hiddenText}>{boundedSelectionText(value.text)}</span></div>
          })}
        </div>
      })}
    </div>
    {props.popup && <div ref={popupRef} role="dialog" aria-label="Column filter" data-officeview-ui="grid-popup"
      onPointerDown={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()}
      style={{ position: 'absolute', zIndex: 3, bottom: 0, maxWidth: '100%', paddingBottom: 'env(safe-area-inset-bottom)' }}>{props.popup.content}</div>}
    <ViewerStatus {...props.status} records={props.status?.records ?? recordCounts(projection, props.body)} announcement={status} />
  </div>
}
