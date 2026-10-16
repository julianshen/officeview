/**
 * View-only row/column resizing in the interactive sheet viewer (spec §5.6
 * amendment): pointer handles on the header strips, the bounded auto-fit
 * reader and the tile-invalidation region of one resize. Sizes are session
 * state (xlsx/view/axisSizes); nothing here writes the file.
 *
 * Handles are presentational (aria-hidden): the accessible route is the
 * Alt+Shift+Arrow / Alt+Shift+Enter keyboard resize in GridSemantics, announced
 * through the status line. A drag shows a guide line and commits once on
 * release, so a drag never rebuilds geometry per pointer move.
 */
import { useEffect, useRef, useState, type CSSProperties, type PointerEvent as ReactPointerEvent, type ReactElement } from 'react'
import type { LogicalCellBox, ScrollWindow } from '../../xlsx/view/contracts'
import type { ProjectedSheetMetrics } from '../../xlsx/view/projection'
import type { SourceCell } from '../../xlsx/view/coordinates'
import { MAX_ROWS } from '../../xlsx/view/coordinates'
import type { ViewportSize } from '../../xlsx/view/scrollWindow'
import { TEXT_OVERFLOW_REACH_CSS, type SheetGeometry } from '../../xlsx/view/viewport'
import { AUTOFIT_MAX_CELLS, AUTOFIT_MAX_ROWS, autoFitRowCandidates, cellFont, clampSize, fitColumnWidth, fitRowHeight, type Measure } from '../../xlsx/view/axisSizes'
import type { XlsxCell, XlsxCellStyle, XlsxSheet } from '../../xlsx/types'
import { cellDisplayText } from '../../xlsx/render'
import { GRID_HEADER_SIZE, gridHeaderBox, viewportHeaderBox } from './GridHeaders'

/** A resize the user asked for: an absolute size, a keyboard step or an auto-fit.
 * `index` is the SOURCE column or SOURCE row. */
export interface ResizeRequest { kind: 'row' | 'column'; index: number; size?: number; delta?: number; fit?: boolean }

/** Fine-pointer hit zone either side of an edge (CSS px). */
const FINE_HALF = 4
/** Coarse-pointer handles: one 44px target each (spec §9.3). The active
 * column's straddles its right edge inside the select lane of the header
 * strip, never over the sort/filter buttons. The active row's sits in the
 * header corner's free second slot, so it never takes a tap from any row
 * header (row-header ownership, §9.3); it is omitted while that slot holds the
 * narrow-column actions lane, or when the strip has no second slot. */
const COARSE_HALF = GRID_HEADER_SIZE / 2
/** Movement below this many CSS px is a click (a double-click never resizes). */
const DRAG_SLOP = 2
const GUIDE = '#1a73e8'

export interface ResizeHandlesProps {
  projection: ProjectedSheetMetrics
  window: ScrollWindow
  viewport: ViewportSize
  headerHeight: number
  touch: boolean
  /** Visible detail columns (source) and row display slots of the frame. */
  columns: readonly number[]
  slots: readonly number[]
  active: SourceCell | null
  onResize: (request: ResizeRequest) => void
}
interface Drag { kind: 'row' | 'column'; index: number; slot: number; start: number; startSize: number; size: number; moved: boolean }

export function ResizeHandles(props: ResizeHandlesProps): ReactElement {
  const { projection, window: scroll, viewport, headerHeight, touch } = props
  const [drag, setDrag] = useState<Drag | null>(null)
  const dragRef = useRef<Drag | null>(null)
  const cleanup = useRef<(() => void) | null>(null)
  useEffect(() => () => cleanup.current?.(), [])
  const half = touch ? COARSE_HALF : FINE_HALF
  const activeSlot = props.active ? projection.order.sourceToDisplay(props.active.row) : null
  const columns = touch ? props.columns.filter(col => col === props.active?.col) : props.columns
  const laneShown = touch && props.active !== null && props.columns.includes(props.active.col)
    && viewportHeaderBox(projection, scroll, viewport, 'column', props.active.col).width < GRID_HEADER_SIZE
  const slots = touch ? props.slots.filter(slot => slot === activeSlot && !laneShown && headerHeight >= 2 * GRID_HEADER_SIZE) : props.slots

  const begin = (event: ReactPointerEvent<HTMLDivElement>, kind: 'row' | 'column', index: number, slot: number) => {
    if (event.button !== 0) return
    event.preventDefault()
    event.stopPropagation()
    try { event.currentTarget.setPointerCapture?.(event.pointerId) } catch { /* capture is an optimization */ }
    cleanup.current?.()
    const startSize = kind === 'row' ? projection.rows.size(slot) : projection.cols.size(index)
    const state: Drag = { kind, index, slot, start: kind === 'row' ? event.clientY : event.clientX, startSize, size: startSize, moved: false }
    dragRef.current = state
    setDrag(state)
    const host = event.currentTarget.ownerDocument.defaultView ?? globalThis
    const move = (e: PointerEvent) => {
      const current = dragRef.current
      if (!current) return
      const delta = ((current.kind === 'row' ? e.clientY : e.clientX) - current.start) / scroll.zoom
      const next = { ...current, moved: current.moved || Math.abs(delta) >= DRAG_SLOP, size: clampSize(current.kind === 'row' ? 'row' : 'col', current.startSize + delta) }
      dragRef.current = next
      setDrag(next)
    }
    const end = (e: PointerEvent) => {
      const current = dragRef.current
      stop()
      if (!current || e.type === 'pointercancel' || !current.moved || current.size === current.startSize) return
      props.onResize({ kind: current.kind, index: current.index, size: current.size })
    }
    const stop = () => {
      host.removeEventListener('pointermove', move)
      host.removeEventListener('pointerup', end)
      host.removeEventListener('pointercancel', end)
      dragRef.current = null
      cleanup.current = null
      setDrag(null)
    }
    host.addEventListener('pointermove', move)
    host.addEventListener('pointerup', end)
    host.addEventListener('pointercancel', end)
    cleanup.current = () => {
      host.removeEventListener('pointermove', move)
      host.removeEventListener('pointerup', end)
      host.removeEventListener('pointercancel', end)
      dragRef.current = null
    }
  }
  const handle = (kind: 'row' | 'column', index: number, slot: number, style: CSSProperties) =>
    <div key={`${kind}:${slot}`} aria-hidden="true" data-officeview-ui="resize-handle" {...(kind === 'row' ? { 'data-resize-row': slot } : { 'data-resize-column': index })}
      title={kind === 'row' ? 'Drag to resize the row; double-click to fit' : 'Drag to resize the column; double-click to fit'}
      onPointerDown={event => begin(event, kind, index, slot)}
      onClick={event => event.stopPropagation()}
      onDoubleClick={event => { event.preventDefault(); event.stopPropagation(); props.onResize({ kind, index, fit: true }) }}
      style={{ position: 'absolute', zIndex: 3, touchAction: 'none', cursor: kind === 'row' ? 'row-resize' : 'col-resize', display: 'flex', alignItems: 'center', justifyContent: 'center', ...style }}>
      {touch && <span style={{ fontSize: 14, lineHeight: 1, color: GUIDE, background: '#ffffffd0', borderRadius: 4, padding: '2px 3px', pointerEvents: 'none', whiteSpace: 'nowrap' }}>{kind === 'row' ? `↕${slot + 1}` : '↔'}</span>}
    </div>
  const handles: ReactElement[] = []
  for (const col of columns) {
    const box = gridHeaderBox(projection, scroll, 'column', col)
    const edge = box.x + box.width
    if (box.width <= 0 || edge < 0 || edge > viewport.width) continue
    handles.push(handle('column', col, col, { left: GRID_HEADER_SIZE + edge - half, top: 0, width: 2 * half, height: touch ? GRID_HEADER_SIZE : headerHeight }))
  }
  for (const slot of slots) {
    const box = gridHeaderBox(projection, scroll, 'row', slot)
    const edge = box.y + box.height
    if (box.height <= 0 || edge < 0 || edge > viewport.height) continue
    handles.push(handle('row', projection.order.displayToSource(slot), slot, touch
      ? { left: 0, top: GRID_HEADER_SIZE, width: GRID_HEADER_SIZE, height: GRID_HEADER_SIZE, background: '#f3f4f6' }
      : { left: 0, top: headerHeight + edge - half, width: GRID_HEADER_SIZE, height: 2 * half }))
  }
  if (drag?.moved) {
    const box = gridHeaderBox(projection, scroll, drag.kind, drag.kind === 'row' ? drag.slot : drag.index)
    const at = (drag.kind === 'row' ? box.y : box.x) + drag.size * scroll.zoom
    const label = `${drag.kind === 'row' ? 'Height' : 'Width'}: ${drag.size} px`
    handles.push(<div key="guide" aria-hidden="true" data-officeview-ui="resize-guide" style={drag.kind === 'row'
      ? { position: 'absolute', zIndex: 3, pointerEvents: 'none', left: 0, top: headerHeight + at - 1, width: GRID_HEADER_SIZE + viewport.width, height: 2, background: GUIDE }
      : { position: 'absolute', zIndex: 3, pointerEvents: 'none', left: GRID_HEADER_SIZE + at - 1, top: 0, width: 2, height: headerHeight + viewport.height, background: GUIDE }}>
      <span style={{ position: 'absolute', left: 4, top: 4, whiteSpace: 'nowrap', font: '12px sans-serif', color: '#ffffff', background: GUIDE, borderRadius: 3, padding: '1px 4px' }}>{label}</span>
    </div>)
  }
  return <>{handles}</>
}

// ---- auto-fit ----------------------------------------------------------------

type FitCell = Pick<XlsxCell, 'col' | 'value'> & { style?: XlsxCellStyle; styleIndex?: number }
export interface AutoFitInput {
  geometry: SheetGeometry
  sheet: XlsxSheet
  /** The pane in logical CSS px (visible rows/columns come first). */
  pane: LogicalCellBox
  /** Rows the source already holds on the main thread, never a read. */
  knownRows: (limit: number) => Iterable<number>
  styleOf?: (styleIndex: number) => XlsxCellStyle | undefined
  measure: Measure
}
function cellIn(cells: readonly FitCell[] | undefined, col: number): FitCell | undefined {
  if (!cells) return undefined
  let low = 0, high = cells.length
  while (low < high) {
    const mid = (low + high) >>> 1
    if (cells[mid].col < col) low = mid + 1
    else high = mid
  }
  return cells[low]?.col === col ? cells[low] : cells.find(cell => cell.col === col)
}
/** The bounded auto-fit size (axisSizes AUTOFIT_MAX_ROWS / AUTOFIT_MAX_CELLS):
 * a column fits its widest formatted value among the visible and known rows; a
 * row fits its tallest populated cell. Merged cells spanning the resized
 * direction are skipped, as Excel does. */
export function autoFitSize(kind: 'row' | 'column', index: number, input: AutoFitInput): number {
  const { geometry, sheet, pane, measure } = input
  const metrics = geometry.metrics
  const styleOf = (cell: FitCell) => cell.style ?? (cell.styleIndex !== undefined ? input.styleOf?.(cell.styleIndex) : undefined)
  const text = (cell: FitCell, style: XlsxCellStyle | undefined) => cellDisplayText(sheet, { value: cell.value, style })
  if (kind === 'column') {
    const visible = function* () {
      for (const run of metrics.rows.visibleBetween(pane.y, pane.y + pane.height).runs) for (let slot = run.start; slot <= run.end; slot++) yield metrics.order.displayToSource(slot)
    }
    const rows = autoFitRowCandidates(visible(), input.knownRows(4 * AUTOFIT_MAX_ROWS), row => row >= 0 && row < MAX_ROWS && metrics.sourceToVisibleRank(row) !== null)
    const cells = function* () {
      for (const row of rows) {
        const cell = cellIn(geometry.cellsOf(row) as readonly FitCell[] | undefined, index)
        if (!cell || geometry.merges.firstWhere({ r0: row, r1: row, c0: index, c1: index }, range => range.minCol !== range.maxCol)) continue
        const style = styleOf(cell)
        yield { text: text(cell, style), font: cellFont(style).font }
      }
    }
    return clampSize('col', fitColumnWidth(cells(), measure))
  }
  const row = index
  const populated = (geometry.cellsOf(row) ?? []) as readonly FitCell[]
  const inPane = (cell: FitCell) => { const x = metrics.cols.start(cell.col); return x < pane.x + pane.width && x + metrics.cols.size(cell.col) > pane.x }
  const ordered = [...populated.filter(inPane), ...populated.filter(cell => !inPane(cell))].slice(0, AUTOFIT_MAX_CELLS)
  const cells = function* () {
    for (const cell of ordered) {
      const widthCss = metrics.cols.size(cell.col)
      if (widthCss <= 0 || geometry.merges.firstWhere({ r0: row, r1: row, c0: cell.col, c1: cell.col }, range => range.minRow !== range.maxRow)) continue
      const style = styleOf(cell), { font, sizePx } = cellFont(style)
      yield { text: text(cell, style), font, sizePx, wrap: style?.wrapText === true, widthCss }
    }
  }
  return clampSize('row', fitRowHeight(cells(), measure))
}

/** Measurement on a scratch 2D context; a character-width estimate when the
 * host has no canvas (the result is still clamped). */
export function scratchMeasure(): Measure {
  let ctx: CanvasRenderingContext2D | null = null
  try { ctx = globalThis.document?.createElement('canvas').getContext('2d') ?? null } catch { ctx = null }
  const cache = new Map<string, number>()
  return (font, text) => {
    const key = `${font}\n${text}`
    let width = cache.get(key)
    if (width === undefined) {
      if (ctx) { ctx.font = font; width = ctx.measureText(text).width }
      else width = text.length * (Number(/([\d.]+)pt/.exec(font)?.[1] ?? 10) * 96 / 72) * 0.55
      cache.set(key, width)
    }
    return width
  }
}

// ---- minimal tile invalidation -------------------------------------------------

/** The logical region whose paint a resize of one line can change, in the OLD
 * geometry: everything right of (below) the line's start, widened left (up) to
 * the first line of any merge crossing it and, for columns, by the text
 * overflow reach. A sheet with drawings repaints entirely (a stretched picture
 * changes on both sides of the line). Null means the whole sheet. */
export function resizeRegion(geometry: SheetGeometry, kind: 'row' | 'column', index: number): { axis: 'x' | 'y'; from: number } | null {
  if (geometry.drawingsIn(-Infinity, -Infinity, Infinity, Infinity, Number.MAX_SAFE_INTEGER)?.length !== 0) return null
  const metrics = geometry.metrics
  if (kind === 'column') {
    let first = index
    for (const range of geometry.merges.collect({ r0: 0, r1: MAX_ROWS - 1, c0: index, c1: index })) first = Math.min(first, range.minCol)
    return { axis: 'x', from: metrics.cols.start(first) - TEXT_OVERFLOW_REACH_CSS }
  }
  let from = metrics.rows.start(metrics.order.sourceToDisplay(index))
  // A sorted view never permutes rows across a merge (spec §6.3), so a merge's
  // first source row is also its first display line.
  for (const range of geometry.merges.collect({ r0: index, r1: index, c0: 0, c1: metrics.cols.logicalCount - 1 })) {
    from = Math.min(from, metrics.rows.start(metrics.order.sourceToDisplay(range.minRow)))
  }
  return { axis: 'y', from }
}
