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
import { frozenViewportSize } from '../../xlsx/view/frozenPanes'
import type { SourceCell } from '../../xlsx/view/coordinates'
import { MAX_ROWS } from '../../xlsx/view/coordinates'
import type { ViewportSize } from '../../xlsx/view/scrollWindow'
import { TEXT_OVERFLOW_REACH_CSS, type SheetGeometry } from '../../xlsx/view/viewport'
import { AUTOFIT_MAX_CELLS, AUTOFIT_MAX_ROWS, RESIZE_LIMITS, autoFitRowCandidates, cellFont, clampSize, fitColumnWidth, fitRowHeight, type Measure } from '../../xlsx/view/axisSizes'
import { createCellLayoutHost } from '../../xlsx/view/renderViewport'
import { createAdmittedLayoutContext } from './measureContext'
import { TILE_CANVAS_BYTES, shapeCacheBytes } from '../../xlsx/view/mainLedger'
import type { XlsxCell, XlsxCellStyle, XlsxSheet } from '../../xlsx/types'
import { cellDisplayText } from '../../xlsx/render'
import { MoveVertical, MoveHorizontal } from 'lucide-react'
import { Icon } from '../Icon'
import type { FormatAllocation } from '../../xlsx/formula/format'
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
  /** Row-header gutter width (GridSemantics' header geometry). */
  gutter?: number
  touch: boolean
  /** Visible detail columns (source) and row display slots of the frame. */
  columns: readonly number[]
  slots: readonly number[]
  active: SourceCell | null
  onResize: (request: ResizeRequest) => void
}
interface Drag { kind: 'row' | 'column'; index: number; slot: number; pointerId: number; start: number; startSize: number; size: number; moved: boolean }

export interface HandleLayout {
  viewport: ViewportSize; headerHeight: number; touch: boolean
  /** Row-header gutter width (the header strip's own geometry); default GRID_HEADER_SIZE. */
  gutter?: number
}
export interface HandleBox { x: number; y: number; width: number; height: number }
/** Grid-local box of one handle for an edge at viewport offset `edge` (the
 * single place handle geometry is decided; a pane-aware caller, e.g. frozen
 * panes, maps `edge` and clips here). Every target keeps its full size inside
 * the visible pane and off the worksheet corner (x < 44, y < header height):
 * a target near a boundary slides inward while still covering its edge. The
 * coarse row target sits in the header corner's second slot (§9.3 ownership). */
export function handlePlacement(kind: 'row' | 'column', edge: number, layout: HandleLayout): HandleBox {
  const { viewport, headerHeight, touch } = layout
  const gutter = layout.gutter ?? GRID_HEADER_SIZE
  const half = touch ? COARSE_HALF : FINE_HALF
  const clamp = (value: number, min: number, max: number) => Math.max(min, Math.min(max, value))
  if (kind === 'column') {
    return { x: clamp(gutter + edge - half, gutter, gutter + viewport.width - 2 * half), y: 0, width: 2 * half, height: touch ? GRID_HEADER_SIZE : headerHeight }
  }
  if (touch) return { x: 0, y: GRID_HEADER_SIZE, width: gutter, height: GRID_HEADER_SIZE }
  return { x: 0, y: clamp(headerHeight + edge - half, headerHeight, headerHeight + viewport.height - 2 * half), width: gutter, height: 2 * half }
}
/** Grid-local box of the drag guide line at viewport offset `at`. */
export function guidePlacement(kind: 'row' | 'column', at: number, layout: HandleLayout): HandleBox {
  const gutter = layout.gutter ?? GRID_HEADER_SIZE
  return kind === 'row'
    ? { x: 0, y: layout.headerHeight + at - 1, width: gutter + layout.viewport.width, height: 2 }
    : { x: gutter + at - 1, y: 0, width: 2, height: layout.headerHeight + layout.viewport.height }
}
// Scroll-driven positions use transforms (not a layout shift, as for the header strips).
const placed = (box: HandleBox): CSSProperties => ({ left: 0, top: 0, transform: `translate(${box.x}px, ${box.y}px)`, width: box.width, height: box.height })

/** Mounted per sheet owner (GridSemantics keys it): an owner change unmounts
 * it, which cancels a drag in progress (never resizing another sheet). */
export function ResizeHandles(props: ResizeHandlesProps): ReactElement {
  const { projection, window: scroll, viewport, headerHeight, touch } = props
  const [drag, setDrag] = useState<Drag | null>(null)
  const dragRef = useRef<Drag | null>(null)
  const cleanup = useRef<(() => void) | null>(null)
  useEffect(() => () => cleanup.current?.(), [])
  const layout: HandleLayout = { viewport, headerHeight, touch, ...(props.gutter !== undefined ? { gutter: props.gutter } : {}) }
  const pinnedCss = frozenViewportSize(scroll.frozen, scroll.zoom)
  const pinnedColumn = (col: number) => Boolean(scroll.frozen && projection.cols.start(col) < scroll.frozen.left)
  const pinnedRow = (slot: number) => Boolean(scroll.frozen && projection.rows.start(slot) < scroll.frozen.top)
  const activeSlot = props.active ? projection.order.sourceToDisplay(props.active.row) : null
  const columns = touch ? props.columns.filter(col => col === props.active?.col) : props.columns
  const laneShown = touch && props.active !== null && props.columns.includes(props.active.col)
    && viewportHeaderBox(projection, scroll, viewport, 'column', props.active.col).width < GRID_HEADER_SIZE
  const slots = touch ? props.slots.filter(slot => slot === activeSlot && !laneShown && headerHeight >= 2 * GRID_HEADER_SIZE) : props.slots

  const begin = (event: ReactPointerEvent<HTMLDivElement>, kind: 'row' | 'column', index: number, slot: number) => {
    if (event.button !== 0 || dragRef.current) return
    event.preventDefault()
    event.stopPropagation()
    try { event.currentTarget.setPointerCapture?.(event.pointerId) } catch { /* capture is an optimization */ }
    cleanup.current?.()
    const startSize = kind === 'row' ? projection.rows.size(slot) : projection.cols.size(index)
    const state: Drag = { kind, index, slot, pointerId: event.pointerId, start: kind === 'row' ? event.clientY : event.clientX, startSize, size: startSize, moved: false }
    dragRef.current = state
    setDrag(state)
    const host = event.currentTarget.ownerDocument.defaultView ?? globalThis
    const zoom = scroll.zoom
    const move = (e: PointerEvent) => {
      const current = dragRef.current
      if (!current || e.pointerId !== current.pointerId) return
      // Click slop is compared in CSS (client) pixels; the logical size divides by zoom.
      const client = (current.kind === 'row' ? e.clientY : e.clientX) - current.start
      const next = { ...current, moved: current.moved || Math.abs(client) >= DRAG_SLOP, size: clampSize(current.kind === 'row' ? 'row' : 'col', current.startSize + client / zoom) }
      dragRef.current = next
      setDrag(next)
    }
    const end = (e: PointerEvent) => {
      const current = dragRef.current
      if (!current || e.pointerId !== current.pointerId) return
      stop()
      setDrag(null)
      if (e.type === 'pointercancel' || !current.moved || current.size === current.startSize) return
      props.onResize({ kind: current.kind, index: current.index, size: current.size })
    }
    const stop = () => {
      host.removeEventListener('pointermove', move)
      host.removeEventListener('pointerup', end)
      host.removeEventListener('pointercancel', end)
      dragRef.current = null
      cleanup.current = null
    }
    host.addEventListener('pointermove', move)
    host.addEventListener('pointerup', end)
    host.addEventListener('pointercancel', end)
    cleanup.current = stop
  }
  const handle = (kind: 'row' | 'column', index: number, slot: number, box: HandleBox, extra: CSSProperties = {}) => {
    const label = kind === 'row' ? `Resize row ${slot + 1}` : `Resize column ${index + 1}`
    return (
      <div key={`${kind}:${slot}`} aria-hidden="true" data-officeview-ui="resize-handle" {...(kind === 'row' ? { 'data-resize-row': slot } : { 'data-resize-column': index })}
        aria-label={label}
        title={kind === 'row' ? 'Drag to resize the row; double-click to fit' : 'Drag to resize the column; double-click to fit'}
        onPointerDown={event => begin(event, kind, index, slot)}
        onClick={event => event.stopPropagation()}
        onDoubleClick={event => { event.preventDefault(); event.stopPropagation(); props.onResize({ kind, index, fit: true }) }}
        style={{ position: 'absolute', zIndex: 3, touchAction: 'none', cursor: kind === 'row' ? 'row-resize' : 'col-resize', display: 'flex', alignItems: 'center', justifyContent: 'center', ...placed(box), ...extra }}>
        {touch && (
          <span
            aria-label={label}
            className="ov:flex ov:items-center ov:gap-0.5 ov:text-xs ov:font-semibold ov:text-muted ov:bg-surface/90 ov:rounded-control ov:px-1 ov:py-0.5 ov:shadow-rest ov:pointer-events-none ov:whitespace-nowrap"
            style={{ fontSize: 14, lineHeight: 1 }}>
            {kind === 'row' ? (
              <>
                <Icon icon={MoveVertical} name="move-vertical" size={14} />
                {slot + 1}
              </>
            ) : (
              <Icon icon={MoveHorizontal} name="move-horizontal" size={14} />
            )}
          </span>
        )}
      </div>
    )
  }
  const handles: ReactElement[] = []
  const fits = (length: number) => length >= (touch ? 2 * COARSE_HALF : 2 * FINE_HALF)
  for (const col of columns) {
    const box = gridHeaderBox(projection, scroll, 'column', col)
    const edge = box.x + box.width
    const start = pinnedColumn(col) ? 0 : pinnedCss.x
    const end = pinnedColumn(col) ? pinnedCss.x : viewport.width
    // Keep the target on its own side of the freeze divider. If that region
    // cannot fit the full target, the keyboard resize route remains available.
    if (box.width <= 0 || edge < start || edge > end || !fits(end - start)) continue
    const regionLayout = { ...layout, gutter: (layout.gutter ?? GRID_HEADER_SIZE) + start, viewport: { ...viewport, width: end - start } }
    handles.push(handle('column', col, col, handlePlacement('column', edge - start, regionLayout)))
  }
  for (const slot of slots) {
    const box = gridHeaderBox(projection, scroll, 'row', slot)
    const edge = box.y + box.height
    const start = pinnedRow(slot) ? 0 : pinnedCss.y
    const end = pinnedRow(slot) ? pinnedCss.y : viewport.height
    if (box.height <= 0 || edge < start || edge > end || (!touch && !fits(end - start))) continue
    const regionLayout = touch ? layout : { ...layout, headerHeight: headerHeight + start, viewport: { ...viewport, height: end - start } }
    handles.push(handle('row', projection.order.displayToSource(slot), slot, handlePlacement('row', touch ? edge : edge - start, regionLayout), touch ? { background: 'var(--ov-color-header, #f3f4f6)' } : {}))
  }
  if (drag?.moved) {
    const box = gridHeaderBox(projection, scroll, drag.kind, drag.kind === 'row' ? drag.slot : drag.index)
    const at = (drag.kind === 'row' ? box.y : box.x) + drag.size * scroll.zoom
    const label = `${drag.kind === 'row' ? 'Height' : 'Width'}: ${drag.size} px`
    const fixed = drag.kind === 'row' ? pinnedRow(drag.slot) : pinnedColumn(drag.index)
    const min = fixed ? 0 : drag.kind === 'row' ? pinnedCss.y : pinnedCss.x
    const max = drag.kind === 'row' ? viewport.height : viewport.width
    // A body edge hidden under the pinned strip has no visible guide. Pinned
    // edges keep their fixed origin even when growing past the old divider.
    if (at >= min && at <= max) handles.push(<div key="guide" aria-hidden="true" data-officeview-ui="resize-guide" style={{ position: 'absolute', zIndex: 3, pointerEvents: 'none', background: GUIDE, ...placed(guidePlacement(drag.kind, at, layout)) }}>
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
  /** Disjoint visible regions in priority order (pinned strips before body). */
  regions?: readonly LogicalCellBox[]
  /** Rows the source already holds on the main thread, never a read. */
  knownRows: (limit: number) => Iterable<number>
  styleOf?: (styleIndex: number) => XlsxCellStyle | undefined
  measure: Measure
  /** Formatting allocation observer (FitMeasure.allocation): every builder admitted first. */
  allocation?: FormatAllocation
  /** Canonical painter height of a wrapped or rotated body (FitMeasure.layoutHeight). */
  layoutHeight?: (cell: XlsxCell, widthCss: number) => number | undefined
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
  const regions = input.regions ?? [pane]
  const metrics = geometry.metrics
  const styleOf = (cell: FitCell) => cell.style ?? (cell.styleIndex !== undefined ? input.styleOf?.(cell.styleIndex) : undefined)
  const text = (cell: FitCell, style: XlsxCellStyle | undefined) => cellDisplayText(sheet, { value: cell.value, style }, input.allocation)
  if (kind === 'column') {
    const visible = function* () {
      for (const region of regions) for (const run of metrics.rows.visibleBetween(region.y, region.y + region.height).runs) for (let slot = run.start; slot <= run.end; slot++) yield metrics.order.displayToSource(slot)
    }
    const rows = autoFitRowCandidates(visible(), input.knownRows(4 * AUTOFIT_MAX_ROWS), row => row >= 0 && row < MAX_ROWS && metrics.sourceToVisibleRank(row) !== null)
    const cells = function* () {
      for (const row of rows) {
        const cell = cellIn(geometry.cellsOf(row) as readonly FitCell[] | undefined, index)
        if (!cell || geometry.merges.firstWhere({ r0: row, r1: row, c0: index, c1: index }, range => range.minCol !== range.maxCol)) continue
        const style = styleOf(cell), shown = text(cell, style), font = cellFont(style).font
        // Excel parity: a wrapping cell never asks for its unwrapped width; it
        // contributes its longest unbreakable word (its own line breaks stay).
        if (style?.wrapText === true) { for (const word of shown.split(/\s+/)) if (word) yield { text: word, font } }
        else yield { text: shown, font }
      }
    }
    return clampSize('col', fitColumnWidth(cells(), measure))
  }
  const row = index
  const populated = (geometry.cellsOf(row) ?? []) as readonly FitCell[]
  const inPane = (cell: FitCell) => { const x = metrics.cols.start(cell.col); return regions.some(region => x < region.x + region.width && x + metrics.cols.size(cell.col) > region.x) }
  const ordered = [...populated.filter(inPane), ...populated.filter(cell => !inPane(cell))].slice(0, AUTOFIT_MAX_CELLS)
  const cells = function* () {
    for (const cell of ordered) {
      const widthCss = metrics.cols.size(cell.col)
      if (widthCss <= 0 || geometry.merges.firstWhere({ r0: row, r1: row, c0: cell.col, c1: cell.col }, range => range.minRow !== range.maxRow)) continue
      const style = styleOf(cell), { font, sizePx } = cellFont(style), wrap = style?.wrapText === true
      const heightPx = wrap || style?.textRotation ? input.layoutHeight?.({ ref: '', row, col: cell.col, value: cell.value, styleIndex: cell.styleIndex ?? 0, ...(style ? { style } : {}) }, widthCss) : undefined
      yield { text: text(cell, style), font, sizePx, wrap, widthCss, ...(heightPx !== undefined ? { heightPx } : {}) }
    }
  }
  return clampSize('row', fitRowHeight(cells(), measure))
}

/** An auto-fit was refused because its measurement scratch would exceed its
 * admitted memory (the caller announces it; nothing is resized). */
export class AutoFitRefusedError extends Error {
  constructor() { super('Auto-fit needs more memory than is available; resize by dragging or with the keyboard instead.'); this.name = 'AutoFitRefusedError' }
}
/** Retained JS bytes per measured (font, text) cache entry beyond its UTF-16 payload. */
const MEASURE_ENTRY_BYTES = 96
/** A canonical body layout's transient JS bytes beyond its text (§11.2 records). */
const LAYOUT_HEAP_BYTES = 512
export interface FitCharge { heap: number; native: number }
export interface FitMeasure {
  measure: Measure
  /** Height of a wrapped or rotated cell body from the canonical painter
   * layout, as its transformed (rotated) extent. */
  layoutHeight?: (cell: XlsxCell, widthCss: number) => number | undefined
  /** Formatting observer: each builder is admitted (heap) before it is made. */
  allocation: FormatAllocation
  bytes(): number
}
/** One auto-fit's measurement owner: a 1×1 scratch canvas whose 2D context and
 * caches live only for that fit (the caller drops them after it). Costs are
 * reserved BEFORE they exist, with the calibrated §12.4 model: the context
 * (TILE_CANVAS_BYTES) before it is created, the browser shaping cache
 * (shapeCacheBytes per font) before each new text is measured, and the JS
 * cache entries. `admit({ heap, native })` receives the running totals and
 * charges them; a refusal throws AutoFitRefusedError and nothing more is
 * created. With an injected `measureText` no context is created or charged. */
export function createFitMeasure(options: { admit: (charge: FitCharge) => boolean; measureText?: Measure; sheet?: XlsxSheet }): FitMeasure {
  let ctx: CanvasRenderingContext2D | null | undefined = options.measureText ? null : undefined
  let host: ReturnType<typeof createCellLayoutHost> | null = null
  const fonts = new Map<string, { entries: number; units: number }>()
  const cache = new Map<string, number>()
  let heap = 0, native = 0
  const reserve = (moreHeap: number, moreNative: number) => {
    if (!options.admit({ heap: heap + moreHeap, native: native + moreNative })) throw new AutoFitRefusedError()
    heap += moreHeap; native += moreNative
  }
  /** Reserves one new shaped text in `font`'s cache before the browser shapes it. */
  const shaped = new Map<string, Set<string>>()
  /** `cache` names the context whose shaping cache grows (ours, or the
   * canonical layout's local measurement context). */
  const reserveShape = (font: string, text: string, cacheName = '') => {
    font = `${cacheName}\u0000${font}`
    const texts = shaped.get(font)
    if (texts?.has(text)) return
    const current = fonts.get(font) ?? { entries: 0, units: 0 }
    const next = { entries: current.entries + 1, units: current.units + text.length }
    reserve(font.length * 2 + text.length * 2 + MEASURE_ENTRY_BYTES, shapeCacheBytes(next.entries, next.units) - shapeCacheBytes(current.entries, current.units))
    fonts.set(font, next)
    if (texts) texts.add(text)
    else shaped.set(font, new Set([text]))
  }
  const context = (): CanvasRenderingContext2D | null => {
    if (ctx !== undefined) return ctx
    // Both admitted contexts (the base and the canonical layout's local
    // identity context) are reserved before they exist; each measures only
    // through its own admitting wrapper (no global state is touched).
    reserve(0, 2 * TILE_CANVAS_BYTES)
    ctx = createAdmittedLayoutContext((cacheName, font, text) => reserveShape(font, text, cacheName))
    return ctx
  }
  // Formatting builders are admitted (heap) before they are made; the fit's
  // owner releases everything when the fit ends.
  const allocation: FormatAllocation = { allocate: bytes => reserve(bytes, 0), release: () => {} }
  const measure: Measure = (font, text) => {
    const key = `${font}\n${text}`
    let width = cache.get(key)
    if (width !== undefined) return width
    const c = context()
    reserve(key.length * 2 + MEASURE_ENTRY_BYTES, 0)
    if (c) {
      c.font = font
      width = c.measureText(text).width
    } else width = options.measureText ? options.measureText(font, text) : text.length * (Number(/([\d.]+)pt/.exec(font)?.[1] ?? 10) * 96 / 72) * 0.55
    cache.set(key, width)
    return width
  }
  const layoutHeight = options.sheet && !options.measureText ? (cell: XlsxCell, widthCss: number): number | undefined => {
    const c = context()
    if (!c) return undefined
    // As paint does: text that cannot change with the width feeds the layout,
    // so only a General number is formatted a second time.
    let widthDependent = false
    const text = cellDisplayText(options.sheet!, cell, allocation, () => { widthDependent = true })
    reserve(text.length * 4 + LAYOUT_HEAP_BYTES, 0)
    host ??= createCellLayoutHost(c, options.sheet!)
    const body = host.layoutBody(cell, widthCss, RESIZE_LIMITS.row.max, undefined, widthDependent ? undefined : text, allocation)
    if (body.text?.kind !== 'layout') return undefined
    return body.text.rotation ? rotatedExtent(body.text.layout, c) : body.text.layout.height
  } : undefined
  return { measure, ...(layoutHeight ? { layoutHeight } : {}), allocation, bytes: () => heap + native }
}
/** Vertical extent of a laid-out body whose segments carry rotated baseline
 * transforms (render.ts rotatedLayout), from the canonical glyph band of each
 * segment's font: measured 'Mg' ascent/descent, at least .85em/.25em
 * (render.ts cellTextMeasurer / textLayoutBounds), for any angle. */
function rotatedExtent(layout: { lines: ReadonlyArray<{ baseline: number; segments: ReadonlyArray<{ text: string; x: number; width: number; style: { fontSizePt?: number; bold?: boolean; italic?: boolean }; transform?: { a: number; b: number; c: number; d: number; e: number; f: number } }> }> }, ctx: CanvasRenderingContext2D): number {
  const bands = new Map<string, { ascent: number; descent: number }>()
  const band = (style: { fontSizePt?: number; bold?: boolean; italic?: boolean }) => {
    const { font, sizePx } = cellFont(style)
    let value = bands.get(font)
    if (!value) {
      ctx.font = font
      const m = ctx.measureText('Mg')
      bands.set(font, value = { ascent: Math.max(m.actualBoundingBoxAscent ?? 0, sizePx * .85), descent: Math.max(m.actualBoundingBoxDescent ?? 0, sizePx * .25) })
    }
    return value
  }
  let top = Infinity, bottom = -Infinity
  for (const line of layout.lines) for (const segment of line.segments) {
    if (segment.text === '\n') continue
    const t = segment.transform ?? { a: 1, b: 0, c: 0, d: 1, e: segment.x, f: line.baseline }
    const { ascent, descent } = band(segment.style)
    for (const [x, y] of [[0, -ascent], [segment.width, -ascent], [segment.width, descent], [0, descent]]) {
      const py = t.b * x + t.d * y + t.f
      top = Math.min(top, py); bottom = Math.max(bottom, py)
    }
  }
  return bottom > top ? bottom - top : 0
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
