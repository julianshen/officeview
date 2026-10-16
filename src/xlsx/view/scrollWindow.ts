import { clampZoom } from '../../core/zoom'
import { MAX_PHYSICAL_SCROLL_CSS_PX, SHEET_MIN_ZOOM, SHEET_MAX_ZOOM, type AxisMetrics, type LogicalCellBox, type ScrollWindow, type SheetMetrics, type SheetView } from './contracts'
import { isValidCoordinate, parseA1, validateCoordinate, type NavigatorRequest, type SourceCell, type DisplaySlot } from './coordinates'
import { normalizeFilter } from './visibility'
import { bodyOriginFor, ensureAxisVisible, frozenExtent, frozenViewportSize, originAtLogical, originAtViewport } from './frozenPanes'

export interface ViewportSize { width: number; height: number }
export interface ScrollPoint { x: number; y: number }
export interface ScrollOptions { zoom?: number; deviceScale?: number; offset?: ScrollPoint; frozen?: { rows: number; cols: number } }
const CAP = MAX_PHYSICAL_SCROLL_CSS_PX
const GUARD_FRACTION = 0.2
function finite(value: number): void { if (!Number.isFinite(value)) throw new RangeError('Scroll coordinates must be finite') }
function point(value: ScrollPoint): void { finite(value.x); finite(value.y) }
function viewportSize(viewport: ViewportSize): void {
  finite(viewport.width); finite(viewport.height)
  if (viewport.width < 0 || viewport.height < 0 || viewport.width > CAP || viewport.height > CAP) throw new RangeError('Viewport exceeds the bounded scroll segment')
}
function state(window: ScrollWindow): void {
  point(window.logicalOrigin); point(window.physicalPosition)
  finite(window.physicalExtent.width); finite(window.physicalExtent.height)
  if (!Number.isFinite(window.zoom) || window.zoom < SHEET_MIN_ZOOM || window.zoom > SHEET_MAX_ZOOM
    || !Number.isFinite(window.deviceScale) || window.deviceScale <= 0
    || window.logicalOrigin.x < 0 || window.logicalOrigin.y < 0
    || window.physicalExtent.width < 0 || window.physicalExtent.height < 0
    || window.physicalExtent.width > CAP || window.physicalExtent.height > CAP
    || window.physicalPosition.x < 0 || window.physicalPosition.y < 0
    || window.physicalPosition.x > window.physicalExtent.width || window.physicalPosition.y > window.physicalExtent.height) throw new RangeError('Invalid sheet scroll window')
}
function clamp(value: number, min: number, max: number): number { return Math.max(min, Math.min(max, value)) }
/** Zoom arguments are document state: non-finite values are rejected instead of
 * silently becoming a bound. Gesture deltas/distances are event inputs and keep
 * their ignore-and-keep-state rule (see wheelSheetZoom/pinchSheetZoom). */
function finiteZoom(zoom: number): number {
  if (!Number.isFinite(zoom)) throw new RangeError('Sheet zoom must be finite')
  return zoom
}
function axisBounds(axis: AxisMetrics, viewportCss: number, zoom: number) {
  const total = axis.totalCss()
  finite(total)
  if (total < 0) throw new RangeError('Negative logical axis extent')
  const extent = Math.max(viewportCss, total >= CAP / zoom ? CAP : total * zoom)
  return { extent, range: Math.max(0, extent - viewportCss), maxOffset: Math.max(0, total - viewportCss / zoom) }
}
function placeAxis(axis: AxisMetrics, viewportCss: number, zoom: number, requested: number) {
  finite(requested)
  const bounds = axisBounds(axis, viewportCss, zoom)
  const offset = clamp(requested, 0, bounds.maxOffset)
  const lower = Math.max(0, bounds.range - Math.min(bounds.range, Math.max(0, bounds.maxOffset - offset) * zoom))
  const upper = Math.min(bounds.range, offset * zoom)
  return { offset, position: clamp(bounds.range / 2, lower, upper), extent: bounds.extent }
}
function makeWindow(metrics: SheetMetrics, viewport: ViewportSize, offset: ScrollPoint, zoom: number, deviceScale: number, freeze?: { rows: number; cols: number }): ScrollWindow {
  viewportSize(viewport); point(offset)
  if (!Number.isFinite(deviceScale) || deviceScale <= 0) throw new RangeError('Device scale must be finite and positive')
  const x = placeAxis(metrics.cols, viewport.width, zoom, offset.x)
  const y = placeAxis(metrics.rows, viewport.height, zoom, offset.y)
  validatePixelPrecision(x.offset, viewport.width, zoom, deviceScale)
  validatePixelPrecision(y.offset, viewport.height, zoom, deviceScale)
  const frozen = frozenExtent(freeze, metrics, viewport, zoom)
  return Object.freeze({ logicalOrigin: Object.freeze({ x: x.offset, y: y.offset }), physicalPosition: Object.freeze({ x: x.position, y: y.position }), physicalExtent: Object.freeze({ width: x.extent, height: y.extent }), zoom, deviceScale, ...(frozen ? { frozen } : {}) })
}
/** logicalOrigin is the actual Float64 viewport offset. physicalPosition is the
 * native segment offset; its derived logical base is origin-position/zoom (§5.4).
 * Recentring changes the native offset, never the viewport's world origin. */
export function createScrollWindow(metrics: SheetMetrics, viewport: ViewportSize, options: ScrollOptions = {}): ScrollWindow {
  return makeWindow(metrics, viewport, options.offset ?? { x: 0, y: 0 }, clampZoom(finiteZoom(options.zoom ?? 1), SHEET_MIN_ZOOM, SHEET_MAX_ZOOM), options.deviceScale ?? 1, options.frozen)
}
export function scrollOffset(window: ScrollWindow): ScrollPoint { state(window); return { ...window.logicalOrigin } }
export function setScrollOffset(window: ScrollWindow, metrics: SheetMetrics, viewport: ViewportSize, offset: ScrollPoint): ScrollWindow {
  state(window)
  return makeWindow(metrics, viewport, offset, window.zoom, window.deviceScale, window.frozen)
}
function applyLocalMotion(window: ScrollWindow, metrics: SheetMetrics, viewport: ViewportSize, delta: ScrollPoint): ScrollWindow {
  const requested = { x: window.logicalOrigin.x + delta.x / window.zoom, y: window.logicalOrigin.y + delta.y / window.zoom }
  const placed = setScrollOffset(window, metrics, viewport, requested)
  const native = { x: window.physicalPosition.x + delta.x, y: window.physicalPosition.y + delta.y }
  const preserve = (position: number, extent: number, viewSize: number, requestedOffset: number, actualOffset: number, fallback: number) =>
    Number.isFinite(position) && position >= 0 && position <= Math.max(0, extent - viewSize) && requestedOffset === actualOffset ? position : fallback
  const candidate = Object.freeze({ ...placed, physicalPosition: Object.freeze({
    x: preserve(native.x, placed.physicalExtent.width, viewport.width, requested.x, placed.logicalOrigin.x, placed.physicalPosition.x),
    y: preserve(native.y, placed.physicalExtent.height, viewport.height, requested.y, placed.logicalOrigin.y, placed.physicalPosition.y),
  }) })
  return recenterScrollWindow(candidate, metrics, viewport)
}
/** Deltas are viewport CSS pixels, independent of DPR and whole-sheet size.
 * The event adapter normalizes wheel deltaMode before calling this helper. */
export function scrollBy(window: ScrollWindow, metrics: SheetMetrics, viewport: ViewportSize, delta: ScrollPoint): ScrollWindow {
  state(window); point(delta)
  return applyLocalMotion(window, metrics, viewport, delta)
}
export function panScrollWindow(window: ScrollWindow, metrics: SheetMetrics, viewport: ViewportSize, delta: ScrollPoint): ScrollWindow {
  point(delta)
  return scrollBy(window, metrics, viewport, { x: -delta.x, y: -delta.y })
}
export function scrollPhysicalTo(window: ScrollWindow, metrics: SheetMetrics, viewport: ViewportSize, position: ScrollPoint): ScrollWindow {
  state(window); viewportSize(viewport); point(position)
  const x = clamp(position.x, 0, Math.max(0, window.physicalExtent.width - viewport.width))
  const y = clamp(position.y, 0, Math.max(0, window.physicalExtent.height - viewport.height))
  return applyLocalMotion(window, metrics, viewport, { x: x - window.physicalPosition.x, y: y - window.physicalPosition.y })
}
export function recenterScrollWindow(window: ScrollWindow, metrics: SheetMetrics, viewport: ViewportSize): ScrollWindow {
  state(window); viewportSize(viewport)
  const x = axisBounds(metrics.cols, viewport.width, window.zoom), y = axisBounds(metrics.rows, viewport.height, window.zoom)
  const guarded = (position: number, range: number) => range > 0 && (position < range * GUARD_FRACTION || position > range * (1 - GUARD_FRACTION))
  const changeX = guarded(window.physicalPosition.x, x.range) || window.physicalExtent.width !== x.extent || window.physicalPosition.x > x.range || window.logicalOrigin.x > x.maxOffset
  const changeY = guarded(window.physicalPosition.y, y.range) || window.physicalExtent.height !== y.extent || window.physicalPosition.y > y.range || window.logicalOrigin.y > y.maxOffset
  if (!changeX && !changeY) return window
  const placed = setScrollOffset(window, metrics, viewport, window.logicalOrigin)
  return Object.freeze({ ...placed, physicalPosition: Object.freeze({ x: changeX ? placed.physicalPosition.x : window.physicalPosition.x, y: changeY ? placed.physicalPosition.y : window.physicalPosition.y }) })
}
/** Size, zoom/DPR or projection publication does not introduce a fit-width scale. */
export function refreshScrollWindow(window: ScrollWindow, metrics: SheetMetrics, viewport: ViewportSize): ScrollWindow {
  return setScrollOffset(window, metrics, viewport, window.logicalOrigin)
}
/** Keep the start anchor's world point under the next anchor (pinch translation).
 * Content boundaries may clamp an otherwise impossible anchor; blank overscroll
 * is never manufactured to hide that unavoidable edge constraint. */
export function setScrollZoom(window: ScrollWindow, metrics: SheetMetrics, viewport: ViewportSize, nextZoom: number, anchor: ScrollPoint, nextAnchor: ScrollPoint = anchor, deviceScale = window.deviceScale): ScrollWindow {
  const world = viewportToLogical(window, anchor)
  point(nextAnchor)
  const zoom = clampZoom(finiteZoom(nextZoom), SHEET_MIN_ZOOM, SHEET_MAX_ZOOM)
  // Reconcile against the strips pinned before AND after the zoom: an anchor
  // that stays inside a pinned strip has no scroll meaning (keep that axis's
  // origin); otherwise the body mapping keeps the anchor's world point,
  // including when the zoom pins or unpins the strip (60% threshold).
  const before = frozenViewportSize(window.frozen, window.zoom)
  const after = frozenViewportSize(frozenExtent(window.frozen, metrics, viewport, zoom), zoom)
  const keepX = anchor.x < before.x && nextAnchor.x < after.x, keepY = anchor.y < before.y && nextAnchor.y < after.y
  return makeWindow(metrics, viewport, { x: keepX ? window.logicalOrigin.x : world.x - nextAnchor.x / zoom, y: keepY ? window.logicalOrigin.y : world.y - nextAnchor.y / zoom }, zoom, deviceScale, window.frozen)
}
export function logicalToViewport(window: ScrollWindow, value: ScrollPoint): ScrollPoint {
  state(window); point(value)
  const origin = originAtLogical(window.frozen, window.logicalOrigin, value)
  const result = { x: (value.x - origin.x) * window.zoom, y: (value.y - origin.y) * window.zoom }
  point(result)
  return result
}
export function viewportToLogical(window: ScrollWindow, value: ScrollPoint): ScrollPoint {
  state(window); point(value)
  const origin = originAtViewport(window.frozen, window.logicalOrigin, window.zoom, value)
  const result = { x: origin.x + value.x / window.zoom, y: origin.y + value.y / window.zoom }
  point(result)
  return result
}
export function logicalToSegment(window: ScrollWindow, value: ScrollPoint): ScrollPoint {
  const viewport = logicalToViewport(window, value)
  return { x: (viewport.x + window.physicalPosition.x) / window.zoom, y: (viewport.y + window.physicalPosition.y) / window.zoom }
}
export function logicalToBitmap(window: ScrollWindow, value: ScrollPoint): ScrollPoint {
  const viewport = logicalToViewport(window, value)
  const result = { x: viewport.x * window.deviceScale, y: viewport.y * window.deviceScale }
  point(result)
  return result
}
export function bitmapToLogical(window: ScrollWindow, value: ScrollPoint): ScrollPoint {
  state(window); point(value)
  return viewportToLogical(window, { x: value.x / window.deviceScale, y: value.y / window.deviceScale })
}
export function sourceToViewport(window: ScrollWindow, metrics: SheetMetrics, at: SourceCell): LogicalCellBox | null {
  const box = metrics.sourceCellToLogical(at)
  if (!box) return null
  const result = { ...logicalToViewport(window, box), width: box.width * window.zoom, height: box.height * window.zoom }
  finite(result.width); finite(result.height)
  return result
}
function axisStart(axis: AxisMetrics, index: number): number { return index === axis.logicalCount ? axis.totalCss() : axis.start(index) }
/** Search canonical projected starts in viewport space. Inverting a fractional
 * scale alone can round an exact border just below its cell, so hit tests compare
 * the same projected boundary used for painting (§5.4). */
function hitAxis(axis: AxisMetrics, origin: number, zoom: number, css: number): number | null {
  if (!axis.logicalCount || !axis.totalCss()) return null
  let low = 0, high = axis.logicalCount
  while (low < high) {
    const mid = Math.floor((low + high) / 2)
    if ((axis.start(mid) - origin) * zoom <= css) low = mid + 1
    else high = mid
  }
  const index = low - 1
  if (index < 0 || axis.size(index) <= 0 || css >= (axisStart(axis, index + 1) - origin) * zoom) return null
  return index
}
export function hitTestViewport(window: ScrollWindow, metrics: SheetMetrics, viewport: ViewportSize, value: ScrollPoint): SourceCell | null {
  state(window); viewportSize(viewport); point(value)
  if (value.x < 0 || value.y < 0 || value.x >= viewport.width || value.y >= viewport.height) return null
  const origin = originAtViewport(window.frozen, window.logicalOrigin, window.zoom, value)
  const col = hitAxis(metrics.cols, origin.x, window.zoom, value.x)
  const rowSlot = hitAxis(metrics.rows, origin.y, window.zoom, value.y)
  return col === null || rowSlot === null ? null : { row: metrics.order.displayToSource(rowSlot), col }
}
export function visibleLogicalViewport(window: ScrollWindow, viewport: ViewportSize): LogicalCellBox {
  state(window); viewportSize(viewport)
  return { ...window.logicalOrigin, width: viewport.width / window.zoom, height: viewport.height / window.zoom }
}
export function navigatorRequestFromA1(address: string, addressSpace: 'display' | 'source' = 'display'): NavigatorRequest {
  const at = parseA1(address)
  if (addressSpace !== 'display' && addressSpace !== 'source') throw new RangeError('Invalid navigation address space')
  return addressSpace === 'display' ? { addressSpace, at: { rowSlot: at.row, col: at.col } } : { addressSpace, at }
}
export interface NavigationContext { metrics: SheetMetrics; viewport: ViewportSize; view: Pick<SheetView, 'visibility' | 'baseHidden' | 'filter' | 'selection'> }
export type ActiveNavigationContext = Pick<NavigationContext, 'metrics' | 'view'>
export type HiddenTargetReason = 'authored-hidden' | 'filter-hidden'
export type NavigationKey = 'ArrowLeft' | 'ArrowRight' | 'ArrowUp' | 'ArrowDown' | 'Enter' | 'PageUp' | 'PageDown' | 'Home' | 'End'
interface NavigationBase { window: ScrollWindow; active: SourceCell | null; target: SourceCell | null; display: { rowSlot: DisplaySlot; col: number } | null; activate: boolean }
export type NavigationResult =
  | (NavigationBase & { kind: 'visible'; active: SourceCell; target: SourceCell; activate: true })
  | (NavigationBase & { kind: 'hidden'; target: SourceCell; hiddenReason: 'authored-hidden' | 'filter-hidden'; activate: false })
  | (NavigationBase & { kind: 'empty'; active: null; target: null; activate: false })
const floatBits = new DataView(new ArrayBuffer(8))
/** Exact preceding Float64 CSS coordinate, including subnormal sizes. */
function previousFloat(value: number): number {
  if (value === 0) return -Number.MIN_VALUE
  floatBits.setFloat64(0, value)
  const bits = floatBits.getBigUint64(0)
  floatBits.setBigUint64(0, value > 0 ? bits - 1n : bits + 1n)
  return floatBits.getFloat64(0)
}
/** Reject a viewport whose global Float64 origin cannot distinguish CSS/device
 * pixels. Otherwise wheel motion can disappear or an endpoint can become blank. */
function floatSpacing(value: number): number {
  if (value === 0) return Number.MIN_VALUE
  floatBits.setFloat64(0, value)
  const bits = floatBits.getBigUint64(0)
  floatBits.setBigUint64(0, bits + 1n)
  return floatBits.getFloat64(0) - value
}
function validatePixelPrecision(origin: number, viewportCss: number, zoom: number, deviceScale: number): void {
  if (viewportCss === 0) return
  const end = origin + viewportCss / zoom
  if (!Number.isFinite(end) || end <= origin) throw new RangeError('Scroll viewport pixel precision is not representable')
  // Float64 spacing can double at a power-of-two boundary inside the pane.
  const spacing = Math.max(floatSpacing(origin), floatSpacing(end))
  const pixelStep = 1 / zoom / Math.max(1, deviceScale)
  if (pixelStep < spacing) throw new RangeError('Scroll viewport pixel precision is not representable')
}
function sourceBodyBounds(context: ActiveNavigationContext): { first: number; last: number } {
  const filter = normalizeFilter(context.view.filter)
  return filter ? { first: filter.range.r0 + filter.range.headerRows, last: filter.range.r1 - filter.range.totalsRows }
    : { first: 0, last: context.metrics.rows.logicalCount - 1 }
}
function firstAxis(axis: AxisMetrics): number | null { return axis.totalCss() > 0 ? axis.locate(0) : null }
function lastAxis(axis: AxisMetrics): number | null { return axis.totalCss() > 0 ? axis.locate(previousFloat(axis.totalCss())) : null }
function nearestColumn(axis: AxisMetrics, col: number): number | null {
  if (!axis.logicalCount || !axis.totalCss()) return null
  if (col >= 0 && col < axis.logicalCount && axis.size(col) > 0) return col
  const at = axisStart(axis, clamp(col, 0, axis.logicalCount))
  const forward = axis.locate(at), backward = axis.locate(previousFloat(at))
  if (forward === null) return backward
  if (backward === null) return forward
  return forward - col <= col - backward ? forward : backward
}
/** Source body membership is independent of where a permutation displays it.
 * Search only its display envelope, choosing the smaller of source membership
 * and visible display candidates so a tiny body in a large band stays bounded. */
function bodyCandidate(context: ActiveNavigationContext, from: number, direction: -1 | 1): number | null {
  const { metrics, view } = context
  const bounds = sourceBodyBounds(context)
  bounds.last = Math.min(bounds.last, metrics.rows.logicalCount - 1)
  if (bounds.first > bounds.last || !metrics.rows.totalCss()) return null
  let first = bounds.first, last = bounds.last
  for (const band of metrics.order.bands) {
    if (band.r0 <= bounds.last && band.r1 >= bounds.first) {
      first = Math.min(first, band.r0)
      last = Math.max(last, band.r1)
    }
  }
  if (direction === 1) first = Math.max(first, from)
  else last = Math.min(last, from)
  if (first > last) return null
  const startRank = view.visibility.rank(first), endRank = view.visibility.rank(last + 1)
  if (bounds.last - bounds.first + 1 < endRank - startRank) {
    let best: number | null = null
    for (let source = bounds.first; source <= bounds.last; source++) {
      const slot = metrics.order.sourceToDisplay(source)
      if (slot >= first && slot <= last && view.visibility.visible(slot) && metrics.rows.size(slot) > 0
        && (best === null || (direction === 1 ? slot < best : slot > best))) best = slot
    }
    return best
  }
  let candidate = view.visibility.select(direction === 1 ? startRank : endRank - 1)
  while (candidate !== null && candidate >= first && candidate <= last) {
    if (metrics.rows.size(candidate) === 0) {
      const at = metrics.rows.start(candidate)
      const next = metrics.rows.locate(direction === 1 ? at : previousFloat(at))
      if (next === null) return null
      candidate = view.visibility.select(direction === 1 ? view.visibility.rank(next) : view.visibility.rank(next + 1) - 1)
      continue
    }
    const source = metrics.order.displayToSource(candidate)
    if (source >= bounds.first && source <= bounds.last) return candidate
    candidate = view.visibility.next(candidate, direction)
  }
  return null
}
/** Nearest visible display slot of any kind in display order, forward on ties.
 * Slots hidden by the mask or collapsed to zero height never qualify. */
function nearestVisibleSlot(context: ActiveNavigationContext, slot: number): DisplaySlot | null {
  if (!context.metrics.rows.totalCss()) return null
  const forward = rowNeighbor(context, slot, 1)
  const backward = rowNeighbor(context, slot, -1)
  if (forward === null) return backward
  if (backward === null) return forward
  return forward - slot <= slot - backward ? forward : backward
}
/** Shared source-focus reconciliation for S2 scrolling and S6 selection. */
export function retainActive(context: ActiveNavigationContext, active: SourceCell | null): SourceCell | null {
  if (!active || !isValidCoordinate(active.row, active.col)) return null
  const { metrics, view } = context
  if (active.row >= metrics.rows.logicalCount || active.col >= metrics.cols.logicalCount) return null
  if (metrics.sourceCellToLogical(active)) return Object.freeze({ row: active.row, col: active.col })
  const slot = metrics.order.sourceToDisplay(active.row)
  const col = nearestColumn(metrics.cols, active.col)
  if (col === null) return null
  // Only the column is hidden: the row itself is visible, so keep the source
  // row and move to the nearest visible column (§8.1 never substitutes rows).
  if (metrics.rows.size(slot) > 0 && view.visibility.visible(slot)) {
    return Object.freeze({ row: active.row, col })
  }
  // The row itself is hidden. A row hidden by the filter body relocates within
  // the body (protected headers/totals are not substitutes); any other hidden
  // row relocates to the nearest visible row of any kind, forward on ties.
  const filter = normalizeFilter(view.filter)
  const bounds = sourceBodyBounds(context)
  const inFilterBody = filter !== null && active.row >= bounds.first && active.row <= bounds.last
  if (!inFilterBody) {
    const targetSlot = nearestVisibleSlot(context, slot)
    return targetSlot === null ? null : Object.freeze({ row: metrics.order.displayToSource(targetSlot), col })
  }
  const forward = bodyCandidate(context, slot, 1), backward = bodyCandidate(context, slot, -1)
  const targetSlot = forward === null ? backward : backward === null ? forward : forward - slot <= slot - backward ? forward : backward
  return targetSlot === null ? null : Object.freeze({ row: metrics.order.displayToSource(targetSlot), col })
}
function validateTarget(metrics: SheetMetrics, target: SourceCell): void {
  validateCoordinate(target.row, target.col)
  if (target.row >= metrics.rows.logicalCount || target.col >= metrics.cols.logicalCount) throw new RangeError('Navigation target outside logical axes')
}
export function navigationTarget(metrics: SheetMetrics, request: NavigatorRequest): SourceCell {
  if (request.addressSpace !== 'source' && request.addressSpace !== 'display') throw new RangeError('Invalid navigation address space')
  const target = request.addressSpace === 'source' ? { ...request.at } : { row: metrics.order.displayToSource(request.at.rowSlot), col: request.at.col }
  validateTarget(metrics, target)
  return Object.freeze(target)
}
export function hiddenReasonForCell(context: ActiveNavigationContext, target: SourceCell, rowSlot = context.metrics.order.sourceToDisplay(target.row)): HiddenTargetReason {
  const filter = normalizeFilter(context.view.filter), bounds = sourceBodyBounds(context)
  return context.metrics.cols.size(target.col) === 0 ? 'authored-hidden'
    : filter && target.row >= bounds.first && target.row <= bounds.last && !context.view.visibility.visible(rowSlot) ? 'filter-hidden' : 'authored-hidden'
}
function visibleResult(window: ScrollWindow, metrics: SheetMetrics, target: SourceCell): NavigationResult {
  return { kind: 'visible', window, active: Object.freeze({ ...target }), target: Object.freeze({ ...target }), display: { rowSlot: metrics.order.sourceToDisplay(target.row), col: target.col }, activate: true }
}
function emptyResult(window: ScrollWindow): NavigationResult { return { kind: 'empty', window, active: null, target: null, display: null, activate: false } }
function ensureVisible(window: ScrollWindow, context: NavigationContext, target: SourceCell): ScrollWindow {
  const box = context.metrics.sourceCellToLogical(target)!
  const ensure = (origin: number, start: number, length: number, viewportCss: number, pinned: number) => {
    const span = viewportCss / window.zoom
    if (pinned > 0) return ensureAxisVisible(origin, start, length, span, pinned)
    if (length > span) return clamp(origin, start, start + length - span)
    return start < origin ? start : start + length > origin + span ? start + length - span : origin
  }
  return setScrollOffset(window, context.metrics, context.viewport, { x: ensure(window.logicalOrigin.x, box.x, box.width, context.viewport.width, window.frozen?.left ?? 0), y: ensure(window.logicalOrigin.y, box.y, box.height, context.viewport.height, window.frozen?.top ?? 0) })
}
/** Hidden requests open an inspector, never activate/unhide their target. A
 * valid existing active source identity is retained; an invalidated active cell
 * relocates to a visible body row with forward tie preference (§§3,8.1). */
export function navigateTo(window: ScrollWindow, context: NavigationContext, request: NavigatorRequest, currentActive: SourceCell | null = null): NavigationResult {
  state(window)
  const target = navigationTarget(context.metrics, request)
  const box = context.metrics.sourceCellToLogical(target)
  if (box) return visibleResult(setScrollOffset(window, context.metrics, context.viewport, bodyOriginFor(window.logicalOrigin, window.frozen, box)), context.metrics, target)
  const rowSlot = context.metrics.order.sourceToDisplay(target.row)
  const hiddenReason = hiddenReasonForCell(context, target, rowSlot)
  return { kind: 'hidden', window, active: retainActive(context, currentActive), target, display: { rowSlot, col: target.col }, activate: false, hiddenReason }
}
function rowNeighbor(context: ActiveNavigationContext, slot: number, direction: -1 | 1): number | null {
  const candidate = context.view.visibility.next(slot, direction)
  if (candidate === null) return null
  if (context.metrics.rows.size(candidate) > 0) return candidate
  const at = context.metrics.rows.start(candidate)
  return context.metrics.rows.locate(direction === 1 ? at : previousFloat(at))
}
function columnNeighbor(axis: AxisMetrics, col: number, direction: -1 | 1): number | null {
  return axis.locate(direction === 1 ? axisStart(axis, col + 1) : previousFloat(axis.start(col)))
}
/** Pure axis-end navigation; Ctrl+arrows jump to visible axis ends. Data-region
 * jumps need source occupancy and belong to a later source-aware controller. */
export function navigateKey(window: ScrollWindow, context: NavigationContext, key: NavigationKey, currentActive: SourceCell | null = null, options: { ctrl?: boolean } = {}): NavigationResult {
  state(window)
  let active = retainActive(context, currentActive)
  const seeded = !active
  let firstRow = firstAxis(context.metrics.rows)
  const firstCol = firstAxis(context.metrics.cols)
  if (firstRow === null || firstCol === null) return emptyResult(window)
  if (!active) {
    const slot = bodyCandidate(context, 0, 1)
    if (slot === null) return emptyResult(window)
    firstRow = slot
    active = { row: context.metrics.order.displayToSource(slot), col: firstCol }
  }
  const slot = context.metrics.order.sourceToDisplay(active.row)
  const lastRow = () => seeded ? bodyCandidate(context, context.metrics.rows.logicalCount - 1, -1)! : lastAxis(context.metrics.rows)!
  let targetSlot = slot, col = active.col
  let pageWindow: ScrollWindow | undefined
  switch (key) {
    case 'ArrowDown': case 'Enter': targetSlot = options.ctrl ? lastRow() : seeded ? slot : rowNeighbor(context, slot, 1) ?? slot; break
    case 'ArrowUp': targetSlot = options.ctrl ? firstRow : seeded ? slot : rowNeighbor(context, slot, -1) ?? slot; break
    case 'ArrowRight': col = options.ctrl ? lastAxis(context.metrics.cols)! : columnNeighbor(context.metrics.cols, col, 1) ?? col; break
    case 'ArrowLeft': col = options.ctrl ? firstCol : columnNeighbor(context.metrics.cols, col, -1) ?? col; break
    case 'Home': col = firstCol; if (options.ctrl) targetSlot = firstRow; break
    case 'End': col = lastAxis(context.metrics.cols)!; if (options.ctrl) targetSlot = lastRow(); break
    case 'PageDown': case 'PageUp': {
      // Initialization creates focus at a body record before page traversal.
      if (seeded) break
      const direction = key === 'PageDown' ? 1 : -1
      // A page is the scrolling body: the pane minus a pinned frozen strip.
      const pinned = window.frozen?.top ?? 0
      const pageCss = Math.max(window.zoom, context.viewport.height - pinned * window.zoom)
      pageWindow = scrollBy(window, context.metrics, context.viewport, { x: 0, y: direction * pageCss })
      // Clamp a far-away active cell into the visible body so the page target
      // always overlaps the moved viewport (an inclusive upper edge would land
      // exactly on its bottom border, outside it).
      const span = context.viewport.height / window.zoom
      const bodyTop = window.logicalOrigin.y + pinned
      const visibleAnchor = clamp(Math.max(context.metrics.rows.start(slot), bodyTop), bodyTop, previousFloat(window.logicalOrigin.y + span))
      const css = clamp(visibleAnchor + direction * pageCss / window.zoom, pinned, previousFloat(context.metrics.rows.totalCss()))
      targetSlot = context.metrics.rows.locate(css) ?? slot
      break
    }
    default: throw new RangeError('Unknown navigation key')
  }
  const target = { row: context.metrics.order.displayToSource(targetSlot), col }
  return visibleResult(pageWindow ?? ensureVisible(window, context, target), context.metrics, target)
}
