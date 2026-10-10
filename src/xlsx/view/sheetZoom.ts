import { clampZoom } from '../../core/zoom'
import { SHEET_MIN_ZOOM, SHEET_MAX_ZOOM } from './contracts'
import type { SheetZoom, SheetMetrics, ScrollWindow } from './contracts'
import { setScrollZoom } from './scrollWindow'

export interface SheetZoomPoint { x: number; y: number }
export interface SheetZoomViewport { width: number; height: number }
/** deltaY is normalized to CSS pixels by the event adapter. */
export interface SheetWheelInput { deltaY: number; ctrlKey?: boolean; metaKey?: boolean }

export interface SheetZoomState {
  readonly docId: string
  readonly activeSheetId: string
  readonly zoomBySheet: ReadonlyMap<string, SheetZoom>
}

export type SheetZoomAction =
  | { kind: 'setZoom'; docId: string; sheetId: string; zoom: SheetZoom }
  | { kind: 'switchSheet'; docId: string; sheetId: string }
  | { kind: 'sheetChanged'; docId: string; sheetId: string; reason: 'sort' | 'filter' | 'values' | 'font' | 'revision' }
  | { kind: 'replaceDocument'; docId: string; nextDocId: string; activeSheetId: string }

export function clampSheetZoom(zoom: number): SheetZoom {
  return clampZoom(zoom, SHEET_MIN_ZOOM, SHEET_MAX_ZOOM)
}

/** Documented keyboard step table: in/out are exact inverses, including at the
 * bounds (off-table zooms snap to the nearest step in the key direction). */
export const SHEET_ZOOM_STEPS: readonly SheetZoom[] = [0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4]

export function snapSheetZoom(zoom: number, direction: 1 | -1): SheetZoom {
  if (direction > 0) {
    for (const step of SHEET_ZOOM_STEPS) if (step > zoom) return step
    return SHEET_MAX_ZOOM
  }
  for (let index = SHEET_ZOOM_STEPS.length - 1; index >= 0; index--) {
    if (SHEET_ZOOM_STEPS[index] < zoom) return SHEET_ZOOM_STEPS[index]
  }
  return SHEET_MIN_ZOOM
}

export function createSheetZoomState(docId: string, activeSheetId: string): SheetZoomState {
  return { docId, activeSheetId, zoomBySheet: new Map() }
}

export function sheetZoomFor(state: SheetZoomState, sheetId = state.activeSheetId): SheetZoom {
  return state.zoomBySheet.get(sheetId) ?? 1
}

export function reduceSheetZoom(state: SheetZoomState, action: SheetZoomAction): SheetZoomState {
  if (action.docId !== state.docId) return state
  switch (action.kind) {
    case 'setZoom': {
      const zoomBySheet = new Map(state.zoomBySheet)
      zoomBySheet.set(action.sheetId, clampSheetZoom(action.zoom))
      return { ...state, zoomBySheet }
    }
    case 'switchSheet': return { ...state, activeSheetId: action.sheetId }
    case 'replaceDocument': return createSheetZoomState(action.nextDocId, action.activeSheetId)
    case 'sheetChanged': return state
  }
}

export function zoomSheetAt(
  start: ScrollWindow,
  metrics: SheetMetrics,
  viewport: SheetZoomViewport,
  nextZoom: number,
  anchorCSS: SheetZoomPoint,
): ScrollWindow {
  if (!Number.isFinite(nextZoom)) throw new RangeError('Sheet zoom must be finite')
  return setScrollZoom(start, metrics, viewport, clampSheetZoom(nextZoom), anchorCSS)
}

export function keyboardSheetZoom(
  start: ScrollWindow,
  metrics: SheetMetrics,
  viewport: SheetZoomViewport,
  direction: 1 | -1,
  anchorCSS: SheetZoomPoint = { x: viewport.width / 2, y: viewport.height / 2 },
): ScrollWindow {
  return zoomSheetAt(start, metrics, viewport, snapSheetZoom(start.zoom, direction), anchorCSS)
}

export function wheelSheetZoom(
  start: ScrollWindow,
  metrics: SheetMetrics,
  viewport: SheetZoomViewport,
  input: SheetWheelInput,
  anchorCSS: SheetZoomPoint,
): ScrollWindow {
  // Event input, not document state: unmodified or non-finite deltas keep the
  // window unchanged instead of throwing (zoom arguments still reject).
  if ((!input.ctrlKey && !input.metaKey) || !Number.isFinite(input.deltaY)) return start
  // Saturate before exponentiation so large finite deltas keep their direction.
  const exponent = -input.deltaY * 0.002
  const nextZoom = exponent >= Math.log(SHEET_MAX_ZOOM / start.zoom) ? SHEET_MAX_ZOOM
    : exponent <= Math.log(SHEET_MIN_ZOOM / start.zoom) ? SHEET_MIN_ZOOM
    : start.zoom * Math.exp(exponent)
  return zoomSheetAt(start, metrics, viewport, nextZoom, anchorCSS)
}

export function pinchSheetZoom(
  start: ScrollWindow,
  metrics: SheetMetrics,
  viewport: SheetZoomViewport,
  startDistance: number,
  currentDistance: number,
  startAnchorCSS: SheetZoomPoint,
  currentAnchorCSS: SheetZoomPoint,
): ScrollWindow {
  // Gesture distances are event inputs like wheel deltas: invalid values keep
  // the window unchanged instead of throwing.
  if (!Number.isFinite(startDistance) || startDistance <= 0
    || !Number.isFinite(currentDistance) || currentDistance <= 0) return start
  const maxRatio = SHEET_MAX_ZOOM / SHEET_MIN_ZOOM
  const ratio = Math.max(1 / maxRatio, Math.min(maxRatio, currentDistance / startDistance))
  return setScrollZoom(start, metrics, viewport, clampSheetZoom(start.zoom * ratio), startAnchorCSS, currentAnchorCSS)
}
