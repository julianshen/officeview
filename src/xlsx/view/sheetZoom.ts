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
  return setScrollZoom(start, metrics, viewport, clampSheetZoom(nextZoom), anchorCSS)
}

export function keyboardSheetZoom(
  start: ScrollWindow,
  metrics: SheetMetrics,
  viewport: SheetZoomViewport,
  direction: 1 | -1,
  anchorCSS: SheetZoomPoint = { x: viewport.width / 2, y: viewport.height / 2 },
): ScrollWindow {
  const nextZoom = direction > 0 ? start.zoom * 1.5 : start.zoom / 1.5
  return zoomSheetAt(start, metrics, viewport, nextZoom, anchorCSS)
}

export function wheelSheetZoom(
  start: ScrollWindow,
  metrics: SheetMetrics,
  viewport: SheetZoomViewport,
  input: SheetWheelInput,
  anchorCSS: SheetZoomPoint,
): ScrollWindow {
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
  if (!Number.isFinite(startDistance) || startDistance <= 0
    || !Number.isFinite(currentDistance) || currentDistance <= 0) return start
  const maxRatio = SHEET_MAX_ZOOM / SHEET_MIN_ZOOM
  const ratio = Math.max(1 / maxRatio, Math.min(maxRatio, currentDistance / startDistance))
  return setScrollZoom(start, metrics, viewport, clampSheetZoom(start.zoom * ratio), startAnchorCSS, currentAnchorCSS)
}
