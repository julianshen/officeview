/**
 * Viewport geometry of a worker-backed (paged) sheet (plan S5.4): the same
 * full-capacity axes and projection the object path builds, from the worker's
 * compact layout, with cell records served by the bounded row cache. The
 * object planner and painter run unchanged on it; rows the cache does not know
 * yet are reported separately (RemoteSheetStore.pendingIn) and paint pending.
 */
import { bindOverflowCache, type OverflowCache } from '../../xlsx/view/overflowCache'
import type { RowOrder, RowVisibility } from '../../xlsx/view/contracts'
import { createSheetMetrics } from '../../xlsx/view/projection'
import { RectangleIndex } from '../../xlsx/view/structure'
import type { SheetGeometry } from '../../xlsx/view/viewport'
import type { RemoteWorkbook } from '../../worker/viewerRemoteSource'
import { axisBytes, projectionBytes } from '../../xlsx/view/mainLedger'
import { AxisMetricsError, createAxisMetrics } from '../../xlsx/view/axes'
import { SIZED_DRAWING_BYTES, hasSessionSizes, mergeAxisOverrides, type SessionSizes } from '../../xlsx/view/axisSizes'
import { MAX_COLS, MAX_ROWS } from '../../xlsx/view/coordinates'
import { axisDiagnostic, drawingQuery, sizedDrawings } from '../../xlsx/view/viewport'
import type { WorkerDiagnostic } from '../../xlsx/view/contracts'
import type { XlsxDrawing } from '../../xlsx/types'

/** `sizes` (spec §5.6) are the sheet's view-only session sizes: the axes are
 * rebuilt from the worker's compact layout with them applied, so every
 * main-thread consumer (planner, hit-testing, overlays, a11y) agrees. Worker
 * viewport plans use the workbook geometry and are not consulted while a
 * sheet has session sizes (InteractiveSheetDoc / tilePlanSource). */
export function pagedSheetGeometry(remote: RemoteWorkbook, sheetId: string, order: RowOrder, visibility: RowVisibility, cache: OverflowCache = remote.overflowFor(), sizes?: SessionSizes): SheetGeometry {
  const axes = remote.geometry(sheetId)
  const store = remote.store(sheetId)
  const merges = remote.merges(sheetId).merges
  const drawings = remote.drawings(sheetId)
  let { rows, cols } = axes
  let sizedBytes = 0
  const plainDrawings: SheetGeometry['drawingsIn'] = (...query) => drawings.drawingsIn(...query)
  let drawingsIn = plainDrawings
  const sizeDiagnostics: WorkerDiagnostic[] = []
  let metrics: ReturnType<typeof createSheetMetrics> | undefined
  if (hasSessionSizes(sizes)) {
    const layout = remote.layout(sheetId)
    const hiddenCols = layout.hiddenCols ?? []
    try {
      if (sizes.rows.size) rows = createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: axes.rows.defaultSizeCss, overrides: mergeAxisOverrides(layout.rowOverrides, sizes.rows) })
      if (sizes.cols.size) cols = createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: axes.cols.defaultSizeCss, overrides: mergeAxisOverrides(layout.colOverrides, sizes.cols), ...(hiddenCols.length ? { collapsed: hiddenCols } : {}) })
      metrics = createSheetMetrics({ rows, cols, order, visibility })
    } catch (error) {
      // A legal session size beside arbitrary authored sizes may still be
      // unrepresentable: keep the authored geometry, diagnosed (§5.6).
      if (!(error instanceof AxisMetricsError)) throw error
      sizeDiagnostics.push(axisDiagnostic(error, 'viewport'))
      rows = axes.rows; cols = axes.cols; metrics = undefined
    }
    sizedBytes = (rows === axes.rows ? 0 : axisBytes(rows)) + (cols === axes.cols ? 0 : axisBytes(cols))
    // Drawing content arrives (and is evicted) after this geometry is built:
    // place each published list once, keyed weakly so the geometry never pins it.
    const placed = new WeakMap<readonly XlsxDrawing[], SheetGeometry['drawingsIn']>()
    const hiddenRows = layout.hidden.map(interval => ({ start: interval.r0, end: interval.r1 }))
    if (metrics) drawingsIn = (...query) => {
      const list = drawings.drawings
      if (!list.length) return drawings.drawingsIn(...query)
      let sized = placed.get(list)
      if (!sized) {
        const moved = sizedDrawings(list, sizes, layout.rowOverrides, hiddenRows, layout.colOverrides, hiddenCols)
        sized = drawingQuery(moved)
        placed.set(list, sized)
        remote.mainLedger.chargeWhileReachable('drawings', [sized], moved.length * SIZED_DRAWING_BYTES)
      }
      return sized(...query)
    }
  }
  metrics ??= createSheetMetrics({ rows, cols, order, visibility })
  const mergeIndex = RectangleIndex.build(merges.map(range => ({ rect: { r0: range.minRow, r1: range.maxRow, c0: range.minCol, c1: range.maxCol }, value: range })))
  const geometry = Object.freeze({
    kind: 'ready' as const,
    metrics,
    diagnostics: Object.freeze([...drawings.diagnostics, ...sizeDiagnostics]),
    cellsOf: (sourceRow: number) => store.cellsOf(sourceRow),
    merges: mergeIndex,
    // Read at query time so geometry does not pin an evicted drawing payload.
    drawingsIn,
  })
  // Projection buffers (dense prefixes of permuted bands) and the merge index
  // are charged while held: by the sheet entry that shows this geometry and any
  // continuation reading it (§12.4). Unheld, it lasts while reachable.
  remote.mainLedger.chargeWhileHeld('geometry', [geometry], projectionBytes(metrics) + mergeIndex.estimatedBytes + sizedBytes)
  bindOverflowCache(geometry, cache)
  return geometry
}
