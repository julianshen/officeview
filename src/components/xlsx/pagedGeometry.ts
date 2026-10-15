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
import { projectionBytes } from '../../xlsx/view/mainLedger'

export function pagedSheetGeometry(remote: RemoteWorkbook, sheetId: string, order: RowOrder, visibility: RowVisibility, cache: OverflowCache = remote.overflowFor()): SheetGeometry {
  const axes = remote.geometry(sheetId)
  const store = remote.store(sheetId)
  const merges = remote.merges(sheetId).merges
  const drawings = remote.drawings(sheetId)
  const metrics = createSheetMetrics({ rows: axes.rows, cols: axes.cols, order, visibility })
  const mergeIndex = RectangleIndex.build(merges.map(range => ({ rect: { r0: range.minRow, r1: range.maxRow, c0: range.minCol, c1: range.maxCol }, value: range })))
  const geometry = Object.freeze({
    kind: 'ready' as const,
    metrics,
    diagnostics: Object.freeze([...drawings.diagnostics]),
    cellsOf: (sourceRow: number) => store.cellsOf(sourceRow),
    merges: mergeIndex,
    // Read at query time so geometry does not pin an evicted drawing payload.
    drawingsIn: (...query: Parameters<SheetGeometry['drawingsIn']>) => drawings.drawingsIn(...query),
  })
  // Projection buffers (dense prefixes of permuted bands) and the merge index
  // are charged while held: by the sheet entry that shows this geometry and any
  // continuation reading it (§12.4). Unheld, it lasts while reachable.
  remote.mainLedger.chargeWhileHeld('geometry', [geometry], projectionBytes(metrics) + mergeIndex.estimatedBytes)
  bindOverflowCache(geometry, cache)
  return geometry
}
