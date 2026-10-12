/**
 * Viewport geometry of a worker-backed (paged) sheet (plan S5.4): the same
 * full-capacity axes and projection the object path builds, from the worker's
 * compact layout, with cell records served by the bounded row cache. The
 * object planner and painter run unchanged on it; rows the cache does not know
 * yet are reported separately (RemoteSheetStore.pendingIn) and paint pending.
 */
import type { RowOrder, RowVisibility } from '../../xlsx/view/contracts'
import { createSheetMetrics } from '../../xlsx/view/projection'
import { RectangleIndex } from '../../xlsx/view/structure'
import type { SheetGeometry } from '../../xlsx/view/viewport'
import type { RemoteWorkbook } from '../../worker/viewerRemoteSource'

export function pagedSheetGeometry(remote: RemoteWorkbook, sheetId: string, order: RowOrder, visibility: RowVisibility): SheetGeometry {
  const axes = remote.geometry(sheetId)
  const store = remote.store(sheetId)
  const merges = remote.merges(sheetId).merges
  return Object.freeze({
    kind: 'ready' as const,
    metrics: createSheetMetrics({ rows: axes.rows, cols: axes.cols, order, visibility }),
    diagnostics: Object.freeze([]),
    cellsOf: (sourceRow: number) => store.cellsOf(sourceRow),
    merges: RectangleIndex.build(merges.map(range => ({ rect: { r0: range.minRow, r1: range.maxRow, c0: range.minCol, c1: range.maxCol }, value: range }))),
    // The paged reader keeps no drawing parts.
    drawingsIn: () => [],
  })
}
