/**
 * Stage 4a (S4 §6.4): the viewport geometry of a stream sheet, rebuilt from its `sheet-layout` alone. The axes are the
 * server's, built the way `pagedSheetGeometry` builds them (`createAxisMetrics`, then `createSheetMetrics`; 4b: with the
 * session's sizes merged in), so every
 * position the client hit-tests or overlays agrees with the tiles the server paints. A stream sheet has no cells and
 * no drawings on the client (they are pixels in the tiles): `cellsOf` is always `undefined` and `drawingsIn` empty.
 */
import type { StreamSheetLayout } from '../../stream/sheet-wire'
import type { XlsxMergeRange } from '../types'
import { createAxisMetrics } from './axes'
import { mergeAxisOverrides } from './axisSizes'
import { MAX_COLS, MAX_ROWS } from './coordinates'
import { createSheetMetrics } from './projection'
import { RectangleIndex } from './structure'
import type { SheetGeometry } from './viewport'

const NO_DRAWINGS: readonly never[] = Object.freeze([])

/** Stage 4b (P11): the layout's authored sizes with the session's merged over them, as `buildSheetGeometry` merges
 * session sizes (`mergeAxisOverrides`); collapsed runs still win in the axis. */
const withSession = (base: StreamSheetLayout['rows'], session: Array<[number, number]> | undefined) => session?.length ? mergeAxisOverrides(base, new Map(session)) : base

export function streamSheetGeometry(l: StreamSheetLayout): SheetGeometry {
  const rows = createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: l.defaults.rowPx, overrides: withSession(l.rows, l.sessionRows), ...(l.collapsedRows.length ? { collapsed: l.collapsedRows } : {}) })
  const cols = createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: l.defaults.colPx, overrides: withSession(l.cols, l.sessionCols), ...(l.collapsedCols.length ? { collapsed: l.collapsedCols } : {}) })
  const metrics = createSheetMetrics({ rows, cols })
  const merges = RectangleIndex.build(l.merges.map((r): { rect: typeof r; value: XlsxMergeRange } => ({ rect: { r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1 }, value: { minRow: r.r0, maxRow: r.r1, minCol: r.c0, maxCol: r.c1 } })))
  return Object.freeze({
    kind: 'ready' as const,
    metrics,
    diagnostics: Object.freeze([]),
    cellsOf: () => undefined,
    merges,
    drawingsIn: () => NO_DRAWINGS,
  })
}
