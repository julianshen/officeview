/** Resolves raw drawing anchor records (S3.3) into structural DrawingAnchor
 * records against one sheet geometry, shared by both source adapters so the
 * same file yields identical bounds (spec section 4.5).
 *
 * Bounds may over-approximate, never under-approximate: a two-cell anchor keeps
 * its `to` cell even at a zero offset, an extent ending exactly on a boundary
 * keeps the next cell, and authored-hidden rows/columns count as zero size (as
 * the static drawing placement does), so a one-cell extent reaches every cell
 * it can occupy. An extent past the last cell clamps to the sheet edge. */
import type { DrawingAnchor, DrawingMarker } from './contracts'
import type { AxisMetrics } from './contracts'
import type { XlsxAnchorMarker, XlsxDrawingAnchorRecord } from '../types'
import type { SheetGeometry } from '../../worker/viewerSession'
import { createAxisMetrics } from './axes'
import { MAX_COLS, MAX_ROWS } from './coordinates'

const EMU_PER_CSS = 9525

const toMarker = (m: XlsxAnchorMarker): DrawingMarker => ({ row: m.row, col: m.col, offsetXEmu: m.colOffEmu, offsetYEmu: m.rowOffEmu })
/** Cell containing an absolute EMU position (the last cell past the axis end). */
function cellAt(axis: AxisMetrics, emu: number, last: number): number {
  const css = emu / EMU_PER_CSS
  if (!(css < axis.totalCss())) return last
  return axis.locate(Math.max(0, css)) ?? last
}

export function resolveDrawingAnchors(records: readonly XlsxDrawingAnchorRecord[], geometry: SheetGeometry): DrawingAnchor[] {
  if (!records.length) return []
  // Placement axes: authored heights/widths with authored-hidden rows (and, when
  // the source reports them, columns) collapsed.
  const rows = createAxisMetrics({
    logicalCount: MAX_ROWS, defaultSizeCss: geometry.rows.defaultSizeCss, overrides: geometry.rowOverrides,
    collapsed: geometry.baseHidden.intervals.map(interval => ({ start: interval.r0, end: interval.r1 })),
  })
  const cols = createAxisMetrics({
    logicalCount: MAX_COLS, defaultSizeCss: geometry.cols.defaultSizeCss, overrides: geometry.colOverrides,
    ...(geometry.hiddenCols?.length ? { collapsed: geometry.hiddenCols } : {}),
  })
  const x = (m: XlsxAnchorMarker) => cols.start(m.col) * EMU_PER_CSS + m.colOffEmu
  const y = (m: XlsxAnchorMarker) => rows.start(m.row) * EMU_PER_CSS + m.rowOffEmu
  return records.map((record): DrawingAnchor => {
    if (record.kind === 'absolute') {
      return { id: record.id, kind: 'absolute', extent: { x: record.x, y: record.y, width: record.cx, height: record.cy } }
    }
    const from = record.from
    if (record.kind === 'oneCell') {
      const r1 = Math.max(from.row, cellAt(rows, y(from) + record.cy, MAX_ROWS - 1))
      const c1 = Math.max(from.col, cellAt(cols, x(from) + record.cx, MAX_COLS - 1))
      return {
        id: record.id, kind: 'oneCell', from: toMarker(from),
        extent: { x: from.colOffEmu, y: from.rowOffEmu, width: record.cx, height: record.cy },
        bounds: { r0: from.row, r1, c0: from.col, c1 },
      }
    }
    const to = record.to
    return {
      id: record.id, kind: 'twoCell', from: toMarker(from), to: toMarker(to),
      extent: { x: from.colOffEmu, y: from.rowOffEmu, width: Math.max(0, x(to) - x(from)), height: Math.max(0, y(to) - y(from)) },
      bounds: { r0: Math.min(from.row, to.row), r1: Math.max(from.row, to.row), c0: Math.min(from.col, to.col), c1: Math.max(from.col, to.col) },
    }
  })
}
