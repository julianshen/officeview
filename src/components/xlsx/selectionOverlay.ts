/**
 * Grid selection overlay geometry for the interactive sheet pane (spec §§9.1,
 * 11.1, 11.3). Below the per-frame cap the overlay uses one record per visible
 * cell (plus a neighbour ring so outlines are computed before clipping). Past
 * the cap the selection is still painted: its source membership is mapped to
 * display slot ranges and coalesced into bounded device-pixel groups.
 */
import type { GridSelection, ProfileId } from '../../xlsx/view/contracts'
import type { SourceCell } from '../../xlsx/view/coordinates'
import type { ProjectedSheetMetrics } from '../../xlsx/view/projection'
import { rowContains } from '../../xlsx/view/selection'
import type { GridOverlayRecord } from '../../core/overlay'

interface Box { x: number; y: number; width: number; height: number }
/** Per-row mapping work inside permuted bands, per frame. Past it a band's selected
 * rows are summarized by the band's display range (declared coarse state). */
export const COARSE_SELECTION_ROW_BUDGET = 65_536
/** Coalesced groups per axis per frame. */
export const COARSE_SELECTION_GROUPS = 1024

/** Bounded per-cell overlay records for the pane plus one neighbour ring; null
 * past the profile's frame cap (the declared coarse state, spec §11.1). */
export function overlayRecords(metrics: ProjectedSheetMetrics, pane: Box, profile: ProfileId): GridOverlayRecord[] | null {
  const cap = profile === 'M' ? 10_000 : 20_000
  const rows = metrics.rows.visibleBetween(pane.y, pane.y + pane.height), cols = metrics.cols.visibleBetween(pane.x, pane.x + pane.width)
  if ((rows.count + 2) * (cols.count + 2) > cap) return null
  const slots = expand([...rows.runs], metrics.rows.logicalCount), columns = expand([...cols.runs], metrics.cols.logicalCount)
  const records: GridOverlayRecord[] = []
  for (const slot of slots) {
    const y = metrics.rows.start(slot), height = metrics.rows.size(slot)
    if (!(height > 0)) continue
    const row = metrics.order.displayToSource(slot)
    for (const col of columns) {
      const width = metrics.cols.size(col)
      if (width > 0) records.push({ source: { row, col }, bounds: { x: metrics.cols.start(col), y, width, height } })
    }
  }
  return records
}
function expand(runs: readonly { start: number; end: number }[], count: number): number[] {
  const out: number[] = []
  if (!runs.length) return out
  const first = runs[0].start, last = runs[runs.length - 1].end
  if (first > 0) out.push(first - 1)
  for (const run of runs) for (let index = run.start; index <= run.end; index++) out.push(index)
  if (last + 1 < count) out.push(last + 1)
  return out
}
export function selectionContains(selection: GridSelection, at: SourceCell): boolean {
  // Worker `ref` row sets need a resolver this in-thread path does not hold yet.
  return selection.ranges.some(part => part.rows.kind !== 'ref' && part.cols.some(c => at.col >= c.c0 && at.col <= c.c1) && rowContains(part.rows, at.row))
}

type Span = [number, number]
/** Sort, snap to device pixels and merge touching spans, then coarsen until at most
 * COARSE_SELECTION_GROUPS remain. */
function coalesce(spans: Span[], lo: number, hi: number, scale: number): Span[] {
  let unit = 1 / scale
  for (;;) {
    const snapped = spans
      .map(([a, b]): Span => [Math.max(lo, Math.floor(a / unit) * unit), Math.min(hi, Math.max(Math.floor(a / unit) * unit + unit, Math.ceil(b / unit) * unit))])
      .filter(([a, b]) => b > a)
      .sort((x, y) => x[0] - y[0])
    const out: Span[] = []
    for (const span of snapped) {
      const last = out[out.length - 1]
      if (last && span[0] <= last[1]) last[1] = Math.max(last[1], span[1])
      else out.push([...span] as Span)
    }
    if (out.length <= COARSE_SELECTION_GROUPS) return out
    unit *= 2
  }
}
/** Selected display-slot ranges of one source row set (bounded work). */
function slotRanges(rows: GridSelection['ranges'][number]['rows'], metrics: ProjectedSheetMetrics, budget: { rows: number }): Array<[number, number]> {
  if (rows.kind === 'ref') return []
  const sourceIntervals: Array<[number, number]> = []
  if (rows.kind === 'intervals') for (const r of rows.intervals) sourceIntervals.push([r.r0, r.r1])
  else if (rows.kind === 'list') for (const row of rows.rows) sourceIntervals.push([row, row])
  else {
    // Walk set bits word by word; empty words cost one test.
    rows.words.forEach((word, index) => {
      for (let bit = 0; word !== 0 && bit < 32; bit++) if (Math.floor(word / 2 ** bit) % 2) { const row = rows.r0 + index * 32 + bit; sourceIntervals.push([row, row]) }
    })
  }
  const bands = [...metrics.order.bands].sort((a, b) => a.r0 - b.r0)
  const out: Array<[number, number]> = []
  for (const [r0, r1] of sourceIntervals) {
    let at = r0
    for (const band of bands) {
      if (band.r1 < at || band.r0 > r1) continue
      if (band.r0 > at) out.push([at, band.r0 - 1]) // identity stretch: slot == source row
      const b0 = Math.max(at, band.r0), b1 = Math.min(r1, band.r1)
      if (b1 - b0 + 1 <= budget.rows) {
        budget.rows -= b1 - b0 + 1
        for (let row = b0; row <= b1; row++) { const slot = metrics.order.sourceToDisplay(row); out.push([slot, slot]) }
      } else out.push([band.r0, band.r1])
      at = b1 + 1
      if (at > r1) break
    }
    if (at <= r1) out.push([at, r1])
  }
  return out
}
/** Coarse selection records: coalesced device-pixel groups of the selected
 * membership clipped to the pane. `contains` is true for every record. */
export function coarseSelectionRecords(selection: GridSelection, metrics: ProjectedSheetMetrics, pane: Box, scale: number): GridOverlayRecord[] {
  const budget = { rows: COARSE_SELECTION_ROW_BUDGET }
  const records: GridOverlayRecord[] = []
  const yLo = pane.y, yHi = pane.y + pane.height, xLo = pane.x, xHi = pane.x + pane.width
  for (const part of selection.ranges) {
    const ys = coalesce(slotRanges(part.rows, metrics, budget)
      .map(([s0, s1]): Span => [metrics.rows.start(s0), metrics.rows.start(s1) + metrics.rows.size(s1)])
      .filter(([a, b]) => b > yLo && a < yHi), yLo, yHi, scale)
    const xs = coalesce(part.cols
      .map((c): Span => [metrics.cols.start(c.c0), metrics.cols.start(c.c1) + metrics.cols.size(c.c1)])
      .filter(([a, b]) => b > xLo && a < xHi), xLo, xHi, scale)
    for (const [y0, y1] of ys) for (const [x0, x1] of xs) records.push({ source: { row: 0, col: 0 }, bounds: { x: x0, y: y0, width: x1 - x0, height: y1 - y0 } })
  }
  return records
}
