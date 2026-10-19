import type { CanvasColors } from '../canvasColors'
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
/** Per-frame membership work (band slots tested, list rows/bitset words visited).
 * Past it the remaining slots are skipped, never painted (declared incomplete). */
export const COARSE_SELECTION_ROW_BUDGET = 65_536

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
/** Total coarse records per frame across every part and both axes. */
export const COARSE_SELECTION_RECORDS = 4096

/** Snap spans to `unit` and merge overlapping/touching ones (clipped to [lo,hi]). */
function snapMerge(spans: readonly Span[], lo: number, hi: number, unit: number): Span[] {
  const snapped = spans
    .map(([a, b]): Span => {
      const start = Math.max(lo, Math.floor(a / unit) * unit)
      return [start, Math.min(hi, Math.max(start + unit, Math.ceil(b / unit) * unit))]
    })
    .filter(([a, b]) => b > a)
    .sort((x, y) => x[0] - y[0])
  const out: Span[] = []
  for (const span of snapped) {
    const last = out[out.length - 1]
    if (last && span[0] <= last[1]) last[1] = Math.max(last[1], span[1])
    else out.push([span[0], span[1]])
  }
  return out
}

type RowSet = GridSelection['ranges'][number]['rows']
function lowerBound(rows: Uint32Array, value: number): number {
  let lo = 0, hi = rows.length
  while (lo < hi) { const mid = Math.floor((lo + hi) / 2); if (rows[mid] < value) lo = mid + 1; else hi = mid }
  return lo
}
/** Selected source rows within the identity stretch [a,b] (slot == source row). */
function identitySpans(rows: RowSet, a: number, b: number, budget: { work: number }, out: Array<[number, number]>): boolean {
  if (rows.kind === 'intervals') {
    for (const r of rows.intervals) { const lo = Math.max(a, r.r0), hi = Math.min(b, r.r1); if (lo <= hi) out.push([lo, hi]) }
    return true
  }
  if (rows.kind === 'list') {
    for (let at = lowerBound(rows.rows, a); at < rows.rows.length && rows.rows[at] <= b; at++) {
      if (--budget.work < 0) return false
      out.push([rows.rows[at], rows.rows[at]])
    }
    return true
  }
  if (rows.kind === 'bitset') {
    const first = Math.max(0, Math.floor((a - rows.r0) / 32)), last = Math.min(rows.words.length - 1, Math.floor((b - rows.r0) / 32))
    for (let index = first; index <= last; index++) {
      if (--budget.work < 0) return false
      const word = rows.words[index]
      if (word === 0) continue
      for (let bit = 0; bit < 32; bit++) {
        const row = rows.r0 + index * 32 + bit
        if (row >= a && row <= b && Math.floor(word / 2 ** bit) % 2) out.push([row, row])
      }
    }
  }
  return true
}

/**
 * Coarse selection records for a pane past the per-cell cap: EXACT visible
 * membership, never an enclosing approximation. Identity stretches intersect the
 * selection's source sets with the visible slot range; permuted band slots are
 * tested one by one against a per-frame work budget, and when it runs out the
 * remaining slots are skipped (`incomplete`), never painted. Output is bounded
 * across all parts and both axes by COARSE_SELECTION_RECORDS through pixel-group
 * coalescing; coarsening beyond one device pixel also marks `incomplete`.
 */
export function coarseSelectionRecords(selection: GridSelection, metrics: ProjectedSheetMetrics, pane: Box, scale: number): { records: GridOverlayRecord[]; incomplete: boolean } {
  const yLo = pane.y, yHi = pane.y + pane.height, xLo = pane.x, xHi = pane.x + pane.width
  const budget = { work: COARSE_SELECTION_ROW_BUDGET }
  let incomplete = false
  // Visible display slot range of the pane (first/last positive-size slots).
  let vs0 = -1, vs1 = -1, runs = 0
  for (const run of metrics.rows.visibleBetween(yLo, yHi).runs) {
    if (vs0 < 0) vs0 = run.start
    vs1 = run.end
    if (++runs > COARSE_SELECTION_ROW_BUDGET) { incomplete = true; break }
  }
  if (vs0 < 0) return { records: [], incomplete }
  const bands = [...metrics.order.bands].filter(band => band.r1 >= vs0 && band.r0 <= vs1).sort((x, y) => x.r0 - y.r0)
  const parts: Array<{ ys: Span[]; xs: Span[] }> = []
  for (const part of selection.ranges) {
    if (part.rows.kind === 'ref') { incomplete = true; continue }
    const slots: Array<[number, number]> = []
    let at = vs0
    const identity = (a: number, b: number) => { if (a <= b && !identitySpans(part.rows, a, b, budget, slots)) incomplete = true }
    for (const band of bands) {
      identity(at, Math.min(vs1, band.r0 - 1))
      const b0 = Math.max(at, band.r0), b1 = Math.min(vs1, band.r1)
      for (let slot = b0; slot <= b1; slot++) {
        if (--budget.work < 0) { incomplete = true; break }
        if (rowContains(part.rows, metrics.order.displayToSource(slot))) slots.push([slot, slot])
      }
      at = b1 + 1
    }
    identity(at, vs1)
    const ys = slots.map(([s0, s1]): Span => [metrics.rows.start(s0), metrics.rows.start(s1) + metrics.rows.size(s1)]).filter(([a, b]) => b > yLo && a < yHi)
    const xs = part.cols.map((c): Span => [metrics.cols.start(c.c0), metrics.cols.start(c.c1) + metrics.cols.size(c.c1)]).filter(([a, b]) => b > xLo && a < xHi)
    if (ys.length && xs.length) parts.push({ ys, xs })
  }
  // One global output budget: coarsen every axis of every part together.
  let unit = 1 / scale
  for (;;) {
    const grouped = parts.map(part => ({ ys: snapMerge(part.ys, yLo, yHi, unit), xs: snapMerge(part.xs, xLo, xHi, unit) }))
    const total = grouped.reduce((sum, part) => sum + part.ys.length * part.xs.length, 0)
    if (total <= COARSE_SELECTION_RECORDS) {
      if (unit > 1 / scale) incomplete = true
      const records: GridOverlayRecord[] = []
      for (const part of grouped) for (const [y0, y1] of part.ys) for (const [x0, x1] of part.xs) records.push({ source: { row: 0, col: 0 }, bounds: { x: x0, y: y0, width: x1 - x0, height: y1 - y0 } })
      return { records, incomplete }
    }
    unit *= 2
  }
}

/** Find highlights for the sheet pane: every hit gets the translucent band; the
 * active hit adds a stronger fill inset from the band (by the inset plus the
 * band's own edge), so its decoration never lands on a cell border or gridline
 * where it would read as an underline of the cell's text. */
const ACTIVE_INSET = 1.5
export function paintFindOverlay(ctx: CanvasRenderingContext2D, highlights: readonly Box[], active: readonly Box[], colors?: Pick<CanvasColors, 'find' | 'find-active'>): void {
  ctx.save()
  ctx.fillStyle = colors?.find ?? 'rgba(255, 214, 0, 0.45)'
  for (const r of highlights) ctx.fillRect(r.x, r.y, r.width, r.height)
  ctx.fillStyle = colors?.['find-active'] ?? 'rgba(255, 140, 0, 0.45)'
  for (const r of active) {
    const inset = Math.min(ACTIVE_INSET, r.width / 4, r.height / 4)
    ctx.fillRect(r.x + inset, r.y + inset, r.width - 2 * inset, r.height - 2 * inset)
  }
  ctx.restore()
}
