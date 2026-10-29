/**
 * Excel's gridlines over styled regions (pivots spec §6.6; measured in docs/validation/xlsx-pivots.md,
 * "Gridlines over pivots and styled tables"):
 * - a pivot table draws no gridline between two cells of its location; the location's outer edges keep
 *   theirs (refs/pivot-light-16.png: no #D4D4D4 at any inner column or row edge, #D4D4D4 at x 0 and
 *   x 728 beside the white body rows);
 * - a cell of a styled table or pivot whose resolved style has a fill covers the gridlines on its four
 *   edges (refs/table-medium-2.png x 53: the band fill in striped rows, #D4D4D4 in white rows;
 *   refs/table-light-1.png x 10: a striped row's fill also covers the row lines above and below it).
 * A gridline is painted on the first pixel of the cell right of (below) it, so a filled right or bottom
 * cell already shows there once the line is hidden; a filled left or top cell beside an unfilled one
 * has its fill painted on the line instead.
 * Own fills outside styled regions keep their gridlines (a cross-format follow-up), so sheets without
 * tables or pivots pay nothing.
 */
import type { XlsxCellStyle, XlsxSheet } from './types'
import type { SheetRect } from './tableStyles/model'
import { regionStylesFor } from './tableStyleLoader'

export interface RegionGrid {
  /** Pivot locations and styled regions; a line touching none of them is drawn as usual. */
  readonly rects: readonly SheetRect[]
  /**
   * The gridline between cell A and its right (or lower) neighbour B: undefined when it is drawn as
   * usual, null when hidden, otherwise the fill (raw style colour) that covers it.
   */
  line(rowA: number, colA: number, rowB: number, colB: number): string | null | undefined
}

const A1 = /^\$?([A-Z]{1,3})\$?(\d{1,7})$/
function cell(ref: string): [number, number] | undefined {
  const m = A1.exec(ref)
  if (!m) return undefined
  let col = 0
  for (const ch of m[1]!) col = col * 26 + ch.charCodeAt(0) - 64
  return [Number(m[2]) - 1, col - 1]
}
/** A1 or A1:B2 → 0-based inclusive rect. */
export function rangeRect(ref: string): SheetRect | undefined {
  const [a, b = a] = ref.split(':')
  const p = cell(a ?? ''), q = cell(b ?? '')
  return p && q ? { minRow: Math.min(p[0], q[0]), minCol: Math.min(p[1], q[1]), maxRow: Math.max(p[0], q[0]), maxCol: Math.max(p[1], q[1]) } : undefined
}

const inside = (r: SheetRect, row: number, col: number) => row >= r.minRow && row <= r.maxRow && col >= r.minCol && col <= r.maxCol

/** The sheet's gridline rule, or undefined when it has no pivot and no styled region. `own` is a cell's own style. */
export function regionGrid(sheet: XlsxSheet, own: (row: number, col: number) => XlsxCellStyle | undefined): RegionGrid | undefined {
  const pivots: SheetRect[] = []
  for (const ref of [...sheet.pivotTables?.map(def => def.location.ref) ?? [], ...sheet.pivotRegions?.map(region => region.ref) ?? []]) {
    const rect = ref && rangeRect(ref)
    if (rect && !pivots.some(p => p.minRow === rect.minRow && p.minCol === rect.minCol && p.maxRow === rect.maxRow && p.maxCol === rect.maxCol)) pivots.push(rect)
  }
  const styles = regionStylesFor(sheet)
  const regions = styles?.rects ?? []
  if (!pivots.length && !regions.length) return undefined
  const memo = new Map<string, string>()
  const fill = (row: number, col: number): string => {
    if (!styles || !regions.some(r => inside(r, row, col))) return ''
    const key = `${row}:${col}`
    let value = memo.get(key)
    if (value === undefined) memo.set(key, value = styles.styleAt(row, col, own(row, col))?.fillColor ?? '')
    return value
  }
  return {
    rects: [...pivots, ...regions],
    line: (rowA, colA, rowB, colB) => {
      if (fill(rowB, colB)) return null
      const covering = fill(rowA, colA)
      if (covering) return covering
      return pivots.some(p => inside(p, rowA, colA) && inside(p, rowB, colB)) ? null : undefined
    },
  }
}

/** [start, end] segments minus hidden intervals (any order). */
export function subtractIntervals(segments: ReadonlyArray<readonly [number, number]>, hidden: Array<[number, number]>): Array<[number, number]> {
  if (!hidden.length) return segments as Array<[number, number]>
  hidden.sort((a, b) => a[0] - b[0])
  const out: Array<[number, number]> = []
  for (const [start, end] of segments) {
    let pos = start
    for (const [h0, h1] of hidden) {
      if (h1 <= pos || h0 >= end) continue
      if (h0 > pos) out.push([pos, h0])
      pos = Math.max(pos, h1)
    }
    if (pos < end) out.push([pos, end])
  }
  return out
}

/** The parts of `segments` inside [from, to]. */
export function intersectInterval(segments: ReadonlyArray<readonly [number, number]>, from: number, to: number): Array<[number, number]> {
  const out: Array<[number, number]> = []
  for (const [start, end] of segments) if (Math.min(end, to) > Math.max(start, from)) out.push([Math.max(start, from), Math.min(end, to)])
  return out
}
