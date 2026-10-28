export interface SparklineInvalidationBox {
  x: number
  y: number
  width: number
  height: number
}

/** `null` is the conservative whole-sheet sentinel for an oversized merge;
 * `undefined` means this host has no visible area and needs no repaint. */
export function sparklineInvalidationBox(
  x0: number,
  x1: number,
  visibleRows: readonly { top: number; bottom: number }[],
  mergedRowSpan: number,
): SparklineInvalidationBox | null | undefined {
  if (mergedRowSpan > 10_000) return null
  let y0 = Number.POSITIVE_INFINITY
  let y1 = Number.NEGATIVE_INFINITY
  for (const row of visibleRows) {
    if (row.bottom > row.top) {
      y0 = Math.min(y0, row.top)
      y1 = Math.max(y1, row.bottom)
    }
  }
  return y1 > y0 && x1 > x0 ? { x: x0, y: y0, width: x1 - x0, height: y1 - y0 } : undefined
}
