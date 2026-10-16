/** Authored frozen panes (`<sheetView><pane state="frozen|frozenSplit">`), shared by
 * the resident parser and the paged importer. A plain split (state="split",
 * twip offsets) is not a freeze and is ignored. */
export interface XlsxFrozenPane {
  /** Frozen leading rows / columns (ySplit / xSplit). */
  rows: number
  cols: number
  /** 0-based top-left cell of the scrolling (bottom-right) pane, when authored. */
  topRow?: number
  leftCol?: number
}

const MAX_FROZEN_ROWS = 1_048_576, MAX_FROZEN_COLS = 16_384
const count = (raw: string | undefined, max: number): number => {
  if (raw === undefined || raw.trim() === '') return 0
  const n = Number(raw)
  return Number.isInteger(n) && n > 0 && n < max ? n : 0
}

export function frozenPaneFrom(a: Record<string, string | undefined>): XlsxFrozenPane | undefined {
  if (a.state !== 'frozen' && a.state !== 'frozenSplit') return undefined
  const rows = count(a.ySplit, MAX_FROZEN_ROWS), cols = count(a.xSplit, MAX_FROZEN_COLS)
  if (!rows && !cols) return undefined
  const pane: XlsxFrozenPane = { rows, cols }
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(a.topLeftCell ?? '')
  if (m) {
    let col = 0
    for (const ch of m[1].toUpperCase()) col = col * 26 + ch.charCodeAt(0) - 64
    const row = Number(m[2]) - 1
    // Both authored coordinates count, whichever axis is frozen (a scrolling
    // pane may start below/right of the strip on the other axis too).
    if (row > 0 && row >= rows && row < MAX_FROZEN_ROWS) pane.topRow = row
    if (col - 1 > 0 && col - 1 >= cols && col - 1 < MAX_FROZEN_COLS) pane.leftCol = col - 1
  }
  return pane
}
