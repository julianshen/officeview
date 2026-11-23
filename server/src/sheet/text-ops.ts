/**
 * Stage 4b (plan P2, P6–P10): the `sheet-text` pool job. Each op runs a local engine over the text view (Task 1), so
 * parity with local mode holds by construction: `fit` (`autoFitSize`) here; `find`, `copy` and `a11y` in Tasks 6–8.
 *
 * `textInputs` is the one place the ops get the local viewer's inputs (`InteractiveSheetDoc.tsx` find ~1814–1857 and
 * copy ~2611–2659): the row order, the row and column visibility, the store and the scope-aware display text. Fit
 * (ruling F4) reads only the view's sheet and geometry, with the base style, as the local fit does (it passes no
 * scope). No op reads a clock: the session clock is in the job and the view.
 */
import type { Box, WireSelection } from '../../../src/stream/sheet-wire'
import { autoFitSize, createFitMeasure } from '../../../src/xlsx/view/autoFit'
import type { Measure } from '../../../src/xlsx/view/axisSizes'
import type { RowOrder, RowVisibility, SourceStore } from '../../../src/xlsx/view/contracts'
import type { EffectiveFormatScope } from '../../../src/xlsx/view/formatWorkspace'
import { createIdentityOrder } from '../../../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../../../src/xlsx/view/visibility'
import type { FormatAllocation } from '../../../src/xlsx/formula/format'
import { cellDisplayText } from '../../../src/xlsx/render'
import type { XlsxCell } from '../../../src/xlsx/types'
import type { WorkbookSource } from '../../../src/worker/viewerSession'
import type { TextView } from './text-view'
import type { WireSizes } from './view-state'

export type SheetTextJob =
  | { op: 'find'; query: string; matchCase: boolean; sheetIndex: number; sizes: WireSizes; limit: number }
  | { op: 'copy'; sheetIndex: number; selection: WireSelection; sizes: WireSizes; maxBytes: number }
  | { op: 'a11y'; sheetIndex: number; rows: Array<[number, number]>; cols: [number, number]; sizes: WireSizes }
  /** `regions` absent (never `[]`): `autoFitSize` then treats the pane as the one visible region, as the local fit does. */
  | { op: 'fit'; sheetIndex: number; axis: 'row' | 'col'; index: number; pane: Box; regions?: Box[]; sizes: WireSizes }
export type SheetTextReply =
  | { op: 'find'; hits: Array<[number, number]>; more: boolean } // [row, col] in local order; more: hits past `limit`
  | { op: 'copy'; text: string; paragraphs: number } | { op: 'copy'; refused: 'too-large'; bytes: number }
  | { op: 'a11y'; cells: Array<[number, number, string]> }
  | { op: 'fit'; size: number }

/** The measurement scratch one fit may hold (`createFitMeasure`'s admitted heap + native bytes); past it the fit is
 * refused (`AutoFitRefusedError`, answered `policy-refused`). */
export const FIT_MEASURE_BYTES = 16 * 1024 * 1024

/** The local viewer's text-op inputs over one text view (`source` is an object source over the text document, so its
 * sheet `view.sheetIndex` is `view.sheet`). */
export interface TextInputs {
  sheetId: string
  /** 4b has no sort: the identity order (display row = source row). */
  order: RowOrder
  /** Display-space row visibility: the authored hidden rows (4b has no filter). */
  visibility: RowVisibility
  /** Column visibility: the collapsed (zero-width) columns of the view's geometry are hidden. */
  columns: RowVisibility
  store(): SourceStore
  /** The cell's style as paint and copy use it: a table or pivot region style resolved over its own. */
  styleOf(cell: Pick<XlsxCell, 'row' | 'col'>): XlsxCell['style']
  /** Display text as the local viewer forms it: the region-resolved style through the scope's CF-effective style. */
  textOf(cell: Pick<XlsxCell, 'row' | 'col' | 'value'>, allocation?: FormatAllocation, scope?: EffectiveFormatScope): string
}

function cellOf(view: TextView, row: number, col: number): XlsxCell | undefined {
  if (view.geometry.kind !== 'ready') return undefined
  return view.geometry.cellsOf(row)?.find((c) => c.col === col)
}

export function textInputs(view: TextView, source: WorkbookSource): TextInputs {
  const entry = source.sheets[view.sheetIndex]
  if (!entry) throw new RangeError(`text view sheet ${view.sheetIndex} is not in the source`)
  const sheetId = entry.sheetId
  const order = createIdentityOrder()
  const authored = source.geometry(sheetId).baseHidden.intervals.map(({ r0, r1 }) => ({ r0: order.sourceToDisplay(r0), r1: order.sourceToDisplay(r1) }))
  const hiddenCols: Array<{ r0: number; r1: number }> = []
  let colCount = 0
  if (view.geometry.kind === 'ready') {
    const cols = view.geometry.metrics.cols
    colCount = cols.logicalCount
    for (const o of cols.overrides) if (o.sizeCss === 0) hiddenCols.push({ r0: o.start, r1: o.end })
  }
  const styleOf = (cell: Pick<XlsxCell, 'row' | 'col'>) => {
    const own = cellOf(view, cell.row, cell.col)?.style
    return view.sheet.regionStyles?.styleAt(cell.row, cell.col, own, 'screen') ?? own
  }
  return {
    sheetId, order,
    visibility: visibilityFromHidden(createIntervalSet(authored)),
    columns: visibilityFromHidden(createIntervalSet(hiddenCols), colCount),
    store: () => source.store(sheetId),
    styleOf,
    textOf(cell, allocation, scope) {
      const base = styleOf(cell)
      return cellDisplayText(view.sheet, { value: cell.value, style: scope?.effectiveStyle(cell.row, cell.col, base) ?? base }, allocation)
    },
  }
}

/** `box` clamped to `extent` (P9): its origin inside the sheet, its size no past the sheet's far edge. */
export function clampBox(box: Box, extent: { width: number; height: number }): Box {
  const x = Math.min(box.x, extent.width), y = Math.min(box.y, extent.height)
  return { x, y, width: Math.min(box.width, extent.width - x), height: Math.min(box.height, extent.height - y) }
}

/**
 * The fit size (P9): local `autoFitSize` over the view's (blanked) sheet and its sized geometry, with the pane and
 * regions clamped to the sheet's extent and `knownRows` the view's first rows, as on the local resident path.
 * `measureText` is the server backend's text measure; with it injected `createFitMeasure` builds no layout context.
 */
export function fitAnswer(view: TextView, job: Extract<SheetTextJob, { op: 'fit' }>, measureText: Measure): number {
  const geometry = view.geometry
  if (geometry.kind !== 'ready') throw new Error('sheet-fit: sheet geometry unavailable')
  const extent = { width: geometry.metrics.cols.totalCss(), height: geometry.metrics.rows.totalCss() }
  const fit = createFitMeasure({ admit: ({ heap, native }) => heap + native <= FIT_MEASURE_BYTES, measureText, sheet: view.sheet })
  return autoFitSize(job.axis === 'col' ? 'column' : 'row', job.index, {
    // Fix round 1: no regions is `undefined`, never `[]` (`autoFitSize` reads `regions ?? [pane]`; `[]` would mean no
    // visible rows at all).
    geometry, sheet: view.sheet, pane: clampBox(job.pane, extent), ...(job.regions?.length ? { regions: job.regions.map((r) => clampBox(r, extent)) } : {}),
    knownRows: (limit) => view.sheet.rows.slice(0, limit).map((r) => r.index),
    measure: fit.measure, allocation: fit.allocation,
  })
}

/** Runs one `sheet-text` op over the text view of `job.sheetIndex` at `job.sizes` (`viewOf`: the host's cached views). */
export async function answerSheetText(job: SheetTextJob, deps: { viewOf: (sheetIndex: number, sizes: WireSizes) => Promise<TextView>; measureText: Measure }): Promise<SheetTextReply> {
  switch (job.op) {
    case 'fit': return { op: 'fit', size: fitAnswer(await deps.viewOf(job.sheetIndex, job.sizes), job, deps.measureText) }
    // Tasks 6–8 replace these.
    case 'find':
    case 'copy':
    case 'a11y':
      throw new Error(`sheet-text: op '${job.op}' is not implemented`)
    default: {
      const never: never = job
      throw new Error(`sheet-text: unknown op ${JSON.stringify((never as { op?: unknown }).op)}`)
    }
  }
}
