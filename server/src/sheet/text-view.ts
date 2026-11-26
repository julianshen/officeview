/**
 * Stage 4b (plan P2–P5, S4 §6.2 steps 3–5): the text view, the one sheet every server text answer (find, copy, a11y,
 * fit) reads. No text op reads `SheetSnapshot.sheet` directly: under unknown coverage that sheet keeps the authored
 * values, so tiles paint as authored (P2).
 *
 * Under `'visible-only'` the view is the snapshot's (base-blanked) sheet with these cells also blank, every one by
 * `blankCell` semantics (value removed, style kept):
 * - P5, current-view coverage: `baseCoverage` over the projected sheet with its drawings placed by the session sizes;
 * - P3, merge tails: every merge cell but the anchor;
 * - zero-height rows (PR #112 review): every cell of a row the geometry collapses at the session sizes (authored `ht="0"`, hidden or
 *   not), but an anchor whose merge reaches a drawn row;
 * - P4, undrawn: a non-empty value whose display text under its CF-effective painted style (`TextView.styleAt`, the table
 *   or pivot region style over its own, through the CF scope) is `''` (`;;;` and the like); a cell
 *   whose CF bar or icon has `showValue="0"` (from the snapshot's CF batch); and, fail closed, every cell of every CF
 *   sqref when the CF text scope is unavailable, and every cell of a `showValue="0"` rule's sqref when the batch is.
 * Unknown coverage (base or current view) blanks every cell, and `'full'` is the identity. Tiles never read the view.
 */
import type { XlsxCell, XlsxCellStyle, XlsxDocument, XlsxSheet } from '../../../src/xlsx/types'
import { cfClockOptions } from '../../../src/components/xlsx/sheetAssets'
import { cfDocumentEvaluationOf } from '../../../src/xlsx/cf/formula'
import { CF_LIMITS } from '../../../src/xlsx/cf/limits'
import { prepareCfTextScope } from '../../../src/xlsx/cf/text'
import { cellDisplayText } from '../../../src/xlsx/render'
import { regionStylesFor } from '../../../src/xlsx/tableStyleLoader'
import type { RegionStyleLayer } from '../../../src/xlsx/tableStyles/regionLayer'
import { hasSessionSizes, type SessionSizes } from '../../../src/xlsx/view/axisSizes'
import { coverCellsToRects, type SourceRect } from '../../../src/xlsx/view/coordinates'
import type { EffectiveFormatScope, FormatBudget } from '../../../src/xlsx/view/formatWorkspace'
import { cellSourceRow } from '../../../src/xlsx/view/sheetCells'
import { buildSheetGeometry, type SheetGeometryResult } from '../../../src/xlsx/view/viewport'
import { clockEpochMs, type SheetClock } from './clock'
import { baseCoverage } from './coverage'
import { blankCell, holdsContent, type SheetSnapshot } from './snapshot'

export interface TextView {
  sheetIndex: number
  /** The sheet text answers read: `projected` under `'full'`, else the snapshot's sheet with P3–P5 blanking applied. */
  sheet: XlsxSheet
  /** Built over `sheet` with the session sizes. */
  geometry: SheetGeometryResult
  /** Unknown coverage (base or current view): every cell is blank. */
  blankAll: boolean
  /** Cells holding content in the projected sheet that hold none here (0 under `'full'`). */
  withheld: number
  /** P4's CF text scope, the `prepareScope`/`textOf` scope of every op; undefined when the sheet has no CF, the scope
   * is unavailable (its sqrefs are then blank) or every cell is blank. */
  scope: EffectiveFormatScope | undefined
  /** The style a cell paints with before CF: its table or pivot region style resolved over `own`, as the tile painter
   * resolves it (`regionStylesFor`, `'screen'`; renderViewport.ts). Blanking, find, copy and a11y all read this one. */
  styleAt(row: number, col: number, own: XlsxCellStyle | undefined): XlsxCellStyle | undefined
  /** Disposes the scope (idempotent). The snapshot is the host's: a view never releases it. */
  release(): void
}

/** A copy of `sheet` with every content cell `blank(row, col, cell)` names blanked (`sheet` itself when none is). Rows
 * without such a cell are shared with the original. */
function blankWhere(sheet: XlsxSheet, blank: (row: number, col: number, cell: XlsxCell) => boolean): XlsxSheet {
  let changed = false
  const rows = sheet.rows.map((row) => {
    let cells: XlsxCell[] | undefined
    row.cells.forEach((cell, i) => {
      if (!holdsContent(cell) || !blank(cellSourceRow(row, cell), cell.col, cell)) return
      cells ??= [...row.cells]
      cells[i] = blankCell(cell)
    })
    if (!cells) return row
    changed = true
    return { ...row, cells }
  })
  return changed ? { ...sheet, rows } : sheet
}

function contentCount(sheet: XlsxSheet): number {
  let n = 0
  for (const row of sheet.rows) for (const cell of row.cells) if (holdsContent(cell)) n++
  return n
}

type Ranged = { firstRow: number; lastRow: number; firstCol: number; lastCol: number }
const inRanges = (ranges: readonly Ranged[], row: number, col: number) => ranges.some((g) => row >= g.firstRow && row <= g.lastRow && col >= g.firstCol && col <= g.lastCol)

/**
 * The exact cover of the sheet's content cells inside any CF rule's sqref, each rect split to at most
 * `CF_LIMITS.preparedCells` cells (F3). `prepareCfTextScope` refuses a single larger rect, and the local viewer never
 * sends one: find and copy prepare per batch over the exact cover of the batch's cells. Splitting the request never
 * changes a rank or aggregate rule's result: the rects only choose which cells get an outcome, while the rules keep
 * their whole sqrefs and one statistics session spans the whole call.
 */
function scopeRects(sheet: XlsxSheet, ranges: readonly Ranged[]): SourceRect[] {
  const cells: Array<{ row: number; col: number }> = []
  for (const row of sheet.rows) for (const cell of row.cells) {
    const r = cellSourceRow(row, cell)
    if (holdsContent(cell) && inRanges(ranges, r, cell.col)) cells.push({ row: r, col: cell.col })
  }
  const max = CF_LIMITS.preparedCells
  const out: SourceRect[] = []
  for (const rect of coverCellsToRects(cells)) {
    const width = Math.min(rect.c1 - rect.c0 + 1, max)
    const rowsPer = Math.max(1, Math.floor(max / width))
    for (let c0 = rect.c0; c0 <= rect.c1; c0 += width) {
      const c1 = Math.min(rect.c1, c0 + width - 1)
      for (let r0 = rect.r0; r0 <= rect.r1; r0 += rowsPer) out.push({ r0, r1: Math.min(rect.r1, r0 + rowsPer - 1), c0, c1 })
    }
  }
  return out
}

/** P4's CF text scope over `rects` of `sheet`, under the session clock (the tile's `cfClock`). `undefined` with rects
 * means unavailable: `prepareCfTextScope` answers a failure that is not a resource limit with `undefined` too. */
async function prepareScope(sheet: XlsxSheet, rects: readonly SourceRect[], o: { workbook: XlsxDocument; clock: SheetClock; budget: FormatBudget }): Promise<EffectiveFormatScope | undefined> {
  if (!rects.length) return undefined
  const clocked = cfClockOptions(cfDocumentEvaluationOf(o.workbook), { epochNowMs: clockEpochMs(o.clock), timeZone: o.clock.timeZone })
  try {
    return await prepareCfTextScope(sheet, {
      rects, signal: new AbortController().signal, locale: sheet.semantics?.locale ?? 'en-US', budget: o.budget,
      documentSheets: o.workbook.sheets, valuePolicy: o.workbook.valuePolicy ?? 'engine',
      evaluation: clocked.evaluation, now: clocked.now, dateContext: clocked.dateContext,
    })
  } catch { return undefined }
}

/** A non-empty value whose display text under `style` is `''`; a formatting failure counts as undrawn (fail closed). */
function paintsNothing(sheet: XlsxSheet, cell: XlsxCell, style: XlsxCellStyle | undefined): boolean {
  if (cell.value === null || cell.value === undefined || cell.value === '') return false
  try { return cellDisplayText(sheet, { value: cell.value, style }) === '' } catch { return true }
}

/** The painter's pre-CF style over `region` (the region layer of the sheet the view is built from). */
function regionStyleAt(region: RegionStyleLayer | undefined): TextView['styleAt'] {
  return (row, col, own) => region?.styleAt(row, col, own, 'screen') ?? own
}

function view(snap: SheetSnapshot, sheet: XlsxSheet, sizes: SessionSizes, projected: XlsxSheet, blankAll: boolean, scope: EffectiveFormatScope | undefined, styleAt: TextView['styleAt']): TextView {
  let released = false
  return {
    sheetIndex: snap.sheetIndex, sheet, geometry: buildSheetGeometry(sheet, { sizes }), blankAll,
    withheld: sheet === projected ? 0 : contentCount(projected) - contentCount(sheet), scope, styleAt,
    release() { if (!released) { released = true; scope?.dispose() } },
  }
}

/**
 * The text view of `snap` (built from the cached snapshot: P4 reads its CF batch) under the session `sizes`.
 * `projected` is the projected (unblanked) sheet the snapshot was built from; `workbook` is the servable workbook (every
 * slot its blanked sheet), which CF's cross-sheet rules read; `budget` admits the scope until `release`.
 */
export async function sheetTextView(snap: SheetSnapshot, projected: XlsxSheet, sizes: SessionSizes, o: { workbook: XlsxDocument; clock: SheetClock; budget: FormatBudget }): Promise<TextView> {
  const ranges = (sheet: XlsxSheet) => (sheet.conditionalFormatting?.rules ?? []).flatMap((rule) => rule.ranges)
  if (snap.contentPolicy === 'full') return view(snap, projected, sizes, projected, false, await prepareScope(projected, scopeRects(projected, ranges(projected)), o), regionStyleAt(regionStylesFor(projected)))

  const blankEverything = () => view(snap, blankWhere(projected, () => true), sizes, projected, true, undefined, regionStyleAt(regionStylesFor(projected)))
  if (snap.coverage === 'off' || snap.coverage.kind !== 'known' || snap.geometry.kind !== 'ready') return blankEverything()
  const merges = snap.geometry.merges

  // P5: current-view coverage, over the unblanked sheet (spill reads its cells) with the drawings the sizes place.
  let currentCovered: ((row: number, col: number) => boolean) | undefined
  // The row sizes the viewer draws at: a zero-height row draws no cell (zero-height rows, below).
  let rows = snap.geometry.metrics.rows
  if (hasSessionSizes(sizes)) {
    const sized = buildSheetGeometry(projected, { sizes })
    if (sized.kind === 'ready') rows = sized.metrics.rows
    const placed = sized.kind === 'ready' ? sized.sizedDrawings ?? projected.drawings : projected.drawings
    const current = baseCoverage({ ...projected, ...(placed ? { drawings: placed as XlsxSheet['drawings'] } : {}) }, sized)
    if (current.kind !== 'known') return blankEverything()
    currentCovered = (row, col) => current.covered(row, col)
  }

  // P4: the CF text scope, over the snapshot's sheet (the values the tile's CF batch was prepared over).
  const sheet = snap.sheet
  // Resolved on the snapshot's sheet before it is copied, so the view's sheet carries the same layer.
  const styleAt = regionStyleAt(regionStylesFor(sheet))
  const rules = sheet.conditionalFormatting?.rules ?? []
  const cfRanges = ranges(sheet)
  const rects = rules.length ? scopeRects(sheet, cfRanges) : []
  const scope = await prepareScope(sheet, rects, o)
  const scopeUnavailable = rects.length > 0 && !scope
  const batch = snap.assets.conditionalFormatting
  const hiddenValueRanges = batch ? [] : rules.filter((rule) => (rule.definition.kind === 'dataBar' || rule.definition.kind === 'iconSet') && rule.definition.showValue === false).flatMap((rule) => rule.ranges)
  // Format diagnostics land on a copy, never on the cached sheet.
  const formatSheet = { ...sheet }

  try {
    const out = blankWhere(sheet, (row, col, cell) => {
      if (currentCovered?.(row, col)) return true
      const merge = merges.first({ r0: row, r1: row, c0: col, c1: col })
      if (merge && (merge.minRow !== row || merge.minCol !== col)) return true
      // Zero-height rows: a row the geometry collapses (an authored `ht="0"` without `hidden`, which `baseHidden`
      // does not hold; a session size never reopens it) draws none of its cells, unless it anchors a merge that reaches
      // a drawn row.
      if (rows.size(row) === 0 && !(merge && rows.start(merge.maxRow) + rows.size(merge.maxRow) > rows.start(row))) return true
      if (scopeUnavailable && inRanges(cfRanges, row, col)) return true
      if (batch && (batch.barAt(row, col)?.showValue === false || batch.iconAt(row, col)?.showValue === false)) return true
      if (hiddenValueRanges.length && inRanges(hiddenValueRanges, row, col)) return true
      // Undrawn under the scope's effective style over the painted base, or under the style the tile's CF batch paints
      // with (either suffices). The batch's layer resolves the region style over the own style itself.
      const base = styleAt(row, col, cell.style)
      if (paintsNothing(formatSheet, cell, scope ? scope.effectiveStyle(row, col, base) : base)) return true
      const painted = batch?.layer.styleAt(row, col, cell.style, 'screen')
      return painted !== undefined && painted !== base && paintsNothing(formatSheet, cell, painted)
    })
    return view(snap, out, sizes, projected, false, scope, styleAt)
  } catch (e) {
    scope?.dispose()
    throw e
  }
}

/** `base` with each sheet that has a view replaced by its view's sheet; a sheet without one stays `base`'s (the caller
 * passes the servable workbook, so such a sheet is at least base-blanked). */
export function textDocument(base: XlsxDocument, views: ReadonlyMap<number, TextView>): XlsxDocument {
  return { ...base, sheets: base.sheets.map((s, i) => views.get(i)?.sheet ?? s) }
}
