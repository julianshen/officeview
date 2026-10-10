/** Operation guard and drawing prefix impact (spec sections 6.3, 8.1; AC-SEM-04, AC-SEM-05).
 * Pure and atomic: guards only read the structural index, populated source and
 * caller-built projections, and return a discriminated outcome. A refusal changes
 * no view generation, mask, selection or cache owner. Incomplete coverage WAITS
 * (reason 'incomplete') and never authorizes. */
import type { FilterRange, OperationRefusal, RowVisibility, SourceRowInterval, SpillExtent } from './contracts'
import { MAX_COLS, MAX_ROWS, validateSourceRect, type SourceRect } from './coordinates'
import type { ProjectedSheetMetrics } from './projection'
import { findOutsideColumnData, type CellAnchoredDrawing, type PopulatedCellSource, type StructuralIndex } from './structure'

export type GuardOutcome = { ok: true } | { ok: false; refusal: OperationRefusal }
export interface SortGuardOptions { pageLimit?: number }
/** Candidate filter plus the before/after geometry the caller built from it.
 * currentVisibility is the committed mask (null when fully visible); the changed
 * band is derived from current vs candidate so the structural query covers at
 * least every changed-height row band across all columns (spec section 6.3). */
export interface FilterGuardRequest {
  range: FilterRange
  candidateVisibility: RowVisibility
  currentVisibility: RowVisibility | null
  projectionBefore: ProjectedSheetMetrics
  projectionAfter: ProjectedSheetMetrics
  structure: StructuralIndex
}
function protectedCounts(range: FilterRange): void {
  if ((range.headerRows !== 0 && range.headerRows !== 1) || (range.totalsRows !== 0 && range.totalsRows !== 1)) {
    throw new RangeError('Invalid protected row counts')
  }
}
/** Affected body rows, or null when protection covers the whole range. Throws
 * RangeError/CoordinateError on invalid input; callers map those to invalid-range. */
function bodyBand(range: FilterRange): SourceRowInterval | null {
  validateSourceRect(range)
  protectedCounts(range)
  const r0 = range.r0 + range.headerRows
  const r1 = range.r1 - range.totalsRows
  return r0 > r1 ? null : { r0, r1 }
}
function invalidRange(error: unknown): GuardOutcome {
  if (error instanceof RangeError) return { ok: false, refusal: { reason: 'invalid-range', message: error.message } }
  throw error
}
function mergeBands(bands: readonly SourceRowInterval[]): SourceRowInterval[] {
  const sorted = [...bands].sort((a, b) => a.r0 - b.r0 || a.r1 - b.r1)
  const out: SourceRowInterval[] = []
  for (const band of sorted) {
    const previous = out[out.length - 1]
    if (previous && previous.r1 + 1 >= band.r0) previous.r1 = Math.max(previous.r1, band.r1)
    else out.push({ ...band })
  }
  return out
}
/** First intersecting saved/evaluated extent, else the anchor cell. The record was
 * found by a band first-hit, so one of them always intersects. */
function spillRefusalRange(spill: SpillExtent, band: SourceRect): SourceRect {
  for (const extent of [spill.saved, spill.evaluated]) {
    if (extent !== undefined && extent.r0 <= band.r1 && band.r0 <= extent.r1
      && extent.c0 <= band.c1 && band.c0 <= extent.c1) return extent
  }
  return { r0: spill.anchor.row, r1: spill.anchor.row, c0: spill.anchor.col, c1: spill.anchor.col }
}
function anchorInside(at: { row: number; col: number }, range: SourceRect): boolean {
  return at.row >= range.r0 && at.row <= range.r1 && at.col >= range.c0 && at.col <= range.c1
}
/** Vertical geometry of a cell-anchored drawing: display Y of the anchor row and
 * display Y just past the occupied row span. */
function drawingSpan(drawing: CellAnchoredDrawing, projection: ProjectedSheetMetrics): { startY: number; endY: number } {
  const startY = projection.rows.start(projection.order.sourceToDisplay(drawing.from.row))
  const past = drawing.bounds.r1 + 1
  const endY = past >= projection.rows.logicalCount
    ? projection.rows.totalCss()
    : projection.rows.start(projection.order.sourceToDisplay(past))
  return { startY, endY }
}
/** Sort permutes whole rows, so the guard queries the entire affected body band
 * across all 16,384 columns: merges, spill extents (child-only included), the
 * occupied bands of cell-anchored drawings, and any nonblank outside-column data. */
export async function guardSort(
  range: FilterRange,
  structure: StructuralIndex,
  occupancy: PopulatedCellSource,
  options?: SortGuardOptions,
): Promise<GuardOutcome> {
  let band: SourceRowInterval | null
  try {
    band = bodyBand(range)
  } catch (error) {
    return invalidRange(error)
  }
  if (band === null) return { ok: true }
  if (structure.completeness !== 'complete') {
    return { ok: false, refusal: { reason: 'incomplete', missing: structure.missing } }
  }
  const fullBand: SourceRect = { r0: band.r0, r1: band.r1, c0: 0, c1: MAX_COLS - 1 }
  const merge = structure.firstWhere(fullBand, hit => hit.kind === 'merge')
  if (merge !== undefined && merge.kind === 'merge') {
    return { ok: false, refusal: { reason: 'merge', range: merge.record.range } }
  }
  const spill = structure.firstWhere(fullBand, hit => hit.kind === 'spill')
  if (spill !== undefined && spill.kind === 'spill') {
    return {
      ok: false,
      refusal: {
        reason: spill.record.kind === 'fixed' ? 'fixed-array' : 'dynamic-spill',
        range: spillRefusalRange(spill.record, fullBand),
      },
    }
  }
  const drawing = structure.firstWhere(fullBand, hit => hit.kind === 'drawing')
  if (drawing !== undefined && drawing.kind === 'drawing') {
    return { ok: false, refusal: { reason: 'drawing-impact', drawingId: drawing.record.id } }
  }
  const built = await findOutsideColumnData(occupancy, band, range.c0, range.c1, options?.pageLimit)
  if (built.kind === 'unknown') return { ok: false, refusal: { reason: 'incomplete', missing: built.missing } }
  if (built.kind === 'occupied') {
    return {
      ok: false,
      refusal: {
        reason: 'outside-column-data',
        range: { r0: built.at.row, r1: built.at.row, c0: built.at.col, c1: built.at.col },
      },
    }
  }
  return { ok: true }
}
/** Filter hides rows in place. The guard refuses merge/spill intersections across
 * the body band plus every changed-height row band (all columns), then refuses
 * when changed row heights alter the start or occupied span of any cell-anchored
 * drawing — including changed rows above the anchor and within it, even when the
 * filter and anchor ranges do not intersect. Absolute anchors are static. */
export async function guardFilter(request: FilterGuardRequest): Promise<GuardOutcome> {
  const { range, candidateVisibility, currentVisibility, projectionBefore, projectionAfter, structure } = request
  let band: SourceRowInterval | null
  try {
    band = bodyBand(range)
  } catch (error) {
    return invalidRange(error)
  }
  if (band === null) return { ok: true }
  if (structure.completeness !== 'complete') {
    return { ok: false, refusal: { reason: 'incomplete', missing: structure.missing } }
  }
  const count = projectionBefore.rows.logicalCount
  const changedSources: number[] = []
  for (let slot = 0; slot < count; slot++) {
    const before = currentVisibility?.visible(slot) ?? true
    if (before !== candidateVisibility.visible(slot)) changedSources.push(projectionBefore.order.displayToSource(slot))
  }
  // Source rows are not display-ordered under a band permutation, so sort before merging.
  changedSources.sort((a, b) => a - b)
  const changed: SourceRowInterval[] = []
  for (const source of changedSources) {
    const previous = changed[changed.length - 1]
    if (previous && previous.r1 + 1 >= source) previous.r1 = Math.max(previous.r1, source)
    else changed.push({ r0: source, r1: source })
  }
  for (const covered of mergeBands([band, ...changed])) {
    const coveredBand: SourceRect = { r0: covered.r0, r1: covered.r1, c0: 0, c1: MAX_COLS - 1 }
    const merge = structure.firstWhere(coveredBand, hit => hit.kind === 'merge')
    if (merge !== undefined && merge.kind === 'merge') {
      return { ok: false, refusal: { reason: 'merge', range: merge.record.range } }
    }
    const spill = structure.firstWhere(coveredBand, hit => hit.kind === 'spill')
    if (spill !== undefined && spill.kind === 'spill') {
      return {
        ok: false,
        refusal: {
          reason: spill.record.kind === 'fixed' ? 'fixed-array' : 'dynamic-spill',
          range: spillRefusalRange(spill.record, coveredBand),
        },
      }
    }
  }
  // Changed rows all lie inside the projection, so a drawing anchored beyond it can
  // only be moved when some changed row precedes it; otherwise it is untouched and
  // its (unmeasurable) geometry must not turn an unrelated filter into a refusal.
  const firstChanged = changedSources.length > 0 ? changedSources[0] : Number.POSITIVE_INFINITY
  for (const drawing of structure.drawingsIn({ r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 })) {
    if (drawing.from.row >= count && firstChanged >= drawing.from.row) continue
    let before: { startY: number; endY: number }
    let after: { startY: number; endY: number }
    try {
      before = drawingSpan(drawing, projectionBefore)
      after = drawingSpan(drawing, projectionAfter)
    } catch (error) {
      // Anchors outside the projection geometry (or any other range-shaped
      // failure while reading it) refuse as invalid ranges, never throw.
      if (error instanceof RangeError) return { ok: false, refusal: { reason: 'invalid-range', message: error.message } }
      throw error
    }
    if (before.startY !== after.startY || before.endY !== after.endY) {
      return { ok: false, refusal: { reason: 'drawing-impact', drawingId: drawing.id } }
    }
  }
  return { ok: true }
}
/** Whether a proposed operation range is an unsupported child-only slice: it
 * intersects a merge/spill extent without containing that record's anchor.
 * Range derivation (S7) consults this before resolving; selection itself stays allowed. */
export function isChildOnlyRange(range: SourceRect, structure: StructuralIndex): OperationRefusal | null {
  validateSourceRect(range)
  if (structure.completeness !== 'complete') return { reason: 'incomplete', missing: structure.missing }
  const merge = structure.firstWhere(range, hit => hit.kind === 'merge' && !anchorInside(hit.record.anchor, range))
  if (merge !== undefined && merge.kind === 'merge') return { reason: 'merge', range: merge.record.range }
  const spill = structure.firstWhere(range, hit => hit.kind === 'spill' && !anchorInside(hit.record.anchor, range))
  if (spill !== undefined && spill.kind === 'spill') {
    return {
      reason: spill.record.kind === 'fixed' ? 'fixed-array' : 'dynamic-spill',
      range: spillRefusalRange(spill.record, range),
    }
  }
  return null
}
/** Read-only selection/highlighting is always safe, including spill/merge children
 * (spec section 8.1). Only operation ranges are refused. */
export function selectionAllowed(_selection: unknown): boolean {
  return true
}
function fnv1a(text: string): string {
  let hash = 0x811c9dc5
  for (let i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i)
    hash = Math.imul(hash, 0x01000193)
  }
  return (hash >>> 0).toString(16)
}
/** Stable deduplication key so one refusal produces one status notice. */
export function refusalKey(refusal: OperationRefusal): string {
  switch (refusal.reason) {
    case 'merge':
    case 'fixed-array':
    case 'dynamic-spill':
    case 'outside-column-data':
      return `${refusal.reason}:${refusal.range.r0}:${refusal.range.r1}:${refusal.range.c0}:${refusal.range.c1}`
    case 'drawing-impact':
      return `drawing-impact:${refusal.drawingId}`
    case 'incomplete':
      return `incomplete:${refusal.missing.map(rect => `${rect.r0}:${rect.r1}:${rect.c0}:${rect.c1}`).join(',')}`
    case 'formula-value-unavailable':
      return `formula-value-unavailable:${refusal.firstAddress.row}:${refusal.firstAddress.col}`
    default:
      return `${refusal.reason}#${fnv1a(refusal.message)}`
  }
}
