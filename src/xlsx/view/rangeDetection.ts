/** Sort/filter operation range detection and structural admission (spec §6.1).
 * Pure module over caller-owned source state. Resolution precedence is strict:
 * a table containing the active cell wins with its metadata header/totals,
 * else the current region grown by sparse adjacency around the active cell,
 * else the authoritative used extent. Selection bounds never define the range:
 * only the active cell seeds detection, so disjoint Shift selections are not
 * coerced into a bounding rectangle. Guards run through the shared S3 guard;
 * nothing here publishes, mutates, or invents values. Unknown coverage waits,
 * refusals carry typed diagnostics, and cancellation maps to cancelled. */
import type {
  FilterRange, GridSelection, OperationRefusal, ReadOptions, SourceStore,
} from './contracts'
import {
  MAX_COLS, MAX_ROWS, isValidCoordinate, isValidSourceRectInput, validateCoordinate,
  type SourceRect,
} from './coordinates'
import {
  ColumnChunkIndex, RowChunkIndex, currentRegion, fromStoreScan,
  type StructuralIndex,
} from './structure'
import { guardSort, isChildOnlyRange } from './operationGuard'

export interface RangeDetectionRequest {
  store: SourceStore
  structure: StructuralIndex
  selection: GridSelection
  options: ReadOptions
  /** Dialog header override for the non-table heuristic; tables keep metadata. */
  override?: { headerRows: 0 | 1; totalsRows: 0 | 1 } | null
  pageLimit?: number
  /** Authoritative used extent override; defaults to the store extent. */
  usedRange?: SourceRect
}
export type RangeDetection =
  | { kind: 'complete'; range: FilterRange }
  | { kind: 'incomplete'; missing: readonly SourceRect[] }
  | { kind: 'refused'; refusal: OperationRefusal }
  | { kind: 'cancelled' }

const SEED_WINDOW_HALF = 2048
const WIDEN_FACTOR = 4
const FULL_DOMAIN: SourceRect = { r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }

function refused(refusal: OperationRefusal): RangeDetection {
  return { kind: 'refused', refusal }
}
/** Guard and child-only outcomes share the refusal shape, but an 'incomplete'
 * reason means unknown coverage (wait and retry), not a terminal refusal. */
function translateRefusal(refusal: OperationRefusal): RangeDetection {
  if (refusal.reason === 'incomplete') return { kind: 'incomplete', missing: refusal.missing }
  return refused(refusal)
}
function freezeRange(rect: SourceRect, headerRows: 0 | 1, totalsRows: 0 | 1): FilterRange {
  return Object.freeze({ r0: rect.r0, r1: rect.r1, c0: rect.c0, c1: rect.c1, headerRows, totalsRows })
}
/** Refusal for a non-complete structural index: a source failure outranks a
 * bare wait (it carries the code and retryability), else the missing ranges. */
function structuralWait(structure: StructuralIndex): RangeDetection {
  const failure = structure.failure
  if (failure !== undefined) {
    return refused({ reason: 'source-error', code: failure.code, retryable: failure.retryable, message: failure.message })
  }
  return { kind: 'incomplete', missing: structure.missing }
}
type HeaderProbe =
  | { kind: 'counts'; headerRows: 0 | 1 }
  | { kind: 'incomplete'; missing: readonly SourceRect[] }
  | { kind: 'refused'; refusal: OperationRefusal }
/** Non-table header heuristic (spec §6.1): the first row holds text and no
 * non-text while the next body row holds at least one non-text value. Blank
 * cells are ignored; errors count as non-text; unknown or cache-less formula
 * cells block inference instead of guessing, so the probe waits on them. */
async function probeHeaderRows(store: SourceStore, options: ReadOptions, rect: SourceRect): Promise<HeaderProbe> {
  const band: SourceRect = { r0: rect.r0, r1: Math.min(rect.r0 + 1, MAX_ROWS - 1), c0: rect.c0, c1: rect.c1 }
  let row0Text = false, row0NonText = false, row1NonText = false
  let cursor: string | undefined
  const seen = new Set<string>()
  for (;;) {
    options.signal.throwIfAborted()
    const page = await store.readRows(band, cursor, options)
    if (page.error !== undefined) {
      return { kind: 'refused', refusal: { reason: 'source-error', code: page.error.code, retryable: page.error.retryable, message: page.error.message } }
    }
    if (page.completeness !== 'complete') return { kind: 'incomplete', missing: Object.freeze([{ ...band }]) }
    for (const item of page.items) {
      if (item.row < band.r0 || item.row > band.r1) throw new RangeError('Read row outside the requested band')
      for (const cell of item.cells) {
        validateCoordinate(cell.row, cell.col)
        if (cell.row < band.r0 || cell.row > band.r1 || cell.col < band.c0 || cell.col > band.c1) {
          throw new RangeError('Read cell outside the requested band')
        }
        if (cell.unavailable) {
          return { kind: 'incomplete', missing: Object.freeze([{ r0: cell.row, r1: cell.row, c0: cell.col, c1: cell.col }]) }
        }
        const isText = !cell.valueIsError && typeof cell.value === 'string'
        const isNonText = cell.valueIsError || typeof cell.value === 'number' || typeof cell.value === 'boolean'
        if (item.row === band.r0) {
          row0Text = row0Text || isText
          row0NonText = row0NonText || isNonText
        } else if (item.row === band.r1) {
          row1NonText = row1NonText || isNonText
        }
      }
    }
    if (page.nextCursor === undefined) break
    if (seen.has(page.nextCursor)) throw new RangeError('Read cursor made no progress')
    seen.add(page.nextCursor)
    cursor = page.nextCursor
  }
  return { kind: 'counts', headerRows: row0Text && !row0NonText && row1NonText ? 1 : 0 }
}
/** Shared tail: child-only refusal, then header protection, then the frozen range. */
async function finishRect(
  request: RangeDetectionRequest,
  rect: SourceRect,
  header: { kind: 'table'; headerRows: 0 | 1; totalsRows: 0 | 1 } | { kind: 'policy' },
): Promise<RangeDetection> {
  const childOnly = isChildOnlyRange(rect, request.structure)
  if (childOnly !== null) return translateRefusal(childOnly)
  let headerRows: 0 | 1, totalsRows: 0 | 1
  if (header.kind === 'table') {
    headerRows = header.headerRows
    totalsRows = header.totalsRows
  } else if (request.override !== undefined && request.override !== null) {
    headerRows = request.override.headerRows
    totalsRows = request.override.totalsRows
  } else if (rect.r0 === rect.r1) {
    headerRows = 0
    totalsRows = 0
  } else {
    const probe = await probeHeaderRows(request.store, request.options, rect)
    if (probe.kind === 'incomplete') return { kind: 'incomplete', missing: probe.missing }
    if (probe.kind === 'refused') return refused(probe.refusal)
    headerRows = probe.headerRows
    totalsRows = 0
  }
  return { kind: 'complete', range: freezeRange(rect, headerRows, totalsRows) }
}
async function detectRange(request: RangeDetectionRequest): Promise<RangeDetection> {
  const { store, structure, selection, options } = request
  const pageLimit = request.pageLimit ?? 4096
  if (request.pageLimit !== undefined && (!Number.isSafeInteger(request.pageLimit) || request.pageLimit <= 0)) {
    return refused({ reason: 'invalid-range', message: 'Invalid chunk page limit' })
  }
  if (request.override !== undefined && request.override !== null
    && ((request.override.headerRows !== 0 && request.override.headerRows !== 1)
      || (request.override.totalsRows !== 0 && request.override.totalsRows !== 1))) {
    return refused({ reason: 'invalid-rule', message: 'Invalid header override counts' })
  }
  const active = selection.active
  if (active === null) {
    return refused({ reason: 'ambiguous-selection', message: 'No active cell to seed range detection' })
  }
  if (!isValidCoordinate(active.row, active.col)) {
    return refused({ reason: 'invalid-range', message: 'Active cell is outside the worksheet domain' })
  }
  options.signal.throwIfAborted()
  if (structure.completeness !== 'complete') return structuralWait(structure)
  const tableHit = structure.firstWhere(
    { r0: active.row, r1: active.row, c0: active.col, c1: active.col },
    hit => hit.kind === 'table',
  )
  if (tableHit !== undefined && tableHit.kind === 'table') {
    const table = tableHit.record
    if (!isValidSourceRectInput(table.range)
      || (table.headerRows !== 0 && table.headerRows !== 1)
      || (table.totalsRows !== 0 && table.totalsRows !== 1)) {
      return refused({ reason: 'invalid-range', message: 'Declared table range is outside the worksheet domain' })
    }
    return finishRect(request, table.range, { kind: 'table', headerRows: table.headerRows, totalsRows: table.totalsRows })
  }
  const used = request.usedRange ?? store.extent.used
  if (used === undefined) {
    return { kind: 'incomplete', missing: Object.freeze([{ ...FULL_DOMAIN }]) }
  }
  if (!isValidSourceRectInput(used)) {
    return refused({ reason: 'invalid-range', message: 'Used range is outside the worksheet domain' })
  }
  if (active.row < used.r0 || active.row > used.r1) {
    return finishRect(request, used, { kind: 'policy' })
  }
  const source = fromStoreScan(store, options, pageLimit)
  const full = { r0: Math.max(0, used.r0 - 1), r1: Math.min(MAX_ROWS - 1, used.r1 + 1) }
  let half = SEED_WINDOW_HALF
  for (;;) {
    options.signal.throwIfAborted()
    const band = { r0: Math.max(full.r0, active.row - half), r1: Math.min(full.r1, active.row + half) }
    const rows = await RowChunkIndex.build(source, band, pageLimit)
    if (!rows.complete) return { kind: 'incomplete', missing: rows.missing }
    const cols = await ColumnChunkIndex.build(source, band, pageLimit)
    if (!cols.complete) return { kind: 'incomplete', missing: cols.missing }
    const region = currentRegion(active, rows.index, cols.index)
    if (region === null) return finishRect(request, used, { kind: 'policy' })
    if (typeof region === 'object' && 'kind' in region) {
      if (band.r0 === full.r0 && band.r1 === full.r1) return { kind: 'incomplete', missing: region.missing }
      half = Math.min(half * WIDEN_FACTOR, MAX_ROWS)
      continue
    }
    return finishRect(request, region, { kind: 'policy' })
  }
}
/** Resolve the operation range: table metadata, current region, or used extent,
 * with header protection and child-only refusal. Never throws for domain
 * inputs; operational failures propagate and cancellation maps to cancelled. */
export async function detectOperationRange(request: RangeDetectionRequest): Promise<RangeDetection> {
  try {
    return await detectRange(request)
  } catch (error) {
    if (request.options.signal?.aborted) return { kind: 'cancelled' }
    throw error
  }
}
/** Detect the range, then admit the sort through the shared guard. A refusal or
 * wait propagates unchanged; only a fully guarded range completes. */
export async function admitSortRange(request: RangeDetectionRequest): Promise<RangeDetection> {
  try {
    const detected = await detectRange(request)
    if (detected.kind !== 'complete') return detected
    const occupancy = fromStoreScan(request.store, request.options, request.pageLimit)
    const guard = await guardSort(detected.range, request.structure, occupancy, { pageLimit: request.pageLimit })
    if (!guard.ok) return translateRefusal(guard.refusal)
    return detected
  } catch (error) {
    if (request.options.signal?.aborted) return { kind: 'cancelled' }
    throw error
  }
}
