/**
 * Source-space overview aggregates (spec §4.1 `queryOverview`, plan S1.2).
 *
 * A group is an exact compact source row set times column intervals. Rows are
 * walked as maximal runs of the set (never its bounding box), and each run is
 * asked for a rectangle aggregate. Adapters answer rectangles from their sparse
 * indexes: the object adapter from its populated-row index, the paged adapter
 * from per-page occupancy summaries, loading a page only when a rectangle cuts
 * it partially. Blank/default spans are analytic: no records means a uniform
 * default style and blank content without visiting a single coordinate.
 */
import type { CompactSourceRowSet, OverviewGroup, OverviewQuery, Page, SourceFailure, SourceStamp } from '../view/contracts'
import { MAX_OVERVIEW_GROUPS_DESKTOP } from '../view/contracts'
import { MAX_COLS, MAX_ROWS, type SourceRect } from '../view/coordinates'

/** Records = stored cells (style-only included); populated = value or formula. */
export interface RectAggregate {
  records: number
  populated: number
  /** Style of every record, 'mixed', or null when there are none. */
  style: number | 'mixed' | null
}
export type AggregateOutcome =
  | { kind: 'ok'; aggregate: RectAggregate }
  | { kind: 'unknown'; pageId: string }
  | { kind: 'error'; failure: SourceFailure }
export interface OverviewCellSource {
  aggregate(rect: SourceRect, signal: AbortSignal): Promise<AggregateOutcome>
}

export const emptyAggregate = (): RectAggregate => ({ records: 0, populated: 0, style: null })
export function addRecord(into: RectAggregate, styleIndex: number, populated: boolean): void {
  into.records++
  if (populated) into.populated++
  if (into.style === null) into.style = styleIndex
  else if (into.style !== styleIndex) into.style = 'mixed'
}
export function mergeAggregate(into: RectAggregate, from: RectAggregate): void {
  into.records += from.records
  into.populated += from.populated
  if (from.style === null) return
  if (into.style === null) into.style = from.style
  else if (into.style !== from.style) into.style = 'mixed'
}

const fail = (code: SourceFailure['code'], message: string, retryable = false): SourceFailure => ({ code, retryable, message })

/** Maximal ascending runs of an exact row set; throws RangeError on a malformed set. */
export function rowRuns(rows: CompactSourceRowSet): { r0: number; r1: number }[] | null {
  const valid = (r: number) => Number.isSafeInteger(r) && r >= 0 && r < MAX_ROWS
  const runs: { r0: number; r1: number }[] = []
  const push = (r: number) => {
    const last = runs[runs.length - 1]
    if (last && last.r1 + 1 === r) last.r1 = r
    else runs.push({ r0: r, r1: r })
  }
  switch (rows.kind) {
    case 'intervals': {
      let previous = -1
      for (const { r0, r1 } of rows.intervals) {
        if (!valid(r0) || !valid(r1) || r1 < r0 || r0 <= previous) throw new RangeError('Row intervals must be ascending and disjoint')
        previous = r1
        runs.push({ r0, r1 })
      }
      return runs
    }
    case 'list': {
      let previous = -1
      for (const r of rows.rows) {
        if (!valid(r) || r <= previous || r < rows.r0 || r > rows.r1) throw new RangeError('Row list must be strictly ascending inside [r0,r1]')
        previous = r
        push(r)
      }
      return runs
    }
    case 'bitset': {
      if (!valid(rows.r0) || !valid(rows.r1) || rows.r1 < rows.r0 || rows.words.length !== Math.ceil((rows.r1 - rows.r0 + 1) / 32)) {
        throw new RangeError('Malformed row bitset')
      }
      for (let w = 0; w < rows.words.length; w++) {
        let word = rows.words[w]
        if (word === 0) continue
        for (let bit = 0; bit < 32 && word !== 0; bit++) {
          if (word & 1) {
            const r = rows.r0 + w * 32 + bit
            if (r > rows.r1) throw new RangeError('Row bitset padding bits must be zero')
            push(r)
          }
          word >>>= 1
        }
      }
      return runs
    }
    case 'ref':
      return null
  }
}

/** Answers one bounded overview page; complete only when every group resolved. */
export async function answerOverview(request: OverviewQuery, source: OverviewCellSource, stamp: SourceStamp, signal: AbortSignal): Promise<Page<OverviewGroup>> {
  const failed = (failure: SourceFailure, items: OverviewGroup[] = []): Page<OverviewGroup> => ({ items, completeness: 'incomplete', stamp, error: failure })
  if (!Number.isSafeInteger(request.maxGroups) || request.maxGroups <= 0) return failed(fail('invalid-source-rect', 'maxGroups must be a positive integer'))
  if (request.maxGroups > MAX_OVERVIEW_GROUPS_DESKTOP) return failed(fail('resource-limit', `maxGroups is capped at ${MAX_OVERVIEW_GROUPS_DESKTOP}`))
  const total = request.groups.length
  let start = 0
  if (request.cursor !== undefined) {
    const parts = request.cursor.split(':')
    const at = Number(parts[1])
    if (parts.length !== 3 || parts[0] !== 'ov1' || Number(parts[2]) !== total || !Number.isSafeInteger(at) || at <= 0 || at >= total) {
      return failed(fail('invalid-cursor', 'Invalid overview cursor'))
    }
    start = at
  }
  const end = Math.min(total, start + request.maxGroups)
  const items: OverviewGroup[] = []
  let loading = false
  for (let g = start; g < end; g++) {
    if (signal.aborted) return failed(fail('cancelled', 'Request cancelled'))
    const group = request.groups[g]
    let runs: { r0: number; r1: number }[] | null
    try { runs = rowRuns(group.rows) } catch (cause) { return failed(fail('invalid-source-rect', (cause as Error).message)) }
    if (runs === null) return failed(fail('unsupported', 'Resident row-set refs are resolved by the worker, not the source'))
    let width = 0, previous = -1
    for (const { c0, c1 } of group.cols) {
      if (!Number.isSafeInteger(c0) || !Number.isSafeInteger(c1) || c0 < 0 || c1 >= MAX_COLS || c1 < c0 || c0 <= previous) {
        return failed(fail('invalid-source-rect', 'Column intervals must be ascending and disjoint'))
      }
      previous = c1
      width += c1 - c0 + 1
    }
    const sum = emptyAggregate()
    let cells = 0
    let unknown: string | undefined
    for (const run of runs) {
      cells += (run.r1 - run.r0 + 1) * width
      for (const { c0, c1 } of group.cols) {
        const outcome = await source.aggregate({ r0: run.r0, r1: run.r1, c0, c1 }, signal)
        if (outcome.kind === 'error') return failed(outcome.failure)
        if (outcome.kind === 'unknown') { unknown ??= outcome.pageId; continue }
        mergeAggregate(sum, outcome.aggregate)
      }
    }
    if (unknown !== undefined) {
      // An unloaded page is never summarized from the pages that did load.
      loading = true
      items.push({ groupId: group.groupId, content: 'unknown', pageId: unknown })
      continue
    }
    // Non-record cells carry the default style 0.
    const style: { kind: 'uniform'; styleIndex: number } | { kind: 'mixed' } =
      sum.records === cells
        ? (typeof sum.style === 'number' ? { kind: 'uniform', styleIndex: sum.style } : { kind: 'mixed' })
        : (sum.style === null || sum.style === 0 ? { kind: 'uniform', styleIndex: 0 } : { kind: 'mixed' })
    const content = sum.populated === 0 ? 'blank' : sum.populated === cells ? 'populated' : 'mixed'
    items.push({ groupId: group.groupId, populatedCount: sum.populated, countKind: 'exact', style, content })
  }
  if (signal.aborted) return failed(fail('cancelled', 'Request cancelled'))
  return {
    items, stamp, completeness: loading ? 'loading' : 'complete',
    ...(end < total ? { nextCursor: `ov1:${end}:${total}` } : {}),
  }
}
