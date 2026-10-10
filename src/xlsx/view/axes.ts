/**
 * Sparse axis geometry (§5.3). Storage is proportional to metadata transitions,
 * never to the blank axis length.
 *
 * Measured cost of pathological layouts (accepted for now): every transition
 * keeps a JS segment object plus a frozen `overrides` copy, ~160 B/transition.
 * 1,048,576 alternating 1 px overrides built in ~430 ms and ~160 MB.
 * TODO(S4.3): move segment storage into typed arrays (the starts/ends/sizes/prefix
 * summaries below already are) and derive `overrides` lazily if worker budgets
 * require it.
 */
import type { AxisMetrics, AxisRun, AxisSizeInterval, AxisWindow } from './contracts'
import { MAX_ROWS } from './coordinates'

/** Stable, machine-readable reasons a caller-supplied axis description is rejected
 * at construction. S2.3/S3 map these to a diagnosed fallback instead of a crash.
 * - CONTRADICTORY_OVERRIDES: overlapping overrides with different sizes
 * - PRECISION_EXCEEDED: a positive cell cannot be told from its neighbour in Float64
 * - TOTAL_NOT_FINITE: the axis total overflows Float64
 * - INVALID_METADATA: counts, sizes or intervals outside the axis domain
 * Extends RangeError so existing `instanceof RangeError` handling keeps working. */
export type AxisMetricsErrorCode =
  | 'CONTRADICTORY_OVERRIDES' | 'PRECISION_EXCEEDED' | 'TOTAL_NOT_FINITE' | 'INVALID_METADATA'
export class AxisMetricsError extends RangeError {
  readonly code: AxisMetricsErrorCode
  constructor(code: AxisMetricsErrorCode, message: string) {
    super(message)
    this.name = 'AxisMetricsError'
    this.code = code
  }
}

export interface AxisMetricsOptions {
  logicalCount: number
  defaultSizeCss: number
  overrides?: readonly AxisSizeInterval[]
  collapsed?: readonly { start: number; end: number }[]
  maxWindowRuns?: number
}

export interface SparseAxisMetrics extends AxisMetrics {
  /** First bounded page of [startCss,endCss): at most `maxWindowRuns` runs. `count` is
   * this page's covered indexes; `nextCursor` resumes via continueVisibleBetween. */
  visiblePageBetween(startCss: number, endCss: number): AxisWindow
  continueVisibleBetween(startCss: number, endCss: number, cursor: string): AxisWindow
  readonly diagnostics: {
    segmentCount: number
    summaryBytes: number
    lastWindowVisitedSegments: number
    lastLocateComparisons: number
  }
}

function validateInteger(value: number, min: number, max: number): void {
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || value < min || value > max) {
    throw new RangeError('Axis index or count is outside the domain')
  }
}
function metadataInteger(value: number, min: number, max: number): void {
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || value < min || value > max) {
    throw new AxisMetricsError('INVALID_METADATA', 'Axis index or count is outside the domain')
  }
}

function validateSize(size: number): void {
  if (!Number.isFinite(size) || size < 0) throw new AxisMetricsError('INVALID_METADATA', 'Axis size must be finite and nonnegative')
}

function floatSpacingAt(value: number): number {
  if (value === 0) return Number.MIN_VALUE
  let exponent = Math.floor(Math.log2(value))
  // log2 can round up immediately below a power of two.
  if (2 ** exponent > value) exponent--
  return 2 ** Math.max(-1074, exponent - 52)
}

function normalize(intervals: readonly AxisSizeInterval[], logicalCount: number): AxisSizeInterval[] {
  const sorted = intervals.map(interval => {
    metadataInteger(interval.start, 0, logicalCount - 1)
    metadataInteger(interval.end, interval.start, logicalCount - 1)
    validateSize(interval.sizeCss)
    return { ...interval }
  }).sort((a, b) => a.start - b.start || a.end - b.end)
  const result: AxisSizeInterval[] = []
  for (const interval of sorted) {
    const previous = result[result.length - 1]
    if (previous && interval.start <= previous.end && interval.sizeCss !== previous.sizeCss) {
      throw new AxisMetricsError('CONTRADICTORY_OVERRIDES', 'Contradictory overlapping axis overrides')
    }
    if (previous && interval.start <= previous.end + 1 && interval.sizeCss === previous.sizeCss) {
      previous.end = Math.max(previous.end, interval.end)
    } else result.push(interval)
  }
  return result
}

let nextAxisId = 0

/** Geometry is proportional to metadata transitions, never the blank axis length.
 * Windows cover [startCss,endCss). */
export function createAxisMetrics(options: AxisMetricsOptions): SparseAxisMetrics {
  const { logicalCount, defaultSizeCss } = options
  const maxWindowRuns = options.maxWindowRuns ?? 256
  metadataInteger(logicalCount, 0, MAX_ROWS)
  validateSize(defaultSizeCss)
  metadataInteger(maxWindowRuns, 1, MAX_ROWS)
  const overrides = normalize(options.overrides ?? [], logicalCount)
  const collapsed = normalize((options.collapsed ?? []).map(interval => ({ ...interval, sizeCss: 0 })), logicalCount)
  const boundaries = [...new Set([0, logicalCount,
    ...overrides.flatMap(interval => [interval.start, interval.end + 1]),
    ...collapsed.flatMap(interval => [interval.start, interval.end + 1]),
  ])].sort((a, b) => a - b)
  const segments: AxisSizeInterval[] = []
  let overrideIndex = 0
  let collapsedIndex = 0
  for (let i = 0; i < boundaries.length - 1; i++) {
    const start = boundaries[i]
    const end = boundaries[i + 1] - 1
    while (overrideIndex < overrides.length && overrides[overrideIndex].end < start) overrideIndex++
    while (collapsedIndex < collapsed.length && collapsed[collapsedIndex].end < start) collapsedIndex++
    const authored = overrides[overrideIndex]
    const hidden = collapsed[collapsedIndex]
    const sizeCss = hidden && hidden.start <= start ? 0
      : authored && authored.start <= start ? authored.sizeCss : defaultSizeCss
    const previous = segments[segments.length - 1]
    if (previous && previous.sizeCss === sizeCss) previous.end = end
    else segments.push({ start, end, sizeCss })
  }
  const starts = Uint32Array.from(segments, segment => segment.start)
  const ends = Uint32Array.from(segments, segment => segment.end)
  const sizes = Float64Array.from(segments, segment => segment.sizeCss)
  const prefix = new Float64Array(segments.length + 1)
  const positive: number[] = []
  for (let i = 0; i < segments.length; i++) {
    const count = ends[i] - starts[i] + 1
    prefix[i + 1] = prefix[i] + count * sizes[i]
    if (!Number.isFinite(prefix[i + 1])) throw new AxisMetricsError('TOTAL_NOT_FINITE', 'Axis total must be finite')
    if (sizes[i] > 0) {
      const lastCellStart = prefix[i] + (count - 1) * sizes[i]
      // A segment total can advance while its individual positive cells alias.
      // Check the largest origin's resolution once, without visiting each cell.
      if (sizes[i] < floatSpacingAt(lastCellStart) || prefix[i] + sizes[i] <= prefix[i]
        || prefix[i + 1] <= lastCellStart) {
        throw new AxisMetricsError('PRECISION_EXCEEDED', 'Axis geometry exceeds Float64 cell precision and cannot be represented')
      }
    }
    if (prefix[i + 1] > prefix[i]) positive.push(i)
  }
  const visibleSegments = Uint32Array.from(positive)
  // cumulative[k] = indexes in the first k positive segments (count of a window in O(log n)).
  const cumulative = new Uint32Array(visibleSegments.length + 1)
  for (let i = 0; i < visibleSegments.length; i++) {
    const segment = visibleSegments[i]
    cumulative[i + 1] = cumulative[i] + ends[segment] - starts[segment] + 1
  }
  const effectiveOverrides = Object.freeze(segments.filter(segment => segment.sizeCss !== defaultSizeCss)
    .map(segment => Object.freeze(segment)))
  const total = prefix[prefix.length - 1]
  const axisId = ++nextAxisId
  let lastWindowVisitedSegments = 0
  let lastLocateComparisons = 0

  function segmentOf(index: number): number {
    validateInteger(index, 0, logicalCount - 1)
    let low = 0
    let high = starts.length
    while (low < high) {
      const mid = Math.floor((low + high) / 2)
      if (starts[mid] <= index) low = mid + 1
      else high = mid
    }
    return low - 1
  }

  function startInSegment(index: number, segment: number): number {
    return prefix[segment] + (index - starts[segment]) * sizes[segment]
  }

  function validateCss(value: number): void {
    if (!Number.isFinite(value)) throw new RangeError('Axis offset must be finite')
  }

  function visibleSegmentAt(css: number, trackComparisons = false): number {
    let low = 0
    let high = visibleSegments.length
    while (low < high) {
      if (trackComparisons) lastLocateComparisons++
      const mid = Math.floor((low + high) / 2)
      if (prefix[visibleSegments[mid] + 1] <= css) low = mid + 1
      else high = mid
    }
    return low
  }

  function indexAt(css: number, segment: number, trackComparisons = false): number {
    const candidate = Math.max(starts[segment], Math.min(ends[segment],
      starts[segment] + Math.floor((css - prefix[segment]) / sizes[segment])))
    // Division can round to a neighbouring index at an exact authored boundary.
    if (trackComparisons) lastLocateComparisons++
    if (startInSegment(candidate, segment) <= css) {
      if (candidate === ends[segment]) return candidate
      if (trackComparisons) lastLocateComparisons++
      if (startInSegment(candidate + 1, segment) > css) return candidate
    }
    let low = starts[segment]
    let high = ends[segment] + 1
    while (low < high) {
      const mid = Math.floor((low + high) / 2)
      if (trackComparisons) lastLocateComparisons++
      if (startInSegment(mid, segment) <= css) low = mid + 1
      else high = mid
    }
    return low - 1
  }

  function locate(css: number): number | null {
    validateCss(css)
    lastLocateComparisons = 0
    if (css < 0 || css >= total) return null
    return indexAt(css, visibleSegments[visibleSegmentAt(css, true)], true)
  }

  /** One bounded page starting at positive-segment `position`; `next` is -1 when exhausted. */
  function pageFrom(lower: number, upper: number, position: number): { runs: AxisRun[]; count: number; next: number } {
    lastWindowVisitedSegments = 0
    const runs: AxisRun[] = []
    let count = 0
    while (position < visibleSegments.length && runs.length < maxWindowRuns) {
      const segment = visibleSegments[position]
      if (prefix[segment] >= upper) { position = visibleSegments.length; break }
      lastWindowVisitedSegments++
      const start = indexAt(Math.max(lower, prefix[segment]), segment)
      let end = ends[segment]
      if (upper < prefix[segment + 1]) {
        end = indexAt(upper, segment)
        if (startInSegment(end, segment) === upper) end--
      }
      if (end >= start) {
        runs.push({ start, end, startCss: startInSegment(start, segment), sizeCss: sizes[segment] })
        count += end - start + 1
      }
      position++
    }
    const more = position < visibleSegments.length && prefix[visibleSegments[position]] < upper
    return { runs, count, next: more ? position : -1 }
  }

  function bounds(startCss: number, endCss: number): { lower: number; upper: number } {
    validateCss(startCss)
    validateCss(endCss)
    if (endCss < startCss) throw new RangeError('Axis window is reversed')
    return { lower: Math.max(0, startCss), upper: Math.min(total, endCss) }
  }

  function firstPosition(lower: number, upper: number): number {
    return lower < upper ? visibleSegmentAt(lower) : visibleSegments.length
  }

  /** Indexes covered by [lower,upper), from the per-segment cumulative counts: O(log n). */
  function coveredCount(lower: number, upper: number): number {
    if (!(lower < upper)) return 0
    const startPosition = visibleSegmentAt(lower)
    if (startPosition >= visibleSegments.length) return 0
    const startSegment = visibleSegments[startPosition]
    const before = cumulative[startPosition] + indexAt(Math.max(lower, prefix[startSegment]), startSegment) - starts[startSegment]
    const endPosition = visibleSegmentAt(upper)
    let through = cumulative[endPosition]
    if (endPosition < visibleSegments.length) {
      const segment = visibleSegments[endPosition]
      if (prefix[segment] < upper) {
        let end = indexAt(upper, segment)
        if (startInSegment(end, segment) === upper) end--
        through += end - starts[segment] + 1
      }
    }
    return through - before
  }

  function boundedPage(startCss: number, endCss: number, cursor?: string): AxisWindow {
    const { lower, upper } = bounds(startCss, endCss)
    let position = firstPosition(lower, upper)
    if (cursor !== undefined) {
      const parts = cursor.split(':')
      const next = Number(parts[4])
      if (parts.length !== 5 || parts[0] !== 'axis' || parts[1] !== String(axisId)
        || parts[2] !== String(startCss) || parts[3] !== String(endCss)
        || !Number.isSafeInteger(next) || parts[4] !== String(next)
        || next <= position || next >= visibleSegments.length || prefix[visibleSegments[next]] >= upper) {
        throw new RangeError('Axis cursor does not belong to this window')
      }
      position = next
    }
    const page = pageFrom(lower, upper, position)
    return { runs: page.runs, count: page.count,
      ...(page.next < 0 ? {} : { nextCursor: `axis:${axisId}:${startCss}:${endCss}:${page.next}` }) }
  }

  /** Complete, lazy window: `runs` is re-iterable, pages internally (at most
   * maxWindowRuns runs resident at a time) until [startCss,endCss) is exhausted, and
   * `count` is the total number of covered indexes. `nextCursor` is never set. */
  function windowBetween(startCss: number, endCss: number): AxisWindow {
    const { lower, upper } = bounds(startCss, endCss)
    return {
      runs: {
        *[Symbol.iterator](): Iterator<AxisRun> {
          let position = firstPosition(lower, upper)
          while (position >= 0) {
            const page = pageFrom(lower, upper, position)
            yield* page.runs
            position = page.next
          }
        },
      },
      count: coveredCount(lower, upper),
    }
  }

  return Object.freeze({
    logicalCount,
    defaultSizeCss,
    overrides: effectiveOverrides,
    totalCss: () => total,
    start: (index: number) => {
      validateInteger(index, 0, logicalCount)
      return index === logicalCount ? total : startInSegment(index, segmentOf(index))
    },
    size: (index: number) => sizes[segmentOf(index)],
    locate,
    visibleBetween: windowBetween,
    visiblePageBetween: (startCss: number, endCss: number) => boundedPage(startCss, endCss),
    continueVisibleBetween: (startCss: number, endCss: number, cursor: string) => boundedPage(startCss, endCss, cursor),
    get diagnostics() {
      return {
        segmentCount: starts.length,
        summaryBytes: starts.byteLength + ends.byteLength + sizes.byteLength + prefix.byteLength + visibleSegments.byteLength + cumulative.byteLength,
        lastWindowVisitedSegments,
        lastLocateComparisons,
      }
    },
  })
}
