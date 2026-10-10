import type { AxisMetrics, AxisRun, AxisSizeInterval, AxisWindow } from './contracts'
import { MAX_ROWS } from './coordinates'

export interface AxisMetricsOptions {
  logicalCount: number
  defaultSizeCss: number
  overrides?: readonly AxisSizeInterval[]
  collapsed?: readonly { start: number; end: number }[]
  maxWindowRuns?: number
}

export interface SparseAxisMetrics extends AxisMetrics {
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

function validateSize(size: number): void {
  if (!Number.isFinite(size) || size < 0) throw new RangeError('Axis size must be finite and nonnegative')
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
    validateInteger(interval.start, 0, logicalCount - 1)
    validateInteger(interval.end, interval.start, logicalCount - 1)
    validateSize(interval.sizeCss)
    return { ...interval }
  }).sort((a, b) => a.start - b.start || a.end - b.end)
  const result: AxisSizeInterval[] = []
  for (const interval of sorted) {
    const previous = result[result.length - 1]
    if (previous && interval.start <= previous.end && interval.sizeCss !== previous.sizeCss) {
      throw new RangeError('Contradictory overlapping axis overrides')
    }
    if (previous && interval.start <= previous.end + 1 && interval.sizeCss === previous.sizeCss) {
      previous.end = Math.max(previous.end, interval.end)
    } else result.push(interval)
  }
  return result
}

let nextAxisId = 0

/** Geometry is proportional to metadata transitions, never the blank axis length.
 * Windows cover [startCss,endCss); count belongs to this page, not later pages. */
export function createAxisMetrics(options: AxisMetricsOptions): SparseAxisMetrics {
  const { logicalCount, defaultSizeCss } = options
  const maxWindowRuns = options.maxWindowRuns ?? 256
  validateInteger(logicalCount, 0, MAX_ROWS)
  validateSize(defaultSizeCss)
  validateInteger(maxWindowRuns, 1, MAX_ROWS)
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
    if (!Number.isFinite(prefix[i + 1])) throw new RangeError('Axis total must be finite')
    if (sizes[i] > 0) {
      const lastCellStart = prefix[i] + (count - 1) * sizes[i]
      // A segment total can advance while its individual positive cells alias.
      // Check the largest origin's resolution once, without visiting each cell.
      if (sizes[i] < floatSpacingAt(lastCellStart) || prefix[i] + sizes[i] <= prefix[i]
        || prefix[i + 1] <= lastCellStart) {
        throw new RangeError('Axis geometry exceeds Float64 cell precision and cannot be represented')
      }
    }
    if (prefix[i + 1] > prefix[i]) positive.push(i)
  }
  const visibleSegments = Uint32Array.from(positive)
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

  function windowBetween(startCss: number, endCss: number, cursor?: string): AxisWindow {
    validateCss(startCss)
    validateCss(endCss)
    if (endCss < startCss) throw new RangeError('Axis window is reversed')
    lastWindowVisitedSegments = 0
    const lower = Math.max(0, startCss)
    const upper = Math.min(total, endCss)
    let position = lower < upper ? visibleSegmentAt(lower) : visibleSegments.length
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
    const runs: AxisRun[] = []
    let count = 0
    while (position < visibleSegments.length && runs.length < maxWindowRuns) {
      const segment = visibleSegments[position]
      if (prefix[segment] >= upper) break
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
    const nextCursor = position < visibleSegments.length && prefix[visibleSegments[position]] < upper
      ? `axis:${axisId}:${startCss}:${endCss}:${position}` : undefined
    return { runs, count, ...(nextCursor === undefined ? {} : { nextCursor }) }
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
    visibleBetween: (startCss: number, endCss: number) => windowBetween(startCss, endCss),
    continueVisibleBetween: (startCss: number, endCss: number, cursor: string) => windowBetween(startCss, endCss, cursor),
    get diagnostics() {
      return {
        segmentCount: starts.length,
        summaryBytes: starts.byteLength + ends.byteLength + sizes.byteLength + prefix.byteLength + visibleSegments.byteLength,
        lastWindowVisitedSegments,
        lastLocateComparisons,
      }
    },
  })
}
