import type { AxisMetrics, AxisRun, AxisSizeInterval, AxisWindow, RowOrder, RowVisibility, SheetMetrics } from './contracts'
import { MAX_COLS, MAX_ROWS, type SourceCell, type SourceRow, type DisplaySlot } from './coordinates'
import { AxisMetricsError, type AxisMetricsErrorCode } from './axes'
import { createIdentityOrder } from './order'
import type { VisibleRank } from './visibility'
export type { VisibleRank } from './visibility'

/** rows are immutable SOURCE-space dimensions; visibility is DISPLAY-space. */
export interface ProjectionOptions { rows: AxisMetrics; cols: AxisMetrics; order?: RowOrder; visibility?: RowVisibility }
export interface ProjectedRows extends AxisMetrics {
  /** First bounded page of [startCss,endCss) (at most 256 visited runs); `count` is this
   * page's rows and `nextCursor` resumes via continueVisibleBetween. `visibleBetween`
   * itself is the complete lazy window. */
  visiblePageBetween(startCss: number, endCss: number): AxisWindow
  continueVisibleBetween(startCss: number, endCss: number, cursor: string): AxisWindow
}
export interface ProjectedSheetMetrics extends SheetMetrics {
  rows: ProjectedRows
  visibleRankToDisplay(ordinal: VisibleRank): DisplaySlot | null
  displayToVisibleRank(slot: DisplaySlot): VisibleRank | null
  visibleRankToSource(ordinal: VisibleRank): SourceRow | null
  sourceToVisibleRank(row: SourceRow): VisibleRank | null
  /** Projection-owned buffers only; caller-owned order, visibility and source axes are excluded. */
  readonly diagnostics: {
    sectionCount: number
    densePrefixBytes: number
    typedArrayBytes: number
    /** Dense row candidates / analytic uniform run starts; excludes index/rank probes. */
    lastWindowVisitedSlots: number
  }
}

interface SourceSegment { start: SourceRow; end: SourceRow; sizeCss: number }
interface Section {
  start: DisplaySlot
  end: DisplaySlot
  sizeCss: number
  cssBase: number
  rankBase: VisibleRank
  maskBase: number
  count: number
  total: number
  sizes?: Float64Array
  prefix?: Float64Array
  positive?: Uint32Array
}
const WINDOW_WORK_LIMIT = 256
let nextProjectionId = 0

function integer(value: number, maximum: number, code: AxisMetricsErrorCode = 'INVALID_QUERY'): void {
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || value < 0 || value > maximum) {
    throw new AxisMetricsError(code, 'Projection index or count is outside the axis')
  }
}
function finite(value: number, code: AxisMetricsErrorCode = 'INVALID_QUERY'): void {
  if (!Number.isFinite(value)) throw new AxisMetricsError(code, 'Projection offset must be finite')
}
function size(value: number): void {
  finite(value, 'INVALID_METADATA')
  if (value < 0) throw new AxisMetricsError('INVALID_METADATA', 'Projection dimensions must be nonnegative')
}
function floatSpacingAt(value: number): number {
  if (value === 0) return Number.MIN_VALUE
  let exponent = Math.floor(Math.log2(value))
  // log2 may round up immediately below a power of two.
  if (2 ** exponent > value) exponent--
  return 2 ** Math.max(-1074, exponent - 52)
}
function validateCellPrecision(origin: number, end: number, height: number): void {
  if (height < floatSpacingAt(origin) || origin + height <= origin || end <= origin) {
    throw new AxisMetricsError('PRECISION_EXCEEDED', 'Projected geometry exceeds Float64 cell precision and cannot be represented')
  }
}
function sourceSegments(axis: AxisMetrics, maximum: number): SourceSegment[] {
  integer(axis.logicalCount, maximum, 'INVALID_METADATA')
  size(axis.defaultSizeCss)
  const segments: SourceSegment[] = []
  function append(start: number, end: number, sizeCss: number): void {
    const previous = segments[segments.length - 1]
    if (previous && previous.end === start && previous.sizeCss === sizeCss) previous.end = end
    else segments.push({ start, end, sizeCss })
  }
  let next = 0
  for (const interval of axis.overrides) {
    integer(interval.start, axis.logicalCount - 1, 'INVALID_METADATA')
    integer(interval.end, axis.logicalCount - 1, 'INVALID_METADATA')
    size(interval.sizeCss)
    if (interval.start < next || interval.end < interval.start) throw new AxisMetricsError('INVALID_METADATA', 'Invalid effective axis overrides')
    if (next < interval.start) append(next, interval.start, axis.defaultSizeCss)
    append(interval.start, interval.end + 1, interval.sizeCss)
    next = interval.end + 1
  }
  if (next < axis.logicalCount) append(next, axis.logicalCount, axis.defaultSizeCss)
  return segments
}
/** Return the first item whose exclusive end lies beyond the query. */
function endingAfter<T>(items: readonly T[], query: number, end: (item: T) => number): number {
  let low = 0, high = items.length
  while (low < high) {
    const mid = Math.floor((low + high) / 2)
    if (end(items[mid]) <= query) low = mid + 1
    else high = mid
  }
  return low
}

/** Construction rejects invalid caller metadata with AxisMetricsError/INVALID_METADATA.
 * Caller-owned order buffers, visibility words and axes must not be mutated after
 * publication; S4 owns their generation lifetimes and replacement (§10.1). */
export function createSheetMetrics(options: ProjectionOptions): ProjectedSheetMetrics {
  const source = sourceSegments(options.rows, MAX_ROWS)
  sourceSegments(options.cols, MAX_COLS)
  const logicalCount = options.rows.logicalCount
  const defaultSizeCss = options.rows.defaultSizeCss
  const order = options.order ?? createIdentityOrder()
  const visibility = options.visibility
  const maskRank = (end: DisplaySlot) => visibility ? visibility.rank(end) : end
  const maskVisible = (slot: DisplaySlot) => !visibility || visibility.visible(slot)
  // Compact masks expose their extent; the base contract also permits custom rank/select implementations.
  const maskCount = (visibility as (RowVisibility & { logicalCount?: number }) | undefined)?.logicalCount
  if (maskCount !== undefined && maskCount !== logicalCount) throw new AxisMetricsError('INVALID_METADATA', 'Visibility extent differs from row axis')
  const maskTotal = maskRank(logicalCount)
  integer(maskTotal, logicalCount, 'INVALID_METADATA')
  if (maskRank(0) !== 0) throw new AxisMetricsError('INVALID_METADATA', 'Invalid visibility rank origin')
  let previousEnd = -1
  if (order.kind !== 'identity' && order.kind !== 'bands') throw new AxisMetricsError('INVALID_METADATA', 'Invalid row order kind')
  if (order.kind === 'identity' && order.bands.length) throw new AxisMetricsError('INVALID_METADATA', 'Identity order cannot contain bands')
  for (const band of order.bands) {
    integer(band.r0, logicalCount - 1, 'INVALID_METADATA')
    integer(band.r1, logicalCount - 1, 'INVALID_METADATA')
    const length = band.r1 - band.r0 + 1
    if (length < 1 || band.r0 <= previousEnd || !(band.sourceAtOffset instanceof Uint32Array)
      || !(band.displayOffsetOfSource instanceof Uint32Array)
      || band.sourceAtOffset.length !== length || band.displayOffsetOfSource.length !== length) {
      throw new AxisMetricsError('INVALID_METADATA', 'Invalid projection permutation band')
    }
    for (let offset = 0; offset < length; offset++) {
      const row = band.sourceAtOffset[offset]
      if (row < band.r0 || row > band.r1 || band.displayOffsetOfSource[row - band.r0] !== offset) {
        throw new AxisMetricsError('INVALID_METADATA', 'Projection permutation is not bijective')
      }
    }
    previousEnd = band.r1
  }

  const sections: Section[] = []
  let densePrefixBytes = 0
  let typedArrayBytes = 0
  function appendUniform(start: number, end: number, sizeCss: number): void {
    if (start === end) return
    const previous = sections[sections.length - 1]
    if (previous && !previous.prefix && previous.end === start && previous.sizeCss === sizeCss) previous.end = end
    else sections.push({ start, end, sizeCss, cssBase: 0, rankBase: 0, maskBase: 0, count: 0, total: 0 })
  }
  function appendSource(start: number, end: number): void {
    let index = endingAfter(source, start, segment => segment.end)
    while (start < end) {
      const segment = source[index++]
      const limit = Math.min(end, segment.end)
      appendUniform(start, limit, segment.sizeCss)
      start = limit
    }
  }
  let displayEnd = 0
  for (const band of order.bands) {
    appendSource(displayEnd, band.r0)
    const first = endingAfter(source, band.r0, segment => segment.end)
    const last = endingAfter(source, band.r1, segment => segment.end)
    const height = source[first].sizeCss
    let uniform = true
    for (let index = first + 1; index <= last; index++) if (source[index].sizeCss !== height) { uniform = false; break }
    if (uniform) appendUniform(band.r0, band.r1 + 1, height)
    else {
      const length = band.r1 - band.r0 + 1
      const sizes = new Float64Array(length)
      const prefix = new Float64Array(length + 1)
      const positive = new Uint32Array(length + 1)
      for (let offset = 0; offset < length; offset++) {
        const row = band.sourceAtOffset[offset]
        const sourceIndex = endingAfter(source, row, segment => segment.end)
        sizes[offset] = maskVisible(band.r0 + offset) ? source[sourceIndex].sizeCss : 0
        prefix[offset + 1] = prefix[offset] + sizes[offset]
        positive[offset + 1] = positive[offset] + (sizes[offset] > 0 ? 1 : 0)
      }
      densePrefixBytes += prefix.byteLength
      typedArrayBytes += sizes.byteLength + prefix.byteLength + positive.byteLength
      sections.push({ start: band.r0, end: band.r1 + 1, sizeCss: 0, sizes, prefix, positive,
        cssBase: 0, rankBase: 0, maskBase: 0, count: positive[length], total: prefix[length] })
    }
    displayEnd = band.r1 + 1
  }
  appendSource(displayEnd, logicalCount)
  let total = 0, visibleCount = 0
  const positiveSections: Section[] = []
  for (const section of sections) {
    section.cssBase = total
    section.rankBase = visibleCount
    section.maskBase = maskRank(section.start)
    if (!section.prefix) {
      section.count = section.sizeCss > 0 ? maskRank(section.end) - section.maskBase : 0
      section.total = section.count * section.sizeCss
    }
    total += section.total
    visibleCount += section.count
    if (!Number.isFinite(total)) throw new AxisMetricsError('TOTAL_NOT_FINITE', 'Projected axis total must be finite')
    if (section.count) {
      if (section.prefix) {
        // A locally valid prefix may alias after another section's global origin
        // is added. Validate once per geometry generation, never during a paint.
        for (let offset = 0; offset < section.sizes!.length; offset++) {
          const height = section.sizes![offset]
          if (height > 0) validateCellPrecision(section.cssBase + section.prefix[offset],
            section.cssBase + section.prefix[offset + 1], height)
        }
      } else {
        // The largest visible origin gives a conservative spacing bound for
        // every row in the uniform span, including spans interrupted by a mask.
        const lastStart = section.cssBase + (section.count - 1) * section.sizeCss
        validateCellPrecision(section.cssBase, section.cssBase + section.sizeCss, section.sizeCss)
        validateCellPrecision(lastStart, total, section.sizeCss)
      }
      positiveSections.push(section)
    }
  }
  const projectionId = ++nextProjectionId
  let lastWindowVisitedSlots = 0
  const sectionOf = (slot: DisplaySlot) => sections[endingAfter(sections, slot, section => section.end)]
  function sectionStart(section: Section, slot: DisplaySlot): number {
    return section.cssBase + (section.prefix ? section.prefix[slot - section.start]
      : section.sizeCss * (maskRank(slot) - section.maskBase))
  }
  function sectionSize(section: Section, slot: DisplaySlot): number {
    return section.sizes ? section.sizes[slot - section.start] : maskVisible(slot) ? section.sizeCss : 0
  }
  function rowSize(slot: DisplaySlot): number {
    integer(slot, logicalCount - 1)
    return sectionSize(sectionOf(slot), slot)
  }
  function rowStart(slot: DisplaySlot): number {
    integer(slot, logicalCount)
    return slot === logicalCount ? total : sectionStart(sectionOf(slot), slot)
  }
  function slotInSection(section: Section, ordinal: VisibleRank): DisplaySlot {
    if (!section.positive) return visibility ? visibility.select(section.maskBase + ordinal)! : section.start + ordinal
    let low = 0, high = section.sizes!.length
    while (low < high) {
      const mid = Math.floor((low + high) / 2)
      if (section.positive[mid + 1] <= ordinal) low = mid + 1
      else high = mid
    }
    return section.start + low
  }
  function visibleRankToDisplay(ordinal: VisibleRank): DisplaySlot | null {
    integer(ordinal, MAX_ROWS)
    if (ordinal >= visibleCount) return null
    const section = positiveSections[endingAfter(positiveSections, ordinal, item => item.rankBase + item.count)]
    return slotInSection(section, ordinal - section.rankBase)
  }
  function displayToVisibleRank(slot: DisplaySlot): VisibleRank | null {
    integer(slot, logicalCount - 1)
    const section = sectionOf(slot)
    if (sectionSize(section, slot) === 0) return null
    return section.rankBase + (section.positive ? section.positive[slot - section.start] : maskRank(slot) - section.maskBase)
  }
  function locate(css: number): DisplaySlot | null {
    finite(css)
    if (css < 0 || css >= total) return null
    const section = positiveSections[endingAfter(positiveSections, css, item => item.cssBase + item.total)]
    // Compare canonical starts rather than trusting division at fractional boundaries.
    // Dense prefixes include plateaus: upper-bound picks the positive row after any zeros.
    if (section.prefix) {
      let low = 0, high = section.prefix.length
      while (low < high) {
        const mid = Math.floor((low + high) / 2)
        if (section.cssBase + section.prefix[mid] <= css) low = mid + 1
        else high = mid
      }
      return section.start + low - 1
    }
    const ordinalStart = (ordinal: number) => section.cssBase + ordinal * section.sizeCss
    const candidate = Math.max(0, Math.min(section.count - 1, Math.floor((css - section.cssBase) / section.sizeCss)))
    if (ordinalStart(candidate) <= css && (candidate === section.count - 1 || ordinalStart(candidate + 1) > css)) {
      return slotInSection(section, candidate)
    }
    let low = 0, high = section.count
    while (low < high) {
      const mid = Math.floor((low + high) / 2)
      if (ordinalStart(mid) <= css) low = mid + 1
      else high = mid
    }
    return slotInSection(section, low - 1)
  }
  /** Maximal same-mask span, using hierarchical ranks rather than visiting hidden slots. */
  function maskSpanEnd(start: number, end: number, visible: boolean): number {
    if (!visibility) return end
    const base = maskRank(start)
    let low = start + 1, high = end + 1
    while (low < high) {
      const mid = Math.floor((low + high) / 2)
      const count = maskRank(mid) - base
      if (count === (visible ? mid - start : 0)) low = mid + 1
      else high = mid
    }
    return low - 1
  }
  /** One bounded page of at most WINDOW_WORK_LIMIT visited runs starting at visible rank
   * `ordinal`; `next` is the rank to resume from, or -1 when the window is exhausted. */
  function pageFrom(upper: number, ordinal: VisibleRank): { runs: AxisRun[]; count: number; next: VisibleRank } {
    lastWindowVisitedSlots = 0
    const runs: AxisRun[] = []
    let count = 0
    while (ordinal < visibleCount && lastWindowVisitedSlots < WINDOW_WORK_LIMIT) {
      const start = visibleRankToDisplay(ordinal)!
      const section = sectionOf(start)
      const startCss = sectionStart(section, start)
      if (startCss >= upper) { ordinal = visibleCount; break }
      lastWindowVisitedSlots++
      let end = start
      const sizeCss = sectionSize(section, start)
      if (!section.prefix) {
        end = maskSpanEnd(start, section.end, true) - 1
        if (sectionStart(section, end) >= upper) {
          let low = start, high = end + 1
          while (low < high) {
            const mid = Math.floor((low + high) / 2)
            if (sectionStart(section, mid) < upper) low = mid + 1
            else high = mid
          }
          end = low - 1
        }
      }
      const previous = runs[runs.length - 1]
      if (previous && previous.end + 1 === start && previous.sizeCss === sizeCss) previous.end = end
      else runs.push({ start, end, startCss, sizeCss })
      const length = end - start + 1
      count += length
      ordinal += length
    }
    const nextSlot = ordinal < visibleCount ? visibleRankToDisplay(ordinal) : null
    return { runs, count, next: nextSlot !== null && rowStart(nextSlot) < upper ? ordinal : -1 }
  }

  function bounds(startCss: number, endCss: number): { lower: number; upper: number; initial: VisibleRank } {
    finite(startCss)
    finite(endCss)
    if (endCss < startCss) throw new AxisMetricsError('INVALID_QUERY', 'Projection window is reversed')
    const lower = Math.max(0, startCss), upper = Math.min(total, endCss)
    const initialSlot = lower < upper ? locate(lower) : null
    return { lower, upper, initial: initialSlot === null ? visibleCount : displayToVisibleRank(initialSlot)! }
  }

  /** Visible rows covered by [lower,upper) as a rank difference: O(log n), no row visits. */
  function coveredCount(upper: number, initial: VisibleRank): number {
    if (initial >= visibleCount) return 0
    if (upper >= total) return visibleCount - initial
    const slot = locate(upper)!
    const rank = displayToVisibleRank(slot)!
    return (rowStart(slot) === upper ? rank : rank + 1) - initial
  }

  function boundedPage(startCss: number, endCss: number, cursor?: string): AxisWindow {
    const { upper, initial } = bounds(startCss, endCss)
    let ordinal = initial
    if (cursor !== undefined) {
      const parts = cursor.split(':')
      const next = Number(parts[4])
      if (parts.length !== 5 || parts[0] !== 'projection' || parts[1] !== String(projectionId)
        || parts[2] !== String(startCss) || parts[3] !== String(endCss)
        || !Number.isSafeInteger(next) || parts[4] !== String(next) || next <= initial || next >= visibleCount
        || rowStart(visibleRankToDisplay(next)!) >= upper) {
        throw new AxisMetricsError('INVALID_QUERY', 'Projection cursor does not belong to this window')
      }
      ordinal = next
    }
    const page = pageFrom(upper, ordinal)
    return { runs: page.runs, count: page.count,
      ...(page.next < 0 ? {} : { nextCursor: `projection:${projectionId}:${startCss}:${endCss}:${page.next}` }) }
  }

  /** Complete, lazy window: `runs` is re-iterable and pages internally (at most one
   * bounded page resident) until [startCss,endCss) is exhausted; adjacent equal runs
   * that straddle a page boundary are coalesced. `count` is the total covered rows and
   * `nextCursor` is never set. */
  function windowBetween(startCss: number, endCss: number): AxisWindow {
    const { upper, initial } = bounds(startCss, endCss)
    return {
      runs: {
        *[Symbol.iterator](): Iterator<AxisRun> {
          let ordinal = initial
          let pending: AxisRun | undefined
          while (ordinal >= 0) {
            const page = pageFrom(upper, ordinal)
            for (const run of page.runs) {
              if (pending && pending.end + 1 === run.start && pending.sizeCss === run.sizeCss) pending.end = run.end
              else {
                if (pending) yield pending
                pending = { ...run }
              }
            }
            ordinal = page.next
          }
          if (pending) yield pending
        },
      },
      count: coveredCount(upper, initial),
    }
  }

  // Display overrides can be as fragmented as a sorted band or visibility mask.
  // Build only a typed boundary index when somebody reads the array; intervals are
  // materialized on demand, never retained as one JavaScript object per row.
  let overrideBoundaries: Uint32Array | undefined
  function visitOverrides(emit: (start: number, end: number) => void): void {
    let pendingStart = -1, pendingEnd = 0, pendingSize = 0
    function visit(start: number, end: number, height: number): void {
      if (pendingStart >= 0 && (pendingEnd !== start || pendingSize !== height)) {
        emit(pendingStart, pendingEnd - 1)
        pendingStart = -1
      }
      if (height !== defaultSizeCss) {
        if (pendingStart < 0) { pendingStart = start; pendingSize = height }
        pendingEnd = end
      }
    }
    for (const section of sections) {
      let slot = section.start
      while (slot < section.end) {
        const height = sectionSize(section, slot)
        let end = slot + 1
        if (section.sizes) {
          while (end < section.end && section.sizes[end - section.start] === height) end++
        } else end = section.sizeCss === 0 ? section.end : maskSpanEnd(slot, section.end, maskVisible(slot))
        visit(slot, end, height)
        slot = end
      }
    }
    if (pendingStart >= 0) emit(pendingStart, pendingEnd - 1)
  }
  function boundaries(): Uint32Array {
    if (!overrideBoundaries) {
      let count = 0
      visitOverrides(() => count++)
      overrideBoundaries = new Uint32Array(count * 2)
      let offset = 0
      visitOverrides((start, end) => { overrideBoundaries![offset++] = start; overrideBoundaries![offset++] = end })
      typedArrayBytes += overrideBoundaries.byteLength
    }
    return overrideBoundaries
  }
  function overrideAt(index: number): AxisSizeInterval | undefined {
    const indexArray = boundaries()
    if (index < 0 || index >= indexArray.length / 2) return undefined
    const start = indexArray[index * 2]
    return Object.freeze({ start, end: indexArray[index * 2 + 1], sizeCss: rowSize(start) })
  }
  const overrides = new Proxy([] as AxisSizeInterval[], {
    get(target, property, receiver) {
      if (property === 'length') return boundaries().length / 2
      if (typeof property === 'string' && /^(0|[1-9][0-9]*)$/.test(property)) return overrideAt(Number(property))
      return Reflect.get(target, property, receiver)
    },
    has(target, property) {
      if (typeof property === 'string' && /^(0|[1-9][0-9]*)$/.test(property)) return Number(property) < boundaries().length / 2
      return Reflect.has(target, property)
    },
    ownKeys() {
      // Enumeration explicitly requests all metadata keys; do not retain them.
      const keys: string[] = []
      const count = boundaries().length / 2
      for (let index = 0; index < count; index++) keys.push(String(index))
      keys.push('length')
      return keys
    },
    getOwnPropertyDescriptor(target, property) {
      if (property === 'length') {
        // The target's nonconfigurable length remains writable internally, so
        // reporting the virtual value respects Proxy descriptor invariants.
        return { ...Reflect.getOwnPropertyDescriptor(target, property)!, value: boundaries().length / 2 }
      }
      if (typeof property === 'string' && /^(0|[1-9][0-9]*)$/.test(property)) {
        const value = overrideAt(Number(property))
        return value === undefined ? undefined : { value, enumerable: true, configurable: true, writable: false }
      }
      return Reflect.getOwnPropertyDescriptor(target, property)
    },
    // Making the empty backing target nonextensible would invalidate virtual keys.
    preventExtensions() { throw new TypeError('Projected overrides are immutable') },
    setPrototypeOf() { throw new TypeError('Projected overrides are immutable') },
    set() { throw new TypeError('Projected overrides are immutable') },
    deleteProperty() { throw new TypeError('Projected overrides are immutable') },
    defineProperty() { throw new TypeError('Projected overrides are immutable') },
  })
  const rows: ProjectedRows = Object.freeze({ logicalCount, defaultSizeCss, overrides,
    totalCss: () => total, start: rowStart, size: rowSize, locate,
    visibleBetween: windowBetween,
    visiblePageBetween: (startCss: number, endCss: number) => boundedPage(startCss, endCss),
    continueVisibleBetween: (startCss: number, endCss: number, cursor: string) => boundedPage(startCss, endCss, cursor),
  })
  return Object.freeze({ rows, cols: options.cols, order,
    sourceCellToLogical(at: SourceCell) {
      integer(at.row, logicalCount - 1)
      integer(at.col, options.cols.logicalCount - 1)
      const slot = order.sourceToDisplay(at.row)
      const height = rowSize(slot), width = options.cols.size(at.col)
      return height > 0 && width > 0 ? { x: options.cols.start(at.col), y: rowStart(slot), width, height } : null
    },
    logicalPointToSource(x: number, y: number) {
      finite(x)
      finite(y)
      const col = options.cols.locate(x), slot = locate(y)
      return col === null || slot === null ? null : { row: order.displayToSource(slot), col }
    },
    visibleRankToDisplay,
    displayToVisibleRank,
    visibleRankToSource(ordinal: VisibleRank): SourceRow | null {
      const slot = visibleRankToDisplay(ordinal)
      return slot === null ? null : order.displayToSource(slot)
    },
    sourceToVisibleRank(row: SourceRow): VisibleRank | null {
      integer(row, logicalCount - 1)
      return displayToVisibleRank(order.sourceToDisplay(row))
    },
    get diagnostics() { return { sectionCount: sections.length, densePrefixBytes, typedArrayBytes, lastWindowVisitedSlots } },
  })
}
