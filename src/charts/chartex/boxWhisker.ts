/**
 * The ChartEx Box & Whisker family painter and preparation (spec 4.3, docs/charts/chartex-map.md §8; lazy chart chunk).
 *
 * Rules:
 * - Statistics use points in `value` state (blank and error points excluded). An unavailable point refuses the chart
 *   (prepare.ts, before this module's preparation runs).
 * - Quartile method: `exclusive` (default; excludes median when splitting halves) or `inclusive` (includes median).
 *   When n is even, both methods agree.
 * - Whiskers reach the furthest data points within [Q1 - 1.5*IQR, Q3 + 1.5*IQR].
 * - Outliers are data points beyond the whiskers, drawn when `visibility.outliers` is true (default true).
 * - Inner points (non-outliers) are drawn when `visibility.nonoutliers` is true (default false).
 * - Mean marker (×) is drawn when `visibility.meanMarker` is true (default true).
 * - Mean line (per series `visibility.meanLine`, default false): without categories it joins the means of consecutive
 *   series (map §8); with categories it joins each series' means across the categories.
 * - Empty series draws no box, and the chart remains supported.
 * - Categories: when a `cat` dimension is present, boxes are grouped by category with series side by side. When it
 *   is absent (or every category is empty), each series is one box along the axis. An empty category among named
 *   ones is its own group (empty label). More than CHART_POINT_LIMIT distinct categories refuse the chart.
 * - Hidden series are skipped.
 * - Stage 1 (prepare) buckets each series' points by category in one pass, sorts each bucket in place (n log n) and
 *   retains only the summaries plus the outlier / inner point lists. Every scratch and retained allocation is
 *   charged through the budget before it is made. Paint only reads the retained records and allocates nothing.
 */
import type { ChartFillStyle } from '../fills'
import type { ChartDiagnostic, ChartGroup, ChartModel, ChartSeries, ChartExGroup } from '../model'
import { applyLineStyle, type ChartLineStyle } from '../style'
import { barGeometry } from '../barGeometry'
import type { GroupAxes } from '../axes'
import type { Box } from '../layout'
import { stringBytes } from '../points'
import { CHART_POINT_LIMIT } from '../parse'

/** Charged bytes per distinct-category entry (a Map/Set entry plus its list slot). */
export const BOX_WHISKER_DISTINCT_BYTES = 64
/** Retained bytes per prepared box record (outlier / inner point lists are priced separately). */
const BOX_RECORD_BYTES = 160

export interface BoxWhiskerStats {
  n: number
  min: number
  max: number
  q1: number
  median: number
  q3: number
  iqr: number
  lowerWhisker: number
  upperWhisker: number
  mean: number
  outliers: number[]
  inliers?: number[]
}

/** Median of the sorted range [lo, hi). */
function medianRange(sorted: ArrayLike<number>, lo: number, hi: number): number {
  const len = hi - lo
  const mid = lo + Math.floor(len / 2)
  return len % 2 === 1 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2
}

/**
 * Statistics of an ascending, non-empty range without copying it. `retain` is told the bytes of the retained
 * outlier (and inner point) lists before they are allocated.
 */
export function boxWhiskerStatsOfSorted(
  sorted: ArrayLike<number>,
  quartileMethod: 'inclusive' | 'exclusive',
  keepInliers: boolean,
  retain: (bytes: number) => void = () => {},
): BoxWhiskerStats {
  const n = sorted.length
  const median = medianRange(sorted, 0, n)
  let lower: [number, number], upper: [number, number]
  if (n % 2 === 0) { lower = [0, n / 2]; upper = [n / 2, n] }
  else {
    const m = Math.floor(n / 2)
    if (quartileMethod === 'inclusive') { lower = [0, m + 1]; upper = [m, n] }
    else { lower = [0, m]; upper = [m + 1, n] }
  }
  const q1 = lower[1] > lower[0] ? medianRange(sorted, lower[0], lower[1]) : median
  const q3 = upper[1] > upper[0] ? medianRange(sorted, upper[0], upper[1]) : median
  const iqr = q3 - q1
  const lowerBound = q1 - 1.5 * iqr
  const upperBound = q3 + 1.5 * iqr
  // Sorted: the outliers are a prefix and a suffix; the inliers are [lo, hi).
  let lo = 0, hi = n
  while (lo < n && sorted[lo] < lowerBound) lo++
  while (hi > lo && sorted[hi - 1] > upperBound) hi--
  retain((lo + n - hi) * 8 + (keepInliers ? (hi - lo) * 8 : 0))
  const outliers: number[] = []
  for (let i = 0; i < lo; i++) outliers.push(sorted[i])
  for (let i = hi; i < n; i++) outliers.push(sorted[i])
  let inliers: number[] | undefined
  if (keepInliers) { inliers = []; for (let i = lo; i < hi; i++) inliers.push(sorted[i]) }
  let sum = 0
  for (let i = 0; i < n; i++) sum += sorted[i]
  return {
    n,
    min: sorted[0],
    max: sorted[n - 1],
    q1,
    median,
    q3,
    iqr,
    lowerWhisker: hi > lo ? sorted[lo] : q1,
    upperWhisker: hi > lo ? sorted[hi - 1] : q3,
    mean: sum / n,
    outliers,
    ...(inliers ? { inliers } : {}),
  }
}

/** Computes box and whisker statistics for a dataset according to the specified quartile method. */
export function computeBoxWhiskerStats(
  values: readonly number[],
  quartileMethod: 'inclusive' | 'exclusive' = 'exclusive',
  keepInliers = false
): BoxWhiskerStats | undefined {
  if (values.length === 0) return undefined
  return boxWhiskerStatsOfSorted(Float64Array.from(values).sort(), quartileMethod, keepInliers)
}

/** True when series `si` of a ChartEx group is hidden (`series@hidden`; map §3). */
export const boxWhiskerSeriesHidden = (group: ChartGroup, si: number): boolean =>
  group.kind === 'chartEx' ? (group as ChartExGroup).chartEx?.series?.[si]?.hidden === true : false
/** A series' category source: its own cat dimension, else the chart's shared categories. */
const categorySource = (model: { categories: readonly string[] }, s: ChartSeries): readonly string[] =>
  s.categories && s.categories.length ? s.categories : model.categories ?? []

/**
 * Distinct category names of the visible series in appearance order, or undefined when there is no category
 * dimension (or every category is empty). Collection stops past `limit` (the caller refuses such a chart).
 */
export function boxWhiskerDistinctCategories(categories: readonly string[], group?: ChartGroup, limit = CHART_POINT_LIMIT): string[] | undefined {
  const sources: Array<readonly string[]> = []
  if (group) {
    group.series.forEach((s, si) => { if (!boxWhiskerSeriesHidden(group, si)) sources.push(categorySource({ categories }, s)) })
  } else sources.push(categories ?? [])
  const seen = new Set<string>()
  const list: string[] = []
  let named = false
  for (const source of sources) {
    for (let j = 0; j < source.length; j++) {
      const c = source[j] ?? ''
      if (c !== '') named = true
      if (seen.has(c)) continue
      seen.add(c)
      list.push(c)
      if (list.length > limit) return list
    }
  }
  return named ? list : undefined
}

/** The entries a distinct-category pass can hold (bounded by the limit plus one). */
export function boxWhiskerDistinctCap(model: { categories: readonly string[] }, group: ChartGroup): number {
  let total = 0
  group.series.forEach((s, si) => { if (!boxWhiskerSeriesHidden(group, si)) total += categorySource(model, s).length })
  return Math.min(total, CHART_POINT_LIMIT + 1)
}

/** Counts non-hidden series in a box & whisker group. */
export function boxWhiskerVisibleSeriesCount(group: ChartGroup): number {
  let n = 0
  for (let si = 0; si < group.series.length; si++) if (!boxWhiskerSeriesHidden(group, si)) n++
  return n
}

/** Formats category axis labels for box & whisker (shared helper for prepare and layout). */
export function boxWhiskerCategoryLabels(categories: readonly string[], group: ChartGroup): string[] {
  const distinct = boxWhiskerDistinctCategories(categories, group)
  if (distinct) return distinct.slice(0, CHART_POINT_LIMIT)
  const out: string[] = []
  group.series.forEach((s, si) => { if (!boxWhiskerSeriesHidden(group, si)) out.push(s.name ?? String(out.length + 1)) })
  return out
}

export interface PreparedBox {
  seriesIndex: number
  visibleSeriesIndex: number
  slotIndex: number // Category index if hasCategories, else visibleSeriesIndex
  seriesSlot: number // visibleSeriesIndex if hasCategories, else 0
  /** The box the mean line joins this one from (same series across categories; else the previous series), or -1. */
  previous: number
  stats: BoxWhiskerStats
  quartileMethod: 'inclusive' | 'exclusive'
  meanMarker: boolean
  meanLine: boolean
  nonoutliers: boolean
  outliers: boolean
}

export interface PreparedBoxWhiskerGroup {
  groupIndex: number
  hasCategories: boolean
  distinctCategories?: string[]
  boxes: PreparedBox[]
  visibleSeriesCount: number
  gapWidth: number
}

/** The stage-1 budget: retained bytes stay with the prepared record; scratch is released after each series. */
export interface BoxWhiskerBudget {
  retain(bytes: number): void
  scratch(bytes: number): void
  releaseScratch(): void
}

export const BOX_WHISKER_TOO_MANY_CATEGORIES = 'too-many-categories' as const

/**
 * Computes and retains all Box & Whisker statistics in stage 1 (prepare). Per series: one counting pass over the
 * points (category -> bucket), one placing pass into a single charged scratch buffer, then an in-place sort of each
 * bucket. Points whose category is missing are not drawn and are diagnosed once.
 */
export function prepareBoxWhisker(model: ChartModel, budget: BoxWhiskerBudget, diagnostics?: ChartDiagnostic[]): PreparedBoxWhiskerGroup[] | typeof BOX_WHISKER_TOO_MANY_CATEGORIES {
  const result: PreparedBoxWhiskerGroup[] = []
  let dropped = 0
  for (let g = 0; g < model.groups.length; g++) {
    const group = model.groups[g]
    if (group.family !== 'boxWhisker') continue
    const exGroup = group.kind === 'chartEx' ? (group as ChartExGroup) : undefined

    budget.scratch(boxWhiskerDistinctCap(model, group) * BOX_WHISKER_DISTINCT_BYTES)
    const distinct = boxWhiskerDistinctCategories(model.categories, group)
    budget.releaseScratch()
    if (distinct && distinct.length > CHART_POINT_LIMIT) return BOX_WHISKER_TOO_MANY_CATEGORIES
    const hasCategories = distinct !== undefined
    const K = distinct ? distinct.length : 1
    budget.retain(128 + (distinct ? distinct.reduce((n, c) => n + stringBytes(c), 64) : 0))

    const visible: number[] = []
    for (let si = 0; si < group.series.length; si++) if (!boxWhiskerSeriesHidden(group, si)) visible.push(si)
    const boxes: PreparedBox[] = []
    for (let vsi = 0; vsi < visible.length; vsi++) {
      const si = visible[vsi]
      const s = group.series[si]
      const exInfo = exGroup?.chartEx?.series?.[si]
      const qMethod = exInfo?.layoutPr?.quartileMethod ?? 'exclusive'
      const vis = exInfo?.layoutPr?.visibility ?? { connectorLines: true, meanLine: false, meanMarker: true, nonoutliers: false, outliers: true }
      const points = s.values
      const n = points?.length ?? 0
      if (n === 0) continue
      const cats = hasCategories ? categorySource(model, s) : undefined
      // Scratch: the category index (map), the bucket bounds and cursors, and one value buffer.
      budget.scratch((distinct ? K * BOX_WHISKER_DISTINCT_BYTES : 0) + (K + 1) * 4 * 2 + n * 8)
      try {
        const index = distinct ? new Map<string, number>() : undefined
        if (index && distinct) for (let k = 0; k < K; k++) index.set(distinct[k], k)
        const bucketOf = (j: number): number => {
          if (!cats || !index) return 0
          const c = cats[j]
          return c === undefined ? -1 : index.get(c) ?? -1
        }
        const starts = new Int32Array(K + 1)
        let missing = 0
        for (let j = 0; j < n; j++) {
          if (points.states[j] !== 1) continue
          const k = bucketOf(j)
          if (k < 0) { missing++; continue }
          starts[k + 1]++
        }
        for (let k = 0; k < K; k++) starts[k + 1] += starts[k]
        dropped += missing
        const total = starts[K]
        if (total === 0) continue
        const buffer = new Float64Array(total)
        const cursor = new Int32Array(K)
        for (let k = 0; k < K; k++) cursor[k] = starts[k]
        for (let j = 0; j < n; j++) {
          if (points.states[j] !== 1) continue
          const k = bucketOf(j)
          if (k < 0) continue
          buffer[cursor[k]++] = points.values[j]
        }
        for (let k = 0; k < K; k++) {
          if (starts[k + 1] === starts[k]) continue
          const bucket = buffer.subarray(starts[k], starts[k + 1]).sort()
          budget.retain(BOX_RECORD_BYTES)
          const stats = boxWhiskerStatsOfSorted(bucket, qMethod, vis.nonoutliers === true, bytes => budget.retain(bytes))
          const last = boxes.length - 1
          // Categories: the previous box of the same series; no categories: the previous series' box.
          const previous = last >= 0 && (!hasCategories || boxes[last].seriesIndex === si) ? last : -1
          boxes.push({
            seriesIndex: si,
            visibleSeriesIndex: vsi,
            slotIndex: hasCategories ? k : vsi,
            seriesSlot: hasCategories ? vsi : 0,
            previous,
            stats,
            quartileMethod: qMethod,
            meanMarker: vis.meanMarker !== false,
            meanLine: vis.meanLine === true,
            nonoutliers: vis.nonoutliers === true,
            outliers: vis.outliers !== false,
          })
        }
      } finally { budget.releaseScratch() }
    }

    result.push({
      groupIndex: g,
      hasCategories,
      distinctCategories: distinct,
      boxes,
      visibleSeriesCount: Math.max(1, visible.length),
      gapWidth: group.gapWidth ?? 100,
    })
  }
  if (dropped && diagnostics) diagnostics.push({ code: 'chart-point-uncategorized', reason: `${dropped} box & whisker ${dropped === 1 ? 'point has' : 'points have'} no category and ${dropped === 1 ? 'is' : 'are'} not drawn` })
  return result
}

export function preparedBoxWhiskerBytes(prepared?: PreparedBoxWhiskerGroup[]): number {
  if (!prepared) return 0
  let bytes = 0
  for (const g of prepared) {
    bytes += 128 + (g.distinctCategories ? g.distinctCategories.reduce((n, c) => n + stringBytes(c), 64) : 0)
    for (const b of g.boxes) {
      bytes += BOX_RECORD_BYTES + b.stats.outliers.length * 8 + (b.stats.inliers ? b.stats.inliers.length * 8 : 0)
    }
  }
  return bytes
}

/**
 * The fixed part of a box & whisker preparation reserved up front: per group a header, per series a record, and the
 * two distinct-category passes outside this module (axis slots and axis labels). Scratch and retained point lists
 * are charged by `prepareBoxWhisker` itself.
 */
export function boxWhiskerBound(model: ChartModel): number {
  let bound = 0
  for (const g of model.groups) {
    if (g.family !== 'boxWhisker') continue
    bound += 2048 + g.series.length * 512 + 2 * boxWhiskerDistinctCap(model, g) * BOX_WHISKER_DISTINCT_BYTES
  }
  return bound
}

export interface BoxWhiskerPaintKit {
  plot: Box
  admitPath(cost: number): boolean
  fillWith(ctx: CanvasRenderingContext2D, fill: ChartFillStyle | undefined, color: string, box: Box): boolean
  strokeBorder(ctx: CanvasRenderingContext2D, line: ChartLineStyle | undefined, color: string, box: Box, rect?: Box): void
  at(v: number): number
}

export interface BoxWhiskerGroupLike {
  family: string
  gapWidth?: number
}

/**
 * The colour of a box's median line and of a mean marker centred over the box (a mean outside the quartiles is on the
 * plot paper and keeps the series colour). Office shows both marks, so they must stay visible: when the box is filled
 * with the series colour (no `spPr` fill: the default, or a solid fill of that
 * colour) they are drawn in white, the plot paper; on any other box fill (another solid colour, a gradient, a
 * pattern, or no fill) they keep the series colour. The exact colour Office uses is **unverified** (C1 phase 2b).
 */
export function boxMarkColor(fill: ChartFillStyle | undefined, color: string): string {
  const fillColor = fill === undefined ? color : fill.kind === 'solid' ? fill.color : undefined
  return fillColor !== undefined && fillColor.toLowerCase() === color.toLowerCase() ? 'FFFFFF' : color
}

/**
 * Paints a Box & Whisker group from its stage-1 records. `seriesList` is the model's own series list (indexed by
 * `PreparedBox.seriesIndex`). Without a prepared record nothing is drawn: statistics are never computed at paint.
 */
export function paintBoxWhisker(
  ctx: CanvasRenderingContext2D,
  group: BoxWhiskerGroupLike,
  seriesList: readonly ChartSeries[],
  axes: GroupAxes,
  kit: BoxWhiskerPaintKit,
  prepared?: PreparedBoxWhiskerGroup
): void {
  if (!prepared || prepared.boxes.length === 0) return
  const catAxis = axes.category ?? axes.x!
  const gap = group.gapWidth ?? 100
  const preparedBoxes = prepared.boxes
  const geom = prepared.hasCategories
    ? barGeometry({ grouping: 'clustered', overlap: 0, gapWidth: gap, series: { length: prepared.visibleSeriesCount } }, catAxis, axes.between, kit.plot)
    : barGeometry({ grouping: 'standard', overlap: 0, gapWidth: gap, series: { length: 1 } }, catAxis, axes.between, kit.plot)
  const W = geom.width
  const centerOf = (box: PreparedBox) => geom.start(box.slotIndex, box.seriesSlot) + W / 2

  for (const box of preparedBoxes) {
    const stats = box.stats
    const xStart = geom.start(box.slotIndex, box.seriesSlot)
    const xCenter = xStart + W / 2
    const src = seriesList[box.seriesIndex]
    const color = src?.color ?? '5b9bd5'

    const yQ3 = kit.at(stats.q3)
    const yQ1 = kit.at(stats.q1)
    const yTop = Math.min(yQ3, yQ1)
    const yBot = Math.max(yQ3, yQ1)
    const h = Math.max(1, yBot - yTop)
    const boxRect: Box = { x: xStart, y: yTop, w: W, h }

    // 1. Whiskers (vertical lines from box edges to whisker endpoints)
    const yLower = kit.at(stats.lowerWhisker)
    const yUpper = kit.at(stats.upperWhisker)
    if (kit.admitPath(4)) {
      ctx.save()
      try {
        ctx.strokeStyle = `#${color}`
        ctx.lineWidth = 1
        applyLineStyle(ctx, src?.lineStyle, kit.plot)
        ctx.beginPath()
        ctx.moveTo(xCenter, yBot)
        ctx.lineTo(xCenter, yLower)
        ctx.moveTo(xCenter, yTop)
        ctx.lineTo(xCenter, yUpper)
        ctx.stroke()
      } finally {
        ctx.restore()
      }
    }

    // 2. Box rectangle
    if (kit.admitPath(5)) {
      if (kit.fillWith(ctx, src?.fill, color, boxRect)) {
        ctx.fillRect(boxRect.x, boxRect.y, boxRect.w, boxRect.h)
      }
      kit.strokeBorder(ctx, src?.lineStyle, color, boxRect, boxRect)
    }

    // 3. Median line (drawn over the box: in the contrasting mark colour when the box is filled with the series colour)
    const mark = boxMarkColor(src?.fill, color)
    const yMedian = kit.at(stats.median)
    if (kit.admitPath(2)) {
      ctx.save()
      try {
        ctx.strokeStyle = `#${color}`
        ctx.lineWidth = 1
        applyLineStyle(ctx, src?.lineStyle, kit.plot)
        if (mark !== color) ctx.strokeStyle = `#${mark}`
        ctx.beginPath()
        ctx.moveTo(xStart, yMedian)
        ctx.lineTo(xStart + W, yMedian)
        ctx.stroke()
      } finally {
        ctx.restore()
      }
    }

    // 4. Outliers
    if (box.outliers && stats.outliers.length > 0 && kit.admitPath(stats.outliers.length * 4)) {
      ctx.save()
      try {
        ctx.fillStyle = `#${color}`
        ctx.strokeStyle = `#${color}`
        ctx.lineWidth = 1
        for (const val of stats.outliers) {
          ctx.beginPath()
          ctx.arc(xCenter, kit.at(val), 3, 0, Math.PI * 2)
          ctx.fill()
          ctx.stroke()
        }
      } finally {
        ctx.restore()
      }
    }

    // 5. Inner points (non-outliers)
    if (box.nonoutliers && stats.inliers && stats.inliers.length > 0 && kit.admitPath(stats.inliers.length * 4)) {
      ctx.save()
      try {
        ctx.fillStyle = `#${color}`
        ctx.strokeStyle = `#${color}`
        ctx.lineWidth = 1
        for (const val of stats.inliers) {
          ctx.beginPath()
          ctx.arc(xCenter, kit.at(val), 2.5, 0, Math.PI * 2)
          ctx.fill()
          ctx.stroke()
        }
      } finally {
        ctx.restore()
      }
    }

    // 6. Mean marker (cross '×')
    const yMean = kit.at(stats.mean)
    if (box.meanMarker && kit.admitPath(4)) {
      ctx.save()
      try {
        // Contrast against what lies under the mark: the box's mark colour only when the cross is centred over the
        // box; a mean outside the quartiles sits on the plot paper and keeps the series colour (C1 review).
        ctx.strokeStyle = `#${yMean >= yTop && yMean <= yBot ? mark : color}`
        ctx.lineWidth = 1.5
        const r = 4
        ctx.beginPath()
        ctx.moveTo(xCenter - r, yMean - r)
        ctx.lineTo(xCenter + r, yMean + r)
        ctx.moveTo(xCenter - r, yMean + r)
        ctx.lineTo(xCenter + r, yMean - r)
        ctx.stroke()
      } finally {
        ctx.restore()
      }
    }
  }

  // 7. Mean line: each box with `meanLine` joins its `previous` box when that one has `meanLine` too, in this
  // box's series colour (no per-paint allocation: the join is stage-1 data).
  for (const box of preparedBoxes) {
    if (!box.meanLine || box.previous < 0) continue
    const from = preparedBoxes[box.previous]
    if (!from?.meanLine || !kit.admitPath(2)) continue
    ctx.save()
    try {
      ctx.strokeStyle = `#${seriesList[box.seriesIndex]?.color ?? '595959'}`
      ctx.lineWidth = 1
      ctx.beginPath()
      ctx.moveTo(centerOf(from), kit.at(from.stats.mean))
      ctx.lineTo(centerOf(box), kit.at(box.stats.mean))
      ctx.stroke()
    } finally {
      ctx.restore()
    }
  }
}
