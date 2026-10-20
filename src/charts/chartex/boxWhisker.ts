/**
 * The ChartEx Box & Whisker family painter and preparation (spec 4.3, docs/charts/chartex-map.md §8; lazy chart chunk).
 *
 * Rules:
 * - Statistics use points in `value` state (blank and error points excluded).
 * - Quartile method: `exclusive` (default; excludes median when splitting halves) or `inclusive` (includes median).
 *   When n is even, both methods agree.
 * - Whiskers reach the furthest data points within [Q1 - 1.5*IQR, Q3 + 1.5*IQR].
 * - Outliers are data points beyond the whiskers, drawn when `visibility.outliers` is true (default true).
 * - Inner points (non-outliers) are drawn when `visibility.nonoutliers` is true (default false).
 * - Mean marker (×) is drawn when `visibility.meanMarker` is true (default true).
 * - Mean line connects consecutive series means when `visibility.meanLine` is true (default false).
 * - Empty series draws no box, and the chart remains supported.
 * - Categories: when a `cat` dimension is present, boxes are grouped by category with series side-by-side.
 *   When absent, each series is drawn as one box along the axis.
 * - Hidden series are skipped.
 * - Prepared in stage 1 (preparation) so statistics are retained without paint-time allocation.
 */
import type { ChartFillStyle } from '../fills'
import type { ChartGroup, ChartModel, ChartSeries, ChartExSeriesInfo, ChartExGroup } from '../model'
import { applyLineStyle, type ChartLineStyle } from '../style'
import { barGeometry } from '../barGeometry'
import type { GroupAxes } from '../axes'
import type { Box } from '../layout'
import { POINT_UNAVAILABLE, stringBytes, valueAt } from '../points'

export class BoxWhiskerUnavailableError extends Error {
  constructor() {
    super("This chart's data is not available")
  }
}

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

function medianOf(arr: readonly number[]): number {
  const n = arr.length
  if (n === 0) return 0
  const mid = Math.floor(n / 2)
  return n % 2 === 1 ? arr[mid] : (arr[mid - 1] + arr[mid]) / 2
}

/** Computes box and whisker statistics for a dataset according to the specified quartile method. */
export function computeBoxWhiskerStats(
  values: readonly number[],
  quartileMethod: 'inclusive' | 'exclusive' = 'exclusive',
  keepInliers = false
): BoxWhiskerStats | undefined {
  if (values.length === 0) return undefined
  const sorted = [...values].sort((a, b) => a - b)
  const n = sorted.length
  const median = medianOf(sorted)

  let lowerHalf: number[]
  let upperHalf: number[]

  if (n % 2 === 0) {
    lowerHalf = sorted.slice(0, n / 2)
    upperHalf = sorted.slice(n / 2)
  } else {
    const m = Math.floor(n / 2)
    if (quartileMethod === 'inclusive') {
      lowerHalf = sorted.slice(0, m + 1)
      upperHalf = sorted.slice(m)
    } else {
      lowerHalf = sorted.slice(0, m)
      upperHalf = sorted.slice(m + 1)
    }
  }

  const q1 = lowerHalf.length > 0 ? medianOf(lowerHalf) : median
  const q3 = upperHalf.length > 0 ? medianOf(upperHalf) : median
  const iqr = q3 - q1

  const lowerBound = q1 - 1.5 * iqr
  const upperBound = q3 + 1.5 * iqr

  const inliers: number[] = []
  const outliers: number[] = []

  let minInlier = Infinity
  let maxInlier = -Infinity

  for (const v of sorted) {
    if (v < lowerBound || v > upperBound) {
      outliers.push(v)
    } else {
      if (keepInliers) inliers.push(v)
      if (v < minInlier) minInlier = v
      if (v > maxInlier) maxInlier = v
    }
  }

  const lowerWhisker = minInlier !== Infinity ? minInlier : q1
  const upperWhisker = maxInlier !== -Infinity ? maxInlier : q3

  let sum = 0
  for (const v of sorted) sum += v
  const mean = sum / n

  return {
    n,
    min: sorted[0],
    max: sorted[n - 1],
    q1,
    median,
    q3,
    iqr,
    lowerWhisker,
    upperWhisker,
    mean,
    outliers,
    ...(keepInliers ? { inliers } : {}),
  }
}

/** Extracts distinct category names in appearance order, or returns undefined if no cat dimension exists. */
export function boxWhiskerDistinctCategories(categories: readonly string[], group?: ChartGroup): string[] | undefined {
  let source: readonly string[] = categories
  if ((!source || source.length === 0) && group) {
    const s = group.series.find(ser => ser.categories && ser.categories.length > 0)
    if (s?.categories) source = s.categories
  }
  if (!source || source.length === 0) return undefined
  const seen = new Set<string>()
  const list: string[] = []
  for (const c of source) {
    if (c !== '' && !seen.has(c)) {
      seen.add(c)
      list.push(c)
    }
  }
  return list.length > 0 ? list : undefined
}

/** Counts non-hidden series in a box & whisker group. */
export function boxWhiskerVisibleSeriesCount(group: ChartGroup): number {
  const ex = group.kind === 'chartEx' ? (group as ChartExGroup) : undefined
  return group.series.filter((_, si) => !ex?.chartEx?.series?.[si]?.hidden).length
}

/** Formats category axis labels for box & whisker (shared helper for prepare and layout). */
export function boxWhiskerCategoryLabels(categories: readonly string[], group: ChartGroup): string[] {
  const distinct = boxWhiskerDistinctCategories(categories, group)
  if (distinct) return distinct
  const ex = group.kind === 'chartEx' ? (group as ChartExGroup) : undefined
  return group.series
    .filter((_, i) => !ex?.chartEx?.series?.[i]?.hidden)
    .map((s, i) => s.name ?? String(i + 1))
}

export interface PreparedBox {
  seriesIndex: number
  visibleSeriesIndex: number
  slotIndex: number // Category index if hasCategories, else visibleSeriesIndex
  seriesSlot: number // visibleSeriesIndex if hasCategories, else 0
  stats: BoxWhiskerStats
  color?: string
  source?: ChartSeries
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

/** Computes and retains all Box & Whisker statistics in stage 1 (prepare). */
export function prepareBoxWhisker(model: ChartModel): PreparedBoxWhiskerGroup[] {
  const result: PreparedBoxWhiskerGroup[] = []

  model.groups.forEach((group, g) => {
    if (group.family !== 'boxWhisker') return

    // An unavailable point in an analytic series refuses the chart.
    for (const s of group.series) {
      if (s.values?.states?.includes(POINT_UNAVAILABLE)) {
        throw new BoxWhiskerUnavailableError()
      }
    }

    const exGroup = group.kind === 'chartEx' ? (group as ChartExGroup) : undefined
    const distinct = boxWhiskerDistinctCategories(model.categories, group)
    const hasCategories = distinct !== undefined

    // Filter out hidden series
    const visibleSeries: Array<{ s: ChartSeries; si: number }> = []
    for (let si = 0; si < group.series.length; si++) {
      const s = group.series[si]
      const exInfo = exGroup?.chartEx?.series?.[si]
      if (exInfo?.hidden) continue
      visibleSeries.push({ s, si })
    }

    const visibleSeriesCount = Math.max(1, visibleSeries.length)
    const boxes: PreparedBox[] = []

    if (hasCategories && distinct) {
      for (let k = 0; k < distinct.length; k++) {
        const catName = distinct[k]
        for (let vsi = 0; vsi < visibleSeries.length; vsi++) {
          const { s, si } = visibleSeries[vsi]
          const exInfo = exGroup?.chartEx?.series?.[si]
          const qMethod = exInfo?.layoutPr?.quartileMethod ?? 'exclusive'
          const vis = exInfo?.layoutPr?.visibility ?? { connectorLines: true, meanLine: false, meanMarker: true, nonoutliers: false, outliers: true }

          const values: number[] = []
          const len = s.values?.length ?? 0
          for (let j = 0; j < len; j++) {
            const pointCat = s.categories?.[j] ?? model.categories?.[j]
            if (pointCat === catName && s.values.states[j] === 1) {
              const val = valueAt(s.values, j)
              if (val !== undefined) values.push(val)
            }
          }

          const stats = computeBoxWhiskerStats(values, qMethod, vis.nonoutliers === true)
          if (!stats) continue

          boxes.push({
            seriesIndex: si,
            visibleSeriesIndex: vsi,
            slotIndex: k,
            seriesSlot: vsi,
            stats,
            color: s.color,
            source: s,
            quartileMethod: qMethod,
            meanMarker: vis.meanMarker !== false,
            meanLine: vis.meanLine === true,
            nonoutliers: vis.nonoutliers === true,
            outliers: vis.outliers !== false,
          })
        }
      }
    } else {
      for (let vsi = 0; vsi < visibleSeries.length; vsi++) {
        const { s, si } = visibleSeries[vsi]
        const exInfo = exGroup?.chartEx?.series?.[si]
        const qMethod = exInfo?.layoutPr?.quartileMethod ?? 'exclusive'
        const vis = exInfo?.layoutPr?.visibility ?? { connectorLines: true, meanLine: false, meanMarker: true, nonoutliers: false, outliers: true }

        const values: number[] = []
        const len = s.values?.length ?? 0
        for (let j = 0; j < len; j++) {
          if (s.values.states[j] === 1) {
            const val = valueAt(s.values, j)
            if (val !== undefined) values.push(val)
          }
        }

        const stats = computeBoxWhiskerStats(values, qMethod, vis.nonoutliers === true)
        if (!stats) continue

        boxes.push({
          seriesIndex: si,
          visibleSeriesIndex: vsi,
          slotIndex: vsi,
          seriesSlot: 0,
          stats,
          color: s.color,
          source: s,
          quartileMethod: qMethod,
          meanMarker: vis.meanMarker !== false,
          meanLine: vis.meanLine === true,
          nonoutliers: vis.nonoutliers === true,
          outliers: vis.outliers !== false,
        })
      }
    }

    result.push({
      groupIndex: g,
      hasCategories,
      distinctCategories: distinct,
      boxes,
      visibleSeriesCount,
      gapWidth: group.gapWidth ?? 100,
    })
  })

  return result
}

export function preparedBoxWhiskerBytes(prepared?: PreparedBoxWhiskerGroup[]): number {
  if (!prepared) return 0
  let bytes = 0
  for (const g of prepared) {
    bytes += 128 + (g.distinctCategories ? g.distinctCategories.reduce((n, c) => n + stringBytes(c), 64) : 0)
    for (const b of g.boxes) {
      bytes += 160 + b.stats.outliers.length * 8 + (b.stats.inliers ? b.stats.inliers.length * 8 : 0)
    }
  }
  return bytes
}

export function boxWhiskerBound(model: ChartModel): number {
  return model.groups.filter(g => g.family === 'boxWhisker').reduce((n, g) => n + 2048 + g.series.length * 512, 0)
}

export interface BoxWhiskerPaintKit {
  plot: Box
  admitPath(cost: number): boolean
  fillWith(ctx: CanvasRenderingContext2D, fill: ChartFillStyle | undefined, color: string, box: Box): boolean
  strokeBorder(ctx: CanvasRenderingContext2D, line: ChartLineStyle | undefined, color: string, box: Box, rect?: Box): void
  at(v: number): number
}

interface BoxWhiskerSeriesLike {
  source?: ChartSeries
  color?: string
}

export interface BoxWhiskerGroupLike {
  family: string
  gapWidth?: number
  kind?: string
  chartEx?: { series: ChartExSeriesInfo[] }
}

/** Paints a Box & Whisker group using precomputed stage-1 summaries. */
export function paintBoxWhisker(
  ctx: CanvasRenderingContext2D,
  group: BoxWhiskerGroupLike,
  seriesList: readonly BoxWhiskerSeriesLike[],
  axes: GroupAxes,
  kit: BoxWhiskerPaintKit,
  prepared?: PreparedBoxWhiskerGroup
): void {
  const catAxis = axes.category ?? axes.x!
  const gap = group.gapWidth ?? 100

  // Fallback preparation if none was provided
  const preparedBoxes = prepared?.boxes ?? []
  const hasCategories = prepared?.hasCategories ?? false
  const visibleSeriesCount = prepared?.visibleSeriesCount ?? Math.max(1, seriesList.length)

  const geom = hasCategories
    ? barGeometry({ grouping: 'clustered', overlap: 0, gapWidth: gap, series: { length: visibleSeriesCount } }, catAxis, axes.between, kit.plot)
    : barGeometry({ grouping: 'standard', overlap: 0, gapWidth: gap, series: { length: 1 } }, catAxis, axes.between, kit.plot)

  const meanPoints: Array<{ x: number; y: number; meanLine: boolean; color: string }> = []

  for (const box of preparedBoxes) {
    const stats = box.stats
    const xStart = geom.start(box.slotIndex, box.seriesSlot)
    const W = geom.width
    const xCenter = xStart + W / 2

    const src = box.source ?? seriesList[box.seriesIndex]?.source ?? (seriesList[box.seriesIndex] as unknown as ChartSeries)
    const color = box.color ?? src?.color ?? '5b9bd5'

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
        // Lower whisker
        ctx.moveTo(xCenter, yBot)
        ctx.lineTo(xCenter, yLower)
        // Upper whisker
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

    // 3. Median line
    const yMedian = kit.at(stats.median)
    if (kit.admitPath(2)) {
      ctx.save()
      try {
        ctx.strokeStyle = `#${color}`
        ctx.lineWidth = 1
        applyLineStyle(ctx, src?.lineStyle, kit.plot)
        ctx.beginPath()
        ctx.moveTo(xStart, yMedian)
        ctx.lineTo(xStart + W, yMedian)
        ctx.stroke()
      } finally {
        ctx.restore()
      }
    }

    // 4. Outliers
    if (box.outliers && stats.outliers.length > 0) {
      if (kit.admitPath(stats.outliers.length * 4)) {
        ctx.save()
        try {
          ctx.fillStyle = `#${color}`
          ctx.strokeStyle = `#${color}`
          ctx.lineWidth = 1
          for (const val of stats.outliers) {
            const y = kit.at(val)
            ctx.beginPath()
            ctx.arc(xCenter, y, 3, 0, Math.PI * 2)
            ctx.fill()
            ctx.stroke()
          }
        } finally {
          ctx.restore()
        }
      }
    }

    // 5. Inner points (non-outliers)
    if (box.nonoutliers && stats.inliers && stats.inliers.length > 0) {
      if (kit.admitPath(stats.inliers.length * 4)) {
        ctx.save()
        try {
          ctx.fillStyle = `#${color}`
          ctx.strokeStyle = `#${color}`
          ctx.lineWidth = 1
          for (const val of stats.inliers) {
            const y = kit.at(val)
            ctx.beginPath()
            ctx.arc(xCenter, y, 2.5, 0, Math.PI * 2)
            ctx.fill()
            ctx.stroke()
          }
        } finally {
          ctx.restore()
        }
      }
    }

    // 6. Mean marker (cross '×')
    const yMean = kit.at(stats.mean)
    meanPoints.push({ x: xCenter, y: yMean, meanLine: box.meanLine, color })

    if (box.meanMarker) {
      if (kit.admitPath(4)) {
        ctx.save()
        try {
          ctx.strokeStyle = `#${color}`
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
  }

  // 7. Mean line connecting consecutive series means that have meanLine enabled (M1)
  for (let i = 0; i + 1 < meanPoints.length; i++) {
    const p1 = meanPoints[i]
    const p2 = meanPoints[i + 1]
    if (p1.meanLine && p2.meanLine) {
      if (kit.admitPath(2)) {
        ctx.save()
        try {
          ctx.strokeStyle = `#${p1.color ?? p2.color ?? '595959'}`
          ctx.lineWidth = 1
          ctx.beginPath()
          ctx.moveTo(p1.x, p1.y)
          ctx.lineTo(p2.x, p2.y)
          ctx.stroke()
        } finally {
          ctx.restore()
        }
      }
    }
  }
}
