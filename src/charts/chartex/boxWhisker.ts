/**
 * The ChartEx Box & Whisker family painter (spec 4.3, docs/charts/chartex-map.md §8; lazy chart chunk).
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
 */
import type { ChartFillStyle } from '../fills'
import type { ChartSeries, ChartExSeriesInfo } from '../model'
import { applyLineStyle, type ChartLineStyle } from '../style'
import { barGeometry } from '../barGeometry'
import type { GroupAxes } from '../axes'
import type { Box } from '../layout'
import { analyticInput } from '../points'

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
  inliers: number[]
}

function medianOf(arr: readonly number[]): number {
  const n = arr.length
  if (n === 0) return 0
  const mid = Math.floor(n / 2)
  return n % 2 === 1 ? arr[mid] : (arr[mid - 1] + arr[mid]) / 2
}

/** Computes box and whisker statistics for a dataset according to the specified quartile method. */
export function computeBoxWhiskerStats(values: readonly number[], quartileMethod: 'inclusive' | 'exclusive' = 'exclusive'): BoxWhiskerStats | undefined {
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
      inliers.push(v)
      if (v < minInlier) minInlier = v
      if (v > maxInlier) maxInlier = v
    }
  }

  const lowerWhisker = inliers.length > 0 ? minInlier : q1
  const upperWhisker = inliers.length > 0 ? maxInlier : q3

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
    inliers,
  }
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

/** Paints a Box & Whisker group. */
export function paintBoxWhisker(
  ctx: CanvasRenderingContext2D,
  group: BoxWhiskerGroupLike,
  seriesList: readonly BoxWhiskerSeriesLike[],
  axes: GroupAxes,
  kit: BoxWhiskerPaintKit
): void {
  const catAxis = axes.category ?? axes.x!
  const gap = group.gapWidth ?? 100
  const geom = barGeometry({ grouping: 'standard', overlap: 0, gapWidth: gap, series: { length: 1 } }, catAxis, axes.between, kit.plot)

  const meanPoints: Array<{ x: number; y: number }> = []

  for (let si = 0; si < seriesList.length; si++) {
    const s = seriesList[si]
    const src = s.source ?? (s as unknown as ChartSeries)
    const exInfo = group.chartEx?.series?.[si]

    const input = analyticInput(src, 'values')
    if (input.kind !== 'input' || input.values.length === 0) continue

    const qMethod = exInfo?.layoutPr?.quartileMethod ?? 'exclusive'
    const vis = exInfo?.layoutPr?.visibility ?? { connectorLines: true, meanLine: false, meanMarker: true, nonoutliers: false, outliers: true }

    const stats = computeBoxWhiskerStats(input.values, qMethod)
    if (!stats) continue

    const xStart = geom.start(si, 0)
    const W = geom.width
    const xCenter = xStart + W / 2
    const color = s.color ?? src.color ?? '5b9bd5'

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
        applyLineStyle(ctx, src.lineStyle, kit.plot)
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
      if (kit.fillWith(ctx, src.fill, color, boxRect)) {
        ctx.fillRect(boxRect.x, boxRect.y, boxRect.w, boxRect.h)
      }
      kit.strokeBorder(ctx, src.lineStyle, color, boxRect, boxRect)
    }

    // 3. Median line
    const yMedian = kit.at(stats.median)
    if (kit.admitPath(2)) {
      ctx.save()
      try {
        ctx.strokeStyle = `#${color}`
        ctx.lineWidth = 1
        applyLineStyle(ctx, src.lineStyle, kit.plot)
        ctx.beginPath()
        ctx.moveTo(xStart, yMedian)
        ctx.lineTo(xStart + W, yMedian)
        ctx.stroke()
      } finally {
        ctx.restore()
      }
    }

    // 4. Outliers
    if (vis.outliers !== false && stats.outliers.length > 0) {
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
    if (vis.nonoutliers === true && stats.inliers.length > 0) {
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
    meanPoints.push({ x: xCenter, y: yMean })

    if (vis.meanMarker !== false) {
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

  // 7. Mean line connecting consecutive series means
  const anyMeanLine = Boolean(group.chartEx?.series.some(s => s.layoutPr?.visibility?.meanLine === true))
  if (anyMeanLine && meanPoints.length >= 2) {
    if (kit.admitPath(meanPoints.length * 2)) {
      ctx.save()
      try {
        ctx.strokeStyle = '#595959'
        ctx.lineWidth = 1
        ctx.beginPath()
        ctx.moveTo(meanPoints[0].x, meanPoints[0].y)
        for (let i = 1; i < meanPoints.length; i++) {
          ctx.lineTo(meanPoints[i].x, meanPoints[i].y)
        }
        ctx.stroke()
      } finally {
        ctx.restore()
      }
    }
  }
}
