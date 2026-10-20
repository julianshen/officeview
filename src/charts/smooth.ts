/**
 * Smooth line generation using Catmull-Rom cubic Bézier curves (spec §5.5).
 * Tangents come from neighbouring points (Catmull-Rom, tension 1/2); the end
 * segments use the end point as its own neighbour.
 *
 * This is an approximation of Office smoothing, unverified against references.
 */

export interface Point2D {
  x: number
  y: number
}

export interface BezierSegment {
  p0: Point2D
  cp1: Point2D
  cp2: Point2D
  p1: Point2D
}

/**
 * Converts a sequence of 2D points into cubic Bézier curve segments using
 * Catmull-Rom spline with tension 1/2 (spec §5.5).
 *
 * For each interior point P_i (1 <= i <= n - 2):
 *   T_i = 1/2 * (P_{i+1} - P_{i-1})
 * End segments use the end point as its own neighbour:
 *   T_0 = 1/2 * (P_1 - P_0)
 *   T_{n-1} = 1/2 * (P_{n-1} - P_{n-2})
 *
 * Each cubic Bézier segment from P_i to P_{i+1} has control points:
 *   cp1 = P_i + T_i / 3 = P_i + (P_{i+1} - P_{i-1}) / 6
 *   cp2 = P_{i+1} - T_{i+1} / 3 = P_{i+1} - (P_{i+2} - P_i) / 6
 */
export function catmullRomToBezier(points: readonly Point2D[]): BezierSegment[] {
  const n = points.length
  if (n < 2) return []
  const segments: BezierSegment[] = []
  for (let i = 0; i < n - 1; i++) {
    const p0 = points[i]
    const p1 = points[i + 1]
    const prev = i === 0 ? p0 : points[i - 1]
    const next = i + 1 === n - 1 ? p1 : points[i + 2]

    const cp1: Point2D = {
      x: p0.x + (p1.x - prev.x) / 6,
      y: p0.y + (p1.y - prev.y) / 6,
    }
    const cp2: Point2D = {
      x: p1.x - (next.x - p0.x) / 6,
      y: p1.y - (next.y - p0.y) / 6,
    }
    segments.push({ p0, cp1, cp2, p1 })
  }
  return segments
}

/**
 * Traces a Catmull-Rom smooth curve through a sequence of points in a Canvas 2D context.
 * Single points issue a moveTo; empty sequences do nothing.
 */
export function traceSmoothCurve(ctx: CanvasRenderingContext2D, points: readonly Point2D[]): void {
  if (points.length === 0) return
  if (points.length === 1) {
    ctx.moveTo(points[0].x, points[0].y)
    return
  }
  const segments = catmullRomToBezier(points)
  ctx.moveTo(segments[0].p0.x, segments[0].p0.y)
  for (let i = 0; i < segments.length; i++) {
    const seg = segments[i]
    ctx.bezierCurveTo(seg.cp1.x, seg.cp1.y, seg.cp2.x, seg.cp2.y, seg.p1.x, seg.p1.y)
  }
}
