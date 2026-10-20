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

export interface LineTracer {
  push(x: number, y: number): void
  break(): void
  end(): void
}

/**
 * Creates a streaming line tracer that emits canvas path operations directly
 * with zero per-point heap allocations.
 *
 * When smooth is false, it emits ctx.moveTo on the first point of each run,
 * and ctx.lineTo on subsequent points.
 *
 * When smooth is true, it buffers a sliding window of at most 3 points
 * (using 6 primitive numbers) and emits Catmull-Rom cubic Bézier curve
 * segments (tension 1/2) directly to ctx.bezierCurveTo as points arrive.
 * End segments use the end point as its own neighbour (spec §5.5).
 */
export function createLineTracer(ctx: CanvasRenderingContext2D, smooth: boolean): LineTracer {
  if (!smooth) {
    let pen = false
    return {
      push(x: number, y: number): void {
        if (!pen) {
          ctx.moveTo(x, y)
          pen = true
        } else {
          ctx.lineTo(x, y)
        }
      },
      break(): void {
        pen = false
      },
      end(): void {
        pen = false
      },
    }
  }

  // Smooth mode: state machine with sliding window of primitive coordinates.
  // count: 0 (no points), 1 (have p0), 2 (have p0, p1), 3 (have prev, p0, p1)
  let count = 0
  let prevX = 0, prevY = 0
  let p0X = 0, p0Y = 0
  let p1X = 0, p1Y = 0

  function flushLast(): void {
    if (count < 2) {
      count = 0
      return
    }
    // Final segment: from p0 to p1.
    // End segment uses p1 as its own neighbour: next = p1.
    // If count == 2 (only p0 and p1 existed), prev was p0 (end segment uses p0 as own neighbour).
    const cp1x = p0X + (p1X - prevX) / 6
    const cp1y = p0Y + (p1Y - prevY) / 6
    const cp2x = p1X - (p1X - p0X) / 6
    const cp2y = p1Y - (p1Y - p0Y) / 6
    ctx.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, p1X, p1Y)
    count = 0
  }

  return {
    push(x: number, y: number): void {
      if (count === 0) {
        p0X = x
        p0Y = y
        ctx.moveTo(x, y)
        count = 1
        return
      }
      if (count === 1) {
        prevX = p0X
        prevY = p0Y
        p1X = x
        p1Y = y
        count = 2
        return
      }
      // count >= 2: we now have prev, p0, p1, and the new point (x, y) which acts as `next`.
      // Emit segment from p0 to p1:
      const cp1x = p0X + (p1X - prevX) / 6
      const cp1y = p0Y + (p1Y - prevY) / 6
      const cp2x = p1X - (x - p0X) / 6
      const cp2y = p1Y - (y - p0Y) / 6
      ctx.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, p1X, p1Y)

      // Slide window:
      prevX = p0X
      prevY = p0Y
      p0X = p1X
      p0Y = p1Y
      p1X = x
      p1Y = y
      count = 3
    },
    break(): void {
      flushLast()
    },
    end(): void {
      flushLast()
    },
  }
}

/**
 * Converts a sequence of 2D points into cubic Bézier curve segments using
 * Catmull-Rom spline with tension 1/2 (spec §5.5).
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
  const tracer = createLineTracer(ctx, true)
  for (let i = 0; i < points.length; i++) tracer.push(points[i].x, points[i].y)
  tracer.end()
}
