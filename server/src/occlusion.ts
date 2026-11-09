/**
 * §7.7 Stage-1-mandatory semantic occlusion EVIDENCE, recorded during the
 * paint pass that serves the unit (fail-closed). Nothing culls or ships it in
 * Stage 1; Stage 3's semantic services consume it.
 *
 * - Every `fillText`/`strokeText` run records device-space ink bounds (the
 *   transformed `measureText` actual bounding box, as an axis-aligned hull).
 * - An opaque-coverage record keeps every op that painted with effective
 *   alpha ≥ 0.95 under `source-over` (device box + sequence number) — the
 *   §7.7 grid at device precision. Bounds are conservative supersets (path points, stroke
 *   width, shadows), and a paint of unknown opacity (gradient, pattern,
 *   image) counts as opaque — so the grid can over-claim coverage, never
 *   under-claim it.
 * - A run is servable iff its ink lies on the surface, inside the current
 *   clip (tracked exactly only while every clip is an axis-aligned rect), is
 *   not fully transparent, and no ink cell was covered — or touched by a
 *   non-`source-over` op or a filter — after the run was drawn.
 */
import type { PaintObserver } from './render'

export type OcclusionReason = 'covered' | 'clipped' | 'off-surface' | 'composite' | 'filter' | 'nonrect-clip' | 'transparent'
export interface OcclusionEvidence {
  runs: number
  servable: number
  withheld: number
  reasons: Partial<Record<OcclusionReason, number>>
  /** The servable runs' text, in paint order (server-side only; never sent in Stage 1). */
  servableText: string[]
  /** Withheld runs and why (server-side only — for audits and tests). */
  withheldRuns: Array<{ text: string; reason: OcclusionReason }>
}

type Rect = [x0: number, y0: number, x1: number, y1: number]
type Clip = Rect | 'nonrect' | undefined
interface Run { seq: number; ink: Rect; text: string; reason?: OcclusionReason }

const EMPTY: Rect = [Infinity, Infinity, -Infinity, -Infinity]
const grow = (r: Rect, x: number, y: number): Rect => [Math.min(r[0], x), Math.min(r[1], y), Math.max(r[2], x), Math.max(r[3], y)]
const pad = (r: Rect, d: number): Rect => [r[0] - d, r[1] - d, r[2] + d, r[3] + d]
const inside = (a: Rect, b: Rect) => a[0] >= b[0] - 0.5 && a[1] >= b[1] - 0.5 && a[2] <= b[2] + 0.5 && a[3] <= b[3] + 0.5
const intersect = (a: Rect, b: Rect): Rect => [Math.max(a[0], b[0]), Math.max(a[1], b[1]), Math.min(a[2], b[2]), Math.min(a[3], b[3])]

/** Alpha of a canvas colour string; undefined for gradients/patterns/unparsed. */
function colorAlpha(style: unknown): number | undefined {
  if (typeof style !== 'string') return undefined
  const s = style.trim().toLowerCase()
  if (s === 'transparent') return 0
  if (/^#[0-9a-f]{6}$/.test(s) || /^#[0-9a-f]{3}$/.test(s)) return 1
  if (/^#[0-9a-f]{8}$/.test(s)) return parseInt(s.slice(7, 9), 16) / 255
  const m = /^rgba?\(([^)]*)\)$/.exec(s) ?? /^hsla?\(([^)]*)\)$/.exec(s)
  if (m) {
    const parts = m[1]!.split(/[\s,/]+/).filter(Boolean)
    if (parts.length < 4) return 1
    const a = parts[3]!
    return a.endsWith('%') ? parseFloat(a) / 100 : parseFloat(a)
  }
  return /^[a-z]+$/.test(s) ? 1 : undefined
}

type Grid = 'cover' | 'erase' | 'filter'
export function occlusionObserver(width: number, height: number): PaintObserver & { wrap(ctx: CanvasRenderingContext2D): CanvasRenderingContext2D; evidence(): OcclusionEvidence } {
  /** The coverage "grid" at full precision: every painting op's device box, in
   * sequence order (a 4 px cell grid withheld every spreadsheet cell next to a
   * gridline; a 1 px grid costs ~40 MB per page). */
  const marks: Array<{ seq: number; box: Rect; grid: Grid }> = []
  const surface: Rect = [0, 0, width, height]
  const runs: Run[] = []
  let seq = 0
  let clip: Clip
  const clipStack: Clip[] = []
  // Current path, in device space: per-subpath bounds (fill) and per-segment
  // bounds (stroke), so a grid of thin lines never claims the cells between them.
  // Join points get the miter allowance; plain segments only half the line width.
  let subpaths: Rect[] = [], segments: Rect[] = [], joins: Array<[number, number]> = [], segInSub = 0
  let cur: [number, number] | undefined, start: [number, number] | undefined
  let pathRects = 0, pathOther = false, pathRect: Rect | undefined

  const mark = (grid: Grid, r: Rect) => {
    const b = intersect(r, surface)
    if (b[2] > b[0] && b[3] > b[1]) marks.push({ seq, box: b, grid })
  }
  const overlaps = (a: Rect, b: Rect) => a[0] < b[2] && b[0] < a[2] && a[1] < b[3] && b[1] < a[3]

  function hull(ctx: CanvasRenderingContext2D, x0: number, y0: number, x1: number, y1: number): Rect {
    const t = ctx.getTransform()
    let r = EMPTY
    for (const [x, y] of [[x0, y0], [x1, y0], [x0, y1], [x1, y1]] as const) r = grow(r, t.a * x + t.c * y + t.e, t.b * x + t.d * y + t.f)
    return r
  }
  const axisAligned = (ctx: CanvasRenderingContext2D) => { const t = ctx.getTransform(); return Math.abs(t.b) < 1e-9 && Math.abs(t.c) < 1e-9 }
  /** Extend the current subpath through device points (one segment's hull). */
  const segment = (pts: Array<[number, number]>) => {
    let r = cur ? grow(EMPTY, cur[0], cur[1]) : EMPTY
    for (const p of pts) r = grow(r, p[0], p[1])
    if (!subpaths.length || !cur) { subpaths.push(r); start = pts[0]; segInSub = 0 }
    else subpaths[subpaths.length - 1] = grow(grow(subpaths[subpaths.length - 1]!, r[0], r[1]), r[2], r[3])
    if (segInSub++ > 0 && cur) joins.push(cur)
    segments.push(r)
    cur = pts[pts.length - 1]
    pathOther = true
  }
  const pts = (ctx: CanvasRenderingContext2D, a: unknown[], n: number) => Array.from({ length: n }, (_, i) => devicePoint(ctx, a[i * 2] as number, a[i * 2 + 1] as number))
  const box = (ctx: CanvasRenderingContext2D, x0: number, y0: number, x1: number, y1: number) => [devicePoint(ctx, x0, y0), devicePoint(ctx, x1, y0), devicePoint(ctx, x1, y1), devicePoint(ctx, x0, y1)]
  const devicePoint = (ctx: CanvasRenderingContext2D, x: number, y: number): [number, number] => { const t = ctx.getTransform(); return [t.a * x + t.c * y + t.e, t.b * x + t.d * y + t.f] }
  const maxScale = (ctx: CanvasRenderingContext2D) => { const t = ctx.getTransform(); return Math.max(Math.hypot(t.a, t.b), Math.hypot(t.c, t.d)) }

  /** One draw op with conservative device bounds `b`; `paint` = its fill/stroke style. */
  function draw(ctx: CanvasRenderingContext2D, b: Rect, paint: unknown, opts: { erase?: boolean } = {}) {
    seq++
    const shadowAlpha = colorAlpha(ctx.shadowColor) ?? 1
    if (shadowAlpha > 0 && (ctx.shadowBlur > 0 || ctx.shadowOffsetX || ctx.shadowOffsetY)) {
      const d = ctx.shadowBlur * 2 + Math.max(Math.abs(ctx.shadowOffsetX), Math.abs(ctx.shadowOffsetY))
      b = pad(b, d)
    }
    if (clip && clip !== 'nonrect') b = intersect(b, clip)
    if (ctx.filter && ctx.filter !== 'none') { mark('filter', b); return }
    if (opts.erase || ctx.globalCompositeOperation !== 'source-over') { mark('erase', b); return }
    const alpha = ctx.globalAlpha * (colorAlpha(paint) ?? 1)
    if (alpha >= 0.95) mark('cover', b)
  }

  // Half the line width covers butt/round/square caps and round/bevel joins;
  // a miter join can spike out to miterLimit × half the width at the join point.
  const strokeHalf = (ctx: CanvasRenderingContext2D) => (ctx.lineWidth * maxScale(ctx)) / 2 * Math.SQRT2
  const joinPad = (ctx: CanvasRenderingContext2D) => strokeHalf(ctx) * (ctx.lineJoin === 'miter' ? Math.max(1, ctx.miterLimit) : 1)
  /** One op painting several boxes (a multi-segment stroke): one sequence number. */
  function draws(ctx: CanvasRenderingContext2D, boxes: Rect[], paint: unknown) {
    if (!boxes.length) return
    draw(ctx, boxes[0]!, paint)
    for (let i = 1; i < boxes.length; i++) { seq--; draw(ctx, boxes[i]!, paint) }
  }

  function text(ctx: CanvasRenderingContext2D, s: unknown, x: number, y: number, stroke: boolean) {
    const str = String(s)
    const m = ctx.measureText(str)
    let ink = hull(ctx, x - m.actualBoundingBoxLeft, y - m.actualBoundingBoxAscent, x + m.actualBoundingBoxRight, y + m.actualBoundingBoxDescent)
    if (stroke) ink = pad(ink, (ctx.lineWidth * maxScale(ctx)) / 2)
    const style = stroke ? ctx.strokeStyle : ctx.fillStyle
    const alpha = ctx.globalAlpha * (colorAlpha(style) ?? 1)
    const run: Run = { seq: seq + 1, ink, text: str }
    if (str.trim() === '' || !(ink[2] > ink[0])) { draw(ctx, ink, style); return }
    if (alpha <= 0.05) run.reason = 'transparent'
    else if (!inside(ink, surface)) run.reason = 'off-surface'
    else if (clip === 'nonrect') run.reason = 'nonrect-clip'
    else if (clip && !inside(ink, clip)) run.reason = 'clipped'
    runs.push(run)
    draw(ctx, ink, style)
  }

  function imageRect(args: unknown[]): [number, number, number, number] {
    const img = args[0] as { width?: number; height?: number }
    if (args.length >= 9) return [args[5], args[6], args[7], args[8]] as number[] as [number, number, number, number]
    if (args.length >= 5) return [args[1], args[2], args[3], args[4]] as number[] as [number, number, number, number]
    return [args[1] as number, args[2] as number, img.width ?? 0, img.height ?? 0]
  }

  const hooks: Record<string, (ctx: CanvasRenderingContext2D, a: unknown[]) => void> = {
    save: () => { clipStack.push(clip) },
    restore: () => { if (clipStack.length) clip = clipStack.pop() },
    beginPath: () => { subpaths = []; segments = []; joins = []; segInSub = 0; cur = start = undefined; pathRects = 0; pathOther = false; pathRect = undefined },
    moveTo: (c, a) => { cur = start = devicePoint(c, a[0] as number, a[1] as number); subpaths.push(grow(EMPTY, cur[0], cur[1])); segInSub = 0 },
    lineTo: (c, a) => segment(pts(c, a, 1)),
    quadraticCurveTo: (c, a) => segment(pts(c, a, 2)),
    bezierCurveTo: (c, a) => segment(pts(c, a, 3)),
    arcTo: (c, a) => segment(pts(c, a, 2)),
    arc: (c, a) => { const [x, y, r] = a as number[]; segment(box(c, x! - r!, y! - r!, x! + r!, y! + r!)) },
    ellipse: (c, a) => { const [x, y, rx, ry] = a as number[]; const r = Math.max(rx!, ry!); segment(box(c, x! - r, y! - r, x! + r, y! + r)) },
    closePath: () => { if (cur && start) { segments.push(grow(grow(EMPTY, cur[0], cur[1]), start[0], start[1])); joins.push(cur, start); cur = start } },
    rect: (c, a) => {
      const [x, y, w, h] = a as number[]
      const [p0, p1, p2, p3] = box(c, x!, y!, x! + w!, y! + h!)
      const wasOther = pathOther
      cur = undefined; segment([p0!]); segment([p1!]); segment([p2!]); segment([p3!]); segment([p0!])
      pathOther = wasOther; pathRects++
      if (axisAligned(c)) pathRect = hull(c, x!, y!, x! + w!, y! + h!); else pathOther = true
    },
    roundRect: (c, a) => { const [x, y, w, h] = a as number[]; cur = undefined; segment(box(c, x!, y!, x! + w!, y! + h!)) },
    clip: (_c, a) => {
      const single = a.length === 0 || typeof a[0] === 'string'
      if (single && !pathOther && pathRects === 1 && pathRect && clip !== 'nonrect') clip = clip ? intersect(clip, pathRect) : pathRect
      else clip = 'nonrect'
    },
    // A Path2D argument has unknown geometry: the whole surface.
    fill: (c, a) => draws(c, a.length && typeof a[0] === 'object' ? [surface] : subpaths, c.fillStyle),
    stroke: (c, a) => draws(c, a.length && typeof a[0] === 'object' ? [surface] : [...segments.map((r) => pad(r, strokeHalf(c))), ...joins.map(([x, y]) => pad([x, y, x, y], joinPad(c)))], c.strokeStyle),
    fillRect: (c, a) => { const [x, y, w, h] = a as number[]; draw(c, hull(c, x!, y!, x! + w!, y! + h!), c.fillStyle) },
    strokeRect: (c, a) => {
      const [x, y, w, h] = a as number[]
      const [p0, p1, p2, p3] = box(c, x!, y!, x! + w!, y! + h!)
      const edges = [[p0, p1], [p1, p2], [p2, p3], [p3, p0]].map(([u, v]) => pad(grow(grow(EMPTY, u![0], u![1]), v![0], v![1]), strokeHalf(c)))
      for (const q of [p0, p1, p2, p3]) edges.push(pad([q![0], q![1], q![0], q![1]], joinPad(c)))
      draws(c, edges, c.strokeStyle)
    },
    clearRect: (c, a) => { const [x, y, w, h] = a as number[]; draw(c, hull(c, x!, y!, x! + w!, y! + h!), undefined, { erase: true }) },
    fillText: (c, a) => text(c, a[0], a[1] as number, a[2] as number, false),
    strokeText: (c, a) => text(c, a[0], a[1] as number, a[2] as number, true),
    drawImage: (c, a) => { const [x, y, w, h] = imageRect(a); draw(c, hull(c, x, y, x + w, y + h), undefined) },
    // putImageData ignores transform, clip and compositing: it replaces pixels.
    putImageData: (_c, a) => { seq++; const img = a[0] as { width: number; height: number }; const x = a[1] as number, y = a[2] as number; mark('erase', [x, y, x + img.width, y + img.height]) },
  }

  return {
    beforeUnit() {},
    afterUnit() {},
    wrap(ctx) {
      return new Proxy(ctx, {
        get(target, key) {
          const v = Reflect.get(target, key, target)
          if (typeof v !== 'function') return v
          const hook = typeof key === 'string' ? hooks[key] : undefined
          return (...args: unknown[]) => {
            // Bounds use the state BEFORE the op runs (transform, styles, clip).
            if (hook) hook(target, args)
            return Reflect.apply(v, target, args)
          }
        },
        set(target, key, value) { return Reflect.set(target, key, value, target) },
      })
    },
    evidence() {
      const reasons: OcclusionEvidence['reasons'] = {}
      const servableText: string[] = [], withheldRuns: OcclusionEvidence['withheldRuns'] = []
      for (const r of runs) {
        let reason = r.reason
        if (!reason) {
          const later = new Set<Grid>()
          for (const m of marks) if (m.seq > r.seq && overlaps(m.box, r.ink)) later.add(m.grid)
          reason = later.has('filter') ? 'filter' : later.has('erase') ? 'composite' : later.has('cover') ? 'covered' : undefined
        }
        if (reason) { reasons[reason] = (reasons[reason] ?? 0) + 1; withheldRuns.push({ text: r.text, reason }) }
        else servableText.push(r.text)
      }
      return { runs: runs.length, servable: servableText.length, withheld: runs.length - servableText.length, reasons, servableText, withheldRuns }
    },
  }
}
