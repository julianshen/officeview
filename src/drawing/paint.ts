/** DrawingML painting in local pixels; adapters own placement/rotation/flips. */
import type { ResolvedCommand, ResolvedGeometry, ResolvedPath } from './geometry'
import { resolveDrawingColor, type ArrowEnd, type DrawingColor, type DrawingFill, type DrawingLine, type DrawingStyle, type DrawingIssue } from './style'

type Point = [number, number]
interface PaintBounds { left: number; top: number; right: number; bottom: number }
interface Endpoint { point: Point; tangent: Point }
const presetDashes: Record<string, number[]> = {
  solid: [], dot: [1, 3], dash: [4, 3], lgDash: [8, 3], dashDot: [4, 3, 1, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDash: [3, 1], sysDot: [1, 1], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1],
}
const clamp = (n: number, max = 1): number => Math.min(max, Math.max(0, n))
const vector = (a: Point, b: Point): Point => [b[0] - a[0], b[1] - a[1]]
const nonzero = (v: Point): boolean => Math.hypot(...v) > 1e-10
function firstVector(...vectors: Point[]): Point | undefined { return vectors.find(nonzero) }
const finiteColor = (c: DrawingColor): boolean => [c.r, c.g, c.b, c.a].every(Number.isFinite)
const effectiveMiterLimit = (line: DrawingLine): number => Number.isFinite(line.miterLimit) && (line.miterLimit ?? 0) > 0 ? line.miterLimit! : 8
const css = (c: DrawingColor): string => `rgba(${clamp(c.r, 255)},${clamp(c.g, 255)},${clamp(c.b, 255)},${clamp(c.a)})`
function shadeColor(color: DrawingColor, mode: string): DrawingColor {
  const amount = ({ lighten: .6, lightenLess: .8, darken: .6, darkenLess: .8 } as Record<string, number>)[mode]
  if (amount === undefined) return color
  const shaded = resolveDrawingColor({
    kind: 'srgb',
    value: [color.r, color.g, color.b].map(v => Math.round(clamp(v, 255)).toString(16).padStart(2, '0')).join(''),
    transforms: [{ name: mode.startsWith('lighten') ? 'tint' : 'shade', value: String(amount * 100000) }],
  })
  return shaded ? { ...shaded, a: color.a } : color
}
function paintFill(ctx: CanvasRenderingContext2D, fill: DrawingFill | undefined, width: number, height: number, mode: string, issues: DrawingIssue[], bounds: PaintBounds): string | CanvasGradient | CanvasPattern | undefined {
  if (!fill || fill.kind === 'none') return undefined
  if (fill.kind === 'solid') {
    if (!finiteColor(fill.color)) { issues.push({ kind: 'invalid-paint', message: 'Nonfinite fill color' }); return undefined }
    return css(shadeColor(fill.color, mode))
  }
  if (!Number.isFinite(fill.angle) || fill.stops.some(stop => !Number.isFinite(stop.position) || !finiteColor(stop.color))) {
    issues.push({ kind: 'invalid-paint', message: 'Nonfinite gradient data' }); return undefined
  }
  if (!fill.stops.length) return undefined
  if (fill.gradient === 'circle') return radialPaint(ctx, fill, width, height, mode, issues, bounds)
  let gradient: CanvasGradient
  {
    let dx = Math.cos(fill.angle), dy = Math.sin(fill.angle)
    // Office uses (h*cos(a), w*sin(a)); the standard's prose swaps the extents.
    // Microsoft confirmed the erratum; the 400x200 native fixture agrees.
    // https://learn.microsoft.com/en-us/answers/questions/2265121/error-in-description-of-attribute-scaled-of-linear
    if (fill.scaled) { dx *= height; dy *= width }
    const norm = Math.hypot(dx, dy)
    if (!Number.isFinite(norm) || norm === 0) { issues.push({ kind: 'invalid-paint', message: 'Invalid linear gradient dimensions' }); return undefined }
    dx /= norm; dy /= norm
    const span = Math.abs(width * dx) + Math.abs(height * dy)
    const coordinates = [width / 2 - dx * span / 2, height / 2 - dy * span / 2, width / 2 + dx * span / 2, height / 2 + dy * span / 2] as [number, number, number, number]
    if (!coordinates.every(Number.isFinite)) { issues.push({ kind: 'invalid-paint', message: 'Invalid linear gradient dimensions' }); return undefined }
    gradient = ctx.createLinearGradient(...coordinates)
  }
  addLinearColorStops(gradient, fill.stops, mode)
  return gradient
}
type GradientFill = Extract<DrawingFill, { kind: 'gradient' }>
/** Canvas interpolates straight RGBA; Office interpolates premultiplied colors.
 * https://html.spec.whatwg.org/multipage/canvas.html#canvasgradient
 * Keep original stops (including ordered duplicates), then approximate only
 * varying-alpha intervals with at most 4096 additional stops per brush.
 * Opaque and constant-alpha intervals retain their original Canvas behavior.
 */
function addLinearColorStops(gradient: CanvasGradient, source: GradientFill['stops'], mode: string): void {
  const stops = source.map(stop => {
    const shaded = shadeColor(stop.color, mode)
    return { position: clamp(stop.position), color: { r: clamp(shaded.r, 255), g: clamp(shaded.g, 255), b: clamp(shaded.b, 255), a: clamp(shaded.a) } }
  })
  // Original insertion order is significant for duplicated hard-stop offsets.
  for (const stop of stops) gradient.addColorStop(stop.position, css(stop.color))
  stops.sort((a, b) => a.position - b.position)
  const varying = (a: typeof stops[number], b: typeof stops[number]): boolean => b.position > a.position && a.color.a !== b.color.a && (a.color.r !== b.color.r || a.color.g !== b.color.g || a.color.b !== b.color.b)
  let intervals = 0
  for (let i = 1; i < stops.length; i++) if (varying(stops[i - 1], stops[i])) intervals++
  if (intervals === 0) return
  const divisions = Math.min(64, Math.floor(4096 / intervals) + 1)
  for (let i = 1; i < stops.length; i++) {
    const a = stops[i - 1], b = stops[i]
    if (!varying(a, b)) continue
    for (let sample = 1; sample < divisions; sample++) {
      const weight = sample / divisions, position = a.position + (b.position - a.position) * weight
      // Very close finite offsets can round to an endpoint; never alter its hard stop.
      if (position <= a.position || position >= b.position) continue
      gradient.addColorStop(position, css(interpolateGradientColor(a.color, b.color, weight)))
    }
  }
}
function interpolateGradientColor(a: DrawingColor, b: DrawingColor, weight: number): DrawingColor {
  const alpha = a.a * (1 - weight) + b.a * weight
  const channel = (key: 'r' | 'g' | 'b'): number => alpha > 0 ? (a[key] * a.a * (1 - weight) + b[key] * b.a * weight) / alpha : 0
  return { r: channel('r'), g: channel('g'), b: channel('b'), a: alpha }
}
export interface PaintSurface { width: number; height: number; getContext(kind: '2d'): unknown }
/** Shared host-neutral offscreen-surface strategy (OffscreenCanvas, owning document,
 * or a node-canvas owner). Used by radial shading and the warp raster painter. */
export function scratchSurface(ctx: CanvasRenderingContext2D, width: number, height: number): PaintSurface | undefined {
  // OffscreenCanvas needs no document. node-canvas exposes its own Canvas constructor.
  if (typeof OffscreenCanvas !== 'undefined') return new OffscreenCanvas(width, height)
  const owner = ctx.canvas.ownerDocument
  if (typeof owner?.createElement === 'function') {
    const surface = owner.createElement('canvas')
    surface.width = width; surface.height = height
    return surface
  }
  try {
    const Constructor = ctx.canvas.constructor as unknown as new (width: number, height: number) => PaintSurface
    const surface = new Constructor(width, height)
    return typeof surface.getContext === 'function' ? surface : undefined
  } catch { return undefined }
}
function gradientColor(stops: GradientFill['stops'], position: number): DrawingColor {
  if (position <= stops[0].position) return stops[0].color
  for (let i = 1; i < stops.length; i++) {
    if (position > stops[i].position) continue
    const a = stops[i - 1], b = stops[i], distance = b.position - a.position
    const weight = distance > 0 ? (position - a.position) / distance : 1
    return interpolateGradientColor(a.color, b.color, weight)
  }
  return stops[stops.length - 1].color
}
/**
 * Office converts fillToRect to a point focus, then interpolates physical
 * circles from that point (stop 0) to the shape's circumcircle (stop 1).
 * Native PDF shading confirms outer center=(w/2,h/2), radius=hypot(w,h)/2.
 * A bounded texture composites every pixel once, including translucent stops.
 * Color interpolation here is premultiplied sRGB; native PDF ICC profiles and
 * sampled nonlinear color ramps can still differ from these colors.
 * https://learn.microsoft.com/en-us/answers/questions/2247174/how-is-attribute-filltorect-evaluated-for-a-gradie
 */
function radialPaint(ctx: CanvasRenderingContext2D, fill: GradientFill, width: number, height: number, mode: string, issues: DrawingIssue[], bounds: PaintBounds): CanvasPattern | undefined {
  const focus = { left: .5, top: .5, right: .5, bottom: .5, ...fill.focus }
  // Shipping fixed-point formula: left/(left+right), top/(top+bottom).
  // Degenerate/inverted inner rectangles retain their starting coordinate.
  const pointFocus = (start: number, end: number): number => 1 - start - end > 0 && start + end !== 0 ? start / (start + end) : start
  const radius = Math.hypot(width / 2, height / 2)
  const cx = width / radius / 2, cy = height / radius / 2
  const fx = pointFocus(focus.left, focus.right) * width / radius, fy = pointFocus(focus.top, focus.bottom) * height / radius
  const vx = cx - fx, vy = cy - fy, a = 1 - vx * vx - vy * vy
  const bw = bounds.right - bounds.left, bh = bounds.bottom - bounds.top
  const left = bounds.left / radius, top = bounds.top / radius, spanX = bw / radius, spanY = bh / radius
  if (![cx, cy, fx, fy, a, bw, bh, left, top, spanX, spanY].every(Number.isFinite) || width <= 0 || height <= 0 || bw <= 0 || bh <= 0) { issues.push({ kind: 'invalid-paint', message: 'Invalid radial gradient dimensions' }); return undefined }
  const sw = Math.min(1024, Math.max(1, Math.ceil(bw))), sh = Math.min(1024, Math.max(1, Math.ceil(bh)))
  const surface = scratchSurface(ctx, sw, sh)
  const scratch = surface?.getContext('2d') as CanvasRenderingContext2D | undefined
  if (!surface || !scratch) { issues.push({ kind: 'invalid-paint', message: 'An offscreen Canvas is required for radial shading' }); return undefined }
  const image = scratch.createImageData(sw, sh)
  const stops = fill.stops.map(stop => ({ position: clamp(stop.position), color: shadeColor(stop.color, mode) })).sort((a, b) => a.position - b.position)
  const position = (x: number, y: number): number => {
    if (Math.hypot(x - cx, y - cy) >= 1) return 1
    const dx = x - fx, dy = y - fy, distance2 = dx * dx + dy * dy
    if (distance2 === 0) return 0
    const b = dx * vx + dy * vy
    if (Math.abs(a) < 1e-12) return b > 0 ? clamp(distance2 / (2 * b)) : 1
    // |P - focus - t*(center-focus)|² = t² in radius-normalized space.
    const discriminant = b * b + a * distance2
    if (discriminant < 0) return 1
    const root = Math.sqrt(discriminant)
    let t = (-b + root) / a
    if (a < 0) t = Math.max(t, (-b - root) / a)
    return Number.isFinite(t) && t >= 0 ? clamp(t) : 1
  }
  for (let y = 0; y < sh; y++) {
    const py = top + (y + .5) / sh * spanY
    for (let x = 0; x < sw; x++) {
      const px = left + (x + .5) / sw * spanX
      const color = gradientColor(stops, position(px, py)), offset = (y * sw + x) * 4
      image.data[offset] = Math.round(clamp(color.r, 255)); image.data[offset + 1] = Math.round(clamp(color.g, 255)); image.data[offset + 2] = Math.round(clamp(color.b, 255)); image.data[offset + 3] = Math.round(clamp(color.a) * 255)
    }
  }
  scratch.putImageData(image, 0, 0)
  const pattern = ctx.createPattern(surface as unknown as CanvasImageSource, 'no-repeat')
  if (!pattern) { issues.push({ kind: 'invalid-paint', message: 'Unable to create radial gradient brush' }); return undefined }
  // Reuse the backend's matrix object, avoiding a global DOMMatrix dependency.
  const matrix = ctx.getTransform()
  matrix.a = bw / sw; matrix.b = 0; matrix.c = 0; matrix.d = bh / sh; matrix.e = bounds.left; matrix.f = bounds.top
  pattern.setTransform(matrix)
  return pattern
}
function arrowDimensions(end: ArrowEnd, lineWidth: number): { length: number; width: number } {
  // POI 5.4.1 corroborates Office's visible minimum: 2.5 points, then 1.5^enum.
  // https://github.com/apache/poi/blob/REL_5_4_1/poi/src/main/java/org/apache/poi/sl/draw/DrawSimpleShape.java
  const base = Math.max(10 / 3, lineWidth), scale = { sm: 1.5, med: 2.25, lg: 3.375 }
  return { length: base * scale[end.length ?? 'med'], width: base * scale[end.width ?? 'med'] }
}
function paintBounds(path: ResolvedPath, line: DrawingLine | undefined): PaintBounds | undefined {
  let left = Infinity, top = Infinity, right = -Infinity, bottom = -Infinity
  const include = (x: number, y: number): boolean => {
    if (!Number.isFinite(x) || !Number.isFinite(y)) return false
    left = Math.min(left, x); top = Math.min(top, y)
    right = Math.max(right, x); bottom = Math.max(bottom, y)
    return true
  }
  for (const command of path.commands) {
    if (command[0] === 'arcTo') {
      if (!include(command[1] - command[3], command[2] - command[4]) || !include(command[1] + command[3], command[2] + command[4])) return undefined
    } else if (command[0] !== 'close') {
      for (let i = 1; i < command.length; i += 2) if (!include(command[i] as number, command[i + 1] as number)) return undefined
    }
  }
  const lineWidth = Number.isFinite(line?.width) ? Math.max(0, line?.width ?? 1) : 1
  let padding = Math.max(1, lineWidth * 6)
  if (line?.join === 'miter') padding = Math.max(padding, lineWidth * effectiveMiterLimit(line) + 1)
  for (const end of [line?.headEnd, line?.tailEnd]) {
    if (!end?.type || end.type === 'none') continue
    const size = arrowDimensions(end, lineWidth)
    // Covers every orientation, including open-arrow pen thickness and AA.
    padding = Math.max(padding, Math.hypot(size.length, size.width / 2) + lineWidth / 2 + 1)
  }
  const bounds = { left: left - padding, top: top - padding, right: right + padding, bottom: bottom + padding }
  return Object.values(bounds).every(Number.isFinite) ? bounds : undefined
}
function validPath(commands: ResolvedCommand[]): boolean {
  const counts: Record<string, number> = { moveTo: 3, lnTo: 3, quadBezTo: 5, cubicBezTo: 7, arcTo: 8, close: 1 }
  let current = false
  for (const command of commands) {
    if (!Array.isArray(command) || command.length !== counts[command[0]]) return false
    const numbers = command[0] === 'arcTo' ? command.slice(1, 7) : command.slice(1)
    if (numbers.some(v => typeof v !== 'number' || !Number.isFinite(v))) return false
    if (command[0] === 'arcTo' && typeof command[7] !== 'boolean') return false
    if (command[0] === 'moveTo') current = true
    else if (!current) return false
    if (command[0] === 'arcTo' && (command[3] < 0 || command[4] < 0)) return false
  }
  return true
}
function drawPath(ctx: CanvasRenderingContext2D, commands: ResolvedCommand[]): void {
  ctx.beginPath()
  for (const command of commands) {
    switch (command[0]) {
      case 'moveTo': ctx.moveTo(command[1], command[2]); break
      case 'lnTo': ctx.lineTo(command[1], command[2]); break
      case 'quadBezTo': ctx.quadraticCurveTo(command[1], command[2], command[3], command[4]); break
      case 'cubicBezTo': ctx.bezierCurveTo(command[1], command[2], command[3], command[4], command[5], command[6]); break
      case 'arcTo':
        if (command[3] === 0 || command[4] === 0) ctx.lineTo(command[1] + command[3] * Math.cos(command[6]), command[2] + command[4] * Math.sin(command[6]))
        else ctx.ellipse(command[1], command[2], command[3], command[4], 0, command[5], command[6], command[7])
        break
      case 'close': ctx.closePath(); break
    }
  }
}
/** Derivatives at a stationary Bézier endpoint fall back to the next control. */
function endpoints(commands: ResolvedCommand[]): { head?: Endpoint; tail?: Endpoint } {
  let current: Point = [0, 0], start: Point = current, head: Endpoint | undefined, tail: Endpoint | undefined
  for (const c of commands) {
    if (c[0] === 'moveTo') { current = [c[1], c[2]]; start = current; continue }
    let end: Point, first: Point | undefined, last: Point | undefined
    if (c[0] === 'lnTo' || c[0] === 'close') {
      end = c[0] === 'close' ? start : [c[1], c[2]]
      first = last = firstVector(vector(current, end))
    } else if (c[0] === 'quadBezTo') {
      const control: Point = [c[1], c[2]]
      end = [c[3], c[4]]
      first = firstVector(vector(current, control), vector(current, end))
      last = firstVector(vector(control, end), vector(current, end))
    } else if (c[0] === 'cubicBezTo') {
      const c1: Point = [c[1], c[2]], c2: Point = [c[3], c[4]]
      end = [c[5], c[6]]
      first = firstVector(vector(current, c1), vector(current, c2), vector(current, end))
      last = firstVector(vector(c2, end), vector(c1, end), vector(current, end))
    } else {
      end = [c[1] + c[3] * Math.cos(c[6]), c[2] + c[4] * Math.sin(c[6])]
      if (c[3] === 0 || c[4] === 0) first = last = firstVector(vector(current, end))
      else if (Math.abs(c[6] - c[5]) > 1e-10) {
        const direction = c[7] ? -1 : 1
        first = firstVector([-c[3] * Math.sin(c[5]) * direction, c[4] * Math.cos(c[5]) * direction], vector(current, end))
        last = firstVector([-c[3] * Math.sin(c[6]) * direction, c[4] * Math.cos(c[6]) * direction], vector(current, end))
      }
    }
    if (first && !head) head = { point: current, tangent: first }
    if (last) tail = { point: end, tangent: last }
    current = end
  }
  return { head, tail }
}
function arrow(ctx: CanvasRenderingContext2D, endpoint: Endpoint | undefined, end: ArrowEnd | undefined, lineWidth: number, head: boolean): void {
  if (!endpoint || !end?.type || end.type === 'none') return
  const { length, width } = arrowDimensions(end, lineWidth)
  if (!Number.isFinite(length) || !Number.isFinite(width)) return
  const direction = head ? -1 : 1
  ctx.save()
  try {
    const angle = Math.atan2(endpoint.tangent[1] * direction, endpoint.tangent[0] * direction)
    const cos = Math.cos(angle), sin = Math.sin(angle)
    const point = (x: number, y: number): Point => [endpoint.point[0] + x * cos - y * sin, endpoint.point[1] + x * sin + y * cos]
    const move = (x: number, y: number): void => ctx.moveTo(...point(x, y))
    const to = (x: number, y: number): void => ctx.lineTo(...point(x, y))
    ctx.setLineDash([]); ctx.lineDashOffset = 0
    ctx.fillStyle = ctx.strokeStyle
    ctx.beginPath()
    // Native DrawingML oval and diamond decorations are centered on the endpoint.
    if (end.type === 'oval') ctx.ellipse(...point(0, 0), length / 2, width / 2, angle, 0, Math.PI * 2)
    else if (end.type === 'diamond') { move(length / 2, 0); to(0, -width / 2); to(-length / 2, 0); to(0, width / 2); ctx.closePath() }
    else if (end.type === 'arrow') { move(-length, -width / 2); to(0, 0); to(-length, width / 2) }
    else { move(0, 0); to(-length, -width / 2); if (end.type === 'stealth') to(-length * .65, 0); to(-length, width / 2); ctx.closePath() }
    if (end.type === 'arrow') ctx.stroke()
    else ctx.fill()
  } finally { ctx.restore() }
}
function setupLine(ctx: CanvasRenderingContext2D, line: DrawingLine, issues: DrawingIssue[]): number | undefined {
  const width = line.width ?? 1
  if (!Number.isFinite(width) || width < 0) { issues.push({ kind: 'invalid-paint', message: 'Invalid stroke width' }); return undefined }
  if (width === 0) return undefined
  ctx.lineWidth = width
  ctx.lineCap = line.cap ?? 'butt'
  ctx.lineJoin = line.join ?? 'round'
  ctx.miterLimit = effectiveMiterLimit(line)
  let dash: number[]
  if (Array.isArray(line.dash)) dash = line.dash
  else {
    const name = line.dash ?? 'solid', known = Object.prototype.hasOwnProperty.call(presetDashes, name)
    dash = known ? presetDashes[name] : []
    if (!known) issues.push({ kind: 'unsupported-line', message: `Unknown dash ${name}; solid used`, feature: name })
  }
  if (dash.some(d => !Number.isFinite(d) || d < 0 || !Number.isFinite(d * width)) || (dash.length > 0 && !dash.some(d => d > 0))) { issues.push({ kind: 'invalid-paint', message: 'Invalid dash pattern' }); dash = [] }
  ctx.setLineDash(dash.map(d => d * width)); ctx.lineDashOffset = 0
  return width
}

/** Does not alter the caller's Canvas drawing state. Returns painting diagnostics. */
export function paintGeometry(ctx: CanvasRenderingContext2D, geometry: ResolvedGeometry | ResolvedPath[], style: DrawingStyle, width: number, height: number): DrawingIssue[] {
  const issues: DrawingIssue[] = []
  if (![width, height].every(n => Number.isFinite(n) && n >= 0)) return [{ kind: 'invalid-paint', message: 'Invalid drawing extents' }]
  const paths = Array.isArray(geometry) ? geometry : geometry.paths
  ctx.save()
  try {
    for (const [pathIndex, path] of paths.entries()) {
      if (!validPath(path.commands)) { issues.push({ kind: 'invalid-paint', message: 'Invalid geometry path', pathIndex }); continue }
      const bounds = paintBounds(path, style.line)
      if (!bounds) { issues.push({ kind: 'invalid-paint', message: 'Invalid derived path bounds', pathIndex }); continue }
      const fill = path.fill !== 'none' && width > 0 && height > 0 ? paintFill(ctx, style.fill, width, height, path.fill, issues, bounds) : undefined
      const stroke = path.stroke && style.line ? paintFill(ctx, style.line.fill, width, height, 'norm', issues, bounds) : undefined
      drawPath(ctx, path.commands)
      if (fill) { ctx.fillStyle = fill; ctx.fill() }
      if (stroke && style.line) {
        const lineWidth = setupLine(ctx, style.line, issues)
        if (lineWidth === undefined) continue
        ctx.strokeStyle = stroke; ctx.stroke()
        const ends = endpoints(path.commands)
        arrow(ctx, ends.head, style.line.headEnd, lineWidth, true)
        arrow(ctx, ends.tail, style.line.tailEnd, lineWidth, false)
      }
    }
  } finally { ctx.restore() }
  return issues
}
