// Bounded motion-path parser for p:animMotion (spec 6.1).
//
// PowerPoint motion paths are a small SVG-like mini-language: M (moveTo), L (lineTo),
// C (cubic bezier), Z (close) and E (end). Upper-case commands are absolute, lower-case
// relative. Coordinates are fractions of the slide.
//
// Decisions (the brief is silent on these):
//  - E ends the path; anything after it is ignored. E is a terminator, not a segment, so it
//    does not count toward the 10,000-command limit (Z does).
//  - A path must be non-empty and start with M; PowerPoint always writes it that way.
//  - Implicit repeated operands (`L 1 0 2 0`) are refused as `syntax`: PowerPoint never
//    writes them. Tokens need no separating whitespace (`M0,0L1,0` is fine).
//  - Numbers are plain decimals (optionally with an exponent), and every resolved coordinate
//    must satisfy |v| <= MAX_COORD, so derived lengths and positions can never overflow.
//  - Segment `pts` are always absolute, whatever case the source used.
//  - `pointAt` returns `angle` in radians (Math.atan2, 0 = +x, y grows downwards on screen),
//    not degrees like `rAngDeg` in types.ts; consumers convert.
import type { MotionPath } from './types'

const MAX_CHARS = 65536
const MAX_COMMANDS = 10000
/** Largest accepted coordinate magnitude, in slide fractions (a slide is 1.0 wide). */
const MAX_COORD = 1e6
/** Chords per cubic segment. */
const CUBIC_SAMPLES = 32

type Seg = MotionPath['segments'][number]
type Fail = { ok: false; reason: 'too-long' | 'too-many-commands' | 'syntax' }

// One token: a command letter, a decimal number, separators, or anything else (stray).
// A number is matched whole (exponent included) before a following `e`/`E` can be read as a command.
const TOKEN = /([MLCZEmlcze])|([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)|[\s,]+|([\s\S])/gy

const ARITY: Readonly<Record<string, number>> = { M: 1, L: 1, C: 3 }

/**
 * Sampled polyline of a path: one entry per non-zero-length drawing edge, with the cumulative
 * length at the edge's end. Built once per path and reused by every `pointAt`.
 */
interface Table {
  x0: Float64Array
  y0: Float64Array
  x1: Float64Array
  y1: Float64Array
  cum: Float64Array
  total: number
  /** Where the pen sits when there is nothing to walk. */
  startX: number
  startY: number
}

const tables = new WeakMap<object, Table>()

function emptyTable(startX: number, startY: number): Table {
  const none = new Float64Array(0)
  return { x0: none, y0: none, x1: none, y1: none, cum: none, total: 0, startX, startY }
}

function buildTable(segments: MotionPath['segments']): Table {
  const x0: number[] = []
  const y0: number[] = []
  const x1: number[] = []
  const y1: number[] = []
  const cum: number[] = []
  let total = 0
  let cx = 0
  let cy = 0
  let sx = 0
  let sy = 0
  let firstX = 0
  let firstY = 0
  let seenM = false
  let bad = false

  const edgeTo = (x: number, y: number): void => {
    if (!Number.isFinite(x) || !Number.isFinite(y)) bad = true
    const d = Math.hypot(x - cx, y - cy)
    if (d > 0) {
      total += d
      x0.push(cx); y0.push(cy); x1.push(x); y1.push(y); cum.push(total)
    }
    cx = x
    cy = y
  }

  for (const seg of segments) {
    const p = seg.pts
    if (seg.cmd === 'M') {
      cx = sx = p[0]
      cy = sy = p[1]
      if (!seenM) { seenM = true; firstX = cx; firstY = cy }
    } else if (seg.cmd === 'L') {
      edgeTo(p[0], p[1])
    } else if (seg.cmd === 'C') {
      const ax = cx
      const ay = cy
      for (let s = 1; s <= CUBIC_SAMPLES; s++) {
        const t = s / CUBIC_SAMPLES
        const u = 1 - t
        const a = u * u * u
        const b = 3 * u * u * t
        const c = 3 * u * t * t
        const d = t * t * t
        edgeTo(a * ax + b * p[0] + c * p[2] + d * p[4], a * ay + b * p[1] + c * p[3] + d * p[5])
      }
    } else {
      edgeTo(sx, sy) // Z: straight edge back to the subpath start
    }
  }

  if (bad || !Number.isFinite(total)) return emptyTable(Number.isFinite(firstX) ? firstX : 0, Number.isFinite(firstY) ? firstY : 0)
  return {
    x0: Float64Array.from(x0),
    y0: Float64Array.from(y0),
    x1: Float64Array.from(x1),
    y1: Float64Array.from(y1),
    cum: Float64Array.from(cum),
    total,
    startX: firstX,
    startY: firstY,
  }
}

export function parseMotionPath(src: string): { ok: true; path: MotionPath } | Fail {
  if (typeof src !== 'string') return { ok: false, reason: 'syntax' }
  if (src.length > MAX_CHARS) return { ok: false, reason: 'too-long' }

  const segments: Seg[] = []
  let cx = 0
  let cy = 0
  let sx = 0
  let sy = 0
  let cmd = ''
  let rel = false
  let need = 0
  let operands: number[] = []

  const syntax: Fail = { ok: false, reason: 'syntax' }
  const commit = (): Fail | null => {
    const pts: number[] = []
    for (let k = 0; k < operands.length; k += 2) {
      // Relative points all start from the point where the command began.
      const x = rel ? cx + operands[k] : operands[k]
      const y = rel ? cy + operands[k + 1] : operands[k + 1]
      if (!(Math.abs(x) <= MAX_COORD) || !(Math.abs(y) <= MAX_COORD)) return syntax
      pts.push(x, y)
    }
    cx = pts[pts.length - 2]
    cy = pts[pts.length - 1]
    if (cmd === 'M') { sx = cx; sy = cy }
    segments.push({ cmd: cmd as 'M' | 'L' | 'C', pts })
    operands = []
    return null
  }

  TOKEN.lastIndex = 0
  for (let m = TOKEN.exec(src); m !== null; m = TOKEN.exec(src)) {
    if (m[3] !== undefined) return syntax
    if (m[1] !== undefined) {
      if (need > 0) return syntax // previous command is missing operands
      const letter = m[1].toUpperCase()
      if (letter === 'E') break
      if (segments.length === 0 && letter !== 'M') return syntax
      cmd = letter
      rel = m[1] !== letter
      if (letter === 'Z') {
        if (segments.length >= MAX_COMMANDS) return { ok: false, reason: 'too-many-commands' }
        segments.push({ cmd: 'Z', pts: [] })
        cx = sx
        cy = sy
        cmd = ''
        continue
      }
      if (segments.length >= MAX_COMMANDS) return { ok: false, reason: 'too-many-commands' }
      need = ARITY[letter] * 2
      continue
    }
    if (m[2] !== undefined) {
      if (need === 0) return syntax // number with no command wanting it
      const v = Number(m[2])
      if (!Number.isFinite(v)) return syntax
      operands.push(v)
      if (--need === 0) {
        const bad = commit()
        if (bad) return bad
      }
    }
  }
  if (need > 0) return syntax
  if (segments.length === 0) return syntax

  const table = buildTable(segments)
  const path: MotionPath = { segments, length: table.total }
  tables.set(path, table)
  return { ok: true, path }
}

/**
 * Position and heading at `fraction` (0..1, clamped; NaN counts as 0) of the path's arc
 * length. `angle` is in radians. The result is always finite.
 */
export function pointAt(path: MotionPath, fraction: number): { x: number; y: number; angle: number } {
  let table = tables.get(path)
  if (!table) {
    table = buildTable(path.segments)
    tables.set(path, table)
  }
  const n = table.cum.length
  if (n === 0) return finiteOr(table.startX, table.startY, 0)

  const f = Number.isNaN(fraction) ? 0 : Math.min(1, Math.max(0, fraction))
  const target = f * table.total
  // First edge whose end is at or beyond the target.
  let lo = 0
  let hi = n - 1
  while (lo < hi) {
    const mid = (lo + hi) >>> 1
    if (table.cum[mid] >= target) hi = mid
    else lo = mid + 1
  }
  const from = lo > 0 ? table.cum[lo - 1] : 0
  const span = table.cum[lo] - from
  const u = span > 0 ? Math.min(1, Math.max(0, (target - from) / span)) : 0
  const ax = table.x0[lo]
  const ay = table.y0[lo]
  const bx = table.x1[lo]
  const by = table.y1[lo]
  return finiteOr(ax + (bx - ax) * u, ay + (by - ay) * u, Math.atan2(by - ay, bx - ax))
}

function finiteOr(x: number, y: number, angle: number): { x: number; y: number; angle: number } {
  return Number.isFinite(x) && Number.isFinite(y) && Number.isFinite(angle) ? { x, y, angle } : { x: 0, y: 0, angle: 0 }
}
