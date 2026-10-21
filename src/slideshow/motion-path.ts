// Bounded motion-path parser for p:animMotion (spec 6.1).
// PowerPoint motion paths are a small SVG-like mini-language: M (moveTo),
// L (lineTo), C (cubic bezier), Z (close) and E (end). Upper-case commands
// are absolute, lower-case relative. Coordinates are fractions of the slide.
import type { MotionPath } from './types'

const MAX_CHARS = 65536
const MAX_COMMANDS = 10000

type Seg = { cmd: 'M' | 'L' | 'C' | 'Z'; pts: readonly number[] }

function isNum(tok: string): boolean {
  if (tok === '' ) return false
  const n = Number(tok)
  return Number.isFinite(n)
}

export function parseMotionPath(src: string): { ok: true; path: MotionPath } | { ok: false; reason: 'too-long' | 'too-many-commands' | 'syntax' } {
  if (typeof src !== 'string') return { ok: false, reason: 'syntax' }
  if (src.length > MAX_CHARS) return { ok: false, reason: 'too-long' }
  const tokens = src.trim().split(/[\s,]+/).filter(t => t !== '')
  const segments: Seg[] = []
  let i = 0
  let cmdCount = 0
  // current point for relative commands
  let cx = 0, cy = 0
  // track start of subpath for Z/close
  let sx = 0, sy = 0
  while (i < tokens.length) {
    const tok = tokens[i]
    let cmd: 'M' | 'L' | 'C' | 'Z' | 'E' | null = null
    if (tok === 'M' || tok === 'm') cmd = 'M'
    else if (tok === 'L' || tok === 'l') cmd = 'L'
    else if (tok === 'C' || tok === 'c') cmd = 'C'
    else if (tok === 'Z' || tok === 'z') cmd = 'Z'
    else if (tok === 'E' || tok === 'e') cmd = 'E'
    else return { ok: false, reason: 'syntax' }
    const rel = tok === tok.toLowerCase() && tok !== tok.toUpperCase() // lowercase
    i++
    if (cmd === 'E') {
      // end marker: no coords
      continue
    }
    if (cmd === 'Z') {
      segments.push({ cmd: 'Z', pts: [] })
      cmdCount++
      if (cmdCount > MAX_COMMANDS) return { ok: false, reason: 'too-many-commands' }
      cx = sx; cy = sy
      continue
    }
    // need coordinates
    let nPts = cmd === 'M' ? 1 : cmd === 'L' ? 1 : 3 // C has 3 points (6 numbers)
    let coords: number[] = []
    for (let p = 0; p < nPts; p++) {
      if (i + 1 >= tokens.length && !(i + 1 === tokens.length && p === nPts-1)) {
        // need two numbers
      }
      if (i + 1 > tokens.length - 1) return { ok: false, reason: 'syntax' }
      const a = tokens[i], b = tokens[i+1]
      if (!isNum(a) || !isNum(b)) return { ok: false, reason: 'syntax' }
      let x = Number(a), y = Number(b)
      if (rel) { x += cx; y += cy }
      // for M/L single point; for C three points
      coords.push(x, y)
      // update current point for subsequent relative? For C, current becomes last control? PowerPoint: current point becomes end point (last pt)
      cx = x; cy = y
      i += 2
    }
    if (cmd === 'M') { sx = coords[0]; sy = coords[1] }
    const mappedCmd = cmd === 'M' ? 'M' : cmd === 'L' ? 'L' : 'C'
    segments.push({ cmd: mappedCmd, pts: coords })
    cmdCount++
    if (cmdCount > MAX_COMMANDS) return { ok: false, reason: 'too-many-commands' }
  }
  // compute length: sum of segment lengths (euclidean for L; cubic arc-length approximated later)
  let length = 0
  // recompute using segments
  // We'll compute length as total path length approximated by sampling in pointAt; store 0? Spec requires length field.
  // Compute polyline-ish length: for M/L use distance; for C use 32 samples.
  let px = 0, py = 0
  for (const seg of segments) {
    if (seg.cmd === 'M') {
      px = seg.pts[0]; py = seg.pts[1]
    } else if (seg.cmd === 'L') {
      const x = seg.pts[0], y = seg.pts[1]
      length += Math.hypot(x - px, y - py)
      px = x; py = y
    } else if (seg.cmd === 'C') {
      const [x1,y1,x2,y2,x,y] = seg.pts
      let lpx = px, lpy = py
      for (let s = 1; s <= 32; s++) {
        const t = s / 32
        const omt = 1 - t
        const bx = omt*omt*omt*px + 3*omt*omt*t*x1 + 3*omt*t*t*x2 + t*t*t*x
        const by = omt*omt*omt*py + 3*omt*omt*t*y1 + 3*omt*t*t*y2 + t*t*t*y
        length += Math.hypot(bx - lpx, by - lpy)
        lpx = bx; lpy = by
      }
      px = x; py = y
    } else if (seg.cmd === 'Z') {
      // close to subpath start; approximate no length added unless we track; skip
    }
  }
  const path: MotionPath = { segments, length }
  return { ok: true, path }
}

export function pointAt(path: MotionPath, fraction: number): { x: number; y: number; angle: number } {
  if (!path.segments.length) return { x: 0, y: 0, angle: 0 }
  const f = Math.max(0, Math.min(1, fraction))
  // build cumulative lengths
  type Pt = { x: number; y: number }
  // sample whole path into polyline with 32 samples per C, 1 per L
  // We'll compute positions along path by cumulative distance.
  // First, generate dense polyline.
  // We'll just iterate segments, track current point, accumulate distance array per vertex.
  let cur: Pt = { x: 0, y: 0 }
  // find first M
  const poly: { x: number; y: number; d: number }[] = [] // d cumulative from start
  let cum = 0
  // helper to add vertex
  const addVertex = (p: Pt) => {
    if (poly.length === 0) { poly.push({ x: p.x, y: p.y, d: 0 }); return }
    const last = poly[poly.length - 1]
    const dd = Math.hypot(p.x - last.x, p.y - last.y)
    cum += dd
    poly.push({ x: p.x, y: p.y, d: cum })
  }
  for (const seg of path.segments) {
    if (seg.cmd === 'M') {
      const p = { x: seg.pts[0], y: seg.pts[1] }
      // move: if poly empty or new subpath, just set current without adding length? PowerPoint moves pen. Add vertex with no length? We'll add vertex at new point with distance from previous? For M, reset cum? Actually M moves without drawing. Treat as new start: add vertex with cum unchanged (zero-length edge).
      if (poly.length > 0) {
        // add zero-length? skip adding distance
        poly.push({ x: p.x, y: p.y, d: cum })
      } else {
        poly.push({ x: p.x, y: p.y, d: 0 })
      }
      cur = p
    } else if (seg.cmd === 'L') {
      const p = { x: seg.pts[0], y: seg.pts[1] }
      addVertex(p)
      cur = p
    } else if (seg.cmd === 'C') {
      const [x1,y1,x2,y2,x,y] = seg.pts
      for (let s = 1; s <= 32; s++) {
        const t = s / 32
        const omt = 1 - t
        const bx = omt*omt*omt*cur.x + 3*omt*omt*t*x1 + 3*omt*t*t*x2 + t*t*t*x
        const by = omt*omt*omt*cur.y + 3*omt*omt*t*y1 + 3*omt*t*t*y2 + t*t*t*y
        addVertex({ x: bx, y: by })
      }
      cur = { x, y }
    } else if (seg.cmd === 'Z') {
      // close: add edge back to subpath start? skip
    }
  }
  if (poly.length === 0) return { x: 0, y: 0, angle: 0 }
  const totalLen = poly[poly.length-1].d
  if (totalLen === 0) return { x: poly[0].x, y: poly[0].y, angle: 0 }
  const target = f * totalLen
  // find segment
  let idx = 0
  for (let k = 0; k < poly.length - 1; k++) {
    if (poly[k+1].d >= target) { idx = k; break }
    idx = k
  }
  if (idx >= poly.length - 1) idx = poly.length - 2
  const a = poly[idx], b = poly[idx+1]
  const segLen = b.d - a.d
  const t = segLen > 0 ? (target - a.d) / segLen : 0
  const x = a.x + (b.x - a.x) * t
  const y = a.y + (b.y - a.y) * t
  const angle = Math.atan2(b.y - a.y, b.x - a.x)
  return { x, y, angle }
}
