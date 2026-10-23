/**
 * Slideshow plan S3-1: the reference models behind
 * docs/validation/slideshow-3d-algorithms.md, and the 50 % oracle values the
 * group tasks (S3-2 to S3-4) assert against. Pure maths on the framework's own
 * helpers (src/slideshow/render/three/mesh.ts); nothing here touches a GPU.
 *
 *   bun scripts/slideshow-3d-oracles.ts          # print every oracle as a markdown table
 *   bun scripts/slideshow-3d-oracles.ts --json   # the same as JSON
 *
 * Oracle inputs: aspect 16 : 9, the default field of view (45°), default
 * attributes unless a row names others, effect seed 2^31 (a shader rng whose
 * first value is 0.5), progress 0.5.
 */
import {
  DEFAULT_FOV_DEG, cellRandom, ease, mul4, project, rotationX, rotationY, rotationZ, scaling, slideCamera, slidePoint, translation, type Mat4,
} from '../src/slideshow/render/three/mesh'

export const ASPECT = 16 / 9
export const SEED = 2 ** 31
const a = ASPECT
const C = slideCamera(DEFAULT_FOV_DEG, a)
const T = translation, S = scaling, Rx = rotationX, Ry = rotationY, Rz = rotationZ
const m = (...ms: Mat4[]): Mat4 => ms.reduce((acc, x) => mul4(acc, x))
const ndc = (model: Mat4, u: number, v: number) => project(mul4(C, model), slidePoint(u, v, a))
const world = (model: Mat4, u: number, v: number): [number, number, number] => {
  const p = slidePoint(u, v, a)
  return [0, 1, 2].map(r => model[r] * p[0] + model[4 + r] * p[1] + model[8 + r] * p[2] + model[12 + r]) as [number, number, number]
}
const r4 = (n: number) => Math.round(n * 1e4) / 1e4
const xy = (p: readonly number[]) => `(${r4(p[0])}, ${r4(p[1])})`
const xyz = (p: readonly number[]) => `(${r4(p[0])}, ${r4(p[1])}, ${r4(p[2])})`
const clamp01 = (x: number) => Math.min(1, Math.max(0, x))

// ---------------------------------------------------------------------------
// Tile template (vortex, shred, fracture, crush, glitter, honeycomb, wind)
// ---------------------------------------------------------------------------
export interface TileSpec { cols: number; rows: number; W: number; J: number; g(i: number, j: number, uc: number, vc: number): number }
/** Delay and local time of cell (i, j): δ = W((1 − J) g + J r0), t = clamp((p − δ) / (1 − W)). */
export function tileTime(spec: TileSpec, seed: number, i: number, j: number, p: number): { delay: number; t: number } {
  const uc = (i + 0.5) / spec.cols, vc = (j + 0.5) / spec.rows
  const delay = spec.W * ((1 - spec.J) * spec.g(i, j, uc, vc) + spec.J * cellRandom(seed, i, j, 0))
  return { delay, t: clamp01((p - delay) / (1 - spec.W)) }
}
function fractions(spec: TileSpec, seed: number, p: number) {
  let rest = 0, done = 0
  for (let j = 0; j < spec.rows; j++) for (let i = 0; i < spec.cols; i++) {
    const { t } = tileTime(spec, seed, i, j, p)
    if (t === 0) rest++
    if (t === 1) done++
  }
  const n = spec.cols * spec.rows
  return { rest: r4(rest / n), done: r4(done / n) }
}
const centreOf = (spec: TileSpec, i: number, j: number) => slidePoint((i + 0.5) / spec.cols, (j + 0.5) / spec.rows, a)
const DIRS: Record<string, [number, number]> = { l: [-1, 0], r: [1, 0], u: [0, 1], d: [0, -1] }
/** Leading-edge order: the edge the motion heads for goes first. */
const leading = (dir: string) => (_i: number, _j: number, uc: number, vc: number) =>
  dir === 'l' ? uc : dir === 'r' ? 1 - uc : dir === 'u' ? vc : 1 - vc
const fromCentre = (_i: number, _j: number, uc: number, vc: number) => Math.min(1, Math.hypot((uc - 0.5) * a, vc - 0.5) / Math.hypot(a / 2, 0.5))
const fade = (t: number) => 1 - ease((t - 0.7) / 0.3)

// ---------------------------------------------------------------------------
// Rows
// ---------------------------------------------------------------------------
export interface Row { id: string; task: 'S3-2' | 'S3-3' | 'S3-4'; attrs: string; oracle(p: number): Record<string, string | number> }

/** prism: a cube (or the inside of a box) turning a quarter. */
export function prism(p: number, dir = 'l', inverted = false, content = false) {
  const th = (Math.PI / 2) * ease(p)
  const horizontal = dir === 'l' || dir === 'r'
  const D = horizontal ? a : 1
  const sigma = dir === 'l' || dir === 'u' ? -1 : 1
  const R = (ang: number) => (horizontal ? Ry(ang) : Rx(ang))
  const K = inverted ? D / 2 : -D / 2
  const s = inverted ? -sigma : sigma
  const lift = content ? T(0, 0, -0.3 * Math.sin(Math.PI * p)) : T(0, 0, 0)
  const from = m(lift, T(0, 0, K), R(s * th), T(0, 0, -K))
  const to = m(lift, T(0, 0, K), R(s * th - s * Math.PI / 2), T(0, 0, -K))
  return { from, to }
}
export function flip(p: number, dir = 'l') {
  const th = Math.PI * ease(p), sg = dir === 'r' ? 1 : -1, z = -0.5 * Math.sin(Math.PI * p)
  return { from: m(T(0, 0, z), Ry(sg * th)), to: m(T(0, 0, z), Ry(sg * th - sg * Math.PI)) }
}
export function gallery(p: number, dir = 'l') {
  const sg = dir === 'r' ? 1 : -1, e = ease(p), s = Math.sin(Math.PI * p)
  return { from: m(T(sg * 1.1 * a * e, 0, -0.6 * s), Ry(sg * 0.35 * s)), to: m(T(sg * 1.1 * a * (e - 1), 0, -0.6 * s), Ry(sg * 0.35 * s)) }
}
/** doors / window: halves swing towards the viewer about their outer edges; window's also part outwards and its incoming frame starts further back. */
export function doors(p: number, dir = 'vert', windowed = false) {
  const th = (Math.PI / 2) * ease(p), e = ease(p)
  const out = windowed ? 0.25 * a * e : 0
  if (dir === 'horz') {
    return {
      first: m(T(0, out / a, 0), T(0, 0.5, 0), Rx(-th), T(0, -0.5, 0)), // top half, uvRect [0, 0, 1, 0.5]
      second: m(T(0, -out / a, 0), T(0, -0.5, 0), Rx(th), T(0, 0.5, 0)), // bottom half, uvRect [0, 0.5, 1, 0.5]
      to: T(0, 0, (windowed ? -1.2 : -0.6) * (1 - e)),
    }
  }
  return {
    first: m(T(-out, 0, 0), T(-a / 2, 0, 0), Ry(-th), T(a / 2, 0, 0)), // left half, uvRect [0, 0, 0.5, 1]
    second: m(T(out, 0, 0), T(a / 2, 0, 0), Ry(th), T(-a / 2, 0, 0)), // right half, uvRect [0.5, 0, 0.5, 1]
    to: T(0, 0, (windowed ? -1.2 : -0.6) * (1 - e)),
  }
}
export function switchRow(p: number, dir = 'l') {
  const sg = dir === 'r' ? 1 : -1, s = Math.sin(Math.PI * p), e = ease(p)
  return { from: m(T(sg * 0.55 * a * s, 0, -1.2 * e), Ry(sg * 0.6 * s)), to: m(T(-sg * 0.55 * a * s, 0, -1.2 * (1 - e)), Ry(-sg * 0.6 * s)) }
}
export function conveyor(p: number, dir = 'l') {
  const sg = dir === 'r' ? 1 : -1, s = Math.sin(Math.PI * p), e = ease(p)
  const z = -0.8 * s, xf = sg * 1.1 * a * e
  return { from: m(T(xf, 0, z), Ry(sg * 0.25 * s)), to: m(T(xf - sg * 1.1 * a, 0, z), Ry(sg * 0.25 * s)) }
}
export function ferris(p: number, dir = 'l') {
  const sg = dir === 'r' ? 1 : -1, R = 0.9 * a, phi = (Math.PI / 3) * ease(p), z = -0.3 * Math.sin(Math.PI * p)
  const at = (psi: number) => T(R * Math.sin(psi), -R + R * Math.cos(psi), z)
  return { from: at(sg * phi), to: at(sg * (phi - Math.PI / 3)) }
}

const VORTEX: TileSpec = { cols: 32, rows: 18, W: 0.6, J: 0.35, g: leading('l') }
export function vortexTile(p: number, i: number, j: number, seed = SEED, dir = 'l') {
  const spec = { ...VORTEX, g: leading(dir) }
  const { t } = tileTime(spec, seed, i, j, p)
  const [mx, my] = DIRS[dir]
  const r1 = cellRandom(seed, i, j, 1), r2 = cellRandom(seed, i, j, 2), r3 = cellRandom(seed, i, j, 3)
  const c = centreOf(spec, i, j)
  const swirl = 0.4 * Math.sin(Math.PI * t) * (2 * r2 - 1)
  const tx = mx * 1.4 * a * t * t - my * swirl, ty = my * 1.4 * a * t * t + mx * swirl
  const model = m(T(c[0] + tx, c[1] + ty, 0.5 * t), Rz(2 * Math.PI * (0.75 + 0.5 * r1) * t), Ry(Math.PI * t * (r3 - 0.5)), S(1 - 0.5 * t, 1 - 0.5 * t, 1), T(-c[0], -c[1], 0))
  return { t, model, opacity: fade(t) }
}
const SHRED_STRIP: TileSpec = { cols: 24, rows: 1, W: 0.5, J: 0.2, g: () => 0 }
const SHRED_RECT: TileSpec = { cols: 24, rows: 14, W: 0.5, J: 0.6, g: () => 0 }
export function shredTile(p: number, i: number, j: number, pattern = 'strip', seed = SEED) {
  const spec = pattern === 'rectangle' ? SHRED_RECT : SHRED_STRIP
  const { t } = tileTime(spec, seed, i, j, p)
  const c = centreOf(spec, i, j)
  if (pattern !== 'rectangle') return { t, model: T(0, (i % 2 === 0 ? 1 : -1) * 1.3 * t * t, 0), opacity: fade(t) }
  const ang = 2 * Math.PI * cellRandom(seed, i, j, 1), r2 = cellRandom(seed, i, j, 2)
  const model = m(T(c[0] + 1.5 * t * Math.cos(ang), c[1] + 1.5 * t * Math.sin(ang), 0.8 * t), Rz(2 * Math.PI * r2 * t), S(1 - 0.6 * t, 1 - 0.6 * t, 1), T(-c[0], -c[1], 0))
  return { t, model, opacity: fade(t) }
}
const FRACTURE: TileSpec = { cols: 12, rows: 7, W: 0.6, J: 0.4, g: fromCentre }
export function fractureTile(p: number, i: number, j: number, seed = SEED) {
  const { t } = tileTime(FRACTURE, seed, i, j, p)
  const c = centreOf(FRACTURE, i, j), len = Math.hypot(c[0], c[1]) || 1
  const model = m(T(c[0] + 0.4 * (c[0] / len) * t, c[1] - 1.6 * t * t, 0.3 * t * cellRandom(seed, i, j, 3)), Rx(Math.PI * cellRandom(seed, i, j, 1) * t), Ry(Math.PI * cellRandom(seed, i, j, 2) * t), T(-c[0], -c[1], 0))
  return { t, model, opacity: fade(t) }
}
const CRUSH: TileSpec = { cols: 16, rows: 9, W: 0.3, J: 1, g: () => 0 }
export function crushTile(p: number, i: number, j: number, seed = SEED) {
  const { t } = tileTime(CRUSH, seed, i, j, p)
  const c = centreOf(CRUSH, i, j), k = 0.9 * ease(t), sc = 1 - 0.8 * ease(t)
  const ball = T(0, 0, -2 * ease((p - 0.6) / 0.4))
  const model = m(ball, T(c[0] * (1 - k), c[1] * (1 - k), 0), Rz(2 * Math.PI * (cellRandom(seed, i, j, 1) - 0.5) * t), Rx(Math.PI * cellRandom(seed, i, j, 2) * t), S(sc, sc, 1), T(-c[0], -c[1], 0))
  return { t, model, opacity: 1 - ease((p - 0.75) / 0.25) }
}
const GLITTER = (pattern: string, dir: string): TileSpec => ({ cols: pattern === 'hexagon' ? 24 : 32, rows: pattern === 'hexagon' ? 14 : 18, W: 0.65, J: 0.4, g: leading(dir) })
export function glitterTile(p: number, i: number, j: number, pattern = 'diamond', dir = 'l', seed = SEED) {
  const spec = GLITTER(pattern, dir)
  const { t } = tileTime(spec, seed, i, j, p)
  const c = centreOf(spec, i, j), angle = Math.PI * ease(t)
  const R = dir === 'l' || dir === 'r' ? Ry((dir === 'l' ? -1 : 1) * angle) : Rx((dir === 'u' ? -1 : 1) * angle)
  return { t, model: m(T(c[0], c[1], 0), R, T(-c[0], -c[1], 0)), showsTo: ease(t) > 0.5, shade: 1 + 0.6 * Math.sin(Math.PI * t) }
}
const HONEYCOMB: TileSpec = { cols: 20, rows: 11, W: 0.6, J: 0.3, g: fromCentre }
export function honeycombTile(p: number, i: number, j: number, seed = SEED) {
  const { t } = tileTime(HONEYCOMB, seed, i, j, p)
  const c = centreOf(HONEYCOMB, i, j), sc = 1 - ease(t)
  return { t, model: m(T(c[0], c[1], 0.6 * t), Rz((Math.PI / 3) * t), S(sc, sc, 1), T(-c[0], -c[1], 0)), opacity: 1 - ease(t) }
}
const RIPPLE_ORIGIN: Record<string, [number, number]> = { center: [0.5, 0.5], lu: [0, 0], ru: [1, 0], ld: [0, 1], rd: [1, 1] }
export function rippleZ(p: number, u: number, v: number, dir = 'center') {
  const [ou, ov] = RIPPLE_ORIGIN[dir] ?? RIPPLE_ORIGIN.center
  const rho = Math.hypot((u - ou) * a, v - ov)
  return 0.06 * 4 * p * (1 - p) * Math.sin(2 * Math.PI * (4 * rho - 3 * p)) * Math.exp(-1.5 * rho)
}
const WIND: TileSpec = { cols: 8, rows: 18, W: 0.55, J: 0.5, g: (_i, _j, uc, vc) => 0.6 * (1 - uc) + 0.4 * vc }
export function windTile(p: number, i: number, j: number, seed = SEED) {
  const { t } = tileTime(WIND, seed, i, j, p)
  const c = centreOf(WIND, i, j)
  const model = m(T(c[0] - 1.6 * a * t * t, c[1] + 0.25 * Math.sin(Math.PI * t) * (cellRandom(seed, i, j, 1) - 0.5), 0), Rz(0.6 * (cellRandom(seed, i, j, 2) - 0.5) * t), Ry((-Math.PI / 3) * t), T(-c[0], -c[1], 0))
  return { t, model, opacity: fade(t) }
}

/** The fold rule of the page curls and peel off, per pixel, in slide units (x right in [0, a], y down in [0, 1]). */
export function fold(p: number, x: number, y: number, page: { x0: number; x1: number; corner: [number, number]; target: [number, number] }): 'to' | 'back' | 'from' {
  if (x < page.x0 || x > page.x1) return 'from'
  const [cx, cy] = page.corner, [tx, ty] = page.target
  const D = Math.hypot(tx - cx, ty - cy), nx = (tx - cx) / D, ny = (ty - cy) / D
  const L = D * ease(p)
  const s = (x - cx) * nx + (y - cy) * ny
  if (s < L) return 'to'
  const mx = x - 2 * (s - L) * nx, my = y - 2 * (s - L) * ny
  return mx >= page.x0 && mx <= page.x1 && my >= 0 && my <= 1 ? 'back' : 'from'
}
function foldFractions(p: number, pages: Array<Parameters<typeof fold>[3]>) {
  const N = 640, M = 360
  let to = 0, back = 0
  for (let j = 0; j < M; j++) for (let i = 0; i < N; i++) {
    const x = ((i + 0.5) / N) * a, y = (j + 0.5) / M
    const page = pages.find(pg => x >= pg.x0 && x <= pg.x1)!
    const r = fold(p, x, y, page)
    if (r === 'to') to++
    if (r === 'back') back++
  }
  return { to: r4(to / (N * M)), back: r4(back / (N * M)) }
}
export const curlSingle = (invX = false, invY = false) => [{ x0: 0, x1: a, corner: [invX ? 0 : a, invY ? 0 : 1] as [number, number], target: [invX ? a : 0, invY ? 1 : 0] as [number, number] }]
export const curlDouble = (invY = false) => [
  { x0: 0, x1: a / 2, corner: [0, invY ? 0 : 1] as [number, number], target: [a / 2, invY ? 1 : 0] as [number, number] },
  { x0: a / 2, x1: a, corner: [a, invY ? 0 : 1] as [number, number], target: [a / 2, invY ? 1 : 0] as [number, number] },
]
export const peel = (invX = false) => [{ x0: 0, x1: a, corner: [invX ? a : 0, 1] as [number, number], target: [invX ? 0 : a, 0] as [number, number] }]

/** Right-handed rotation by `th` about the axis through world points `from` → `to`. */
export function axisRotation(from: readonly number[], to: readonly number[], th: number): Mat4 {
  const d = [to[0] - from[0], to[1] - from[1], to[2] - from[2]], n = Math.hypot(d[0], d[1], d[2])
  const [x, y, z] = d.map(v => v / n), c = Math.cos(th), s = Math.sin(th), t = 1 - c
  const r = new Float32Array([
    t * x * x + c, t * x * y + s * z, t * x * z - s * y, 0,
    t * x * y - s * z, t * y * y + c, t * y * z + s * x, 0,
    t * x * z + s * y, t * y * z - s * x, t * z * z + c, 0,
    0, 0, 0, 1,
  ])
  return m(T(from[0], from[1], from[2]), r, T(-from[0], -from[1], -from[2]))
}
/** airplane's four rigid triangles, in slide uv: the body halves either side of the spine and the wings under the creases from the top corners to the nose. */
export const AIRPLANE_TRIS = {
  bodyL: [[0, 0], [0.5, 0], [0.5, 1]], bodyR: [[0.5, 0], [1, 0], [0.5, 1]],
  wingL: [[0, 0], [0, 1], [0.5, 1]], wingR: [[1, 0], [1, 1], [0.5, 1]],
} as const
export type AirplanePanel = keyof typeof AIRPLANE_TRIS
/** airplane: a paper dart folded about diagonal creases (nose at the bottom centre), then banked, pitched and flown away into depth. */
export function airplane(p: number, invX = false) {
  const f = ease(2 * p), tau = ease(2 * p - 1), sx = invX ? -1 : 1
  const beta = (Math.PI / 4) * f, th = 0.9 * Math.PI * f
  const N = slidePoint(0.5, 1, a), TL = slidePoint(0, 0, a), TR = slidePoint(1, 0, a)
  const bodyL = Ry(-beta), bodyR = Ry(beta)
  const wingL = m(bodyL, axisRotation(N, TL, th)), wingR = m(bodyR, axisRotation(N, TR, -th))
  const k = 1 - 0.4 * tau
  const G = m(T(sx * 0.9 * a * tau, -2.4 * tau, -3 * tau), Rz(sx * 0.3 * tau), Ry(sx * 0.5 * tau), Rx(0.6 * tau), S(k, k, k))
  return { panels: { bodyL: m(G, bodyL), bodyR: m(G, bodyR), wingL: m(G, wingL), wingR: m(G, wingR) } as Record<AirplanePanel, Mat4>, opacity: 1 - ease((p - 0.85) / 0.15), wingShade: 1 - 0.1 * f }
}
/** The projected outline of the dart at `p`: its convex hull in NDC (counter-clockwise) and that hull's area. */
export function airplaneSilhouette(p: number, invX = false): { hull: number[][]; area: number } {
  const { panels } = airplane(p, invX)
  const pts = (Object.keys(AIRPLANE_TRIS) as AirplanePanel[]).flatMap(k => AIRPLANE_TRIS[k].map(([u, v]) => ndc(panels[k], u, v).slice(0, 2)))
  return hullOf(pts)
}
export function hullOf(points: number[][]): { hull: number[][]; area: number } {
  const ps = [...points].sort((p, q) => p[0] - q[0] || p[1] - q[1])
  const cross = (o: number[], p: number[], q: number[]) => (p[0] - o[0]) * (q[1] - o[1]) - (p[1] - o[1]) * (q[0] - o[0])
  const half = (list: number[][]) => { const h: number[][] = []; for (const q of list) { while (h.length >= 2 && cross(h[h.length - 2], h[h.length - 1], q) <= 1e-12) h.pop(); h.push(q) } h.pop(); return h }
  const hull = [...half(ps), ...half([...ps].reverse())]
  let area = 0
  for (let i = 0; i < hull.length; i++) { const [x0, y0] = hull[i], [x1, y1] = hull[(i + 1) % hull.length]; area += x0 * y1 - x1 * y0 }
  return { hull, area: area / 2 }
}
export function origami(p: number, invX = false) {
  const th = 0.9 * (Math.PI / 2) * ease(p), w = a / 4, sx = invX ? -1 : 1
  const panels: Mat4[] = []
  let px = -sx * a / 2, pz = 0
  for (let k = 0; k < 4; k++) {
    const alpha = (k % 2 === 0 ? 1 : -1) * th
    // Panel k's own left edge (in its rest frame) is at x = -a/2 + k w; move it to the hinge, then turn it.
    const restLeft = invX ? a / 2 - k * w : -a / 2 + k * w
    panels.push(m(T(px, 0, pz), Ry(sx * alpha), T(-restLeft, 0, 0)))
    px += sx * w * Math.cos(alpha); pz += -w * Math.sin(alpha)
  }
  return { panels, opacity: 1 - ease((p - 0.7) / 0.3) }
}
/** curtains: the vertex x of a half (left: uLocal from the outer edge) at progress p. */
export function curtainX(p: number, u: number, v: number) {
  const left = u <= 0.5, outer = left ? -a / 2 : a / 2
  const x = (u - 0.5) * a
  const gather = 0.85 * ease(p) * (0.75 + 0.25 * v)
  const slide = (left ? -1 : 1) * 0.2 * a * ease((p - 0.6) / 0.4)
  return outer + (x - outer) * (1 - gather) + slide
}
export function drapeY(p: number, v: number) {
  const y = 0.5 - v
  return y - 1.4 * ease(p) * (0.8 + 0.2 * (0.5 - y))
}
export function fallOver(p: number, invX = false) {
  const th = (Math.PI / 2) * ease(p), sg = invX ? -1 : 1
  return { from: m(T(0, -0.5, 0), Rz(sg * 0.08 * Math.sin(Math.PI * p)), Rx(-th), T(0, 0.5, 0)), opacity: 1 - ease((p - 0.8) / 0.2) }
}
export function prestige(p: number) {
  const e = ease(p)
  return { from: m(T(0, 0, -1.2 * e), T(0, -0.5, 0), Rx(-0.6 * e), T(0, 0.5, 0)), fromOpacity: 1 - e, to: T(0, -1.1 * (1 - e), 0) }
}
export function warpK(p: number, rho: number) {
  return Math.pow(1 - ease(p), 1 + 2 * rho)
}

export const ROWS: Row[] = [
  { id: 'transition:p14:prism', task: 'S3-2', attrs: 'dir l', oracle: p => ({ 'from uv(1,0) NDC': xy(ndc(prism(p).from, 1, 0)), 'to uv(1,1) NDC': xy(ndc(prism(p).to, 1, 1)), 'isInverted=1 from uv(1,0) NDC': xy(ndc(prism(p, 'l', true).from, 1, 0)), 'dir u from uv(0,0) NDC': xy(ndc(prism(p, 'u').from, 0, 0)) }) },
  { id: 'transition:p14:flip', task: 'S3-2', attrs: 'dir l', oracle: p => ({ 'from uv(0,0) NDC': xy(ndc(flip(p).from, 0, 0)), 'from uv(0,0) world': xyz(world(flip(p).from, 0, 0)) }) },
  { id: 'transition:p14:gallery', task: 'S3-2', attrs: 'dir l', oracle: p => ({ 'from uv(1,0.5) NDC': xy(ndc(gallery(p).from, 1, 0.5)), 'to uv(0,0.5) NDC': xy(ndc(gallery(p).to, 0, 0.5)) }) },
  { id: 'transition:p14:doors', task: 'S3-2', attrs: 'dir vert', oracle: p => ({ 'left half uv(0.5,0) NDC': xy(ndc(doors(p).first, 0.5, 0)), 'to uv(0,0) NDC': xy(ndc(doors(p).to, 0, 0)), 'horz top half uv(0,0.5) NDC': xy(ndc(doors(p, 'horz').first, 0, 0.5)) }) },
  { id: 'transition:p14:window', task: 'S3-2', attrs: 'dir vert', oracle: p => ({ 'left half uv(0.5,0) NDC': xy(ndc(doors(p, 'vert', true).first, 0.5, 0)), 'to uv(0,0) NDC': xy(ndc(doors(p, 'vert', true).to, 0, 0)) }) },
  { id: 'transition:p14:switch', task: 'S3-2', attrs: 'dir l', oracle: p => ({ 'from uv(0.5,0.5) NDC': xy(ndc(switchRow(p).from, 0.5, 0.5)), 'to uv(0.5,0.5) NDC': xy(ndc(switchRow(p).to, 0.5, 0.5)), 'from uv(1,0) NDC': xy(ndc(switchRow(p).from, 1, 0)) }) },
  { id: 'transition:p14:conveyor', task: 'S3-2', attrs: 'dir l', oracle: p => ({ 'from uv(1,0.5) NDC': xy(ndc(conveyor(p).from, 1, 0.5)), 'to uv(0,0.5) NDC': xy(ndc(conveyor(p).to, 0, 0.5)) }) },
  { id: 'transition:p14:ferris', task: 'S3-2', attrs: 'dir l', oracle: p => ({ 'from uv(0.5,0.5) NDC': xy(ndc(ferris(p).from, 0.5, 0.5)), 'to uv(0.5,0.5) NDC': xy(ndc(ferris(p).to, 0.5, 0.5)) }) },
  { id: 'transition:p14:vortex', task: 'S3-3', attrs: 'dir l', oracle: p => ({ ...pref(fractions(VORTEX, SEED, p)), 'cell (0,0) centre NDC': xy(ndc(vortexTile(p, 0, 0).model, 0.5 / 32, 0.5 / 18)), 'cell (0,0) t': r4(vortexTile(p, 0, 0).t) }) },
  { id: 'transition:p14:shred', task: 'S3-3', attrs: 'pattern strip, dir out', oracle: p => ({ ...pref(fractions(SHRED_STRIP, SEED, p)), 'strip 0 t': r4(shredTile(p, 0, 0).t), 'strip 0 dy': r4(shredTile(p, 0, 0).model[13]), ...pref(fractions(SHRED_RECT, SEED, p), 'rectangle ') }) },
  { id: 'transition:p15:prstTrans:fracture', task: 'S3-3', attrs: '—', oracle: p => ({ ...pref(fractions(FRACTURE, SEED, p)), 'cell (0,0) centre NDC': xy(ndc(fractureTile(p, 0, 0).model, 0.5 / 12, 0.5 / 7)) }) },
  { id: 'transition:p15:prstTrans:crush', task: 'S3-3', attrs: '—', oracle: p => ({ ...pref(fractions(CRUSH, SEED, p)), 'cell (0,0) centre NDC': xy(ndc(crushTile(p, 0, 0).model, 0.5 / 16, 0.5 / 9)), 'opacity': r4(crushTile(p, 0, 0).opacity) }) },
  { id: 'transition:p14:glitter', task: 'S3-3', attrs: 'pattern diamond, dir l', oracle: p => ({ ...pref(fractions(GLITTER('diamond', 'l'), SEED, p)), 'cells showing to': r4(countTo(p)), 'cell (0,0) t': r4(glitterTile(p, 0, 0).t) }) },
  { id: 'transition:p14:honeycomb', task: 'S3-3', attrs: '—', oracle: p => ({ ...pref(fractions(HONEYCOMB, SEED, p)), 'cell (10,5) t': r4(honeycombTile(p, 10, 5).t), 'cell (0,0) t': r4(honeycombTile(p, 0, 0).t) }) },
  { id: 'transition:p14:ripple', task: 'S3-3', attrs: 'dir center', oracle: p => ({ 'z at uv(0.75,0.5)': r4(rippleZ(p, 0.75, 0.5)), 'z at uv(0.5,0.5)': r4(rippleZ(p, 0.5, 0.5)), 'from opacity': r4(1 - ease(p)), 'dir lu z at uv(0.25,0.25)': r4(rippleZ(p, 0.25, 0.25, 'lu')) }) },
  { id: 'transition:p15:prstTrans:pageCurlSingle', task: 'S3-4', attrs: '—', oracle: p => pref(foldFractions(p, curlSingle()), 'fraction ') },
  { id: 'transition:p15:prstTrans:pageCurlDouble', task: 'S3-4', attrs: '—', oracle: p => pref(foldFractions(p, curlDouble()), 'fraction ') },
  { id: 'transition:p15:prstTrans:peelOff', task: 'S3-4', attrs: '—', oracle: p => pref(foldFractions(p, peel()), 'fraction ') },
  { id: 'transition:p15:prstTrans:airplane', task: 'S3-4', attrs: '—', oracle: p => ({ 'nose uv(0.5,1) NDC': xy(ndc(airplane(p).panels.bodyL, 0.5, 1)), 'left wing uv(0,1) world': xyz(world(airplane(p).panels.wingL, 0, 1)), 'silhouette area (NDC)': r4(airplaneSilhouette(p).area), 'p 0.75 nose NDC': xy(ndc(airplane(0.75).panels.bodyL, 0.5, 1)), 'p 0.75 silhouette area (NDC)': r4(airplaneSilhouette(0.75).area) }) },
  { id: 'transition:p15:prstTrans:origami', task: 'S3-4', attrs: '—', oracle: p => ({ 'panel 2 uv(0.5,0) world': xyz(world(origami(p).panels[2], 0.5, 0)), 'panel 3 uv(1,0) NDC': xy(ndc(origami(p).panels[3], 1, 0)) }) },
  { id: 'transition:p15:prstTrans:curtains', task: 'S3-4', attrs: '—', oracle: p => ({ 'x at uv(0.5,0.5) (left half)': r4(curtainX(p, 0.5, 0.5)), 'x at uv(0.25,1)': r4(curtainX(p, 0.25, 1)), 'p 1 x at uv(0.5,0)': r4(curtainX(1, 0.5, 0)) }) },
  { id: 'transition:p15:prstTrans:drape', task: 'S3-4', attrs: '—', oracle: p => ({ 'y at uv(0.5,0)': r4(drapeY(p, 0)), 'y at uv(0.5,1)': r4(drapeY(p, 1)) }) },
  { id: 'transition:p15:prstTrans:wind', task: 'S3-4', attrs: '—', oracle: p => ({ ...pref(fractions(WIND, SEED, p)), 'cell (7,0) t': r4(windTile(p, 7, 0).t), 'cell (7,0) centre NDC': xy(ndc(windTile(p, 7, 0).model, 7.5 / 8, 0.5 / 18)) }) },
  { id: 'transition:p15:prstTrans:fallOver', task: 'S3-4', attrs: '—', oracle: p => ({ 'from uv(0.5,0) NDC': xy(ndc(fallOver(p).from, 0.5, 0)), 'from uv(0.5,0) world': xyz(world(fallOver(p).from, 0.5, 0)) }) },
  { id: 'transition:p15:prstTrans:prestige', task: 'S3-4', attrs: '—', oracle: p => ({ 'from uv(0.5,0) NDC': xy(ndc(prestige(p).from, 0.5, 0)), 'to uv(0.5,0) NDC': xy(ndc(prestige(p).to, 0.5, 0)), 'from opacity': r4(prestige(p).fromOpacity) }) },
  { id: 'transition:p14:warp', task: 'S3-4', attrs: 'dir in', oracle: p => ({ 'k at uv(1,0.5)': r4(warpK(p, a / 2)), 'world x at uv(1,0.5)': r4((a / 2) * warpK(p, a / 2)), 'k at uv(0.5,0.5)': r4(warpK(p, 0)) }) },
]
function pref(o: Record<string, number>, prefix = ''): Record<string, number> {
  const out: Record<string, number> = {}
  for (const [k, v] of Object.entries(o)) out[`${prefix}${k}`] = v
  return out
}
function countTo(p: number) {
  const spec = GLITTER('diamond', 'l')
  let n = 0
  for (let j = 0; j < spec.rows; j++) for (let i = 0; i < spec.cols; i++) if (glitterTile(p, i, j).showsTo) n++
  return n / (spec.cols * spec.rows)
}

if (import.meta.main) {
  const json = process.argv.includes('--json')
  const table = ROWS.map(r => ({ id: r.id, task: r.task, attrs: r.attrs, oracle: r.oracle(0.5) }))
  if (json) console.log(JSON.stringify(table, null, 2))
  else {
    console.log('| Row | Task | Attributes | Oracle at 50 % |')
    console.log('| --- | --- | --- | --- |')
    for (const r of table) console.log(`| ${r.id} | ${r.task} | ${r.attrs} | ${Object.entries(r.oracle).map(([k, v]) => `${k} = ${v}`).join('; ')} |`)
  }
}
