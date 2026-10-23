/**
 * The deformation 3-D transitions (slideshow spec 6.3, S3-4): pageCurlSingle,
 * pageCurlDouble, peelOff, airplane, origami, curtains, drape, wind,
 * fallOver, prestige (all `p15:prstTrans` presets) and `p14:warp`.
 *
 * Every row follows its entry in docs/validation/slideshow-3d-algorithms.md
 * (the reference script `scripts/slideshow-3d-oracles.ts` wins over the
 * prose) and is a closest look-alike, never Office-matched. drape, wind,
 * fallOver and warp also carry an unverified D4 classification in the
 * catalogue: their models are kept at parity with the reference script until
 * Office captures exist.
 *
 * Shapes of the work:
 * - **Fold rows** (the page curls, peelOff): one 1 × 1 grid and a fragment
 *   shader with both frames bound, evaluating the sheet's fold template per
 *   pixel.
 * - **Rigid rows** (airplane, origami, fallOver, prestige): CPU matrices
 *   (`airplaneModel` …); airplane's four triangles use a triangle program,
 *   the others the standard program with pieces cut out by
 *   `uvRect`, back to front.
 * - **Deformed rows** (curtains, drape, warp): a 32 × 18 grid moved per
 *   vertex by a custom vertex shader mirroring `curtainVertex`,
 *   `drapeVertex`, `warpK`.
 * - **wind**: the tile template on an 8 × 18 grid in one indexed draw, keyed
 *   on `a_cell`, seeded from the shader rng's first value (`effectSeed`).
 *
 * Each row leases and compiles everything it draws in `prepare` (so a refusal
 * sends the whole effect to its look-alike), asks the pool again every frame,
 * and draws O(1) indexed draws. Nothing here is registered on import:
 * `registerDeform3D()` registers the group (S3-7 wires it).
 */
import type { Rng } from '../../types'
import {
  DEFAULT_FOV_DEG, project, FRAGMENT_3D, GRID_ATTRIBUTES_GLSL, HASH_GLSL, SLIDE_GLSL, STANDARD_PROGRAM, VERTEX_3D,
  cellRandom, ease, effectSeed, mul4, rotationX, rotationY, rotationZ, scaling, slideCamera, slidePoint, translation,
  type Mat4, type MeshPool, type Vec3,
} from './mesh'
import { registerTransition3D, type Transition3D } from './registry'

type Attrs = Readonly<Record<string, string>>
const flag = (attrs: Attrs, name: string) => attrs[name] === '1' || attrs[name] === 'true'
const m = (...ms: Mat4[]): Mat4 => ms.reduce((acc, x) => mul4(acc, x))
const clamp01 = (x: number) => Math.min(1, Math.max(0, x))
/** Deformed-row grid: fine enough for smooth folds and ripples, bounded (32 × 18 × 76 B ≈ 43 KiB). */
const DEFORM_GRID_COLS = 32, DEFORM_GRID_ROWS = 18

// ---------------------------------------------------------------------------
// Pure models (the oracles' own formulas)
// ---------------------------------------------------------------------------

/** A fold-template page: x range, start corner and target corner, in slide units (x right in [0, a], y down in [0, 1]). */
export interface FoldPage { x0: number; x1: number; corner: [number, number]; target: [number, number] }

/** The pages of a fold row (`pageCurlSingle`, `pageCurlDouble`, `peelOff`) for its attributes. */
export function foldPages(preset: string, attrs: Attrs, aspect: number): FoldPage[] {
  const a = aspect
  const invX = flag(attrs, 'invX'), invY = flag(attrs, 'invY')
  if (preset === 'pageCurlDouble') {
    return [
      { x0: 0, x1: a / 2, corner: [0, invY ? 0 : 1], target: [a / 2, invY ? 1 : 0] },
      { x0: a / 2, x1: a, corner: [a, invY ? 0 : 1], target: [a / 2, invY ? 1 : 0] },
    ]
  }
  if (preset === 'peelOff') return [{ x0: 0, x1: a, corner: [invX ? a : 0, 1], target: [invX ? 0 : a, 0] }]
  return [{ x0: 0, x1: a, corner: [invX ? 0 : a, invY ? 0 : 1], target: [invX ? a : 0, invY ? 1 : 0] }]
}

/** The fold rule at pixel (x, y) (slide units): what the fragment shader shows there. */
export function foldClassify(p: number, x: number, y: number, pages: readonly FoldPage[]): 'to' | 'back' | 'from' {
  const page = pages.length > 1 && x > pages[0].x1 ? pages[1] : pages[0]
  if (x < page.x0 || x > page.x1) return 'from'
  const [cx, cy] = page.corner, [tx, ty] = page.target
  const D = Math.hypot(tx - cx, ty - cy), nx = (tx - cx) / D, ny = (ty - cy) / D
  const L = D * ease(p)
  const s = (x - cx) * nx + (y - cy) * ny
  if (s < L) return 'to'
  const mx = x - 2 * (s - L) * nx, my = y - 2 * (s - L) * ny
  return mx >= page.x0 && mx <= page.x1 && my >= 0 && my <= 1 ? 'back' : 'from'
}

/** Right-handed rotation by `th` about the axis through world points `from` → `to`. */
export function axisRotation(from: Vec3, to: Vec3, th: number): Mat4 {
  const d = [to[0] - from[0], to[1] - from[1], to[2] - from[2]], n = Math.hypot(d[0], d[1], d[2])
  const [x, y, z] = d.map(v => v / n), c = Math.cos(th), s = Math.sin(th), t = 1 - c
  const r = new Float32Array([
    t * x * x + c, t * x * y + s * z, t * x * z - s * y, 0,
    t * x * y - s * z, t * y * y + c, t * y * z + s * x, 0,
    t * x * z + s * y, t * y * z - s * x, t * z * z + c, 0,
    0, 0, 0, 1,
  ])
  return m(translation(from[0], from[1], from[2]), r, translation(-from[0], -from[1], -from[2]))
}

/** airplane's four rigid triangles, in slide uv: the body halves either side of the spine, and the wings under the diagonal creases from the top corners to the nose (bottom centre). */
export const AIRPLANE_TRIS = Object.freeze({
  bodyL: [[0, 0], [0.5, 0], [0.5, 1]], bodyR: [[0.5, 0], [1, 0], [0.5, 1]],
  wingL: [[0, 0], [0, 1], [0.5, 1]], wingR: [[1, 0], [1, 1], [0.5, 1]],
} as const)
export type AirplanePanel = keyof typeof AIRPLANE_TRIS
const AIRPLANE_PANELS = Object.keys(AIRPLANE_TRIS) as AirplanePanel[]

/**
 * airplane: a paper dart. Over the first half the body halves open a dihedral
 * about the spine and the wings fold over them about the diagonal creases
 * (nose at the bottom centre); over the second half the folded dart banks,
 * pitches its nose into depth and flies away, never edge-on.
 */
export function airplaneModel(p: number, aspect: number, invX: boolean): { panels: Record<AirplanePanel, Mat4>; opacity: number; wingShade: number } {
  const f = ease(2 * p), tau = ease(2 * p - 1), sx = invX ? -1 : 1
  const beta = (Math.PI / 4) * f, th = 0.9 * Math.PI * f
  const N = slidePoint(0.5, 1, aspect), TL = slidePoint(0, 0, aspect), TR = slidePoint(1, 0, aspect)
  const bodyL = rotationY(-beta), bodyR = rotationY(beta)
  const wingL = mul4(bodyL, axisRotation(N, TL, th)), wingR = mul4(bodyR, axisRotation(N, TR, -th))
  const k = 1 - 0.4 * tau
  const G = m(translation(sx * 0.9 * aspect * tau, -2.4 * tau, -3 * tau), rotationZ(sx * 0.3 * tau), rotationY(sx * 0.5 * tau), rotationX(0.6 * tau), scaling(k, k, k))
  return { panels: { bodyL: mul4(G, bodyL), bodyR: mul4(G, bodyR), wingL: mul4(G, wingL), wingR: mul4(G, wingR) }, opacity: 1 - ease((p - 0.85) / 0.15), wingShade: 1 - 0.1 * f }
}

/** The dart's projected outline at `p` under the slide camera: the convex hull (NDC, counter-clockwise) of its panel corners, and its area. */
export function airplaneSilhouette(p: number, aspect: number, invX: boolean): { hull: number[][]; area: number } {
  const { panels } = airplaneModel(p, aspect, invX)
  const C = slideCamera(DEFAULT_FOV_DEG, aspect)
  const pts = AIRPLANE_PANELS.flatMap(k => AIRPLANE_TRIS[k].map(([u, v]) => project(mul4(C, panels[k]), slidePoint(u, v, aspect)).slice(0, 2)))
  pts.sort((a, b) => a[0] - b[0] || a[1] - b[1])
  const cross = (o: number[], a: number[], b: number[]) => (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])
  const half = (list: number[][]) => {
    const h: number[][] = []
    for (const q of list) { while (h.length >= 2 && cross(h[h.length - 2], h[h.length - 1], q) <= 1e-12) h.pop(); h.push(q) }
    h.pop()
    return h
  }
  const hull = [...half(pts), ...half([...pts].reverse())]
  let area = 0
  for (let i = 0; i < hull.length; i++) { const [x0, y0] = hull[i], [x1, y1] = hull[(i + 1) % hull.length]; area += x0 * y1 - x1 * y0 }
  return { hull, area: area / 2 }
}

/** origami: four vertical panels as an accordion hinged in a chain from the left edge (`invX`: the right edge). Panel k is slide uv x ∈ [k/4, (k+1)/4] (`invX`: counted from the right). */
export function origamiModel(p: number, aspect: number, invX: boolean): { panels: Mat4[]; opacity: number } {
  const th = 0.9 * (Math.PI / 2) * ease(p), w = aspect / 4, sx = invX ? -1 : 1
  const panels: Mat4[] = []
  let px = (-sx * aspect) / 2, pz = 0
  for (let k = 0; k < 4; k++) {
    const alpha = (k % 2 === 0 ? 1 : -1) * th
    const restHinge = invX ? aspect / 2 - k * w : -aspect / 2 + k * w
    panels.push(m(translation(px, 0, pz), rotationY(sx * alpha), translation(-restHinge, 0, 0)))
    px += sx * w * Math.cos(alpha)
    pz += -w * Math.sin(alpha)
  }
  return { panels, opacity: 1 - ease((p - 0.7) / 0.3) }
}

/** curtains: where slide uv (u, v) is (world), each half gathering to its outer edge and sliding off. */
export function curtainVertex(p: number, u: number, v: number, aspect: number): Vec3 {
  const left = u <= 0.5, side = left ? -1 : 1, outer = (side * aspect) / 2
  const e = ease(p)
  const x = (u - 0.5) * aspect
  const gather = 0.85 * e * (0.75 + 0.25 * v)
  const slide = side * 0.2 * aspect * ease((p - 0.6) / 0.4)
  const ul = left ? u / 0.5 : (1 - u) / 0.5
  return [outer + (x - outer) * (1 - gather) + slide, 0.5 - v, 0.035 * e * Math.sin(16 * Math.PI * ul)]
}

/** drape: the slide falls like cloth, the bottom edge faster, with a ripple and a sway. */
export function drapeVertex(p: number, u: number, v: number, aspect: number): Vec3 {
  const e = ease(p), s = Math.sin(Math.PI * p), y = 0.5 - v
  return [
    (u - 0.5) * aspect + 0.03 * s * Math.sin(2 * Math.PI * v),
    y - 1.4 * e * (0.8 + 0.2 * (0.5 - y)),
    0.08 * s * Math.sin(3 * Math.PI * u + 2 * Math.PI * p) * (0.5 - y),
  ]
}

/** warp: the pull towards the centre of a vertex at world distance `rho` from it. */
export const warpK = (p: number, rho: number): number => Math.pow(1 - ease(p), 1 + 2 * rho)

/** wind's tile grid and template constants (the sheet's 8 × 18, W = 0.55, J = 0.5). */
export const WIND = Object.freeze({ cols: 8, rows: 18, W: 0.55, J: 0.5 })

/** wind: cell (i, j)'s local time, matrix and opacity. `invX` is the exact mirror image (cell i mirrors cell cols − 1 − i). */
export function windTile(p: number, i: number, j: number, seed: number, aspect: number, invX: boolean): { t: number; model: Mat4; opacity: number } {
  const im = invX ? WIND.cols - 1 - i : i
  const uc = (im + 0.5) / WIND.cols, vc = (j + 0.5) / WIND.rows
  const g = 0.6 * (1 - uc) + 0.4 * vc
  const delay = WIND.W * ((1 - WIND.J) * g + WIND.J * cellRandom(seed, im, j, 0))
  const t = clamp01((p - delay) / (1 - WIND.W))
  const c = slidePoint(uc, vc, aspect)
  let model = m(
    translation(c[0] - 1.6 * aspect * t * t, c[1] + 0.25 * Math.sin(Math.PI * t) * (cellRandom(seed, im, j, 1) - 0.5), 0),
    rotationZ(0.6 * (cellRandom(seed, im, j, 2) - 0.5) * t), rotationY((-Math.PI / 3) * t), translation(-c[0], -c[1], 0),
  )
  if (invX) model = m(scaling(-1, 1, 1), model, scaling(-1, 1, 1))
  return { t, model, opacity: 1 - ease((t - 0.7) / 0.3) }
}

/** fallOver: the slide falls backwards about its bottom edge, leaning a little. */
export function fallOverModel(p: number, invX: boolean): { from: Mat4; opacity: number } {
  const th = (Math.PI / 2) * ease(p), sg = invX ? -1 : 1
  return { from: m(translation(0, -0.5, 0), rotationZ(sg * 0.08 * Math.sin(Math.PI * p)), rotationX(-th), translation(0, 0.5, 0)), opacity: 1 - ease((p - 0.8) / 0.2) }
}

/** prestige: the outgoing slide tips back and recedes, fading; the incoming one rises from below over it. */
export function prestigeModel(p: number): { from: Mat4; fromOpacity: number; to: Mat4 } {
  const e = ease(p)
  return { from: m(translation(0, 0, -1.2 * e), translation(0, -0.5, 0), rotationX(-0.6 * e), translation(0, 0.5, 0)), fromOpacity: 1 - e, to: translation(0, -1.1 * (1 - e), 0) }
}

// ---------------------------------------------------------------------------
// Shaders
// ---------------------------------------------------------------------------

const FOLD_FS = `#version 300 es
precision highp float;
in vec2 v_uv;
uniform sampler2D u_from;
uniform sampler2D u_to;
uniform float u_aspect;
uniform float u_e;
uniform int u_peel;
uniform int u_pages;
uniform vec4 u_range[2];
uniform vec4 u_line[2];
out vec4 outColor;
void main() {
  vec2 q = vec2(v_uv.x * u_aspect, v_uv.y);
  int k = (u_pages > 1 && q.x > u_range[0].y) ? 1 : 0;
  vec2 xr = u_range[k].xy;
  vec2 c = u_line[k].xy;
  vec2 d = u_line[k].zw - c;
  float D = length(d);
  vec2 n = d / D;
  float L = D * u_e;
  float s = dot(q - c, n);
  float crease = 1.0 - 0.3 * exp(-abs(s - L) / 0.05);
  if (s < L) {
    vec4 t = texture(u_to, v_uv);
    if (u_peel == 1 && s >= L - 0.04) t.rgb *= 0.6;
    outColor = t;
    return;
  }
  vec2 mq = q - 2.0 * (s - L) * n;
  if (mq.x >= xr.x && mq.x <= xr.y && mq.y >= 0.0 && mq.y <= 1.0) {
    vec4 b = u_peel == 1 ? vec4(0.92, 0.92, 0.92, 1.0) : mix(texture(u_from, vec2(mq.x / u_aspect, mq.y)), vec4(1.0), 0.5);
    outColor = vec4(b.rgb * crease, b.a);
    return;
  }
  vec4 f = texture(u_from, v_uv);
  outColor = vec4(f.rgb * crease, f.a);
}
`

/** The deformed rows' fragment shader: the standard one plus a per-vertex shade and opacity. */
const DEFORM_FS = `#version 300 es
precision highp float;
in vec2 v_uv;
in float v_shade;
in float v_alpha;
uniform sampler2D u_tex;
uniform float u_opacity;
out vec4 outColor;
void main() {
  vec4 c = texture(u_tex, v_uv);
  outColor = vec4(c.rgb * v_shade, c.a) * (u_opacity * v_alpha);
}
`

const DEFORM_HEAD = `#version 300 es
precision highp float;
${GRID_ATTRIBUTES_GLSL}uniform mat4 u_mvp;
uniform float u_aspect;
uniform float u_p;
uniform vec4 u_uvRect;
out vec2 v_uv;
out float v_shade;
out float v_alpha;
${SLIDE_GLSL}${HASH_GLSL}const float PI = 3.14159265358979;
`

/** `curtainVertex`, per half (`u_side` −1 left, +1 right). Folds shade their far-facing slopes. */
const CURTAINS_VS = `${DEFORM_HEAD}uniform float u_side;
void main() {
  vec2 uv = u_uvRect.xy + a_uv * u_uvRect.zw;
  v_uv = uv;
  float e = ease(u_p);
  float outer = u_side * 0.5 * u_aspect;
  float x = (uv.x - 0.5) * u_aspect;
  float gather = 0.85 * e * (0.75 + 0.25 * uv.y);
  float slide = u_side * 0.2 * u_aspect * ease((u_p - 0.6) / 0.4);
  float ul = u_side < 0.0 ? uv.x / 0.5 : (1.0 - uv.x) / 0.5;
  vec3 q = vec3(outer + (x - outer) * (1.0 - gather) + slide, 0.5 - uv.y, 0.035 * e * sin(16.0 * PI * ul));
  v_shade = 1.0 - 0.25 * e * (0.5 - 0.5 * cos(16.0 * PI * ul));
  v_alpha = 1.0;
  gl_Position = u_mvp * vec4(q, 1.0);
}
`

/** `drapeVertex`. */
const DRAPE_VS = `${DEFORM_HEAD}void main() {
  vec2 uv = a_uv;
  v_uv = uv;
  float e = ease(u_p);
  float s = sin(PI * u_p);
  float y = 0.5 - uv.y;
  vec3 q = vec3((uv.x - 0.5) * u_aspect + 0.03 * s * sin(2.0 * PI * uv.y), y - 1.4 * e * (0.8 + 0.2 * (0.5 - y)), 0.08 * s * sin(3.0 * PI * uv.x + 2.0 * PI * u_p) * (0.5 - y));
  v_shade = 1.0;
  v_alpha = 1.0;
  gl_Position = u_mvp * vec4(q, 1.0);
}
`

/** warp: q' = q · (1 − e(u_kp))^(1 + 2|q|); `u_kp` is p (`in`) or 1 − p (`out`). */
const WARP_VS = `${DEFORM_HEAD}uniform float u_kp;
void main() {
  vec2 uv = a_uv;
  v_uv = uv;
  vec3 q = slidePos(uv, u_aspect);
  float k = pow(max(1.0 - ease(u_kp), 0.0), 1.0 + 2.0 * length(q.xy));
  v_shade = 1.0;
  v_alpha = 1.0;
  gl_Position = u_mvp * vec4(q.xy * k, 0.0, 1.0);
}
`

/** wind: `windTile` per vertex, keyed on its cell (`u_grid` = cols, rows, W, J; `u_invX` mirrors). */
const WIND_VS = `${DEFORM_HEAD}uniform uint u_seed;
uniform vec4 u_grid;
uniform int u_invX;
void main() {
  v_uv = a_uv;
  float mx = u_invX == 1 ? -1.0 : 1.0;
  vec2 uv = u_invX == 1 ? vec2(1.0 - a_uv.x, a_uv.y) : a_uv;
  vec2 cell = u_invX == 1 ? vec2(1.0 - a_cell.x, a_cell.y) : a_cell;
  uint i = uint(floor(cell.x * u_grid.x));
  uint j = uint(floor(cell.y * u_grid.y));
  float g = 0.6 * (1.0 - cell.x) + 0.4 * cell.y;
  float delay = u_grid.z * ((1.0 - u_grid.w) * g + u_grid.w * cellRandom(u_seed, i, j, 0u));
  float t = clamp((u_p - delay) / (1.0 - u_grid.z), 0.0, 1.0);
  vec3 c = slidePos(cell, u_aspect);
  vec3 l = slidePos(uv, u_aspect) - c;
  float ay = -PI / 3.0 * t;
  l = vec3(cos(ay) * l.x + sin(ay) * l.z, l.y, -sin(ay) * l.x + cos(ay) * l.z);
  float az = 0.6 * (cellRandom(u_seed, i, j, 2u) - 0.5) * t;
  l = vec3(cos(az) * l.x - sin(az) * l.y, sin(az) * l.x + cos(az) * l.y, l.z);
  vec3 q = l + c + vec3(-1.6 * u_aspect * t * t, 0.25 * sin(PI * t) * (cellRandom(u_seed, i, j, 1u) - 0.5), 0.0);
  q.x *= mx;
  v_shade = 1.0 - 0.3 * t;
  v_alpha = 1.0 - ease((t - 0.7) / 0.3);
  gl_Position = u_mvp * vec4(q, 1.0);
}
`

/** airplane's panels: a 1 × 1 grid whose corners map to triangle `u_tri` (slide uv; the fourth corner collapses onto the third). */
const TRIANGLE_VS = `#version 300 es
precision highp float;
${GRID_ATTRIBUTES_GLSL}uniform mat4 u_mvp;
uniform float u_aspect;
uniform vec2 u_tri[3];
out vec2 v_uv;
${SLIDE_GLSL}void main() {
  vec2 uv = (1.0 - a_uv.y) * mix(u_tri[0], u_tri[1], a_uv.x) + a_uv.y * u_tri[2];
  v_uv = uv;
  gl_Position = u_mvp * vec4(slidePos(uv, u_aspect), 1.0);
}
`

const PROGRAMS = {
  fold: ['three:deform:fold', VERTEX_3D, FOLD_FS],
  curtains: ['three:deform:curtains', CURTAINS_VS, DEFORM_FS],
  drape: ['three:deform:drape', DRAPE_VS, DEFORM_FS],
  warp: ['three:deform:warp', WARP_VS, DEFORM_FS],
  wind: ['three:deform:wind', WIND_VS, DEFORM_FS],
  standard: [STANDARD_PROGRAM, VERTEX_3D, FRAGMENT_3D],
  triangle: ['three:deform:triangle', TRIANGLE_VS, FRAGMENT_3D],
} as const satisfies Record<string, readonly [string, string, string]>
type ProgramName = keyof typeof PROGRAMS
const programOf = (mesh: MeshPool, name: ProgramName) => { const [key, vs, fs] = PROGRAMS[name]; return mesh.program(key, vs, fs) }

// ---------------------------------------------------------------------------
// Drawing
// ---------------------------------------------------------------------------

const camera = (mesh: MeshPool) => slideCamera(DEFAULT_FOV_DEG, mesh.aspect)
function premultipliedBlend(gl: WebGL2RenderingContext) {
  gl.enable(gl.BLEND)
  gl.blendEquation(gl.FUNC_ADD)
  gl.blendFunc(gl.ONE, gl.ONE_MINUS_SRC_ALPHA)
  gl.disable(gl.CULL_FACE)
}
/** Rigid pieces, farthest (most negative world z of its centre) first. */
interface Piece { model: Mat4; tex: WebGLTexture; uvRect: readonly [number, number, number, number]; opacity: number }
function drawPieces(gl: WebGL2RenderingContext, mesh: MeshPool, pieces: Piece[]) {
  const C = camera(mesh), quad = mesh.grid(1, 1), a = mesh.aspect
  const depth = ({ model, uvRect: r }: Piece) => {
    const c = slidePoint(r[0] + r[2] / 2, r[1] + r[3] / 2, a)
    return model[2] * c[0] + model[6] * c[1] + model[10] * c[2] + model[14]
  }
  const order = pieces.map((piece, i) => ({ piece, i, z: depth(piece) })).sort((x, y) => x.z - y.z || x.i - y.i)
  premultipliedBlend(gl)
  for (const { piece } of order) {
    if (piece.opacity <= 0) continue
    mesh.draw(quad, piece.tex, mul4(C, piece.model), { opacity: piece.opacity, uvRect: piece.uvRect })
  }
}
const FULL = [0, 0, 1, 1] as const
const IDENTITY = translation(0, 0, 0)

/** Common uniforms of the deformed programs, with `tex` on unit 0. */
function useDeform(gl: WebGL2RenderingContext, mesh: MeshPool, name: ProgramName, tex: WebGLTexture, progress: number, mvp: Mat4, opacity = 1) {
  const p = programOf(mesh, name)
  gl.useProgram(p.program)
  gl.activeTexture(gl.TEXTURE0)
  gl.bindTexture(gl.TEXTURE_2D, tex)
  gl.uniform1i(p.uniform('u_tex'), 0)
  gl.uniformMatrix4fv(p.uniform('u_mvp'), false, mvp)
  gl.uniform1f(p.uniform('u_aspect'), mesh.aspect)
  gl.uniform1f(p.uniform('u_p'), progress)
  gl.uniform1f(p.uniform('u_opacity'), opacity)
  return p
}

function drawFold(preset: string) {
  return (gl: WebGL2RenderingContext, from: WebGLTexture, to: WebGLTexture, progress: number, attrs: Attrs, _rng: Rng, mesh: MeshPool) => {
    const pages = foldPages(preset, attrs, mesh.aspect)
    const quad = mesh.grid(1, 1)
    const p = programOf(mesh, 'fold')
    premultipliedBlend(gl)
    gl.useProgram(p.program)
    gl.activeTexture(gl.TEXTURE1)
    gl.bindTexture(gl.TEXTURE_2D, to)
    gl.activeTexture(gl.TEXTURE0)
    gl.bindTexture(gl.TEXTURE_2D, from)
    gl.uniform1i(p.uniform('u_from'), 0)
    gl.uniform1i(p.uniform('u_to'), 1)
    gl.uniformMatrix4fv(p.uniform('u_mvp'), false, camera(mesh))
    gl.uniform1f(p.uniform('u_aspect'), mesh.aspect)
    gl.uniform4f(p.uniform('u_uvRect'), 0, 0, 1, 1)
    gl.uniform1f(p.uniform('u_e'), ease(progress))
    gl.uniform1i(p.uniform('u_peel'), preset === 'peelOff' ? 1 : 0)
    gl.uniform1i(p.uniform('u_pages'), pages.length)
    const lines = new Float32Array(8), ranges = new Float32Array(8)
    pages.forEach((pg, k) => {
      lines.set([pg.corner[0], pg.corner[1], pg.target[0], pg.target[1]], 4 * k)
      ranges.set([pg.x0, pg.x1, 0, 0], 4 * k)
    })
    gl.uniform4fv(p.uniform('u_line[0]'), lines)
    gl.uniform4fv(p.uniform('u_range[0]'), ranges)
    mesh.drawMesh(quad)
    // Leave unit 1 empty: the compositor binds its own mask texture there only when it uses one.
    gl.activeTexture(gl.TEXTURE1)
    gl.bindTexture(gl.TEXTURE_2D, null)
    gl.activeTexture(gl.TEXTURE0)
  }
}

type Draw = Transition3D['draw']
interface Row { element: string; preset?: string; grids: Array<[number, number]>; programs: ProgramName[]; draw: Draw }

const ROWS: Row[] = [
  { element: 'prstTrans', preset: 'pageCurlSingle', grids: [[1, 1]], programs: ['fold'], draw: drawFold('pageCurlSingle') },
  { element: 'prstTrans', preset: 'pageCurlDouble', grids: [[1, 1]], programs: ['fold'], draw: drawFold('pageCurlDouble') },
  { element: 'prstTrans', preset: 'peelOff', grids: [[1, 1]], programs: ['fold'], draw: drawFold('peelOff') },
  {
    element: 'prstTrans', preset: 'airplane', grids: [[1, 1]], programs: ['standard', 'triangle'],
    draw(gl, from, to, progress, attrs, _rng, mesh) {
      const a = airplaneModel(progress, mesh.aspect, flag(attrs, 'invX'))
      const C = camera(mesh), quad = mesh.grid(1, 1)
      premultipliedBlend(gl)
      mesh.draw(quad, to, C)
      if (a.opacity <= 0) return
      const p = programOf(mesh, 'triangle')
      gl.useProgram(p.program)
      gl.activeTexture(gl.TEXTURE0)
      gl.bindTexture(gl.TEXTURE_2D, from)
      gl.uniform1i(p.uniform('u_tex'), 0)
      gl.uniform1f(p.uniform('u_aspect'), mesh.aspect)
      gl.uniform1f(p.uniform('u_opacity'), a.opacity)
      // The bodies, then the wings folded over them; each pair farthest (centroid world z) first.
      const depth = (k: AirplanePanel) => {
        const c = AIRPLANE_TRIS[k].reduce<[number, number]>((s, [u, v]) => [s[0] + u / 3, s[1] + v / 3], [0, 0])
        const q = slidePoint(c[0], c[1], mesh.aspect), M = a.panels[k]
        return M[2] * q[0] + M[6] * q[1] + M[10] * q[2] + M[14]
      }
      for (const pair of [['bodyL', 'bodyR'], ['wingL', 'wingR']] as const) {
        for (const k of [...pair].sort((x, y) => depth(x) - depth(y))) {
          gl.uniformMatrix4fv(p.uniform('u_mvp'), false, mul4(C, a.panels[k]))
          gl.uniform2fv(p.uniform('u_tri[0]'), new Float32Array(AIRPLANE_TRIS[k].flat()))
          gl.uniform1f(p.uniform('u_shade'), k.startsWith('wing') ? a.wingShade : 1)
          mesh.drawMesh(quad)
        }
      }
    },
  },
  {
    element: 'prstTrans', preset: 'origami', grids: [[1, 1]], programs: ['standard'],
    draw(gl, from, to, progress, attrs, _rng, mesh) {
      const invX = flag(attrs, 'invX')
      const o = origamiModel(progress, mesh.aspect, invX)
      premultipliedBlend(gl)
      mesh.draw(mesh.grid(1, 1), to, camera(mesh))
      drawPieces(gl, mesh, o.panels.map((model, k) => ({ model, tex: from, uvRect: [invX ? (3 - k) / 4 : k / 4, 0, 0.25, 1] as const, opacity: o.opacity })))
    },
  },
  {
    element: 'prstTrans', preset: 'curtains', grids: [[1, 1], [DEFORM_GRID_COLS, DEFORM_GRID_ROWS]], programs: ['standard', 'curtains'],
    draw(gl, from, to, progress, _attrs, _rng, mesh) {
      const C = camera(mesh)
      premultipliedBlend(gl)
      mesh.draw(mesh.grid(1, 1), to, C)
      const grid = mesh.grid(DEFORM_GRID_COLS, DEFORM_GRID_ROWS)
      for (const side of [-1, 1]) {
        const p = useDeform(gl, mesh, 'curtains', from, progress, C)
        gl.uniform4f(p.uniform('u_uvRect'), side < 0 ? 0 : 0.5, 0, 0.5, 1)
        gl.uniform1f(p.uniform('u_side'), side)
        mesh.drawMesh(grid)
      }
    },
  },
  {
    element: 'prstTrans', preset: 'drape', grids: [[1, 1], [DEFORM_GRID_COLS, DEFORM_GRID_ROWS]], programs: ['standard', 'drape'],
    draw(gl, from, to, progress, _attrs, _rng, mesh) {
      const C = camera(mesh)
      premultipliedBlend(gl)
      mesh.draw(mesh.grid(1, 1), to, C)
      const p = useDeform(gl, mesh, 'drape', from, progress, C)
      gl.uniform4f(p.uniform('u_uvRect'), 0, 0, 1, 1)
      mesh.drawMesh(mesh.grid(DEFORM_GRID_COLS, DEFORM_GRID_ROWS))
    },
  },
  {
    element: 'prstTrans', preset: 'wind', grids: [[1, 1], [WIND.cols, WIND.rows]], programs: ['standard', 'wind'],
    draw(gl, from, to, progress, attrs, rng, mesh) {
      const seed = effectSeed(rng)
      const C = camera(mesh)
      premultipliedBlend(gl)
      mesh.draw(mesh.grid(1, 1), to, C)
      const p = useDeform(gl, mesh, 'wind', from, progress, C)
      gl.uniform4f(p.uniform('u_uvRect'), 0, 0, 1, 1)
      gl.uniform1ui(p.uniform('u_seed'), seed)
      gl.uniform4f(p.uniform('u_grid'), WIND.cols, WIND.rows, WIND.W, WIND.J)
      gl.uniform1i(p.uniform('u_invX'), flag(attrs, 'invX') ? 1 : 0)
      mesh.drawMesh(mesh.grid(WIND.cols, WIND.rows))
    },
  },
  {
    element: 'prstTrans', preset: 'fallOver', grids: [[1, 1]], programs: ['standard'],
    draw(gl, from, to, progress, attrs, _rng, mesh) {
      const f = fallOverModel(progress, flag(attrs, 'invX'))
      premultipliedBlend(gl)
      mesh.draw(mesh.grid(1, 1), to, camera(mesh))
      drawPieces(gl, mesh, [{ model: f.from, tex: from, uvRect: FULL, opacity: f.opacity }])
    },
  },
  {
    element: 'prstTrans', preset: 'prestige', grids: [[1, 1]], programs: ['standard'],
    draw(gl, from, to, progress, _attrs, _rng, mesh) {
      const s = prestigeModel(progress)
      const C = camera(mesh), quad = mesh.grid(1, 1)
      premultipliedBlend(gl)
      // from, then to over it (the sheet's order, not depth order: to rises in front).
      if (s.fromOpacity > 0) mesh.draw(quad, from, mul4(C, s.from), { opacity: s.fromOpacity })
      mesh.draw(quad, to, mul4(C, s.to))
    },
  },
  {
    element: 'warp', grids: [[1, 1], [DEFORM_GRID_COLS, DEFORM_GRID_ROWS]], programs: ['standard', 'warp'],
    draw(gl, from, to, progress, attrs, _rng, mesh) {
      // `out` is `in` run backwards with the frames swapped.
      const out = attrs.dir === 'out'
      const kp = out ? 1 - progress : progress
      const warped = out ? to : from, flat = out ? from : to
      const C = camera(mesh), s = 0.7 + 0.3 * ease(kp)
      premultipliedBlend(gl)
      mesh.draw(mesh.grid(1, 1), flat, mul4(C, scaling(s, s, s)))
      const opacity = 1 - ease((kp - 0.5) / 0.5)
      if (opacity <= 0) return
      const p = useDeform(gl, mesh, 'warp', warped, kp, mul4(C, IDENTITY), opacity)
      gl.uniform4f(p.uniform('u_uvRect'), 0, 0, 1, 1)
      gl.uniform1f(p.uniform('u_kp'), kp)
      mesh.drawMesh(mesh.grid(DEFORM_GRID_COLS, DEFORM_GRID_ROWS))
    },
  },
]

function toTransition(row: Row): Transition3D {
  return {
    element: row.element,
    ...(row.preset !== undefined ? { preset: row.preset } : {}),
    prepare(mesh) {
      for (const [c, r] of row.grids) mesh.grid(c, r)
      for (const name of row.programs) programOf(mesh, name)
    },
    draw: row.draw,
  }
}

const TRANSITIONS: readonly Transition3D[] = ROWS.map(toTransition)

/** The group's rows (element plus `p15:prstTrans` preset), in sheet order. */
export const DEFORM_ROWS: ReadonlyArray<{ element: string; preset?: string }> = ROWS.map(r => (r.preset !== undefined ? { element: r.element, preset: r.preset } : { element: r.element }))

/** Registers every deformation row (idempotent: the same objects every call). */
export function registerDeform3D(): void {
  for (const t of TRANSITIONS) registerTransition3D(t)
}
