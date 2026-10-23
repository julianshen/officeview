/**
 * The tile and particle 3-D transitions (slideshow spec 6.3, plan S3-3):
 * vortex, shred, fracture, crush, glitter, honeycomb and ripple. The numbers
 * are the algorithm sheet's (docs/validation/slideshow-3d-algorithms.md,
 * "S3-3"); `scripts/slideshow-3d-oracles.ts` is the executable reference and
 * wins any disagreement. `flash` and `reveal` are in the plan's S3-3 line but
 * are exact 2-D rows (S1-15), so they have no shader and are not registered.
 *
 * - **One draw per row.** Every tile of a row moves in a single indexed draw:
 *   the grid's unshared quads carry the cell centre (`a_cell`), the vertex
 *   shader derives the cell's index, its seeded randoms (`cellRandom`, the
 *   framework's exact integer hash) and its local time from it, and moves the
 *   vertex as part of its tile. Draw calls per frame: a flat backdrop plus
 *   one (glitter: two culled passes). No call is per tile.
 * - **A pure plan.** `particlePlan` turns the attributes, the progress, the
 *   aspect and the effect seed into the grid, the programs and every uniform.
 *   `draw` only executes the plan, so the same inputs give the same frame;
 *   the shader rng is the framework's re-seeded one and only its first value
 *   (`effectSeed`) is read. Nothing here reads `Math.random` or a clock.
 * - **Ordering.** The drawing buffer has no depth (mesh.ts): the flat frame
 *   goes first, the tiles over it in cell order, and glitter draws the
 *   incoming faces before the outgoing ones, each culled to its own side.
 * - **Resources.** `prepare` leases every grid the row can use (it sees no
 *   attributes: glitter and shred lease both variants) and compiles the
 *   row's program before the effect starts; a refusal sends the whole effect
 *   to its look-alike. Every grid is at most 64 × 36 cells (2,304 quads), far
 *   under `GRID_MAX`, and its lease is charged to the slide's workspace.
 * - **Registration.** `registerParticle3D()` registers the seven rows; the
 *   wiring in `three/index.ts` is S3-7's.
 */
import { DEFAULT_FOV_DEG, GRID_ATTRIBUTES_GLSL, HASH_GLSL, SLIDE_GLSL, effectSeed, slideCamera, type MeshHandle, type MeshPool, type Program3D } from './mesh'
import { registerTransition3D, type Transition3D } from './registry'
import type { Rng } from '../../types'

export type ShaderKey = 'vortex' | 'shred' | 'fracture' | 'crush' | 'glitter' | 'honeycomb' | 'ripple'
export type Uniform = number | readonly number[] | Float32Array

export interface ParticlePass {
  shader: ShaderKey
  /** The frame the pass's tiles show. */
  texture: 'from' | 'to'
  uniforms: Record<string, Uniform>
  /** Which faces are discarded (glitter's two-sided tiles); none when absent. */
  cull?: 'front' | 'back'
}

export interface ParticlePlan {
  /** The tile grid (cols, rows) every pass draws. */
  grid: readonly [number, number]
  /** A frame drawn flat first (the backdrop the tiles reveal), if any. */
  flat?: 'from' | 'to'
  passes: ParticlePass[]
  progress: number
}

// ---------------------------------------------------------------------------
// Shaders
// ---------------------------------------------------------------------------

const VERTEX_PRELUDE = `#version 300 es
${GRID_ATTRIBUTES_GLSL}uniform mat4 u_cam;
uniform float u_aspect;
uniform float u_p;
uniform uint u_seed;
uniform vec2 u_grid;
uniform vec2 u_wj;
uniform vec2 u_gw;
uniform float u_g0;
uniform float u_centre;
out vec2 v_uv;
flat out float v_t;
out float v_opacity;
out float v_shade;
out vec2 v_local;
out vec3 v_world;
out vec4 v_clip;
${SLIDE_GLSL}${HASH_GLSL}const float PI = 3.14159265358979;
vec3 rotX(vec3 v, float a) { float c = cos(a), s = sin(a); return vec3(v.x, c * v.y - s * v.z, s * v.y + c * v.z); }
vec3 rotY(vec3 v, float a) { float c = cos(a), s = sin(a); return vec3(c * v.x + s * v.z, v.y, -s * v.x + c * v.z); }
vec3 rotZ(vec3 v, float a) { float c = cos(a), s = sin(a); return vec3(c * v.x - s * v.y, s * v.x + c * v.y, v.z); }
float fade(float t) { return 1.0 - ease((t - 0.7) / 0.3); }
uvec2 cellIndex() { return uvec2(floor(a_cell * u_grid)); }
float rnd(uvec2 c, uint k) { return cellRandom(u_seed, c.x, c.y, k); }
float order() {
  if (u_centre > 0.5) return min(1.0, length((a_cell - 0.5) * vec2(u_aspect, 1.0)) / length(vec2(u_aspect * 0.5, 0.5)));
  return dot(a_cell, u_gw) + u_g0;
}
float tileT(float r0) {
  float delay = u_wj.x * ((1.0 - u_wj.y) * order() + u_wj.y * r0);
  return clamp((u_p - delay) / (1.0 - u_wj.x), 0.0, 1.0);
}
void emit(vec3 world, float t, float opacity, float shade, vec2 uv) {
  v_uv = uv;
  v_t = t;
  v_opacity = opacity;
  v_shade = shade;
  v_local = (a_uv - a_cell) * u_grid;
  v_world = world;
  vec4 clip = u_cam * vec4(world, 1.0);
  v_clip = clip;
  gl_Position = clip;
}
`

/** vortex (sheet "vortexTile"): tiles spiral off towards `u_m`, turning and shrinking. */
const VORTEX_VS = `${VERTEX_PRELUDE}uniform vec2 u_m;
void main() {
  uvec2 ij = cellIndex();
  float t = tileT(rnd(ij, 0u));
  float r1 = rnd(ij, 1u), r2 = rnd(ij, 2u), r3 = rnd(ij, 3u);
  vec3 c = slidePos(a_cell, u_aspect);
  float swirl = 0.4 * sin(PI * t) * (2.0 * r2 - 1.0);
  vec2 off = u_m * (1.4 * u_aspect * t * t) + swirl * vec2(-u_m.y, u_m.x);
  vec3 q = slidePos(a_uv, u_aspect) - c;
  q.xy *= 1.0 - 0.5 * t;
  q = rotY(q, PI * t * (r3 - 0.5));
  q = rotZ(q, 2.0 * PI * (0.75 + 0.5 * r1) * t);
  q += vec3(c.xy + off, 0.5 * t);
  emit(q, t, fade(t), 1.0, a_uv);
}
`

/** shred (sheet "shredTile"): strips slide off up and down; rectangles scatter. `u_mode` 0 is strip, 1 is rectangle. */
const SHRED_VS = `${VERTEX_PRELUDE}uniform float u_mode;
void main() {
  uvec2 ij = cellIndex();
  float t = tileT(rnd(ij, 0u));
  vec3 p = slidePos(a_uv, u_aspect);
  vec3 q;
  if (u_mode < 0.5) {
    float up = (ij.x % 2u == 0u) ? 1.0 : -1.0;
    q = p + vec3(0.0, up * 1.3 * t * t, 0.0);
  } else {
    vec3 c = slidePos(a_cell, u_aspect);
    float ang = 2.0 * PI * rnd(ij, 1u);
    q = p - c;
    q.xy *= 1.0 - 0.6 * t;
    q = rotZ(q, 2.0 * PI * rnd(ij, 2u) * t);
    q += vec3(c.xy + 1.5 * t * vec2(cos(ang), sin(ang)), 0.8 * t);
  }
  emit(q, t, fade(t), 1.0, a_uv);
}
`

/** fracture (sheet "fractureTile"): shards fall away from the centre, tumbling. */
const FRACTURE_VS = `${VERTEX_PRELUDE}void main() {
  uvec2 ij = cellIndex();
  float t = tileT(rnd(ij, 0u));
  float r1 = rnd(ij, 1u), r2 = rnd(ij, 2u), r3 = rnd(ij, 3u);
  vec3 c = slidePos(a_cell, u_aspect);
  float len = length(c.xy);
  if (len == 0.0) len = 1.0;
  vec3 q = slidePos(a_uv, u_aspect) - c;
  q = rotY(q, PI * r2 * t);
  q = rotX(q, PI * r1 * t);
  q += vec3(c.x + 0.4 * (c.x / len) * t, c.y - 1.6 * t * t, 0.3 * t * r3);
  emit(q, t, fade(t), 1.0, a_uv);
}
`

/** crush (sheet "crushTile"): tiles gather into a ball that recedes and fades. */
const CRUSH_VS = `${VERTEX_PRELUDE}void main() {
  uvec2 ij = cellIndex();
  float t = tileT(rnd(ij, 0u));
  float r1 = rnd(ij, 1u), r2 = rnd(ij, 2u);
  vec3 c = slidePos(a_cell, u_aspect);
  float k = 0.9 * ease(t), sc = 1.0 - 0.8 * ease(t);
  vec3 q = slidePos(a_uv, u_aspect) - c;
  q.xy *= sc;
  q = rotX(q, PI * r2 * t);
  q = rotZ(q, 2.0 * PI * (r1 - 0.5) * t);
  q += vec3(c.xy * (1.0 - k), -2.0 * ease((u_p - 0.6) / 0.4));
  emit(q, t, 1.0 - ease((u_p - 0.75) / 0.25), 1.0, a_uv);
}
`

/**
 * glitter (sheet "glitterTile"): every cell turns over about its centre. The
 * outgoing frame is on the front face and the incoming one on the back face,
 * so the back pass (`u_back` 1) samples the incoming frame mirrored about the
 * cell, which is the right way round once the cell has turned a half turn.
 */
const GLITTER_VS = `${VERTEX_PRELUDE}uniform float u_horiz;
uniform float u_sgn;
uniform float u_back;
void main() {
  uvec2 ij = cellIndex();
  float t = tileT(rnd(ij, 0u));
  vec3 c = slidePos(a_cell, u_aspect);
  float angle = PI * ease(t);
  vec3 q = slidePos(a_uv, u_aspect) - c;
  q = u_horiz > 0.5 ? rotY(q, u_sgn * angle) : rotX(q, u_sgn * angle);
  q += c;
  vec2 uv = a_uv;
  if (u_back > 0.5) uv = u_horiz > 0.5 ? vec2(2.0 * a_cell.x - a_uv.x, a_uv.y) : vec2(a_uv.x, 2.0 * a_cell.y - a_uv.y);
  emit(q, t, 1.0, 1.0 + 0.6 * sin(PI * t), uv);
}
`

/** honeycomb (sheet "honeycombTile"): cells shrink, turn and rise, masked to a hexagon in flight. */
const HONEYCOMB_VS = `${VERTEX_PRELUDE}void main() {
  uvec2 ij = cellIndex();
  float t = tileT(rnd(ij, 0u));
  vec3 c = slidePos(a_cell, u_aspect);
  float sc = 1.0 - ease(t);
  vec3 q = slidePos(a_uv, u_aspect) - c;
  q.xy *= sc;
  q = rotZ(q, (PI / 3.0) * t);
  q += vec3(c.xy, 0.6 * t);
  emit(q, t, 1.0 - ease(t), 1.0, a_uv);
}
`

/** ripple (sheet "rippleZ"): the outgoing frame ripples in z from `u_origin` while it fades. */
const RIPPLE_VS = `${VERTEX_PRELUDE}uniform vec2 u_origin;
void main() {
  float rho = length((a_uv - u_origin) * vec2(u_aspect, 1.0));
  float z = 0.06 * 4.0 * u_p * (1.0 - u_p) * sin(2.0 * PI * (4.0 * rho - 3.0 * u_p)) * exp(-1.5 * rho);
  emit(slidePos(a_uv, u_aspect) + vec3(0.0, 0.0, z), u_p, 1.0 - ease(u_p), 1.0, a_uv);
}
`

const FRAGMENT_PRELUDE = `#version 300 es
precision highp float;
in vec2 v_uv;
flat in float v_t;
in float v_opacity;
in float v_shade;
in vec2 v_local;
uniform sampler2D u_tex;
out vec4 outColor;
bool inDiamond(vec2 q) { return abs(q.x) + abs(q.y) <= 0.5; }
bool inHexagon(vec2 q) { vec2 a = abs(q); return a.x <= 0.4330127 && a.y + a.x * 0.5773503 <= 0.5; }
`

/** Premultiplied in, premultiplied out; \`body\` may change \`shade\` or discard. */
const fragment = (uniforms: string, body: string) => `${FRAGMENT_PRELUDE}${uniforms}void main() {
  float shade = v_shade;
${body}  vec4 c = texture(u_tex, v_uv);
  outColor = vec4(c.rgb * shade, c.a) * v_opacity;
}
`

const TILE_FS = fragment('', '')
/** Glitter's sparkle mask changes the shading only; cells at rest stay whole squares. */
const GLITTER_FS = fragment('uniform float u_pattern;\n', '  if (!(u_pattern > 0.5 ? inHexagon(v_local) : inDiamond(v_local))) shade = 1.0;\n')
/** Honeycomb keeps only the hexagon inscribed in a cell while it is in flight. */
const HONEYCOMB_FS = fragment('', '  if (v_t > 0.0 && v_t < 1.0 && !inHexagon(v_local)) discard;\n')

export const PARTICLE_SHADERS: Readonly<Record<ShaderKey, { vertex: string; fragment: string }>> = Object.freeze({
  vortex: { vertex: VORTEX_VS, fragment: TILE_FS },
  shred: { vertex: SHRED_VS, fragment: TILE_FS },
  fracture: { vertex: FRACTURE_VS, fragment: TILE_FS },
  crush: { vertex: CRUSH_VS, fragment: TILE_FS },
  glitter: { vertex: GLITTER_VS, fragment: GLITTER_FS },
  honeycomb: { vertex: HONEYCOMB_VS, fragment: HONEYCOMB_FS },
  ripple: { vertex: RIPPLE_VS, fragment: TILE_FS },
})

// ---------------------------------------------------------------------------
// Plans
// ---------------------------------------------------------------------------

type Dir4 = 'l' | 'u' | 'r' | 'd'
const dir4 = (v: string | undefined): Dir4 => (v === 'u' || v === 'r' || v === 'd' ? v : 'l')

/** Leading-edge order g = gw · uv + g0: the edge the motion heads for goes first (sheet "leading"). */
const LEADING: Record<Dir4, { gw: readonly [number, number]; g0: number }> = {
  l: { gw: [1, 0], g0: 0 }, r: { gw: [-1, 0], g0: 1 }, u: { gw: [0, 1], g0: 0 }, d: { gw: [0, -1], g0: 1 },
}
/** The unit heading of each direction in world xy (y up). */
const HEADING: Record<Dir4, readonly [number, number]> = { l: [-1, 0], r: [1, 0], u: [0, 1], d: [0, -1] }
const RIPPLE_ORIGIN: Record<string, readonly [number, number]> = { center: [0.5, 0.5], lu: [0, 0], ru: [1, 0], ld: [0, 1], rd: [1, 1] }

interface TileConfig { grid: readonly [number, number]; W: number; J: number; order: 'centre' | Dir4 | 'none' }

/** The uniforms every tile program reads, for `cfg`. */
function tileUniforms(cfg: TileConfig, p: number, aspect: number, seed: number): Record<string, Uniform> {
  const lead = cfg.order === 'centre' || cfg.order === 'none' ? { gw: [0, 0] as const, g0: 0 } : LEADING[cfg.order]
  return {
    u_cam: slideCamera(DEFAULT_FOV_DEG, aspect), u_aspect: aspect, u_p: p, u_seed: seed,
    u_grid: [cfg.grid[0], cfg.grid[1]], u_wj: [cfg.W, cfg.J], u_gw: [lead.gw[0], lead.gw[1]], u_g0: lead.g0, u_centre: cfg.order === 'centre' ? 1 : 0,
  }
}

interface RowDef {
  element: string
  preset?: string
  /** Every grid the row can draw with, for `prepare`. */
  grids: ReadonlyArray<readonly [number, number]>
  /** Whether the row draws a flat frame (a 1 × 1 grid) first. */
  flat: boolean
  shaders: readonly ShaderKey[]
  plan(attrs: Readonly<Record<string, string>>, p: number, aspect: number, seed: number): ParticlePlan
}

const VORTEX: TileConfig = { grid: [32, 18], W: 0.6, J: 0.35, order: 'l' }
const SHRED_STRIP: TileConfig = { grid: [24, 1], W: 0.5, J: 0.2, order: 'none' }
const SHRED_RECT: TileConfig = { grid: [24, 14], W: 0.5, J: 0.6, order: 'none' }
const FRACTURE: TileConfig = { grid: [12, 7], W: 0.6, J: 0.4, order: 'centre' }
const CRUSH: TileConfig = { grid: [16, 9], W: 0.3, J: 1, order: 'none' }
const GLITTER_DIAMOND: TileConfig = { grid: [32, 18], W: 0.65, J: 0.4, order: 'l' }
const GLITTER_HEXAGON: TileConfig = { grid: [24, 14], W: 0.65, J: 0.4, order: 'l' }
const HONEYCOMB: TileConfig = { grid: [20, 11], W: 0.6, J: 0.3, order: 'centre' }
const RIPPLE_GRID: readonly [number, number] = [64, 36]

const ROWS: readonly RowDef[] = [
  {
    element: 'vortex', grids: [VORTEX.grid], flat: true, shaders: ['vortex'],
    plan(attrs, p, aspect, seed) {
      const dir = dir4(attrs.dir)
      const cfg = { ...VORTEX, order: dir }
      return { grid: cfg.grid, flat: 'to', progress: p, passes: [{ shader: 'vortex', texture: 'from', uniforms: { ...tileUniforms(cfg, p, aspect, seed), u_m: [HEADING[dir][0], HEADING[dir][1]] } }] }
    },
  },
  {
    element: 'shred', grids: [SHRED_STRIP.grid, SHRED_RECT.grid], flat: true, shaders: ['shred'],
    plan(attrs, p, aspect, seed) {
      const rect = attrs.pattern === 'rectangle'
      const inward = attrs.dir === 'in'
      const cfg = rect ? SHRED_RECT : SHRED_STRIP
      // `in`: the roles swap and time runs backwards, so the incoming frame's tiles fly in.
      return {
        grid: cfg.grid, flat: inward ? 'from' : 'to', progress: p,
        passes: [{ shader: 'shred', texture: inward ? 'to' : 'from', uniforms: { ...tileUniforms(cfg, inward ? 1 - p : p, aspect, seed), u_mode: rect ? 1 : 0 } }],
      }
    },
  },
  {
    element: 'prstTrans', preset: 'fracture', grids: [FRACTURE.grid], flat: true, shaders: ['fracture'],
    plan: (_attrs, p, aspect, seed) => ({ grid: FRACTURE.grid, flat: 'to', progress: p, passes: [{ shader: 'fracture', texture: 'from', uniforms: tileUniforms(FRACTURE, p, aspect, seed) }] }),
  },
  {
    element: 'prstTrans', preset: 'crush', grids: [CRUSH.grid], flat: true, shaders: ['crush'],
    plan: (_attrs, p, aspect, seed) => ({ grid: CRUSH.grid, flat: 'to', progress: p, passes: [{ shader: 'crush', texture: 'from', uniforms: tileUniforms(CRUSH, p, aspect, seed) }] }),
  },
  {
    element: 'glitter', grids: [GLITTER_DIAMOND.grid, GLITTER_HEXAGON.grid], flat: false, shaders: ['glitter'],
    plan(attrs, p, aspect, seed) {
      const hexagon = attrs.pattern === 'hexagon'
      const dir = dir4(attrs.dir)
      const cfg = { ...(hexagon ? GLITTER_HEXAGON : GLITTER_DIAMOND), order: dir }
      const common = { ...tileUniforms(cfg, p, aspect, seed), u_horiz: dir === 'l' || dir === 'r' ? 1 : 0, u_sgn: dir === 'l' || dir === 'u' ? -1 : 1, u_pattern: hexagon ? 1 : 0 }
      // The incoming faces (turned half way round) first, then the outgoing ones.
      return {
        grid: cfg.grid, progress: p,
        passes: [
          { shader: 'glitter', texture: 'to', cull: 'front', uniforms: { ...common, u_back: 1 } },
          { shader: 'glitter', texture: 'from', cull: 'back', uniforms: { ...common, u_back: 0 } },
        ],
      }
    },
  },
  {
    element: 'honeycomb', grids: [HONEYCOMB.grid], flat: true, shaders: ['honeycomb'],
    plan: (_attrs, p, aspect, seed) => ({ grid: HONEYCOMB.grid, flat: 'to', progress: p, passes: [{ shader: 'honeycomb', texture: 'from', uniforms: tileUniforms(HONEYCOMB, p, aspect, seed) }] }),
  },
  {
    element: 'ripple', grids: [RIPPLE_GRID], flat: true, shaders: ['ripple'],
    plan(attrs, p, aspect, seed) {
      const o = RIPPLE_ORIGIN[attrs.dir ?? ''] ?? RIPPLE_ORIGIN.center
      return {
        grid: RIPPLE_GRID, flat: 'to', progress: p,
        passes: [{ shader: 'ripple', texture: 'from', uniforms: { u_cam: slideCamera(DEFAULT_FOV_DEG, aspect), u_aspect: aspect, u_p: p, u_seed: seed, u_origin: [o[0], o[1]] } }],
      }
    },
  },
]

export const PARTICLE_ROWS: ReadonlyArray<{ readonly element: string; readonly preset?: string }> = ROWS.map(r => ({ element: r.element, ...(r.preset !== undefined ? { preset: r.preset } : {}) }))

const rowOf = (element: string, preset?: string): RowDef => {
  const row = ROWS.find(r => r.element === element && r.preset === preset)
  if (!row) throw new RangeError(`${element}${preset ? ' ' + preset : ''} is not a tile or particle transition`)
  return row
}

/** The grid, passes and uniforms of `element` at `progress` for a slide of `aspect` and an effect `seed` (the shader rng's first value). */
export function particlePlan(element: string, preset: string | undefined, attrs: Readonly<Record<string, string>>, progress: number, aspect: number, seed: number): ParticlePlan {
  return rowOf(element, preset).plan(attrs, progress, aspect, seed)
}

// ---------------------------------------------------------------------------
// Drawing
// ---------------------------------------------------------------------------

/** Sets `uniforms` on `program` (a seed is an unsigned integer uniform, a `Float32Array(16)` a matrix). */
export function setParticleUniforms(gl: WebGL2RenderingContext, program: Pick<Program3D, 'uniform'>, uniforms: Readonly<Record<string, Uniform>>): void {
  for (const [name, v] of Object.entries(uniforms)) {
    const loc = program.uniform(name)
    if (!loc) continue
    if (typeof v === 'number') {
      if (name === 'u_seed') gl.uniform1ui(loc, v)
      else gl.uniform1f(loc, v)
    } else if (v.length === 16) gl.uniformMatrix4fv(loc, false, v as Float32Array)
    else gl.uniform2f(loc, v[0], v[1])
  }
}

function programOf(mesh: MeshPool, key: ShaderKey): Program3D {
  const s = PARTICLE_SHADERS[key]
  return mesh.program(`particles:${key}`, s.vertex, s.fragment)
}

function transitionFor(row: RowDef): Transition3D {
  return {
    element: row.element,
    ...(row.preset !== undefined ? { preset: row.preset } : {}),
    prepare(mesh) {
      if (row.flat) mesh.grid(1, 1)
      for (const [c, r] of row.grids) mesh.grid(c, r)
      for (const key of row.shaders) programOf(mesh, key)
    },
    draw(gl: WebGL2RenderingContext, from: WebGLTexture, to: WebGLTexture, progress: number, attrs: Readonly<Record<string, string>>, rng: Rng, mesh: MeshPool) {
      const plan = row.plan(attrs, progress, mesh.aspect, effectSeed(rng))
      gl.enable(gl.BLEND)
      gl.blendFunc(gl.ONE, gl.ONE_MINUS_SRC_ALPHA)
      let flat: MeshHandle | undefined
      if (plan.flat) {
        flat = mesh.grid(1, 1)
        mesh.draw(flat, plan.flat === 'to' ? to : from, slideCamera(DEFAULT_FOV_DEG, mesh.aspect))
      }
      const grid = mesh.grid(plan.grid[0], plan.grid[1])
      for (const pass of plan.passes) {
        const program = programOf(mesh, pass.shader)
        gl.useProgram(program.program)
        gl.activeTexture(gl.TEXTURE0)
        gl.bindTexture(gl.TEXTURE_2D, pass.texture === 'to' ? to : from)
        gl.uniform1i(program.uniform('u_tex'), 0)
        setParticleUniforms(gl, program, pass.uniforms)
        if (pass.cull) {
          gl.enable(gl.CULL_FACE)
          gl.cullFace(pass.cull === 'front' ? gl.FRONT : gl.BACK)
        } else gl.disable(gl.CULL_FACE)
        mesh.drawMesh(grid)
      }
      gl.disable(gl.CULL_FACE)
    },
  }
}

const TRANSITIONS: readonly Transition3D[] = ROWS.map(transitionFor)

/** Registers the seven tile and particle rows (idempotent). S3-7 calls this from `three/index.ts`. */
export function registerParticle3D(): void {
  for (const t of TRANSITIONS) registerTransition3D(t)
}
