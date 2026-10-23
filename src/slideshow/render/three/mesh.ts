/**
 * The 3-D transition mesh pool, camera maths and standard program (slideshow
 * spec 6.3 "3-D transitions run as GPU compositor shaders", S3-1).
 *
 * Conventions every 3-D transition (S3-2 to S3-4) builds on:
 * - **World space.** The slide is the rectangle `x ∈ [-aspect/2, aspect/2]`,
 *   `y ∈ [-1/2, 1/2]` on the plane `z = 0`, y up, the viewer on +z.
 *   `aspect` is the slide's width / height (`MeshPool.aspect`). Slide uv
 *   `(u, v)` has its origin at the top-left corner (v down), the same as
 *   the uploaded frame textures (unflipped), so `slidePoint(u, v)` and the
 *   GLSL `slidePos(uv, aspect)` both map uv `(0, 0)` to `(-aspect/2, 1/2, 0)`.
 * - **Camera.** `slideCamera(fov, aspect)` is `perspective(fov, aspect)`
 *   behind `translation(0, 0, -cameraDistance(fov))`, the distance at which
 *   the `z = 0` slide exactly fills the viewport. The compositor sets the GL
 *   viewport to the slide's device rect, so an identity model draws the
 *   frame exactly where the flat quad would (progress 0 and 1 match the 2-D
 *   frames). The default field of view is `DEFAULT_FOV_DEG` (45°).
 * - **No depth buffer.** The drawing buffer has none (S1-14 context
 *   attributes) and an offscreen one would cost a full surface at 4K, so the
 *   3-D path allocates no surface at all: a transition orders its draws back
 *   to front (painter's order) and may enable `CULL_FACE`. Front faces are
 *   counter-clockwise (GL's default) and the grid's faces front the viewer at
 *   rest. The compositor resets every GL state the draw may touch afterwards.
 * - **Matrices** are column-major `Float32Array(16)`, as GLSL `mat4` reads
 *   them: `mul4(a, b)` applies `b`, then `a`.
 * - **Grid.** `grid(cols, rows)` is `cols × rows` unshared quads (4 vertices,
 *   6 `UNSIGNED_SHORT` indices each), so one mesh serves both smooth
 *   deformation (a vertex shader moving each vertex by its uv) and tiles
 *   (each cell moving as a unit, keyed by its centre). Vertex attributes are
 *   fixed by location, so a vertex array works with any program:
 *   `layout(location = 0) in vec2 a_uv` (the vertex's slide uv) and
 *   `layout(location = 1) in vec2 a_cell` (its cell's centre uv). Draw with
 *   `drawMesh(mesh)` or `gl.drawElements(gl.TRIANGLES, mesh.count,
 *   gl.UNSIGNED_SHORT, 0)` after binding `mesh.vao`. Cols and rows are
 *   integers in `1..GRID_MAX`.
 * - **Workspace.** Every grid leases `gridBytes(cols, rows)` (`webgl2:mesh`)
 *   before any GL object is created; a refusal throws `Transition3DError`
 *   (`'workspace'`) holding nothing, and the compositor draws the row's
 *   look-alike instead, diagnosed. Grids are cached (at most
 *   `MAX_CACHED_GRIDS`, least recently used evicted). Programs hold no
 *   client memory and take no lease (as in S1-14); a failed compile throws
 *   `Transition3DError` (`'compile'`) and is remembered until the pool is
 *   rebuilt. Handles are valid for the current frame only: a transition
 *   asks `grid()` / `program()` again every frame (a cache hit) and never
 *   keeps one, because context loss drops them all.
 */
import type { AllocationOwner, Lease } from '../../types'

export type Mat4 = Float32Array
export type Vec3 = readonly [number, number, number]

export const DEFAULT_FOV_DEG = 45
/** Largest grid side; `GRID_MAX²` cells × 4 vertices still fit 16-bit indices. */
export const GRID_MAX = 128
/**
 * Grids the pool keeps (least recently used evicted beyond). S3-7: with all three groups registered the rows
 * use 9 distinct grids (1×1, 32×18, 8×18, 64×36, 24×1, 24×14, 12×7, 16×9, 20×11; 291,308 B leased in all),
 * so 12 holds every one with room for a new row: a slide change never evicts a grid it is about to use.
 * Each grid is still leased (`webgl2:mesh`) on first use only, so the cache costs only the grids a show plays.
 */
export const MAX_CACHED_GRIDS = 12
/** Bytes per vertex: `a_uv` and `a_cell`, two `vec2` of float. */
export const GRID_VERTEX_BYTES = 16
export const ATTR_UV = 0
export const ATTR_CELL = 1

/** Client bytes a `cols × rows` grid leases: its vertex buffer plus its 16-bit index buffer. */
export const gridBytes = (cols: number, rows: number): number => cols * rows * (4 * GRID_VERTEX_BYTES + 6 * 2)

export interface MeshHandle { vao: WebGLVertexArrayObject; count: number }

export interface Program3D {
  readonly program: WebGLProgram
  /** The uniform's location (looked up once), or null if the program has no such active uniform. */
  uniform(name: string): WebGLUniformLocation | null
}

export interface DrawOptions {
  /** Multiplies the premultiplied colour; default 1. */
  opacity?: number
  /** Multiplies rgb only (a lighting term for turned faces); default 1. */
  shade?: number
  /** The slide-uv rect `[u, v, w, h]` the mesh's uv `[0, 1]²` maps to, for both position and texture; default the whole slide. */
  uvRect?: readonly [number, number, number, number]
}

export interface MeshPool {
  /** A cached `cols × rows` grid; throws `RangeError` (bad size) or `Transition3DError` (workspace refused). */
  grid(cols: number, rows: number): MeshHandle
  /** Additive (S3-1): a linked program cached by `key`; throws `Transition3DError` (`'compile'`). */
  program(key: string, vertex: string, fragment: string): Program3D
  /** Additive (S3-1): draws `mesh` textured with `texture` under `mvp`, with the standard program. */
  draw(mesh: MeshHandle, texture: WebGLTexture, mvp: Mat4, options?: DrawOptions): void
  /** Additive (S3-1): binds `mesh` and issues its indexed draw with whatever program is in use. */
  drawMesh(mesh: MeshHandle): void
  /** Additive (S3-1): the slide's width / height for this frame. */
  readonly aspect: number
}

export class Transition3DError extends Error {
  constructor(readonly reason: 'workspace' | 'compile' | 'lost', message: string) {
    super(message)
    this.name = 'Transition3DError'
  }
}

// ---------------------------------------------------------------------------
// Pure maths
// ---------------------------------------------------------------------------

export function identity4(): Mat4 {
  const m = new Float32Array(16)
  m[0] = m[5] = m[10] = m[15] = 1
  return m
}

/** `a ∘ b`: apply `b`, then `a`. */
export function mul4(a: Mat4, b: Mat4): Mat4 {
  const out = new Float32Array(16)
  for (let c = 0; c < 4; c++) {
    for (let r = 0; r < 4; r++) {
      let s = 0
      for (let k = 0; k < 4; k++) s += a[k * 4 + r] * b[c * 4 + k]
      out[c * 4 + r] = s
    }
  }
  return out
}

export function translation(x: number, y: number, z: number): Mat4 {
  const m = identity4()
  m[12] = x; m[13] = y; m[14] = z
  return m
}

export function scaling(x: number, y: number, z: number): Mat4 {
  const m = identity4()
  m[0] = x; m[5] = y; m[10] = z
  return m
}

/** Right-handed rotation by `rad` about +x (y turns towards +z). */
export function rotationX(rad: number): Mat4 {
  const c = Math.cos(rad), s = Math.sin(rad), m = identity4()
  m[5] = c; m[6] = s; m[9] = -s; m[10] = c
  return m
}

/** Right-handed rotation by `rad` about +y (z turns towards +x). */
export function rotationY(rad: number): Mat4 {
  const c = Math.cos(rad), s = Math.sin(rad), m = identity4()
  m[0] = c; m[2] = -s; m[8] = s; m[10] = c
  return m
}

/** Right-handed rotation by `rad` about +z (x turns towards +y: counter-clockwise on screen). */
export function rotationZ(rad: number): Mat4 {
  const c = Math.cos(rad), s = Math.sin(rad), m = identity4()
  m[0] = c; m[1] = s; m[4] = -s; m[5] = c
  return m
}

/** GL perspective projection: vertical field of view `fovDeg`, width / height `aspect`, looking down -z. */
export function perspective(fovDeg: number, aspect: number, near = 0.01, far = 100): Mat4 {
  const f = 1 / Math.tan((fovDeg * Math.PI) / 360)
  const m = new Float32Array(16)
  m[0] = f / aspect
  m[5] = f
  m[10] = (far + near) / (near - far)
  m[11] = -1
  m[14] = (2 * far * near) / (near - far)
  return m
}

/** How far the camera stands from the `z = 0` slide for it to fill a `fovDeg` vertical view. */
export const cameraDistance = (fovDeg: number): number => 0.5 / Math.tan((fovDeg * Math.PI) / 360)

/** The view-projection that shows the `z = 0` slide filling the viewport. */
export function slideCamera(fovDeg: number, aspect: number): Mat4 {
  return mul4(perspective(fovDeg, aspect), translation(0, 0, -cameraDistance(fovDeg)))
}

/** The world position of slide uv `(u, v)` at rest. */
export const slidePoint = (u: number, v: number, aspect: number): Vec3 => [(u - 0.5) * aspect, 0.5 - v, 0]

/** `m · (p, 1)` divided by w: normalised device coordinates (x right, y up, z depth). */
export function project(m: Mat4, p: Vec3): [number, number, number] {
  const x = m[0] * p[0] + m[4] * p[1] + m[8] * p[2] + m[12]
  const y = m[1] * p[0] + m[5] * p[1] + m[9] * p[2] + m[13]
  const z = m[2] * p[0] + m[6] * p[1] + m[10] * p[2] + m[14]
  const w = m[3] * p[0] + m[7] * p[1] + m[11] * p[2] + m[15]
  return [x / w, y / w, z / w]
}

/** The shared timing curve (smoothstep): 0 → 0, 1/2 → 1/2, 1 → 1, flat at both ends. */
export function ease(t: number): number {
  const x = Math.min(1, Math.max(0, t))
  return x * x * (3 - 2 * x)
}

/** lowbias32 integer hash (exact on every GPU as GLSL ES 3.0 `uint` ops; see `HASH_GLSL`). */
export function hash32(x: number): number {
  let h = x >>> 0
  h ^= h >>> 16
  h = Math.imul(h, 0x7feb352d) >>> 0
  h ^= h >>> 15
  h = Math.imul(h, 0x846ca68b) >>> 0
  h ^= h >>> 16
  return h >>> 0
}

/**
 * The per-cell random in [0, 1) every seeded 3-D row uses, identical on the
 * CPU and in GLSL (`cellRandom(seed, i, j, k)`): `seed` is the effect seed
 * (`Math.floor(rng() * 2 ** 32)`, the shader rng's first value), `(i, j)`
 * the cell's column and row, `k` which value of that cell.
 */
export function cellRandom(seed: number, i: number, j: number, k: number): number {
  const key = ((i >>> 0) + Math.imul(1024, ((j >>> 0) + Math.imul(1024, k >>> 0)) >>> 0)) >>> 0
  return hash32((seed >>> 0) ^ hash32(key)) / 4294967296
}

/** The effect seed a 3-D row derives from its (re-seeded) rng: its first value. */
export const effectSeed = (rng: () => number): number => Math.floor(rng() * 4294967296) >>> 0

// ---------------------------------------------------------------------------
// Shaders
// ---------------------------------------------------------------------------

/** GLSL `ease`, `hash32` and `cellRandom`, matching the CPU functions bit for bit (the hash) and to float precision (the rest). */
export const HASH_GLSL = `float ease(float t) { float x = clamp(t, 0.0, 1.0); return x * x * (3.0 - 2.0 * x); }
uint hash32(uint x) { x ^= x >> 16; x *= 0x7feb352du; x ^= x >> 15; x *= 0x846ca68bu; x ^= x >> 16; return x; }
float cellRandom(uint seed, uint i, uint j, uint k) { return float(hash32(seed ^ hash32(i + 1024u * (j + 1024u * k)))) / 4294967296.0; }
`

/** GLSL: the rest position of slide uv `uv` (paste after the declarations of a custom vertex shader). */
export const SLIDE_GLSL = `vec3 slidePos(vec2 uv, float aspect) { return vec3((uv.x - 0.5) * aspect, 0.5 - uv.y, 0.0); }
`

/** The grid's attribute declarations, by fixed location. */
export const GRID_ATTRIBUTES_GLSL = `layout(location = 0) in vec2 a_uv;
layout(location = 1) in vec2 a_cell;
`

export const VERTEX_3D = `#version 300 es
${GRID_ATTRIBUTES_GLSL}uniform mat4 u_mvp;
uniform float u_aspect;
uniform vec4 u_uvRect;
out vec2 v_uv;
${SLIDE_GLSL}void main() {
  vec2 uv = u_uvRect.xy + a_uv * u_uvRect.zw;
  v_uv = uv;
  gl_Position = u_mvp * vec4(slidePos(uv, u_aspect), 1.0);
}
`

/** Premultiplied in, premultiplied out (blend `ONE, ONE_MINUS_SRC_ALPHA`). */
export const FRAGMENT_3D = `#version 300 es
precision highp float;
in vec2 v_uv;
uniform sampler2D u_tex;
uniform float u_opacity;
uniform float u_shade;
out vec4 outColor;
void main() {
  vec4 c = texture(u_tex, v_uv);
  outColor = vec4(c.rgb * u_shade, c.a) * u_opacity;
}
`

export const STANDARD_PROGRAM = 'three:standard'

// ---------------------------------------------------------------------------
// The pool
// ---------------------------------------------------------------------------

export interface MeshPoolOwner {
  readonly pool: MeshPool
  /** The slide aspect the next draw uses. */
  setAspect(aspect: number): void
  /** Client bytes the pool holds. */
  bytes(): number
  /** Returns every lease; deletes GL objects only when `deleteObjects` (not after context loss). */
  dispose(deleteObjects: boolean): void
}

interface Grid { handle: MeshHandle; buffers: WebGLBuffer[]; lease: Lease; used: number }

function gridData(cols: number, rows: number): { vertices: Float32Array; indices: Uint16Array } {
  const vertices = new Float32Array(cols * rows * 16)
  const indices = new Uint16Array(cols * rows * 6)
  let v = 0, i = 0, base = 0
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const u0 = c / cols, u1 = (c + 1) / cols, v0 = r / rows, v1 = (r + 1) / rows
      const cu = (c + 0.5) / cols, cv = (r + 0.5) / rows
      // Top-left, top-right, bottom-left, bottom-right.
      for (const [u, w] of [[u0, v0], [u1, v0], [u0, v1], [u1, v1]]) {
        vertices[v++] = u; vertices[v++] = w; vertices[v++] = cu; vertices[v++] = cv
      }
      // Counter-clockwise seen from +z (world y is up, uv v is down).
      indices[i++] = base; indices[i++] = base + 2; indices[i++] = base + 1
      indices[i++] = base + 1; indices[i++] = base + 2; indices[i++] = base + 3
      base += 4
    }
  }
  return { vertices, indices }
}

export function createMeshPool(gl: WebGL2RenderingContext, owner: AllocationOwner): MeshPoolOwner {
  const grids = new Map<string, Grid>()
  const programs = new Map<string, { p: Program3D; shaders: WebGLShader[]; uniforms: Map<string, WebGLUniformLocation | null> }>()
  const failed = new Set<string>()
  let aspect = 16 / 9
  let tick = 0

  function deleteGrid(g: Grid, deleteObjects: boolean) {
    if (deleteObjects) {
      gl.deleteVertexArray(g.handle.vao)
      for (const b of g.buffers) gl.deleteBuffer(b)
    }
    g.lease.release()
  }

  function grid(cols: number, rows: number): MeshHandle {
    if (!Number.isInteger(cols) || !Number.isInteger(rows) || cols < 1 || rows < 1 || cols > GRID_MAX || rows > GRID_MAX) {
      throw new RangeError(`A 3-D grid is 1..${GRID_MAX} cells per side, not ${cols} × ${rows}`)
    }
    const key = `${cols}x${rows}`
    const hit = grids.get(key)
    if (hit) { hit.used = ++tick; return hit.handle }
    if (gl.isContextLost()) throw new Transition3DError('lost', 'The WebGL2 context is lost')
    if (grids.size >= MAX_CACHED_GRIDS) {
      let oldest: [string, Grid] | undefined
      for (const e of grids) if (!oldest || e[1].used < oldest[1].used) oldest = e
      if (oldest) { grids.delete(oldest[0]); deleteGrid(oldest[1], true) }
    }
    const lease = owner.lease(gridBytes(cols, rows), 'webgl2:mesh')
    if (!lease) throw new Transition3DError('workspace', `The ${cols} × ${rows} transition mesh did not fit the slideshow workspace`)
    const vao = gl.createVertexArray()
    const vbo = gl.createBuffer()
    const ibo = gl.createBuffer()
    if (!vao || !vbo || !ibo) {
      if (vao) gl.deleteVertexArray(vao)
      if (vbo) gl.deleteBuffer(vbo)
      if (ibo) gl.deleteBuffer(ibo)
      lease.release()
      throw new Transition3DError('lost', 'The transition mesh could not be created')
    }
    const { vertices, indices } = gridData(cols, rows)
    gl.bindVertexArray(vao)
    gl.bindBuffer(gl.ARRAY_BUFFER, vbo)
    gl.bufferData(gl.ARRAY_BUFFER, vertices, gl.STATIC_DRAW)
    gl.enableVertexAttribArray(ATTR_UV)
    gl.vertexAttribPointer(ATTR_UV, 2, gl.FLOAT, false, GRID_VERTEX_BYTES, 0)
    gl.enableVertexAttribArray(ATTR_CELL)
    gl.vertexAttribPointer(ATTR_CELL, 2, gl.FLOAT, false, GRID_VERTEX_BYTES, 8)
    gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, ibo)
    gl.bufferData(gl.ELEMENT_ARRAY_BUFFER, indices, gl.STATIC_DRAW)
    gl.bindVertexArray(null)
    const g: Grid = { handle: { vao, count: indices.length }, buffers: [vbo, ibo], lease, used: ++tick }
    grids.set(key, g)
    return g.handle
  }

  function compile(type: number, source: string): WebGLShader | undefined {
    const s = gl.createShader(type)
    if (!s) return undefined
    gl.shaderSource(s, source)
    gl.compileShader(s)
    if (!gl.getShaderParameter(s, gl.COMPILE_STATUS)) { gl.deleteShader(s); return undefined }
    return s
  }

  function program(key: string, vertex: string, fragment: string): Program3D {
    const hit = programs.get(key)
    if (hit) return hit.p
    if (failed.has(key)) throw new Transition3DError('compile', `The 3-D program ${key} did not compile`)
    if (gl.isContextLost()) throw new Transition3DError('lost', 'The WebGL2 context is lost')
    const vs = compile(gl.VERTEX_SHADER, vertex)
    const fs = compile(gl.FRAGMENT_SHADER, fragment)
    const prog = vs && fs ? gl.createProgram() : null
    if (prog && vs && fs) {
      gl.attachShader(prog, vs)
      gl.attachShader(prog, fs)
      gl.linkProgram(prog)
      if (gl.getProgramParameter(prog, gl.LINK_STATUS)) {
        const uniforms = new Map<string, WebGLUniformLocation | null>()
        const p: Program3D = {
          program: prog,
          uniform(name) {
            if (!uniforms.has(name)) uniforms.set(name, gl.getUniformLocation(prog, name))
            return uniforms.get(name) ?? null
          },
        }
        programs.set(key, { p, shaders: [vs, fs], uniforms })
        return p
      }
    }
    if (prog) gl.deleteProgram(prog)
    if (vs) gl.deleteShader(vs)
    if (fs) gl.deleteShader(fs)
    failed.add(key)
    throw new Transition3DError('compile', `The 3-D program ${key} did not compile`)
  }

  function drawMesh(mesh: MeshHandle) {
    gl.bindVertexArray(mesh.vao)
    gl.drawElements(gl.TRIANGLES, mesh.count, gl.UNSIGNED_SHORT, 0)
  }

  const pool: MeshPool = {
    grid,
    program,
    drawMesh,
    get aspect() { return aspect },
    draw(mesh, texture, mvp, options = {}) {
      const p = program(STANDARD_PROGRAM, VERTEX_3D, FRAGMENT_3D)
      gl.useProgram(p.program)
      gl.activeTexture(gl.TEXTURE0)
      gl.bindTexture(gl.TEXTURE_2D, texture)
      gl.uniform1i(p.uniform('u_tex'), 0)
      gl.uniformMatrix4fv(p.uniform('u_mvp'), false, mvp)
      gl.uniform1f(p.uniform('u_aspect'), aspect)
      const r = options.uvRect ?? [0, 0, 1, 1]
      gl.uniform4f(p.uniform('u_uvRect'), r[0], r[1], r[2], r[3])
      gl.uniform1f(p.uniform('u_opacity'), options.opacity ?? 1)
      gl.uniform1f(p.uniform('u_shade'), options.shade ?? 1)
      drawMesh(mesh)
    },
  }

  return {
    pool,
    setAspect(a) { if (Number.isFinite(a) && a > 0) aspect = a },
    bytes() { let n = 0; for (const g of grids.values()) n += g.lease.bytes; return n },
    dispose(deleteObjects) {
      for (const g of grids.values()) deleteGrid(g, deleteObjects)
      grids.clear()
      for (const e of programs.values()) {
        if (deleteObjects) {
          gl.deleteProgram(e.p.program)
          for (const s of e.shaders) gl.deleteShader(s)
        }
      }
      programs.clear()
      failed.clear()
    },
  }
}
