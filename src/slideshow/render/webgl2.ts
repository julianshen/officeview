/**
 * The WebGL2 compositor (slideshow spec 6.3, 7): the required backend behind
 * the same `Compositor` contract as the Canvas 2D fallback, with which it
 * agrees within the catalogue row's tolerance (tests/browser/slideshow-gpu).
 *
 * Conventions are the Canvas 2D backend's (canvas2d.ts), verbatim:
 * - **Spaces.** Layer `bounds` are slide px, the bitmap is `bounds × scale`.
 *   A draw lands at `view ∘ translate(motion) ∘ state.transform`; a layer
 *   mask covers the bounds box and moves with that matrix. Transition `from`
 *   / `to` affines are in unit slide space; a transition mask is in device
 *   px over the incoming layer's device rect and does not move with `to`.
 *   The transition composites first, then `draws` on top. S1-15: the
 *   transition is scissored to the incoming layer's device rect, edges
 *   rounded (`transitionRect`), as Canvas 2D clips it. Additive params:
 *   `backdrop` clears that rect to an opaque colour under both frames (a
 *   scissored `clear`: no texture, no lease); `fromOnTop` draws the outgoing
 *   frame over the incoming one.
 * - **Colour pipeline.** Uploads use `UNPACK_PREMULTIPLY_ALPHA_WEBGL = true`
 *   (and no colour-space conversion), blending is `ONE, ONE_MINUS_SRC_ALPHA`,
 *   and a part layer's colour matrix runs un-premultiplied in the fragment
 *   shader (shaders.ts). A whole layer (no `part`) ignores `state.color`.
 * - **Masks.** Geometry comes from the shared mask library: a path kind is
 *   traced with `tracePath` onto a 2-D canvas, the same rasteriser the
 *   Canvas 2D backend clips with, and uploaded as a coverage texture (linear
 *   for a layer mask at texture resolution, nearest for a transition mask,
 *   whose texels are device pixels); a noise kind's `MaskRaster` is uploaded
 *   as an `R8` texture sampled nearest, like `imageSmoothingEnabled = false`.
 *   Seeds, op caps (`mask-ops`), unknown kinds (`mask-unknown`) and refused
 *   leases (`compositor-workspace`, a cut at the midpoint) behave exactly as
 *   in the Canvas 2D backend, under the same diagnostic codes.
 * - **Workspace.** Everything is leased before it is allocated: the drawing
 *   buffer (`webgl2:framebuffer`, w × h × 4), the quad buffer
 *   (`webgl2:buffer`), each layer texture (`upload` grows the caller's lease
 *   by the texture, then creates and fills it, then shrinks the lease by the
 *   CPU copy, which the caller drops), the mask canvas plus its texture
 *   (`webgl2:mask`, 8 B/px), the noise texture (`webgl2:noise`) and the
 *   transient noise arrays (`webgl2:mask-raster`). Programs, shaders and the
 *   vertex array hold no client memory and take no lease. The 2-D
 *   transitions (S1-15) need no offscreen framebuffer (their backdrop is a
 *   scissored clear), and neither do the 3-D ones (S3-1). The one offscreen
 *   target is the Morph pair sum (PR #20 review 2, `webgl2:morph-sum`,
 *   w × h × 4): leased before it is created, reused across pairs and
 *   frames, dropped on resize, context loss and dispose.
 * - **3-D transitions (S3-1).** Params carrying `three` (framework.ts) are
 *   drawn by the registered `Transition3D` strictly between the ends (at 0
 *   and 1 the flat outgoing / incoming frame): the slide rect is scissored
 *   and cleared to opaque black, the viewport is that rect, and the shader
 *   gets a rng re-seeded from the effect's seed every frame and the mesh
 *   pool (three/mesh.ts: grids leased `webgl2:mesh` before creation, cached
 *   programs), which is part of `build`/`teardown`, so context loss drops it
 *   and a restore rebuilds it on demand. Every GL state is reset after the
 *   draw. A throwing shader or an unregistered reference draws the params'
 *   2-D look-alike, diagnosed (`transition-shader-fallback`).
 * - **Context loss.** A lost context keeps nothing: `contextLost` drops every
 *   lease and handle. `contextRestored` recreates the program, the buffer,
 *   the vertex array and all GL state, re-leasing each, and returns `false`
 *   (holding nothing) on a refusal. Layer textures are re-rasterised and
 *   re-uploaded by the caller (`BackendManager`'s `rebuildLayers`). Every
 *   entry point is inert while the context is lost.
 */
import type {
  Affine, CanvasFactory, ColorMatrix, Compositor, CompositorViewport, DiagnosticSink, LayerBitmap, LayerDraw, LayerHandle, Lease, MaskState,
  Rng, SlideShowWorkspace, TransitionDraw,
} from '../types'
import { createRng } from '../rng'
import { BACKDROP_RGB, MASK_MAX_OPS, transitionRect, isMaskKind, isMaskRaster, maskOpCount, maskPath, maskRasterBytes, tracePath, type MaskPath, type MaskRaster } from './masks'
import { FRAGMENT_SHADER, UNIFORMS, VERTEX_SHADER, colorUniforms, mulAffine, quadMatrix, type UniformName } from './shaders'
import { createMeshPool, type MeshPoolOwner } from './three/mesh'
import { transition3DFor, type Transition3D } from './three/registry'
import { shader3DOf } from './three/framework'

/** Context attributes: premultiplied output like a 2-D canvas, no MSAA, depth or stencil buffers to pay for. */
export const WEBGL2_CONTEXT_ATTRIBUTES: WebGLContextAttributes = {
  alpha: true, premultipliedAlpha: true, antialias: false, depth: false, stencil: false, preserveDrawingBuffer: false,
}

/** The quad: four `vec2` corners of the unit square, as a triangle strip. */
const QUAD = new Float32Array([0, 0, 1, 0, 0, 1, 1, 1])

export interface WebGL2Compositor extends Compositor {
  readonly backend: 'webgl2'
  /** The context is gone: every lease is returned and every handle is invalid. Idempotent. */
  contextLost(): void
  /** Recreates programs, buffers, vertex arrays and GL state, re-leasing each. `false` holds nothing and stays inert. */
  contextRestored(): boolean
}

export interface WebGL2Options {
  /** Where the 2-D mask canvas comes from; default `canvas.ownerDocument` (realm-safe). */
  factory?: CanvasFactory
  /** Where degraded-effect diagnostics go (the Canvas 2D backend's codes). */
  sink?: DiagnosticSink
}

type Ctx = CanvasRenderingContext2D
type AnyCanvas = HTMLCanvasElement | OffscreenCanvas
type Bounds = LayerBitmap['bounds']

interface Texture { tex: WebGLTexture; bounds: Bounds; lease: Lease; w: number; h: number }
interface Resources {
  program: WebGLProgram
  shaders: WebGLShader[]
  vao: WebGLVertexArrayObject
  buffer: WebGLBuffer
  u: Record<UniformName, WebGLUniformLocation | null>
  frame: Lease
  quad: Lease
  /** S3-1: the 3-D transitions' grids and programs (leased per grid, `webgl2:mesh`). */
  three: MeshPoolOwner
}
interface MaskSurface { canvas: AnyCanvas; ctx: Ctx; tex: WebGLTexture; lease: Lease; w: number; h: number; key?: string }
interface NoiseSurface { tex: WebGLTexture; lease: Lease; w: number; h: number; key?: string }
/** The Morph pair sum target (PR #20 review 2): an output-sized texture behind a framebuffer. */
interface SumTarget { fbo: WebGLFramebuffer; tex: WebGLTexture; lease: Lease; w: number; h: number }
interface TransitionParams {
  mask?: MaskState; from?: Affine; to?: Affine; fromOpacity?: number; toOpacity?: number
  fromOnTop?: boolean; backdrop?: 'black' | 'white'
}
/** A ready mask: mode 1 samples at the quad's uv, mode 2 over `rect` in device px. */
type MaskBinding = { tex: WebGLTexture; mode: 1 | 2; rect: [number, number, number, number] }

const IDENTITY: Affine = [1, 0, 0, 1, 0, 0]
const clamp01 = (v: unknown, fallback: number) => {
  const n = typeof v === 'number' && Number.isFinite(v) ? v : fallback
  return Math.min(1, Math.max(0, n))
}
const isAffine = (v: unknown): v is Affine => Array.isArray(v) && v.length === 6 && v.every(n => typeof n === 'number' && Number.isFinite(n))
const isFullBox = (m: MaskPath, w: number, h: number) =>
  m.ops.length === 1 && m.ops[0].op === 'rect' && m.ops[0].x <= 0 && m.ops[0].y <= 0 && m.ops[0].x + m.ops[0].w >= w && m.ops[0].y + m.ops[0].h >= h

export function createWebGL2Compositor(
  canvas: HTMLCanvasElement,
  workspace: SlideShowWorkspace,
  /** Additive to the brief, like S1-13's optional `sink`. */
  options: WebGL2Options = {},
): WebGL2Compositor | undefined {
  const context = canvas.getContext('webgl2', WEBGL2_CONTEXT_ATTRIBUTES) as WebGL2RenderingContext | null
  if (!context) return undefined
  const gl: WebGL2RenderingContext = context
  const sink = options.sink
  const factory: CanvasFactory = options.factory ?? ((w, h) => {
    const c = canvas.ownerDocument.createElement('canvas')
    c.width = w; c.height = h
    return c
  })

  const textures = new Map<number, Texture>()
  const seeds = new WeakMap<Rng, number>()
  let nextId = 1
  let disposed = false
  let lost = false
  let res: Resources | undefined
  let mask: MaskSurface | undefined
  let noise: NoiseSurface | undefined
  let sum: SumTarget | undefined
  let maxTextureSize = 0
  let view: Affine = IDENTITY
  let slideIndex: number | undefined

  const diagnose = (code: string, effect: string, message: string) =>
    sink?.add(slideIndex === undefined ? { code, effect, message } : { code, slideIndex, effect, message })
  const inert = () => disposed || lost || !res || gl.isContextLost()

  // -------------------------------------------------------------------------
  // Resources
  // -------------------------------------------------------------------------
  function compile(type: number, source: string): WebGLShader | undefined {
    const s = gl.createShader(type)
    if (!s) return undefined
    gl.shaderSource(s, source)
    gl.compileShader(s)
    if (!gl.getShaderParameter(s, gl.COMPILE_STATUS)) { gl.deleteShader(s); return undefined }
    return s
  }

  function applyState() {
    gl.pixelStorei(gl.UNPACK_PREMULTIPLY_ALPHA_WEBGL, true)
    gl.pixelStorei(gl.UNPACK_COLORSPACE_CONVERSION_WEBGL, gl.NONE)
    gl.pixelStorei(gl.UNPACK_FLIP_Y_WEBGL, false)
    gl.pixelStorei(gl.UNPACK_ALIGNMENT, 1)
    gl.disable(gl.DEPTH_TEST)
    gl.disable(gl.CULL_FACE)
    gl.enable(gl.BLEND)
    gl.blendFunc(gl.ONE, gl.ONE_MINUS_SRC_ALPHA)
  }

  /** Lease, then create; on any refusal or failure nothing is held. */
  function build(): boolean {
    const frame = workspace.lease(canvas.width * canvas.height * 4, 'webgl2:framebuffer')
    if (!frame) return false
    const quad = workspace.lease(QUAD.byteLength, 'webgl2:buffer')
    if (!quad) { frame.release(); return false }
    const vs = compile(gl.VERTEX_SHADER, VERTEX_SHADER)
    const fs = compile(gl.FRAGMENT_SHADER, FRAGMENT_SHADER)
    const program = vs && fs ? gl.createProgram() : null
    const fail = () => {
      if (program) gl.deleteProgram(program)
      if (vs) gl.deleteShader(vs)
      if (fs) gl.deleteShader(fs)
      frame.release(); quad.release()
      return false
    }
    if (!program || !vs || !fs) return fail()
    gl.attachShader(program, vs)
    gl.attachShader(program, fs)
    gl.linkProgram(program)
    if (!gl.getProgramParameter(program, gl.LINK_STATUS)) return fail()
    const vao = gl.createVertexArray()
    const buffer = gl.createBuffer()
    if (!vao || !buffer) {
      if (vao) gl.deleteVertexArray(vao)
      if (buffer) gl.deleteBuffer(buffer)
      return fail()
    }
    gl.bindVertexArray(vao)
    gl.bindBuffer(gl.ARRAY_BUFFER, buffer)
    gl.bufferData(gl.ARRAY_BUFFER, QUAD, gl.STATIC_DRAW)
    const loc = gl.getAttribLocation(program, 'a_unit')
    gl.enableVertexAttribArray(loc)
    gl.vertexAttribPointer(loc, 2, gl.FLOAT, false, 0, 0)
    gl.bindVertexArray(null)
    const u = {} as Record<UniformName, WebGLUniformLocation | null>
    for (const name of UNIFORMS) u[name] = gl.getUniformLocation(program, name)
    gl.useProgram(program)
    gl.uniform1i(u.u_layer, 0)
    gl.uniform1i(u.u_mask, 1)
    applyState()
    maxTextureSize = Number(gl.getParameter(gl.MAX_TEXTURE_SIZE)) || 2048
    res = { program, shaders: [vs, fs], vao, buffer, u, frame, quad, three: createMeshPool(gl, workspace) }
    return true
  }

  /** Returns every lease; deletes GL objects unless the context is already gone. */
  function teardown(deleteObjects: boolean) {
    for (const t of textures.values()) {
      if (deleteObjects) gl.deleteTexture(t.tex)
      t.lease.release()
    }
    textures.clear()
    if (mask) { if (deleteObjects) gl.deleteTexture(mask.tex); mask.lease.release(); mask = undefined }
    if (noise) { if (deleteObjects) gl.deleteTexture(noise.tex); noise.lease.release(); noise = undefined }
    dropSum(deleteObjects)
    if (res) {
      res.three.dispose(deleteObjects)
      if (deleteObjects) {
        gl.deleteVertexArray(res.vao)
        gl.deleteBuffer(res.buffer)
        gl.deleteProgram(res.program)
        for (const s of res.shaders) gl.deleteShader(s)
      }
      res.frame.release()
      res.quad.release()
      res = undefined
    }
  }

  function dropSum(deleteObjects: boolean) {
    if (!sum) return
    if (deleteObjects) { gl.deleteFramebuffer(sum.fbo); gl.deleteTexture(sum.tex) }
    sum.lease.release()
    sum = undefined
  }

  /** The output-sized sum target, leased before it is created (`webgl2:morph-sum`); undefined when refused. */
  function sumTarget(): SumTarget | undefined {
    const w = canvas.width, h = canvas.height
    if (sum && sum.w === w && sum.h === h) return sum
    dropSum(true)
    const lease = workspace.lease(w * h * 4, 'webgl2:morph-sum')
    if (!lease) return undefined
    gl.activeTexture(gl.TEXTURE0)
    const tex = newTexture(gl.NEAREST)
    const fbo = tex ? gl.createFramebuffer() : null
    if (!tex || !fbo) { if (tex) gl.deleteTexture(tex); lease.release(); return undefined }
    gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, w, h, 0, gl.RGBA, gl.UNSIGNED_BYTE, null)
    gl.bindFramebuffer(gl.FRAMEBUFFER, fbo)
    gl.framebufferTexture2D(gl.FRAMEBUFFER, gl.COLOR_ATTACHMENT0, gl.TEXTURE_2D, tex, 0)
    const ok = gl.checkFramebufferStatus(gl.FRAMEBUFFER) === gl.FRAMEBUFFER_COMPLETE
    gl.bindFramebuffer(gl.FRAMEBUFFER, null)
    if (!ok) { gl.deleteFramebuffer(fbo); gl.deleteTexture(tex); lease.release(); return undefined }
    sum = { fbo, tex, lease, w, h }
    return sum
  }

  function newTexture(filter: number): WebGLTexture | null {
    const tex = gl.createTexture()
    if (!tex) return null
    gl.bindTexture(gl.TEXTURE_2D, tex)
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, filter)
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, filter)
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE)
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE)
    return tex
  }

  // -------------------------------------------------------------------------
  // Masks
  // -------------------------------------------------------------------------
  function seedOf(rng: Rng): number {
    let seed = seeds.get(rng)
    if (seed === undefined) { seed = Math.floor(rng() * 0x100000000); seeds.set(rng, seed) }
    return seed
  }

  /** The mask canvas and its texture at `w × h`, leased together before either is created. */
  function maskSurface(w: number, h: number): MaskSurface | undefined {
    if (mask && mask.w === w && mask.h === h) return mask
    if (mask) { gl.deleteTexture(mask.tex); mask.lease.release(); mask = undefined }
    const lease = workspace.lease(w * h * 8, 'webgl2:mask')
    if (!lease) return undefined
    let c: AnyCanvas
    try { c = factory(w, h) }
    catch (error) { lease.release(); throw error }
    const tex = gl.createTexture()
    if (!tex) { lease.release(); return undefined }
    mask = { canvas: c, ctx: c.getContext('2d') as unknown as Ctx, tex, lease, w, h }
    return mask
  }

  function noiseSurface(w: number, h: number): NoiseSurface | undefined {
    if (noise && noise.w === w && noise.h === h) return noise
    if (noise) { gl.deleteTexture(noise.tex); noise.lease.release(); noise = undefined }
    const lease = workspace.lease(w * h * 4, 'webgl2:noise')
    if (!lease) return undefined
    const tex = gl.createTexture()
    if (!tex) { lease.release(); return undefined }
    noise = { tex, lease, w, h }
    return noise
  }

  function bindMaskTexture(tex: WebGLTexture, filter: number) {
    gl.activeTexture(gl.TEXTURE1)
    gl.bindTexture(gl.TEXTURE_2D, tex)
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, filter)
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, filter)
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE)
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE)
  }

  /**
   * A path mask traced into the mask canvas (`pw × ph` px; box units scaled
   * by `(sx, sy)` and offset by `(ox, oy)` px) and uploaded.
   */
  function pathMask(shape: MaskPath, key: string, pw: number, ph: number, sx: number, sy: number, ox: number, oy: number, filter: number): WebGLTexture | undefined {
    const s = maskSurface(pw, ph)
    if (!s) return undefined
    bindMaskTexture(s.tex, filter)
    if (s.key === key) return s.tex
    const c = s.ctx
    c.setTransform(1, 0, 0, 1, 0, 0)
    c.globalCompositeOperation = 'source-over'
    c.globalAlpha = 1
    c.clearRect(0, 0, pw, ph)
    c.setTransform(sx, 0, 0, sy, ox, oy)
    c.beginPath()
    tracePath(c, shape)
    c.fillStyle = '#ffffff'
    c.fill(shape.fillRule)
    c.setTransform(1, 0, 0, 1, 0, 0)
    gl.pixelStorei(gl.UNPACK_PREMULTIPLY_ALPHA_WEBGL, true)
    gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, gl.RGBA, gl.UNSIGNED_BYTE, s.canvas as TexImageSource)
    s.key = key
    return s.tex
  }

  /** A noise raster uploaded as `R8`, or undefined when a lease is refused. */
  function noiseMask(state: MaskState, box: { w: number; h: number }, seed: number, key: string): WebGLTexture | undefined {
    const temp = workspace.lease(maskRasterBytes(state.kind, state.params, box.w, box.h), 'webgl2:mask-raster')
    if (!temp) return undefined
    try {
      const r = maskPath(state.kind, state.progress, state.params, box.w, box.h, createRng(seed)) as MaskRaster
      const s = noiseSurface(r.width, r.height)
      if (!s) return undefined
      bindMaskTexture(s.tex, gl.NEAREST)
      if (s.key !== key) {
        gl.texImage2D(gl.TEXTURE_2D, 0, gl.R8, r.width, r.height, 0, gl.RED, gl.UNSIGNED_BYTE, new Uint8Array(r.alpha.buffer, r.alpha.byteOffset, r.alpha.length))
        s.key = key
      }
      return s.tex
    } finally {
      temp.release()
    }
  }

  /**
   * Resolve a mask for one quad. `layer` gives the layer-mask geometry (box =
   * bounds in slide px, rasterised at texture size); `device` the transition
   * mask's device rect. Returns the binding, `'none'` (draw unmasked),
   * `'skip'` (nothing shows), mirroring the Canvas 2D backend's choices.
   */
  function resolveMask(
    state: MaskState, seed: number, effect: string,
    geometry: { kind: 'layer'; box: Bounds; tw: number; th: number } | { kind: 'device'; box: Bounds },
  ): MaskBinding | 'none' | 'skip' {
    const box = geometry.box
    const key = `${geometry.kind}|${state.kind}|${state.progress}|${JSON.stringify(state.params)}|${box.x},${box.y},${box.w},${box.h}|${seed}`
    if (maskRasterBytes(state.kind, state.params, box.w, box.h) === 0) {
      if (maskOpCount(state.kind, state.params, box.w, box.h) > MASK_MAX_OPS) {
        diagnose('mask-ops', effect, `The ${state.kind} mask needs more than ${MASK_MAX_OPS} sub-paths at this size; it cuts at its midpoint`)
      }
      const shape = maskPath(state.kind, state.progress, state.params, box.w, box.h, createRng(seed))
      if (isMaskRaster(shape) || shape.ops.length === 0) return 'skip'
      if (isFullBox(shape, box.w, box.h)) return 'none'
      if (geometry.kind === 'layer') {
        const tex = pathMask(shape, `${key}|${geometry.tw}x${geometry.th}`, geometry.tw, geometry.th, geometry.tw / box.w, geometry.th / box.h, 0, 0, gl.LINEAR)
        return tex ? { tex, mode: 1, rect: [0, 0, 1, 1] } : cut(state, effect)
      }
      // Texels on device pixels: an integer rect around the box, the path offset into it.
      const x0 = Math.floor(box.x), y0 = Math.floor(box.y)
      const x1 = Math.max(x0 + 1, Math.ceil(box.x + box.w)), y1 = Math.max(y0 + 1, Math.ceil(box.y + box.h))
      const tex = pathMask(shape, key, x1 - x0, y1 - y0, 1, 1, box.x - x0, box.y - y0, gl.NEAREST)
      return tex ? { tex, mode: 2, rect: [x0, y0, x1 - x0, y1 - y0] } : cut(state, effect)
    }
    const tex = noiseMask(state, box, seed, key)
    if (!tex) return cut(state, effect)
    return geometry.kind === 'layer' ? { tex, mode: 1, rect: [0, 0, 1, 1] } : { tex, mode: 2, rect: [box.x, box.y, box.w, box.h] }
  }

  function cut(state: MaskState, effect: string): 'none' | 'skip' {
    diagnose('compositor-workspace', effect, `The ${state.kind} mask did not fit the slideshow workspace; it cuts at its midpoint`)
    return clamp01(state.progress, 0) >= 0.5 ? 'none' : 'skip'
  }

  // -------------------------------------------------------------------------
  // Drawing
  // -------------------------------------------------------------------------
  function drawQuad(t: Texture, m: Affine, opacity: number, tint: ColorMatrix | undefined, binding: MaskBinding | undefined) {
    const u = res!.u
    gl.activeTexture(gl.TEXTURE0)
    gl.bindTexture(gl.TEXTURE_2D, t.tex)
    gl.uniformMatrix3fv(u.u_matrix, false, quadMatrix(m, t.bounds))
    gl.uniform1f(u.u_opacity, opacity)
    if (tint) {
      const { matrix, offset } = colorUniforms(tint)
      gl.uniform1i(u.u_tint, 1)
      gl.uniformMatrix4fv(u.u_colorM, false, matrix)
      gl.uniform4fv(u.u_colorOff, offset)
    } else {
      gl.uniform1i(u.u_tint, 0)
    }
    if (binding) {
      gl.activeTexture(gl.TEXTURE1)
      gl.bindTexture(gl.TEXTURE_2D, binding.tex)
      gl.uniform1i(u.u_maskMode, binding.mode)
      gl.uniform4f(u.u_maskRect, binding.rect[0], binding.rect[1], binding.rect[2], binding.rect[3])
    } else {
      gl.uniform1i(u.u_maskMode, 0)
    }
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4)
  }

  function lookup(handle: LayerHandle, effect: string): Texture | undefined {
    const t = textures.get(handle.id)
    if (!t) diagnose('compositor-handle', effect, `Layer ${handle.id} is not uploaded; it is skipped`)
    return t
  }

  function checkedMask(state: MaskState | undefined, effect: string): MaskState | undefined {
    if (state && !isMaskKind(state.kind)) {
      diagnose('mask-unknown', effect, `Unknown mask kind ${String(state.kind)}; the layer is shown unmasked`)
      return undefined
    }
    return state
  }

  /** The axis-aligned device rect of a layer's bounds under the viewport. */
  function deviceRect(b: Bounds): Bounds {
    const [a, bb, c, d, e, f] = view
    const px: number[] = [], py: number[] = []
    for (const x of [b.x, b.x + b.w]) for (const y of [b.y, b.y + b.h]) { px.push(a * x + c * y + e); py.push(bb * x + d * y + f) }
    const x0 = Math.min(...px), y0 = Math.min(...py)
    return { x: x0, y: y0, w: Math.max(...px) - x0, h: Math.max(...py) - y0 }
  }

  /**
   * S3-1: one frame of a 3-D transition, viewport-mapped to the slide's device
   * rect `rect` (already scissored) over opaque black. Returns false (state
   * reset, diagnosed once) when the shader throws, so the look-alike draws.
   */
  function drawShader(t: Transition3D, attrs: Readonly<Record<string, string>>, from: Texture, to: Texture, progress: number, tr: TransitionDraw, rect: Bounds): boolean {
    const three = res!.three
    gl.clearColor(0, 0, 0, 1)
    gl.clear(gl.COLOR_BUFFER_BIT)
    gl.viewport(rect.x, canvas.height - (rect.y + rect.h), rect.w, rect.h)
    three.setAspect(to.bounds.w / to.bounds.h)
    let ok = true
    try {
      t.draw(gl, from.tex, to.tex, progress, attrs, createRng(seedOf(tr.rng)), three.pool)
    } catch (error) {
      ok = false
      const why = error instanceof Error ? error.message : String(error)
      diagnose('transition-shader-fallback', tr.effect, `The ${tr.effect} 3-D shader could not draw (${why}); it draws its 2-D look-alike`)
    }
    resetState(rect)
    return ok
  }

  /** Every GL state a 3-D draw may have changed, back to the compositor's own (scissored to the slide rect `rect`). */
  function resetState(rect: Bounds) {
    gl.bindFramebuffer(gl.FRAMEBUFFER, null)
    gl.viewport(0, 0, canvas.width, canvas.height)
    gl.useProgram(res!.program)
    gl.bindVertexArray(res!.vao)
    gl.activeTexture(gl.TEXTURE0)
    gl.colorMask(true, true, true, true)
    gl.frontFace(gl.CCW)
    gl.blendEquation(gl.FUNC_ADD)
    gl.clearColor(0, 0, 0, 0)
    applyState()
    const x0 = Math.max(0, rect.x), y0 = Math.max(0, rect.y)
    const x1 = Math.min(canvas.width, rect.x + rect.w), y1 = Math.min(canvas.height, rect.y + rect.h)
    gl.enable(gl.SCISSOR_TEST)
    gl.scissor(x0, canvas.height - y1, Math.max(0, x1 - x0), Math.max(0, y1 - y0))
  }

  function drawTransition(tr: TransitionDraw) {
    const params = tr.params as TransitionParams
    const fit = (t: Texture, a: unknown): Affine => {
      const b = t.bounds
      const unit: Affine = [1 / b.w, 0, 0, 1 / b.h, -b.x / b.w, -b.y / b.h]
      const back: Affine = [b.w, 0, 0, b.h, b.x, b.y]
      return mulAffine(view, mulAffine(back, mulAffine(isAffine(a) ? a : IDENTITY, unit)))
    }
    const from = lookup(tr.from, tr.effect), to = lookup(tr.to, tr.effect)
    const backdrop = params.backdrop === 'black' || params.backdrop === 'white' ? BACKDROP_RGB[params.backdrop] : undefined
    const slide = to ?? from
    if (!slide) return
    const r = transitionRect(deviceRect(slide.bounds))
    const x0 = Math.max(0, r.x), y0 = Math.max(0, r.y)
    const x1 = Math.min(canvas.width, r.x + r.w), y1 = Math.min(canvas.height, r.y + r.h)
    if (!(x1 > x0 && y1 > y0)) return
    gl.enable(gl.SCISSOR_TEST)
    // GL's window origin is the bottom-left corner.
    gl.scissor(x0, canvas.height - y1, x1 - x0, y1 - y0)
    // S3-1: a 3-D row dispatched to its shader. At the ends it is exactly the flat outgoing / incoming frame.
    const ref = shader3DOf(tr.params)
    const progress = clamp01(tr.progress, 0)
    if (ref && from && to && (progress <= 0 || progress >= 1)) {
      const end = progress <= 0 ? from : to
      drawQuad(end, fit(end, IDENTITY), 1, undefined, undefined)
      gl.disable(gl.SCISSOR_TEST)
      return
    }
    if (ref && from && to) {
      const t = transition3DFor(ref)
      if (t && drawShader(t, ref.attrs, from, to, progress, tr, r)) { gl.disable(gl.SCISSOR_TEST); return }
      if (!t) diagnose('transition-shader-fallback', tr.effect, `No 3-D shader is registered for ${ref.element}; it draws its 2-D look-alike`)
      // The look-alike draws instead, on a clean slide rect.
      gl.clearColor(0, 0, 0, 0)
      gl.clear(gl.COLOR_BUFFER_BIT)
    }
    if (backdrop) {
      gl.clearColor(backdrop[0] / 255, backdrop[1] / 255, backdrop[2] / 255, 1)
      gl.clear(gl.COLOR_BUFFER_BIT)
    }
    const drawFrom = () => {
      if (!from) return
      const alpha = clamp01(params.fromOpacity, 1)
      if (alpha > 0) drawQuad(from, fit(from, params.from), alpha, undefined, undefined)
    }
    const drawTo = () => {
      if (!to) return
      const alpha = clamp01(params.toOpacity, 1)
      if (alpha <= 0) return
      const m = checkedMask(params.mask, tr.effect)
      let binding: MaskBinding | undefined
      if (m) {
        const r = resolveMask(m, seedOf(tr.rng), tr.effect, { kind: 'device', box: deviceRect(to.bounds) })
        if (r === 'skip') return
        if (r !== 'none') binding = r
      }
      drawQuad(to, fit(to, params.to), alpha, undefined, binding)
    }
    if (params.fromOnTop) { drawTo(); drawFrom() } else { drawFrom(); drawTo() }
    gl.disable(gl.SCISSOR_TEST)
  }

  /** A box with positive width and height from a corner and a signed size. */
  function normBox(x: number, y: number, w: number, h: number): Bounds {
    return { x: w < 0 ? x + w : x, y: h < 0 ? y + h : y, w: Math.abs(w), h: Math.abs(h) }
  }

  /** An unmasked draw's texture, tint and device matrix; undefined when it draws nothing. */
  function resolveDraw(d: LayerDraw, effect: string): { t: Texture; tint: ColorMatrix | undefined; m: Affine } | undefined {
    const { state } = d
    if (!state.visible || !(state.opacity > 0)) return undefined
    const t = lookup(d.handle, effect)
    if (!t) return undefined
    const matrix = d.part ? state.color?.[d.part] : undefined
    const tint = matrix && matrix.length >= 20 ? matrix : undefined
    const transform = isAffine(state.transform) ? state.transform : IDENTITY
    const local = state.motion ? mulAffine([1, 0, 0, 1, state.motion.dx, state.motion.dy], transform) : transform
    return { t, tint, m: mulAffine(view, local) }
  }

  /** The integer device rect covering the parts, clipped to the output; undefined when empty. */
  function coverRect(parts: ReadonlyArray<{ bounds: Bounds; m: Affine }>): Bounds | undefined {
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity
    for (const { bounds: b, m } of parts) {
      for (const x of [b.x, b.x + b.w]) for (const y of [b.y, b.y + b.h]) {
        const px = m[0] * x + m[2] * y + m[4], py = m[1] * x + m[3] * y + m[5]
        x0 = Math.min(x0, px); y0 = Math.min(y0, py); x1 = Math.max(x1, px); y1 = Math.max(y1, py)
      }
    }
    const l = Math.max(0, Math.floor(x0)), t = Math.max(0, Math.floor(y0))
    const r = Math.min(canvas.width, Math.ceil(x1)), b = Math.min(canvas.height, Math.ceil(y1))
    return r > l && b > t ? { x: l, y: t, w: r - l, h: b - t } : undefined
  }

  /**
   * A Morph pair summed alpha-correctly (`LayerDraw.sum`, PR #20 review 2),
   * as the Canvas 2D backend does: both sides drawn additively
   * (`ONE, ONE`) at their weights into the sum target, scissored to the
   * device rect they cover, then that rect composites source-over
   * (`ONE, ONE_MINUS_SRC_ALPHA`). A masked side, or a refused target
   * (diagnosed once), holds the outgoing side under instead.
   */
  function drawSum(d: LayerDraw, s: NonNullable<LayerDraw['sum']>) {
    const held = () => {
      drawOne({ handle: d.handle, state: { ...d.state, opacity: s.heldOpacity }, ...(d.part ? { part: d.part } : {}) })
      drawOne(s.with)
    }
    if (d.state.mask || s.with.state.mask) return held()
    const a = resolveDraw(d, 'layer'), b = resolveDraw(s.with, 'layer')
    if (!a || !b) return held()
    const r = coverRect([{ bounds: a.t.bounds, m: a.m }, { bounds: b.t.bounds, m: b.m }])
    if (!r) return
    const target = sumTarget()
    if (!target) {
      diagnose('compositor-workspace', 'morph-blend', 'The Morph blend surface did not fit the slideshow workspace; matched shapes hold the outgoing side under (a translucent pair can pop on its last frame)')
      return held()
    }
    const H = canvas.height
    gl.bindFramebuffer(gl.FRAMEBUFFER, target.fbo)
    gl.enable(gl.SCISSOR_TEST)
    gl.scissor(r.x, H - r.y - r.h, r.w, r.h)
    gl.clearColor(0, 0, 0, 0)
    gl.clear(gl.COLOR_BUFFER_BIT)
    gl.blendFunc(gl.ONE, gl.ONE)
    drawQuad(a.t, a.m, clamp01(d.state.opacity, 1), a.tint, undefined)
    drawQuad(b.t, b.m, clamp01(s.with.state.opacity, 1), b.tint, undefined)
    gl.bindFramebuffer(gl.FRAMEBUFFER, null)
    gl.blendFunc(gl.ONE, gl.ONE_MINUS_SRC_ALPHA)
    // The target's rows run bottom-up (GL): the quad samples it flipped onto the output rect.
    drawQuad({ tex: target.tex, bounds: { x: 0, y: H, w: canvas.width, h: -H }, lease: target.lease, w: target.w, h: target.h }, IDENTITY, 1, undefined, undefined)
    gl.disable(gl.SCISSOR_TEST)
  }

  function drawOne(d: LayerDraw) {
    if (d.sum) return drawSum(d, d.sum)
    const { state } = d
    const effect = state.mask ? `mask:${state.mask.kind}` : 'layer'
    const resolved = resolveDraw(d, effect)
    if (!resolved) return
    const { t, tint, m } = resolved
    let binding: MaskBinding | undefined
    const ms = checkedMask(state.mask, effect)
    if (ms) {
      const seed = Number(ms.params.seed)
      // A mask over another box than the layer's (a split text item under its shape's mask): rasterised over that
      // box at the layer's texel density (capped), and sampled through the layer's place inside it.
      const b = t.bounds, mb = state.maskBox
      const own = !(mb && [mb.x, mb.y, mb.w, mb.h].every(Number.isFinite) && mb.w > 0 && mb.h > 0 && b.w > 0 && b.h > 0)
      const box = own ? b : mb!
      const k = own ? 1 : Math.min(t.w / b.w, t.h / b.h, maxTextureSize / box.w, maxTextureSize / box.h)
      const tw = own ? t.w : Math.max(1, Math.round(box.w * k)), th = own ? t.h : Math.max(1, Math.round(box.h * k))
      // A shared parent mask placed by the parent's own mapping (PR #19 review 7): fixed in device space when that
      // mapping is axis-aligned, so it does not follow the layer's transform; otherwise it moves with the layer.
      const pm = own || !isAffine(state.maskTransform) ? undefined : mulAffine(view, state.maskTransform)
      const device = pm && Math.abs(pm[1]) < 1e-9 && Math.abs(pm[2]) < 1e-9 && pm[0] !== 0 && pm[3] !== 0
        ? normBox(pm[0] * box.x + pm[4], pm[3] * box.y + pm[5], pm[0] * box.w, pm[3] * box.h) : undefined
      const r = device
        ? resolveMask(ms, Number.isFinite(seed) ? seed : 0, effect, { kind: 'device', box: device })
        : resolveMask(ms, Number.isFinite(seed) ? seed : 0, effect, { kind: 'layer', box, tw, th })
      if (r === 'skip') return
      if (r !== 'none') binding = own || device ? r : { ...r, rect: [(box.x - b.x) / b.w, (box.y - b.y) / b.h, box.w / b.w, box.h / b.h] }
    }
    drawQuad(t, m, clamp01(state.opacity, 1), tint, binding)
  }

  /** False when the new drawing buffer did not fit: the output keeps its size. */
  function resizeOutput(widthPx: number, heightPx: number): boolean {
    const w = Math.max(1, Math.floor(widthPx)), h = Math.max(1, Math.floor(heightPx))
    if (!Number.isFinite(w) || !Number.isFinite(h) || (w === canvas.width && h === canvas.height)) return true
    if (res) {
      const next = workspace.lease(w * h * 4, 'webgl2:framebuffer')
      if (!next) {
        diagnose('compositor-workspace', 'output', `The ${w}×${h} WebGL2 drawing buffer did not fit the slideshow workspace; the output keeps its size`)
        return false
      }
      res.frame.release()
      res.frame = next
    }
    canvas.width = w
    canvas.height = h
    // The output-sized Morph sum target is re-leased at the new size on next use.
    dropSum(!lost && !gl.isContextLost())
    return true
  }

  if (!build()) return undefined

  return {
    backend: 'webgl2',
    get maxTextureSize() { return maxTextureSize },

    upload(layer, lease) {
      if (inert()) return { refused: 'workspace' }
      const w = layer.canvas.width, h = layer.canvas.height
      if (!(w >= 1 && h >= 1) || w > maxTextureSize || h > maxTextureSize) return { refused: 'texture-size' }
      const texBytes = w * h * 4
      if (!lease.grow(texBytes)) return { refused: 'workspace' }
      gl.activeTexture(gl.TEXTURE0)
      const tex = newTexture(gl.LINEAR)
      if (!tex) { lease.shrink(texBytes); return { refused: 'workspace' } }
      gl.pixelStorei(gl.UNPACK_PREMULTIPLY_ALPHA_WEBGL, true)
      gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, gl.RGBA, gl.UNSIGNED_BYTE, layer.canvas as TexImageSource)
      // The CPU copy is released after upload: its bytes leave the lease, the texture's stay.
      lease.shrink(Math.min(Math.max(0, layer.bytes), lease.bytes - texBytes))
      const id = nextId++
      textures.set(id, { tex, bounds: { ...layer.bounds }, lease, w, h })
      return { id }
    },

    release(handle) {
      const t = textures.get(handle.id)
      if (!t) return
      textures.delete(handle.id)
      if (!lost) gl.deleteTexture(t.tex)
      t.lease.release()
    },

    frame(draws, transition) {
      if (inert()) return
      gl.bindFramebuffer(gl.FRAMEBUFFER, null)
      gl.viewport(0, 0, canvas.width, canvas.height)
      gl.clearColor(0, 0, 0, 0)
      gl.clear(gl.COLOR_BUFFER_BIT)
      gl.useProgram(res!.program)
      gl.bindVertexArray(res!.vao)
      gl.enable(gl.BLEND)
      gl.blendFunc(gl.ONE, gl.ONE_MINUS_SRC_ALPHA)
      gl.uniform2f(res!.u.u_output, canvas.width, canvas.height)
      if (transition) drawTransition(transition)
      for (const d of draws) drawOne(d)
      gl.bindVertexArray(null)
    },

    resize(widthPx, heightPx) {
      if (!disposed) resizeOutput(widthPx, heightPx)
    },

    setViewport(next: CompositorViewport) {
      if (disposed) return
      view = isAffine(next.matrix) ? next.matrix : IDENTITY
      return resizeOutput(next.deviceWidth, next.deviceHeight)
    },

    setSlideIndex(index) {
      slideIndex = index
    },

    prepareTransition3D(t) {
      if (inert()) return false
      try {
        t.prepare?.(res!.three.pool)
        return true
      } catch {
        return false
      } finally {
        gl.bindVertexArray(null)
      }
    },

    contextLost() {
      if (disposed || lost) return
      lost = true
      teardown(false)
    },

    contextRestored() {
      if (disposed) return false
      teardown(false)
      if (gl.isContextLost() || !build()) { lost = true; return false }
      lost = false
      return true
    },

    dispose() {
      if (disposed) return
      teardown(!lost && !gl.isContextLost())
      disposed = true
    },
  }
}
