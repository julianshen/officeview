/**
 * The Canvas 2D compositor (slideshow spec 6.3 fallback, 7): the same
 * `Compositor` contract as the WebGL2 backend, drawn with 2-D transforms,
 * `globalAlpha`, clip paths and pre-rendered masks. It never uses
 * `ctx.filter`, which is not available everywhere; colour matrices run on
 * pixels (`applyColorMatrix`).
 *
 * Conventions (shared with the WebGL2 backend, S1-14):
 * - **Layer space.** A `LayerBitmap`'s `bounds` are in slide px; its canvas
 *   is `bounds × scale` device px (S1-12). A reduced-scale layer is drawn at
 *   its full bounds, so it scales up.
 * - **`LayerDraw`.** `state.transform` maps slide px to output device px,
 *   the view (letterbox and backing-store scale) included. `state.motion` is
 *   a further translation in output device px, applied after the transform.
 *   `state.mask` covers the layer's `bounds` box in slide px and moves with
 *   the transform. Draws composite in array order (z-order), source-over on
 *   premultiplied colour (WebGL2's `ONE, ONE_MINUS_SRC_ALPHA`).
 * - **Colour.** A draw with `part` gets `state.color[part]`; a whole layer
 *   (no `part`) ignores `state.color`, because an unsplit target is
 *   re-rasterised instead (spec 6.3, part layers).
 * - **`TransitionDraw`.** `params` may carry `mask` (a `MaskState`), `from`
 *   and `to` (`Affine`s in unit slide space: the layer's bounds map to
 *   [0, 1]², the affine applies there, then [0, 1]² fills the output), and
 *   `fromOpacity` / `toOpacity` (default 1). The outgoing frame is drawn,
 *   then the incoming one through the mask, which is in output device px.
 *   The transition composites first and `draws` go on top (the watermark is a
 *   top layer no effect moves).
 * - **Noise determinism.** A transition's `rng` is stateful; one seed is drawn
 *   from it once and every frame re-seeds from that, so a dissolve keeps its
 *   pattern. A layer mask (no rng) uses `params.seed` (default 0).
 * - **Workspace.** `upload` grows the caller's lease by the texture before
 *   creating it (CPU copy and texture both charged during upload, spec 7),
 *   then shrinks it by the CPU copy's `bytes`: on success the lease belongs to
 *   the compositor and charges only the texture, and the caller drops
 *   `layer.canvas`. A refusal leaves the lease untouched and creates nothing.
 *   Scratch, mask and recolour surfaces, and the buffers behind them, are
 *   leased before they are created and released on `release`, `resize` or
 *   `dispose`. A refused lease degrades the effect (a raster mask becomes a
 *   cut at its midpoint; a recolour is skipped) with a diagnostic, never
 *   silently.
 */
import type {
  Affine, CanvasFactory, ColorMatrix, Compositor, DiagnosticSink, LayerBitmap, LayerDraw, LayerHandle, Lease, MaskState,
  Rng, SlideShowWorkspace, TransitionDraw,
} from '../types'
import { createRng } from '../rng'
import { isMaskKind, isMaskRaster, maskPath, maskRasterBytes, tracePath, type MaskRaster } from './masks'

/**
 * The largest texture side the Canvas 2D backend accepts. 8192 is the
 * conservative side for canvases across browsers (iOS Safari caps canvas
 * area at about 16.7 M px) and matches SwiftShader's WebGL2 limit, so a layer
 * that fits one backend fits the other.
 */
export const CANVAS2D_MAX_TEXTURE_SIZE = 8192

type Ctx = CanvasRenderingContext2D
type AnyCanvas = HTMLCanvasElement | OffscreenCanvas

interface Surface { canvas: AnyCanvas; ctx: Ctx; lease: Lease; w: number; h: number }
interface Texture {
  canvas: AnyCanvas
  bounds: LayerBitmap['bounds']
  lease: Lease
  tint?: { key: string; surface: Surface }
}
interface TransitionParams { mask?: MaskState; from?: Affine; to?: Affine; fromOpacity?: number; toOpacity?: number }

const IDENTITY: Affine = [1, 0, 0, 1, 0, 0]

/** `m1 ∘ m2`: apply `m2`, then `m1`. */
function mul(m1: Affine, m2: Affine): Affine {
  const [a, b, c, d, e, f] = m1, [A, B, C, D, E, F] = m2
  return [a * A + c * B, b * A + d * B, a * C + c * D, b * C + d * D, a * E + c * F + e, b * E + d * F + f]
}

const clamp01 = (v: unknown, fallback: number) => {
  const n = typeof v === 'number' && Number.isFinite(v) ? v : fallback
  return Math.min(1, Math.max(0, n))
}
const isAffine = (v: unknown): v is Affine => Array.isArray(v) && v.length === 6 && v.every(n => typeof n === 'number' && Number.isFinite(n))

/**
 * A 4×5 row-major colour matrix on un-premultiplied colour, values and
 * offsets normalised to [0, 1] (the SVG `feColorMatrix` convention), clamped.
 * `getImageData` already returns un-premultiplied colour. Returns a new
 * `ImageData` of the same realm as `src`.
 */
export function applyColorMatrix(src: ImageData, m: ColorMatrix): ImageData {
  const s = src.data
  const d = new Uint8ClampedArray(s.length)
  const k = (i: number) => (Number.isFinite(m[i]) ? m[i] : 0)
  const r0 = k(0), r1 = k(1), r2 = k(2), r3 = k(3), r4 = k(4) * 255
  const g0 = k(5), g1 = k(6), g2 = k(7), g3 = k(8), g4 = k(9) * 255
  const b0 = k(10), b1 = k(11), b2 = k(12), b3 = k(13), b4 = k(14) * 255
  const a0 = k(15), a1 = k(16), a2 = k(17), a3 = k(18), a4 = k(19) * 255
  for (let i = 0; i < s.length; i += 4) {
    const r = s[i], g = s[i + 1], b = s[i + 2], a = s[i + 3]
    // Uint8ClampedArray rounds half to even; round half up so 127.5 → 128 like the GPU path.
    d[i] = Math.floor(r0 * r + r1 * g + r2 * b + r3 * a + r4 + 0.5)
    d[i + 1] = Math.floor(g0 * r + g1 * g + g2 * b + g3 * a + g4 + 0.5)
    d[i + 2] = Math.floor(b0 * r + b1 * g + b2 * b + b3 * a + b4 + 0.5)
    d[i + 3] = Math.floor(a0 * r + a1 * g + a2 * b + a3 * a + a4 + 0.5)
  }
  const Ctor = src.constructor as new (data: Uint8ClampedArray, w: number, h: number) => ImageData
  return new Ctor(d, src.width, src.height)
}

export function createCanvas2DCompositor(
  canvas: AnyCanvas,
  factory: CanvasFactory,
  workspace: SlideShowWorkspace,
  /** Additive to the brief: where degraded-effect diagnostics go. */
  sink?: DiagnosticSink,
): Compositor {
  const out = canvas.getContext('2d') as unknown as Ctx
  const textures = new Map<number, Texture>()
  const seeds = new WeakMap<Rng, number>()
  let nextId = 1
  let disposed = false
  let scratch: Surface | undefined
  let mask: (Surface & { image: ImageData }) | undefined

  const diagnose = (code: string, effect: string, message: string) => sink?.add({ code, effect, message })

  /** Lease, then create; nothing is created when the lease is refused. */
  function surface(w: number, h: number, bytes: number, tag: string): Surface | undefined {
    const lease = workspace.lease(bytes, tag)
    if (!lease) return undefined
    let c: AnyCanvas
    try { c = factory(w, h) }
    catch (error) { lease.release(); throw error }
    return { canvas: c, ctx: c.getContext('2d') as unknown as Ctx, lease, w, h }
  }

  function scratchSurface(): Surface | undefined {
    const w = canvas.width, h = canvas.height
    if (scratch && scratch.w === w && scratch.h === h) return scratch
    scratch?.lease.release()
    scratch = surface(w, h, w * h * 4, 'canvas2d:scratch')
    return scratch
  }

  /** The mask surface plus its reusable `ImageData`, both charged. */
  function maskSurface(w: number, h: number) {
    if (mask && mask.w === w && mask.h === h) return mask
    mask?.lease.release()
    mask = undefined
    const s = surface(w, h, w * h * 8, 'canvas2d:mask')
    if (!s) return undefined
    mask = { ...s, image: s.ctx.createImageData(w, h) }
    return mask
  }

  /** The texture recoloured by `m`, cached per texture and matrix; undefined when refused. */
  function tinted(t: Texture, m: ColorMatrix): AnyCanvas | undefined {
    const key = m.join(',')
    if (t.tint?.key === key) return t.tint.surface.canvas
    const w = t.canvas.width, h = t.canvas.height
    // The canvas, the read-back and the recoloured copy coexist while recolouring.
    const s = t.tint?.surface ?? surface(w, h, w * h * 12, 'canvas2d:tint')
    if (!s) return undefined
    s.ctx.setTransform(1, 0, 0, 1, 0, 0)
    s.ctx.globalCompositeOperation = 'copy'
    s.ctx.drawImage(t.canvas as CanvasImageSource, 0, 0)
    s.ctx.globalCompositeOperation = 'source-over'
    s.ctx.putImageData(applyColorMatrix(s.ctx.getImageData(0, 0, w, h), m), 0, 0)
    t.tint = { key, surface: s }
    return s.canvas
  }

  function seedOf(rng: Rng): number {
    let seed = seeds.get(rng)
    if (seed === undefined) { seed = Math.floor(rng() * 0x100000000); seeds.set(rng, seed) }
    return seed
  }

  function place(ctx: Ctx, m: Affine) { ctx.setTransform(m[0], m[1], m[2], m[3], m[4], m[5]) }

  /** The mask's raster in the mask surface, or undefined when refused. */
  function rasterSurface(state: MaskState, box: { w: number; h: number }, seed: number) {
    const bytes = maskRasterBytes(state.kind, state.params, box.w, box.h)
    const temp = workspace.lease(bytes, 'canvas2d:mask-raster')
    if (!temp) return undefined
    try {
      const r = maskPath(state.kind, state.progress, state.params, box.w, box.h, createRng(seed)) as MaskRaster
      const s = maskSurface(r.width, r.height)
      if (!s) return undefined
      const data = s.image.data
      for (let i = 0; i < r.alpha.length; i++) {
        data[i * 4] = 0; data[i * 4 + 1] = 0; data[i * 4 + 2] = 0; data[i * 4 + 3] = r.alpha[i]
      }
      s.ctx.putImageData(s.image, 0, 0)
      return s
    } finally {
      temp.release()
    }
  }

  /**
   * Draw `source` at `bounds` under `m`. A mask covers `box` under `maskM`
   * (the layer's box and matrix for a layer mask; the output for a transition).
   */
  function drawLayer(
    source: AnyCanvas, bounds: LayerBitmap['bounds'], m: Affine, opacity: number,
    maskState: MaskState | undefined, maskM: Affine, box: LayerBitmap['bounds'], seed: number, effect: string,
  ) {
    const alpha = clamp01(opacity, 1)
    if (alpha <= 0) return
    const img = source as CanvasImageSource
    if (maskState && !isMaskKind(maskState.kind)) {
      diagnose('mask-unknown', effect, `Unknown mask kind ${String(maskState.kind)}; the layer is shown unmasked`)
      maskState = undefined
    }
    if (!maskState) {
      out.globalAlpha = alpha
      place(out, m)
      out.drawImage(img, bounds.x, bounds.y, bounds.w, bounds.h)
      return
    }
    if (maskRasterBytes(maskState.kind, maskState.params, box.w, box.h) === 0) {
      const shape = maskPath(maskState.kind, maskState.progress, maskState.params, box.w, box.h, createRng(seed))
      if (isMaskRaster(shape) || shape.ops.length === 0) return
      out.save()
      place(out, maskM)
      out.beginPath()
      tracePath(out, shape, box.x, box.y)
      out.clip(shape.fillRule)
      out.globalAlpha = alpha
      place(out, m)
      out.drawImage(img, bounds.x, bounds.y, bounds.w, bounds.h)
      out.restore()
      return
    }
    // Noise: layer into the scratch surface, keep it where the raster is opaque, then composite.
    const s = scratchSurface()
    const r = s && rasterSurface(maskState, box, seed)
    if (!s || !r) {
      diagnose('compositor-workspace', effect, `The ${maskState.kind} mask did not fit the slideshow workspace; it cuts at its midpoint`)
      if (clamp01(maskState.progress, 0) >= 0.5) {
        out.globalAlpha = alpha
        place(out, m)
        out.drawImage(img, bounds.x, bounds.y, bounds.w, bounds.h)
      }
      return
    }
    const c = s.ctx
    c.setTransform(1, 0, 0, 1, 0, 0)
    c.globalAlpha = 1
    c.globalCompositeOperation = 'source-over'
    c.clearRect(0, 0, s.w, s.h)
    place(c, m)
    c.drawImage(img, bounds.x, bounds.y, bounds.w, bounds.h)
    c.globalCompositeOperation = 'destination-in'
    c.imageSmoothingEnabled = false
    place(c, maskM)
    c.drawImage(r.canvas as CanvasImageSource, box.x, box.y, box.w, box.h)
    c.imageSmoothingEnabled = true
    c.globalCompositeOperation = 'source-over'
    out.globalAlpha = alpha
    out.setTransform(1, 0, 0, 1, 0, 0)
    out.drawImage(s.canvas as CanvasImageSource, 0, 0)
  }

  function lookup(handle: LayerHandle, effect: string): Texture | undefined {
    const t = textures.get(handle.id)
    if (!t) diagnose('compositor-handle', effect, `Layer ${handle.id} is not uploaded; it is skipped`)
    return t
  }

  function drawTransition(tr: TransitionDraw) {
    const params = tr.params as TransitionParams
    const W = canvas.width, H = canvas.height
    const fit = (t: Texture, a: unknown): Affine => {
      const b = t.bounds
      const unit: Affine = [1 / b.w, 0, 0, 1 / b.h, -b.x / b.w, -b.y / b.h]
      return mul([W, 0, 0, H, 0, 0], mul(isAffine(a) ? a : IDENTITY, unit))
    }
    const from = lookup(tr.from, tr.effect), to = lookup(tr.to, tr.effect)
    if (from) drawLayer(from.canvas, from.bounds, fit(from, params.from), clamp01(params.fromOpacity, 1), undefined, IDENTITY, from.bounds, 0, tr.effect)
    if (to) {
      const seed = params.mask ? seedOf(tr.rng) : 0
      drawLayer(to.canvas, to.bounds, fit(to, params.to), clamp01(params.toOpacity, 1), params.mask, IDENTITY, { x: 0, y: 0, w: W, h: H }, seed, tr.effect)
    }
  }

  function drawOne(d: LayerDraw) {
    const { state } = d
    if (!state.visible || !(state.opacity > 0)) return
    const effect = state.mask ? `mask:${state.mask.kind}` : 'layer'
    const t = lookup(d.handle, effect)
    if (!t) return
    let source: AnyCanvas = t.canvas
    const matrix = d.part ? state.color?.[d.part] : undefined
    if (matrix && matrix.length >= 20) {
      const tint = tinted(t, matrix)
      if (tint) source = tint
      else diagnose('compositor-workspace', `colour:${d.part}`, 'A colour transform did not fit the slideshow workspace; the layer keeps its colour')
    }
    const transform = isAffine(state.transform) ? state.transform : IDENTITY
    const m = state.motion ? mul([1, 0, 0, 1, state.motion.dx, state.motion.dy], transform) : transform
    const seed = Number(state.mask?.params.seed)
    drawLayer(source, t.bounds, m, state.opacity, state.mask, m, t.bounds, Number.isFinite(seed) ? seed : 0, effect)
  }

  function dropTexture(t: Texture) {
    t.tint?.surface.lease.release()
    t.tint = undefined
    t.lease.release()
  }

  return {
    backend: 'canvas2d',
    maxTextureSize: CANVAS2D_MAX_TEXTURE_SIZE,

    upload(layer, lease) {
      if (disposed) return { refused: 'workspace' }
      const w = layer.canvas.width, h = layer.canvas.height
      if (!(w >= 1 && h >= 1) || w > CANVAS2D_MAX_TEXTURE_SIZE || h > CANVAS2D_MAX_TEXTURE_SIZE) return { refused: 'texture-size' }
      const texBytes = w * h * 4
      if (!lease.grow(texBytes)) return { refused: 'workspace' }
      let tex: AnyCanvas
      try { tex = factory(w, h) }
      catch (error) { lease.shrink(texBytes); throw error }
      const c = tex.getContext('2d') as unknown as Ctx
      c.drawImage(layer.canvas as CanvasImageSource, 0, 0)
      // The CPU copy is released after upload: its bytes leave the lease, the texture's stay.
      lease.shrink(Math.min(Math.max(0, layer.bytes), lease.bytes - texBytes))
      const id = nextId++
      textures.set(id, { canvas: tex, bounds: { ...layer.bounds }, lease })
      return { id }
    },

    release(handle) {
      const t = textures.get(handle.id)
      if (!t) return
      textures.delete(handle.id)
      dropTexture(t)
    },

    frame(draws, transition) {
      if (disposed) return
      out.setTransform(1, 0, 0, 1, 0, 0)
      out.globalAlpha = 1
      out.globalCompositeOperation = 'source-over'
      out.clearRect(0, 0, canvas.width, canvas.height)
      if (transition) drawTransition(transition)
      for (const d of draws) drawOne(d)
      out.setTransform(1, 0, 0, 1, 0, 0)
      out.globalAlpha = 1
    },

    resize(widthPx, heightPx) {
      if (disposed) return
      canvas.width = Math.max(1, Math.floor(widthPx))
      canvas.height = Math.max(1, Math.floor(heightPx))
      // Output-sized scratch is re-leased at the new size on next use.
      scratch?.lease.release()
      scratch = undefined
    },

    dispose() {
      if (disposed) return
      disposed = true
      for (const t of textures.values()) dropTexture(t)
      textures.clear()
      scratch?.lease.release()
      mask?.lease.release()
      scratch = undefined
      mask = undefined
    },
  }
}
