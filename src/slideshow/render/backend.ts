/**
 * Backend selection and WebGL2 context-loss recovery (slideshow spec 6.3
 * "Backends" and "Context loss", round 2; spec 7 permanent reserve).
 *
 * `createBackendManager` mounts a WebGL2 canvas in `host` and composites on
 * it. Without WebGL2 (or with `force: 'canvas2d'`) it starts on the Canvas 2D
 * compositor instead, with a diagnostic, on `deps.output`. `current` is one stable
 * `Compositor` that delegates to whichever backend is active, remembers the
 * last viewport and slide index, and replays them onto a new backend.
 *
 * On `webglcontextlost`:
 * 1. the handler calls `preventDefault()` (without it the context is never
 *    restored);
 * 2. in the same task, `onLoss` hides the WebGL canvas and paints on
 *    `deps.output`, the terminal reserve's output canvas, which the
 *    permanent reserve already pays for, so it needs no admission and no
 *    second full-size canvas exists (PR #18 review P2). A 2-D context
 *    cannot share the WebGL canvas, but the output is a canvas of its own.
 *    While the context is lost every
 *    `current.frame()` paints live state on it through `deps.paintLive`,
 *    under the viewport (slide px in, as from `createLivePainter`). The show
 *    clock and media are not touched, so they keep running;
 * 3. restored within `RESTORE_DEADLINE_MS` of `deps.now()` (the monotonic
 *    `performance.now`, never the show clock, so pausing does not stretch
 *    it): every GPU resource is recreated and re-leased, `rebuildLayers`
 *    re-rasterises and re-uploads, and WebGL2 takes over when that resolves.
 *    A refused lease or failed rebuild drops to Canvas 2D with a diagnostic;
 * 4. not restored by the deadline: the Canvas 2D compositor takes over on the
 *    output canvas once `rebuildLayers` has built its layers, and the
 *    WebGL canvas is removed at once;
 * 5. a loss beyond `MAX_CONTEXT_LOSSES` pins Canvas 2D for the rest of the
 *    show;
 * 6. restores (and rebuild results) that arrive after a switch, a newer loss
 *    or `dispose` are ignored.
 *
 * Handles do not survive a switch: after one, draw only the handles that
 * `rebuildLayers` uploaded to the compositor it was given. While the context
 * is lost, `current.upload` refuses with `'workspace'` without touching the
 * lease (the caller keeps and releases it); `state` tells that apart from
 * pressure. The manager never releases `terminal`: its owner does.
 */
import type { PptxSlide } from '../../pptx/types'
import type {
  Affine, CanvasFactory, Compositor, CompositorViewport, DiagnosticSink, RenderState, SlideShowWorkspace, TerminalReserve,
} from '../types'
import { CANVAS2D_MAX_TEXTURE_SIZE, createCanvas2DCompositor } from './canvas2d'
import { paintDirect, type DirectAssets } from './direct'
import { mulAffine } from './shaders'
import { createWebGL2Compositor, type WebGL2Compositor } from './webgl2'
import { registerTransitions3D } from './three'

/** Context-restore deadline (spec 6.3), on `deps.now()`. */
export const RESTORE_DEADLINE_MS = 3000
/** Losses tolerated per show; the next one pins Canvas 2D. */
export const MAX_CONTEXT_LOSSES = 2

export interface BackendDeps {
  /** Monotonic milliseconds: `performance.now` in production. */
  now: () => number
  schedule: (cb: () => void, ms: number) => unknown
  /** Paints the live state, in slide px, onto the loss-time 2-D canvas (see `createLivePainter`). */
  paintLive: (ctx: CanvasRenderingContext2D) => void
  /** Re-rasterise and re-upload the current slide's layers to `c`; the switch to `c` waits for it. */
  rebuildLayers: (c: Compositor) => Promise<void>
  workspace: SlideShowWorkspace
  terminal: TerminalReserve
  /**
   * The terminal reserve's output canvas (2-D only, never given a WebGL context). The Canvas 2D compositor and
   * the loss-time painter draw on it, so the fallback holds no surface of its own. Its owner sizes it; the
   * manager never resizes it.
   */
  output: HTMLCanvasElement
  sink: DiagnosticSink
  /** Additive to the brief: start on Canvas 2D without trying WebGL2 (tests, user setting). */
  force?: 'canvas2d'
}

export type BackendState = 'webgl2' | 'lost' | 'canvas2d' | 'disposed'

export interface BackendManager {
  readonly current: Compositor
  /** Additive to the brief: `'lost'` while the 2-D canvas paints live state. */
  readonly state: BackendState
  onLoss(): void
  dispose(): void
}

/** What the live painter needs from the session each frame; `scratch` and `scale` are the painter's. */
export interface LiveScene {
  slide: PptxSlide
  state: RenderState
  assets: Omit<DirectAssets, 'scratch' | 'scale'>
}

/**
 * The production `paintLive`: `paintDirect` with `DirectAssets.scratch` set
 * to `terminal.scratch` (the permanent reserve, so a loss never needs a fresh
 * admission) and `scale` 1, because the manager has already applied the
 * viewport to the context. `source` returning undefined paints nothing.
 */
export function createLivePainter(
  terminal: TerminalReserve,
  source: () => LiveScene | undefined,
  paint: typeof paintDirect = paintDirect,
): (ctx: CanvasRenderingContext2D) => void {
  return ctx => {
    const scene = source()
    if (!scene) return
    const assets: DirectAssets = { ...scene.assets, scale: 1, scratch: terminal.scratch }
    paint(ctx, scene.slide, scene.state, assets)
  }
}

const IDENTITY: Affine = [1, 0, 0, 1, 0, 0]

export function createBackendManager(host: HTMLElement, deps: BackendDeps): BackendManager {
  // S3-7: every 3-D row has its shader before any WebGL2 compositor (first or restored) can dispatch one.
  registerTransitions3D()
  const doc = host.ownerDocument
  const factory: CanvasFactory = (w, h) => {
    const c = doc.createElement('canvas')
    c.width = w; c.height = h
    return c
  }

  let phase: BackendState = 'webgl2'
  let gl: WebGL2Compositor | undefined
  let glCanvas: HTMLCanvasElement | undefined
  let live: { canvas: HTMLCanvasElement; ctx: CanvasRenderingContext2D | null; k: number } | undefined
  let c2d: Compositor | undefined
  let active: Compositor | 'live'
  let view: CompositorViewport | undefined
  let slideIndex: number | undefined
  let losses = 0
  let lostAt = 0
  let restored = false
  /** Bumped on every loss, switch and dispose; stale callbacks compare against it. */
  let gen = 0

  const diagnose = (code: string, effect: string, message: string) =>
    deps.sink.add(slideIndex === undefined ? { code, effect, message } : { code, slideIndex, effect, message })

  function style(c: HTMLCanvasElement, role: 'webgl2' | 'canvas2d') {
    c.dataset.slideshowBackend = role
    const s = c.style
    s.position = 'absolute'; s.left = '0'; s.top = '0'; s.width = '100%'; s.height = '100%'; s.display = 'block'
  }

  function deviceSize(): { w: number; h: number } {
    if (view && view.deviceWidth >= 1 && view.deviceHeight >= 1) return { w: Math.floor(view.deviceWidth), h: Math.floor(view.deviceHeight) }
    if (glCanvas) return { w: glCanvas.width, h: glCanvas.height }
    return { w: 300, h: 150 }
  }

  /** Scale the device size down to fit the output canvas (its owner sizes it to the backing store). */
  function fitLive() {
    if (!live) return
    const { w, h } = deviceSize()
    live.k = Math.min(1, live.canvas.width / Math.max(1, w), live.canvas.height / Math.max(1, h))
  }

  /** Composite on the output canvas: nothing is created or admitted (PR #18 review P2). */
  function mountLive() {
    if (live) return
    const canvas = deps.output
    canvas.dataset.slideshowBackend = 'canvas2d'
    live = { canvas, ctx: canvas.getContext('2d'), k: 1 }
    fitLive()
    // A lost WebGL canvas stacks over the output: hide it while live state paints beneath.
    if (glCanvas) glCanvas.style.visibility = 'hidden'
  }

  function unmountLive() {
    if (!live) return
    delete live.canvas.dataset.slideshowBackend
    if (glCanvas) glCanvas.style.visibility = ''
    live = undefined
  }

  /** The viewport as the live canvas sees it. */
  function fitted(): CompositorViewport {
    const k = live?.k ?? 1
    return {
      matrix: mulAffine([k, 0, 0, k, 0, 0], view?.matrix ?? IDENTITY),
      deviceWidth: live?.canvas.width ?? 1,
      deviceHeight: live?.canvas.height ?? 1,
    }
  }

  function paintLiveFrame() {
    const ctx = live?.ctx
    if (!ctx || !live) return
    ctx.setTransform(1, 0, 0, 1, 0, 0)
    ctx.globalAlpha = 1
    ctx.globalCompositeOperation = 'source-over'
    ctx.clearRect(0, 0, live.canvas.width, live.canvas.height)
    const m = fitted().matrix
    ctx.setTransform(m[0], m[1], m[2], m[3], m[4], m[5])
    try { deps.paintLive(ctx) }
    finally { ctx.setTransform(1, 0, 0, 1, 0, 0) }
  }

  const onLost = (e: Event) => {
    e.preventDefault()
    manager.onLoss()
  }
  const onRestored = () => handleRestore()

  function dropGL() {
    if (glCanvas) {
      glCanvas.removeEventListener('webglcontextlost', onLost)
      glCanvas.removeEventListener('webglcontextrestored', onRestored)
      glCanvas.remove()
    }
    gl?.dispose()
    gl = undefined
    glCanvas = undefined
  }

  function startCanvas2D() {
    mountLive()
    c2d = createCanvas2DCompositor(live!.canvas, factory, deps.workspace, deps.sink)
    c2d.setSlideIndex(slideIndex)
    if (view) c2d.setViewport(fitted())
    return c2d
  }

  function switchToCanvas2D(reason: string, message: string) {
    if (phase === 'disposed') return
    const token = ++gen
    diagnose('backend-fallback', reason, `${message}; the show continues on the Canvas 2D compositor`)
    phase = 'canvas2d'
    dropGL()
    active = 'live'
    const c = startCanvas2D()
    deps.rebuildLayers(c).then(
      () => { if (token === gen && phase === 'canvas2d') active = c },
      () => {
        if (token === gen && phase === 'canvas2d') {
          diagnose('backend-fallback', 'canvas2d-rebuild', 'The Canvas 2D layers could not be built; live state is painted directly')
        }
      },
    )
  }

  function handleRestore() {
    if (phase !== 'lost' || restored || !gl) return
    if (deps.now() - lostAt >= RESTORE_DEADLINE_MS) {
      switchToCanvas2D('restore-deadline', `The WebGL2 context was not restored within ${RESTORE_DEADLINE_MS} ms`)
      return
    }
    if (!gl.contextRestored()) {
      switchToCanvas2D('restore-workspace', 'The restored WebGL2 resources did not fit the slideshow workspace')
      return
    }
    restored = true
    const c = gl
    c.setSlideIndex(slideIndex)
    if (view) c.setViewport(view)
    const token = gen
    deps.rebuildLayers(c).then(
      () => {
        if (token !== gen || phase !== 'lost') return
        active = c
        phase = 'webgl2'
        unmountLive()
      },
      () => {
        if (token !== gen || phase !== 'lost') return
        switchToCanvas2D('restore-rebuild', 'The layers could not be rebuilt on the restored WebGL2 context')
      },
    )
  }

  /** False when the WebGL2 output could not take the new size (PR #18 review 4: the caller keeps the old one). */
  function forward(): boolean {
    if (!view) return true
    const ok = gl?.setViewport(view) !== false
    if (live) {
      fitLive()
      c2d?.setViewport(fitted())
    }
    return ok
  }

  const facade: Compositor = {
    get backend() { return active === 'live' ? 'canvas2d' : active.backend },
    get maxTextureSize() { return active === 'live' ? gl?.maxTextureSize ?? CANVAS2D_MAX_TEXTURE_SIZE : active.maxTextureSize },
    upload(layer, lease) { return active === 'live' || phase === 'disposed' ? { refused: 'workspace' } : active.upload(layer, lease) },
    release(handle) { if (active !== 'live') active.release(handle) },
    frame(draws, transition) {
      if (phase === 'disposed') return
      if (active === 'live') paintLiveFrame()
      else active.frame(draws, transition)
    },
    resize(widthPx, heightPx) {
      view = { matrix: view?.matrix ?? IDENTITY, deviceWidth: widthPx, deviceHeight: heightPx }
      forward()
    },
    setViewport(next) {
      view = next
      return forward()
    },
    setSlideIndex(index) {
      slideIndex = index
      gl?.setSlideIndex(index)
      c2d?.setSlideIndex(index)
    },
    // S3-1: only the painting WebGL2 compositor prepares 3-D transitions (never while lost or on Canvas 2D).
    prepareTransition3D(t) { return active !== 'live' && phase !== 'disposed' && active.backend === 'webgl2' && active.prepareTransition3D?.(t) === true },
    dispose() { manager.dispose() },
  }

  const manager: BackendManager = {
    get current() { return facade },
    get state() { return phase },

    onLoss() {
      if (phase === 'disposed' || phase === 'canvas2d') return
      losses++
      const token = ++gen
      restored = false
      gl?.contextLost()
      mountLive()
      active = 'live'
      phase = 'lost'
      diagnose('context-lost', 'webgl2', 'The WebGL2 context was lost; live state is painted on a 2-D canvas')
      paintLiveFrame()
      if (losses > MAX_CONTEXT_LOSSES) {
        switchToCanvas2D('repeated-loss', `The WebGL2 context was lost more than ${MAX_CONTEXT_LOSSES} times`)
        return
      }
      lostAt = deps.now()
      const check = () => {
        if (token !== gen || phase !== 'lost' || restored) return
        const left = RESTORE_DEADLINE_MS - (deps.now() - lostAt)
        if (left > 0) { deps.schedule(check, left); return }
        switchToCanvas2D('restore-deadline', `The WebGL2 context was not restored within ${RESTORE_DEADLINE_MS} ms`)
      }
      deps.schedule(check, RESTORE_DEADLINE_MS)
    },

    dispose() {
      if (phase === 'disposed') return
      phase = 'disposed'
      gen++
      dropGL()
      c2d?.dispose()
      c2d = undefined
      unmountLive()
    },
  }

  if (deps.force !== 'canvas2d') {
    glCanvas = doc.createElement('canvas')
    style(glCanvas, 'webgl2')
    host.appendChild(glCanvas)
    gl = createWebGL2Compositor(glCanvas, deps.workspace, { factory, sink: deps.sink })
    if (gl) {
      glCanvas.addEventListener('webglcontextlost', onLost)
      glCanvas.addEventListener('webglcontextrestored', onRestored)
    } else {
      glCanvas.remove()
      glCanvas = undefined
    }
  }
  if (gl) {
    active = gl
  } else {
    diagnose('backend-fallback', deps.force === 'canvas2d' ? 'forced' : 'webgl2-unavailable',
      deps.force === 'canvas2d' ? 'Canvas 2D was requested' : 'WebGL2 is unavailable; the show runs on the Canvas 2D compositor')
    phase = 'canvas2d'
    active = startCanvas2D()
  }
  return manager
}
