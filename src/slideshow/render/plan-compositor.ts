/**
 * The painter-free compositor for streamed shows (piece 4 plan D17).
 *
 * `render/backend.ts` cannot serve the streamed show as it is: its live
 * painter reaches `paintDirect` (and through it the package). This module is
 * the same facade with no painter at all, for a stage that composites only
 * bitmaps it already holds (`PlanStage`):
 *
 * - it mounts a WebGL2 canvas in `host` and composites on it; without WebGL2
 *   (or with `force: 'canvas2d'`) it composites on `output` with the Canvas
 *   2D compositor, diagnosed (`backend-fallback`);
 * - on `webglcontextlost` (or `onLoss()`) it switches to Canvas 2D at once,
 *   for the rest of the show: there is no live painter to bridge a restore,
 *   and the stage holds every bitmap it uploaded, so `rebuild(c)` re-uploads
 *   them to the new compositor synchronously before the next frame;
 * - `current` is one stable `Compositor` that delegates to the active backend
 *   and replays the last viewport and slide index onto a new one.
 *
 * Handles do not survive a switch: after `rebuild`, draw only the handles it
 * uploaded to the compositor it was given. Package-free (D25): nothing here
 * reaches the reader, the preparation or the direct painter.
 */
import type { Affine, CanvasFactory, Compositor, CompositorViewport, DiagnosticSink, SlideShowWorkspace } from '../types'
import { CANVAS2D_MAX_TEXTURE_SIZE, createCanvas2DCompositor } from './canvas2d'
import { registerTransitions3D } from './three'
import { createWebGL2Compositor, type WebGL2Compositor } from './webgl2'

export interface PlanCompositorDeps {
  /** The 2-D output canvas (never given a WebGL context): the Canvas 2D compositor draws on it. Its owner sizes it. */
  output: HTMLCanvasElement
  workspace: SlideShowWorkspace
  sink: DiagnosticSink
  /** A backend switch: every handle is void; re-upload what is drawn to `c` (synchronously). */
  rebuild(c: Compositor): void
  /** Start on Canvas 2D without trying WebGL2 (tests, user setting). */
  force?: 'canvas2d'
}

export type PlanCompositorState = 'webgl2' | 'canvas2d' | 'disposed'

export interface PlanCompositor {
  readonly current: Compositor
  readonly state: PlanCompositorState
  /** The WebGL2 context is gone: switch to Canvas 2D now (idempotent). */
  onLoss(): void
  dispose(): void
}

const IDENTITY: Affine = [1, 0, 0, 1, 0, 0]

export function createPlanCompositor(host: HTMLElement, deps: PlanCompositorDeps): PlanCompositor {
  // Every 3-D row has its shader before a WebGL2 compositor can dispatch one (S3-7).
  registerTransitions3D()
  const doc = host.ownerDocument
  const factory: CanvasFactory = (w, h) => {
    const c = doc.createElement('canvas')
    c.width = Math.max(1, Math.ceil(w)); c.height = Math.max(1, Math.ceil(h))
    return c
  }
  let phase: PlanCompositorState = 'webgl2'
  let gl: WebGL2Compositor | undefined
  let glCanvas: HTMLCanvasElement | undefined
  let active: Compositor
  let view: CompositorViewport | undefined
  let slideIndex: number | undefined

  const diagnose = (effect: string, message: string) =>
    deps.sink.add(slideIndex === undefined ? { code: 'backend-fallback', effect, message } : { code: 'backend-fallback', slideIndex, effect, message })

  const onLost = (e: Event) => { e.preventDefault(); manager.onLoss() }

  function dropGL() {
    if (glCanvas) {
      glCanvas.removeEventListener('webglcontextlost', onLost)
      glCanvas.remove()
    }
    gl?.dispose()
    gl = undefined
    glCanvas = undefined
  }

  function startCanvas2D(): Compositor {
    deps.output.dataset.slideshowBackend = 'canvas2d'
    const c = createCanvas2DCompositor(deps.output, factory, deps.workspace, deps.sink)
    c.setSlideIndex(slideIndex)
    if (view) c.setViewport(view)
    return c
  }

  const facade: Compositor = {
    get backend() { return active.backend },
    get maxTextureSize() { return phase === 'disposed' ? CANVAS2D_MAX_TEXTURE_SIZE : active.maxTextureSize },
    upload(layer, lease) { return phase === 'disposed' ? { refused: 'workspace' } : active.upload(layer, lease) },
    release(handle) { if (phase !== 'disposed') active.release(handle) },
    frame(draws, transition) { if (phase !== 'disposed') active.frame(draws, transition) },
    resize(widthPx, heightPx) {
      view = { matrix: view?.matrix ?? IDENTITY, deviceWidth: widthPx, deviceHeight: heightPx }
      return facade.setViewport(view)
    },
    setViewport(next) {
      view = next
      return phase === 'disposed' ? true : active.setViewport(next) !== false
    },
    setSlideIndex(index) {
      slideIndex = index
      if (phase !== 'disposed') active.setSlideIndex(index)
    },
    prepareTransition3D(t) { return phase === 'webgl2' && active.prepareTransition3D?.(t) === true },
    dispose() { manager.dispose() },
  }

  const manager: PlanCompositor = {
    get current() { return facade },
    get state() { return phase },
    onLoss() {
      if (phase !== 'webgl2') return
      gl?.contextLost()
      diagnose('context-lost', 'The WebGL2 context was lost; the show continues on the Canvas 2D compositor')
      phase = 'canvas2d'
      dropGL()
      active = startCanvas2D()
      deps.rebuild(facade)
    },
    dispose() {
      if (phase === 'disposed') return
      const was = active
      phase = 'disposed'
      if (gl) dropGL()
      else was.dispose()
      delete deps.output.dataset.slideshowBackend
    },
  }

  if (deps.force !== 'canvas2d') {
    glCanvas = doc.createElement('canvas')
    glCanvas.dataset.slideshowBackend = 'webgl2'
    const s = glCanvas.style
    s.position = 'absolute'; s.left = '0'; s.top = '0'; s.width = '100%'; s.height = '100%'; s.display = 'block'
    host.appendChild(glCanvas)
    gl = createWebGL2Compositor(glCanvas, deps.workspace, { factory, sink: deps.sink })
    if (gl) glCanvas.addEventListener('webglcontextlost', onLost)
    else { glCanvas.remove(); glCanvas = undefined }
  }
  if (gl) active = gl
  else {
    phase = 'canvas2d'
    diagnose(deps.force === 'canvas2d' ? 'forced' : 'webgl2-unavailable',
      deps.force === 'canvas2d' ? 'Canvas 2D was requested' : 'WebGL2 is unavailable; the show runs on the Canvas 2D compositor')
    active = startCanvas2D()
  }
  return manager
}
