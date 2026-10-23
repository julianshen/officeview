/**
 * The 3-D transition framework (slideshow spec 6.3, S3-1): which path draws a
 * transition, and the params that carry a shader reference to the
 * compositor.
 *
 * - **Dispatch.** `dispatchTransition(spec, compositor)` is `'shader'` when
 *   the spec's row is 3-D, a `Transition3D` is registered for it, the
 *   compositor is WebGL2 (not lost) and the transition prepared (its grids
 *   leased, its programs compiled). Otherwise it is `'lookalike'`: the spec
 *   is drawn by `transitionDraw`, which draws a 3-D row's catalogued 2-D
 *   look-alike, and an approximated row is diagnosed once
 *   (`transition-approximate`, deduplicated per slide and effect by the
 *   sink). The caller applies reduced motion first (`reducedTransition`
 *   turns every real transition into a fade), so reduced motion never
 *   reaches a shader.
 * - **Params.** `transition3DDraw(spec, progress, rng, aspect)` returns the
 *   row's look-alike params (from `transitionDraw`, with the same rng) plus
 *   `three`, the shader reference. The WebGL2 compositor draws `three` with
 *   the registered shader; Canvas 2D ignores it and draws the look-alike, so
 *   the same params stay correct across a backend switch. The WebGL2
 *   compositor also falls back to the look-alike params, diagnosed
 *   (`transition-shader-fallback`), when the shader throws in a frame.
 * - **Ends.** At progress 0 and 1 the compositor draws the flat outgoing and
 *   incoming frames (the look-alike's ends) without calling the shader.
 * - **Seeds.** The compositor re-seeds the shader's rng from the effect rng's
 *   seed (the same memo its masks use) on every frame.
 */
import { resolveTransition2D, transitionDraw } from '../transitions2d'
import type { Compositor, DiagnosticSink, Rng, TransitionSpec } from '../../types'
import { transition3DFor } from './registry'

/** The shader reference carried in `TransitionDraw.params.three`. */
export interface Shader3DRef {
  ns: TransitionSpec['ns']
  element: string
  preset?: string
  attrs: Readonly<Record<string, string>>
}

/** Which path draws `spec` on `compositor`; diagnoses an approximated look-alike into `sink`.
 * A `p159:morph` row on WebGL2 dispatches to `'morph'`: the composited
 * object-interpolation path (`morphFrame`), itself always approximate and
 * diagnosed here; Canvas 2D keeps the catalogued fade look-alike (diagnosed). */
export function dispatchTransition(spec: TransitionSpec, compositor: Compositor, sink?: DiagnosticSink, slideIndex?: number): 'shader' | 'morph' | 'lookalike' {
  if (spec.ns === 'p159' && spec.element === 'morph') {
    if (compositor.backend === 'webgl2') {
      const name = `${spec.ns}:${spec.element}${spec.preset ? ' ' + spec.preset : ''}`
      sink?.add({
        code: 'transition-approximate',
        ...(slideIndex !== undefined ? { slideIndex } : {}),
        effect: 'morph',
        message: `Transition ${name} plays as an approximation: matched objects interpolate, content crossfades`,
      })
      return 'morph'
    }
  }
  const t = transition3DFor(spec)
  let reason = 'WebGL2 is not available'
  if (t) {
    if (compositor.backend === 'webgl2' && compositor.prepareTransition3D?.(t) === true) return 'shader'
    if (compositor.backend === 'webgl2') reason = 'its 3-D shader could not be prepared'
  } else if (compositor.backend === 'webgl2') reason = 'it has no 3-D shader yet'
  const drawn = resolveTransition2D(spec)
  if (drawn.approximated && sink) {
    const name = `${spec.ns}:${spec.element}${spec.preset ? ' ' + spec.preset : ''}`
    sink.add({
      code: 'transition-approximate',
      ...(slideIndex !== undefined ? { slideIndex } : {}),
      effect: spec.preset ? `${spec.element}:${spec.preset}` : spec.element,
      message: `Transition ${name} is drawn as its 2-D look-alike ${drawn.spec.element} (${reason})`,
    })
  }
  return 'lookalike'
}

/** The look-alike's params plus the shader reference, for a transition dispatched to `'shader'`. */
export function transition3DDraw(spec: TransitionSpec, progress: number, rng: Rng, aspect = 16 / 9): ReturnType<typeof transitionDraw> & { three: Shader3DRef } {
  const params = transitionDraw(spec, progress, rng, aspect)
  const three: Shader3DRef = { ns: spec.ns, element: spec.element, ...(spec.preset !== undefined ? { preset: spec.preset } : {}), attrs: spec.attrs }
  return { ...params, three }
}

/** The shader reference in `params`, if well formed. */
export function shader3DOf(params: Readonly<Record<string, unknown>>): Shader3DRef | undefined {
  const v = params.three as Partial<Shader3DRef> | undefined
  if (!v || typeof v !== 'object' || typeof v.element !== 'string' || typeof v.ns !== 'string') return undefined
  if (v.preset !== undefined && typeof v.preset !== 'string') return undefined
  const attrs = v.attrs && typeof v.attrs === 'object' ? v.attrs : {}
  return { ns: v.ns, element: v.element, ...(v.preset !== undefined ? { preset: v.preset } : {}), attrs }
}
