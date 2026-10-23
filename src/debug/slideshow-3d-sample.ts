/**
 * Slideshow plan S3-1: a SAMPLE 3-D transition for the framework's tests
 * (tests/slideshow-3d-framework.test.ts and, through
 * src/debug/slideshow-gpu.ts, tests/browser/slideshow-3d.spec.ts). It is
 * never registered by the library: the group tasks (S3-2 to S3-4) own every
 * real row.
 *
 * The sample is a card flip about the slide's vertical centre line: over the
 * first half the outgoing frame turns from 0° to 90° (edge-on), over the
 * second half the incoming frame turns from -90° to 0°. It uses only the
 * framework's helpers (a 1 × 1 grid, the slide camera, the standard program),
 * so its pure model (`sampleFace`) is the oracle the tests project with.
 */
import type { Rng } from '../slideshow/types'
import { mul4, rotationY, slideCamera, DEFAULT_FOV_DEG, type Mat4, type MeshPool } from '../slideshow/render/three/mesh'
import type { Transition3D } from '../slideshow/render/three/registry'

/** Which frame the sample shows at `progress` and the matrix it draws it with. */
export function sampleFace(progress: number, aspect: number): { face: 'from' | 'to'; mvp: Mat4 } {
  const p = Math.min(1, Math.max(0, progress))
  const face = p < 0.5 ? 'from' : 'to'
  const angle = face === 'from' ? Math.PI * p : Math.PI * (p - 1)
  return { face, mvp: mul4(slideCamera(DEFAULT_FOV_DEG, aspect), rotationY(angle)) }
}

export interface SampleCall { progress: number; attrs: Readonly<Record<string, string>>; first: number; meshAspect: number }

/** A fresh sample bound to a catalogued 3-D row (default `p14:vortex`), recording every draw. */
export function createSampleTransition3D(element = 'vortex', preset?: string): { transition: Transition3D; calls: SampleCall[] } {
  const calls: SampleCall[] = []
  const transition: Transition3D = {
    element,
    ...(preset !== undefined ? { preset } : {}),
    prepare(mesh) { mesh.grid(1, 1) },
    draw(_gl: WebGL2RenderingContext, from: WebGLTexture, to: WebGLTexture, progress: number, params: Readonly<Record<string, string>>, rng: Rng, mesh: MeshPool) {
      calls.push({ progress, attrs: params, first: rng(), meshAspect: mesh.aspect })
      const { face, mvp } = sampleFace(progress, mesh.aspect)
      mesh.draw(mesh.grid(1, 1), face === 'from' ? from : to, mvp)
    },
  }
  return { transition, calls }
}
