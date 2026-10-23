/**
 * The geometric 3-D transitions (slideshow spec 6.3, plan S3-2): `prism`
 * (the Cube / Box / Rotate / Orbit variants, by `dir`, `isInverted` and
 * `isContent`), `flip`, `gallery`, `doors`, `window`, `switch`, `conveyor`
 * and `ferris`. `pan` is a 2-D row (S1-15) and has no shader.
 *
 * Each row follows its entry in docs/validation/slideshow-3d-algorithms.md
 * ("S3-2: geometric rows"); the reference models are the functions of the
 * same names in scripts/slideshow-3d-oracles.ts, and the tests hold these
 * models to that script and to the sheet's 50 % oracles. Every row is a
 * closest look-alike, not an Office match.
 *
 * - **Pieces.** A row is a few slide-sized (or half-slide, via `uvRect`)
 *   rigid quads: `geometricPieces` returns each one's frame, model matrix
 *   and faces, in draw order (painter's order, farthest centre first; the
 *   drawing buffer has no depth). All of them draw the one cached 1 × 1 grid
 *   with the framework's standard program, so a frame is 2 to 5 draw calls.
 * - **Faces.** prism and flip cull back faces (a turned-away face is the
 *   cube's hidden side or the card's other frame). The doors and window
 *   halves are two-sided: a half that swings past edge-on shows its back,
 *   shaded. The rest never turn far enough to face away.
 * - **Light.** A face turned by angle φ from the viewer is shaded by
 *   `0.7 + 0.3 |cos φ|`; it is 1 at rest, so the ends stay exact.
 * - **Seeds.** No geometric row is seeded: a frame is a pure function of
 *   progress, attributes and aspect, and the rng is never drawn from.
 * - **Leases.** `prepare` leases the 1 × 1 grid (`webgl2:mesh`, before any
 *   GL object exists) and compiles the standard program, so a refusal or a
 *   failed compile sends the effect to its look-alike before it starts.
 *
 * Registration: `registerGeometric3D()`; S3-7 wires it into the library.
 */
import type { Rng } from '../../types'
import {
  DEFAULT_FOV_DEG, FRAGMENT_3D, STANDARD_PROGRAM, VERTEX_3D, mul4, rotationX, rotationY, slideCamera, translation, type Mat4, type MeshPool,
} from './mesh'
import { registerTransition3D, type Transition3D } from './registry'

export const GEOMETRIC_ELEMENTS = ['prism', 'flip', 'gallery', 'doors', 'window', 'switch', 'conveyor', 'ferris'] as const
export type GeometricElement = typeof GEOMETRIC_ELEMENTS[number]

export type UvRect = readonly [number, number, number, number]

export interface GeometricPiece {
  /** `from` / `to` are whole frames; `first` / `second` are the doors' and window's halves (left / right, or top / bottom). */
  role: 'from' | 'to' | 'first' | 'second'
  /** The frame texture the piece shows. */
  face: 'from' | 'to'
  model: Mat4
  /** The slide-uv sub-rect the piece covers; omitted for the whole slide. */
  uvRect?: UvRect
  /** `front`: back faces culled. `both`: the back is drawn too, shaded. */
  sides: 'front' | 'both'
}

/** How much a piece's back face is darkened relative to its front. */
export const BACK_SHADE = 0.55

const T = translation, Rx = rotationX, Ry = rotationY
const chain = (...ms: Mat4[]): Mat4 => ms.reduce((acc, x) => mul4(acc, x))
const ease = (t: number): number => { const x = Math.min(1, Math.max(0, t)); return x * x * (3 - 2 * x) }
const bulge = (p: number): number => Math.sin(Math.PI * p)
const bool = (v: string | undefined): boolean => v === '1' || v === 'true'
const leftRight = (attrs: Readonly<Record<string, string>>): number => (attrs.dir === 'r' ? 1 : -1)

/** The pieces of `element` at `progress` (clamped to [0, 1]), in draw order. Pure: the tests project these. */
export function geometricPieces(element: string, progress: number, attrs: Readonly<Record<string, string>>, aspect: number): GeometricPiece[] {
  const p = Math.min(1, Math.max(0, Number.isFinite(progress) ? progress : 0))
  const a = aspect
  const e = ease(p), s = bulge(p)
  const whole = (role: 'from' | 'to', model: Mat4, sides: GeometricPiece['sides'] = 'front'): GeometricPiece => ({ role, face: role, model, sides })
  switch (element) {
    case 'prism': {
      const dir = attrs.dir === 'r' || attrs.dir === 'u' || attrs.dir === 'd' ? attrs.dir : 'l'
      const horizontal = dir === 'l' || dir === 'r'
      const D = horizontal ? a : 1
      const sigma = dir === 'l' || dir === 'u' ? -1 : 1
      const R = horizontal ? Ry : Rx
      const inverted = bool(attrs.isInverted)
      const K = inverted ? D / 2 : -D / 2
      const turn = inverted ? -sigma : sigma
      const th = (Math.PI / 2) * e
      const lift = T(0, 0, bool(attrs.isContent) ? -0.3 * s : 0)
      return byDepth([
        whole('from', chain(lift, T(0, 0, K), R(turn * th), T(0, 0, -K))),
        whole('to', chain(lift, T(0, 0, K), R(turn * th - turn * Math.PI / 2), T(0, 0, -K))),
      ])
    }
    case 'flip': {
      const sg = leftRight(attrs), th = Math.PI * e, z = -0.5 * s
      return [whole('from', chain(T(0, 0, z), Ry(sg * th))), whole('to', chain(T(0, 0, z), Ry(sg * th - sg * Math.PI)))]
    }
    case 'gallery': {
      const sg = leftRight(attrs)
      return [
        whole('from', chain(T(sg * 1.1 * a * e, 0, -0.6 * s), Ry(sg * 0.35 * s))),
        whole('to', chain(T(sg * 1.1 * a * (e - 1), 0, -0.6 * s), Ry(sg * 0.35 * s))),
      ]
    }
    case 'doors':
    case 'window': {
      const windowed = element === 'window'
      const th = (Math.PI / 2) * e
      const out = windowed ? 0.25 * a * e : 0
      const to = whole('to', T(0, 0, (windowed ? -1.2 : -0.6) * (1 - e)))
      const half = (role: 'first' | 'second', model: Mat4, uvRect: UvRect): GeometricPiece => ({ role, face: 'from', model, uvRect, sides: 'both' })
      if (attrs.dir === 'horz') {
        return [
          to,
          half('first', chain(T(0, out / a, 0), T(0, 0.5, 0), Rx(-th), T(0, -0.5, 0)), [0, 0, 1, 0.5]),
          half('second', chain(T(0, -out / a, 0), T(0, -0.5, 0), Rx(th), T(0, 0.5, 0)), [0, 0.5, 1, 0.5]),
        ]
      }
      return [
        to,
        half('first', chain(T(-out, 0, 0), T(-a / 2, 0, 0), Ry(-th), T(a / 2, 0, 0)), [0, 0, 0.5, 1]),
        half('second', chain(T(out, 0, 0), T(a / 2, 0, 0), Ry(th), T(-a / 2, 0, 0)), [0.5, 0, 0.5, 1]),
      ]
    }
    case 'switch': {
      const sg = leftRight(attrs)
      return byDepth([
        whole('from', chain(T(sg * 0.55 * a * s, 0, -1.2 * e), Ry(sg * 0.6 * s))),
        whole('to', chain(T(-sg * 0.55 * a * s, 0, -1.2 * (1 - e)), Ry(-sg * 0.6 * s))),
      ])
    }
    case 'conveyor': {
      const sg = leftRight(attrs), z = -0.8 * s, xf = sg * 1.1 * a * e
      return [whole('from', chain(T(xf, 0, z), Ry(sg * 0.25 * s))), whole('to', chain(T(xf - sg * 1.1 * a, 0, z), Ry(sg * 0.25 * s)))]
    }
    case 'ferris': {
      const sg = leftRight(attrs), R = 0.9 * a, phi = (Math.PI / 3) * e, z = -0.3 * s
      const at = (psi: number) => T(R * Math.sin(psi), -R + R * Math.cos(psi), z)
      return [whole('from', at(sg * phi)), whole('to', at(sg * (phi - Math.PI / 3)))]
    }
    default:
      return []
  }
}

/** Farthest first by the world z of each piece's centre (the camera looks down -z); stable, so a tie keeps the given order. */
function byDepth(pieces: GeometricPiece[]): GeometricPiece[] {
  const z = (pc: GeometricPiece) => {
    const r = pc.uvRect ?? [0, 0, 1, 1]
    // The centre's rest position is on z = 0; only its x, y (in slide units) and the model's z row matter.
    const cx = r[0] + r[2] / 2 - 0.5, cy = 0.5 - (r[1] + r[3] / 2), m = pc.model
    return m[2] * cx + m[6] * cy + m[14]
  }
  return pieces.map((pc, i) => ({ pc, i, z: z(pc) })).sort((x, y) => x.z - y.z || x.i - y.i).map(x => x.pc)
}

/** `0.7 + 0.3 |cos φ|`, φ the angle between the piece's turned normal and the viewer axis. */
const shadeOf = (m: Mat4): number => 0.7 + 0.3 * Math.min(1, Math.abs(m[10]))

function geometricTransition(element: GeometricElement): Transition3D {
  return {
    element,
    prepare(mesh: MeshPool) {
      mesh.grid(1, 1)
      mesh.program(STANDARD_PROGRAM, VERTEX_3D, FRAGMENT_3D)
    },
    draw(gl: WebGL2RenderingContext, from: WebGLTexture, to: WebGLTexture, progress: number, params: Readonly<Record<string, string>>, _rng: Rng, mesh: MeshPool) {
      const camera = slideCamera(DEFAULT_FOV_DEG, mesh.aspect)
      const quad = mesh.grid(1, 1)
      gl.enable(gl.BLEND)
      gl.blendFunc(gl.ONE, gl.ONE_MINUS_SRC_ALPHA)
      gl.frontFace(gl.CCW)
      for (const pc of geometricPieces(element, progress, params, mesh.aspect)) {
        const mvp = mul4(camera, pc.model)
        const texture = pc.face === 'from' ? from : to
        const shade = shadeOf(pc.model)
        const uvRect = pc.uvRect
        gl.enable(gl.CULL_FACE)
        gl.cullFace(gl.BACK)
        mesh.draw(quad, texture, mvp, { shade, ...(uvRect ? { uvRect } : {}) })
        if (pc.sides === 'both') {
          gl.cullFace(gl.FRONT)
          mesh.draw(quad, texture, mvp, { shade: shade * BACK_SHADE, ...(uvRect ? { uvRect } : {}) })
        }
      }
    },
  }
}

/** One `Transition3D` per geometric row. */
export const GEOMETRIC_TRANSITIONS: readonly Transition3D[] = Object.freeze(GEOMETRIC_ELEMENTS.map(geometricTransition))

/** Registers every geometric row (idempotent). */
export function registerGeometric3D(): void {
  for (const t of GEOMETRIC_TRANSITIONS) registerTransition3D(t)
}
