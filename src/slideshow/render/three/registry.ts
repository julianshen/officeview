/**
 * The 3-D transition registry (slideshow spec 6.3, S3-1).
 *
 * One `Transition3D` per catalogued 3-D row (`THREE_D_IDS`: every stage-3
 * transition row except Morph, each of which names a 2-D look-alike). Group
 * tasks never edit this file or an index: each exports `registerGeometric3D()`,
 * `registerParticle3D()` or `registerDeform3D()`, which calls
 * `registerTransition3D` for its rows, and S3-7 wires those calls in
 * `three/index.ts`. Until then the registry is empty and every 3-D row draws
 * its look-alike (diagnosed by `dispatchTransition`).
 *
 * A row is keyed by `element` plus, for `p15:prstTrans`, `preset`: the 3-D
 * elements are unique across the p14 and p15 namespaces.
 */
import { CATALOGUE, rowForTransition } from '../../catalogue'
import type { CatalogueRow, Rng, TransitionSpec } from '../../types'
import type { MeshPool } from './mesh'

export interface Transition3D {
  element: string
  preset?: string
  /**
   * Draws one frame at `progress` in (0, 1) (the compositor draws the flat
   * outgoing / incoming frame at exactly 0 / 1). The GL viewport is the
   * slide's device rect, scissored, cleared to opaque black; `from` / `to`
   * are the premultiplied frame textures; `params` are the transition's
   * attributes; `rng` is re-seeded from the effect's seed on every frame, so
   * the same draws give the same values at every progress. Any GL state may
   * be changed: the compositor resets it afterwards. Throwing (for example a
   * `Transition3DError` from the pool) makes the compositor draw the row's
   * look-alike for that frame, diagnosed once.
   */
  draw(gl: WebGL2RenderingContext, from: WebGLTexture, to: WebGLTexture, progress: number, params: Readonly<Record<string, string>>, rng: Rng, mesh: MeshPool): void
  /**
   * Additive (S3-1): called once when the transition is dispatched, to lease
   * its grids and compile its programs before it starts. Throwing sends the
   * transition to its look-alike for the whole effect.
   */
  prepare?(mesh: MeshPool): void
}

const THREE_D_ROWS: readonly CatalogueRow[] = CATALOGUE.filter(r => r.kind === 'transition' && r.stage === 3 && r.ns !== 'p159')

/** The catalogue rows a `Transition3D` may be registered for. */
export const THREE_D_IDS: readonly string[] = Object.freeze(THREE_D_ROWS.map(r => r.id))

const keyOf = (element: string, preset?: string) => `${element}|${preset ?? ''}`
const ALLOWED: ReadonlySet<string> = new Set(THREE_D_ROWS.map(r => keyOf(r.element!, presetOf(r))))
function presetOf(r: CatalogueRow): string | undefined {
  // Transition rows carry their prstTrans preset in the id (`transition:p15:prstTrans:<preset>`).
  const parts = r.id.split(':')
  return parts.length > 3 ? parts.slice(3).join(':') : undefined
}

const registry = new Map<string, Transition3D>()

/** Registers `t` for its catalogued 3-D row. Throws for a row that is not 3-D or is already taken by another object. */
export function registerTransition3D(t: Transition3D): void {
  const key = keyOf(t.element, t.preset)
  if (!ALLOWED.has(key)) throw new Error(`${t.element}${t.preset ? ' ' + t.preset : ''} is not a catalogued 3-D transition row`)
  const had = registry.get(key)
  if (had === t) return
  if (had) throw new Error(`A 3-D transition is already registered for ${t.element}${t.preset ? ' ' + t.preset : ''}`)
  registry.set(key, t)
}

/** The registered `Transition3D` that draws `spec`, if its row is 3-D and registered. */
export function transition3DFor(spec: Pick<TransitionSpec, 'ns' | 'element' | 'preset'>): Transition3D | undefined {
  const row = rowForTransition(spec.ns, spec.element, spec.preset)
  if (!row || row.stage !== 3 || row.ns === 'p159') return undefined
  return registry.get(keyOf(spec.element, spec.preset))
}

/** Every registered transition, in registration order. */
export function registeredTransitions3D(): readonly Transition3D[] {
  return [...registry.values()]
}

/** Test-only: empties the registry. */
export function resetTransitions3D(): void {
  registry.clear()
}
