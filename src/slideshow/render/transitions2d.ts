/**
 * The 2-D slide transitions (slideshow spec 6.3 transitions, 4.5 reduced
 * motion, 10 seeded randomness).
 *
 * `transitionDraw` maps a `TransitionSpec` at a progress onto the params both
 * compositors draw (canvas2d.ts, webgl2.ts): a transition `mask` over the
 * incoming frame, `from` / `to` affines in unit slide space, `fromOpacity` /
 * `toOpacity`, and two additive params, `fromOnTop` (the outgoing frame is
 * drawn over the incoming one, for uncover) and `backdrop` (an opaque colour
 * painted over the slide rect under both frames, for "through black" and
 * flash). It is pure and allocates nothing the workspace charges; no
 * offscreen framebuffer is needed.
 *
 * Conventions:
 * - **Directions.** ECMA-376 `dir` on `push`, `cover`, `pull`, `wipe`,
 *   `strips` (and `p14:pan`, `p14:reveal`) is the direction of motion
 *   (LibreOffice's importer reads it the same way): `wipe dir="d"` moves
 *   down, so the reveal starts at the top edge; `push dir="l"` moves left, so
 *   the incoming slide enters from the right. The mask library's `from` is
 *   the edge or corner a reveal starts at, so it is the opposite of `dir`.
 *   `split` `orient`/`dir` and `blinds`, `comb`, `randomBar` `dir`
 *   (`horz`/`vert`) map onto the mask params of the same meaning; `checker`
 *   `horz` is the mask's `across`.
 * - **Ends.** Progress is clamped to [0, 1]. Every effect shows exactly the
 *   outgoing frame at 0 and the incoming frame at 1 (identity affines,
 *   opacity 1, or a full mask); a zero-scale incoming frame is hidden by
 *   opacity, never by a degenerate matrix.
 * - **Rotation** (`newsflash`) is conformal on the slide, which needs the
 *   slide's aspect (width / height): the optional fourth argument, default
 *   16 : 9.
 * - **Defaults** for an omitted attribute are PowerPoint's (probed
 *   `PpEntryEffect` write-outs, S1-15 review): push, pan, cover, pull, wipe
 *   `l`; strips `lu`; zoom `out`; flythrough `in`; wheel 4 spokes.
 * - **`random`** picks one 2-D effect and its attributes from the show's
 *   seeded `rng`, once per rng identity, so every frame of the effect draws
 *   the same pick (`randomPick`).
 * - **Not 2-D.** A catalogued stand-in (3-D and texture-heavy rows) draws its
 *   row's 2-D look-alike; an uncatalogued element draws a fade, and `none`
 *   (no effect element) a cut. Diagnosing those is the caller's job
 *   (`resolveTransition2D` says what is drawn), never per frame.
 */
import { CATALOGUE, rowForTransition } from '../catalogue'
import type { Affine, MaskKind, MaskState, Rng, TransitionDraw, TransitionSpec } from '../types'

/** Reduced motion (spec 4.5): transitions become a crossfade of at most this long. */
export const REDUCED_TRANSITION_MS = 200

/** Params `transitionDraw` produces; every field is optional and drawn by both compositors. */
export interface TransitionParams2D {
  mask?: MaskState
  from?: Affine
  to?: Affine
  fromOpacity?: number
  toOpacity?: number
  /** Additive (S1-15): draw the outgoing frame over the incoming one. */
  fromOnTop?: boolean
  /** Additive (S1-15): an opaque colour over the slide rect, under both frames. */
  backdrop?: 'black' | 'white'
}

/** The catalogue rows drawn natively here: every stage-1 transition row (spec 6.3 2-D list). */
export const TRANSITION_2D_IDS: readonly string[] = Object.freeze(
  CATALOGUE.filter(r => r.kind === 'transition' && r.stage === 1).map(r => r.id))

const NATIVE: ReadonlySet<string> = new Set(
  CATALOGUE.filter(r => r.kind === 'transition' && r.stage === 1).map(r => r.element!))

type Params = TransitionDraw['params'] & TransitionParams2D

const clamp01 = (v: number) => (Number.isFinite(v) ? Math.min(1, Math.max(0, v)) : 0)
/** PowerPoint movie probes (0.5–4 s fades) fit a sine ease-in-out, not linear opacity. */
export const easeFade = (progress: number) => 0.5 - 0.5 * Math.cos(Math.PI * clamp01(progress))
const pick = <T extends string>(v: string | undefined, allowed: readonly T[], fallback: T): T =>
  v !== undefined && (allowed as readonly string[]).includes(v) ? (v as T) : fallback
const flag = (v: string | undefined) => v === '1' || v === 'true'

const SIDES = ['l', 'u', 'r', 'd'] as const
const EIGHT = ['l', 'u', 'r', 'd', 'lu', 'ru', 'ld', 'rd'] as const
const CORNERS = ['lu', 'ru', 'ld', 'rd'] as const
const ORIENT = ['horz', 'vert'] as const
const IN_OUT = ['in', 'out'] as const

/** Unit motion vector of a direction of motion. */
function motion(dir: string): [number, number] {
  const x = dir.includes('l') ? -1 : dir.includes('r') ? 1 : 0
  const y = dir.includes('u') ? -1 : dir.includes('d') ? 1 : 0
  return [x, y]
}
const translate = (x: number, y: number): Affine => [1, 0, 0, 1, x, y]
/** Uniform scale `s` and clockwise rotation `angle` about the slide centre, conformal for aspect `aspect`. */
function aboutCentre(s: number, angle = 0, aspect = 1): Affine {
  const cos = angle === 0 ? 1 : Math.cos(angle), sin = angle === 0 ? 0 : Math.sin(angle)
  const a = s * cos, b = s * sin * aspect, c = (-s * sin) / aspect, d = s * cos
  return [a, b, c, d, 0.5 - (a + c) / 2, 0.5 - (b + d) / 2]
}
const OPPOSITE_SIDE: Record<string, string> = { l: 'r', r: 'l', u: 'b', d: 't' }
const OPPOSITE_CORNER: Record<string, string> = { lu: 'rd', ru: 'ld', ld: 'ru', rd: 'lu' }

const masked = (kind: MaskKind, p: number, params: MaskState['params'] = {}): Params => ({ mask: { kind, progress: p, params } })
const spokes = (v: string | undefined) => {
  const n = Number(v)
  return Number.isInteger(n) && n >= 1 ? n : 4
}

/** Fade out to the backdrop over the first half, fade in from it over the second. */
function through(backdrop: 'black' | 'white', p: number, from?: Affine, to?: Affine): Params {
  return {
    backdrop,
    fromOpacity: Math.max(0, 1 - 2 * p),
    toOpacity: Math.max(0, 2 * p - 1),
    ...(from ? { from } : {}),
    ...(to ? { to } : {}),
  }
}

/** The spec that is actually drawn: the native effect, a stand-in's look-alike, a fade, or a cut for `none`. */
export function resolveTransition2D(spec: TransitionSpec): { spec: TransitionSpec; approximated: boolean } {
  if (NATIVE.has(spec.element) && rowForTransition(spec.ns, spec.element, spec.preset)) return { spec, approximated: false }
  if (spec.element === 'none') return { spec: { ...spec, element: 'cut', ns: 'p', attrs: {} }, approximated: false }
  const row = rowForTransition(spec.ns, spec.element, spec.preset)
  const look = row?.lookalike ? CATALOGUE.find(r => r.id === row.lookalike) : undefined
  const element = look?.element && NATIVE.has(look.element) ? look.element : 'fade'
  const { preset: _preset, ...rest } = spec
  return { spec: { ...rest, element, ns: 'p', attrs: {} }, approximated: true }
}

/** Effects `random` chooses from, with the attribute values each may take. */
const RANDOM_POOL: ReadonlyArray<{ element: string; attrs: Readonly<Record<string, readonly string[]>> }> = [
  { element: 'blinds', attrs: { dir: ORIENT } },
  { element: 'checker', attrs: { dir: ORIENT } },
  { element: 'circle', attrs: {} },
  { element: 'comb', attrs: { dir: ORIENT } },
  { element: 'cover', attrs: { dir: EIGHT } },
  { element: 'diamond', attrs: {} },
  { element: 'dissolve', attrs: {} },
  { element: 'fade', attrs: { thruBlk: ['0', '1'] } },
  { element: 'newsflash', attrs: {} },
  { element: 'plus', attrs: {} },
  { element: 'pull', attrs: { dir: EIGHT } },
  { element: 'push', attrs: { dir: SIDES } },
  { element: 'randomBar', attrs: { dir: ORIENT } },
  { element: 'split', attrs: { orient: ORIENT, dir: IN_OUT } },
  { element: 'strips', attrs: { dir: CORNERS } },
  { element: 'wedge', attrs: {} },
  { element: 'wheel', attrs: { spokes: ['1', '2', '3', '4', '8'] } },
  { element: 'wipe', attrs: { dir: SIDES } },
  { element: 'zoom', attrs: { dir: IN_OUT } },
]
const picks = new WeakMap<Rng, TransitionSpec>()

/**
 * The effect `random` draws for this rng: chosen on first use, then the same
 * for every later call with the same rng. Only `element`, `ns` and `attrs`
 * matter to `transitionDraw`.
 */
export function randomPick(rng: Rng): TransitionSpec {
  const cached = picks.get(rng)
  if (cached) return cached
  const draw = (n: number) => Math.min(n - 1, Math.floor(clamp01(rng()) * n))
  const choice = RANDOM_POOL[draw(RANDOM_POOL.length)]
  const attrs: Record<string, string> = {}
  for (const [name, values] of Object.entries(choice.attrs)) attrs[name] = values[draw(values.length)]
  const spec: TransitionSpec = Object.freeze({ element: choice.element, ns: 'p', attrs: Object.freeze(attrs), durationMs: 0, advClick: true })
  picks.set(rng, spec)
  return spec
}

/**
 * Compositor params for `spec` at `progress` (0 = outgoing, 1 = incoming).
 * `aspect` is the slide's width / height, used only for rotation.
 */
export function transitionDraw(spec: TransitionSpec, progress: number, rng: Rng, aspect = 16 / 9): Params {
  const p = clamp01(progress)
  const resolved = spec.element === 'random' ? randomPick(rng) : resolveTransition2D(spec).spec
  const a = resolved.attrs
  switch (resolved.element) {
    case 'blinds': return masked('blinds', p, { orient: pick(a.dir, ORIENT, 'horz') })
    case 'checker': return masked('checkerboard', p, { dir: pick(a.dir, ORIENT, 'horz') === 'horz' ? 'across' : 'down' })
    case 'circle': return masked('circle', p)
    case 'comb': return masked('comb', p, { orient: pick(a.dir, ORIENT, 'horz') })
    case 'diamond': return masked('diamond', p)
    case 'dissolve': return masked('dissolve', p)
    case 'plus': return masked('plus', p)
    case 'randomBar': return masked('randomBar', p, { orient: pick(a.dir, ORIENT, 'horz') })
    case 'split': return masked('split', p, { orient: pick(a.orient, ORIENT, 'horz'), dir: pick(a.dir, IN_OUT, 'out') })
    case 'strips': return masked('strips', p, { from: OPPOSITE_CORNER[pick(a.dir, CORNERS, 'lu')] })
    case 'wedge': return masked('wedge', p)
    case 'wheel': return masked('wheel', p, { spokes: spokes(a.spokes) })
    case 'wheelReverse': return masked('wheel', p, { spokes: spokes(a.spokes), dir: 'ccw' })
    case 'wipe': return masked('wipe', p, { from: OPPOSITE_SIDE[pick(a.dir, SIDES, 'l')] })

    case 'cut':
      if (flag(a.thruBlk)) return { backdrop: 'black', fromOpacity: p < 1 / 3 ? 1 : 0, toOpacity: p >= 2 / 3 ? 1 : 0 }
      return { toOpacity: p >= 0.5 ? 1 : 0 }
    case 'fade':
      return flag(a.thruBlk) ? through('black', p) : { toOpacity: easeFade(p) }
    case 'flash':
      return through('white', p)

    case 'push':
    case 'pan': {
      // Omitted dir is "Push Left" / "Pan Left" in PowerPoint (PpEntryEffect 3853 / 3930).
      const [x, y] = motion(pick(a.dir, SIDES, 'l'))
      return { from: translate(x * p, y * p), to: translate(x * (p - 1), y * (p - 1)) }
    }
    case 'cover': {
      const [x, y] = motion(pick(a.dir, EIGHT, 'l'))
      return { to: translate(x * (p - 1), y * (p - 1)) }
    }
    case 'pull': {
      const [x, y] = motion(pick(a.dir, EIGHT, 'l'))
      return { from: translate(x * p, y * p), fromOnTop: true }
    }
    case 'reveal': {
      const [x] = motion(pick(a.dir, ['l', 'r'] as const, 'l'))
      const from = translate(0.25 * x * p, 0), to = translate(0.25 * x * (p - 1), 0)
      return flag(a.thruBlk) ? through('black', p, from, to) : { to, toOpacity: p }
    }
    case 'zoom':
      if (pick(a.dir, IN_OUT, 'out') === 'in') return { to: aboutCentre(p), toOpacity: p > 0 ? 1 : 0 }
      return { to: aboutCentre(2 - p), toOpacity: p }
    case 'newsflash':
      return { to: aboutCentre(p, 2 * Math.PI * (1 - p), aspect), toOpacity: p > 0 ? 1 : 0 }
    case 'flythrough':
      if (pick(a.dir, IN_OUT, 'in') === 'in') return { from: aboutCentre(1 + 0.5 * p), to: aboutCentre(0.5 + 0.5 * p), toOpacity: p }
      return { to: aboutCentre(1.5 - 0.5 * p), toOpacity: p }
  }
  // resolveTransition2D only returns native elements; a fade keeps the owner rule if one slips through.
  return { toOpacity: p }
}

/**
 * Reduced motion (spec 4.5): any transition becomes a crossfade of at most
 * `REDUCED_TRANSITION_MS`. The advance (`advClick`, `advTmMs`) and the sound
 * are kept: they are not motion. Reduced motion never adds motion (S1-15
 * fix round 1 ruling): `cut` (with or without `thruBlk`) and `none` are
 * already instant and come back unchanged.
 */
export function reducedTransition(spec: TransitionSpec): TransitionSpec {
  if (spec.element === 'none' || (spec.element === 'cut' && spec.ns === 'p')) return spec
  const { preset: _preset, ...rest } = spec
  const d = Number.isFinite(spec.durationMs) ? Math.max(0, spec.durationMs) : REDUCED_TRANSITION_MS
  return { ...rest, element: 'fade', ns: 'p', attrs: {}, durationMs: Math.min(d, REDUCED_TRANSITION_MS) }
}
