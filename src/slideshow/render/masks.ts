/**
 * The mask library (slideshow spec 6.3): procedural geometry for every
 * `MaskKind`, shared by the Canvas 2D compositor and the WebGL2 compositor's
 * tests.
 *
 * A mask is the region where the incoming (or masked) layer shows, in a box
 * `w × h` with its origin at the top-left. Progress 0 shows nothing and
 * progress 1 shows the whole box; values outside [0, 1] are clamped.
 *
 * Geometry is a backend-neutral `MaskPath` (a list of closed sub-paths with a
 * fill rule), not a `Path2D`: node-canvas and jsdom have no `Path2D`, and the
 * WebGL2 backend needs the shapes, not a browser object. `tracePath` replays a
 * path onto any canvas context or `Path2D`; `toPath2D` builds one in a
 * browser. `dissolve` and `randomBar` are noise a path cannot express and
 * come back as a `MaskRaster`, stretched over the box with nearest-neighbour
 * sampling.
 *
 * Params use geometric names, never ECMA-376 attribute values; mapping a
 * `p:transition` or `p:animEffect` filter onto them is the caller's job
 * (S1-15, S1-11):
 * - `wipe`: `from` `'l' | 'r' | 't' | 'b'` (default `'l'`), the edge the reveal starts at.
 * - `split`, `barn`: `orient` `'horz' | 'vert'` (default `'horz'`: horizontal
 *   bands, the seam runs left to right), `dir` `'out' | 'in'` (default `'out'`:
 *   from the centre outward).
 * - `box`, `circle`, `diamond`, `plus`: `dir` `'out' | 'in'` (default `'out'`:
 *   grows from the centre; `'in'` closes from the edges toward the centre).
 * - `wedge`: two arms sweeping from 12 o'clock both ways around the centre.
 * - `wheel`: `spokes` (default 1, at most 64), each sector sweeping clockwise from 12 o'clock;
 *   `dir` `'cw' | 'ccw'` (default `'cw'`; `'ccw'` sweeps counter-clockwise, S1-15 `wheelReverse`).
 * - `checkerboard`: `dir` `'across' | 'down'` (default `'across'`), `cells`
 *   (default 8, at most 128) squares across the box; odd rows are offset by
 *   half a cell.
 * - `blinds`: `orient` (default `'horz'`), `count` (default 6, at most 256)
 *   bands, each revealed from its top (left).
 * - `comb`: `orient` (default `'horz'`), `count` (default 8, at most 256) bands
 *   revealed from alternate sides, the first from the left (top).
 * - `strips`: `from` `'lu' | 'ru' | 'ld' | 'rd'` (default `'lu'`), the corner a
 *   diagonal reveal starts at.
 * - `dissolve`: `cell` (default 8) box units per square cell.
 * - `randomBar`: `orient` (default `'horz'`: horizontal bars), `bar` (default 4)
 *   box units per bar.
 *
 * Op counts are bounded: a path never has more than `MASK_MAX_OPS` sub-paths.
 * A mask whose params would need more (`maskOpCount`) is a cut instead:
 * nothing before progress 0.5, the whole box from 0.5. Callers diagnose it.
 *
 * Noise masks are monotone in progress and exact in count: a fixed shuffle
 * from `rng` ranks the cells, and progress `p` shows the first
 * `floor(p × cells)`. The shuffle consumes the same draws at every progress,
 * so one seed gives one pattern for the whole effect.
 */
import type { MaskKind, MaskState, Rng } from '../types'

export interface MaskRaster { width: number; height: number; alpha: Uint8ClampedArray }

/** A closed sub-path. `sector` is a pie slice clockwise from `a0` to `a1` (radians, canvas convention). */
export type MaskOp =
  | { op: 'rect'; x: number; y: number; w: number; h: number }
  | { op: 'poly'; pts: readonly number[] }
  | { op: 'sector'; cx: number; cy: number; r: number; a0: number; a1: number }

export interface MaskPath { readonly fillRule: 'nonzero' | 'evenodd'; readonly ops: readonly MaskOp[] }

export type MaskParams = MaskState['params']

/** The `CanvasPath` subset a path is traced onto: a 2-D context or a `Path2D`. */
export interface PathSink {
  moveTo(x: number, y: number): void
  lineTo(x: number, y: number): void
  arc(x: number, y: number, r: number, a0: number, a1: number, ccw?: boolean): void
  rect(x: number, y: number, w: number, h: number): void
  closePath(): void
}

const KIND_SET: Record<MaskKind, true> = {
  wipe: true, split: true, circle: true, diamond: true, plus: true, wedge: true, wheel: true, barn: true,
  checkerboard: true, dissolve: true, blinds: true, box: true, comb: true, randomBar: true, strips: true,
}
/** Every `MaskKind`, at run time. */
export const MASK_KINDS = Object.keys(KIND_SET) as readonly MaskKind[]

export function isMaskKind(kind: unknown): kind is MaskKind {
  return typeof kind === 'string' && Object.prototype.hasOwnProperty.call(KIND_SET, kind)
}

export function isMaskRaster(m: MaskPath | MaskRaster): m is MaskRaster {
  return 'alpha' in m
}

/** The most sub-paths one mask path may hold; more degrades to a cut (bounded per-frame allocation). */
export const MASK_MAX_OPS = 4096
const MAX_CELLS = 128
const MAX_BANDS = 256
const MAX_SPOKES = 64

/** Shapes that grow from the centre overshoot the box by this much, so anti-aliasing never leaks at 0 or 1. */
const PAD = 1
const TAU = Math.PI * 2
const TOP = -Math.PI / 2

function str<T extends string>(params: MaskParams, key: string, allowed: readonly T[], fallback: T): T {
  const v = params[key]
  return typeof v === 'string' && (allowed as readonly string[]).includes(v) ? (v as T) : fallback
}
function int(params: MaskParams, key: string, fallback: number, max: number): number {
  const v = Number(params[key])
  return Number.isFinite(v) && v >= 1 ? Math.min(max, Math.floor(v)) : fallback
}
function size(params: MaskParams, key: string, fallback: number): number {
  const v = Number(params[key])
  return Number.isFinite(v) && v > 0 ? v : fallback
}

const path = (ops: MaskOp[], fillRule: MaskPath['fillRule'] = 'nonzero'): MaskPath => ({ fillRule, ops })
const rect = (x: number, y: number, w: number, h: number): MaskOp => ({ op: 'rect', x, y, w, h })

/** Raster grid size for a noise kind, or undefined for a path kind. */
function rasterGrid(kind: MaskKind, params: MaskParams, w: number, h: number): { cols: number; rows: number } | undefined {
  if (kind === 'dissolve') {
    const cell = size(params, 'cell', 8)
    return { cols: Math.max(1, Math.ceil(w / cell)), rows: Math.max(1, Math.ceil(h / cell)) }
  }
  if (kind === 'randomBar') {
    const bar = size(params, 'bar', 4)
    return str(params, 'orient', ['horz', 'vert'], 'horz') === 'horz'
      ? { cols: 1, rows: Math.max(1, Math.ceil(h / bar)) }
      : { cols: Math.max(1, Math.ceil(w / bar)), rows: 1 }
  }
  return undefined
}

/** Bytes `maskPath` allocates for a noise kind (the shuffle plus the alpha), 0 for a path kind. Lease this first. */
export function maskRasterBytes(kind: MaskKind, params: MaskParams, w: number, h: number): number {
  const g = rasterGrid(kind, params, w, h)
  return g ? g.cols * g.rows * 5 : 0
}

/** Sub-paths `maskPath` would build for a path kind at a progress strictly between 0 and 1; 0 for a noise kind. */
export function maskOpCount(kind: MaskKind, params: MaskParams, w: number, h: number): number {
  switch (kind) {
    case 'dissolve': case 'randomBar': return 0
    case 'split': case 'barn': return str(params, 'dir', ['out', 'in'], 'out') === 'out' ? 1 : 2
    case 'box': case 'circle': case 'diamond': case 'plus': return str(params, 'dir', ['out', 'in'], 'out') === 'out' ? 1 : 2
    case 'wheel': return int(params, 'spokes', 1, MAX_SPOKES)
    case 'blinds': case 'comb': return int(params, 'count', kind === 'blinds' ? 6 : 8, MAX_BANDS)
    case 'checkerboard': {
      const across = str(params, 'dir', ['across', 'down'], 'across') === 'across'
      const cells = int(params, 'cells', 8, MAX_CELLS)
      const [len, breadth] = across ? [w, h] : [h, w]
      if (!(len > 0 && breadth > 0)) return 0
      return Math.ceil(breadth / (len / cells)) * (cells + 1)
    }
    default: return 1
  }
}

/** Mirror a polygon of `[x, y, …]` points about the centre lines. */
function mirror(pts: number[], w: number, h: number, flipX: boolean, flipY: boolean): number[] {
  return pts.map((v, i) => (i % 2 === 0 ? (flipX ? w - v : v) : (flipY ? h - v : v)))
}

/** The `'out'` shape at extent `t` (0 empty, 1 covers the box), centred. */
function centred(kind: 'box' | 'circle' | 'diamond' | 'plus', t: number, w: number, h: number): MaskOp {
  const cx = w / 2, cy = h / 2
  const hx = w / 2 + PAD, hy = h / 2 + PAD
  switch (kind) {
    case 'box': return rect(cx - t * hx, cy - t * hy, 2 * t * hx, 2 * t * hy)
    case 'circle': return { op: 'sector', cx, cy, r: t * (Math.hypot(w, h) / 2 + PAD), a0: TOP, a1: TOP + TAU }
    case 'diamond': {
      const a = t * 2 * hx, b = t * 2 * hy
      return { op: 'poly', pts: [cx, cy - b, cx + a, cy, cx, cy + b, cx - a, cy] }
    }
    case 'plus': {
      const tx = t * hx, ty = t * hy, reach = Math.min(1, 2 * t), ax = reach * hx, ay = reach * hy
      return {
        op: 'poly',
        pts: [
          cx - tx, cy - ay, cx + tx, cy - ay, cx + tx, cy - ty, cx + ax, cy - ty, cx + ax, cy + ty, cx + tx, cy + ty,
          cx + tx, cy + ay, cx - tx, cy + ay, cx - tx, cy + ty, cx - ax, cy + ty, cx - ax, cy - ty, cx - tx, cy - ty,
        ],
      }
    }
  }
}

function noise(grid: { cols: number; rows: number }, p: number, rng: Rng): MaskRaster {
  const n = grid.cols * grid.rows
  const order = new Uint32Array(n)
  for (let i = 0; i < n; i++) order[i] = i
  for (let i = n - 1; i > 0; i--) {
    const j = Math.floor(rng() * (i + 1))
    const t = order[i]; order[i] = order[j]; order[j] = t
  }
  const alpha = new Uint8ClampedArray(n)
  const shown = Math.min(n, Math.floor(p * n + 1e-9))
  for (let k = 0; k < shown; k++) alpha[order[k]] = 255
  return { width: grid.cols, height: grid.rows, alpha }
}

/**
 * The region of a `w × h` box shown at `progress`. Noise kinds use `rng`
 * (pass a freshly seeded generator per frame for a stable pattern) and
 * allocate `maskRasterBytes`; path kinds never call it.
 */
export function maskPath(kind: MaskKind, progress: number, params: MaskParams, w: number, h: number, rng: Rng): MaskPath | MaskRaster {
  const p = Number.isFinite(progress) ? Math.min(1, Math.max(0, progress)) : 0
  const grid = rasterGrid(kind, params, w, h)
  if (grid) return noise(grid, p, rng)
  if (p <= 0) return path([])
  if (p >= 1) return path([rect(0, 0, w, h)])
  if (maskOpCount(kind, params, w, h) > MASK_MAX_OPS) return p < 0.5 ? path([]) : path([rect(0, 0, w, h)])

  switch (kind) {
    case 'wipe': {
      switch (str(params, 'from', ['l', 'r', 't', 'b'], 'l')) {
        case 'l': return path([rect(0, 0, p * w, h)])
        case 'r': return path([rect(w - p * w, 0, p * w, h)])
        case 't': return path([rect(0, 0, w, p * h)])
        case 'b': return path([rect(0, h - p * h, w, p * h)])
      }
      break
    }
    case 'split':
    case 'barn': {
      const horz = str(params, 'orient', ['horz', 'vert'], 'horz') === 'horz'
      const out = str(params, 'dir', ['out', 'in'], 'out') === 'out'
      if (horz) {
        return out
          ? path([rect(0, h / 2 - (p * h) / 2, w, p * h)])
          : path([rect(0, 0, w, (p * h) / 2), rect(0, h - (p * h) / 2, w, (p * h) / 2)])
      }
      return out
        ? path([rect(w / 2 - (p * w) / 2, 0, p * w, h)])
        : path([rect(0, 0, (p * w) / 2, h), rect(w - (p * w) / 2, 0, (p * w) / 2, h)])
    }
    case 'box':
    case 'circle':
    case 'diamond':
    case 'plus': {
      if (str(params, 'dir', ['out', 'in'], 'out') === 'out') return path([centred(kind, p, w, h)])
      // Closing from the edges: everything outside the shape at 1 − p.
      return path([rect(-PAD, -PAD, w + 2 * PAD, h + 2 * PAD), centred(kind, 1 - p, w, h)], 'evenodd')
    }
    case 'wedge': {
      const r = Math.hypot(w, h) / 2 + PAD
      return path([{ op: 'sector', cx: w / 2, cy: h / 2, r, a0: TOP - p * Math.PI, a1: TOP + p * Math.PI }])
    }
    case 'wheel': {
      // 1 spoke is the layer-mask (wheel animation) default; transitionDraw passes 4, PowerPoint's wheel-transition default.
      const spokes = int(params, 'spokes', 1, MAX_SPOKES)
      const r = Math.hypot(w, h) / 2 + PAD
      const step = TAU / spokes
      const ops: MaskOp[] = []
      const ccw = str(params, 'dir', ['cw', 'ccw'], 'cw') === 'ccw'
      for (let i = 0; i < spokes; i++) {
        const at = TOP + i * step
        ops.push({ op: 'sector', cx: w / 2, cy: h / 2, r, a0: ccw ? at - p * step : at, a1: ccw ? at : at + p * step })
      }
      return path(ops)
    }
    case 'checkerboard': {
      const across = str(params, 'dir', ['across', 'down'], 'across') === 'across'
      const cells = int(params, 'cells', 8, MAX_CELLS)
      const ops: MaskOp[] = []
      const [len, breadth] = across ? [w, h] : [h, w]
      const cell = len / cells
      const lines = Math.ceil(breadth / cell)
      for (let r = 0; r < lines; r++) {
        const offset = (r % 2) * (cell / 2)
        for (let k = -1; k < cells; k++) {
          const a = k * cell + offset
          ops.push(across ? rect(a, r * cell, p * cell, cell) : rect(r * cell, a, cell, p * cell))
        }
      }
      return path(ops)
    }
    case 'blinds':
    case 'comb': {
      const horz = str(params, 'orient', ['horz', 'vert'], 'horz') === 'horz'
      const count = int(params, 'count', kind === 'blinds' ? 6 : 8, MAX_BANDS)
      const ops: MaskOp[] = []
      const band = (horz ? h : w) / count
      for (let i = 0; i < count; i++) {
        const at = i * band
        if (kind === 'blinds') ops.push(horz ? rect(0, at, w, p * band) : rect(at, 0, p * band, h))
        else {
          const lead = i % 2 === 0
          ops.push(horz
            ? rect(lead ? 0 : w - p * w, at, p * w, band)
            : rect(at, lead ? 0 : h - p * h, band, p * h))
        }
      }
      return path(ops)
    }
    case 'strips': {
      const from = str(params, 'from', ['lu', 'ru', 'ld', 'rd'], 'lu')
      const tri = [0, 0, 2 * p * w, 0, 0, 2 * p * h]
      return path([{ op: 'poly', pts: mirror(tri, w, h, from[0] === 'r', from[1] === 'd') }])
    }
  }
  // Unreachable for a MaskKind; a caller with an unchecked string should test isMaskKind first.
  return path([rect(0, 0, w, h)])
}

/**
 * Transition backdrops (S1-15): the colours a `backdrop` param paints, shared
 * by both compositors (Canvas 2D `fillRect`, WebGL2 scissored `clear`).
 */
export const BACKDROP_RGB = Object.freeze({ black: [0, 0, 0], white: [255, 255, 255] } as const)
/**
 * The integer device rect a transition is clipped to and its backdrop fills:
 * the slide's device rect with its edges rounded, so both backends cover the
 * same pixels and no frame spills into the letterbox.
 */
export function transitionRect(rect: { x: number; y: number; w: number; h: number }): { x: number; y: number; w: number; h: number } {
  const x = Math.round(rect.x), y = Math.round(rect.y)
  return { x, y, w: Math.max(0, Math.round(rect.x + rect.w) - x), h: Math.max(0, Math.round(rect.y + rect.h) - y) }
}

/** Replay a path onto a context or `Path2D`, offset by `(dx, dy)`. The caller begins the path and fills or clips it. */
export function tracePath(target: PathSink, m: MaskPath, dx = 0, dy = 0): void {
  for (const op of m.ops) {
    switch (op.op) {
      case 'rect':
        target.rect(op.x + dx, op.y + dy, op.w, op.h)
        break
      case 'poly':
        target.moveTo(op.pts[0] + dx, op.pts[1] + dy)
        for (let i = 2; i < op.pts.length; i += 2) target.lineTo(op.pts[i] + dx, op.pts[i + 1] + dy)
        target.closePath()
        break
      case 'sector':
        target.moveTo(op.cx + dx, op.cy + dy)
        target.arc(op.cx + dx, op.cy + dy, op.r, op.a0, op.a1, false)
        target.closePath()
        break
    }
  }
}

/** A browser `Path2D` for a mask path. Pass the constructor from the target realm. */
export function toPath2D(m: MaskPath, Ctor: typeof Path2D = globalThis.Path2D): Path2D {
  const p = new Ctor()
  tracePath(p, m)
  return p
}
