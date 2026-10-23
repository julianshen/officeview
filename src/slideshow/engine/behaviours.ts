// S1-11: behaviour evaluation to render state (slideshow spec 6.2; §6.1 cBhvr
// composition; §6.3 part-addressed colour). The frame loop feeds a scheduler
// snapshot; this module resolves composition, formulas, masks and colour.
//
// Reading rules settled here (documented in the task report):
//  - Formula env values are fractions: ppt_x/y/w/h are fractions of the slide
//    (ECMA-376 §19.5.4), ppt_x/ppt_y the target's CENTRE (S1-22: as
//    LibreOffice reads X/Y, so "from (0-#ppt_w/2)" starts just off the left
//    edge and "1+#ppt_h/2" just below the slide), ppt_r is degrees, `$` = the behaviour's eased local
//    progress. Raw anim numbers (from/by/to, tav values, set `to`) are
//    fractions; animRot carries degrees, animScale factors, animMotion
//    slide-fraction offsets.
//  - `#name` formula references read the target's UNDERLYING value (the base
//    geometry with no animation applied).
//  - Additive composition in document order:
//      base/repl/none → C = V; sum → C = C + V; mult → C = C × V.
//    The FIRST edit seeds the chain from the track's base value (the geometry
//    fraction, geo rotation, opacity 1, scale 1).
//  - accumulate=always adds k·(V(1) − V(0)) of one iteration for repeat k.
//  - ppt_x/y/w/h tracks are absolute positions/sizes; the final offsets are
//    (C·slide − base) px. animRot's by value is relative (base rotation +
//    by·p); from/to are absolute. animScale factors scale about the base
//    centre. animMotion emits px offsets directly (origin layout subtracts the
//    base position, parent does not).
//  - animEffect filters map onto MaskState kinds; the fade filter (and
//    unknown filters, diagnosed) render as an opacity ramp.
//  - animClr writes colour matrices (SVG feColorMatrix 4x5 row-major) onto
//    color.fill / color.line / color.text per attr name.
import { compileFormula } from '../formula'
import { pointAt } from '../motion-path'
import { targetKey } from '../target'
import { easeProgress } from './ease'
import { rebuildFromPreset } from './presets'
import { standInFor, transitionStandIn } from './stand-ins'
import type { FormulaEnv } from '../formula'
import type { SchedulerSnapshot } from './scheduler'
import type {
  Affine, AnimBehaviour, AnimClrBehaviour, AnimEffectBehaviour, AnimMotionBehaviour, AnimRotBehaviour,
  AnimScaleBehaviour, BehaviourNode, ColorMatrix, DiagnosticSink, MaskState, Rng,
  ContainerNode, Iterate, SetBehaviour, TargetKey, TargetState, TimeNode, TimingTree,
} from '../types'

export interface BaseGeometry { x: number; y: number; w: number; h: number; rot: number }
export { rebuildFromPreset, standInFor, transitionStandIn }

const RGB_IDENTITY: ColorMatrix = [1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0]
type Slide = { w: number; h: number }

// ---------------------------------------------------------------------------
// Colour helpers (SVG feColorMatrix convention: rows r/g/b/a, 5th column offsets)
// ---------------------------------------------------------------------------

type Rgb = { r: number; g: number; b: number }

/** `#hex`, `rgb(f,f,f)` and `hslClr:h,s,l` fractional colour strings (S1-8 output). */
export function readRgb(src: string): Rgb | undefined {
  if (/^#[0-9a-fA-F]{6}$/.test(src)) {
    return { r: parseInt(src.slice(1, 3), 16) / 255, g: parseInt(src.slice(3, 5), 16) / 255, b: parseInt(src.slice(5, 7), 16) / 255 }
  }
  if (/^rgb\(|^hslClr:/.test(src)) {
    const parts = src.replace(/^(rgb\(|hslClr:)/, '').replace(/\)$/, '').split(',').map(v => Number(v.trim()))
    if (parts.length === 3 && parts.every(Number.isFinite)) return { r: parts[0], g: parts[1], b: parts[2] }
  }
  return undefined
}

/** Hue-rotation degrees of an `hsl(h,s,l)`/`hslClr:h,...` by-offset, or undefined for rgb. */
export function hueDegrees(src: string): number | undefined {
  if (!/^hsl\(|^hslClr:/.test(src)) return undefined
  const parts = src.replace(/^(hsl\(|hslClr:)/, '').replace(/\)$/, '').split(',').map(v => Number(v.trim()))
  return parts.length === 3 && Number.isFinite(parts[0]) ? parts[0] : undefined
}

/** The SVG feColorMatrix hueRotate matrix for a rotation of `deg` degrees. */
export function hueMatrix(deg: number): ColorMatrix {
  const a = (deg * Math.PI) / 180
  const c = Math.cos(a), s = Math.sin(a)
  return [
    0.213 + 0.787 * c - 0.213 * s, 0.715 - 0.715 * c - 0.715 * s, 0.072 - 0.072 * c + 0.928 * s, 0, 0,
    0.213 - 0.213 * c + 0.143 * s, 0.715 + 0.285 * c + 0.14 * s, 0.072 - 0.072 * c - 0.283 * s, 0, 0,
    0.213 - 0.213 * c - 0.787 * s, 0.715 - 0.715 * c + 0.715 * s, 0.072 + 0.928 * c - 0.072 * s, 0, 0,
    0, 0, 0, 1, 0,
  ]
}

/** Compose m2 after m1 (apply m1, then m2). */
export function composeColor(m1: ColorMatrix, m2: ColorMatrix): ColorMatrix {
  const out: number[] = new Array(20)
  for (let row = 0; row < 4; row++) {
    for (let col = 0; col < 5; col++) {
      let sum = 0
      for (let k = 0; k < 4; k++) sum += m2[row * 5 + k] * m1[k * 5 + col]
      if (col === 4) sum += m2[row * 5 + 4]
      out[row * 5 + col] = sum
    }
  }
  return out
}

/** A scale-toward-`target` matrix expressed against a `from` baseline source. */
function matrixScale(target: Rgb, from: Rgb): ColorMatrix {
  const m = [...RGB_IDENTITY] as number[]
  const chans = [target.r, target.g, target.b]
  const froms = [from.r, from.g, from.b]
  for (let i = 0; i < 3; i++) {
    const row = i * 5
    if (chans[i] === 0 && froms[i] === 0) {
      m[row + i] = 1
    } else if (froms[i] > 0) {
      m[row + i] = chans[i] / froms[i]
    } else {
      m[row + 4] = chans[i]
    }
  }
  return m
}

/** Compose one affine transform on top of another (b applied first). */
function mulAffine(b: Affine, a: Affine): Affine {
  return [
    a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1],
    a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3],
    a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5],
  ]
}

// ---------------------------------------------------------------------------
// Formula evaluation (compile once per behaviour slot; NaN/unsupported diagnosed)
// ---------------------------------------------------------------------------

type Compiled = (env: FormulaEnv) => number
interface Slot { ok?: Compiled; refused?: string }
const memo = new WeakMap<object, Map<string, Slot>>()

function slotOf(node: object, id: string): Slot {
  let slots = memo.get(node)
  if (!slots) memo.set(node, slots = new Map())
  let found = slots.get(id)
  if (!found) slots.set(id, found = {})
  return found
}

const HEX = /^#[0-9a-fA-F]{3,8}$/

/** Bare words the formula grammar reads as constants rather than literal text. */
const FORMULA_WORDS: ReadonlySet<string> = new Set(['pi', 'e'])

/**
 * Literal text, never compiled: hex colours, `kind:value` colour choices
 * (`schemeClr:accent1`, `hslClr:…`) and bare words (`visible`, `hidden`,
 * `true`). Underscored names (`ppt_x`, `no_such_name`) are formula
 * identifiers, so a typo still refuses and diagnoses.
 */
function isLiteral(text: string): boolean {
  if (HEX.test(text)) return true
  if (/^[a-z][a-z0-9]*:/i.test(text)) return true
  return /^[a-z][a-z-]*$/i.test(text) && !FORMULA_WORDS.has(text.toLowerCase())
}

class Once {
  private fired = false
  next(): boolean { if (this.fired) return false; this.fired = true; return true }
}

/** Evaluates a formula-ish string, dropping the behaviour on refusal/NaN contents. */
function evalSlot(
  node: object, id: string, src: string, env: FormulaEnv,
  sink: DiagnosticSink, effect: string, once: Once,
): number | undefined {
  const s = slotOf(node, id)
  if (s.refused !== undefined) {
    if (once.next()) sink.add({ code: 'behaviour-invalid-formula', effect, message: `Formula ${JSON.stringify(src)} refused (${s.refused}); behaviour dropped` })
    return undefined
  }
  if (!s.ok) {
    const c = compileFormula(src)
    if (c.ok) {
      s.ok = c.fn
    } else {
      s.refused = c.reason
      if (once.next()) sink.add({ code: 'behaviour-invalid-formula', effect, message: `Formula ${JSON.stringify(src)} refused (${c.reason}); behaviour dropped` })
      return undefined
    }
  }
  const v = s.ok(env)
  if (!Number.isFinite(v)) {
    if (once.next()) sink.add({ code: 'behaviour-invalid-formula', effect, message: 'Formula evaluated to a non-finite value; behaviour dropped' })
    return undefined
  }
  return v
}

/**
 * A raw value: literal text, a plain number, or a formula (`#ppt_x-1`,
 * `$+0.2`). A formula that refuses or evaluates non-finite returns undefined
 * (already diagnosed) so the caller drops the behaviour; its text never leaks
 * through as a value.
 */
function readValue(
  node: object, id: string, src: string | number | undefined, env: FormulaEnv,
  sink: DiagnosticSink, effect: string, once: Once,
): number | string | undefined {
  if (src === undefined) return undefined
  if (typeof src === 'number') return Number.isFinite(src) ? src : undefined
  const text = src.trim()
  if (text === '') return undefined
  if (isLiteral(text)) return text
  const n = Number(text)
  if (Number.isFinite(n)) return n
  return evalSlot(node, id, text, env, sink, effect, once)
}

/** Segment `i` of a `;`-separated per-attribute list (segment 0 when there are fewer). */
function segment(src: string | undefined, i: number): string | undefined {
  if (src === undefined) return undefined
  const parts = src.split(';')
  return parts[i] ?? parts[0]
}

// ---------------------------------------------------------------------------
// Per-behaviour reading
// ---------------------------------------------------------------------------

type Channel = 'fill' | 'line' | 'text'
const CHANNEL_FOR = (name: string): Channel | undefined =>
  name === 'fillcolor' ? 'fill' : name === 'stroke.color' ? 'line' : name === 'style.color' ? 'text' : undefined

interface Reads {
  /** Track name → candidate value; ppt_* fractions, angle degrees, scale factors, motion px, strings, matrices. */
  edits: Map<string, number | string | ColorMatrix>
  colors: Map<Channel, ColorMatrix>
  /** Set when the read is relative by definition (`by` animations): composes as `sum`. */
  additive?: 'sum'
  mask?: MaskState
  opacityCap?: number
  failed?: true
}

function evalBehaviour(
  node: BehaviourNode,
  env: FormulaEnv,
  progress: number,
  slide: Slide,
  geo: BaseGeometry,
  sink: DiagnosticSink,
): Reads {
  const effect = `behaviour ${node.ctn.id}`
  const once = new Once()
  const reads: Reads = { edits: new Map(), colors: new Map() }
  const edit = (name: string, v: number | string | ColorMatrix): void => void reads.edits.set(name, v)
  const fail = (): Reads => ({ ...reads, failed: true })

  switch (node.kind) {
    case 'set': {
      const set = node as SetBehaviour
      if (set.to === undefined) return fail()
      const values: (number | string)[] = []
      for (let i = 0; i < set.attrNames.length; i++) {
        const src = typeof set.to === 'number' ? set.to : segment(String(set.to), i)
        const v = readValue(node, `set:${i}`, src, env, sink, effect, once)
        if (v === undefined) {
          if (once.next()) sink.add({ code: 'behaviour-invalid-formula', effect, message: `set to ${JSON.stringify(src)} is unreadable; behaviour dropped` })
          return fail()
        }
        values.push(typeof v === 'number' ? v : String(v))
      }
      set.attrNames.forEach((n, i) => edit(n, values[i] ?? values[0]))
      return reads
    }
    case 'anim': {
      // Each attribute reads its own segment of `;` lists (S1-5: split before compiling).
      for (let i = 0; i < node.attrNames.length; i++) {
        const name = String(node.attrNames[i])
        const read = node.keys.length > 0
          ? { value: tavValue(node, i, env, progress, sink, effect, once) }
          : endpointValue(node, i, name, env, progress, geo, slide, sink, effect, once)
        const value = read.value
        if (value === undefined || (typeof value === 'number' && !Number.isFinite(value))) return fail()
        if (typeof value !== 'number' && !/^(visible|hidden)$/.test(value)) {
          unsupported(node, name, sink)
          return fail()
        }
        if (read.sum) reads.additive = 'sum'
        edit(name, value)
      }
      return reads
    }
    case 'animRot': {
      const v = rotValue(node, progress)
      if (v === undefined || !Number.isFinite(v.angle)) return fail()
      if (v.relative) reads.additive = 'sum'
      edit('angle', v.angle)
      return reads
    }
    case 'animScale': {
      const factor = scaleValue(node, progress)
      edit('scale:x', factor.x)
      edit('scale:y', factor.y)
      return reads
    }
    case 'animMotion': {
      const offsets = motionOffsets(node, progress, slide, geo, sink, effect, once)
      if (offsets === undefined) return fail()
      edit('motion:dx', offsets.dx)
      edit('motion:dy', offsets.dy)
      return reads
    }
    case 'animEffect': {
      const state = effectState(node, progress, sink)
      if (state.mask) reads.mask = state.mask
      else if (state.opacity !== undefined) reads.opacityCap = state.opacity
      return reads
    }
    case 'animClr': {
      const matrix = clrMatrix(node, progress, sink, effect, once)
      if (matrix === undefined) return fail()
      for (const name of node.attrNames) {
        const channel = CHANNEL_FOR(name)
        if (channel) reads.colors.set(channel, matrix)
      }
      return reads
    }
    case 'audio': case 'video': case 'cmd':
      // Media has no render edits. The session's media player follows shape media, effect sounds (`p:sndTgt`,
      // S2-10) and commands on the timeline, plays them and diagnoses any refusal itself.
      return fail()
    default:
      unsupported(node, (node as BehaviourNode).kind, sink)
      return fail()
  }
}

// -- anim (tav keys) ---------------------------------------------------------

function tavValue(
  node: AnimBehaviour, attr: number, env: FormulaEnv, progress: number,
  sink: DiagnosticSink, effect: string, once: Once,
): number | string | undefined {
  const keys = node.keys.map(x => ({
    tm: x.tm === 'indefinite' ? 1 : Math.min(1, Math.max(0, Number(x.tm))),
    value: typeof x.value === 'string' ? segment(x.value, attr) : x.value,
    ...(x.fmla ? { fmla: segment(x.fmla, attr) } : {}),
  }))
  if (keys.length === 0) return undefined
  keys.sort((a, b) => a.tm - b.tm)
  // Each key's value is a snapshot at its own template time; a formula or
  // value that cannot be read fails the behaviour rather than degrading.
  const snap = (i: number): number | string | undefined => {
    const key = keys[i]
    const at = { ...env, $: key.tm }
    if (key.fmla !== undefined) return evalSlot(node, `key:${i}:${attr}`, key.fmla.trim(), at, sink, effect, once)
    return readValue(node, `key:${i}:${attr}`, key.value, at, sink, effect, once)
  }
  if (node.calcmode === 'discrete') {
    let pick = 0
    for (let i = 0; i < keys.length; i++) {
      if (keys[i].tm <= progress) pick = i
      else break
    }
    return snap(pick)
  }
  if (keys.length === 1) return snap(0)
  for (let i = 0; i < keys.length - 1; i++) {
    const a = keys[i].tm, b = keys[i + 1].tm
    if (progress >= a && progress <= b) {
      const va = snap(i), vb = snap(i + 1)
      const na = typeof va === 'number' ? va : NaN
      const nb = typeof vb === 'number' ? vb : NaN
      if (Number.isFinite(na) && Number.isFinite(nb)) {
        const t = b - a <= 0 ? 1 : (progress - a) / (b - a)
        return na + (nb - na) * t
      }
      return undefined
    }
  }
  return snap(keys.length - 1)
}

/**
 * `p:anim` without a `p:tavLst` (SMIL from/to/by): from→to interpolates;
 * from+by runs from→from+by; by alone is relative (composes as `sum` onto
 * the underlying value); to alone starts from the underlying value. Values
 * that cannot interpolate (text) step at the midpoint, or hold `to` alone.
 */
function endpointValue(
  node: AnimBehaviour, attr: number, name: string, env: FormulaEnv, progress: number,
  geo: BaseGeometry, slide: Slide, sink: DiagnosticSink, effect: string, once: Once,
): { value: number | string | undefined; sum?: true } {
  const read = (which: 'from' | 'to' | 'by'): number | string | undefined | null => {
    const src = segment(node[which], attr)
    if (src === undefined) return null
    return readValue(node, `${which}:${attr}`, src, env, sink, effect, once)
  }
  const from = read('from'), to = read('to'), by = read('by')
  if (from === undefined || to === undefined || by === undefined) return { value: undefined }
  const discrete = node.calcmode === 'discrete'
  const mix = (a: number | string, b: number | string): number | string =>
    typeof a === 'number' && typeof b === 'number' && !discrete ? a + (b - a) * progress : progress < 0.5 ? a : b
  if (from !== null && to !== null) return { value: mix(from, to) }
  if (by !== null && typeof by !== 'number') return { value: undefined }
  if (from !== null && by !== null) return typeof from === 'number' ? { value: mix(from, from + by) } : { value: undefined }
  if (by !== null) return { value: discrete ? (progress < 0.5 ? 0 : by) : by * progress, sum: true }
  if (to !== null) {
    const under = typeof to === 'number' ? trackBase(name, geo, slide) : NaN
    return { value: Number.isFinite(under) ? mix(under, to) : to }
  }
  return { value: from ?? undefined }
}

/** from/to are absolute degrees; `by` is relative (composes as `sum` onto the underlying rotation). */
function rotValue(node: AnimRotBehaviour, p: number): { angle: number; relative: boolean } | undefined {
  if (node.to !== undefined || node.from !== undefined) {
    const a = node.from ?? 0
    const b = node.to ?? a
    return { angle: a + (b - a) * p, relative: false }
  }
  if (node.by === undefined) return undefined
  return { angle: node.by * p, relative: true }
}

function scaleValue(node: AnimScaleBehaviour, p: number): { x: number; y: number } {
  const from = node.from ?? { x: 1, y: 1 }
  const to = node.to ?? node.by ?? { x: 1, y: 1 }
  return { x: from.x + (to.x - from.x) * p, y: from.y + (to.y - from.y) * p }
}

function motionOffsets(
  node: AnimMotionBehaviour, p: number, slide: Slide, geo: BaseGeometry,
  sink: DiagnosticSink, effect: string, once: Once,
): { dx: number; dy: number } | undefined {
  void once; void sink; void effect
  if (node.path) {
    const at = pointAt(node.path, p)
    let dx = at.x * slide.w, dy = at.y * slide.h
    if (node.origin === 'layout') {
      dx -= geo.x
      dy -= geo.y
    }
    const deg = node.rAngDeg ?? 0
    if (deg !== 0) {
      const a = (deg * Math.PI) / 180
      const c = Math.cos(a), s = Math.sin(a)
      const [x, y] = [dx, dy]
      dx = x * c - y * s
      dy = x * s + y * c
    }
    return { dx, dy }
  }
  const a = node.from ?? { x: 0, y: 0 }
  const b = node.to ?? node.by ?? a
  return { dx: (a.x + (b.x - a.x) * p) * slide.w, dy: (a.y + (b.y - a.y) * p) * slide.h }
}

function effectState(node: AnimEffectBehaviour, progress: number, sink: DiagnosticSink): { mask?: MaskState; opacity?: number } {
  const stated = node.progress === undefined ? undefined : Number(node.progress)
  const p = stated !== undefined && Number.isFinite(stated) ? Math.min(1, Math.max(0, stated)) : progress
  const raw = node.filter?.trim().toLowerCase() ?? ''
  const matched = /^([a-z]+)\s*\(([^)]*)\)$/.exec(raw)
  const name = matched?.[1] ?? raw
  const arg = (matched?.[2] ?? '').trim()
  const ramp = (diag = false): { opacity: number } => {
    if (diag) unsupported(node, node.filter ?? 'animEffect', sink, 'anim-effect-unknown-filter')
    return node.transition === 'out' ? { opacity: 1 - p } : { opacity: p }
  }
  if (!name) return ramp()
  if (name === 'fade') return ramp()
  if (name === 'wipe') {
    // `dir` is the motion direction; the reveal comes from the opposite edge.
    const from = ({ left: 'r', right: 'l', up: 'b', down: 't' })[arg]
    if (!from) return ramp(true)
    return { mask: { kind: 'wipe', progress: p, params: { from } } }
  }
  if (name === 'slide') {
    const from = ({ fromleft: 'l', fromright: 'r', fromtop: 't', frombottom: 'b' })[arg.toLowerCase()]
    return { mask: { kind: 'wipe', progress: p, params: { from: from ?? 'l' } } }
  }
  if (name === 'wheel') {
    // The mask library bounds spokes at 64; a count outside 1..64 refuses the mask (diagnosed, rendered as a fade).
    const spokes = arg === '' ? 1 : Number(arg)
    if (!Number.isInteger(spokes) || spokes < 1 || spokes > 64) return ramp(true)
    return { mask: { kind: 'wheel', progress: p, params: { spokes } } }
  }
  if (name === 'blinds') return { mask: { kind: 'blinds', progress: p, params: { orient: /vert/i.test(arg) ? 'vert' : 'horz' } } }
  if (name === 'barn') return { mask: { kind: 'barn', progress: p, params: { orient: /vert/i.test(arg) ? 'vert' : 'horz', dir: /^in/i.test(arg) ? 'in' : 'out' } } }
  if (name === 'randombar') return { mask: { kind: 'randomBar', progress: p, params: { orient: /vert/i.test(arg) ? 'vert' : 'horz' } } }
  if (name === 'strips') {
    // Like wipe, the argument names the motion direction; the reveal starts at the opposite corner.
    const from = ({ downleft: 'ru', downright: 'lu', upleft: 'rd', upright: 'ld' })[arg.toLowerCase()]
    if (!from) return ramp(true)
    return { mask: { kind: 'strips', progress: p, params: { from } } }
  }
  if (name === 'box' || name === 'circle' || name === 'diamond' || name === 'plus')
    return { mask: { kind: name, progress: p, params: { dir: /^in/i.test(arg) ? 'in' : 'out' } } }
  if (name === 'checkerboard') return { mask: { kind: 'checkerboard', progress: p, params: { dir: arg === 'down' ? 'down' : 'across' } } }
  if (name === 'wedge' || name === 'dissolve') return { mask: { kind: name, progress: p, params: {} } }
  return ramp(true)
}

function clrMatrix(node: AnimClrBehaviour, p: number, sink: DiagnosticSink, effect: string, once: Once): ColorMatrix | undefined {
  if (node.by !== undefined) {
    const hue = hueDegrees(node.by)
    if (hue !== undefined) return hueMatrix(node.dir === 'ccw' ? -hue : hue)
    const d = readRgb(node.by)
    if (!d) {
      if (once.next()) sink.add({ code: 'anim-clr-unreadable', effect, message: `animClr by ${JSON.stringify(node.by)} is unreadable; the track is dropped` })
      return undefined
    }
    const m = [...RGB_IDENTITY] as number[]
    m[0] = 1 + d.r * p
    if (d.g !== 0) m[6] = 1 + d.g * p
    if (d.b !== 0) m[12] = 1 + d.b * p
    return m
  }
  if (node.from !== undefined || node.to !== undefined) {
    const from = readRgb(node.from ?? '#000000')
    const to = readRgb(node.to ?? '#000000')
    if (!from || !to) {
      if (once.next()) sink.add({ code: 'anim-clr-unreadable', effect, message: `animClr ${JSON.stringify([node.from, node.to])} needs the theme; the track is dropped` })
      return undefined
    }
    const at: Rgb = { r: from.r + (to.r - from.r) * p, g: from.g + (to.g - from.g) * p, b: from.b + (to.b - from.b) * p }
    return matrixScale(at, from)
  }
  return RGB_IDENTITY
}

// ---------------------------------------------------------------------------
// Track accumulation (cBhvr additive/accumulate, document order)
// ---------------------------------------------------------------------------

interface Track { own: number; base: number; set: boolean }

function compose(step: Track, v: number, additive: BehaviourNode['additive'], iteration: number, span: number, accumulate: BehaviourNode['accumulate']): Track {
  if (!Number.isFinite(v)) return step
  let next: number
  if (!step.set) {
    next = additive === 'sum' ? step.base + v : additive === 'mult' ? step.base * v : 0 + v
  } else if (additive === 'sum') next = step.own + v
  else if (additive === 'mult') next = step.own * v
  else next = v
  if (accumulate === 'always' && Number.isFinite(span)) next += iteration * span
  return { own: next, base: step.base, set: true }
}

const warned = new WeakMap<BehaviourNode, Set<string>>()
function unsupported(node: BehaviourNode, name: string, sink: DiagnosticSink, code = 'behaviour-unsupported'): void {
  let set = warned.get(node)
  if (!set) warned.set(node, set = new Set())
  if (!set.has(name)) {
    set.add(name)
    const label = code === 'behaviour-unsupported' ? `Attr ${name} has no render mapping; dropped`
      : `Effect filter ${JSON.stringify(name)} is not in the mask catalogue; rendering as a fade`
    sink.add({ code, effect: `behaviour ${node.ctn.id}`, message: label })
  }
}

/** The target's centre as a fraction of the slide: what ppt_x / ppt_y read and write. */
const centreX = (geo: BaseGeometry, slide: Slide) => (geo.x + geo.w / 2) / slide.w
const centreY = (geo: BaseGeometry, slide: Slide) => (geo.y + geo.h / 2) / slide.h

function trackBase(name: string, geo: BaseGeometry, slide: Slide): number {
  switch (name) {
    case 'ppt_x': return centreX(geo, slide)
    case 'ppt_y': return centreY(geo, slide)
    case 'ppt_w': return geo.w / slide.w
    case 'ppt_h': return geo.h / slide.h
    case 'angle': return geo.rot
    case 'scale:x': case 'scale:y': return 1
    case 'style.opacity': return 1
    default: return NaN
  }
}

function envOf(progress: number, geo: BaseGeometry, slide: Slide, rng: Rng): FormulaEnv {
  return {
    $: progress,
    ppt_x: centreX(geo, slide),
    ppt_y: centreY(geo, slide),
    ppt_w: geo.w / slide.w,
    ppt_h: geo.h / slide.h,
    ppt_r: geo.rot,
    current: (name: string) => {
      switch (name) {
        case 'ppt_x': case 'x': return centreX(geo, slide)
        case 'ppt_y': case 'y': return centreY(geo, slide)
        case 'ppt_w': case 'w': return geo.w / slide.w
        case 'ppt_h': case 'h': return geo.h / slide.h
        case 'ppt_r': case 'r': return geo.rot
        default: return NaN
      }
    },
    rand: rng,
  }
}

/** The snapshot's behaviours in document order (each list is already ordered; merged by entry order). */
export function docOrder(snapshot: SchedulerSnapshot): readonly { node: BehaviourNode; progress: number; iteration: number; order: number }[] {
  const active = snapshot.active.map((e, i) => ({ e, rank: e.order ?? i }))
  const frozen = snapshot.frozen.map((e, i) => ({ e, rank: e.order ?? i }))
  return [...active, ...frozen].sort((a, b) => a.rank - b.rank).map(x => ({
    node: x.e.node, progress: x.e.progress, iteration: x.e.iteration, order: x.rank,
  }))
}

interface Accumulated {
  visible: boolean
  opacity: number
  tracks: Map<string, Track>
  motion?: { dx: number; dy: number }
  mask?: MaskState
  color: { fill?: ColorMatrix; line?: ColorMatrix; text?: ColorMatrix }
  opacityCap: number
}

/**
 * Resolves a scheduler snapshot into per-target states (the brief signature,
 * plus an additive trailing `slide` in px — the ppt_x/y/w/h fraction→px
 * conversion needs the slide size; with the unit default {w:1,h:1}, 1.0 = one
 * slide). `rng` feeds formulas that call `rand`.
 */
export function renderStateAt(
  snapshot: SchedulerSnapshot,
  base: ReadonlyMap<TargetKey, BaseGeometry>,
  rng: Rng,
  sink: DiagnosticSink,
  slide: Slide = { w: 1, h: 1 },
  fan?: ReadonlyMap<BehaviourNode, IterateFan>,
): Map<TargetKey, TargetState> {
  const out = new Map<TargetKey, Accumulated>()
  // Split text items (`sp:<id>/…` from a word or letter fan) belong to their shape: whole-shape behaviours that are
  // not fanned (a later exit or emphasis) are also accumulated apart, as each item's parent state (PR #19 review 6).
  const itemsOf = new Map<TargetKey, Set<TargetKey>>()
  for (const f of fan?.values() ?? []) {
    for (const { key } of f.items) {
      const cut = key.indexOf('/')
      if (cut < 0) continue
      const shape = key.slice(0, cut)
      const set = itemsOf.get(shape) ?? new Set<TargetKey>()
      set.add(key)
      itemsOf.set(shape, set)
    }
  }
  const parents = new Map<TargetKey, Accumulated>()
  const geoOf = (key: TargetKey): BaseGeometry | undefined => base.get(key) ?? (key.includes('/') ? base.get(key.slice(0, key.indexOf('/'))) : undefined)
  const apply = (node: BehaviourNode, progress: number, iteration: number, key: TargetKey, into = out): void => {
    const geo = geoOf(key) ?? { x: 0, y: 0, w: 0, h: 0, rot: 0 }
    const env = envOf(progress, geo, slide, rng)
    const reads = evalBehaviour(node, env, progress, slide, geo, sink)
    if (reads.failed) return
    const additive = reads.additive ?? node.additive
    let acc = into.get(key) ?? { visible: true, opacity: 1, tracks: new Map(), color: {}, opacityCap: 1 } as Accumulated
    into.set(key, acc)
    for (const [name, value] of reads.edits) {
      if (name === 'motion:dx' || name === 'motion:dy') {
        const px = typeof value === 'number' ? value : Number(value)
        if (!Number.isFinite(px)) continue
        const axis = name === 'motion:dx' ? 'dx' as const : 'dy' as const
        const prior = axis === 'dx' ? acc.motion?.dx ?? 0 : acc.motion?.dy ?? 0
        const step: Track = { own: prior, base: 0, set: acc.motion !== undefined }
        const span = node.accumulate === 'always' ? motionSpan(node, axis, geo, slide, rng) : 0
        acc.motion = axis === 'dx'
          ? { dx: compose(step, px, additive, iteration, span, node.accumulate).own, dy: acc.motion?.dy ?? 0 }
          : { dx: acc.motion?.dx ?? 0, dy: compose(step, px, additive, iteration, span, node.accumulate).own }
        continue
      }
      if (name === 'style.visibility') {
        acc.visible = value === 'visible'
        continue
      }
      if (typeof value === 'number' && name === 'style.opacity') {
        acc.opacity = compose({ own: acc.opacity, base: 1, set: true }, Math.min(1, Math.max(0, value)), additive, iteration, 0, node.accumulate).own
        continue
      }
      if (!/^(ppt_[xywh]|angle|scale:[xy])$/.test(name)) {
        unsupported(node, name, sink)
        continue
      }
      if (typeof value !== 'number') {
        unsupported(node, name, sink)
        continue
      }
      const fresh: Track = { own: trackBase(name, geo, slide), base: trackBase(name, geo, slide), set: false }
      const step = acc.tracks.get(name) ?? fresh
      let span = 0
      if (node.accumulate === 'always') {
        const v0 = readTrackValue(node, 0, geo, slide, name, rng)
        const v1 = readTrackValue(node, 1, geo, slide, name, rng)
        if (v0 !== undefined && v1 !== undefined) span = v1 - v0
      }
      acc.tracks.set(name, compose(step, value, additive, iteration, span, node.accumulate))
    }
    if (reads.mask) acc.mask = reads.mask
    if (reads.opacityCap !== undefined) acc.opacityCap = Math.min(acc.opacityCap, reads.opacityCap)
    for (const [channel, matrix] of reads.colors) {
      const prior = acc.color[channel]
      acc.color[channel] = prior ? composeColor(prior, matrix) : matrix
    }
  }
  // Iterated entrance items that have not begun under any behaviour stay hidden (S2-8).
  const begun = new Set<TargetKey>()
  const waiting = new Set<TargetKey>()
  for (const { node, progress, iteration } of docOrder(snapshot)) {
    const f = fan?.get(node)
    if (!f) {
      const key = targetKey(node.target)
      apply(node, progress, iteration, key)
      if (itemsOf.has(key)) apply(node, progress, iteration, key, parents)
      continue
    }
    const elapsed = progress * f.totalMs
    for (const { key, delayMs } of f.items) {
      const raw = f.spanMs > 0 ? (elapsed - delayMs) / f.spanMs : elapsed >= delayMs ? 1 : -1
      if (raw < 0) { if (f.entrance) waiting.add(key); continue }
      begun.add(key)
      apply(node, easeProgress(Math.min(1, raw), f.accel, f.decel, f.autoRev), iteration, key)
    }
  }
  for (const key of waiting) {
    if (begun.has(key)) continue
    const acc = out.get(key) ?? { visible: false, opacity: 1, tracks: new Map(), color: {}, opacityCap: 1 } as Accumulated
    acc.visible = false
    out.set(key, acc)
  }
  const states = new Map<TargetKey, TargetState>()
  for (const [key, acc] of out) states.set(key, finalize(acc, geoOf(key), slide))
  // Each split item composes with its shape's whole-target state as its parent: visibility, opacity and transform.
  for (const [shape, acc] of parents) {
    const parent = finalize(acc, geoOf(shape), slide)
    for (const key of itemsOf.get(shape) ?? []) {
      const own = states.get(key)
      // An item with no state of its own keeps the stage's default (hidden until its entrance).
      if (own) states.set(key, withParent(own, parent))
    }
  }
  return states
}

/** `child` drawn inside `parent`: the parent's full mapping (motion after transform) applied after the child's. */
function withParent(child: TargetState, parent: TargetState): TargetState {
  const full = (s: TargetState): Affine => s.motion ? mulAffine(s.transform, [1, 0, 0, 1, s.motion.dx, s.motion.dy]) : s.transform
  const { motion: _motion, ...rest } = child
  void _motion
  return {
    ...rest,
    visible: child.visible && parent.visible,
    opacity: child.opacity * parent.opacity,
    transform: mulAffine(full(child), full(parent)),
  }
}

/** The track's value at a posed progress, evaluated on a THROWAWAY sink (for accumulate spans). */
function readTrackValue(node: BehaviourNode, progress: number, geo: BaseGeometry, slide: Slide, name: string, rng: Rng): number | undefined {
  const throwaway = { add: () => {}, list: () => [] } as unknown as DiagnosticSink
  const env = envOf(progress, geo, slide, rng)
  const reads = evalBehaviour(node, env, progress, slide, geo, throwaway)
  if (reads.failed) return undefined
  const v = reads.edits.get(name)
  return typeof v === 'number' && Number.isFinite(v) ? v : undefined
}

function motionSpan(node: BehaviourNode, axis: 'dx' | 'dy', geo: BaseGeometry, slide: Slide, rng: Rng): number {
  const v0 = readTrackValue(node, 0, geo, slide, `motion:${axis}`, rng)
  const v1 = readTrackValue(node, 1, geo, slide, `motion:${axis}`, rng)
  return v0 !== undefined && v1 !== undefined ? v1 - v0 : 0
}

function finalize(acc: Accumulated, geo: BaseGeometry | undefined, slide: Slide): TargetState {
  const g = geo ?? { x: 0, y: 0, w: 0, h: 0, rot: 0 }
  const movedX = acc.tracks.get('ppt_x')?.set === true
  const movedY = acc.tracks.get('ppt_y')?.set === true
  const scaledW = acc.tracks.get('ppt_w')?.set === true
  const scaledH = acc.tracks.get('ppt_h')?.set === true
  const factorX = acc.tracks.get('scale:x')?.set === true ? Math.max(1e-6, acc.tracks.get('scale:x')!.own) : 1
  const factorY = acc.tracks.get('scale:y')?.set === true ? Math.max(1e-6, acc.tracks.get('scale:y')!.own) : 1
  const x = acc.tracks.get('ppt_x')?.own ?? centreX(g, slide)
  const y = acc.tracks.get('ppt_y')?.own ?? centreY(g, slide)
  const w = acc.tracks.get('ppt_w')?.own ?? g.w / slide.w
  const h = acc.tracks.get('ppt_h')?.own ?? g.h / slide.h
  const sx = (scaledW ? Math.max(1e-6, (w * slide.w) / Math.max(1e-6, g.w)) : 1) * factorX
  const sy = (scaledH ? Math.max(1e-6, (h * slide.h) / Math.max(1e-6, g.h)) : 1) * factorY
  const angle = acc.tracks.get('angle')?.own ?? g.rot
  const cx = g.x + g.w / 2, cy = g.y + g.h / 2
  const dx = movedX ? x * slide.w - cx : 0
  const dy = movedY ? y * slide.h - cy : 0
  const scale: Affine = [sx, 0, 0, sy, cx * (1 - sx), cy * (1 - sy)]
  const a = (angle * Math.PI) / 180
  const c = Math.cos(a), s = Math.sin(a)
  const rotAbout: Affine = [c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy]
  const transform = mulAffine(scale, rotAbout)
  // animMotion translates on top of the ppt_x/ppt_y position: the offsets add.
  const motion = acc.motion
    ? { dx: acc.motion.dx + dx, dy: acc.motion.dy + dy }
    : movedX || movedY ? { dx, dy } : undefined
  const colour = acc.color.fill || acc.color.line || acc.color.text ? acc.color : undefined
  return {
    visible: acc.visible,
    opacity: Math.min(1, Math.max(0, acc.opacity * acc.opacityCap)),
    transform,
    ...(acc.mask ? { mask: acc.mask } : {}),
    ...(colour ? { color: colour } : {}),
    ...(motion ? { motion } : {}),
  }
}

// ---------------------------------------------------------------------------
// Text by word and by letter: `p:iterate type="wd"/"lt"` (S2-8; spec 6.5)
// ---------------------------------------------------------------------------

/**
 * How one iterated behaviour fans out over its items. `items` are evaluated
 * under their own keys, each `delayMs` after the behaviour begins, for
 * `spanMs` each (one simple duration; `autoRev` doubles it), eased per item.
 * `totalMs` is the behaviour's extended duration on the working copy.
 */
export interface IterateFan {
  items: ReadonlyArray<{ key: TargetKey; delayMs: number }>
  spanMs: number
  totalMs: number
  accel: number
  decel: number
  autoRev: boolean
  /** Under an entrance effect: an item that has not begun is hidden. */
  entrance: boolean
}

export interface IteratedTiming {
  /** The working copy to schedule (the source tree itself when nothing fans out). */
  tree: TimingTree
  /** Pass to `createScheduler` so `SchedulerEntry.order` stays the source pre-order index. */
  orderOf: ReadonlyMap<TimeNode, number>
  /** Pass to `renderStateAt`. */
  fan: ReadonlyMap<BehaviourNode, IterateFan>
  /** Iterated target key → its item keys, for callers that track per-target visibility. */
  items: ReadonlyMap<TargetKey, readonly TargetKey[]>
}

export interface IterateOptions {
  sink?: DiagnosticSink
  slideIndex?: number
  /** When `tree` is itself a working copy: its nodes' source pre-order indices. */
  orderOf?: ReadonlyMap<TimeNode, number>
  /** Reserves `bytes` before the working copy is built; `false` refuses (the tree stays as it is, diagnosed). */
  charge?: (bytes: number, tag: string) => boolean
}

/** Item evaluations per slide and frame (fanned behaviours × items); past it an effect plays whole, diagnosed. */
export const MAX_FAN_ITEMS = 4096
const FAN_NODE_BYTES = 256
const FAN_ITEM_BYTES = 48
const MAX_TREE_DEPTH = 64

/**
 * Start offsets (ms) of `n` items: `tmAbs` is the interval itself, `tmPct`
 * a fraction of `durMs`; `backwards` starts the last item first. A missing,
 * negative or non-finite interval is 0 (all items at once).
 */
export function iterateDelays(iterate: Iterate, n: number, durMs: number): number[] {
  const raw = iterate.tmAbsMs !== undefined ? iterate.tmAbsMs : iterate.tmPct !== undefined ? iterate.tmPct * durMs : 0
  const step = Number.isFinite(raw) && raw > 0 ? raw : 0
  return Array.from({ length: Math.max(0, n) }, (_, i) => (iterate.backwards ? n - 1 - i : i) * step)
}

const finiteMs = (v: number | 'indefinite' | undefined): number | undefined => typeof v === 'number' && Number.isFinite(v) && v >= 0 ? v : undefined

/**
 * Fans every behaviour under a `wd`/`lt` iterate out over the boxes of its
 * target. `splits` maps an iterated target key (a shape, or a paragraph
 * range) to its box keys in reading order, from `splitTextBoxes`.
 *
 * The result is a working copy (the source tree is never mutated): each
 * fanned behaviour lasts `D + max(delay)` (`D` its simple duration, doubled
 * by `autoRev`), with its accel, decel and autoRev moved into the fan so they
 * apply per item; an iterated container with a finite `dur` is extended by
 * its largest delay. So the scheduler sees the true length and successors
 * (`afterEffect`, `onEnd`) wait for the last item. A whole-shape target also
 * keeps its own key as an item for the body layer: at delay 0, or under an
 * exit effect at the last item's delay.
 *
 * `tmPct` is a fraction of the iterated container's duration: its finite
 * `dur`, else its longest behaviour's. Targets without boxes and behaviours
 * without a finite duration play whole; everything is diagnosed.
 */
export function expandIterations(tree: TimingTree, splits: ReadonlyMap<TargetKey, readonly TargetKey[]>, options: IterateOptions = {}): IteratedTiming {
  const { sink, slideIndex } = options
  const diag = (code: string, effect: string, message: string): void => { sink?.add({ code, slideIndex, effect, message }) }
  const srcOrder = new Map<TimeNode, number>()
  interface Plan { node: BehaviourNode; owner?: ContainerNode; effect: string; keys: readonly TargetKey[]; whole?: TargetKey; delays: number[]; entrance: boolean; exit: boolean }
  const plans: Plan[] = []
  const containerDur = new Map<ContainerNode, number | undefined>()
  const longest = (n: TimeNode, depth: number): number | undefined => {
    if (depth > MAX_TREE_DEPTH) return undefined
    if ('children' in n) {
      let best: number | undefined
      for (const c of n.children) { const d = longest(c, depth + 1); if (d !== undefined) best = Math.max(best ?? 0, d) }
      return best
    }
    return finiteMs(n.ctn.durMs)
  }
  let fanned = 0
  const walk = (n: TimeNode, iterated: { owner: ContainerNode; iterate: Iterate } | undefined, cls: string | undefined, depth: number): void => {
    if (srcOrder.has(n) || depth > MAX_TREE_DEPTH) return
    srcOrder.set(n, srcOrder.size)
    const here = n.ctn.presetClass ?? cls
    const own = n.ctn.iterate && n.ctn.iterate.type !== 'el' ? n.ctn.iterate : undefined
    if ('children' in n) {
      const next = own ? { owner: n, iterate: own } : iterated
      for (const c of n.children) walk(c, next, here, depth + 1)
      return
    }
    // A behaviour's own iterate wins over its container's; it then staggers by its own duration.
    const it = own ?? iterated?.iterate
    if (!it) return
    const owner = own ? undefined : iterated?.owner
    const effect = `iterate:${(owner ?? n).ctn.id}`
    if (!('spid' in n.target)) return
    const key = targetKey(n.target)
    const keys = splits.get(key)
    const unit = it.type === 'wd' ? 'word' : 'letter'
    if (!keys || keys.length === 0) {
      diag('text-build-unsplit', effect, `Text by ${unit} on ${key} has no word or letter boxes; the effect plays on the whole target`)
      return
    }
    const dur = finiteMs(n.ctn.durMs)
    if (dur === undefined) {
      diag('text-build-unsplit', effect, `Text by ${unit} on ${key}: a behaviour without a finite duration cannot be staggered; it plays on the whole target`)
      return
    }
    const whole = n.target.text ? undefined : key
    const cost = keys.length + (whole ? 1 : 0)
    if (fanned + cost > MAX_FAN_ITEMS) {
      diag('text-build-refused', effect, `Text by ${unit} on ${key} would exceed ${MAX_FAN_ITEMS} item evaluations per slide; the effect plays on the whole target`)
      return
    }
    if (owner && !containerDur.has(owner)) containerDur.set(owner, finiteMs(owner.ctn.durMs) ?? longest(owner, depth))
    const base = owner ? containerDur.get(owner) : dur
    if (it.tmAbsMs === undefined && it.tmPct !== undefined && base === undefined) {
      diag('text-build-unsplit', effect, `Text by ${unit} on ${key}: tmPct needs a finite effect duration; the items start together`)
    }
    fanned += cost
    plans.push({ node: n, owner, effect, keys, whole, delays: iterateDelays(it, keys.length, base ?? 0), entrance: here === 'entr', exit: here === 'exit' })
  }
  walk(tree.root, undefined, undefined, 0)
  const unchanged = (): IteratedTiming => ({ tree, orderOf: options.orderOf ?? srcOrder, fan: new Map(), items: new Map() })
  if (plans.length === 0) return unchanged()

  const bytes = srcOrder.size * FAN_NODE_BYTES + fanned * FAN_ITEM_BYTES
  if (options.charge && !options.charge(bytes, 'iterate')) {
    for (const p of plans) diag('text-build-refused', p.effect, `Text by word or letter: the slideshow workspace refused ${bytes} B for the fan-out; the effect plays on the whole target`)
    return unchanged()
  }

  // The working copy: every node cloned once (shared node objects stay shared).
  const copies = new Map<TimeNode, TimeNode>()
  const orderOf = new Map<TimeNode, number>()
  const copy = (n: TimeNode, depth: number): TimeNode => {
    const have = copies.get(n)
    if (have) return have
    const c: TimeNode = 'children' in n
      ? { ...n, ctn: { ...n.ctn }, children: [] as TimeNode[] }
      : { ...n, ctn: { ...n.ctn } } as BehaviourNode
    copies.set(n, c)
    orderOf.set(c, options.orderOf?.get(n) ?? srcOrder.get(n) ?? srcOrder.size)
    if ('children' in n && 'children' in c) c.children = depth >= MAX_TREE_DEPTH ? [] : n.children.map(k => copy(k, depth + 1))
    return c
  }
  const root = copy(tree.root, 0)
  const fan = new Map<BehaviourNode, IterateFan>()
  const items = new Map<TargetKey, TargetKey[]>()
  const extend = new Map<ContainerNode, number>()
  for (const p of plans) {
    const c = copies.get(p.node) as BehaviourNode
    const src = p.node.ctn
    const span = finiteMs(src.durMs)! * (src.autoRev ? 2 : 1)
    const last = Math.max(0, ...p.delays)
    c.ctn.durMs = span + last
    c.ctn.accel = 0
    c.ctn.decel = 0
    c.ctn.autoRev = false
    const list = p.keys.map((key, i) => ({ key, delayMs: p.delays[i] }))
    // The body appears with the first item and, under an exit, leaves with the last.
    if (p.whole) list.unshift({ key: p.whole, delayMs: p.exit ? last : 0 })
    fan.set(c, { items: list, spanMs: span, totalMs: span + last, accel: src.accel, decel: src.decel, autoRev: src.autoRev, entrance: p.entrance })
    items.set(targetKey(p.node.target), [...p.keys])
    if (p.owner) extend.set(p.owner, Math.max(extend.get(p.owner) ?? 0, last))
    diag('text-build-approximate', p.effect, 'Text by word or letter: each word or letter box is its own layer (closest look-alike)')
  }
  for (const [owner, last] of extend) {
    const dur = finiteMs(owner.ctn.durMs)
    if (dur !== undefined) (copies.get(owner) as ContainerNode).ctn.durMs = dur + last
  }
  return { tree: { ...tree, root }, orderOf, fan, items }
}
