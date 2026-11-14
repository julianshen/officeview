/**
 * OVSP: the SHOW_PLAN (kind 6) wire codec (streaming Stage 3 piece 4, D1/D2/D11/D14/D22).
 *
 * Shared by the client and the server, DOM-free and package-free: nothing here
 * reaches `prepare/`, `morph/match`, `pptx/`, `core/package` or `render/` at
 * runtime (those are `import type` only), so the client bundle never pulls the
 * server's preparation code in through this module.
 *
 * One charged request returns several kind-6 messages under one reqId. Each
 * payload is `u8 tag` + body:
 *   1  the plan record (UTF-8 JSON, one per slide bundle)
 *   2  a layer asset        3  the entry frame        4  a keyframe
 *   5  a Morph record (UTF-8 JSON)
 * An asset body is `u32 LE unit` + `u32 LE assetId` + `encodeRasterBody(w, h, png)`.
 * `unit` is a server unit index. In a show-morph reply every asset carries the
 * `to` slide's server unit; this module enforces only the generic rule that an
 * asset's unit equals `expect.unit`.
 *
 * Fail closed. `decodeShowMessage` never returns a record that violates the
 * closed key sets, the D2 caps, the unique-assetId rule, the Σ decodedBytes
 * budget or the timing-tree limits: it throws `PlanRejected(path)` naming the
 * first offending field. The JSON text is parsed only after its byte cap is
 * checked. `encodeShowMessage` runs the same validation on what it produces, so
 * a server can never frame what a client would refuse.
 *
 * Paths are dotted field paths from the record root (`layers[1].assetId`,
 * `show.timing.root.children[0].ctn.extra`); a top-level unknown key is the key
 * itself. Message-level failures are `payload`, `tag`, `plan`/`morph` (byte cap,
 * UTF-8, JSON syntax, non-object), `body`, `png`, `unit`, `w`, `h`, `decodedBytes`.
 */
import { READER_LIMITS } from '../slideshow/read/limits'
import { decodeRasterBody, encodeRasterBody, RasterBodyError } from './raster'
import type { BaseGeometry } from '../slideshow/engine/behaviours'
import type { MorphRule } from '../slideshow/morph/match'
import type { MorphItemRef } from '../slideshow/prepare/morph'
import type { MaskKind, ShowSlide, TargetState } from '../slideshow/types'

export const SHOW_PLAN_LIMITS = Object.freeze({
  layers: 128, steps: 64, trackFrames: 90, slideKeyframes: 512, rectsPerStep: 32, triggers: 64, links: 64, morphItems: 512,
  planRecordBytes: 4194304, assetDecoded: 67108864, showDecoded: 134217728, str: 256,
} as const)
/** D22: every `TargetKey` string in `iterate`, keys and groups together. Not in `SHOW_PLAN_LIMITS`, whose shape is fixed. */
export const SHOW_PLAN_ITERATE_ITEMS = 128
/** D21. A diagnostic outside this list rejects the record (fail closed). */
export const SHOW_DIAGNOSTIC_CODES = [
  'step-composited', 'step-truncated', 'step-deferred', 'trigger-cut', 'morph-late', 'morph-composited', 'morph-fallback',
  'show-budget', 'show-limits', 'show-plan-timeout', 'link-withheld', 'preview-reduced', 'plan-stage-pressure', 'mark-rest', 'cmd-dropped',
] as const

export interface Rect { x: number; y: number; w: number; h: number }
export interface PlanLayerRecord { id: number; key?: string; part?: 'fill' | 'line' | 'text' | 'body' | 'range'; parent?: number; bounds: Rect; z: number; assetId: number; decodedBytes: number; scale: number; maskBox?: Rect; background?: true }
export interface PlanStepRecord { trigger: 'click' | 'after' | 'with' | 'shape'; eligibility: 'layers' | 'composited'; reason?: string; track?: { fps: 15; count: number } }
export interface ShowPlanRecord {
  v: 1
  slide: { index: number; w: number; h: number }
  options: { showAnimation: boolean; useTimings: boolean }
  show: Pick<ShowSlide, 'transition' | 'timing' | 'builds'>
  layers: PlanLayerRecord[]
  base: Array<[string, BaseGeometry]>
  steps: PlanStepRecord[]
  triggers: Array<{ spid: string; rects: Rect[] }>
  linkRegions: Array<{ rect: Rect; action: { kind: 'slide'; index: number } | { kind: 'jump'; verb: string } | { kind: 'withheld' } }>
  iterate: { keys: Array<[string, string[]]>; groups: Array<[string, { wd?: string[][]; lt?: string[][] }]> }
  entry: number
  diagnostics: Array<{ code: string; step?: number }>
}
export interface MorphSideRecord { assetId: number; geometry: BaseGeometry; state?: TargetState; part?: 'fill' | 'line' | 'text'; key?: string }
export interface MorphPlanRecord {
  v: 1; from: number; to: number; fromStep: number
  items: Array<Omit<MorphItemRef, 'from' | 'to'> & { from?: MorphSideRecord; to?: MorphSideRecord }>
  over: number[]
  pairs: Array<{ fromKey: string; toKey: string; rule: MorphRule }>
  /** 1..90 frames at 15 fps; `items` is empty when present (D11). */
  composited?: { fps: 15; count: number }
}

export class PlanRejected extends Error {
  constructor(readonly path: string) { super(`show plan rejected at ${path}`); this.name = 'PlanRejected' }
}

export type ShowMessage =
  | { tag: 1; plan: ShowPlanRecord }
  | { tag: 5; morph: MorphPlanRecord }
  | { tag: 2 | 3 | 4; unit: number; assetId: number; w: number; h: number; png: Uint8Array }

const MAX_PX = 8192
const MASK_KINDS = ['wipe', 'split', 'circle', 'diamond', 'plus', 'wedge', 'wheel', 'barn', 'checkerboard', 'dissolve', 'blinds', 'box', 'comb', 'randomBar', 'strips'] as const
// Compile-time: the runtime list names every MaskKind.
const _masks: Exclude<MaskKind, (typeof MASK_KINDS)[number]> extends never ? true : never = true
void _masks
const PARTS = ['fill', 'line', 'text', 'body', 'range'] as const
const SIDE_PARTS = ['fill', 'line', 'text'] as const
const RULES = ['!!', 'name-kind', 'text', 'geometry'] as const
const VERBS = ['nextslide', 'previousslide', 'firstslide', 'lastslide', 'endshow'] as const

// ---------------------------------------------------------------------------
// Walker primitives. Every failure is a PlanRejected carrying the field path.
// ---------------------------------------------------------------------------

type O = Record<string, unknown>
const fail = (path: string): never => { throw new PlanRejected(path) }
const at = (p: string, k: string) => (p ? `${p}.${k}` : k)
const ix = (p: string, i: number) => `${p}[${i}]`
const has = (o: O, k: string) => Object.prototype.hasOwnProperty.call(o, k) && o[k] !== undefined
const isObj = (v: unknown): v is O => !!v && typeof v === 'object' && !Array.isArray(v)

/** A plain object whose defined keys are all in `req ∪ opt` and which holds every `req` key. */
function obj(v: unknown, p: string, req: readonly string[], opt: readonly string[] = []): O {
  if (!isObj(v)) return fail(p)
  for (const k of Object.keys(v)) if (v[k] !== undefined && !req.includes(k) && !opt.includes(k)) fail(at(p, k))
  for (const k of req) if (!has(v, k)) fail(at(p, k))
  return v
}
const num = (v: unknown, p: string, lo = -1e9, hi = 1e9): number => {
  if (typeof v !== 'number' || !Number.isFinite(v) || v < lo || v > hi) fail(p)
  return v as number
}
const int = (v: unknown, p: string, lo: number, hi: number): number => {
  if (!Number.isInteger(v) || (v as number) < lo || (v as number) > hi) fail(p)
  return v as number
}
const u32 = (v: unknown, p: string) => int(v, p, 0, 0xffffffff)
const bool = (v: unknown, p: string) => { if (typeof v !== 'boolean') fail(p) }
const yes = (v: unknown, p: string) => { if (v !== true) fail(p) }
const str = (v: unknown, p: string, max: number = SHOW_PLAN_LIMITS.str): string => {
  if (typeof v !== 'string' || v.length > max) fail(p)
  return v as string
}
const oneOf = (v: unknown, p: string, set: readonly string[]) => { if (typeof v !== 'string' || !set.includes(v)) fail(p) }
const arr = (v: unknown, p: string, max: number): unknown[] => {
  if (!Array.isArray(v) || v.length > max) fail(p)
  return v as unknown[]
}
const numOrStr = (v: unknown, p: string) => { if (typeof v === 'string') str(v, p); else num(v, p) }
/** An animation value: a number, or a string that may be a formula (the reader's `formulaBytes`, not the 256 of D2). */
const numOrFormula = (v: unknown, p: string) => { if (typeof v === 'string') str(v, p, READER_LIMITS.formulaBytes); else num(v, p) }
/** `string -> string` map (transition attrs). */
function strMap(v: unknown, p: string, max: number, value: (x: unknown, q: string) => void) {
  if (!isObj(v)) return fail(p)
  const keys = Object.keys(v)
  if (keys.length > max) fail(p)
  for (const k of keys) { if (k === '__proto__' || k.length > SHOW_PLAN_LIMITS.str) fail(at(p, k)); value(v[k], at(p, k)) }
}
const opt = <T>(o: O, k: string, p: string, f: (v: unknown, q: string) => T): T | undefined => (o[k] === undefined ? undefined : f(o[k], at(p, k)))

/** A box in slide px: finite, x/y within [-4·slide, 5·slide], and the far edge within 5·slide. */
function rect(v: unknown, p: string, W: number, H: number) {
  const o = obj(v, p, ['x', 'y', 'w', 'h'])
  const x = num(o.x, at(p, 'x'), -4 * W, 5 * W), y = num(o.y, at(p, 'y'), -4 * H, 5 * H)
  const w = num(o.w, at(p, 'w'), 0, 9 * W), h = num(o.h, at(p, 'h'), 0, 9 * H)
  if (x + w > 5 * W) fail(at(p, 'w'))
  if (y + h > 5 * H) fail(at(p, 'h'))
}
function geometry(v: unknown, p: string, W: number, H: number) {
  const o = obj(v, p, ['x', 'y', 'w', 'h', 'rot'])
  rect({ x: o.x, y: o.y, w: o.w, h: o.h }, p, W, H)
  // `rect` names the offending axis under `p`, which is the geometry's own path.
  num(o.rot, at(p, 'rot'), -1e7, 1e7)
}
const affine = (v: unknown, p: string) => {
  const a = arr(v, p, 6)
  if (a.length !== 6) fail(p)
  a.forEach((n, i) => num(n, ix(p, i), -1e7, 1e7))
}
const point = (v: unknown, p: string) => { const o = obj(v, p, ['x', 'y']); num(o.x, at(p, 'x')); num(o.y, at(p, 'y')) }

/** The D11 / compositor `TargetState`, sanity-bounded (a Morph record has no slide size: the bound is the worst-case unit). */
function targetState(v: unknown, p: string) {
  const o = obj(v, p, ['visible', 'opacity', 'transform'], ['mask', 'maskBox', 'maskTransform', 'color', 'motion'])
  bool(o.visible, at(p, 'visible')); num(o.opacity, at(p, 'opacity'))
  affine(o.transform, at(p, 'transform'))
  opt(o, 'mask', p, (m, q) => {
    const mo = obj(m, q, ['kind', 'progress', 'params'])
    oneOf(mo.kind, at(q, 'kind'), MASK_KINDS); num(mo.progress, at(q, 'progress'))
    strMap(mo.params, at(q, 'params'), 32, numOrStr)
  })
  opt(o, 'maskBox', p, (b, q) => rect(b, q, MAX_PX, MAX_PX))
  opt(o, 'maskTransform', p, affine)
  opt(o, 'color', p, (c, q) => {
    const co = obj(c, q, [], ['fill', 'line', 'text'])
    for (const k of ['fill', 'line', 'text']) opt(co, k, q, (m, r) => { if (arr(m, r, 20).length !== 20) fail(r); (m as unknown[]).forEach((n, i) => num(n, ix(r, i), -1e6, 1e6)) })
  })
  opt(o, 'motion', p, (m, q) => { const mo = obj(m, q, ['dx', 'dy']); num(mo.dx, at(q, 'dx')); num(mo.dy, at(q, 'dy')) })
}

// ---------------------------------------------------------------------------
// The timing tree (the reader's own limits; READER_LIMITS)
// ---------------------------------------------------------------------------

interface TimingCtx { nodes: number; keyframes: number; conditions: number }
const TARGET_SUBS = ['bg', 'subSp', 'chart', 'graphic', 'ink']

function targetRef(v: unknown, p: string) {
  if (!isObj(v)) return fail(p)
  if (has(v, 'bg')) yes(obj(v, p, ['bg']).bg, at(p, 'bg'))
  else if (has(v, 'sld')) yes(obj(v, p, ['sld']).sld, at(p, 'sld'))
  else if (has(v, 'snd')) str(obj(v, p, ['snd']).snd, at(p, 'snd'))
  else {
    const o = obj(v, p, ['spid'], ['text', 'sub'])
    str(o.spid, at(p, 'spid'))
    opt(o, 'text', p, (t, q) => { const to = obj(t, q, ['kind', 'st', 'end']); oneOf(to.kind, at(q, 'kind'), ['p', 'char']); int(to.st, at(q, 'st'), 0, 1e7); int(to.end, at(q, 'end'), 0, 1e7) })
    opt(o, 'sub', p, (s, q) => oneOf(s, q, TARGET_SUBS))
  }
}
function condition(v: unknown, p: string, ctx: TimingCtx) {
  if (++ctx.conditions > READER_LIMITS.conditions) fail(p)
  const o = obj(v, p, ['delayMs'], ['evt', 'tn', 'target', 'rtn'])
  if (o.delayMs !== 'indefinite') num(o.delayMs, at(p, 'delayMs'))
  opt(o, 'evt', p, (e, q) => str(e, q)); opt(o, 'tn', p, (n, q) => int(n, q, 0, 0xffffffff)); opt(o, 'target', p, targetRef); opt(o, 'rtn', p, (r, q) => str(r, q))
}
function ctnRecord(v: unknown, p: string, ctx: TimingCtx) {
  const o = obj(v, p, ['id', 'accel', 'decel', 'autoRev', 'start', 'end'],
    ['presetClass', 'presetId', 'presetSubtype', 'nodeType', 'grpId', 'durMs', 'repeatCount', 'repeatDurMs', 'fill', 'restart', 'syncBehavior', 'masterRel', 'afterEffect', 'endSync', 'iterate'])
  int(o.id, at(p, 'id'), -0x7fffffff, 0x7fffffff)
  num(o.accel, at(p, 'accel'), 0, 1); num(o.decel, at(p, 'decel'), 0, 1); bool(o.autoRev, at(p, 'autoRev'))
  opt(o, 'presetClass', p, (s, q) => str(s, q)); opt(o, 'nodeType', p, (s, q) => str(s, q))
  for (const k of ['presetId', 'presetSubtype', 'grpId']) opt(o, k, p, (n, q) => int(n, q, -0x7fffffff, 0x7fffffff))
  for (const k of ['durMs', 'repeatCount', 'repeatDurMs']) opt(o, k, p, (n, q) => { if (n !== 'indefinite') num(n, q, 0, 1e12) })
  opt(o, 'fill', p, (s, q) => oneOf(s, q, ['hold', 'remove', 'freeze', 'transition']))
  opt(o, 'restart', p, (s, q) => oneOf(s, q, ['always', 'whenNotActive', 'never']))
  opt(o, 'syncBehavior', p, (s, q) => oneOf(s, q, ['canSlip', 'locked']))
  opt(o, 'masterRel', p, (s, q) => oneOf(s, q, ['sameClick', 'lastClick', 'nextClick']))
  opt(o, 'afterEffect', p, bool)
  for (const k of ['start', 'end'] as const) arr(o[k], at(p, k), READER_LIMITS.conditions).forEach((c, i) => condition(c, ix(at(p, k), i), ctx))
  opt(o, 'endSync', p, (c, q) => condition(c, q, ctx))
  opt(o, 'iterate', p, (it, q) => {
    const io = obj(it, q, ['type', 'backwards'], ['tmAbsMs', 'tmPct'])
    oneOf(io.type, at(q, 'type'), ['el', 'wd', 'lt']); bool(io.backwards, at(q, 'backwards'))
    opt(io, 'tmAbsMs', q, (n, r) => num(n, r, 0, 1e12)); opt(io, 'tmPct', q, (n, r) => num(n, r, 0, 1e6))
  })
}

const BASE_REQ = ['kind', 'ctn', 'target', 'attrNames', 'additive', 'accumulate']
const BASE_OPT = ['override', 'xfrmType', 'rctx']
const BEHAVIOURS: Record<string, { req: string[]; opt: string[] }> = {
  set: { req: [], opt: ['to'] },
  anim: { req: ['keys'], opt: ['calcmode', 'valueType', 'by', 'from', 'to'] },
  animEffect: { req: [], opt: ['transition', 'filter', 'prLst', 'progress'] },
  animMotion: { req: [], opt: ['path', 'origin', 'rAngDeg', 'ptsTypes', 'by', 'from', 'to', 'rCtr'] },
  animRot: { req: [], opt: ['by', 'from', 'to'] },
  animScale: { req: [], opt: ['by', 'from', 'to', 'zoomContents'] },
  animClr: { req: ['clrSpc', 'dir'], opt: ['by', 'from', 'to'] },
  cmd: { req: ['cmdType', 'cmd'], opt: [] },
  audio: { req: [], opt: ['numSld', 'volume', 'muted', 'isNarration'] },
  video: { req: [], opt: ['numSld', 'volume', 'muted', 'fullScrn'] },
}
function behaviourFields(kind: string, o: O, p: string, ctx: TimingCtx) {
  const s = (k: string) => opt(o, k, p, (v, q) => str(v, q, READER_LIMITS.formulaBytes))
  const n = (k: string) => opt(o, k, p, (v, q) => num(v, q))
  const pt = (k: string) => opt(o, k, p, point)
  switch (kind) {
    case 'set': opt(o, 'to', p, numOrFormula); break
    case 'anim': {
      const keys = arr(o.keys, at(p, 'keys'), READER_LIMITS.keyframesPerBehaviour)
      if ((ctx.keyframes += keys.length) > READER_LIMITS.keyframesPerSlide) fail(at(p, 'keys'))
      keys.forEach((k, i) => {
        const q = ix(at(p, 'keys'), i), ko = obj(k, q, ['tm', 'value'], ['fmla'])
        if (ko.tm !== 'indefinite') num(ko.tm, at(q, 'tm'))
        numOrFormula(ko.value, at(q, 'value'))
        opt(ko, 'fmla', q, (f, r) => str(f, r, READER_LIMITS.formulaBytes))
      })
      opt(o, 'calcmode', p, (v, q) => oneOf(v, q, ['discrete', 'lin', 'fmla'])); opt(o, 'valueType', p, (v, q) => oneOf(v, q, ['str', 'num', 'clr']))
      s('by'); s('from'); s('to'); break
    }
    case 'animEffect':
      opt(o, 'transition', p, (v, q) => oneOf(v, q, ['in', 'out', 'none'])); s('filter'); s('prLst'); opt(o, 'progress', p, numOrFormula); break
    case 'animMotion':
      opt(o, 'path', p, (v, q) => {
        const po = obj(v, q, ['segments', 'length'])
        num(po.length, at(q, 'length'), 0, 1e9)
        arr(po.segments, at(q, 'segments'), READER_LIMITS.motionPathCommands).forEach((sg, i) => {
          const r = ix(at(q, 'segments'), i), so = obj(sg, r, ['cmd', 'pts'])
          oneOf(so.cmd, at(r, 'cmd'), ['M', 'L', 'C', 'Z'])
          arr(so.pts, at(r, 'pts'), 6).forEach((x, j) => num(x, ix(at(r, 'pts'), j), -1e6, 1e6))
        })
      })
      opt(o, 'origin', p, (v, q) => oneOf(v, q, ['parent', 'layout'])); n('rAngDeg'); opt(o, 'ptsTypes', p, (v, q) => str(v, q, READER_LIMITS.motionPathCommands))
      pt('by'); pt('from'); pt('to'); pt('rCtr'); break
    case 'animRot': n('by'); n('from'); n('to'); break
    case 'animScale': pt('by'); pt('from'); pt('to'); opt(o, 'zoomContents', p, bool); break
    case 'animClr':
      oneOf(o.clrSpc, at(p, 'clrSpc'), ['rgb', 'hsl']); oneOf(o.dir, at(p, 'dir'), ['cw', 'ccw']); s('by'); s('from'); s('to'); break
    case 'cmd': oneOf(o.cmdType, at(p, 'cmdType'), ['evt', 'call', 'verb']); str(o.cmd, at(p, 'cmd')); break
    default: // audio, video
      opt(o, 'numSld', p, (v, q) => int(v, q, 0, 1e6)); opt(o, 'volume', p, (v, q) => num(v, q, 0, 1)); opt(o, 'muted', p, bool)
      opt(o, 'isNarration', p, bool); opt(o, 'fullScrn', p, bool)
  }
}
function timeNode(v: unknown, p: string, depth: number, ctx: TimingCtx) {
  if (depth > READER_LIMITS.nesting) fail(p)
  if (++ctx.nodes > READER_LIMITS.timeNodes) fail(p)
  if (!isObj(v)) return fail(p)
  const kind = v.kind
  if (kind === 'par' || kind === 'seq' || kind === 'excl') {
    const o = obj(v, p, ['kind', 'ctn', 'children'], ['seq'])
    ctnRecord(o.ctn, at(p, 'ctn'), ctx)
    opt(o, 'seq', p, (s, q) => { const so = obj(s, q, ['concurrent', 'nextAc', 'prevAc']); bool(so.concurrent, at(q, 'concurrent')); str(so.nextAc, at(q, 'nextAc')); str(so.prevAc, at(q, 'prevAc')) })
    arr(o.children, at(p, 'children'), READER_LIMITS.timeNodes).forEach((c, i) => timeNode(c, ix(at(p, 'children'), i), depth + 1, ctx))
    return
  }
  const spec = typeof kind === 'string' && Object.prototype.hasOwnProperty.call(BEHAVIOURS, kind) ? BEHAVIOURS[kind]! : fail(at(p, 'kind'))
  const o = obj(v, p, [...BASE_REQ, ...spec.req], [...BASE_OPT, ...spec.opt])
  ctnRecord(o.ctn, at(p, 'ctn'), ctx)
  targetRef(o.target, at(p, 'target'))
  arr(o.attrNames, at(p, 'attrNames'), 64).forEach((a, i) => str(a, ix(at(p, 'attrNames'), i)))
  oneOf(o.additive, at(p, 'additive'), ['base', 'sum', 'repl', 'mult', 'none']); oneOf(o.accumulate, at(p, 'accumulate'), ['always', 'none'])
  opt(o, 'override', p, (s, q) => oneOf(s, q, ['normal', 'childStyle'])); opt(o, 'xfrmType', p, (s, q) => oneOf(s, q, ['pt', 'img'])); opt(o, 'rctx', p, (s, q) => str(s, q))
  behaviourFields(kind as string, o, p, ctx)
}
function buildTemplate(v: unknown, p: string) {
  if (!isObj(v)) return fail(p)
  if (v.kind === 'bldP') {
    const o = obj(v, p, ['kind', 'spid', 'grpId', 'build', 'bldLvl', 'animBg', 'rev'])
    str(o.spid, at(p, 'spid')); int(o.grpId, at(p, 'grpId'), 0, 0x7fffffff); oneOf(o.build, at(p, 'build'), ['whole', 'p', 'cust', 'allAtOnce'])
    int(o.bldLvl, at(p, 'bldLvl'), 0, 1000); bool(o.animBg, at(p, 'animBg')); bool(o.rev, at(p, 'rev'))
  } else if (v.kind === 'bldGraphic' || v.kind === 'bldOleChart' || v.kind === 'bldDgm') {
    const o = obj(v, p, ['kind', 'spid', 'grpId'], ['bld'])
    str(o.spid, at(p, 'spid')); int(o.grpId, at(p, 'grpId'), 0, 0x7fffffff); opt(o, 'bld', p, (s, q) => str(s, q))
  } else fail(at(p, 'kind'))
}
function timing(v: unknown, p: string) {
  if (isObj(v) && has(v, 'refused')) { str(obj(v, p, ['refused']).refused, at(p, 'refused')); return }
  const o = obj(v, p, ['root', 'build'])
  timeNode(o.root, at(p, 'root'), 1, { nodes: 0, keyframes: 0, conditions: 0 })
  arr(o.build, at(p, 'build'), READER_LIMITS.timeNodes).forEach((b, i) => buildTemplate(b, ix(at(p, 'build'), i)))
}
function transition(v: unknown, p: string) {
  const o = obj(v, p, ['element', 'ns', 'attrs', 'durationMs', 'advClick'], ['preset', 'advTmMs', 'sound'])
  str(o.element, at(p, 'element')); oneOf(o.ns, at(p, 'ns'), ['p', 'p14', 'p15', 'p159'])
  strMap(o.attrs, at(p, 'attrs'), 64, (x, q) => { str(x, q) })
  num(o.durationMs, at(p, 'durationMs'), 0, 1e9); bool(o.advClick, at(p, 'advClick'))
  opt(o, 'preset', p, (s, q) => str(s, q)); opt(o, 'advTmMs', p, (n, q) => num(n, q, 0, 1e9))
  opt(o, 'sound', p, (s, q) => {
    const so = obj(s, q, [], ['start', 'end'])
    opt(so, 'start', q, (st, r) => { const to = obj(st, r, ['rId', 'loop']); str(to.rId, at(r, 'rId')); bool(to.loop, at(r, 'loop')) })
    opt(so, 'end', q, yes)
  })
}
function builds(v: unknown, p: string) {
  arr(v, p, READER_LIMITS.timeNodes).forEach((b, i) => {
    const q = ix(p, i), o = obj(b, q, ['spid', 'mode', 'ranges', 'animBg', 'body'])
    str(o.spid, at(q, 'spid')); oneOf(o.mode, at(q, 'mode'), ['whole', 'paragraph']); bool(o.animBg, at(q, 'animBg')); str(o.body, at(q, 'body'))
    arr(o.ranges, at(q, 'ranges'), READER_LIMITS.timeNodes).forEach((r, j) => {
      const rp = ix(at(q, 'ranges'), j), ro = obj(r, rp, ['key', 'st', 'end'])
      str(ro.key, at(rp, 'key')); int(ro.st, at(rp, 'st'), 0, 1e7); int(ro.end, at(rp, 'end'), 0, 1e7)
    })
  })
}

// ---------------------------------------------------------------------------
// The plan record (tag 1)
// ---------------------------------------------------------------------------

const PLAN_KEYS = ['v', 'slide', 'options', 'show', 'layers', 'base', 'steps', 'triggers', 'linkRegions', 'iterate', 'entry', 'diagnostics']

function validatePlan(v: unknown, expect: { unit: number; maxDecoded: number }): ShowPlanRecord {
  if (!isObj(v)) return fail('plan')
  const o = obj(v, '', PLAN_KEYS)
  if (o.v !== 1) fail('v')
  const sl = obj(o.slide, 'slide', ['index', 'w', 'h'])
  const index = u32(sl.index, 'slide.index'), W = num(sl.w, 'slide.w', 1e-6, MAX_PX), H = num(sl.h, 'slide.h', 1e-6, MAX_PX)
  if (index !== expect.unit) fail('slide.index')
  const op = obj(o.options, 'options', ['showAnimation', 'useTimings'])
  bool(op.showAnimation, 'options.showAnimation'); bool(op.useTimings, 'options.useTimings')
  const sh = obj(o.show, 'show', ['builds'], ['transition', 'timing'])
  opt(sh, 'transition', 'show', transition); opt(sh, 'timing', 'show', timing); builds(sh.builds, 'show.builds')

  const ids = new Set<number>()
  const layers = arr(o.layers, 'layers', SHOW_PLAN_LIMITS.layers)
  let sum = 0
  const sumCap = Math.min(SHOW_PLAN_LIMITS.showDecoded, expect.maxDecoded)
  layers.forEach((l, i) => {
    const p = ix('layers', i)
    const lo = obj(l, p, ['id', 'bounds', 'z', 'assetId', 'decodedBytes', 'scale'], ['key', 'part', 'parent', 'maskBox', 'background'])
    u32(lo.id, at(p, 'id'))
    opt(lo, 'key', p, (s, q) => str(s, q)); opt(lo, 'part', p, (s, q) => oneOf(s, q, PARTS)); opt(lo, 'parent', p, u32)
    rect(lo.bounds, at(p, 'bounds'), W, H); opt(lo, 'maskBox', p, (b, q) => rect(b, q, W, H))
    num(lo.z, at(p, 'z'), -1e9, 1e9); num(lo.scale, at(p, 'scale'), 1e-6, 64); opt(lo, 'background', p, yes)
    const id = u32(lo.assetId, at(p, 'assetId'))
    if (ids.has(id)) fail(at(p, 'assetId'))
    ids.add(id)
    sum += int(lo.decodedBytes, at(p, 'decodedBytes'), 1, SHOW_PLAN_LIMITS.assetDecoded)
    if (sum > sumCap) fail('layers.decodedBytes')
  })
  const entry = u32(o.entry, 'entry')
  if (ids.has(entry)) fail('entry')

  arr(o.base, 'base', READER_LIMITS.timeNodes).forEach((b, i) => {
    const p = ix('base', i)
    if (arr(b, p, 2).length !== 2) fail(p)
    str((b as unknown[])[0], ix(p, 0)); geometry((b as unknown[])[1], ix(p, 1), W, H)
  })
  let frames = 0
  arr(o.steps, 'steps', SHOW_PLAN_LIMITS.steps).forEach((s, i) => {
    const p = ix('steps', i), so = obj(s, p, ['trigger', 'eligibility'], ['reason', 'track'])
    oneOf(so.trigger, at(p, 'trigger'), ['click', 'after', 'with', 'shape']); oneOf(so.eligibility, at(p, 'eligibility'), ['layers', 'composited'])
    opt(so, 'reason', p, (r, q) => str(r, q))
    opt(so, 'track', p, (t, q) => {
      const to = obj(t, q, ['fps', 'count'])
      if (to.fps !== 15) fail(at(q, 'fps'))
      frames += int(to.count, at(q, 'count'), 1, SHOW_PLAN_LIMITS.trackFrames)
    })
  })
  if (frames > SHOW_PLAN_LIMITS.slideKeyframes) fail('steps')
  arr(o.triggers, 'triggers', SHOW_PLAN_LIMITS.triggers).forEach((t, i) => {
    const p = ix('triggers', i), to = obj(t, p, ['spid', 'rects'])
    str(to.spid, at(p, 'spid'))
    arr(to.rects, at(p, 'rects'), SHOW_PLAN_LIMITS.rectsPerStep).forEach((r, j) => rect(r, ix(at(p, 'rects'), j), W, H))
  })
  arr(o.linkRegions, 'linkRegions', SHOW_PLAN_LIMITS.links).forEach((l, i) => {
    const p = ix('linkRegions', i), lo = obj(l, p, ['rect', 'action'])
    rect(lo.rect, at(p, 'rect'), W, H)
    if (!isObj(lo.action)) return fail(at(p, 'action'))
    const ap = at(p, 'action')
    if (lo.action.kind === 'slide') u32(obj(lo.action, ap, ['kind', 'index']).index, at(ap, 'index'))
    else if (lo.action.kind === 'jump') oneOf(obj(lo.action, ap, ['kind', 'verb']).verb, at(ap, 'verb'), VERBS)
    else if (lo.action.kind === 'withheld') obj(lo.action, ap, ['kind'])
    else fail(at(ap, 'kind'))
  })

  const it = obj(o.iterate, 'iterate', ['keys', 'groups'])
  let items = 0
  const key = (x: unknown, p: string) => { str(x, p); if (++items > SHOW_PLAN_ITERATE_ITEMS) fail('iterate') }
  // A list's own length is bounded by the running total in `key` (the 129th key names `iterate`).
  const keyList = (x: unknown, p: string) => arr(x, p, Infinity).forEach((k, i) => key(k, ix(p, i)))
  arr(it.keys, 'iterate.keys', SHOW_PLAN_ITERATE_ITEMS).forEach((k, i) => {
    const p = ix('iterate.keys', i)
    if (arr(k, p, 2).length !== 2) fail(p)
    key((k as unknown[])[0], ix(p, 0)); keyList((k as unknown[])[1], ix(p, 1))
  })
  arr(it.groups, 'iterate.groups', SHOW_PLAN_ITERATE_ITEMS).forEach((g, i) => {
    const p = ix('iterate.groups', i)
    if (arr(g, p, 2).length !== 2) fail(p)
    key((g as unknown[])[0], ix(p, 0))
    const go = obj((g as unknown[])[1], ix(p, 1), [], ['wd', 'lt'])
    for (const k of ['wd', 'lt']) opt(go, k, ix(p, 1), (l, q) => arr(l, q, SHOW_PLAN_ITERATE_ITEMS).forEach((x, j) => keyList(x, ix(q, j))))
  })
  arr(o.diagnostics, 'diagnostics', 1024).forEach((d, i) => {
    const p = ix('diagnostics', i), dobj = obj(d, p, ['code'], ['step'])
    oneOf(dobj.code, at(p, 'code'), SHOW_DIAGNOSTIC_CODES); opt(dobj, 'step', p, (n, q) => int(n, q, 0, SHOW_PLAN_LIMITS.steps))
  })
  return o as unknown as ShowPlanRecord
}

// ---------------------------------------------------------------------------
// The Morph record (tag 5)
// ---------------------------------------------------------------------------

function validateMorph(v: unknown): MorphPlanRecord {
  if (!isObj(v)) return fail('morph')
  const o = obj(v, '', ['v', 'from', 'to', 'fromStep', 'items', 'over', 'pairs'], ['composited'])
  if (o.v !== 1) fail('v')
  u32(o.from, 'from'); u32(o.to, 'to'); int(o.fromStep, 'fromStep', 0, SHOW_PLAN_LIMITS.steps)
  const composited = opt(o, 'composited', '', (c, q) => {
    const co = obj(c, q, ['fps', 'count'])
    if (co.fps !== 15) fail(at(q, 'fps'))
    return int(co.count, at(q, 'count'), 1, SHOW_PLAN_LIMITS.trackFrames)
  })
  const ids = new Set<number>()
  const claim = (x: unknown, p: string) => { const id = u32(x, p); if (ids.has(id)) fail(p); ids.add(id) }
  const items = arr(o.items, 'items', SHOW_PLAN_LIMITS.morphItems)
  if (composited !== undefined && items.length) fail('composited')
  items.forEach((it, i) => {
    const p = ix('items', i), io = obj(it, p, [], ['from', 'to', 'under', 'background', 'path', 'pathStates'])
    if (!has(io, 'from') && !has(io, 'to')) fail(p)
    for (const side of ['from', 'to']) {
      opt(io, side, p, (s, q) => {
        const so = obj(s, q, ['assetId', 'geometry'], ['state', 'part', 'key'])
        claim(so.assetId, at(q, 'assetId')); geometry(so.geometry, at(q, 'geometry'), MAX_PX, MAX_PX)
        opt(so, 'state', q, targetState); opt(so, 'part', q, (x, r) => oneOf(x, r, SIDE_PARTS)); opt(so, 'key', q, (x, r) => str(x, r))
      })
    }
    opt(io, 'under', p, bool); opt(io, 'background', p, bool)
    opt(io, 'path', p, (s, q) => { const so = obj(s, q, ['from', 'to']); geometry(so.from, at(q, 'from'), MAX_PX, MAX_PX); geometry(so.to, at(q, 'to'), MAX_PX, MAX_PX) })
    opt(io, 'pathStates', p, (s, q) => { const so = obj(s, q, [], ['from', 'to']); opt(so, 'from', q, targetState); opt(so, 'to', q, targetState) })
  })
  arr(o.over, 'over', SHOW_PLAN_LIMITS.morphItems).forEach((x, i) => claim(x, ix('over', i)))
  arr(o.pairs, 'pairs', SHOW_PLAN_LIMITS.morphItems).forEach((pr, i) => {
    const p = ix('pairs', i), po = obj(pr, p, ['fromKey', 'toKey', 'rule'])
    str(po.fromKey, at(p, 'fromKey')); str(po.toKey, at(p, 'toKey')); oneOf(po.rule, at(p, 'rule'), RULES)
  })
  return o as unknown as MorphPlanRecord
}

// ---------------------------------------------------------------------------
// Messages
// ---------------------------------------------------------------------------

const utf8 = new TextEncoder()
const fatal = new TextDecoder('utf-8', { fatal: true })

function parseRecord(payload: Uint8Array, label: 'plan' | 'morph'): unknown {
  // The byte cap is checked before any decode or parse.
  if (payload.length - 1 > SHOW_PLAN_LIMITS.planRecordBytes) fail(label)
  try { return JSON.parse(fatal.decode(payload.subarray(1))) } catch { return fail(label) }
}

export function decodeShowMessage(payload: Uint8Array, expect: { unit: number; maxDecoded: number }): ShowMessage {
  if (payload.length === 0) fail('payload')
  const tag = payload[0]!
  if (tag === 1) return { tag, plan: validatePlan(parseRecord(payload, 'plan'), expect) }
  if (tag === 5) return { tag, morph: validateMorph(parseRecord(payload, 'morph')) }
  if (tag < 2 || tag > 4) return fail('tag')
  if (payload.length < 1 + 8 + 12) fail('body')
  const v = new DataView(payload.buffer, payload.byteOffset, payload.byteLength)
  const unit = v.getUint32(1, true)
  if (unit !== expect.unit) fail('unit')
  const assetId = v.getUint32(5, true), w = v.getUint32(9, true), h = v.getUint32(13, true)
  if (w < 1 || w > MAX_PX) fail('w')
  if (h < 1 || h > MAX_PX) fail('h')
  const cap = Math.min(expect.maxDecoded, SHOW_PLAN_LIMITS.assetDecoded)
  if (w * h * 4 > cap) fail('decodedBytes')
  try {
    return { tag: tag as 2 | 3 | 4, unit, assetId, ...decodeRasterBody(payload.subarray(9), { maxRgbaBytes: cap }) }
  } catch (e) {
    if (!(e instanceof RasterBodyError)) throw e
    return fail(e.message === 'decoded-cap' ? 'decodedBytes' : e.message === 'truncated' || e.message === 'length' ? 'body' : 'png')
  }
}

export function encodeShowMessage(m: ShowMessage): Uint8Array {
  let out: Uint8Array, expect: { unit: number; maxDecoded: number }
  if (m.tag === 1 || m.tag === 5) {
    const json = utf8.encode(JSON.stringify(m.tag === 1 ? m.plan : m.morph))
    out = new Uint8Array(1 + json.length); out[0] = m.tag; out.set(json, 1)
    expect = { unit: m.tag === 1 ? m.plan.slide.index : 0, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }
  } else {
    if (!(m.tag >= 2 && m.tag <= 4)) fail('tag')
    const body = encodeRasterBody(m.w, m.h, m.png)
    out = new Uint8Array(9 + body.length)
    const v = new DataView(out.buffer)
    out[0] = m.tag; v.setUint32(1, u32(m.unit, 'unit'), true); v.setUint32(5, u32(m.assetId, 'assetId'), true); out.set(body, 9)
    expect = { unit: m.unit, maxDecoded: SHOW_PLAN_LIMITS.assetDecoded }
  }
  // What this module frames is exactly what its decoder accepts.
  decodeShowMessage(out, expect)
  return out
}
