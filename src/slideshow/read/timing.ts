/**
 * Timing-tree reader (slideshow spec 6.1 semantics table, unknown-node policy
 * and reader limits; §6.2 model).
 *
 * Units are normalised here: times in ms, `repeatCount` in iterations (the file
 * writes thousandths), `accel`/`decel` and keyframe times as fractions (1/1000
 * of a percent in the file), angles in degrees (1/60000°), scale and motion
 * points as fractions (1/1000 of a percent).
 *
 * Policy:
 * - An unknown element inside a known container is dropped with a diagnostic
 *   naming it; its siblings stay. An unknown element with element children is
 *   an unknown container: its whole subtree is dropped.
 * - An invalid or unknown attribute value is dropped (the field is absent) with
 *   a diagnostic; an unknown condition event drops that condition.
 * - Exceeding a reader limit refuses the slide's timing (`{ refused }`, the
 *   slide is static) and its builds.
 * - A behaviour whose target is missing is dropped (`unresolved-target`), or
 *   found only on the layout/master (`master-target`); its container stays.
 * - A `tn` reference cycle is broken at its closing condition, which becomes a
 *   condition that never fires; diagnosed once.
 * - `syncBehavior`, `masterRel` and `rctx` are carried but not acted on
 *   (`timing-carried`, once per slide each); attributes the model has no field
 *   for (`display`, `evtFilter`, `numSld`, ...) are `timing-ignored`.
 *
 * Shape (the S1-9 scheduler's contract): the `p:tnLst` root `par` is `root`;
 * sequences keep `nodeType` verbatim; `delay="indefinite"` with no event is
 * `{ delayMs: 'indefinite' }` exactly; absent `dur`/`fill` stay undefined.
 * `p:subTnLst` children join `children` after the `childTnLst` ones.
 */
import { attrs, orderedChildren, textOf, type XmlNode } from '../../core/xml'
import { targetKey } from '../target'
import type {
  AnimBehaviour, AnimClrBehaviour, AnimEffectBehaviour, AnimMotionBehaviour, AnimRotBehaviour, AnimScaleBehaviour,
  AudioBehaviour, BehaviourBase, BehaviourNode, BuildOwnership, BuildTemplate, CmdBehaviour, CTn, Condition, ContainerNode,
  DiagnosticSink, Iterate, Keyframe, MotionPath, Point, SetBehaviour, SlideTimingReader, TargetRef, TargetResolver, TimeNode,
  TimingTree, TransitionSpec, VideoBehaviour,
} from '../types'
import { compatibilityBranches } from './requirements'
import { readTransition } from './transition'

/** Spec 6.1 reader limits for one slide's timing. */
export const TIMING_LIMITS = Object.freeze({ timeNodes: 2000, nesting: 64, conditions: 10000, keyframesPerBehaviour: 1000, keyframesPerSlide: 20000, motionPathBytes: 65536 })

/** `p:attrName` values the reader accepts (ECMA-376 Part 1 §19.5.2 list plus the PowerPoint `ppt_*` forms); others are dropped with a diagnostic. */
export const ATTRIBUTE_NAMES: ReadonlySet<string> = new Set([
  'ppt_x', 'ppt_y', 'ppt_w', 'ppt_h', 'ppt_c', 'ppt_r', 'r', 'xshear', 'yshear', 'scaleX', 'scaleY', 'image',
  'style.color', 'style.fontSize', 'style.fontStyle', 'style.fontWeight', 'style.fontFamily', 'style.opacity', 'style.rotation',
  'style.sRotation', 'style.visibility', 'style.textDecorationUnderline', 'style.textDecorationLineThrough', 'style.textEffectEmboss',
  'style.textEffectOutline', 'style.textShadow', 'style.textTransform',
  'fillcolor', 'fillColor', 'fill.type', 'fill.on', 'fill.color', 'fill.color2', 'fill.opacity', 'fill.opacity2', 'fill.method',
  'fill.angle', 'fill.focus', 'fill.focusposition', 'fill.focussize', 'fill.radialfocus',
  'stroke.on', 'stroke.color', 'stroke.color2', 'stroke.weight', 'stroke.opacity', 'stroke.linestyle', 'stroke.dashstyle',
  'stroke.filltype', 'stroke.src', 'stroke.imagesize', 'stroke.startArrow', 'stroke.endArrow', 'stroke.startArrowWidth',
  'stroke.startArrowLength', 'stroke.endArrowWidth', 'stroke.endArrowLength',
  'shadow.on', 'shadow.type', 'shadow.color', 'shadow.color2', 'shadow.opacity', 'shadow.offset', 'shadow.offset2', 'shadow.origin',
  'shadow.matrix', 'shadow.visible',
  'imageData.cropTop', 'imageData.cropBottom', 'imageData.cropLeft', 'imageData.cropRight', 'imageData.gain', 'imageData.blackLevel',
  'imageData.gamma', 'imageData.grayscale', 'imageData.chromakey',
  'skew.on', 'skew.offset', 'skew.origin', 'skew.matrix',
  'extrusion.on', 'extrusion.type', 'extrusion.render', 'extrusion.viewpointorigin', 'extrusion.viewpoint', 'extrusion.plane',
  'extrusion.skewangle', 'extrusion.skewamt', 'extrusion.backdepth', 'extrusion.foredepth', 'extrusion.orientation',
  'extrusion.orientationangle', 'extrusion.rotationangle', 'extrusion.color', 'extrusion.shininess',
])

type MotionPathParser = (src: string) => { ok: true; path: MotionPath } | { ok: false; reason: string }
export interface TimingReaderOptions {
  /** S1-6's bounded `parseMotionPath`; without it `animMotion@path` is diagnosed `motion-path-deferred`. */
  parseMotionPath?: MotionPathParser
}

class Refused extends Error { constructor(readonly why: string) { super(why) } }

interface Ctx {
  slideIndex: number
  sink: DiagnosticSink
  resolve: TargetResolver
  parsePath?: MotionPathParser
  nodes: number
  conds: number
  tavs: number
  synthetic: number
  ids: Map<number, CTn>
  textRefs: Map<string, Array<{ st: number; end: number }>>
}

const CONTAINERS = new Set(['par', 'seq', 'excl'])
const BEHAVIOURS = new Set(['set', 'anim', 'animEffect', 'animMotion', 'animRot', 'animScale', 'animClr', 'cmd', 'audio', 'video'])
const EVENTS = new Set(['onBegin', 'onEnd', 'onClick', 'onNext', 'onPrev', 'begin', 'end', 'onMouseOver', 'onMouseOut', 'onStopAudio', 'onDblClick'])
const PRESET_CLASSES = new Set(['entr', 'exit', 'emph', 'path', 'verb', 'mediacall'])
const NODE_TYPES = new Set(['clickEffect', 'withEffect', 'afterEffect', 'mainSeq', 'interactiveSeq', 'clickPar', 'withGroup', 'afterGroup', 'tmRoot'])
const CARRIED_CTN = new Set(['syncBehavior', 'masterRel'])
const IGNORED_CTN = new Set(['spd', 'tmFilter', 'evtFilter', 'display', 'bldLvl', 'nodePh', 'presetBounceEnd', 'bounceEnd'])
const CTN_ATTRS = new Set(['id', 'presetID', 'presetClass', 'presetSubtype', 'dur', 'repeatCount', 'repeatDur', 'accel', 'decel', 'autoRev', 'restart', 'fill', 'grpId', 'afterEffect', 'nodeType', ...CARRIED_CTN, ...IGNORED_CTN])

const kids = (node: XmlNode | undefined) => orderedChildren(node).filter(([name]) => name !== '#text')
const child = (node: XmlNode | undefined, name: string) => kids(node).find(([n]) => n === name)?.[1]
const hasElements = (node: XmlNode) => kids(node).length > 0

function add(ctx: Ctx, code: string, effect: string, message: string): void {
  ctx.sink.add({ code, slideIndex: ctx.slideIndex, effect, message })
}
const unknownNode = (ctx: Ctx, name: string, node: XmlNode, where: string) => hasElements(node)
  ? add(ctx, 'timing-unknown-container', name, `Unknown container ${name} in ${where}; subtree dropped`)
  : add(ctx, 'timing-unknown-element', name, `Unknown element ${name} in ${where}; dropped`)

// ---------------------------------------------------------------------------
// Values
// ---------------------------------------------------------------------------

const UINT = /^\s*\d+\s*$/
const INT = /^\s*[-+]?\d+\s*$/
function uint(v: string | undefined): number | undefined { return v !== undefined && UINT.test(v) && Number.isSafeInteger(Number(v)) ? Number(v) : undefined }
function int(v: string | undefined): number | undefined { return v !== undefined && INT.test(v) && Number.isSafeInteger(Number(v)) ? Number(v) : undefined }
function bool(v: string | undefined): boolean | undefined { return v === '1' || v === 'true' ? true : v === '0' || v === 'false' ? false : undefined }
function time(v: string | undefined): number | 'indefinite' | undefined { return v === 'indefinite' ? v : uint(v) }

/** Reads an attribute with `parse`; an invalid value is diagnosed and absent. */
function value<T>(ctx: Ctx, a: Record<string, string>, el: string, name: string, parse: (v: string) => T | undefined): T | undefined {
  const raw = a[name]
  if (raw === undefined) return undefined
  const out = parse(raw)
  if (out === undefined) add(ctx, 'timing-invalid-value', `${el}@${name}`, `Invalid ${el}@${name}="${raw}"; dropped`)
  return out
}
const oneOf = <T extends string>(...allowed: T[]) => (v: string): T | undefined => (allowed as string[]).includes(v) ? v as T : undefined
const fraction = (v: string): number | undefined => { const n = int(v); return n === undefined ? undefined : n / 100000 }
const degrees = (v: string): number | undefined => { const n = int(v); return n === undefined ? undefined : n / 60000 }
/** `ST_TLTimeAnimateValueTime`: 1/1000 % (or `n%`), or `indefinite`; as a fraction. */
function keyTime(v: string): number | 'indefinite' | undefined {
  if (v === 'indefinite') return v
  const pct = /^\s*(\d+(?:\.\d+)?)%\s*$/.exec(v)
  if (pct) return Number(pct[1]) / 100
  const n = uint(v)
  return n === undefined || n > 100000 ? undefined : n / 100000
}

function point(ctx: Ctx, node: XmlNode | undefined, el: string): Point | undefined {
  if (!node) return undefined
  const a = attrs(node)
  const x = value(ctx, a, el, 'x', fraction)
  const y = value(ctx, a, el, 'y', fraction)
  return x === undefined || y === undefined ? undefined : { x, y }
}

/** `CT_TLAnimVariant`: number for int/float, string otherwise. */
function variant(ctx: Ctx, node: XmlNode | undefined, el: string): string | number | undefined {
  const first = kids(node)[0]
  if (!first) return undefined
  const [name, v] = first
  const val = attrs(v).val
  switch (name) {
    case 'strVal': return val ?? ''
    case 'boolVal': return bool(val) === undefined ? undefined : String(bool(val))
    case 'intVal': { const n = int(val); if (n === undefined) add(ctx, 'timing-invalid-value', `${el}/intVal`, `Invalid intVal ${val}`); return n }
    case 'fltVal': { const n = Number(val); if (val === undefined || val.trim() === '' || !Number.isFinite(n)) { add(ctx, 'timing-invalid-value', `${el}/fltVal`, `Invalid fltVal ${val}`); return undefined } return n }
    case 'clrVal': return color(ctx, v, el)
    default: unknownNode(ctx, name, v, el); return undefined
  }
}

/** A DrawingML colour choice as `#rrggbb` (sRGB) or `kind:value` as written. */
function color(ctx: Ctx, holder: XmlNode | undefined, el: string): string | undefined {
  const first = kids(holder)[0]
  if (!first) return undefined
  const [name, c] = first
  const a = attrs(c)
  if (kids(c).length) add(ctx, 'timing-ignored', `${el}/${name}`, `Colour transforms on ${el} are not applied`)
  if (name === 'srgbClr' && a.val && /^[0-9a-fA-F]{6}$/.test(a.val)) return `#${a.val.toLowerCase()}`
  if (name === 'scrgbClr') return `scrgbClr:${a.r},${a.g},${a.b}`
  if (name === 'hslClr') return `hslClr:${a.hue},${a.sat},${a.lum}`
  if (['schemeClr', 'sysClr', 'prstClr'].includes(name) && a.val) return `${name}:${a.val}`
  add(ctx, 'timing-invalid-value', `${el}/${name}`, `Unreadable colour ${name}`)
  return undefined
}

// ---------------------------------------------------------------------------
// Conditions, targets, cTn
// ---------------------------------------------------------------------------

function countCond(ctx: Ctx): void {
  if (++ctx.conds > TIMING_LIMITS.conditions) throw new Refused('conditions')
}

function target(ctx: Ctx, tgtEl: XmlNode | undefined): TargetRef | undefined {
  const first = kids(tgtEl)[0]
  if (!first) return undefined
  const [name, node] = first
  const a = attrs(node)
  switch (name) {
    case 'sldTgt': return { sld: true }
    case 'sndTgt': return a.embed ? { snd: a.embed } : a.name ? { snd: `builtIn:${a.name}` } : undefined
    case 'inkTgt': return a.spid ? { spid: a.spid, sub: 'ink' } : undefined
    case 'spTgt': {
      if (!a.spid) return undefined
      const inner = kids(node)[0]
      if (!inner) return { spid: a.spid }
      const [sub, el] = inner
      switch (sub) {
        case 'bg': return { spid: a.spid, sub: 'bg' }
        case 'subSpId': return { spid: a.spid, sub: 'subSp' }
        case 'oleChartEl': add(ctx, 'timing-approximated', 'oleChartEl', 'Chart element targets animate the whole chart'); return { spid: a.spid, sub: 'chart' }
        case 'graphicEl': add(ctx, 'timing-approximated', 'graphicEl', 'Graphic element targets animate the whole graphic'); return { spid: a.spid, sub: 'graphic' }
        case 'txEl': {
          const range = kids(el)[0]
          if (!range || (range[0] !== 'pRg' && range[0] !== 'charRg')) { if (range) unknownNode(ctx, range[0], range[1], 'txEl'); return undefined }
          const st = value(ctx, attrs(range[1]), range[0], 'st', uint)
          const end = value(ctx, attrs(range[1]), range[0], 'end', uint)
          if (st === undefined || end === undefined || end < st) return undefined
          return { spid: a.spid, text: { kind: range[0] === 'pRg' ? 'p' : 'char', st, end } }
        }
        default: unknownNode(ctx, sub, el, 'spTgt'); return undefined
      }
    }
    default: unknownNode(ctx, name, node, 'tgtEl'); return undefined
  }
}

function condition(ctx: Ctx, node: XmlNode): Condition | undefined {
  countCond(ctx)
  const a = attrs(node)
  const delay = a.delay === undefined ? 0 : time(a.delay)
  if (delay === undefined) { add(ctx, 'timing-invalid-value', 'cond@delay', `Invalid cond@delay="${a.delay}"; condition dropped`); return undefined }
  if (a.evt !== undefined && !EVENTS.has(a.evt)) { add(ctx, 'timing-invalid-value', 'cond@evt', `Unknown event ${a.evt}; condition dropped`); return undefined }
  const out: Condition = { delayMs: delay }
  if (a.evt !== undefined) out.evt = a.evt
  for (const [name, c] of kids(node)) {
    if (name === 'tn') { const v = value(ctx, attrs(c), 'tn', 'val', uint); if (v !== undefined) out.tn = v }
    else if (name === 'rtn') { const v = value(ctx, attrs(c), 'rtn', 'val', oneOf('first', 'last', 'all')); if (v !== undefined) out.rtn = v }
    else if (name === 'tgtEl') { const t = target(ctx, c); if (t) out.target = t }
    else unknownNode(ctx, name, c, 'cond')
  }
  return out
}

function conditionList(ctx: Ctx, node: XmlNode | undefined): Condition[] {
  const out: Condition[] = []
  for (const [name, c] of kids(node)) {
    if (name !== 'cond') { unknownNode(ctx, name, c, 'condition list'); continue }
    const cond = condition(ctx, c)
    if (cond) out.push(cond)
  }
  return out
}

function iterate(ctx: Ctx, node: XmlNode): Iterate {
  const a = attrs(node)
  const out: Iterate = { type: value(ctx, a, 'iterate', 'type', oneOf('el', 'wd', 'lt')) ?? 'el', backwards: value(ctx, a, 'iterate', 'backwards', bool) ?? false }
  for (const [name, c] of kids(node)) {
    if (name === 'tmAbs') { const v = value(ctx, attrs(c), 'tmAbs', 'val', time); if (typeof v === 'number') out.tmAbsMs = v }
    else if (name === 'tmPct') { const v = value(ctx, attrs(c), 'tmPct', 'val', uint); if (v !== undefined) out.tmPct = v / 100000 }
    else unknownNode(ctx, name, c, 'iterate')
  }
  return out
}

/** Reads a `p:cTn`; time-node children (`childTnLst`, `subTnLst`) are returned for the caller. */
function ctnOf(ctx: Ctx, node: XmlNode): { ctn: CTn; childTn: Array<[string, XmlNode]>; subTn: Array<[string, XmlNode]> } {
  const a = attrs(node)
  for (const name of Object.keys(a)) {
    if (!CTN_ATTRS.has(name)) add(ctx, 'timing-unknown-attribute', `cTn@${name}`, `Unknown attribute cTn@${name}; dropped`)
    else if (CARRIED_CTN.has(name)) add(ctx, 'timing-carried', name, `cTn@${name} is carried but not acted on`)
    else if (IGNORED_CTN.has(name)) add(ctx, 'timing-ignored', name, `cTn@${name} is not applied`)
  }
  let id = value(ctx, a, 'cTn', 'id', uint)
  if (id === undefined) id = ctx.synthetic--
  else if (ctx.ids.has(id)) add(ctx, 'timing-duplicate-id', String(id), `Duplicate cTn id ${id}`)
  let accel = value(ctx, a, 'cTn', 'accel', fraction) ?? 0
  let decel = value(ctx, a, 'cTn', 'decel', fraction) ?? 0
  if (accel < 0 || accel > 1) { add(ctx, 'timing-invalid-value', 'cTn@accel', `accel out of range`); accel = 0 }
  if (decel < 0 || decel > 1) { add(ctx, 'timing-invalid-value', 'cTn@decel', `decel out of range`); decel = 0 }
  if (accel + decel > 1) {
    add(ctx, 'timing-invalid-value', 'cTn@accel+decel', 'accel + decel exceeds 100 %; scaled to fit')
    const sum = accel + decel
    accel /= sum; decel /= sum
  }
  const ctn: CTn = {
    id,
    durMs: value(ctx, a, 'cTn', 'dur', time),
    accel, decel,
    autoRev: value(ctx, a, 'cTn', 'autoRev', bool) ?? false,
    start: [], end: [],
  }
  const set = <K extends keyof CTn>(k: K, v: CTn[K] | undefined) => { if (v !== undefined) ctn[k] = v }
  set('presetClass', value(ctx, a, 'cTn', 'presetClass', v => PRESET_CLASSES.has(v) ? v : undefined))
  set('presetId', value(ctx, a, 'cTn', 'presetID', int))
  set('presetSubtype', value(ctx, a, 'cTn', 'presetSubtype', int))
  set('nodeType', value(ctx, a, 'cTn', 'nodeType', v => NODE_TYPES.has(v) ? v : undefined))
  set('grpId', value(ctx, a, 'cTn', 'grpId', uint))
  set('repeatCount', value(ctx, a, 'cTn', 'repeatCount', v => v === 'indefinite' ? v : uint(v) === undefined ? undefined : uint(v)! / 1000))
  set('repeatDurMs', value(ctx, a, 'cTn', 'repeatDur', time))
  set('fill', value(ctx, a, 'cTn', 'fill', oneOf('hold', 'remove', 'freeze', 'transition')))
  set('restart', value(ctx, a, 'cTn', 'restart', oneOf('always', 'whenNotActive', 'never')))
  set('syncBehavior', value(ctx, a, 'cTn', 'syncBehavior', oneOf('canSlip', 'locked')))
  set('masterRel', value(ctx, a, 'cTn', 'masterRel', oneOf('sameClick', 'lastClick', 'nextClick')))
  set('afterEffect', value(ctx, a, 'cTn', 'afterEffect', bool))
  ctx.ids.set(id, ctn)
  let childTn: Array<[string, XmlNode]> = []
  let subTn: Array<[string, XmlNode]> = []
  for (const [name, c] of kids(node)) {
    switch (name) {
      case 'stCondLst': ctn.start = conditionList(ctx, c); break
      case 'endCondLst': ctn.end = conditionList(ctx, c); break
      case 'endSync': { const cond = condition(ctx, c); if (cond) ctn.endSync = cond; break }
      case 'iterate': ctn.iterate = iterate(ctx, c); break
      case 'childTnLst': childTn = kids(c); break
      case 'subTnLst': subTn = kids(c); break
      default: unknownNode(ctx, name, c, 'cTn')
    }
  }
  return { ctn, childTn, subTn }
}

// ---------------------------------------------------------------------------
// Time nodes
// ---------------------------------------------------------------------------

function enter(ctx: Ctx, depth: number): void {
  if (++ctx.nodes > TIMING_LIMITS.timeNodes) throw new Refused('time-nodes')
  if (depth > TIMING_LIMITS.nesting) throw new Refused('nesting')
}

function timeNode(ctx: Ctx, name: string, node: XmlNode, depth: number): TimeNode | undefined {
  if (CONTAINERS.has(name)) return container(ctx, name as ContainerNode['kind'], node, depth)
  if (BEHAVIOURS.has(name)) return behaviour(ctx, name, node, depth)
  unknownNode(ctx, name, node, 'time node list')
  return undefined
}

function container(ctx: Ctx, kind: ContainerNode['kind'], node: XmlNode, depth: number): ContainerNode | undefined {
  enter(ctx, depth)
  const ctnNode = child(node, 'cTn')
  if (!ctnNode) { add(ctx, 'timing-malformed', kind, `${kind} without cTn; dropped`); return undefined }
  const { ctn, childTn, subTn } = ctnOf(ctx, ctnNode)
  const out: ContainerNode = { kind, ctn, children: [] }
  const a = attrs(node)
  let nextConds: Condition[] = []
  let prevConds: Condition[] = []
  for (const [name, c] of kids(node)) {
    if (name === 'cTn') continue
    if (kind === 'seq' && name === 'nextCondLst') nextConds = conditionList(ctx, c)
    else if (kind === 'seq' && name === 'prevCondLst') prevConds = conditionList(ctx, c)
    else unknownNode(ctx, name, c, kind)
  }
  if (kind === 'seq') {
    out.seq = {
      concurrent: value(ctx, a, 'seq', 'concurrent', bool) ?? false,
      nextAc: value(ctx, a, 'seq', 'nextAc', oneOf('none', 'seek')) ?? 'none',
      prevAc: value(ctx, a, 'seq', 'prevAc', oneOf('none', 'skipTimed')) ?? 'none',
    }
    // The model advances sequences by nodeType: next/previous for the main
    // sequence, the start trigger's clicks for an interactive one. Other
    // next/previous conditions are not represented.
    const triggers = new Set(ctn.start.flatMap(c => c.evt === 'onClick' && c.target && 'spid' in c.target ? [c.target.spid] : []))
    const nextOk = (c: Condition) => ctn.nodeType === 'mainSeq' ? c.evt === 'onNext' : ctn.nodeType === 'interactiveSeq' && c.evt === 'onClick' && !!c.target && 'spid' in c.target && triggers.has(c.target.spid)
    if (!nextConds.every(nextOk)) add(ctx, 'timing-ignored', 'nextCondLst', 'Sequence next conditions other than the main or interactive form are not applied')
    if (!prevConds.every(c => c.evt === 'onPrev')) add(ctx, 'timing-ignored', 'prevCondLst', 'Sequence previous conditions other than onPrev are not applied')
  }
  for (const [name, c] of childTn) {
    const n = timeNode(ctx, name, c, depth + 1)
    if (n) out.children.push(n)
  }
  if (subTn.length && kind !== 'par') add(ctx, 'timing-approximated', 'subTnLst', `subTnLst under ${kind} joins its children`)
  for (const [name, c] of subTn) {
    const n = timeNode(ctx, name, c, depth + 1)
    if (n) out.children.push(n)
  }
  return out
}

/** Resolves a behaviour's target; undefined drops the behaviour. */
function resolveTarget(ctx: Ctx, t: TargetRef | undefined, el: string): TargetRef | undefined {
  if (!t) { add(ctx, 'timing-malformed', el, `${el} without a readable target; dropped`); return undefined }
  if (!('spid' in t)) return t
  const where = ctx.resolve.shape(t.spid)
  if (where === 'missing') { add(ctx, 'unresolved-target', t.spid, `Target shape ${t.spid} is not on the slide; behaviour dropped`); return undefined }
  if (where === 'master-only') { add(ctx, 'master-target', t.spid, `Target shape ${t.spid} is on the layout or master; behaviour dropped`); return undefined }
  if (t.text?.kind === 'p') {
    const count = ctx.resolve.paragraphs(t.spid).length
    if (t.text.end >= count) { add(ctx, 'unresolved-target', `${t.spid}/p:${t.text.st}-${t.text.end}`, `Paragraph range past the shape's ${count} paragraphs; behaviour dropped`); return undefined }
    const refs = ctx.textRefs.get(t.spid) ?? []
    refs.push({ st: t.text.st, end: t.text.end })
    ctx.textRefs.set(t.spid, refs)
  }
  return t
}

const CBHVR_ATTRS = new Set(['additive', 'accumulate', 'xfrmType', 'from', 'to', 'by', 'rctx', 'override'])

function cBhvr(ctx: Ctx, node: XmlNode | undefined, el: string): BehaviourBase | undefined {
  if (!node) { add(ctx, 'timing-malformed', el, `${el} without cBhvr; dropped`); return undefined }
  const a = attrs(node)
  for (const name of Object.keys(a)) {
    if (!CBHVR_ATTRS.has(name)) add(ctx, 'timing-unknown-attribute', `cBhvr@${name}`, `Unknown attribute cBhvr@${name}; dropped`)
    else if (name === 'from' || name === 'to' || name === 'by') add(ctx, 'timing-ignored', `cBhvr@${name}`, `cBhvr@${name} is not applied`)
  }
  let ctn: CTn | undefined
  let tgt: TargetRef | undefined
  const attrNames: string[] = []
  for (const [name, c] of kids(node)) {
    if (name === 'cTn') {
      const r = ctnOf(ctx, c)
      ctn = r.ctn
      for (const [n] of [...r.childTn, ...r.subTn]) add(ctx, 'timing-unknown-container', `${el}/${n}`, `Time nodes under a behaviour's cTn are dropped`)
    } else if (name === 'tgtEl') tgt = target(ctx, c)
    else if (name === 'attrNameLst') {
      for (const [n, x] of kids(c)) {
        if (n !== 'attrName') { unknownNode(ctx, n, x, 'attrNameLst'); continue }
        const v = textOf(x).trim()
        if (ATTRIBUTE_NAMES.has(v)) attrNames.push(v)
        else add(ctx, 'timing-unknown-attribute-name', v, `Unknown animated attribute ${v}; dropped`)
      }
    } else unknownNode(ctx, name, c, 'cBhvr')
  }
  if (!ctn) { add(ctx, 'timing-malformed', el, `${el} without cTn; dropped`); return undefined }
  const resolved = resolveTarget(ctx, tgt, el)
  // A dropped behaviour's id no longer resolves: references to it are diagnosed missing.
  if (!resolved) { if (ctx.ids.get(ctn.id) === ctn) ctx.ids.delete(ctn.id); return undefined }
  const base: BehaviourBase = {
    ctn, target: resolved, attrNames,
    additive: value(ctx, a, 'cBhvr', 'additive', oneOf('base', 'sum', 'repl', 'mult', 'none')) ?? 'base',
    accumulate: value(ctx, a, 'cBhvr', 'accumulate', oneOf('always', 'none')) ?? 'none',
  }
  const override = value(ctx, a, 'cBhvr', 'override', oneOf('normal', 'childStyle'))
  if (override) base.override = override
  const xfrmType = value(ctx, a, 'cBhvr', 'xfrmType', oneOf('pt', 'img'))
  if (xfrmType) base.xfrmType = xfrmType
  if (a.rctx !== undefined) { base.rctx = a.rctx; add(ctx, 'timing-carried', 'rctx', 'cBhvr@rctx is carried but not acted on') }
  return base
}

function mediaBase(ctx: Ctx, node: XmlNode, el: string): BehaviourBase | undefined {
  const media = child(node, 'cMediaNode')
  if (!media) { add(ctx, 'timing-malformed', el, `${el} without cMediaNode; dropped`); return undefined }
  for (const name of Object.keys(attrs(media))) {
    if (['vol', 'mute', 'numSld', 'showWhenStopped'].includes(name)) add(ctx, 'timing-ignored', name, `cMediaNode@${name} is not applied`)
    else add(ctx, 'timing-unknown-attribute', `cMediaNode@${name}`, `Unknown attribute cMediaNode@${name}; dropped`)
  }
  let ctn: CTn | undefined
  let tgt: TargetRef | undefined
  for (const [name, c] of kids(media)) {
    if (name === 'cTn') {
      const r = ctnOf(ctx, c)
      ctn = r.ctn
      for (const [n] of [...r.childTn, ...r.subTn]) add(ctx, 'timing-unknown-container', `${el}/${n}`, `Time nodes under a media node are dropped`)
    } else if (name === 'tgtEl') tgt = target(ctx, c)
    else unknownNode(ctx, name, c, 'cMediaNode')
  }
  if (!ctn) { add(ctx, 'timing-malformed', el, `${el} without cTn; dropped`); return undefined }
  const resolved = resolveTarget(ctx, tgt, el)
  // A dropped behaviour's id no longer resolves: references to it are diagnosed missing.
  if (!resolved) { if (ctx.ids.get(ctn.id) === ctn) ctx.ids.delete(ctn.id); return undefined }
  return { ctn, target: resolved, attrNames: [], additive: 'base', accumulate: 'none' }
}

function behaviour(ctx: Ctx, el: string, node: XmlNode, depth: number): BehaviourNode | undefined {
  enter(ctx, depth)
  const a = attrs(node)
  const known: Record<string, readonly string[]> = {
    set: [], anim: ['by', 'from', 'to', 'calcmode', 'valueType'], animEffect: ['transition', 'filter', 'prLst'],
    animMotion: ['origin', 'path', 'pathEditMode', 'rAng', 'ptsTypes'], animRot: ['by', 'from', 'to'], animScale: ['zoomContents'],
    animClr: ['clrSpc', 'dir'], cmd: ['type', 'cmd'], audio: ['isNarration'], video: ['fullScrn'],
  }
  for (const name of Object.keys(a)) if (!known[el].includes(name)) add(ctx, 'timing-unknown-attribute', `${el}@${name}`, `Unknown attribute ${el}@${name}; dropped`)
  if (el === 'audio' || el === 'video') {
    for (const [name, c] of kids(node)) if (name !== 'cMediaNode') unknownNode(ctx, name, c, el)
    const base = mediaBase(ctx, node, el)
    if (!base) return undefined
    if (el === 'audio') {
      const out: AudioBehaviour = { ...base, kind: 'audio' }
      const n = value(ctx, a, el, 'isNarration', bool)
      if (n !== undefined) out.isNarration = n
      return out
    }
    const out: VideoBehaviour = { ...base, kind: 'video' }
    const f = value(ctx, a, el, 'fullScrn', bool)
    if (f !== undefined) out.fullScrn = f
    return out
  }
  // Element children other than cBhvr, by kind.
  const extra: Record<string, readonly string[]> = {
    set: ['to'], anim: ['tavLst'], animEffect: ['progress'], animMotion: ['by', 'from', 'to', 'rCtr'], animRot: [],
    animScale: ['by', 'from', 'to'], animClr: ['by', 'from', 'to'], cmd: [],
  }
  const parts = new Map<string, XmlNode>()
  for (const [name, c] of kids(node)) {
    if (name === 'cBhvr') continue
    if (extra[el].includes(name) && !parts.has(name)) parts.set(name, c)
    else unknownNode(ctx, name, c, el)
  }
  // Keyframes are counted (and limits applied) before the target can drop the behaviour.
  const keys: Keyframe[] = []
  if (el === 'anim') {
    let count = 0
    for (const [name, tav] of kids(parts.get('tavLst'))) {
      if (name !== 'tav') { unknownNode(ctx, name, tav, 'tavLst'); continue }
      if (++count > TIMING_LIMITS.keyframesPerBehaviour) throw new Refused('keyframes-per-behaviour')
      if (++ctx.tavs > TIMING_LIMITS.keyframesPerSlide) throw new Refused('keyframes')
      const ta = attrs(tav)
      const tm = ta.tm === undefined ? undefined : value(ctx, ta, 'tav', 'tm', keyTime)
      let v: string | number | undefined
      for (const [n, x] of kids(tav)) {
        if (n === 'val') v = variant(ctx, x, 'tav')
        else unknownNode(ctx, n, x, 'tav')
      }
      if (ta.tm !== undefined && tm === undefined) continue
      if (v === undefined && ta.fmla === undefined) { add(ctx, 'timing-malformed', 'tav', 'Keyframe without a value; dropped'); continue }
      const key: Keyframe = { tm: tm ?? 0, value: v ?? '' }
      if (ta.fmla !== undefined) key.fmla = ta.fmla
      keys.push(key)
    }
  }
  let path: MotionPath | undefined
  if (el === 'animMotion' && a.path !== undefined) {
    if (a.path.length > TIMING_LIMITS.motionPathBytes) throw new Refused('motion-path')
    if (!ctx.parsePath) add(ctx, 'motion-path-deferred', 'animMotion', 'Motion paths need the S1-6 parser; path not read')
    else {
      const parsed = ctx.parsePath(a.path)
      if (parsed.ok) path = parsed.path
      else if (parsed.reason === 'too-long' || parsed.reason === 'too-many-commands') throw new Refused('motion-path')
      else add(ctx, 'timing-invalid-value', 'animMotion@path', `Unreadable motion path (${parsed.reason}); dropped`)
    }
  }
  const base = cBhvr(ctx, child(node, 'cBhvr'), el)
  if (!base) return undefined
  const opt = <T extends object, K extends keyof T>(o: T, k: K, v: T[K] | undefined) => { if (v !== undefined) o[k] = v }
  switch (el) {
    case 'set': {
      const out: SetBehaviour = { ...base, kind: 'set' }
      opt(out, 'to', parts.has('to') ? variant(ctx, parts.get('to'), 'set/to') : undefined)
      return out
    }
    case 'anim': {
      const out: AnimBehaviour = { ...base, kind: 'anim', keys }
      opt(out, 'calcmode', value(ctx, a, el, 'calcmode', oneOf('discrete', 'lin', 'fmla')))
      opt(out, 'valueType', value(ctx, a, el, 'valueType', oneOf('str', 'num', 'clr')))
      opt(out, 'by', a.by); opt(out, 'from', a.from); opt(out, 'to', a.to)
      return out
    }
    case 'animEffect': {
      const out: AnimEffectBehaviour = { ...base, kind: 'animEffect', transition: value(ctx, a, el, 'transition', oneOf('in', 'out', 'none')) ?? 'in' }
      opt(out, 'filter', a.filter); opt(out, 'prLst', a.prLst)
      opt(out, 'progress', parts.has('progress') ? variant(ctx, parts.get('progress'), 'animEffect/progress') : undefined)
      return out
    }
    case 'animMotion': {
      const out: AnimMotionBehaviour = { ...base, kind: 'animMotion' }
      if (a.pathEditMode !== undefined) add(ctx, 'timing-ignored', 'pathEditMode', 'animMotion@pathEditMode is not applied')
      opt(out, 'path', path)
      opt(out, 'origin', value(ctx, a, el, 'origin', oneOf('parent', 'layout')))
      opt(out, 'rAngDeg', value(ctx, a, el, 'rAng', degrees))
      opt(out, 'ptsTypes', a.ptsTypes)
      for (const k of ['by', 'from', 'to', 'rCtr'] as const) opt(out, k, point(ctx, parts.get(k), `animMotion/${k}`))
      return out
    }
    case 'animRot': {
      const out: AnimRotBehaviour = { ...base, kind: 'animRot' }
      for (const k of ['by', 'from', 'to'] as const) opt(out, k, value(ctx, a, el, k, degrees))
      return out
    }
    case 'animScale': {
      const out: AnimScaleBehaviour = { ...base, kind: 'animScale' }
      opt(out, 'zoomContents', value(ctx, a, el, 'zoomContents', bool))
      for (const k of ['by', 'from', 'to'] as const) opt(out, k, point(ctx, parts.get(k), `animScale/${k}`))
      return out
    }
    case 'animClr': {
      const out: AnimClrBehaviour = {
        ...base, kind: 'animClr',
        clrSpc: value(ctx, a, el, 'clrSpc', oneOf('rgb', 'hsl')) ?? 'rgb',
        dir: value(ctx, a, el, 'dir', oneOf('cw', 'ccw')) ?? 'cw',
      }
      const by = parts.get('by')
      if (by) {
        const [kind, x] = kids(by)[0] ?? []
        const xa = attrs(x)
        const num = (v: string | undefined, div: number) => { const n = int(v); return n === undefined ? undefined : n / div }
        if (kind === 'hsl' && [xa.h, xa.s, xa.l].every(v => int(v) !== undefined)) out.by = `hsl(${num(xa.h, 60000)},${num(xa.s, 100000)},${num(xa.l, 100000)})`
        else if (kind === 'rgb' && [xa.r, xa.g, xa.b].every(v => int(v) !== undefined)) out.by = `rgb(${num(xa.r, 100000)},${num(xa.g, 100000)},${num(xa.b, 100000)})`
        else add(ctx, 'timing-invalid-value', 'animClr/by', 'Unreadable colour offset; dropped')
      }
      opt(out, 'from', parts.has('from') ? color(ctx, parts.get('from'), 'animClr/from') : undefined)
      opt(out, 'to', parts.has('to') ? color(ctx, parts.get('to'), 'animClr/to') : undefined)
      return out
    }
    case 'cmd': {
      const cmdType = value(ctx, a, el, 'type', oneOf('evt', 'call', 'verb'))
      if (!cmdType || a.cmd === undefined) { add(ctx, 'timing-malformed', 'cmd', 'cmd without type or command; dropped'); ctx.ids.delete(base.ctn.id); return undefined }
      const out: CmdBehaviour = { ...base, kind: 'cmd', cmdType, cmd: a.cmd }
      return out
    }
  }
  return undefined
}

// ---------------------------------------------------------------------------
// Cycles
// ---------------------------------------------------------------------------

/**
 * Breaks `tn` reference cycles. Vertices are each node's begin and end; a start
 * condition on `tn` makes this begin depend on that node's begin or end, an end
 * condition does the same for this end, and every end depends on its own begin.
 * Each round finds one cycle (iterative DFS) and replaces its closing condition
 * with one that never fires; more than 100 rounds refuse the timing.
 */
function breakCycles(ctx: Ctx): void {
  interface Edge { to: string; list?: Condition[]; index?: number; owner?: number }
  const edges = new Map<string, Edge[]>()
  const push = (from: string, e: Edge) => { const l = edges.get(from); if (l) l.push(e); else edges.set(from, [e]) }
  const missing = new Set<number>()
  for (const [id, ctn] of ctx.ids) {
    push(`${id}:e`, { to: `${id}:b` })
    for (const [list, side] of [[ctn.start, 'b'], [ctn.end, 'e']] as const) {
      list.forEach((c, index) => {
        if (c.tn === undefined) return
        if (!ctx.ids.has(c.tn)) { if (!missing.has(c.tn)) { missing.add(c.tn); add(ctx, 'timing-missing-reference', String(c.tn), `Condition references missing time node ${c.tn}`) } return }
        const which = c.evt === 'end' || c.evt === 'onEnd' ? 'e' : c.evt === 'begin' || c.evt === 'onBegin' ? 'b' : undefined
        if (which) push(`${id}:${side}`, { to: `${c.tn}:${which}`, list, index, owner: id })
      })
    }
  }
  for (let round = 0; round <= 100; round++) {
    const cyc = findCycle(edges)
    if (!cyc) return
    if (round === 100) throw new Refused('cycles')
    // The closing-most condition edge on the cycle.
    const e = [...cyc].reverse().find(x => x.list)!
    const c = e.list![e.index!]
    e.list![e.index!] = { delayMs: 'indefinite', ...(c.evt !== undefined ? { evt: c.evt } : {}) }
    const from = [...edges.entries()].find(([, l]) => l.includes(e))!
    from[1].splice(from[1].indexOf(e), 1)
    add(ctx, 'timing-cycle', `tn:${e.owner}`, `Time node ${e.owner} closes a reference cycle; that condition never fires`)
  }

  function findCycle(graph: Map<string, Edge[]>): Edge[] | undefined {
    const state = new Map<string, 1 | 2>()
    for (const startVertex of graph.keys()) {
      if (state.has(startVertex)) continue
      const stack: Array<{ v: string; i: number; via?: Edge }> = [{ v: startVertex, i: 0 }]
      state.set(startVertex, 1)
      while (stack.length) {
        const top = stack[stack.length - 1]
        const out = graph.get(top.v) ?? []
        if (top.i >= out.length) { state.set(top.v, 2); stack.pop(); continue }
        const e = out[top.i++]
        const s = state.get(e.to)
        if (s === 1) {
          const at = stack.findIndex(f => f.v === e.to)
          const cycle = [...stack.slice(at + 1).map(f => f.via!), e]
          if (cycle.some(x => x.list)) return cycle
          continue
        }
        if (s === 2) continue
        state.set(e.to, 1)
        stack.push({ v: e.to, i: 0, via: e })
      }
    }
    return undefined
  }
}

// ---------------------------------------------------------------------------
// Builds
// ---------------------------------------------------------------------------

function buildList(ctx: Ctx, node: XmlNode | undefined): BuildTemplate[] {
  const out: BuildTemplate[] = []
  for (const [name, b] of kids(node)) {
    const a = attrs(b)
    if (!['bldP', 'bldGraphic', 'bldOleChart', 'bldDgm'].includes(name)) { unknownNode(ctx, name, b, 'bldLst'); continue }
    const grpId = value(ctx, a, name, 'grpId', uint)
    if (!a.spid || grpId === undefined) { add(ctx, 'timing-malformed', name, `${name} without spid or grpId; dropped`); continue }
    if (name === 'bldP') {
      for (const k of ['uiExpand', 'autoUpdateAnimBg', 'advAuto']) if (a[k] !== undefined) add(ctx, 'timing-ignored', `bldP@${k}`, `bldP@${k} is not applied`)
      for (const [n, x] of kids(b)) n === 'tmplLst' ? add(ctx, 'timing-ignored', 'tmplLst', 'Build templates (tmplLst) are not applied') : unknownNode(ctx, n, x, 'bldP')
      out.push({
        kind: 'bldP', spid: a.spid, grpId,
        build: value(ctx, a, name, 'build', oneOf('whole', 'p', 'cust', 'allAtOnce')) ?? 'whole',
        bldLvl: value(ctx, a, name, 'bldLvl', uint) ?? 1,
        animBg: value(ctx, a, name, 'animBg', bool) ?? false,
        rev: value(ctx, a, name, 'rev', bool) ?? false,
      })
    } else if (name === 'bldGraphic') {
      const [n, x] = kids(b)[0] ?? []
      let bld: string | undefined
      if (n === 'bldAsOne') bld = 'asOne'
      else if (n === 'bldSub') { const [sn, sx] = kids(x)[0] ?? []; bld = sn ? `${sn}:${attrs(sx).bld ?? ''}` : 'sub' }
      else if (n) unknownNode(ctx, n, x!, 'bldGraphic')
      out.push({ kind: 'bldGraphic', spid: a.spid, grpId, ...(bld ? { bld } : {}) })
    } else {
      out.push({ kind: name as 'bldOleChart' | 'bldDgm', spid: a.spid, grpId, ...(a.bld !== undefined ? { bld: a.bld } : {}) })
    }
  }
  return out
}

export interface ExpandOptions {
  sink?: DiagnosticSink
  slideIndex?: number
  /** Paragraph ranges the timing tree targets, by spid: the units of a `build="cust"` template. */
  referenced?: ReadonlyMap<string, ReadonlyArray<{ st: number; end: number }>>
}

/**
 * Normalises build templates for the render planner (S1-12 contract): range
 * `st..end` are inclusive paragraph indices (as `p:pRg`), keys come from
 * `targetKey`, and `body` is the whole shape's key. A `build="p"` template
 * gives one range per non-empty paragraph whose level is above `bldLvl`
 * (paragraph levels are 0-based, `bldLvl` 1-based); deeper and empty
 * paragraphs join the preceding range. `cust` uses the ranges the timing tree
 * targets. Other builds, and text-less shapes, are `whole`.
 */
export function expandBuilds(build: readonly BuildTemplate[], resolve: TargetResolver, options: ExpandOptions = {}): BuildOwnership[] {
  const out: BuildOwnership[] = []
  for (const b of build) {
    const where = resolve.shape(b.spid)
    if (where !== 'slide') {
      options.sink?.add({ code: where === 'master-only' ? 'master-target' : 'unresolved-target', slideIndex: options.slideIndex, effect: b.spid, message: `Build target ${b.spid} is not on the slide; build dropped` })
      continue
    }
    const body = targetKey({ spid: b.spid })
    const key = (st: number, end: number) => ({ key: targetKey({ spid: b.spid, text: { kind: 'p', st, end } }), st, end })
    const ranges: Array<{ key: string; st: number; end: number }> = []
    if (b.kind === 'bldP' && (b.build === 'p' || b.build === 'cust')) {
      const refs = b.build === 'cust' ? options.referenced?.get(b.spid) : undefined
      if (refs?.length) {
        const seen = new Set<string>()
        for (const r of [...refs].sort((x, y) => x.st - y.st || x.end - y.end)) {
          const k = key(r.st, r.end)
          if (!seen.has(k.key)) { seen.add(k.key); ranges.push(k) }
        }
      } else {
        const paragraphs = resolve.paragraphs(b.spid)
        let st = -1
        paragraphs.forEach((p, i) => {
          if (!p.empty && p.level < b.bldLvl) {
            if (st >= 0) ranges.push(key(st, i - 1))
            st = i
          }
        })
        if (st >= 0) ranges.push(key(st, paragraphs.length - 1))
      }
    }
    out.push({ spid: b.spid, mode: ranges.length ? 'paragraph' : 'whole', ranges, animBg: b.kind === 'bldP' ? b.animBg : false, body })
  }
  return out
}

// ---------------------------------------------------------------------------
// Reader
// ---------------------------------------------------------------------------

/** The slide's `p:timing`, honouring `mc:AlternateContent` with the requirements table. */
function timingElement(root: XmlNode, ctx: Ctx): XmlNode | undefined {
  for (const [name, node] of kids(root)) {
    if (name === 'timing') return node
    if (name === 'AlternateContent') {
      const { branches, rejected } = compatibilityBranches(node)
      const found = branches.map(b => child(b, 'timing')).find(Boolean)
      if (found) {
        for (const prefix of rejected) add(ctx, 'timing-requires', prefix, `mc:Choice requires unsupported namespace prefix ${prefix}; next branch used`)
        return found
      }
    }
  }
  return undefined
}

export function createSlideTimingReader(options: TimingReaderOptions = {}): SlideTimingReader {
  return (slideXml, slideIndex, resolve, sink) => {
    const transition: TransitionSpec | undefined = readTransition(slideXml, slideIndex, sink)
    const ctx: Ctx = { slideIndex, sink, resolve, parsePath: options.parseMotionPath, nodes: 0, conds: 0, tavs: 0, synthetic: -1, ids: new Map(), textRefs: new Map() }
    const withTransition = <T extends object>(r: T) => (transition ? { transition, ...r } : r)
    const timingNode = timingElement(slideXml, ctx)
    if (!timingNode) return withTransition({ timing: undefined, builds: [] })
    try {
      let root: TimeNode | undefined
      let build: BuildTemplate[] = []
      for (const [name, node] of kids(timingNode)) {
        if (name === 'tnLst') {
          for (const [n, x] of kids(node)) {
            if (root === undefined && n === 'par') root = container(ctx, 'par', x, 1)
            else add(ctx, 'timing-unknown-element', `tnLst/${n}`, `Only one root par is read; ${n} dropped`)
          }
        } else if (name === 'bldLst') build = buildList(ctx, node)
        else if (name === 'extLst') add(ctx, 'timing-ignored', 'timing/extLst', 'Timing extensions are not applied')
        else unknownNode(ctx, name, node, 'timing')
      }
      if (!root) return withTransition({ timing: undefined, builds: [] })
      breakCycles(ctx)
      const tree: TimingTree = { root, build }
      return withTransition({ timing: tree, builds: expandBuilds(build, resolve, { sink, slideIndex, referenced: ctx.textRefs }) })
    } catch (e) {
      if (!(e instanceof Refused)) throw e
      sink.add({ code: 'timing-limit', slideIndex, effect: e.why, message: `Timing exceeds the ${e.why} limit; the slide is static` })
      return withTransition({ timing: { refused: `limit:${e.why}` }, builds: [] })
    }
  }
}

/** The default reader. S1-10 builds one with S1-6's motion path parser. */
export const readSlideTiming: SlideTimingReader = createSlideTimingReader()
