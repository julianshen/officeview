/**
 * DOM-free Morph preparation (piece 4 spec §4), shared by the local stage
 * (`player/stage.ts`) and the stream server (bun, `@napi-rs/canvas`).
 *
 * From the outgoing slide as it was left and the incoming one at its live
 * state: the matcher's pairs, the keyed re-plan of both slides (Morph-only
 * text unit splits included), the layers the Morph rasterises for itself
 * (`refs`) and its frame items in z order, each side naming its layer by
 * index into `refs`.
 *
 * Nothing here touches `document`, `window` or a compositor: canvases come
 * from `MorphEnv.canvas`, and rasterising and uploading the refs is the
 * caller's (`MorphEnv.raster`, whose outcome feeds back into the plan: a
 * layer that paints nothing drops its pair, and workspace pressure coarsens
 * the text units).
 */
import type { FontResolver } from '../../core/fonts/register'
import type { PptxShape, PptxSlide } from '../../pptx/types'
import type { BaseGeometry } from '../engine/behaviours'
import { matchMorph, type MorphOption, type MorphPair, type MorphRule } from '../morph/match'
import type { MorphFrameItem, MorphLayerFrame } from '../morph/render'
import { mul, stateMatrix } from '../player/hit-test'
import { textBases } from '../player/text-builds'
import { buildRenderPlan, shapeKey, type ColourPart, type PlanLayer, type RenderPlan } from '../render/plan'
import { slideSize } from '../render/scene'
import { TEXT_BOX_BYTES, textItems, type TextItem } from '../render/text-split'
import type { Affine, AllocationOwner, CanvasFactory, DiagnosticSink, Lease, RenderState, TargetKey, TargetState, TransitionSpec } from '../types'
import { rasterLayers, type PreparedSlide } from './slide'

const IDENTITY: Affine = [1, 0, 0, 1, 0, 0]
export const PLAIN: TargetState = Object.freeze({ visible: true, opacity: 1, transform: IDENTITY })

/** The group state composed over a sub-layer's own. */
export function compose(parent: TargetState | undefined, own: TargetState | undefined): TargetState {
  if (!parent) return own ?? PLAIN
  const p = stateMatrix(parent), o = stateMatrix(own)
  return {
    visible: parent.visible && (own?.visible ?? true),
    opacity: parent.opacity * (own?.opacity ?? 1),
    transform: mul(p, o),
    ...(own?.mask ? { mask: own.mask, ...(own.maskBox ? { maskBox: own.maskBox } : {}), ...(own.maskTransform ? { maskTransform: mul(p, own.maskTransform) } : {}) } : {}),
    ...(own?.color ?? parent.color ? { color: own?.color ?? parent.color } : {}),
  }
}

/** One layer the Morph rasterises for itself: a plan layer of the outgoing or incoming slide, at `scale`. */
export interface MorphLayerRef { side: 'from' | 'to'; layer: PlanLayer; scale: number }

type SideRef = Omit<MorphLayerFrame, 'handle'> & { ref: number }

/** A `MorphFrameItem` whose sides name their layer by index into `MorphPlan.refs`. */
export interface MorphItemRef extends Omit<MorphFrameItem, 'from' | 'to'> { from?: SideRef; to?: SideRef }

export interface MorphPlan {
  /** Every layer that painted (rasterised by the caller), outgoing then incoming, each in paint order. */
  refs: MorphLayerRef[]
  /** Items bottom-most first (z order). */
  items: MorphItemRef[]
  /** Refs drawn unchanged on top (the incoming slide's watermark). */
  over: number[]
  /** Workspace charges held while the Morph runs (the Morph-only text unit records): the caller releases them. */
  leases: Lease[]
  /** The matcher's pairs (the parity harness compares them, §8). */
  pairs: Array<{ fromKey: TargetKey; toKey: TargetKey; rule: MorphRule }>
  /** Additive to the brief: each item's kind (`morph:background`, `morph:out:…`, `morph:in:…`, `morph:pair:…`), as `inspect().lastDraws`. */
  kinds: string[]
}

export interface MorphEnv {
  sink: DiagnosticSink
  workspace: AllocationOwner
  canvas: CanvasFactory
  resolveFont: FontResolver
  /** The most layers one Morph rasterises for itself (both slides). */
  maxLayers: number
  /** Additive to the brief: device pixels per slide px of every ref. */
  scale: number
  /** Additive to the brief: whether the incoming slide's watermark layer rasterises (a mark is configured). */
  watermark: boolean
  /**
   * Additive to the brief: rasterises `refs` (in order) and returns the indices that painted, or `'pressure'` when
   * they did not fit (having released whatever it rasterised). Called again after a coarsening re-plan, with the
   * new refs. A throw releases the plan's leases and propagates. Absent: every ref paints.
   */
  raster?: (refs: readonly MorphLayerRef[]) => ReadonlySet<number> | 'pressure'
}

/** The catalogue's Morph option from `p159:morph`'s `option` attribute. */
export function morphOptionOf(spec: TransitionSpec): MorphOption {
  const o = spec.attrs.option
  return o === 'byWord' || o === 'byChar' ? o : 'byObject'
}

/** The slide's top-level objects by key: what the matcher pairs (first occurrence of a duplicate id). */
export function topShapes(slide: PptxSlide): Map<TargetKey, PptxShape> {
  const out = new Map<TargetKey, PptxShape>()
  for (const shape of slide.shapes) {
    const k = shapeKey(shape)
    if (k !== undefined && !out.has(k)) out.set(k, shape)
  }
  return out
}

const flatText = (shape: PptxShape): string => (shape.textBody?.paragraphs ?? []).map(p => p.runs.map(r => r.text ?? '').join('')).join('\n')

/**
 * How a matched pair blends: `under` (the same outline and content: the
 * outgoing side holds under the incoming one, so colour interpolates and
 * an opaque shape never dips), `split` (the same outline, changed text:
 * fill and line blend `under`, the text crossfades) or `cross` (the
 * outline, picture or kind differs: the content crossfades).
 */
export function blendOf(a: PptxShape, b: PptxShape): 'under' | 'split' | 'cross' {
  const plain = (s: PptxShape) => !s.group && !s.children?.length && !s.table && !s.content && !s.chartFallback
  if (!plain(a) || !plain(b)) return 'cross'
  if (a.geometry !== b.geometry || (a.presetName ?? '') !== (b.presetName ?? '') || (a.geometry === 'other' && !a.presetName)) return 'cross'
  if (a.imageIndex !== b.imageIndex) return 'cross'
  return flatText(a) === flatText(b) ? 'under' : 'split'
}

/** The top-level object a text unit key (`sp:<spid>/c:…`, `/p:…`) belongs to; undefined for an object key. */
export const containerOf = (k: TargetKey): TargetKey | undefined => {
  const i = k.indexOf('/')
  return i < 0 ? undefined : k.slice(0, i)
}

/** Whether the Morph may split `shape` into its own text units: not when an animation already splits it (its layers carry the build's state). */
export const unitSplittable = (q: PreparedSlide, key: TargetKey, shape: PptxShape): boolean =>
  !q.splits?.has(key) && !q.meta.builds.some(b => b.mode === 'paragraph' && b.ranges.length > 0 && b.spid === shape.source?.id)

/** One drawn layer of a Morph side: its ref, state, plan kind, target key and top-level object. */
interface Entry { ref: number; state: TargetState; part?: 'fill' | 'line' | 'text'; kind: string; key?: TargetKey; root?: TargetKey; background?: boolean }

/** Every drawn layer of `layers` at `state`, in paint order; invisible targets and layers without a ref are left out. */
function entriesOf(layers: readonly PlanLayer[], refs: ReadonlyMap<PlanLayer, number>, state: RenderState): Entry[] {
  const out: Entry[] = []
  const walk = (layers: readonly PlanLayer[], parent: TargetState | undefined, root: TargetKey | undefined) => {
    for (const l of layers) {
      if (l.kind === 'watermark') {
        const r = refs.get(l)
        if (r !== undefined) out.push({ ref: r, state: PLAIN, kind: 'watermark' })
        continue
      }
      if (l.kind === 'static') {
        const r = refs.get(l)
        if (r !== undefined) out.push({ ref: r, state: compose(parent, undefined), kind: 'static', ...(root ? { root } : {}), ...(l.background ? { background: true } : {}) })
        continue
      }
      const s = compose(parent, state.get(l.key))
      if (!s.visible) continue
      if (l.sub) { walk(l.sub, s, root ?? l.key); continue }
      const r = refs.get(l)
      if (r === undefined) continue
      // A text range (a word, letter or paragraph clone) is text only: its colour is the text channel's.
      const part = l.part === 'fill' || l.part === 'line' || l.part === 'text' ? l.part : l.part === 'range' && s.color?.text ? 'text' : undefined
      out.push({ ref: r, state: s, ...(part ? { part } : {}), kind: `target:${l.key}${l.part ? `:${l.part}` : ''}`, key: l.key, root: root ?? l.key })
    }
  }
  walk(layers, undefined, undefined)
  return out
}

/**
 * The Morph's own layers and frame items. Both slides are re-planned with
 * every top-level object keyed (animated or not), rasterised by the caller
 * (`env.raster`) and drawn at their states: the outgoing slide as it was
 * left, the incoming at `live`. Items go in z order: the backgrounds
 * (outgoing under incoming), the outgoing objects that fade out, then the
 * incoming slide's layers in its own order (matched pairs at the incoming
 * object's z, fade-ins), and the incoming watermark on top (`over`).
 *
 * Unit pairs (`byWord`/`byChar`, PR #20 review): each paired container
 * that owns a matched unit is split into its word or letter items (the
 * text-build planner's `splits`, Morph-only), laid out for their own base
 * boxes. A matched unit then morphs on its own path (summed alpha-correctly
 * with its partner, the content being equal); an unmatched unit fades while moving with its
 * container; the container's body (geometry, fill, line, text in no item)
 * pairs as any object does. A shape an animation already splits keeps its
 * own layers (its units move with it). When the split slides would pass
 * `env.maxLayers`, or the workspace refuses the items' records or
 * measuring, or their layers do not fit, the units coarsen to whole objects
 * (`morph-units-coarsened`).
 *
 * A string is why the Morph cannot run (the caller plays the fade look-alike).
 */
export function planMorph(from: { ps: PreparedSlide; state: RenderState }, to: PreparedSlide, live: RenderState, spec: TransitionSpec, env: MorphEnv): MorphPlan | string {
  const { sink, workspace, resolveFont, maxLayers } = env
  const p = to
  const option = morphOptionOf(spec)
  const match = matchMorph(from.ps.slide, p.slide, option, sink)
  const fromShapes = topShapes(from.ps.slide)
  const toShapes = topShapes(p.slide)
  const pairOf = new Map<TargetKey, TargetKey>()
  const blend = new Map<TargetKey, 'under' | 'split' | 'cross'>()
  const fromSplit = new Set<TargetKey>()
  const toSplit = new Set<TargetKey>()
  const unitPairs: MorphPair[] = []
  for (const m of match.pairs) {
    if (containerOf(m.from) !== undefined || containerOf(m.to) !== undefined) { unitPairs.push(m); continue }
    const a = fromShapes.get(m.from), b = toShapes.get(m.to)
    if (!a || !b || pairOf.has(m.to)) continue
    pairOf.set(m.to, m.from)
    const how = blendOf(a, b)
    blend.set(m.to, how)
    if (how === 'split') { fromSplit.add(m.from); toSplit.add(m.to) }
  }

  // Morph-only unit splits: the paired containers that own a matched unit, both sides splittable.
  const unit = option === 'byWord' ? 'wd' : 'lt'
  const fromUnits = new Map<TargetKey, readonly TextItem[]>()
  const toUnits = new Map<TargetKey, readonly TextItem[]>()
  if (option !== 'byObject') {
    for (const m of unitPairs) {
      const fc = containerOf(m.from), tc = containerOf(m.to)
      if (fc === undefined || tc === undefined || pairOf.get(tc) !== fc) continue
      if (fromUnits.has(fc) && toUnits.has(tc)) continue
      const a = fromShapes.get(fc), b = toShapes.get(tc)
      if (!a || !b || !unitSplittable(from.ps, fc, a) || !unitSplittable(p, tc, b)) continue
      // The matcher's own item keys (it split with the same defaults and diagnosed already).
      fromUnits.set(fc, textItems(a, unit))
      toUnits.set(tc, textItems(b, unit))
    }
  }
  const coarsen = (why: string) => {
    sink.add({ code: 'morph-units-coarsened', slideIndex: p.index, effect: 'sld', message: `Morph text units match as whole objects: ${why}` })
    fromUnits.clear()
    toUnits.clear()
  }

  const planOf = (q: PreparedSlide, shapes: Map<TargetKey, PptxShape>, split: Set<TargetKey>, units: ReadonlyMap<TargetKey, readonly TextItem[]>): RenderPlan => {
    const parts = new Map(q.parts)
    for (const k of split) if (!parts.has(k)) parts.set(k, new Set<ColourPart>(['fill', 'line']))
    const splits = units.size ? new Map([...(q.splits ?? []), ...units]) : q.splits
    return buildRenderPlan(q.slide, new Set([...q.animated, ...shapes.keys()]), parts, q.meta.builds, splits)
  }
  const plans = () => {
    const fromPlan = planOf(from.ps, fromShapes, fromSplit, fromUnits)
    const toPlan = planOf(p, toShapes, toSplit, toUnits)
    // The outgoing watermark never draws: the incoming one draws unchanged on top.
    return { fromPlan, toPlan, fromLayers: rasterLayers(fromPlan, false), toLayers: rasterLayers(toPlan, env.watermark) }
  }
  let planned = plans()
  const leases: Lease[] = []
  const fromGeo = new Map<TargetKey, BaseGeometry>()
  const toGeo = new Map<TargetKey, BaseGeometry>()
  if (fromUnits.size) {
    const count = planned.fromLayers.length + planned.toLayers.length
    const items = [...fromUnits.values(), ...toUnits.values()].reduce((n, xs) => n + xs.length, 0)
    const records = count <= maxLayers ? workspace.lease(items * TEXT_BOX_BYTES, `morph-units:${p.index}`) : undefined
    const measuring = records && workspace.lease(4, `morph-text-bases:${p.index}`)
    if (count > maxLayers) coarsen(`${count} layers would exceed the ${maxLayers} layer bound`)
    else if (!records || !measuring) { records?.release(); coarsen('the slideshow workspace refused their records') }
    else {
      leases.push(records)
      try {
        // Item geometry with the real fonts, as text builds lay theirs out (S2-8).
        const ctx = env.canvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D | null
        if (ctx) {
          for (const [k, g] of textBases(from.ps.slide, fromUnits, from.ps.base, ctx, resolveFont, true)) fromGeo.set(k, g)
          for (const [k, g] of textBases(p.slide, toUnits, p.base, ctx, resolveFont, true)) toGeo.set(k, g)
        } else coarsen('no context to measure them')
      } finally {
        measuring.release()
      }
    }
    if (!fromUnits.size) planned = plans()
  }
  const unleash = () => { for (const l of leases.splice(0)) l.release() }
  let refs: MorphLayerRef[]
  let drawn: ReadonlySet<number>
  for (;;) {
    const count = planned.fromLayers.length + planned.toLayers.length
    if (count > maxLayers) { unleash(); return `${count} layers exceed the ${maxLayers} layer bound` }
    refs = [
      ...planned.fromLayers.map(layer => ({ side: 'from' as const, layer, scale: env.scale })),
      ...planned.toLayers.map(layer => ({ side: 'to' as const, layer, scale: env.scale })),
    ]
    let got: ReadonlySet<number> | 'pressure'
    try {
      got = env.raster ? env.raster(refs) : new Set(refs.keys())
    } catch (e) {
      unleash()
      throw e
    }
    if (got !== 'pressure') { drawn = got; break }
    unleash()
    if (!fromUnits.size) return 'its layers do not fit the slideshow workspace'
    // The unit layers did not fit: the containers morph whole instead (the closest effect).
    coarsen('their layers do not fit the slideshow workspace')
    fromGeo.clear()
    toGeo.clear()
    planned = plans()
  }
  const { fromPlan, toPlan } = planned
  // Only the layers that painted have a ref in the items (renumbered densely, in raster order).
  const kept: MorphLayerRef[] = []
  const fromRef = new Map<PlanLayer, number>(), toRef = new Map<PlanLayer, number>()
  refs.forEach((r, i) => {
    if (!drawn.has(i)) return
    ;(r.side === 'from' ? fromRef : toRef).set(r.layer, kept.length)
    kept.push(r)
  })
  // A Morph-only unit has no state of its own: it shows, hides and moves as its container does.
  const unitOwner = new Map<TargetKey, TargetKey>()
  for (const units of [fromUnits, toUnits]) for (const [k, items] of units) for (const it of items) unitOwner.set(it.key, k)
  const withUnits = (state: RenderState, units: ReadonlyMap<TargetKey, readonly TextItem[]>): RenderState => {
    if (!units.size) return state
    const out = new Map(state)
    for (const [k, items] of units) {
      const s = state.get(k)
      if (s) for (const it of items) if (!out.has(it.key)) out.set(it.key, s)
    }
    return out
  }
  const fe = entriesOf(fromPlan.layers, fromRef, withUnits(from.state, fromUnits))
  const te = entriesOf(toPlan.layers, toRef, withUnits(live, toUnits))
  const box = (q: PreparedSlide): BaseGeometry => ({ x: 0, y: 0, ...slideSize(q.slide), rot: 0 })
  const baseOf = (q: PreparedSlide, k: TargetKey): BaseGeometry | undefined => (q === p ? toGeo : fromGeo).get(k) ?? q.base.get(k)
  const side = (q: PreparedSlide, e: Entry, geometry?: BaseGeometry): SideRef => ({
    ref: e.ref, geometry: geometry ?? (e.root ? baseOf(q, e.root) : undefined) ?? box(q), state: e.state,
    ...(e.part ? { part: e.part } : {}), ...(e.root ? { key: e.root } : {}),
  })
  // Matched units of split containers pair like same-outline objects (summed: their content is equal).
  for (const m of unitPairs) {
    const fc = containerOf(m.from), tc = containerOf(m.to)
    if (fc === undefined || tc === undefined || !fromUnits.has(fc) || !toUnits.has(tc) || pairOf.has(m.to)) continue
    pairOf.set(m.to, m.from)
    blend.set(m.to, 'under')
  }
  // Pairs whose two objects both draw (an exited or not-yet-entered object fades instead).
  const fromRoots = new Set(fe.flatMap(e => e.root ? [e.root] : []))
  const toRoots = new Set(te.flatMap(e => e.root ? [e.root] : []))
  for (const [t, f] of [...pairOf]) if (!toRoots.has(t) || !fromRoots.has(f)) pairOf.delete(t)
  const pairedFrom = new Set(pairOf.values())
  const containerTo = new Map<TargetKey, TargetKey>()
  for (const [t, f] of pairOf) if (containerOf(t) === undefined) containerTo.set(f, t)
  /** The container a text layer rides with: a Morph-only unit's, or (PR #20 review 4) a text build's range or item layer's, by its key. */
  const ownerOf = (q: PreparedSlide, root: TargetKey | undefined): TargetKey | undefined => {
    if (root === undefined) return undefined
    const own = unitOwner.get(root)
    if (own !== undefined) return own
    const c = containerOf(root)
    return c !== undefined && (q === p ? toShapes : fromShapes).has(c) ? c : undefined
  }
  // A text build's layer on both sides (the same range or item of a paired container) pairs, riding the
  // container's path with its own animation state; Morph-only units pair only as the matcher said.
  const ride = new Map<TargetKey, { f: TargetKey; t: TargetKey }>()
  for (const r of toRoots) {
    const t = containerOf(r)
    if (t === undefined || unitOwner.has(r) || pairOf.has(r) || !toShapes.has(t)) continue
    const f = pairOf.get(t)
    if (f === undefined || containerOf(f) !== undefined) continue
    const fr = f + r.slice(t.length)
    if (!fromRoots.has(fr) || unitOwner.has(fr) || pairedFrom.has(fr)) continue
    pairOf.set(r, fr)
    pairedFrom.add(fr)
    ride.set(r, { f, t })
    blend.set(r, blend.get(t) === 'under' ? 'under' : 'cross')
  }
  /** An unmatched unit or text-build layer of a paired container fades while moving with it. */
  const carried = (q: PreparedSlide, e: Entry): Pick<MorphItemRef, 'path' | 'pathStates'> & { geometry?: BaseGeometry } => {
    const owner = ownerOf(q, e.root)
    if (owner === undefined) return {}
    const f = q === p ? pairOf.get(owner) : owner
    const t = q === p ? owner : containerTo.get(owner)
    if (f === undefined || t === undefined || pairOf.get(t) !== f) return {}
    const fg = baseOf(from.ps, f) ?? box(from.ps), tg = baseOf(p, t) ?? box(p)
    // The containers' drawn states render the path's ends (PR #20 review 3: a held animation moves the box).
    const fs = fe.find(x => x.root === f)?.state, ts = te.find(x => x.root === t)?.state
    return { path: { from: fg, to: tg }, geometry: q === p ? tg : fg, pathStates: { ...(fs ? { from: fs } : {}), ...(ts ? { to: ts } : {}) } }
  }
  const items: MorphItemRef[] = []
  const kinds: string[] = []
  const push = (item: MorphItemRef, kind: string) => { items.push(item); kinds.push(kind) }
  const fbg = fe.find(e => e.background), tbg = te.find(e => e.background)
  if (fbg || tbg) push({ ...(fbg ? { from: side(from.ps, fbg) } : {}), ...(tbg ? { to: side(p, tbg) } : {}), under: true, background: true }, 'morph:background')
  for (const e of fe) {
    if (e.background || e.kind === 'watermark' || (e.root && pairedFrom.has(e.root))) continue
    const { path, geometry, pathStates } = carried(from.ps, e)
    push({ from: side(from.ps, e, geometry), ...(path ? { path, pathStates } : {}) }, `morph:out:${e.root ?? 'static'}`)
  }
  const done = new Set<TargetKey>()
  for (const e of te) {
    if (e.background || e.kind === 'watermark') continue
    const f = e.root ? pairOf.get(e.root) : undefined
    if (!e.root || f === undefined) {
      const { path, geometry, pathStates } = carried(p, e)
      push({ to: side(p, e, geometry), ...(path ? { path, pathStates } : {}) }, `morph:in:${e.root ?? 'static'}`)
      continue
    }
    if (done.has(e.root)) continue
    done.add(e.root)
    const t = e.root
    const ts = te.filter(x => x.root === t), fs = fe.filter(x => x.root === f)
    // A text-build layer pair rides its containers' boxes (its layers paint where the container does).
    const rides = ride.get(t)
    const fg = baseOf(from.ps, rides?.f ?? f) ?? box(from.ps), tg = baseOf(p, rides?.t ?? t) ?? box(p)
    const how = blend.get(t)
    const pairUp = (a: Entry | undefined, b: Entry | undefined, under: boolean) => {
      if (!a && !b) return
      push({
        ...(a ? { from: side(from.ps, a, rides ? fg : undefined) } : {}), ...(b ? { to: side(p, b, rides ? tg : undefined) } : {}),
        ...(under ? { under: true } : {}), path: { from: fg, to: tg },
      }, `morph:pair:${t}`)
    }
    if (how === 'split') {
      for (const part of ['fill', 'line', 'text'] as const) pairUp(fs.find(x => x.part === part), ts.find(x => x.part === part), part !== 'text')
      // Layers of other kinds (none expected) cross over too.
      const rest = (xs: Entry[]) => xs.filter(x => !x.part)
      const rf = rest(fs), rt = rest(ts)
      for (let i = 0; i < Math.max(rf.length, rt.length); i++) pairUp(rf[i], rt[i], false)
    } else {
      // Layers of one object (sub-layers, text-build items) pair in reading order.
      for (let i = 0; i < Math.max(fs.length, ts.length); i++) pairUp(fs[i], ts[i], how === 'under')
    }
  }
  const over = te.filter(e => e.kind === 'watermark').map(e => e.ref)
  const pairs = match.pairs.map(m => ({ fromKey: m.from, toKey: m.to, rule: m.rule }))
  return { refs: kept, items, over, leases, pairs, kinds }
}
