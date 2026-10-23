/**
 * Morph object matching (slideshow spec 6.3): which objects of two slides
 * pair into one morphed target, and which fade out/in.
 *
 * Microsoft's "Morph transition: Tips and tricks" publishes the `!!` naming
 * scheme: objects renamed to a common name starting with `!!` are treated as
 * the same object (even across kinds here), a `!!` object never pairs with a
 * non-`!!` object, and Morph expects a 1:1 mapping. It also publishes that
 * names without `!!` never influence matching, and that charts cross-fade
 * rather than morph. The remaining rules below are the staged plan's
 * approximation, recorded in the task report: same name and kind, then same
 * text, then same preset geometry within 25% of the slide diagonal.
 *
 * Bounds: slides beyond `MORPH_MAX_OBJECTS` refuse and diagnose
 * (`morph-refused`); more text unit pairs than `MORPH_MAX_UNIT_PAIRS` coarsen
 * the remaining text pairs back to object level (`morph-units-coarsened`).
 * Text units reuse the engine's word/letter item keys (`textItems`), so the
 * transition consumes pairs and builds verbatim.
 */
import type { PptxShape, PptxSlide } from '../../pptx/types'
import { textItems } from '../render/text-split'
import { targetKey } from '../target'
import type { DiagnosticSink, TargetKey } from '../types'

/** The most objects one slide contributes to a morph match (refusal beyond). */
export const MORPH_MAX_OBJECTS = 256
/** The most text unit pairs one match emits; further pairs match as objects. */
export const MORPH_MAX_UNIT_PAIRS = 2048

export type MorphRule = '!!' | 'name-kind' | 'text' | 'geometry'

export interface MorphPair {
  from: TargetKey
  to: TargetKey
  rule: MorphRule
}

export interface MorphMatch {
  pairs: MorphPair[]
  fadeOut: TargetKey[]
  fadeIn: TargetKey[]
}

export type MorphOption = 'byObject' | 'byWord' | 'byChar'

interface Entry {
  key: TargetKey
  shape: PptxShape
  name: string
  kind: string
  text: string
  geometry: string
  cx: number
  cy: number
  matched: boolean
}

interface Ctx {
  pairs: MorphPair[]
  unitPairs: number
  coarsened: boolean
  sink?: DiagnosticSink
  slideIndex: number
  /** Text unit granularity, or undefined under byObject. */
  unit?: 'wd' | 'lt'
}

/** The object's flat text: run texts joined within paragraphs, `\n` between. */
function flatText(shape: PptxShape): string {
  const body = shape.textBody
  if (!body) return ''
  let out = ''
  for (const p of body.paragraphs) {
    if (out) out += '\n'
    for (const r of p.runs) out += r.text ?? ''
  }
  return out
}

/** Coarse object class for name-kind matching (Microsoft's taxonomy is unpublished). */
function kindOf(shape: PptxShape): string {
  if (shape.children?.length || shape.group) return 'group'
  if (shape.image) return 'picture'
  if (shape.table || shape.content || shape.chartFallback) return 'graphic'
  if (shape.textBody?.paragraphs.some(p => p.runs.some(r => (r.text ?? '').length > 0))) return 'text'
  return 'shape'
}

/** Preset geometry bucket: the geometry class, or the preset name for other/absent ones. */
function geometryOf(shape: PptxShape): string {
  const g = shape.geometry
  if (g && g !== 'other') return g
  return `:${shape.presetName ?? ''}`
}

function collect(slide: PptxSlide): Entry[] {
  const out: Entry[] = []
  for (const shape of slide.shapes) {
    const id = shape.source?.id
    if (id === undefined) continue
    out.push({
      key: targetKey({ spid: id }),
      shape,
      name: shape.source?.name ?? '',
      kind: kindOf(shape),
      text: flatText(shape),
      geometry: geometryOf(shape),
      cx: shape.xEmu + shape.widthEmu / 2,
      cy: shape.yEmu + shape.heightEmu / 2,
      matched: false,
    })
  }
  return out
}

function push(map: Map<string, Entry[]>, key: string, e: Entry): void {
  const q = map.get(key)
  if (q) q.push(e)
  else map.set(key, [e])
}

function queueBy(entries: readonly Entry[], key: (e: Entry) => string, only?: (e: Entry) => boolean): Map<string, Entry[]> {
  const out = new Map<string, Entry[]>()
  for (const e of entries) if (!e.matched && (!only || only(e))) push(out, key(e), e)
  return out
}

/** Pairs equal-bucket entries in reading order, one-to-one. */
function pairQueues(ctx: Ctx, fromQ: Map<string, Entry[]>, toQ: Map<string, Entry[]>, rule: MorphRule): void {
  for (const [bucket, fq] of fromQ) {
    const tq = toQ.get(bucket)
    if (!tq) continue
    const n = Math.min(fq.length, tq.length)
    for (let i = 0; i < n; i++) {
      if (fq[i].matched || tq[i].matched) continue
      fq[i].matched = true
      tq[i].matched = true
      ctx.pairs.push({ from: fq[i].key, to: tq[i].key, rule })
    }
  }
}

/** Same flat text: whole-object pairs under byObject, per-unit pairs otherwise. !! objects never reach this. */
function pairByText(ctx: Ctx, fromEntries: readonly Entry[], toEntries: readonly Entry[]): void {
  const untagged = (e: Entry): boolean => !e.name.startsWith('!!')
  const fromQ = queueBy(fromEntries, e => e.text, e => e.text !== '' && untagged(e))
  const toQ = queueBy(toEntries, e => e.text, e => e.text !== '' && untagged(e))
  for (const [text, fq] of fromQ) {
    const tq = toQ.get(text)
    if (!tq) continue
    const n = Math.min(fq.length, tq.length)
    for (let i = 0; i < n; i++) {
      const a = fq[i]
      const b = tq[i]
      if (a.matched || b.matched) continue
      if (!ctx.unit) {
        add(ctx, a, b, 'text')
        continue
      }
      const fa = textItems(a.shape, ctx.unit, { sink: ctx.sink, slideIndex: ctx.slideIndex })
      const fb = textItems(b.shape, ctx.unit, { sink: ctx.sink, slideIndex: ctx.slideIndex })
      // Equal text yields identical text-item sequences (equal lengths), so
      // unit k pairs with unit k; otherwise the pair falls back to the
      // whole object. Beyond the unit-pair cap, further pairs coarsen.
      if (fa.length > 0 && fa.length === fb.length && ctx.unitPairs + fa.length <= MORPH_MAX_UNIT_PAIRS) {
        a.matched = true
        b.matched = true
        ctx.unitPairs += fa.length
        for (let k = 0; k < fa.length; k++) ctx.pairs.push({ from: fa[k].key, to: fb[k].key, rule: 'text' })
      } else {
        if (fa.length > 0 && fa.length === fb.length && !ctx.coarsened) {
          ctx.coarsened = true
          ctx.sink?.add({
            code: 'morph-units-coarsened', slideIndex: ctx.slideIndex, effect: 'sld',
            message: `Morph text units beyond ${MORPH_MAX_UNIT_PAIRS} pairs match as objects`,
          })
        }
        add(ctx, a, b, 'text')
      }
    }
  }
}

/** Same preset geometry, nearer than a quarter of the slide diagonal; greedy nearest in reading order. */
function pairByGeometry(ctx: Ctx, fromEntries: readonly Entry[], toEntries: readonly Entry[], threshold: number): void {
  const fromQ = queueBy(fromEntries, e => e.geometry, e => !e.name.startsWith('!!'))
  const toQ = queueBy(toEntries, e => e.geometry, e => !e.name.startsWith('!!'))
  for (const [bucket, fq] of fromQ) {
    const tq = toQ.get(bucket)
    if (!tq) continue
    for (const a of fq) {
      if (a.matched) continue
      let best: Entry | undefined
      let bestD = Infinity
      for (const b of tq) {
        if (b.matched) continue
        const d = Math.hypot(a.cx - b.cx, a.cy - b.cy)
        if (d <= threshold && d < bestD) {
          best = b
          bestD = d
        }
      }
      if (best) add(ctx, a, best, 'geometry')
    }
  }
}

function add(ctx: Ctx, a: Entry, b: Entry, rule: MorphRule): void {
  a.matched = true
  b.matched = true
  ctx.pairs.push({ from: a.key, to: b.key, rule })
}

export function matchMorph(
  from: PptxSlide,
  to: PptxSlide,
  option: MorphOption,
  sink?: DiagnosticSink,
): MorphMatch {
  const refuse = (why: string): MorphMatch => {
    sink?.add({
      code: 'morph-refused', slideIndex: to.index, effect: 'sld',
      message: `Morph matching refused: ${why}; every object fades and the transition crossfades`,
    })
    return { pairs: [], fadeOut: [], fadeIn: [] }
  }
  if (from.shapes.length > MORPH_MAX_OBJECTS || to.shapes.length > MORPH_MAX_OBJECTS) {
    return refuse(`${from.shapes.length} and ${to.shapes.length} objects exceed the ${MORPH_MAX_OBJECTS} object bound`)
  }

  const ctx: Ctx = {
    pairs: [],
    unitPairs: 0,
    coarsened: false,
    sink,
    slideIndex: to.index,
    unit: option === 'byWord' ? 'wd' : option === 'byChar' ? 'lt' : undefined,
  }
  const fromEntries = collect(from)
  const toEntries = collect(to)

  // Rule 1: `!!` names, exactly as Microsoft publishes. Marked pairQueues
  // never run again on matched entries, so a !! object cannot pair with a
  // non-!! object below.
  pairQueues(
    ctx,
    queueBy(fromEntries, e => e.name, e => e.name.startsWith('!!')),
    queueBy(toEntries, e => e.name, e => e.name.startsWith('!!')),
    '!!',
  )

  // Rule 2: same untagged name and kind.
  pairQueues(
    ctx,
    queueBy(fromEntries, e => `${e.name}\u0000${e.kind}`, e => e.name !== '' && !e.name.startsWith('!!')),
    queueBy(toEntries, e => `${e.name}\u0000${e.kind}`, e => e.name !== '' && !e.name.startsWith('!!')),
    'name-kind',
  )

  // Rule 3: same flat text (published detail absent; this is the plan's rule).
  pairByText(ctx, fromEntries, toEntries)

  // Rule 4: same preset geometry near a quarter of the slide diagonal.
  const threshold = 0.25 * Math.hypot(Math.max(from.widthEmu, to.widthEmu), Math.max(from.heightEmu, to.heightEmu))
  pairByGeometry(ctx, fromEntries, toEntries, threshold)

  return {
    pairs: ctx.pairs,
    fadeOut: fromEntries.filter(e => !e.matched).map(e => e.key),
    fadeIn: toEntries.filter(e => !e.matched).map(e => e.key),
  }
}
