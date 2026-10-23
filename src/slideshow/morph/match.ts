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
 * approximation, recorded in the task report: same name and kind, then word
 * or character units, then same preset geometry within 25% of the slide
 * diagonal.
 *
 * Unit mode (byWord/byChar): every object pair from rules 1 and 2 whose
 * shapes both carry text decomposes into unit pairs (LCS-aligned over the
 * engine's word/letter item keys); the container pair with its rule of
 * origin keeps the frame animatable, and unmatched units fade individually.
 * Remaining text-bearing objects pair their units across objects by equal
 * content, and each container follows the partner that carries most of its
 * matched units. Unmatched objects fade whole; unmatched text units fade
 * only when their container is paired.
 *
 * Bounds, monotonic: slides beyond `MORPH_MAX_OBJECTS` refuse and diagnose
 * (`morph-refused`); once text unit pairs pass `MORPH_MAX_UNIT_PAIRS`, later
 * pairs coarsen to object level (`morph-units-coarsened`) and cross-object
 * matching stops — never silently.
 */
import type { PptxShape, PptxSlide } from '../../pptx/types'
import { textItems, type TextItem } from '../render/text-split'
import { targetKey } from '../target'
import type { DiagnosticSink, TargetKey } from '../types'

/** The most objects one slide contributes to a morph match (refusal beyond). */
export const MORPH_MAX_OBJECTS = 256
/** The most text unit pairs one match emits; beyond, later pairs coarsen. */
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
  /** The shape's text units, read lazily once in unit mode. */
  units?: ItemView[]
}

/** One text unit: its engine item key and content. */
interface ItemView {
  key: TargetKey
  text: string
  matched: boolean
}

interface Ctx {
  pairs: MorphPair[]
  unitPairs: number
  /** Set once the unit-pair cap is passed: later pairs coarsen. */
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

/** Structural object class for name-kind and geometry matching. Text
 * presence is not part of it: adding text does not change a shape's type. */
function kindOf(shape: PptxShape): string {
  if (shape.children?.length || shape.group) return 'group'
  if (shape.image) return 'picture'
  if (shape.table || shape.content || shape.chartFallback) return 'graphic'
  return 'shape'
}

/** Preset geometry bucket: same kind and geometry class, and the preset name
 * for other/absent geometries so two freeform presets do not pair. */
function geometryOf(shape: PptxShape): string {
  const g = shape.geometry
  const class_ = g && g !== 'other' ? g : shape.presetName ? `:${shape.presetName}` : ':'
  return `${kindOf(shape)}|${class_}`
}

function collect(slide: PptxSlide): Entry[] {
  const out: Entry[] = []
  const seen = new Set<string>()
  for (const shape of slide.shapes) {
    const id = shape.source?.id
    // Duplicate cNvPr ids collapse to the first occurrence, as the engine
    // treats them everywhere; later duplicates claim nothing (no pair, no
    // fade entry) because they have no key of their own.
    if (id === undefined || seen.has(id)) continue
    seen.add(id)
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
      add(ctx, fq[i], tq[i], rule)
    }
  }
}

function add(ctx: Ctx, a: Entry, b: Entry, rule: MorphRule): void {
  a.matched = true
  b.matched = true
  ctx.pairs.push({ from: a.key, to: b.key, rule })
}

/** The text units of a shape, from the engine's own item splitter. */
function unitsOf(ctx: Ctx, e: Entry): ItemView[] {
  if (e.units) return e.units
  const body = e.shape.textBody
  const units: ItemView[] = []
  if (body) {
    const paragraphText = (p: number): string => body.paragraphs[p].runs.map(r => r.text ?? '').join('')
    for (const item of textItems(e.shape, ctx.unit!, { sink: ctx.sink, slideIndex: ctx.slideIndex })) {
      units.push({ key: item.key, text: itemText(paragraphText, item), matched: false })
    }
  }
  e.units = units
  return units
}

/** A text item's content: paragraph-local character range, or the paragraph. */
function itemText(paragraphText: (p: number) => string, item: TextItem): string {
  const whole = paragraphText(item.paragraph)
  if (item.st === undefined) return whole
  return whole.slice(item.st, item.end ?? item.st)
}

/** LCS alignment of equal-content units, in reading order, one-to-one. */
function alignUnits(a: readonly ItemView[], b: readonly ItemView[]): Array<[ItemView, ItemView]> {
  const n = a.length
  const m = b.length
  if (!n || !m) return []
  const dp: number[][] = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(0))
  for (let i = 1; i <= n; i++) {
    for (let j = 1; j <= m; j++) {
      dp[i][j] = a[i - 1].text === b[j - 1].text ? dp[i - 1][j - 1] + 1 : Math.max(dp[i - 1][j], dp[i][j - 1])
    }
  }
  const out: Array<[ItemView, ItemView]> = []
  let i = n
  let j = m
  while (i > 0 && j > 0) {
    if (a[i - 1].text === b[j - 1].text && dp[i][j] === dp[i - 1][j - 1] + 1) {
      out.push([a[i - 1], b[j - 1]])
      a[i - 1].matched = true
      b[j - 1].matched = true
      i--
      j--
    } else if (dp[i - 1][j] >= dp[i][j - 1]) i--
    else j--
  }
  out.reverse()
  return out
}

/** Decomposes a rules-1/2 object pair into text units: the container pair
 * keeps its rule of origin, then aligned unit pairs follow. Coarsening is
 * monotonic: past the cap, this and every later pair stays object-level. */
function decompose(ctx: Ctx, a: Entry, b: Entry): void {
  if (!a.shape.textBody || !b.shape.textBody) return
  const unitsA = unitsOf(ctx, a)
  const unitsB = unitsOf(ctx, b)
  const aligned = alignUnits(unitsA, unitsB)
  if (aligned.length === 0) return
  if (ctx.coarsened || ctx.unitPairs + aligned.length > MORPH_MAX_UNIT_PAIRS) {
    // Coarsening must leave nothing behind (N2): every unit of the
    // coarsened pair is marked matched, so it cannot linger as an
    // individual fade alongside the object pair.
    ctx.coarsened = true
    for (const u of unitsA) u.matched = true
    for (const u of unitsB) u.matched = true
    ctx.sink?.add({
      code: 'morph-units-coarsened', slideIndex: ctx.slideIndex, effect: 'sld',
      message: `Morph text units beyond ${MORPH_MAX_UNIT_PAIRS} pairs match as objects`,
    })
    return
  }
  ctx.unitPairs += aligned.length
  for (const [fa, fb] of aligned) ctx.pairs.push({ from: fa.key, to: fb.key, rule: 'text' })
}

/** Rule 3 in unit mode: the remaining text-bearing objects pair by position
 * first (nearest-greedy in reading order, one to one, ties to the first
 * candidate), so equal-text repeats cannot scatter a unit across the slide
 * (N3); units then align by equal content within each pair. Cross-object
 * content pairing is gone. */
function pairUnitsAcrossObjects(ctx: Ctx, fromEntries: readonly Entry[], toEntries: readonly Entry[]): void {
  // Past the unit cap the cross-object stage stops entirely (N1): remaining
  // containers fade whole, which the unit-bound diagnostic already recorded.
  if (ctx.coarsened) return
  const untagged = (e: Entry): boolean => !e.name.startsWith('!!')
  const fromF = fromEntries.filter(e => !e.matched && untagged(e) && e.text !== '')
  const toF = toEntries.filter(e => !e.matched && untagged(e) && e.text !== '')
  if (!fromF.length || !toF.length) return

  for (const a of fromF) {
    if (a.matched) continue
    let best: Entry | undefined
    let bestD = Infinity
    for (const b of toF) {
      if (b.matched) continue
      const dx = a.cx - b.cx, dy = a.cy - b.cy, d = dx * dx + dy * dy
      if (d < bestD) { bestD = d; best = b }
    }
    if (!best) continue
    a.matched = true
    best.matched = true
    const before = ctx.pairs.length
    alignPairUnits(ctx, a, best)
    // Nothing shared content: the position pairing lets both sides rest
    // unmatched, and they fade whole.
    if (ctx.pairs.length === before) {
      a.matched = false
      best.matched = false
    }
  }
}

/** Aligns the text units of one position-paired container pair by equal
 * content (reading order, one to one): swapping two units inside one box
 * pairs them, where an LCS would not. On success the container pair follows
 * its unit pairs; past the cap the pair coarsens to object level and none
 * of its units can linger as individual fades (N2). */
function alignPairUnits(ctx: Ctx, a: Entry, b: Entry): void {
  const aUnits = unitsOf(ctx, a)
  const bUnits = unitsOf(ctx, b)
  if (!aUnits.length || !bUnits.length) return
  const toByText = new Map<string, ItemView[]>()
  for (const u of bUnits) {
    const q = toByText.get(u.text)
    if (q) q.push(u)
    else toByText.set(u.text, [u])
  }
  const aligned: Array<[ItemView, ItemView]> = []
  for (const u of aUnits) {
    const q = toByText.get(u.text)?.find(v => !v.matched)
    if (!q) continue
    u.matched = true
    q.matched = true
    aligned.push([u, q])
  }
  if (!aligned.length) return
  if (ctx.coarsened || ctx.unitPairs + aligned.length > MORPH_MAX_UNIT_PAIRS) {
    ctx.coarsened = true
    for (const u of aUnits) u.matched = true
    for (const u of bUnits) u.matched = true
    ctx.pairs.push({ from: a.key, to: b.key, rule: 'text' })
    ctx.sink?.add({
      code: 'morph-units-coarsened', slideIndex: ctx.slideIndex, effect: 'sld',
      message: `Morph text units beyond ${MORPH_MAX_UNIT_PAIRS} pairs match as objects`,
    })
    return
  }
  ctx.unitPairs += aligned.length
  for (const [fa, fb] of aligned) ctx.pairs.push({ from: fa.key, to: fb.key, rule: 'text' })
  ctx.pairs.push({ from: a.key, to: b.key, rule: 'text' })
  a.matched = true
  b.matched = true
}

/** Rule 4: same preset geometry and kind, nearer than a quarter of the
 * slide diagonal; greedy nearest in reading order. */
function pairByGeometry(ctx: Ctx, fromEntries: readonly Entry[], toEntries: readonly Entry[], threshold: number): void {
  const fromQ = queueBy(fromEntries, e => e.geometry, e => !e.name.startsWith('!!'))
  const toQ = queueBy(toEntries, e => e.geometry, e => !e.name.startsWith('!!'))
  for (const fq of fromQ.values()) {
    for (const a of fq) {
      if (a.matched) continue
      const tq = toQ.get(a.geometry)
      if (!tq) continue
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

/** Whole-object fade plus, for paired containers, leftover unit fades. */
function fades(entries: readonly Entry[]): TargetKey[] {
  const out: TargetKey[] = []
  for (const e of entries) if (!e.matched) out.push(e.key)
  for (const e of entries) if (e.matched && e.units) for (const u of e.units) if (!u.matched) out.push(u.key)
  return out
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

  // Rule 1: `!!` names, exactly as Microsoft publishes. These entries are
  // matched now and excluded from every rule below, so a !! object cannot
  // pair with a non-!! object.
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

  if (option === 'byObject') {
    // Rule 3: same flat text, whole objects. !! entries never participate.
    const untagged = (e: Entry): boolean => !e.name.startsWith('!!')
    pairQueues(
      ctx,
      queueBy(fromEntries, e => e.text, e => e.text !== '' && untagged(e)),
      queueBy(toEntries, e => e.text, e => e.text !== '' && untagged(e)),
      'text',
    )
  } else {
    // Unit mode: the pairs made by rules 1-2 decompose into text units, the
    // container pair already emitted with its rule of origin, then remaining
    // objects pair units across objects.
    const rulePairs = [...ctx.pairs]
    const fromByKey = new Map(fromEntries.map(e => [e.key, e]))
    const toByKey = new Map(toEntries.map(e => [e.key, e]))
    for (const pair of rulePairs) {
      const a = fromByKey.get(pair.from)
      const b = toByKey.get(pair.to)
      if (a && b) decompose(ctx, a, b)
    }
    pairUnitsAcrossObjects(ctx, fromEntries, toEntries)
  }

  // Rule 4: same preset geometry and kind near a quarter of the slide diagonal.
  const threshold = 0.25 * Math.hypot(Math.max(from.widthEmu, to.widthEmu), Math.max(from.heightEmu, to.heightEmu))
  pairByGeometry(ctx, fromEntries, toEntries, threshold)

  return {
    pairs: ctx.pairs,
    fadeOut: fades(fromEntries),
    fadeIn: fades(toEntries),
  }
}
