/**
 * Text builds by word and by letter in the live show (S2-8; spec 6.5, 7).
 *
 * From a slide's timing, finds every target under a `p:iterate type="wd"/"lt"`,
 * splits its text into items (`textItems`: text only, no layout, so the
 * timeline never waits on fonts), and fans the timing out over them
 * (`expandIterations`). The stage feeds `splits` to the render plan and
 * `fan` to `renderStateAt`; the session schedules `tree` with `orderOf`.
 */
import { emuToPx } from '../../core/geometry'
import type { FontResolver } from '../../core/fonts/register'
import { layoutTextBody } from '../../drawing/text-layout'
import { createTextBodyMeasurer } from '../../drawing/text-paint'
import type { PptxShape, PptxSlide } from '../../pptx/types'
import { expandIterations, type BaseGeometry, type IteratedTiming } from '../engine/behaviours'
import { boxesFor, MAX_TEXT_BOXES, textBoxBases, textItems, type TextItem } from '../render/text-split'
import { targetKey } from '../target'
import type { DiagnosticSink, Iterate, TargetKey, TargetRef, TimeNode, TimingTree } from '../types'

export interface TextTiming extends IteratedTiming {
  /** Iterated target key → its items, for `buildRenderPlan`. */
  splits: ReadonlyMap<TargetKey, readonly TextItem[]>
}

export interface TextTimingOptions {
  sink?: DiagnosticSink
  slideIndex?: number
  /** Reserves bytes before items and the working copy are allocated (the slide's metadata handle). */
  charge?: (bytes: number, tag: string) => boolean
  /** Layers the slide can still give to items (default 128). */
  layers?: number
}

const MAX_DEPTH = 64

/** Every behaviour target under a `wd`/`lt` iterate, by key, with every unit its effects use. */
function iteratedTargets(tree: TimingTree): Map<TargetKey, { ref: Extract<TargetRef, { spid: string }>; units: Set<'wd' | 'lt'> }> {
  const out = new Map<TargetKey, { ref: Extract<TargetRef, { spid: string }>; units: Set<'wd' | 'lt'> }>()
  const seen = new Set<TimeNode>()
  const walk = (n: TimeNode, it: Iterate | undefined, depth: number): void => {
    if (depth > MAX_DEPTH || seen.has(n)) return
    seen.add(n)
    const own = n.ctn.iterate && n.ctn.iterate.type !== 'el' ? n.ctn.iterate : undefined
    const here = own ?? it
    if ('children' in n) { for (const c of n.children) walk(c, here, depth + 1); return }
    if (!here || here.type === 'el' || !('spid' in n.target)) return
    const key = targetKey(n.target)
    const have = out.get(key)
    if (have) have.units.add(here.type)
    else out.set(key, { ref: n.target, units: new Set([here.type]) })
  }
  walk(tree.root, undefined, 0)
  return out
}

/** Shapes by `source.id`, groups searched (the first wins). */
export function shapesById(slide: PptxSlide): Map<string, PptxShape> {
  const out = new Map<string, PptxShape>()
  const walk = (list: readonly PptxShape[], depth: number): void => {
    if (depth > MAX_DEPTH) return
    for (const s of list) {
      const id = s.source?.id
      if (id !== undefined && !out.has(id)) out.set(id, s)
      if (s.group) walk(s.children ?? [], depth + 1)
    }
  }
  walk(slide.shapes ?? [], 0)
  return out
}

/** The slide's text builds, or undefined when its timing iterates no text by word or letter. */
export function textTimingFor(slide: PptxSlide, tree: TimingTree, options: TextTimingOptions = {}): TextTiming | undefined {
  const targets = iteratedTargets(tree)
  if (targets.size === 0) return undefined
  const shapes = shapesById(slide)
  const splits = new Map<TargetKey, readonly TextItem[]>()
  const keys = new Map<TargetKey, TargetKey[]>()
  const groups = new Map<TargetKey, Partial<Record<'wd' | 'lt', TargetKey[][]>>>()
  let budget = Math.max(0, Math.min(MAX_TEXT_BOXES, options.layers ?? MAX_TEXT_BOXES))
  for (const [key, { ref, units }] of targets) {
    const shape = shapes.get(ref.spid)
    // No text, or a character range: expandIterations diagnoses it and the effect plays whole.
    if (!shape?.textBody || ref.text?.kind === 'char') continue
    const common = {
      sink: options.sink, slideIndex: options.slideIndex, charge: options.charge,
      ...(ref.text ? { paragraphs: { st: ref.text.st, end: ref.text.end } } : {}),
    }
    // Layers follow the finest unit any effect on the target uses (PR #19 review 9).
    const items = textItems(shape, units.has('lt') ? 'lt' : 'wd', { ...common, limit: budget })
    if (items.length === 0) continue
    if (units.has('lt') && units.has('wd')) {
      // Built by word as well: each word is the group of letters inside it. The word items only derive the groups
      // and are dropped at once, so they take no retained charge (and repeat no diagnostic) (PR #19 review 10).
      const words = textItems(shape, 'wd', { slideIndex: options.slideIndex, ...(ref.text ? { paragraphs: { st: ref.text.st, end: ref.text.end } } : {}) })
      if (words.length === 0) {
        options.sink?.add({ code: 'text-build-refused', slideIndex: options.slideIndex, effect: key, message: `Text by word and by letter on ${key}: the words could not be found; its effects play on the whole target` })
        continue
      }
      groups.set(key, { wd: wordGroups(items, words) })
    }
    const unique = [...new Set(items.map(i => i.key))]
    budget = Math.max(0, budget - unique.length)
    splits.set(key, items)
    keys.set(key, unique)
  }
  const x = expandIterations(tree, keys, { sink: options.sink, slideIndex: options.slideIndex, charge: options.charge, groups })
  // Only targets that fanned out are split; a refused fan-out leaves the text in its shape (PR #19 review 9).
  const fanned = new Map([...splits].filter(([k]) => x.items.has(k)))
  return { ...x, splits: fanned }
}

/**
 * The letter keys of each word, in word order: a letter belongs to the first word of its own paragraph whose
 * character range (paragraph-local) holds it (or, for paragraph items of a coarsened split, the first in its
 * paragraph). Words left empty are dropped.
 */
function wordGroups(letters: readonly TextItem[], words: readonly TextItem[]): TargetKey[][] {
  const order: TargetKey[] = []
  const byWord = new Map<TargetKey, TargetKey[]>()
  for (const w of words) if (!byWord.has(w.key)) { byWord.set(w.key, []); order.push(w.key) }
  const placed = new Set<TargetKey>()
  for (const l of letters) {
    if (placed.has(l.key)) continue
    // Offsets are paragraph-local: a letter joins a word of its own paragraph only (PR #19 review 12).
    const w = words.find(w => l.paragraph === w.paragraph && (l.st !== undefined && l.end !== undefined && w.st !== undefined && w.end !== undefined
      ? l.st >= w.st && l.end <= w.end
      : true))
    if (!w) continue
    byWord.get(w.key)!.push(l.key)
    placed.add(l.key)
  }
  const out = order.map(k => byWord.get(k)!).filter(g => g.length > 0)
  // A letter in no word (whitespace never is an item; a stray one) moves with the last word.
  const loose = [...new Set(letters.map(l => l.key))].filter(k => !placed.has(k))
  if (loose.length) { if (out.length) out[out.length - 1].push(...loose); else out.push(loose) }
  return out
}

/**
 * Base geometry for every item of `splits`, laid out with the real fonts:
 * each shape's text is laid out once on `measure` and its item boxes placed
 * relative to the shape's own base (`base`, group transforms composed).
 * `tight` gives each word or letter its own measured span (`boxesFor`).
 */
export function textBases(slide: PptxSlide, splits: ReadonlyMap<TargetKey, readonly TextItem[]>, base: ReadonlyMap<TargetKey, BaseGeometry>,
  measure: CanvasRenderingContext2D, resolveFont: FontResolver, tight = false): Map<TargetKey, BaseGeometry> {
  const out = new Map<TargetKey, BaseGeometry>()
  const shapes = shapesById(slide)
  const measurer = createTextBodyMeasurer(measure, resolveFont)
  for (const items of splits.values()) {
    const spid = items[0]?.key.slice(3, items[0].key.indexOf('/'))
    const shape = spid !== undefined ? shapes.get(decodeURIComponent(spid)) : undefined
    if (!shape?.textBody) continue
    const layout = layoutTextBody(shape.textBody, emuToPx(shape.widthEmu), emuToPx(shape.heightEmu), measurer, slide.theme)
    const own = base.get(targetKey({ spid: shape.source!.id! }))
    for (const [k, g] of textBoxBases(shape, boxesFor(shape, layout, items, measurer, tight), own)) out.set(k, g)
  }
  return out
}
