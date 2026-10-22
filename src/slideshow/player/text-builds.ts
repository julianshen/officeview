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

/** Every behaviour target under a `wd`/`lt` iterate, by key (the first unit wins). */
function iteratedTargets(tree: TimingTree): Map<TargetKey, { ref: Extract<TargetRef, { spid: string }>; unit: 'wd' | 'lt' }> {
  const out = new Map<TargetKey, { ref: Extract<TargetRef, { spid: string }>; unit: 'wd' | 'lt' }>()
  const seen = new Set<TimeNode>()
  const walk = (n: TimeNode, it: Iterate | undefined, depth: number): void => {
    if (depth > MAX_DEPTH || seen.has(n)) return
    seen.add(n)
    const own = n.ctn.iterate && n.ctn.iterate.type !== 'el' ? n.ctn.iterate : undefined
    const here = own ?? it
    if ('children' in n) { for (const c of n.children) walk(c, here, depth + 1); return }
    if (!here || here.type === 'el' || !('spid' in n.target)) return
    const key = targetKey(n.target)
    if (!out.has(key)) out.set(key, { ref: n.target, unit: here.type })
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
  let budget = Math.max(0, Math.min(MAX_TEXT_BOXES, options.layers ?? MAX_TEXT_BOXES))
  for (const [key, { ref, unit }] of targets) {
    const shape = shapes.get(ref.spid)
    // No text, or a character range: expandIterations diagnoses it and the effect plays whole.
    if (!shape?.textBody || ref.text?.kind === 'char') continue
    const items = textItems(shape, unit, {
      sink: options.sink, slideIndex: options.slideIndex, limit: budget, charge: options.charge,
      ...(ref.text ? { paragraphs: { st: ref.text.st, end: ref.text.end } } : {}),
    })
    if (items.length === 0) continue
    const unique = [...new Set(items.map(i => i.key))]
    budget = Math.max(0, budget - unique.length)
    splits.set(key, items)
    keys.set(key, unique)
  }
  const x = expandIterations(tree, keys, { sink: options.sink, slideIndex: options.slideIndex, charge: options.charge })
  return { ...x, splits }
}

/**
 * Base geometry for every item of `splits`, laid out with the real fonts:
 * each shape's text is laid out once on `measure` and its item boxes placed
 * relative to the shape's own base (`base`, group transforms composed).
 */
export function textBases(slide: PptxSlide, splits: ReadonlyMap<TargetKey, readonly TextItem[]>, base: ReadonlyMap<TargetKey, BaseGeometry>,
  measure: CanvasRenderingContext2D, resolveFont: FontResolver): Map<TargetKey, BaseGeometry> {
  const out = new Map<TargetKey, BaseGeometry>()
  const shapes = shapesById(slide)
  const measurer = createTextBodyMeasurer(measure, resolveFont)
  for (const items of splits.values()) {
    const spid = items[0]?.key.slice(3, items[0].key.indexOf('/'))
    const shape = spid !== undefined ? shapes.get(decodeURIComponent(spid)) : undefined
    if (!shape?.textBody) continue
    const layout = layoutTextBody(shape.textBody, emuToPx(shape.widthEmu), emuToPx(shape.heightEmu), measurer, slide.theme)
    const own = base.get(targetKey({ spid: shape.source!.id! }))
    for (const [k, g] of textBoxBases(shape, boxesFor(shape, layout, items, measurer), own)) out.set(k, g)
  }
  return out
}
