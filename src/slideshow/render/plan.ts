/**
 * The ordered render plan (slideshow spec 6.3): authored z-order kept as
 * maximal static runs interleaved with animated targets, then the watermark.
 *
 * Nodes are pruned trees: the wanted nodes plus clones of their ancestor
 * groups (children filtered), so ancestor transforms are kept and every leaf
 * is painted by exactly one layer. The slide model is never mutated.
 */
import type { PptxShape, PptxSlide } from '../../pptx/types'
import { targetKey } from '../target'
import type { BuildOwnership, TargetKey } from '../types'
import { textClip, textFilter, underlayOf } from './scene'
import type { TextItem } from './text-split'

export type ColourPart = 'fill' | 'line' | 'text'

export type PlanLayer =
  /** `background` marks the first static run: it also paints the slide's solid fill, full-slide. */
  | { kind: 'static'; nodes: PptxShape[]; background?: true }
  /**
   * With `sub`, the target is a group whose animated descendants are
   * sub-targets: `sub` holds the group's content in z-order (static runs and
   * sub-targets), its `nodes` are empty, and the group's state composes over
   * every sub-layer.
   */
  | { kind: 'target'; key: TargetKey; nodes: PptxShape[]; part?: 'fill' | 'line' | 'text' | 'body' | 'range'; sub?: PlanLayer[] }
  | { kind: 'watermark' }

export interface RenderPlan { layers: readonly PlanLayer[] }

interface Owner { key: TargetKey; part?: 'fill' | 'line' | 'text' | 'body' | 'range' }
interface Leaf {
  /** Ancestor groups, outermost first. */
  path: PptxShape[]
  node: PptxShape
  /** The leaf's own target (a whole shape, a part, a body or a range); undefined is static. */
  own?: Owner
  /** Underlay: never a target. */
  fixed?: true
}

/**
 * Plan nodes painted from the slide's underlay (master and layout artwork, and their group clones). Shape ids are
 * scoped per part, so an underlay shape can share a slide shape's id: these nodes are never targets, triggers or
 * links (PR #18 review 7).
 */
export const fromUnderlay = new WeakSet<PptxShape>()

export function shapeKey(shape: PptxShape): TargetKey | undefined {
  return shape.source?.id !== undefined ? targetKey({ spid: shape.source.id }) : undefined
}

/** A clone painting only the fill (with the shape's shadow, pictures and content). */
function fillPart(shape: PptxShape): PptxShape {
  return { ...shape, textBody: undefined, line: undefined, drawingStyle: shape.drawingStyle && { ...shape.drawingStyle, line: undefined } }
}
function linePart(shape: PptxShape): PptxShape {
  return {
    ...shape, textBody: undefined, fill: undefined, useBgFill: undefined, content: undefined, table: undefined, imageIndex: undefined, chartFallback: undefined,
    drawingStyle: shape.drawingStyle && { line: shape.drawingStyle.line, issues: shape.drawingStyle.issues },
  }
}
function textPart(shape: PptxShape): PptxShape {
  return {
    ...shape, drawingGeometry: undefined, fill: undefined, line: undefined, drawingStyle: undefined,
    content: undefined, table: undefined, imageIndex: undefined, chartFallback: undefined,
  }
}
function filtered(shape: PptxShape, keep: (paragraph: number) => boolean): PptxShape {
  textFilter.set(shape, keep)
  return shape
}
/** One text-only clone per word or letter item (S2-8): its paragraph, clipped to the item's boxes at paint time. */
function boxLeaves(node: PptxShape, path: PptxShape[], items: readonly TextItem[]): Leaf[] {
  const seen = new Set<TargetKey>()
  const out: Leaf[] = []
  for (const item of items) {
    if (seen.has(item.key)) continue
    seen.add(item.key)
    const clone = filtered(textPart(node), i => i === item.paragraph)
    textClip.set(clone, { items, key: item.key })
    out.push({ path, node: clone, own: { key: item.key, part: 'range' as const } })
  }
  return out
}

/**
 * `colourParts` splits a shape into fill, line and text layers when a colour
 * animation addresses some but not all of them. Paragraph splitting comes
 * only from `builds`: a `paragraph` build gives a body layer (geometry, fill,
 * line, and text outside every range) and one text layer per range; range
 * `st..end` are inclusive paragraph indices (as `p:pRg`). With `animBg` off
 * the body is static and joins the surrounding static run.
 *
 * `splits` (S2-8, text by word or letter) maps an iterated target, a shape
 * key or a paragraph-range key, to its items (`textItems`, or the boxes of
 * `splitTextBoxes`). A split shape gives a body layer (geometry, fill, line,
 * and text in no item; keyed by the shape when it is animated) and one text
 * layer per item, in reading order, clipped at paint time to the item's
 * boxes on the painted layout. A split range replaces that range's single
 * layer with its items.
 */
export function buildRenderPlan(
  slide: PptxSlide,
  animated: ReadonlySet<TargetKey>,
  colourParts: ReadonlyMap<TargetKey, ReadonlySet<ColourPart>>,
  builds: readonly BuildOwnership[],
  splits?: ReadonlyMap<TargetKey, readonly TextItem[]>,
): RenderPlan {
  const paragraphBuilds = new Map<string, BuildOwnership>()
  for (const b of builds) if (b.mode === 'paragraph' && b.ranges.length > 0) paragraphBuilds.set(b.spid, b)

  const groupTarget = (shape: PptxShape): TargetKey | undefined => {
    const key = shapeKey(shape)
    return key !== undefined && animated.has(key) ? key : undefined
  }

  const leaves: Leaf[] = []
  const flatten = (list: readonly PptxShape[], path: PptxShape[], fixed: boolean): void => {
    for (const node of list) {
      if (node.group) { if (fixed) fromUnderlay.add(node); flatten(node.children ?? [], [...path, node], fixed); continue }
      if (fixed) { fromUnderlay.add(node); leaves.push({ path, node, fixed: true }); continue }
      const key = shapeKey(node)
      const whole = key !== undefined ? splits?.get(key) : undefined
      if (key !== undefined && whole && whole.length > 0) {
        const boxed = new Set(whole.map(b => b.paragraph))
        const keep = (i: number) => !boxed.has(i)
        const parts = colourParts.get(key)
        if (parts && parts.size > 0) {
          // A colour animation on the split shape (PR #19 review 7): the body is its fill, line and residual text
          // layers, so fill and line take their colour channels; the words keep theirs (text). The residual text
          // keeps the paragraph filter, so the item paragraphs are not painted twice (PR #19 review 8).
          leaves.push({ path, node: fillPart(node), own: { key, part: 'fill' } })
          leaves.push({ path, node: linePart(node), own: { key, part: 'line' } })
          leaves.push({ path, node: filtered(textPart(node), keep), own: { key, part: 'text' } })
        } else leaves.push({ path, node: filtered({ ...node }, keep), own: animated.has(key) ? { key, part: 'body' } : undefined })
        leaves.push(...boxLeaves(node, path, whole))
        continue
      }
      const build = node.source?.id !== undefined ? paragraphBuilds.get(node.source.id) : undefined
      if (build) {
        const inRange = (i: number): boolean => build.ranges.some(r => i >= r.st && i <= r.end)
        leaves.push({ path, node: filtered({ ...node }, i => !inRange(i)), own: build.animBg || animated.has(build.body) ? { key: build.body, part: 'body' } : undefined })
        for (const r of build.ranges) {
          const boxes = splits?.get(r.key)
          if (boxes && boxes.length > 0) leaves.push(...boxLeaves(node, path, boxes))
          else leaves.push({ path, node: filtered(textPart(node), i => i >= r.st && i <= r.end), own: { key: r.key, part: 'range' } })
        }
        continue
      }
      const parts = key !== undefined ? colourParts.get(key) : undefined
      if (key !== undefined && parts && parts.size > 0 && parts.size < 3) {
        leaves.push({ path, node: fillPart(node), own: { key, part: 'fill' } })
        leaves.push({ path, node: linePart(node), own: { key, part: 'line' } })
        leaves.push({ path, node: textPart(node), own: { key, part: 'text' } })
        continue
      }
      leaves.push({ path, node, own: key !== undefined && animated.has(key) ? { key } : undefined })
    }
  }
  flatten(underlayOf(slide), [], true)
  flatten(slide.shapes, [], false)

  const layers = runs(leaves, 0, groupTarget)
  const first = layers[0]
  if (first?.kind === 'static') first.background = true
  else layers.unshift({ kind: 'static', nodes: [], background: true })
  layers.push({ kind: 'watermark' })
  return { layers }
}

/**
 * Split consecutive leaves into layers. Ownership goes to the outermost
 * animated group below `depth`, else to the leaf's own target.
 */
function runs(leaves: readonly Leaf[], depth: number, groupTarget: (shape: PptxShape) => TargetKey | undefined): PlanLayer[] {
  type Seg = { token: object | undefined; key?: TargetKey; part?: Owner['part']; groupDepth?: number; leaves: Leaf[] }
  const segs: Seg[] = []
  for (const leaf of leaves) {
    let token: object | undefined, key: TargetKey | undefined, part: Owner['part'], groupDepth: number | undefined
    if (!leaf.fixed) {
      for (let i = depth; i < leaf.path.length; i++) {
        const k = groupTarget(leaf.path[i])
        if (k !== undefined) { token = leaf.path[i]; key = k; groupDepth = i; break }
      }
      if (!token && leaf.own) { token = leaf.own; key = leaf.own.key; part = leaf.own.part }
    }
    const last = segs[segs.length - 1]
    if (last && last.token === token) last.leaves.push(leaf)
    else segs.push({ token, key, part, groupDepth, leaves: [leaf] })
  }
  return segs.map((seg): PlanLayer => {
    if (!seg.token || seg.key === undefined) return { kind: 'static', nodes: prune(seg.leaves) }
    if (seg.groupDepth === undefined) return { kind: 'target', key: seg.key, nodes: prune(seg.leaves), ...(seg.part ? { part: seg.part } : {}) }
    const inner = runs(seg.leaves, seg.groupDepth + 1, groupTarget)
    if (inner.length === 1 && inner[0].kind === 'static') return { kind: 'target', key: seg.key, nodes: inner[0].nodes }
    return { kind: 'target', key: seg.key, nodes: [], sub: inner }
  })
}

/** Rebuild the consecutive leaves as trees, cloning each ancestor group once per run. */
function prune(leaves: readonly Leaf[]): PptxShape[] {
  const roots: PptxShape[] = []
  let chain: Array<{ source: PptxShape; clone: PptxShape }> = []
  for (const leaf of leaves) {
    let common = 0
    while (common < chain.length && common < leaf.path.length && chain[common].source === leaf.path[common]) common++
    chain = chain.slice(0, common)
    for (let i = common; i < leaf.path.length; i++) {
      const clone: PptxShape = { ...leaf.path[i], children: [] }
      if (fromUnderlay.has(leaf.path[i])) fromUnderlay.add(clone)
      const parent = chain[chain.length - 1]
      if (parent) parent.clone.children!.push(clone)
      else roots.push(clone)
      chain.push({ source: leaf.path[i], clone })
    }
    const parent = chain[chain.length - 1]
    if (parent) parent.clone.children!.push(leaf.node)
    else roots.push(leaf.node)
  }
  return roots
}
