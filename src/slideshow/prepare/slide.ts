/**
 * DOM-free slide preparation (piece 4 spec §4 item 1), shared by the local
 * stage (`player/stage.ts`) and the stream server (bun, `@napi-rs/canvas`).
 *
 * From a slide and its show metadata: the timing targets (`engine/targets`),
 * the text builds by word or letter, the render plan, its rasterising
 * layers and the base geometry. Then, on demand, a layer's raster, the
 * text items' measured geometry and the painted hit regions of shapes.
 *
 * Nothing here touches `document`, `window` or a compositor: canvases come
 * from the caller's `CanvasFactory` (`RasterContext.canvas`) and text is
 * measured on the caller's context. Uploading, image decoding and workspace
 * leases stay with the caller.
 */
import type { FontResolver } from '../../core/fonts/register'
import type { PptxShape, PptxSlide } from '../../pptx/types'
import type { BaseGeometry } from '../engine/behaviours'
import { timingTargets } from '../engine/targets'
import { apply, groupChildMatrix, nodeMatrix } from '../player/hit-test'
import { textBases, textTimingFor, type TextTiming } from '../player/text-builds'
import { buildRenderPlan, shapeKey, type ColourPart, type PlanLayer, type RenderPlan } from '../render/plan'
import { rasterLayer, type RasterContext, type RasterResult } from '../render/raster'
import { paintedBounds } from '../render/scene'
import type { TextItem } from '../render/text-split'
import type { Affine, DiagnosticSink, ShowSlide, TargetKey, TimeNode } from '../types'
import { LAYER_LIMIT } from '../workspace'

export interface PrepareInput {
  index: number
  slide: PptxSlide
  meta: ShowSlide
  /** Whether the plan's watermark layer rasterises (a mark is configured). */
  watermark: boolean
  /** Reserves the text-build working copy's bytes (the slide's metadata handle). */
  charge?: (bytes: number, tag: string) => boolean
  /** Additive to the brief: where text-build diagnostics go. */
  sink?: DiagnosticSink
  /**
   * Additive to the brief: the slide's text timing when the caller already holds it (the stage caches one
   * working copy per `ShowSlide`, shared by the scheduled timeline); `null` for none. Absent: computed here.
   */
  textTiming?: TextTiming | null
}

export interface PreparedSlide {
  index: number
  slide: PptxSlide
  meta: ShowSlide
  plan: RenderPlan
  layers: PlanLayer[]
  /** The plan's inputs (S3-6: a Morph re-plans the slide with every object keyed). */
  animated: ReadonlySet<TargetKey>
  parts: ReadonlyMap<TargetKey, ReadonlySet<ColourPart>>
  splits?: ReadonlyMap<TargetKey, readonly TextItem[]>
  base: Map<TargetKey, BaseGeometry>
  hidden: Set<TargetKey>
  /** The entrance behaviours of each hidden target: it shows once one of them has begun. */
  entrances: Map<TargetKey, Set<TimeNode>>
  /** The slide's text builds by word or letter (S2-8), when it has any. */
  textTiming?: TextTiming
}

const IDENTITY: Affine = [1, 0, 0, 1, 0, 0]

/** Base geometry (slide px, group transforms composed) per shape key, groups included; ranges and parts share their shape's. */
export function baseGeometry(slide: PptxSlide): Map<TargetKey, BaseGeometry> {
  const out = new Map<TargetKey, BaseGeometry>()
  const walk = (nodes: readonly PptxShape[], parent: Affine, rot: number, depth: number) => {
    if (depth > 64) return
    for (const shape of nodes) {
      const local = nodeMatrix(shape, parent)
      if (!local) continue
      const r = rot + (shape.rotationDeg ?? 0)
      const key = shapeKey(shape)
      if (key !== undefined) out.set(key, geometryOf(local.m, local.w, local.h, r))
      if (shape.group) {
        const child = groupChildMatrix(shape, local.m)
        if (child) walk(shape.children ?? [], child, r, depth + 1)
      }
    }
  }
  walk(slide.shapes, IDENTITY, 0, 0)
  return out
}

function geometryOf(m: Affine, w: number, h: number, rot = 0): BaseGeometry {
  const [cx, cy] = apply(m, w / 2, h / 2)
  const sx = Math.hypot(m[0], m[1]), sy = Math.hypot(m[2], m[3])
  const ww = w * sx, hh = h * sy
  return { x: cx - ww / 2, y: cy - hh / 2, w: ww, h: hh, rot: Number.isFinite(rot) ? rot : 0 }
}

/** Every plan layer that rasterises on its own, in paint order (sub-layers included; a group-with-sub paints through them). */
export function rasterLayers(plan: RenderPlan, watermark: boolean): PlanLayer[] {
  const out: PlanLayer[] = []
  const walk = (layers: readonly PlanLayer[]) => {
    for (const l of layers) {
      if (l.kind === 'watermark') { if (watermark) out.push(l); continue }
      if (l.kind === 'target' && l.sub) { walk(l.sub); continue }
      out.push(l)
    }
  }
  walk(plan.layers)
  return out
}

/**
 * The slide's text builds (S2-8), uncached: the layers its items may take are the layer bound less what the
 * slide plans without them (and a margin per split). `undefined` when its timing builds no text by word or letter.
 */
export function computeTextTiming(index: number, slide: PptxSlide | undefined, meta: ShowSlide, charge: ((bytes: number, tag: string) => boolean) | undefined, sink?: DiagnosticSink): TextTiming | undefined {
  const timing = meta.timing
  if (!timing || 'refused' in timing || !slide) return undefined
  const t = timingTargets(meta)
  const plain = buildRenderPlan(slide, t.animated, new Map(), meta.builds).layers.length
  return textTimingFor(slide, timing, { ...(sink ? { sink } : {}), slideIndex: index, charge: charge ?? (() => false), layers: LAYER_LIMIT - plain - 4 })
}

/** Timing targets of the scheduled tree, with every text item hidden and entered as its iterated target is. */
export function targetsOf(meta: ShowSlide, tt: TextTiming | undefined): ReturnType<typeof timingTargets> {
  const t = timingTargets(tt ? { ...meta, timing: tt.tree } : meta)
  if (tt) {
    for (const [target, keys] of tt.items) {
      for (const k of keys) {
        t.animated.add(k)
        if (t.hidden.has(target)) {
          t.hidden.add(k)
          t.entrances.set(k, t.entrances.get(target) ?? new Set())
        }
      }
    }
  }
  return t
}

/** Plans slide `index`: its targets, text builds, render plan, rasterising layers and base geometry. */
export function prepareSlide(input: PrepareInput): PreparedSlide {
  const { index, slide, meta } = input
  const tt = input.textTiming !== undefined ? input.textTiming ?? undefined : computeTextTiming(index, slide, meta, input.charge, input.sink)
  const { animated, colour, hidden, entrances } = targetsOf(meta, tt)
  // Any colour animation splits its shape into fill, line and text layers (a two-part set forces the split).
  const parts = new Map<TargetKey, ReadonlySet<ColourPart>>([...colour].map(k => [k, new Set<ColourPart>(['fill', 'line'])]))
  const plan = buildRenderPlan(slide, animated, parts, meta.builds, tt?.splits)
  return {
    index, slide, meta, plan, layers: rasterLayers(plan, input.watermark), animated, parts, ...(tt?.splits ? { splits: tt.splits } : {}),
    base: baseGeometry(slide), hidden, entrances, ...(tt ? { textTiming: tt } : {}),
  }
}

/** Lays out the text items' base geometry (S2-8) with the real fonts on `measure`, into `ps.base`. */
export function layoutTextBases(ps: PreparedSlide, measure: CanvasRenderingContext2D, resolveFont: FontResolver): void {
  const tt = ps.textTiming
  if (!tt || tt.splits.size === 0) return
  for (const [k, g] of textBases(ps.slide, tt.splits, ps.base, measure, resolveFont)) ps.base.set(k, g)
}

/** Rasterises one of `ps`'s plan layers (the caller uploads or releases its lease). */
export function rasterPreparedLayer(ps: PreparedSlide, layer: PlanLayer, ctx: RasterContext): RasterResult {
  return rasterLayer(layer, ps.slide, ctx)
}

/** A leaf with clones of its ancestor groups (children filtered to the one path), so their transforms apply. */
function chain(path: readonly PptxShape[], leaf: PptxShape): PptxShape {
  return path.reduceRight<PptxShape>((child, group) => ({ ...group, children: [child] }), leaf)
}

/**
 * The painted bounds (`paintedBounds`) of every leaf owned by each of `spids` (the shape itself, or every leaf of
 * a group), in slide px. Slide shapes only: underlay ids are scoped to their own part and never hit targets. A
 * spid held by several shapes is the first one's (as `shapesById`); a spid with nothing painted has no entry.
 */
export function hitRegionsOf(ps: PreparedSlide, spids: readonly string[], measure: CanvasRenderingContext2D, resolveFont: FontResolver): Map<string, Array<{ x: number; y: number; w: number; h: number }>> {
  const out = new Map<string, Array<{ x: number; y: number; w: number; h: number }>>()
  const want = new Set(spids)
  const claimed = new Set<string>()
  const walk = (nodes: readonly PptxShape[], path: PptxShape[], owner: string | undefined, depth: number): void => {
    if (depth > 64) return
    for (const shape of nodes) {
      let own = owner
      const id = shape.source?.id
      if (own === undefined && id !== undefined && want.has(id) && !claimed.has(id)) { claimed.add(id); own = id }
      if (shape.group) { walk(shape.children ?? [], [...path, shape], own, depth + 1); continue }
      if (own === undefined) continue
      const b = paintedBounds([chain(path, shape)], ps.slide, resolveFont, measure)
      if (!b) continue
      let list = out.get(own)
      if (!list) out.set(own, list = [])
      list.push({ x: b.minX, y: b.minY, w: b.maxX - b.minX, h: b.maxY - b.minY })
    }
  }
  walk(ps.slide.shapes ?? [], [], undefined, 0)
  return out
}
