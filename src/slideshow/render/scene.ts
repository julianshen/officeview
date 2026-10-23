/**
 * Shared painting helpers for the slideshow's layer rasteriser and direct
 * painter (slideshow spec 6.3, 7). Lazy chunk only.
 *
 * Painting goes through the unchanged `paintScene`; the per-node callbacks
 * mirror `renderSlide`'s. Content (tables, charts, diagrams, text boxes) is
 * delegated to `renderSlide` itself on a one-node synthetic slide with a
 * transparent background, so the private table painter stays private and
 * the shared painter is not touched.
 */
import type { ContentPaintAssets } from '../../drawing/content'
import type { ScenePaintOptions } from '../../drawing/scene-paint'
import { layoutTextBody } from '../../drawing/text-layout'
import { createTextBodyMeasurer, paintTextBody } from '../../drawing/text-paint'
import type { FontResolver } from '../../core/fonts/register'
import { emuToPx } from '../../core/geometry'
import { renderSlide } from '../../pptx/render'
import type { PptxShape, PptxSlide } from '../../pptx/types'

/** Stage 0 (S0-7) adds `PptxSlide.underlay`; read it structurally until then. */
export function underlayOf(slide: PptxSlide): readonly PptxShape[] {
  return (slide as PptxSlide & { underlay?: PptxShape[] }).underlay ?? []
}

/** The slide in CSS px, rounded exactly as `renderSlide` rounds it. */
export function slideSize(slide: PptxSlide): { w: number; h: number } {
  return { w: Math.round(emuToPx(slide.widthEmu)), h: Math.round(emuToPx(slide.heightEmu)) }
}

/**
 * Paragraph filters for planner-made clones: a clone listed here paints only
 * the laid-out lines whose paragraph index passes. The layout is computed on
 * the whole body, so kept lines stay exactly where they were.
 */
export const textFilter = new WeakMap<PptxShape, (paragraph: number) => boolean>()

const TRANSPARENT = 'rgba(0,0,0,0)'
const PX_TO_EMU = 914400 / 96

export interface SceneAssets {
  images: ReadonlyArray<CanvasImageSource | undefined>
  resolveFont: FontResolver
  contentAssets?: ContentPaintAssets
  /** Return true when the image was withheld on purpose (the caller diagnoses it); the static diagnostic is then skipped. */
  withheld?: (shape: PptxShape) => boolean
}

/** The `renderSlide` scene options, with paragraph filtering for clones. */
export function sceneOptions(slide: PptxSlide, assets: SceneAssets): ScenePaintOptions<PptxShape> {
  const { resolveFont } = assets
  return {
    images: assets.images,
    missingImage(shape) {
      if (assets.withheld?.(shape)) return
      const message = 'Embedded image could not be decoded; static representation is unavailable'
      if (shape.diagnostics?.some(issue => issue.kind === 'missing-image' && issue.message === message)) return
      const issue = { kind: 'missing-image' as const, message, feature: 'blipFill', source: shape.source }
      shape.diagnostics ??= []
      shape.diagnostics.push(issue)
      slide.diagnostics ??= []
      slide.diagnostics.push(issue)
    },
    paintContent(shape, ctx, w, h) {
      if (!shape.content && !shape.table) return
      // Only the content survives on this clone; renderSlide's transparent
      // background fill is a no-op under source-over.
      const node: PptxShape = {
        xEmu: 0, yEmu: 0, widthEmu: w * PX_TO_EMU, heightEmu: h * PX_TO_EMU,
        geometry: 'rect', content: shape.content, table: shape.table, source: shape.source,
      }
      renderSlide({ index: slide.index, widthEmu: slide.widthEmu, heightEmu: slide.heightEmu, shapes: [node], background: TRANSPARENT, theme: slide.theme },
        ctx, { widthPx: 0, heightPx: 0 }, undefined, undefined, resolveFont, assets.contentAssets)
    },
    paintText(shape, ctx, w, h) {
      const body = shape.textBody
      if (!body) return
      const keep = textFilter.get(shape)
      if (!keep) { paintTextBody(body, ctx, 0, 0, w, h, resolveFont, slide.theme); return }
      const layout = layoutTextBody(body, w, h, createTextBodyMeasurer(ctx, resolveFont), slide.theme)
      paintTextBody(body, ctx, 0, 0, w, h, resolveFont, slide.theme, { layout: { ...layout, lines: layout.lines.filter(line => keep(line.paragraphIndex)) } })
    },
  }
}

/** Every image index the nodes would paint, in paint order. */
export function imageIndices(nodes: readonly PptxShape[], out: number[] = []): number[] {
  for (const node of nodes) {
    if (node.group) imageIndices(node.children ?? [], out)
    else if (node.imageIndex !== undefined) out.push(node.imageIndex)
  }
  return out
}

// ---------------------------------------------------------------------------
// Painted bounds (slide px), following paintNode's transform chain.
// ---------------------------------------------------------------------------

type M = [number, number, number, number, number, number]
const translate = (m: M, x: number, y: number): M => [m[0], m[1], m[2], m[3], m[4] + m[0] * x + m[2] * y, m[5] + m[1] * x + m[3] * y]
const scale = (m: M, x: number, y: number): M => [m[0] * x, m[1] * x, m[2] * y, m[3] * y, m[4], m[5]]
const rotate = (m: M, rad: number): M => {
  const c = Math.cos(rad), s = Math.sin(rad)
  return [m[0] * c + m[2] * s, m[1] * c + m[3] * s, m[2] * c - m[0] * s, m[3] * c - m[1] * s, m[4], m[5]]
}

export interface Box { minX: number; minY: number; maxX: number; maxY: number }

/** Anti-aliasing slack added to every painted box. */
const AA_PAD = 2

/**
 * The union of the nodes' painted boxes in slide px, padded for stroke, arrow
 * heads, shadow and laid-out text (measured with `measure` when given).
 * Undefined when nothing paints.
 */
export function paintedBounds(nodes: readonly PptxShape[], slide: PptxSlide, resolveFont: FontResolver, measure?: CanvasRenderingContext2D): Box | undefined {
  let box: Box | undefined
  const add = (m: M, l: number, t: number, r: number, b: number): void => {
    for (const [x, y] of [[l, t], [r, t], [r, b], [l, b]]) {
      const px = m[0] * x + m[2] * y + m[4], py = m[1] * x + m[3] * y + m[5]
      if (!Number.isFinite(px) || !Number.isFinite(py)) continue
      if (!box) box = { minX: px, minY: py, maxX: px, maxY: py }
      else { box.minX = Math.min(box.minX, px); box.minY = Math.min(box.minY, py); box.maxX = Math.max(box.maxX, px); box.maxY = Math.max(box.maxY, py) }
    }
  }
  const walk = (list: readonly PptxShape[], parent: M): void => {
    for (const shape of list) {
      const x = emuToPx(shape.xEmu), y = emuToPx(shape.yEmu), w = emuToPx(shape.widthEmu), h = emuToPx(shape.heightEmu)
      if (shape.transformValid === false || ![x, y, w, h, shape.rotationDeg ?? 0].every(Number.isFinite) || w < 0 || h < 0) continue
      let m = translate(parent, x + w / 2, y + h / 2)
      if (shape.rotationDeg) m = rotate(m, shape.rotationDeg * Math.PI / 180)
      if (shape.flipH || shape.flipV) m = scale(m, shape.flipH ? -1 : 1, shape.flipV ? -1 : 1)
      m = translate(m, -w / 2, -h / 2)
      if (shape.group) {
        const g = shape.group
        if (!(g.chExt.width > 0 && g.chExt.height > 0)) continue
        m = translate(scale(m, g.ext.width / g.chExt.width, g.ext.height / g.chExt.height), -emuToPx(g.chOff.x), -emuToPx(g.chOff.y))
        walk(shape.children ?? [], m)
        continue
      }
      leafBox(shape, w, h, m)
    }
  }
  const leafBox = (shape: PptxShape, w: number, h: number, m: M): void => {
    const style = shape.drawingStyle
    const hasGeometry = !!(shape.drawingGeometry ? style?.fill || style?.line || style?.shadow : shape.fill || shape.line)
    const hasImage = shape.imageIndex !== undefined
    if (hasGeometry || hasImage || shape.content || shape.table) {
      const lineWidth = shape.drawingGeometry ? (style?.line ? style.line.width ?? 1 : 0) : shape.line ? Math.max(1, emuToPx(shape.line.widthEmu ?? 12700)) : 0
      // Miter joins and arrow heads reach several line widths past the path.
      let pad = lineWidth * 6
      const sh = style?.shadow
      if (sh && [sh.offsetX, sh.offsetY, sh.blurPx].every(Number.isFinite)) pad += Math.abs(sh.offsetX) + Math.abs(sh.offsetY) + sh.blurPx * 2
      let l = 0, t = 0, r = w, b = h
      const fill = shape.image?.fillRect
      if (hasImage && fill && [fill.l, fill.t, fill.r, fill.b].every(Number.isFinite)) {
        l = Math.min(l, w * fill.l); t = Math.min(t, h * fill.t); r = Math.max(r, w * (1 - fill.r)); b = Math.max(b, h * (1 - fill.b))
      }
      add(m, l - pad, t - pad, r + pad, b + pad)
    }
    const body = shape.textBody
    if (!body || body.paragraphs.length === 0) return
    if (!measure || (body.direction && body.direction !== 'horz') || body.textWarp) { add(m, 0, 0, w, h); return }
    const keep = textFilter.get(shape)
    const layout = layoutTextBody(body, w, h, createTextBodyMeasurer(measure, resolveFont), slide.theme)
    for (const line of layout.lines) {
      if (keep && !keep(line.paragraphIndex)) continue
      let l = Math.min(0, line.x), r = w, slack = 0
      for (const s of line.segments) {
        l = Math.min(l, s.x); r = Math.max(r, s.x + s.width)
        const size = (s.style.fontSizePt ?? 12) * 96 / 72
        const sh = s.style.textShadow
        const shadow = sh ? Math.min(Math.abs(sh.offsetX) || 0, 200) + Math.min(Math.abs(sh.offsetY) || 0, 200) + Math.min(Math.max(0, sh.blurPx || 0), 100) * 2 : 0
        slack = Math.max(slack, size / 2 + shadow + Math.min(s.style.textOutline?.widthPx ?? 0, 100))
      }
      add(m, l - slack, line.y - slack, r + slack, line.y + line.height + slack)
    }
  }
  walk(nodes, [1, 0, 0, 1, 0, 0])
  if (!box) return undefined
  const b: Box = box
  return { minX: b.minX - AA_PAD, minY: b.minY - AA_PAD, maxX: b.maxX + AA_PAD, maxY: b.maxY + AA_PAD }
}

/** True when any node carries text that `paintedBounds` would measure. */
export function hasText(nodes: readonly PptxShape[]): boolean {
  return nodes.some(node => node.group ? hasText(node.children ?? []) : (node.textBody?.paragraphs.length ?? 0) > 0)
}
