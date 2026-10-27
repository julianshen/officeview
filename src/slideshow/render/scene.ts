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
import { layoutTextBody, type MeasureText, type TextLayout } from '../../drawing/text-layout'
import { createTextBodyMeasurer, paintTextBody } from '../../drawing/text-paint'
import type { FontResolver } from '../../core/fonts/register'
import { emuToPx } from '../../core/geometry'
import { renderSlide } from '../../pptx/render'
import type { PptxShape, PptxSlide } from '../../pptx/types'
import type { TargetKey } from '../types'
import { boxesFor, type TextItem } from './text-split'

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

/**
 * Text clips for planner-made word and letter clones (S2-8): a clone listed
 * here paints its text only inside the boxes of item `key` among `items`,
 * placed by `boxesFor` on the layout it paints with (so the clip follows the
 * painted glyphs), and its painted bounds are cut to them.
 */
export const textClip = new WeakMap<PptxShape, { items: readonly TextItem[]; key: TargetKey }>()

/** The clip rects of a clone on `layout` (shape-local px). */
function clipRects(shape: PptxShape, layout: TextLayout, clip: { items: readonly TextItem[]; key: TargetKey }, measure: MeasureText): Array<{ x: number; y: number; w: number; h: number }> {
  return boxesFor(shape, layout, clip.items, measure).filter(b => b.key === clip.key).map(b => b.box)
}

const TRANSPARENT = 'rgba(0,0,0,0)'
const PX_TO_EMU = 914400 / 96

export interface SceneAssets {
  images: ReadonlyArray<CanvasImageSource | undefined>
  resolveFont: FontResolver
  contentAssets?: ContentPaintAssets
  /** Return true when the image was withheld on purpose (the caller diagnoses it); the static diagnostic is then skipped. */
  withheld?: (shape: PptxShape) => boolean
}

/**
 * The `renderSlide` scene options, with paragraph filtering for clones.
 * `root` is the context's transform at slide px (0, 0): `useBgFill` shapes paint the slide's backdrop through it.
 */
export function sceneOptions(slide: PptxSlide, assets: SceneAssets, root?: DOMMatrix): ScenePaintOptions<PptxShape> {
  const { resolveFont } = assets
  return {
    images: assets.images,
    pixelSnap: true,
    ...(slide.backdrop && root ? { backdrop: [slide.backdrop], root } : {}),
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
      const clip = textClip.get(shape)
      if (!keep && !clip) { paintTextBody(body, ctx, 0, 0, w, h, resolveFont, slide.theme); return }
      const measure = createTextBodyMeasurer(ctx, resolveFont)
      const layout = layoutTextBody(body, w, h, measure, slide.theme)
      const kept = { ...layout, lines: keep ? layout.lines.filter(line => keep(line.paragraphIndex)) : layout.lines }
      if (!clip) { paintTextBody(body, ctx, 0, 0, w, h, resolveFont, slide.theme, { layout: kept }); return }
      const rects = clipRects(shape, layout, clip, measure)
      if (rects.length === 0) return
      ctx.save()
      try {
        ctx.beginPath()
        for (const r of rects) ctx.rect(r.x, r.y, r.w, r.h)
        ctx.clip()
        paintTextBody(body, ctx, 0, 0, w, h, resolveFont, slide.theme, { layout: kept })
      } finally {
        ctx.restore()
      }
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

/** A box accumulator: `add` grows it by a local rect's corners under `m`. */
function boxUnion(): { add(m: M, l: number, t: number, r: number, b: number): void; box(): Box | undefined } {
  let box: Box | undefined
  return {
    add(m, l, t, r, b) {
      for (const [x, y] of [[l, t], [r, t], [r, b], [l, b]]) {
        const px = m[0] * x + m[2] * y + m[4], py = m[1] * x + m[3] * y + m[5]
        if (!Number.isFinite(px) || !Number.isFinite(py)) continue
        if (!box) box = { minX: px, minY: py, maxX: px, maxY: py }
        else { box.minX = Math.min(box.minX, px); box.minY = Math.min(box.minY, py); box.maxX = Math.max(box.maxX, px); box.maxY = Math.max(box.maxY, py) }
      }
    },
    box: () => box,
  }
}

/** Every leaf (non-group) node with its local size and paintNode's local-to-slide matrix. */
function walkLeaves(list: readonly PptxShape[], parent: M, visit: (shape: PptxShape, w: number, h: number, m: M) => void): void {
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
      walkLeaves(shape.children ?? [], m, visit)
      continue
    }
    visit(shape, w, h, m)
  }
}

/** A line's em in slide px: its largest font size (an empty line: its 1.2 em height). */
const lineEm = (line: TextLayout['lines'][number]): number =>
  line.segments.length ? Math.max(...line.segments.map(s => (s.style.fontSizePt ?? 12) * 96 / 72)) : line.height / 1.2

/**
 * The union of the nodes' painted boxes in slide px, padded for stroke, arrow
 * heads, shadow and laid-out text (measured with `measure` when given).
 * Undefined when nothing paints.
 */
export function paintedBounds(nodes: readonly PptxShape[], slide: PptxSlide, resolveFont: FontResolver, measure?: CanvasRenderingContext2D): Box | undefined {
  const { add, box: union } = boxUnion()
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
    const measurer = createTextBodyMeasurer(measure, resolveFont)
    const layout = layoutTextBody(body, w, h, measurer, slide.theme)
    const clipped = textClip.get(shape)
    let clip: { x: number; y: number; w: number; h: number } | undefined
    if (clipped) {
      // The bounding box of the item's boxes (a word broken across lines has several).
      const rects = clipRects(shape, layout, clipped, measurer)
      if (rects.length === 0) return
      const x0 = Math.min(...rects.map(r => r.x)), y0 = Math.min(...rects.map(r => r.y))
      clip = { x: x0, y: y0, w: Math.max(...rects.map(r => r.x + r.w)) - x0, h: Math.max(...rects.map(r => r.y + r.h)) - y0 }
    }
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
      let t = line.y - slack, b = line.y + line.height + slack
      if (clip) {
        // A word or letter clone paints only inside its clip.
        l = Math.max(l - slack, clip.x); r = Math.min(r + slack, clip.x + clip.w)
        t = Math.max(t, clip.y); b = Math.min(b, clip.y + clip.h)
        if (!(r > l && b > t)) continue
        add(m, l, t, r, b)
        continue
      }
      add(m, l - slack, t, r + slack, b)
    }
  }
  walkLeaves(nodes, [1, 0, 0, 1, 0, 0], leafBox)
  const b = union()
  if (!b) return undefined
  return { minX: b.minX - AA_PAD, minY: b.minY - AA_PAD, maxX: b.maxX + AA_PAD, maxY: b.maxY + AA_PAD }
}

/**
 * The box PowerPoint sweeps a text-range effect over (a paragraph build under blinds, wipe, …), in slide px:
 * from 1.3 em above the first laid-out baseline to 0.9 em below the last, across the shape's width and the
 * lines' extent, em being each line's largest font size. Fit to PowerPoint's blinds frames (lo-tdf93868:
 * two 20 pt lines of Calibri to within a pixel, one 40 pt line to about two); the horizontal extent is unmeasured
 * and keeps the painted bounds' em/2 slack. Word and letter clones (`textClip`) are left out: their sweep box is
 * unmeasured and stays the painted bounds. So is anything painted past the box (a shadow, an outline, a
 * highlight's line band, an inline object): a mask clips outside its box even when complete, which would cut it
 * off for good. Undefined in all those cases.
 */
export function textRangeBox(nodes: readonly PptxShape[], slide: PptxSlide, resolveFont: FontResolver, measure: CanvasRenderingContext2D): Box | undefined {
  const { add, box } = boxUnion()
  const measurer = createTextBodyMeasurer(measure, resolveFont)
  let clips = false
  walkLeaves(nodes, [1, 0, 0, 1, 0, 0], (shape, w, h, m) => {
    const body = shape.textBody
    if (!body || body.paragraphs.length === 0 || (body.direction && body.direction !== 'horz') || body.textWarp || textClip.has(shape)) return
    const keep = textFilter.get(shape)
    const lines = layoutTextBody(body, w, h, measurer, slide.theme).lines.filter(line => !keep || keep(line.paragraphIndex))
    if (lines.length === 0) return
    const first = lines[0], last = lines[lines.length - 1]
    const top = first.baseline - 1.3 * lineEm(first), bottom = last.baseline + 0.9 * lineEm(last)
    let l = 0, r = w, slack = 0
    for (const line of lines) {
      slack = Math.max(slack, lineEm(line) / 2)
      l = Math.min(l, line.x)
      for (const s of line.segments) {
        l = Math.min(l, s.x); r = Math.max(r, s.x + s.width)
        // How far the segment's shadow and outline reach past its line box, as paintedBounds measures them.
        const sh = s.style.textShadow
        const reach = (sh ? Math.min(Math.abs(sh.offsetX) || 0, 200) + Math.min(Math.abs(sh.offsetY) || 0, 200) + Math.min(Math.max(0, sh.blurPx || 0), 100) * 2 : 0)
          + Math.min(s.style.textOutline?.widthPx ?? 0, 100)
        if (reach > 0 && (line.y - reach < top || line.y + line.height + reach > bottom || reach > lineEm(line) / 2)) clips = true
        // A highlight fills the whole line band, which wide line spacing takes above the box.
        if (s.style.highlight && s.text !== '\n' && (line.y < top || line.y + line.height > bottom)) clips = true
      }
      for (const slot of line.inlineSlots ?? []) if (slot.y < top || slot.y + slot.height > bottom) clips = true
    }
    add(m, l - slack, top, r + slack, bottom)
  })
  return clips ? undefined : box()
}

/** True when any node carries text that `paintedBounds` would measure. */
export function hasText(nodes: readonly PptxShape[]): boolean {
  return nodes.some(node => node.group ? hasText(node.children ?? []) : (node.textBody?.paragraphs.length ?? 0) > 0)
}
