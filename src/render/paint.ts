import { acquireFonts, type RegisterFont } from '../core/fonts/register'
import { withFallbackFonts, type FallbackFontsOptions } from '../core/fonts/fallback'
import type { FontDiagnostic } from '../core/fonts/types'
/**
 * Shared "document → canvas paintables" extraction. Single source of truth
 * used by <OfficeDoc> (React) and the pixel-diff/golden harness.
 */
import { collectDocImages, createMeasurer, layoutDocx, renderPages } from '../docx/layout'
import { fontFamilyCss as fontFamilyCssDocx } from '../docx/styles'
import { decodeImageAsset, releaseDecodedImage, type ImageCandidates, type ImageDecodeFn } from '../core/images'
import { normalizeWatermark, type ResolvedWatermark, type WatermarkOptions } from '../core/watermark'
import { computeMetrics, renderSheet } from '../xlsx/render'
import { renderSlide, slideMetrics } from '../pptx/render'
import type { DocxDocument } from '../docx/types'
import type { ContentImageAsset, DrawingContent, DrawingContentShape } from '../drawing/content'
import type { XlsxDocument } from '../xlsx/types'
import type { PptxDocument } from '../pptx/types'

export interface PageSpec {
  widthPx: number
  heightPx: number
}

/** Each extraction owns a lease. Release when its paint closures are no longer used. */
export interface PaintableArray extends Array<Paintable> {
  dispose: () => void
  fontDiagnostics: FontDiagnostic[]
}
function withLease(pages: Paintable[], dispose: () => void = () => {}, fontDiagnostics: FontDiagnostic[] = [], decoded: Array<CanvasImageSource | undefined> = []): PaintableArray {
  let released = false
  return Object.assign(pages, { fontDiagnostics, dispose: () => {
    if (released) return
    released = true
    dispose()
    for (const image of new Set(decoded)) releaseDecodedImage(image)
  } })
}
export interface Paintable {
  spec: PageSpec
  /** Paints one white-background unit into a ctx scaled so 1 unit = 1 spec px. */
  paint: (ctx: CanvasRenderingContext2D) => void
}

async function measurerFromDoc(resolve: (family: string) => string): Promise<ReturnType<typeof createMeasurer>> {
  // headless-safe: DOM canvas when available (browser/jsdom), node-canvas fallback
  let ctx: CanvasRenderingContext2D | null = null
  try {
    if (typeof document !== 'undefined') {
      const c = document.createElement('canvas')
      c.width = 4
      c.height = 4
      ctx = c.getContext('2d')
    }
  } catch {
    // fall through to node-canvas
  }
  if (!ctx) {
    // @vite-ignore: resolved only in node/bun (golden harness), never bundled
    const { createCanvas } = (await import(/* @vite-ignore */ 'canvas')) as unknown as { createCanvas: (w: number, h: number) => { getContext: (t: string) => CanvasRenderingContext2D | null } }
    ctx = createCanvas(4, 4).getContext('2d')
  }
  if (!ctx) throw new Error('no 2d context available for measurement')
  return createMeasurer(ctx, resolve)
}

/** Options for stamping a watermark across every painted unit. */
export interface PaintOptions {
  /** Optional headless font adapter; its identity is the registration cache key. */
  registerFont?: RegisterFont
  /**
   * Applied to every page/sheet/slide. Resolved once here rather than per
   * canvas, so a bad watermark is rejected before any painting happens.
   */
  watermark?: WatermarkOptions | ResolvedWatermark
  /** Optional image decoder override (tests count decodes). Defaults to decodeImage. */
  decodeImage?: ImageDecodeFn
  /**
   * Optional explicit fallback chain/region, composed onto every resolver.
   * Absent by default: resolvers pass through byte-identical, so default
   * rendering (and goldens) never move; pinned chains are host-deterministic.
   */
  fallbackFonts?: FallbackFontsOptions
}

/** Byte equality for shared-asset decode reuse (length pre-checked by callers). */
function bytesEqual(a: Uint8Array | undefined, b: Uint8Array | undefined): boolean {
  if (a === b) return true
  if (!a || !b || a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false
  return true
}

/**
 * Decode one bitmap per COMPLETE candidate pair (raster bytes + SVG bytes).
 * Same PNG with a different SVG is a distinct asset; identical pairs decode
 * once. The selected representation/reason is propagated to every alias's
 * selection record so coverage stays consistent.
 */
async function decodeImageAssets(assets: Array<ImageCandidates & { drawing?: unknown }>, decode?: ImageDecodeFn): Promise<Array<CanvasImageSource | undefined>> {
  // Complete-pair identity: raster bytes + SVG bytes + SVG verdict state +
  // primary-SVG verdict state + fallback availability. Missing/external
  // candidates (byte-less) and ordinary rasters therefore never collapse.
  const svgState = (svg: ImageCandidates['svg']): string => svg ? (svg.verdict.ok ? 'ok' : `rej:${(svg.verdict as { reason: string }).reason}`) : 'none'
  const primaryState = (verdict: ImageCandidates['primarySvgVerdict']): string => verdict ? (verdict.ok ? 'ok' : `rej:${(verdict as { reason: string }).reason}`) : 'none'
  const canonical: Array<{ data?: Uint8Array; svg?: Uint8Array; svgState: string; primaryState: string; svgOnly: boolean; decoded: Promise<CanvasImageSource | undefined> }> = []
  const representative: number[] = []
  for (const asset of assets) {
    if (asset.drawing || (!asset.data && !asset.svg)) { representative.push(-1); continue }
    const state = svgState(asset.svg), primary = primaryState(asset.primarySvgVerdict), svgOnly = asset.hasRaster === false
    const found = canonical.findIndex(entry => entry.svgState === state && entry.primaryState === primary && entry.svgOnly === svgOnly && bytesEqual(entry.data, asset.data) && bytesEqual(entry.svg, asset.svg?.bytes))
    if (found >= 0) { representative.push(found); continue }
    canonical.push({ data: asset.data, svg: asset.svg?.bytes, svgState: state, primaryState: primary, svgOnly, decoded: decodeImageAsset(asset, decode) })
    representative.push(canonical.length - 1)
  }
  const decoded = await Promise.all(assets.map((_, index) => representative[index] < 0 ? undefined : canonical[representative[index]].decoded))
  // Propagate the decoded outcome from each group's representative (the asset
  // whose selection decodeImageAsset mutated) to every alias's selection.
  const groups = new Map<number, number[]>()
  for (let index = 0; index < representative.length; index++) {
    const rep = representative[index]
    if (rep < 0) continue
    const list = groups.get(rep) ?? []
    list.push(index)
    groups.set(rep, list)
  }
  for (const members of groups.values()) {
    const lead = assets[members[0]].imageSelection
    if (!lead) continue
    for (const index of members) {
      const selection = assets[index].imageSelection
      if (selection && selection !== lead) { selection.phase = lead.phase; selection.representation = lead.representation; selection.reason = lead.reason }
    }
  }
  return decoded
}

interface ContentCarrier { content?: DrawingContent<unknown, unknown>; children?: ContentCarrier[] }
/** Append cached pictures after indexed assets so ordinary scene image indices
 * stay stable. Identity deduplication also bounds handwritten cyclic models. */
function collectCachedPictures(nodes: ContentCarrier[]): ContentImageAsset[] {
  const images: ContentImageAsset[] = [], seenImages = new Set<ContentImageAsset>()
  const seenNodes = new Set<ContentCarrier>(), seenShapes = new Set<DrawingContentShape<unknown>>()
  const shapes = (items: DrawingContentShape<unknown>[]) => {
    for (const shape of items) {
      if (seenShapes.has(shape)) continue
      seenShapes.add(shape)
      if (shape.image && !seenImages.has(shape.image)) { seenImages.add(shape.image); images.push(shape.image) }
      if (shape.children) shapes(shape.children)
    }
  }
  const walk = (items: ContentCarrier[]) => {
    for (const node of items) {
      if (seenNodes.has(node)) continue
      seenNodes.add(node)
      if (node.content?.kind === 'diagram') shapes(node.content.shapes)
      if (node.children) walk(node.children)
    }
  }
  walk(nodes)
  return images
}

/** Prepare embedded fonts before measurement/painting. Low-level renderSlide stays synchronous. */
export function getPaintables(
  doc: DocxDocument | XlsxDocument | PptxDocument,
  options?: PaintOptions,
): Promise<PaintableArray> {
  const watermark = normalizeWatermark(options?.watermark as WatermarkOptions | undefined)
  if ('sections' in doc) {
    return acquireFonts(doc, options?.registerFont).then(async (lease) => {
      try {
        const baseFamilyCss = (family: string) => fontFamilyCssDocx(lease.resolve(family))
        const measure = await measurerFromDoc(withFallbackFonts(baseFamilyCss, options?.fallbackFonts))
        const pages = layoutDocx(doc, measure)
        // One decode per unique candidate pair; failures degrade to a blank slot.
        const docImages = collectDocImages(doc)
        const decoded = await decodeImageAssets(docImages, options?.decodeImage)
        const decodedByObject = new Map<ContentImageAsset, CanvasImageSource | undefined>(docImages.map((img, index) => [img, decoded[index]]))
        // Each page paints on its own canvas, so PAGE/NUMPAGES fields must be
        // resolved against the whole document, not the single-page array.
        const total = pages.length
        return withLease(pages.map((page, index) => ({
          spec: { widthPx: Math.ceil(page.widthPx), heightPx: Math.ceil(page.heightPx) },
          paint: (ctx) => {
            ctx.fillStyle = '#ffffff'
            ctx.fillRect(0, 0, page.widthPx, page.heightPx)
            renderPages([page], ctx, decoded, {
              pageNumberStart: index + 1,
              totalPages: total,
              resolveFont: lease.resolve,
              ...(watermark ? { watermark } : {}),
              assets: {
                imageFor: (image) => decodedByObject.get(image),
                ...(options?.fallbackFonts ? { fallbackFonts: options.fallbackFonts } : {}),
                resolveFont: lease.resolve,
              },
            })
          },
        })), lease.dispose, lease.diagnostics, decoded)
      } catch (error) { lease.dispose(); throw error }
    })
  }
  if ('sheets' in doc) {
    return acquireFonts(doc, options?.registerFont).then(async lease => {
      try {
        const cached = collectCachedPictures(doc.sheets.flatMap(sheet => sheet.drawings ?? []))
        const indexed = doc.images ?? []
        const decoded = await decodeImageAssets([...indexed, ...cached], options?.decodeImage)
        const decodedByObject = new Map(cached.map((image, index) => [image, decoded[indexed.length + index]]))
        return withLease(doc.sheets.map((sheet) => {
          const m = computeMetrics(sheet)
          return {
            spec: { widthPx: m.widthPx, heightPx: m.heightPx },
            paint: (ctx) => renderSheet(sheet, ctx, m, watermark, { images: decoded, imageFor: image => decodedByObject.get(image), resolveFont: withFallbackFonts(lease.resolve, options?.fallbackFonts) }),
          }
        }), lease.dispose, lease.diagnostics, decoded)
      } catch (error) { lease.dispose(); throw error }
    })
  }
  return acquireFonts(doc, options?.registerFont).then(async lease => {
    try {
      const sm = slideMetrics(doc)
      const cached = collectCachedPictures(doc.slides.flatMap(slide => slide.shapes))
      const decoded = await decodeImageAssets([...doc.images, ...cached], options?.decodeImage)
      const decodedByObject = new Map(cached.map((image, index) => [image, decoded[doc.images.length + index]]))
      return withLease(doc.slides.map(slide => ({
        spec: { widthPx: sm.widthPx, heightPx: sm.heightPx },
        paint: ctx => renderSlide(slide, ctx, sm, decoded, watermark, withFallbackFonts(lease.resolve, options?.fallbackFonts), { imageFor: image => decodedByObject.get(image) }),
      })), lease.dispose, lease.diagnostics, decoded)
    } catch (error) { lease.dispose(); throw error }
  })
}
