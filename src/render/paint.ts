import { acquireFonts, type RegisterFont } from '../core/fonts/register'
import type { FontDiagnostic } from '../core/fonts/types'
/**
 * Shared "document → canvas paintables" extraction. Single source of truth
 * used by <OfficeDoc> (React) and the pixel-diff/golden harness.
 */
import { collectDocImages, createMeasurer, layoutDocx, renderPages } from '../docx/layout'
import { decodeImage } from '../core/images'
import { normalizeWatermark, type ResolvedWatermark, type WatermarkOptions } from '../core/watermark'
import { computeMetrics, renderSheet } from '../xlsx/render'
import { renderSlide, slideMetrics } from '../pptx/render'
import type { DocxDocument } from '../docx/types'
import type { ContentImageAsset } from '../drawing/content'
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
function withLease(pages: Paintable[], dispose: () => void = () => {}, fontDiagnostics: FontDiagnostic[] = []): PaintableArray {
  let released = false
  return Object.assign(pages, { fontDiagnostics, dispose: () => { if (!released) { released = true; dispose() } } })
}
export interface Paintable {
  spec: PageSpec
  /** Paints one white-background unit into a ctx scaled so 1 unit = 1 spec px. */
  paint: (ctx: CanvasRenderingContext2D) => void
}

async function measurerFromDoc(): Promise<ReturnType<typeof createMeasurer>> {
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
  return createMeasurer(ctx)
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
  decodeImage?: (bytes: Uint8Array, mime?: string) => Promise<CanvasImageSource | undefined>
}

/** Byte equality for shared-asset decode reuse (length pre-checked by callers). */
function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  return a.length === b.length && a.every((value, index) => value === b[index])
}

/** Prepare embedded fonts before measurement/painting. Low-level renderSlide stays synchronous. */
export function getPaintables(
  doc: DocxDocument | XlsxDocument | PptxDocument,
  options?: PaintOptions,
): Promise<PaintableArray> {
  const watermark = normalizeWatermark(options?.watermark as WatermarkOptions | undefined)
  if ('sections' in doc) {
    return measurerFromDoc().then(async (measure) => {
      const pages = layoutDocx(doc, measure)
      // decode embedded images once per unique asset; failures degrade to a blank slot.
      // Distinct wrappers may reference identical bytes (reused raster part):
      // they keep stable per-object indices while sharing one decode.
      const decode = options?.decodeImage ?? decodeImage
      const docImages = collectDocImages(doc)
      const canonical: Array<{ data: Uint8Array; decoded: Promise<CanvasImageSource | undefined> }> = []
      const decoded = await Promise.all(
        docImages.map((img) => {
          if (img.drawing) return undefined
          const found = canonical.find((entry) => entry.data.byteLength === img.data.byteLength && bytesEqual(entry.data, img.data))
          if (found) return found.decoded
          const decoded = decode(img.data, img.mime).catch(() => undefined)
          canonical.push({ data: img.data, decoded })
          return decoded
        })
      )
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
            ...(watermark ? { watermark } : {}),
            assets: { imageFor: (image) => decodedByObject.get(image) },
          })
        },
      })))
    })
  }
  if ('sheets' in doc) {
    return acquireFonts(doc, options?.registerFont).then(async lease => {
      try {
        const images = await Promise.all((doc.images ?? []).map(img => decodeImage(img.data, img.mime).catch(() => undefined)))
        return withLease(doc.sheets.map((sheet) => {
          const m = computeMetrics(sheet)
          return {
            spec: { widthPx: m.widthPx, heightPx: m.heightPx },
            paint: (ctx) => renderSheet(sheet, ctx, m, watermark, { images, resolveFont: lease.resolve }),
          }
        }), lease.dispose, lease.diagnostics)
      } catch (error) { lease.dispose(); throw error }
    })
  }
  return acquireFonts(doc, options?.registerFont).then(async lease => {
    try {
      const sm = slideMetrics(doc)
      const images = await Promise.all(doc.images.map(img => decodeImage(img.data, img.mime).catch(() => undefined)))
      return withLease(doc.slides.map(slide => ({
        spec: { widthPx: sm.widthPx, heightPx: sm.heightPx },
        paint: ctx => renderSlide(slide, ctx, sm, images, watermark, lease.resolve),
      })), lease.dispose, lease.diagnostics)
    } catch (error) { lease.dispose(); throw error }
  })
}
