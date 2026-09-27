/**
 * Shared "document → canvas paintables" extraction. Single source of truth
 * used by <OfficeDoc> (React) and the pixel-diff/golden harness.
 */
import { collectDocImages, createMeasurer, layoutDocx, renderPages } from '../docx/layout'
import { decodeImage } from '../core/images'
import { computeMetrics, renderSheet } from '../xlsx/render'
import { renderSlide, slideMetrics } from '../pptx/render'
import type { DocxDocument } from '../docx/types'
import type { XlsxDocument } from '../xlsx/types'
import type { PptxDocument } from '../pptx/types'

export interface PageSpec {
  widthPx: number
  heightPx: number
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

/** Collect the canvas-paintable units of a parsed document: pages/sheets/slides. */
export function getPaintables(doc: DocxDocument | XlsxDocument | PptxDocument): Promise<Paintable[]> {
  if ('sections' in doc) {
    return measurerFromDoc().then(async (measure) => {
      const pages = layoutDocx(doc, measure)
      // decode embedded images once; failures degrade to a blank slot
      const decoded = await Promise.all(
        collectDocImages(doc).map((img) => decodeImage(img.data, img.mime).catch(() => undefined)),
      )
      return pages.map((page) => ({
        spec: { widthPx: Math.ceil(page.widthPx), heightPx: Math.ceil(page.heightPx) },
        paint: (ctx) => {
          ctx.fillStyle = '#ffffff'
          ctx.fillRect(0, 0, page.widthPx, page.heightPx)
          renderPages([page], ctx, decoded)
        },
      }))
    })
  }
  if ('sheets' in doc) {
    return Promise.resolve(doc.sheets.map((sheet) => {
      const m = computeMetrics(sheet)
      return {
        spec: { widthPx: m.widthPx, heightPx: m.heightPx },
        paint: (ctx) => renderSheet(sheet, ctx, m),
      }
    }))
  }
  const sm = slideMetrics(doc)
  return Promise.all(doc.images.map((img) => decodeImage(img.data, img.mime).catch(() => undefined))).then((images) =>
    doc.slides.map((slide) => ({
      spec: { widthPx: sm.widthPx, heightPx: sm.heightPx },
      paint: (ctx) => renderSlide(slide, ctx, sm, images),
    })),
  )
}
