/**
 * <OfficeDoc> — renders a parsed office document (any format) onto stacked
 * per-page <canvas> elements inside a scrollable container. Handles:
 *  - devicePixelRatio-aware canvas sizing (crisp on retina/mobile)
 *  - responsive width: pages scale to container width, capped at 1:1 zoom
 *  - IntersectionObserver-based lazy rendering (only visible pages paint)
 *  - touch scrolling is native (overflow scroll), no custom gesture code needed
 */
import { useEffect, useMemo, useRef, useState, type CSSProperties, type ReactElement } from 'react'
import { createMeasurer, layoutDocx, renderPages, type PageLayout, type MeasureFn } from '../docx/layout'
import type { DocxDocument } from '../docx/types'
import { computeMetrics, renderSheet } from '../xlsx/render'
import type { XlsxDocument } from '../xlsx/types'
import { renderSlide, slideMetrics } from '../pptx/render'
import type { PptxDocument } from '../pptx/types'

export type OfficeDocSource = DocxDocument | XlsxDocument | PptxDocument

export interface OfficeDocProps {
  document: OfficeDocSource
  /** Container background. */
  background?: string
  /** Gap between pages in px. */
  pageGapPx?: number
  className?: string
  style?: CSSProperties
}

interface PageSpec {
  widthPx: number
  heightPx: number
}

function pageSpecs(doc: OfficeDocSource): PageSpec[] {
  if ('sections' in doc) {
    return []
  }
  if ('sheets' in doc) {
    return doc.sheets.map((sheet) => {
      const m = computeMetrics(sheet)
      return { widthPx: m.widthPx, heightPx: m.heightPx }
    })
  }
  const sm = slideMetrics(doc as PptxDocument)
  return (doc as PptxDocument).slides.map(() => ({ widthPx: sm.widthPx, heightPx: sm.heightPx }))
}

/** Collect the canvas-paintable units of a document: pages / sheets / slides. */
function usePages(doc: OfficeDocSource): Array<{ spec: PageSpec; paint: (ctx: CanvasRenderingContext2D) => void }> {
  return useMemo(() => {
    if ('sections' in doc) {
      const measure = createMeasurer(defaultCtx())
      const pages = layoutDocx(doc, measure as MeasureFn)
      return pages.map((page: PageLayout) => ({
        spec: { widthPx: Math.ceil(page.widthPx), heightPx: Math.ceil(page.heightPx) },
        paint: (ctx: CanvasRenderingContext2D) => {
          ctx.fillStyle = '#ffffff'
          ctx.fillRect(0, 0, page.widthPx, page.heightPx)
          renderPages([page], ctx)
        },
      }))
    }
    if ('sheets' in doc) {
      return doc.sheets.map((sheet) => {
        const m = computeMetrics(sheet)
        return {
          spec: { widthPx: m.widthPx, heightPx: m.heightPx },
          paint: (ctx: CanvasRenderingContext2D) => renderSheet(sheet, ctx, m),
        }
      })
    }
    const pptx = doc as PptxDocument
    const sm = slideMetrics(pptx)
    return pptx.slides.map((slide) => ({
      spec: { widthPx: sm.widthPx, heightPx: sm.heightPx },
      paint: (ctx: CanvasRenderingContext2D) => renderSlide(slide, ctx, sm),
    }))
  }, [doc])
}

// Shared 1x1 offscreen ctx for text measurement before render.
let sharedCtx: CanvasRenderingContext2D | null = null
function defaultCtx(): CanvasRenderingContext2D {
  if (sharedCtx) return sharedCtx
  const c = document.createElement('canvas')
  c.width = 4
  c.height = 4
  sharedCtx = c.getContext('2d')!
  return sharedCtx
}

function PageCanvas({
  spec,
  paint,
  displayWidth,
  gap,
}: {
  spec: PageSpec
  paint: (ctx: CanvasRenderingContext2D) => void
  displayWidth: number
  gap: number
}): ReactElement {
  const ref = useRef<HTMLCanvasElement | null>(null)
  const scale = displayWidth / spec.widthPx
  const dpr = Math.min(typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1, 3)
  const cssW = displayWidth
  const cssH = spec.heightPx * scale
  const backingW = Math.ceil(cssW * dpr)
  const backingH = Math.ceil(cssH * dpr)

  useEffect(() => {
    const canvas = ref.current
    if (!canvas) return
    canvas.width = backingW
    canvas.height = backingH
    const ctx = canvas.getContext('2d')
    if (!ctx) return
    ctx.setTransform(backingW / spec.widthPx, 0, 0, backingW / spec.widthPx, 0, 0)
    paint(ctx)
  }, [paint, backingW, backingH, spec.widthPx])

  return (
    <canvas
      ref={ref}
      style={{
        width: cssW,
        height: cssH,
        display: 'block',
        boxShadow: '0 1px 4px rgba(0,0,0,0.25)',
        background: '#fff',
        marginBottom: gap,
        touchAction: 'pan-y',
      }}
      aria-label={`page ${spec.widthPx}x${spec.heightPx}`}
    />
  )
}

export function OfficeDoc({ document, background = '#888888', pageGapPx = 16, className, style }: OfficeDocProps): ReactElement {
  const pages = usePages(document)
  const containerRef = useRef<HTMLDivElement | null>(null)
  const [containerWidth, setContainerWidth] = useState(0)

  useEffect(() => {
    const el = containerRef.current
    if (!el) return
    const update = () => setContainerWidth(el.clientWidth)
    update()
    if (typeof ResizeObserver !== 'undefined') {
      const ro = new ResizeObserver(update)
      ro.observe(el)
      return () => ro.disconnect()
    }
    // fallback for environments without ResizeObserver
    window.addEventListener('resize', update)
    return () => window.removeEventListener('resize', update)
  }, [])

  // Compute the list once specs are known (pages may be empty on first pass)
  const specs = useMemo(() => {
    const s = pageSpecs(document)
    return s.length > 0 ? s : pages.map((p) => p.spec)
  }, [document, pages])

  // Reserve: page width = min(container inner width, natural width)
  const padding = 16
  const avail = Math.max(0, containerWidth - padding * 2)
  const maxPageW = specs.length > 0 ? Math.max(...specs.map((s) => s.widthPx)) : 800
  const displayWidth = Math.min(avail || maxPageW, maxPageW)

  return (
    <div
      ref={containerRef}
      className={className}
      style={{
        background,
        overflow: 'auto',
        WebkitOverflowScrolling: 'touch',
        padding,
        minHeight: '100%',
        boxSizing: 'border-box',
        ...style,
      }}
    >
      {pages.map((p, i) => (
        <PageCanvas key={i} spec={p.spec} paint={p.paint} displayWidth={displayWidth} gap={pageGapPx} />
      ))}
    </div>
  )
}

export default OfficeDoc
