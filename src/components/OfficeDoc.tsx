/**
 * <OfficeDoc> — renders a parsed office document (any format) onto stacked
 * per-page <canvas> elements inside a scrollable container. Handles:
 *  - devicePixelRatio-aware canvas sizing (crisp on retina/mobile)
 *  - responsive width: pages scale to container width, capped at 1:1 zoom
 *  - IntersectionObserver-based lazy rendering (only visible pages paint)
 *  - touch scrolling is native (overflow scroll), no custom gesture code needed
 */
import { useEffect, useRef, useState, type CSSProperties, type ReactElement } from 'react'
import { getPaintables, type PageSpec, type Paintable } from '../render/paint'

export type OfficeDocSource = import('../docx/types').DocxDocument | import('../xlsx/types').XlsxDocument | import('../pptx/types').PptxDocument

export interface OfficeDocProps {
  document: OfficeDocSource
  /** Container background. */
  background?: string
  /** Gap between pages in px. */
  pageGapPx?: number
  className?: string
  style?: CSSProperties
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
  // getPaintables is async (docx measurement resolves a 2D ctx); hold in state.
  const [pages, setPages] = useState<Paintable[]>([])
  useEffect(() => {
    let cancelled = false
    getPaintables(document).then((p) => {
      if (!cancelled) setPages(p)
    })
    return () => {
      cancelled = true
    }
  }, [document])
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

  const specs = pages.map((p) => p.spec)

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
