/**
 * <OfficeDoc> — renders a parsed office document (any format) onto stacked
 * per-page <canvas> elements inside a scrollable container. Handles:
 *  - devicePixelRatio-aware canvas sizing (crisp on retina/mobile)
 *  - responsive width: pages scale to container width, capped at 1:1 zoom
 *  - IntersectionObserver-based lazy rendering (only visible pages paint)
 *  - touch scrolling is native (overflow scroll), no custom gesture code needed
 */
import { useEffect, useMemo, useRef, useState, type CSSProperties, type ReactElement } from 'react'
import { getPaintables, type PageSpec, type Paintable } from '../render/paint'
import {
  distance,
  midpoint,
  panTransform,
  pinchTransform,
  stepZoom,
  transformCss,
  zoomAt,
  type Point,
  type Transform,
  type Viewport,
} from '../core/zoom'

export type OfficeDocSource = import('../docx/types').DocxDocument | import('../xlsx/types').XlsxDocument | import('../pptx/types').PptxDocument

export interface OfficeDocProps {
  document: OfficeDocSource
  /** Container background. */
  background?: string
  /** Gap between pages in px. */
  pageGapPx?: number
  className?: string
  style?: CSSProperties
  /** Show the zoom control cluster (default true). */
  showZoomControls?: boolean
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

export function OfficeDoc({
  document,
  background = '#888888',
  pageGapPx = 16,
  className,
  style,
  showZoomControls = true,
}: OfficeDocProps): ReactElement {
  // getPaintables is async (docx measurement resolves a 2D ctx); hold in state.
  const [pages, setPages] = useState<Paintable[]>([])
  useEffect(() => {
    let cancelled = false
    // a new document starts at fit
    setTransform({ zoom: 1, panX: 0, panY: 0 })
    getPaintables(document).then((p) => {
      if (!cancelled) setPages(p)
    })
    return () => {
      cancelled = true
    }
  }, [document])
  const containerRef = useRef<HTMLDivElement | null>(null)
  const [containerWidth, setContainerWidth] = useState(0)
  const [viewportHeight, setViewportHeight] = useState(0)
  const [transform, setTransform] = useState<Transform>({ zoom: 1, panX: 0, panY: 0 })
  // gesture bookkeeping
  const pointers = useRef(new Map<number, Point>())
  const gestureStart = useRef<{ transform: Transform; distance: number; center: Point } | null>(null)
  const panStart = useRef<{ transform: Transform; point: Point } | null>(null)

  useEffect(() => {
    const el = containerRef.current
    if (!el) return
    const update = () => {
      setContainerWidth(el.clientWidth)
      setViewportHeight(el.clientHeight)
    }
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

  // content size at fit scale, used for pan clamping
  const contentWidth = displayWidth || 1
  const contentHeight = useMemo(
    () => specs.reduce((sum, s) => sum + (displayWidth > 0 ? (s.heightPx / s.widthPx) * displayWidth : s.heightPx), 0) + pageGapPx * Math.max(0, specs.length - 1),
    [specs, displayWidth, pageGapPx],
  )

  const viewportRef = useRef<Viewport>({ width: 0, height: 0, contentWidth, contentHeight })
  viewportRef.current = {
    width: containerWidth,
    height: viewportHeight,
    contentWidth,
    contentHeight,
  }

  const localPoint = (e: { clientX: number; clientY: number }): Point => {
    const rect = containerRef.current?.getBoundingClientRect()
    return { x: e.clientX - (rect?.left ?? 0), y: e.clientY - (rect?.top ?? 0) }
  }

  const onPointerDown = (e: React.PointerEvent) => {
    if (e.pointerType === 'mouse' && e.button !== 0) return
    pointers.current.set(e.pointerId, localPoint(e))
    const pts = [...pointers.current.values()]
    if (pts.length === 2) {
      // start a pinch from the current transform
      gestureStart.current = {
        transform,
        distance: Math.max(1, distance(pts[0], pts[1])),
        center: midpoint(pts[0], pts[1]),
      }
      panStart.current = null
    } else if (transform.zoom > 1) {
      // only capture drags once zoomed; at fit the container scrolls natively
      panStart.current = { transform, point: localPoint(e) }
    }
  }

  const onPointerMove = (e: React.PointerEvent) => {
    if (!pointers.current.has(e.pointerId)) return
    pointers.current.set(e.pointerId, localPoint(e))
    const pts = [...pointers.current.values()]
    if (pts.length >= 2 && gestureStart.current) {
      const g = gestureStart.current
      setTransform(
        pinchTransform(g.transform, g.distance, Math.max(1, distance(pts[0], pts[1])), midpoint(pts[0], pts[1]), viewportRef.current),
      )
    } else if (pts.length === 1 && panStart.current) {
      const start = panStart.current
      const now = pts[0]
      setTransform(panTransform(start.transform, { x: now.x - start.point.x, y: now.y - start.point.y }, viewportRef.current))
    }
  }

  const endPointer = (e: React.PointerEvent) => {
    pointers.current.delete(e.pointerId)
    if (pointers.current.size < 2) gestureStart.current = null
    if (pointers.current.size === 0) panStart.current = null
  }

  const onWheel = (e: React.WheelEvent) => {
    // trackpad pinch and ctrl+wheel zoom; plain wheel keeps native scrolling
    if (!e.ctrlKey && !e.metaKey) return
    e.preventDefault()
    const factor = Math.exp(-e.deltaY / 250)
    setTransform((t) => zoomAt(t, t.zoom * factor, localPoint(e), viewportRef.current))
  }

  const onDoubleClick = (e: React.MouseEvent) => {
    const at = localPoint(e)
    setTransform((t) => (t.zoom > 1 ? { zoom: 1, panX: 0, panY: 0 } : zoomAt(t, 2, at, viewportRef.current)))
  }

  const zoomBy = (direction: 1 | -1) => {
    setTransform((t) => {
      const next = stepZoom(t.zoom, direction)
      const anchor = { x: viewportRef.current.width / 2, y: viewportRef.current.height / 2 }
      return next <= 1 ? { zoom: 1, panX: 0, panY: 0 } : zoomAt(t, next, anchor, viewportRef.current)
    })
  }

  const reset = () => setTransform({ zoom: 1, panX: 0, panY: 0 })

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === '+' || e.key === '=') { e.preventDefault(); zoomBy(1) }
    else if (e.key === '-') { e.preventDefault(); zoomBy(-1) }
    else if (e.key === '0') { e.preventDefault(); reset() }
  }

  const zoomed = transform.zoom > 1

  return (
    <div style={{ position: 'relative', display: 'flex', flexDirection: 'column', minHeight: 0 }}>
      <div
        ref={containerRef}
        className={className}
        role="region"
        aria-label="Document viewer"
        tabIndex={0}
        onKeyDown={onKeyDown}
        onDoubleClick={onDoubleClick}
        onWheel={onWheel}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={endPointer}
        onPointerCancel={endPointer}
        style={{
          background,
          // at fit we let the browser do momentum scrolling; when zoomed we
          // own the pan and must not scroll the container
          overflow: zoomed ? 'hidden' : 'auto',
          WebkitOverflowScrolling: 'touch',
          touchAction: zoomed ? 'none' : 'pan-y',
          padding,
          minHeight: 0,
          flex: '1 1 auto',
          position: 'relative',
          boxSizing: 'border-box',
          cursor: zoomed ? 'grab' : 'auto',
          ...style,
        }}
      >
        <div
          data-testid="officeview-content"
          style={{
            transform: transformCss(transform),
            transformOrigin: '0 0',
            width: zoomed ? contentWidth * transform.zoom : undefined,
            willChange: transformCss(transform) === 'none' ? undefined : 'transform',
          }}
        >
          {pages.map((p, i) => (
            <PageCanvas key={i} spec={p.spec} paint={p.paint} displayWidth={displayWidth} gap={pageGapPx} />
          ))}
        </div>
      </div>
      {showZoomControls && (
        <ZoomControls
          zoom={transform.zoom}
          onIn={() => zoomBy(1)}
          onOut={() => zoomBy(-1)}
          onReset={reset}
          atFit={!zoomed}
        />
      )}
    </div>
  )
}

function ZoomControls({
  zoom,
  onIn,
  onOut,
  onReset,
  atFit,
}: {
  zoom: number
  onIn: () => void
  onOut: () => void
  onReset: () => void
  atFit: boolean
}): ReactElement {
  const btn: CSSProperties = {
    width: 36,
    height: 36,
    borderRadius: 18,
    border: '1px solid rgba(0,0,0,0.15)',
    background: '#fff',
    color: '#111',
    fontSize: 18,
    lineHeight: 1,
    cursor: 'pointer',
    boxShadow: '0 1px 4px rgba(0,0,0,0.25)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  }
  return (
    <div
      data-testid="officeview-zoom-controls"
      style={{
        position: 'absolute',
        right: 12,
        bottom: 'calc(env(safe-area-inset-bottom) + 12px)',
        display: 'flex',
        flexDirection: 'column',
        gap: 8,
        zIndex: 5,
      }}
    >
      <button type="button" onClick={onIn} aria-label="Zoom in" style={btn}>+</button>
      <button type="button" onClick={onOut} aria-label="Zoom out" disabled={atFit} style={{ ...btn, opacity: atFit ? 0.5 : 1 }}>−</button>
      <button
        type="button"
        onClick={onReset}
        aria-label="Reset zoom"
        style={{ ...btn, width: 36, height: 28, fontSize: 11, fontWeight: 600 }}
      >
        {Math.round(zoom * 100)}%
      </button>
    </div>
  )
}

export default OfficeDoc
