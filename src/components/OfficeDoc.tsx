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
import { buildTextIndex, findMatches, stepMatch, type SearchMatch, type TextIndex } from '../core/search'
import {
  hitTest,
  isEmptyRange,
  lineRangeAt,
  normalizeRange,
  rectsForSelectionOnPage,
  textForRange,
  wordRangeAt,
  type CaretPos,
  type SelectionRange,
} from '../core/selection'
import type { PageOverlay } from '../core/overlay'
import { copyText } from '../core/clipboard'
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
  /** Show the in-viewer search bar (default true). */
  showSearch?: boolean
  /**
   * Allow the reader to select document text and copy it (default true).
   * Setting this to false removes text selection, the copy button and the
   * Cmd/Ctrl+C handler, and suppresses native copy of the container.
   *
   * This is a deterrent, not DRM: the document is drawn to a canvas, so the
   * browser holds no selectable text for it, but a reader can still capture
   * pixels via screenshots or devtools.
   */
  allowCopy?: boolean
  /**
   * Allow the document to appear in the browser's print output (default true).
   * Setting this to false hides the viewer with a print-only stylesheet and
   * suppresses the Cmd/Ctrl+P shortcut while the viewer has focus.
   *
   * This is likewise a deterrent: it stops the browser's own print pipeline
   * and a careless Cmd+P, not a determined user with a screenshot.
   */
  allowPrint?: boolean
}

/** Cap the effective device scale so a 6x zoom can't allocate absurd canvases. */
const MAX_EFFECTIVE_SCALE = 3

export interface HighlightRect {
  x: number
  y: number
  width: number
  height: number
}

interface SearchOverlay extends PageOverlay {
  /** Alias kept for the paint call sites. */
  rects: HighlightRect[]
  active: HighlightRect[]
}

/**
 * Draw search highlights and the text selection over the page content. Called
 * while the ctx still carries the page transform, so rectangles are in page
 * coordinates.
 */
function paintHighlights(ctx: CanvasRenderingContext2D, overlay: SearchOverlay): void {
  ctx.save()
  ctx.fillStyle = 'rgba(255, 214, 0, 0.45)'
  for (const r of overlay.rects) ctx.fillRect(r.x, r.y, r.width, r.height)
  // selection sits on top of the search wash, as it does in a browser
  ctx.fillStyle = 'rgba(64, 128, 255, 0.35)'
  for (const r of overlay.selection) ctx.fillRect(r.x, r.y, r.width, r.height)
  ctx.strokeStyle = '#ff8c00'
  ctx.lineWidth = 1.5
  for (const r of overlay.active) ctx.strokeRect(r.x, r.y, r.width, r.height)
  ctx.restore()
}

function PageCanvas({
  spec,
  paint,
  displayWidth,
  gap,
  baseScale,
  boostedScale,
  boosted,
  rootRef,
  overlay,
}: {
  spec: PageSpec
  paint: (ctx: CanvasRenderingContext2D) => void
  displayWidth: number
  gap: number
  /** device scale at fit (devicePixelRatio, capped). */
  baseScale: number
  /** device scale to use while zoomed and visible. */
  boostedScale: number
  /** whether the viewer is currently zoomed. */
  boosted: boolean
  /** scroll container used as the IntersectionObserver root. */
  rootRef?: React.RefObject<HTMLElement | null>
  /** search highlights drawn on top of the page content. */
  overlay?: SearchOverlay
}): ReactElement {
  const ref = useRef<HTMLCanvasElement | null>(null)
  const [visible, setVisible] = useState(true)

  // Only pages the user can actually see get the high-resolution backing
  // store; off-screen pages stay cheap so long documents remain affordable.
  useEffect(() => {
    if (!boosted) {
      setVisible(true) // no need to observe while at fit
      return
    }
    const el = ref.current
    if (!el || typeof IntersectionObserver === 'undefined') {
      setVisible(true)
      return
    }
    // start pessimistic: the observer's first callback decides
    setVisible(false)
    const root = rootRef?.current ?? null
    const io = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) setVisible(entry.isIntersecting)
      },
      { root, rootMargin: '50% 0px' },
    )
    io.observe(el)
    return () => io.disconnect()
  }, [boosted, rootRef])

  const scale = displayWidth / spec.widthPx
  // at fit (or off-screen while zoomed) use the base scale; zoomed + visible
  // pages render at the zoom-aware scale so CSS upscaling stays sharp
  const effective = boosted && visible ? Math.min(boostedScale, MAX_EFFECTIVE_SCALE) : baseScale
  const cssW = displayWidth
  const cssH = spec.heightPx * scale
  const backingW = Math.ceil(cssW * effective)
  const backingH = Math.ceil(cssH * effective)

  useEffect(() => {
    const canvas = ref.current
    if (!canvas) return
    canvas.width = backingW
    canvas.height = backingH
    const ctx = canvas.getContext('2d')
    if (!ctx) return
    ctx.setTransform(backingW / spec.widthPx, 0, 0, backingW / spec.widthPx, 0, 0)
    paint(ctx)
    if (overlay) paintHighlights(ctx, overlay)
    // overlay.key changes whenever the match set or active match changes
  }, [paint, backingW, backingH, spec.widthPx, overlay?.key])

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
  showSearch = true,
  allowCopy = true,
  allowPrint = true,
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

  // Swallow a native copy / context-menu gesture while copy prevention is on.
  // The canvas has no selectable text of its own, so this only stops the
  // browser from copying whatever page chrome happens to sit around it.
  const preventNativeCopy = (e: React.ClipboardEvent | React.MouseEvent): void => {
    e.preventDefault()
  }

  const onPointerDown = (e: React.PointerEvent) => {
    if (e.pointerType === 'mouse' && e.button !== 0) return
    const point = localPoint(e)
    pointers.current.set(e.pointerId, point)
    const pts = [...pointers.current.values()]
    if (pts.length === 2) {
      // start a pinch from the current transform
      gestureStart.current = {
        transform,
        distance: Math.max(1, distance(pts[0], pts[1])),
        center: midpoint(pts[0], pts[1]),
      }
      panStart.current = null
      selectionDrag.current = null
    } else if (transform.zoom > 1) {
      // only capture drags once zoomed; at fit the container scrolls natively
      panStart.current = { transform, point }
      dragStart.current = point
      selectionDrag.current = null
    } else {
      // at fit, dragging is text selection
      if (!allowCopy) return
      selectionDrag.current = { origin: point, moved: false }
      selectionStartRef.current = { clientX: e.clientX, clientY: e.clientY, target: e.target }
      void caretFromEvent(e).then((caret) => {
        selectionAnchor.current = caret ?? null
        if (caret) {
          // a plain click clears any existing selection
          setSelection({
            start: caret,
            end: caret,
          })
        }
      })
    }
  }

  const onPointerMove = (e: React.PointerEvent) => {
    if (!pointers.current.has(e.pointerId)) return
    const point = localPoint(e)
    pointers.current.set(e.pointerId, point)
    const pts = [...pointers.current.values()]
    if (pts.length >= 2 && gestureStart.current) {
      const g = gestureStart.current
      setTransform(
        pinchTransform(g.transform, g.distance, Math.max(1, distance(pts[0], pts[1])), midpoint(pts[0], pts[1]), viewportRef.current),
      )
      return
    }
    // a drag past the threshold while zoomed pans instead of selecting
    if (pts.length === 1 && panStart.current) {
      if (dragStart.current && Math.hypot(point.x - dragStart.current.x, point.y - dragStart.current.y) > 6) {
        selectionDrag.current = null
      }
      const start = panStart.current
      setTransform(panTransform(start.transform, { x: point.x - start.point.x, y: point.y - start.point.y }, viewportRef.current))
      return
    }
    if (pts.length === 1 && selectionDrag.current && !selectionDrag.current.moved) {
      if (Math.hypot(point.x - selectionDrag.current.origin.x, point.y - selectionDrag.current.origin.y) > 3) {
        selectionDrag.current.moved = true
      }
    }
    if (selectionDrag.current?.moved) {
      extendSelectionTo(e)
    }
  }

  const endPointer = (e: React.PointerEvent) => {
    pointers.current.delete(e.pointerId)
    if (pointers.current.size < 2) gestureStart.current = null
    if (pointers.current.size === 0) {
      panStart.current = null
      selectionDrag.current = null
      dragStart.current = null
    }
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
    // double-click on text selects a word; on blank space it zooms, so both
    // behaviours stay available
    void caretFromEvent(e, true).then((caret) => {
      if (caret) {
        selectWordAt(caret)
        return
      }
      setTransform((t) => (t.zoom > 1 ? { zoom: 1, panX: 0, panY: 0 } : zoomAt(t, 2, at, viewportRef.current)))
    })
  }

  const onClick = (e: React.MouseEvent) => {
    if (e.detail < 3) return // double-click already handled words
    void caretFromEvent(e).then((caret) => caret && selectLineAt(caret))
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
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'f') {
      e.preventDefault()
      inputRef.current?.focus()
      inputRef.current?.select()
      return
    }
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'c') {
      // when copying is off, swallow the shortcut rather than letting the
      // browser act on whatever happens to be selected around the canvas
      if (!allowCopy) {
        e.preventDefault()
        return
      }
      if (selection && !isEmptyRange(selection)) {
        e.preventDefault()
        void doCopy()
        return
      }
    }
    if (!allowPrint && (e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'p') {
      e.preventDefault()
      return
    }
    if (e.key === 'Escape' && selection) {
      setSelection(undefined)
      selectionAnchor.current = null
      return
    }
    if (e.key === '+' || e.key === '=') { e.preventDefault(); zoomBy(1) }
    else if (e.key === '-') { e.preventDefault(); zoomBy(-1) }
    else if (e.key === '0') { e.preventDefault(); reset() }
  }

  const zoomed = transform.zoom > 1
  const dpr = Math.min(typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1, 3)

  // ---- search ----
  const [query, setQuery] = useState('')
  const [matches, setMatches] = useState<SearchMatch[]>([])
  const [active, setActive] = useState(-1)
  const [indexing, setIndexing] = useState(false)
  const indexRef = useRef<TextIndex | null>(null)
  const inputRef = useRef<HTMLInputElement | null>(null)

  // The index replays each page's paint once; built lazily on first search.
  const ensureIndex = async (): Promise<TextIndex> => {
    if (indexRef.current) return indexRef.current
    setIndexing(true)
    try {
      const paintables = pages.length > 0 ? pages : await getPaintables(document)
      const built = await buildTextIndex(paintables as Paintable[])
      indexRef.current = built
      return built
    } finally {
      setIndexing(false)
    }
  }

  useEffect(() => {
    indexRef.current = null // a new document needs a fresh index
    setMatches([])
    setActive(-1)
  }, [document])

  // Print suppression. The document is drawn to canvases, so it reaches the
  // print pipeline as pixels; hiding the viewer with a print-only rule is what
  // actually removes it from a printout. The rule is installed only while a
  // viewer with allowPrint={false} is mounted, and removed on unmount so a
  // later print of the host page is unaffected. `document` is shadowed by the
  // prop of the same name, hence globalThis.
  useEffect(() => {
    if (allowPrint) return
    const doc = globalThis.document
    if (!doc) return
    const style = doc.createElement('style')
    style.setAttribute('data-officeview-print-guard', '')
    style.textContent = '@media print { [data-officeview-root] { display: none !important; } }'
    doc.head.appendChild(style)
    return () => {
      style.remove()
    }
  }, [allowPrint])

  useEffect(() => {
    if (query.trim().length === 0) {
      setMatches([])
      setActive(-1)
      return
    }
    let cancelled = false
    ensureIndex().then((index) => {
      if (cancelled) return
      const found = findMatches(index, query)
      setMatches(found)
      setActive(found.length > 0 ? 0 : -1)
    })
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [query, pages])

  const goTo = (direction: 1 | -1) => {
    setActive((current) => stepMatch(matches, current, direction))
  }

  // bring the active match's page into view
  useEffect(() => {
    const match = matches[active]
    if (!match) return
    const container = containerRef.current
    if (!container) return
    const canvas = container.querySelectorAll('canvas')[match.pageIndex] as HTMLElement | undefined
    if (!canvas) return

    // "visible" means inside the region's box *and* the viewport — the region
    // itself may not be the scrolling element (it can grow to fit all pages).
    const regionRect = container.getBoundingClientRect?.()
    const canvasRect = canvas.getBoundingClientRect?.()
    const viewportHeight = typeof window !== 'undefined' ? window.innerHeight : Infinity
    const viewTop = Math.max(regionRect?.top ?? 0, 0)
    const viewBottom = Math.min(regionRect?.bottom ?? viewportHeight, viewportHeight)
    const pageVisible =
      !!canvasRect && canvasRect.height > 0 && canvasRect.top >= viewTop && canvasRect.bottom <= viewBottom

    if (!pageVisible) {
      if (typeof canvas.scrollIntoView === 'function') {
        // scrolls whichever ancestor actually scrolls
        canvas.scrollIntoView({ block: 'start' })
      } else {
        const target = Math.max(0, canvas.offsetTop - padding)
        if (typeof container.scrollTo === 'function') {
          container.scrollTo({ top: target, behavior: 'smooth' })
        } else {
          container.scrollTop = target
        }
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [active, matches])

  // ---- text selection ----
  const [selection, setSelection] = useState<SelectionRange | undefined>(undefined)
  const selectionKey = selection
    ? `${selection.start.pageIndex}:${selection.start.lineIndex}:${selection.start.charIndex}` +
      `-${selection.end.pageIndex}:${selection.end.lineIndex}:${selection.end.charIndex}`
    : ''
  const [copied, setCopied] = useState<'idle' | 'done' | 'failed'>('idle')
  const selectionAnchor = useRef<CaretPos | null>(null)
  const selectionStartRef = useRef<{ clientX: number; clientY: number; target: EventTarget | null } | null>(null)
  const selectionDrag = useRef<{ origin: { x: number; y: number }; moved: boolean } | null>(null)
  const dragStart = useRef<{ x: number; y: number } | null>(null)

  /** Resolve a pointer/mouse event to a caret, mapping page pixels into the
   * document coordinate space (the canvas is displayed scaled to fit). */
  const caretFromEvent = async (e: { clientX: number; clientY: number; target: EventTarget | null }, strict = false): Promise<CaretPos | undefined> => {
    const container = containerRef.current
    const target = e.target
    if (!container || !(target instanceof HTMLCanvasElement)) return undefined
    const pageIndex = Array.from(container.querySelectorAll('canvas')).indexOf(target)
    if (pageIndex < 0) return undefined
    const index = await ensureIndex()
    const rect = target.getBoundingClientRect()
    const natural = pages[pageIndex]?.spec.widthPx ?? rect.width
    const scale = rect.width > 0 ? natural / rect.width : 1
    return hitTest(index, pageIndex, (e.clientX - rect.left) * scale, (e.clientY - rect.top) * scale, { strict })
  }

  const extendSelectionTo = async (e: React.PointerEvent) => {
    // the anchor resolves asynchronously (the index may still be building), so
    // resolve it here too if a fast drag got ahead of it
    let anchor = selectionAnchor.current
    if (!anchor && selectionStartRef.current) {
      anchor = (await caretFromEvent(selectionStartRef.current)) ?? null
      selectionAnchor.current = anchor
    }
    const caret = await caretFromEvent(e)
    // a drag may cross a page boundary — normalizeRange orders by page
    if (!caret || !anchor) return
    setSelection(normalizeRange(anchor, caret))
  }

  const selectWordAt = (caret: CaretPos) => {
    const index = indexRef.current
    const line = index?.pages.find((p) => p.index === caret.pageIndex)?.lines[caret.lineIndex]
    if (!index || !line) return
    const { start, end } = wordRangeAt(line, caret.charIndex)
    selectionAnchor.current = caret
    setSelection({
      start: { pageIndex: caret.pageIndex, lineIndex: caret.lineIndex, charIndex: start },
      end: { pageIndex: caret.pageIndex, lineIndex: caret.lineIndex, charIndex: end },
    })
  }

  const selectLineAt = (caret: CaretPos) => {
    const index = indexRef.current
    const line = index?.pages.find((p) => p.index === caret.pageIndex)?.lines[caret.lineIndex]
    if (!index || !line) return
    const { start, end } = lineRangeAt(line, caret.charIndex)
    selectionAnchor.current = caret
    setSelection({
      start: { pageIndex: caret.pageIndex, lineIndex: caret.lineIndex, charIndex: start },
      end: { pageIndex: caret.pageIndex, lineIndex: caret.lineIndex, charIndex: end },
    })
  }

  const doCopy = async () => {
    if (!allowCopy) return
    const index = indexRef.current
    if (!index || !selection) return
    const text = textForRange(index, selection)
    if (text.length === 0) return
    const outcome = await copyText(text)
    // be honest about failure: the selection stays visible so the reader can
    // still press the platform copy shortcut themselves
    setCopied(outcome === 'failed' ? 'failed' : 'done')
    setTimeout(() => setCopied('idle'), 1600)
  }

  const overlayFor = (pageIndex: number): SearchOverlay | undefined => {
    const searchRects =
      matches.length === 0 ? [] : matches.filter((m) => m.pageIndex === pageIndex).flatMap((m) => m.rects)
    const activeMatch = active >= 0 ? matches[active] : undefined
    const activeRects = activeMatch && activeMatch.pageIndex === pageIndex ? activeMatch.rects : []
    const selRects =
      selection && indexRef.current ? rectsForSelectionOnPage(indexRef.current, pageIndex, selection) : []
    if (searchRects.length === 0 && activeRects.length === 0 && selRects.length === 0) return undefined
    return {
      rects: searchRects,
      active: activeRects,
      selection: selRects,
      highlights: searchRects,
      activeMatch: activeRects,
      key: `${searchRects.length}:${active}:${selectionKey}:${query}`,
    }
  }

  return (
    <div style={{ position: 'relative', display: 'flex', flexDirection: 'column', minHeight: 0 }}>
      <div
        ref={containerRef}
        className={className}
        role="region"
        aria-label="Document viewer"
        data-officeview-root=""
        tabIndex={0}
        onKeyDown={onKeyDown}
        onDoubleClick={onDoubleClick}
        onClick={onClick}
        onWheel={onWheel}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={endPointer}
        onPointerCancel={endPointer}
        onCopyCapture={allowCopy ? undefined : preventNativeCopy}
        onContextMenu={allowCopy ? undefined : preventNativeCopy}
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
          // there is no DOM text to select, but block the browser's own
          // selection gesture so drag-selecting the page chrome is not possible
          WebkitUserSelect: allowCopy ? undefined : 'none',
          userSelect: allowCopy ? undefined : 'none',
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
            <PageCanvas
              key={i}
              spec={p.spec}
              paint={p.paint}
              displayWidth={displayWidth}
              gap={pageGapPx}
              baseScale={dpr}
              boostedScale={dpr * transform.zoom}
              boosted={zoomed}
              rootRef={containerRef}
              overlay={overlayFor(i)}
            />
          ))}
        </div>
      </div>
      {showSearch && (
        <SearchBar
          inputRef={inputRef}
          query={query}
          onQuery={(q) => setQuery(q)}
          count={matches.length}
          active={active}
          indexing={indexing}
          onPrev={() => goTo(-1)}
          onNext={() => goTo(1)}
        />
      )}
      {selection && allowCopy && !isEmptyRange(selection) && (
        <button
          type="button"
          data-testid="officeview-copy-selection"
          onClick={() => void doCopy()}
          style={{
            position: 'absolute',
            bottom: 'calc(env(safe-area-inset-bottom) + 16px)',
            left: '50%',
            transform: 'translateX(-50%)',
            padding: '8px 16px',
            borderRadius: 18,
            border: 'none',
            background: copied === 'failed' ? '#8a1f11' : copied === 'done' ? '#1a7f37' : 'rgba(20,22,28,0.92)',
            color: '#fff',
            fontSize: 14,
            cursor: 'pointer',
            boxShadow: '0 2px 10px rgba(0,0,0,0.35)',
            zIndex: 6,
          }}
        >
          {copied === 'done' ? 'Copied ✓' : copied === 'failed' ? 'Copy blocked — press ⌘C' : 'Copy selection'}
        </button>
      )}
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

function SearchBar({
  inputRef,
  query,
  onQuery,
  count,
  active,
  indexing,
  onPrev,
  onNext,
}: {
  inputRef: React.RefObject<HTMLInputElement | null>
  query: string
  onQuery: (q: string) => void
  count: number
  active: number
  indexing: boolean
  onPrev: () => void
  onNext: () => void
}): ReactElement {
  const status = indexing ? 'indexing…' : count === 0 ? (query ? 'no matches' : '') : `${active + 1} of ${count}`
  const btn: CSSProperties = {
    height: 28,
    minWidth: 28,
    padding: '0 8px',
    borderRadius: 6,
    border: '1px solid rgba(0,0,0,0.2)',
    background: count > 0 ? '#f0f0f0' : 'rgba(0,0,0,0.06)',
    color: count > 0 ? '#111' : '#888',
    fontSize: 14,
    cursor: count > 0 ? 'pointer' : 'default',
  }
  return (
    <div
      data-testid="officeview-search"
      style={{
        position: 'absolute',
        top: 'calc(env(safe-area-inset-top) + 8px)',
        right: 12,
        display: 'flex',
        alignItems: 'center',
        gap: 6,
        padding: 6,
        borderRadius: 10,
        background: 'rgba(255,255,255,0.95)',
        boxShadow: '0 2px 8px rgba(0,0,0,0.3)',
        zIndex: 6,
      }}
    >
      <input
        ref={inputRef}
        type="search"
        value={query}
        placeholder="Search…"
        aria-label="Search document"
        onChange={(e) => onQuery(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') { e.preventDefault(); if (e.shiftKey) onPrev(); else onNext() }
          else if (e.key === 'Escape') { e.preventDefault(); onQuery('') }
        }}
        style={{
          height: 28,
          width: 148,
          padding: '0 8px',
          borderRadius: 6,
          border: '1px solid rgba(0,0,0,0.2)',
          fontSize: 14,
          outline: 'none',
        }}
      />
      <span
        data-testid="officeview-search-status"
        style={{ minWidth: 64, fontSize: 12, color: '#555', textAlign: 'center' }}
      >
        {status}
      </span>
      <button type="button" onClick={onPrev} disabled={count === 0} aria-label="Previous match" style={btn}>↑</button>
      <button type="button" onClick={onNext} disabled={count === 0} aria-label="Next match" style={btn}>↓</button>
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
