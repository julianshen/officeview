import type { CanvasColors } from './canvasColors'
/**
 * <OfficeDoc> — renders a parsed office document (any format) onto stacked
 * per-page <canvas> elements inside a scrollable container. Handles:
 *  - devicePixelRatio-aware canvas sizing (crisp on retina/mobile)
 *  - responsive width: pages scale to container width, capped at 1:1 zoom
 *  - IntersectionObserver-based lazy rendering (only visible pages paint)
 *  - touch scrolling is native (overflow scroll), no custom gesture code needed
 */
import { forwardRef, lazy, Suspense, useEffect, useId, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState, type CSSProperties, type ReactElement, type Ref } from 'react'
import { getPaintables, type PageSpec, type Paintable, type PaintableArray } from '../render/paint'
import type { WatermarkOptions } from '../core/watermark'
import type { FallbackFontsOptions } from '../core/fonts/fallback'
import { buildTextIndex, findMatches, type SearchMatch, type TextIndex } from '../core/search'
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
import { isViewerUiTarget } from './xlsx/gridEvents'
import type { PageOverlay } from '../core/overlay'
import type { PagedChromeProps } from './paged/PagedChrome'
import { PAGE_PADDING_PX as PAD, PAGED_STATUS_HEIGHT, PAGED_TOOLBAR_HEIGHT, COMPACT_PANE_PX, isCompact } from './paged/metrics'
import { useCoarsePointer } from './pointer'
import type { ZoomContext, ZoomOutcome } from './zoomRouting'
import { copyText } from '../core/clipboard'
import { pagedWorkbookOf } from './pagedDocument'
import { themeRootProps, type ThemeMode } from './theme'
import { defaultPlatform, SlideShowError, startPresent, type OfficeViewerHandle, type PresentOptions, type PresentPlatform, type ReexportableSource, type SlideShowDiagnostic, type SlideShowMotion } from './present'
import {
  distance,
  MAX_ZOOM,
  midpoint,
  panTransform,
  stepZoom,
  transformCss,
  type Point,
  type Transform,
  type Viewport,
} from '../core/zoom'

export type OfficeDocSource = import('../docx/types').DocxDocument | import('../xlsx/types').XlsxDocument | import('../pptx/types').PptxDocument

export interface OfficeDocProps {
  document: OfficeDocSource
  /** Container background (default: the --ov-color-backdrop token). */
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
  /**
   * Stamp a watermark on every page (default none).
   *
   * Painted into the page bitmap rather than laid over it, so it travels with
   * the document instead of being removable by hiding an element. See
   * <WatermarkOptions> for placement, rotation and opacity.
   *
   * This is a deterrent, not DRM: it marks a page, it does not protect a
   * secret inside it.
   */
  watermark?: WatermarkOptions
  /**
   * Pin explicit font fallback chains/regions (default none).
   *
   * Absent, resolvers pass through byte-identical and rendering never moves.
   * Pinned, every resolved face carries the ordered chain before its generic,
   * so hosts with those faces installed paint deterministically. See
   * <FallbackFontsOptions>.
   */
  fallbackFonts?: FallbackFontsOptions
  /**
   * Opt in to the interactive spreadsheet viewport for XLSX documents (default
   * false). Off, spreadsheets render exactly as before (fit-to-width pages).
   * On, an XLSX document opens in sheet mode: absolute 25%–400% zoom with
   * scrolling, bounded canvas tiles, grid selection, worksheet navigation and
   * workbook-wide find. Other formats ignore the flag. The interactive code is
   * loaded on demand, only when the flag is on for a spreadsheet.
   *
   * One exception: a worker-backed (large-file) workbook that <OfficeFile>
   * opened with the flag on stays in sheet mode (and keeps the interactive code
   * loaded) if the flag later turns off, because the legacy renderer has no
   * resident rows to show for it. Such a document only arises when the flag was
   * on at load time.
   *
   * Experimental: this flag stays opt-in until the interactive viewer's release
   * review; its behaviour may change between minor versions.
   */
  interactiveSheets?: boolean
  /**
   * Colour theme (default 'system', which follows prefers-color-scheme). Set on
   * the viewer's theme root as data-ov-theme; hosts re-brand by overriding the
   * --ov-* tokens from any rule that targets [data-ov-theme-root].
   */
  theme?: ThemeMode
  /** Show the Word/PowerPoint toolbar and status line (default true). */
  showToolbar?: boolean
  /** File name shown in the toolbar; <OfficeFile> derives one when unset. */
  title?: string
  /** Allow presenting a PPTX as a slideshow (default true); false hides Present and makes `present()` refuse. */
  allowPresent?: boolean
  /** F5 / Shift+F5 start the show while focus is inside the viewer (default true). */
  presentShortcuts?: boolean
  /** Slideshow motion (default 'auto': follows prefers-reduced-motion). */
  slideShowMotion?: SlideShowMotion
  onSlideShowStart?: () => void
  /** Fires exactly once per started show, before `session.ended` resolves. */
  onSlideShowEnd?: () => void
  onSlideShowSlideChange?: (slideIndex: number, buildIndex: number) => void
  onSlideShowDiagnostic?: (diagnostic: SlideShowDiagnostic) => void
  /** @internal */
  __presentPlatform?: PresentPlatform
  /** @internal The re-readable source <OfficeFile> loaded from. */
  __presentSource?: ReexportableSource
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
  colors?: CanvasColors
}

/**
 * Draw search highlights and the text selection over the page content. Called
 * while the ctx still carries the page transform, so rectangles are in page
 * coordinates.
 */
function paintHighlights(ctx: CanvasRenderingContext2D, overlay: SearchOverlay, colors: Pick<CanvasColors, 'find' | 'find-active' | 'text-selection'> = {}): void {
  ctx.save()
  ctx.fillStyle = colors.find ?? 'rgba(255,214,0,.45)'
  for (const r of overlay.highlights) ctx.fillRect(r.x, r.y, r.width, r.height)
  // selection sits on top of the search wash, as it does in a browser
  ctx.fillStyle = colors['text-selection'] ?? 'rgba(64,128,255,.35)'
  for (const r of overlay.selection) ctx.fillRect(r.x, r.y, r.width, r.height)
  ctx.strokeStyle = colors['find-active'] ?? '#ff8c00'
  ctx.lineWidth = 1.5
  for (const r of overlay.activeMatch) ctx.strokeRect(r.x, r.y, r.width, r.height)
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
    if (overlay) paintHighlights(ctx, overlay, overlay.colors)
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

const LegacyOfficeDoc = forwardRef(function LegacyOfficeDoc({
  document,
  background,
  pageGapPx = 16,
  className,
  style,
  showZoomControls = true,
  showSearch = true,
  allowCopy = true,
  allowPrint = true,
  watermark,
  fallbackFonts,
  theme,
  showToolbar = true,
  title,
  allowPresent = true,
  presentShortcuts = true,
  slideShowMotion = 'auto',
  onSlideShowStart,
  onSlideShowEnd,
  onSlideShowSlideChange,
  onSlideShowDiagnostic,
  __presentPlatform = defaultPlatform,
  __presentSource,
}: OfficeDocProps, ref: Ref<OfficeViewerHandle>): ReactElement {
  const instance = useId()
  // The slideshow host (spec 4.1): always mounted, empty and hidden, so full
  // screen can be requested on it synchronously inside the gesture. A start
  // belongs to the document epoch: unmount, a new document or allowPresent
  // withdrawn aborts it (cancel during load, exit once running).
  const hostRef = useRef<HTMLDivElement>(null)
  const epoch = useRef<AbortController>(null)
  const gen = useRef(0)
  const [notice, setNotice] = useState('')
  useLayoutEffect(() => {
    const c = epoch.current = new AbortController()
    gen.current++
    return () => c.abort()
  }, [document, allowPresent])
  const present = (opts?: PresentOptions) => {
    setNotice('')
    const run = startPresent({
      doc: document, host: hostRef.current!, opts, platform: __presentPlatform, generation: gen.current, source: __presentSource, allow: allowPresent, motion: slideShowMotion, watermark, signal: epoch.current!.signal,
      callbacks: { onStart: onSlideShowStart, onEnd: onSlideShowEnd, onSlideChange: onSlideShowSlideChange, onDiagnostic: onSlideShowDiagnostic },
    })
    run.catch((e) => e.code == 'load-failed' && setNotice(e.message))
    return run
  }
  // The Present button and F5 both start here. Everything about its offer
  // (preload hint, busy state) lives in the lazy toolbar; see PagedToolbar.
  const onPresent = (from: 'beginning' | 'current') => void present({ from }).catch(() => {})
  useImperativeHandle(ref, () => ({ present }))
  // getPaintables is async (docx measurement resolves a 2D ctx); hold in state.
  const [pages, setPages] = useState<Paintable[]>([])
  // The watermark is painted into the page bitmap, so changing it has to repaint.
  // A stable string key keeps the effect from re-running on a fresh object
  // identity every render (and from looping).
  const markKey = watermark ? JSON.stringify(watermark) : ''
  const fontKey = fallbackFonts ? JSON.stringify(fallbackFonts) : ''
  const owner = useMemo(() => ({ document, markKey, fontKey }), [document, markKey, fontKey])
  const [printRequest, requestPrint] = useState<{ owner: typeof owner }>()
  const ownerRef = useRef<typeof owner | null>(owner)
  ownerRef.current = owner
  const copyAllowed = useRef(allowCopy)
  copyAllowed.current = allowCopy
  useEffect(() => {
    let cancelled = false
    let lease: PaintableArray | undefined
    const abort = new AbortController()
    ownerRef.current = owner
    setPages([])
    commit(IDENTITY)
    pending.current = 0
    getPaintables(document, { signal: abort.signal, ...(markKey ? { watermark } : {}), ...(fontKey ? { fallbackFonts } : {}) }).then(p => {
      if (cancelled) { p.dispose(); return }
      lease = p
      setPages(p)
    })
    return () => {
      cancelled = true
      abort.abort()
      lease?.dispose()
      if (ownerRef.current === owner) ownerRef.current = null
    }
  }, [owner])
  const containerRef = useRef<HTMLDivElement | null>(null)
  const [containerWidth, setContainerWidth] = useState(0)
  const [viewportHeight, setViewportHeight] = useState(0)
  const [transform, setTransform] = useState<Transform>(IDENTITY)
  // Every transform change goes through commit, so handlers firing before the
  // next render (a burst of wheel events) build on the latest value.
  const tRef = useRef(transform)
  const commit = (t: Transform) => { tRef.current = t; setTransform(t) }
  // Fit-to-width is a layout mode (spec §6.9): pages are laid out at the pane
  // width, the transform stays at zoom 1 and the container scrolls natively.
  const [fitMode, setFit] = useState(false)
  const fitRef = useRef(false)
  const setFitMode = (on: boolean) => { fitRef.current = on; setFit(on) }
  // A scrollTop applied in the commit that switches overflow (spec §6.11):
  // scroll and pan trade places atomically.
  const pending = useRef<number | null>(null)
  useLayoutEffect(() => {
    const el = containerRef.current
    if (el && pending.current !== null) { el.scrollTop = pending.current; pending.current = null }
  })
  // gesture bookkeeping
  const pointers = useRef(new Map<number, Point>())
  const gestureStart = useRef<{ context: ZoomContext; at: Point; distance: number } | null>(null)
  const panStart = useRef<{ transform: Transform; point: Point } | null>(null)

  // React attaches wheel listeners as passive, so onWheel cannot stop the
  // browser's own Ctrl/Cmd+wheel page zoom; a native listener on the outer
  // wrapper can, over the pages and every control (search, copy, zoom).
  const [palette, setPalette] = useState<CanvasColors & {key: string}>()
  const wrapperRef = useRef<HTMLDivElement | null>(null)
  // Once per mount: request the zoom routing. (The lazy chrome checks that
  // officeview.css is loaded, spec §5, R10.)
  useEffect(() => { void loadRouter() }, [])
  useEffect(() => {
    const el = wrapperRef.current
    if (!el) return
    const suppressPageZoom = (e: WheelEvent) => { if (e.ctrlKey || e.metaKey) e.preventDefault() }
    el.addEventListener('wheel', suppressPageZoom, { passive: false })
    return () => el.removeEventListener('wheel', suppressPageZoom)
  }, [])

  // Layout effect: the observer is disconnected inside the commit that removes
  // the container, and a report for a detached container is ignored, so no
  // update is queued for an unmounted document.
  useLayoutEffect(() => {
    const el = containerRef.current
    if (!el) return
    const update = () => {
      if (!el.isConnected) return
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

  const specs = useMemo(() => pages.map((p) => p.spec), [pages])

  // Natural: page width = min(container inner width, natural width). Fit:
  // the pane width, capped at MAX_ZOOM x natural; containers 0-32px wide
  // keep the natural width.
  const avail = Math.max(0, containerWidth - PAD * 2)
  const maxPageW = specs.length > 0 ? Math.max(...specs.map((s) => s.widthPx)) : 800
  const naturalWidth = Math.min(avail || maxPageW, maxPageW)
  const fitWidth = avail > 0 ? Math.min(avail, maxPageW * MAX_ZOOM) : naturalWidth
  const displayWidth = fitMode ? fitWidth : naturalWidth
  const layout = { specs, naturalWidth, fitWidth, gap: pageGapPx }
  const layoutRef = useRef(layout)
  layoutRef.current = layout

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
    if (isViewerUiTarget(e.target)) return
    if (e.pointerType === 'mouse' && e.button !== 0) return
    const point = localPoint(e)
    pointers.current.set(e.pointerId, point)
    const pts = [...pointers.current.values()]
    if (pts.length === 2) {
      // start a pinch from the current transform
      // the full start context: its view is derived once the routing is in
      gestureStart.current = {
        context: zoomContext(),
        at: midpoint(pts[0], pts[1]),
        distance: Math.max(1, distance(pts[0], pts[1])),
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
    // ownership was decided at pointerdown; a tracked finger may cross viewer UI
    if (!pointers.current.has(e.pointerId)) return
    const point = localPoint(e)
    pointers.current.set(e.pointerId, point)
    const pts = [...pointers.current.values()]
    if (pts.length >= 2 && gestureStart.current) {
      const g = gestureStart.current
      const ratio = Math.max(1, distance(pts[0], pts[1])) / g.distance, at = midpoint(pts[0], pts[1])
      routed((m) => { const view = m.viewAt(g.context, g.at); apply(m.zoomTo(zoomContext(), view, view.zoom * ratio, at)) })
      return
    }
    // a drag past the threshold while zoomed pans instead of selecting
    if (pts.length === 1 && panStart.current) {
      if (dragStart.current && Math.hypot(point.x - dragStart.current.x, point.y - dragStart.current.y) > 6) {
        selectionDrag.current = null
      }
      const start = panStart.current
      commit(panTransform(start.transform, { x: point.x - start.point.x, y: point.y - start.point.y }, viewportRef.current))
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
    if (isViewerUiTarget(e.target)) return
    e.preventDefault()
    const at = localPoint(e)
    const factor = Math.exp(-e.deltaY / 250)
    zoomAround(at, (z) => z * factor)
  }

  const onDoubleClick = (e: React.MouseEvent) => {
    if (isViewerUiTarget(e.target)) return
    const at = localPoint(e)
    // double-click on text selects a word; on blank space it zooms, so both
    // behaviours stay available
    void caretFromEvent(e, true).then((caret) => {
      if (caret) {
        selectWordAt(caret)
        return
      }
      zoomAround(at, (z) => (z > 1 ? 1 : 2))
    })
  }

  const onClick = (e: React.MouseEvent) => {
    if (isViewerUiTarget(e.target)) return
    if (e.detail < 3) return // double-click already handled words
    void caretFromEvent(e).then((caret) => caret && selectLineAt(caret))
  }

  // Every zoom input goes through zoomRouting (scroll/pan model, fit exit).
  const zoomContext = (): ZoomContext => ({
    transform: tRef.current,
    fit: fitRef.current,
    scrollTop: pending.current ?? containerRef.current?.scrollTop ?? 0,
    ...layoutRef.current,
    viewport: viewportRef.current,
  })
  const apply = (o: ZoomOutcome | null) => {
    if (!o) return
    // No transform or mode change means no commit to carry the scrollTop.
    if (o.fit === fitRef.current && transformCss(o.transform) === transformCss(tRef.current)) { if (containerRef.current) containerRef.current.scrollTop = o.scrollTop; return }
    setFitMode(o.fit)
    pending.current = o.scrollTop
    commit(o.transform)
  }
  // Input queued for the routing applies only to the document it was made on.
  const routed = (fn: (m: Router) => void) => { const o = ownerRef.current; route((m) => { if (o && ownerRef.current === o) fn(m) }) }
  const zoomAround = (at: Point, next: (zoom: number) => number) => routed((m) => {
    const c = zoomContext()
    const view = m.viewAt(c, at)
    apply(m.zoomTo(c, view, next(view.zoom), at))
  })
  const center = (): Point => ({ x: viewportRef.current.width / 2, y: viewportRef.current.height / 2 })
  const zoomBy = (direction: 1 | -1) => zoomAround(center(), (z) => stepZoom(z, direction))
  const reset = () => zoomAround(center(), () => 1)
  const toggleFit = () => routed((m) => apply(m.toggleFit(zoomContext())))

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.key == 'F5' && presentShortcuts && allowPresent && !e.altKey && 'slides' in document) {
      e.preventDefault()
      return onPresent(e.shiftKey ? 'current' : 'beginning')
    }
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'p') {
      e.preventDefault()
      if (allowPrint) requestPrint({ owner })
    }
    if (e.defaultPrevented) return
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'f') {
      e.preventDefault()
      if (inputRef.current) {
        inputRef.current.focus()
        inputRef.current.select()
      } else if (showSearch) {
        // Find is collapsed or its chunk is still loading: the toolbar
        // expands and focuses it. With search off nothing is queued.
        setFocusSearchRequest((n) => n + 1)
      }
      return
    }
    // Protection policy and Escape fall-through apply even when the event comes
    // from viewer UI such as an owned grid; a control that consumed Escape
    // already prevented it.
    if (!allowCopy && (e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'c') {
      // when copying is off, swallow the shortcut rather than letting the
      // browser act on whatever happens to be selected around the canvas
      e.preventDefault()
      return
    }
    if (e.key === 'Escape' && selection) {
      setSelection(undefined)
      selectionAnchor.current = null
      return
    }
    if (isViewerUiTarget(e.target)) return
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'c' && selection && !isEmptyRange(selection)) {
      e.preventDefault()
      void doCopy()
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
  const [focusSearchRequest, setFocusSearchRequest] = useState(0)

  const indexOwnerRef = useRef<typeof owner | null>(null)
  const indexPendingRef = useRef<{ owner: typeof owner; promise: Promise<TextIndex> } | null>(null)
  // Indexing can race initial font loading. Its fallback extraction owns a
  // separate lease, and an obsolete build must never publish into the new doc.
  const ensureIndex = (): Promise<TextIndex> => {
    if (indexRef.current && indexOwnerRef.current === owner) return Promise.resolve(indexRef.current)
    if (indexPendingRef.current?.owner === owner) return indexPendingRef.current.promise
    setIndexing(true)
    const promise = (async () => {
      let fallback: PaintableArray | undefined
      try {
        // Own a consumer even when pages already exist: recording awaits its
        // canvas and may outlive the viewer that owns those page closures.
        const paintables = fallback = await getPaintables(document, { ...(markKey ? { watermark } : {}), ...(fontKey ? { fallbackFonts } : {}) })
        const built = await buildTextIndex(paintables)
        if (ownerRef.current === owner) { indexRef.current = built; indexOwnerRef.current = owner }
        return built
      } finally {
        fallback?.dispose()
        if (indexPendingRef.current?.owner === owner) indexPendingRef.current = null
        if (ownerRef.current === owner) setIndexing(false)
      }
    })()
    indexPendingRef.current = { owner, promise }
    return promise
  }

  useEffect(() => {
    indexRef.current = null
    indexOwnerRef.current = null
    setIndexing(false)
    setMatches([])
    setActive(-1)
  }, [owner])

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
  }, [query, pages, owner])

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
    const outcome = await copyText(text, { allowed: () => copyAllowed.current && ownerRef.current === owner })
    if (!copyAllowed.current || ownerRef.current !== owner) return
    // be honest about failure: the selection stays visible so the reader can
    // still press the platform copy shortcut themselves
    setCopied(outcome === 'failed' ? 'failed' : 'done')
    setTimeout(() => setCopied('idle'), 1600)
  }

  const overlayFor = (pageIndex: number): SearchOverlay | undefined => {
    const searchRects =
      matches.filter((m) => m.pageIndex === pageIndex).flatMap((m) => m.rects)
    const activeMatch = matches[active]
    const activeRects = activeMatch?.pageIndex === pageIndex ? activeMatch.rects : []
    const selRects =
      selection && indexRef.current ? rectsForSelectionOnPage(indexRef.current, pageIndex, selection) : []
    // The active match is already a member of searchRects.
    if (!searchRects.length && !selRects.length) return undefined
    return {
      colors: palette,
      selection: selRects,
      highlights: searchRects,
      activeMatch: activeRects,
      key: `${searchRects.length}:${active}:${selectionKey}:${query}:${palette?.key}`,
    }
  }

  // Unmeasured (0) counts as regular: jsdom and display:none panes have no width.
  const compact = isCompact(containerWidth || COMPACT_PANE_PX, useCoarsePointer())
  const [statusSlot, setStatusSlot] = useState<HTMLDivElement | null>(null)
  const chrome: PagedChromeProps = {
    onPalette: setPalette,
    document,
    title,
    compact,
    statusSlot,
    containerRef,
    layout: { specs, displayWidth, gap: pageGapPx },
    transform,
    setTransform: commit,
    zoom: fitMode ? fitWidth / naturalWidth : transform.zoom,
    onZoom: zoomBy,
    onReset: reset,
    fitActive: fitMode,
    onFit: toggleFit,
    showZoom: showZoomControls,
    search: showSearch ? { inputRef, query, onQuery: setQuery, matches, active, setActive, indexing } : null,
    copy: allowCopy ? { state: copied, enabled: !!selection && !isEmptyRange(selection), onCopy: () => void doCopy() } : undefined,
    allowPrint,
    present: allowPresent && 'slides' in document ? { start: onPresent } : undefined,
    pages,
    owner,
    showToolbar,
    print: printRequest,
    focusSearchRequest,
  }

  // The pages container's style: a themed backdrop (a host background prop
  // wins); zoomed, we own the pan and the container never scrolls; with copy
  // off there is no DOM text, so the browser's selection gesture is blocked.
  return (
    <div ref={wrapperRef} {...themeRootProps(theme, instance)} className="ov:relative ov:flex ov:flex-col ov:min-h-0"
      data-ov-print={allowPrint} onKeyDown={onKeyDown}>
      {/* The lazy chrome also owns printing with showToolbar=false (if its
          chunk fails, pages render without chrome, Find or print). Fixed
          placeholders reserve visible chrome without shifting pages (§6.10). */}
      <Suspense fallback={showToolbar && <div data-testid="officeview-toolbar-placeholder" className={SLOT} style={{ height: PAGED_TOOLBAR_HEIGHT[compact ? 'compact' : 'regular'] }} />}>
        <PagedChrome {...chrome} />
      </Suspense>
      <div
        ref={containerRef}
        className={className}
        role="region"
        aria-label="Document viewer"
        data-officeview-root=""
        tabIndex={0}
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
          background: background ?? 'var(--ov-color-backdrop)',
          overflow: zoomed ? 'hidden' : 'auto',
          WebkitOverflowScrolling: 'touch',
          touchAction: zoomed ? 'none' : 'pan-y',
          padding: PAD,
          minHeight: 0,
          flex: '1 1 auto',
          position: 'relative',
          boxSizing: 'border-box',
          cursor: zoomed ? 'grab' : 'auto',
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
              baseScale={dpr * naturalWidth / displayWidth}
              boostedScale={dpr * transform.zoom}
              boosted={zoomed || fitMode}
              rootRef={containerRef}
              overlay={overlayFor(i)}
            />
          ))}
        </div>
      </div>
      {notice && <div role="alert" data-officeview-notice="">{notice}</div>}
      {showToolbar && <div ref={setStatusSlot} data-testid="officeview-status-slot" className={SLOT} style={{ height: PAGED_STATUS_HEIGHT }} />}
      <div ref={hostRef} data-officeview-show="" tabIndex={-1} style={{ visibility: 'hidden' }} />
    </div>
  )
})

// A chunk that fails to load (a redeploy removed it, a CSP blocks it) warns
// once and degrades inside the viewer rather than unmounting the host.
const warnChunk = (e: unknown) => console.warn('[officeview] chunk not loaded', e)
/** The sheet viewer's loading or failure status, inside its own theme root. */
const SheetStatus = (p: { theme?: ThemeMode; text: string }) => <div {...themeRootProps(p.theme, useId())}><div role="status" data-officeview-root="" style={{ padding: 16 }}>{p.text}</div></div>
// Loaded only when a host opts in for a spreadsheet (or a worker-backed shell
// needs it); the legacy path never pulls it. On failure: a visible status, never
// the legacy renderer (a worker-backed shell has no resident rows).
const InteractiveSheetDoc = lazy(() => import('./xlsx/InteractiveSheetDoc').catch((e) => (warnChunk(e), { default: (p: { theme?: ThemeMode }) => <SheetStatus theme={p.theme} text="Spreadsheet viewer could not load. Reload to retry." /> as never })))

// The toolbar, status line, page tracking and word count (spec §6.10). Key
// handlers and zoom state stay here, eager.
const PagedChrome = lazy(() => import('./paged/PagedChrome').catch((e) => (warnChunk(e), { default: () => null as never })))
// The zoom routing is lazy too, requested by the first paged render: its
// gestures need painted pages, so it is loaded before they can happen; an
// input that still beats it is applied once it arrives.
type Router = typeof import('./zoomRouting')
let router: Router | undefined
let routerP: Promise<Router | void> | undefined
/** @internal Tests delay the routing module through this. */
export const routerLoader = { load: (): Promise<Router> => import('./zoomRouting') }
// One attempt: a failure warns once and zoom stays off (no retry per gesture).
const loadRouter = () => routerP ??= routerLoader.load().then((m) => (router = m), warnChunk)
const route = (fn: (m: Router) => void) => { if (router) fn(router); else void loadRouter().then((m) => m && fn(m)) }
const SLOT = 'ov:shrink-0 ov:bg-chrome'
const IDENTITY: Transform = { zoom: 1, panX: 0, panY: 0 }

// A spreadsheet viewer is never presentable: the ref still gets a handle that says so (PR #18 review 3).
const SHEET: OfficeViewerHandle = { present: () => Promise.reject(new SlideShowError('not-presentable')) }

export const OfficeDoc = forwardRef(function OfficeDoc(props: OfficeDocProps, ref: Ref<OfficeViewerHandle>): ReactElement {
  const { interactiveSheets = false, document, ...rest } = props
  // A worker-backed (paged) shell document carries its source (S5.4) and has no
  // rows on the main thread: only the interactive viewer can show it.
  const paged = pagedWorkbookOf(document)
  const sheet = (interactiveSheets || paged) && 'sheets' in document
  // The spreadsheet's handle; the legacy viewer sets its own.
  useImperativeHandle(sheet ? ref : null, () => SHEET)
  if (sheet) {
    return (
      <Suspense fallback={<SheetStatus theme={props.theme} text="Loading spreadsheet viewer…" />}>
        <InteractiveSheetDoc doc={document} {...(paged ? { source: paged.source } : {})} {...rest} />
      </Suspense>
    )
  }
  return <LegacyOfficeDoc ref={ref} document={document} {...rest} />
})

export default OfficeDoc
