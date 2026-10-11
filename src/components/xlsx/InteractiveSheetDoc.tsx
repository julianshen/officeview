/**
 * Interactive sheet viewport (plan S5.2; spec §§5.5, 9, 10.1, 11, 12.4), mounted
 * by <OfficeDoc interactiveSheets> for XLSX documents and loaded on demand.
 *
 * Position is a ScrollWindow plus an absolute SheetZoom (0.25–4), never the
 * page viewer's fit-to-width pan transform. Three independent layers:
 *  - base: 512px device tiles painted by renderSheetViewport into the bitmap
 *    pool (SheetTileRenderer) and drawn into one pane surface;
 *  - overlay: grid selection and search highlights, repainted alone when only
 *    they change (the watermark is stamped into the pane bitmap after tiles);
 *  - DOM: GridSemantics (owned ARIA grid, headers, status) and SheetNavigator.
 *
 * Ownership (§10.1): a document replacement remounts the viewer and resets zoom;
 * view commits (the `views` seam S9.1 drives), font publication and watermark
 * changes keep zoom and scroll. A committed view bumps that sheet's viewGen and
 * the workbook globalViewGen, which re-keys base tiles, find and hit geometry.
 *
 * TODO(S1.2/S4): source reads, geometry and find run in-thread over the interim
 * object source; the worker `viewport`/`search` commands replace them for the
 * paged source without changing this component's layers.
 */
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type CSSProperties, type ReactElement } from 'react'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../../xlsx/types'
import type {
  AuthoritativeCell, CacheIdentity, FilterSpec, GridSelection, ProfileId, RowOrder, RowVisibility, ScrollWindow, SheetView, SortSpec, SourceRowInterval, SourceStore, WorkbookOwner,
} from '../../xlsx/view/contracts'
import { SHEET_MIN_ZOOM } from '../../xlsx/view/contracts'
import type { NavigatorRequest, SourceCell } from '../../xlsx/view/coordinates'
import { buildSheetGeometry, MergeQueryCache, sheetBaseHidden, type SheetGeometry } from '../../xlsx/view/viewport'
import type { ProjectedSheetMetrics } from '../../xlsx/view/projection'
import { createIdentityOrder } from '../../xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../../xlsx/view/visibility'
import {
  createScrollWindow, navigateTo as scrollNavigateTo, refreshScrollWindow, scrollPhysicalTo, setScrollZoom, visibleLogicalViewport, type ViewportSize,
} from '../../xlsx/view/scrollWindow'
import { keyboardSheetZoom, pinchSheetZoom, wheelSheetZoom } from '../../xlsx/view/sheetZoom'
import { SheetNavigationSession, emptyGridSelection, navigateSelection } from '../../xlsx/view/selection'
import { FIND_PAGE_HITS, FindController, hitGeometry, navigateToHit, type FindHit, type FindInput, type FindOptions, type FindPublication } from '../../xlsx/view/find'
import { LayoutCache } from '../../xlsx/view/textLayout'
import { cellDisplayText } from '../../xlsx/render'
import { createCellLayoutHost } from '../../xlsx/view/renderViewport'
import { createObjectWorkbookSource } from '../../worker/viewerObjectSource'
import type { TileCacheStats } from '../../xlsx/view/tileCache'
import { gridGeometryToViewport, gridSelectionOverlay, paintGridSelectionOverlay, paintSearchOverlay, type HighlightRect } from '../../core/overlay'
import { normalizeWatermark, type WatermarkOptions } from '../../core/watermark'
import type { FallbackFontsOptions } from '../../core/fonts/fallback'
import { prepareSheetAssets, type SheetAssetsLease } from '../../render/paint'
import { GridSemantics, type GridTransition } from './GridSemantics'
import { GRID_HEADER_SIZE } from './GridHeaders'
import { coarseSelectionRecords, overlayRecords, selectionContains } from './selectionOverlay'
import { SheetNavigator } from './SheetNavigator'
import type { HiddenCellInspection } from './ViewerStatus'
import { SheetTileRenderer, admitDeviceScale, paneSurfaceBytes, type SheetFrame } from './sheetTiles'
import { SearchBar, ZoomControls } from '../viewerControls'

/** A committed sheet view (the S9.1 sort/filter commit seam). `revision` must
 * increase with every commit; it becomes the sheet's viewGen. */
export interface SheetViewInput {
  revision: number
  order?: RowOrder
  /** DISPLAY-space mask (authored hidden ∪ filter); absent = authored hidden. */
  visibility?: RowVisibility
  filter?: FilterSpec | null
  sort?: SortSpec | null
}
export interface SheetFrameStats {
  basePaints: number
  tileHits: number
  prefetched: number
  refused: number
  overlayPaints: number
  /** Logical boxes the latest base frame requested (one per missing visible tile). */
  lastRequests: readonly { x: number; y: number; width: number; height: number }[]
  tiles: TileCacheStats
  /** Accounted bitmap bytes now and at peak (pane + overlay + tiles + in-flight). */
  bitmapBytes: number
  peakBitmapBytes: number
  budgetBytes: number
  /** Detail records presented by the latest frame / painted by the latest prefetch batch. */
  frameRecords: number
  frameGridlines: number
  prefetchRecords: number
  activeSheetId: string
  zoom: number
  identity: CacheIdentity
}
export interface InteractiveSheetDocProps {
  doc: XlsxDocument
  background?: string
  className?: string
  style?: CSSProperties
  showZoomControls?: boolean
  showSearch?: boolean
  allowCopy?: boolean
  allowPrint?: boolean
  watermark?: WatermarkOptions
  fallbackFonts?: FallbackFontsOptions
  pageGapPx?: number
  /** @internal Committed sort/filter views per sheet id (S9.1 seam). */
  views?: Readonly<Record<string, SheetViewInput>>
  /** @internal Instrumentation for tests and the browser harness. */
  onFrame?: (stats: SheetFrameStats) => void
  /** @internal Resource profile (desktop D / mobile M); default D. */
  profile?: ProfileId
}

// Document identity alone owns zoom reset (§10.1): a stable id per document object.
const documentIds = new WeakMap<object, string>()
let nextDocumentId = 0
function documentIdOf(doc: XlsxDocument): string {
  let id = documentIds.get(doc)
  if (!id) documentIds.set(doc, id = `xlsx-${++nextDocumentId}`)
  return id
}
export function sheetIdOf(sheet: XlsxSheet, index: number): string { return sheet.sheetId ?? String(index + 1) }

export default function InteractiveSheetDoc(props: InteractiveSheetDocProps): ReactElement {
  const docId = documentIdOf(props.doc)
  return <SheetViewer key={docId} docId={docId} {...props} />
}

interface SheetEntry {
  sheet: XlsxSheet
  revision: number
  view: SheetView
  geometry: SheetGeometry
  metrics: ProjectedSheetMetrics
  body: SourceRowInterval | null
  layoutCache: LayoutCache
  merges: MergeQueryCache
}
/** Fallback pane when the host has not laid the viewer out (jsdom, hidden tab). */
const FALLBACK_PANE = { width: 640, height: 400 }

function SheetViewer(props: InteractiveSheetDocProps & { docId: string }): ReactElement {
  const { doc, docId, views, watermark, fallbackFonts, allowCopy = true, allowPrint = true, showSearch = true, showZoomControls = true } = props
  const profile = props.profile ?? 'D'
  const sheetIds = useMemo(() => doc.sheets.map(sheetIdOf), [doc])
  const [activeSheetId, setActiveSheetId] = useState(sheetIds[0] ?? '')
  const [, setTick] = useState(0)
  const rerender = useCallback(() => setTick(tick => tick + 1), [])
  const session = useRef(new SheetNavigationSession(docId)).current
  const windows = useRef(new Map<string, ScrollWindow>()).current
  const entries = useRef(new Map<string, SheetEntry>()).current
  const source = useMemo(() => createObjectWorkbookSource(doc, docId), [doc, docId])
  useEffect(() => () => { void source.release() }, [source])

  // ---- fonts/images: a font publication bumps fontGen (re-keys tiles, keeps zoom) ----
  const fontKey = fallbackFonts ? JSON.stringify(fallbackFonts) : ''
  const [assets, setAssets] = useState<{ lease: SheetAssetsLease; fontGen: number } | null>(null)
  const fontGenRef = useRef(0)
  useEffect(() => {
    let cancelled = false, lease: SheetAssetsLease | undefined
    void prepareSheetAssets(doc, fontKey ? { fallbackFonts } : {}).then(next => {
      if (cancelled) { next.dispose(); return }
      lease = next
      setAssets({ lease: next, fontGen: ++fontGenRef.current })
    })
    return () => { cancelled = true; lease?.dispose() }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [doc, fontKey])
  const fontGen = assets?.fontGen ?? 0

  // ---- per-sheet committed views and geometry (one geometry per viewGen) ----
  const viewInput = (sheetId: string): SheetViewInput => views?.[sheetId] ?? { revision: 0 }
  const entryFor = (sheetId: string): SheetEntry => {
    const index = sheetIds.indexOf(sheetId)
    const sheet = doc.sheets[index]
    const input = viewInput(sheetId)
    const previous = entries.get(sheetId)
    if (previous && previous.revision === input.revision) return previous
    if (previous && input.revision < previous.revision) return previous // stale commit: never publish an older view
    const order = input.order ?? createIdentityOrder()
    const baseHidden = sheetBaseHidden(sheet)
    const visibility = input.visibility ?? visibilityFromHidden(createIntervalSet(baseHidden.intervals.map(({ r0, r1 }) => ({ r0: order.sourceToDisplay(r0), r1: order.sourceToDisplay(r1) }))))
    const built = buildSheetGeometry(sheet, { order, ...(input.visibility ? { visibility } : {}) })
    if (built.kind !== 'ready') throw new Error(built.diagnostics.map(d => d.message).join('; ') || 'This sheet cannot be laid out')
    const selection = previous?.view.selection ?? emptyGridSelection()
    const view: SheetView = { sheetId, revision: input.revision, order, baseHidden, visibility, sort: input.sort ?? null, filter: input.filter ?? null, selection }
    const entry: SheetEntry = {
      sheet, revision: input.revision, view, geometry: built, metrics: built.metrics, body: usedBody(sheet),
      // Layout records are keyed by complete identity; a new view gets its own merge cache.
      layoutCache: previous?.layoutCache ?? new LayoutCache(), merges: new MergeQueryCache(),
    }
    entries.set(sheetId, entry)
    if (!previous) session.register(sheetId, { view, projection: entry.metrics, body: entry.body })
    else session.updateProjection(sheetId, entry.metrics, entry.body, view)
    return entry
  }
  // The committed revision: a stale (lower) prop revision never regresses the
  // owner that keys tiles, find and hit geometry; a newer one is built on use.
  const revisionOf = (sheetId: string): number => Math.max(entries.get(sheetId)?.revision ?? 0, viewInput(sheetId).revision)
  const identityOf = (sheetId: string): CacheIdentity => ({ docId, sheetId, valueGen: 0, fontGen, viewGen: revisionOf(sheetId) })
  const globalViewGen = sheetIds.reduce((sum, id) => sum + revisionOf(id), 0)
  const ownerOf = (): WorkbookOwner => ({
    docId, globalViewGen, workbookValueGen: 0, fontGen,
    sheetVector: sheetIds.map(sheetId => ({ sheetId, revision: revisionOf(sheetId), valueGen: 0 })),
  })
  const ownerRef = useRef(ownerOf)
  ownerRef.current = ownerOf
  const ownerKey = JSON.stringify([globalViewGen, fontGen, sheetIds.map(revisionOf)])

  const entry = activeSheetId ? entryFor(activeSheetId) : null
  const identity = identityOf(activeSheetId)
  const identityKey = JSON.stringify([identity.docId, identity.sheetId, identity.valueGen, identity.fontGen, identity.viewGen])

  // ---- pane size and device scale ----
  const paneRef = useRef<HTMLDivElement | null>(null)
  const [pane, setPane] = useState(FALLBACK_PANE)
  useEffect(() => {
    const el = paneRef.current
    if (!el) return
    const update = () => {
      const width = el.clientWidth || FALLBACK_PANE.width, height = el.clientHeight || FALLBACK_PANE.height
      setPane(current => current.width === width && current.height === height ? current : { width, height })
    }
    update()
    if (typeof ResizeObserver === 'undefined') return
    const ro = new ResizeObserver(update)
    ro.observe(el)
    return () => ro.disconnect()
  }, [])
  const viewport: ViewportSize = useMemo(() => ({ width: Math.max(1, pane.width - GRID_HEADER_SIZE), height: Math.max(1, pane.height - GRID_HEADER_SIZE) }), [pane])
  const nativeDpr = Math.min(typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1, 3)
  // Admit the pane surfaces before sizing any canvas; a lower device scale keeps
  // the pane within the bitmap budget, and nothing fits → a visible resource outcome.
  const admittedDpr = admitDeviceScale(viewport, nativeDpr, profile)
  const dpr = admittedDpr ?? nativeDpr
  const resourceNotice = admittedDpr === null ? 'This view is too large for the display memory budget. Make the viewer smaller.'
    : admittedDpr < nativeDpr ? 'Display resolution reduced to fit the memory budget.' : ''

  // ---- scroll window per sheet; zoom is retained per sheet in the session ----
  const windowFor = (sheetId: string, sheetEntry: SheetEntry): ScrollWindow => {
    const zoom = session.get(sheetId).zoom
    let current = windows.get(sheetId)
    if (!current) current = createScrollWindow(sheetEntry.metrics, viewport, { zoom, deviceScale: dpr })
    else if (current.deviceScale !== dpr) current = createScrollWindow(sheetEntry.metrics, viewport, { zoom: current.zoom, deviceScale: dpr, offset: current.logicalOrigin })
    else current = refreshScrollWindow(current, sheetEntry.metrics, viewport)
    windows.set(sheetId, current)
    return current
  }
  // Refresh against the CURRENT projection/viewport (a view commit or resize keeps
  // the logical origin and zoom; refresh only clamps).
  const windowKey = useRef('')
  const nextWindowKey = JSON.stringify([activeSheetId, entry?.revision, viewport.width, viewport.height, dpr])
  if (entry && windowKey.current !== nextWindowKey) { windowKey.current = nextWindowKey; windowFor(activeSheetId, entry) }
  const scroll = entry ? windows.get(activeSheetId) ?? windowFor(activeSheetId, entry) : null
  const setWindow = (sheetId: string, next: ScrollWindow) => {
    windows.set(sheetId, next)
    session.setZoom(sheetId, next.zoom)
    rerender()
  }
  const selection: GridSelection = entry ? session.get(activeSheetId).selection : emptyGridSelection()

  // ---- layers ----
  const baseRef = useRef<HTMLCanvasElement | null>(null)
  const overlayRef = useRef<HTMLCanvasElement | null>(null)
  const scrollerRef = useRef<HTMLDivElement | null>(null)
  const device = { width: Math.ceil(viewport.width * dpr), height: Math.ceil(viewport.height * dpr) }
  const surfaceBytes = paneSurfaceBytes(viewport, dpr)
  const rendererRef = useRef<SheetTileRenderer | null>(null)
  // Created lazily in effects: an unmount (or StrictMode's dev re-mount) disposes
  // it, and the next effect run starts a fresh pool instead of a closed one.
  const rendererFor = (): SheetTileRenderer => {
    const current = rendererRef.current
    if (current && current.surfaceBytes === surfaceBytes && current.profile === profile) return current
    current?.dispose()
    return rendererRef.current = new SheetTileRenderer(profile, surfaceBytes)
  }
  useEffect(() => () => { rendererRef.current?.dispose(); rendererRef.current = null }, [])
  const markKey = watermark ? JSON.stringify(watermark) : ''
  const mark = useMemo(() => normalizeWatermark(watermark), [markKey]) // eslint-disable-line react-hooks/exhaustive-deps
  const biasRef = useRef({ x: 0, y: 0 })
  const overlayPaints = useRef(0)
  const frameOf = (): SheetFrame | null => entry && scroll ? {
    sheet: entry.sheet, geometry: entry.geometry, identity, origin: scroll.logicalOrigin, viewport, zoom: scroll.zoom, dpr,
    extent: { width: entry.metrics.cols.totalCss(), height: entry.metrics.rows.totalCss() },
    assets: assets?.lease.assets, layoutCache: entry.layoutCache, merges: entry.merges, ...(mark ? { watermark: mark } : {}),
  } : null
  const report = () => {
    if (!props.onFrame || !scroll) return
    const renderer = rendererFor(), stats = renderer.stats
    props.onFrame({
      basePaints: stats.basePaints, tileHits: stats.tileHits, prefetched: stats.prefetched, refused: stats.refused,
      overlayPaints: overlayPaints.current, lastRequests: stats.lastRequests.map(request => ({ ...request.box })),
      tiles: renderer.cache.stats(), bitmapBytes: renderer.bitmapBytes(), peakBitmapBytes: stats.peakBitmapBytes,
      budgetBytes: renderer.budgetBytes, frameRecords: stats.frameRecords, frameGridlines: stats.frameGridlines, prefetchRecords: stats.prefetchRecords, activeSheetId, zoom: scroll.zoom, identity,
    })
  }

  // Base layer: keyed by owner/position/zoom/DPR/pane only; overlay changes never re-run it.
  useLayoutEffect(() => {
    const canvas = baseRef.current, frame = frameOf()
    if (!canvas || !frame || !assets || admittedDpr === null) return
    // The pool (and its admission) exists before the pane surface is sized.
    const renderer = rendererFor()
    if (canvas.width !== device.width) canvas.width = device.width
    if (canvas.height !== device.height) canvas.height = device.height
    const ctx = canvas.getContext('2d')
    if (!ctx) return
    biasRef.current = renderer.present(ctx, frame).bias
    report()
    // Bounded adjacent prefetch after the visible frame is up.
    const timer = setTimeout(() => { renderer.prefetch(frame); report() }, 0)
    return () => clearTimeout(timer)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [identityKey, scroll, viewport, dpr, assets, surfaceBytes, profile, mark])

  // ---- find (workbook-wide, bounded pages) ----
  const [query, setQuery] = useState('')
  const [page, setPage] = useState<FindPublication | null>(null)
  const [activeHit, setActiveHit] = useState(-1)
  const [searching, setSearching] = useState(false)
  const [announcement, setAnnouncement] = useState('')
  const [inspection, setInspection] = useState<HiddenCellInspection | undefined>(undefined)
  const inspectHidden = (sheetId: string, target: SourceCell, reason: string) =>
    setInspection({ target, hiddenReason: reason === 'filter-hidden' ? 'hidden by filter' : 'authored hidden', source: source.store(sheetId) })
  const inputRef = useRef<HTMLInputElement | null>(null)
  const finder = useMemo(() => new FindController({ currentOwner: () => ownerRef.current() }), [])
  useEffect(() => () => finder.dispose(), [finder])
  const cellAt = (sheetEntry: SheetEntry, at: SourceCell): XlsxCell | undefined => sheetEntry.geometry.cellsOf(at.row)?.find(cell => cell.col === at.col)
  const findInput = (): FindInput => ({
    owner: ownerOf(),
    sheets: sheetIds.map(sheetId => {
      const sheetEntry = entryFor(sheetId)
      return {
        sheetId, store: withUnavailableFormulas(source.store(sheetId)), identity: identityOf(sheetId), order: sheetEntry.view.order, visibility: sheetEntry.view.visibility,
        // Authored-hidden (zero-width) columns are outside the visible search domain (§11.4).
        isColumnVisible: (col: number) => sheetEntry.metrics.cols.size(col) > 0,
      }
    }),
  })
  // The arguments of the current search, so a backward/wrapping step can restart it.
  const runArgs = useRef<{ input: FindInput; query: string; options: FindOptions } | null>(null)
  useEffect(() => {
    // Only an empty query cancels: whitespace is a literal search (core/search parity).
    if (query.length === 0) { finder.cancel(); runArgs.current = null; setPage(null); setActiveHit(-1); setSearching(false); return }
    let cancelled = false
    setSearching(true)
    const input = findInput()
    const byId = new Map(sheetIds.map(id => [id, entryFor(id)]))
    const options: FindOptions = {
      // The same canonical display text the cell paints (authored formats, semantics).
      textOf: (cell, sheetId) => {
        const sheetEntry = byId.get(sheetId)
        const styled = sheetEntry && cellAt(sheetEntry, cell)
        return sheetEntry ? cellDisplayText(sheetEntry.sheet, styled ?? cell) : String(cell.value ?? '')
      },
    }
    runArgs.current = { input, query, options }
    void finder.run(input, { text: query, caseSensitive: false }, options).then(published => {
      if (cancelled) return
      setPage(published)
      setActiveHit(published && published.hits.length ? 0 : -1)
      setSearching(false)
    }).catch(error => {
      if (cancelled) return
      setSearching(false)
      setAnnouncement(error instanceof Error ? error.message.slice(0, 256) : 'Search failed.')
    })
    return () => { cancelled = true }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [query, ownerKey])
  const pageOffset = page ? page.pageIndex * FIND_PAGE_HITS : 0
  const matchCount = page ? page.totalKnown ?? pageOffset + page.hits.length : 0
  // Partial results present a lower bound, never a final-looking count (§10.4).
  // A total learned by an earlier complete traversal of the same run stays final.
  const knownTotal = useRef<{ key: string; total: number } | null>(null)
  if (page?.complete && page.totalKnown !== undefined) knownTotal.current = { key: page.key, total: page.totalKnown }
  const learned = page && knownTotal.current?.key === page.key ? knownTotal.current.total : undefined
  const countLabel = page && !page.complete ? (learned !== undefined ? String(learned) : `${pageOffset + page.hits.length}+`) : undefined
  const searchNote = page ? findNote(page) : undefined
  /** Bring page `target` (or the last page) in, one bounded page at a time: forward
   * pulls continue the run; anything earlier restarts it from page 0. */
  const loadPage = async (target: number | 'last'): Promise<FindPublication | null> => {
    let current = page
    const done = (candidate: FindPublication | null) => !candidate || (target === 'last' ? candidate.complete : candidate.pageIndex === target)
    if (done(current)) return current
    if (!current || (target !== 'last' && current.pageIndex > target) || !finder.running) {
      const args = runArgs.current
      if (!args) return null
      current = await finder.run(args.input, { text: args.query, caseSensitive: false }, args.options)
    }
    while (current && !done(current) && finder.running) current = await finder.next()
    return current
  }
  const stepHit = async (direction: 1 | -1) => {
    if (!page || !page.hits.length) return
    const next = activeHit + direction
    if (next >= 0 && next < page.hits.length) { setActiveHit(next); return }
    let target: FindPublication | null
    if (direction > 0) target = page.complete ? await loadPage(0) : await loadPage(page.pageIndex + 1)
    else target = page.pageIndex > 0 ? await loadPage(page.pageIndex - 1) : await loadPage('last')
    if (!target || !target.hits.length) return
    setPage(target)
    setActiveHit(direction > 0 ? 0 : target.hits.length - 1)
  }
  const goToHit = (hit: FindHit) => {
    if (!sheetIds.includes(hit.sheetId)) return
    const sheetEntry = entryFor(hit.sheetId)
    if (!windows.has(hit.sheetId)) windowFor(hit.sheetId, sheetEntry)
    const result = navigateToHit({ session, windows, viewport }, hit)
    if (result.kind === 'needs-projection') return
    if (result.kind === 'hidden') inspectHidden(hit.sheetId, hit.source, result.hiddenReason)
    else setInspection(undefined)
    session.activate(hit.sheetId)
    setActiveSheetId(hit.sheetId)
    setAnnouncement(result.announcement)
    rerender()
  }
  useEffect(() => {
    const hit = page?.hits[activeHit]
    if (hit) goToHit(hit)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [page, activeHit])

  // Overlay layer: selection and search highlights in viewport space.
  useLayoutEffect(() => {
    const canvas = overlayRef.current
    if (!canvas || !entry || !scroll || admittedDpr === null) return
    if (canvas.width !== device.width) canvas.width = device.width
    if (canvas.height !== device.height) canvas.height = device.height
    const ctx = canvas.getContext('2d')
    if (!ctx) return
    overlayPaints.current++
    ctx.setTransform(1, 0, 0, 1, 0, 0)
    ctx.clearRect(0, 0, canvas.width, canvas.height)
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
    const pane = visibleLogicalViewport(scroll, viewport)
    const bias = biasRef.current
    const toViewport = (rect: HighlightRect): HighlightRect => ({ x: (rect.x - pane.x) * scroll.zoom + bias.x, y: (rect.y - pane.y) * scroll.zoom + bias.y, width: rect.width * scroll.zoom, height: rect.height * scroll.zoom })
    // Search: lazily materialized geometry for this page's hits on the pane only.
    const highlights: HighlightRect[] = [], active: HighlightRect[] = []
    if (page) {
      const host = createCellLayoutHost(ctx, entry.sheet)
      page.hits.forEach((hit, index) => {
        if (hit.sheetId !== activeSheetId) return
        const box = entry.metrics.sourceCellToLogical(hit.source)
        if (!box || box.x >= pane.x + pane.width || box.x + box.width <= pane.x || box.y >= pane.y + pane.height || box.y + box.height <= pane.y) return
        const cell = cellAt(entry, hit.source)
        const geometry = cell && hitGeometry(hit, cell, entry.metrics, identity, host, entry.layoutCache)
        if (!geometry) return
        const rects = geometry.rects.map(toViewport)
        highlights.push(...rects)
        if (index === activeHit) active.push(...rects)
      })
    }
    paintSearchOverlay(ctx, highlights, active)
    if (selection.ranges.length) {
      const records = overlayRecords(entry.metrics, pane, profile)
      // Past the per-cell cap the membership is coalesced into pixel groups, never dropped.
      const geometry = records
        ? gridSelectionOverlay({ contains: at => selectionContains(selection, at) }, records, pane, scroll.zoom * dpr)
        : gridSelectionOverlay({ contains: () => true }, coarseSelectionRecords(selection, entry.metrics, pane, scroll.zoom * dpr), pane, scroll.zoom * dpr)
      paintGridSelectionOverlay(ctx, gridGeometryToViewport(geometry, pane, scroll.zoom, bias), dpr)
    }
    report()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [scroll, selection, page, activeHit, viewport, dpr, identityKey])

  // Keep the native scroller on the window's physical segment (Go To, keys, recenter).
  const syncing = useRef(false)
  useLayoutEffect(() => {
    const el = scrollerRef.current
    if (!el || !scroll) return
    if (Math.abs(el.scrollLeft - scroll.physicalPosition.x) >= 0.5 || Math.abs(el.scrollTop - scroll.physicalPosition.y) >= 0.5) {
      syncing.current = true
      el.scrollLeft = scroll.physicalPosition.x
      el.scrollTop = scroll.physicalPosition.y
    }
  }, [scroll])
  const onScroll = () => {
    const el = scrollerRef.current
    if (!el || !entry || !scroll) return
    if (syncing.current) { syncing.current = false; if (Math.abs(el.scrollLeft - scroll.physicalPosition.x) < 0.5 && Math.abs(el.scrollTop - scroll.physicalPosition.y) < 0.5) return }
    setWindow(activeSheetId, scrollPhysicalTo(scroll, entry.metrics, viewport, { x: el.scrollLeft, y: el.scrollTop }))
  }

  // ---- zoom: Ctrl/Cmd+wheel, pinch and keys, about an anchor ----
  const zoomTo = (next: ScrollWindow | null) => { if (next && scroll && next !== scroll) setWindow(activeSheetId, next) }
  const scrollRef = useRef({ scroll, entry, viewport, activeSheetId })
  scrollRef.current = { scroll, entry, viewport, activeSheetId }
  useEffect(() => {
    const el = scrollerRef.current
    if (!el) return
    const onWheel = (event: WheelEvent) => {
      if (!event.ctrlKey && !event.metaKey) return
      event.preventDefault()
      const { scroll: current, entry: sheetEntry, viewport: size, activeSheetId: sheetId } = scrollRef.current
      if (!current || !sheetEntry) return
      const bounds = el.getBoundingClientRect()
      const deltaY = event.deltaMode === 1 ? event.deltaY * 16 : event.deltaMode === 2 ? event.deltaY * size.height : event.deltaY
      const next = wheelSheetZoom(current, sheetEntry.metrics, size, { deltaY, ctrlKey: event.ctrlKey, metaKey: event.metaKey }, { x: event.clientX - bounds.left, y: event.clientY - bounds.top })
      if (next !== current) setWindow(sheetId, next)
    }
    el.addEventListener('wheel', onWheel, { passive: false })
    return () => el.removeEventListener('wheel', onWheel)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [entry !== null])
  const pointers = useRef(new Map<number, { x: number; y: number }>())
  const pinch = useRef<{ start: ScrollWindow; distance: number; anchor: { x: number; y: number } } | null>(null)
  const local = (event: { clientX: number; clientY: number }) => {
    const bounds = scrollerRef.current?.getBoundingClientRect()
    return { x: event.clientX - (bounds?.left ?? 0), y: event.clientY - (bounds?.top ?? 0) }
  }
  const pinchState = () => {
    const [a, b] = [...pointers.current.values()]
    return { distance: Math.hypot(a.x - b.x, a.y - b.y), anchor: { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 } }
  }
  const onPointerDown = (event: React.PointerEvent) => {
    if (event.pointerType !== 'touch') return
    pointers.current.set(event.pointerId, local(event))
    if (pointers.current.size === 2 && scroll) pinch.current = { start: scroll, ...pinchState() }
  }
  const onPointerMove = (event: React.PointerEvent) => {
    if (!pointers.current.has(event.pointerId)) return
    pointers.current.set(event.pointerId, local(event))
    const gesture = pinch.current
    if (pointers.current.size !== 2 || !gesture || !entry) return
    event.preventDefault()
    const now = pinchState()
    zoomTo(pinchSheetZoom(gesture.start, entry.metrics, viewport, gesture.distance, now.distance, gesture.anchor, now.anchor))
  }
  const onPointerEnd = (event: React.PointerEvent) => {
    pointers.current.delete(event.pointerId)
    if (pointers.current.size < 2) pinch.current = null
  }
  const zoomBy = (direction: 1 | -1) => { if (entry && scroll) zoomTo(keyboardSheetZoom(scroll, entry.metrics, viewport, direction)) }
  const resetZoom = () => { if (entry && scroll) zoomTo(setScrollZoom(scroll, entry.metrics, viewport, 1, { x: viewport.width / 2, y: viewport.height / 2 })) }

  const onKeyDown = (event: React.KeyboardEvent) => {
    const mod = event.metaKey || event.ctrlKey, key = event.key.toLowerCase()
    if (mod && key === 'f') { event.preventDefault(); inputRef.current?.focus(); inputRef.current?.select(); return }
    if (!allowCopy && mod && key === 'c') { event.preventDefault(); return }
    if (!allowPrint && mod && key === 'p') { event.preventDefault(); return }
    // Zoom keys belong to the grid/viewer, never to inputs or header controls.
    const target = event.target as HTMLElement
    if (!(target === event.currentTarget || target.hasAttribute?.('data-officeview-grid')) || mod || event.altKey) return
    if (event.key === '+' || event.key === '=') { event.preventDefault(); zoomBy(1) }
    else if (event.key === '-') { event.preventDefault(); zoomBy(-1) }
    else if (event.key === '0') { event.preventDefault(); resetZoom() }
  }

  // Ctrl/Cmd+wheel anywhere over the viewer (headers, navigator) never zooms the
  // browser page; React's wheel listeners are passive, so this one is native.
  const rootRef = useRef<HTMLDivElement | null>(null)
  useEffect(() => {
    const el = rootRef.current
    if (!el) return
    const suppress = (event: WheelEvent) => { if (event.ctrlKey || event.metaKey) event.preventDefault() }
    el.addEventListener('wheel', suppress, { passive: false })
    return () => el.removeEventListener('wheel', suppress)
  }, [])

  // Print guard while this viewer is mounted (same rule as the page viewer).
  useEffect(() => {
    if (allowPrint) return
    const host = globalThis.document
    const guard = host.createElement('style')
    guard.setAttribute('data-officeview-print-guard', '')
    guard.textContent = '@media print { [data-officeview-root] { display: none !important; } }'
    host.head.appendChild(guard)
    return () => guard.remove()
  }, [allowPrint])

  const onGridChange = (transition: GridTransition) => {
    session.setSelection(activeSheetId, transition.selection)
    setAnnouncement(transition.announcement)
    setInspection(undefined)
    if (transition.window !== scroll) setWindow(activeSheetId, transition.window)
    else rerender()
  }
  const onNavigate = (request: NavigatorRequest) => {
    if (!entry || !scroll) return
    const state = session.get(activeSheetId)
    const result = navigateSelection(state.selection, entry.metrics, request, entry.view)
    setAnnouncement(result.announcement)
    if (result.kind === 'hidden') { inspectHidden(activeSheetId, result.target, result.hiddenReason ?? 'authored-hidden'); return }
    setInspection(undefined)
    const moved = scrollNavigateTo(scroll, { metrics: entry.metrics, viewport, view: { ...entry.view, selection: result.selection } }, request, state.selection.active)
    session.setSelection(activeSheetId, result.selection)
    setWindow(activeSheetId, moved.window)
  }
  const onSheetChange = (sheetId: string) => {
    if (!sheetIds.includes(sheetId)) return
    entryFor(sheetId)
    setInspection(undefined)
    session.activate(sheetId)
    setActiveSheetId(sheetId)
  }

  const zoom = scroll?.zoom ?? 1
  const canvasStyle: CSSProperties = { position: 'absolute', left: 0, top: 0, width: viewport.width, height: viewport.height, pointerEvents: 'none' }
  return (
    <div style={{ position: 'relative', display: 'flex', flexDirection: 'column', minHeight: 0 }}>
      <div
        ref={rootRef}
        className={props.className}
        role="region"
        aria-label="Spreadsheet viewer"
        data-officeview-root=""
        data-officeview-sheet-mode=""
        tabIndex={-1}
        onKeyDown={onKeyDown}
        onCopyCapture={allowCopy ? undefined : event => event.preventDefault()}
        onContextMenu={allowCopy ? undefined : event => event.preventDefault()}
        style={{
          background: props.background ?? '#888888', display: 'flex', flexDirection: 'column', minHeight: 0, flex: '1 1 auto',
          WebkitUserSelect: allowCopy ? undefined : 'none', userSelect: allowCopy ? undefined : 'none', ...props.style,
        }}
      >
        <SheetNavigator documentId={docId} sheets={doc.sheets.map((sheet, index) => ({ id: sheetIds[index], name: sheet.name }))}
          activeSheetId={activeSheetId} onSheetChange={onSheetChange} onNavigate={onNavigate} onNotice={setAnnouncement} disabled={!entry} />
        <div ref={paneRef} data-testid="officeview-sheet-pane" style={{ position: 'relative', flex: '1 1 auto', minHeight: 0, overflow: 'hidden', background: '#ffffff' }}>
          {entry && scroll && (
            <GridSemantics owner={identity} projection={entry.metrics} view={entry.view} selection={selection} window={scroll} viewport={viewport}
              body={entry.body} onChange={onGridChange} announcement={[resourceNotice, announcement].filter(Boolean).join(' ')} inspection={inspection} profile={profile} label={`Worksheet ${entry.sheet.name}`}
              cellPresentation={at => {
                const cell = cellAt(entry, at)
                return { text: cell ? cellDisplayText(entry.sheet, cell) : '' }
              }}>
              <canvas ref={baseRef} data-testid="officeview-sheet-base" style={canvasStyle} />
              <canvas ref={overlayRef} data-testid="officeview-sheet-overlay" style={canvasStyle} />
              <div ref={scrollerRef} data-testid="officeview-sheet-scroller" onScroll={onScroll}
                onPointerDown={onPointerDown} onPointerMove={onPointerMove} onPointerUp={onPointerEnd} onPointerCancel={onPointerEnd}
                style={{ position: 'absolute', inset: 0, overflow: 'auto', touchAction: 'pan-x pan-y', overscrollBehavior: 'contain' }}>
                <div style={{ width: scroll.physicalExtent.width, height: scroll.physicalExtent.height }} />
              </div>
            </GridSemantics>
          )}
        </div>
      </div>
      {showSearch && (
        <SearchBar inputRef={inputRef} query={query} onQuery={setQuery} count={matchCount} active={activeHit < 0 ? -1 : pageOffset + activeHit}
          countLabel={countLabel} note={searchNote}
          indexing={searching} onPrev={() => void stepHit(-1)} onNext={() => void stepHit(1)} />
      )}
      {showZoomControls && (
        <ZoomControls zoom={zoom} onIn={() => zoomBy(1)} onOut={() => zoomBy(-1)} onReset={resetZoom} atFit={zoom <= SHEET_MIN_ZOOM} />
      )}
    </div>
  )
}

/** TODO(S1.3): the interim object source reports a formula without a cached or
 * evaluated value as a null cell; find must count it as unavailable (§4.4), never
 * as a blank that silently cannot match. */
function withUnavailableFormulas(store: SourceStore): SourceStore {
  const mark = (cell: AuthoritativeCell): AuthoritativeCell =>
    cell.formula !== undefined && cell.value === null && cell.hasCachedValue === false && !cell.unavailable ? { ...cell, unavailable: true } : cell
  return Object.assign(Object.create(store) as SourceStore, {
    async *scanPopulated(request: Parameters<SourceStore['scanPopulated']>[0], options: Parameters<SourceStore['scanPopulated']>[1]) {
      for await (const page of store.scanPopulated(request, options)) yield { ...page, items: page.items.map(mark) }
    },
  })
}

/** Status qualifiers retained with a published find page (§§10.4, 11.4). */
function findNote(page: FindPublication): string | undefined {
  const notes: string[] = []
  if (page.unavailableCount > 0) notes.push(`${page.unavailableCount} value${page.unavailableCount === 1 ? '' : 's'} unavailable`)
  if (page.failures.length) notes.push(`${page.failures.length} source failure${page.failures.length === 1 ? '' : 's'}; retry may help`)
  else if (page.incomplete.length) notes.push('some sheets still loading')
  return notes.length ? notes.join('; ') : undefined
}

/** Resolved used-data body for record counts (never the implicit million rows). */
function usedBody(sheet: XlsxSheet): SourceRowInterval | null {
  let r0 = Number.POSITIVE_INFINITY, r1 = -1
  for (const row of sheet.rows) {
    if (!row.cells.some(cell => cell.value !== null || cell.formula !== undefined)) continue
    r0 = Math.min(r0, row.index)
    r1 = Math.max(r1, row.index)
  }
  return r1 < 0 ? null : { r0, r1 }
}
