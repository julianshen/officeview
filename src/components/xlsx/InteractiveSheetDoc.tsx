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
 * Worker-backed sheets (S5.4): a paged admission arrives as a shell document
 * plus a RemoteWorkbook source. Its cells come from bounded worker reads into a
 * row cache (unknown rows paint pending), its views commit in the worker (sort/
 * filter via viewerControls' remote port) and are adopted here, and find and
 * copy read through the same source; the object path is unchanged.
 */
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type CSSProperties, type ReactElement } from 'react'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../../xlsx/types'
import type {
  AuthoritativeCell, CacheIdentity, FilterSpec, GridSelection, LogicalCellBox, ProfileId, RowOrder, RowVisibility, ScrollWindow, SheetView, SortSpec, SourceRowInterval, SourceStore, WorkbookOwner,
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
import { SheetNavigationSession, emptyGridSelection, navigateSelection, recordCounts, boundedSelectionText } from '../../xlsx/view/selection'
import { FIND_PAGE_HITS, FindController, hitGeometry, navigateToHit, type FindHit, type FindInput, type FindOptions, type FindPublication } from '../../xlsx/view/find'
import { LayoutCache } from '../../xlsx/view/textLayout'
import { cellDisplayText } from '../../xlsx/render'
import { createCellLayoutHost } from '../../xlsx/view/renderViewport'
import { createObjectWorkbookSource } from '../../worker/viewerObjectSource'
import { BITMAP_BUDGETS, type TileCacheStats } from '../../xlsx/view/tileCache'
import { gridGeometryToViewport, gridSelectionOverlay, paintGridSelectionOverlay, paintSearchOverlay, type HighlightRect } from '../../core/overlay'
import { normalizeWatermark, type WatermarkOptions } from '../../core/watermark'
import type { FallbackFontsOptions } from '../../core/fonts/fallback'
import type { FontDiagnostic } from '../../core/fonts/types'
import { SheetImageCache, imageBudgetFor, preparePrintUnits, prepareSheetAssets, type PrintUnitsOptions, type SheetAssetsLease } from './sheetAssets'
import type { Paintable, PaintableArray } from '../../render/paint'
import { GridSemantics, type GridTransition } from './GridSemantics'
import { GRID_HEADER_SIZE, columnHeaderHeight } from './GridHeaders'
import { useControlMedia } from './controlMedia'
import { useDevicePixelRatio } from './devicePixelRatio'
import { FilterDropdown } from './FilterDropdown'
import { SortDialog } from './SortDialog'
import { createViewerControlResources, viewerControls, type ViewerControlResources, type ViewerControls } from './viewerControlsClient'
import type { ControlSnapshot } from './controlsClient'
import { failureText } from './controlUi'
import { copySelection, copySessionBudget } from '../../xlsx/view/copy'
import { LARGE_EXPORT_MESSAGE } from '../../render/paint'
import { SAVED_VALUES_DIAGNOSTIC } from '../../xlsx/view/values'
import type { WorkbookSource } from '../../worker/viewerSession'
import { coarseSelectionRecords, overlayRecords, selectionContains } from './selectionOverlay'
import { SheetNavigator } from './SheetNavigator'
import type { HiddenCellInspection } from './ViewerStatus'
import { SheetTileRenderer, TILE_BYTES, admitDeviceScale, paneSurfaceBytes, type SheetFrame } from './sheetTiles'
import { SearchBar, ZoomControls } from '../viewerControls'
import { isRemoteWorkbook } from '../../worker/viewerRemoteSource'
import { pagedSheetGeometry } from './pagedGeometry'

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
  maxTileRecords: number
  /** Visible tiles of the latest frame still showing pending (loading) records. */
  pendingTiles: number
  activeSheetId: string
  zoom: number
  identity: CacheIdentity
}
export interface InteractiveSheetDocProps {
  doc: XlsxDocument
  /** @internal Authoritative source port; the viewer owns its release. */
  source?: WorkbookSource
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
  /** @internal Asset lease factory (tests inject delayed leases). */
  assetLoader?: typeof prepareSheetAssets
  /** @internal Static print unit factory (tests inject delayed units). */
  printLoader?: (doc: XlsxDocument, options: PrintUnitsOptions) => Promise<PaintableArray & { refusedImages?: number[] }>
}

// Document identity alone owns zoom reset (§10.1): a stable id per document object.
const documentIds = new WeakMap<object, string>()
let nextDocumentId = 0
function documentIdOf(doc: object): string {
  let id = documentIds.get(doc)
  if (!id) documentIds.set(doc, id = `xlsx-${++nextDocumentId}`)
  return id
}
export function sheetIdOf(sheet: XlsxSheet, index: number): string { return sheet.sheetId ?? String(index + 1) }

/** The failed-load notice with its recovery path: a storage failure ended the
 * import and removed its pages, so only reopening after freeing space helps. */
function loadFailureNotice(message: string, code?: string): string {
  const recovery = code === 'quota' || code === 'storage-error'
    ? ' Free some browser storage for this site (or use a smaller file), then reopen the file.'
    : ' Reopen the file to try again.'
  return `This sheet could not be fully loaded: ${message}.${recovery}`
}
const MAX_FONT_NOTICE_FAMILIES = 3
function fontNotice(diagnostics: readonly FontDiagnostic[]): string[] {
  const families = [...new Set(diagnostics.map(d => d.family).filter((family): family is string => Boolean(family)))]
  if (!families.length) return []
  const named = families.slice(0, MAX_FONT_NOTICE_FAMILIES).map(family => boundedSelectionText(family)).join(', ')
  const more = families.length > MAX_FONT_NOTICE_FAMILIES ? ` and ${families.length - MAX_FONT_NOTICE_FAMILIES} more` : ''
  return [`Font${families.length === 1 ? '' : 's'} not available, shown with a fallback: ${named}${more}`]
}

export default function InteractiveSheetDoc(props: InteractiveSheetDocProps): ReactElement {
  const docId = props.source?.store(props.source.sheets[0].sheetId).stamp.docId ?? documentIdOf(props.doc)
  return <SheetViewer key={`${documentIdOf(props.doc)}:${props.source ? documentIdOf(props.source) : ''}`} docId={docId} {...props} />
}

interface SheetEntry {
  sheet: XlsxSheet
  revision: number
  /** Worker-backed sheets: generation of the merged ranges the geometry holds. */
  structureGen: number
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
  const source = useMemo(() => props.source ?? createObjectWorkbookSource(doc, docId), [doc, docId, props.source])
  useEffect(() => () => { void source.release() }, [source])
  // Worker-backed source: row loads, merges and adopted worker views re-render.
  const remote = isRemoteWorkbook(source) ? source : null
  const [loadTick, setLoadTick] = useState(0)
  useEffect(() => remote?.subscribe(() => setLoadTick(tick => tick + 1)), [remote])
  // Tiles painted under an older worker layout/merge generation are stale (§10.1).
  const structureGens = useRef(new Map<string, number>()).current

  // ---- fonts/images: one lease per (fallback fonts, active sheet) ----
  // A font publication bumps fontGen (re-keys tiles, keeps zoom). Leases swap
  // atomically: the current one stays alive, and keeps painting, until its
  // admitted replacement has committed; only then is it disposed (§10.1).
  const fontKey = fallbackFonts ? JSON.stringify(fallbackFonts) : ''
  const [assets, setAssets] = useState<{ lease: SheetAssetsLease; fontGen: number; fontKey: string; sheetId: string } | null>(null)
  const fontGenRef = useRef<{ key: string; gen: number }>({ key: fontKey, gen: 0 })
  const loadAssets = props.assetLoader ?? prepareSheetAssets
  // One image budget shared by every lease of this viewer (overlapping swaps included).
  const imageCache = useRef<SheetImageCache | null>(null)
  const imageCacheFor = () => {
    if (!imageCache.current || imageCache.current.limitBytes !== imageBudgetFor(profile)) imageCache.current = new SheetImageCache(imageBudgetFor(profile))
    return imageCache.current
  }
  // A sheet switch stops presenting the old sheet at once (no stale raster under the
  // new sheet's headers), so its lease is released before the new sheet decodes.
  useEffect(() => { setAssets(current => current && current.sheetId !== activeSheetId ? null : current) }, [activeSheetId])
  useEffect(() => {
    if (!activeSheetId) return
    let cancelled = false
    const sheetIndex = sheetIds.indexOf(activeSheetId)
    void loadAssets(doc, { ...(fontKey ? { fallbackFonts } : {}), sheetIndex, profile, imageCache: imageCacheFor() }).then(next => {
      if (cancelled) { next.dispose(); return }
      if (fontGenRef.current.key !== fontKey || fontGenRef.current.gen === 0) fontGenRef.current = { key: fontKey, gen: fontGenRef.current.gen + 1 }
      setAssets({ lease: next, fontGen: fontGenRef.current.gen, fontKey, sheetId: activeSheetId })
    })
    return () => { cancelled = true }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [doc, fontKey, activeSheetId, profile])
  const committedLease = useRef<SheetAssetsLease | null>(null)
  useEffect(() => {
    const previous = committedLease.current
    committedLease.current = assets?.lease ?? null
    if (previous && previous !== committedLease.current) previous.dispose()
  }, [assets])
  useEffect(() => () => { committedLease.current?.dispose(); committedLease.current = null }, [])
  // The published font generation; a sheet switch (lease pending) never resets it.
  const fontGen = fontGenRef.current.gen

  // ---- per-sheet committed views and geometry (one geometry per viewGen) ----
  const localViews = useRef(new Map<string, SheetViewInput>()).current
  const viewInput = (sheetId: string): SheetViewInput => {
    if (remote) {
      // The worker owns a paged sheet's view; its revision is the worker viewGen.
      const view = remote.view(sheetId)
      return { revision: view.revision, order: view.order, ...(view.visibility ? { visibility: view.visibility } : {}), filter: view.filter, sort: view.sort }
    }
    const local = localViews.get(sheetId), external = views?.[sheetId]
    return local && (!external || local.revision >= external.revision) ? local : external ?? { revision: 0 }
  }
  const entryFor = (sheetId: string): SheetEntry => {
    const index = sheetIds.indexOf(sheetId)
    const sheet = doc.sheets[index]
    const input = viewInput(sheetId)
    const previous = entries.get(sheetId)
    const structureGen = remote ? remote.merges(sheetId).gen : 0
    if (previous && previous.revision === input.revision && previous.structureGen === structureGen) return previous
    if (previous && input.revision < previous.revision) return previous // stale commit: never publish an older view
    const order = input.order ?? createIdentityOrder()
    const baseHidden = remote ? remote.geometry(sheetId).baseHidden : sheetBaseHidden(sheet)
    const visibility = input.visibility ?? visibilityFromHidden(createIntervalSet(displayIntervals(baseHidden.intervals, order)))
    const built = remote ? pagedSheetGeometry(remote, sheetId, order, visibility) : buildSheetGeometry(sheet, { order, ...(input.visibility ? { visibility } : {}) })
    if (built.kind !== 'ready') throw new Error(built.diagnostics.map(d => d.message).join('; ') || 'This sheet cannot be laid out')
    const selection = previous?.view.selection ?? emptyGridSelection()
    const view: SheetView = { sheetId, revision: input.revision, order, baseHidden, visibility, sort: input.sort ?? null, filter: input.filter ?? null, selection }
    const used = remote?.layout(sheetId).used
    const entry: SheetEntry = {
      sheet, revision: input.revision, structureGen, view, geometry: built, metrics: built.metrics, body: input.filter || input.sort ? { r0: (input.filter?.range ?? input.sort!.range).r0 + (input.filter?.range ?? input.sort!.range).headerRows, r1: (input.filter?.range ?? input.sort!.range).r1 - (input.filter?.range ?? input.sort!.range).totalsRows } : remote ? used ? { r0: used.r0, r1: used.r1 } : null : usedBody(sheet),
      // Layout records are keyed by complete identity; a new view gets its own merge cache.
      // A worker layout/merge publication changes record boxes: fresh layouts then.
      layoutCache: previous && previous.structureGen === structureGen ? previous.layoutCache : new LayoutCache(), merges: new MergeQueryCache(),
    }
    entries.set(sheetId, entry)
    if (!previous) session.register(sheetId, { view, projection: entry.metrics, body: entry.body })
    else session.updateProjection(sheetId, entry.metrics, entry.body, view)
    return entry
  }
  // The committed revision: a stale (lower) prop revision never regresses the
  // owner that keys tiles, find and hit geometry; a newer one is built on use.
  const revisionOf = (sheetId: string): number => Math.max(entries.get(sheetId)?.revision ?? 0, viewInput(sheetId).revision)
  const identityOf = (sheetId: string): CacheIdentity => ({ docId, sheetId, valueGen: source.store(sheetId).stamp.valueGen, fontGen, viewGen: revisionOf(sheetId) })
  const globalViewGen = sheetIds.reduce((sum, id) => sum + revisionOf(id), 0)
  const ownerOf = (): WorkbookOwner => ({
    docId, globalViewGen, workbookValueGen: 0, fontGen,
    sheetVector: sheetIds.map(sheetId => ({ sheetId, revision: revisionOf(sheetId), valueGen: source.store(sheetId).stamp.valueGen })),
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
    if (typeof ResizeObserver === 'undefined') {
      window.addEventListener('resize', update)
      return () => window.removeEventListener('resize', update)
    }
    const ro = new ResizeObserver(update)
    ro.observe(el)
    return () => ro.disconnect()
  }, [])
  const touchControls = useControlMedia('(pointer: coarse)')
  const headerHeight = columnHeaderHeight(touchControls, true, true)
  const viewport: ViewportSize = useMemo(() => ({ width: Math.max(1, pane.width - GRID_HEADER_SIZE), height: Math.max(1, pane.height - headerHeight) }), [pane, headerHeight])
  // Live ratio: a display move re-renders here, re-admits the pool and re-keys tiles (S13 D-DPR).
  const nativeDpr = Math.min(useDevicePixelRatio(), 3)
  // Admit the pane surfaces before sizing any canvas; a lower device scale keeps
  // the pane within the bitmap budget, and nothing fits → a visible resource outcome.
  // Live print stores (while a print is prepared/open) are reserved beside the panes.
  const [printReserved, setPrintReserved] = useState(0)
  const admittedDpr = admitDeviceScale(viewport, nativeDpr, profile, printReserved)
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
  const nextWindowKey = JSON.stringify([activeSheetId, entry?.revision, entry?.structureGen, viewport.width, viewport.height, dpr])
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
  const surfaceBytes = paneSurfaceBytes(viewport, dpr) + printReserved
  const latestGeometry = useRef({ viewport, dpr })
  latestGeometry.current = { viewport, dpr }
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
  // Published by the overlay pass as state, so the status follows it (a declared
  // coarse state); an empty selection always clears it.
  const [selectionIncomplete, setSelectionIncomplete] = useState(false)
  const frameOf = (): SheetFrame | null => entry && scroll ? {
    sheet: entry.sheet, geometry: entry.geometry, identity, origin: scroll.logicalOrigin, viewport, zoom: scroll.zoom, dpr,
    extent: { width: entry.metrics.cols.totalCss(), height: entry.metrics.rows.totalCss() },
    assets: assets?.lease.assets, layoutCache: entry.layoutCache, merges: entry.merges, ...(mark ? { watermark: mark } : {}),
    structureGen: entry.structureGen,
    ...(remote ? { pending: (box: LogicalCellBox) => remote.store(activeSheetId).pendingIn(entry.metrics, box, scroll.zoom * dpr) } : {}),
  } : null
  // Purely observational: it never creates or resizes the pool.
  const report = () => {
    const renderer = rendererRef.current
    if (!props.onFrame || !scroll || !renderer) return
    const stats = renderer.stats
    props.onFrame({
      basePaints: stats.basePaints, tileHits: stats.tileHits, prefetched: stats.prefetched, refused: stats.refused,
      overlayPaints: overlayPaints.current, lastRequests: stats.lastRequests.map(request => ({ ...request.box })),
      tiles: renderer.cache.stats(), bitmapBytes: renderer.bitmapBytes(), peakBitmapBytes: stats.peakBitmapBytes,
      budgetBytes: renderer.budgetBytes, frameRecords: stats.frameRecords, frameGridlines: stats.frameGridlines, prefetchRecords: stats.prefetchRecords, maxTileRecords: stats.maxTileRecords, pendingTiles: stats.pendingTiles, activeSheetId, zoom: scroll.zoom, identity,
    })
  }

  // Base layer: keyed by owner/position/zoom/DPR/pane only; overlay changes never re-run it.
  useLayoutEffect(() => {
    const canvas = baseRef.current, frame = frameOf()
    if (!canvas || !frame || admittedDpr === null) return
    if (!assets || assets.sheetId !== activeSheetId) {
      // The active sheet's lease is pending: present a declared loading state, never
      // another sheet's (or an unleased) raster. The pool is re-admitted for the
      // current surfaces (releasing tiles that no longer fit) before any sizing.
      rendererFor()
      sizeCanvas(canvas, device.width, device.height)
      const ctx = canvas.getContext('2d')
      if (!ctx) return
      ctx.setTransform(1, 0, 0, 1, 0, 0)
      ctx.fillStyle = '#ffffff'
      ctx.fillRect(0, 0, canvas.width, canvas.height)
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
      ctx.fillStyle = '#666666'
      ctx.font = '12px sans-serif'
      ctx.fillText('Loading…', 8, 20)
      return
    }
    // The pool (and its admission) exists before the pane surface is sized.
    const renderer = rendererFor()
    const seenGen = structureGens.get(activeSheetId)
    if (seenGen !== undefined && seenGen !== frame.structureGen) renderer.cache.invalidate(key => key.identity.sheetId === activeSheetId)
    structureGens.set(activeSheetId, frame.structureGen ?? 0)
    sizeCanvas(canvas, device.width, device.height)
    const ctx = canvas.getContext('2d')
    if (!ctx) return
    biasRef.current = renderer.present(ctx, frame).bias
    report()
    // Bounded adjacent prefetch after the visible frame is up.
    const timer = setTimeout(() => { renderer.prefetch(frame); report() }, 0)
    return () => clearTimeout(timer)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [identityKey, scroll, viewport, dpr, assets, surfaceBytes, profile, mark, loadTick, entry?.structureGen])
  // A restored 2D context starts blank, and the cached tile canvases may have
  // been lost with it: drop the tiles and repaint both layers at once.
  const surfacesMounted = Boolean(entry && scroll)
  useEffect(() => {
    const surfaces = [baseRef.current, overlayRef.current].filter((canvas): canvas is HTMLCanvasElement => canvas !== null)
    const restored = () => { rendererRef.current?.cache.invalidate(() => true); setLoadTick(tick => tick + 1) }
    for (const canvas of surfaces) canvas.addEventListener('contextrestored', restored)
    return () => { for (const canvas of surfaces) canvas.removeEventListener('contextrestored', restored) }
  }, [surfacesMounted])

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
        // Worker-backed cells carry a style id; the workbook table resolves it.
        if (sheetEntry && !styled && remote) return cellDisplayText(sheetEntry.sheet, { value: cell.value, style: remote.styleOf(cell.styleIndex) })
        return sheetEntry ? cellDisplayText(sheetEntry.sheet, styled ?? cell) : String(cell.value ?? '')
      },
    }
    runArgs.current = { input, query, options }
    void finder.run(input, { text: query, caseSensitive: false }, options).then(published => {
      if (cancelled) return
      pageRef.current = published
      activeRef.current = published && published.hits.length ? 0 : -1
      setPage(published)
      setActiveHit(activeRef.current)
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
  // Navigation is serialized: each step starts from the position the previous
  // step committed (refs, not render closures), and a step whose search was
  // replaced meanwhile publishes nothing (§10.1).
  const pageRef = useRef(page), activeRef = useRef(activeHit), searchGen = useRef(0), navQueue = useRef<Promise<void>>(Promise.resolve())
  pageRef.current = page
  activeRef.current = activeHit
  useLayoutEffect(() => { searchGen.current++; setFindFailure(null) }, [query, ownerKey])
  // A paging failure of the CURRENT search (e.g. FindError resource-limit) is shown
  // with the results; the queue stays usable for later steps (§§10.4, 11.4).
  const [findFailure, setFindFailure] = useState<{ generation: number; message: string } | null>(null)
  const searchNote = [page ? findNote(page) : undefined, findFailure && findFailure.generation === searchGen.current ? `search stopped: ${findFailure.message}` : undefined].filter(Boolean).join('; ') || undefined
  /** Bring page `target` (or the last page) in, one bounded page at a time: forward
   * pulls continue the run; anything earlier restarts it from page 0. */
  const loadPage = async (from: FindPublication, target: number | 'last'): Promise<FindPublication | null> => {
    let current: FindPublication | null = from
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
  const step = async (direction: 1 | -1, generation: number) => {
    const current = pageRef.current, active = activeRef.current
    if (!current || !current.hits.length) return
    const next = active + direction
    if (next >= 0 && next < current.hits.length) { activeRef.current = next; setActiveHit(next); return }
    let target: FindPublication | null
    if (direction > 0) target = current.complete ? await loadPage(current, 0) : await loadPage(current, current.pageIndex + 1)
    else target = current.pageIndex > 0 ? await loadPage(current, current.pageIndex - 1) : await loadPage(current, 'last')
    if (generation !== searchGen.current || !target || !target.hits.length) return
    const index = direction > 0 ? 0 : target.hits.length - 1
    pageRef.current = target
    activeRef.current = index
    setPage(target)
    setActiveHit(index)
  }
  const stepHit = (direction: 1 | -1) => {
    // Ownership is captured when the step is ENQUEUED: a step queued for a search
    // that has since been replaced is dropped before it runs.
    const generation = searchGen.current
    const run = navQueue.current.then(async () => {
      if (generation !== searchGen.current) return
      try {
        setFindFailure(null)
        await step(direction, generation)
      } catch (error) {
        if (generation !== searchGen.current) return
        setFindFailure({ generation, message: error instanceof Error ? error.message.slice(0, 256) : 'Search paging failed.' })
      }
    })
    navQueue.current = run
    return run
  }
  const goToHit = (hit: FindHit) => {
    if (!sheetIds.includes(hit.sheetId)) return
    cancelCopy()
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

  // ---- atomic control transport and copy ownership ----
  const current = useRef({ entryFor, identityOf, activeSheetId, allowCopy, allowPrint })
  current.current = { entryFor, identityOf, activeSheetId, allowCopy, allowPrint }
  const snapshot = (sheetId: string): ControlSnapshot => {
    const e = current.current.entryFor(sheetId), body = e.body
    return { owner: current.current.identityOf(current.current.activeSheetId), filter: e.view.filter, sort: e.view.sort, records: { ...recordCounts(e.metrics, body), completeness: 'complete' } }
  }
  // Only the active sheet keeps a worker row cache (§12.4).
  useEffect(() => { remote?.activate(activeSheetId) }, [remote, activeSheetId])
  const [controls, setControls] = useState<ViewerControls | null>(null)
  const controlsReady = useRef<Promise<ViewerControls> | null>(null)
  const controlPorts = useRef(new Map<string, ViewerControls>()).current
  const controlResources = useRef<ViewerControlResources | null>(null)
  useEffect(() => {
    const reused = controlPorts.get(activeSheetId)
    setControls(reused ?? null)
    if (reused) {
      controlPorts.delete(activeSheetId); controlPorts.set(activeSheetId, reused)
      controlsReady.current = Promise.resolve(reused)
      return
    }
    const abort = new AbortController()
    const resources = controlResources.current ??= createViewerControlResources(profile)
    const sheetId = activeSheetId
    const ready = viewerControls({ doc, source, resources, sheetId, snapshot: () => snapshot(sheetId),
      view: () => ({ ...current.current.entryFor(sheetId).view, selection: session.get(sheetId).selection }),
      publish: draft => {
        cancelCopy()
        // A worker-owned view was already adopted from the worker's delta.
        if (!remote) localViews.set(sheetId, { revision: current.current.identityOf(sheetId).viewGen + 1, ...draft })
        current.current.entryFor(sheetId)
        rerender()
        return snapshot(sheetId)
      },
      ...(remote ? { styleOf: (index: number) => remote.styleOf(index), remote: {
        sort: (spec, signal, progress) => remote.sort(remote.identity(sheetId), spec, signal, progress),
        filter: (spec, keySets, signal, progress) => remote.filter(remote.identity(sheetId), spec, keySets, signal, progress),
      } } : {}),
    }, abort.signal)
    controlsReady.current = ready
    void ready.then(port => {
      if (abort.signal.aborted) { port.dispose(); return }
      controlPorts.set(activeSheetId, port)
      // Retain the active port and one recent sheet, not a full index per visit.
      while (controlPorts.size > 2) {
        const oldest = controlPorts.keys().next().value!
        controlPorts.get(oldest)!.dispose(); controlPorts.delete(oldest)
      }
      setControls(port)
    })
      .catch(error => { if (!abort.signal.aborted) setAnnouncement(error instanceof Error ? error.message : 'Controls unavailable.') })
    return () => abort.abort()
  }, [activeSheetId, source])
  useEffect(() => () => { for (const port of controlPorts.values()) port.dispose(); controlPorts.clear(); controlResources.current?.dispose(); controlResources.current = null }, [])
  const [popup, setPopup] = useState<{ kind: 'filter' | 'sort'; col: number; invoker: HTMLElement } | null>(null)
  const closePopup = () => setPopup(null)
  const controlJob = useRef<AbortController | null>(null)
  useEffect(() => { closePopup(); controlJob.current?.abort() }, [activeSheetId])
  useEffect(() => () => controlJob.current?.abort(), [])
  const sortCycle = async (col: number, shift: boolean) => {
    const sheetId = activeSheetId
    controlJob.current?.abort()
    const abort = new AbortController(); controlJob.current = abort
    try {
      const client = (controls ?? await controlsReady.current)?.client
      if (!client || abort.signal.aborted || current.current.activeSheetId !== sheetId) return
      const owner = client.owner(), previous = current.current.entryFor(sheetId).view.sort
      const resolved = await client.resolve(owner, 'sort', abort.signal, undefined, col)
      if (resolved.kind !== 'complete') { if (!abort.signal.aborted) setAnnouncement(resolved.kind === 'refused' ? failureText(resolved) : 'Source range is incomplete.'); return }
      const key = previous?.keys.find(key => key.col === col)
      const order = !key ? 'asc' : key.order === 'asc' ? 'desc' : null
      const keys = shift ? [...previous?.keys.filter(key => key.col !== col) ?? []] : []
      if (order) keys.push({ col, order })
      const result = await client.sort(owner, keys.length ? { range: resolved.range, keys, caseSensitive: previous?.caseSensitive ?? false } : null, abort.signal)
      if (!abort.signal.aborted) setAnnouncement('snapshot' in result ? 'View-only sort applied; formulas unchanged.' : failureText(result))
    } catch (error) { if (!abort.signal.aborted) setAnnouncement(error instanceof Error ? error.message : 'Sort unavailable.') }
  }
  const copyJob = useRef<AbortController | null>(null), copyToken = useRef(0)
  const [copyPending, setCopyPending] = useState(false)
  const cancelCopy = () => { copyToken.current++; copyJob.current?.abort(); setCopyPending(false) }
  useEffect(() => { cancelCopy() }, [identityKey, allowCopy, selection])
  useEffect(() => () => { copyToken.current++; copyJob.current?.abort() }, [])
  const copyGrid = async () => {
    if (!current.current.allowCopy || !entry) return
    cancelCopy()
    const token = ++copyToken.current, abort = new AbortController(); copyJob.current = abort
    setCopyPending(true)
    // Column visibility is compact authored geometry, not the selected cell product.
    const hiddenCols: SourceRowInterval[] = []
    let start: number | null = null
    for (let col = 0; col < entry.metrics.cols.logicalCount; col++) {
      if (entry.metrics.cols.size(col) === 0) { start ??= col }
      else if (start !== null) { hiddenCols.push({ r0: start, r1: col - 1 }); start = null }
    }
    if (start !== null) hiddenCols.push({ r0: start, r1: entry.metrics.cols.logicalCount - 1 })
    const result = await copySelection({ store: source.store(activeSheetId), owner: { ...identity }, currentOwner: () => current.current.identityOf(current.current.activeSheetId),
      sessionBudgetBytes: copySessionBudget(profile),
      selection, order: entry.view.order, visibility: entry.view.visibility, columns: visibilityFromHidden(createIntervalSet(hiddenCols), entry.metrics.cols.logicalCount),
      allowed: () => current.current.allowCopy, signal: abort.signal, onPending: () => { if (token === copyToken.current) setAnnouncement('Copy is waiting for values; cancel to stop.') },
      format: cell => { const style = remote ? remote.styleOf(cell.styleIndex) : cellAt(entry, cell)?.style; return cellDisplayText(entry.sheet, { value: cell.value, style }) },
    })
    if (token !== copyToken.current) return
    setCopyPending(false)
    setAnnouncement(result.kind === 'copied' ? 'Selection copied.' : result.kind === 'refused' ? result.message : result.kind === 'stale' ? 'Source or view changed; copy again.' : 'Copy cancelled.')
  }
  const [valueNotice, setValueNotice] = useState(() => source.valueStatus?.().message)
  useEffect(() => {
    const refresh = () => {
      const status = source.valueStatus?.()
      setValueNotice(status?.message)
      return status?.policy === 'cached' && status.formulas === undefined && !status.ended
    }
    if (!refresh()) return
    const timer = setInterval(() => { if (!refresh()) clearInterval(timer) }, 100)
    return () => clearInterval(timer)
  }, [source])
  const notices = [...new Set([...(doc.diagnostics ?? []).filter(d => d.feature !== SAVED_VALUES_DIAGNOSTIC && ['alternate-content-data', 'large-file-compatibility', 'possibly-damaged'].includes(d.feature)).map(d => d.message), ...(source.openDiagnostics ?? []).filter(d => d.code !== SAVED_VALUES_DIAGNOSTIC).map(d => d.message), ...(valueNotice ? [valueNotice] : [])])]
  // A worker-backed sheet is complete only once its import finished (LIFE-03).
  const loadState = remote?.loadState(activeSheetId)
  // Retained font diagnostics (§9.1): bounded and deduplicated by family.
  const fontNotices = fontNotice(assets && assets.sheetId === activeSheetId ? assets.lease.fontDiagnostics : [])

  // Overlay layer: selection and search highlights in viewport space.
  useLayoutEffect(() => {
    const canvas = overlayRef.current
    if (!canvas || !entry || !scroll || admittedDpr === null) return
    rendererFor() // admit retained tiles + surfaces before sizing this surface
    sizeCanvas(canvas, device.width, device.height)
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
    if (!selection.ranges.length) setSelectionIncomplete(false)
    else {
      const records = overlayRecords(entry.metrics, pane, profile)
      // Past the per-cell cap the membership is coalesced into pixel groups, never dropped.
      const coarse = records ? null : coarseSelectionRecords(selection, entry.metrics, pane, scroll.zoom * dpr)
      setSelectionIncomplete(Boolean(coarse?.incomplete))
      const geometry = records
        ? gridSelectionOverlay({ contains: at => selectionContains(selection, at) }, records, pane, scroll.zoom * dpr)
        : gridSelectionOverlay({ contains: () => true }, coarse!.records, pane, scroll.zoom * dpr)
      paintGridSelectionOverlay(ctx, gridGeometryToViewport(geometry, pane, scroll.zoom, bias), dpr)
    }
    report()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [scroll, selection, page, activeHit, viewport, dpr, identityKey, loadTick])

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
    if (key === 'escape' && copyPending) { event.preventDefault(); cancelCopy(); setAnnouncement('Copy cancelled.'); return }
    if (allowCopy && mod && key === 'c' && ((event.target as HTMLElement).hasAttribute?.('data-officeview-grid') || event.target === event.currentTarget)) { event.preventDefault(); void copyGrid(); return }
    if (!allowCopy && mod && key === 'c') { event.preventDefault(); return }
    if (!allowPrint && mod && key === 'p') { event.preventDefault(); return }
    if (allowPrint && mod && key === 'p') { event.preventDefault(); void preparePrint(); return }
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

  // ---- printing (spec §11.5): the S2.3 budgeted static pages, never the bare pane ----
  // Ctrl/Cmd+P prepares one static unit per sheet through getPaintables (which
  // refuses oversized sheets with a typed notice instead of truncating), paints
  // them into print-only canvases, prints, and releases them afterwards. A print
  // started from the browser menu shows an explicit print-only notice instead of
  // an unlabeled viewport subset.
  // Aggregate admission: the units' backing stores together must fit the bitmap
  // budget left beside the pane surfaces (tiles are released first); units that do
  // not fit are replaced by an explicit "omitted" notice, never silently dropped.
  // Every publication re-checks its request token, permission and mount (§10.1).
  const [printPages, setPrintPages] = useState<ReadonlyArray<{ name: string; page: Paintable | null; refusedImages: number }> | null>(null)
  const printLease = useRef<PaintableArray | null>(null)
  const printToken = useRef(0)
  const allowPrintRef = useRef(allowPrint)
  allowPrintRef.current = allowPrint
  const releasePrint = () => { setPrintPages(null); setPrintReserved(0); printLease.current?.dispose(); printLease.current = null }
  const preparePrint = async () => {
    if (!allowPrintRef.current) return
    // Static print needs every cell on the main thread; a worker-backed workbook never has them.
    if (remote) { setAnnouncement(LARGE_EXPORT_MESSAGE); return }
    const token = ++printToken.current
    const loader = props.printLoader ?? preparePrintUnits
    let pages: Awaited<ReturnType<typeof loader>>
    try { pages = await loader(doc, {
      ...(markKey ? { watermark } : {}), ...(fontKey ? { fallbackFonts } : {}), profile, imageCache: imageCacheFor(), assetLoader: loadAssets, allowed: () => allowPrintRef.current,
    }) } catch (error) {
      if (token === printToken.current && allowPrintRef.current) setAnnouncement(error instanceof Error ? error.message : 'Print unavailable; try again.')
      return
    }
    if (token !== printToken.current || !allowPrintRef.current) { pages.dispose(); return }
    // Admission against the CURRENT pane geometry (a resize may have landed during
    // the await): print stores + pane surfaces + the pool's minimum tiles <= budget.
    rendererRef.current?.cache.invalidate(() => true)
    const { viewport: current, dpr: scale } = latestGeometry.current
    let remaining = BITMAP_BUDGETS[profile] - paneSurfaceBytes(current, scale) - 2 * TILE_BYTES
    let reserved = 0
    if (pages.some(page => page.refusal?.length)) setAnnouncement(LARGE_EXPORT_MESSAGE)
    const units = pages.map((page, index) => {
      const bytes = page.spec.widthPx * page.spec.heightPx * 4
      const admitted = bytes <= remaining
      if (admitted) { remaining -= bytes; reserved += bytes }
      return { name: doc.sheets[index]?.name ?? `Sheet ${index + 1}`, page: admitted ? page : null, refusedImages: pages.refusedImages?.[index] ?? 0 }
    })
    printLease.current?.dispose()
    printLease.current = pages
    // Reserve before the stores exist: the pool and pane admission shrink first.
    setPrintReserved(reserved)
    setPrintPages(units)
  }
  const printRefs = useRef<Array<HTMLCanvasElement | null>>([])
  useLayoutEffect(() => {
    if (!printPages) return
    const token = printToken.current
    // One unit at a time: each admitted store is sized and painted in turn.
    printPages.forEach((unit, index) => {
      const canvas = printRefs.current[index]
      if (!unit.page || !canvas) return
      sizeCanvas(canvas, unit.page.spec.widthPx, unit.page.spec.heightPx)
      const ctx = canvas.getContext('2d')
      if (ctx) unit.page.paint(ctx)
    })
    if (token === printToken.current && allowPrintRef.current) window.print?.()
  }, [printPages])
  useEffect(() => {
    if (!allowPrint) return
    window.addEventListener('afterprint', releasePrint)
    const host = globalThis.document
    const media = host.createElement('style')
    media.setAttribute('data-officeview-sheet-print-media', '')
    media.textContent = '[data-officeview-sheet-print] { display: none; } @media print { [data-officeview-sheet-screen] { display: none !important; } [data-officeview-sheet-print] { display: block !important; } }'
    host.head.appendChild(media)
    return () => {
      // Permission withdrawn or unmounted: any in-flight preparation is now stale.
      printToken.current++
      window.removeEventListener('afterprint', releasePrint)
      media.remove()
      releasePrint()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allowPrint])

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
    cancelCopy()
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
    cancelCopy()
    session.setSelection(activeSheetId, result.selection)
    setWindow(activeSheetId, moved.window)
  }
  const onSheetChange = (sheetId: string) => {
    if (!sheetIds.includes(sheetId)) return
    cancelCopy()
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
        onCopyCapture={event => { if (!allowCopy) event.preventDefault(); else if ((event.target as HTMLElement).hasAttribute?.('data-officeview-grid')) { event.preventDefault(); void copyGrid() } }}
        onContextMenu={allowCopy ? undefined : event => event.preventDefault()}
        style={{
          background: props.background ?? '#888888', display: 'flex', flexDirection: 'column', minHeight: 0, flex: '1 1 auto',
          WebkitUserSelect: allowCopy ? undefined : 'none', userSelect: allowCopy ? undefined : 'none', ...props.style,
        }}
      >
        <div data-officeview-sheet-screen="" style={{ display: 'flex', flexDirection: 'column', minHeight: 0, flex: '1 1 auto' }}>
        <SheetNavigator documentId={docId} sheets={doc.sheets.map((sheet, index) => ({ id: sheetIds[index], name: sheet.name }))}
          activeSheetId={activeSheetId} onSheetChange={onSheetChange} onNavigate={onNavigate} onNotice={setAnnouncement} disabled={!entry} />
        <div data-officeview-ui="sheet-actions"><button type="button" disabled={!allowCopy || !selection.ranges.length || copyPending} onClick={() => void copyGrid()}>Copy selection</button>{copyPending && <button type="button" onClick={cancelCopy}>Cancel copy</button>}<button type="button" disabled={!allowPrint} onClick={() => void preparePrint()}>Print</button></div>
        <div ref={paneRef} data-testid="officeview-sheet-pane" style={{ position: 'relative', flex: '1 1 auto', minHeight: 0, overflow: 'hidden', background: '#ffffff' }}>
          {entry && scroll && (
            <GridSemantics owner={identity} projection={entry.metrics} view={entry.view} selection={selection} window={scroll} viewport={viewport}
              body={entry.body} onChange={onGridChange} status={{ records: { ...snapshot(activeSheetId).records, completeness: loadState?.completeness ?? 'complete' }, countMode: 'body', loadReport: { status: 'ready', notices: [...notices, ...fontNotices, ...(loadState?.failure ? [loadFailureNotice(loadState.failure, loadState.code)] : [])] } }}
              onSortCycle={(col, shift) => void sortCycle(col, shift)} onFilter={(col, invoker) => { setPopup({ kind: 'filter', col, invoker }) }} onSortDialog={(col, invoker) => { setPopup({ kind: 'sort', col, invoker }) }} onNotice={setAnnouncement}
              popup={popup && controls ? { invoker: popup.invoker, close: closePopup, label: popup.kind === 'filter' ? 'Column filter' : 'Sort worksheet', content: popup.kind === 'filter' ? <FilterDropdown owner={identity} client={controls.client} applied={entry.view.filter} appliedSort={entry.view.sort} col={popup.col} invoker={popup.invoker} embedded onClose={closePopup} onApplied={() => rerender()} onNotice={setAnnouncement} /> : <SortDialog owner={identity} client={controls.client} applied={entry.view.sort} col={popup.col} invoker={popup.invoker} embedded onClose={closePopup} onApplied={() => rerender()} onNotice={setAnnouncement} /> } : undefined} announcement={[resourceNotice, assets?.sheetId !== activeSheetId ? 'Loading sheet…' : '', assets && assets.sheetId === activeSheetId && assets.lease.refusedImages > 0 ? `${assets.lease.refusedImages} picture${assets.lease.refusedImages === 1 ? ' is' : 's are'} not shown (image memory budget).` : '', selectionIncomplete ? 'Selection highlight is incomplete at this zoom; zoom in for detail.' : '', announcement].filter(Boolean).join(' ')} inspection={inspection} profile={profile} label={`Worksheet ${entry.sheet.name}`}
              cellPresentation={at => {
                const cell = cellAt(entry, at)
                if (!cell && remote && !remote.store(activeSheetId).isKnown(at)) return { text: 'Loading…' }
                return { text: cell ? cellDisplayText(entry.sheet, cell) : '' }
              }}>
              <canvas ref={baseRef} data-testid="officeview-sheet-base" style={canvasStyle} />
              <canvas ref={overlayRef} data-testid="officeview-sheet-overlay" style={canvasStyle} />
              {/* tabIndex -1: Chromium would make this scroller a second Tab stop inside the grid (spec §9.2: one Tab entry). */}
              <div ref={scrollerRef} data-testid="officeview-sheet-scroller" tabIndex={-1} onScroll={onScroll}
                onPointerDown={onPointerDown} onPointerMove={onPointerMove} onPointerUp={onPointerEnd} onPointerCancel={onPointerEnd}
                style={{ position: 'absolute', inset: 0, overflow: 'auto', touchAction: 'pan-x pan-y', overscrollBehavior: 'contain' }}>
                <div style={{ width: scroll.physicalExtent.width, height: scroll.physicalExtent.height }} />
              </div>
            </GridSemantics>
          )}
        </div>
        </div>
        {allowPrint && (
          <div data-officeview-sheet-print="" aria-hidden="true">
            {!printPages && <p data-testid="officeview-sheet-print-notice">To print this spreadsheet, press Ctrl/Cmd+P inside the viewer.</p>}
            <div data-testid="officeview-sheet-print-pages">
              {printPages?.map((unit, index) => unit.page
                ? <div key={index} style={{ breakAfter: 'page' }}>
                    {unit.page.refusal?.length ? <p>{LARGE_EXPORT_MESSAGE}</p> : null}
                    {unit.refusedImages > 0 && <p>Sheet “{unit.name}”: {unit.refusedImages} picture{unit.refusedImages === 1 ? ' is' : 's are'} omitted from print (image memory budget).</p>}
                    <canvas ref={el => { printRefs.current[index] = el }} style={{ display: 'block', maxWidth: '100%' }} />
                  </div>
                : <p key={index}>Sheet “{unit.name}” is omitted from print: it exceeds the print memory budget.</p>)}
            </div>
          </div>
        )}
      </div>
      {showSearch && (
        <SearchBar inputRef={inputRef} query={query} onQuery={text => { searchGen.current++; setQuery(text) }} count={matchCount} active={activeHit < 0 ? -1 : pageOffset + activeHit}
          countLabel={countLabel} note={searchNote}
          indexing={searching} onPrev={() => void stepHit(-1)} onNext={() => void stepHit(1)} />
      )}
      {showZoomControls && (
        <ZoomControls zoom={zoom} onIn={() => zoomBy(1)} onOut={() => zoomBy(-1)} onReset={resetZoom} atFit={zoom <= SHEET_MIN_ZOOM} />
      )}
    </div>
  )
}

/** Source row intervals → display slot intervals under `order`. Outside permuted
 * bands slots equal source rows; inside a band each source row maps to its own
 * slot (a contiguous source run is not contiguous in display order). */
function displayIntervals(intervals: readonly SourceRowInterval[], order: RowOrder): SourceRowInterval[] {
  const bands = [...order.bands].sort((a, b) => a.r0 - b.r0)
  const slots: Array<[number, number]> = []
  for (const { r0, r1 } of intervals) {
    let at = r0
    for (const band of bands) {
      if (band.r1 < at || band.r0 > r1) continue
      if (band.r0 > at) slots.push([at, band.r0 - 1])
      const b1 = Math.min(r1, band.r1)
      for (let row = Math.max(at, band.r0); row <= b1; row++) { const slot = order.sourceToDisplay(row); slots.push([slot, slot]) }
      at = b1 + 1
      if (at > r1) break
    }
    if (at <= r1) slots.push([at, r1])
  }
  slots.sort((a, b) => a[0] - b[0])
  const merged: SourceRowInterval[] = []
  for (const [lo, hi] of slots) {
    const last = merged[merged.length - 1]
    if (last && lo <= last.r1 + 1) last.r1 = Math.max(last.r1, hi)
    else merged.push({ r0: lo, r1: hi })
  }
  return merged
}

/** Resize a backing store without an intermediate W_new x H_old surface: release
 * it first (0 bytes), then assign the admitted final dimensions (§12.4). */
function sizeCanvas(canvas: HTMLCanvasElement, width: number, height: number): void {
  if (canvas.width === width && canvas.height === height) return
  canvas.height = 0
  canvas.width = width
  canvas.height = height
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
