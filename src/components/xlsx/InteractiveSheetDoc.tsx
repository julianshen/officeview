import { useCanvasColors } from '../canvasColors'
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
import { tilePlanSource, WorkerViewportPlans } from './workerViewport'
import { useCallback, useEffect, useId, useLayoutEffect, useMemo, useRef, useState, type CSSProperties, type ReactElement } from 'react'
import { Copy, Printer, Minus, Plus, RotateCcw } from 'lucide-react'
import { Icon } from '../Icon'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../../xlsx/types'
import type {
  AuthoritativeCell, CacheIdentity, FilterSpec, GridSelection, LogicalCellBox, ProfileId, RowOrder, RowVisibility, ScrollWindow, SheetView, SortSpec, SourceRowInterval, SourceStore, WorkbookOwner,
} from '../../xlsx/view/contracts'
import { SHEET_MIN_ZOOM } from '../../xlsx/view/contracts'
import type { NavigatorRequest, SourceCell } from '../../xlsx/view/coordinates'
import { columnToLetters, validateCoordinate } from '../../xlsx/view/coordinates'
import { createDocumentOverflowCache } from '../../xlsx/view/documentOverflow'
import { bindOverflowCache, OverflowResourceError, type OverflowCache } from '../../xlsx/view/overflowCache'
import { MainMemoryLedger, axisBytes, estimatePayloadBytes, projectionBytes, type MainPool } from '../../xlsx/view/mainLedger'
import { buildSheetGeometry, cellTextOverflow, MergeQueryCache, sheetBaseHidden, type SheetGeometry } from '../../xlsx/view/viewport'
import type { ProjectedSheetMetrics } from '../../xlsx/view/projection'
import { createIdentityOrder } from '../../xlsx/view/order'
import { CompactRowVisibility, createIntervalSet, createRowVisibility, visibilityFromHidden } from '../../xlsx/view/visibility'
import {
  createScrollWindow, navigateTo as scrollNavigateTo, refreshScrollWindow, scrollIntoView, scrollPhysicalTo, setScrollOffset, setScrollZoom, visibleLogicalViewport, type ViewportSize,
} from '../../xlsx/view/scrollWindow'
import { keyboardSheetZoom, pinchSheetZoom, wheelSheetZoom } from '../../xlsx/view/sheetZoom'
import { SheetNavigationSession, emptyGridSelection, navigateSelection, recordCounts, boundedSelectionText } from '../../xlsx/view/selection'
import { FIND_PAGE_HITS, FindController, hitGeometry, navigateToHit, type FindHit, type FindInput, type FindOptions, type FindPublication } from '../../xlsx/view/find'
import { LayoutCache, type CellLayoutHost, type LayoutAdmission } from '../../xlsx/view/textLayout'
import { SharedMeasureContext } from './measureContext'
import { cellDisplayText } from '../../xlsx/render'
import { SemanticPresentationCache } from './formatPresentation'
import { mainFormatBudget } from './mainFormatBudget'
import type { FormatBudget } from '../../xlsx/view/formatWorkspace'
import { createObjectWorkbookSource } from '../../worker/viewerObjectSource'
import { detectProfile } from '../../xlsx/storage/importSession'
import { BITMAP_BUDGETS, MAIN_THREAD_BUDGETS, type TileCacheStats } from '../../xlsx/view/tileCache'
import { gridGeometryToViewport, gridSelectionOverlay, paintGridSelectionOverlay, type HighlightRect } from '../../core/overlay'
import { normalizeWatermark, type WatermarkOptions } from '../../core/watermark'
import type { FallbackFontsOptions } from '../../core/fonts/fallback'
import type { FontDiagnostic } from '../../core/fonts/types'
import { SheetImageCache, imageBudgetFor, mainChartParent, preparePrintUnits, prepareSheetAssets, type PrintUnitsOptions, type SheetAssetsLease } from './sheetAssets'
import type { Paintable, PaintableArray } from '../../render/paint'
import { GridSemantics, STATUS_BAR_HEIGHT, type GridTransition } from './GridSemantics'
import { GRID_HEADER_SIZE, columnHeaderHeight, rowGutterWidth } from './GridHeaders'
import { useControlMedia, useLive } from './controlMedia'
import { useDevicePixelRatio } from './devicePixelRatio'
import { domClassBytes, domClassCounts } from './domCost'
import { FilterDropdown } from './FilterDropdown'
import { SortDialog } from './SortDialog'
import { createViewerControlResources, viewerControls, type ViewerControlResources, type ViewerControls } from './viewerControlsClient'
import type { ControlSnapshot } from './controlsClient'
import { failureText } from './controlUi'
import { copySelection, copySessionBudget } from '../../xlsx/view/copy'
import { LARGE_EXPORT_MESSAGE } from '../../render/paint'
import { SAVED_VALUES_DIAGNOSTIC, SAVED_VALUES_STATUS } from '../../xlsx/view/values'
import type { WorkbookSource } from '../../worker/viewerSession'
import { coarseSelectionRecords, overlayRecords, paintFindOverlay, selectionContains } from './selectionOverlay'
import { SheetNavigator } from './SheetNavigator'
import type { HiddenCellInspection } from './ViewerStatus'
import { SheetTileRenderer, TILE_BYTES, admitDeviceScale, paneSurfaceBytes, type SheetFrame } from './sheetTiles'
import { SearchBar, ZoomControls } from '../viewerControls'
import { MotionRoot, Tick } from '../motion/presets'
import { themeRootProps, type ThemeMode } from '../theme'
import { warnIfStylesMissing } from '../styleWarning'
import { isRemoteWorkbook } from '../../worker/viewerRemoteSource'
import { documentWithDrawings, drawingNotices, sheetWithDrawings } from '../../worker/viewerRemoteDrawings'
import { pagedSheetGeometry } from './pagedGeometry'
import type { LedgerHold } from '../../xlsx/view/mainLedger'
import { DisplayHolds } from './displayHolds'
import { SheetChartDescriptions, chartNotices, sheetCharts } from './sheetCharts'
import type { XlsxFrozenPane } from '../../xlsx/frozenPane'
import { initialOrigin, paneLogicalBoxes, paneRegions } from '../../xlsx/view/frozenPanes'
import { SheetSizes, NO_SESSION_SIZES, RESIZE_LIMITS, SIZED_DRAWING_BYTES, axisUpperBytes, type SessionSizes } from '../../xlsx/view/axisSizes'
import { AutoFitRefusedError, autoFitSize, createFitMeasure, resizeRegion, type ResizeRequest } from './sheetResize'
import { tileLogicalBox } from '../../xlsx/view/tileCache'

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
  theme?: ThemeMode
  /** @internal Committed sort/filter views per sheet id (S9.1 seam). */
  views?: Readonly<Record<string, SheetViewInput>>
  /** @internal Instrumentation for tests and the browser harness. */
  onFrame?: (stats: SheetFrameStats) => void
  /** @internal S10.2 harness: read-only access to the document's main-thread ledger. */
  onMainLedger?: (ledger: MainMemoryLedger) => void
  /** @internal Resource profile (desktop D / mobile M); default: the admitted source's, else the detected device profile. */
  profile?: ProfileId
  /** @internal Asset lease factory (tests inject delayed leases). */
  assetLoader?: typeof prepareSheetAssets
  /** @internal Static print unit factory (tests inject delayed units). */
  printLoader?: (doc: XlsxDocument, options: PrintUnitsOptions) => Promise<PaintableArray & { refusedImages?: number[] }>
}

// Document identity alone owns zoom reset (§10.1): a stable id per document object.
const documentIds = new WeakMap<object, string>()
let nextDocumentId = 0
/** Same requested freeze (counts only; extents follow the pane size). */
function sameFreeze(window: ScrollWindow['frozen'], authored: XlsxFrozenPane | undefined): boolean {
  return (window?.rows ?? 0) === (authored?.rows ?? 0) && (window?.cols ?? 0) === (authored?.cols ?? 0)
}
/**
 * The active sheet's semantic presentations (AC-LIFE-02). Ownership:
 * - Committed and replaced caches are released explicitly: each commit
 *   disposes the superseded cache and any cache a render made without
 *   committing while this viewer stayed mounted; unmount disposes the committed
 *   cache and those still tracked.
 * - Caches from renders that never reach an owner's cleanup are released by
 *   the GC finalization backstop only: an initial render that suspends and
 *   never commits (React runs no cleanup for it), or a cache a commit already
 *   disposed and dropped that a still-pending render then refills.
 * - The committed cache is published to callbacks in the commit phase only: a
 *   speculative render never changes what the visible view's callbacks read.
 *   During its own render a view reads the cache it rendered with.
 * dispose() settles each presentation's lifetime charge; a later finalizer for
 * a settled target is a no-op in the ledger, so a charge ends exactly once.
 * No closure here captures a cache directly: a stale callback (or a retained
 * closure context) reaches at most the committed one through the owner.
 */
export function useSemanticPresentations(sheet: XlsxSheet | undefined, valueGen: number, budget: FormatBudget): (cell: XlsxCell) => { text: string; state?: 'error' } {
  const cache = useMemo(() => sheet ? new SemanticPresentationCache(sheet, budget) : null, [sheet, valueGen, budget])
  const owner = useRef<{ committed: SemanticPresentationCache | null; rendered: Set<SemanticPresentationCache> } | null>(null)
  const own = owner.current ??= { committed: null, rendered: new Set() }
  if (cache) own.rendered.add(cache)
  const slot = { cache, pending: true }
  useLayoutEffect(() => {
    // StrictMode re-runs this effect for the same commit: publish once.
    if (!slot.pending) return
    const next = slot.cache
    slot.pending = false
    slot.cache = null
    // The superseded cache and every rendered-but-uncommitted one end here.
    if (own.committed) own.rendered.add(own.committed)
    own.committed = next
    for (const rendered of own.rendered) if (rendered !== next) rendered.dispose()
    own.rendered.clear()
  })
  // Disposal is not terminal: a StrictMode remount refills the committed cache.
  useEffect(() => () => {
    own.committed?.dispose()
    for (const rendered of own.rendered) rendered.dispose()
    own.rendered.clear()
  }, [])
  return cell => (slot.pending ? slot.cache : own.committed)!.get(cell)
}
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
/** Worker-backed (paged) sheets have no main-thread cells or static drawing paint to print. */
const PAGED_PRINT_MESSAGE = 'This spreadsheet cannot be printed from the large-file view: neither its cells nor its drawings are printed. Copy a selected range instead.'
const MAX_FONT_NOTICE_FAMILIES = 3
/** Print rules scoped to one viewer instance: its screen and floating chrome
 * never print; its print container does. */
function sheetPrintCss(id: string): string {
  const own = `[data-officeview-sheet-instance="${id}"]`
  // In print the viewer is a plain white block: a flex container would ignore
  // the per-sheet page breaks, and the screen background must not print.
  return `${own} [data-officeview-sheet-print]{display:none}@media print{${own} [data-officeview-sheet-screen],${own} [data-officeview-chrome]{display:none!important}${own}{display:block!important}${own} [data-officeview-root]{display:block!important;background:#fff!important}${own} [data-officeview-sheet-print]{display:block!important}}`
}
/** One document-wide print owner: the viewer running its own print leaves every
 * other spreadsheet viewer out of the printout. The latest starter takes over
 * (the previous owner's exclusion goes with it), so a print whose afterprint
 * never arrived cannot combine with a later one to hide every viewer. Released
 * on the owner's afterprint, permission withdrawal or unmount. */
let printOwner: { id: string; style: HTMLStyleElement } | null = null
function claimPrint(id: string): void {
  printOwner?.style.remove()
  const style = globalThis.document.createElement('style')
  style.setAttribute('data-officeview-sheet-print-owner', '')
  style.textContent = `@media print{[data-officeview-sheet-instance]:not([data-officeview-sheet-instance="${id}"]){display:none!important}}`
  globalThis.document.head.appendChild(style)
  printOwner = { id, style }
}
function yieldPrint(id: string): void {
  if (printOwner?.id !== id) return
  printOwner.style.remove()
  printOwner = null
}
/** Print units render at this device scale when the print bitmap budget admits it. */
const PRINT_SCALE = 2
/** Per-axis canvas ceiling (§11.5 raster admission). */
const PRINT_MAX_AXIS = 16384
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
  /** View-only session size generation the geometry holds (spec §5.6). */
  resizeGen: number
  /** The view input this entry shows (rebuilt with it when only sizes change). */
  input: SheetViewInput
  view: SheetView
  geometry: SheetGeometry
  metrics: ProjectedSheetMetrics
  body: SourceRowInterval | null
  layoutCache: LayoutCache
  merges: MergeQueryCache
  /** Authored frozen panes (resident sheet or worker layout). */
  frozen?: XlsxFrozenPane
}
/** Fallback pane when the host has not laid the viewer out (jsdom, hidden tab). */
const FALLBACK_PANE = { width: 640, height: 400 }
/** Prefetch slice: leaves most of a 16.7 ms frame to input and presentation. */
const PREFETCH_SLICE_MS = 6
// Ledger owners must stay unique across remounts with the same remote workbook.
let copySerial = 0

function SheetViewer(props: InteractiveSheetDocProps & { docId: string }): ReactElement {
  const { doc, docId, views, watermark, fallbackFonts, allowCopy = true, allowPrint = true, showSearch = true, showZoomControls = true } = props
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
  // README / spec §12.4: admission selects D (desktop) or M (phones and tablets,
  // iPadOS included); presentation (bitmap pool, DPR admission, tiles, images)
  // follows the same device profile, not a fixed D. A worker-backed source
  // carries the profile it was admitted with; a resident one is detected here.
  const profile = props.profile ?? remote?.profile ?? detectProfile()
  // S10.2 §12.4: one main-thread ledger per document. A worker-backed workbook
  // owns it (row cache, view buffers, replies); a resident one gets its own.
  const mainLedger = useMemo(() => remote?.mainLedger ?? new MainMemoryLedger(), [remote, source])
  // One budget profile across every main-thread ledger consumer: if a worker-backed
  // remote workbook already established a profile (D 64 MiB / M 32 MiB), sibling or
  // child viewers sharing its mainLedger must respect the remote's budget cap for
  // controls, overflow, and copy admission, while retaining their own props for tile
  // rendering, DPR, and presentation.
  const budgetProfile = remote?.profile ?? profile
  const semanticFormatBudget = useMemo(() => mainFormatBudget(mainLedger, MAIN_THREAD_BUDGETS[budgetProfile]), [mainLedger, budgetProfile])
  const mergeFormatBudget = useMemo(() => mainFormatBudget(mainLedger, MAIN_THREAD_BUDGETS[budgetProfile], 'merges'), [mainLedger, budgetProfile])
  const overflowContext = useMemo(() => ({ cache: undefined as OverflowCache | undefined }), [mainLedger, budgetProfile])
  const overflowFor = (): OverflowCache => {
    if (!overflowContext.cache || overflowContext.cache.closed) overflowContext.cache = remote ? remote.overflowFor(budgetProfile) : createDocumentOverflowCache(mainLedger, budgetProfile)
    return overflowContext.cache
  }
  useEffect(() => () => { overflowContext.cache?.close(); overflowContext.cache = undefined }, [overflowContext])
  // View buffers already charged in this ledger (a rebuilt entry may reuse them).
  const viewCharged = useMemo(() => new WeakSet<object>(), [mainLedger])
  // §§10.1, 12.4: each sheet entry holds the charged buffers it shows (view
  // order/mask, row projection) until a committed display no longer shows it;
  // a worker view also stays held by the workbook, sibling viewers and copies.
  const holds = useMemo(() => new DisplayHolds<SheetEntry>(held => [held.view.order, held.view.visibility, held.geometry, held.geometry.sizedDrawings].flatMap(target => target ? mainLedger.hold(target) ?? [] : [])), [mainLedger])
  // Display serials: an entry replaced while rendering display N is released
  // once N (or a later display) commits; one replaced outside a render (an
  // event or publication) once the next display commits.
  const serials = useRef({ render: 0, rendering: false }).current
  const displaySerial = ++serials.render
  serials.rendering = true
  useLayoutEffect(() => { serials.rendering = false })
  useEffect(() => { holds.committed(displaySerial) })
  useEffect(() => { holds.restore(); return () => holds.dispose() }, [holds])
  // §12.4 retained views: every external `views` input in the committed props is
  // charged and held while any sheet's prop still references it, whether or not
  // it is displayed (a newer local view may win). Reconciled after each commit
  // against the committed props, never from a speculative render; a new holder
  // is taken before an old one is released, so a fresh wrapper of the same
  // buffers never drops them.
  const inputOwners = useMemo(() => new Map<SheetViewInput, LedgerHold[]>(), [mainLedger])
  useEffect(() => {
    const provided = new Set(remote ? [] : Object.values(views ?? {}))
    for (const input of provided) if (!inputOwners.has(input)) {
      chargeViewBuffers(input.order ?? null, input.visibility ?? null)
      inputOwners.set(input, [input.order, input.visibility].flatMap(target => target ? mainLedger.hold(target) ?? [] : []))
    }
    for (const [input, held] of inputOwners) if (!provided.has(input)) { for (const hold of held) hold.release(); inputOwners.delete(input) }
  })
  useEffect(() => () => { for (const held of inputOwners.values()) for (const hold of held) hold.release(); inputOwners.clear() }, [inputOwners])
  // The viewer's own pools end with the viewer; lifetime charges (geometry,
  // views, replies, the resident document) end when nothing reaches them.
  // Submitted clipboard jobs retain their own charge until they settle.
  useEffect(() => () => {
    for (const pool of ['dom', 'canvasText', 'layout'] as const) mainLedger.releaseAll(pool)
    // Adopted merge pages may still have callers after unmount. Only the
    // legacy aggregate ends here; each adopted page ends with its aliases.
    mainLedger.release('merges', 'caches')
  }, [mainLedger])
  // A resident document's parsed model lives on the main thread while reachable.
  useMemo(() => { if (!remote) mainLedger.chargeWhileReachable('document', [doc], estimatePayloadBytes(doc)) }, [mainLedger, remote, doc])
  const onMainLedger = props.onMainLedger
  useEffect(() => { onMainLedger?.(mainLedger) }, [mainLedger, onMainLedger])
  const [loadTick, setLoadTick] = useState(0)
  useEffect(() => remote?.subscribe(() => setLoadTick(tick => tick + 1)), [remote])
  // Tiles painted under an older worker layout/merge generation are stale (§10.1).
  const structureGens = useRef(new Map<string, number>()).current
  // View-only session sizes per sheet (spec §5.6): they outlive scroll, zoom,
  // sort/filter commits and sheet switches, and end with the viewer.
  const sheetSizes = useRef(new Map<string, SheetSizes>()).current
  const sizesOf = (sheetId: string): SheetSizes => {
    let sizes = sheetSizes.get(sheetId)
    if (!sizes) sheetSizes.set(sheetId, sizes = new SheetSizes())
    return sizes
  }

  // ---- fonts/images: one lease per (fallback fonts, active sheet) ----
  // A font publication bumps fontGen (re-keys tiles, keeps zoom). Leases swap
  // atomically: the current one stays alive, and keeps painting, until its
  // admitted replacement has committed; only then is it disposed (§10.1).
  const fontKey = fallbackFonts ? JSON.stringify(fallbackFonts) : ''
  const [assets, setAssets] = useState<{ lease: SheetAssetsLease; fontGen: number; fontKey: string; sheetId: string } | null>(null)
  const fontGenRef = useRef<{ key: string; gen: number }>({ key: fontKey, gen: 0 })
  const loadAssets = props.assetLoader ?? prepareSheetAssets
  // Transferred chart models are charged to this viewer's live main-thread budget (spec 6.4), like its drawings.
  const chartParent = useMemo(() => mainChartParent(mainLedger, MAIN_THREAD_BUDGETS[budgetProfile]), [mainLedger, budgetProfile])
  // One image budget shared by every lease of this viewer (overlapping swaps included).
  const imageCache = useRef<SheetImageCache | null>(null)
  const imageCacheFor = () => {
    if (!imageCache.current || imageCache.current.limitBytes !== imageBudgetFor(profile)) imageCache.current = new SheetImageCache(imageBudgetFor(profile))
    return imageCache.current
  }
  // A sheet switch stops presenting the old sheet at once (no stale raster under the
  // new sheet's headers), so its lease is released before the new sheet decodes.
  useEffect(() => { setAssets(current => current && current.sheetId !== activeSheetId ? null : current) }, [activeSheetId])
  // A worker-backed sheet's drawing content arrives after its structure (S13.2);
  // its pictures are leased like resident ones once it does.
  const activeDrawings = remote && activeSheetId ? remote.drawings(activeSheetId) : null
  useEffect(() => {
    if (!activeSheetId) return
    let cancelled = false
    const sheetIndex = sheetIds.indexOf(activeSheetId)
    void loadAssets(activeDrawings ? documentWithDrawings(doc, sheetIndex, activeDrawings) : doc, { ...(fontKey ? { fallbackFonts } : {}), sheetIndex, profile, imageCache: imageCacheFor(), chartOwner: doc, chartParent }).then(next => {
      if (cancelled) { next.dispose(); return }
      if (fontGenRef.current.key !== fontKey || fontGenRef.current.gen === 0) fontGenRef.current = { key: fontKey, gen: fontGenRef.current.gen + 1 }
      setAssets({ lease: next, fontGen: fontGenRef.current.gen, fontKey, sheetId: activeSheetId })
    })
    return () => { cancelled = true }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [doc, fontKey, activeSheetId, profile, activeDrawings?.gen, chartParent])
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
  const chargeViewBuffers = (order: SheetView['order'] | null, visibility: SheetView['visibility'] | null) => {
    const targets: object[] = []
    let bytes = 0
    const orderBytes = (order as { diagnostics?: { typedArrayBytes?: number } } | null)?.diagnostics?.typedArrayBytes ?? 0
    if (order && orderBytes > 0 && !viewCharged.has(order) && !mainLedger.hasLifetime(order)) { targets.push(order, ...order.bands.flatMap(band => [band.sourceAtOffset, band.sourceAtOffset.buffer, band.displayOffsetOfSource, band.displayOffsetOfSource.buffer])); bytes += orderBytes; viewCharged.add(order) }
    if (visibility instanceof CompactRowVisibility && !viewCharged.has(visibility) && !mainLedger.hasLifetime(visibility)) { targets.push(visibility, visibility.words, visibility.words.buffer, visibility.summaries, visibility.summaries.buffer); bytes += visibility.words.byteLength + visibility.summaries.byteLength; viewCharged.add(visibility) }
    if (targets.length) mainLedger.chargeWhileHeld('viewBuffers', targets, bytes)
  }
  /** Bytes of order/visibility buffers this view would newly charge. */
  const viewBufferBytes = (order: SheetView['order'] | null, visibility: SheetView['visibility'] | null): number => {
    let bytes = 0
    const orderBytes = (order as { diagnostics?: { typedArrayBytes?: number } } | null)?.diagnostics?.typedArrayBytes ?? 0
    if (order && orderBytes > 0 && !viewCharged.has(order) && !mainLedger.hasLifetime(order)) bytes += orderBytes
    if (visibility instanceof CompactRowVisibility && !viewCharged.has(visibility) && !mainLedger.hasLifetime(visibility)) bytes += visibility.words.byteLength + visibility.summaries.byteLength
    return bytes
  }
  /** A sheet entry built off to the side: nothing is charged, held or published. */
  interface Candidate {
    sheetId: string; sheet: XlsxSheet; input: SheetViewInput; structureGen: number; previous: SheetEntry | undefined
    order: SheetView['order']; baseHidden: SheetView['baseHidden']; visibility: SheetView['visibility']
    built: SheetGeometry & { readonly chargeBytes?: number }
    /** Exactly what publishing charges (the entry charge functions). */
    bytes: number
  }
  const sameInput = (a: SheetViewInput, b: SheetViewInput) => a === b || (a.revision === b.revision && a.order === b.order && a.visibility === b.visibility && a.filter === b.filter && a.sort === b.sort)
  /** The current entry, or a candidate for a new one. `sizes` forces a
   * candidate with those session sizes (a resize or reset being admitted). */
  const prepareEntry = (sheetId: string, override?: SessionSizes): { entry: SheetEntry } | { candidate: Candidate } => {
    const index = sheetIds.indexOf(sheetId)
    const sheet = remote ? sheetWithDrawings(doc.sheets[index], remote.drawings(sheetId)) : doc.sheets[index]
    let input = viewInput(sheetId)
    const previous = entries.get(sheetId)
    const structureGen = remote ? remote.merges(sheetId).gen : 0
    const sized = sheetSizes.get(sheetId), resizeGen = sized?.gen ?? 0
    if (!override && previous && previous.revision === input.revision && previous.structureGen === structureGen && previous.resizeGen === resizeGen) return { entry: previous }
    if (previous && input.revision < previous.revision) {
      // Stale commit: never publish an older view. Session sizes or structure
      // that changed since are still applied, to the view already shown.
      if (!override && previous.resizeGen === resizeGen && previous.structureGen === structureGen) return { entry: previous }
      input = previous.input
    }
    // A rebuild for session sizes alone keeps the view buffers it already shows
    // (charged and held). A structure change (a worker layout publication)
    // refreshes the authored-hidden set and the mask derived from it.
    const sameView = Boolean(previous && previous.revision === input.revision && (remote !== null || sameInput(previous.input, input)))
    const sameStructure = sameView && previous!.structureGen === structureGen
    const order = sameView ? previous!.view.order : input.order ?? createIdentityOrder()
    const baseHidden = sameStructure ? previous!.view.baseHidden : remote ? remote.geometry(sheetId).baseHidden : sheetBaseHidden(sheet)
    const authored = () => visibilityFromHidden(createIntervalSet(displayIntervals(baseHidden.intervals, order)))
    // An adopted worker (filter) mask predates a newer authored-hidden set: the
    // shown mask is that mask AND NOT authored-hidden (it never revives a
    // filtered row; source identities are unchanged).
    const visibility = sameStructure ? previous!.view.visibility
      : input.visibility ? (remote && sameView ? intersectVisibility(input.visibility, authored()) : input.visibility)
        : authored()
    const sizes = override ?? (sized && !sized.empty ? sized.snapshot() : NO_SESSION_SIZES)
    const built = remote ? pagedSheetGeometry(remote, sheetId, order, visibility, overflowFor(), sizes, true) : buildSheetGeometry(sheet, { order, ...(input.visibility ? { visibility } : {}), sizes })
    if (built.kind !== 'ready') throw new Error(built.diagnostics.map(d => d.message).join('; ') || 'This sheet cannot be laid out')
    const geometryBytes = remote ? (built as unknown as { chargeBytes: number }).chargeBytes : projectionBytes(built.metrics) + axisBytes(built.metrics.cols) + (built.cellIndexBytes ?? 0) + built.merges.estimatedBytes
    const bytes = geometryBytes + (built.sizedDrawings?.length ?? 0) * SIZED_DRAWING_BYTES + viewBufferBytes(remote ? null : order, remote && visibility === input.visibility ? null : visibility)
    return { candidate: { sheetId, sheet, input, structureGen, previous, order, baseHidden, visibility, built, bytes } }
  }
  /** Publishes a candidate: charges exactly its measured owners, holds it and
   * retires the previous entry on the next display commit. */
  const publishEntry = (candidate: Candidate, resizeGen: number): SheetEntry => {
    const { sheetId, sheet, input, structureGen, previous, order, baseHidden, visibility, built } = candidate
    bindOverflowCache(built, overflowFor())
    // A resident geometry is charged for its lifetime: the row projection (it
    // keeps its own copied segments; the build's row axis is dropped), the
    // column axis, the per-row cell index and merges. A paged one: its
    // projection, merge index and rebuilt axes (chargeBytes).
    if (remote) remote.mainLedger.chargeWhileHeld('geometry', [built], built.chargeBytes ?? 0)
    else mainLedger.chargeWhileHeld('geometry', [built], projectionBytes(built.metrics) + axisBytes(built.metrics.cols) + (built.cellIndexBytes ?? 0) + built.merges.estimatedBytes)
    // Drawings repositioned for session sizes (§5.6) are this geometry's own copies.
    if (built.sizedDrawings) mainLedger.chargeWhileReachable('drawings', [built.sizedDrawings], built.sizedDrawings.length * SIZED_DRAWING_BYTES)
    // Order and visibility buffers built on the main thread are charged once each,
    // while a view holds them: a resident sort/filter's, and the base visibility
    // words built here for any sheet (1,048,576 rows: ~132 KiB). A worker-backed
    // view's adopted buffers are charged by the workbook (adopt).
    chargeViewBuffers(remote ? null : order, remote && visibility === input.visibility ? null : visibility)
    const selection = previous?.view.selection ?? emptyGridSelection()
    const view: SheetView = { sheetId, revision: input.revision, order, baseHidden, visibility, sort: input.sort ?? null, filter: input.filter ?? null, selection }
    const used = remote?.layout(sheetId).used
    const entry: SheetEntry = {
      sheet, revision: input.revision, structureGen, resizeGen, input, view, geometry: built, metrics: built.metrics, body: input.filter || input.sort ? { r0: (input.filter?.range ?? input.sort!.range).r0 + (input.filter?.range ?? input.sort!.range).headerRows, r1: (input.filter?.range ?? input.sort!.range).r1 - (input.filter?.range ?? input.sort!.range).totalsRows } : remote ? used ? { r0: used.r0, r1: used.r1 } : null : usedBody(sheet),
      // Layout records are keyed by complete identity; a new view gets its own merge cache.
      // A worker layout/merge publication changes record boxes: fresh layouts then.
      layoutCache: previous && previous.structureGen === structureGen ? previous.layoutCache : new LayoutCache(), merges: new MergeQueryCache(64, mergeFormatBudget),
      ...((remote ? remote.layout(sheetId).frozen : sheet.frozen) ? { frozen: remote ? remote.layout(sheetId).frozen : sheet.frozen } : {}),
    }
    entries.set(sheetId, entry)
    holds.own(entry)
    if (previous) holds.retire(previous, serials.rendering ? serials.render : serials.render + 1)
    if (!previous) session.register(sheetId, { view, projection: entry.metrics, body: entry.body })
    else session.updateProjection(sheetId, entry.metrics, entry.body, view)
    return entry
  }
  const entryFor = (sheetId: string): SheetEntry => {
    const prepared = prepareEntry(sheetId)
    if ('entry' in prepared) { bindOverflowCache(prepared.entry.geometry, overflowFor()); return prepared.entry }
    return publishEntry(prepared.candidate, sheetSizes.get(sheetId)?.gen ?? 0)
  }
  // The committed revision: a stale (lower) prop revision never regresses the
  // owner that keys tiles, find and hit geometry; a newer one is built on use.
  const revisionOf = (sheetId: string): number => Math.max(entries.get(sheetId)?.revision ?? 0, viewInput(sheetId).revision)
  const identityOf = (sheetId: string): CacheIdentity => ({ docId, sheetId, valueGen: source.store(sheetId).stamp.valueGen, fontGen, viewGen: revisionOf(sheetId) })
  const globalViewGen = sheetIds.reduce((sum, id) => sum + revisionOf(id), 0)
  const ownerOf = (): WorkbookOwner => ({
    docId, globalViewGen: remote?.workbookOwner().globalViewGen ?? globalViewGen, workbookValueGen: remote?.workbookOwner().workbookValueGen ?? 0, fontGen,
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
  const live = useLive()
  // A layout effect: its cleanup disconnects the observer inside the commit
  // that removes the pane, before the browser can report the removal. A report
  // already queued for a detached pane is ignored, so no update is ever queued
  // for an unmounted viewer (React would keep it, and the tree, until the
  // host's root renders again).
  useLayoutEffect(() => {
    const el = paneRef.current
    if (!el) return
    const update = () => {
      if (!live.current || !el.isConnected) return
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
  }, [live])
  const touchControls = useControlMedia('(pointer: coarse)')
  const headerHeight = columnHeaderHeight(touchControls, true, true)
  // The row gutter fits the largest row number shown (it widens past row 9999),
  // measured on the window this sheet last presented.
  const viewportHeight = Math.max(1, pane.height - headerHeight - STATUS_BAR_HEIGHT)
  const gutterOf = (sheetEntry: SheetEntry, win: ScrollWindow | undefined) => win ? rowGutterWidth(lastVisibleRowLabel(sheetEntry.metrics, win, viewportHeight)) : GRID_HEADER_SIZE
  const presented = entry ? windows.get(activeSheetId) : undefined
  const gutter = entry ? gutterOf(entry, presented) : GRID_HEADER_SIZE
  const scrollInset = useMemo(() => showZoomControls && !touchControls ? Object.freeze({ right: 60, bottom: 136 }) : undefined, [showZoomControls, touchControls])
  const viewportOf = (g: number): ViewportSize => ({ width: Math.max(1, pane.width - g), height: viewportHeight, ...(scrollInset ? { inset: scrollInset } : {}) })
  const viewport: ViewportSize = useMemo(() => viewportOf(gutter), [pane, headerHeight, gutter, scrollInset]) // eslint-disable-line react-hooks/exhaustive-deps
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
    const frozen = sheetEntry.frozen ? { frozen: { rows: sheetEntry.frozen.rows, cols: sheetEntry.frozen.cols } } : {}
    let current = windows.get(sheetId)
    if (!current) {
      const created = createScrollWindow(sheetEntry.metrics, viewport, { zoom, deviceScale: dpr, ...frozen })
      // The authored top-left cell of the scrolling pane is the first position.
      current = sheetEntry.frozen ? setScrollOffset(created, sheetEntry.metrics, viewport, initialOrigin(sheetEntry.frozen, created.frozen, sheetEntry.metrics)) : created
    } else if (current.deviceScale !== dpr || !sameFreeze(current.frozen, sheetEntry.frozen)) current = createScrollWindow(sheetEntry.metrics, viewport, { zoom: current.zoom, deviceScale: dpr, offset: current.logicalOrigin, ...frozen })
    else current = refreshScrollWindow(current, sheetEntry.metrics, viewport)
    windows.set(sheetId, current)
    return current
  }
  // Refresh against the CURRENT projection/viewport (a view commit or resize keeps
  // the logical origin and zoom; refresh only clamps).
  const windowKey = useRef('')
  const nextWindowKey = JSON.stringify([activeSheetId, entry?.revision, entry?.structureGen, entry?.resizeGen, viewport.width, viewport.height, dpr, Boolean(scrollInset)])
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
  const wrapperRef = useRef<HTMLDivElement | null>(null)
  const palette = useCanvasColors(wrapperRef, props.theme ?? 'system')
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
  const workerPlans = useMemo(() => remote ? new WorkerViewportPlans(remote) : null, [remote])
  // Repaint memory admitted per record at construction (spec §§5.6, 12.4): O(1)
  // per record against the hard main budget; the cache charge follows each
  // insertion/eviction at once, and the paint batch holds what it keeps alive.
  const layoutAdmission = useMemo<LayoutAdmission>(() => {
    let held = 0
    const publish = () => mainLedger.set('layout', 'paint-records', held)
    return {
      admit: bytes => { if (mainLedger.bytes + bytes > MAIN_THREAD_BUDGETS[budgetProfile]) return false; held += bytes; publish(); return true },
      release: bytes => { held = Math.max(0, held - bytes); publish() },
      hold: bytes => { held += bytes; publish() },
      cacheChanged: delta => mainLedger.set('layout', 'caches', Math.max(0, mainLedger.ownerBytes('layout', 'caches') + delta)),
      nativeChanged: delta => mainLedger.set('canvasText', 'tiles', Math.max(0, mainLedger.ownerBytes('canvasText', 'tiles') + delta)),
      measureChanged: delta => mainLedger.set('canvasText', 'measure', Math.max(0, mainLedger.ownerBytes('canvasText', 'measure') + delta)),
    }
  }, [mainLedger, budgetProfile])
  // One shared, admitted measurement context for every record layout (§12.4).
  const measure = useMemo(() => new SharedMeasureContext(() => layoutAdmission), [layoutAdmission])
  measure.setFontGen(fontGen)
  const layoutHostFor = useMemo(() => new WeakMap<XlsxSheet, CellLayoutHost>(), [measure])
  const presentCell = useSemanticPresentations(entry?.sheet, identity.valueGen, semanticFormatBudget)
  const layoutHost = (sheet: XlsxSheet): CellLayoutHost => {
    let host = layoutHostFor.get(sheet)
    if (!host) layoutHostFor.set(sheet, host = measure.host(sheet))
    return host
  }
  useEffect(() => () => workerPlans?.dispose(), [workerPlans])
  // Two sources: the base present and the find overlay each own (and clear)
  // only their own refusal, so an unrelated present never hides a find one.
  const [overflowNotice, setOverflowNotice] = useState('')
  const [findOverflowNotice, setFindOverflowNotice] = useState('')
  const frameOf = (): SheetFrame | null => {
    // Not redundant with entryFor(): a sibling viewer unmounting or a workbook
    // reset may have closed the shared cache between this render and its layout
    // effect. Rebinding at frame-time ensures entry.geometry paints against a
    // live, open owner instead of a closed one.
    if (entry) bindOverflowCache(entry.geometry, overflowFor())
    if (!entry || !scroll) return null
    return {
      sheet: entry.sheet, geometry: entry.geometry, identity, colors: palette.colors, paletteKey: palette.key, origin: scroll.logicalOrigin, viewport, zoom: scroll.zoom, dpr,
      extent: { width: entry.metrics.cols.totalCss(), height: entry.metrics.rows.totalCss() },
      assets: assets?.lease.assets, layoutCache: entry.layoutCache, layoutAdmission, layoutHost: layoutHost(entry.sheet), merges: entry.merges, ...(mark ? { watermark: mark } : {}),
      structureGen: entry.structureGen,
      ...(scroll.frozen ? { frozen: scroll.frozen } : {}),
      // Row-cache proof gates worker requests: tiles the cache already proves
      // paint from the main-thread planner with no worker wait and no worker
      // command; worker plans are only for unknown or oversized tiles, and the
      // row reads `pending` requests keep blank cells known (S13.3).
      // While the sheet has view-only session sizes the main-thread planner is
      // authoritative: worker plans are laid out in the workbook geometry (§5.6).
      ...(workerPlans && remote ? { plan: tilePlanSource({ store: remote.store(activeSheetId), metrics: entry.metrics, scale: scroll.zoom * dpr, identity, extent: { width: entry.metrics.cols.totalCss(), height: entry.metrics.rows.totalCss() }, plans: workerPlans, resized: !sizesOf(activeSheetId).empty }) } : {}),
      ...(remote ? { pending: (box: LogicalCellBox, visible: boolean) => remote.store(activeSheetId).pendingIn(entry.metrics, box, scroll.zoom * dpr, !visible) } : {}),
    }
  }
  // Viewer-owned main pools (layout/merge caches, controls caches, owned DOM),
  // re-measured after every present and render; the row cache, view buffers,
  // replies and copy sessions charge themselves where they allocate.
  const chargeMain = () => {
    let layout = 0, merges = 0
    for (const sheetEntry of entries.values()) { layout += sheetEntry.layoutCache.stats.bytes; if (!sheetEntry.merges.budgetOwned) merges += sheetEntry.merges.bytes }
    let sized = 0
    for (const sizes of sheetSizes.values()) sized += sizes.bytes
    const updates: Array<[MainPool, string, number]> = [
      ['layout', 'caches', layout], ['merges', 'caches', merges], ['layout', 'resize-sizes', sized],
      // Every element the viewer renders (its toolbar, navigation, headers and
      // status as well as the pane), each with its React fiber and props.
      ['dom', 'grid', heldDomBytes(domClassBytes(domClassCounts(rootRef.current ?? paneRef.current)))],
      ['canvasText', 'tiles', rendererRef.current?.textCacheBytes() ?? 0],
    ]
    if (domReservation.current) { domReservation.current = 0; updates.push(['dom', 'resize-reservation', 0]) }
    // Releases before growth: a re-measure never passes through a transient
    // peak above what it settles at (§12.4 hard limit).
    updates.sort((x, y) => (x[2] - mainLedger.ownerBytes(x[0], x[1])) - (y[2] - mainLedger.ownerBytes(y[0], y[1])))
    for (const [pool, owner, bytes] of updates) mainLedger.set(pool, owner, bytes)
  }
  useEffect(() => { chargeMain() })
  // Purely observational: it never creates or resizes the pool.
  const report = () => {
    chargeMain()
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
    // Tile planning lifecycle runs here, post-commit: frameOf itself only
    // describes the frame, so a render that never commits cannot cancel
    // requests for frames that do.
    measure.trim()
    if (scroll) {
      // Pinned strips paint from origin 0 on their axis: their plans stay alive,
      // but only the visible regions themselves (never their enclosing hull).
      workerPlans?.beginFrame(identity, { box: { ...scroll.logicalOrigin, width: viewport.width / scroll.zoom, height: viewport.height / scroll.zoom }, zoom: scroll.zoom, deviceScale: dpr }, paneLogicalBoxes(scroll.frozen, scroll.logicalOrigin, scroll.zoom, viewport), entry?.resizeGen ?? 0)
    }
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
      ctx.fillStyle = palette.colors['pending-text'] ?? '#666666'
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
    const presented = renderer.present(ctx, frame)
    biasRef.current = presented.bias
    setOverflowNotice(presented.resourceNotice ?? '')
    report()
    // Bounded adjacent prefetch after the visible frame is up, in slices of at
    // most PREFETCH_SLICE_MS so a costly ring never blocks a frame; the batch
    // (and its single detail budget) resumes until done or superseded.
    const batch = renderer.beginPrefetch(frame)
    let timer: ReturnType<typeof setTimeout>
    const step = () => {
      const until = performance.now() + PREFETCH_SLICE_MS
      const done = renderer.prefetchSlice(batch, () => performance.now() >= until)
      report()
      if (!done) timer = setTimeout(step, 0)
    }
    timer = setTimeout(step, 0)
    return () => clearTimeout(timer)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [identityKey, scroll, viewport, dpr, assets, surfaceBytes, profile, mark, palette.key, loadTick, entry?.structureGen, entry?.resizeGen])
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
  const finder = useMemo(() => new FindController({ currentOwner: () => ownerRef.current(),
    ...(remote ? { createSession: (input, query, options) => remote.findSession(input.owner, query, options) } : {}),
  }), [remote])
  useEffect(() => () => finder.dispose(), [finder])
  const cellAt = (sheetEntry: SheetEntry, at: SourceCell): XlsxCell | undefined => sheetEntry.geometry.cellsOf(at.row)?.find(cell => cell.col === at.col)
  const findInput = (): FindInput => ({
    owner: ownerOf(),
    sheets: remote ? [] : sheetIds.map(sheetId => {
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
    const byId = new Map(remote ? [] : sheetIds.map(id => [id, entryFor(id)]))
    const options: FindOptions = {
      formatBudget: mainFormatBudget(mainLedger, MAIN_THREAD_BUDGETS[budgetProfile]),
      // The same canonical display text the cell paints (authored formats, semantics).
      textOf: (cell, sheetId, allocation) => {
        const sheetEntry = byId.get(sheetId)
        const styled = sheetEntry && cellAt(sheetEntry, cell)
        // Worker-backed cells carry a style id; the workbook table resolves it.
        if (sheetEntry && !styled && remote) return cellDisplayText(sheetEntry.sheet, { value: cell.value, style: remote.styleOf(cell.styleIndex) }, allocation)
        return sheetEntry ? cellDisplayText(sheetEntry.sheet, styled ?? cell, allocation) : String(cell.value ?? '')
      },
    }
    runArgs.current = { input, query, options }
    // A resident search reads every sheet's order and mask across awaits: its
    // run holds those entries until it is replaced and has settled (§12.4).
    const searched = [...byId.values()].map(sheetEntry => holds.extra(sheetEntry))
    const run = finder.run(input, { text: query, caseSensitive: false }, options)
    void run.then(published => {
      if (cancelled) return
      pageRef.current = published
      activeRef.current = published && published.hits.length ? 0 : -1
      setPage(published)
      setActiveHit(activeRef.current)
      setSearching(false)
    }).catch(error => {
      if (cancelled) return
      setSearching(false)
      pageRef.current = null
      activeRef.current = -1
      setPage(null)
      setActiveHit(-1)
      setFindFailure({ generation: searchGen.current, message: error instanceof Error ? error.message.slice(0, 256) : 'Search failed.' })
      setAnnouncement(error instanceof Error ? error.message.slice(0, 256) : 'Search failed.')
    })
    return () => { cancelled = true; void run.catch(() => {}).finally(() => { for (const release of searched) release() }) }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [query, ownerKey])
  // A finished page over sheets that were still importing says so instead of
  // pretending completeness ("some sheets still loading"). Re-run it once its
  // listed sheets complete, so the results catch up with the import (NIT-1).
  // Sheet IDs are arbitrary strings: the key is a lossless JSON array.
  const incompleteKey = page && page.complete && page.incomplete.length ? JSON.stringify([...page.incomplete].sort()) : ''
  const refreshAttempt = useRef<{ args: object; key: string } | null>(null)
  useEffect(() => {
    if (!remote || !incompleteKey || !runArgs.current) return
    const args = runArgs.current
    const listed = JSON.parse(incompleteKey) as string[]
    // Cleared on cleanup (unmount, a new page or owner): a refresh that settles
    // later publishes nothing, and one still running is cancelled.
    let live = true, refreshing = false
    const check = () => {
      if (!live || runArgs.current !== args || finder.running) return
      const done = listed.filter(sheetId => remote.loadState(sheetId).completeness === 'complete')
      if (!done.length) return
      const key = JSON.stringify([listed, done])
      if (refreshAttempt.current?.args === args && refreshAttempt.current.key === key) return
      refreshAttempt.current = { args, key }
      refreshing = true
      setSearching(true)
      const input = findInput()
      void finder.run(input, { text: args.query, caseSensitive: false }, args.options).then(published => {
        refreshing = false
        // Unmounted, superseded, or a newer search (or a cleared query) owns the results now.
        if (!live || runArgs.current !== args) return
        pageRef.current = published
        activeRef.current = published && published.hits.length ? 0 : -1
        setPage(published)
        setActiveHit(activeRef.current)
        setSearching(false)
      }).catch(error => {
        refreshing = false
        if (!live || runArgs.current !== args) return
        setSearching(false)
        setFindFailure({ generation: searchGen.current, message: error instanceof Error ? error.message.slice(0, 256) : 'Search failed.' })
        setAnnouncement(error instanceof Error ? error.message.slice(0, 256) : 'Search failed.')
      })
    }
    // The import may already have finished before this page was published.
    const unsubscribe = remote.subscribe(check)
    check()
    return () => {
      live = false
      unsubscribe()
      // Only the refresh this effect started, and only while it still owns the search.
      if (refreshing && runArgs.current === args) finder.cancel()
    }
    // Re-bound to every published page and query: a new search with the same
    // incomplete-sheet list must be refreshed with ITS arguments, and a page
    // published after its sheets completed is checked at once.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [remote, incompleteKey, ownerKey, query, page])
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
  const searchFailure = findFailure && findFailure.generation === searchGen.current ? `search stopped${remote?.queryNotice === findFailure.message ? '' : `: ${findFailure.message}`}` : undefined
  const searchNote = [page ? findNote(page) : undefined, searchFailure].filter(Boolean).join('; ') || undefined
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
    let result = navigateToHit({ session, windows, viewport, frozen: sheetId => entries.get(sheetId)?.frozen }, hit)
    if (result.kind === 'visible') {
      const nextGutter = gutterOf(sheetEntry, result.window)
      if (nextGutter !== gutter) result = navigateToHit({ session, windows, viewport: viewportOf(nextGutter), frozen: sheetId => entries.get(sheetId)?.frozen }, hit)
    }
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
    const resources = controlResources.current ??= createViewerControlResources(budgetProfile, undefined, mainLedger)
    const sheetId = activeSheetId
    const ready = viewerControls({ doc, source, resources, sheetId, ledger: mainLedger, snapshot: () => snapshot(sheetId),
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
  // The header sort in flight (spec §12.2 acknowledgment): shown with the click,
  // cleared when its own job ends; any newer job or a sheet switch replaces it.
  const [sortPending, setSortPending] = useState<{ col: number; job: AbortController } | null>(null)
  const domHold = useRef<{ job: AbortController; bytes: number } | null>(null)
  const heldDomBytes = (measured: number) => Math.max(measured, domHold.current?.bytes ?? 0)
  useEffect(() => { closePopup(); controlJob.current?.abort(); setSortPending(null) }, [activeSheetId])
  // A job's hold ends in the commit that no longer shows its pending glyph
  // (completion, refusal, cancellation or a sheet switch), never when the job
  // merely settles: until React commits, the DOM still holds the smaller
  // spinner and a re-measure would publish the dip. A newer job keeps its own
  // hold. Layout phase: before the observer's microtask and chargeMain, which
  // then measure the restored DOM.
  useLayoutEffect(() => {
    const hold = domHold.current
    if (!hold || sortPending?.job === hold.job) return
    domHold.current = null
    const el = rootRef.current ?? paneRef.current
    if (live.current && el) mainLedger.set('dom', 'grid', domClassBytes(domClassCounts(el)))
  }, [sortPending, mainLedger, live])
  useEffect(() => () => controlJob.current?.abort(), [])
  const sortCycle = async (col: number, shift: boolean) => {
    const sheetId = activeSheetId
    controlJob.current?.abort()
    const abort = new AbortController(); controlJob.current = abort
    // The pending glyph has fewer elements than the resting one, and the DOM
    // charge is unconditional: hold it at its current bytes until this job
    // ends, so admitted work (a ring-tile prefetch) cannot take the dip and the
    // restored glyph never re-measures the ledger past the hard budget (§12.4).
    domHold.current = { job: abort, bytes: Math.max(domHold.current?.bytes ?? 0, mainLedger.ownerBytes('dom', 'grid')) }
    // Acknowledge before any await: admission over a large paged range can take long.
    setSortPending({ col, job: abort })
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
    // The DOM hold ends in the commit that clears this pending state (above).
    finally { setSortPending(pending => (pending?.job === abort ? null : pending)) }
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
    // The copy reads this entry's order, mask and projection across awaits.
    const releaseEntry = holds.extra(entry)
    const copyOwner = `session:${++copySerial}`
    let held = 0, closed = false, columnBytes = 0
    const admit = (next: number): boolean => {
      if (closed || !Number.isSafeInteger(next) || next < 0 || next > copySessionBudget(budgetProfile)) return false
      if (next > held && mainLedger.bytes - held + next > MAIN_THREAD_BUDGETS[budgetProfile]) return false
      mainLedger.set('copy', copyOwner, next)
      held = next
      return true
    }
    const reserveColumns = (next: number) => {
      if (!admit(next)) throw new Error('Copy is too large. Select a smaller range.')
      columnBytes = next
    }
    setCopyPending(true)
    try {
      // Column visibility is compact authored geometry, not the selected cell product.
      const hiddenCols: SourceRowInterval[] = []
      const pushHidden = (r0: number, r1: number) => {
        reserveColumns((hiddenCols.length + 1) * 64)
        hiddenCols.push({ r0, r1 })
      }
      let start: number | null = null
      for (let col = 0; col < entry.metrics.cols.logicalCount; col++) {
        if (entry.metrics.cols.size(col) === 0) { start ??= col }
        else if (start !== null) { pushHidden(start, col - 1); start = null }
      }
      if (start !== null) pushHidden(start, entry.metrics.cols.logicalCount - 1)
      const count = entry.metrics.cols.logicalCount
      let columns: RowVisibility
      if (!hiddenCols.length) {
        const slot = (at: number) => { validateCoordinate(at, 0); if (at >= count) throw new RangeError('Display slot outside mask') }
        columns = { visible: at => { slot(at); return true },
          rank: end => { if (!Number.isSafeInteger(end) || end < 0 || end > count) throw new RangeError('Invalid rank endpoint'); return end },
          select: rank => Number.isSafeInteger(rank) && rank >= 0 && rank < count ? rank : null,
          next: (at, direction) => { slot(at); if (direction !== 1 && direction !== -1) throw new RangeError('Invalid traversal direction'); const next = at + direction; return next >= 0 && next < count ? next : null } }
      } else {
        const words = Math.ceil(count / 32), summaries = Math.ceil(words / 32) + 1
        // Input intervals + cloned/sorted intervals + snapshot slots, sorting
        // scratch, two word arrays and summaries are simultaneously live.
        reserveColumns(hiddenCols.length * 160 + 8 * words + 4 * summaries)
        columns = visibilityFromHidden(createIntervalSet(hiddenCols), count)
        hiddenCols.length = 0
        reserveColumns(4 * words + 4 * summaries)
      }
      const result = await copySelection({ store: source.store(activeSheetId), owner: { ...identity }, currentOwner: () => current.current.identityOf(current.current.activeSheetId),
        sessionBudgetBytes: copySessionBudget(budgetProfile) - columnBytes,
        selection, order: entry.view.order, visibility: entry.view.visibility, columns,
        allowed: () => current.current.allowCopy, signal: abort.signal, admitBytes: bytes => admit(columnBytes + bytes),
        retainBytes: (targets, bytes, commit) => { mainLedger.adoptWhileReachable('copy', copyOwner, 'document', targets, bytes, commit); held -= bytes }, onPending: () => { if (live.current && token === copyToken.current) setAnnouncement('Copy is waiting for values; cancel to stop.') },
        format: (cell, allocation) => { const style = remote ? remote.styleOf(cell.styleIndex) : cellAt(entry, cell)?.style; return cellDisplayText(entry.sheet, { value: cell.value, style }, allocation) },
      })
      if (!live.current || token !== copyToken.current) return
      setAnnouncement(result.kind === 'copied' ? 'Selection copied.' : result.kind === 'refused' ? result.message : result.kind === 'stale' ? 'Source or view changed; copy again.' : 'Copy cancelled.')
    } catch (error) {
      if (live.current && token === copyToken.current) setAnnouncement(error instanceof Error ? error.message : 'Copy unavailable; retry.')
    } finally {
      closed = true
      held = 0
      mainLedger.release('copy', copyOwner)
      releaseEntry()
      if (live.current && token === copyToken.current) setCopyPending(false)
    }
  }
  const [valueNotice, setValueNotice] = useState(() => savedValuesNotice(source.valueStatus?.()))
  useEffect(() => {
    const refresh = () => {
      if (!live.current) return false
      const status = source.valueStatus?.()
      setValueNotice(savedValuesNotice(status))
      return status?.policy === 'cached' && status.formulas === undefined && !status.ended
    }
    if (!refresh()) return
    const timer = setInterval(() => { if (!refresh()) clearInterval(timer) }, 100)
    return () => clearInterval(timer)
  }, [source, live])
  const notices = [...new Set([...(doc.diagnostics ?? []).filter(d => d.feature !== SAVED_VALUES_DIAGNOSTIC && ['alternate-content-data', 'large-file-compatibility', 'possibly-damaged'].includes(d.feature)).map(d => d.message), ...(source.openDiagnostics ?? []).filter(d => d.code !== SAVED_VALUES_DIAGNOSTIC).map(d => d.message), ...(valueNotice ? [valueNotice] : [])])]
  // A worker-backed sheet is complete only once its import finished (LIFE-03).
  const loadState = remote?.loadState(activeSheetId)
  // Retained font diagnostics (§9.1): bounded and deduplicated by family.
  const pagedDrawingNotices = activeDrawings ? drawingNotices(activeDrawings) : []
  const activeCharts = entry ? sheetCharts(entry.sheet.drawings) : []
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
    // Frozen panes: each region (body, pinned strips) paints its own clip with
    // its effective origin; without panes this is the single body region.
    const regions = paneRegions(scroll.frozen, scroll.logicalOrigin, scroll.zoom, viewport)
    const pinned = regions.length > 1
    let findRefusal = '', incomplete = false
    for (const region of regions) {
    const origin = region.origin
    const pane = pinned ? { x: origin.x + region.clip.x / scroll.zoom, y: origin.y + region.clip.y / scroll.zoom, width: region.clip.width / scroll.zoom, height: region.clip.height / scroll.zoom } : visibleLogicalViewport(scroll, viewport)
    const snap = (value: number) => (value * scroll.zoom * dpr - Math.round(value * scroll.zoom * dpr)) / dpr
    const bias = region.kind === 'body' ? biasRef.current : { x: region.kind === 'top' ? biasRef.current.x : snap(origin.x), y: region.kind === 'left' ? biasRef.current.y : snap(origin.y) }
    if (pinned) { ctx.save(); ctx.beginPath(); ctx.rect(region.clip.x, region.clip.y, region.clip.width, region.clip.height); ctx.clip() }
    const toViewport = (rect: HighlightRect): HighlightRect => ({ x: (rect.x - origin.x) * scroll.zoom + bias.x, y: (rect.y - origin.y) * scroll.zoom + bias.y, width: rect.width * scroll.zoom, height: rect.height * scroll.zoom })
    // Search: lazily materialized geometry for this page's hits on the pane only.
    const highlights: HighlightRect[] = [], active: HighlightRect[] = []
    if (page) {
      const host = layoutHost(entry.sheet)
      page.hits.forEach((hit, index) => {
        if (hit.sheetId !== activeSheetId) return
        const box = entry.metrics.sourceCellToLogical(hit.source)
        if (!box || box.x >= pane.x + pane.width || box.x + box.width <= pane.x || box.y >= pane.y + pane.height || box.y + box.height <= pane.y) return
        const cell = cellAt(entry, hit.source)
        let geometry
        try { geometry = cell && hitGeometry(hit, cell, entry.metrics, identity, host, entry.layoutCache, cellTextOverflow(entry.geometry, hit.source), layoutAdmission) }
        catch (error) { if (!(error instanceof OverflowResourceError)) throw error; findRefusal ||= error.message; return }
        if (!geometry) return
        const rects = geometry.rects.map(toViewport)
        highlights.push(...rects)
        if (index === activeHit) active.push(...rects)
      })
    }
    paintFindOverlay(ctx, highlights, active, palette.colors)
    if (selection.ranges.length) {
      const records = overlayRecords(entry.metrics, pane, profile)
      // Past the per-cell cap the membership is coalesced into pixel groups, never dropped.
      const coarse = records ? null : coarseSelectionRecords(selection, entry.metrics, pane, scroll.zoom * dpr)
      incomplete ||= Boolean(coarse?.incomplete)
      const geometry = records
        ? gridSelectionOverlay({ contains: at => selectionContains(selection, at) }, records, pane, scroll.zoom * dpr)
        : gridSelectionOverlay({ contains: () => true }, coarse!.records, pane, scroll.zoom * dpr)
      paintGridSelectionOverlay(ctx, gridGeometryToViewport(geometry, origin, scroll.zoom, bias), dpr, palette.colors)
    }
    if (pinned) ctx.restore()
    }
    setFindOverflowNotice(findRefusal)
    setSelectionIncomplete(selection.ranges.length ? incomplete : false)
    report()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [scroll, selection, page, activeHit, viewport, dpr, identityKey, loadTick, palette.key])

  // Keep the native scroller on the window's physical segment (Go To, keys, recenter).
  const syncing = useRef(false)
  // The window the latest native scroll event produced. Committing it must not
  // write the scroller back: the user may have scrolled further while the commit
  // was pending, and writing the older position would revert that motion (its own
  // scroll event then arrives as an echo and is dropped), losing every other step.
  const nativeWindow = useRef<ScrollWindow | null>(null)
  useLayoutEffect(() => {
    const el = scrollerRef.current
    if (!el || !scroll) return
    if (scroll === nativeWindow.current) return
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
    const next = scrollPhysicalTo(scroll, entry.metrics, viewport, { x: el.scrollLeft, y: el.scrollTop })
    // A recentered window moved its physical segment: that one must be written.
    nativeWindow.current = Math.abs(next.physicalPosition.x - el.scrollLeft) < 0.5 && Math.abs(next.physicalPosition.y - el.scrollTop) < 0.5 ? next : null
    setWindow(activeSheetId, next)
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

  // Ctrl/Cmd+wheel anywhere over the viewer (headers, navigator, search and zoom
  // controls) never zooms the browser page; React's wheel listeners are passive,
  // so this one is native, on the outer wrapper that also holds the controls.
  const rootRef = useRef<HTMLDivElement | null>(null)
  // DOM a child commits on its own (a popup's value list arriving through its
  // own state) is charged when it lands, not at the next parent render: one
  // MutationObserver callback per task re-counts the viewer's elements by class.
  useLayoutEffect(() => {
    const el = rootRef.current
    if (!el || typeof MutationObserver === 'undefined') return
    const observer = new MutationObserver(() => { if (live.current && el.isConnected) mainLedger.set('dom', 'grid', heldDomBytes(domClassBytes(domClassCounts(el)))) })
    observer.observe(el, { childList: true, subtree: true })
    return () => observer.disconnect()
  }, [mainLedger, live])
  // Once per mount: is officeview.css loaded? (spec §5, R10)
  useEffect(() => { if (wrapperRef.current) warnIfStylesMissing(wrapperRef.current) }, [])
  useEffect(() => {
    const el = wrapperRef.current
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
  const [printPages, setPrintPages] = useState<ReadonlyArray<{ name: string; page: Paintable | null; scale: number; refusedImages: number }> | null>(null)
  const printLease = useRef<PaintableArray | null>(null)
  const printToken = useRef(0)
  const allowPrintRef = useRef(allowPrint)
  allowPrintRef.current = allowPrint
  const releasePrint = () => { setPrintPages(null); setPrintReserved(0); printLease.current?.dispose(); printLease.current = null }
  const preparePrint = async () => {
    if (!allowPrintRef.current) return
    // Static print needs every cell on the main thread; a worker-backed workbook never has them.
    if (remote) { setAnnouncement(PAGED_PRINT_MESSAGE); return }
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
      // Crisp print: PRINT_SCALE device pixels per spec px when that store fits
      // (budget and axis ceiling), else 1x; neither fits → omitted with a notice.
      const bytesAt = (scale: number) => page.spec.widthPx * scale * page.spec.heightPx * scale * 4
      const fits = (scale: number) => bytesAt(scale) <= remaining && Math.max(page.spec.widthPx, page.spec.heightPx) * scale <= PRINT_MAX_AXIS
      const scale = fits(PRINT_SCALE) ? PRINT_SCALE : fits(1) ? 1 : 0
      if (scale) { remaining -= bytesAt(scale); reserved += bytesAt(scale) }
      return { name: doc.sheets[index]?.name ?? `Sheet ${index + 1}`, page: scale ? page : null, scale: scale || 1, refusedImages: pages.refusedImages?.[index] ?? 0 }
    })
    printLease.current?.dispose()
    printLease.current = pages
    // Reserve before the stores exist: the pool and pane admission shrink first.
    setPrintReserved(reserved)
    setPrintPages(units)
  }
  const printRefs = useRef<Array<HTMLCanvasElement | null>>([])
  const instance = useId()
  useLayoutEffect(() => {
    if (!printPages) { yieldPrint(instance); return }
    const token = printToken.current
    // One unit at a time: each admitted store is sized and painted in turn.
    printPages.forEach((unit, index) => {
      const canvas = printRefs.current[index]
      if (!unit.page || !canvas) return
      sizeCanvas(canvas, unit.page.spec.widthPx * unit.scale, unit.page.spec.heightPx * unit.scale)
      const ctx = canvas.getContext('2d')
      if (ctx) { ctx.setTransform(unit.scale, 0, 0, unit.scale, 0, 0); unit.page.paint(ctx) }
    })
    if (token === printToken.current && allowPrintRef.current) { claimPrint(instance); window.print?.() }
  }, [printPages])
  useEffect(() => {
    if (!allowPrint) return
    window.addEventListener('afterprint', releasePrint)
    const host = globalThis.document
    const media = host.createElement('style')
    media.setAttribute('data-officeview-sheet-print-media', '')
    media.textContent = sheetPrintCss(instance)
    host.head.appendChild(media)
    return () => {
      // Permission withdrawn or unmounted: any in-flight preparation is now stale.
      printToken.current++
      window.removeEventListener('afterprint', releasePrint)
      media.remove()
      yieldPrint(instance)
      releasePrint()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allowPrint])

  // Print guard while this viewer is mounted: hides this viewer (its outer
  // wrapper, so the floating chrome too) and never another, allowed viewer.
  useEffect(() => {
    if (allowPrint) return
    const host = globalThis.document
    const guard = host.createElement('style')
    guard.setAttribute('data-officeview-print-guard', '')
    guard.textContent = `@media print{[data-officeview-sheet-instance="${instance}"]{display:none!important}}`
    host.head.appendChild(guard)
    return () => guard.remove()
  }, [allowPrint, instance])

  const onGridChange = (transition: GridTransition) => {
    cancelCopy()
    session.setSelection(activeSheetId, transition.selection)
    setAnnouncement(transition.announcement)
    setInspection(undefined)
    let nextWindow = transition.window
    if (entry && nextWindow !== scroll && transition.selection.active) {
      const nextGutter = gutterOf(entry, nextWindow)
      // Re-fit the step at the new gutter width only: a frozen sheet must not
      // take the explicit Go To/Find body top-left policy here (that would jump).
      if (nextGutter !== gutter) nextWindow = scrollIntoView(nextWindow, { metrics: entry.metrics, viewport: viewportOf(nextGutter), view: { ...entry.view, selection: transition.selection } }, transition.selection.active)
    }
    if (nextWindow !== scroll) setWindow(activeSheetId, nextWindow)
    else rerender()
  }
  const onNavigate = (request: NavigatorRequest) => {
    if (!entry || !scroll) return
    const state = session.get(activeSheetId)
    const result = navigateSelection(state.selection, entry.metrics, request, entry.view)
    setAnnouncement(result.announcement)
    if (result.kind === 'hidden') { inspectHidden(activeSheetId, result.target, result.hiddenReason ?? 'authored-hidden'); return }
    setInspection(undefined)
    let moved = scrollNavigateTo(scroll, { metrics: entry.metrics, viewport, view: { ...entry.view, selection: result.selection } }, request, state.selection.active)
    const nextGutter = gutterOf(entry, moved.window)
    if (nextGutter !== gutter) moved = scrollNavigateTo(scroll, { metrics: entry.metrics, viewport: viewportOf(nextGutter), view: { ...entry.view, selection: result.selection } }, request, state.selection.active)
    cancelCopy()
    session.setSelection(activeSheetId, result.selection)
    setWindow(activeSheetId, moved.window)
  }
  // ---- view-only row/column resizing (spec §5.6) ----
  // One commit per gesture: the active sheet's session size changes, the tiles
  // its paint can reach are dropped, and the rebuilt entry (new geometry, new
  // window) repaints and re-lays-out everything that depends on geometry.
  const dropTiles = (sheetId: string, region: { axis: 'x' | 'y'; from: number } | null) => {
    rendererRef.current?.cache.invalidate(key => {
      if (key.identity.sheetId !== sheetId) return false
      if (!region) return true
      const box = tileLogicalBox(key.col, key.row, key.zoom, key.dpr)
      return region.axis === 'x' ? box.x + box.width > region.from : box.y + box.height > region.from
    })
  }
  const resizeLine = (request: ResizeRequest) => {
    if (!entry || !scroll) return
    const sheetId = activeSheetId, metrics = entry.metrics
    const column = request.kind === 'column'
    const slot = column ? -1 : metrics.order.sourceToDisplay(request.index)
    const name = column ? `Column ${columnToLetters(request.index)}` : `Row ${slot + 1}`
    const current = column ? metrics.cols.size(request.index) : metrics.rows.size(slot)
    if (!(current > 0)) { setAnnouncement(`${name} is hidden; it cannot be resized.`); return }
    let size: number
    if (request.fit) {
      // A per-fit measurement owner, charged while it runs and dropped after (§12.4).
      let heldHeap = 0, heldNative = 0
      const fit = createFitMeasure({ sheet: entry.sheet, admit: ({ heap, native }) => {
        if (mainLedger.bytes - heldHeap - heldNative + heap + native > MAIN_THREAD_BUDGETS[budgetProfile]) return false
        mainLedger.set('layout', 'resize-measure', heldHeap = heap)
        mainLedger.set('canvasText', 'resize-measure', heldNative = native)
        return true
      } })
      try {
        size = autoFitSize(request.kind, request.index, {
          geometry: entry.geometry, sheet: entry.sheet, pane: visibleLogicalViewport(scroll, viewport), regions: paneLogicalBoxes(scroll.frozen, scroll.logicalOrigin, scroll.zoom, viewport).reverse(), measure: fit.measure, allocation: fit.allocation, ...(fit.layoutHeight ? { layoutHeight: fit.layoutHeight } : {}),
          knownRows: limit => remote ? remote.store(sheetId).knownRows(limit) : entry.sheet.rows.slice(0, limit).map(row => row.index),
          ...(remote ? { styleOf: (index: number) => remote.styleOf(index) } : {}),
        })
      } catch (error) {
        if (!(error instanceof AutoFitRefusedError)) throw error
        setAnnouncement(`${name} was not fitted: ${error.message}`)
        return
      } finally { mainLedger.release('layout', 'resize-measure'); mainLedger.release('canvasText', 'resize-measure') }
    } else size = request.size ?? current + (request.delta ?? 0)
    const sizes = sizesOf(sheetId), axis = column ? 'col' : 'row'
    const previousSize = sizes.get(axis, request.index)
    const prospective = sizes.preview(axis, request.index, size)
    if (!prospective) { setAnnouncement(`${name} was not resized: this sheet already has the maximum number of resized ${column ? 'columns' : 'rows'}. Reset sizes to start over.`); return }
    const applied = (column ? prospective.cols : prospective.rows).get(request.index)!
    const limit = RESIZE_LIMITS[axis]
    const bound = applied === limit.max ? ' (maximum)' : applied === limit.min ? ' (minimum)' : ''
    const announcement = `${name} ${column ? 'width' : 'height'} ${applied} pixels${bound}${request.fit ? ', fitted to content' : ''}. View only; the file is unchanged.`
    if (previousSize === applied) { setAnnouncement(announcement); return }
    const region = resizeRegion(entry.geometry, request.kind, request.index)
    const candidate = admitCandidate(sheetId, prospective, name, shown => (column ? shown.cols.size(request.index) : shown.rows.size(shown.order.sourceToDisplay(request.index))) === applied, `${applied} pixels cannot be represented beside this sheet's sizes`)
    if (!candidate) return
    sizes.set(axis, request.index, applied)
    publishEntry(candidate, sizes.gen)
    setAnnouncement(announcement)
    dropTiles(sheetId, region)
    rerender()
  }
  /**
   * Admission of a session-size rebuild (§12.4): the candidate is built off to
   * the side and measured with the entry's own charge functions; that exact
   * amount, plus the semantic DOM the new geometry reveals, must fit beside
   * everything the ledger holds now (the current entry stays held until the
   * next display commits). A candidate that does not fit, or does not show the
   * size, is discarded unpublished and the refusal is announced.
   */
  const admitCandidate = (sheetId: string, sizes: SessionSizes, name: string, represents: (metrics: ProjectedSheetMetrics) => boolean, unrepresentable: string): Candidate | null => {
    const budget = MAIN_THREAD_BUDGETS[budgetProfile], current = entries.get(sheetId)!
    const refuse = () => { setAnnouncement(`${name} was not resized: not enough memory to lay the sheet out again (main memory budget).`); return null }
    // 1. Provisional reservation BEFORE construction: an upper bound of what
    //    building the candidate allocates, derived from its inputs alone.
    const upper = constructionUpperBytes(sheetId, current, sizes)
    if (mainLedger.bytes + upper > budget) return refuse()
    mainLedger.set('geometry', 'resize-candidate', upper)
    let candidate: Candidate | null = null
    try {
      const prepared = prepareEntry(sheetId, sizes)
      if ('candidate' in prepared) candidate = prepared.candidate
    } finally { if (!candidate) mainLedger.release('geometry', 'resize-candidate') }
    if (!candidate) return null
    if (!represents(candidate.built.metrics)) { mainLedger.release('geometry', 'resize-candidate'); setAnnouncement(`${name} was not resized: ${unrepresentable}.`); return null }
    // 2. Exact admission: the candidate's measured charge plus the semantic DOM
    //    it reveals, beside everything held now. Repaint memory (layout records,
    //    formatted text) is admitted per record where it is built (LayoutCache
    //    batches, §12.4), never predicted here.
    const growth = Math.max(0, semanticDomBytes(candidate.built.metrics) - semanticDomBytes(current.metrics))
    if (mainLedger.bytes - upper + candidate.bytes + growth > budget) { mainLedger.release('geometry', 'resize-candidate'); return refuse() }
    // 3. Commit: the provisional reservation is released first, then the repaint
    //    growth is reserved (until chargeMain measures it) and the caller
    //    charges the exact candidate: the ledger never holds both reservations.
    mainLedger.release('geometry', 'resize-candidate')
    domReservation.current = growth
    mainLedger.set('dom', 'resize-reservation', growth)
    return candidate
  }
  /** Upper bound of what building a session-size candidate allocates (§12.4),
   * from its inputs only: both axes (segments per authored/hidden/session
   * interval, temporary normalization copies included), the row projection
   * with every permuted band dense (20 B/row) and its sections, the resident
   * cell and merge index or the paged merge index, repositioned drawing copies
   * with their placement axes, and a refreshed visibility mask when the
   * structure changed (authored mask, intersection and their copies). */
  const constructionUpperBytes = (sheetId: string, current: SheetEntry, sizes: SessionSizes): number => {
    let rowIntervals = 0, colIntervals = 0, drawingCount = 0
    if (remote) {
      const layout = remote.layout(sheetId)
      rowIntervals = layout.rowOverrides.length + layout.hidden.length
      colIntervals = layout.colOverrides.length + (layout.hiddenCols?.length ?? 0)
      drawingCount = remote.drawings(sheetId).drawings.length
    } else {
      for (const row of current.sheet.rows) if (row.heightPt !== undefined || row.customHeight || row.hidden) rowIntervals++
      colIntervals = current.sheet.cols.length
      drawingCount = current.sheet.drawings?.length ?? 0
    }
    const axis = axisUpperBytes
    const rows = rowIntervals + sizes.rows.size, cols = colIntervals + sizes.cols.size
    let projection = (2 * current.view.order.bands.length + 2 * rows + 8) * 176
    for (const band of current.view.order.bands) projection += 20 * (band.r1 - band.r0 + 1) + 12
    const drawings = drawingCount ? drawingCount * SIZED_DRAWING_BYTES + 2 * (axis(rows) + axis(cols)) : 0
    const index = (remote ? 0 : current.geometry.cellIndexBytes ?? 0) + current.geometry.merges.estimatedBytes
    const structureChanged = remote !== null && remote.merges(sheetId).gen !== current.structureGen
    const mask = structureChanged ? 4 * (Math.ceil(current.metrics.rows.logicalCount / 32) * 4 + 8192) : 0
    return axis(rows) + axis(cols) + projection + drawings + index + mask
  }
  /** Semantic DOM a geometry shows in the current pane (GridSemantics rows,
   * cells, headers and resize handles), in the dom ledger's class bytes. */
  const semanticDomBytes = (metrics: ProjectedSheetMetrics): number => {
    if (!scroll) return 0
    const pane = visibleLogicalViewport(scroll, viewport)
    const rows = metrics.rows.visibleBetween(pane.y, pane.y + pane.height).count, cols = metrics.cols.visibleBetween(pane.x, pane.x + pane.width).count
    const limits = profile === 'M' ? { cells: 10_000, lines: 2048 } : { cells: 20_000, lines: 4096 }
    if (rows * cols >= limits.cells || rows + cols >= limits.lines) return domClassBytes({ cell: 0, header: 8, other: 2 })
    return domClassBytes({ cell: 2 * rows * cols, header: 3 * rows + 8 * cols, other: 2 * rows + cols })
  }
  const domReservation = useRef(0)
  const resetSizes = () => {
    const sheetId = activeSheetId, sizes = sizesOf(sheetId)
    if (sizes.empty) return
    const candidate = admitCandidate(sheetId, NO_SESSION_SIZES, 'Sizes were', () => true, '')
    if (!candidate) return
    sizes.reset()
    publishEntry(candidate, sizes.gen)
    dropTiles(sheetId, null)
    setAnnouncement('Column widths and row heights restored to the workbook sizes.')
    rerender()
  }
  const resized = Boolean(activeSheetId && !sizesOf(activeSheetId).empty)

  const onSheetChange = (sheetId: string) => {
    if (!sheetIds.includes(sheetId)) return
    cancelCopy()
    entryFor(sheetId)
    setInspection(undefined)
    session.activate(sheetId)
    setActiveSheetId(sheetId)
  }

  const zoom = scroll?.zoom ?? 1
  // Stream F: on a coarse pointer or a narrow pane the zoom controls dock into the
  // toolbar; floating, they covered visible cells on a phone.
  const dockZoom = touchControls || pane.width < NARROW_PANE_PX
  const zoomProps = { zoom, onIn: () => zoomBy(1), onOut: () => zoomBy(-1), onReset: resetZoom, atFit: zoom <= SHEET_MIN_ZOOM }
  const canvasStyle: CSSProperties = { position: 'absolute', left: 0, top: 0, width: viewport.width, height: viewport.height, pointerEvents: 'none' }
  // Coarse pointer: 44px targets (spec §9.3) shown icon-only; the label stays in
  // the accessible name through an sr-only span, so names are unchanged.
  const toolbarBtn = 'ov:rounded-control ov:border ov:border-border ov:bg-field ov:text-text ov:enabled:hover:bg-hover ov:disabled:text-muted ov:disabled:opacity-50 ov:disabled:cursor-default ov:px-2 ov:py-1 ov:text-xs ov:font-medium ov:flex ov:items-center ov:justify-center ov:gap-1 ov:coarse:min-h-11 ov:coarse:min-w-11 ov:focus-visible:outline-2 ov:focus-visible:outline-solid ov:focus-visible:outline-offset-1 ov:focus-visible:outline-accent'
  const toolbarLabel = 'ov:coarse:sr-only'
  return (
    // The outer wrapper is the host's flex item: it fills a sized flex-column host;
    // an unsized host gets the default 400px pane. Chrome motion only (spec §6.4).
    <MotionRoot>
    <div ref={wrapperRef} data-officeview-sheet-instance={instance} {...themeRootProps(props.theme, instance)} style={{ position: 'relative', display: 'flex', flexDirection: 'column', minHeight: 0, flex: '1 1 auto' }}>
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
          background: props.background ?? 'var(--ov-color-backdrop)', display: 'flex', flexDirection: 'column', minHeight: 0, flex: '1 1 auto',
          WebkitUserSelect: allowCopy ? undefined : 'none', userSelect: allowCopy ? undefined : 'none', ...props.style,
        }}
      >
        <div data-officeview-sheet-screen="" style={{ display: 'flex', flexDirection: 'column', minHeight: 0, flex: '1 1 auto' }}>
        <div data-officeview-ui="sheet-toolbar" className="ov:bg-chrome ov:border-b ov:border-border ov:flex ov:flex-wrap ov:items-center ov:gap-x-3 ov:gap-y-1 ov:px-2 ov:py-1">
        <SheetNavigator documentId={docId} sheets={doc.sheets.map((sheet, index) => ({ id: sheetIds[index], name: sheet.name }))}
          activeSheetId={activeSheetId} onSheetChange={onSheetChange} onNavigate={onNavigate} onNotice={setAnnouncement} disabled={!entry} />
        <div data-officeview-ui="sheet-actions" className="ov:flex ov:items-center ov:gap-1">
          <button type="button" className={toolbarBtn} disabled={!allowCopy || !selection.ranges.length || copyPending} onClick={() => void copyGrid()}>
            <Icon icon={Copy} name="copy" size={14} />
            <span className={toolbarLabel}>Copy selection</span>
          </button>
          {copyPending && (
            <button type="button" className={toolbarBtn} onClick={cancelCopy}>
              Cancel copy
            </button>
          )}
          {allowPrint && (
            <button type="button" className={toolbarBtn} onClick={() => void preparePrint()}>
              <Icon icon={Printer} name="printer" size={14} />
              <span className={toolbarLabel}>Print</span>
            </button>
          )}
          <button type="button" className={toolbarBtn} disabled={!resized} aria-description="Restore this sheet's workbook column widths and row heights" onClick={resetSizes}>
            <Icon icon={RotateCcw} name="rotate-ccw" size={14} />
            <span className={toolbarLabel}>Reset sizes</span>
          </button>
        </div>
        {showZoomControls && dockZoom && !popup && <DockedZoomControls {...zoomProps} touch={touchControls} />}
        {showSearch && (
          <div data-officeview-ui="sheet-search" style={{ flex: '1 1 260px', minWidth: 0, display: 'flex', justifyContent: 'flex-end' }}>
            <SearchBar inputRef={inputRef} query={query} onQuery={text => { searchGen.current++; setQuery(text) }} count={matchCount} active={activeHit < 0 ? -1 : pageOffset + activeHit}
              countLabel={countLabel} note={searchNote}
              failure={!page ? searchFailure : undefined}
              indexing={searching} onPrev={() => void stepHit(-1)} onNext={() => void stepHit(1)} />
          </div>
        )}
        </div>
        <div ref={paneRef} data-testid="officeview-sheet-pane" style={{ position: 'relative', flex: `1 1 ${FALLBACK_PANE.height}px`, minHeight: 0, overflow: 'hidden', background: '#ffffff' }}>
          {entry && scroll && (
            <GridSemantics fill gutter={gutter} owner={identity} projection={entry.metrics} view={entry.view} selection={selection} window={scroll} viewport={viewport}
              body={entry.body} onChange={onGridChange} status={{ records: { ...snapshot(activeSheetId).records, completeness: loadState?.completeness ?? 'complete' }, countMode: 'body', loadReport: { status: 'ready', notices: [...notices, ...fontNotices, ...pagedDrawingNotices, ...chartNotices(activeCharts), ...(loadState?.failure ? [loadFailureNotice(loadState.failure, loadState.code)] : [])] } }}
              onResize={resizeLine}
              onSortCycle={(col, shift) => void sortCycle(col, shift)} pendingSortColumn={sortPending?.col ?? null} onFilter={(col, invoker) => { setPopup({ kind: 'filter', col, invoker }) }} onSortDialog={(col, invoker) => { setPopup({ kind: 'sort', col, invoker }) }} onNotice={setAnnouncement}
              popup={popup && controls ? { invoker: popup.invoker, close: closePopup, label: popup.kind === 'filter' ? 'Column filter' : 'Sort worksheet', content: popup.kind === 'filter' ? <FilterDropdown owner={identity} client={controls.client} applied={entry.view.filter} appliedSort={entry.view.sort} col={popup.col} invoker={popup.invoker} embedded onClose={closePopup} onApplied={() => rerender()} onNotice={setAnnouncement} /> : <SortDialog owner={identity} client={controls.client} applied={entry.view.sort} col={popup.col} invoker={popup.invoker} embedded onClose={closePopup} onApplied={() => rerender()} onNotice={setAnnouncement} /> } : undefined} announcement={[sortPending ? `Sorting column ${columnToLetters(sortPending.col)}…` : '', resourceNotice, overflowNotice, findOverflowNotice === overflowNotice ? '' : findOverflowNotice, assets?.sheetId !== activeSheetId ? 'Loading sheet…' : '', assets && assets.sheetId === activeSheetId && assets.lease.refusedImages > 0 ? `${assets.lease.refusedImages} picture${assets.lease.refusedImages === 1 ? ' is' : 's are'} not shown (image memory budget).` : '', assets && assets.sheetId === activeSheetId && assets.lease.chartsUnavailable ? 'Charts are shown as placeholders: the chart engine could not be loaded.' : '', selectionIncomplete ? 'Selection highlight is incomplete at this zoom; zoom in for detail.' : '', remote?.queryNotice === announcement ? '' : remote?.queryNotice, announcement].filter(Boolean).join(' ')} inspection={inspection} profile={profile} label={`Worksheet ${entry.sheet.name}`}
              cellPresentation={at => {
                const cell = cellAt(entry, at)
                if (!cell && remote && !remote.store(activeSheetId).isKnown(at)) return { text: 'Loading…' }
                return cell ? presentCell(cell) : { text: '' }
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
          <SheetChartDescriptions charts={activeCharts} />
        </div>
        </div>
        {allowPrint && (
          <div data-officeview-sheet-print="" aria-hidden="true">
            {!printPages && <p data-testid="officeview-sheet-print-notice">{remote ? PAGED_PRINT_MESSAGE : 'To print this spreadsheet, press Ctrl/Cmd+P inside the viewer.'}</p>}
            <div data-testid="officeview-sheet-print-pages">
              {printPages?.map((unit, index) => unit.page
                ? <div key={index} style={index > 0 ? { breakBefore: 'page', pageBreakBefore: 'always' } : undefined}>
                    <h2 data-officeview-print-sheet-title="">{unit.name}</h2>
                    {unit.page.refusal?.length ? <p>{LARGE_EXPORT_MESSAGE}</p> : null}
                    {unit.refusedImages > 0 && <p>Sheet “{unit.name}”: {unit.refusedImages} picture{unit.refusedImages === 1 ? ' is' : 's are'} omitted from print (image memory budget).</p>}
                    <canvas ref={el => { printRefs.current[index] = el }} style={{ display: 'block', width: unit.page.spec.widthPx, maxWidth: '100%', height: 'auto' }} />
                  </div>
                : <p key={index} style={index > 0 ? { breakBefore: 'page', pageBreakBefore: 'always' } : undefined}>Sheet “{unit.name}” is omitted from print: it exceeds the print memory budget.</p>)}
            </div>
          </div>
        )}
      </div>
      {/* A modal popup (the mobile filter sheet) owns the bottom of the viewer. */}
      {showZoomControls && !dockZoom && !popup && <ZoomControls {...zoomProps} />}
    </div>
    </MotionRoot>
  )
}

/** A display mask AND another (both display-space, same extent). */
function intersectVisibility(mask: RowVisibility, other: CompactRowVisibility): CompactRowVisibility {
  if (mask instanceof CompactRowVisibility && mask.logicalCount === other.logicalCount) {
    const words = new Uint32Array(other.words.length)
    for (let word = 0; word < words.length; word++) words[word] = mask.words[word] & other.words[word]
    return new CompactRowVisibility(other.logicalCount, words)
  }
  return createRowVisibility(other.logicalCount, slot => mask.visible(slot) && other.visible(slot))
}

/** Control-utility rule (spec §6.8): colour, background, padding and border as ov: utilities. */
const DOCKED = 'ov:rounded-control ov:border ov:border-border ov:bg-field ov:text-text ov:enabled:hover:bg-hover ov:disabled:text-muted ov:disabled:opacity-50 ov:disabled:cursor-default ov:px-1.5 ov:py-1 ov:text-xs ov:font-semibold ov:flex ov:items-center ov:justify-center'
/** Below this pane width (CSS px) the zoom controls dock into the toolbar. */
const NARROW_PANE_PX = 600
/** The zoom controls as a toolbar group (same names and test id as the floating
 * cluster); 44px targets on a coarse pointer (spec §9.3). */
function DockedZoomControls({ zoom, onIn, onOut, onReset, atFit, touch }: { zoom: number; onIn: () => void; onOut: () => void; onReset: () => void; atFit: boolean; touch: boolean }): ReactElement {
  const side = touch ? 44 : 28
  const btn: CSSProperties = { minWidth: side, minHeight: side }
  return <div role="group" aria-label="Zoom" data-testid="officeview-zoom-controls" data-officeview-ui="sheet-zoom" className="ov:flex ov:items-center ov:gap-1">
    <button type="button" onClick={onOut} aria-label="Zoom out" disabled={atFit} className={DOCKED} style={btn}><Icon icon={Minus} name="minus" size={14} /></button>
    <button type="button" onClick={onReset} aria-label="Reset zoom" className={DOCKED} style={btn}><Tick k={Math.round(zoom * 100)}>{Math.round(zoom * 100)}%</Tick></button>
    <button type="button" onClick={onIn} aria-label="Zoom in" className={DOCKED} style={btn}><Icon icon={Plus} name="plus" size={14} /></button>
  </div>
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

/** The status-line saved-values notice: with nothing affected, the policy
 * statement alone (no "0 formulas … (file too large)"). */
function savedValuesNotice(status: { formulas: number | undefined; message: string | undefined } | undefined): string | undefined {
  return status?.message !== undefined && status.formulas === 0 ? SAVED_VALUES_STATUS : status?.message
}

/** 1-based label of the last display row a window shows (its row-gutter width). */
function lastVisibleRowLabel(metrics: ProjectedSheetMetrics, window: ScrollWindow, height: number): number {
  const rows = metrics.rows.visibleBetween(window.logicalOrigin.y, window.logicalOrigin.y + height / window.zoom)
  let last = 0
  for (const run of rows.runs) last = Math.max(last, run.end + 1)
  return last || 1
}
