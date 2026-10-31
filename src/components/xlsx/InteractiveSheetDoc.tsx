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
import type { XlsxCell, XlsxCellStyle, XlsxDocument, XlsxDrawing, XlsxSheet } from '../../xlsx/types'
import { markPagedDependentsStale, recomputeDependents, restoreDependentFormulaValues, savedDependentFormulaCells } from '../../xlsx/formula/workbook'
import { formulaError } from '../../xlsx/formula/evaluator'
import type {
  AuthoritativeCell, CacheIdentity, FilterSpec, GridSelection, LogicalCellBox, ProfileId, RowOrder, RowVisibility, ScrollWindow, SheetView, SortSpec, SourceRowInterval, SourceStore, WorkbookOwner,
} from '../../xlsx/view/contracts'
import { SHEET_MIN_ZOOM } from '../../xlsx/view/contracts'
import type { NavigatorRequest, SourceCell } from '../../xlsx/view/coordinates'
import { columnToLetters, parseA1, validateCoordinate } from '../../xlsx/view/coordinates'
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
import { SheetImageCache, imageBudgetFor, installPivotCharts, mainChartParent, preparePrintUnits, prepareSheetAssets, type PrintUnitsOptions, type SheetAssetsLease } from './sheetAssets'
import type { ChartModel } from '../../charts/model'
import { chartLinksPivot } from '../../xlsx/pivotChartLink'
import { overlayChartCell, rememberDependentChart, restoreDependentCharts } from '../../xlsx/pivots/session/dependentCharts'
import { mergeOwedRects } from '../../xlsx/pivots/session/dependentRects'
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
import { documentWithSparklines, sheetWithSparklines, sheetWithoutResolvedSparklines, sparklineNotices } from '../../worker/viewerRemoteSparklines'
import { sheetWithRegionStyles, sheetWithoutRegionStyles } from '../../worker/viewerRemoteRegionStyles'
import { pagedSheetGeometry } from './pagedGeometry'
import type { LedgerHold } from '../../xlsx/view/mainLedger'
import { DisplayHolds } from './displayHolds'
import { SheetChartDescriptions, chartNotices, sheetCharts } from './sheetCharts'
import type { XlsxFrozenPane } from '../../xlsx/frozenPane'
import { initialOrigin, paneLogicalBoxes, paneRegions } from '../../xlsx/view/frozenPanes'
import { SheetSizes, NO_SESSION_SIZES, RESIZE_LIMITS, SIZED_DRAWING_BYTES, axisUpperBytes, cellFont, type SessionSizes } from '../../xlsx/view/axisSizes'
import { columnWidthPx } from '../../xlsx/columnWidth'
import { AutoFitRefusedError, autoFitSize, createFitMeasure, resizeRegion, type ResizeRequest } from './sheetResize'
import { tileLogicalBox } from '../../xlsx/view/tileCache'
import { sparklineInvalidationBox } from '../../xlsx/view/sparklineInvalidation'
import type { ValueOverlay } from '../../xlsx/view/contracts'
import { invalidateRegionStyles, regionStylesFor } from '../../xlsx/tableStyleLoader'
import { pivotRuntimeFor } from '../../xlsx/pivotLoader'
// Type-only: the pivot controls (client, layer, menus) are a lazy chunk loaded on a sheet with pivots (Task 28).
import type * as PivotUi from './pivot/PivotButtonLayer'
import type { PivotDeltaKind } from '../../xlsx/pivots/session/delta'

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
  /** @internal The pivot interaction kinds the menus offer (default: the session's oracle-verified set). */
  pivotVerifiedKinds?: ReadonlySet<PivotDeltaKind>
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
  /** Generation of the worker-resolved sparkline map read by the tile painter. */
  sparklineGen: number
  /** Generation of the worker-resolved table-style layer. */
  regionStyleGen: string
  revision: number
  /** Generation of the merged ranges the geometry holds (worker merges; a resident pivot overlay's merges). */
  structureGen: number
  /** The source store's value generation the geometry holds (a committed pivot overlay, spec §7.9). */
  valueGen: number
  /** Worker-backed sheets: the drawing generation the entry's sheet holds (a pivot chart swap changes only this). */
  drawingGen: number
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
  useEffect(() => {
    if (!remote) return
    return remote.subscribeSparklineEvictions(sheetId => {
      const retained = entries.get(sheetId)
      if (retained) retained.sheet = sheetWithoutResolvedSparklines(retained.sheet)
    })
  }, [remote, entries])
  const hasResidentSparklines = useMemo(() => doc.sheets.some(sheet => Boolean(sheet.resolvedSparklines?.size)), [doc])
  const hasActiveWorkerSparklines = Boolean(remote && remote.layout(activeSheetId).hasSparklines)
  const hasResolvedSparklines = hasResidentSparklines || hasActiveWorkerSparklines
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
  useMemo(() => {
    if (remote) return
    mainLedger.chargeWhileReachable('document', [doc], estimatePayloadBytes(doc))
    // Ruling S3-F: the pivot runtime's retained records copy (Ruling S3-E) is charged from mount, not first lease.
    if (doc.sheets.some(sheet => sheet.pivotTables?.length)) pivotRuntimeFor(doc)?.chargeRetained(mainLedger)
  }, [mainLedger, remote, doc])
  const onMainLedger = props.onMainLedger
  useEffect(() => { onMainLedger?.(mainLedger) }, [mainLedger, onMainLedger])
  const [loadTick, setLoadTick] = useState(0)
  useEffect(() => remote?.subscribe(() => setLoadTick(tick => tick + 1)), [remote])
  useEffect(() => {
    if (!remote) return
    return remote.subscribeRegionStyleEvictions(sheetId => {
      const retained = entries.get(sheetId)
      if (retained) retained.sheet = sheetWithoutRegionStyles(retained.sheet)
    })
  }, [remote, entries])
  // Tiles painted under an older worker layout/merge generation are stale (§10.1).
  const structureGens = useRef(new Map<string, number>()).current
  const presentedLease = useRef<{ sheetId: string; lease: SheetAssetsLease } | null>(null)
  // The drawing generation each sheet's cached tiles were painted under: a worker sheet's drawings can change while
  // it is not shown (its pivot chart replaced by a commit on another sheet), and its tiles are then stale.
  const paintedDrawingGens = useRef(new Map<string, number>()).current
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
  const [assets, setAssets] = useState<{ lease: SheetAssetsLease; fontGen: number; fontKey: string; sheetId: string; sparklineGen: number } | null>(null)
  const [residentDrawingGen, setResidentDrawingGen] = useState(0)
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
  const activeSparklines = remote && activeSheetId && remote.layout(activeSheetId).hasSparklines ? remote.sparklines(activeSheetId) : null
  useEffect(() => {
    if (!activeSheetId) return
    let cancelled = false
    const sheetIndex = sheetIds.indexOf(activeSheetId)
    let assetDoc = activeDrawings ? documentWithDrawings(doc, sheetIndex, activeDrawings) : doc
    if (activeSparklines) assetDoc = documentWithSparklines(assetDoc, sheetIndex, activeSparklines)
    void loadAssets(assetDoc, { ...(fontKey ? { fallbackFonts } : {}), sheetIndex, profile, imageCache: imageCacheFor(), chartOwner: doc, chartParent }).then(next => {
      if (cancelled) { next.dispose(); return }
      if (fontGenRef.current.key !== fontKey || fontGenRef.current.gen === 0) fontGenRef.current = { key: fontKey, gen: fontGenRef.current.gen + 1 }
      setAssets({ lease: next, fontGen: fontGenRef.current.gen, fontKey, sheetId: activeSheetId, sparklineGen: activeSparklines?.gen ?? 0 })
    })
    return () => { cancelled = true }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [doc, fontKey, activeSheetId, profile, activeDrawings?.gen, activeSparklines?.gen, chartParent, residentDrawingGen])
  const committedLease = useRef<SheetAssetsLease | null>(null)
  useEffect(() => {
    const previous = committedLease.current
    committedLease.current = assets?.lease ?? null
    if (previous && previous !== committedLease.current) previous.dispose()
  }, [assets])
  useEffect(() => () => { committedLease.current?.dispose(); committedLease.current = null }, [])
  // The published font generation; a sheet switch (lease pending) never resets it.
  const fontGen = fontGenRef.current.gen

  // ---- pivots (Task 28): committed overlays per resident sheet; the saved definitions the host's document gets back ----
  const pivotOverlays = useRef(new Map<string, ValueOverlay>()).current
  const pivotFormulaContexts = useRef(new Map<string, { sheet: string; def: import('../../xlsx/pivots/model').PivotTableDefinition; cache: import('../../xlsx/pivots/model').PivotCacheDefinition; layout: import('../../xlsx/pivots/engine/place').PivotLayout; layoutCells: Map<string, import('../../xlsx/pivots/engine/place').PivotLayout['cells'][number]>; date1904: boolean; rects: readonly { r0: number; r1: number; c0: number; c1: number }[]; /** A pivot reset to its saved view while others are live (S3-AF): its saved layout answers GETPIVOTDATA; its cells are the sheet's own. */ saved?: true }>()).current
  /** Document-wide dependent generation: a restore of every pivot drops all in-flight dependent work. */
  const pivotLifecycle = useRef(0)
  /** Per-pivot dependent generation (S3-AF): a pivot's newer commit or reset drops only that pivot's in-flight work. */
  const pivotFormulaGens = useRef(new Map<string, number>()).current
  /** Changed rectangles whose recompute a newer commit of the same pivot dropped: the newer one recomputes them too. */
  const pendingDependentRects = useRef(new Map<string, { r0: number; r1: number; c0: number; c1: number }[]>()).current
  /** Which pivots' commits refilled each dependent chart: a reset puts back only the charts no other live pivot feeds. */
  const dependentChartOwners = useRef(new Map<XlsxDrawing, Set<string>>()).current
  /** The resident stale marks each pivot's commits caused, by sheet index: `unsupported` readers keep their saved
   * value (S3-T, S3-Z); `outdated` ones were over the work budget and keep whatever they showed (S3-AE, M-4). */
  const dependentStaleMarks = useRef(new Map<string, { unsupported: Set<number>; /** reader key → its sheet index */ outdated: Map<string, number> }>()).current
  const [pagedDependentNoticeSheets, setPagedDependentNoticeSheets] = useState<ReadonlySet<number>>(() => new Set())
  const [residentDependentNotices, setResidentDependentNotices] = useState<{ unsupported: ReadonlySet<number>; outdated: ReadonlySet<number> }>(() => ({ unsupported: new Set(), outdated: new Set() }))
  const publishDependentNotices = () => {
    const unsupported = new Set<number>(), outdated = new Set<number>()
    for (const marks of dependentStaleMarks.values()) { for (const index of marks.unsupported) unsupported.add(index); for (const index of marks.outdated.values()) outdated.add(index) }
    setResidentDependentNotices({ unsupported, outdated })
  }
  const hasPivots = useMemo(() => doc.sheets.some(sheet => sheet.pivotTables?.length || sheet.pivotRegions?.length), [doc])
  const savedPivotDefs = useMemo(() => doc.sheets.map(sheet => sheet.pivotTables), [doc])
  const savedPivotRegions = useMemo(() => doc.sheets.map(sheet => sheet.pivotRegions), [doc])
  // The saved models of pivot charts a commit replaced (Task 30), by pivot name, in installPivotCharts' order.
  const savedPivotCharts = useRef(new Map<string, { sheet: string; name: string; charts: ChartModel[] }>()).current
  // Non-pivot chart models rebuilt from dependent formulas are restored with the
  // saved workbook on pivot-client teardown, just like pivot-linked charts.
  const savedDependentCharts = useRef(new Map<XlsxDrawing, ChartModel>()).current
  /** Back to the file's pivots: the saved definitions on the host's document and no overlay. A live definition never
   * outlives the sessions that made it (unmount, or a new client). */
  /** Columns a pivot autofit sized, per sheet, in characters (Task 29): their widths are the "current" a later
   * autofit compares with, and they go back to the workbook's when the pivots do. */
  const pivotAutofitChars = useRef(new Map<string, Map<number, number>>()).current
  /** Resizes the changed pivot's columns; returns why it did not, when memory refused it (announced by the caller). */
  const autofitPivot = (sheetId: string, sheet: XlsxSheet, before: Parameters<typeof PivotUi.pivotAutofitColumns>[0]['before'], after: Parameters<typeof PivotUi.pivotAutofitColumns>[0]['after'], shownChars?: (col: number) => number | undefined): string | undefined => {
    // No laid-out sheet to resize (it is not shown): the change keeps the widths.
    if (!pivotUi || !entries.get(sheetId)) return undefined
    const mdw = sheet.columnMetrics?.maxDigitWidthPx ?? 7, sizes = sizesOf(sheetId)
    let fitted = pivotAutofitChars.get(sheetId)
    if (!fitted) pivotAutofitChars.set(sheetId, fitted = new Map())
    const authored = shownChars ?? ((col: number) => sheet.cols.find(spec => col >= spec.min && col <= spec.max)?.widthChars)
    // The measurement, as a fit gesture's (codex review T29-4): a per-fit context whose texts and shaping are admitted
    // to the main ledger as they grow, released when the fit ends; the context goes with it.
    let heldHeap = 0, heldNative = 0
    const fit = createFitMeasure({ sheet, admit: ({ heap, native }) => {
      if (mainLedger.bytes - heldHeap - heldNative + heap + native > MAIN_THREAD_BUDGETS[budgetProfile]) return false
      mainLedger.set('layout', 'pivot-autofit-measure', heldHeap = heap)
      mainLedger.set('canvasText', 'pivot-autofit-measure', heldNative = native)
      return true
    } })
    let widths: Map<number, number>
    try {
      widths = pivotUi.pivotAutofitColumns({ sheet, before, after, mdwPx: mdw,
        measure: fit.measure,
        font: style => cellFont(style).font,
        current: col => fitted!.get(col) ?? authored(col),
        userSet: col => sizes.userSet('col', col) })
    } catch (error) {
      if (!(error instanceof AutoFitRefusedError)) throw error
      return PIVOT_FIT_REFUSED
    } finally { mainLedger.release('layout', 'pivot-autofit-measure'); mainLedger.release('canvasText', 'pivot-autofit-measure') }
    if (!widths.size) return undefined
    // Like a resize (codex review T29-5): the geometry for the new sizes is reserved and admitted before it is built,
    // and published only if it fits; the sizes are never user-set.
    const { sizes: prospective, kept } = sizes.previewMany('col', [...widths].map(([col, chars]) => [col, columnWidthPx(chars, mdw)] as const))
    if (!kept.size) return undefined
    // A hidden column keeps its fitted size for when it is shown; while hidden it projects to 0, which is not a refusal.
    const hiddenBefore = entries.get(sheetId)!.metrics.cols
    let unrepresentable = false
    const candidate = admitCandidate(sheetId, prospective, 'The pivot table\'s column width', shown => {
      const ok = [...kept].every(([col, px]) => { const size = shown.cols.size(col); return size === px || (size === 0 && hiddenBefore.size(col) === 0) })
      unrepresentable = !ok
      return ok
    }, '', true)
    if (!candidate) return unrepresentable ? PIVOT_FIT_UNREPRESENTABLE : PIVOT_FIT_REFUSED
    for (const col of kept.keys()) if (sizes.set('col', col, kept.get(col)!, 'autofit')) fitted.set(col, widths.get(col)!)
    publishEntry(candidate, sizes.gen)
    return undefined
  }
  /** A resident autofit's refusal, said with the change's own announcement (it is made during the commit). */
  const autofitNotice = useRef<string | undefined>(undefined)
  /** Paged commits waiting for their sheet's refreshed pivot display before the autofit (Task 29). */
  // One per pivot, so a second pivot's commit before the refresh does not drop the first's fit (codex review T29-6).
  const pendingPagedFits = useRef<PivotUi.PendingPivotFits<ReturnType<typeof pivotStyleAt>> | null>(null)
  /** The paged fit cells the main thread keeps (a port's shown layouts, waiting fits), admitted to the main ledger. */
  const pivotFitMemory: PivotUi.PivotFitMemory = {
    set: (owner, bytes) => {
      mainLedger.release('layout', owner)
      if (mainLedger.bytes + bytes > MAIN_THREAD_BUDGETS[budgetProfile]) return false
      mainLedger.set('layout', owner, bytes)
      return true
    },
    release: owner => mainLedger.release('layout', owner),
  }
  useEffect(() => {
    const pending = pendingPagedFits.current
    if (!remote || !pivotUi || !pending?.size) return
    for (const sheetId of pending.sheets()) {
      const shownEntry = entries.get(sheetId)
      if (!shownEntry || !shownEntry.sheet.regionStyles) continue
      const ready = pending.take(sheetId, shownEntry.regionStyleGen)
      if (!ready.length) continue
      const mdw = shownEntry.sheet.columnMetrics?.maxDigitWidthPx ?? 7
      try {
        for (const { fit: { before, after }, styleAt } of ready) {
          const metrics = entries.get(sheetId)!.metrics
          const refused = autofitPivot(sheetId, shownEntry.sheet, { def: before.def, read: pivotCellsReader(before.cells), ...(styleAt ? { styleAt } : {}) },
            { def: after.def, read: pivotCellsReader(after.cells), ...(shownEntry.sheet.regionStyles ? { styleAt: pivotStyleAt(shownEntry.sheet.regionStyles)! } : {}) },
            col => metrics.cols.size(col) / mdw)
          if (refused) setAnnouncement(refused)
        }
      } finally { for (const fit of ready) fit.release() }   // charged through the fit (codex re-review T29-R2-3)
      dropTiles(sheetId, null)
      rerender()
    }
  })
  const restorePivots = () => {
    pivotLifecycle.current++
    pendingPagedFits.current?.clear()
    // The autofit widths go back to the workbook's, admitted like a reset (codex re-review T29-R2-1): a shown sheet's
    // geometry for the restored sizes is reserved and admitted before it is built; refused, the widths stay (as
    // session sizes "Reset sizes" can restore). Unmounting rebuilds nothing: the sizes go with the viewer.
    if (!unmounting.current) for (const [sheetId, fitted] of pivotAutofitChars) {
      const sizes = sheetSizes.get(sheetId)
      const cols = sizes ? [...fitted.keys()].filter(col => !sizes.userSet('col', col) && sizes.get('col', col) !== undefined) : []
      if (!sizes || !cols.length) continue
      if (entries.get(sheetId)) {
        const candidate = admitCandidate(sheetId, sizes.previewWithout('col', cols), 'The pivot table\'s column width', () => true, '', true)
        if (!candidate) { setAnnouncement('The pivot table\'s fitted column widths were kept: not enough memory to lay the sheet out again (main memory budget).'); continue }
        for (const col of cols) sizes.restore('col', col, undefined)
        publishEntry(candidate, sizes.gen)
        dropTiles(sheetId, null)
      } else for (const col of cols) sizes.restore('col', col, undefined)
    }
    pivotAutofitChars.clear()
    restoreAllDependents()
    doc.sheets.forEach((sheet, index) => {
      if (sheet.pivotTables === savedPivotDefs[index]) return
      sheet.pivotTables = savedPivotDefs[index]
      sheet.pivotRegions = savedPivotRegions[index]
      invalidateRegionStyles(sheet)
    })
    for (const sheetId of pivotOverlays.keys()) { try { source.store(sheetId).setValueOverlay?.(undefined) } catch { /* released with the viewer */ } }
    pivotOverlays.clear()
    for (const { sheet, name, charts } of savedPivotCharts.values()) for (const index of installPivotCharts(doc, name, charts, sheet)) { const id = sheetIds[index]; if (id !== undefined) dropTiles(id, null) }
    savedPivotCharts.clear()
  }
  /** The document-wide restore (S3-AF): every saved formula value and dependent chart back, every stale mark gone,
   * every in-flight dependent recompute dropped. Only when every pivot is back to its saved view, or on teardown. */
  const restoreAllDependents = () => {
    pivotLifecycle.current++
    pivotFormulaContexts.clear(); pivotFormulaGens.clear(); pendingDependentRects.clear(); dependentChartOwners.clear(); dependentStaleMarks.clear()
    for (const index of restoreDependentFormulaValues(doc)) { const id = sheetIds[index]; if (id !== undefined) dropTiles(id, null) }
    dropDependentChartTiles(restoreDependentCharts(savedDependentCharts))
    if (!unmounting.current) setResidentDependentNotices({ unsupported: new Set(), outdated: new Set() })
  }
  const dropDependentChartTiles = (restored: readonly XlsxDrawing[]) => {
    if (!restored.length) return
    for (const [index, sheet] of doc.sheets.entries()) if (restored.some(node => sheet.drawings?.includes(node))) {
      const id = sheetIds[index]
      if (id !== undefined) dropTiles(id, null)
    }
    if (!unmounting.current) setResidentDrawingGen(generation => generation + 1)
  }
  /** One pivot back to its saved view while others stay live (S3-AF): its stale marks go, and the dependent charts only
   * it fed get their saved models back; charts another live pivot also feeds are refilled by the reset's recompute. */
  const releaseDependentsOf = (key: string) => {
    if (dependentStaleMarks.delete(key)) publishDependentNotices()
    const restored: XlsxDrawing[] = []
    for (const [node, owners] of [...dependentChartOwners]) {
      owners.delete(key)
      if (owners.size) continue
      dependentChartOwners.delete(node)
      const saved = savedDependentCharts.get(node)
      savedDependentCharts.delete(node)
      if (saved && node.content?.kind === 'xlsxChart') { node.content.chart = saved; restored.push(node) }
    }
    dropDependentChartTiles(restored)
  }

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
    sheetId: string; sheet: XlsxSheet; sparklineGen: number; regionStyleGen: string; input: SheetViewInput; structureGen: number; valueGen: number; drawingGen: number; previous: SheetEntry | undefined
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
    let sheet = remote ? sheetWithDrawings(doc.sheets[index], remote.drawings(sheetId)) : doc.sheets[index]
    const layout = remote?.layout(sheetId)
    const regionStyleContents = remote && layout ? [
      ...(layout.hasRegionStyles ? [remote.regionStyles(sheetId)] : []),
      ...(layout.hasPivotDisplay ? [remote.pivotDisplay(sheetId)] : []),
    ] : []
    const regionStyleGen = JSON.stringify(regionStyleContents.map(content => content.gen))
    for (const content of regionStyleContents) sheet = sheetWithRegionStyles(sheet, content)
    const sparklineContent = remote && remote.layout(sheetId).hasSparklines ? remote.sparklines(sheetId) : null
    const sparklineGen = sparklineContent?.gen ?? 0
    if (sparklineContent) sheet = sheetWithSparklines(sheet, sparklineContent)
    let input = viewInput(sheetId)
    const previous = entries.get(sheetId)
    // A resident pivot overlay (Task 28) moves the store's value and merge generations: the geometry is rebuilt over it.
    const structureGen = remote ? remote.merges(sheetId).gen : source.store(sheetId).structureGen ?? 0
    const valueGen = remote ? 0 : source.store(sheetId).stamp.valueGen
    const drawingGen = remote ? remote.drawings(sheetId).gen : 0
    const sized = sheetSizes.get(sheetId), resizeGen = sized?.gen ?? 0
    const same = (entry: SheetEntry) => entry.structureGen === structureGen && entry.valueGen === valueGen && entry.drawingGen === drawingGen && entry.resizeGen === resizeGen && entry.sparklineGen === sparklineGen && entry.regionStyleGen === regionStyleGen
    if (!override && previous && previous.revision === input.revision && same(previous)) return { entry: previous }
    if (previous && input.revision < previous.revision) {
      // Stale commit: never publish an older view. Session sizes or structure
      // that changed since are still applied, to the view already shown.
      if (!override && same(previous)) return { entry: previous }
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
    const overlay = remote ? undefined : pivotOverlays.get(sheetId)
    const built = remote ? pagedSheetGeometry(remote, sheetId, order, visibility, overflowFor(), sizes, true) : buildSheetGeometry(sheet, { order, ...(input.visibility ? { visibility } : {}), sizes, ...(overlay ? { overlay } : {}) })
    if (built.kind !== 'ready') throw new Error(built.diagnostics.map(d => d.message).join('; ') || 'This sheet cannot be laid out')
    const geometryBytes = remote ? (built as unknown as { chargeBytes: number }).chargeBytes : projectionBytes(built.metrics) + axisBytes(built.metrics.cols) + (built.cellIndexBytes ?? 0) + built.merges.estimatedBytes
    const bytes = geometryBytes + (built.sizedDrawings?.length ?? 0) * SIZED_DRAWING_BYTES + viewBufferBytes(remote ? null : order, remote && visibility === input.visibility ? null : visibility)
    return { candidate: { sheetId, sheet, sparklineGen, regionStyleGen, input, structureGen, valueGen, drawingGen, previous, order, baseHidden, visibility, built, bytes } }
  }
  /** Publishes a candidate: charges exactly its measured owners, holds it and
   * retires the previous entry on the next display commit. */
  const publishEntry = (candidate: Candidate, resizeGen: number): SheetEntry => {
    const { sheetId, sheet, sparklineGen, regionStyleGen, input, structureGen, valueGen, drawingGen, previous, order, baseHidden, visibility, built } = candidate
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
      sheet, sparklineGen, regionStyleGen, revision: input.revision, structureGen, valueGen, drawingGen, resizeGen, input, view, geometry: built, metrics: built.metrics, body: input.filter || input.sort ? { r0: (input.filter?.range ?? input.sort!.range).r0 + (input.filter?.range ?? input.sort!.range).headerRows, r1: (input.filter?.range ?? input.sort!.range).r1 - (input.filter?.range ?? input.sort!.range).totalsRows } : remote ? used ? { r0: used.r0, r1: used.r1 } : null : usedBody(sheet),
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
  // Value generations too: a committed pivot change re-runs find (and every owner-keyed consumer) over the new values.
  const ownerKey = JSON.stringify([globalViewGen, fontGen, sheetIds.map(revisionOf), sheetIds.map(sheetId => source.store(sheetId).stamp.valueGen)])
  const sparklineViewKey = hasResolvedSparklines ? ownerKey : ''
  const sparklineViewKeyRef = useRef(sparklineViewKey)

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
    let sparklineViews: Map<string, import('../../xlsx/view/renderViewport').SparklineViewState> | undefined
    if (hasResolvedSparklines) {
      sparklineViews = new Map()
      for (const dataSheet of doc.sheets) {
        const sheetId = dataSheet.sheetId ?? dataSheet.name
        const sheetEntry = entries.get(sheetId)
        const input = viewInput(sheetId)
        // A stale external prop cannot regress a view the workbook has already
        // committed for this source sheet.
        const committedView = sheetEntry && sheetEntry.revision >= input.revision ? sheetEntry.view : undefined
        const visibility = committedView?.visibility ?? input.visibility
        if (!visibility) continue
        let baseHidden = committedView?.baseHidden
        if (!baseHidden) {
          if (remote) { try { baseHidden = remote.geometry(sheetId).baseHidden } catch { continue } }
          else baseHidden = sheetBaseHidden(dataSheet)
        }
        sparklineViews.set(dataSheet.name.toLowerCase(), {
          order: committedView?.order ?? input.order ?? createIdentityOrder(),
          visibility,
          baseHidden,
        })
      }
    }
    return {
      sheet: entry.sheet, geometry: entry.geometry, ...(sparklineViews ? { sparklineViews } : {}), identity, colors: palette.colors, paletteKey: palette.key, origin: scroll.logicalOrigin, viewport, zoom: scroll.zoom, dpr,
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
    // A worker sparkline generation is paintable only after its lazy painter
    // has joined that generation's asset lease. Keep host invalidations queued
    // until then; otherwise an unavailable-painter tile would be cached blank.
    const sparklineAssetsReady = !hasActiveWorkerSparklines || assets?.sparklineGen === (activeSparklines?.gen ?? 0)
    if (!assets || assets.sheetId !== activeSheetId || !sparklineAssetsReady) {
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
    const changedSparkHosts = remote?.takeSparklineChanges(activeSheetId) ?? []
    if (changedSparkHosts.length && entry) {
      // Resolve paged merge extents once per changed host, outside the per-tile
      // predicate; shell sheet.mergeRanges is empty on the worker path.
      const changedBoxes = changedSparkHosts.map(host => {
        const merge = entry.geometry.merges.query({ r0: host.row, r1: host.row, c0: host.col, c1: host.col }, { limit: 1 }).items[0]
        const anchored = merge?.minRow === host.row && merge.minCol === host.col ? merge : undefined
        const r0 = anchored?.minRow ?? host.row, r1 = anchored?.maxRow ?? host.row
        const c0 = anchored?.minCol ?? host.col, c1 = anchored?.maxCol ?? host.col
        if (r1 - r0 > 10_000) return null
        const x0 = Math.min(entry.metrics.cols.start(c0), entry.metrics.cols.start(c1))
        const x1 = Math.max(entry.metrics.cols.start(c0) + entry.metrics.cols.size(c0), entry.metrics.cols.start(c1) + entry.metrics.cols.size(c1))
        const visibleRows: Array<{ top: number; bottom: number }> = []
        for (let sourceRow = r0; sourceRow <= r1; sourceRow++) {
          const slot = entry.metrics.order.sourceToDisplay(sourceRow)
          const top = entry.metrics.rows.start(slot), bottom = top + entry.metrics.rows.size(slot)
          if (bottom > top) visibleRows.push({ top, bottom })
        }
        return sparklineInvalidationBox(x0, x1, visibleRows, r1 - r0)
      })
      renderer.cache.invalidate(key => {
        if (key.identity.sheetId !== activeSheetId) return false
        const tile = tileLogicalBox(key.col, key.row, key.zoom, key.dpr)
        return changedBoxes.some(box => box === null || (box !== undefined && tile.x < box.x + box.width && box.x < tile.x + tile.width && tile.y < box.y + box.height && box.y < tile.y + tile.height))
      })
    }
    if (sparklineViewKeyRef.current !== sparklineViewKey) {
      renderer.cache.invalidate(key => key.identity.sheetId === activeSheetId)
      sparklineViewKeyRef.current = sparklineViewKey
    }
    const seenGen = structureGens.get(activeSheetId)
    if (seenGen !== undefined && seenGen !== frame.structureGen) renderer.cache.invalidate(key => key.identity.sheetId === activeSheetId)
    // A new asset lease for the same sheet (a worker sheet's drawings arrived, a pivot's charts were replaced) may paint
    // what an older one could not (the lazy chart painter, pictures): tiles painted under the older lease are stale.
    const drawingGen = entry?.drawingGen ?? 0, paintedGen = paintedDrawingGens.get(activeSheetId)
    if (paintedGen !== undefined && paintedGen !== drawingGen) renderer.cache.invalidate(key => key.identity.sheetId === activeSheetId)
    paintedDrawingGens.set(activeSheetId, drawingGen)
    const leased = presentedLease.current
    if (leased && leased.sheetId === activeSheetId && leased.lease !== assets.lease) renderer.cache.invalidate(key => key.identity.sheetId === activeSheetId)
    presentedLease.current = { sheetId: activeSheetId, lease: assets.lease }
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
  }, [identityKey, ownerKey, sparklineViewKey, activeSparklines?.gen, scroll, viewport, dpr, assets, surfaceBytes, profile, mark, palette.key, loadTick, entry?.structureGen, entry?.resizeGen])
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
  const styleForCell = (sheetEntry: SheetEntry, cell: Pick<XlsxCell, 'row' | 'col' | 'styleIndex'> & { style?: XlsxCell['style'] }): import('../../xlsx/types').XlsxCellStyle | undefined => {
    const own = remote ? remote.styleOf(cell.styleIndex) : cell.style
    return sheetEntry.sheet.regionStyles?.styleAt(cell.row, cell.col, own, 'screen') ?? own
  }
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
        if (!sheetEntry) return String(cell.value ?? '')
        const styled = cellAt(sheetEntry, cell)
        // The style paint and copy use: a styled table or pivot region resolved over the cell's own.
        const style = styleForCell(sheetEntry, { row: cell.row, col: cell.col, styleIndex: cell.styleIndex, style: styled?.style })
        return cellDisplayText(sheetEntry.sheet, { value: (styled ?? cell).value, style }, allocation)
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
  const [popup, setPopup] = useState<{ kind: 'filter' | 'sort'; col: number; invoker: HTMLElement } | { kind: 'pivot'; request: PivotUi.PivotMenuRequest; invoker: HTMLElement } | null>(null)
  const closePopup = () => setPopup(null)

  // ---- pivot controls (Task 28): a lazy chunk with the client, the button layer and the menus ----
  const [pivotUi, setPivotUi] = useState<typeof PivotUi | null>(null)
  // A paged shell's pivots arrive with the worker's display parts: the active sheet's entry says so.
  const wantsPivotUi = hasPivots || Boolean(entry?.sheet.pivotTables?.length)
  useEffect(() => {
    if (!wantsPivotUi) return
    let live = true
    import('./pivot/PivotButtonLayer').then(module => { if (live) setPivotUi(module) }, () => { /* buttons stay painted, display only */ })
    return () => { live = false }
  }, [wantsPivotUi])
  /** Each sheet's committed composite overlay, by lower-cased sheet name, as installed now (every pivot of the sheet). */
  const overlaysByName = () => {
    const out = new Map<string, ValueOverlay>()
    doc.sheets.forEach((owner, index) => { const id = sheetIds[index], overlay = id === undefined ? undefined : pivotOverlays.get(id); if (overlay) out.set(owner.name.toLocaleLowerCase(), overlay) })
    return out
  }
  const overlayValue = (cell: NonNullable<ReturnType<ValueOverlay['cell']>>) => cell.valueIsError ? formulaError(String(cell.value) as Parameters<typeof formulaError>[0]) : cell.value
  /** Task 31 resident dependents of one pivot commit (Ruling S3-AF). A change recomputes the readers of the pivot's
   * saved and shown areas against every pivot's live layout. A reset does the same with this pivot's overlay gone (its
   * cells are the sheet's own again; its saved layout answers GETPIVOTDATA) while the other pivots stay live: only
   * the charts and stale marks this pivot caused go, and no other pivot's in-flight work is dropped. When every pivot
   * is back to its saved view, the document-wide restore runs instead. */
  const installDependents = (commit: PivotUi.PivotCommit, sheet: XlsxSheet, layout: NonNullable<PivotUi.PivotCommit['layout']>, cache: NonNullable<PivotUi.PivotCommit['cache']>, changedRects: NonNullable<PivotUi.PivotCommit['changedRects']>) => {
    const contextKey = `${commit.sheetId}:${commit.pivot}`
    const generation = (pivotFormulaGens.get(contextKey) ?? 0) + 1
    pivotFormulaGens.set(contextKey, generation)
    if (commit.reset && commit.allSaved) { restoreAllDependents(); return }
    pivotFormulaContexts.set(contextKey, { sheet: sheet.name, def: commit.def, cache, layout, layoutCells: new Map(layout.cells.map(cell => [`${cell.row}:${cell.col}`, cell])), date1904: commit.date1904 ?? false, rects: changedRects, ...(commit.reset ? { saved: true as const } : {}) })
    if (commit.reset) releaseDependentsOf(contextKey)
    // A superseded commit's rectangles stay owed until a recompute of this pivot runs: one canonical set, without
    // duplicates or contained rectangles, so repeated over-budget retries stay bounded (codex r5 finding 1).
    const rects = mergeOwedRects(pendingDependentRects.get(contextKey) ?? [], changedRects)
    pendingDependentRects.set(contextKey, rects)
    const changed = rects.map(rect => ({ sheetId: commit.sheetId, rect: { minRow: rect.r0, minCol: rect.c0, maxRow: rect.r1, maxCol: rect.c1 } }))
    const lifecycle = pivotLifecycle.current
    const current = () => lifecycle === pivotLifecycle.current && pivotFormulaGens.get(contextKey) === generation
    // The viewer's pivot chunk is already loaded here; keep formula support out
    // of the eager graph by resolving its adapter from a lazy import.
    void import('../../xlsx/pivots/session/getpivotdata').then(({ getPivotData }) => {
      if (!current()) return
      pendingDependentRects.delete(contextKey)
      // The pivots as they are when this runs (a later commit of another pivot included).
      const contexts = [...pivotFormulaContexts.values()], live = contexts.filter(context => !context.saved), overlays = overlaysByName()
      const overlayEntries = live.flatMap(context => context.layout.cells.map(cell => ({ sheetName: context.sheet, col: cell.col, row: cell.row, value: cell.value.kind === 'error' ? formulaError(cell.value.value as Parameters<typeof formulaError>[0]) : cell.value.kind === 'empty' ? null : cell.value.value })))
      const result = recomputeDependents(doc, changed, {
        getPivotData: (args, evaluate, currentSheet, markStale) => getPivotData(contexts, args, evaluate, currentSheet, markStale),
        overlayEntries,
        readCellValue: (name, col, row) => {
          const lower = name.toLocaleLowerCase()
          // The sheet's composite overlay, sheet-scoped: never a same-coordinate cell of another worksheet (R2-1).
          const shown = overlays.get(lower)?.cell(row, col)
          if (shown !== undefined) return shown === null ? null : overlayValue(shown)
          for (const context of live) {
            if (context.sheet.toLocaleLowerCase() !== lower) continue
            const cell = context.layoutCells.get(`${row}:${col}`)
            if (!cell) {
              if (context.rects.some(rect => row >= rect.r0 && row <= rect.r1 && col >= rect.c0 && col <= rect.c1)) return null
              continue
            }
            if (cell.value.kind === 'empty') return null
            if (cell.value.kind === 'error') return formulaError(cell.value.value as Parameters<typeof formulaError>[0])
            return cell.value.value
          }
          return undefined
        },
      })
      // Over the budget nothing ran: the rectangles stay owed to this pivot's next recompute. Nothing was owed meanwhile
      // (the recompute is synchronous after the delete above), so this is the same canonical set, not a copy appended.
      if (result.overBudget) pendingDependentRects.set(contextKey, mergeOwedRects(pendingDependentRects.get(contextKey) ?? [], rects))
      const sheetOf = (key: string) => {
        const bang = key.lastIndexOf('!')
        return doc.sheets.findIndex(candidate => candidate.name.toLocaleLowerCase() === key.slice(0, bang))
      }
      let marksChanged = false
      // A recompute within budget ran every reader it covered against every live pivot: whichever pivot's over-budget
      // commit marked those readers out of date, they are current now (codex r5 finding 2). Readers it did not cover
      // keep their marks.
      if (!result.overBudget) for (const [owner, marks] of dependentStaleMarks) {
        for (const key of result.readers) if (marks.outdated.delete(key)) marksChanged = true
        if (!marks.unsupported.size && !marks.outdated.size) dependentStaleMarks.delete(owner)
      }
      // Stale marks are this pivot's: a change adds them; a reset recomputes against saved cells, so only an
      // over-budget reset leaves readers out of date (an unsupported reader keeps its saved value, which is right now).
      if (result.stale.length && (result.overBudget || !commit.reset)) {
        const marks = dependentStaleMarks.get(contextKey) ?? { unsupported: new Set<number>(), outdated: new Map<string, number>() }
        for (const key of result.stale) {
          const index = sheetOf(key)
          if (index < 0) continue
          if (result.overBudget) marks.outdated.set(key, index); else marks.unsupported.add(index)
          marksChanged = true
        }
        if (marks.unsupported.size || marks.outdated.size) dependentStaleMarks.set(contextKey, marks)
      }
      if (marksChanged) publishDependentNotices()
      if (result.recomputed) for (const index of result.sheets) { const id = sheetIds[index]; if (id !== undefined) dropTiles(id, null) }
      if (result.charts.length) void import('../../charts/fill').then(({ fillChartRanges, refillChartModel }) => {
        if (!current()) return
        const wanted = new Set(result.charts), changedSheets = new Set<number>(), chartOverlays = overlaysByName()
        let chartIndex = 0
        const refill = (nodes: XlsxSheet['drawings'], sheetIndex: number) => {
          for (const node of nodes ?? []) {
            if (node.content?.kind === 'xlsxChart' && node.content.chart) {
              const owners = dependentChartOwners.get(node)
              // A reset refills only the charts another live pivot still feeds; the rest show their saved models.
              if (wanted.has(chartIndex) && (!commit.reset || owners?.size)) {
                const ownerSheet = doc.sheets[sheetIndex]!
                const sourceChart = node.content.chart
                const pendingChart = { ...sourceChart,
                  categoryPending: sourceChart.categoryRef ?? sourceChart.categoryPending,
                  groups: sourceChart.groups.map(group => ({ ...group, series: group.series.map(series => ({ ...series, pending: series.refs ?? series.pending })) })),
                }
                const refilled = refillChartModel(sourceChart, () => fillChartRanges(pendingChart, ownerSheet.name, (name, row, col, needsText) => {
                  const target = doc.sheets.find(candidate => candidate.name.toLocaleLowerCase() === name.toLocaleLowerCase())
                  const shown = chartOverlays.get(name.toLocaleLowerCase())?.cell(row, col)
                  if (shown === null) return { value: null }
                  const cell = target?.rows.find(entry => entry.index === row)?.cells.find(entry => entry.col === col)
                  if (shown && target) {
                    return overlayChartCell(target, row, col, shown.value, shown.valueIsError ? String(shown.value) : undefined, needsText)
                  }
                  if (!cell) return undefined
                  const style = cell.style
                  return { value: cell.value, ...(cell.valueIsError ? { error: String(cell.value) } : {}),
                    ...(style?.formatCode ? { formatCode: style.formatCode } : {}),
                    ...(needsText ? { text: cellDisplayText(target!, { value: cell.value, style }) } : {}) }
                }))
                rememberDependentChart(savedDependentCharts, node, sourceChart)
                if (!commit.reset) { if (owners) owners.add(contextKey); else dependentChartOwners.set(node, new Set([contextKey])) }
                mainLedger.chargeWhileReachable('drawings', [refilled], estimatePayloadBytes(refilled))
                node.content.chart = refilled
                changedSheets.add(sheetIndex)
              }
              chartIndex++
            }
            if (node.children?.length) refill(node.children, sheetIndex)
          }
        }
        doc.sheets.forEach((owner, index) => refill(owner.drawings, index))
        for (const index of changedSheets) { const id = sheetIds[index]; if (id !== undefined) dropTiles(id, null) }
        if (changedSheets.size) setResidentDrawingGen(generation => generation + 1)
        rerender()
      }).catch(() => setAnnouncement('Some charts that depend on this pivot may still show saved values.'))
      rerender()
    }).catch(() => setAnnouncement('Some formulas that depend on this pivot may still show saved values.'))
  }
  /** Installs a committed pivot change (spec §7.9): 1. the sheet's live definition; 2. its region styles rebuilt;
   * 3. the overlay on the source store (resident), which 4. bumps the store's valueGen (tiles, find, copy and the
   * geometry re-key); 5. the sheet's tiles dropped. The paged path's overlay is the worker's (Task 27). */
  const installPivot = (commit: PivotUi.PivotCommit) => {
    const sheet = doc.sheets[commit.sheetId], sheetId = sheetIds[commit.sheetId]
    if (sheetId === undefined) return
    // Paged: the workbook already published the worker's commit (merges, extent, valueGen, pivot display and charts).
    if (remote) {
      const defs = [...(sheet?.pivotTables ?? [])]
      defs[commit.pivot] = commit.def
      const pivotRects = defs.flatMap(def => {
        const parts = def.location.ref.split(':')
        if (parts.length < 1 || parts.length > 2) return []
        try {
          const a = parseA1(parts[0]!.replace(/\$/g, ''))
          const b = parseA1((parts[1] ?? parts[0]!).replace(/\$/g, ''))
          return [{ minRow: Math.min(a.row, b.row), maxRow: Math.max(a.row, b.row), minCol: Math.min(a.col, b.col), maxCol: Math.max(a.col, b.col) }]
        } catch { return [] }
      })
      if (sheet) setPagedDependentNoticeSheets(current => markPagedDependentsStale(current, commit.sheetId, sheet, pivotRects, remote.hasFormulaCellsOutsidePivots ?? false))
      // §7.8 (Task 29): autofit once the worker's new pivot display (its styles) has arrived; the old styles now.
      const shownEntry = entries.get(sheetId)
      if (pivotUi && commit.fit && shownEntry) (pendingPagedFits.current ??= pivotUi.pendingPivotFits(pivotFitMemory)).add(sheetId, { pivot: commit.pivot, fit: commit.fit, styleAt: pivotStyleAt(shownEntry.sheet.regionStyles), styleGen: shownEntry.regionStyleGen })
      cancelCopy(); dropTiles(sheetId, null); rerender(); return
    }
    if (!sheet?.pivotTables) return
    // §7.8 (Task 29): the cells and pivot styles shown before the change, read before they go.
    const beforeFit = pivotUi ? { def: sheet.pivotTables[commit.pivot]!, read: pivotShownReader(sheet, pivotOverlays.get(sheetId)), styleAt: pivotStyleAt(regionStylesFor(sheet)) } : undefined
    const defs = sheet.pivotTables.map((def, index) => index === commit.pivot ? commit.def : def), saved = savedPivotDefs[commit.sheetId]
    // Every pivot of the sheet back to its saved definition (a reset): the sheet gets the saved arrays themselves, so
    // print and teardown see it as saved again (final review minor 13).
    const allSaved = saved !== undefined && defs.length === saved.length && defs.every((def, index) => def === saved[index])
    sheet.pivotTables = allSaved ? saved : defs
    // The pivot's region (gridlines hidden inside it, regionGrid) follows its live location.
    if (sheet.pivotRegions) sheet.pivotRegions = allSaved ? savedPivotRegions[commit.sheetId] : sheet.pivotRegions.map(region => region.partPath === commit.def.partPath ? { ...region, ref: commit.def.location.ref } : region)
    invalidateRegionStyles(sheet)
    regionStylesFor(sheet)
    // Task 30: the refilled pivot charts replace their saved models (charts may sit on other sheets: their tiles go too).
    if (commit.charts?.length) {
      const name = String(commit.def.name ?? ''), key = JSON.stringify([commit.def.sheet, name])
      if (!savedPivotCharts.has(key)) savedPivotCharts.set(key, { sheet: commit.def.sheet, name, charts: linkedPivotCharts(doc, commit.def.sheet, name) })
      for (const index of installPivotCharts(doc, name, commit.charts, commit.def.sheet)) { const id = sheetIds[index]; if (id !== undefined && id !== sheetId) dropTiles(id, null) }
    }
    if (!remote && commit.overlay !== undefined) {
      if (commit.overlay) pivotOverlays.set(sheetId, commit.overlay); else pivotOverlays.delete(sheetId)
      source.store(sheetId).setValueOverlay?.(commit.overlay ?? undefined)
    }
    if (pivotUi && beforeFit) autofitNotice.current = autofitPivot(sheetId, sheet, beforeFit, { def: commit.def, read: pivotShownReader(sheet, pivotOverlays.get(sheetId)), styleAt: pivotStyleAt(regionStylesFor(sheet)) })
    if (commit.layout && commit.cache && commit.changedRects?.length) installDependents(commit, sheet, commit.layout, commit.cache, commit.changedRects)
    cancelCopy()
    dropTiles(sheetId, null)
    rerender()
  }
  const pivotHost = useRef({ installPivot, restore: restorePivots, announce: setAnnouncement, openMenu: (request: PivotUi.PivotMenuRequest) => setPopup({ kind: 'pivot', request, invoker: request.invoker }), closeMenu: closePopup })
  pivotHost.current = { installPivot, restore: restorePivots, announce: message => { const note = autofitNotice.current; autofitNotice.current = undefined; setAnnouncement(note ? `${message} ${note}` : message) }, openMenu: request => setPopup({ kind: 'pivot', request, invoker: request.invoker }), closeMenu: closePopup }
  // Declared before the pivot controller's effect, so on unmount this cleanup runs first (restorePivots reads it).
  const unmounting = useRef(false)
  useEffect(() => { unmounting.current = false; return () => { unmounting.current = true } }, [])
  const pivotKinds = props.pivotVerifiedKinds
  // One controller (and client, with its sessions) per mount of these inputs; created in an effect so a StrictMode
  // re-run gets a fresh one instead of a disposed one.
  const [pivotController, setPivotController] = useState<PivotUi.PivotController | null>(null)
  useEffect(() => {
    if (!pivotUi) return
    // A resident workbook parsed on this thread has its runtime; a paged or worker-parsed one runs its sessions in
    // the worker (Task 27's RemotePivotSessions, addressed by the worker's numeric sheet id).
    const client = remote ? pivotUi.createRemotePivotClient(pivotUi.remotePivotPort(remote.pivots, {
      id: index => { const id = Number(sheetIds[index]); return Number.isSafeInteger(id) ? id : undefined },
      sheet: index => doc.sheets[index],
    }, pivotFitMemory), commit => pivotHost.current.installPivot(commit))
      : pivotUi.createResidentPivotClient({ doc, saved: savedPivotDefs, ledger: mainLedger, profile: budgetProfile, install: commit => pivotHost.current.installPivot(commit), ...(pivotKinds ? { verifiedKinds: pivotKinds } : {}) })
    const controller = new pivotUi.PivotController({ client, announce: message => pivotHost.current.announce(message), openMenu: request => pivotHost.current.openMenu(request), closeMenu: () => pivotHost.current.closeMenu() })
    setPivotController(controller)
    return () => {
      controller.dispose()
      pivotHost.current.restore()
      setPivotController(current => current === controller ? null : current)
    }
  }, [pivotUi, doc, remote, mainLedger, budgetProfile, pivotKinds])
  useEffect(() => pivotController?.subscribe(rerender), [pivotController, rerender])
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
        format: (cell, allocation) => { const style = styleForCell(entry, { ...cell, style: cellAt(entry, cell)?.style }); return cellDisplayText(entry.sheet, { value: cell.value, style }, allocation) },
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
  // Pivots on the active sheet (Task 28): load-time pivot-* diagnostics, and one notice per pivot not interactive.
  const activeIndex = sheetIds.indexOf(activeSheetId)
  const activePivots = entry?.sheet.pivotTables
  const pivotCellText = (row: number, col: number): string => { if (!entry) return ''; const cell = cellAt(entry, { row, col }); return cell ? presentCell(cell).text : '' }
  const pivotNotices = entry ? [...new Set([
    ...(entry.sheet.diagnostics ?? []).filter(d => d.feature.startsWith('pivot-')).map(d => d.message),
    ...(pivotController && activePivots ? pivotController.notices(activeIndex, activePivots.length) : []),
    ...(pagedDependentNoticeSheets.has(activeIndex) || Boolean(activeSheetId && remote?.pivotFormulaNoticeSheets.has(activeSheetId)) ? ['Formulas outside this pivot may still show saved values (paged workbook).'] : []),
    ...(residentDependentNotices.unsupported.has(activeIndex) ? ['Some formulas still show saved values because their pivot reference is unsupported.'] : []),
    // M-4: over the work budget a reader keeps what it showed, which may be an earlier change's live value.
    ...(residentDependentNotices.outdated.has(activeIndex) ? ['Some formulas that depend on a pivot may be out of date: recalculating them would take too long.'] : []),
  ])] : []

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
  const [printPivotNote, setPrintPivotNote] = useState('')
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
    // Print never runs a pivot session (§7.10): a changed pivot prints as saved, its saved definition over its saved cells.
    const changed = doc.sheets.some((sheet, index) => sheet.pivotTables !== savedPivotDefs[index])
    // Ruling S3-AL: so do its dependents. The formula cells Task 31 recomputed print their saved values, and the charts
    // it refilled their saved models, so the paper never shows a saved pivot beside live numbers that read it.
    const savedCells = savedDependentFormulaCells(doc)
    // Its linked charts print as saved too: each live chart model maps back to the saved one it replaced.
    const savedChart = new Map<ChartModel, ChartModel>()
    let dependentsSaved = Boolean(savedCells)
    for (const [node, saved] of savedDependentCharts) if (node.content?.kind === 'xlsxChart' && node.content.chart && node.content.chart !== saved) { savedChart.set(node.content.chart, saved); dependentsSaved = true }
    // A pivot chart's saved model (captured before any refill) wins over a dependent refill's snapshot of the same chart.
    for (const { sheet, name, charts } of savedPivotCharts.values()) linkedPivotCharts(doc, sheet, name).forEach((live, k) => { const saved = charts[k]; if (saved && saved !== live) savedChart.set(live, saved) })
    setPrintPivotNote(dependentsSaved ? 'Printing shows the pivot, and the values that depend on it, as saved.' : changed ? 'Printing shows the pivot as saved.' : '')
    const printed = changed || savedChart.size || savedCells ? { ...doc, sheets: doc.sheets.map((sheet, index) => {
      const drawings = savedChart.size ? withSavedCharts(sheet.drawings, savedChart) : sheet.drawings
      const rows = savedCells ? withSavedCells(sheet.rows, savedCells) : sheet.rows
      if (sheet.pivotTables === savedPivotDefs[index] && drawings === sheet.drawings && rows === sheet.rows) return sheet
      return { ...sheet, ...(drawings !== sheet.drawings ? { drawings } : {}), ...(rows !== sheet.rows ? { rows } : {}), ...(sheet.pivotTables !== savedPivotDefs[index] ? { pivotTables: savedPivotDefs[index], pivotRegions: savedPivotRegions[index], regionStyles: undefined, regionStylesResolved: false } : {}) }
    }) } : doc
    try { pages = await loader(printed, {
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
  const admitCandidate = (sheetId: string, sizes: SessionSizes, name: string, represents: (metrics: ProjectedSheetMetrics) => boolean, unrepresentable: string, quiet = false): Candidate | null => {
    const budget = MAIN_THREAD_BUDGETS[budgetProfile], current = entries.get(sheetId)!
    const refuse = () => { if (!quiet) setAnnouncement(`${name} was not resized: not enough memory to lay the sheet out again (main memory budget).`); return null }
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
    if (!represents(candidate.built.metrics)) { mainLedger.release('geometry', 'resize-candidate'); if (!quiet) setAnnouncement(`${name} was not resized: ${unrepresentable}.`); return null }
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
              body={entry.body} onChange={onGridChange} status={{ records: { ...snapshot(activeSheetId).records, completeness: loadState?.completeness ?? 'complete' }, countMode: 'body', loadReport: { status: 'ready', notices: [...notices, ...pivotNotices, ...fontNotices, ...pagedDrawingNotices, ...(activeSparklines ? sparklineNotices(activeSparklines) : []), ...chartNotices(activeCharts), ...(loadState?.failure ? [loadFailureNotice(loadState.failure, loadState.code)] : [])] } }}
              onResize={resizeLine}
              onSortCycle={(col, shift) => void sortCycle(col, shift)} pendingSortColumn={sortPending?.col ?? null} onFilter={(col, invoker) => { setPopup({ kind: 'filter', col, invoker }) }} onSortDialog={(col, invoker) => { setPopup({ kind: 'sort', col, invoker }) }} onNotice={setAnnouncement}
              popup={popup?.kind === 'pivot' ? pivotUi && pivotController ? { invoker: popup.invoker, close: closePopup, anchor: 'invoker', label: pivotMenuLabel(pivotController, popup.request), ...(popup.request.anchor ? { anchorElement: popup.request.anchor } : {}), content: <pivotUi.PivotMenu controller={pivotController} request={popup.request} owner={identity} onClose={closePopup} onNotice={setAnnouncement} /> } : undefined
                : popup && controls ? { invoker: popup.invoker, close: closePopup, label: popup.kind === 'filter' ? 'Column filter' : 'Sort worksheet', content: popup.kind === 'filter' ? <FilterDropdown owner={identity} client={controls.client} applied={entry.view.filter} appliedSort={entry.view.sort} col={popup.col} invoker={popup.invoker} embedded onClose={closePopup} onApplied={() => rerender()} onNotice={setAnnouncement} /> : <SortDialog owner={identity} client={controls.client} applied={entry.view.sort} col={popup.col} invoker={popup.invoker} embedded onClose={closePopup} onApplied={() => rerender()} onNotice={setAnnouncement} /> } : undefined}
              cellControls={pivotUi && pivotController && activePivots?.length && scroll ? <pivotUi.PivotButtonLayer controller={pivotController} sheet={activeIndex} defs={activePivots} projection={entry.metrics} window={scroll} viewport={viewport} touch={touchControls}
                cellText={pivotCellText} /> : undefined}
              onCellControl={pivotUi && pivotController && activePivots?.length ? (at, action, grid) => {
                // Ruling S3-J: the keyboard route to a cell's pivot control (focus returns to the grid).
                const found = pivotUi.controlsAt(activePivots, at.row, at.col)
                const drill = found.find(entry => entry.button.kind === 'expand' || entry.button.kind === 'collapse')
                const target = action === 'open' ? found.find(entry => entry.button.kind === 'filter' || entry.button.kind === 'page') : drill
                if (!target) return false
                if (action !== 'open' && target.button.kind !== action) { setAnnouncement(action === 'expand' ? 'Already expanded.' : 'Already collapsed.'); return true }
                const anchor = rootRef.current?.querySelector<HTMLElement>(`[data-pivot-key="${target.pivot}:${at.row}:${at.col}"]`) ?? undefined
                void pivotController.activate(activeIndex, target.pivot, target.button, grid, anchor)
                return true
              } : undefined} announcement={[sortPending ? `Sorting column ${columnToLetters(sortPending.col)}…` : '', resourceNotice, overflowNotice, findOverflowNotice === overflowNotice ? '' : findOverflowNotice, assets?.sheetId !== activeSheetId ? 'Loading sheet…' : '', assets && assets.sheetId === activeSheetId && assets.lease.refusedImages > 0 ? `${assets.lease.refusedImages} picture${assets.lease.refusedImages === 1 ? ' is' : 's are'} not shown (image memory budget).` : '', assets && assets.sheetId === activeSheetId && assets.lease.chartsUnavailable ? 'Charts are shown as placeholders: the chart engine could not be loaded.' : '', selectionIncomplete ? 'Selection highlight is incomplete at this zoom; zoom in for detail.' : '', remote?.queryNotice === announcement ? '' : remote?.queryNotice, announcement].filter(Boolean).join(' ')} inspection={inspection} profile={profile} label={`Worksheet ${entry.sheet.name}`}
              cellPresentation={at => {
                const cell = cellAt(entry, at)
                if (!cell && remote && !remote.store(activeSheetId).isKnown(at)) return { text: 'Loading…' }
                const shown = cell ? presentCell(cell) : { text: '' }
                if (!pivotUi || !pivotController || !activePivots?.length) return shown
                const expanded = pivotUi.expandedAt(activePivots, at.row, at.col)
                const control = pivotUi.cellControlDescription(pivotController, activeIndex, activePivots, at.row, at.col, pivotCellText)
                return { ...shown, ...(expanded !== undefined ? { expanded } : {}), ...(control ? { control } : {}) }
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
              {printPages && printPivotNote && <p>{printPivotNote}</p>}
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

/** A pivot menu's accessible name: its field's. */
function pivotMenuLabel(controller: PivotUi.PivotController, request: PivotUi.PivotMenuRequest): string {
  const state = controller.state(request.sheet, request.pivot)
  return (state.status === 'open' ? state.model.fields.find(field => field.index === request.field)?.name : undefined) ?? 'Pivot table field'
}

/** A pivot's linked charts in `installPivotCharts`' order (sheets, then drawing order), by the name after the last `!`. */
function linkedPivotCharts(doc: XlsxDocument, sheetName: string, name: string): ChartModel[] {
  return doc.sheets.flatMap(sheet => sheetCharts(sheet.drawings)).filter(chart => chartLinksPivot(chart, sheetName, name))
}

/** Rows with the given cells swapped for their saved copies (copies only the rows that hold one; print, Ruling S3-AL). */
function withSavedCells(rows: XlsxSheet['rows'], saved: ReadonlyMap<XlsxCell, XlsxCell>): XlsxSheet['rows'] {
  let changed = false
  const out = rows.map(row => {
    if (!row.cells.some(cell => saved.has(cell))) return row
    changed = true
    return { ...row, cells: row.cells.map(cell => saved.get(cell) ?? cell) }
  })
  return changed ? out : rows
}

/** Drawing nodes with the given chart models swapped back (copies only along the changed paths; print of a changed pivot). */
function withSavedCharts(nodes: XlsxSheet['drawings'], saved: ReadonlyMap<ChartModel, ChartModel>): XlsxSheet['drawings'] {
  if (!nodes) return nodes
  let changed = false
  const out = nodes.map(node => {
    const content = node.content?.kind === 'xlsxChart' ? saved.get(node.content.chart) : undefined
    const children = node.children?.length ? withSavedCharts(node.children, saved) : node.children
    if (!content && children === node.children) return node
    changed = true
    return { ...node, ...(content ? { content: { ...node.content!, chart: content } } : {}), ...(children !== node.children ? { children } : {}) } as typeof node
  })
  return changed ? out : nodes
}

/** The cells a resident pivot shows: its committed overlay's (value, number format and role style), else the sheet's. */
function pivotShownReader(sheet: XlsxSheet, overlay: ValueOverlay | undefined): (row: number, col: number) => { value: unknown; style?: XlsxCellStyle } | undefined {
  let rows: Map<number, XlsxSheet['rows'][number]> | undefined
  return (row, col) => {
    const shown = overlay?.cell(row, col)
    if (shown === null) return undefined
    if (shown !== undefined) {
      const own = (shown as { style?: XlsxCellStyle }).style
      const style = own ? { ...own, numFmtId: shown.numFmtId ?? 0 } : shown.numFmtId !== undefined ? { numFmtId: shown.numFmtId } : undefined
      return { value: shown.value, ...(style ? { style } : {}) }
    }
    rows ??= new Map(sheet.rows.map(entry => [entry.index, entry]))
    const cell = rows.get(row)?.cells.find(entry => entry.col === col)
    return cell ? { value: cell.value, ...(cell.style ? { style: cell.style } : {}) } : undefined
  }
}
const PIVOT_FIT_REFUSED = 'Its columns were not fitted: not enough memory to measure them or lay the sheet out again (main memory budget).'
const PIVOT_FIT_UNREPRESENTABLE = 'Its columns were not fitted: the fitted widths cannot be represented beside this sheet\'s sizes.'
/** A region style layer's resolution for screen, as a pivot autofit reads it. */
function pivotStyleAt(layer: ReturnType<typeof regionStylesFor>): ((row: number, col: number, own: XlsxCellStyle | undefined) => XlsxCellStyle | undefined) | undefined {
  return layer ? (row, col, own) => layer.styleAt(row, col, own, 'screen') ?? own : undefined
}

/** A worker reply's layout cells as a pivot autofit reads them (values and number formats; no own styles). */
function pivotCellsReader(cells: readonly { row: number; col: number; value: { value: unknown; numFmtId?: number } | null }[]): (row: number, col: number) => { value: unknown; style?: XlsxCellStyle } | undefined {
  const byCell = new Map(cells.map(cell => [`${cell.row},${cell.col}`, cell.value]))
  return (row, col) => {
    const shown = byCell.get(`${row},${col}`)
    return shown ? { value: shown.value, ...(shown.numFmtId !== undefined ? { style: { numFmtId: shown.numFmtId } } : {}) } : undefined
  }
}
