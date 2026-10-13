/**
 * Dev-only performance page for the real interactive viewer (plan S10.2).
 *
 * It mounts the shipped public component as a host would:
 *   <OfficeFile data={file} interactiveSheets viewerWorker={factory} showSearch />
 * with the workbook supplied as a LOCAL File through #perf-file (the bench sets
 * it with Playwright's setInputFiles), so the first-viewport clock starts when
 * a local seekable Blob is available (spec §12.2) and no fixture route disables
 * the browser's module cache. The viewer worker is src/debug/viewerPerfWorker.ts:
 * the shipped kernel plus a read-only `__perf` ledger/import side channel.
 *
 * Instrumentation stays at public boundaries: input events (capturing document
 * listeners installed before React mounts), DOM mutations, the viewer's
 * @internal onFrame stats and the worker port. `window.__viewerPerf` exposes
 * the drivers the bench CLI calls; every series it returns is raw samples, the
 * statistics are computed by the CLI (src/debug/harness/viewerPerf.ts).
 * Excluded from the library build (tsconfig.build.json excludes src/debug).
 */
import { useEffect, useMemo, useState, type JSX } from 'react'
import { createRoot } from 'react-dom/client'
import { OfficeFile, type ViewerWorkerPort } from '../components/OfficeFile'
import type { SheetFrameStats } from '../components/xlsx/InteractiveSheetDoc'
import type { MainLedgerSnapshot, MainMemoryLedger } from '../xlsx/view/mainLedger'
import type { PerfReply, PerfRequest } from './viewerPerfWorker'
import { classifyPresent, pairInputsToPresents, type PresentCounters } from './harness/viewerPerf'
// Statically named so the locked-down harness server serves the on-demand modules.
import '../components/largeSource'
import '../core/zip'
import '../components/xlsx/pagedWorkbook'
import InteractiveSheetDoc from '../components/xlsx/InteractiveSheetDoc'
import { GridSemantics } from '../components/xlsx/GridSemantics'
import { domClassCounts, type DomClassCounts } from '../components/xlsx/domCost'
import { createAxisMetrics } from '../xlsx/view/axes'
import { createSheetMetrics } from '../xlsx/view/projection'
import { createIdentityOrder } from '../xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../xlsx/view/visibility'
import { createScrollWindow } from '../xlsx/view/scrollWindow'
import { emptyGridSelection, selectCell } from '../xlsx/view/selection'

export interface ScrollRunOptions { warmup: number; frames: number; dx: number; dy: number }
export interface ScrollRunResult {
  /** rAF begin-frame deltas over the measured frames (ms). */
  frameIntervals: number[]
  /** Native scroll event timeStamp -> first base present after it (ms). */
  inputToPaint: number[]
  /** Scroll step applied (in rAF) -> first base present after it (ms), informational. */
  stepToPaint: number[]
  scrollEvents: number
  presents: number
  /** Measured scroll events that shared a present with a later event. */
  coalesced: number
  longTasks: number[]
  /** Max detail records any measured present painted. */
  maxFrameRecords: number
  pendingPresents: number
  /** Raw timelines relative to the first measured frame (ms), retained in the receipt. */
  timeline: { frames: number[]; steps: number[]; events: number[]; presents: number[] }
}
export interface AckRecord { inputAt: number; inputType: string; mutationAt: number | null; presentAt: number | null; what: string | null; statusBefore: string; statusAfter: string | null }
export interface FirstViewport {
  blobAt: number | null
  firstPresentAt: number | null
  firstUsableAt: number | null
  statusAt: number | null
  statusText: string | null
  /** Status first reports the complete body (the viewer's import finished). */
  completeAt: number | null
  /** Status text when the first usable present happened. */
  usableStatus: string | null
}
export interface ViewerPerf {
  booted: boolean
  error: string | null
  fileName: string | null
  fileBytes: number
  marks: FirstViewport
  lastFrame: SheetFrameStats | null
  peakBitmapBytes: number
  workersCreated: number
  workersTerminated: number
  commands: Record<string, number>
  /** Worker command/reply timeline (post and terminal-reply times by requestId). */
  wire: WireRecord[]
  close(): Promise<void>
  scrollRun(options: ScrollRunOptions): Promise<ScrollRunResult>
  /** Starts/stops recording native wheel events and base presents (real-input diagnostic). */
  wheelRecord(start: boolean): { wheel: number[]; presents: number[] }
  /** Arms one acknowledgment probe (listeners are installed when this returns). */
  ackArm(kind: AckKind, timeoutMs: number): void
  /** Resolves when the armed probe saw a qualifying mutation presented, or on timeout. */
  ackWait(): Promise<AckRecord>
  /** DOM cost calibration (S10.2 ledger estimate): mounts `cells` owned-gridcell-like
   * React elements (div + text span, as GridSemantics renders them) or unmounts them. */
  calibrateDom(cells: number, updates?: number): Promise<number>
  /** Viewer-representative DOM calibration: mounts the real GridSemantics tree
   * (cells, header select/sort/filter/more controls, status and live regions)
   * for a `width` x `height` viewport, re-renders it `updates` times with a
   * moved selection, and returns its element counts by class; 0 x 0 unmounts. */
  calibrateGrid(width: number, height: number, updates?: number): Promise<DomClassCounts>
  /** Real-viewer calibration (scripts/calibrate-viewer-tiles.ts): mounts the
   * actual InteractiveSheetDoc (resident, one value at A1) in a `width` x
   * `height` box and scrolls it `updates` steps of 600 px (negative: keep the
   * mounted viewer and scroll further); returns its DOM class counts, main
   * ledger and retained tile count. 0 x 0 unmounts. */
  calibrateViewer(width: number, height: number, updates?: number): Promise<{ counts: DomClassCounts; ledger: MainLedgerSnapshot | null; tiles: number | null }>
  /** The open document's main-thread application ledger (null before a viewer mounts). */
  mainLedger(): MainLedgerSnapshot | null
  /** Drops the page's own instrumentation (the wire timeline) before a memory
   * checkpoint, so harness records are not read as viewer memory; returns how
   * many records were dropped. */
  dropInstrumentation(): { wire: number }
  /** Layout-shift entries the page has buffered so far (null: API missing). */
  layoutShifts(): number | null
  /** Their attribution sources (each holds two DOMRectReadOnly). */
  layoutShiftSources(): number | null
  /** Resolves true once the viewer has presented nothing for `idleMs` and no
   * worker request is open (false after `timeoutMs`): memory checkpoints are
   * taken at rest, not while prefetch or row loads still allocate. */
  quiet(idleMs: number, timeoutMs: number): Promise<boolean>
  /** Holds worker replies of these command kinds until called with null, which
   * releases every held reply in order (test hook). */
  holdReplies(kinds: string[] | null): number
  workerStats(): Promise<PerfReply | null>
  workerGc(): Promise<boolean>
  /** Dedicated importer pass over #perf-import-file in a fresh perf worker (the viewer is not involved). */
  importRun(profile: 'D' | 'M'): Promise<PerfReply>
  canvasBytes(): { count: number; bytes: number }
  storageEstimate(): Promise<{ usage: number | null; quota: number | null }>
  measureMemory(): Promise<unknown>
}
export interface WireRecord { requestId: string; kind: string; op?: string; detail?: string; postedAt: number; progress: number; repliedAt: number | null; reply: string | null }
export type AckKind = 'status' | 'dialog-open' | 'dialog-close' | 'apply' | 'find'
declare global { interface Window { __viewerPerf?: ViewerPerf } }

const raf = () => new Promise<number>(resolve => requestAnimationFrame(resolve))

// ---- capturing input/scroll recorders: installed before React mounts ----
let scrollLog: number[] | null = null
document.addEventListener('scroll', event => { scrollLog?.push(event.timeStamp) }, { capture: true, passive: true })
let presentLog: { at: number; pending: number; records: number }[] | null = null
let wheelLog: number[] | null = null
document.addEventListener('wheel', event => { wheelLog?.push(event.timeStamp) }, { capture: true, passive: true })
let lastCounters: PresentCounters | null = null
// Layout-shift entries the page's performance timeline buffers (capped at 150
// by spec; page-owned, not the viewer's): counted so the memory gate can
// exclude exactly their measured growth. null where the API is missing.
let layoutShiftEntries: number | null = null
let layoutShiftSources = 0
let openRequests: Map<string, WireRecord> | null = null
let lastFrameAt = 0
// Reply kinds currently held back from the viewer, and the held deliveries.
let heldKinds: Set<string> | null = null
const held: (() => void)[] = []
const LAYOUT_SHIFT_BUFFER = 150
try {
  new PerformanceObserver(list => {
    for (const entry of list.getEntries() as unknown as { sources?: readonly unknown[] }[]) {
      // The buffer keeps the first LAYOUT_SHIFT_BUFFER entries only.
      if ((layoutShiftEntries ?? 0) >= LAYOUT_SHIFT_BUFFER) break
      layoutShiftEntries = (layoutShiftEntries ?? 0) + 1
      layoutShiftSources += entry.sources?.length ?? 0
    }
  }).observe({ type: 'layout-shift', buffered: true })
  layoutShiftEntries = 0
} catch { layoutShiftEntries = null }

const state: ViewerPerf = {
  booted: false,
  error: null,
  fileName: null,
  fileBytes: 0,
  marks: { blobAt: null, firstPresentAt: null, firstUsableAt: null, statusAt: null, statusText: null, completeAt: null, usableStatus: null },
  lastFrame: null,
  peakBitmapBytes: 0,
  workersCreated: 0,
  workersTerminated: 0,
  commands: {},
  wire: [],
  close: async () => {},
  scrollRun,
  wheelRecord(start) {
    if (start) { wheelLog = []; presentLog = []; return { wheel: [], presents: [] } }
    const out = { wheel: wheelLog ?? [], presents: (presentLog ?? []).map(p => p.at) }
    wheelLog = null; presentLog = null
    return out
  },
  ackArm(kind, timeoutMs) { pendingAck = ackProbe(kind, timeoutMs) },
  ackWait: () => pendingAck ?? Promise.reject(new Error('no acknowledgment probe armed')),
  mainLedger: () => currentLedger?.snapshot() ?? null,
  dropInstrumentation() { const wire = state.wire.length; state.wire = []; return { wire } },
  layoutShifts: () => layoutShiftEntries,
  layoutShiftSources: () => layoutShiftEntries === null ? null : layoutShiftSources,
  holdReplies(kinds) {
    if (kinds) { heldKinds = new Set(kinds); return 0 }
    heldKinds = null
    const release = held.splice(0)
    for (const deliver of release) deliver()
    return release.length
  },
  quiet(idleMs, timeoutMs) {
    const started = performance.now()
    return new Promise(resolve => {
      const check = () => {
        const now = performance.now()
        if (now - lastFrameAt >= idleMs && (openRequests?.size ?? 0) === 0) { resolve(true); return }
        if (now - started > timeoutMs) { resolve(false); return }
        setTimeout(check, 50)
      }
      check()
    })
  },
  calibrateDom,
  calibrateGrid,
  calibrateViewer,
  workerStats: () => probe(currentWorker, { __perf: 'stats' }),
  workerGc: async () => { const reply = await probe(currentWorker, { __perf: 'gc' }); return reply?.kind === 'gc' && reply.collected },
  importRun,
  canvasBytes() {
    let bytes = 0
    const canvases = [...document.querySelectorAll('canvas')]
    for (const canvas of canvases) bytes += canvas.width * canvas.height * 4
    return { count: canvases.length, bytes }
  },
  async storageEstimate() {
    try {
      const estimate = await navigator.storage.estimate()
      return { usage: estimate.usage ?? null, quota: estimate.quota ?? null }
    } catch { return { usage: null, quota: null } }
  },
  async measureMemory() {
    const api = (performance as unknown as { measureUserAgentSpecificMemory?: () => Promise<unknown> }).measureUserAgentSpecificMemory
    if (!globalThis.crossOriginIsolated) return { unavailable: 'page is not crossOriginIsolated' }
    if (typeof api !== 'function') return { unavailable: 'performance.measureUserAgentSpecificMemory is not implemented' }
    // Chrome may defer the result to a later GC; a bounded wait keeps a hung
    // probe labelled unavailable instead of stalling the run.
    try {
      return await Promise.race([api.call(performance), new Promise(resolve => setTimeout(() => resolve({ unavailable: 'measureUserAgentSpecificMemory did not resolve within 30 s' }), 30_000))])
    } catch (error) { return { unavailable: error instanceof Error ? error.message : String(error) } }
  },
}
window.__viewerPerf = state

let currentLedger: MainMemoryLedger | null = null
let calibration: { root: ReturnType<typeof createRoot>; host: HTMLElement } | null = null
function calibrateDom(cells: number, updates = 0): Promise<number> {
  calibration?.root.unmount(); calibration?.host.remove(); calibration = null
  if (cells <= 0) return Promise.resolve(0)
  const host = document.createElement('div')
  host.style.cssText = 'position:absolute;left:0;top:0;width:1px;height:1px;overflow:hidden'
  document.body.append(host)
  const root = createRoot(host)
  const hidden = { position: 'absolute', width: 1, height: 1, overflow: 'hidden', clipPath: 'inset(50%)', whiteSpace: 'nowrap' } as const
  const tree = (pass: number) => <div role="rowgroup">{Array.from({ length: cells }, (_, i) => (
    <div key={i} role="gridcell" id={`cal-cell-${i * 16384 + 3}`} aria-colindex={(i % 40) + 2} aria-label={`C${i}, value ${i * 7.25}`} aria-selected={pass % 2 === 1} aria-description="value"
      style={{ position: 'absolute', top: 0, left: 44 + (i % 40) * 64, width: 64, height: 20 }}><span style={hidden}>{String(i * 7.25)}</span></div>
  ))}</div>
  root.render(tree(0))
  calibration = { root, host }
  // `updates` re-renders change every cell's props, as scrolling and selection
  // do in the viewer: React then keeps an alternate fiber per updated element.
  return new Promise(resolve => {
    let pass = 0
    const step = () => {
      if (pass >= updates) { setTimeout(() => resolve(host.getElementsByTagName('*').length), 100); return }
      pass++
      root.render(tree(pass))
      setTimeout(step, 50)
    }
    setTimeout(step, 50)
  })
}
let gridCalibration: { root: ReturnType<typeof createRoot>; host: HTMLElement } | null = null
function calibrateGrid(width: number, height: number, updates = 0): Promise<DomClassCounts> {
  gridCalibration?.root.unmount(); gridCalibration?.host.remove(); gridCalibration = null
  if (width <= 0 || height <= 0) return Promise.resolve({ cell: 0, header: 0, other: 0 })
  const host = document.createElement('div')
  host.style.cssText = `position:absolute;left:0;top:0;width:${width}px;height:${height}px;overflow:hidden`
  document.body.append(host)
  const root = createRoot(host)
  const projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 1048576, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: 16384, defaultSizeCss: 64 }), order: createIdentityOrder(), visibility: visibilityFromHidden(createIntervalSet([])) })
  const viewport = { width, height }
  const window = createScrollWindow(projection, viewport)
  const props = {
    owner: { docId: 'calibration', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }, projection,
    view: { baseHidden: createIntervalSet([]), visibility: visibilityFromHidden(createIntervalSet([])), filter: null, sort: null },
    window, viewport, body: { r0: 0, r1: 1048575 }, onChange: () => {},
    cellPresentation: (at: { row: number; col: number }) => ({ text: String(at.row * 7.25 + at.col) }),
    onSortCycle: () => {}, onFilter: () => {}, onSortDialog: () => {}, onNotice: () => {}, pendingSortColumn: null,
    status: { records: { visible: 1048575, complete: 1048575, completeness: 'complete' as const }, countMode: 'body' as const, loadReport: { status: 'ready' as const, notices: ['Showing saved values.'] } },
    announcement: '', label: 'Worksheet Calibration', profile: 'D' as const,
  }
  const tree = (pass: number) => <GridSemantics {...(props as unknown as Parameters<typeof GridSemantics>[0])} selection={selectCell(emptyGridSelection(), projection, { row: pass, col: pass % 3 })} />
  root.render(tree(0))
  gridCalibration = { root, host }
  return new Promise(resolve => {
    let pass = 0
    const step = () => {
      if (pass >= updates) { setTimeout(() => resolve(domClassCounts(host)), 100); return }
      pass++
      root.render(tree(pass))
      setTimeout(step, 50)
    }
    setTimeout(step, 50)
  })
}
let viewerCalibration: { root: ReturnType<typeof createRoot>; host: HTMLElement; ledger: MainLedgerSnapshot | null; read: () => MainLedgerSnapshot | null; tiles: number | null; live: MainMemoryLedger | null } | null = null
const calibrationDoc = (() => {
  // One populated cell: every other tile draws gridlines only, so tile cost
  // is isolated from the shaping cache (no text to shape).
  const rows = [{ index: 0, cells: [{ ref: 'A1', row: 0, col: 0, styleIndex: 0, value: 'x', hasCachedValue: true }] }]
  return { sheets: [{ name: 'Calibration', sheetId: 'cal', rows, cols: [], merges: [], mergeRanges: [] }], tables: [] }
})()
function calibrateViewer(width: number, height: number, updates = 0): Promise<{ counts: DomClassCounts; ledger: MainLedgerSnapshot | null; tiles: number | null }> {
  // updates < 0: keep the mounted viewer and scroll it -updates more steps.
  if (updates >= 0) { viewerCalibration?.root.unmount(); viewerCalibration?.host.remove(); viewerCalibration = null }
  if (updates >= 0 && (width <= 0 || height <= 0)) return Promise.resolve({ counts: { cell: 0, header: 0, other: 0 }, ledger: null, tiles: null })
  if (!viewerCalibration) {
    const host = document.createElement('div')
    host.style.cssText = `position:absolute;left:0;top:0;width:${width}px;height:${height}px;display:flex;flex-direction:column`
    document.body.append(host)
    const root = createRoot(host)
    const cal = { root, host, ledger: null as MainLedgerSnapshot | null, read: () => null as MainLedgerSnapshot | null, tiles: null as number | null, live: null as MainLedger | null }
    const internal = { onMainLedger: (l: MainLedger) => { cal.live = l }, onFrame: (stats: SheetFrameStats) => { cal.tiles = stats.tiles.tiles } } as Record<string, unknown>
    root.render(<InteractiveSheetDoc doc={calibrationDoc as never} style={{ flex: '1 1 auto', minHeight: 0 }} {...internal} />)
    cal.read = () => cal.live?.snapshot() ?? null
    viewerCalibration = cal
  }
  const cal = viewerCalibration
  const steps = Math.abs(updates)
  return new Promise(resolve => {
    let pass = 0
    const step = () => {
      const scroller = cal.host.querySelector<HTMLElement>('[data-testid=officeview-sheet-scroller]')
      if (pass >= steps || !scroller) { setTimeout(() => resolve({ counts: domClassCounts(cal.host.firstElementChild ?? cal.host), ledger: cal.read(), tiles: cal.tiles }), 800); return }
      pass++
      scroller.scrollTop += 600
      setTimeout(step, 250)
    }
    setTimeout(step, updates >= 0 ? 1500 : 50)
  })
}
type MainLedger = MainMemoryLedger
// ---- worker port: the shipped kernel behind a counting port ----
let currentWorker: Worker | null = null
let perfIds = 0
function probe(worker: Worker | null, request: { __perf: PerfRequest['__perf'] } & Record<string, unknown>, transfer: Transferable[] = []): Promise<PerfReply | null> {
  if (!worker) return Promise.resolve(null)
  const id = ++perfIds
  return new Promise(resolve => {
    const done = (reply: PerfReply) => {
      worker.removeEventListener('message', listener)
      worker.removeEventListener('error', onError)
      clearTimeout(timer)
      resolve(reply)
    }
    const listener = (event: MessageEvent) => {
      const data = event.data as { __perfReply?: number }
      if (data?.__perfReply === id) done(event.data as PerfReply)
    }
    // A worker that fails to load or throws never replies: report it, never hang.
    const onError = (event: ErrorEvent) => done({ __perfReply: id, kind: 'error', message: `worker error: ${event.message || 'failed to load'}` })
    const limitMs = request.__perf === 'import' ? 600_000 : 30_000
    const timer = setTimeout(() => done({ __perfReply: id, kind: 'error', message: `worker probe timed out after ${limitMs / 1000} s` }), limitMs)
    worker.addEventListener('message', listener)
    worker.addEventListener('error', onError)
    worker.postMessage({ ...request, id }, transfer)
  })
}
function perfWorker(): Worker {
  return new Worker(new URL('./viewerPerfWorker.ts', import.meta.url), { type: 'module' })
}
function countingWorker(): ViewerWorkerPort {
  const worker = perfWorker()
  currentWorker = worker
  state.workersCreated++
  const open = new Map<string, WireRecord>()
  openRequests = open
  const requestKinds = new Map<string, string>()
  const gatedListeners = new Map<unknown, (event: MessageEvent) => void>()
  worker.addEventListener('message', event => {
    const data = event.data as { requestId?: unknown; kind?: unknown }
    if (typeof data?.requestId !== 'string') return
    const record = open.get(data.requestId)
    if (!record) return
    if (data.kind === 'progress') { record.progress++; return }
    record.repliedAt = performance.now()
    record.reply = String(data.kind)
    open.delete(data.requestId)
  })
  return {
    postMessage(message: unknown, transfer: Transferable[]) {
      const command = message as { kind?: string; requestId?: unknown }
      const kind = String(command.kind)
      state.commands[kind] = (state.commands[kind] ?? 0) + 1
      if (typeof command.requestId === 'string') requestKinds.set(command.requestId, kind)
      if (typeof command.requestId === 'string' && state.wire.length < 20_000) {
        const query = (command as { query?: { op?: string; request?: { limit?: number; range?: { r0: number; r1: number; c0: number; c1: number } }; range?: { r0: number; r1: number; c0: number; c1: number } } }).query
        const rect = query?.request?.range ?? query?.range
        const record: WireRecord = {
          requestId: command.requestId, kind, postedAt: performance.now(), progress: 0, repliedAt: null, reply: null,
          ...(query?.op ? { op: query.op } : {}),
          ...(rect ? { detail: `${rect.r1 - rect.r0 + 1}x${rect.c1 - rect.c0 + 1}${query?.request?.limit !== undefined ? ` limit ${query.request.limit}` : ''}` } : {}),
        }
        state.wire.push(record)
        open.set(command.requestId, record)
      }
      worker.postMessage(message, transfer)
    },
    // Message listeners go through a holding gate (holdReplies): a held reply
    // kind reaches the viewer only when released (layout-stability test hook).
    addEventListener: (type, listener) => {
      if (type !== 'message') { worker.addEventListener(type, listener); return }
      const gated = (event: MessageEvent) => {
        const data = event.data as { requestId?: unknown; kind?: unknown }
        const kind = typeof data?.requestId === 'string' ? requestKinds.get(data.requestId) : undefined
        if (heldKinds && kind && heldKinds.has(kind)) { held.push(() => (listener as (event: MessageEvent) => void)(event)); return }
        ;(listener as (event: MessageEvent) => void)(event)
      }
      gatedListeners.set(listener, gated)
      worker.addEventListener('message', gated)
    },
    removeEventListener: (type, listener) => {
      if (type !== 'message') { worker.removeEventListener(type, listener); return }
      const gated = gatedListeners.get(listener)
      if (gated) { worker.removeEventListener('message', gated); gatedListeners.delete(listener) }
    },
    terminate() {
      state.workersTerminated++
      if (currentWorker === worker) currentWorker = null
      worker.terminate()
    },
  }
}

async function importRun(profile: 'D' | 'M'): Promise<PerfReply> {
  const file = (document.getElementById('perf-import-file') as HTMLInputElement).files?.[0]
  if (!file) return { __perfReply: 0, kind: 'error', message: 'no #perf-import-file selected' }
  const worker = perfWorker()
  try {
    const reply = await probe(worker, { __perf: 'import', blob: file, profile })
    return reply ?? { __perfReply: 0, kind: 'error', message: 'no reply' }
  } finally {
    worker.terminate()
  }
}

// ---- frames: present detection and first usable viewport ----
function onFrame(stats: SheetFrameStats): void {
  state.lastFrame = stats
  lastFrameAt = performance.now()
  state.peakBitmapBytes = Math.max(state.peakBitmapBytes, stats.peakBitmapBytes)
  const counters: PresentCounters = { basePaints: stats.basePaints, tileHits: stats.tileHits, prefetched: stats.prefetched, overlayPaints: stats.overlayPaints }
  const present = classifyPresent(lastCounters, counters)
  lastCounters = counters
  if (!present) return
  const at = performance.now()
  presentLog?.push({ at, pending: stats.pendingTiles, records: stats.frameRecords })
  const marks = state.marks
  if (marks.blobAt === null) return
  marks.firstPresentAt ??= at
  if (marks.firstUsableAt === null && stats.pendingTiles === 0 && stats.frameRecords > 0) {
    marks.firstUsableAt = at
    marks.usableStatus = statusText().slice(0, 300)
  }
}
const statusText = () => document.querySelector('[data-officeview-ui=viewer-status]')?.textContent ?? ''
const statusObserver = new MutationObserver(() => {
  const marks = state.marks
  if (marks.blobAt === null || marks.completeAt !== null) return
  const text = statusText()
  if (marks.statusAt === null && /Visible body|records/.test(text)) { marks.statusAt = performance.now(); marks.statusText = text.slice(0, 300) }
  // Worker-backed sheets say "complete body N" once imported; resident (object)
  // documents are complete when first shown and count records instead.
  if (/Visible body \d+ \/ (complete )?body [1-9]|\d+ of [1-9]\d* records/.test(text) && !/Loading sheet|\(loading\)/.test(text)) marks.completeAt = performance.now()
})
statusObserver.observe(document.getElementById('root')!, { subtree: true, childList: true, characterData: true })

// ---- sustained scroll: open loop, one step per animation frame ----
async function scrollRun(options: ScrollRunOptions): Promise<ScrollRunResult> {
  const el = document.querySelector('[data-testid=officeview-sheet-scroller]') as HTMLElement | null
  if (!el) throw new Error('viewer scroller not mounted')
  const longTasks: number[] = []
  let observer: PerformanceObserver | null = null
  try {
    observer = new PerformanceObserver(list => { for (const entry of list.getEntries()) longTasks.push(entry.duration) })
    observer.observe({ type: 'longtask', buffered: false })
  } catch { observer = null }
  const frameTimes: number[] = []
  const steps: number[] = []
  let measuring = false
  let dirY = 1, dirX = 1
  const startScroll: number[] = [], startPresent: { at: number; pending: number; records: number }[] = []
  scrollLog = startScroll
  presentLog = startPresent
  let measuredFrom = 0
  for (let k = 0; k < options.warmup + options.frames + 1; k++) {
    const ts = await raf()
    if (k === options.warmup) { measuring = true; measuredFrom = performance.now(); longTasks.length = 0 }
    if (measuring) frameTimes.push(ts)
    if (k === options.warmup + options.frames) break
    // Bounce at the scroller's physical edges so every step moves the view.
    if (options.dy) {
      if (el.scrollTop + el.clientHeight + options.dy >= el.scrollHeight) dirY = -1
      else if (el.scrollTop - options.dy <= 0) dirY = 1
      el.scrollTop += dirY * options.dy
    }
    if (options.dx) {
      if (el.scrollLeft + el.clientWidth + options.dx >= el.scrollWidth) dirX = -1
      else if (el.scrollLeft - options.dx <= 0) dirX = 1
      el.scrollLeft += dirX * options.dx
    }
    if (measuring) steps.push(performance.now())
  }
  // Trailing presents for the last steps.
  await new Promise(resolve => setTimeout(resolve, 250))
  scrollLog = null
  presentLog = null
  observer?.disconnect()
  const events = startScroll.filter(t => t >= measuredFrom)
  const presents = startPresent.filter(p => p.at >= measuredFrom)
  const paired = pairInputsToPresents(events, presents.map(p => p.at))
  const stepPaired = pairInputsToPresents(steps, presents.map(p => p.at))
  const frameIntervals: number[] = []
  for (let i = 1; i < frameTimes.length; i++) frameIntervals.push(frameTimes[i] - frameTimes[i - 1])
  return {
    frameIntervals,
    inputToPaint: paired.latencies,
    stepToPaint: stepPaired.latencies,
    scrollEvents: events.length,
    presents: presents.length,
    coalesced: paired.coalesced,
    longTasks,
    maxFrameRecords: presents.reduce((max, p) => Math.max(max, p.records), 0),
    pendingPresents: presents.filter(p => p.pending > 0).length,
    timeline: {
      frames: frameTimes.map(t => t - measuredFrom),
      steps: steps.map(t => t - measuredFrom),
      events: events.map(t => t - measuredFrom),
      presents: presents.map(p => p.at - measuredFrom),
    },
  }
}

// ---- input acknowledgment: input event -> qualifying DOM mutation -> next frame ----
function describe(mutation: MutationRecord): string {
  const target = mutation.target instanceof Element ? mutation.target : mutation.target.parentElement
  const label = target ? `${target.tagName.toLowerCase()}${target.getAttribute('role') ? `[role=${target.getAttribute('role')}]` : ''}${target.getAttribute('data-officeview-ui') ? `[ui=${target.getAttribute('data-officeview-ui')}]` : ''}` : '?'
  return mutation.type === 'attributes' ? `${label} @${mutation.attributeName}` : `${label} ${mutation.type}`
}
function removesDialog(mutation: MutationRecord): boolean {
  return mutation.type === 'childList' && [...mutation.removedNodes].some(node => node instanceof Element && (node.matches('[role=dialog]') || node.querySelector('[role=dialog]') !== null))
}
function qualifies(kind: AckKind, mutation: MutationRecord): boolean {
  const el = mutation.target instanceof Element ? mutation.target : mutation.target.parentElement
  if (!el) return false
  if (mutation.type === 'attributes' && (mutation.attributeName === 'aria-busy' || mutation.attributeName === 'aria-sort' || mutation.attributeName === 'aria-disabled')) return true
  const inStatus = el.closest('[role=status],[role=progressbar],[role=alert]') !== null
  switch (kind) {
    case 'status': return inStatus || el.closest('[role=progressbar]') !== null
    case 'dialog-open': return mutation.type === 'childList' && [...mutation.addedNodes].some(node => node instanceof Element && (node.matches('[role=dialog]') || node.querySelector('[role=dialog]') !== null))
    case 'dialog-close': return removesDialog(mutation)
    // Applying from a dialog: the dialog closing, or a status/aria change, acknowledges it.
    case 'apply': return removesDialog(mutation) || inStatus
    case 'find': return inStatus || (el.closest('[data-officeview-ui]') !== null && mutation.type !== 'attributes')
  }
}
let pendingAck: Promise<AckRecord> | null = null
function ackProbe(kind: AckKind, timeoutMs: number): Promise<AckRecord> {
  return new Promise(resolve => {
    const record: AckRecord = { inputAt: -1, inputType: '', mutationAt: null, presentAt: null, what: null, statusBefore: statusText().slice(0, 300), statusAfter: null }
    const types = ['pointerdown', 'mousedown', 'keydown', 'input'] as const
    const onInput = (event: Event) => { if (record.inputAt < 0) { record.inputAt = event.timeStamp; record.inputType = event.type } }
    for (const type of types) document.addEventListener(type, onInput, { capture: true })
    const root = document.getElementById('root')!
    // The find box echoes typed text without a DOM mutation: its frame is the
    // acknowledgment, so the input event itself qualifies for 'find'.
    const echo = (event: Event) => { if (kind === 'find' && record.mutationAt === null && record.inputAt >= 0) { record.mutationAt = performance.now(); record.what = `${event.type} echo`; finishSoon() } }
    document.addEventListener('input', echo, { capture: false })
    const observer = new MutationObserver(mutations => {
      if (record.inputAt < 0 || record.mutationAt !== null) return
      const hit = mutations.find(mutation => qualifies(kind, mutation))
      if (!hit) return
      record.mutationAt = performance.now()
      record.what = describe(hit)
      record.statusAfter = statusText().slice(0, 300)
      finishSoon()
    })
    observer.observe(root, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ['aria-busy', 'aria-sort', 'aria-disabled', 'hidden'] })
    let done = false
    const finish = () => {
      if (done) return
      done = true
      clearTimeout(timer)
      observer.disconnect()
      for (const type of types) document.removeEventListener(type, onInput, { capture: true })
      document.removeEventListener('input', echo)
      resolve(record)
    }
    function finishSoon() {
      // The acknowledgment is visible once the frame holding the mutation is produced.
      void raf().then(() => { record.presentAt = performance.now(); finish() })
    }
    const timer = setTimeout(finish, timeoutMs)
  })
}

// ---- page ----
function PerfPage(): JSX.Element {
  const [file, setFile] = useState<File | null>(null)
  useEffect(() => {
    const input = document.getElementById('perf-file') as HTMLInputElement
    const onChange = () => {
      const picked = input.files?.[0] ?? null
      if (!picked) return
      // The first-viewport clock: a local seekable Blob is now available.
      state.marks = { blobAt: performance.now(), firstPresentAt: null, firstUsableAt: null, statusAt: null, statusText: null, completeAt: null, usableStatus: null }
      state.fileName = picked.name
      state.fileBytes = picked.size
      state.peakBitmapBytes = 0
      lastCounters = null
      setFile(picked)
    }
    input.addEventListener('change', onChange)
    state.close = () => new Promise<void>(resolve => {
      const terminated = state.workersTerminated
      const hadWorker = currentWorker !== null
      setFile(null)
      input.value = ''
      // Instrumentation residue must not count as viewer memory after close.
      state.wire = []
      state.lastFrame = null
      const started = performance.now()
      const wait = () => {
        // The page then stays idle: no re-render and no input after the close,
        // so collection measures what the unmounted viewer leaves behind.
        if (!hadWorker || state.workersTerminated > terminated || performance.now() - started > 5000) resolve()
        else setTimeout(wait, 20)
      }
      setTimeout(wait, 0)
    })
    state.booted = true
    return () => input.removeEventListener('change', onChange)
  }, [])
  const factory = useMemo(() => countingWorker, [])
  const internal = { onFrame, onMainLedger: (ledger: MainMemoryLedger) => { currentLedger = ledger } } as Record<string, unknown>
  if (!file) return <div data-testid="perf-idle">No workbook</div>
  return (
    <OfficeFile
      data={file}
      interactiveSheets
      showSearch
      viewerWorker={factory}
      loading={<div role="status">Loading workbook…</div>}
      error={message => { state.error = message; return <div role="alert" data-testid="perf-error">{message}</div> }}
      style={{ flex: '1 1 auto', minHeight: 0 }}
      {...internal}
    />
  )
}

const container = document.getElementById('root')
if (!container) throw new Error('perf page requires #root')
createRoot(container).render(<PerfPage />)
