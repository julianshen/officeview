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
import type { PerfReply, PerfRequest } from './viewerPerfWorker'
import { classifyPresent, pairInputsToPresents, type PresentCounters } from './harness/viewerPerf'
// Statically named so the locked-down harness server serves the on-demand modules.
import '../components/largeSource'
import '../core/zip'
import '../components/xlsx/pagedWorkbook'
import '../components/xlsx/InteractiveSheetDoc'

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
  workerStats(): Promise<PerfReply | null>
  workerGc(): Promise<boolean>
  /** Dedicated importer pass over #perf-import-file in a fresh perf worker (the viewer is not involved). */
  importRun(profile: 'D' | 'M'): Promise<PerfReply>
  canvasBytes(): { count: number; bytes: number }
  storageEstimate(): Promise<{ usage: number | null; quota: number | null }>
  measureMemory(): Promise<unknown>
}
export interface WireRecord { requestId: string; kind: string; op?: string; detail?: string; postedAt: number; progress: number; repliedAt: number | null; reply: string | null }
export type AckKind = 'status' | 'dialog-open' | 'dialog-close' | 'find'
declare global { interface Window { __viewerPerf?: ViewerPerf } }

const raf = () => new Promise<number>(resolve => requestAnimationFrame(resolve))

// ---- capturing input/scroll recorders: installed before React mounts ----
let scrollLog: number[] | null = null
document.addEventListener('scroll', event => { scrollLog?.push(event.timeStamp) }, { capture: true, passive: true })
let presentLog: { at: number; pending: number; records: number }[] | null = null
let wheelLog: number[] | null = null
document.addEventListener('wheel', event => { wheelLog?.push(event.timeStamp) }, { capture: true, passive: true })
let lastCounters: PresentCounters | null = null

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
    try { return await api.call(performance) } catch (error) { return { unavailable: error instanceof Error ? error.message : String(error) } }
  },
}
window.__viewerPerf = state

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
    const timer = setTimeout(() => done({ __perfReply: id, kind: 'error', message: 'worker probe timed out after 600 s' }), 600_000)
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
    addEventListener: (type, listener) => worker.addEventListener(type, listener),
    removeEventListener: (type, listener) => worker.removeEventListener(type, listener),
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
function qualifies(kind: AckKind, mutation: MutationRecord): boolean {
  const el = mutation.target instanceof Element ? mutation.target : mutation.target.parentElement
  if (!el) return false
  if (mutation.type === 'attributes' && (mutation.attributeName === 'aria-busy' || mutation.attributeName === 'aria-sort' || mutation.attributeName === 'aria-disabled')) return true
  const inStatus = el.closest('[role=status],[role=progressbar],[role=alert]') !== null
  switch (kind) {
    case 'status': return inStatus || el.closest('[role=progressbar]') !== null
    case 'dialog-open': return mutation.type === 'childList' && [...mutation.addedNodes].some(node => node instanceof Element && (node.matches('[role=dialog]') || node.querySelector('[role=dialog]') !== null))
    case 'dialog-close': return mutation.type === 'childList' && [...mutation.removedNodes].some(node => node instanceof Element && (node.matches('[role=dialog]') || node.querySelector('[role=dialog]') !== null))
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
      const started = performance.now()
      const wait = () => {
        if (!hadWorker || state.workersTerminated > terminated || performance.now() - started > 5000) resolve()
        else setTimeout(wait, 20)
      }
      setTimeout(wait, 0)
    })
    state.booted = true
    return () => input.removeEventListener('change', onChange)
  }, [])
  const factory = useMemo(() => countingWorker, [])
  const internal = { onFrame } as Record<string, unknown>
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
