/**
 * Dev-only end-to-end page for the real interactive viewer (plan S10.1).
 *
 * Unlike the S0.3 measurement scaffold (interactive-viewer.tsx), this page
 * mounts the shipped public component exactly as a host would:
 *   <OfficeFile data={blob} interactiveSheets viewerWorker={factory} showSearch />
 * Query parameters:
 *   src=<url>     workbook bytes (the Playwright spec routes /__fixtures/* to
 *                 committed or freshly generated XLSX packages)
 *   worker=1      supply a real module Worker running src/worker/worker-entry.ts,
 *                 so a paged admission opens worker-backed (S5.4); without it a
 *                 large file keeps compatibility mode, as for any host
 *   copy=0        allowCopy={false}
 *   wcaps=a,b     capabilities the viewer worker denies itself (workerCaps.ts),
 *                 passed as the Worker name since init scripts never reach it
 *   focus=grid    focus the grid once it has painted (real-keyboard runs)
 *   name=1        hand the bytes over as a File named after the src path's last
 *                 segment (paged documents: the toolbar shows the file name)
 *   toolbar=0     showToolbar={false} at first; window.__viewerE2E.setToolbar(on)
 *                 toggles it at run time (paged toolbar CLS checks)
 *   theme=light|dark  the viewer theme (default system)
 *   paged=1       with worker=1: debug-only admission override (forcePaged.ts).
 *                 The page loads through loadOfficeSource with the object cell
 *                 cap lowered to 1 (and the worker kernel likewise), so even a
 *                 small file takes the worker-backed paged path. The document
 *                 is rendered by <OfficeDoc>, exactly as OfficeFile would.
 *
 * Instrumentation lives at public boundaries only: the worker port (commands
 * by kind / read op, transferred bytes, workers created/terminated), the
 * clipboard (captured text) and the viewer's @internal onFrame stats. Nothing
 * here reaches into viewer internals. Excluded from the library build
 * (tsconfig.build.json excludes src/debug).
 */
import { useEffect, useMemo, useState, type JSX } from 'react'
import { createRoot } from 'react-dom/client'
import { OfficeFile, type ViewerWorkerPort } from '../components/OfficeFile'
import { OfficeDoc } from '../components/OfficeDoc'
import '../styles/officeview.css'
import { loadOfficeSource, type AnyDoc } from '../components/loadOfficeSource'
import { openPagedWorkbook } from '../components/xlsx/pagedWorkbook'
import { pagedWorkbookOf } from '../components/pagedDocument'
import { FORCE_PAGED_LIMITS } from './forcePaged'
import type { SheetFrameStats } from '../components/xlsx/InteractiveSheetDoc'
import type { MainMemoryLedger } from '../xlsx/view/mainLedger'
// The library loads these on demand. The harness server only serves modules it
// has reached through the module graph from src/debug (LAN lockdown, see
// vite.harness.config.ts), so the page names them statically.
import '../components/largeSource'
import '../core/zip'
import '../components/xlsx/pagedWorkbook'
import '../components/xlsx/InteractiveSheetDoc'

interface CommandRecord { kind: string; op?: string; at: number }
export interface ViewerE2E {
  loaded: boolean
  error: string | null
  frames: number
  lastFrame: SheetFrameStats | null
  /** Peak bitmap bytes over every frame since the last reset. */
  peakBitmapBytes: number
  maxFrameRecords: number
  commands: CommandRecord[]
  replies: number
  transferBytes: number
  workersCreated: number
  workersTerminated: number
  mainLedger: MainMemoryLedger | null
  clipboard: string[]
  /** Paged documents: show or hide the toolbar and status line. */
  setToolbar(on: boolean): void
  counts(): Record<string, number>
  reset(): void
}

declare global {
  interface Window { __viewerE2E?: ViewerE2E }
}

const params = new URLSearchParams(window.location.search)
const state: ViewerE2E = {
  loaded: false,
  error: null,
  frames: 0,
  lastFrame: null,
  peakBitmapBytes: 0,
  maxFrameRecords: 0,
  commands: [],
  replies: 0,
  transferBytes: 0,
  workersCreated: 0,
  workersTerminated: 0,
  mainLedger: null,
  clipboard: [],
  setToolbar() {},
  counts() {
    const out: Record<string, number> = {}
    for (const command of state.commands) {
      const key = command.op ? `${command.kind}:${command.op}` : command.kind
      out[key] = (out[key] ?? 0) + 1
    }
    return out
  },
  reset() {
    state.commands = []
    state.replies = 0
    state.transferBytes = 0
    state.frames = 0
    state.peakBitmapBytes = 0
    state.maxFrameRecords = 0
  },
}
window.__viewerE2E = state

// The clipboard is captured, never written: copy is asserted by its exact text.
Object.defineProperty(navigator, 'clipboard', {
  configurable: true,
  value: { writeText: async (text: string) => { state.clipboard.push(text) } },
})

function transferSize(transfer: Transferable[]): number {
  let bytes = 0
  for (const item of transfer) if (item instanceof ArrayBuffer) bytes += item.byteLength
  return bytes
}

/** A real dedicated module worker behind a counting port. */
function countingWorker(): ViewerWorkerPort {
  const caps = params.get('wcaps')
  const name = [caps ? `caps:${caps}` : '', params.get('paged') === '1' ? 'paged' : ''].filter(Boolean).join(';')
  const worker = new Worker(new URL('./viewerWorker.ts', import.meta.url), { name, type: 'module' })
  state.workersCreated++
  worker.addEventListener('message', (event) => { state.replies++; if (params.get('log') === '1') console.log('reply', JSON.stringify(event.data).slice(0, 400)) })
  return {
    postMessage(message: unknown, transfer: Transferable[]) {
      const command = message as { kind?: string; query?: { op?: string } }
      state.commands.push({ kind: String(command.kind), ...(command.query?.op ? { op: command.query.op } : {}), at: performance.now() })
      state.transferBytes += transferSize(transfer)
      worker.postMessage(message, transfer)
    },
    addEventListener: (type, listener) => worker.addEventListener(type, listener),
    removeEventListener: (type, listener) => worker.removeEventListener(type, listener),
    terminate() { state.workersTerminated++; worker.terminate() },
  }
}

function onFrame(stats: SheetFrameStats): void {
  state.frames++
  state.lastFrame = stats
  state.peakBitmapBytes = Math.max(state.peakBitmapBytes, stats.bitmapBytes)
  state.maxFrameRecords = Math.max(state.maxFrameRecords, stats.frameRecords)
}

function onMainLedger(ledger: MainMemoryLedger): void {
  state.mainLedger = ledger
}

function E2EPage(): JSX.Element {
  const [blob, setBlob] = useState<Blob | null>(null)
  const src = params.get('src')
  useEffect(() => {
    if (!src) { state.error = 'missing ?src='; return }
    void fetch(src).then(async response => {
      if (!response.ok) throw new Error(`fixture ${src}: HTTP ${response.status}`)
      const blob = await response.blob()
      setBlob(params.get('name') === '1' ? new File([blob], decodeURIComponent(src.split(/[?#]/)[0].split('/').pop() ?? 'file')) : blob)
      state.loaded = true
    }).catch(error => { state.error = error instanceof Error ? error.message : String(error) })
  }, [src])
  const worker = params.get('worker') === '1'
  // OfficeFile's paged route keys on a stable factory identity.
  const factory = useMemo(() => (worker ? countingWorker : undefined), [worker])
  // onFrame and onMainLedger are the viewer's @internal instrumentation props;
  // OfficeFile forwards unknown props to the interactive viewer unchanged.
  const internal = { onFrame, onMainLedger } as Record<string, unknown>
  const [toolbar, setToolbar] = useState(params.get('toolbar') !== '0')
  state.setToolbar = setToolbar
  const theme = params.get('theme')
  if (worker && params.get('paged') === '1') return <ForcedPagedDoc blob={blob} internal={internal} />
  return (
    <OfficeFile
      data={blob}
      interactiveSheets
      showSearch
      showToolbar={toolbar}
      {...(theme === 'light' || theme === 'dark' ? { theme } : {})}
      allowCopy={params.get('copy') !== '0'}
      {...(factory ? { viewerWorker: factory } : {})}
      loading={<div role="status">Loading workbook…</div>}
      error={message => <div role="alert" data-testid="e2e-error">{message}</div>}
      style={{ flex: '1 1 auto', minHeight: 0 }}
      {...internal}
    />
  )
}

// focus=grid: focus the grid once it is up, for runs that send real OS key
// events to a normal (non-automation) Safari window, where no script can.
if (params.get('focus') === 'grid') {
  const started = Date.now()
  const tryFocus = () => {
    const grid = document.querySelector<HTMLElement>('[role=grid]')
    if (grid && state.frames > 0) grid.focus()
    else if (Date.now() - started < 60_000) setTimeout(tryFocus, 200)
  }
  tryFocus()
}

/** `?paged=1&worker=1`: the paged route of useOfficeFile, with admission forced paged. */
function ForcedPagedDoc({ blob, internal }: { blob: Blob | null; internal: Record<string, unknown> }): JSX.Element {
  const [doc, setDoc] = useState<AnyDoc | null>(null)
  const [error, setError] = useState<string | null>(null)
  useEffect(() => {
    if (!blob) return
    const controller = new AbortController()
    let loaded: AnyDoc | undefined
    let cancelled = false
    void loadOfficeSource(blob, {
      signal: controller.signal,
      limits: FORCE_PAGED_LIMITS,
      paged: async (source, ctx, release) => openPagedWorkbook(source, { port: countingWorker(), profile: ctx.profile, signal: ctx.signal, ...(release ? { release } : {}) }),
    }).then(document => {
      if (cancelled) { void pagedWorkbookOf(document)?.release(); return }
      loaded = document
      setDoc(document)
    }).catch(cause => { if (!cancelled) setError(cause instanceof Error ? cause.message : String(cause)) })
    return () => { cancelled = true; controller.abort(); void pagedWorkbookOf(loaded)?.release() }
  }, [blob])
  if (error) return <div role="alert" data-testid="e2e-error">{error}</div>
  if (!doc) return <div role="status">Loading workbook…</div>
  return (
    <OfficeDoc
      document={doc}
      interactiveSheets
      showSearch
      allowCopy={params.get('copy') !== '0'}
      style={{ flex: '1 1 auto', minHeight: 0 }}
      {...internal}
    />
  )
}

const container = document.getElementById('root')
if (!container) throw new Error('e2e page requires #root')
createRoot(container).render(<E2EPage />)
