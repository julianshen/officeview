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
import type { SheetFrameStats } from '../components/xlsx/InteractiveSheetDoc'
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
  clipboard: string[]
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
  clipboard: [],
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
  const worker = new Worker(new URL('./viewerWorker.ts', import.meta.url), { type: 'module' })
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

function E2EPage(): JSX.Element {
  const [blob, setBlob] = useState<Blob | null>(null)
  const src = params.get('src')
  useEffect(() => {
    if (!src) { state.error = 'missing ?src='; return }
    void fetch(src).then(async response => {
      if (!response.ok) throw new Error(`fixture ${src}: HTTP ${response.status}`)
      setBlob(await response.blob())
      state.loaded = true
    }).catch(error => { state.error = error instanceof Error ? error.message : String(error) })
  }, [src])
  const worker = params.get('worker') === '1'
  // OfficeFile's paged route keys on a stable factory identity.
  const factory = useMemo(() => (worker ? countingWorker : undefined), [worker])
  // onFrame is the viewer's @internal instrumentation prop; OfficeFile forwards
  // unknown props to the interactive viewer unchanged.
  const internal = { onFrame } as Record<string, unknown>
  return (
    <OfficeFile
      data={blob}
      interactiveSheets
      showSearch
      allowCopy={params.get('copy') !== '0'}
      {...(factory ? { viewerWorker: factory } : {})}
      loading={<div role="status">Loading workbook…</div>}
      error={message => <div role="alert" data-testid="e2e-error">{message}</div>}
      style={{ flex: '1 1 auto', minHeight: 0 }}
      {...internal}
    />
  )
}

const container = document.getElementById('root')
if (!container) throw new Error('e2e page requires #root')
createRoot(container).render(<E2EPage />)
