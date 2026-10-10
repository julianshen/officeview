/**
 * Dev-only interactive-viewer harness page (S0.3 scaffold).
 *
 * Served by the vite dev server at /src/debug/interactive-viewer.html — never
 * part of the library build (excluded from tsconfig.build.json, no dist input).
 *
 * This is an HONEST measurement frame, not a viewer: the interactive viewer
 * does not exist yet. The page loads a tiny synthetic placeholder fixture
 * (a real minimal XLSX generated in-page), renders it with a minimal canvas
 * grid, and exposes window.__viewerHarness so real browsers can measure what
 * exists today. Viewer-kernel metrics (frames, cold import, first viewport)
 * are reported unavailable with reasons — never as zeros. Later packages
 * (S1 importer, S5 viewport, S7/S8 kernels, S10 gates) plug scenarios in.
 *
 * The harness logic lives in ./harness/harness.ts (unit-tested); this file is
 * only the React shell and the canvas placeholder renderer.
 */
import { useEffect, useRef, useState, type JSX } from 'react'
import { createRoot } from 'react-dom/client'
import { parseHarnessParams } from './harness/environment'
import { fixtureLoader } from './harness/fixture'
import { createViewerHarness, type ViewerHarness } from './harness/harness'
import { postResult } from './harness/collect'
import type { HarnessResult } from './harness/types'

declare global {
  interface Window {
    __viewerHarness?: ViewerHarness
  }
}

const MAX_CANVAS_DIMENSION = 2560

function renderGrid(canvas: HTMLCanvasElement | null, grid: string[][], harness: ViewerHarness): void {
  if (!canvas) return
  const dpr = Math.min(window.devicePixelRatio || 1, 3)
  const cssWidth = Math.min(window.innerWidth, MAX_CANVAS_DIMENSION)
  const cssHeight = Math.min(window.innerHeight - 120, MAX_CANVAS_DIMENSION)
  canvas.width = Math.round(cssWidth * dpr)
  canvas.height = Math.round(cssHeight * dpr)
  canvas.style.width = `${cssWidth}px`
  canvas.style.height = `${cssHeight}px`
  harness.ledger.release('canvas') // a re-render replaces the backing store, it does not add one
  harness.ledger.charge('bitmap', canvas.width * canvas.height * 4, 'canvas')
  const context = canvas.getContext('2d')
  if (!context) return
  const started = harness.deterministicClock.now()
  context.scale(dpr, dpr)
  context.fillStyle = '#ffffff'
  context.fillRect(0, 0, cssWidth, cssHeight)
  const rows = Math.min(grid.length, 60)
  const cols = Math.min(grid[0]?.length ?? 0, 10)
  const cellWidth = Math.max(40, Math.floor(cssWidth / cols))
  const cellHeight = 24
  context.strokeStyle = '#c8c8c8'
  context.fillStyle = '#111111'
  context.font = '12px sans-serif'
  for (let r = 0; r <= rows; r++) {
    const y = r * cellHeight
    context.beginPath()
    context.moveTo(0, y)
    context.lineTo(cols * cellWidth, y)
    context.stroke()
  }
  for (let c = 0; c <= cols; c++) {
    const x = c * cellWidth
    context.beginPath()
    context.moveTo(x, 0)
    context.lineTo(x, rows * cellHeight)
    context.stroke()
  }
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      context.fillText(grid[r][c], c * cellWidth + 4, r * cellHeight + 16)
    }
  }
  const painted = harness.deterministicClock.now()
  harness.recordOneShot('scaffold-render', 'ms', 'operation', painted - started)
}

function HarnessPage(): JSX.Element {
  const [status, setStatus] = useState('booting')
  const [fixtureNote, setFixtureNote] = useState('')
  const [exported, setExported] = useState<HarnessResult | null>(null)
  const [posted, setPosted] = useState<boolean | null>(null)
  const [error, setError] = useState<string | null>(null)
  const harnessRef = useRef<ViewerHarness | null>(null)
  const canvasRef = useRef<HTMLCanvasElement | null>(null)

  useEffect(() => {
    let cancelled = false
    void (async () => {
      const harness = createViewerHarness()
      harnessRef.current = harness
      try {
        parseHarnessParams(window.location.search)
      } catch (paramError) {
        setStatus('invalid parameters')
        setError(paramError instanceof Error ? paramError.message : String(paramError))
        return
      }
      try {
        const { info, grid } = await fixtureLoader.loadPlaceholder()
        if (cancelled) return
        harness.setFixture(info, grid)
        renderGrid(canvasRef.current, grid ?? [], harness)
        harness.startLongTaskCollection()
        harness.booted = true
        window.__viewerHarness = harness
        setStatus('ready')
        setFixtureNote(info.note ?? '')
      } catch (catchError) {
        if (!cancelled) {
          setStatus('fixture error')
          setError(catchError instanceof Error ? catchError.message : String(catchError))
        }
      }
    })()
    return () => {
      cancelled = true
    }
  }, [])

  async function handleExport(): Promise<void> {
    const harness = harnessRef.current
    if (!harness) return
    try {
      const result = await harness.exportResult()
      setExported(result)
      const outcome = await postResult(result, window.location, window.fetch.bind(window))
      if (outcome.status === 'posted') setPosted(true)
      else if (outcome.status === 'rejected') {
        setPosted(false)
        setError(`collector rejected the result: ${outcome.errors.join('; ')}`)
      }
    } catch (catchError) {
      setError(catchError instanceof Error ? catchError.message : String(catchError))
    }
  }

  return (
    <div style={{ fontFamily: 'system-ui, sans-serif', margin: '16px' }}>
      <h1 style={{ fontSize: 18 }}>Interactive viewer harness (S0.3 scaffold)</h1>
      <p>
        status: <strong>{status}</strong>
        {fixtureNote ? ` — ${fixtureNote}` : ''}
      </p>
      {error ? <p style={{ color: '#c00' }}>error: {error}</p> : null}
      <canvas ref={canvasRef} style={{ border: '1px solid #999', display: status === 'ready' ? 'block' : 'none' }} />
      <p>
        <button onClick={handleExport} disabled={status !== 'ready'}>
          Export result
        </button>
        {posted === true ? ' — posted to collector' : null}
        {posted === false ? ' — collector rejected the result' : null}
      </p>
      {exported ? (
        <details>
          <summary>result JSON (hash {exported.hash.slice(0, 12)}…)</summary>
          <pre style={{ fontSize: 11, maxHeight: 320, overflow: 'auto' }}>{JSON.stringify(exported, null, 2)}</pre>
        </details>
      ) : null}
    </div>
  )
}

const container = document.getElementById('root')
if (!container) throw new Error('harness page requires #root')
createRoot(container).render(<HarnessPage />)
