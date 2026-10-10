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
 */
import { useEffect, useRef, useState, type JSX } from 'react'
import { createRoot } from 'react-dom/client'
import JSZip from 'jszip'
import { AllocationLedger } from './harness/ledger'
import { createDeterministicClock } from './harness/clock'
import { deriveMetric, finalizeResult, unavailableMetric, FIXTURE_IDS } from './harness/result'
import type {
  FixtureId,
  HarnessEnvironment,
  HarnessMetric,
  HarnessResult,
  HarnessSample,
} from './harness/types'

declare global {
  interface Window {
    __viewerHarness?: ViewerHarness
  }
}

const PLACEHOLDER_ROWS = 120
const PLACEHOLDER_COLS = 10
const MAX_CANVAS_DIMENSION = 2560

interface FixtureInfo {
  adapter: string
  bytes: number
  rows: number | null
  cols: number | null
  note?: string
}

interface LoadedFixture {
  info: FixtureInfo
  grid: string[][] | null
}

/** A real minimal XLSX, generated deterministically in-page. */
async function buildPlaceholderXlsx(rows: number, cols: number): Promise<{ bytes: Uint8Array; grid: string[][] }> {
  const grid: string[][] = []
  const sheetRows: string[] = []
  for (let r = 0; r < rows; r++) {
    const row: string[] = []
    const cells: string[] = []
    for (let c = 0; c < cols; c++) {
      const value = r === 0 ? `Column ${c + 1}` : String((r * 31 + c * 17) % 997)
      row.push(value)
      cells.push(`<c r="${columnName(c)}${r + 1}" t="inlineStr"><is><t>${value}</t></is></c>`)
    }
    grid.push(row)
    sheetRows.push(`<row r="${r + 1}">${cells.join('')}</row>`)
  }
  const sheet =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<sheetData>${sheetRows.join('')}</sheetData></worksheet>`
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Placeholder" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '</Relationships>',
  )
  zip.file('xl/worksheets/sheet1.xml', sheet)
  const bytes = await zip.generateAsync({ type: 'uint8array' })
  return { bytes, grid }
}

function columnName(index: number): string {
  let name = ''
  let n = index
  while (n >= 0) {
    name = String.fromCharCode(65 + (n % 26)) + name
    n = Math.floor(n / 26) - 1
  }
  return name
}

function detectGpu(): string | null {
  try {
    const canvas = document.createElement('canvas')
    const gl = canvas.getContext('webgl2') ?? canvas.getContext('webgl')
    if (!gl) return null
    const info = gl.getExtension('WEBGL_debug_renderer_info')
    const renderer = info ? gl.getParameter(info.UNMASKED_RENDERER_WEBGL) : gl.getParameter(gl.RENDERER)
    return typeof renderer === 'string' ? renderer : null
  } catch {
    return null
  }
}

const fixtureLoader = {
  async loadPlaceholder(): Promise<LoadedFixture> {
    const { bytes, grid } = await buildPlaceholderXlsx(PLACEHOLDER_ROWS, PLACEHOLDER_COLS)
    return {
      info: {
        adapter: 'xlsx-placeholder',
        bytes: bytes.byteLength,
        rows: PLACEHOLDER_ROWS,
        cols: PLACEHOLDER_COLS,
        note: 'tiny synthetic XLSX generated in-page; real fixtures arrive with S0.2',
      },
      grid,
    }
  },
  async load(url: string): Promise<LoadedFixture> {
    const response = await fetch(url)
    if (!response.ok) throw new Error(`fixture fetch failed with status ${response.status}`)
    const buffer = await response.arrayBuffer()
    return this.loadBuffer(new Uint8Array(buffer))
  },
  async loadBuffer(bytes: Uint8Array): Promise<LoadedFixture> {
    const isZip = bytes.length > 4 && bytes[0] === 0x50 && bytes[1] === 0x4b
    return {
      info: {
        adapter: isZip ? 'xlsx' : 'unknown',
        bytes: bytes.byteLength,
        rows: null,
        cols: null,
        note: isZip
          ? 'source bytes recorded; parsing/rendering awaits the S1 importer and S5 viewport'
          : 'unrecognized fixture format',
      },
      grid: null,
    }
  },
}

interface SampleState {
  name: string
  unit: string
  kind: 'operation' | 'frame'
  values: number[]
  warmupExcluded: boolean
}

interface SampleWaiter {
  name: string
  target: number
  resolve: () => void
}

/**\ * The harness API exposed on window.__viewerHarness. A plain object (not a
 * class) so the exposed surface is explicit and enumerable.
 */
export interface ViewerHarness {
  booted: boolean
  markInput(label: string): void
  markPresent(label: string, discard?: boolean): void
  recordRoundTrip(label: string, warmup?: boolean): Promise<void>
  sampleCount(name: string): number
  startLongTaskCollection(): boolean
  stopLongTaskCollection(): void
  recordOneShot(name: string, unit: string, kind: 'operation' | 'frame', value: number): void
  exportResult(): Promise<HarnessResult>
  ledger: AllocationLedger
  deterministicClock: { now(): number; inject(now: () => number): void }
  fixtureLoader: typeof fixtureLoader
  setFixture(info: FixtureInfo, grid: string[][] | null, fixtureId: FixtureId): void
}

export function createViewerHarness(): ViewerHarness {
  let booted = false
  let clock = createDeterministicClock()
  const ledger = new AllocationLedger()
  const inputMarkers = new Map<string, number>()
  const samples = new Map<string, SampleState>()
  const sampleWaiters: SampleWaiter[] = []
  let longtaskObserver: PerformanceObserver | null = null
  let longtaskEntries: number[] = []
  let longtaskSupported = true
  let longtaskRunning = false
  let excludedWarmupRuns = false
  let fixture: FixtureInfo | null = null
  let fixtureId: FixtureId = 'F1'

  function pushSample(name: string, unit: string, kind: 'operation' | 'frame', value: number): void {
    const sample =
      samples.get(name) ?? ({ name, unit, kind, values: [], warmupExcluded: excludedWarmupRuns } as SampleState)
    sample.values.push(value)
    samples.set(name, sample)
    const stillWaiting: SampleWaiter[] = []
    for (const waiter of sampleWaiters) {
      if (waiter.name === name && sampleCount(name) >= waiter.target) waiter.resolve()
      else stillWaiting.push(waiter)
    }
    sampleWaiters.length = 0
    sampleWaiters.push(...stillWaiting)
  }

  function waitForSampleCount(name: string, target: number): Promise<void> {
    if (sampleCount(name) >= target) return Promise.resolve()
    return new Promise((resolve) => {
      sampleWaiters.push({ name, target, resolve })
    })
  }

  function sampleCount(name: string): number {
    return samples.get(name)?.values.length ?? 0
  }

  /**
   * Mark presentation: records performance.now() at scheduling time and uses
   * a double requestAnimationFrame so the recorded present time is the frame
   * after the one that painted the acknowledgment.
   */
  function markPresent(label: string, discard = false): void {
    const inputMs = inputMarkers.get(label)
    if (inputMs === undefined) return
    clock.now()
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        const presentMs = clock.now()
        if (!discard) pushSample('input-to-present', 'ms', 'operation', presentMs - inputMs)
      })
    })
  }

  function startLongTaskCollection(): boolean {
    if (longtaskRunning) return true
    if (typeof PerformanceObserver === 'undefined') {
      longtaskSupported = false
      return false
    }
    try {
      longtaskObserver = new PerformanceObserver((list) => {
        for (const entry of list.getEntries()) longtaskEntries.push(entry.duration)
      })
      longtaskObserver.observe({ entryTypes: ['longtask'] })
      longtaskRunning = true
      return true
    } catch {
      longtaskSupported = false
      return false
    }
  }

  function stopLongTaskCollection(): void {
    longtaskObserver?.disconnect()
    longtaskObserver = null
    longtaskRunning = false
  }

  function gatherEnvironment(): HarnessEnvironment {
    const params = new URLSearchParams(window.location.search)
    const nav = window.navigator as Navigator & {
      deviceMemory?: number
      userAgentData?: { platform?: string; brands?: Array<{ name: string; version: string }> }
    }
    const userAgent = nav.userAgent
    const isChrome = /Chrome/.test(userAgent)
    const isSafari = /Safari/.test(userAgent) && !isChrome
    const versionMatch = userAgent.match(/(?:Chrome|Version|Firefox)\/([\d.]+)/)
    const cacheParam = params.get('cache')
    const profileParam = params.get('profile')
    const modeParam = params.get('mode')
    return {
      buildCommit: params.get('commit') ?? 'unknown',
      hardware: {
        logicalCores: nav.hardwareConcurrency ?? null,
        deviceMemoryGiB: nav.deviceMemory ?? null,
        gpu: detectGpu(),
      },
      os: { platform: nav.userAgentData?.platform ?? 'unknown', userAgent },
      browser: {
        name: isChrome ? 'chromium' : isSafari ? 'safari' : /Firefox/.test(userAgent) ? 'firefox' : 'unknown',
        version: versionMatch?.[1] ?? 'unknown',
        engine: isChrome ? 'blink' : isSafari ? 'webkit' : /Firefox/.test(userAgent) ? 'gecko' : 'unknown',
        userAgent,
        automation: nav.webdriver ? 'playwright' : 'none',
      },
      dpr: window.devicePixelRatio,
      viewport: { width: window.innerWidth, height: window.innerHeight },
      fonts: document.fonts
        ? {
            status: document.fonts.status === 'loaded' ? 'ready' : 'loading',
            resolved: document.fonts.status === 'loaded',
          }
        : { status: 'unavailable', resolved: false },
      cache: {
        state: cacheParam === 'warm' ? 'warm' : 'cold',
        note:
          cacheParam === 'warm'
            ? 'warm run requested; scenario-level warmup is the caller responsibility'
            : 'fresh browser context; no prior page state reused',
      },
      source: {
        adapter: fixture?.adapter ?? 'none',
        bytes: fixture?.bytes ?? 0,
        rows: fixture?.rows ?? null,
        cols: fixture?.cols ?? null,
      },
      backend: 'canvas2d-placeholder',
      fixtureId,
      profile: profileParam === 'mobile' ? 'mobile' : 'desktop',
      mode: modeParam === 'release' ? 'release' : 'smoke',
      note: 'S0.3 scaffold: synthetic placeholder fixture rendered by a minimal canvas grid; interactive viewer kernels are not implemented yet',
    }
  }

  return {
    get booted() {
      return booted
    },
    set booted(value: boolean) {
      booted = value
    },
    ledger,
    deterministicClock: {
      now: () => clock.now(),
      inject: (now: () => number) => {
        clock = createDeterministicClock(now)
      },
    },
    fixtureLoader,
    setFixture(info: FixtureInfo, grid: string[][] | null, id: FixtureId): void {
      fixture = info
      fixtureId = id
      ledger.charge('main', info.bytes, 'fixture-bytes')
      if (grid) {
        ledger.charge('main', grid.length * Math.max(1, grid[0]?.length ?? 0) * 8, 'grid-values')
      }
    },
    markInput(label: string): void {
      inputMarkers.set(label, clock.now())
    },
    markPresent,
    /** Run one input→present round-trip; resolves once the sample is recorded. */
    async recordRoundTrip(label: string, warmup = false): Promise<void> {
      if (warmup) excludedWarmupRuns = true
      inputMarkers.set(label, clock.now())
      const target = sampleCount('input-to-present') + 1
      const recorded = waitForSampleCount('input-to-present', target)
      markPresent(label, warmup)
      await recorded
    },
    sampleCount,
    recordOneShot(name: string, unit: string, kind: 'operation' | 'frame', value: number): void {
      pushSample(name, unit, kind, value)
    },
    startLongTaskCollection,
    stopLongTaskCollection,
    async exportResult(): Promise<HarnessResult> {
      stopLongTaskCollection()
      const exportedSamples: HarnessSample[] = [...samples.values()].map((sample) => ({
        name: sample.name,
        unit: sample.unit,
        kind: sample.kind,
        values: sample.values,
        warmupExcluded: sample.warmupExcluded,
      }))
      const longtaskSample: HarnessSample = {
        name: 'longtask-duration',
        unit: 'ms',
        kind: 'operation',
        values: longtaskEntries,
        warmupExcluded: false,
      }
      const withLongtask = longtaskEntries.length > 0 ? [...exportedSamples, longtaskSample] : exportedSamples
      const extraMetrics: Record<string, HarnessMetric> = {
        'viewport-frame': unavailableMetric(
          'viewport-frame',
          'no scroll/frame kernel exists in the S0.3 scaffold; plugs in with the S5 viewport',
        ),
        'cold-import': unavailableMetric(
          'cold-import',
          'the fixture import pipeline is not part of the S0.3 scaffold; plugs in with S1',
        ),
        'progressive-first-viewport': unavailableMetric(
          'progressive-first-viewport',
          'first-viewport measurement needs the S1 importer and S5 viewport',
        ),
        'longtask-duration': !longtaskSupported
          ? unavailableMetric('longtask-duration', 'PerformanceObserver longtask is not supported in this browser')
          : longtaskEntries.length === 0
            ? unavailableMetric('longtask-duration', 'no longtask entries observed during the collection window')
            : deriveMetric(longtaskSample),
      }
      return finalizeResult({
        environment: gatherEnvironment(),
        samples: withLongtask,
        ledger: ledger.accounting(),
        extraMetrics,
      })
    },
  }
}

function renderGrid(canvas: HTMLCanvasElement | null, grid: string[][], harness: ViewerHarness): void {
  if (!canvas) return
  const dpr = Math.min(window.devicePixelRatio || 1, 3)
  const cssWidth = Math.min(window.innerWidth, MAX_CANVAS_DIMENSION)
  const cssHeight = Math.min(window.innerHeight - 120, MAX_CANVAS_DIMENSION)
  canvas.width = Math.round(cssWidth * dpr)
  canvas.height = Math.round(cssHeight * dpr)
  canvas.style.width = `${cssWidth}px`
  canvas.style.height = `${cssHeight}px`
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
      const params = new URLSearchParams(window.location.search)
      const fixtureParam = params.get('fixture') ?? 'F1'
      if (!FIXTURE_IDS.includes(fixtureParam as FixtureId)) {
        setStatus(`invalid fixture id ${fixtureParam}; expected F1..F8`)
        return
      }
      const fixtureId = fixtureParam as FixtureId
      try {
        const { info, grid } = await fixtureLoader.loadPlaceholder()
        if (cancelled) return
        harness.setFixture(info, grid, fixtureId)
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
      const collect = new URLSearchParams(window.location.search).get('collect')
      if (collect) {
        const response = await fetch(collect, {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify(result),
        })
        setPosted(response.ok)
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
      {error ? <p style={{ color: 'c00' }}>error: {error}</p> : null}
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
