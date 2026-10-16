/**
 * The window.__viewerHarness implementation (S0.3), kept free of React/DOM
 * side effects at import time so it is unit-testable with injected clock,
 * requestAnimationFrame and PerformanceObserver.
 *
 * It is a plain object (not a class) so the exposed surface is explicit.
 * Honesty rules: a marker without a matching input is an error, a round-trip
 * can never hang (it times out), warmup runs are counted, exports are
 * snapshots, and long-task collection distinguishes "engine cannot observe it"
 * (unavailable) from "observed nothing" (a measured empty set).
 */
import { AllocationLedger } from './ledger'
import { createDeterministicClock } from './clock'
import { gatherEnvironment, parseHarnessParams, readBrowserContext } from './environment'
import { fixtureLoader, type FixtureInfo } from './fixture'
import { finalizeResult, unavailableMetric } from './result'
import type { HarnessEnvironment, HarnessMetric, HarnessResult, HarnessSample, SampleKind, ScenarioRefusal } from './types'

/** environment.note bound enforced by result.ts. */
const NOTE_LIMIT = 1024

export interface HarnessDeps {
  /** Millisecond clock; defaults to performance.now(). */
  clock?: () => number
  /** requestAnimationFrame; defaults to window.requestAnimationFrame. */
  raf?: (callback: () => void) => void
  /** PerformanceObserver constructor; pass `undefined` to model its absence. */
  performanceObserver?: typeof PerformanceObserver
  /** Environment provider; defaults to reading the live browser and URL. */
  environment?: (fixture: FixtureInfo | null) => HarnessEnvironment
  /** Upper bound for one marker round-trip (rAF never fires in hidden tabs). */
  timeoutMs?: number
}

export interface ViewerHarness {
  booted: boolean
  markInput(label: string): void
  /** Resolves once the presentation frame was reached; throws for unknown labels. */
  markPresent(label: string, discard?: boolean): Promise<void>
  recordRoundTrip(label: string, warmup?: boolean): Promise<void>
  sampleCount(name: string): number
  warmupCount(name: string): number
  startLongTaskCollection(): boolean
  stopLongTaskCollection(): void
  recordOneShot(name: string, unit: string, kind: SampleKind, value: number): void
  /** Counts one warmup run of a one-shot series (run, then dropped from values). */
  recordWarmup(name: string, unit: string, kind: SampleKind): void
  exportResult(): Promise<HarnessResult>
  ledger: AllocationLedger
  deterministicClock: { now(): number; inject(now: () => number): void }
  fixtureLoader: typeof fixtureLoader
  setFixture(info: FixtureInfo, grid: string[][] | null): void
  /** Marks the current kernel scenario "refused (by design)" (identity + note). */
  recordScenarioRefusal(refusal: ScenarioRefusal): void
}

interface SampleState {
  name: string
  unit: string
  kind: SampleKind
  values: number[]
  warmupExcluded: number
}

type LongTaskState = 'idle' | 'running' | 'stopped' | 'unsupported'

const DEFAULT_TIMEOUT_MS = 10_000

export function createViewerHarness(deps: HarnessDeps = {}): ViewerHarness {
  let booted = false
  let clock = createDeterministicClock(deps.clock)
  const raf = deps.raf ?? ((callback: () => void) => void window.requestAnimationFrame(callback))
  const timeoutMs = deps.timeoutMs ?? DEFAULT_TIMEOUT_MS
  const ledger = new AllocationLedger()
  const inputMarkers = new Map<string, number>()
  const samples = new Map<string, SampleState>()
  /** Labels with an input marker not yet consumed by markPresent. A second
   * concurrent use of a pending label is rejected so overlapping pairs can
   * never share (and silently overwrite) one start time. */
  const pendingInputs = new Set<string>()
  /** Labels with a markPresent flight in progress (double rAF outstanding). */
  const inflightPresent = new Set<string>()

  function markInput(label: string): void {
    if (pendingInputs.has(label)) {
      throw new Error(`markInput: label ${JSON.stringify(label)} is already pending; overlapping pairs must use distinct labels`)
    }
    inputMarkers.set(label, clock.now())
    pendingInputs.add(label)
  }
  let observer: PerformanceObserver | null = null
  let longtaskState: LongTaskState = 'idle'
  const longtaskEntries: number[] = []
  let fixture: FixtureInfo | null = null
  /** The scenario note before any refusal was recorded on it. */
  let baseNote = ''

  function sampleState(name: string, unit: string, kind: SampleKind): SampleState {
    const existing = samples.get(name)
    if (!existing) {
      const created: SampleState = { name, unit, kind, values: [], warmupExcluded: 0 }
      samples.set(name, created)
      return created
    }
    if (existing.unit !== unit) throw new Error(`sample ${JSON.stringify(name)} already uses unit ${existing.unit}, cannot record ${unit}`)
    if (existing.kind !== kind) throw new Error(`sample ${JSON.stringify(name)} already has kind ${existing.kind}, cannot record ${kind}`)
    return existing
  }

  function pushSample(name: string, unit: string, kind: SampleKind, value: number): void {
    sampleState(name, unit, kind).values.push(value)
  }

  function sampleCount(name: string): number {
    return samples.get(name)?.values.length ?? 0
  }

  function warmupCount(name: string): number {
    return samples.get(name)?.warmupExcluded ?? 0
  }

  /**
   * Mark presentation: the present time is read after a double
   * requestAnimationFrame, i.e. the frame after the one that painted the
   * acknowledgment. `discard` runs the same path but counts the run as warmup.
   */
  function markPresent(label: string, discard = false): Promise<void> {
    const inputMs = inputMarkers.get(label)
    if (inputMs === undefined) {
      throw new Error(`markPresent: unknown input label ${JSON.stringify(label)}; call markInput first`)
    }
    if (inflightPresent.has(label)) {
      throw new Error(`markPresent: label ${JSON.stringify(label)} already has a presentation flight in progress`)
    }
    inflightPresent.add(label)
    return new Promise<void>((resolve, reject) => {
      let settled = false
      const finish = (action: () => void): void => {
        if (settled) return
        settled = true
        clearTimeout(timer)
        pendingInputs.delete(label)
        inflightPresent.delete(label)
        inputMarkers.delete(label)
        action()
      }
      const timer = setTimeout(() => {
        finish(() => reject(new Error(`markPresent(${label}) timed out after ${timeoutMs} ms waiting for two animation frames (hidden tab?)`)))
      }, timeoutMs)
      raf(() => {
        raf(() => {
          if (settled) return
          try {
            const presentMs = clock.now()
            const state = sampleState('input-to-present', 'ms', 'operation')
            if (discard) state.warmupExcluded += 1
            else state.values.push(presentMs - inputMs)
            finish(resolve)
          } catch (error) {
            finish(() => reject(error))
          }
        })
      })
    })
  }

  function flushObserver(): void {
    if (!observer) return
    for (const entry of observer.takeRecords()) longtaskEntries.push(entry.duration)
  }

  function startLongTaskCollection(): boolean {
    if (longtaskState === 'running') return true
    const ctor = 'performanceObserver' in deps ? deps.performanceObserver : typeof PerformanceObserver === 'undefined' ? undefined : PerformanceObserver
    // WebKit has PerformanceObserver but not the longtask entry type, and
    // observe() on an unsupported type is silently ignored there — so check
    // supportedEntryTypes instead of trusting observe() not to throw.
    if (!ctor || !ctor.supportedEntryTypes?.includes('longtask')) {
      longtaskState = 'unsupported'
      return false
    }
    try {
      observer = new ctor((list) => {
        for (const entry of list.getEntries()) longtaskEntries.push(entry.duration)
      })
      observer.observe({ entryTypes: ['longtask'] })
    } catch {
      observer = null
      longtaskState = 'unsupported'
      return false
    }
    longtaskState = 'running'
    return true
  }

  function stopLongTaskCollection(): void {
    if (longtaskState !== 'running') return
    flushObserver()
    observer?.disconnect()
    observer = null
    longtaskState = 'stopped'
  }

  function longtaskSeries(): { sample: HarnessSample | null; metric: HarnessMetric | null } {
    if (longtaskState === 'unsupported') {
      return { sample: null, metric: unavailableMetric('longtask-duration', 'PerformanceObserver longtask is not supported in this browser') }
    }
    if (longtaskState === 'idle') {
      return { sample: null, metric: unavailableMetric('longtask-duration', 'long-task collection was not started') }
    }
    // Running or stopped: a measured set, possibly empty.
    return {
      sample: { name: 'longtask-duration', unit: 'ms', kind: 'event', values: [...longtaskEntries], warmupExcluded: 0 },
      metric: null,
    }
  }

  const environment = deps.environment ?? ((current: FixtureInfo | null) => gatherEnvironment(readBrowserContext(), parseHarnessParams(window.location.search), current))

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
    setFixture(info: FixtureInfo, grid: string[][] | null): void {
      // Replacing a fixture must not double-charge the previous one.
      ledger.release('fixture-bytes')
      ledger.release('grid-values')
      fixture = info
      ledger.charge('main', info.bytes, 'fixture-bytes')
      if (grid) ledger.charge('main', grid.length * Math.max(1, grid[0]?.length ?? 0) * 8, 'grid-values')
    },
    recordScenarioRefusal(refusal: ScenarioRefusal): void {
      if (!fixture?.scenario) throw new Error('recordScenarioRefusal: no kernel scenario fixture is set')
      const refusals = [...(fixture.scenario.refusals ?? []), { ...refusal }]
      if (!fixture.scenario.refusals) baseNote = fixture.note ?? ''
      // The structured refusals carry the reason; the bounded note leads with the outcome.
      const mark = `REFUSED (by design): ${refusals.map(r => `${r.boundary} ${r.kind} ${r.requiredBytes} B > ${r.limitBytes} B`).join('; ')} (aggregate at the first refused charge, not the peak; see environment.scenario.refusals)`
      const note = baseNote ? `${mark} | ${baseNote}` : mark
      fixture = { ...fixture, scenario: { ...fixture.scenario, refusals }, note: note.length > NOTE_LIMIT ? `${note.slice(0, NOTE_LIMIT - 1)}…` : note }
    },
    markInput,
    markPresent,
    /** One input→present round-trip; resolves when the sample (or warmup) is recorded. */
    async recordRoundTrip(label: string, warmup = false): Promise<void> {
      markInput(label)
      await markPresent(label, warmup)
    },
    sampleCount,
    warmupCount,
    recordOneShot(name: string, unit: string, kind: SampleKind, value: number): void {
      pushSample(name, unit, kind, value)
    },
    recordWarmup(name: string, unit: string, kind: SampleKind): void {
      sampleState(name, unit, kind).warmupExcluded += 1
    },
    startLongTaskCollection,
    stopLongTaskCollection,
    async exportResult(): Promise<HarnessResult> {
      // Flush but keep observing: exporting must not end the collection window.
      flushObserver()
      const exported: HarnessSample[] = [...samples.values()].map((sample) => ({
        name: sample.name,
        unit: sample.unit,
        kind: sample.kind,
        values: [...sample.values],
        warmupExcluded: sample.warmupExcluded,
      }))
      const longtask = longtaskSeries()
      if (longtask.sample) exported.push(longtask.sample)
      const extraMetrics: Record<string, HarnessMetric> = {
        'viewport-frame': unavailableMetric('viewport-frame', 'no scroll/frame kernel exists in the S0.3 scaffold; plugs in with the S5 viewport'),
        'cold-import': unavailableMetric('cold-import', 'the fixture import pipeline is not part of the S0.3 scaffold; plugs in with S1'),
        'progressive-first-viewport': unavailableMetric('progressive-first-viewport', 'first-viewport measurement needs the S1 importer and S5 viewport'),
      }
      if (longtask.metric) extraMetrics['longtask-duration'] = longtask.metric
      return finalizeResult({
        environment: environment(fixture),
        samples: exported,
        ledger: ledger.accounting(),
        extraMetrics,
      })
    },
  }
}
