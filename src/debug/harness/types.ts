/**
 * Shared types for the interactive-viewer real-browser harness (S0.3 scaffold).
 * Pure type module — no runtime code — so it can be imported from the dev page,
 * the bench CLI and unit tests alike.
 */

export type HarnessProfile = 'desktop' | 'mobile'
export type HarnessMode = 'smoke' | 'release'
export type FixtureId = 'F1' | 'F2' | 'F3' | 'F4' | 'F5' | 'F6' | 'F7' | 'F8'
export type SampleKind = 'operation' | 'frame'
export type AllocationPool = 'main' | 'worker' | 'bitmap'
export type BrowserEngine = 'blink' | 'webkit' | 'gecko' | 'unknown'
export type AutomationKind = 'playwright' | 'none' | 'unknown'
export type FontStatus = 'ready' | 'loading' | 'unavailable'
export type CacheState = 'cold' | 'warm'

/**
 * One named measurement series. `warmupExcluded` records that warmup
 * iterations were dropped from `values` before export (spec §12.6 requires
 * sample counts and excluded warmup to be reported, never hidden).
 */
export interface HarnessSample {
  name: string
  unit: string
  values: number[]
  warmupExcluded: boolean
  kind: SampleKind
}

/**
 * A derived metric. `ok` carries p50/p95/max with the sample count.
 * `insufficient-samples` refuses p95 below the spec minimum (30 operation
 * samples, 300 frames) but still reports p50/max. `unavailable` means the
 * metric could not be measured at all — it NEVER fabricates a 0.
 */
export type HarnessMetric =
  | { status: 'ok'; p50: number; p95: number; max: number; samples: number }
  | { status: 'insufficient-samples'; p50: number; max: number; samples: number; reason: string }
  | { status: 'unavailable'; reason: string }

export interface PoolAccounting {
  retained: number
  peak: number
  charged: number
}

export interface HarnessEnvironment {
  buildCommit: string
  hardware: {
    logicalCores: number | null
    deviceMemoryGiB: number | null
    gpu: string | null
  }
  os: {
    platform: string
    userAgent: string
  }
  browser: {
    name: string
    version: string
    engine: BrowserEngine
    userAgent: string
    automation: AutomationKind
  }
  dpr: number
  viewport: { width: number; height: number }
  fonts: { status: FontStatus; resolved: boolean }
  cache: { state: CacheState; note: string }
  source: { adapter: string; bytes: number; rows: number | null; cols: number | null }
  backend: string
  fixtureId: FixtureId
  profile: HarnessProfile
  mode: HarnessMode
  note: string
}

export interface HarnessResult {
  schemaVersion: 1
  environment: HarnessEnvironment
  samples: HarnessSample[]
  metrics: Record<string, HarnessMetric>
  ledger: Record<AllocationPool, PoolAccounting>
  /**
   * SHA-256 hex over the canonical JSON (sorted keys) of this result with the
   * `hash` field removed. Integrity tag only — see result.ts header.
   */
  hash: string
}

export interface ResultDraft {
  environment: HarnessEnvironment
  samples: HarnessSample[]
  ledger: Record<AllocationPool, PoolAccounting>
  extraMetrics?: Record<string, HarnessMetric>
}

export type ValidationOutcome = { ok: true; result: HarnessResult } | { ok: false; errors: string[] }
