/**
 * Shared types for the interactive-viewer real-browser harness (S0.3 scaffold).
 * Pure type module — no runtime code — so it can be imported from the dev page,
 * the bench CLI and unit tests alike.
 */

export type HarnessProfile = 'desktop' | 'mobile'
export type HarnessMode = 'smoke' | 'release'
export type FixtureId = 'F1' | 'F2' | 'F3' | 'F4' | 'F5' | 'F6' | 'F7' | 'F8'
export type SampleKind = 'operation' | 'frame' | 'event'
export type AllocationPool = 'main' | 'worker' | 'bitmap'
export type BrowserEngine = 'blink' | 'webkit' | 'gecko' | 'unknown'
/** `webdriver`: the page saw navigator.webdriver but only the driver knows which tool. */
export type AutomationKind = 'playwright' | 'webdriver' | 'none'
/**
 * Generic labels: `safari` is claimed only for a real, non-automated Safari
 * (see classifyBrowser); every other WebKit is plain `webkit`, and
 * `webkit-playwright` is applied only by the Playwright driver.
 */
export type BrowserName = 'chromium' | 'safari' | 'webkit' | 'firefox' | 'webkit-playwright' | 'unknown'
/**
 * What kind of evidence a result can ever count as. Derived from the
 * environment (see deriveEvidence) and re-checked on ingest.
 */
export type EvidenceLabel = 'scaffold-placeholder' | 'emulated-webkit' | 'automated-browser' | 'device'
export type FontStatus = 'ready' | 'loading' | 'unavailable'
export type CacheState = 'cold' | 'warm'

/**
 * One named measurement series. `warmupExcluded` is the COUNT of warmup
 * iterations that were run and dropped from `values` (spec §12.6 requires
 * sample counts and excluded warmup to be reported, never hidden). `event`
 * series (e.g. long tasks) are occurrences in a window: an empty event series
 * is a measured empty set, not an unavailable metric.
 */
export interface HarnessSample {
  name: string
  unit: string
  values: number[]
  warmupExcluded: number
  kind: SampleKind
}

/**
 * A derived metric. `ok` carries p50/p95/max with the sample count.
 * `insufficient-samples` refuses p95 below the spec minimum (30 operation
 * samples, 300 frames) but still reports p50/max. `none-observed` is a
 * measured empty event set (collection ran in a supporting engine and nothing
 * occurred). `unavailable` means the metric could not be measured at all — it
 * NEVER fabricates a 0.
 */
export type HarnessMetric =
  | { status: 'ok'; p50: number; p95: number; max: number; samples: number }
  | { status: 'insufficient-samples'; p50: number; max: number; samples: number; reason: string }
  | { status: 'none-observed'; samples: 0; reason: string }
  | { status: 'unavailable'; reason: string }

export interface PoolAccounting {
  retained: number
  peak: number
  charged: number
}

export interface ProfileConformance {
  conforms: boolean
  deviations: string[]
}

/** Operator-supplied facts the page cannot know (spec §12.2 device runs). */
export interface OperatorMetadata {
  deviceModel: string | null
  osVersion: string | null
  thermalState: string | null
  batteryState: string | null
}

export interface HarnessEnvironment {
  buildCommit: string
  hardware: {
    logicalCores: number | null
    deviceMemoryGiB: number | null
    gpu: string | null
  }
  /** `platformNote` is mandatory (non-empty) whenever `platform` is 'unknown'. */
  os: {
    platform: string
    platformNote: string | null
    userAgent: string
  }
  browser: {
    name: BrowserName
    version: string
    engine: BrowserEngine
    userAgent: string
    automation: AutomationKind
  }
  dpr: number
  viewport: { width: number; height: number }
  profileConformance: ProfileConformance
  fonts: { status: FontStatus; resolved: boolean }
  cache: { state: CacheState; note: string }
  source: { adapter: string; bytes: number; rows: number | null; cols: number | null }
  backend: string
  /** The fixture the operator asked for (F1..F8); never implies content. */
  requestedFixtureId: FixtureId
  /** The spec fixture actually loaded; null while only the placeholder exists. */
  fixtureId: FixtureId | null
  evidence: EvidenceLabel
  operator: OperatorMetadata
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
