/**
 * Shared builders for the harness unit tests. `rehash` models an attacker (or
 * a buggy device) that edits a result and recomputes the integrity tag: the
 * hash only detects accidents, so ingest has to enforce the honesty
 * invariants itself.
 */
import { AllocationLedger } from '../../src/debug/harness/ledger'
import { canonicalize, checkProfileConformance, deriveEvidence, finalizeResult, hashResult } from '../../src/debug/harness/result'
import type { HarnessEnvironment, HarnessResult, HarnessSample } from '../../src/debug/harness/types'

export function makeEnvironment(overrides: Partial<HarnessEnvironment> = {}): HarnessEnvironment {
  const base: HarnessEnvironment = {
    buildCommit: 'deadbeef',
    hardware: { logicalCores: 8, deviceMemoryGiB: 16, gpu: 'Apple M4 Pro' },
    os: { platform: 'macOS', platformNote: null, userAgent: 'HarnessTest' },
    browser: { name: 'chromium', version: '1.0', engine: 'blink', userAgent: 'HarnessTest', automation: 'none' },
    dpr: 2,
    viewport: { width: 1440, height: 900 },
    profileConformance: { conforms: true, deviations: [] },
    fonts: { status: 'ready', resolved: true },
    cache: { state: 'cold', note: 'fresh context' },
    source: { adapter: 'xlsx', bytes: 1024, rows: 120, cols: 10 },
    backend: 'canvas2d-placeholder',
    requestedFixtureId: 'F1',
    fixtureId: 'F1',
    evidence: 'device',
    operator: { deviceModel: null, osVersion: null, thermalState: null, batteryState: null },
    profile: 'desktop',
    mode: 'smoke',
    note: 'test environment',
  }
  const env = { ...base, ...overrides }
  // Keep derived facts consistent unless a test overrides them on purpose.
  if (!('profileConformance' in overrides)) env.profileConformance = checkProfileConformance(env.profile, env.viewport, env.dpr)
  if (!('evidence' in overrides)) env.evidence = deriveEvidence(env)
  return env
}

export function makeSample(values: number[], kind: HarnessSample['kind'] = 'operation', name = 'op'): HarnessSample {
  return { name, unit: 'ms', values, warmupExcluded: 0, kind }
}

export function range(n: number): number[] {
  return Array.from({ length: n }, (_, i) => i + 1)
}

export async function buildResult(
  samples: HarnessSample[] = [makeSample(range(30))],
  env: Partial<HarnessEnvironment> = {},
): Promise<HarnessResult> {
  return finalizeResult({ environment: makeEnvironment(env), samples, ledger: new AllocationLedger().accounting() })
}

/** Deep clone through JSON — also what a posted body looks like after parsing. */
export function clone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T
}

/** Apply `mutate` to a clone and recompute the hash, as a forger would. */
export async function rehash(result: HarnessResult, mutate: (draft: Record<string, any>) => void): Promise<unknown> {
  const draft = clone(result) as Record<string, any>
  mutate(draft)
  const { hash: _drop, ...rest } = draft
  return { ...rest, hash: await hashResult(rest) }
}

export { canonicalize }
