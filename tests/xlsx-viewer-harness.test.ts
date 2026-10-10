/**
 * Unit tests for the S0.3 real-browser harness scaffold.
 *
 * These assert schema, accounting and canonicalization semantics only.
 * NO timing or SLA assertions live here — performance gates belong to the
 * controlled real-browser harness (spec §12.6), never to Vitest.
 */
import { describe, expect, it } from 'vitest'
import { AllocationLedger } from '../src/debug/harness/ledger'
import {
  MIN_FRAME_SAMPLES,
  MIN_OPERATION_SAMPLES,
  canonicalize,
  deriveMetric,
  finalizeResult,
  hashResult,
  percentile,
  unavailableMetric,
  validateHarnessResult,
  verifyHarnessResult,
} from '../src/debug/harness/result'
import type { AllocationPool, FixtureId, HarnessEnvironment, HarnessMetric, HarnessSample } from '../src/debug/harness/types'

function makeEnvironment(overrides: Partial<HarnessEnvironment> = {}): HarnessEnvironment {
  return {
    buildCommit: 'deadbeef',
    hardware: { logicalCores: 8, deviceMemoryGiB: 16, gpu: 'Apple M4 Pro' },
    os: { platform: 'macOS', userAgent: 'HarnessTest' },
    browser: { name: 'chromium', version: '1.0', engine: 'blink', userAgent: 'HarnessTest', automation: 'none' },
    dpr: 2,
    viewport: { width: 1440, height: 900 },
    fonts: { status: 'ready', resolved: true },
    cache: { state: 'cold', note: 'fresh context' },
    source: { adapter: 'synthetic-placeholder', bytes: 1024, rows: 120, cols: 10 },
    backend: 'canvas2d-placeholder',
    fixtureId: 'F1',
    profile: 'desktop',
    mode: 'smoke',
    note: 'test environment',
    ...overrides,
  }
}

function makeSample(values: number[], kind: 'operation' | 'frame' = 'operation'): HarnessSample {
  return { name: 'op', unit: 'ms', values, warmupExcluded: false, kind }
}

function range(n: number): number[] {
  return Array.from({ length: n }, (_, i) => i + 1)
}

describe('AllocationLedger', () => {
  it('charges bytes per pool and owner with retained/peak/charged accounting', () => {
    const ledger = new AllocationLedger()
    ledger.charge('bitmap', 100, 'canvas')
    ledger.charge('bitmap', 50, 'overlay')
    ledger.charge('main', 10, 'fixture')
    const accounting = ledger.accounting()
    expect(accounting.bitmap).toEqual({ retained: 150, peak: 150, charged: 150 })
    expect(accounting.main).toEqual({ retained: 10, peak: 10, charged: 10 })
    expect(accounting.worker).toEqual({ retained: 0, peak: 0, charged: 0 })
  })

  it('tracks peak retained even after release', () => {
    const ledger = new AllocationLedger()
    ledger.charge('main', 100, 'a')
    ledger.charge('main', 100, 'b')
    expect(ledger.accounting().main.peak).toBe(200)
    expect(ledger.release('a')).toBe(true)
    expect(ledger.accounting().main).toEqual({ retained: 100, peak: 200, charged: 200 })
  })

  it('release frees every charge owned by the owner and rejects unknown owners', () => {
    const ledger = new AllocationLedger()
    ledger.charge('worker', 10, 'w1')
    ledger.charge('worker', 5, 'w1')
    expect(ledger.release('w1')).toBe(true)
    expect(ledger.accounting().worker.retained).toBe(0)
    expect(ledger.release('w1')).toBe(false)
    expect(ledger.release('never-seen')).toBe(false)
  })

  it('rejects non-finite or negative byte counts', () => {
    const ledger = new AllocationLedger()
    expect(() => ledger.charge('main', -1, 'x')).toThrow()
    expect(() => ledger.charge('main', Number.NaN, 'x')).toThrow()
    expect(() => ledger.charge('main', Number.POSITIVE_INFINITY, 'x')).toThrow()
  })

  it('exposes every pool even when untouched', () => {
    const ledger = new AllocationLedger()
    const pools = Object.keys(ledger.accounting()) as AllocationPool[]
    expect(pools.sort()).toEqual(['bitmap', 'main', 'worker'])
  })
})

describe('percentile (nearest-rank)', () => {
  it('returns the nearest-rank value for p50 and p95', () => {
    const values = range(100)
    expect(percentile(values, 50)).toBe(50)
    expect(percentile(values, 95)).toBe(95)
  })

  it('clamps to the last value when the rank exceeds the sample count', () => {
    expect(percentile([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], 95)).toBe(10)
  })

  it('sorts before ranking and handles a single value', () => {
    expect(percentile([30, 10, 20], 50)).toBe(20)
    expect(percentile([42], 95)).toBe(42)
  })

  it('refuses an empty sample set', () => {
    expect(() => percentile([], 50)).toThrow()
  })
})

describe('deriveMetric sample-count rules', () => {
  it('reports ok with p50/p95/max at the 30-sample operation minimum', () => {
    const metric = deriveMetric(makeSample(range(MIN_OPERATION_SAMPLES)))
    expect(metric).toEqual({ status: 'ok', p50: 15, p95: 29, max: 30, samples: 30 })
  })

  it('refuses p95 below the 30-sample operation minimum', () => {
    const metric = deriveMetric(makeSample(range(MIN_OPERATION_SAMPLES - 1)))
    expect(metric.status).toBe('insufficient-samples')
    if (metric.status === 'insufficient-samples') {
      expect(metric.p50).toBe(15)
      expect(metric.max).toBe(29)
      expect(metric.samples).toBe(29)
      expect(metric.reason).toMatch(/30/)
    }
  })

  it('applies the 300-sample frame minimum to frame samples', () => {
    const below = deriveMetric(makeSample(range(MIN_FRAME_SAMPLES - 1), 'frame'))
    expect(below.status).toBe('insufficient-samples')
    const atMinimum = deriveMetric(makeSample(range(MIN_FRAME_SAMPLES), 'frame'))
    expect(atMinimum.status).toBe('ok')
    if (atMinimum.status === 'ok') expect(atMinimum.p95).toBe(285)
  })

  it('reports unavailable with a reason — never 0 — when no samples exist', () => {
    const metric = deriveMetric(makeSample([]))
    expect(metric).toEqual({ status: 'unavailable', reason: expect.stringContaining('no samples') })
  })

  it('unavailableMetric never fabricates numeric values', () => {
    const metric = unavailableMetric('viewport-frame', 'no kernel yet')
    expect(metric).toEqual({ status: 'unavailable', reason: 'viewport-frame: no kernel yet' })
  })
})

describe('canonicalization', () => {
  it('is independent of key insertion order, including nested objects', () => {
    expect(canonicalize({ a: 1, b: { c: 2, d: 3 } })).toBe(canonicalize({ b: { d: 3, c: 2 }, a: 1 }))
  })

  it('preserves array order and scalar encoding', () => {
    expect(canonicalize({ v: [3, 1, 2] })).toBe('{"v":[3,1,2]}')
    expect(canonicalize({ s: 'x', n: 1.5, z: null })).toBe('{"n":1.5,"s":"x","z":null}')
  })
})

describe('hash tag', () => {
  it('matches the published SHA-256 test vector', async () => {
    const { sha256Hex } = await import('../src/debug/harness/result')
    await expect(sha256Hex('abc')).resolves.toBe(
      'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad',
    )
  })

  it('is stable for identical input and changes when any value changes', async () => {
    const draft = { environment: makeEnvironment(), samples: [], metrics: {}, ledger: {} }
    const a = await hashResult(draft)
    const b = await hashResult(draft)
    expect(a).toBe(b)
    const tampered = await hashResult({ ...draft, environment: makeEnvironment({ fixtureId: 'F2' }) })
    expect(tampered).not.toBe(a)
  })
})

describe('validateHarnessResult / verifyHarnessResult / finalizeResult', () => {
  it('finalizeResult derives metrics, ledger and a verifiable hash', async () => {
    const ledger = new AllocationLedger()
    ledger.charge('bitmap', 2048, 'canvas')
    const result = await finalizeResult({
      environment: makeEnvironment(),
      samples: [makeSample(range(MIN_OPERATION_SAMPLES))],
      ledger: ledger.accounting(),
    })
    expect(result.schemaVersion).toBe(1)
    expect(result.metrics['op'].status).toBe('ok')
    expect(result.ledger.bitmap.retained).toBe(2048)
    expect(result.hash).toMatch(/^[0-9a-f]{64}$/)
    expect(validateHarnessResult(result)).toEqual({ ok: true, result })
    await expect(verifyHarnessResult(result)).resolves.toEqual({ ok: true, errors: [] })
  })

  it('detects accidental modification through the hash', async () => {
    const result = await finalizeResult({ environment: makeEnvironment(), samples: [], ledger: new AllocationLedger().accounting() })
    const tampered = { ...result, environment: { ...result.environment, dpr: 1 } }
    const outcome = await verifyHarnessResult(tampered)
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.errors[0]).toMatch(/hash mismatch/)
  })

  it('rejects invalid fixture ids, profiles and modes', async () => {
    const result = await finalizeResult({ environment: makeEnvironment(), samples: [], ledger: new AllocationLedger().accounting() })
    for (const bad of [
      { ...result, environment: { ...result.environment, fixtureId: 'F9' as FixtureId } },
      { ...result, environment: { ...result.environment, profile: 'tablet' as HarnessEnvironment['profile'] } },
      { ...result, environment: { ...result.environment, mode: 'debug' as HarnessEnvironment['mode'] } },
    ]) {
      const outcome = validateHarnessResult(bad)
      expect(outcome.ok).toBe(false)
    }
  })

  it('rejects malformed samples, metrics and hashes', async () => {
    const result = await finalizeResult({ environment: makeEnvironment(), samples: [], ledger: new AllocationLedger().accounting() })
    const badSamples = validateHarnessResult({ ...result, samples: [{ name: 'x' }] })
    expect(badSamples.ok).toBe(false)
    const badMetric: HarnessMetric = { status: 'unavailable' } as HarnessMetric
    expect(validateHarnessResult({ ...result, metrics: { m: badMetric } }).ok).toBe(false)
    const okWithoutP95 = { status: 'ok', p50: 1, max: 2, samples: 3 } as HarnessMetric
    expect(validateHarnessResult({ ...result, metrics: { m: okWithoutP95 } }).ok).toBe(false)
    expect(validateHarnessResult({ ...result, hash: 'deadbeef' }).ok).toBe(false)
    expect(validateHarnessResult(null).ok).toBe(false)
    expect(validateHarnessResult({ schemaVersion: 2 }).ok).toBe(false)
  })

  it('accepts every fixture id F1..F8', async () => {
    for (const fixtureId of ['F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8'] as FixtureId[]) {
      const result = await finalizeResult({
        environment: makeEnvironment({ fixtureId }),
        samples: [],
        ledger: new AllocationLedger().accounting(),
      })
      expect(validateHarnessResult(result).ok).toBe(true)
    }
  })
})
