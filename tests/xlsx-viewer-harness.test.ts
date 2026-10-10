/**
 * Unit tests for the S0.3 real-browser harness scaffold.
 *
 * These assert schema, accounting and canonicalization semantics only.
 * NO timing or SLA assertions live here — performance gates belong to the
 * controlled real-browser harness (spec §12.6), never to Vitest.
 */
import { describe, expect, it, vi } from 'vitest'
import { createHash } from 'node:crypto'
import { AllocationLedger } from '../src/debug/harness/ledger'
import {
  MIN_FRAME_SAMPLES,
  MIN_OPERATION_SAMPLES,
  canonicalize,
  deriveMetric,
  finalizeResult,
  hashResult,
  percentile,
  sha256Hex,
  unavailableMetric,
  validateHarnessResult,
  verifyHarnessResult,
} from '../src/debug/harness/result'
import type { AllocationPool, HarnessEnvironment, HarnessMetric } from '../src/debug/harness/types'
import { buildResult, clone, makeEnvironment, makeSample, range, rehash } from './helpers/harnessResult'

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

  it('releases one owner that charged several pools without touching other owners', () => {
    const ledger = new AllocationLedger()
    ledger.charge('main', 10, 'view')
    ledger.charge('worker', 20, 'view')
    ledger.charge('bitmap', 40, 'view')
    ledger.charge('bitmap', 4, 'other')
    expect(ledger.release('view')).toBe(true)
    expect(ledger.accounting()).toEqual({
      main: { retained: 0, peak: 10, charged: 10 },
      worker: { retained: 0, peak: 20, charged: 20 },
      bitmap: { retained: 4, peak: 44, charged: 44 },
    })
  })

  it('releaseAll zeroes retained in every pool, forgets owners, and keeps peak/charged', () => {
    const ledger = new AllocationLedger()
    ledger.charge('main', 10, 'a')
    ledger.charge('worker', 20, 'b')
    ledger.charge('bitmap', 30, 'c')
    ledger.releaseAll()
    expect(ledger.accounting()).toEqual({
      main: { retained: 0, peak: 10, charged: 10 },
      worker: { retained: 0, peak: 20, charged: 20 },
      bitmap: { retained: 0, peak: 30, charged: 30 },
    })
    expect(ledger.release('a')).toBe(false)
    ledger.charge('main', 5, 'a')
    expect(ledger.retained('main')).toBe(5)
  })

  it('rejects non-finite or negative byte counts', () => {
    const ledger = new AllocationLedger()
    expect(() => ledger.charge('main', -1, 'x')).toThrow()
    expect(() => ledger.charge('main', Number.NaN, 'x')).toThrow()
    expect(() => ledger.charge('main', Number.POSITIVE_INFINITY, 'x')).toThrow()
    expect(ledger.accounting().main).toEqual({ retained: 0, peak: 0, charged: 0 })
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
    expect(() => percentile([], 50)).toThrow(/at least one value/)
  })

  it('refuses out-of-range or non-numeric p', () => {
    for (const p of [0, -1, 100.0001, 101, Number.NaN, Number.POSITIVE_INFINITY]) {
      expect(() => percentile([1, 2, 3], p), `p=${p}`).toThrow(/percentile p must be in \(0, 100\]/)
    }
    expect(percentile([1, 2, 3], 100)).toBe(3)
  })

  it('refuses non-finite values instead of ranking NaN', () => {
    expect(() => percentile([1, Number.NaN, 3], 50)).toThrow(/finite/)
    expect(() => percentile([1, Number.POSITIVE_INFINITY], 50)).toThrow(/finite/)
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
      expect('p95' in metric).toBe(false)
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

  it('an empty EVENT series is a measured empty set (none-observed), not unavailable', () => {
    const metric = deriveMetric(makeSample([], 'event', 'longtask-duration'))
    expect(metric).toEqual({ status: 'none-observed', samples: 0, reason: expect.stringContaining('longtask-duration') })
  })

  it('non-empty event series follow the operation minimum', () => {
    expect(deriveMetric(makeSample([60, 80], 'event')).status).toBe('insufficient-samples')
    expect(deriveMetric(makeSample(range(30), 'event')).status).toBe('ok')
  })

  it('unavailableMetric never fabricates numeric values', () => {
    const metric = unavailableMetric('viewport-frame', 'no kernel yet')
    expect(metric).toEqual({ status: 'unavailable', reason: 'viewport-frame: no kernel yet' })
  })
})

describe('canonicalization', () => {
  it('is independent of key insertion order, including nested objects and objects inside arrays', () => {
    const a = { a: 1, b: { c: 2, d: 3 }, e: [{ y: 1, x: 2 }] }
    const b = { e: [{ x: 2, y: 1 }], b: { d: 3, c: 2 }, a: 1 }
    expect(canonicalize(a)).toBe(canonicalize(b))
    expect(canonicalize(a)).toBe('{"a":1,"b":{"c":2,"d":3},"e":[{"x":2,"y":1}]}')
  })

  it('preserves array order and scalar encoding', () => {
    expect(canonicalize({ v: [3, 1, 2] })).toBe('{"v":[3,1,2]}')
    expect(canonicalize({ s: 'x', n: 1.5, z: null })).toBe('{"n":1.5,"s":"x","z":null}')
  })

  it('keeps a JSON-parsed "__proto__" key in the canonical form (cannot smuggle content past the hash)', () => {
    const withProto = JSON.parse('{"a":1,"__proto__":{"smuggled":true}}')
    const without = { a: 1 }
    expect(canonicalize(withProto)).toBe('{"__proto__":{"smuggled":true},"a":1}')
    expect(canonicalize(withProto)).not.toBe(canonicalize(without))
    expect(canonicalize(JSON.parse('{"__proto__":{"x":1}}'))).not.toBe(canonicalize(JSON.parse('{"__proto__":{"x":2}}')))
  })

  it('normalizes -0 to 0 (JSON cannot distinguish them)', () => {
    expect(canonicalize({ v: -0 })).toBe('{"v":0}')
    expect(canonicalize([-0])).toBe('[0]')
  })

  it('refuses values JSON cannot represent faithfully instead of silently changing them', () => {
    for (const bad of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY, undefined, () => 1, 10n, Symbol('s'), new Date(0)]) {
      expect(() => canonicalize({ v: bad }), String(typeof bad)).toThrow(TypeError)
    }
    expect(() => canonicalize([undefined])).toThrow(TypeError)
    expect(() => canonicalize({ nested: { deep: [1, { v: Number.NaN }] } })).toThrow(/nested\.deep\[1\]\.v/)
  })
})

describe('hash tag', () => {
  it('matches the published SHA-256 test vector', async () => {
    await expect(sha256Hex('abc')).resolves.toBe('ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad')
  })

  it('matches independently computed known answers for canonical JSON', async () => {
    await expect(hashResult({ b: [2, 3], a: 1 })).resolves.toBe('efbd0040190fb0871831e606c581f8a66db79d8e2bb836745a70051306956070')
    await expect(hashResult({ b: [1, { d: 2, c: 1 }], a: 'x' })).resolves.toBe(
      '49e7cd01071d3312826d46d6de0396d723ae28ed33557f1665120af2b53f4f70',
    )
  })

  it('is invariant under key-order permutation of a full result and changes with any value', async () => {
    const result = await buildResult()
    const { hash: _hash, ...rest } = result
    const reordered = {
      samples: rest.samples,
      ledger: rest.ledger,
      environment: Object.fromEntries(Object.entries(rest.environment).reverse()),
      metrics: rest.metrics,
      schemaVersion: rest.schemaVersion,
    }
    expect(await hashResult(reordered)).toBe(result.hash)
    expect(await hashResult({ ...rest, environment: { ...rest.environment, dpr: 3 } })).not.toBe(result.hash)
  })

  it('falls back to the pure-JS implementation when crypto.subtle is absent (insecure origin)', async () => {
    const expected = createHash('sha256').update('{"a":1}').digest('hex')
    vi.stubGlobal('crypto', {})
    try {
      expect(globalThis.crypto.subtle).toBeUndefined()
      await expect(sha256Hex('{"a":1}')).resolves.toBe(expected)
      const result = await buildResult()
      await expect(verifyHarnessResult(result)).resolves.toEqual({ ok: true, errors: [] })
    } finally {
      vi.unstubAllGlobals()
    }
  })
})

describe('finalizeResult / validateHarnessResult / verifyHarnessResult basics', () => {
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

  it('detects modification of several different fields through the hash', async () => {
    const result = await buildResult()
    const mutations: Array<[string, (draft: Record<string, any>) => void]> = [
      ['logical cores', (d) => (d.environment.hardware.logicalCores = 4)],
      ['buildCommit', (d) => (d.environment.buildCommit = 'cafebabe')],
      ['note', (d) => (d.environment.note = 'edited')],
      ['os platform', (d) => (d.environment.os.platform = 'Linux')],
      ['backend', (d) => (d.environment.backend = 'webgl')],
      ['cache note', (d) => (d.environment.cache.note = 'edited')],
      ['gpu', (d) => (d.environment.hardware.gpu = 'Other GPU')],
    ]
    for (const [label, mutate] of mutations) {
      const tampered = clone(result) as Record<string, any>
      mutate(tampered)
      const outcome = await verifyHarnessResult(tampered)
      expect(outcome.ok, label).toBe(false)
      expect(outcome.errors[0], label).toMatch(/hash mismatch/)
    }
  })

  it('detects modification of raw samples, metrics and the ledger even when the hash is not recomputed', async () => {
    const result = await buildResult()
    for (const mutate of [
      (d: Record<string, any>) => (d.samples[0].values[0] = 999),
      (d: Record<string, any>) => (d.metrics.op.p50 = 1),
      (d: Record<string, any>) => (d.ledger.main.charged = 7),
    ]) {
      const tampered = clone(result) as Record<string, any>
      mutate(tampered)
      expect((await verifyHarnessResult(tampered)).ok).toBe(false)
    }
  })

  it('rejects invalid fixture ids, profiles and modes', async () => {
    const result = await buildResult()
    for (const bad of [
      { ...result, environment: { ...result.environment, fixtureId: 'F9' } },
      { ...result, environment: { ...result.environment, requestedFixtureId: 'F9' } },
      { ...result, environment: { ...result.environment, profile: 'tablet' as HarnessEnvironment['profile'] } },
      { ...result, environment: { ...result.environment, mode: 'debug' as HarnessEnvironment['mode'] } },
    ]) {
      expect(validateHarnessResult(bad).ok).toBe(false)
    }
  })

  it('rejects malformed samples, metrics and hashes', async () => {
    const result = await buildResult([])
    expect(validateHarnessResult({ ...result, samples: [{ name: 'x' }] }).ok).toBe(false)
    const badMetric: HarnessMetric = { status: 'unavailable' } as HarnessMetric
    expect(validateHarnessResult({ ...result, metrics: { m: badMetric } }).ok).toBe(false)
    const okWithoutP95 = { status: 'ok', p50: 1, max: 2, samples: 3 } as HarnessMetric
    expect(validateHarnessResult({ ...result, metrics: { m: okWithoutP95 } }).ok).toBe(false)
    expect(validateHarnessResult({ ...result, hash: 'deadbeef' }).ok).toBe(false)
    expect(validateHarnessResult(null).ok).toBe(false)
    expect(validateHarnessResult({ schemaVersion: 2 }).ok).toBe(false)
  })

  it('accepts every fixture id F1..F8 for a real (non-placeholder) source', async () => {
    for (const fixtureId of ['F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8'] as const) {
      const result = await buildResult([], { fixtureId, requestedFixtureId: fixtureId })
      expect(validateHarnessResult(result).ok, fixtureId).toBe(true)
    }
  })

  it('finalizeResult refuses to emit a result its own validator would reject', async () => {
    const ledger = new AllocationLedger().accounting()
    await expect(finalizeResult({ environment: makeEnvironment(), samples: [makeSample([1]), makeSample([2])], ledger })).rejects.toThrow(
      /duplicate/,
    )
    await expect(
      finalizeResult({
        environment: makeEnvironment(),
        samples: [makeSample([1])],
        ledger,
        extraMetrics: { op: { status: 'unavailable', reason: 'shadowing a sample' } },
      }),
    ).rejects.toThrow(/op/)
    await expect(
      finalizeResult({ environment: makeEnvironment({ mode: 'release', fixtureId: null, source: { adapter: 'xlsx-placeholder', bytes: 1, rows: 1, cols: 1 } }), samples: [], ledger }),
    ).rejects.toThrow(/release/)
  })
})

// Referenced so the helper stays exercised even if a test above is trimmed.
void rehash
