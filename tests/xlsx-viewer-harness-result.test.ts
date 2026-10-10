/**
 * Ingest honesty invariants (finding 5). The integrity hash is recomputable by
 * anyone, so a result that passes the hash must still be rejected when it is
 * internally dishonest. Every case here models a forger who EDITS the result
 * and RECOMPUTES the hash.
 */
import { describe, expect, it } from 'vitest'
import {
  MAX_SAMPLE_SERIES,
  MAX_SAMPLE_VALUES,
  checkProfileConformance,
  deriveEvidence,
  verifyHarnessResult,
} from '../src/debug/harness/result'
import { buildResult, makeEnvironment, makeSample, rehash } from './helpers/harnessResult'

type Mutator = (draft: Record<string, any>) => void

async function expectRejected(mutate: Mutator, pattern: RegExp, base = buildResult()): Promise<void> {
  const forged = await rehash(await base, mutate)
  const outcome = await verifyHarnessResult(forged)
  expect(outcome.ok, `forged result was accepted: ${JSON.stringify(outcome)}`).toBe(false)
  expect(outcome.errors.join('\n')).toMatch(pattern)
}

describe('metrics must be exactly derivable from the raw samples', () => {
  it('rejects ok-with-p95 when only 3 samples exist', async () => {
    await expectRejected(
      (d) => {
        d.samples[0].values = [1, 2, 3]
        d.metrics.op = { status: 'ok', p50: 2, p95: 3, max: 3, samples: 3 }
      },
      /metrics\.op.*derived/,
    )
  })

  it('rejects a p50 that differs from the samples by even 1', async () => {
    await expectRejected((d) => (d.metrics.op.p50 += 1), /metrics\.op.*derived/)
  })

  it('rejects p95 < p50', async () => {
    await expectRejected((d) => ((d.metrics.op.p95 = 1), (d.metrics.op.p50 = 20)), /metrics\.op/)
  })

  it('rejects negative or fractional metric sample counts', async () => {
    await expectRejected((d) => (d.metrics.op.samples = -5), /metrics\.op\.samples/)
    await expectRejected((d) => (d.metrics.op.samples = 1.5), /metrics\.op\.samples/)
  })

  it('rejects a metric that is missing for a sample series', async () => {
    await expectRejected((d) => delete d.metrics.op, /metrics\.op.*missing/)
  })

  it('rejects an unavailable metric that carries numeric fields', async () => {
    const base = buildResult([])
    await expectRejected(
      (d) => (d.metrics['viewport-frame'] = { status: 'unavailable', reason: 'x', p50: 0, p95: 0 }),
      /metrics\.viewport-frame.*(unexpected|unknown)/i,
      base,
    )
  })

  it('rejects unavailable with an empty or non-string reason', async () => {
    await expectRejected((d) => (d.metrics.m = { status: 'unavailable', reason: '' }), /metrics\.m/, buildResult([]))
    await expectRejected((d) => (d.metrics.m = { status: 'unavailable', reason: 7 }), /metrics\.m/, buildResult([]))
  })

  it('rejects insufficient-samples with non-numeric fields', async () => {
    const base = buildResult([makeSample([1, 2, 3])])
    await expectRejected((d) => (d.metrics.op = { status: 'insufficient-samples', p50: 'a', max: null, samples: '3', reason: 5 }), /metrics\.op/, base)
  })

  it('rejects a non-unavailable metric that no sample backs', async () => {
    await expectRejected(
      (d) => (d.metrics.ghost = { status: 'ok', p50: 1, p95: 2, max: 3, samples: 99 }),
      /metrics\.ghost.*no backing sample/,
    )
  })

  it('rejects unknown metric statuses and unexpected extra fields on an honest metric', async () => {
    await expectRejected((d) => (d.metrics.op.status = 'great'), /metrics\.op/)
    await expectRejected((d) => (d.metrics.op.bonus = 1), /metrics\.op.*unexpected/i)
  })

  it('accepts none-observed only for an empty event series', async () => {
    const base = await buildResult([makeSample([], 'event', 'longtask-duration')])
    expect(base.metrics['longtask-duration'].status).toBe('none-observed')
    await expect(verifyHarnessResult(base)).resolves.toEqual({ ok: true, errors: [] })
    await expectRejected(
      (d) => (d.metrics.op = { status: 'none-observed', samples: 0, reason: 'claimed' }),
      /metrics\.op.*derived/,
    )
  })
})

describe('numeric and size sanity', () => {
  it('rejects negative, non-finite-looking or wrongly typed sample values', async () => {
    await expectRejected((d) => (d.samples[0].values[0] = -1), /samples\[0\]\.values/)
    await expectRejected((d) => (d.samples[0].values[0] = '5'), /samples\[0\]\.values/)
  })

  it('requires warmupExcluded to be a non-negative integer count', async () => {
    await expectRejected((d) => (d.samples[0].warmupExcluded = true), /warmupExcluded/)
    await expectRejected((d) => (d.samples[0].warmupExcluded = -1), /warmupExcluded/)
    await expectRejected((d) => (d.samples[0].warmupExcluded = 1.5), /warmupExcluded/)
  })

  it('rejects negative, fractional or inconsistent ledger byte counts', async () => {
    await expectRejected((d) => (d.ledger.main.retained = -1), /ledger\.main\.retained/)
    await expectRejected((d) => (d.ledger.bitmap.peak = 1.5), /ledger\.bitmap\.peak/)
    await expectRejected((d) => ((d.ledger.worker.retained = 10), (d.ledger.worker.peak = 5), (d.ledger.worker.charged = 20)), /retained.*peak/)
    await expectRejected((d) => ((d.ledger.worker.peak = 10), (d.ledger.worker.charged = 5)), /peak.*charged/)
  })

  it('rejects negative, fractional or oversized source bytes and bad row/col counts', async () => {
    await expectRejected((d) => (d.environment.source.bytes = -1), /source\.bytes/)
    await expectRejected((d) => (d.environment.source.bytes = 2.5), /source\.bytes/)
    await expectRejected((d) => (d.environment.source.rows = -3), /source\.rows/)
    await expectRejected((d) => (d.environment.source.cols = 'x'), /source\.cols/)
  })

  it('caps series count, values per series and duplicate series names', async () => {
    await expectRejected(
      (d) => (d.samples[0].values = Array.from({ length: MAX_SAMPLE_VALUES + 1 }, () => 1)),
      /at most/,
    )
    await expectRejected(
      (d) => {
        d.samples = Array.from({ length: MAX_SAMPLE_SERIES + 1 }, (_, i) => ({ name: `s${i}`, unit: 'ms', kind: 'operation', warmupExcluded: 0, values: [1] }))
        d.metrics = {}
      },
      /at most/,
    )
    await expectRejected((d) => d.samples.push({ ...d.samples[0] }), /duplicate/)
  })

  it('caps free-text lengths', async () => {
    await expectRejected((d) => (d.environment.note = 'x'.repeat(5000)), /environment\.note/)
    await expectRejected((d) => (d.environment.browser.userAgent = 'x'.repeat(5000)), /userAgent/)
  })
})

describe('prototype pollution and unknown keys', () => {
  it('rejects a JSON-parsed __proto__ key anywhere, even with a recomputed hash', async () => {
    const forged = JSON.parse(
      JSON.stringify(await buildResult()).replace('"schemaVersion":1', '"schemaVersion":1,"__proto__":{"x":1}'),
    )
    expect(Object.prototype.hasOwnProperty.call(forged, '__proto__')).toBe(true)
    const outcome = await verifyHarnessResult(forged)
    expect(outcome.ok).toBe(false)
    expect(outcome.errors.join()).toMatch(/__proto__/)

    const nested = JSON.parse(JSON.stringify(await buildResult()).replace('"cache":{', '"cache":{"__proto__":{"y":1},'))
    expect((await verifyHarnessResult(nested)).errors.join()).toMatch(/__proto__/)
  })

  it('rejects unknown keys at the root, in the environment and in samples', async () => {
    await expectRejected((d) => (d.extra = 1), /unexpected key.*extra/i)
    await expectRejected((d) => (d.environment.extra = 1), /unexpected key.*extra/i)
    await expectRejected((d) => (d.samples[0].extra = 1), /unexpected key.*extra/i)
    await expectRejected((d) => (d.ledger.main.extra = 1), /unexpected key.*extra/i)
  })

  it('rejects absurdly deep nesting without overflowing the stack', async () => {
    const deep: Record<string, unknown> = {}
    let cursor = deep
    for (let i = 0; i < 5000; i++) cursor = (cursor.n = {}) as Record<string, unknown>
    const forged = await rehash(await buildResult(), (d) => (d.note = 'x'))
    const outcome = await verifyHarnessResult({ ...(forged as object), deep })
    expect(outcome.ok).toBe(false)
  })
})

describe('environment consistency', () => {
  it('rejects browser.name safari with automation playwright', async () => {
    await expectRejected(
      (d) => {
        d.environment.browser.name = 'safari'
        d.environment.browser.engine = 'webkit'
        d.environment.browser.automation = 'playwright'
        d.environment.evidence = 'automated-browser'
      },
      /safari.*automation/i,
    )
  })

  it('rejects engine/name mismatches and webkit-playwright without playwright automation', async () => {
    await expectRejected((d) => (d.environment.browser.engine = 'webkit'), /engine/)
    await expectRejected(
      (d) => {
        d.environment.browser.name = 'webkit-playwright'
        d.environment.browser.engine = 'webkit'
        d.environment.evidence = 'emulated-webkit'
      },
      /webkit-playwright.*playwright/i,
    )
  })

  it('rejects an evidence label that disagrees with what the environment proves', async () => {
    await expectRejected((d) => (d.environment.evidence = 'emulated-webkit'), /evidence/)
    await expectRejected((d) => (d.environment.evidence = 'scaffold-placeholder'), /evidence/)
  })

  it('placeholder content never carries a spec fixture id and is labeled scaffold-placeholder', async () => {
    const placeholder = {
      source: { adapter: 'xlsx-placeholder', bytes: 100, rows: 120, cols: 10 },
      fixtureId: null,
      requestedFixtureId: 'F1' as const,
    }
    const ok = await buildResult([], placeholder)
    expect(ok.environment.evidence).toBe('scaffold-placeholder')
    await expect(verifyHarnessResult(ok)).resolves.toEqual({ ok: true, errors: [] })
    await expectRejected((d) => (d.environment.fixtureId = 'F1'), /fixtureId/, buildResult([], placeholder))
    await expectRejected((d) => (d.environment.evidence = 'device'), /evidence/, buildResult([], placeholder))
    await expectRejected((d) => (d.environment.fixtureId = null), /fixtureId/) // a real adapter needs its fixture id
  })

  it('refuses release-mode results built on placeholder content', async () => {
    await expectRejected(
      (d) => (d.environment.mode = 'release'),
      /release.*placeholder/i,
      buildResult([], { source: { adapter: 'xlsx-placeholder', bytes: 100, rows: 1, cols: 1 }, fixtureId: null }),
    )
  })

  it('requires a reason when the OS platform is unknown', async () => {
    await expectRejected((d) => (d.environment.os.platform = 'unknown'), /platformNote/)
    const withReason = await buildResult([], { os: { platform: 'unknown', platformNote: 'userAgentData unavailable on insecure origin', userAgent: 'x' } })
    await expect(verifyHarnessResult(withReason)).resolves.toEqual({ ok: true, errors: [] })
  })

  it('profile conformance is recomputed, not trusted', async () => {
    await expectRejected((d) => (d.environment.viewport.width = 800), /profileConformance/)
    await expectRejected((d) => (d.environment.profileConformance.conforms = false), /profileConformance/)
    const nonConforming = await buildResult([], { viewport: { width: 800, height: 600 }, dpr: 1 })
    expect(nonConforming.environment.profileConformance.conforms).toBe(false)
    await expect(verifyHarnessResult(nonConforming)).resolves.toEqual({ ok: true, errors: [] })
  })

  it('checks operator metadata types', async () => {
    await expectRejected((d) => (d.environment.operator.deviceModel = 5), /operator\.deviceModel/)
  })
})

describe('checkProfileConformance (spec 12.2)', () => {
  it('accepts the desktop 1440x900 @2x and mobile 1024x768 @2x profiles', () => {
    expect(checkProfileConformance('desktop', { width: 1440, height: 900 }, 2)).toEqual({ conforms: true, deviations: [] })
    expect(checkProfileConformance('mobile', { width: 1024, height: 768 }, 2)).toEqual({ conforms: true, deviations: [] })
  })
  it('lists each deviation', () => {
    const outcome = checkProfileConformance('desktop', { width: 1024, height: 768 }, 1)
    expect(outcome.conforms).toBe(false)
    expect(outcome.deviations).toEqual([expect.stringContaining('1440x900'), expect.stringContaining('DPR')])
  })
  it('does not cross-accept profiles', () => {
    expect(checkProfileConformance('mobile', { width: 1440, height: 900 }, 2).conforms).toBe(false)
  })
})

describe('deriveEvidence', () => {
  const env = (overrides: Parameters<typeof makeEnvironment>[0]) => makeEnvironment(overrides)
  it('placeholder wins over every other label', () => {
    expect(deriveEvidence(env({ source: { adapter: 'xlsx-placeholder', bytes: 1, rows: 1, cols: 1 }, fixtureId: null }))).toBe('scaffold-placeholder')
  })
  it('Playwright WebKit is emulated-webkit', () => {
    expect(
      deriveEvidence(env({ browser: { name: 'webkit-playwright', version: '1', engine: 'webkit', userAgent: 'x', automation: 'playwright' } })),
    ).toBe('emulated-webkit')
  })
  it('automated Chromium on real content is automated-browser; an unautomated browser is device', () => {
    expect(deriveEvidence(env({ browser: { name: 'chromium', version: '1', engine: 'blink', userAgent: 'x', automation: 'playwright' } }))).toBe('automated-browser')
    expect(deriveEvidence(env({}))).toBe('device')
  })
})
