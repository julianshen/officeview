/**
 * Interactive-viewer real-browser harness — performance result schema (S0.3).
 *
 * This module is the single source of truth for result shape, validation,
 * canonicalization and the integrity hash. It is environment-agnostic (DOM,
 * Bun, jsdom) so the dev page, the bench CLI and unit tests share it.
 *
 * HASH: SHA-256 over canonical JSON (recursively sorted keys). The hash is an
 * INTEGRITY TAG ONLY — it detects accidental modification of a result file.
 * It is NOT a signature: there is no key management, no authenticity claim,
 * and anyone with the file can recompute it. Do not present it as signed
 * evidence.
 *
 * Percentile method: nearest-rank. Values are sorted ascending; the rank is
 * ceil(p/100 * n) clamped to [1, n]; the result is values[rank - 1]. The
 * method is deterministic for any n and stable under repeated runs.
 *
 * Sample-count rules (spec §12.6): p95 is refused below the minimum sample
 * count — 30 for operation samples, 300 for frame samples. A metric that
 * cannot be measured at all is { status: 'unavailable', reason }, never 0.
 */
import type {
  AllocationPool,
  FixtureId,
  HarnessEnvironment,
  HarnessMetric,
  HarnessResult,
  HarnessSample,
  PoolAccounting,
  ResultDraft,
  ValidationOutcome,
} from './types'

export const MIN_OPERATION_SAMPLES = 30
export const MIN_FRAME_SAMPLES = 300

export const FIXTURE_IDS: readonly FixtureId[] = ['F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8']

export type { AllocationPool, HarnessEnvironment, HarnessMetric, HarnessResult, HarnessSample, PoolAccounting, ResultDraft, ValidationOutcome }

/** Nearest-rank percentile; throws on an empty sample set. */
export function percentile(values: readonly number[], p: number): number {
  if (values.length === 0) throw new Error('percentile requires at least one value')
  if (!(p > 0 && p <= 100)) throw new Error(`percentile p must be in (0, 100], got ${p}`)
  const sorted = [...values].sort((a, b) => a - b)
  const rank = Math.min(Math.max(Math.ceil((p / 100) * sorted.length), 1), sorted.length)
  return sorted[rank - 1]
}

export function unavailableMetric(name: string, reason: string): HarnessMetric {
  return { status: 'unavailable', reason: `${name}: ${reason}` }
}

function sampleMinimum(kind: HarnessSample['kind']): number {
  return kind === 'frame' ? MIN_FRAME_SAMPLES : MIN_OPERATION_SAMPLES
}

/** Derive the reportable metric for one sample series. */
export function deriveMetric(sample: HarnessSample): HarnessMetric {
  const values = sample.values
  if (values.length === 0) {
    return unavailableMetric(sample.name, 'no samples recorded during the collection window')
  }
  const sorted = [...values].sort((a, b) => a - b)
  const p50 = percentile(sorted, 50)
  const max = sorted[sorted.length - 1]
  const minimum = sampleMinimum(sample.kind)
  if (values.length < minimum) {
    return {
      status: 'insufficient-samples',
      p50,
      max,
      samples: values.length,
      reason: `${sample.name}: ${values.length} sample(s) is below the ${minimum}-sample minimum for ${sample.kind} metrics; p95 refused`,
    }
  }
  return { status: 'ok', p50, p95: percentile(sorted, 95), max, samples: values.length }
}

/** Canonical JSON: object keys sorted recursively, arrays keep their order. */
export function canonicalize(value: unknown): string {
  return JSON.stringify(sortKeysDeep(value))
}

function sortKeysDeep(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(sortKeysDeep)
  if (value !== null && typeof value === 'object') {
    const source = value as Record<string, unknown>
    const sorted: Record<string, unknown> = {}
    for (const key of Object.keys(source).sort()) sorted[key] = sortKeysDeep(source[key])
    return sorted
  }
  return value
}

export async function sha256Hex(text: string): Promise<string> {
  const subtle = globalThis.crypto?.subtle
  if (!subtle) throw new Error('crypto.subtle is unavailable in this environment')
  const digest = await subtle.digest('SHA-256', new TextEncoder().encode(text))
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, '0')).join('')
}

/** Hash a result draft (or a result object with `hash` stripped) canonically. */
export async function hashResult(result: unknown): Promise<string> {
  return sha256Hex(canonicalize(result))
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function validateMetric(name: string, metric: unknown, errors: string[]): void {
  if (!isPlainObject(metric)) {
    errors.push(`metrics.${name}: must be an object`)
    return
  }
  const status = metric.status
  if (status === 'ok') {
    for (const field of ['p50', 'p95', 'max', 'samples'] as const) {
      if (typeof metric[field] !== 'number' || !Number.isFinite(metric[field])) {
        errors.push(`metrics.${name}: status 'ok' requires finite numeric ${field}`)
      }
    }
    return
  }
  if (status === 'insufficient-samples') {
    for (const field of ['p50', 'max', 'samples', 'reason'] as const) {
      if (metric[field] === undefined) errors.push(`metrics.${name}: status 'insufficient-samples' requires ${field}`)
    }
    if ('p95' in metric) errors.push(`metrics.${name}: insufficient-samples must not report p95`)
    return
  }
  if (status === 'unavailable') {
    if (typeof metric.reason !== 'string' || metric.reason.length === 0) {
      errors.push(`metrics.${name}: status 'unavailable' requires a non-empty reason`)
    }
    return
  }
  errors.push(`metrics.${name}: unknown status ${String(status)}`)
}

/** Runtime validation for ingested result JSON (bench collector, tests). */
export function validateHarnessResult(input: unknown): ValidationOutcome {
  const errors: string[] = []
  if (!isPlainObject(input)) return { ok: false, errors: ['result: must be an object'] }
  if (input.schemaVersion !== 1) errors.push('schemaVersion: must be 1')

  const env = input.environment
  if (!isPlainObject(env)) {
    errors.push('environment: must be an object')
  } else {
    if (typeof env.buildCommit !== 'string') errors.push('environment.buildCommit: must be a string')
    if (!isPlainObject(env.hardware)) errors.push('environment.hardware: must be an object')
    if (!isPlainObject(env.os) || typeof env.os.userAgent !== 'string') errors.push('environment.os: must carry userAgent')
    if (!isPlainObject(env.browser)) errors.push('environment.browser: must be an object')
    else {
      for (const field of ['name', 'version', 'engine', 'userAgent', 'automation'] as const) {
        if (typeof env.browser[field] !== 'string') errors.push(`environment.browser.${field}: must be a string`)
      }
    }
    if (typeof env.dpr !== 'number' || !Number.isFinite(env.dpr)) errors.push('environment.dpr: must be a finite number')
    if (!isPlainObject(env.viewport)) errors.push('environment.viewport: must be an object')
    if (!isPlainObject(env.fonts)) errors.push('environment.fonts: must be an object')
    if (!isPlainObject(env.cache)) errors.push('environment.cache: must be an object')
    if (!isPlainObject(env.source) || typeof env.source.bytes !== 'number') {
      errors.push('environment.source: must carry numeric bytes')
    }
    if (typeof env.backend !== 'string' || env.backend.length === 0) errors.push('environment.backend: must be a non-empty string')
    if (!FIXTURE_IDS.includes(env.fixtureId as FixtureId)) {
      errors.push(`environment.fixtureId: must be one of ${FIXTURE_IDS.join(', ')}`)
    }
    if (env.profile !== 'desktop' && env.profile !== 'mobile') errors.push('environment.profile: must be desktop|mobile')
    if (env.mode !== 'smoke' && env.mode !== 'release') errors.push('environment.mode: must be smoke|release')
    if (typeof env.note !== 'string') errors.push('environment.note: must be a string')
  }

  if (!Array.isArray(input.samples)) {
    errors.push('samples: must be an array')
  } else {
    for (const [index, sample] of input.samples.entries()) {
      if (!isPlainObject(sample)) {
        errors.push(`samples[${index}]: must be an object`)
        continue
      }
      if (typeof sample.name !== 'string' || sample.name.length === 0) errors.push(`samples[${index}].name: required`)
      if (typeof sample.unit !== 'string' || sample.unit.length === 0) errors.push(`samples[${index}].unit: required`)
      if (!Array.isArray(sample.values) || sample.values.some((v) => typeof v !== 'number' || !Number.isFinite(v))) {
        errors.push(`samples[${index}].values: must be an array of finite numbers`)
      }
      if (typeof sample.warmupExcluded !== 'boolean') errors.push(`samples[${index}].warmupExcluded: must be boolean`)
      if (sample.kind !== 'operation' && sample.kind !== 'frame') errors.push(`samples[${index}].kind: must be operation|frame`)
    }
  }

  if (!isPlainObject(input.metrics)) {
    errors.push('metrics: must be an object')
  } else {
    for (const [name, metric] of Object.entries(input.metrics)) validateMetric(name, metric, errors)
  }

  if (!isPlainObject(input.ledger)) {
    errors.push('ledger: must be an object')
  } else {
    for (const pool of ['main', 'worker', 'bitmap'] as const) {
      const accounting = (input.ledger as Record<string, unknown>)[pool]
      if (!isPlainObject(accounting)) {
        errors.push(`ledger.${pool}: must be an object`)
        continue
      }
      for (const field of ['retained', 'peak', 'charged'] as const) {
        if (typeof accounting[field] !== 'number' || !Number.isFinite(accounting[field])) {
          errors.push(`ledger.${pool}.${field}: must be a finite number`)
        }
      }
    }
  }

  if (typeof input.hash !== 'string' || !/^[0-9a-f]{64}$/.test(input.hash)) {
    errors.push('hash: must be a 64-character lowercase hex SHA-256 tag')
  }

  if (errors.length > 0) return { ok: false, errors }
  return { ok: true, result: input as unknown as HarnessResult }
}

/** Validate then recompute the integrity hash; any mismatch fails closed. */
export async function verifyHarnessResult(input: unknown): Promise<{ ok: boolean; errors: string[] }> {
  const outcome = validateHarnessResult(input)
  if (!outcome.ok) return { ok: false, errors: outcome.errors }
  const { hash, ...rest } = outcome.result
  const recomputed = await hashResult(rest)
  if (recomputed !== hash) {
    return { ok: false, errors: [`hash mismatch: stored ${hash}, recomputed ${recomputed}`] }
  }
  return { ok: true, errors: [] }
}

/** Build a complete result: derive metrics, attach the ledger, compute the hash. */
export async function finalizeResult(draft: ResultDraft): Promise<HarnessResult> {
  const metrics: Record<string, HarnessMetric> = {}
  for (const sample of draft.samples) metrics[sample.name] = deriveMetric(sample)
  for (const [name, metric] of Object.entries(draft.extraMetrics ?? {})) metrics[name] = metric
  const withoutHash = {
    schemaVersion: 1 as const,
    environment: draft.environment,
    samples: draft.samples,
    metrics,
    ledger: draft.ledger,
  }
  const hash = await hashResult(withoutHash)
  return { ...withoutHash, hash }
}
