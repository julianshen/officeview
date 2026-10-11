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
 * evidence. Because a forger can recompute it, ingest (`verifyHarnessResult`)
 * does not stop at the hash: every derived metric must equal what the raw
 * samples imply, every count must be a sane integer, and the labels
 * (browser, evidence, profile conformance) must agree with each other.
 *
 * Canonical JSON refuses what JSON cannot carry faithfully (undefined, NaN,
 * Infinity, bigint, symbols, non-plain objects) and normalizes -0 to 0 rather
 * than silently dropping or rewriting it. Objects are rebuilt with
 * Object.fromEntries so a JSON-parsed "__proto__" key stays an own property
 * (assignment would have dropped it); the validator additionally rejects
 * "__proto__" keys outright.
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
  BrowserName,
  EvidenceLabel,
  FixtureId,
  HarnessEnvironment,
  HarnessMetric,
  HarnessProfile,
  HarnessResult,
  HarnessSample,
  PoolAccounting,
  ProfileConformance,
  ResultDraft,
  ValidationOutcome,
} from './types'
import { sha256HexJs } from './sha256'

export const MIN_OPERATION_SAMPLES = 30
export const MIN_FRAME_SAMPLES = 300

/** Ingest caps: a result is a few KiB of numbers, never megabytes. */
export const MAX_SAMPLE_SERIES = 64
export const MAX_SAMPLE_VALUES = 5000
export const MAX_TOTAL_SAMPLE_VALUES = 20_000
const MAX_METRICS = 128
const MAX_NESTING_DEPTH = 16
const MAX_TEXT = 1024
const MAX_UA = 512

/** The single canonical fixture-id list (spec §12.5). */
export const FIXTURE_IDS: readonly FixtureId[] = ['F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8']

/** Adapters whose content is NOT a spec fixture (no fixture id, never release). */
export const PLACEHOLDER_ADAPTERS: readonly string[] = ['xlsx-placeholder', 'none']

/** Spec §12.2 acceptance profiles: CSS viewport and DPR. */
export const PROFILE_EXPECTATIONS: Record<HarnessProfile, { width: number; height: number; dpr: number }> = {
  desktop: { width: 1440, height: 900, dpr: 2 },
  mobile: { width: 1024, height: 768, dpr: 2 },
}

const BROWSER_ENGINE: Record<BrowserName, readonly string[]> = {
  chromium: ['blink'],
  safari: ['webkit'],
  webkit: ['webkit'],
  'webkit-playwright': ['webkit'],
  firefox: ['gecko'],
  unknown: ['unknown'],
}

export type { AllocationPool, HarnessEnvironment, HarnessMetric, HarnessResult, HarnessSample, PoolAccounting, ResultDraft, ValidationOutcome }

/** Nearest-rank percentile; throws on an empty sample set or bad arguments. */
export function percentile(values: readonly number[], p: number): number {
  if (values.length === 0) throw new Error('percentile requires at least one value')
  if (!(p > 0 && p <= 100)) throw new Error(`percentile p must be in (0, 100], got ${p}`)
  if (values.some((value) => !Number.isFinite(value))) throw new Error('percentile requires finite values')
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
    if (sample.kind === 'event') {
      return { status: 'none-observed', samples: 0, reason: `${sample.name}: collection ran and observed no events in the window` }
    }
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

/**
 * Canonical JSON: object keys sorted recursively, arrays keep their order.
 * Throws on values JSON cannot represent faithfully; -0 becomes 0.
 */
export function canonicalize(value: unknown): string {
  return JSON.stringify(toCanonical(value, '$'))
}

function toCanonical(value: unknown, path: string): unknown {
  if (value === null) return null
  switch (typeof value) {
    case 'string':
    case 'boolean':
      return value
    case 'number':
      if (!Number.isFinite(value)) throw new TypeError(`canonicalize: ${path} is ${value}, which JSON cannot represent`)
      return Object.is(value, -0) ? 0 : value
    case 'object': {
      if (Array.isArray(value)) return value.map((item, index) => toCanonical(item, `${path}[${index}]`))
      const prototype = Object.getPrototypeOf(value)
      if (prototype !== Object.prototype && prototype !== null) {
        throw new TypeError(`canonicalize: ${path} is a non-plain object`)
      }
      const source = value as Record<string, unknown>
      // fromEntries defines own properties, so a "__proto__" key survives.
      return Object.fromEntries(
        Object.keys(source)
          .sort()
          .map((key) => [key, toCanonical(source[key], path === '$' ? key : `${path}.${key}`)]),
      )
    }
    default:
      throw new TypeError(`canonicalize: ${path} is ${typeof value}, which JSON cannot represent`)
  }
}

/**
 * SHA-256 hex of UTF-8 text. Uses crypto.subtle when the context has it and
 * the pure-JS implementation otherwise (insecure origins such as the LAN URL
 * of the manual device path have no crypto.subtle).
 */
export async function sha256Hex(text: string): Promise<string> {
  const subtle = globalThis.crypto?.subtle
  if (subtle) {
    try {
      const digest = await subtle.digest('SHA-256', new TextEncoder().encode(text))
      return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, '0')).join('')
    } catch {
      // fall through to the dependency-free implementation
    }
  }
  return sha256HexJs(text)
}

/** Hash a result draft (or a result object with `hash` stripped) canonically. */
export async function hashResult(result: unknown): Promise<string> {
  return sha256Hex(canonicalize(result))
}

/** Spec §12.2 profile check; a mismatch is recorded, not hidden. */
export function checkProfileConformance(
  profile: HarnessProfile,
  viewport: { width: number; height: number },
  dpr: number,
): ProfileConformance {
  const expected = PROFILE_EXPECTATIONS[profile]
  const deviations: string[] = []
  if (viewport.width !== expected.width || viewport.height !== expected.height) {
    deviations.push(`viewport ${viewport.width}x${viewport.height} differs from the ${profile} profile ${expected.width}x${expected.height} CSS px`)
  }
  if (dpr !== expected.dpr) deviations.push(`DPR ${dpr} differs from the ${profile} profile DPR ${expected.dpr}`)
  return { conforms: deviations.length === 0, deviations }
}

/** What kind of evidence this environment can ever count as. */
export function deriveEvidence(env: Pick<HarnessEnvironment, 'source' | 'browser'>): EvidenceLabel {
  if (PLACEHOLDER_ADAPTERS.includes(env.source.adapter)) return 'scaffold-placeholder'
  if (env.browser.name === 'webkit-playwright') return 'emulated-webkit'
  if (env.browser.automation !== 'none') return 'automated-browser'
  return 'device'
}

type Bag = Record<string, unknown>

function isPlainObject(value: unknown): value is Bag {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function isCount(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value)
}

function boundedString(value: unknown, max: number, min = 0): value is string {
  return typeof value === 'string' && value.length >= min && value.length <= max
}

function exactKeys(value: Bag, allowed: readonly string[], path: string, errors: string[]): void {
  for (const key of Object.keys(value)) {
    if (!allowed.includes(key)) errors.push(`${path}: unexpected key ${JSON.stringify(key)}`)
  }
}

/** Reject "__proto__" own keys and pathological nesting before anything else. */
function scanStructure(value: unknown, path: string, depth: number, errors: string[]): void {
  if (depth > MAX_NESTING_DEPTH) {
    errors.push(`${path}: nesting deeper than ${MAX_NESTING_DEPTH} levels`)
    return
  }
  if (Array.isArray(value)) {
    value.forEach((item, index) => scanStructure(item, `${path}[${index}]`, depth + 1, errors))
  } else if (isPlainObject(value)) {
    for (const key of Object.keys(value)) {
      if (key === '__proto__') errors.push(`${path}: forbidden key "__proto__"`)
      scanStructure(value[key], `${path}.${key}`, depth + 1, errors)
    }
  }
}

const METRIC_NAME = /^[A-Za-z0-9][A-Za-z0-9._:/ -]{0,127}$/
const UNIT = /^[A-Za-z%/]{1,16}$/
const COMMIT = /^(unknown|[0-9a-f]{7,64}(-dirty)?)$/

function validateMetricShape(name: string, metric: unknown, errors: string[]): void {
  const path = `metrics.${name}`
  if (!isPlainObject(metric)) {
    errors.push(`${path}: must be an object`)
    return
  }
  switch (metric.status) {
    case 'ok': {
      exactKeys(metric, ['status', 'p50', 'p95', 'max', 'samples'], path, errors)
      for (const field of ['p50', 'p95', 'max'] as const) {
        if (!isFiniteNumber(metric[field]) || (metric[field] as number) < 0) errors.push(`${path}.${field}: must be a finite non-negative number`)
      }
      if (!isCount(metric.samples) || metric.samples === 0) errors.push(`${path}.samples: must be a positive integer`)
      if (isFiniteNumber(metric.p50) && isFiniteNumber(metric.p95) && isFiniteNumber(metric.max) && !(metric.p50 <= metric.p95 && metric.p95 <= metric.max)) {
        errors.push(`${path}: requires p50 <= p95 <= max`)
      }
      return
    }
    case 'insufficient-samples': {
      exactKeys(metric, ['status', 'p50', 'max', 'samples', 'reason'], path, errors)
      for (const field of ['p50', 'max'] as const) {
        if (!isFiniteNumber(metric[field]) || (metric[field] as number) < 0) errors.push(`${path}.${field}: must be a finite non-negative number`)
      }
      if (!isCount(metric.samples) || metric.samples === 0) errors.push(`${path}.samples: must be a positive integer`)
      if (!boundedString(metric.reason, MAX_TEXT, 1)) errors.push(`${path}.reason: must be a non-empty string`)
      if (isFiniteNumber(metric.p50) && isFiniteNumber(metric.max) && metric.p50 > metric.max) errors.push(`${path}: requires p50 <= max`)
      return
    }
    case 'none-observed': {
      exactKeys(metric, ['status', 'samples', 'reason'], path, errors)
      if (metric.samples !== 0) errors.push(`${path}.samples: must be 0`)
      if (!boundedString(metric.reason, MAX_TEXT, 1)) errors.push(`${path}.reason: must be a non-empty string`)
      return
    }
    case 'unavailable': {
      exactKeys(metric, ['status', 'reason'], path, errors)
      if (!boundedString(metric.reason, MAX_TEXT, 1)) errors.push(`${path}.reason: must be a non-empty string`)
      return
    }
    default:
      errors.push(`${path}: unknown status ${JSON.stringify(metric.status)}`)
  }
}

function validateEnvironment(env: unknown, errors: string[]): void {
  if (!isPlainObject(env)) {
    errors.push('environment: must be an object')
    return
  }
  exactKeys(
    env,
    ['buildCommit', 'hardware', 'os', 'browser', 'dpr', 'viewport', 'profileConformance', 'fonts', 'cache', 'source', 'backend', 'requestedFixtureId', 'fixtureId', 'evidence', 'operator', 'profile', 'mode', 'note', 'scenario'],
    'environment',
    errors,
  )
  if (env.scenario !== undefined) {
    const scenario = env.scenario
    if (!isPlainObject(scenario)) errors.push('environment.scenario: must be an object when present')
    else {
      exactKeys(scenario, ['id', 'kind', 'keys', 'rows', 'admitted'], 'environment.scenario', errors)
      if (typeof scenario.admitted !== 'boolean') errors.push('environment.scenario.admitted: must be boolean')
      if (!boundedString(scenario.id, 64, 1)) errors.push('environment.scenario.id: must be a non-empty string')
      if (scenario.kind !== 'sort' && scenario.kind !== 'filter' && scenario.kind !== 'value-index') errors.push('environment.scenario.kind: must be sort|filter|value-index')
      if (!Array.isArray(scenario.keys) || scenario.keys.length > 8 || !scenario.keys.every(isCount)) errors.push('environment.scenario.keys: must be at most 8 column indexes')
      if (!isCount(scenario.rows) || scenario.rows === 0) errors.push('environment.scenario.rows: must be a positive integer')
    }
  }
  if (typeof env.buildCommit !== 'string' || !COMMIT.test(env.buildCommit)) errors.push('environment.buildCommit: must be "unknown" or a hex commit id, optionally "-dirty"')

  const hardware = env.hardware
  if (!isPlainObject(hardware)) errors.push('environment.hardware: must be an object')
  else {
    exactKeys(hardware, ['logicalCores', 'deviceMemoryGiB', 'gpu'], 'environment.hardware', errors)
    if (hardware.logicalCores !== null && !(isCount(hardware.logicalCores) && hardware.logicalCores > 0 && hardware.logicalCores <= 4096)) errors.push('environment.hardware.logicalCores: must be null or a positive integer')
    if (hardware.deviceMemoryGiB !== null && !(isFiniteNumber(hardware.deviceMemoryGiB) && hardware.deviceMemoryGiB > 0 && hardware.deviceMemoryGiB <= 65536)) errors.push('environment.hardware.deviceMemoryGiB: must be null or a positive number')
    if (hardware.gpu !== null && !boundedString(hardware.gpu, 256)) errors.push('environment.hardware.gpu: must be null or a string')
  }

  const os = env.os
  if (!isPlainObject(os)) errors.push('environment.os: must be an object')
  else {
    exactKeys(os, ['platform', 'platformNote', 'userAgent'], 'environment.os', errors)
    if (!boundedString(os.platform, 64, 1)) errors.push('environment.os.platform: must be a non-empty string')
    if (os.platformNote !== null && !boundedString(os.platformNote, MAX_TEXT)) errors.push('environment.os.platformNote: must be null or a string')
    if (!boundedString(os.userAgent, MAX_UA)) errors.push('environment.os.userAgent: must be a string')
    if (os.platform === 'unknown' && !boundedString(os.platformNote, MAX_TEXT, 1)) {
      errors.push('environment.os.platformNote: a reason is required when platform is "unknown"')
    }
  }

  const browser = env.browser
  if (!isPlainObject(browser)) errors.push('environment.browser: must be an object')
  else {
    exactKeys(browser, ['name', 'version', 'engine', 'userAgent', 'automation'], 'environment.browser', errors)
    const names = Object.keys(BROWSER_ENGINE)
    if (typeof browser.name !== 'string' || !names.includes(browser.name)) errors.push(`environment.browser.name: must be one of ${names.join(', ')}`)
    if (!boundedString(browser.version, 64, 1)) errors.push('environment.browser.version: must be a non-empty string')
    if (!boundedString(browser.userAgent, MAX_UA)) errors.push('environment.browser.userAgent: must be a string')
    if (browser.automation !== 'playwright' && browser.automation !== 'webdriver' && browser.automation !== 'none') errors.push('environment.browser.automation: must be playwright|webdriver|none')
    if (typeof browser.name === 'string' && names.includes(browser.name)) {
      if (!BROWSER_ENGINE[browser.name as BrowserName].includes(browser.engine as string)) {
        errors.push(`environment.browser.engine: ${String(browser.engine)} is inconsistent with browser ${browser.name}`)
      }
      if (browser.name === 'safari' && browser.automation !== 'none') errors.push('environment.browser: a safari label requires automation "none" (automated WebKit is never Safari)')
      if (browser.name === 'webkit-playwright' && browser.automation !== 'playwright') errors.push('environment.browser: webkit-playwright requires automation "playwright"')
    }
  }

  if (!isFiniteNumber(env.dpr) || env.dpr <= 0 || env.dpr > 16) errors.push('environment.dpr: must be a finite number in (0, 16]')
  const viewport = env.viewport
  if (!isPlainObject(viewport)) errors.push('environment.viewport: must be an object')
  else {
    exactKeys(viewport, ['width', 'height'], 'environment.viewport', errors)
    for (const field of ['width', 'height'] as const) {
      if (!isCount(viewport[field]) || viewport[field] === 0 || (viewport[field] as number) > 32768) errors.push(`environment.viewport.${field}: must be a positive integer`)
    }
  }

  const fonts = env.fonts
  if (!isPlainObject(fonts)) errors.push('environment.fonts: must be an object')
  else {
    exactKeys(fonts, ['status', 'resolved'], 'environment.fonts', errors)
    if (fonts.status !== 'ready' && fonts.status !== 'loading' && fonts.status !== 'unavailable') errors.push('environment.fonts.status: must be ready|loading|unavailable')
    if (typeof fonts.resolved !== 'boolean') errors.push('environment.fonts.resolved: must be boolean')
    else if (fonts.resolved !== (fonts.status === 'ready')) errors.push('environment.fonts: resolved must be true exactly when status is ready')
  }

  const cache = env.cache
  if (!isPlainObject(cache)) errors.push('environment.cache: must be an object')
  else {
    exactKeys(cache, ['state', 'note'], 'environment.cache', errors)
    if (cache.state !== 'cold' && cache.state !== 'warm') errors.push('environment.cache.state: must be cold|warm')
    if (!boundedString(cache.note, MAX_TEXT)) errors.push('environment.cache.note: must be a string')
  }

  const source = env.source
  if (!isPlainObject(source)) errors.push('environment.source: must be an object')
  else {
    exactKeys(source, ['adapter', 'bytes', 'rows', 'cols'], 'environment.source', errors)
    if (!boundedString(source.adapter, 64, 1)) errors.push('environment.source.adapter: must be a non-empty string')
    if (!isCount(source.bytes)) errors.push('environment.source.bytes: must be a non-negative integer')
    for (const field of ['rows', 'cols'] as const) {
      if (source[field] !== null && !isCount(source[field])) errors.push(`environment.source.${field}: must be null or a non-negative integer`)
    }
  }

  if (!boundedString(env.backend, 64, 1)) errors.push('environment.backend: must be a non-empty string')
  if (!FIXTURE_IDS.includes(env.requestedFixtureId as FixtureId)) errors.push(`environment.requestedFixtureId: must be one of ${FIXTURE_IDS.join(', ')}`)
  if (env.fixtureId !== null && !FIXTURE_IDS.includes(env.fixtureId as FixtureId)) errors.push(`environment.fixtureId: must be null or one of ${FIXTURE_IDS.join(', ')}`)
  if (env.profile !== 'desktop' && env.profile !== 'mobile') errors.push('environment.profile: must be desktop|mobile')
  if (env.mode !== 'smoke' && env.mode !== 'release') errors.push('environment.mode: must be smoke|release')
  if (!boundedString(env.note, MAX_TEXT)) errors.push('environment.note: must be a string')

  const operator = env.operator
  if (!isPlainObject(operator)) errors.push('environment.operator: must be an object')
  else {
    exactKeys(operator, ['deviceModel', 'osVersion', 'thermalState', 'batteryState'], 'environment.operator', errors)
    for (const field of ['deviceModel', 'osVersion', 'thermalState', 'batteryState'] as const) {
      if (operator[field] !== null && !boundedString(operator[field], 128)) errors.push(`environment.operator.${field}: must be null or a string`)
    }
  }

  const conformance = env.profileConformance
  if (!isPlainObject(conformance)) errors.push('environment.profileConformance: must be an object')
  else exactKeys(conformance, ['conforms', 'deviations'], 'environment.profileConformance', errors)

  // Cross-field facts: derived values are recomputed, never trusted.
  if (!isPlainObject(source) || !isPlainObject(browser) || typeof source.adapter !== 'string') return
  const placeholder = PLACEHOLDER_ADAPTERS.includes(source.adapter)
  if (placeholder && env.fixtureId !== null) errors.push(`environment.fixtureId: placeholder content (adapter ${source.adapter}) must not carry a spec fixture id`)
  if (!placeholder && env.fixtureId === null) errors.push('environment.fixtureId: real fixture content requires a spec fixture id')
  if (placeholder && env.mode === 'release') errors.push('environment.mode: release results cannot be built on placeholder content')
  if (browser.automation === 'none' || browser.automation === 'playwright' || browser.automation === 'webdriver') {
    if (typeof browser.name === 'string') {
      const expected = deriveEvidence({ source: source as HarnessEnvironment['source'], browser: browser as HarnessEnvironment['browser'] })
      if (env.evidence !== expected) errors.push(`environment.evidence: ${JSON.stringify(env.evidence)} does not match the derived label ${JSON.stringify(expected)}`)
    }
  }
  if (isPlainObject(viewport) && isCount(viewport.width) && isCount(viewport.height) && isFiniteNumber(env.dpr) && (env.profile === 'desktop' || env.profile === 'mobile') && isPlainObject(conformance)) {
    const expected = checkProfileConformance(env.profile, { width: viewport.width, height: viewport.height }, env.dpr)
    if (canonicalize(conformance) !== canonicalize(expected)) {
      errors.push(`environment.profileConformance: does not match the value recomputed from profile, viewport and DPR (${JSON.stringify(expected)})`)
    }
  }
}

function validateSamples(samples: unknown, errors: string[]): HarnessSample[] {
  const valid: HarnessSample[] = []
  if (!Array.isArray(samples)) {
    errors.push('samples: must be an array')
    return valid
  }
  if (samples.length > MAX_SAMPLE_SERIES) {
    errors.push(`samples: at most ${MAX_SAMPLE_SERIES} series are allowed, got ${samples.length}`)
    return valid
  }
  const seen = new Set<string>()
  let total = 0
  for (const [index, sample] of samples.entries()) {
    const path = `samples[${index}]`
    if (!isPlainObject(sample)) {
      errors.push(`${path}: must be an object`)
      continue
    }
    const before = errors.length
    exactKeys(sample, ['name', 'unit', 'values', 'warmupExcluded', 'kind'], path, errors)
    if (typeof sample.name !== 'string' || !METRIC_NAME.test(sample.name)) errors.push(`${path}.name: must match ${METRIC_NAME}`)
    else if (seen.has(sample.name)) errors.push(`${path}.name: duplicate series name ${JSON.stringify(sample.name)}`)
    else seen.add(sample.name)
    if (typeof sample.unit !== 'string' || !UNIT.test(sample.unit)) errors.push(`${path}.unit: required short unit`)
    if (!isCount(sample.warmupExcluded) || sample.warmupExcluded > 1_000_000) errors.push(`${path}.warmupExcluded: must be a non-negative integer count`)
    if (sample.kind !== 'operation' && sample.kind !== 'frame' && sample.kind !== 'event') errors.push(`${path}.kind: must be operation|frame|event`)
    if (!Array.isArray(sample.values)) {
      errors.push(`${path}.values: must be an array of finite non-negative numbers`)
    } else if (sample.values.length > MAX_SAMPLE_VALUES) {
      errors.push(`${path}.values: at most ${MAX_SAMPLE_VALUES} values are allowed, got ${sample.values.length}`)
    } else {
      total += sample.values.length
      if (sample.values.some((v) => !isFiniteNumber(v) || v < 0)) errors.push(`${path}.values: must be an array of finite non-negative numbers`)
    }
    if (errors.length === before) valid.push(sample as unknown as HarnessSample)
  }
  if (total > MAX_TOTAL_SAMPLE_VALUES) errors.push(`samples: at most ${MAX_TOTAL_SAMPLE_VALUES} values in total are allowed, got ${total}`)
  return valid
}

function validateMetrics(metrics: unknown, samples: HarnessSample[], samplesValid: boolean, errors: string[]): void {
  if (!isPlainObject(metrics)) {
    errors.push('metrics: must be an object')
    return
  }
  const names = Object.keys(metrics)
  if (names.length > MAX_METRICS) {
    errors.push(`metrics: at most ${MAX_METRICS} metrics are allowed, got ${names.length}`)
    return
  }
  const bySample = new Map(samples.map((sample) => [sample.name, sample]))
  for (const name of names) {
    if (!METRIC_NAME.test(name)) {
      errors.push(`metrics: invalid metric name ${JSON.stringify(name)}`)
      continue
    }
    const before = errors.length
    validateMetricShape(name, metrics[name], errors)
    if (errors.length > before) continue
    const sample = bySample.get(name)
    if (sample) {
      const derived = deriveMetric(sample)
      if (canonicalize(metrics[name]) !== canonicalize(derived)) {
        errors.push(`metrics.${name}: does not match the metric derived from its raw samples (expected ${JSON.stringify(derived)})`)
      }
    } else if ((metrics[name] as Bag).status !== 'unavailable' && samplesValid) {
      errors.push(`metrics.${name}: has no backing sample series, so only status "unavailable" is allowed`)
    }
  }
  if (samplesValid) {
    for (const sample of samples) {
      if (!Object.prototype.hasOwnProperty.call(metrics, sample.name)) errors.push(`metrics.${sample.name}: missing for the sample series of the same name`)
    }
  }
}

function validateLedger(ledger: unknown, errors: string[]): void {
  if (!isPlainObject(ledger)) {
    errors.push('ledger: must be an object')
    return
  }
  exactKeys(ledger, ['main', 'worker', 'bitmap'], 'ledger', errors)
  for (const pool of ['main', 'worker', 'bitmap'] as const) {
    const accounting = ledger[pool]
    if (!isPlainObject(accounting)) {
      errors.push(`ledger.${pool}: must be an object`)
      continue
    }
    exactKeys(accounting, ['retained', 'peak', 'charged'], `ledger.${pool}`, errors)
    let counts = true
    for (const field of ['retained', 'peak', 'charged'] as const) {
      if (!isCount(accounting[field])) {
        errors.push(`ledger.${pool}.${field}: must be a non-negative integer byte count`)
        counts = false
      }
    }
    if (counts) {
      if ((accounting.retained as number) > (accounting.peak as number)) errors.push(`ledger.${pool}: retained must not exceed peak`)
      if ((accounting.peak as number) > (accounting.charged as number)) errors.push(`ledger.${pool}: peak must not exceed charged`)
    }
  }
}

/**
 * Runtime validation for ingested result JSON (bench collector, tests). This
 * is where the honesty invariants live; see the header comment.
 */
export function validateHarnessResult(input: unknown): ValidationOutcome {
  const errors: string[] = []
  if (!isPlainObject(input)) return { ok: false, errors: ['result: must be an object'] }
  scanStructure(input, 'result', 0, errors)
  if (errors.length > 0) return { ok: false, errors }

  exactKeys(input, ['schemaVersion', 'environment', 'samples', 'metrics', 'ledger', 'hash'], 'result', errors)
  if (input.schemaVersion !== 1) errors.push('schemaVersion: must be 1')
  validateEnvironment(input.environment, errors)
  const samplesBefore = errors.length
  const samples = validateSamples(input.samples, errors)
  validateMetrics(input.metrics, samples, errors.length === samplesBefore, errors)
  validateLedger(input.ledger, errors)
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

/** A metric/sample name that can never round-trip JSON as an own key. The
 * ingest validator rejects it too (METRIC_NAME/scanStructure); this guard
 * makes finalizeResult fail fast with a clear error instead of silently
 * dropping the series through prototype assignment. */
function assertSafeMetricName(name: string, role: string): void {
  if (name === '__proto__') throw new Error(`${role} name "__proto__" is forbidden (it cannot round-trip JSON safely)`)
}

/**
 * Build a complete result: derive metrics, attach the ledger, compute the
 * hash. Refuses to emit anything the ingest validator would reject, so a
 * device never exports a result the collector must bounce.
 */
export async function finalizeResult(draft: ResultDraft): Promise<HarnessResult> {
  // Null-prototype accumulator: inherited-looking names (constructor, toString)
  // behave as ordinary own keys, and no assignment can touch Object.prototype.
  const metrics: Record<string, HarnessMetric> = Object.create(null)
  for (const sample of draft.samples) {
    assertSafeMetricName(sample.name, 'sample series')
    if (Object.prototype.hasOwnProperty.call(metrics, sample.name)) throw new Error(`finalizeResult: duplicate sample series ${JSON.stringify(sample.name)}`)
    metrics[sample.name] = deriveMetric({ ...sample, values: [...sample.values] })
  }
  for (const [name, metric] of Object.entries(draft.extraMetrics ?? {})) {
    assertSafeMetricName(name, 'extra metric')
    if (Object.prototype.hasOwnProperty.call(metrics, name)) throw new Error(`finalizeResult: extra metric ${JSON.stringify(name)} collides with a sample series`)
    metrics[name] = metric
  }
  const withoutHash = {
    schemaVersion: 1 as const,
    environment: draft.environment,
    samples: draft.samples,
    metrics,
    ledger: draft.ledger,
  }
  const hash = await hashResult(withoutHash)
  const result = { ...withoutHash, hash }
  const verification = validateHarnessResult(JSON.parse(JSON.stringify(result)))
  if (!verification.ok) throw new Error(`finalizeResult: refusing to emit an invalid result: ${verification.errors.join('; ')}`)
  return result
}
