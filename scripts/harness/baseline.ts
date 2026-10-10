/**
 * Baseline loading and comparison. A baseline is validated and checked for
 * comparability BEFORE any browser launches, and comparison failures can never
 * cost a run its artifacts (callers record the error and still write them).
 */
import { verifyHarnessResult } from '../../src/debug/harness/result'
import type { HarnessMetric, HarnessResult } from '../../src/debug/harness/types'

export interface BaselineExpectation {
  browserLabel: string
  profile: string
  requestedFixtureId: string
  mode: string
}

export interface ComparisonRow {
  metric: string
  baselineP50: number | null
  p50: number | null
  delta: number | null
}

export interface BaselineComparison {
  rows: ComparisonRow[]
  warnings: string[]
}

export async function parseBaseline(text: string): Promise<HarnessResult[]> {
  let raw: unknown
  try {
    raw = JSON.parse(text)
  } catch {
    throw new Error('baseline is not valid JSON')
  }
  const list = Array.isArray(raw) ? raw : typeof raw === 'object' && raw !== null && Array.isArray((raw as { results?: unknown }).results) ? (raw as { results: unknown[] }).results : null
  if (!list || list.length === 0) throw new Error('baseline contains no results')
  for (const [index, entry] of list.entries()) {
    const verification = await verifyHarnessResult(entry)
    if (!verification.ok) throw new Error(`baseline result ${index} is invalid: ${verification.errors.join('; ')}`)
  }
  return list as HarnessResult[]
}

/** Differences knowable before launching; any entry means "not comparable". */
export function checkBaselineComparable(baseline: HarnessResult, expected: BaselineExpectation): string[] {
  const env = baseline.environment
  const reasons: string[] = []
  if (env.browser.name !== expected.browserLabel) reasons.push(`baseline browser ${env.browser.name} differs from this run's ${expected.browserLabel}`)
  if (env.profile !== expected.profile) reasons.push(`baseline profile ${env.profile} differs from this run's ${expected.profile}`)
  if (env.requestedFixtureId !== expected.requestedFixtureId) reasons.push(`baseline fixture ${env.requestedFixtureId} differs from this run's ${expected.requestedFixtureId}`)
  if (env.mode !== expected.mode) reasons.push(`baseline mode ${env.mode} differs from this run's ${expected.mode}`)
  return reasons
}

function p50Of(metric: HarnessMetric | undefined): number | null {
  return metric && (metric.status === 'ok' || metric.status === 'insufficient-samples') ? metric.p50 : null
}

export function compareResults(baseline: HarnessResult, current: HarnessResult): BaselineComparison {
  const rows: ComparisonRow[] = []
  for (const [name, metric] of Object.entries(current.metrics)) {
    const baselineP50 = p50Of(baseline.metrics[name])
    const p50 = p50Of(metric)
    rows.push({ metric: name, baselineP50, p50, delta: baselineP50 !== null && p50 !== null ? p50 - baselineP50 : null })
  }
  const a = baseline.environment
  const b = current.environment
  const warnings: string[] = []
  const differs = (label: string, left: unknown, right: unknown): void => {
    if (JSON.stringify(left) !== JSON.stringify(right)) warnings.push(`${label} differs: baseline ${JSON.stringify(left)}, current ${JSON.stringify(right)}`)
  }
  differs('dpr', a.dpr, b.dpr)
  differs('viewport', a.viewport, b.viewport)
  differs('browser engine', a.browser.engine, b.browser.engine)
  differs('browser version', a.browser.version, b.browser.version)
  differs('backend', a.backend, b.backend)
  differs('source adapter', a.source.adapter, b.source.adapter)
  differs('os platform', a.os.platform, b.os.platform)
  differs('evidence', a.evidence, b.evidence)
  differs('build commit (informational: a baseline normally predates the change)', a.buildCommit, b.buildCommit)
  return { rows, warnings }
}
