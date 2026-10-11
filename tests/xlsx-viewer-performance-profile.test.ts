/** S8.4 threshold freeze: tests/fixtures/xlsx-viewer/performance-profile.json
 * must carry every spec §12.2 operation for both profiles, keep spec values
 * verbatim next to the frozen ones, never loosen past 2× without an amendment
 * (spec §12.1), and record the F4 headroom from its own measurement. Timing is
 * NOT asserted here (spec §12.6: no wall-clock assertions in unit CI). */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const PROFILE_PATH = join(process.cwd(), 'tests/fixtures/xlsx-viewer/performance-profile.json')
const RESULTS_PATH = join(process.cwd(), 'docs/validation/interactive-viewer-performance-results.json')
const MIN_OPERATION_SAMPLES = 30

/** Spec §12.2 acceptance table, copied verbatim (ms, p95). */
const SPEC_TABLE: Record<string, { desktop: number; mobile: number }> = {
  'warm-viewport-frame': { desktop: 16.7, mobile: 33.3 },
  'input-acknowledgment': { desktop: 50, mobile: 50 },
  'numeric-sort-100k': { desktop: 100, mobile: 300 },
  'numeric-sort-1m': { desktop: 1000, mobile: 3000 },
  'numeric-filter-1m': { desktop: 100, mobile: 300 },
  'natural-text-sort-1m': { desktop: 4000, mobile: 10000 },
  'cold-value-index-100k': { desktop: 750, mobile: 2000 },
  'progressive-first-viewport': { desktop: 1500, mobile: 3000 },
}
const MAX_LOOSEN_RATIO = 2

interface Summary { samples: number; warmupExcluded: number; p50: number; p95: number | null; max: number }
interface Measurement {
  scenario: string
  series: string
  /** False: resident harness data was not charged to the worker budget. */
  admitted: boolean
  /** Prepared key columns admitted by the session (null for non-sort scenarios). */
  preparedColumns: number | null
  keys: number | null
  summary: Summary
}
interface ProfileEntry {
  specMs: number
  frozenMs: number
  basis: 'spec' | 'measured'
  outcome: 'pass' | 'miss' | 'unmeasured'
  justification: string
  measurement: Measurement | null
}
interface Threshold { id: string; operation: string; boundary: string; desktop: ProfileEntry; mobile: ProfileEntry }
interface Profile {
  schemaVersion: number
  spec: string
  frozenAt: string
  maxLoosenRatio: number
  evidence: { doc: string; commit: string; browser: string; machine: string; profile: string }
  thresholds: Threshold[]
  diagnostics: (Measurement & { threshold: string; reason: string })[]
  headroom: Record<string, { ceilingMs: number; specEstimateMs: number; measuredP95Ms: number; headroomMs: number; scenario: string }>
}

const profile = JSON.parse(readFileSync(PROFILE_PATH, 'utf8')) as Profile
interface ResultRow {
  admitted: boolean | null
  note: string
  scenario: { keys: number[]; admitted: boolean } | null
  series: Record<string, { samples: number; warmupExcluded: number; p50: number; p95: number | null; max: number }>
}
const results = JSON.parse(readFileSync(RESULTS_PATH, 'utf8')) as Record<string, ResultRow>
const WARM_SORTS = ['numeric-sort-100k', 'numeric-sort-1m']

/** A frozen summary must be exactly the retained result it cites. */
function checkAgainstResults(measurement: Measurement, where: string): void {
  const row = results[measurement.scenario]
  expect(row, `${where}: results JSON has scenario ${measurement.scenario}`).toBeDefined()
  const series = row.series[measurement.series]
  expect(series, `${where}: results JSON has series ${measurement.series}`).toBeDefined()
  expect(measurement.summary, `${where}: summary matches results JSON`).toEqual({
    samples: series.samples, warmupExcluded: series.warmupExcluded, p50: series.p50, p95: series.p95, max: series.max,
  })
  // Runs before resident accounting carry no admitted flag: they are admitted
  // only if no resident harness data existed to charge.
  const admitted = row.admitted ?? (!/UNADMITTED/.test(row.note) && row.series['prep:resident-source-bytes'] === undefined && row.series['prep:resident-source-materialize'] === undefined)
  expect(measurement.admitted, `${where}: admitted flag matches the run`).toBe(admitted)
  if (measurement.preparedColumns !== null) {
    expect(measurement.preparedColumns, `${where}: prepared columns match the run`).toBe(row.series['prep:prepared-columns-admitted']?.p50)
  }
}

function checkSummary(summary: Summary, where: string): void {
  expect(Number.isInteger(summary.samples) && summary.samples >= 1, `${where}: samples`).toBe(true)
  expect(Number.isInteger(summary.warmupExcluded) && summary.warmupExcluded >= 0, `${where}: warmup`).toBe(true)
  expect(summary.p50 > 0 && summary.max >= summary.p50, `${where}: p50 <= max`).toBe(true)
  if (summary.p95 === null) expect(summary.samples, `${where}: p95 refused only below 30 samples`).toBeLessThan(MIN_OPERATION_SAMPLES)
  else {
    // Spec §12.6: a p95 needs at least 30 operation samples.
    expect(summary.samples, `${where}: a p95 needs >= ${MIN_OPERATION_SAMPLES} samples`).toBeGreaterThanOrEqual(MIN_OPERATION_SAMPLES)
    expect(summary.p50 <= summary.p95 && summary.p95 <= summary.max, `${where}: p50 <= p95 <= max`).toBe(true)
  }
}

describe('performance-profile.json (S8.4 threshold freeze)', () => {
  it('declares its schema, spec source and evidence record', () => {
    expect(profile.schemaVersion).toBe(1)
    expect(profile.spec).toContain('§12.2')
    expect(profile.frozenAt).toMatch(/^\d{4}-\d{2}-\d{2}$/)
    expect(profile.maxLoosenRatio).toBe(MAX_LOOSEN_RATIO)
    expect(profile.evidence.doc).toBe('docs/validation/interactive-viewer-performance.md')
    expect(profile.evidence.commit).toMatch(/^[0-9a-f]{7,40}$/)
    expect(profile.evidence.profile).toBe('D')
    expect(profile.evidence.browser.length).toBeGreaterThan(0)
    expect(profile.evidence.machine.length).toBeGreaterThan(0)
  })

  it('lists every §12.2 operation exactly once with the spec values verbatim', () => {
    expect(profile.thresholds.map(t => t.id).sort()).toEqual(Object.keys(SPEC_TABLE).sort())
    for (const threshold of profile.thresholds) {
      expect(threshold.operation.length, threshold.id).toBeGreaterThan(0)
      expect(threshold.boundary.length, threshold.id).toBeGreaterThan(0)
      expect(threshold.desktop.specMs, threshold.id).toBe(SPEC_TABLE[threshold.id].desktop)
      expect(threshold.mobile.specMs, threshold.id).toBe(SPEC_TABLE[threshold.id].mobile)
    }
  })

  it('freezes each value at spec or a justified measured replacement, never loosened past 2×', () => {
    for (const threshold of profile.thresholds) {
      for (const key of ['desktop', 'mobile'] as const) {
        const entry = threshold[key], where = `${threshold.id}.${key}`
        expect(entry.frozenMs > 0, where).toBe(true)
        expect(entry.frozenMs, `${where}: loosening past ${MAX_LOOSEN_RATIO}× needs a spec amendment`).toBeLessThanOrEqual(entry.specMs * MAX_LOOSEN_RATIO)
        expect(['spec', 'measured'], where).toContain(entry.basis)
        expect(['pass', 'miss', 'unmeasured'], where).toContain(entry.outcome)
        expect(entry.justification.trim().length, `${where}: justification`).toBeGreaterThan(0)
        if (entry.basis === 'spec') expect(entry.frozenMs, where).toBe(entry.specMs)
        if (entry.basis === 'measured') expect(entry.measurement, `${where}: a measured replacement cites its measurement`).not.toBeNull()
        if (entry.outcome === 'unmeasured') expect(entry.measurement, where).toBeNull()
        if (entry.measurement) {
          checkSummary(entry.measurement.summary, where)
          checkAgainstResults(entry.measurement, where)
          // Only an admitted-profile run can pass a threshold.
          if (entry.outcome === 'pass') expect(entry.measurement.admitted, `${where}: pass needs an admitted run`).toBe(true)
          expect(entry.measurement.scenario.length, where).toBeGreaterThan(0)
          const p95 = entry.measurement.summary.p95
          // pass/miss is against the FROZEN value, and only a real p95 can pass.
          if (entry.outcome === 'pass') expect(p95 !== null && p95 <= entry.frozenMs, `${where}: pass`).toBe(true)
          if (entry.outcome === 'miss') expect(p95 === null || p95 > entry.frozenMs, `${where}: miss`).toBe(true)
        }
      }
    }
  })

  it('bases the warm sort thresholds on genuinely prepared key columns', () => {
    for (const id of WARM_SORTS) {
      const measurement = profile.thresholds.find(t => t.id === id)!.desktop.measurement!
      expect(measurement.preparedColumns, id).not.toBeNull()
      expect(measurement.preparedColumns, `${id}: every key column was prepared`).toBe(measurement.keys)
      expect(measurement.preparedColumns! >= 1, id).toBe(true)
      expect(measurement.admitted, id).toBe(true)
    }
  })

  it('keeps fallback diagnostics separate and consistent with the retained results', () => {
    for (const diagnostic of profile.diagnostics) {
      const where = `diagnostic ${diagnostic.scenario}`
      expect(profile.thresholds.some(t => t.id === diagnostic.threshold), where).toBe(true)
      expect(diagnostic.reason.trim().length, where).toBeGreaterThan(0)
      checkSummary(diagnostic.summary, where)
      checkAgainstResults(diagnostic, where)
      for (const threshold of profile.thresholds) expect(threshold.desktop.measurement?.scenario === diagnostic.scenario && threshold.desktop.measurement?.series === diagnostic.series, where).toBe(false)
    }
  })

  it('claims no mobile measurement: no physical iPad/device run exists', () => {
    for (const threshold of profile.thresholds) {
      expect(threshold.mobile.measurement, threshold.id).toBeNull()
      expect(threshold.mobile.outcome, threshold.id).toBe('unmeasured')
      expect(threshold.mobile.basis, threshold.id).toBe('spec')
    }
  })

  it('measures the S8.4 kernel operations on the desktop profile', () => {
    const measured = ['numeric-sort-100k', 'numeric-sort-1m', 'numeric-filter-1m', 'natural-text-sort-1m', 'cold-value-index-100k']
    for (const id of measured) {
      const entry = profile.thresholds.find(t => t.id === id)!.desktop
      expect(entry.measurement, id).not.toBeNull()
      expect(entry.outcome, id).not.toBe('unmeasured')
    }
  })

  it('records the 1M natural-text-sort headroom from its own desktop measurement', () => {
    const headroom = profile.headroom['natural-text-sort-1m']
    const entry = profile.thresholds.find(t => t.id === 'natural-text-sort-1m')!.desktop
    expect(headroom.ceilingMs).toBe(4000)
    expect(headroom.specEstimateMs).toBe(3100)
    expect(headroom.measuredP95Ms).toBe(entry.measurement!.summary.p95 ?? entry.measurement!.summary.max)
    expect(headroom.headroomMs).toBeCloseTo(headroom.ceilingMs - headroom.measuredP95Ms, 6)
    expect(headroom.scenario).toBe(entry.measurement!.scenario)
  })
})
