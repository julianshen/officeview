/** S8.4 threshold freeze: tests/fixtures/xlsx-viewer/performance-profile.json
 * must carry every spec §12.2 operation for both profiles, keep spec values
 * verbatim next to the frozen ones, never loosen past 2× without an amendment
 * (spec §12.1), and record the F4 headroom from its own measurement. Timing is
 * NOT asserted here (spec §12.6: no wall-clock assertions in unit CI). */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const PROFILE_PATH = join(process.cwd(), 'tests/fixtures/xlsx-viewer/performance-profile.json')

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
interface ProfileEntry {
  specMs: number
  frozenMs: number
  basis: 'spec' | 'measured'
  outcome: 'pass' | 'miss' | 'unmeasured'
  justification: string
  measurement: null | { scenario: string; series: string; summary: Summary }
}
interface Threshold { id: string; operation: string; boundary: string; desktop: ProfileEntry; mobile: ProfileEntry }
interface Profile {
  schemaVersion: number
  spec: string
  frozenAt: string
  maxLoosenRatio: number
  evidence: { doc: string; commit: string; browser: string; machine: string; profile: string }
  thresholds: Threshold[]
  headroom: Record<string, { ceilingMs: number; specEstimateMs: number; measuredP95Ms: number; headroomMs: number; scenario: string }>
}

const profile = JSON.parse(readFileSync(PROFILE_PATH, 'utf8')) as Profile

function checkSummary(summary: Summary, where: string): void {
  expect(Number.isInteger(summary.samples) && summary.samples >= 1, `${where}: samples`).toBe(true)
  expect(Number.isInteger(summary.warmupExcluded) && summary.warmupExcluded >= 0, `${where}: warmup`).toBe(true)
  expect(summary.p50 > 0 && summary.max >= summary.p50, `${where}: p50 <= max`).toBe(true)
  if (summary.p95 === null) expect(summary.samples, `${where}: p95 refused only below 30 samples`).toBeLessThan(30)
  else expect(summary.p50 <= summary.p95 && summary.p95 <= summary.max, `${where}: p50 <= p95 <= max`).toBe(true)
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
          expect(entry.measurement.scenario.length, where).toBeGreaterThan(0)
          const p95 = entry.measurement.summary.p95
          // pass/miss is against the FROZEN value, and only a real p95 can pass.
          if (entry.outcome === 'pass') expect(p95 !== null && p95 <= entry.frozenMs, `${where}: pass`).toBe(true)
          if (entry.outcome === 'miss') expect(p95 === null || p95 > entry.frozenMs, `${where}: miss`).toBe(true)
        }
      }
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
